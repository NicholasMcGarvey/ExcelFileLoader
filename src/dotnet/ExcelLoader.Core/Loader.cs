
using System.Data;
using System.Text;
using Dapper;
using Microsoft.Data.SqlClient;

namespace ExcelLoader.Core;

public sealed class LoaderEngine : ILoaderEngine
{
    private readonly IExcelParser _parser;
    private readonly ITransformer _tx;

    public LoaderEngine(IExcelParser parser, ITransformer tx)
    {
        _parser = parser;
        _tx = tx;
    }

    public async Task<LoadResult> RunAsync(SqlConnection conn, LoadContext ctx, CancellationToken ct = default)
    {
        await conn.OpenAsync(ct);

        // Look up Entity/LoadType/FileSpec
        var entityId = await conn.ExecuteScalarAsync<int?>(
            "SELECT EntityId FROM repo.Entity WHERE Code=@code", new { code = ctx.EntityCode });
        var loadTypeId = await conn.ExecuteScalarAsync<int?>(
            "SELECT LoadTypeId FROM repo.LoadType WHERE Name=@name", new { name = ctx.LoadTypeName });

        int fileSpecId;
        if (ctx.FileSpecId.HasValue) fileSpecId = ctx.FileSpecId.Value;
        else
        {
            var fileGroupId = await conn.ExecuteScalarAsync<int>(
                "SELECT FileGroupId FROM repo.FileGroup WHERE Code=@code", new { code = ctx.FileGroupCode });
            // Choose active FileSpec by group; for real usage, filter by pattern or name
            fileSpecId = await conn.ExecuteScalarAsync<int>(@"
                SELECT TOP 1 FileSpecId FROM repo.FileSpec 
                WHERE FileGroupId=@fg AND Active=1 ORDER BY FileSpecId",
                new { fg = fileGroupId });
        }

        var spec = await conn.QuerySingleAsync<(int HeaderRow, int FirstDataRow, string? SheetHint)>(@"
            SELECT HeaderRow, FirstDataRow, SheetHint FROM repo.FileSpec WHERE FileSpecId=@id", new { id = fileSpecId });

        var parsed = _parser.Parse(ctx.FilePath, spec.HeaderRow, spec.FirstDataRow, ctx.SheetHint ?? spec.SheetHint);
        var headers = parsed.Headers.Select(h => h.Trim()).ToList();

        // Load effective mappings
        var maps = (await conn.QueryAsync<MappingRow>(
            "SELECT * FROM repo.fn_GetEffectiveFieldMap(@fs, @ent, @lt)",
            new { fs = fileSpecId, ent = entityId, lt = loadTypeId })).ToList();

        if (maps.Count == 0)
            throw new InvalidOperationException("No mappings defined for FileSpec.");

        // Pull target tables/fields
        var targetTableIds = maps.Select(m => m.TargetTableId).Distinct().ToArray();
        var tables = (await conn.QueryAsync<TargetTableInfo>(@"
            SELECT TargetTableId, Name, KeyStrategy, StageFirst, UniqueKeyCsv
            FROM repo.TargetTable WHERE TargetTableId IN @ids", new { ids = targetTableIds })).ToDictionary(t => t.TargetTableId);

        var fields = (await conn.QueryAsync<TargetFieldInfo>(@"
            SELECT TargetFieldId, TargetTableId, Name, DataType, IsNullable, Ordinal, RequiredForKey
            FROM repo.TargetField WHERE TargetTableId IN @ids", new { ids = targetTableIds })).ToList();

        // Audit start
        var loadRunId = await conn.ExecuteScalarAsync<long>(@"
            INSERT repo.LoadRun(FileSpecId, EntityId, LoadTypeId, SourcePath, DetectedSheet, RowCountRead, RowCountLoaded, ErrorCount, [Status]) 
            OUTPUT INSERTED.LoadRunId
            VALUES (@fs, @ent, @lt, @path, @sheet, NULL, NULL, NULL, 'RUNNING')",
            new { fs = fileSpecId, ent = entityId, lt = loadTypeId, path = ctx.FilePath, sheet = parsed.SheetName });

        var errors = new List<(int? RowNumber, string? ColumnName, string Code, string Message, string? Raw)>();

        // Build per-table row buffers
        var byTable = targetTableIds.ToDictionary(id => id, id => new List<Dictionary<string, object?>>());

        int rowsRead = 0, rowsLoaded = 0;
        foreach (var row in parsed.Rows)
        {
            rowsRead++;
            // Create mapped outputs per table
            foreach (var tableId in targetTableIds)
            {
                var tableFields = fields.Where(f => f.TargetTableId == tableId).ToDictionary(f => f.TargetFieldId);
                var tableMaps = maps.Where(m => m.TargetTableId == tableId).ToList();
                var outRow = new Dictionary<string, object?>();

                foreach (var m in tableMaps)
                {
                    var srcHeader = m.SourceHeader.Trim();
                    row.TryGetValue(srcHeader, out var raw);
                    var value = _tx.Apply(m.TransformChain, raw);
                    if ((value is null || (value is string s && string.IsNullOrWhiteSpace(s))) 
                        && !string.IsNullOrWhiteSpace(m.DefaultValue))
                        value = m.DefaultValue;

                    var field = tableFields[m.TargetFieldId];
                    outRow[field.Name] = value;

                    if (m.Required && (value is null || (value is string sv && string.IsNullOrWhiteSpace(sv))))
                        errors.Add((rowsRead, field.Name, "REQUIRED", $"Field '{field.Name}' is required.", raw?.ToString()));
                }

                byTable[tableId].Add(outRow);
            }
        }

        // If dry-run, only persist audit + errors
        if (ctx.DryRun)
        {
            await PersistErrors(conn, loadRunId, errors);
            await conn.ExecuteAsync(@"
                UPDATE repo.LoadRun SET RowCountRead=@rr, RowCountLoaded=0, ErrorCount=@ec,
                       [Status]=CASE WHEN @ec=0 THEN 'SUCCEEDED' ELSE 'PARTIAL' END,
                       CompletedUtc=SYSUTCDATETIME()
                WHERE LoadRunId=@id",
                new { rr = rowsRead, ec = errors.Count, id = loadRunId });
            return new LoadResult(loadRunId, rowsRead, 0, errors.Count, "Dry run");
        }

        // Execute per table: stage + merge/append
        using var tx = conn.BeginTransaction();
        try
        {
            foreach (var tableId in targetTableIds)
            {
                var info = tables[tableId];
                var tableFields = fields.Where(f => f.TargetTableId == tableId).OrderBy(f => f.Ordinal).ToList();
                var data = byTable[tableId];

                // Create #temp with typed columns
                var tempName = $"#Stage_{tableId}_{Guid.NewGuid().ToString("N")[..8]}";
                var createSql = $"CREATE TABLE {tempName} ({string.Join(", ", tableFields.Select(f => $"[{f.Name}] {f.DataType} {(f.IsNullable ? "NULL" : "NOT NULL")}"))});";
                await conn.ExecuteAsync(createSql, transaction: tx);

                // Bulk copy
                using (var bulk = new Microsoft.Data.SqlClient.SqlBulkCopy(conn, SqlBulkCopyOptions.Default, tx))
                {
                    bulk.DestinationTableName = tempName;
                    var dt = new DataTable();
                    foreach (var f in tableFields) dt.Columns.Add(f.Name, typeof(object));
                    foreach (var r in data)
                    {
                        var dr = dt.NewRow();
                        foreach (var f in tableFields) dr[f.Name] = r.TryGetValue(f.Name, out var v) && v != null ? v : DBNull.Value;
                        dt.Rows.Add(dr);
                    }
                    await bulk.WriteToServerAsync(dt, CancellationToken.None);
                }

                // Merge/Append
                if (string.Equals(info.KeyStrategy, "APPEND", StringComparison.OrdinalIgnoreCase))
                {
                    var cols = string.Join(", ", tableFields.Select(f => $"[{f.Name}]"));
                    var sql = $"INSERT INTO {info.Name} ({cols}) SELECT {cols} FROM {tempName};";
                    rowsLoaded += await conn.ExecuteAsync(sql, transaction: tx);
                }
                else // UPSERT default
                {
                    var keyCols = (info.UniqueKeyCsv ?? string.Join(',', tableFields.Where(f => f.RequiredForKey).Select(f => f.Name)))
                                  .Split(',', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries);

                    var nonKeys = tableFields.Where(f => !keyCols.Contains(f.Name, StringComparer.OrdinalIgnoreCase)).Select(f => f.Name).ToArray();

                    var onClause = string.Join(" AND ", keyCols.Select(k => $"T.[{k}] = S.[{k}]"));
                    var updateSet = string.Join(", ", nonKeys.Select(c => $"T.[{c}] = S.[{c}]"));
                    var insertCols = string.Join(", ", tableFields.Select(f => $"[{f.Name}]"));
                    var insertVals = string.Join(", ", tableFields.Select(f => $"S.[{f.Name}]"));

                    var sql = $@"
MERGE {info.Name} AS T
USING {tempName} AS S
  ON ({onClause})
WHEN MATCHED THEN UPDATE SET {updateSet}
WHEN NOT MATCHED THEN INSERT ({insertCols}) VALUES ({insertVals});";

                    rowsLoaded += await conn.ExecuteAsync(sql, transaction: tx);
                }
            }

            await PersistErrors(conn, loadRunId, errors, tx);
            await conn.ExecuteAsync(@"
                UPDATE repo.LoadRun SET RowCountRead=@rr, RowCountLoaded=@rl, ErrorCount=@ec,
                       [Status]=CASE WHEN @ec=0 THEN 'SUCCEEDED' ELSE 'PARTIAL' END,
                       CompletedUtc=SYSUTCDATETIME()
                WHERE LoadRunId=@id",
                new { rr = rowsRead, rl = rowsLoaded, ec = errors.Count, id = loadRunId }, tx);

            tx.Commit();
        }
        catch (Exception ex)
        {
            tx.Rollback();
            await conn.ExecuteAsync(@"
                UPDATE repo.LoadRun SET [Status]='FAILED', [Message]=@msg, CompletedUtc=SYSUTCDATETIME()
                WHERE LoadRunId=@id", new { id = loadRunId, msg = ex.ToString()});
            throw;
        }

        return new LoadResult(loadRunId, rowsRead, rowsLoaded, errors.Count, null);
    }

    private static async Task PersistErrors(SqlConnection conn, long loadRunId, List<(int? RowNumber, string? ColumnName, string Code, string Message, string? Raw)> errors, SqlTransaction? tx = null)
    {
        if (errors.Count == 0) return;
        var sql = @"INSERT repo.LoadRunError(LoadRunId, RowNumber, ColumnName, ErrorCode, ErrorMessage, RawValue)
                    VALUES (@LoadRunId, @RowNumber, @ColumnName, @ErrorCode, @ErrorMessage, @RawValue)";
        foreach (var e in errors)
            await conn.ExecuteAsync(sql, new { LoadRunId = loadRunId, RowNumber = e.RowNumber, ColumnName = e.ColumnName, ErrorCode = e.Code, ErrorMessage = e.Message, RawValue = e.Raw }, tx);
    }
}
