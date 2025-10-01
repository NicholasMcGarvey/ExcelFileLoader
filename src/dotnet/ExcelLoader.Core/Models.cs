
using System.Data;
using Microsoft.Data.SqlClient;

namespace ExcelLoader.Core;

public record LoadContext(
    string FilePath,
    string FileGroupCode,
    string? EntityCode,
    string? LoadTypeName,
    int? FileSpecId = null,
    bool DryRun = false,
    string? SheetHint = null
);

public record TargetTableInfo(
    int TargetTableId,
    string Name,
    string KeyStrategy,
    bool StageFirst,
    string? UniqueKeyCsv
);

public record TargetFieldInfo(
    int TargetFieldId,
    int TargetTableId,
    string Name,
    string DataType,
    bool IsNullable,
    int Ordinal,
    bool RequiredForKey
);

public record MappingRow(
    int FieldMapId,
    string SourceHeader,
    int TargetTableId,
    int TargetFieldId,
    string? TransformChain,
    string? DefaultValue,
    bool Required
);

public interface IExcelParser {
    ParsedFile Parse(string path, int headerRow, int firstDataRow, string? sheetHint);
}

public sealed class ParsedFile {
    public string? SheetName { get; init; }
    public List<string> Headers { get; } = new();
    public List<Dictionary<string, object?>> Rows { get; } = new();
}

public interface ITransformer {
    object? Apply(string? chain, object? value);
}

public interface ILoaderEngine {
    Task<LoadResult> RunAsync(SqlConnection conn, LoadContext ctx, CancellationToken ct = default);
}

public sealed record LoadResult(long LoadRunId, int RowsRead, int RowsLoaded, int ErrorCount, string? Message);
