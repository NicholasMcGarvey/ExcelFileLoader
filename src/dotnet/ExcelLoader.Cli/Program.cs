
using System.CommandLine;
using ExcelLoader.Core;
using ExcelLoader.Excel;
using Microsoft.Data.SqlClient;

var connOpt = new Option<string>("--connection", "SQL Server connection string") { IsRequired = true };
var fileOpt = new Option<string>("--file", "Path to Excel file") { IsRequired = true };
var fileGroupOpt = new Option<string>("--fileGroup", description: "FileGroup code, e.g. COL") { IsRequired = true };
var entityOpt = new Option<string?>("--entity", () => null, "Entity code (optional)");
var loadTypeOpt = new Option<string?>("--loadType", () => null, "Load type name (daily, cyclical, etc.)");
var sheetOpt = new Option<string?>("--sheet", () => null, "Sheet name hint");
var dryRunOpt = new Option<bool>("--dryRun", () => false, "If set, no DB writes");

var root = new RootCommand("ExcelLoader CLI");
root.AddOption(connOpt);
root.AddOption(fileOpt);
root.AddOption(fileGroupOpt);
root.AddOption(entityOpt);
root.AddOption(loadTypeOpt);
root.AddOption(sheetOpt);
root.AddOption(dryRunOpt);

root.SetHandler(async (string connection, string file, string fileGroup, string? entity, string? loadType, string? sheet, bool dryRun) =>
{
    var parser = new ClosedXmlParser();
    var tx = new Transformer();
    var engine = new LoaderEngine(parser, tx);

    await using var conn = new SqlConnection(connection);
    var ctx = new LoadContext(
        FilePath: file,
        FileGroupCode: fileGroup,
        EntityCode: entity,
        LoadTypeName: loadType,
        DryRun: dryRun,
        SheetHint: sheet
    );

    var result = await engine.RunAsync(conn, ctx);
    Console.WriteLine($"{(dryRun ? "[DRY]" : "[RUN]")} LoadRunId={result.LoadRunId}, RowsRead={result.RowsRead}, RowsLoaded={result.RowsLoaded}, Errors={result.ErrorCount}");
}, connOpt, fileOpt, fileGroupOpt, entityOpt, loadTypeOpt, sheetOpt, dryRunOpt);

return await root.InvokeAsync(args);
