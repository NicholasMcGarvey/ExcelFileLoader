
using ClosedXML.Excel;
using ExcelLoader.Core;

namespace ExcelLoader.Excel;

public sealed class ClosedXmlParser : IExcelParser
{
    public ParsedFile Parse(string path, int headerRow, int firstDataRow, string? sheetHint)
    {
        using var wb = new XLWorkbook(path);
        var ws = sheetHint != null ? wb.Worksheets.FirstOrDefault(s => s.Name.Equals(sheetHint, StringComparison.OrdinalIgnoreCase)) : null;
        ws ??= wb.Worksheets.First();

        var pf = new ParsedFile { SheetName = ws.Name };

        // Header
        var hr = ws.Row(headerRow);
        foreach (var cell in hr.CellsUsed())
            pf.Headers.Add(cell.GetString().Trim());

        // Rows
        var lastCol = pf.Headers.Count;
        for (int r = firstDataRow; r <= ws.LastRowUsed().RowNumber(); r++)
        {
            var row = ws.Row(r);
            if (row.IsEmpty()) continue;
            var obj = new Dictionary<string, object?>();
            var hasAny = false;
            for (int c = 1; c <= lastCol; c++)
            {
                var h = pf.Headers[c-1];
                var cell = row.Cell(c);
                object? v = null;
                if (cell.DataType == XLDataType.DateTime) v = cell.GetDateTime();
                else if (cell.DataType == XLDataType.Number) v = cell.GetDouble();
                else v = cell.GetString();

                if (v is string sv) v = sv.Trim();
                if (v is string s2 && string.IsNullOrWhiteSpace(s2)) v = null;
                if (v != null) hasAny = true;
                obj[h] = v;
            }
            if (hasAny) pf.Rows.Add(obj);
        }
        return pf;
    }
}
