
using System.Globalization;
using System.Text;
using System.Text.RegularExpressions;

namespace ExcelLoader.Core;

public sealed class Transformer : ITransformer
{
    public object? Apply(string? chain, object? value)
    {
        if (string.IsNullOrWhiteSpace(chain)) return value;
        var parts = chain.Split('|', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries);
        object? v = value;

        foreach (var part in parts)
        {
            var call = part.Trim();
            if (call.Equals("trim", StringComparison.OrdinalIgnoreCase))
            {
                v = v?.ToString()?.Trim();
            }
            else if (call.Equals("upper", StringComparison.OrdinalIgnoreCase))
            {
                v = v?.ToString()?.ToUpperInvariant();
            }
            else if (call.Equals("lower", StringComparison.OrdinalIgnoreCase))
            {
                v = v?.ToString()?.ToLowerInvariant();
            }
            else if (call.StartsWith("coalesce(", StringComparison.OrdinalIgnoreCase))
            {
                var arg = ExtractArgs(call);
                if (IsNullOrEmpty(v)) v = arg;
            }
            else if (call.Equals("strip_nonnum", StringComparison.OrdinalIgnoreCase))
            {
                var s = v?.ToString();
                if (s is not null) v = Regex.Replace(s, "[^0-9]", "");
            }
            else if (call.StartsWith("int", StringComparison.OrdinalIgnoreCase))
            {
                v = ToInt(v);
            }
            else if (call.StartsWith("decimal", StringComparison.OrdinalIgnoreCase))
            {
                // decimal(p,s) or just decimal
                v = ToDecimal(v);
            }
            else if (call.Equals("money", StringComparison.OrdinalIgnoreCase))
            {
                v = ToDecimal(v);
            }
            else if (call.StartsWith("bool", StringComparison.OrdinalIgnoreCase))
            {
                v = ToBool(v);
            }
            else if (call.StartsWith("date(", StringComparison.OrdinalIgnoreCase))
            {
                var fmt = ExtractArgs(call);
                v = ToDate(v, fmt);
            }
            else
            {
                // Unknown transform => pass-through
            }
        }

        return v;
    }

    private static bool IsNullOrEmpty(object? v)
        => v is null || (v is string s && string.IsNullOrWhiteSpace(s));

    private static object? ToInt(object? v)
    {
        if (v is null) return null;
        if (v is int i) return i;
        if (int.TryParse(v.ToString(), NumberStyles.Integer, CultureInfo.InvariantCulture, out var iv))
            return iv;
        var dec = ToDecimal(v);
        return dec is null ? null : Convert.ToInt32(dec);
    }

    private static object? ToDecimal(object? v)
    {
        if (v is null) return null;
        if (v is decimal d) return d;
        var s = v.ToString();
        if (string.IsNullOrWhiteSpace(s)) return null;
        s = s.Replace(",", "");
        if (decimal.TryParse(s, NumberStyles.Any, CultureInfo.InvariantCulture, out var dv))
            return dv;
        return null;
    }

    private static object? ToBool(object? v)
    {
        if (v is null) return null;
        var s = v.ToString()?.Trim().ToLowerInvariant();
        if (string.IsNullOrEmpty(s)) return null;
        return s switch
        {
            "1" or "y" or "yes" or "true" => true,
            "0" or "n" or "no" or "false" => false,
            _ => null
        };
    }

    private static object? ToDate(object? v, string? format)
    {
        if (v is null) return null;
        if (v is DateTime dt) return dt.Date;
        var s = v.ToString();
        if (string.IsNullOrWhiteSpace(s)) return null;

        // Try provided format first
        if (!string.IsNullOrWhiteSpace(format) &&
            DateTime.TryParseExact(s, format, CultureInfo.InvariantCulture, DateTimeStyles.None, out var d0))
            return d0.Date;

        // Try common formats
        string[] fmts = new[] { "M/d/yyyy", "M/d/yy", "yyyy-MM-dd", "MM/dd/yyyy", "dd-MMM-yyyy" };
        if (DateTime.TryParseExact(s, fmts, CultureInfo.InvariantCulture, DateTimeStyles.None, out var d1))
            return d1.Date;

        // Excel serial?
        if (double.TryParse(s, NumberStyles.Any, CultureInfo.InvariantCulture, out var serial))
        {
            // Excel serial date: days since 1899-12-30
            var origin = new DateTime(1899, 12, 30);
            return origin.AddDays(serial).Date;
        }
        return null;
    }

    private static string? ExtractArgs(string call)
    {
        var i = call.IndexOf('(');
        var j = call.LastIndexOf(')');
        if (i >= 0 && j > i) return call[(i+1)..j].Trim('\'', ' ');
        return null;
    }
}
