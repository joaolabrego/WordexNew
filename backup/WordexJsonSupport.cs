using System;
using System.Collections.Generic;
using System.Data;
using System.Globalization;
using System.Linq;
using System.Text.Json;
using System.Text.Json.Serialization;

namespace Wordex.Json;

internal static class WordexJsonSupport
{
    internal static readonly CultureInfo PtBr = CultureInfo.GetCultureInfo("pt-BR");
    internal const string FormatoCount = "###,###,###,##0";

    internal static JsonSerializerOptions SerializerOptions(bool indented = false) =>
        new()
        {
            WriteIndented = indented,
            DefaultIgnoreCondition = JsonIgnoreCondition.WhenWritingNull,
            Encoder = System.Text.Encodings.Web.JavaScriptEncoder.UnsafeRelaxedJsonEscaping
        };

    internal static string ToJson(object? value, bool indented = false) =>
        JsonSerializer.Serialize(value, SerializerOptions(indented));

    internal static string ErrorArray(string message) =>
        ToJson(new[] { new Dictionary<string, object?> { ["Erro"] = message } });

    internal static string ErrorObject(string message) =>
        ToJson(new Dictionary<string, object?> { ["Erro"] = message });

    internal static string[] SplitList(string? csv)
    {
        if (string.IsNullOrWhiteSpace(csv))
        {
            return Array.Empty<string>();
        }

        return csv
            .Split(',')
            .Select(part => part.Trim())
            .Where(part => part.Length > 0)
            .ToArray();
    }

    internal static IReadOnlyList<(string Name, object? Value)> ParseCriteria(params object?[] criterios)
    {
        if (criterios == null || criterios.Length == 0)
        {
            return Array.Empty<(string, object?)>();
        }

        if (criterios.Length % 2 != 0)
        {
            throw new ArgumentException("Critérios devem vir em pares NomeColuna, Valor.");
        }

        var list = new List<(string, object?)>();

        for (var i = 0; i < criterios.Length; i += 2)
        {
            list.Add((Convert.ToString(criterios[i], PtBr) ?? "", criterios[i + 1]));
        }

        return list;
    }

    internal static bool ValuesEqual(object? actual, object? expected)
    {
        if (actual == DBNull.Value) actual = null;
        if (expected == DBNull.Value) expected = null;

        if (actual == null && expected == null) return true;
        if (actual == null || expected == null) return false;

        return string.Equals(
            Convert.ToString(actual, CultureInfo.InvariantCulture),
            Convert.ToString(expected, CultureInfo.InvariantCulture),
            StringComparison.OrdinalIgnoreCase);
    }

    internal static string FormatValue(object? raw, string? format)
    {
        if (raw == null || raw == DBNull.Value)
        {
            return "";
        }

        if (raw is IFormattable formattable && !string.IsNullOrWhiteSpace(format))
        {
            return formattable.ToString(format, PtBr) ?? "";
        }

        if (raw is DateTime dt)
        {
            return dt.ToString(string.IsNullOrWhiteSpace(format) ? "dd/MM/yyyy" : format, PtBr);
        }

        if (raw is decimal or double or float or int or long)
        {
            return Convert.ToDecimal(raw).ToString(
                string.IsNullOrWhiteSpace(format) ? "#,##0.##" : format,
                PtBr);
        }

        return Convert.ToString(raw, PtBr) ?? "";
    }

    internal static Dictionary<string, object?> TextField(string value) =>
        new()
        {
            ["Kind"] = "string",
            ["Value"] = value ?? ""
        };

    internal static string GetColumnFormat(DataColumn column, string operation)
    {
        if (string.Equals(operation, "Count", StringComparison.OrdinalIgnoreCase))
        {
            return FormatoCount;
        }

        if (column.ExtendedProperties.Contains("Format"))
        {
            return Convert.ToString(column.ExtendedProperties["Format"], PtBr) ?? "#,##0.##";
        }

        return "#,##0.##";
    }

    internal static string RegistroParaArray(string conteudo)
    {
        conteudo = (conteudo ?? "").Trim();

        if (conteudo.Length == 0 || conteudo == "{}")
        {
            return "[]";
        }

        if (conteudo.StartsWith("[", StringComparison.Ordinal))
        {
            return conteudo;
        }

        return "[" + conteudo + "]";
    }

    internal static object WrapDatasource(string kind, string jsonArrayOrObject)
    {
        jsonArrayOrObject = (jsonArrayOrObject ?? "").Trim();

        if (jsonArrayOrObject.StartsWith("{", StringComparison.Ordinal) &&
            jsonArrayOrObject.Contains("\"Kind\"", StringComparison.OrdinalIgnoreCase))
        {
            return JsonSerializer.Deserialize<object>(jsonArrayOrObject, SerializerOptions())!;
        }

        if (jsonArrayOrObject.StartsWith("[", StringComparison.Ordinal))
        {
            var items = JsonSerializer.Deserialize<List<Dictionary<string, object?>>>(jsonArrayOrObject, SerializerOptions())
                ?? new List<Dictionary<string, object?>>();
            return new Dictionary<string, object?> { ["Kind"] = kind, ["Items"] = items };
        }

        if (jsonArrayOrObject.StartsWith("{", StringComparison.Ordinal))
        {
            var item = JsonSerializer.Deserialize<Dictionary<string, object?>>(jsonArrayOrObject, SerializerOptions())
                ?? new Dictionary<string, object?>();
            return new Dictionary<string, object?> { ["Kind"] = kind, ["Items"] = new[] { item } };
        }

        return new Dictionary<string, object?> { ["Kind"] = kind, ["Items"] = Array.Empty<object>() };
    }
}
