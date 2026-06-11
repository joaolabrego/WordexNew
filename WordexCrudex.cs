using System;
using System.Collections.Generic;
using System.Data;
using System.Globalization;
using System.Linq;
using System.Text.Json;
using System.Threading;

namespace Wordex.Json;

/// <summary>
/// Equivalente C# de Wordex.bas — ObterRegistro / ObterRegistros.
/// </summary>
public sealed class WordexCrudex
{
    private readonly DataSet _dataSet;
    private readonly WordexConsulta _consulta;
    private readonly WordexChromePdf _pdf;

    public WordexCrudex(DataSet dataSet, WordexChromePdf? pdf = null)
    {
        _dataSet = dataSet ?? throw new ArgumentNullException(nameof(dataSet));
        _consulta = new WordexConsulta(dataSet);
        _pdf = pdf ?? WordexChromePdf.Default;
    }

    public WordexConsulta Consulta => _consulta;
    public WordexChromePdf Pdf => _pdf;

    /// <summary>=ObterRegistro(nomeAba, criterios...)</summary>
    public string ObterRegistro(string nomeAba, params object?[] criterios)
    {
        try
        {
            var table = RequireTable(nomeAba);
            var pares = WordexJsonSupport.ParseCriteria(criterios);
            var row = FindFirstRow(table, pares);

            return row == null
                ? ObterRegistroJson(table, null, vazio: true)
                : ObterRegistroJson(table, row, vazio: false);
        }
        catch (Exception ex)
        {
            return WordexJsonSupport.ErrorObject(ex.Message);
        }
    }

    /// <summary>=ObterRegistros(nomeAba, criterios...)</summary>
    public string ObterRegistros(string nomeAba, params object?[] criterios)
    {
        try
        {
            var table = RequireTable(nomeAba);
            var pares = WordexJsonSupport.ParseCriteria(criterios);
            var rows = table.Rows.Cast<DataRow>()
                .Where(r => r.RowState != DataRowState.Deleted && RegistroAtendeCriterios(r, pares))
                .Select(r => ObterRegistroJson(table, r, vazio: false))
                .ToList();

            var parsed = rows
                .Select(json => JsonSerializer.Deserialize<Dictionary<string, object?>>(json, WordexJsonSupport.SerializerOptions())!)
                .ToList();

            return WordexJsonSupport.ToJson(parsed);
        }
        catch (Exception ex)
        {
            return WordexJsonSupport.ErrorArray(ex.Message);
        }
    }

    private DataTable RequireTable(string name)
    {
        if (!_dataSet.Tables.Contains(name))
        {
            throw new ArgumentException($"Tabela '{name}' não existe no DataSet.");
        }

        return _dataSet.Tables[name]!;
    }

    private static DataRow? FindFirstRow(DataTable table, IReadOnlyList<(string Name, object? Value)> criterios)
    {
        return table.Rows.Cast<DataRow>()
            .FirstOrDefault(r => r.RowState != DataRowState.Deleted && RegistroAtendeCriterios(r, criterios));
    }

    private static bool RegistroAtendeCriterios(DataRow row, IReadOnlyList<(string Name, object? Value)> criterios)
    {
        if (criterios.Count == 0)
        {
            return true;
        }

        foreach (var (name, expected) in criterios)
        {
            if (!row.Table.Columns.Contains(name))
            {
                throw new ArgumentException($"Coluna '{name}' não existe em '{row.Table.TableName}'.");
            }

            if (!WordexJsonSupport.ValuesEqual(row[name], expected))
            {
                return false;
            }
        }

        return true;
    }

    private string ObterRegistroJson(DataTable table, DataRow? row, bool vazio)
    {
        var obj = new Dictionary<string, object?>(StringComparer.OrdinalIgnoreCase);

        foreach (DataColumn column in table.Columns)
        {
            var titulo = column.ColumnName;

            if (titulo.EndsWith("_Kind", StringComparison.OrdinalIgnoreCase))
            {
                continue;
            }

            if (vazio)
            {
                obj[titulo] = WordexJsonSupport.TextField("");
                continue;
            }

            var raw = row![column];

            if (TryGetJsonRaw(raw, out var jsonRaw))
            {
                obj[titulo] = EmbalarJsonColuna(titulo, jsonRaw, ObterKindExplicito(row, titulo));
                continue;
            }

            obj[titulo] = JsonCampoTipado(column, raw, vazio: false);
        }

        return WordexJsonSupport.ToJson(obj);
    }

    private static bool TryGetJsonRaw(object? raw, out string json)
    {
        json = Convert.ToString(raw, CultureInfo.InvariantCulture) ?? "";
        json = json.Trim();

        if (json.StartsWith("{", StringComparison.Ordinal) || json.StartsWith("[", StringComparison.Ordinal))
        {
            return true;
        }

        json = "";
        return false;
    }

    private object EmbalarJsonColuna(string tituloColuna, string valorTexto, string? kindExplicito)
    {
        valorTexto = valorTexto.Trim();

        if (JaEmbaladoComoDatasource(valorTexto))
        {
            return JsonSerializer.Deserialize<object>(valorTexto, WordexJsonSupport.SerializerOptions())!;
        }

        var kind = string.IsNullOrWhiteSpace(kindExplicito)
            ? InferirKindDatasource(tituloColuna)
            : kindExplicito;

        if (valorTexto.StartsWith("[", StringComparison.Ordinal))
        {
            var items = JsonSerializer.Deserialize<object>(valorTexto, WordexJsonSupport.SerializerOptions());
            return new Dictionary<string, object?> { ["Kind"] = kind, ["Items"] = items! };
        }

        if (valorTexto.StartsWith("{", StringComparison.Ordinal))
        {
            var item = JsonSerializer.Deserialize<object>(valorTexto, WordexJsonSupport.SerializerOptions());
            return new Dictionary<string, object?> { ["Kind"] = kind, ["Items"] = new[] { item! } };
        }

        return valorTexto;
    }

    private static bool JaEmbaladoComoDatasource(string conteudo)
    {
        var lower = conteudo.ToLowerInvariant();
        return conteudo.StartsWith("{", StringComparison.Ordinal) &&
               lower.Contains("\"kind\"") &&
               lower.Contains("\"items\"");
    }

    private static string InferirKindDatasource(string tituloColuna)
    {
        var nome = tituloColuna.Trim().ToLowerInvariant();

        if (nome.EndsWith("grafico", StringComparison.Ordinal))
        {
            return "histogram";
        }

        if (nome.StartsWith("totais", StringComparison.Ordinal))
        {
            return "total";
        }

        return "collection";
    }

    private static string? ObterKindExplicito(DataRow row, string tituloColuna)
    {
        var kindColumn = tituloColuna + "_Kind";

        if (!row.Table.Columns.Contains(kindColumn))
        {
            return null;
        }

        return Convert.ToString(row[kindColumn], CultureInfo.InvariantCulture)?.Trim();
    }

    private static Dictionary<string, object?> JsonCampoTipado(DataColumn column, object raw, bool vazio)
    {
        if (vazio || raw == null || raw == DBNull.Value)
        {
            return new Dictionary<string, object?> { ["Kind"] = "string", ["Value"] = null };
        }

        var kind = InferirKind(column, raw);

        return kind switch
        {
            "number" => new Dictionary<string, object?>
            {
                ["Kind"] = "number",
                ["Value"] = Convert.ToDecimal(raw, CultureInfo.InvariantCulture)
            },
            "boolean" => new Dictionary<string, object?>
            {
                ["Kind"] = "boolean",
                ["Value"] = Convert.ToBoolean(raw, CultureInfo.InvariantCulture)
            },
            "datetime" => new Dictionary<string, object?>
            {
                ["Kind"] = "datetime",
                ["Value"] = Convert.ToDateTime(raw, CultureInfo.InvariantCulture).ToString("yyyy-MM-dd")
            },
            _ => WordexJsonSupport.TextField(WordexJsonSupport.FormatValue(raw, GetFormat(column)))
        };
    }

    private static string InferirKind(DataColumn column, object raw)
    {
        if (column.ExtendedProperties.Contains("Kind"))
        {
            var explicitKind = Convert.ToString(column.ExtendedProperties["Kind"], CultureInfo.InvariantCulture)?.ToLowerInvariant()
                ?? "string";
            return explicitKind == "text" ? "string" : explicitKind;
        }

        if (raw is bool)
        {
            return "boolean";
        }

        if (raw is DateTime)
        {
            return "datetime";
        }

        if (raw is byte or short or int or long or float or double or decimal)
        {
            return "number";
        }

        return "string";
    }

    private static string? GetFormat(DataColumn column) =>
        column.ExtendedProperties.Contains("Format")
            ? Convert.ToString(column.ExtendedProperties["Format"], WordexJsonSupport.PtBr)
            : null;

    /// <summary>
    /// Atalho: monta campo total/histogram embalado como no ROOT do Excel
    /// (=ObterRegistroTotal / =ObterRegistroGrafico + Kind wrapper).
    /// </summary>
    public object ObterCampoTotalEmbalado(
        string nomeAbaOrigem,
        string colunasAgrupadoras,
        string colunasTotalizadoras,
        string valoresCalcular,
        params object?[] criterios) =>
        WordexJsonSupport.WrapDatasource("total",
            _consulta.ObterRegistroTotal(nomeAbaOrigem, colunasAgrupadoras, colunasTotalizadoras, valoresCalcular, criterios));

    public object ObterCampoGraficoEmbalado(
        string nomeAbaOrigem,
        string colunasAgrupadoras,
        string colunasTotalizadoras,
        string valoresCalcular,
        params object?[] criterios) =>
        WordexJsonSupport.WrapDatasource("histogram",
            _consulta.ObterRegistroGrafico(nomeAbaOrigem, colunasAgrupadoras, colunasTotalizadoras, valoresCalcular, criterios));

    /// <summary>
    /// HTML paginado Wordex (saída de "Montar HTML" / .pagex-page) servido por URL → PDF.
    /// Não use WORDEX.html (editor/template) nem wordexDocument.
    /// </summary>
    public void GerarPdf(string htmlUrl, string caminhoPdf, CancellationToken cancellationToken = default) =>
        _pdf.ConvertUrlToPdf(htmlUrl, caminhoPdf, cancellationToken);

    /// <summary>HTML paginado (URL http) → bytes do PDF.</summary>
    public byte[] GerarPdfBytes(string htmlUrl, CancellationToken cancellationToken = default) =>
        _pdf.ConvertUrlToPdfBytes(htmlUrl, cancellationToken);

    /// <summary>HTML paginado (URL http) → data URI base64 (formato ObterPDF).</summary>
    public string GerarPdfDataUri(string htmlUrl, CancellationToken cancellationToken = default) =>
        _pdf.ConvertUrlToPdfDataUri(htmlUrl, cancellationToken);

    /// <summary>
    /// Arquivo .html da janela paginada (Salvar HTML) → PDF.
    /// Rejeita WORDEX.html e documento não paginado (RequirePaginatedHtml).
    /// </summary>
    public void GerarPdfDeArquivoHtml(string caminhoHtml, string caminhoPdf, CancellationToken cancellationToken = default) =>
        _pdf.ConvertHtmlFileToPdf(caminhoHtml, caminhoPdf, cancellationToken);

    /// <summary>
    /// Conteúdo HTML paginado em memória → PDF.
    /// Deve conter .pagex-page (equivalente ao HTML salvo da janela paginada).
    /// </summary>
    public void GerarPdfDeHtml(string html, string caminhoPdf, CancellationToken cancellationToken = default) =>
        _pdf.ConvertHtmlContentToPdf(html, caminhoPdf, cancellationToken);
}
