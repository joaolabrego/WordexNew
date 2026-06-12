using System;
using System.Collections.Generic;
using System.Data;
using System.Text.Json;

namespace Wordex.Json;

/// <summary>
/// Monta o JSON ROOT [{ ... }] para o Wordex usando as mesmas funções do crudex.xlsm.
/// </summary>
public sealed class WordexReportJsonBuilder
{
    private readonly DataSet _dataSet;
    private readonly WordexConsulta _consulta;
    private readonly WordexCrudex _crudex;
    private readonly Dictionary<string, object?> _root = new(StringComparer.OrdinalIgnoreCase);

    public WordexReportJsonBuilder(DataSet dataSet)
    {
        _dataSet = dataSet ?? throw new ArgumentNullException(nameof(dataSet));
        _consulta = new WordexConsulta(dataSet);
        _crudex = new WordexCrudex(dataSet);
    }

    public WordexConsulta Consulta => _consulta;
    public WordexCrudex Crudex => _crudex;

    public WordexReportJsonBuilder SetField(string jsonName, object value)
    {
        _root[jsonName] = value;
        return this;
    }

    /// <summary>Campo escalar formatado (equivalente a célula tipada no ROOT).</summary>
    public WordexReportJsonBuilder AddField(string jsonName, string text)
    {
        _root[jsonName] = WordexJsonSupport.TextField(text);
        return this;
    }

    /// <summary>=ObterRegistros(...) embalado como Kind collection.</summary>
    public WordexReportJsonBuilder AddCollection(
        string jsonName,
        string nomeAba,
        params object?[] criterios)
    {
        _root[jsonName] = WordexJsonSupport.WrapDatasource("collection", _crudex.ObterRegistros(nomeAba, criterios));
        return this;
    }

    /// <summary>
    /// Coleção com totais/gráficos por linha — mesmas assinaturas VBA nos delegates.
    /// </summary>
    public WordexReportJsonBuilder AddCollection(
        string jsonName,
        string nomeAba,
        Action<WordexCollectionRowBuilder> mapRow,
        params object?[] criterios)
    {
        var table = RequireTable(nomeAba);
        var pares = WordexJsonSupport.ParseCriteria(criterios);
        var items = new List<Dictionary<string, object?>>();

        foreach (DataRow row in table.Rows)
        {
            if (row.RowState == DataRowState.Deleted)
            {
                continue;
            }

            if (pares.Count > 0 && !RowMatches(row, pares))
            {
                continue;
            }

            var rowBuilder = new WordexCollectionRowBuilder(_consulta, _crudex, row);
            mapRow(rowBuilder);
            items.Add(rowBuilder.Build());
        }

        _root[jsonName] = new Dictionary<string, object?>
        {
            ["Kind"] = "collection",
            ["Items"] = items
        };

        return this;
    }

    /// <summary>=ObterRegistroTotal(...) + Kind totals.</summary>
    public WordexReportJsonBuilder AddTotal(
        string jsonName,
        string nomeAbaOrigem,
        string colunasAgrupadoras,
        string colunasTotalizadoras,
        string valoresCalcular,
        params object?[] criterios)
    {
        _root[jsonName] = _crudex.ObterCampoTotalEmbalado(
            nomeAbaOrigem, colunasAgrupadoras, colunasTotalizadoras, valoresCalcular, criterios);
        return this;
    }

    /// <summary>=ObterRegistroGrafico(...) + Kind graph.</summary>
    public WordexReportJsonBuilder AddHistogram(
        string jsonName,
        string nomeAbaOrigem,
        string colunasAgrupadoras,
        string colunasTotalizadoras,
        string valoresCalcular,
        params object?[] criterios)
    {
        _root[jsonName] = _crudex.ObterCampoGraficoEmbalado(
            nomeAbaOrigem, colunasAgrupadoras, colunasTotalizadoras, valoresCalcular, criterios);
        return this;
    }

    public string BuildJson(bool indented = false) =>
        WordexJsonSupport.ToJson(new[] { _root }, indented);

    private DataTable RequireTable(string name)
    {
        if (!_dataSet.Tables.Contains(name))
        {
            throw new ArgumentException($"Tabela '{name}' não existe no DataSet.");
        }

        return _dataSet.Tables[name]!;
    }

    private static bool RowMatches(DataRow row, IReadOnlyList<(string Name, object? Value)> criterios)
    {
        foreach (var (name, expected) in criterios)
        {
            if (!WordexJsonSupport.ValuesEqual(row[name], expected))
            {
                return false;
            }
        }

        return true;
    }
}

public sealed class WordexCollectionRowBuilder
{
    private readonly WordexConsulta _consulta;
    private readonly WordexCrudex _crudex;
    private readonly DataRow _row;
    private readonly Dictionary<string, object?> _fields = new(StringComparer.OrdinalIgnoreCase);

    internal WordexCollectionRowBuilder(WordexConsulta consulta, WordexCrudex crudex, DataRow row)
    {
        _consulta = consulta;
        _crudex = crudex;
        _row = row;
    }

    public WordexCollectionRowBuilder AddField(string jsonName, string column, string? format = null)
    {
        var value = _row[column];
        _fields[jsonName] = WordexJsonSupport.TextField(
            WordexJsonSupport.FormatValue(value, format));
        return this;
    }

    public WordexCollectionRowBuilder AddTotal(
        string jsonName,
        string nomeAbaOrigem,
        string colunasAgrupadoras,
        string colunasTotalizadoras,
        string valoresCalcular,
        params object?[] criteriosExtra)
    {
        _fields[jsonName] = _crudex.ObterCampoTotalEmbalado(
            nomeAbaOrigem,
            colunasAgrupadoras,
            colunasTotalizadoras,
            valoresCalcular,
            MergeParentCriteria(nomeAbaOrigem, criteriosExtra));
        return this;
    }

    public WordexCollectionRowBuilder AddHistogram(
        string jsonName,
        string nomeAbaOrigem,
        string colunasAgrupadoras,
        string colunasTotalizadoras,
        string valoresCalcular,
        params object?[] criteriosExtra)
    {
        _fields[jsonName] = _crudex.ObterCampoGraficoEmbalado(
            nomeAbaOrigem,
            colunasAgrupadoras,
            colunasTotalizadoras,
            valoresCalcular,
            MergeParentCriteria(nomeAbaOrigem, criteriosExtra));
        return this;
    }

    internal Dictionary<string, object?> Build() => _fields;

    private object?[] MergeParentCriteria(string nomeAbaOrigem, object?[] criteriosExtra)
    {
        var list = new List<object?>(criteriosExtra ?? Array.Empty<object?>());

        if (!_row.Table.DataSet?.Tables.Contains(nomeAbaOrigem) ?? true)
        {
            return list.ToArray();
        }

        var childTable = _row.Table.DataSet!.Tables[nomeAbaOrigem]!;

        foreach (DataColumn column in _row.Table.Columns)
        {
            if (!childTable.Columns.Contains(column.ColumnName))
            {
                continue;
            }

            list.Add(column.ColumnName);
            list.Add(_row[column] == DBNull.Value ? null : _row[column]);
        }

        return list.ToArray();
    }
}
