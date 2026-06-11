using System;
using System.Collections.Generic;
using System.Data;
using System.Globalization;
using System.Linq;
using System.Text.Json;

namespace Wordex.Json;

/// <summary>
/// Equivalente C# de WordexConsulta.bas — mesmas assinaturas públicas do crudex.xlsm.
/// A planilha Excel vira <see cref="DataSet"/> (nome da tabela = nome da aba).
/// Formato de coluna: DataColumn.ExtendedProperties["Format"] (como NumberFormat do Excel).
/// </summary>
public sealed class WordexConsulta
{
    private readonly DataSet _dataSet;

    public WordexConsulta(DataSet dataSet)
    {
        _dataSet = dataSet ?? throw new ArgumentNullException(nameof(dataSet));
    }

    /// <summary>=ObterTotal(...)</summary>
    public string ObterTotal(
        string nomeAbaOrigem,
        string colunasAgrupadoras,
        string colunasTotalizadoras,
        string valoresCalcular,
        params object?[] criterios)
    {
        try
        {
            return GerarJson(
                nomeAbaOrigem,
                colunasAgrupadoras,
                colunasTotalizadoras,
                valoresCalcular,
                gerarParaGrafico: false,
                criterios,
                retornarArray: true);
        }
        catch (Exception ex)
        {
            return WordexJsonSupport.ErrorArray(ex.Message);
        }
    }

    /// <summary>=ObterGrafico(...)</summary>
    public string ObterGrafico(
        string nomeAbaOrigem,
        string colunasAgrupadoras,
        string colunasTotalizadoras,
        string valoresCalcular,
        params object?[] criterios)
    {
        try
        {
            return GerarJson(
                nomeAbaOrigem,
                colunasAgrupadoras,
                colunasTotalizadoras,
                valoresCalcular,
                gerarParaGrafico: true,
                criterios,
                retornarArray: true);
        }
        catch (Exception ex)
        {
            return WordexJsonSupport.ErrorArray(ex.Message);
        }
    }

    /// <summary>=ObterTotais(...)</summary>
    public string ObterTotais(
        string nomeAbaOrigem,
        string colunasAgrupadoras,
        string colunasTotalizadoras,
        string valoresCalcular,
        params object?[] criterios)
    {
        try
        {
            var json = GerarJson(
                nomeAbaOrigem,
                colunasAgrupadoras,
                colunasTotalizadoras,
                valoresCalcular,
                gerarParaGrafico: false,
                criterios,
                retornarArray: false);

            return WordexJsonSupport.RegistroParaArray(json);
        }
        catch (Exception ex)
        {
            return WordexJsonSupport.ErrorArray(ex.Message);
        }
    }

    /// <summary>=ObterRegistroTotal(...) — alias legado de ObterTotais.</summary>
    public string ObterRegistroTotal(
        string nomeAbaOrigem,
        string colunasAgrupadoras,
        string colunasTotalizadoras,
        string valoresCalcular,
        params object?[] criterios) =>
        ObterTotais(nomeAbaOrigem, colunasAgrupadoras, colunasTotalizadoras, valoresCalcular, criterios);

    /// <summary>=ObterRegistroGrafico(...)</summary>
    public string ObterRegistroGrafico(
        string nomeAbaOrigem,
        string colunasAgrupadoras,
        string colunasTotalizadoras,
        string valoresCalcular,
        params object?[] criterios)
    {
        try
        {
            var json = GerarJson(
                nomeAbaOrigem,
                colunasAgrupadoras,
                colunasTotalizadoras,
                valoresCalcular,
                gerarParaGrafico: true,
                criterios,
                retornarArray: false);

            return WordexJsonSupport.RegistroParaArray(json);
        }
        catch (Exception ex)
        {
            return WordexJsonSupport.ErrorObject(ex.Message);
        }
    }

    private string GerarJson(
        string nomeAbaOrigem,
        string colunasAgrupadoras,
        string colunasTotalizadoras,
        string valoresCalcular,
        bool gerarParaGrafico,
        object?[] criterios,
        bool retornarArray)
    {
        var table = RequireTable(nomeAbaOrigem);
        var agrupadoras = WordexJsonSupport.SplitList(colunasAgrupadoras);
        var totalizadoras = WordexJsonSupport.SplitList(colunasTotalizadoras);
        var calcular = WordexJsonSupport.SplitList(valoresCalcular);
        var criteriosPares = WordexJsonSupport.ParseCriteria(criterios);

        if (totalizadoras.Length == 0)
        {
            throw new InvalidOperationException("Selecione pelo menos uma coluna para totalizar.");
        }

        if (calcular.Length == 0)
        {
            throw new InvalidOperationException("Selecione pelo menos um valor para calcular: Sum, Min, Max, Avg ou Count.");
        }

        foreach (var op in calcular)
        {
            NormalizarTipoTotal(op);
        }

        var grupos = CalcularGrupos(table, agrupadoras, totalizadoras, criteriosPares);
        var objetos = new List<Dictionary<string, object?>>();

        foreach (var grupo in grupos)
        {
            if (!GrupoAtendeCriterios(table, agrupadoras, grupo, totalizadoras, calcular, gerarParaGrafico, criteriosPares))
            {
                continue;
            }

            var obj = GrupoParaObjeto(table, agrupadoras, grupo, totalizadoras, calcular, gerarParaGrafico);
            objetos.Add(obj);

            if (!retornarArray)
            {
                return WordexJsonSupport.ToJson(obj);
            }
        }

        if (!retornarArray)
        {
            return "{}";
        }

        return WordexJsonSupport.ToJson(objetos);
    }

    private DataTable RequireTable(string name)
    {
        if (!_dataSet.Tables.Contains(name))
        {
            throw new ArgumentException($"Tabela '{name}' não existe no DataSet.");
        }

        return _dataSet.Tables[name]!;
    }

    private static List<GroupAccumulator> CalcularGrupos(
        DataTable table,
        IReadOnlyList<string> agrupadoras,
        IReadOnlyList<string> totalizadoras,
        IReadOnlyList<(string Name, object? Value)> criterios)
    {
        var map = new Dictionary<string, GroupAccumulator>(StringComparer.Ordinal);

        foreach (DataRow row in table.Rows)
        {
            if (row.RowState == DataRowState.Deleted)
            {
                continue;
            }

            if (!RegistroAtendeCriterios(row, criterios, ignorarColunaAusente: true))
            {
                continue;
            }

            var key = ObterChaveGrupo(row, agrupadoras);

            if (!map.TryGetValue(key, out var grupo))
            {
                grupo = new GroupAccumulator(ObterValoresGrupo(row, agrupadoras));
                map[key] = grupo;
            }

            AcumularLinha(row, totalizadoras, grupo.Stats);
        }

        return map.Values.ToList();
    }

    private static string ObterChaveGrupo(DataRow row, IReadOnlyList<string> agrupadoras)
    {
        if (agrupadoras.Count == 0)
        {
            return string.Empty;
        }

        return string.Join("\u001f", agrupadoras.Select(col =>
            Convert.ToString(row[col], CultureInfo.InvariantCulture) ?? ""));
    }

    private static object?[] ObterValoresGrupo(DataRow row, IReadOnlyList<string> agrupadoras) =>
        agrupadoras.Select(col => row[col] == DBNull.Value ? null : row[col]).ToArray();

    private static void AcumularLinha(DataRow row, IReadOnlyList<string> totalizadoras, GroupStats stats)
    {
        foreach (var campo in totalizadoras)
        {
            var raw = row[campo];

            if (raw == null || raw == DBNull.Value)
            {
                continue;
            }

            if (!TryToDouble(raw, out var numero))
            {
                continue;
            }

            stats.Acumular(campo, numero);
        }
    }

    private static bool TryToDouble(object raw, out double numero)
    {
        try
        {
            numero = Convert.ToDouble(raw, CultureInfo.InvariantCulture);
            return true;
        }
        catch
        {
            return double.TryParse(
                Convert.ToString(raw, WordexJsonSupport.PtBr),
                NumberStyles.Number,
                WordexJsonSupport.PtBr,
                out numero);
        }
    }

    private static bool RegistroAtendeCriterios(
        DataRow row,
        IReadOnlyList<(string Name, object? Value)> criterios,
        bool ignorarColunaAusente)
    {
        if (criterios.Count == 0)
        {
            return true;
        }

        foreach (var (name, expected) in criterios)
        {
            if (ignorarColunaAusente && !row.Table.Columns.Contains(name))
            {
                continue;
            }

            var actual = row[name];
            if (!WordexJsonSupport.ValuesEqual(actual, expected))
            {
                return false;
            }
        }

        return true;
    }

    private static bool GrupoAtendeCriterios(
        DataTable table,
        IReadOnlyList<string> agrupadoras,
        GroupAccumulator grupo,
        IReadOnlyList<string> totalizadoras,
        IReadOnlyList<string> calcular,
        bool gerarParaGrafico,
        IReadOnlyList<(string Name, object? Value)> criterios)
    {
        if (criterios.Count == 0)
        {
            return true;
        }

        foreach (var (name, expected) in criterios)
        {
            if (!TryObterValorCampoGrupo(table, agrupadoras, grupo, totalizadoras, calcular, gerarParaGrafico, name, out var actual))
            {
                continue;
            }

            if (!WordexJsonSupport.ValuesEqual(actual, expected))
            {
                return false;
            }
        }

        return true;
    }

    private static bool TryObterValorCampoGrupo(
        DataTable table,
        IReadOnlyList<string> agrupadoras,
        GroupAccumulator grupo,
        IReadOnlyList<string> totalizadoras,
        IReadOnlyList<string> calcular,
        bool gerarParaGrafico,
        string nomeCampo,
        out object? actual)
    {
        actual = null;

        for (var i = 0; i < agrupadoras.Count; i++)
        {
            if (string.Equals(agrupadoras[i], nomeCampo, StringComparison.OrdinalIgnoreCase))
            {
                actual = grupo.GroupValues[i];
                return true;
            }
        }

        foreach (var campo in totalizadoras)
        {
            foreach (var opRaw in calcular)
            {
                var op = NormalizarTipoTotal(opRaw);
                var nomeTotal = campo + op;

                if (string.Equals(nomeCampo, nomeTotal, StringComparison.OrdinalIgnoreCase))
                {
                    actual = CalcularValorTotal(grupo.Stats, campo, op);
                    return true;
                }

                if (gerarParaGrafico && string.Equals(nomeCampo, nomeTotal + "_Label", StringComparison.OrdinalIgnoreCase))
                {
                    actual = LegendaGrafico(nomeTotal);
                    return true;
                }
            }
        }

        return false;
    }

    private static Dictionary<string, object?> GrupoParaObjeto(
        DataTable table,
        IReadOnlyList<string> agrupadoras,
        GroupAccumulator grupo,
        IReadOnlyList<string> totalizadoras,
        IReadOnlyList<string> calcular,
        bool gerarParaGrafico)
    {
        var obj = new Dictionary<string, object?>(StringComparer.OrdinalIgnoreCase);

        for (var i = 0; i < agrupadoras.Count; i++)
        {
            var campo = agrupadoras[i];
            var column = table.Columns[campo];
            var format = WordexJsonSupport.GetColumnFormat(column, "text");
            var text = WordexJsonSupport.FormatValue(grupo.GroupValues[i], format);
            obj[campo] = WordexJsonSupport.TextField(text);
        }

        foreach (var campo in totalizadoras)
        {
            var column = table.Columns[campo];
            var fonteEhData = ColunaPareceData(column);

            foreach (var opRaw in calcular)
            {
                var op = NormalizarTipoTotal(opRaw);
                var nomeTotal = campo + op;
                var valorTotal = CalcularValorTotal(grupo.Stats, campo, op);
                var format = WordexJsonSupport.GetColumnFormat(column, op);
                var texto = WordexJsonSupport.FormatValue(valorTotal, format);

                obj[nomeTotal] = WordexJsonSupport.TextField(texto);

                if (gerarParaGrafico)
                {
                    obj[nomeTotal + "_Label"] = WordexJsonSupport.TextField(LegendaGrafico(nomeTotal));
                }
            }
        }

        return obj;
    }

    private static string NormalizarTipoTotal(string tipoTotal)
    {
        return (tipoTotal ?? "").Trim().ToLowerInvariant() switch
        {
            "sum" => "Sum",
            "min" => "Min",
            "max" => "Max",
            "avg" => "Avg",
            "count" => "Count",
            _ => throw new InvalidOperationException(
                $"Valor a calcular inválido: '{tipoTotal}'. Use Sum, Min, Max, Avg ou Count.")
        };
    }

    private static object? CalcularValorTotal(GroupStats stats, string campo, string tipoTotal)
    {
        return tipoTotal switch
        {
            "Count" => stats.GetCount(campo),
            "Sum" => stats.GetSum(campo),
            "Min" => stats.GetMin(campo),
            "Max" => stats.GetMax(campo),
            "Avg" => stats.GetAvg(campo),
            _ => throw new ArgumentOutOfRangeException(nameof(tipoTotal))
        };
    }

    private static string LegendaGrafico(string nomeTotal)
    {
        foreach (var sufixo in new[] { "Count", "Sum", "Min", "Max", "Avg" })
        {
            if (nomeTotal.EndsWith(sufixo, StringComparison.OrdinalIgnoreCase) &&
                nomeTotal.Length > sufixo.Length)
            {
                return nomeTotal[..^sufixo.Length];
            }
        }

        return nomeTotal;
    }

    private static bool ColunaPareceData(DataColumn column)
    {
        if (column.DataType == typeof(DateTime))
        {
            return true;
        }

        var format = Convert.ToString(column.ExtendedProperties["Format"], WordexJsonSupport.PtBr) ?? "";
        format = format.ToLowerInvariant();
        return format.Contains('d') || format.Contains('y') || format.Contains('m');
    }

    private sealed class GroupAccumulator
    {
        public GroupAccumulator(object?[] groupValues)
        {
            GroupValues = groupValues;
            Stats = new GroupStats();
        }

        public object?[] GroupValues { get; }
        public GroupStats Stats { get; }
    }

    private sealed class GroupStats
    {
        private readonly Dictionary<string, MetricStats> _metrics = new(StringComparer.OrdinalIgnoreCase);

        public void Acumular(string campo, double numero)
        {
            if (!_metrics.TryGetValue(campo, out var stats))
            {
                stats = new MetricStats();
                _metrics[campo] = stats;
            }

            stats.Add(numero);
        }

        public long GetCount(string campo) => _metrics.TryGetValue(campo, out var s) ? s.Count : 0;
        public double GetSum(string campo) => _metrics.TryGetValue(campo, out var s) ? s.Sum : 0;
        public object? GetMin(string campo) => _metrics.TryGetValue(campo, out var s) ? s.Min : null;
        public object? GetMax(string campo) => _metrics.TryGetValue(campo, out var s) ? s.Max : null;

        public object? GetAvg(string campo)
        {
            if (!_metrics.TryGetValue(campo, out var s) || s.Count == 0)
            {
                return null;
            }

            return s.Sum / s.Count;
        }
    }

    private sealed class MetricStats
    {
        public long Count { get; private set; }
        public double Sum { get; private set; }
        public double Min { get; private set; }
        public double Max { get; private set; }
        private bool _initialized;

        public void Add(double value)
        {
            if (!_initialized)
            {
                Min = value;
                Max = value;
                _initialized = true;
            }
            else
            {
                if (value < Min) Min = value;
                if (value > Max) Max = value;
            }

            Count++;
            Sum += value;
        }
    }
}
