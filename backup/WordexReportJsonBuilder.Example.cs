using System.Data;
using Wordex.Json;

namespace Wordex.Examples;

public static class CrudexJsonExample
{
    public static string BuildFromDataSet(DataSet ds)
    {
        var consulta = new WordexConsulta(ds);
        var crudex = new WordexCrudex(ds);

        // Mesmas chamadas das fórmulas Excel:
        // =ObterRegistroTotal("Produtos"; ""; "Preço, Quantidade"; "Sum")
        // =ObterRegistroGrafico("Produtos"; "ClienteId"; "Preço"; "Sum"; "ClienteId"; A2)

        return new WordexReportJsonBuilder(ds)
            .AddField("Nome", "Evadin S/A")
            .AddField("CNPJ", "12.345.678/0001-90")
            .AddCollection("Clientes", "Clientes", MapCliente)
            .AddTotal(
                "TotaisProdutosGerais",
                nomeAbaOrigem: "Produtos",
                colunasAgrupadoras: "",
                colunasTotalizadoras: "Preço, Quantidade",
                valoresCalcular: "Sum")
            .AddHistogram(
                "TotaisProdutosGrafico",
                nomeAbaOrigem: "Produtos",
                colunasAgrupadoras: "",
                colunasTotalizadoras: "Preço, Quantidade",
                valoresCalcular: "Sum")
            .BuildJson();
    }

    private static void MapCliente(WordexCollectionRowBuilder row)
    {
        row.AddField("ClienteId", "ClienteId")
            .AddField("Nome", "Nome")
            .AddField("Email", "Email")
            .AddField("Telefone", "Telefone")
            .AddTotal(
                "TotaisProdutos",
                nomeAbaOrigem: "Produtos",
                colunasAgrupadoras: "",
                colunasTotalizadoras: "Preço, Quantidade",
                valoresCalcular: "Sum")
            .AddHistogram(
                "TotaisProdutosGrafico",
                nomeAbaOrigem: "Produtos",
                colunasAgrupadoras: "",
                colunasTotalizadoras: "Preço, Quantidade",
                valoresCalcular: "Sum");
    }
}

/// <summary>Exemplo de uso direto do Chrome (sem WordexCrudex).</summary>
public static class WordexPdfExample
{
    public static void HtmlParaPdf(string htmlUrl, string caminhoPdf)
    {
        WordexChromePdf.Default.ConvertUrlToPdf(htmlUrl, caminhoPdf);
    }

    public static void ArquivoHtmlParaPdf(string caminhoHtml, string caminhoPdf)
    {
        WordexChromePdf.Default.ConvertHtmlFileToPdf(caminhoHtml, caminhoPdf);
    }
}
