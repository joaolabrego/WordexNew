using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Net.Sockets;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace Wordex.Json;

/// <summary>
/// Converte o <strong>HTML paginado</strong> Wordex (saída de "Montar HTML" / wordex-paged)
/// em PDF via Chrome/Edge headless (--print-to-pdf).
/// Não use WORDEX.html (template/editor) nem documento não paginado.
/// Ver helpCHROME.txt na raiz do projeto.
/// </summary>
public sealed class WordexChromePdf
{
    private const string PaginatedHtmlRequiredMessage =
        "O PDF deve ser gerado a partir do HTML paginado Wordex (com .pagex-page), " +
        "produzido por \"Montar template HTML\" e paginação (wordex-paged.html). " +
        "Não use WORDEX.html (editor/template) nem o documento interno wordexDocument.";

    private static readonly HttpClient SharedHttpClient = new()
    {
        Timeout = TimeSpan.FromSeconds(120)
    };
    private static readonly string[] ChromeCandidates =
    {
        @"C:\Program Files\Google\Chrome\Application\chrome.exe",
        @"C:\Program Files (x86)\Google\Chrome\Application\chrome.exe",
        @"C:\Program Files\Microsoft\Edge\Application\msedge.exe",
        @"C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe"
    };

    public WordexChromePdfOptions Options { get; }

    public WordexChromePdf(WordexChromePdfOptions? options = null)
    {
        Options = options ?? new WordexChromePdfOptions();
    }

    public static WordexChromePdf Default { get; } = new();

    /// <summary>Localiza chrome.exe ou msedge.exe (registro Windows + caminhos padrão).</summary>
    public static string? FindBrowserExecutable()
    {
        try
        {
            var fromRegistry = Microsoft.Win32.Registry.GetValue(
                @"HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\chrome.exe",
                "",
                null) as string;

            if (!string.IsNullOrWhiteSpace(fromRegistry) && File.Exists(fromRegistry))
            {
                return fromRegistry;
            }
        }
        catch
        {
            /* tenta caminhos fixos */
        }

        return ChromeCandidates.FirstOrDefault(File.Exists);
    }

    /// <summary>
    /// Verifica se o HTML é saída paginada Wordex (contém páginas A4 .pagex-page).
    /// </summary>
    public static bool IsPaginatedWordexHtml(string? html) =>
        !string.IsNullOrWhiteSpace(html) &&
        html.Contains("pagex-page", StringComparison.OrdinalIgnoreCase);

    /// <summary>
    /// Garante HTML paginado; lança se for template/editor ou documento sem paginação.
    /// </summary>
    public static void EnsurePaginatedWordexHtml(string? html)
    {
        if (IsPaginatedWordexHtml(html))
        {
            return;
        }

        var hint = LooksLikeWordexTemplate(html)
            ? " O arquivo parece ser o template/editor (WORDEX.html), não o relatório paginado."
            : "";

        throw new InvalidOperationException(PaginatedHtmlRequiredMessage + hint);
    }

    private static bool LooksLikeWordexTemplate(string? html)
    {
        if (string.IsNullOrWhiteSpace(html))
        {
            return false;
        }

        return html.Contains("wordexDocument", StringComparison.OrdinalIgnoreCase) ||
               html.Contains("wordex-editor", StringComparison.OrdinalIgnoreCase) ||
               html.Contains("class=\"body-flow\"", StringComparison.OrdinalIgnoreCase) ||
               html.Contains("WordexEditor", StringComparison.OrdinalIgnoreCase);
    }

    /// <summary>URL http(s) do HTML paginado servido → PDF em disco.</summary>
    public void ConvertPaginatedUrlToPdf(string paginatedHtmlUrl, string pdfPath, CancellationToken cancellationToken = default) =>
        ConvertUrlToPdf(paginatedHtmlUrl, pdfPath, cancellationToken);

    /// <summary>URL http(s) do HTML paginado servido → PDF em disco.</summary>
    public void ConvertUrlToPdf(string htmlUrl, string pdfPath, CancellationToken cancellationToken = default)
    {
        var bytes = ConvertUrlToPdfBytes(htmlUrl, cancellationToken);
        Directory.CreateDirectory(Path.GetDirectoryName(Path.GetFullPath(pdfPath))!);
        File.WriteAllBytes(pdfPath, bytes);
    }

    /// <summary>URL http(s) do HTML paginado servido → bytes do PDF.</summary>
    public byte[] ConvertPaginatedUrlToPdfBytes(string paginatedHtmlUrl, CancellationToken cancellationToken = default) =>
        ConvertUrlToPdfBytes(paginatedHtmlUrl, cancellationToken);

    /// <summary>URL http(s) do HTML paginado servido → bytes do PDF.</summary>
    public byte[] ConvertUrlToPdfBytes(string htmlUrl, CancellationToken cancellationToken = default)
    {
        ValidateHtmlUrl(htmlUrl);

        if (Options.RequirePaginatedHtml)
        {
            EnsurePaginatedHtmlFromUrl(htmlUrl, cancellationToken);
        }

        var browser = ResolveBrowserPath();
        var tempPdf = Path.Combine(
            Path.GetTempPath(),
            $"wordex-{Guid.NewGuid():N}.pdf");

        try
        {
            RunChromePrint(browser, htmlUrl, tempPdf, cancellationToken);

            if (!File.Exists(tempPdf))
            {
                throw new InvalidOperationException("Chrome não gerou o arquivo PDF.");
            }

            return File.ReadAllBytes(tempPdf);
        }
        finally
        {
            TryDelete(tempPdf);
        }
    }

    /// <summary>data:application/pdf;base64,... (mesmo formato do ObterPDF).</summary>
    public string ConvertUrlToPdfDataUri(string htmlUrl, CancellationToken cancellationToken = default)
    {
        var bytes = ConvertUrlToPdfBytes(htmlUrl, cancellationToken);
        return "data:application/pdf;base64," + Convert.ToBase64String(bytes);
    }

    /// <summary>
    /// Arquivo .html da janela paginada (Salvar HTML) → PDF.
    /// Sobe HTTP temporário em localhost — preferível a file://.
    /// </summary>
    public byte[] ConvertPaginatedHtmlFileToPdfBytes(string paginatedHtmlFilePath, CancellationToken cancellationToken = default) =>
        ConvertHtmlFileToPdfBytes(paginatedHtmlFilePath, cancellationToken);

    public byte[] ConvertHtmlFileToPdfBytes(string htmlFilePath, CancellationToken cancellationToken = default)
    {
        htmlFilePath = Path.GetFullPath(htmlFilePath);

        if (!File.Exists(htmlFilePath))
        {
            throw new FileNotFoundException("HTML não encontrado.", htmlFilePath);
        }

        if (Options.RequirePaginatedHtml)
        {
            EnsurePaginatedWordexHtml(File.ReadAllText(htmlFilePath, Encoding.UTF8));
        }

        using var server = WordexTempHtmlServer.Start(htmlFilePath, Options.AllowedHosts);

        return ConvertUrlToPdfBytes(server.BaseUrl, cancellationToken);
    }

    public void ConvertHtmlFileToPdf(string htmlFilePath, string pdfPath, CancellationToken cancellationToken = default)
    {
        var bytes = ConvertHtmlFileToPdfBytes(htmlFilePath, cancellationToken);
        Directory.CreateDirectory(Path.GetDirectoryName(Path.GetFullPath(pdfPath))!);
        File.WriteAllBytes(pdfPath, bytes);
    }

    /// <summary>Conteúdo HTML paginado em memória → PDF (servidor HTTP temporário).</summary>
    public byte[] ConvertPaginatedHtmlContentToPdfBytes(string paginatedHtmlContent, CancellationToken cancellationToken = default) =>
        ConvertHtmlContentToPdfBytes(paginatedHtmlContent, cancellationToken);

    public byte[] ConvertHtmlContentToPdfBytes(string htmlContent, CancellationToken cancellationToken = default)
    {
        if (Options.RequirePaginatedHtml)
        {
            EnsurePaginatedWordexHtml(htmlContent);
        }

        var tempHtml = Path.Combine(Path.GetTempPath(), $"wordex-{Guid.NewGuid():N}.html");

        try
        {
            File.WriteAllText(tempHtml, htmlContent ?? "", new UTF8Encoding(encoderShouldEmitUTF8Identifier: false));
            return ConvertHtmlFileToPdfBytes(tempHtml, cancellationToken);
        }
        finally
        {
            TryDelete(tempHtml);
        }
    }

    public void ConvertHtmlContentToPdf(string htmlContent, string pdfPath, CancellationToken cancellationToken = default)
    {
        var bytes = ConvertHtmlContentToPdfBytes(htmlContent, cancellationToken);
        Directory.CreateDirectory(Path.GetDirectoryName(Path.GetFullPath(pdfPath))!);
        File.WriteAllBytes(pdfPath, bytes);
    }

    private string ResolveBrowserPath()
    {
        var path = !string.IsNullOrWhiteSpace(Options.BrowserExecutablePath)
            ? Options.BrowserExecutablePath
            : FindBrowserExecutable();

        if (string.IsNullOrWhiteSpace(path) || !File.Exists(path))
        {
            throw new FileNotFoundException(
                "Chrome/Edge não encontrado. Instale o Google Chrome ou defina BrowserExecutablePath.");
        }

        return path;
    }

    private void EnsurePaginatedHtmlFromUrl(string htmlUrl, CancellationToken cancellationToken)
    {
        cancellationToken.ThrowIfCancellationRequested();

        using var request = new HttpRequestMessage(HttpMethod.Get, htmlUrl);
        using var response = SharedHttpClient.Send(request, cancellationToken);

        if (!response.IsSuccessStatusCode)
        {
            throw new InvalidOperationException(
                $"Não foi possível validar o HTML paginado em {htmlUrl}: HTTP {(int)response.StatusCode}.");
        }

        var html = response.Content.ReadAsStringAsync(cancellationToken).GetAwaiter().GetResult();
        EnsurePaginatedWordexHtml(html);
    }

    private void RunChromePrint(
        string browserPath,
        string htmlUrl,
        string outputPdfPath,
        CancellationToken cancellationToken)
    {
        var args = BuildChromeArguments(htmlUrl, outputPdfPath);
        var psi = new ProcessStartInfo
        {
            FileName = browserPath,
            Arguments = args,
            UseShellExecute = false,
            CreateNoWindow = true,
            RedirectStandardOutput = true,
            RedirectStandardError = true
        };

        using var process = Process.Start(psi)
            ?? throw new InvalidOperationException("Não foi possível iniciar o Chrome headless.");

        var stdout = process.StandardOutput.ReadToEndAsync();
        var stderr = process.StandardError.ReadToEndAsync();

        if (!process.WaitForExit(Options.TimeoutMs))
        {
            try { process.Kill(entireProcessTree: true); } catch { /* ignore */ }
            throw new TimeoutException(
                $"Chrome excedeu o timeout de {Options.TimeoutMs} ms ao gerar o PDF.");
        }

        cancellationToken.ThrowIfCancellationRequested();

        if (process.ExitCode != 0)
        {
            var err = stderr.GetAwaiter().GetResult();
            throw new InvalidOperationException(
                $"Chrome retornou código {process.ExitCode}. {err}".Trim());
        }

        _ = stdout.GetAwaiter().GetResult();
    }

    private string BuildChromeArguments(string htmlUrl, string outputPdfPath)
    {
        var parts = new List<string>
        {
            "--headless=new",
            "--disable-gpu",
            "--run-all-compositor-stages-before-draw",
            $"--virtual-time-budget={Options.VirtualTimeBudgetMs}",
            $"--print-to-pdf={Quote(outputPdfPath)}",
            Quote(htmlUrl)
        };

        if (Options.NoPdfHeaderFooter)
        {
            parts.Insert(parts.Count - 1, "--print-to-pdf-no-header");
        }

        if (Options.DisableSandbox)
        {
            parts.Insert(1, "--no-sandbox");
        }

        return string.Join(" ", parts);
    }

    private void ValidateHtmlUrl(string htmlUrl)
    {
        if (string.IsNullOrWhiteSpace(htmlUrl))
        {
            throw new ArgumentException("htmlUrl é obrigatório.", nameof(htmlUrl));
        }

        if (!Uri.TryCreate(htmlUrl, UriKind.Absolute, out var uri))
        {
            throw new ArgumentException("htmlUrl deve ser uma URL absoluta (http/https).", nameof(htmlUrl));
        }

        if (uri.Scheme != Uri.UriSchemeHttp && uri.Scheme != Uri.UriSchemeHttps)
        {
            throw new ArgumentException("Use http:// ou https:// (prefira servir o HTML localmente).", nameof(htmlUrl));
        }

        if (Options.AllowedHosts.Count > 0)
        {
            var host = uri.Host.ToLowerInvariant();

            if (!Options.AllowedHosts.Any(h => string.Equals(h, host, StringComparison.OrdinalIgnoreCase)))
            {
                throw new InvalidOperationException(
                    $"Host não permitido para PDF: {uri.Host}. Ajuste AllowedHosts.");
            }
        }
    }

    private static string Quote(string value) =>
        value.Contains(' ') || value.Contains('"')
            ? "\"" + value.Replace("\"", "\\\"") + "\""
            : value;

    private static void TryDelete(string path)
    {
        try
        {
            if (File.Exists(path))
            {
                File.Delete(path);
            }
        }
        catch
        {
            /* temp cleanup best-effort */
        }
    }
}

public sealed class WordexChromePdfOptions
{
    /// <summary>Se vazio, detecta Chrome/Edge automaticamente.</summary>
    public string? BrowserExecutablePath { get; set; }

    public int VirtualTimeBudgetMs { get; set; } = 15000;

    public int TimeoutMs { get; set; } = 120000;

    public bool NoPdfHeaderFooter { get; set; } = true;

    public bool DisableSandbox { get; set; }

    /// <summary>
    /// Exige HTML paginado Wordex (.pagex-page) antes de converter.
    /// Rejeita WORDEX.html (editor/template) e documento não paginado.
    /// </summary>
    public bool RequirePaginatedHtml { get; set; } = true;

    /// <summary>
    /// Hosts permitidos em htmlUrl (proteção SSRF). Vazio = qualquer host.
    /// Para HTML temporário local, use localhost / 127.0.0.1.
    /// </summary>
    public IList<string> AllowedHosts { get; } = new List<string>();
}

internal sealed class WordexTempHtmlServer : IDisposable
{
    private readonly HttpListener _listener;
    private readonly string _filePath;
    private readonly Thread _thread;
    private readonly ManualResetEventSlim _ready = new(false);
    private volatile bool _disposed;

    public string BaseUrl { get; }

    private WordexTempHtmlServer(HttpListener listener, string filePath, string baseUrl)
    {
        _listener = listener;
        _filePath = filePath;
        BaseUrl = baseUrl;
        _thread = new Thread(ListenLoop)
        {
            IsBackground = true,
            Name = "WordexTempHtmlServer"
        };
        _thread.Start();
        _ready.Wait(TimeSpan.FromSeconds(5));
    }

    public static WordexTempHtmlServer Start(string htmlFilePath, IList<string> allowedHosts)
    {
        htmlFilePath = Path.GetFullPath(htmlFilePath);
        var port = GetFreePort();
        var prefix = $"http://127.0.0.1:{port}/";
        var listener = new HttpListener();
        listener.Prefixes.Add(prefix);

        if (allowedHosts.Count == 0)
        {
            allowedHosts.Add("127.0.0.1");
            allowedHosts.Add("localhost");
        }

        listener.Start();
        return new WordexTempHtmlServer(listener, htmlFilePath, prefix);
    }

    private void ListenLoop()
    {
        _ready.Set();

        while (!_disposed)
        {
            HttpListenerContext? ctx = null;

            try
            {
                ctx = _listener.GetContext();
                var path = ctx.Request.Url?.AbsolutePath.TrimStart('/') ?? "";

                if (path.Length == 0 ||
                    string.Equals(path, "index.html", StringComparison.OrdinalIgnoreCase) ||
                    string.Equals(Uri.UnescapeDataString(path), Path.GetFileName(_filePath), StringComparison.OrdinalIgnoreCase))
                {
                    var bytes = File.ReadAllBytes(_filePath);
                    ctx.Response.ContentType = "text/html; charset=utf-8";
                    ctx.Response.ContentLength64 = bytes.Length;
                    ctx.Response.OutputStream.Write(bytes, 0, bytes.Length);
                }
                else
                {
                    ctx.Response.StatusCode = 404;
                }
            }
            catch (HttpListenerException)
            {
                break;
            }
            catch (ObjectDisposedException)
            {
                break;
            }
            finally
            {
                try { ctx?.Response.OutputStream.Close(); } catch { /* ignore */ }
                try { ctx?.Response.Close(); } catch { /* ignore */ }
            }
        }
    }

    private static int GetFreePort()
    {
        var listener = new TcpListener(System.Net.IPAddress.Loopback, 0);
        listener.Start();
        var port = ((System.Net.IPEndPoint)listener.LocalEndpoint).Port;
        listener.Stop();
        return port;
    }

    public void Dispose()
    {
        _disposed = true;

        try { _listener.Stop(); } catch { /* ignore */ }
        try { _listener.Close(); } catch { /* ignore */ }
    }
}
