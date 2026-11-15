// The viewer service: <comofficekit-viewer> sends it a document, and gets back each page as a PNG image and the whole
// document as PDF, which browsers show themselves. Its contract is api/openapi.yaml.
//
//   comofficekit-viewer-service --urls http://localhost:5000 [--AllowedOrigins https://app.example.com] [--WorkingDirectory <folder>]
//
// The licence comes from COMOFFICEKIT_LICENSE_KEY (the key) or COMOFFICEKIT_LICENSE_FILE (a licence file).
//
//   POST   /documents?name=<file name>      the document as the request body; answers its id, pages, kind and fonts
//   GET    /documents/{id}                  the same, for a document already open
//   GET    /documents/{id}/pages/{n}?dpi=96 page n, counting from 1, as PNG
//   GET    /documents/{id}/pages/{n}/text   page n's text and where it is drawn, to lay over the image
//   GET    /documents/{id}/text             every page's text, for searching the document
//   GET    /documents/{id}/pdf              the document as PDF
//   DELETE /documents/{id}                  closes it and deletes the service's copy
//   GET    /comofficekit-viewer.js          the web component
//
// Settings, from the command line, the environment or appsettings.json:
//   AllowedOrigins          other origins whose pages may call the service, comma-separated; none by default
//   AllowCredentials        true to let those pages send cookies
//   WorkingDirectory        where documents are kept while open; the temporary folder by default
//   MaxOpenDocuments        16 by default; opening one more closes the one open longest
//   TimeLimitSeconds        the longest one request may keep the engine, 60 by default
//   Kestrel:Limits:MaxRequestBodySize   the largest document accepted, 30,000,000 bytes by default
//
// The engine lays documents out one at a time within a process, so the service takes them to it one at a time, each
// call under a time limit and stopped when the browser goes away. A service with many users runs several processes
// and sends each document's requests to the process that opened it. There is no sign-in: put the service behind your
// own before it faces the internet.
using System.Globalization;
using ComOfficeKit;
using Microsoft.AspNetCore.Http.Features;

if (string.IsNullOrWhiteSpace(Environment.GetEnvironmentVariable("COMOFFICEKIT_LICENSE_KEY"))
    && string.IsNullOrWhiteSpace(Environment.GetEnvironmentVariable("COMOFFICEKIT_LICENSE_FILE")))
{
    Console.Error.WriteLine("Set COMOFFICEKIT_LICENSE_KEY to your licence key, or COMOFFICEKIT_LICENSE_FILE to a licence file.");
    return 2;
}

var builder = WebApplication.CreateBuilder(new WebApplicationOptions
{
    Args = args,
    // Published, the page is beside the program; under dotnet run, in the project folder, the current one.
    ContentRootPath = Directory.Exists(Path.Combine(AppContext.BaseDirectory, "wwwroot")) ? AppContext.BaseDirectory : null,
});
var origins = (builder.Configuration["AllowedOrigins"] ?? "")
    .Split(',', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries);
if (origins.Length > 0)
{
    builder.Services.AddCors(cors => cors.AddDefaultPolicy(policy =>
    {
        policy.WithOrigins(origins).WithMethods("GET", "POST", "DELETE").WithHeaders("Content-Type");
        if (builder.Configuration.GetValue<bool>("AllowCredentials"))
        {
            policy.AllowCredentials();
        }
    }));
}
var app = builder.Build();

// A folder of the service's own, inside WorkingDirectory or the temporary folder, deleted when it stops: the engine's
// working directory, with a quota, and the copies of the documents it has open.
var parent = app.Configuration["WorkingDirectory"] is { Length: > 0 } configured ? configured : Path.GetTempPath();
var workRoot = Directory.CreateDirectory(Path.Combine(parent, $"comofficekit-viewer-{Guid.NewGuid():N}")).FullName;
var copies = Directory.CreateDirectory(Path.Combine(workRoot, "documents")).FullName;
Engine.Configure(new EngineOptions { WorkingDirectory = Path.Combine(workRoot, "engine"), WorkingDirectoryQuota = 2L << 30 });
License.Configure(LicenseStorage.Memory);
if (Environment.GetEnvironmentVariable("COMOFFICEKIT_LICENSE_FILE") is { Length: > 0 } file)
{
    License.ActivateFromFile(file);
}
else
{
    License.Activate(Environment.GetEnvironmentVariable("COMOFFICEKIT_LICENSE_KEY")!);
}

var documents = new OpenDocuments(
    capacity: app.Configuration.GetValue("MaxOpenDocuments", 16),
    timeLimit: TimeSpan.FromSeconds(app.Configuration.GetValue("TimeLimitSeconds", 60)));

if (origins.Length > 0)
{
    app.UseCors();
}
app.UseDefaultFiles();
app.UseStaticFiles();

var component = Path.Combine(AppContext.BaseDirectory, "comofficekit-viewer.js");
app.MapGet("/comofficekit-viewer.js", () => Results.File(component, "text/javascript; charset=utf-8"));

app.MapPost("/documents", async (HttpRequest request, string? name, CancellationToken aborted) =>
{
    var extension = Path.GetExtension(name ?? "").ToLowerInvariant();
    if (extension is not (".docx" or ".xlsx" or ".xlsm" or ".pptx" or ".csv" or ".tsv"))
    {
        return Results.Problem("Send a .docx, .xlsx, .xlsm, .pptx, .csv or .tsv file, with its name in ?name=.", statusCode: 400);
    }
    // The service names its copy itself: the name the browser sends gives the format, and nothing else.
    var id = Guid.NewGuid().ToString("N");
    var copy = Path.Combine(copies, id + extension);
    try
    {
        await using (var stream = File.Create(copy))
        {
            await request.Body.CopyToAsync(stream, aborted);
        }
    }
    catch (BadHttpRequestException e) when (e.StatusCode == StatusCodes.Status413PayloadTooLarge)
    {
        // The server stops reading a body past its limit. Answered here, the page can say why, and the service's log
        // keeps no error for what is the sender's mistake.
        File.Delete(copy);
        var most = request.HttpContext.Features.Get<IHttpMaxRequestBodySizeFeature>()?.MaxRequestBodySize;
        return Results.Problem(most is { } bytes
            ? $"The document is larger than this service accepts: at most {bytes.ToString("N0", CultureInfo.InvariantCulture)} bytes."
            : "The document is larger than this service accepts.", statusCode: 413);
    }
    catch
    {
        File.Delete(copy);
        throw;
    }
    return await documents.Run(aborted, limit => documents.Open(id, copy, limit));
});

app.MapGet("/documents/{id}", (string id, CancellationToken aborted) =>
    documents.Run(aborted, _ => documents.Find(id) is { } entry ? Results.Ok(entry.Info) : Results.NotFound()));

app.MapGet("/documents/{id}/pages/{page:int}", (string id, int page, double? dpi, CancellationToken aborted) =>
    documents.Run(aborted, limit =>
    {
        if (documents.Find(id) is not { } entry || page < 1 || page > entry.Info.Pages)
        {
            return Results.NotFound();
        }
        var image = entry.Document.RenderPage(page - 1, Math.Clamp(dpi ?? 96, 24, 300), limit);
        using var png = new MemoryStream();
        image.SavePng(png);
        return Results.File(png.ToArray(), "image/png");
    }));

// The text a page draws and where, in points (a point is dpi / 72 pixels of the page's image): what the component lays
// over the image so it can be selected and found.
app.MapGet("/documents/{id}/pages/{page:int}/text", (string id, int page, CancellationToken aborted) =>
    documents.Run(aborted, limit =>
        documents.Find(id) is { } entry && page >= 1 && page <= entry.Info.Pages
            ? Results.Ok(entry.Document.GetPageText(page - 1, limit))
            : Results.NotFound()));

app.MapGet("/documents/{id}/text", (string id, CancellationToken aborted) =>
    documents.Run(aborted, limit =>
    {
        if (documents.Find(id) is not { } entry)
        {
            return Results.NotFound();
        }
        var pages = new List<PageText>(entry.Info.Pages);
        for (int page = 0; page < entry.Info.Pages; page++)
        {
            pages.Add(entry.Document.GetPageText(page, limit));
        }
        return Results.Ok(new { pages });
    }));

app.MapGet("/documents/{id}/pdf", (string id, CancellationToken aborted) =>
    documents.Run(aborted, limit =>
    {
        if (documents.Find(id) is not { } entry)
        {
            return Results.NotFound();
        }
        using var pdf = new MemoryStream();
        entry.Document.ExportPdf(pdf, limit);
        // No file name, so the browser shows it instead of saving it.
        return Results.File(pdf.ToArray(), "application/pdf");
    }));

app.MapDelete("/documents/{id}", (string id, CancellationToken aborted) =>
    documents.Run(aborted, _ => documents.Close(id) ? Results.NoContent() : Results.NotFound()));

await app.RunAsync();
documents.Dispose();
Directory.Delete(workRoot, recursive: true);
return 0;

/// <summary>What the service says about an open document.</summary>
sealed record DocumentInfo(string Id, int Pages, string Kind, IReadOnlyList<FontInfo> Fonts);

/// <summary>A font the document asked for, and the one it was drawn in instead.</summary>
sealed record FontInfo(string Requested, string Used);

// The documents this process has open, and the gate that takes them to the engine one at a time. Everything but
// Dispose runs inside Run, behind the gate.
sealed class OpenDocuments : IDisposable
{
    public sealed record Entry(Document Document, string Copy, DocumentInfo Info);

    private readonly SemaphoreSlim gate = new(1, 1);
    private readonly Dictionary<string, Entry> open = new();
    private readonly Queue<string> opened = new();
    private readonly int capacity;
    private readonly TimeSpan timeLimit;

    public OpenDocuments(int capacity, TimeSpan timeLimit)
    {
        this.capacity = Math.Max(1, capacity);
        this.timeLimit = timeLimit;
    }

    // Waits for the engine, then runs work with a token that the browser going away or the time limit cancels. A
    // cancelled call changes nothing, so the document stays open and usable.
    public async Task<IResult> Run(CancellationToken aborted, Func<CancellationToken, IResult> work)
    {
        await gate.WaitAsync(aborted);
        try
        {
            using var limit = CancellationTokenSource.CreateLinkedTokenSource(aborted);
            limit.CancelAfter(timeLimit);
            return work(limit.Token);
        }
        catch (OperationCanceledException) when (!aborted.IsCancellationRequested)
        {
            return Results.Problem($"The document took longer than {timeLimit.TotalSeconds:0} seconds.", statusCode: 504);
        }
        catch (ComOfficeKitException e)
        {
            // The engine's own words, which can name this service's files. A service for others logs them and answers
            // in its own.
            return Results.Problem(e.Message, statusCode: e.Kind switch
            {
                ErrorKind.UnsupportedFormat or ErrorKind.Parse or ErrorKind.Unsupported => 422,
                ErrorKind.LimitExceeded => 413,
                ErrorKind.License => 503,
                _ => 500,
            });
        }
        finally
        {
            gate.Release();
        }
    }

    // Opens the copy and keeps it open under id, closing the one open longest past capacity. Counting the pages lays
    // the document out, which also tells which fonts it was drawn in instead of its own.
    public IResult Open(string id, string copy, CancellationToken limit)
    {
        Document document;
        try
        {
            document = Document.Open(copy, limit);
        }
        catch
        {
            File.Delete(copy);
            throw;
        }
        DocumentInfo info;
        try
        {
            info = new DocumentInfo(id, document.GetPageCount(limit), document.Kind.ToString(),
                                    document.FontSubstitutions.Select(s => new FontInfo(s.Requested, s.Used)).ToArray());
        }
        catch
        {
            document.Dispose();
            File.Delete(copy);
            throw;
        }
        open[id] = new Entry(document, copy, info);
        opened.Enqueue(id);
        while (open.Count > capacity && opened.TryDequeue(out var oldest))
        {
            Close(oldest);
        }
        return Results.Created($"documents/{id}", info);
    }

    public Entry? Find(string id) => open.TryGetValue(id, out var entry) ? entry : null;

    public bool Close(string id)
    {
        if (!open.Remove(id, out var entry))
        {
            return false;
        }
        entry.Document.Dispose();
        File.Delete(entry.Copy);
        return true;
    }

    public void Dispose()
    {
        foreach (var entry in open.Values)
        {
            entry.Document.Dispose();
        }
        open.Clear();
        gate.Dispose();
    }
}
