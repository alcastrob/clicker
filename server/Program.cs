using System.Collections.Concurrent;
using System.Net.NetworkInformation;
using System.Net.Sockets;
using System.Net.WebSockets;
using System.Security.Cryptography;
using System.Text;
using System.Text.Json;
using Clicker.Server;

var builder = WebApplication.CreateBuilder(args);
int port = builder.Configuration.GetValue("port", 5055);
string token = builder.Configuration["token"] ?? RandomNumberGenerator.GetInt32(100000, 1000000).ToString();
builder.WebHost.UseUrls($"http://0.0.0.0:{port}");

var app = builder.Build();
app.UseWebSockets();
app.UseDefaultFiles();
app.UseStaticFiles();

// --keys true: avanza simulando la pulsación de flechas (como un clicker físico) en vez de la API COM.
using var bridge = new PowerPointBridge(builder.Configuration.GetValue("keys", false));
var timer = new SessionTimer();
var clients = new ConcurrentDictionary<Guid, WebSocket>();
int totalSeconds = 0; // duración total prevista; 0 = solo etiquetas [m:ss] de las notas
var json = new JsonSerializerOptions(JsonSerializerDefaults.Web);

app.Map("/ws", async ctx =>
{
    if (!ctx.WebSockets.IsWebSocketRequest) { ctx.Response.StatusCode = 400; return; }

    var given = Encoding.UTF8.GetBytes(ctx.Request.Query["token"].ToString());
    if (!CryptographicOperations.FixedTimeEquals(given, Encoding.UTF8.GetBytes(token)))
    {
        ctx.Response.StatusCode = 401;
        return;
    }

    using var ws = await ctx.WebSockets.AcceptWebSocketAsync();
    var id = Guid.NewGuid();
    clients[id] = ws;
    var buffer = new byte[4096];
    try
    {
        while (ws.State == WebSocketState.Open)
        {
            var r = await ws.ReceiveAsync(buffer, ctx.RequestAborted);
            if (r.MessageType == WebSocketMessageType.Close) break;
            if (!r.EndOfMessage) continue; // las órdenes son diminutas
            HandleCommand(Encoding.UTF8.GetString(buffer, 0, r.Count));
        }
    }
    catch (Exception e) when (e is OperationCanceledException or WebSocketException) { }
    finally { clients.TryRemove(id, out _); }
});

void HandleCommand(string text)
{
    try
    {
        using var doc = JsonDocument.Parse(text);
        var root = doc.RootElement;
        switch (root.GetProperty("cmd").GetString())
        {
            case "next":       bridge.Send("next"); break;
            case "prev":       bridge.Send("prev"); break;
            case "start":      bridge.Send("start"); break;
            case "end":        bridge.Send("end"); break;
            case "goto":       bridge.Send("goto", root.GetProperty("slide").GetInt32()); break;
            case "timerStart": timer.Start(); break;
            case "timerPause": timer.Pause(); break;
            case "timerReset": timer.Reset(); break;
            case "setTotal":   totalSeconds = Math.Max(0, root.GetProperty("minutes").GetInt32()) * 60; break;
        }
    }
    catch (Exception e) when (e is JsonException or KeyNotFoundException or InvalidOperationException) { }
}

// Difusión del estado a todos los clientes, 5 veces por segundo.
_ = Task.Run(async () =>
{
    var ct = app.Lifetime.ApplicationStopping;
    int[] budgets = Array.Empty<int>();
    string[]? budgetsFor = null;
    int budgetsTotal = -1;

    while (!ct.IsCancellationRequested)
    {
        var ppt = bridge.Snapshot;
        var t = timer.Update(ppt.InShow, ppt.Slide);

        if (!ReferenceEquals(budgetsFor, ppt.AllNotes) || budgetsTotal != totalSeconds)
        {
            budgets = Budgets.Compute(ppt.AllNotes, totalSeconds);
            budgetsFor = ppt.AllNotes;
            budgetsTotal = totalSeconds;
        }

        int slideBudget = ppt.Slide is > 0 && ppt.Slide <= budgets.Length ? budgets[ppt.Slide - 1] : 0;
        int plannedBefore = 0;
        for (int i = 0; i < ppt.Slide - 1 && i < budgets.Length; i++) plannedBefore += budgets[i];
        double slideElapsed = t.PerSlide.GetValueOrDefault(ppt.Slide);

        var state = new
        {
            connected = ppt.Connected,
            inSlideShow = ppt.InShow,
            presentation = ppt.Presentation,
            slide = ppt.Slide,
            slideCount = ppt.Count,
            title = ppt.Title,
            notes = Budgets.ForDisplay(ppt.Notes),
            running = t.Running,
            totalElapsed = t.Total,
            totalBudget = totalSeconds > 0 ? totalSeconds : budgets.Sum(),
            slideElapsed,
            slideBudget,
            slideOvertime = slideBudget > 0 && slideElapsed > slideBudget,
            // > 0: vas con retraso; < 0: vas adelantado (tiempo gastado antes de esta diapositiva vs. previsto)
            behindSeconds = ppt.Slide > 0 ? (t.Total - slideElapsed) - plannedBefore : 0,
        };

        var bytes = JsonSerializer.SerializeToUtf8Bytes(state, json);
        foreach (var (id, ws) in clients)
        {
            try
            {
                if (ws.State == WebSocketState.Open)
                {
                    // Un cliente lento no debe bloquear al resto: 2 s máximo por envío.
                    using var send = CancellationTokenSource.CreateLinkedTokenSource(ct);
                    send.CancelAfter(TimeSpan.FromSeconds(2));
                    await ws.SendAsync(bytes, WebSocketMessageType.Text, true, send.Token);
                }
            }
            catch { clients.TryRemove(id, out _); }
        }

        try { await Task.Delay(200, ct); } catch (OperationCanceledException) { }
    }
});

Console.WriteLine($"Clicker listo. Código de acceso: {token}");
foreach (var ip in NetworkInterface.GetAllNetworkInterfaces()
             .Where(n => n.OperationalStatus == OperationalStatus.Up
                         && n.NetworkInterfaceType != NetworkInterfaceType.Loopback)
             .SelectMany(n => n.GetIPProperties().UnicastAddresses)
             .Where(a => a.Address.AddressFamily == AddressFamily.InterNetwork)
             .Select(a => a.Address))
{
    Console.WriteLine($"  http://{ip}:{port}/?token={token}");
}

app.Run();
