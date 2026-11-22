using System.Collections.Concurrent;
using System.Runtime.InteropServices;

namespace Clicker.Server;

/// <summary>Estado de PowerPoint en un instante.</summary>
public sealed record PptSnapshot(
    bool Connected,
    bool InShow,
    string Presentation,
    int Slide,
    int Count,
    string Title,
    string Notes,
    string[] AllNotes,
    string Key = "",
    int ClickIndex = 0,
    int ClickCount = 0)
{
    public static readonly PptSnapshot Disconnected =
        new(false, false, "", 0, 0, "", "", Array.Empty<string>());
}

/// <summary>
/// Puente con PowerPoint de escritorio por COM (enlace tardío, sin ensamblados de interop).
/// Todas las llamadas COM ocurren en un hilo STA propio; el resto del programa solo
/// lee <see cref="Snapshot"/> y encola órdenes con <see cref="Send"/>.
/// </summary>
public sealed class PowerPointBridge : IDisposable
{
    private const int PollMs = 200;

    private readonly ConcurrentQueue<(string Cmd, int Arg)> _commands = new();
    private readonly Thread _thread;
    private volatile bool _stop;
    private volatile PptSnapshot _snapshot = PptSnapshot.Disconnected;

    public PptSnapshot Snapshot => _snapshot;

    private readonly bool _useKeys;

    /// <param name="useKeys">Avanzar simulando la pulsación de flechas en vez de la API COM.</param>
    public PowerPointBridge(bool useKeys = false)
    {
        _useKeys = useKeys;
        _thread = new Thread(Run) { IsBackground = true, Name = "PowerPointBridge" };
        _thread.SetApartmentState(ApartmentState.STA);
        _thread.Start();
    }

    public void Send(string cmd, int arg = 0) => _commands.Enqueue((cmd, arg));

    public void Dispose()
    {
        _stop = true;
        _thread.Join(2000);
    }

    // ---------------------------------------------------------------- hilo STA

    private void Run()
    {
        dynamic? app = null;
        string notesKey = "";
        string[] allNotes = Array.Empty<string>();

        while (!_stop)
        {
            try
            {
                app ??= GetPowerPoint();
                if (app is null)
                {
                    _snapshot = PptSnapshot.Disconnected;
                    _commands.Clear();
                    while (_exports.TryDequeue(out var lost)) lost.Tcs.TrySetResult(null);
                }
                else
                {
                    Drain(app);
                    DrainExports(app);
                    _snapshot = Read(app, ref notesKey, ref allNotes);
                }
            }
            catch (COMException)
            {
                // PowerPoint se cerró o está ocupado: reintentar la conexión.
                app = null;
                _snapshot = PptSnapshot.Disconnected;
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine($"[ppt] {ex.GetType().Name}: {ex.Message}");
                app = null;
                _snapshot = PptSnapshot.Disconnected;
            }

            Thread.Sleep(PollMs);
        }
    }

    private void Drain(dynamic app)
    {
        while (_commands.TryDequeue(out var c))
        {
            try
            {
                dynamic? view = ShowView(app);
                switch (c.Cmd)
                {
                    case "next":
                        if (view is null) break;
                        if (!(_useKeys && PressKey(app, VK_RIGHT))) view.Next();   // animaciones y diapositivas, como un clicker
                        break;
                    case "prev":
                        if (view is null) break;
                        if (!(_useKeys && PressKey(app, VK_LEFT))) view.Previous();
                        break;
                    case "goto":  view?.GotoSlide(c.Arg); break;
                    case "end":   view?.Exit(); break;
                    case "start":
                        if (view is null && (int)app.Presentations.Count > 0)
                            app.ActivePresentation.SlideShowSettings.Run();
                        break;
                }
            }
            catch (COMException ex)
            {
                Console.Error.WriteLine($"[ppt] comando '{c.Cmd}' falló: {ex.Message}");
            }
        }
    }

    // ------------------------------------------------------ miniaturas de diapositivas

    private readonly ConcurrentQueue<(int Slide, int Click, TaskCompletionSource<string?> Tcs)> _exports = new();
    private string _copyKey = "";
    private string _copyPath = "";
    private static readonly string CacheRoot = Path.Combine(Path.GetTempPath(), "clicker-slides");

    /// <summary>Identificador corto y estable de una presentación (para cachés y URLs).</summary>
    public static string Hash(string key) =>
        Convert.ToHexString(System.Security.Cryptography.SHA1.HashData(System.Text.Encoding.UTF8.GetBytes(key)))[..8];

    /// <summary>Borra miniaturas de ejecuciones anteriores (la presentación pudo cambiar).</summary>
    public static void ClearCache()
    {
        try { if (Directory.Exists(CacheRoot)) Directory.Delete(CacheRoot, true); }
        catch (IOException) { }
    }

    /// <summary>
    /// Exporta la diapositiva a PNG (con caché) tal como se vería tras <paramref name="click"/> clics
    /// de animación (0 = al llegar a ella). Devuelve la ruta, o null si no se pudo.
    /// </summary>
    public Task<string?> ExportSlideAsync(int slide, int click = 0)
    {
        var tcs = new TaskCompletionSource<string?>(TaskCreationOptions.RunContinuationsAsynchronously);
        _exports.Enqueue((slide, Math.Max(0, click), tcs));
        return tcs.Task;
    }

    /// <summary>
    /// Copia de la presentación (SaveCopyAs) sobre la que se oculta lo que aún no ha aparecido.
    /// Así nunca se toca la presentación en pantalla.
    /// </summary>
    private string EnsureCopy(dynamic pres, string key)
    {
        if (_copyKey == key && File.Exists(_copyPath)) return _copyPath;
        string dir = Path.Combine(CacheRoot, Hash(key));
        Directory.CreateDirectory(dir);
        string path = Path.Combine(dir, "deck.pptx");
        pres.SaveCopyAs(path);
        _copyKey = key;
        _copyPath = path;
        return path;
    }

    private void ExportWithState(dynamic app, dynamic pres, string key, int slideIdx, int click, string file, int px, int py)
    {
        string copyPath = EnsureCopy(pres, key);
        dynamic copy = app.Presentations.Open(copyPath, -1 /*solo lectura*/, 0 /*con título*/, 0 /*sin ventana*/);
        try
        {
            dynamic slide = copy.Slides.Item(slideIdx);
            ApplyClickState(slide, click);
            slide.Export(file, "PNG", px, py);
        }
        finally
        {
            try { copy.Saved = -1; copy.Close(); } catch (COMException) { }
        }
    }

    /// <summary>
    /// Deja la diapositiva (de la copia) como se vería tras <paramref name="clicks"/> clics:
    /// oculta lo que aún no ha entrado y lo que ya ha salido. Solo efectos de entrada y salida
    /// (tipos 1–53); los de énfasis y trayectorias no cambian la visibilidad.
    /// </summary>
    private static void ApplyClickState(dynamic slide, int clicks)
    {
        dynamic seq = slide.TimeLine.MainSequence;
        int n = (int)seq.Count;
        var effects = new List<(dynamic Shape, int Id, int Para, bool Entrance, int ClickNo)>();
        int click = 0;
        for (int i = 1; i <= n; i++)
        {
            try
            {
                dynamic e = seq.Item(i);
                if ((int)e.Timing.TriggerType == 1) click++;   // 1 = al hacer clic
                int type = (int)e.EffectType;
                if (type < 1 || type > 53) continue;
                dynamic shape = e.Shape;

                // Effect.Paragraph falla en formas sin texto (imágenes, gráficos...): efecto sobre toda la forma.
                int para = 0;
                try { para = (int)e.Paragraph; } catch (COMException) { }

                effects.Add((shape, (int)shape.Id, para, (int)e.Exit == 0, click));
            }
            catch (Exception ex) when (ex is COMException or Microsoft.CSharp.RuntimeBinder.RuntimeBinderException)
            {
                Console.Error.WriteLine($"[ppt] efecto {i} ignorado: {ex.Message}");
            }
        }

        // Visibilidad inicial: lo que tiene efecto de entrada empieza oculto.
        var visible = new Dictionary<(int, int), bool>();
        var targets = new Dictionary<(int, int), dynamic>();
        foreach (var e in effects)
        {
            var k = (e.Id, e.Para);
            if (!visible.ContainsKey(k)) { visible[k] = !e.Entrance; targets[k] = e.Shape; }
        }
        foreach (var e in effects)
            if (e.ClickNo <= clicks) visible[(e.Id, e.Para)] = e.Entrance;

        foreach (var (k, isVisible) in visible)
        {
            if (isVisible) continue;
            dynamic shape = targets[k];
            try
            {
                if (k.Item2 == 0) shape.Visible = 0;
                else
                {
                    dynamic para = shape.TextFrame2.TextRange.Paragraphs(k.Item2, 1);
                    para.Font.Fill.Transparency = 1f;           // conserva el hueco de la línea
                    para.ParagraphFormat.Bullet.Visible = 0;
                }
            }
            catch (Exception ex) when (ex is COMException or Microsoft.CSharp.RuntimeBinder.RuntimeBinderException) { }
        }
    }

    private void DrainExports(dynamic app)
    {
        while (_exports.TryDequeue(out var req))
        {
            try
            {
                dynamic pres = (int)app.SlideShowWindows.Count > 0
                    ? app.SlideShowWindows.Item(1).Presentation
                    : app.ActivePresentation;
                int count = (int)pres.Slides.Count;
                if (req.Slide < 1 || req.Slide > count) { req.Tcs.SetResult(null); continue; }

                string key = $"{(string)pres.FullName}|{count}";
                string dir = Path.Combine(CacheRoot, Hash(key));
                Directory.CreateDirectory(dir);
                string file = Path.Combine(dir, $"{req.Slide}_{req.Click}.png");
                if (!File.Exists(file))
                {
                    const int width = 640;
                    double ratio = (double)pres.PageSetup.SlideHeight / (double)pres.PageSetup.SlideWidth;
                    int height = (int)Math.Round(width * ratio);
                    dynamic live = pres.Slides.Item(req.Slide);
                    if ((int)live.TimeLine.MainSequence.Count == 0)
                        live.Export(file, "PNG", width, height);          // sin animaciones: exportación directa
                    else
                        ExportWithState(app, pres, key, req.Slide, req.Click, file, width, height);
                }
                req.Tcs.SetResult(File.Exists(file) ? file : null);
            }
            catch (Exception e)
            {
                Console.Error.WriteLine($"[ppt] exportar diapositiva {req.Slide} falló: {e.Message}");
                req.Tcs.TrySetResult(null);
            }
        }
    }

    private const byte VK_LEFT = 0x25, VK_RIGHT = 0x27;

    /// <summary>
    /// Pulsa una flecha como lo haría un clicker físico (útil si la API COM no reproduce
    /// igual alguna transición o animación). Devuelve false si no pudo hacerlo.
    /// </summary>
    private static bool PressKey(dynamic app, byte vk)
    {
        try
        {
            var hwnd = new IntPtr((long)app.SlideShowWindows.Item(1).HWND);
            if (hwnd == IntPtr.Zero) return false;
            if (GetForegroundWindow() != hwnd)
            {
                SetForegroundWindow(hwnd);
                Thread.Sleep(50);
                if (GetForegroundWindow() != hwnd) return false; // sin foco, mejor usar COM
            }
            const uint EXTENDED = 1, KEYUP = 2;
            keybd_event(vk, 0, EXTENDED, UIntPtr.Zero);
            keybd_event(vk, 0, EXTENDED | KEYUP, UIntPtr.Zero);
            return true;
        }
        catch (Exception e) when (e is COMException or Microsoft.CSharp.RuntimeBinder.RuntimeBinderException)
        {
            return false;
        }
    }

    [DllImport("user32.dll")] private static extern IntPtr GetForegroundWindow();
    [DllImport("user32.dll")] private static extern bool SetForegroundWindow(IntPtr hWnd);
    [DllImport("user32.dll")] private static extern void keybd_event(byte vk, byte scan, uint flags, UIntPtr extra);

    private static dynamic? ShowView(dynamic app)
    {
        if ((int)app.SlideShowWindows.Count > 0)
            return app.SlideShowWindows.Item(1).View;
        return null;
    }

    private static PptSnapshot Read(dynamic app, ref string notesKey, ref string[] allNotes)
    {
        dynamic? view = ShowView(app);
        bool inShow = view is not null;

        if ((int)app.Presentations.Count == 0)
            return new PptSnapshot(true, false, "", 0, 0, "", "", Array.Empty<string>());

        dynamic pres = inShow ? app.SlideShowWindows.Item(1).Presentation : app.ActivePresentation;
        int count = (int)pres.Slides.Count;

        int idx;
        if (inShow)
        {
            idx = (int)view!.CurrentShowPosition;
        }
        else
        {
            try { idx = (int)app.ActiveWindow.View.Slide.SlideIndex; }
            catch { idx = 1; }
        }
        idx = Math.Clamp(idx, 1, Math.Max(count, 1));

        string name = (string)pres.Name;
        string key = $"{(string)pres.FullName}|{count}";
        if (key != notesKey)
        {
            var notes = new string[count];
            for (int i = 1; i <= count; i++)
                notes[i - 1] = NotesOf(pres.Slides.Item(i));
            allNotes = notes;
            notesKey = key;
        }

        dynamic slide = pres.Slides.Item(idx);
        int clickIndex = 0, clickCount = 0;
        if (inShow)
        {
            try { clickIndex = (int)view!.GetClickIndex(); clickCount = (int)view!.GetClickCount(); }
            catch (COMException) { }
        }

        return new PptSnapshot(true, inShow, name, idx, count, TitleOf(slide), NotesOf(slide), allNotes,
            key, clickIndex, clickCount);
    }

    private static string TitleOf(dynamic slide)
    {
        try
        {
            if ((int)slide.Shapes.HasTitle != 0)
                return (string)slide.Shapes.Title.TextFrame.TextRange.Text;
        }
        catch (COMException) { }
        return "";
    }

    private static string NotesOf(dynamic slide)
    {
        try
        {
            foreach (dynamic shape in slide.NotesPage.Shapes)
            {
                // msoPlaceholder = 14, ppPlaceholderBody = 2
                if ((int)shape.Type == 14 && (int)shape.PlaceholderFormat.Type == 2
                    && (int)shape.HasTextFrame != 0)
                    return (string)shape.TextFrame.TextRange.Text;
            }
        }
        catch (COMException) { }
        return "";
    }

    // ------------------------------------------------------- enganche a PowerPoint

    [DllImport("ole32.dll")]
    private static extern int CLSIDFromProgID([MarshalAs(UnmanagedType.LPWStr)] string progId, out Guid clsid);

    [DllImport("oleaut32.dll")]
    private static extern int GetActiveObject(ref Guid rclsid, IntPtr pvReserved,
        [MarshalAs(UnmanagedType.IUnknown)] out object ppunk);

    private static dynamic? GetPowerPoint()
    {
        if (CLSIDFromProgID("PowerPoint.Application", out var clsid) != 0) return null;
        return GetActiveObject(ref clsid, IntPtr.Zero, out var obj) == 0 ? obj : null;
    }
}
