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
    string[] AllNotes)
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
                }
                else
                {
                    Drain(app);
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
        return new PptSnapshot(true, inShow, name, idx, count, TitleOf(slide), NotesOf(slide), allNotes);
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
