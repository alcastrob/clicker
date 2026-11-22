using System.Text.RegularExpressions;

namespace Clicker.Server;

public sealed record TimerSnapshot(bool Running, double Total, IReadOnlyDictionary<int, double> PerSlide);

/// <summary>
/// Cronómetro de la sesión: tiempo total y tiempo acumulado por diapositiva
/// (si vuelves a una diapositiva, su tiempo sigue sumando).
/// Arranca solo al iniciar la presentación, salvo que lo pauses a mano.
/// </summary>
public sealed class SessionTimer
{
    private readonly object _lock = new();
    private readonly Dictionary<int, double> _per = new();
    private bool _running, _manualPaused, _wasInShow;
    private double _total;
    private int _slide;
    private DateTime _last = DateTime.UtcNow;

    private void Tick()
    {
        var now = DateTime.UtcNow;
        if (_running)
        {
            var d = (now - _last).TotalSeconds;
            _total += d;
            if (_slide > 0) _per[_slide] = _per.GetValueOrDefault(_slide) + d;
        }
        _last = now;
    }

    public TimerSnapshot Update(bool inShow, int slide)
    {
        lock (_lock)
        {
            Tick();
            _slide = slide;
            if (inShow && !_running && !_manualPaused && _total == 0) _running = true;
            // Al terminar la presentación el reloj se detiene (se puede reanudar a mano).
            if (_wasInShow && !inShow) { _running = false; _manualPaused = true; }
            _wasInShow = inShow;
            return new TimerSnapshot(_running, _total, new Dictionary<int, double>(_per));
        }
    }

    public void Start()  { lock (_lock) { Tick(); _running = true;  _manualPaused = false; } }
    public void Pause()  { lock (_lock) { Tick(); _running = false; _manualPaused = true; } }
    public void Reset()  { lock (_lock) { _running = false; _manualPaused = false; _total = 0; _per.Clear(); _last = DateTime.UtcNow; } }
}

public static partial class Budgets
{
    // Etiqueta [m:ss] en las notas de la diapositiva, p. ej. "[3:00]" o "[0:45]".
    [GeneratedRegex(@"\[(\d{1,3}):([0-5]\d)\]")]
    private static partial Regex Tag();

    /// <summary>
    /// Notas listas para mostrar: sin la etiqueta [m:ss] y con saltos de línea normalizados
    /// (PowerPoint usa \r para el salto de párrafo y \v para el salto suave).
    /// </summary>
    public static string ForDisplay(string notes)
    {
        var text = Tag().Replace(notes, "")
            .Replace("\r\n", "\n").Replace('\r', '\n').Replace('\v', '\n');
        return text.Trim();
    }

    /// <summary>
    /// Segundos previstos por diapositiva. Las que llevan etiqueta usan su valor; si hay
    /// duración total, el resto se reparte a partes iguales entre las que no la llevan.
    /// </summary>
    public static int[] Compute(string[] notes, int totalSeconds)
    {
        var result = new int[notes.Length];
        var untagged = new List<int>();
        int tagged = 0;

        for (int i = 0; i < notes.Length; i++)
        {
            var m = Tag().Match(notes[i]);
            if (m.Success)
            {
                result[i] = int.Parse(m.Groups[1].Value) * 60 + int.Parse(m.Groups[2].Value);
                tagged += result[i];
            }
            else untagged.Add(i);
        }

        if (totalSeconds > 0 && untagged.Count > 0)
        {
            int each = Math.Max(0, (totalSeconds - tagged) / untagged.Count);
            foreach (var i in untagged) result[i] = each;
        }
        return result;
    }
}
