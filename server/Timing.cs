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
    private static readonly TimeSpan EndGrace = TimeSpan.FromSeconds(3);
    private bool _running, _active;
    private DateTime? _notInShowSince;
    private string _deckKey = "";
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

    /// <summary>
    /// Actualiza el cronómetro con el estado actual de PowerPoint.
    /// Al empezar una presentación el reloj arranca; si empieza desde la diapositiva 1 (o es otra
    /// presentación) se pone a cero, y si empieza desde otra se reanuda lo acumulado.
    /// Al terminarla se detiene, pero solo tras <see cref="EndGrace"/> sin presentación, para que un
    /// fallo momentáneo de lectura no pare el reloj en mitad de la charla.
    /// </summary>
    public TimerSnapshot Update(bool inShow, int slide, string deckKey)
    {
        lock (_lock)
        {
            Tick();
            if (slide > 0) _slide = slide;   // sin presentación se conserva la última diapositiva
            var now = DateTime.UtcNow;

            if (inShow)
            {
                _notInShowSince = null;
                if (!_active)
                {
                    _active = true;
                    if (slide == 1 || deckKey != _deckKey)
                    {
                        _total = 0;
                        _per.Clear();
                    }
                    _deckKey = deckKey;
                    _running = true;
                }
            }
            else if (_active)
            {
                _notInShowSince ??= now;
                var away = now - _notInShowSince.Value;
                if (away >= EndGrace)
                {
                    // El tiempo de espera no era charla: se descuenta del total y de la diapositiva.
                    double s = away.TotalSeconds;
                    if (_running)
                    {
                        _total = Math.Max(0, _total - s);
                        if (_slide > 0) _per[_slide] = Math.Max(0, _per.GetValueOrDefault(_slide) - s);
                    }
                    _active = false;
                    _running = false;
                }
            }

            return new TimerSnapshot(_running, _total, new Dictionary<int, double>(_per));
        }
    }

    public void Start()  { lock (_lock) { Tick(); _running = true; } }
    public void Pause()  { lock (_lock) { Tick(); _running = false; } }
    public void Reset()  { lock (_lock) { _running = false; _active = false; _notInShowSince = null; _total = 0; _per.Clear(); _last = DateTime.UtcNow; } }
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
