# Clicker – servidor del portátil

Puente entre PowerPoint de escritorio (Windows) y el móvil, por WebSocket.

```powershell
cd server
dotnet run -- --port 5055 --token 123456   # --token es opcional (si no, se genera uno)
```

Abre en el móvil `http://<IP-del-portátil>:5055/?token=<código>` (la consola muestra las URLs).
La primera vez, Windows pedirá permitir el servidor en el firewall (red privada).

## Protocolo (`ws://host:5055/ws?token=...`)
Servidor → cliente, 5 veces/s: JSON con `slide`, `slideCount`, `title`, `notes`, `slideElapsed`,
`slideBudget`, `totalElapsed`, `totalBudget`, `behindSeconds` (>0 = retraso), `slideOvertime`, etc.

Cliente → servidor: `{"cmd":"next"|"prev"|"start"|"end"|"timerStart"|"timerPause"|"timerReset"}`,
`{"cmd":"goto","slide":N}`, `{"cmd":"setTotal","minutes":N}`.

## Tiempos previstos
Pon `[m:ss]` en las notas de cada diapositiva (p. ej. `[3:00]`). Con `setTotal`, las que no
lleven etiqueta se reparten el tiempo restante. Las etiquetas se leen al abrir la presentación.
