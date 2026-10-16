# Pruebas del contrato de apps-script/Code.gs contra la URL real (plan, sección 4.10).
# La URL y el token se pasan por variables de entorno; nunca se escriben en este fichero:
#   $env:SHEET_API = 'https://script.google.com/macros/s/…/exec'
#   $env:COACH_TOKEN = '…'
#   powershell -ExecutionPolicy Bypass -File apps-script\test-api.ps1
# Deja filas de prueba en la hoja (no se borran filas) y revoca al final los deportistas que crea.

$ErrorActionPreference = 'Stop'
$api = $env:SHEET_API
$coach = $env:COACH_TOKEN
if (-not $api -or -not $coach) { throw 'Faltan las variables de entorno SHEET_API y COACH_TOKEN' }

$script:fails = 0
function Check($name, [bool]$ok, $detail) {
  if ($ok) { Write-Host "OK    $name" -ForegroundColor Green }
  else { $script:fails++; Write-Host "FALLO $name  $detail" -ForegroundColor Red }
}

function Call($payload) {
  $json = $payload | ConvertTo-Json -Depth 20 -Compress
  $bytes = [Text.Encoding]::UTF8.GetBytes($json)
  # text/plain: petición simple, igual que la app (sin preflight CORS).
  # Google responde 302 hacia script.googleusercontent.com; se sigue a mano con un GET, porque la
  # redirección automática de Invoke-RestMethod (PowerShell 5.1) a veces acaba en 404.
  $req = [Net.HttpWebRequest]::Create($api)
  $req.Method = 'POST'; $req.AllowAutoRedirect = $false; $req.ContentType = 'text/plain;charset=utf-8'
  $stream = $req.GetRequestStream(); $stream.Write($bytes, 0, $bytes.Length); $stream.Close()
  try { $res = $req.GetResponse() } catch [Net.WebException] { $res = $_.Exception.Response; if (-not $res) { throw } }
  $location = $res.Headers['Location']
  if ($location) { $res.Close(); $res = [Net.WebRequest]::Create($location).GetResponse() }
  $text = (New-Object IO.StreamReader($res.GetResponseStream(), [Text.Encoding]::UTF8)).ReadToEnd()
  $res.Close()
  $text | ConvertFrom-Json
}

function PullAll($token) {
  $since = 0; $all = @()
  do {
    $res = Call @{ action = 'pull'; token = $token; since = $since }
    if (-not $res.ok) { throw "pull falló: $($res.error)" }
    $all += @($res.records); $since = $res.cursor
  } while ($res.more)
  return $all
}

$run = [DateTimeOffset]::UtcNow.ToUnixTimeMilliseconds()

# T1 — GET a la URL
$res = Invoke-RestMethod -Uri $api -Method Get
Check 'T1 GET devuelve ok' ($res.ok -eq $true) ($res | ConvertTo-Json -Compress)

# T2 — ping con token inventado
$res = Call @{ action = 'ping'; token = 'token-inventado' }
Check 'T2 ping con token inventado: error token' ($res.ok -eq $false -and $res.error -eq 'token') ($res | ConvertTo-Json -Compress)

# T3 — ping con el token de entrenador
$res = Call @{ action = 'ping'; token = $coach }
Check 'T3 ping de entrenador: role coach' ($res.ok -and $res.role -eq 'coach') ($res | ConvertTo-Json -Compress)

# T4 — invite (deportista A)
$a = Call @{ action = 'invite'; token = $coach }
Check 'T4 invite devuelve id_unico y token' ($a.ok -and $a.id_unico -match '^d[0-9a-f]{8}$' -and $a.token) ($a | ConvertTo-Json -Compress)
$res = Call @{ action = 'ping'; token = $a.token }
Check 'T4 el token nuevo es de deportista' ($res.ok -and $res.role -eq 'athlete' -and $res.id_unico -eq $a.id_unico) ($res | ConvertTo-Json -Compress)

# T5 — perfil con name y birthDate: se guarda sin ellos
$profile = @{ store = 'profiles'; id = $a.id_unico; userId = $a.id_unico; updatedAt = $run; deleted = $false
  data = @{ hrMax = 190; hrRest = 50; weight = 70; method = 'karvonen'; name = 'Prueba'; birthDate = '1990-01-01' } }
$res = Call @{ action = 'push'; token = $a.token; records = @($profile) }
Check 'T5 push de perfil aceptado' ($res.ok -and @($res.accepted).Count -eq 1) ($res | ConvertTo-Json -Compress -Depth 10)
$p = PullAll $a.token | Where-Object { $_.store -eq 'profiles' -and $_.id -eq $a.id_unico }
$names = @($p.data.PSObject.Properties.Name)
Check 'T5 el perfil no tiene name ni birthDate' ($p -and $names -contains 'hrMax' -and $names -notcontains 'name' -and $names -notcontains 'birthDate') ($names -join ',')

# T6 — sesión con updatedAt más antiguo que el guardado: older con current
$sid = "t$run" + 's'
$sess = @{ store = 'sessions'; id = $sid; userId = $a.id_unico; updatedAt = $run; deleted = $false
  data = @{ id = $sid; userId = $a.id_unico; type = 'Carrera'; status = 'done'; rpePost = 6 } }
$res = Call @{ action = 'push'; token = $a.token; records = @($sess) }
Check 'T6 primera versión de la sesión aceptada' ($res.ok -and @($res.accepted).Count -eq 1) ($res | ConvertTo-Json -Compress -Depth 10)
$sess.updatedAt = $run - 1000; $sess.data.rpePost = 2
$res = Call @{ action = 'push'; token = $a.token; records = @($sess) }
$r = @($res.rejected)[0]
Check 'T6 versión antigua rechazada como older con current' ($res.ok -and $r.reason -eq 'older' -and $r.current.data.rpePost -eq 6) ($res | ConvertTo-Json -Compress -Depth 10)

# T7 — deportista A escribe con userId de B
$b = Call @{ action = 'invite'; token = $coach }
$bid = "t$run" + 'b'
$res = Call @{ action = 'push'; token = $coach; records = @(@{ store = 'tests'; id = $bid; userId = $b.id_unico; updatedAt = $run; deleted = $false; data = @{ id = $bid; note = 'de B' } }) }
Check 'T7 el entrenador puede escribir para B' ($res.ok -and @($res.accepted).Count -eq 1) ($res | ConvertTo-Json -Compress -Depth 10)
$aid = "t$run" + 'x'
$res = Call @{ action = 'push'; token = $a.token; records = @(@{ store = 'tests'; id = $aid; userId = $b.id_unico; updatedAt = $run; deleted = $false; data = @{ id = $aid } }) }
Check 'T7 A no puede escribir con userId de B' ($res.ok -and @($res.rejected)[0].reason -eq 'forbidden') ($res | ConvertTo-Json -Compress -Depth 10)

# T8 — pull de A no trae nada de B (y el del entrenador sí)
$recsA = PullAll $a.token
Check 'T8 pull de A sin registros de B' (@($recsA | Where-Object { $_.userId -ne $a.id_unico }).Count -eq 0 -and @($recsA).Count -ge 2) ("$(@($recsA).Count) registros")
$recsC = PullAll $coach
Check 'T8 pull del entrenador trae lo de A y lo de B' ((@($recsC | Where-Object { $_.id -eq $bid }).Count -eq 1) -and (@($recsC | Where-Object { $_.id -eq $sid }).Count -eq 1)) ''

# T9 — track de 100 000 caracteres: 3 partes, vuelve entero e idéntico
$chars = 'ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789-_'
$rnd = New-Object System.Random 42
$sb = New-Object System.Text.StringBuilder 100000
for ($i = 0; $i -lt 100000; $i++) { [void]$sb.Append($chars[$rnd.Next($chars.Length)]) }
$big = $sb.ToString()
$res = Call @{ action = 'push'; token = $a.token; records = @(@{ store = 'tracks'; id = $sid; userId = $a.id_unico; updatedAt = $run; deleted = $false; data = $big }) }
Check 'T9 track de 100 000 caracteres aceptado' ($res.ok -and @($res.accepted).Count -eq 1) ($res | ConvertTo-Json -Compress -Depth 10)
$t = PullAll $a.token | Where-Object { $_.store -eq 'tracks' -and $_.id -eq $sid }
Check 'T9 el track vuelve entero e idéntico' ($t -and $t.data -ceq $big) ("longitud $($t.data.Length)")
Write-Host "      (en la hoja, pestaña tracks, el id $sid debe tener 3 filas: part 0, 1 y 2)"

# T10 — invite con token de deportista
$res = Call @{ action = 'invite'; token = $a.token }
Check 'T10 invite con token de deportista: forbidden' ($res.ok -eq $false -and $res.error -eq 'forbidden') ($res | ConvertTo-Json -Compress)

# T11 — revoke de A y ping con su token
$res = Call @{ action = 'revoke'; token = $coach; id_unico = $a.id_unico }
Check 'T11 revoke aceptado' ($res.ok -eq $true) ($res | ConvertTo-Json -Compress)
$res = Call @{ action = 'ping'; token = $a.token }
Check 'T11 el token revocado ya no vale' ($res.ok -eq $false -and $res.error -eq 'token') ($res | ConvertTo-Json -Compress)

# Limpieza: B también queda revocado
$null = Call @{ action = 'revoke'; token = $coach; id_unico = $b.id_unico }

Write-Host ''
if ($script:fails) { Write-Host "$script:fails comprobaciones fallidas" -ForegroundColor Red; exit 1 }
Write-Host 'Todas las comprobaciones pasan' -ForegroundColor Green
