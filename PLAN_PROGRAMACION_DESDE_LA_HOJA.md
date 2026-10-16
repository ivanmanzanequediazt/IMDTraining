# Plan: programar entrenamientos desde la hoja de Google

| | |
|---|---|
| Qué es | Guía para que Iván escriba los entrenamientos en una pestaña de su hoja «IMDTraining datos» y, al pulsar **Publicar**, aparezcan solos en la app de cada deportista |
| Para quién | Una sesión de Claude Code que trabaja junto a **Iván** (entrenador, sin conocimientos de informática) |
| Estado | Diseño acordado con Iván, **sin empezar** |
| Base de código | Commit `f81049f` de `main` («Sincronización con Google Sheets»): `index.html` de 3164 líneas, `apps-script/Code.gs` de 534 líneas |
| Requisito previo | La sincronización de `PLAN_SINCRONIZACION_GOOGLE_SHEETS.md` funcionando (lo está, según Iván) |
| Fecha | 2026-10-16 |

## Cómo leer este documento

Igual que el plan de sincronización: **Para Iván** (lenguaje sencillo) y **Para Claude** (detalle técnico).
Las referencias del tipo `Code.gs:240` o `index.html:1617` son del commit `f81049f`. Si han cambiado,
localiza por nombre de función (`grep -n "function writeRecord_" apps-script/Code.gs`).

---

## 0. Instrucciones para la sesión de Claude

**Para Claude.** Lee el documento entero antes de hablar con Iván. Aplican las mismas reglas que en la
sección 0 de `PLAN_SINCRONIZACION_GOOGLE_SHEETS.md`:

- Una fase cada vez, un paso cada vez, esperando a que Iván confirme.
- Cada paso de Iván con las cinco partes: qué, por qué, qué haces tú (clic a clic, menús de Google en
  español con el inglés entre paréntesis), cómo sabes que ha ido bien, si algo sale distinto.
- Si una pantalla de Google no coincide, **pide captura** antes de improvisar.
- Rama de trabajo: `programacion-hoja`. Nada en `main` hasta la fase final.
- **Nunca borres** filas de la hoja ni datos de los móviles sin confirmación explícita.
- **Nunca subas** nada a GitHub sin permiso explícito.
- No toques el contrato de `doPost` (`ping`, `push`, `pull`, `invite`, `revoke`, `rotateCoachToken`).
  Esta función es **aditiva**: funciones nuevas que se ejecutan desde la hoja, no desde las apps.
- El ordenador de Iván es **Windows** (en el repo hay `test-api.ps1`): usa PowerShell para los comandos
  que le pidas ejecutar.
- Al acabar cada fase, marca la checklist (sección 8) y díselo a Iván.

Decisiones ya tomadas con Iván (no las vuelvas a preguntar):

1. El deportista se elige por un **apodo o iniciales** que Iván asigna a cada código, nunca por el nombre
   completo.
2. Los entrenamientos pasan a la app **al pulsar «Publicar»** en un menú de la hoja, no de forma automática.
3. **Una fila por ejercicio** (series × reps × carga). Para series distintas (pirámides), varias filas
   seguidas del mismo ejercicio.

---

## 1. Explicación para Iván

**Para Iván.**

**Qué vas a poder hacer.** En tu hoja «IMDTraining datos» habrá una pestaña nueva, **programacion**,
donde escribes los entrenamientos como en cualquier Excel: una fila por ejercicio, con el deportista, el
día, el nombre de la sesión, series, repeticiones, kilos o %1RM y RPE o RIR. También sesiones aeróbicas,
continuas o por intervalos, y pliometría.

Cuando termines, pulsas **IMDTraining → Publicar** en el menú de la hoja. En unos segundos (lo que tarde
cada móvil en sincronizar) el entrenamiento aparece en la app del deportista, en «Programado por tu
entrenador», y en tu calendario semanal. La columna **Estado** te dice si se publicó bien y, más
adelante, si el deportista ya lo hizo.

**Lo bueno de hacerlo en la hoja:** copiar y pegar semanas enteras, arrastrar fechas, programar a varios
deportistas a la vez y tener toda la planificación a la vista.

**Lo que no cambia:** puedes seguir programando desde la app como hasta ahora.

**Privacidad:** en la hoja seguirá sin haber nombres completos. Tú decides el apodo de cada deportista
(por ejemplo «Dani G.» o «DG»). Te recomiendo iniciales o un apodo, no el nombre y los dos apellidos.

**Cuánto cuesta:** 0 €. **Tu tiempo:** unos 20 minutos para prepararlo y 30 minutos de pruebas.

---

## 2. Cómo será la hoja (lo que verá Iván)

### 2.1 Pestaña `deportistas`

| Columna | Qué es | Quién la rellena |
|---|---|---|
| A `codigo` | El código de cada deportista vinculado (`d7f3a9c2`) | El programa, solo (sale de `athletes`, sin los desvinculados) |
| B `apodo` | Cómo quieres llamarle en la hoja | **Iván** |

Para saber qué código es cada deportista, la app mostrará el código en su perfil (fase 2).

### 2.2 Pestaña `programacion`

Fila 1 = cabecera, congelada. Desde la fila 2, una fila por ejercicio.

| Col | Cabecera | Para fuerza | Para aeróbico | Validación |
|---|---|---|---|---|
| A | Deportista | apodo | apodo | Desplegable con `deportistas!B2:B` |
| B | Fecha | día | día | Fecha válida |
| C | Sesión | título («Pierna A») | título («Rodaje») | Texto |
| D | Tipo | Fuerza | Aeróbico | Desplegable: Fuerza, Aeróbico (vacío = Fuerza) |
| E | Ejercicio | Sentadilla | Actividad: Carrera, Ciclismo… | Texto |
| F | Variante | opcional | — | Texto |
| G | Unilateral | casilla | — | Casilla |
| H | Pliometría | casilla | — | Casilla |
| I | Series | 4 | — | Número entero 1-20 |
| J | Reps | 8 (o contactos si es pliometría) | — | Número entero |
| K | Kg | 80 | — | Número ≥ 0 |
| L | %1RM | 75 | — | Número 1-120 |
| M | RPE | 7 | — | Número 0-10 |
| N | RIR | 2 | — | Número 0-5 |
| O | Intensidad | Baja/Media/Alta (pliometría) | — | Desplegable |
| P | Minutos | — | 40 (continuo) | Número |
| Q | Zona | — | Z2 (continuo) | Desplegable Z1-Z5 |
| R | Intervalos | — | `10 Z2, 3 Z4, 10 Z2` | Texto |
| S | Indicaciones | notas para el deportista | notas | Texto |
| T | Acción | vacío o «Borrar» | vacío o «Borrar» | Desplegable: Borrar |
| U | Estado | lo escribe el programa | lo escribe el programa | — (gris) |
| V | id | lo escribe el programa | lo escribe el programa | Oculta |

Reglas que verá Iván en una nota en la cabecera:

- **Una sesión** = todas las filas con el mismo **Deportista + Fecha + Sesión**. El orden de las filas es
  el orden de los ejercicios.
- **Pirámides:** filas seguidas del mismo ejercicio se juntan en uno (por ejemplo, 1 fila «3 × 8 × 70 kg»
  y otra «2 × 5 × 80 kg» → un ejercicio con 5 series).
- **Kg o %1RM:** si pones %1RM, la app calcula los kilos con el 1RM de cada deportista.
- **Aeróbico:** o Minutos + Zona (continuo) o Intervalos. Los intervalos se escriben como bloques
  «minutos Zona» separados por comas o «+»: `10 Z2, 3 Z4, 10 Z2` o `10'Z2+3'Z4+10'Z2`. Medio minuto =
  `0,5`.
- **Indicaciones:** se usa la primera celda no vacía de la sesión.
- **Cambiar algo ya publicado:** cámbialo en la hoja y vuelve a pulsar Publicar. Se actualiza en la app.
- **Quitar algo ya publicado:** en **Acción** elige «Borrar» y pulsa Publicar. Borrar la fila sin más
  **no** lo quita de la app.

Ejemplo que el programa dejará escrito en las filas 2-5 la primera vez (con el apodo «Ejemplo» y fechas
de la semana siguiente), para que Iván vea el formato:

| Deportista | Fecha | Sesión | Tipo | Ejercicio | Variante | Unil. | Plio | Series | Reps | Kg | %1RM | RPE | RIR | Intens. | Min | Zona | Intervalos | Indicaciones |
|---|---|---|---|---|---|---|---|---|---|---|---|---|---|---|---|---|---|---|
| Ejemplo | 20/10/2026 | Pierna A | Fuerza | Sentadilla | | | | 4 | 6 | | 80 | | 2 | | | | | Calienta 10 min |
| Ejemplo | 20/10/2026 | Pierna A | Fuerza | Zancada | | ✓ | | 3 | 8 | 12 | | 7 | | | | | | |
| Ejemplo | 20/10/2026 | Pierna A | Fuerza | Drop jump | | | ✓ | 3 | 6 | | | | | Alta | | | | |
| Ejemplo | 22/10/2026 | Series | Aeróbico | Carrera | | | | | | | | | | | | | 10 Z2, 3 Z4, 2 Z1, 3 Z4, 10 Z2 | |

(Las filas de ejemplo usan un apodo inexistente, así que al publicar salen como «Error: deportista
desconocido» hasta que Iván las borre o cambie el apodo. Así no se publica nada sin querer.)

### 2.3 Columna Estado

| Valor | Significado |
|---|---|
| `✓ Publicado · Pendiente` / `· En curso` / `· Hecho` | En la app; el estado viene de lo que marque el deportista |
| `Sin cambios` | Ya estaba publicado igual; no se ha tocado |
| `Borrado` | Quitado de la app (la fila puede borrarse ya) |
| `Error: …` | No se ha publicado esa sesión; el texto dice por qué |

Todas las filas de una sesión llevan el mismo estado.

---

## 3. Fase 1 — El programa de la hoja (Claude escribe, Iván pega)

### 3.1 Fichero nuevo `apps-script/Programacion.gs`

**Para Claude.** Va en un **fichero aparte** del proyecto de Apps Script (comparte ámbito global con
`Code.gs`, así que puede usar `sheet_`, `withLock_`, `readAthletes_`, `readMeta_`, `readRecord_`,
`newWriter_`, `newClock_`, `bool_`). **No modifica `Code.gs`.** Las dos pestañas nuevas **no** se añaden a
`TABS`: las crea su propia función, con su formato.

Funciones públicas (sin guion bajo final, para que salgan en el desplegable del editor):

| Función | Para qué |
|---|---|
| `onOpen()` | Menú **IMDTraining** con: *Publicar programación*, *Actualizar estados*, *Preparar pestañas de programación* |
| `prepararProgramacion()` | Crea `deportistas` y `programacion` si faltan (cabecera, anchos, congelado, validaciones, columna V oculta, nota de ayuda en A1, filas de ejemplo solo si la pestaña es nueva). Idempotente |
| `publicarProgramacion()` | Lee la pestaña, valida, escribe planes en `plans` y rellena Estado |
| `actualizarEstados()` | Solo lee `plans` y rellena Estado (no escribe planes) |

Atención: `onOpen` es un *trigger* simple. **No existe ya en `Code.gs`** (comprobado en `f81049f`): si en
algún momento se añade otro `onOpen`, hay que fusionarlos, porque dos funciones con el mismo nombre en el
mismo proyecto se pisan.

### 3.2 Núcleo puro y testeable

Separa la lógica en funciones **puras** (sin `SpreadsheetApp`), para probarlas con Node en el
ordenador de Iván:

```
parseRows_(values, tz)         -> [{row, apodo, fecha:'YYYY-MM-DD', sesion, tipo, ejercicio, variante,
                                    unilateral, plyo, series, reps, kg, pct, rpe, rir, int, minutos, zona,
                                    intervalos, notas, accion, id}]   // filas vacías fuera
groupSessions_(rows)           -> [{key, rows:[...], id (el primero no vacío de V)}]   // key = apodo|fecha|sesion
buildPlan_(group, cloudId, existing, now) -> {ok:true, data, changed} | {ok:false, error}
parseIntervals_(text)          -> [{min, zone}] | null
contentOf_(planData)           -> lo comparable: {title, kind, date, notes, exercises, hr}
```

Mantén `Programacion.gs` como JavaScript compatible con V8 de Apps Script y con Node (sin `import`). Para
las pruebas, `apps-script/test-programacion.js` lo carga con `vm` o `eval` y comprueba los casos de 3.6
(`node apps-script/test-programacion.js`). Comprueba antes que Iván tiene Node (`node -v`); si no, pide
permiso para instalarlo o haz las pruebas en la propia hoja con una función `probarProgramacion_()`.

### 3.3 Conversión fila → plan (el formato que ya entiende la app)

**Para Claude.** El plan tiene la forma que crea `newPlanFor` (`index.html:1991`) y que ya circula por la
sincronización. Verifica los campos con el código antes de escribir (`grep -n "function newPlanFor\|plyo\|load==='pct'\|rir" index.html`).

```
data = {
  id, userId: cloudId,                       // userId = id_unico; la app lo traduce al perfil local (Sync.merge)
  title: sesion, kind: tipo==='Aeróbico' ? 'hr' : 'strength',
  date: 'YYYY-MM-DD', notes,
  status: existing?.status || 'pending',      // se conservan los campos que pone el deportista:
  doneId: existing?.doneId, doneAt: existing?.doneAt,   // status, doneId, doneAt (y cualquier otro no listado en contentOf_)
  createdAt: existing?.createdAt || now,
  src: 'hoja',                                // marca para la app (fase 3)
  exercises: [...],                           // fuerza
  hr: {...},                                  // aeróbico
}
```

Fuerza — cada bloque de filas consecutivas con la misma (ejercicio, variante, unilateral, plyo) es un
ejercicio:

```
{ name, variant, unilateral, plyo,            // plyo solo si true
  load: (alguna fila con %1RM) ? 'pct' : undefined,
  sets: por cada fila, `series` copias de
        { reps, kg: pct ? null : kg, pct, rpe, rir, int } }   // int: Baja=1, Media=2, Alta=3; vacíos = null
```

- Una fila con %1RM y otra con kg dentro del mismo ejercicio → error «mezcla kg y %1RM en <ejercicio>».
- En unilaterales, las series del plan son **por lado** (la app crea I/D al empezar). No duplicar aquí.

Aeróbico — se usa la **primera** fila de la sesión; si hay más filas de tipo Aeróbico en la misma sesión,
error «una sesión aeróbica va en una sola fila»:

```
hr = intervalos ? {type: ejercicio || 'Carrera', mode:'int', blocks: parseIntervals_(intervalos)}
                : {type: ejercicio || 'Carrera', minutes, zone}
```

`parseIntervals_`: separa por `,`, `+` o `;`; en cada trozo, número (con `,` o `.` decimal) seguido de
`'`, `min` o espacio y `Z1`-`Z5`. Cualquier trozo que no encaje → `null` → error «intervalos no válidos:
escribe por ejemplo 10 Z2, 3 Z4».

Mezcla de Fuerza y Aeróbico en una misma sesión → error «una sesión es de fuerza o aeróbica, no las dos».

Validaciones (error en todas las filas de la sesión, y esa sesión no se publica; las demás sí):

| Caso | Mensaje |
|---|---|
| Apodo vacío o que no está en `deportistas` | `Error: deportista desconocido` |
| Apodo duplicado en `deportistas` | `Error: el apodo «X» está repetido en deportistas` |
| Código desvinculado (`revoked`) | `Error: deportista desvinculado` |
| Fecha vacía o no válida | `Error: falta la fecha` |
| Sesión vacía | `Error: falta el nombre de la sesión` |
| Fuerza sin ejercicio, series o reps | `Error: fila N sin ejercicio, series o reps` |
| Aeróbico sin minutos+zona ni intervalos | `Error: pon minutos y zona, o intervalos` |

Fechas: convierte con `Utilities.formatDate(value, tz, 'yyyy-MM-dd')` usando
`SpreadsheetApp.getActive().getSpreadsheetTimeZone()`. **Comprueba** con Iván que la zona horaria de la
hoja es Madrid (*Archivo → Configuración → Zona horaria*, *File → Settings → Time zone*): si no, las
fechas pueden salir un día antes.

### 3.4 Publicar

```
function publicarProgramacion():
  res = withLock_(() => {
    syncDeportistas_()                               // añade a `deportistas` los id_unico nuevos no revocados
    apodos  = leer deportistas -> {apodo -> id_unico}; detectar duplicados
    athletes = readAthletes_()
    meta    = readMeta_('plans')                     // id -> {row, userId, updatedAt, deleted...}
    rows    = parseRows_(programacion.getDataRange().getValues(), tz)
    writer  = newWriter_(); clock = newClock_(); now = Date.now()
    estados = {}                                     // fila -> texto
    for g of groupSessions_(rows):
      if g tiene acción Borrar:
        if g.id && meta[g.id] && !meta[g.id].deleted:
          writeRecord_(writer, meta[g.id], {store:'plans', id:g.id, userId:meta[g.id].userId, updatedAt:now, deleted:true},
                       meta[g.id].author, clock.next(), '')
        estado = 'Borrado'; continuar
      validar -> si error: estado = 'Error: …'; continuar
      cid = apodos[g.apodo]
      id  = g.id || ('h' + Utilities.getUuid().replace(/-/g,'').slice(0,12))
      old = meta[id]
      if old && old.userId !== cid:                  // se ha cambiado el deportista de una sesión ya publicada
        tombstone del antiguo (como Borrar) e id nuevo   // los planes no pueden cambiar de dueño (Code.gs:177)
      existing = old && !old.deleted ? readRecord_(old, {plans:meta, sessions:{}}).data : null
      plan = buildPlan_(g, cid, existing, now)
      if !plan.changed: estado = 'Sin cambios · ' + estadoApp(existing); continuar
      writeRecord_(writer, old, {store:'plans', id, userId:cid, updatedAt:now, deleted:false},
                   old ? old.author : 'coach', clock.next(), JSON.stringify(plan.data))
      estado = '✓ Publicado · ' + estadoApp(plan.data); g.id = id
    clock.save(); writer.flush()
    return {estados, ids}
  })
  if (!res || res.ok === false) -> SpreadsheetApp.getUi().alert('La hoja está ocupada sincronizando. Vuelve a pulsar Publicar en unos segundos.')
  escribir columnas U (Estado) y V (id) en bloque (setValues), fuera del bloqueo
  toast: 'Publicadas N sesiones · M errores'
```

Notas:

- `writeRecord_(writer, old, rec, author, serverAt, data)` es la de `Code.gs:240`: reutilízala tal cual
  para que las filas de `plans` tengan exactamente el formato que lee `pull_`. `author` = `'coach'` en
  planes nuevos.
- `updatedAt = now` hace que la versión de la hoja gane (LWW) a lo que hubiera en los móviles. Por eso se
  conservan `status`, `doneId` y `doneAt` del registro existente leído **dentro del bloqueo**.
- `changed`: compara `JSON.stringify(contentOf_(nuevo))` con el del existente. Sin cambios no se escribe
  nada: así Publicar se puede pulsar todas las veces que se quiera sin pisar lo que marque el deportista.
- El id del plan se guarda en la columna V de **todas** las filas de la sesión.
- `estadoApp(d)`: `pending` → Pendiente, `inprogress` → En curso, `done` → Hecho.

### 3.5 Actualizar estados

`actualizarEstados()`: con el bloqueo, lee `plans` (solo los ids presentes en la columna V) y reescribe
Estado (`✓ Publicado · Hecho`, etc.). Sin escribir en `plans`.

### 3.6 Casos de prueba del núcleo (Node)

| # | Entrada | Esperado |
|---|---|---|
| N1 | 2 filas «Sentadilla» seguidas: 3×8×70 y 2×5×80 | 1 ejercicio, 5 series en ese orden |
| N2 | Fila con %1RM 75 | `load:'pct'`, sets con `pct:75`, `kg:null` |
| N3 | Mismo ejercicio con kg en una fila y %1RM en otra | error de mezcla |
| N4 | Pliometría, Intensidad Alta | `plyo:true`, sets con `int:3` |
| N5 | Intervalos `10 Z2, 3 Z4, 0,5 Z5` y `10'Z2+3'Z4` | bloques correctos; `0,5` = 0.5 |
| N6 | Intervalos `10 Z7` | `null` → error |
| N7 | Mismo grupo, se cambia un kg | `changed:true`; se conserva `status:'done'` del existente |
| N8 | Mismo grupo sin cambios | `changed:false` |
| N9 | Sesión con Fuerza y Aeróbico | error |
| N10 | Fila vacía entre medias | ignorada |

---

## 4. Fase 2 — Cambios pequeños en la app (`index.html`)

**Para Claude.** Solo tres cambios, todos aditivos:

1. **Mostrar el código del deportista** en su tarjeta de sincronización (`shareCard`, rama
   `Sync.linked(p)`, `index.html:1617`), en modo entrenador: «Código en la hoja: **d7f3a9c2**», con
   botón para copiarlo (`navigator.clipboard.writeText`, con `toast`). Es lo que Iván necesita para
   rellenar los apodos.
2. **Planes que vienen de la hoja** (`data.src === 'hoja'`): en el editor del entrenador (`V.plan`), un
   aviso arriba: «Programado desde la hoja de Google. Si lo cambias aquí, la próxima vez que pulses
   Publicar en la hoja se sustituirá por lo de la hoja.» No bloquees la edición.
3. En la vista del deportista no hace falta nada: ya muestra los planes del entrenador.

Verifica en dos navegadores (normal = entrenador, incógnito = deportista) como en el plan anterior, sin
tocar la sincronización.

---

## 5. Fase 3 — Montarlo en Google (Iván, guiado)

### Paso 3.1 — Copia de seguridad (Iván)
Como siempre: app → menú → Copia de seguridad → Exportar copia → Drive. Y en la hoja de Google:
*Archivo → Crear una copia* (*File → Make a copy*), por si acaso.

### Paso 3.2 — Añadir el fichero al proyecto (Iván)
- **Qué:** crear un segundo fichero en el editor de Apps Script y pegar el programa nuevo.
- **Qué haces tú:** en la hoja, **Extensiones → Apps Script** (*Extensions → Apps Script*) → a la
  izquierda, junto a **Archivos** (*Files*), pulsa **+** → **Secuencia de comandos** (*Script*) → llámalo
  `Programacion` → borra lo que traiga y pega el contenido que te doy → **Guardar proyecto** (Ctrl + S).
- **No hace falta** volver a implementar: lo nuevo se ejecuta desde la hoja, no desde las apps. La
  dirección `/exec` no cambia.
- **Cómo sabes que ha ido bien:** a la izquierda ves `Código.gs` y `Programacion.gs`; al guardar no sale
  nada en rojo.

### Paso 3.3 — Preparar las pestañas (Iván)
- En el desplegable de funciones elige `prepararProgramacion` → **Ejecutar**. Si pide permisos, los mismos
  pasos que la primera vez (Revisar permisos → tu cuenta → Configuración avanzada → Ir a… → Permitir).
- **Cómo sabes que ha ido bien:** en la hoja hay pestañas nuevas `deportistas` y `programacion`, y al
  recargar la hoja (F5) aparece arriba un menú **IMDTraining**.

### Paso 3.4 — Poner apodos (Iván)
- En la app, abre el perfil de cada deportista vinculado → pestaña Datos → copia su **código en la hoja**.
- En la pestaña `deportistas`, junto a ese código, escribe su apodo.

### Paso 3.5 — Primera publicación (Iván)
- En `programacion`, cambia el apodo de las filas de ejemplo por el de un deportista de prueba (o tu
  propio perfil vinculado) y pon fechas de esta semana.
- **IMDTraining → Publicar programación**.
- **Cómo sabes que ha ido bien:** la columna Estado pone `✓ Publicado · Pendiente`; en unos segundos (o
  al abrir la app) el entrenamiento aparece en tu calendario semanal y en la app del deportista.

---

## 6. Fase 4 — Pruebas

| # | Qué se hace | Qué se tiene que ver |
|---|---|---|
| H1 | Publicar una sesión de fuerza con kg, %1RM, RIR y una unilateral | En la app del deportista, mismo contenido; el %1RM convertido a kg con su 1RM |
| H2 | Publicar una de pliometría y una aeróbica por intervalos | Contactos e intensidad; bloques de intervalos en la sesión en directo |
| H3 | Pulsar Publicar otra vez sin cambiar nada | Estado «Sin cambios»; nada se mueve en la app |
| H4 | Cambiar los kg y publicar | Se actualiza en la app del deportista |
| H5 | El deportista hace el entreno; Iván pulsa *Actualizar estados* | Estado «✓ Publicado · Hecho» |
| H6 | Tras H5, Iván cambia una indicación y publica | El plan sigue «Hecho» en la app (no vuelve a Pendiente) |
| H7 | Acción «Borrar» + Publicar | Desaparece de la app del deportista; Estado «Borrado» |
| H8 | Apodo inexistente, fecha vacía, intervalos mal escritos | Estado «Error: …» en esas filas; el resto se publica |
| H9 | Cambiar el deportista de una sesión ya publicada | Desaparece del primero y aparece en el segundo |
| H10 | Buscar en la hoja (Ctrl + H) el nombre completo de un deportista | Ninguna coincidencia fuera de lo que Iván haya escrito como apodo |
| H11 | Publicar mientras un móvil sincroniza | O se publica, o sale el aviso «La hoja está ocupada»; nunca datos a medias |

---

## 7. Publicar la nueva versión de la app

Igual que la fase 5 del plan anterior (permiso explícito; `git push` si hay credenciales, si no subida por
la web de GitHub). Sube también `apps-script/Programacion.gs` y el test.

## 8. Checklist

### Fase 1 — Programa
- [ ] `Programacion.gs` con núcleo puro y funciones de hoja
- [ ] `test-programacion.js` pasa N1-N10

### Fase 2 — App
- [ ] Código del deportista visible y copiable en su perfil
- [ ] Aviso en planes con `src:'hoja'`

### Fase 3 — Google
- [ ] Copia de la app y de la hoja
- [ ] `Programacion.gs` añadido y guardado
- [ ] `prepararProgramacion` ejecutado; menú IMDTraining visible
- [ ] Zona horaria de la hoja: Madrid
- [ ] Apodos puestos
- [ ] Primera publicación correcta

### Fase 4 — Pruebas
- [ ] H1 · [ ] H2 · [ ] H3 · [ ] H4 · [ ] H5 · [ ] H6 · [ ] H7 · [ ] H8 · [ ] H9 · [ ] H10 · [ ] H11

### Fase 5 — Publicar
- [ ] Permiso de Iván · [ ] Subido a GitHub · [ ] Visto en el móvil

## Anexo — Fuera de alcance (no lo hagas aunque parezca buena idea)

- Publicar automáticamente con temporizador o al editar (Iván eligió el botón).
- Copiar a la hoja los planes creados desde la app (la dirección es solo hoja → app).
- Poner nombres completos en la hoja.
- Tocar `doPost`, el contrato de sincronización o el formato de `plans` más allá del campo `src`.
- Borrar filas de `plans` o de `programacion`.
