# job-radar

Radar personal de búsqueda laboral para perfiles junior / trainee de backend, IA y automatización. Corre solo en tu PC (Windows con Task Scheduler, o Linux/macOS con cron) y:

- **baja vacantes nuevas** de 10 bolsas con API o RSS (Get on Board, Himalayas, We Work Remotely, Remote OK, Remotive y otras), filtra las que encajan con tu perfil y desde dónde aceptan candidatos, y las agrega a un **tracker en Excel** sin pisar lo que ya cargaste;
- crea una **nota en Obsidian por vacante**, con una checklist de preparación;
- te avisa por **Telegram** de las vacantes nuevas, de un resumen semanal y de cualquier **fuente que deje de andar**;
- arma un **radar de entrevistas**: lee foros (HN, dev.to, n8n, Reddit), un LLM (Groq, free tier) extrae preguntas técnicas, live coding y proyectos que se valoran, y lo cruza con lo que piden tus vacantes.

Es un proyecto personal: no hay servidor ni base de datos, los datos son tuyos y quedan en tu máquina.

## Empezar de cero

Requisitos: Python 3.11 o superior y, si querés las notas, [Obsidian](https://obsidian.md).

```bat
git clone <este repo> job-radar
cd job-radar
pip install -r requirements.txt

python radar.py setup                     :: crea el tracker, el .env y el vault; te dice qué falta completar
python scripts\fetch_jobs.py --dry-run    :: muestra qué agregaría, sin escribir nada
python radar.py run                       :: la cadena completa
python radar.py setup --schedule          :: que corra solo (tareas de Windows; en Linux/macOS imprime el cron)
```

Antes de la primera corrida real, ajustá `config/sources.yaml` a tu perfil:

- `filter.home`: dónde podés ir presencial o híbrido (por ejemplo `['mendoza']`). Vacío = solo remoto o avisos que no aclaran modalidad.
- `filter.remote_from`: desde dónde postulás a remotos (por defecto Argentina, LATAM y Worldwide).
- `filter.tech` / `filter.exclude`: qué títulos cuentan como tu perfil y cuáles no.
- `sources`: qué fuentes y consultas usar; cada una se puede desactivar con `enabled: false`.
- `http.user_agent`: poné tu repo o un contacto; los sitios bloquean scrapers anónimos.

`setup` nunca pisa nada que ya exista: ni el tracker, ni el `.env`, ni tus notas. Si ya tenés tu propio tracker, usalo con las mismas 13 columnas de la hoja **Vacantes**.

## Uso

```bat
python radar.py run          :: fetch → Excel → notas → Telegram
python radar.py weekly       :: resumen semanal por Telegram + radar de entrevistas
python radar.py interviews   :: solo el radar de entrevistas
python radar.py status       :: tracker, última corrida, salud de cada fuente, qué falta configurar
```

`--log` (en `run`, `weekly` e `interviews`) manda la salida a `data\radar.log`, que rota a `radar.log.1` al pasar 1 MB; es lo que usan las tareas programadas. `status` nunca muestra tokens ni claves: solo dice si están.

Si el tracker está abierto en Excel cuando corre `run`, no se agrega nada, te llega un aviso por Telegram para que lo cierres y las mismas vacantes entran en la corrida siguiente.

## Fuentes

| Fuente | Cómo se lee |
|---|---|
| n8n Community — Jobs | RSS de Discourse (`/c/jobs/13.rss`) |
| python.org.ar/trabajo | HTML de la lista |
| Get on Board | API pública v0 (JSON) |
| Remote OK | API pública (JSON), por tags |
| Becas IT | HTML de la página de pasantías |
| Selección manual | `data/curated_jobs.json`: avisos que cargás vos |
| Himalayas | API de búsqueda (JSON), 5 consultas × 2 páginas |
| Jobicy | API (JSON), tag `python`; pide citar la fuente y no consultar más de una vez por hora |
| Remotive | API (JSON), categoría software; máx. 4 consultas por día, avisos con 24 h de demora |
| Working Nomads | API (JSON) |
| We Work Remotely | RSS de backend y full-stack |

Las bolsas remotas tienen `location_filter: true`: si el aviso dice desde qué países acepta candidatos, se guarda solo si alguno matchea `filter.remote_from`. Si no dice nada, se guarda.

### Sumar una bolsa sin escribir código

Si la bolsa tiene RSS, alcanza con `type: rss` y su `url` (o `urls`, para varios feeds). Si tiene una API JSON, `type: json` se configura entero en `sources.yaml`:

```yaml
- name: Mi bolsa
  type: json
  url: https://ejemplo.com/api/jobs
  params: {category: dev}            # fijos en cada pedido
  queries: ["python", "backend"]     # opcional: un pedido por consulta...
  query_param: q                     # ...en este parámetro
  page_param: page                   # paginación por número (o cursor_field + cursor_param)
  pages: 2                           # tope de páginas por consulta; una página vacía corta antes
  items: data.jobs                   # dónde está la lista (omitir si la respuesta es la lista)
  modalidad: Remoto
  location_filter: true
  fields: {titulo: title, empresa: company, link: url, resumen: description,
           fecha: published_at, ubicacion: countries, nivel: seniority}
```

Las fechas pueden venir como epoch, ISO 8601 o RFC 822; las listas se unen con comas. El aviso de fuente rota también cubre a las fuentes nuevas.

Sin API ni RSS (verificado 2026-09): Computrabajo, Indeed, Talent, Laborum, WeRemoto, Torre, Jooble sin key: usá sus alertas por email. LinkedIn no se scrapea a propósito.

## Pasos sueltos

```bat
python scripts\fetch_jobs.py               :: deja las nuevas en data\pending_jobs.json
python scripts\update_excel.py             :: las agrega al Excel y las marca como vistas
python scripts\update_excel.py --manual    :: cargar una vacante a mano (por ejemplo, de una alerta de LinkedIn)
```

Los `.bat` que usan las tareas programadas llaman a `radar.py` con el lanzador `py` o `python` del PATH; para forzar otro (por ejemplo, el de un virtualenv), definí la variable `PY` con la ruta a `python.exe`.

## Cómo decide qué agregar

1. **Relevancia**: el **título** tiene que matchear un patrón de `filter.tech` (backend, python, n8n, IA, developer…), nada de título+cuerpo puede matchear `filter.exclude` (senior, lead, diseño, ventas, `[For Hire]`…), el nivel no puede estar en `exclude_levels` (Senior, Experto) y no puede ser presencial o híbrido fuera de `filter.home`. Si no aclara la modalidad, se guarda. "Junior" solo no alcanza: los portales generalistas llenarían el tracker de "Junior Designer".
   Las **pasantías** siguen otra regla (`filter.internship`): cuentan como tales si el título lo dice o Get on Board las marca así. Solo el título puede excluirlas, la señal de software puede venir del cuerpo y rige la misma regla de ubicación. Quedan con `Nivel = Pasantía` y salen primero en el aviso de Telegram.
2. **Deduplicación**: una vacante se considera vista si su **link** ya está en `data/seen_jobs.json` o en el Excel, o si coincide **título + empresa + fuente** con una fila existente (para las cargadas a mano con el link genérico de la lista).
3. **Escritura**: agrega al final de la hoja Vacantes copiando el estilo de la fila 2 (si tenés una nota de ayuda combinada A:M debajo de los datos, la corre hacia abajo). Por defecto `Estado publicación = Abierta` y `Estado postulación = No postulado`. Nunca modifica ni borra filas; la hoja Resumen se recalcula sola al abrir el archivo.
4. **Seguridad**: lo que viene de afuera se guarda como texto; un título que empieza con `=`, `+`, `-` o `@` nunca se ejecuta como fórmula en Excel.

Si una fuente falla (cambió el HTML, está caída), se informa en el log y las demás siguen.

## Selección manual

Lo que aparece en sitios que el radar no lee se agrega a `data/curated_jobs.json`, una lista JSON con los campos de la vacante (ver `curated_jobs.example.json`). Esos avisos no pasan por el filtro, pero sí por la deduplicación, y llegan al Excel y a Telegram en la próxima corrida. `estado_publicacion`, `estado_postulacion` y `notas` son opcionales y se copian tal cual.

## Notas por vacante

`python scripts\write_job_notes.py` crea una nota corta por cada vacante que agregó la última corrida en `obsidian_vault\Entrevistas-IA-Backend\Vacantes\<empresa>-<aaaa-mm-dd>.md`: frontmatter con los datos del tracker, una checklist de preparación, una sección de apuntes y el link a `[[00 Índice]]`. Corre antes de `notify.py`; si falla, el aviso por Telegram sale igual.

- Nunca pisa una nota: si el nombre ya existe (dos vacantes de la misma empresa el mismo día, o una nota tuya), usa `-2`, `-3`…
- Cada nota guarda la `clave` de la vacante; correrlo dos veces no duplica nada.
- Si la empresa no se conoce, usa el título; los caracteres que Windows no acepta en nombres de archivo se reemplazan.
- `--dry-run` muestra los nombres sin escribir.
- `--backfill AAAA-MM-DD` arma las notas que falten para las filas del tracker registradas desde esa fecha, con la fecha de registro de cada fila y su columna Notas como resumen.

## Radar de entrevistas

Busca en foros qué preguntan, qué live coding o take-home piden y qué proyectos valoran, lo clasifica con un LLM y lo escribe en el vault.

```
fetch_interviews.py   HN (posts y comentarios), dev.to, foro de n8n, Reddit (RSS) → data\interview_posts.json
extract_interviews.py Groq → ítems {tipo, sección, texto, detalle, fuentes, veces} → data\interview_items.json
update_obsidian.py    → Radar\Preguntas técnicas · Live coding · Proyectos más pedidos · Comportamental
```

- `veces` = en cuántos posts distintos apareció el mismo ítem; ordena cada sección y marca con ⭐ los repetidos.
- "Proyectos más pedidos" suma, antes del ranking de foros, **lo que piden las vacantes de tu tracker**: las tecnologías de `stack_vocabulary` contadas por vacante (según título y columna Stack), las combinaciones que se repiten en 2 o más vacantes y las frases de la columna Notas que mencionan GitHub, portfolio o proyectos propios. No usa el LLM.
- Cada ítem lleva `idioma` (`es`, `en-bajo` = alcanzó inglés básico, `en`). Las notas muestran solo `obsidian.idiomas` de `config/interviews.yaml`; con `idiomas: []` se ve todo.
- Las notas de `Radar/` se regeneran completas en cada corrida: no las edites. Tus notas propias no se tocan.
- Necesita `GROQ_API_KEY` en `.env` (gratis en console.groq.com). Tope de 40 posts por corrida (`llm.max_posts_per_run`), unos 20k tokens, para no salir del free tier; la cola se drena en corridas sucesivas.
- Si Groq falla, las notas se regeneran igual con los ítems ya guardados.
- X/Twitter y LinkedIn no se consultan: API paga o scraping prohibido. Reddit va por RSS público y a veces devuelve 429; esa consulta se reintenta la próxima vez.

Corre los viernes dentro de `radar.py weekly`; a mano, `python radar.py interviews`.

## Aviso por Telegram

1. En Telegram, hablale a **@BotFather** → `/newbot` → te da el token.
2. Mandale cualquier mensaje a tu bot nuevo y abrí `https://api.telegram.org/bot<TOKEN>/getUpdates`: el `chat.id` que aparece es tu `TELEGRAM_CHAT_ID`.
3. Completá las dos variables en `.env`.
4. `python scripts\notify.py --test` tiene que llegarte al celular.

Sin `.env` el radar sigue actualizando el Excel; solo salta el aviso (queda en el log).

`notify.py` manda un mensaje por corrida con las vacantes recién agregadas (nada si no hubo), y `--weekly` un resumen de lo registrado en los últimos 7 días.

### Aviso por fuente rota

`fetch_jobs.py` anota en `data/source_health.json`, por fuente, cuántas corridas seguidas lleva sin traer avisos (0 avisos o error), el último error y la última vez que anduvo. Cuenta lo que la fuente trajo antes de filtrar: 0 relevantes es una semana floja, 0 traídos es una página que cambió o una API caída.

- Cuando una fuente llega a `health.max_fallas` corridas seguidas (`config/sources.yaml`, 3 = una semana), `notify.py` manda **un** aviso con el error; no se repite mientras siga rota.
- Cuando vuelve a traer avisos, avisa que volvió.
- Si Telegram falla, el aviso no se marca como enviado y sale en la corrida siguiente.
- `Selección manual` no cuenta (vacía es normal), ni las fuentes desactivadas.

## Tareas programadas

`python radar.py setup --schedule` crea estas dos tareas en Windows; en Linux/macOS imprime las líneas equivalentes para `crontab -e`.

| Tarea | Cuándo | Corre |
|---|---|---|
| `job-radar` | lun, mié y vie 09:00 | `run_radar.bat` |
| `job-radar-weekly` | vie 18:00 | `run_weekly.bat` (resumen por Telegram + radar de entrevistas) |

Probar a mano: `schtasks /Run /TN job-radar` y mirar `data\radar.log`. Si la PC está apagada a esa hora, Task Scheduler la corre al próximo inicio solo si activás "Ejecutar la tarea lo antes posible después de un inicio programado perdido" en las propiedades de la tarea.

## Tests

Tests end-to-end: corren la cadena real (fetch → Excel → notas → Telegram) contra un tracker y un vault temporales, con las respuestas HTTP y Telegram simuladas. No tocan tus datos ni la red.

```bat
python -m unittest discover -s tests -t .
```

## Estructura

```
config/sources.yaml            fuentes, filtro, vocabulario de stack, umbral de fuentes rotas
config/interviews.yaml         radar de entrevistas: fuentes, LLM, secciones, idiomas
radar.py                       punto de entrada: setup, run, weekly, interviews, status
scripts/common.py              modelo de vacante, filtro, dedup, helpers
scripts/create_tracker.py      crea el tracker vacío (lo usa radar.py setup)
scripts/fetch_jobs.py          un fetcher por tipo de fuente
scripts/update_excel.py        escritura al tracker + carga manual
scripts/write_job_notes.py     una nota por vacante en el vault
scripts/notify.py              Telegram: vacantes nuevas, resumen semanal, fuentes rotas
scripts/fetch_interviews.py    posts de foros sobre entrevistas
scripts/extract_interviews.py  Groq extrae preguntas / live coding / proyectos en JSON
scripts/update_obsidian.py     regenera las notas de Radar/
run_*.bat, find_python.bat     envoltorios para Task Scheduler (llaman a radar.py --log)
data/                          estado local (vistas, cola, salud de fuentes, log); no se versiona
obsidian_vault/                vault; solo se escriben Radar/ (se regenera) y Vacantes/ (solo agrega)
```
