# Monitor Retail LATAM: actualización automática

Proyecto en Python que actualiza `excel/Monitor_Retail_LATAM.xlsx`. En cada corrida:

1. Descarga cada indicador de su fuente oficial (bancos centrales, institutos de estadística y BIS).
2. Lo transforma a la unidad del Excel (variación a 12 meses o índice 0-100, normalizado si hace falta).
3. **Lo valida contra el histórico del propio Excel.** Si no cuadra, no escribe nada y deja el motivo en el log.
4. Escribe los datos y rueda el heatmap a los últimos meses. También extiende los gráficos y actualiza promedios y rangos con nombre, **sin perder el formato**.
5. Arma la hoja **Noticias**, con titulares recientes por país y una lectura de qué se movió y por qué, y la hoja **Log_Actualizacion**, con la trazabilidad.

Hojas que se actualizan: `1_Actividad_Economica`, `2_Ventas_Retail`, `3_Confianza_Empresarial`,
`4_Confianza_Consumidor`, `5_Inflacion`, `6_ANTAD` y `7_TPM's`. Las demás hojas no se tocan.

---

## Inicio rápido (Windows + VS Code)

**Requisito único:** Python 3.9 o superior instalado. En equipos corporativos suele estar en el Centro de Software; si no, está en python.org. Al instalarlo marca *Add python.exe to PATH*.

1. **Descarga el repo:** en GitHub, botón *Code → Download ZIP*, y descomprímelo en una carpeta local. También sirve `git clone`.
2. **Ábrelo en VS Code:** *File → Open Folder…* Si VS Code sugiere instalar la extensión de Python, acepta.
3. **Instala el entorno:** *Terminal → Run Task… → "1. Instalar / reparar entorno"*. También puedes dar doble clic a `setup.bat`.
   Esto crea `.venv`, instala `requirements.txt` y crea el archivo `.env`. Se hace una sola vez.
4. **Tokens gratuitos (2 minutos):** abre `.env` y pega:
   - `INEGI_TOKEN`: [registro INEGI](https://www.inegi.org.mx/app/desarrolladores/generatoken/Usuarios/token_Verify). Lo usan IGAE, ventas minoristas y confianza de México.
   - `BANXICO_TOKEN`: [token Banxico SIE](https://www.banxico.org.mx/SieAPIRest/service/v1/token). Lo usan el INPC y la tasa objetivo; si falta, se usa el BIS.
   - `BCCH_TOKEN` es opcional: sin él, Chile sale de las tablas públicas del BCCh.
5. **Corre:** `Ctrl+Shift+B`, que equivale a *Actualizar Monitor*, o doble clic a `run.bat`.
6. Abre el Excel y mira la hoja **Log_Actualizacion**.

> La primera vez, Windows puede mostrar *"Windows protegió su PC"* al abrir un `.bat` descargado. Usa *Más información → Ejecutar de todas formas*, o corre todo desde las *Tasks* de VS Code.

**Cada mes:** cierra el Excel y corre *Actualizar Monitor*. No hay que pegar nada a mano.

| Tarea de VS Code | Qué hace |
|---|---|
| Actualizar Monitor | Corrida normal. Antes de escribir hace un respaldo del Excel en `output/respaldos/` |
| Vista previa | Igual, pero guarda `output/reportes/PREVIEW_….xlsx` y **no toca** el Excel maestro |
| Actualizar solo datos | Sin la hoja Noticias (más rápido) |
| Diagnóstico de fuentes | Muestra la estructura de los archivos que no se pudieron leer |
| Buscar id INEGI | Busca la clave de un indicador del BIE por texto |
| Tests | Pruebas automáticas del proyecto |

Desde una terminal equivale a `run.bat`, `run.bat --prueba`, `run.bat --sin-noticias` o `run.bat --solo 5_Inflacion`.

---

## De dónde sale cada dato

Cada serie puede tener varias fuentes. Se descargan todas, se validan y gana la válida con el dato más reciente; en caso de empate, la más parecida al histórico. Todo está en `config/series.yaml`.

| Hoja | Brasil | Chile | Colombia | México | Perú | EE.UU. |
|---|---|---|---|---|---|---|
| 1 Actividad (a/a, sin desestacionalizar) | IBC-Br, BCB SGS 24363 · API | IMACEC, tabla BCCh · web | ISE serie original: anexo DANE · archivo; respaldo boletín PDF | IGAE cifras originales, INEGI · token | PBI mensual, BCRP PN01728AM · API | |
| 2 Ventas retail (a/a) | PMC volumen, IBGE SIDRA t8880 · API | IAC por menor, INE · archivo | EMC total con vehículos, DANE · archivo; respaldo boletín PDF | EMEC sector 46 cifras originales, INEGI · token | | |
| 3 Confianza empresarial (0-100) | ICEI Indústria Geral, CNI · archivo; respaldos: texto de la página CNI, resumen PDF mensual, SGS 7341 | IMCE Comercio, BCCh · web | ICCO Fedesarrollo (X+100)/2 · archivo o titulares | ICE manufacturas cifras originales, INEGI · token | Expectativas a 12m, BCRP PD39752AM · API | |
| 4 Confianza consumidor (0-100) | ICC, BCB SGS 4393 ÷2 · API | ICC Ipsos, BCCh · web | ICC Fedesarrollo (X+100)/2 · archivo o titulares | ICC, INEGI 454186 · token; respaldo boletín PDF sin token | | |
| 5 Inflación (a/a) | IPCA 12m, SGS 13522 · API; respaldo BIS | IPC, BCCh · web; respaldo BIS | IPC, DANE · archivo; respaldo BIS | INPC, Banxico SP1 · token; respaldo BIS | IPC Lima, BCRP PN01273PM · API; respaldo BIS | CPI-U NSA, FRED · API; respaldo BIS |
| 6 ANTAD (tiendas iguales) | | | | Total: titulares de prensa (naranja); segmentos: manual | | |
| 7 TPM (fin de mes) | Selic meta, SGS 432; respaldo BIS | BIS; opcional API BCCh | BIS | Banxico SF61745; respaldo BIS | BCRP PD04722MM; respaldo BIS | FEDFUNDS, FRED |

**Nivel de automatización:**
- *API* y *web BCCh* son estables.
- *Archivo* significa que el código busca el enlace del Excel del mes en la página del instituto y lo lee con un lector flexible. Si el instituto cambia el formato, la validación lo detecta y la serie queda en "ERROR" en el log, sin dañar el Excel.
- *Boletín PDF* (tipo `texto_web`): cuando la API o el anexo fallan, el código descarga los boletines mensuales del instituto (URL predecible por mes), lee el texto y toma el dato de la frase clave ("se ubicó en 45.1 puntos", "serie original … 1,11 %"). Lee varios meses hacia atrás, así que rellena huecos solo, y guarda cada PDF para no descargarlo de nuevo. Se pinta naranja.
- *Titulares* es un respaldo: el dato se extrae de la prensa ("confianza comercial bajó a 23,3"), se pinta naranja y hay que verificarlo.
- ANTAD no publica datos descargables (su web los carga con JavaScript), así que los segmentos van por `data/manual_inputs.csv`.

**Colores en las tablas de datos:**
- **Verde:** dato oficial descargado.
- **Naranja:** manual sin verificar o tomado de prensa.
- **Amarillo:** mes sin publicar.

Los colores originales (azul PPTX, púrpura visual) se conservan en los meses que ninguna fuente oficial cubre.

---

## Por qué el formato no se pierde

El código no reescribe el libro (lo que sí hace openpyxl, que descarta partes del formato de los gráficos). Edita **directamente el XML del .xlsx**:

- Solo toca las celdas de datos, las del heatmap y los rangos y cachés de las series de los gráficos.
- Estilos, colores de series, ejes, formatos condicionales, notas, comentarios y las hojas `PIB_Corp_*` (con fórmulas de Capital IQ) se copian tal cual.
- Si llega un mes nuevo, inserta la fila al final de la tabla y desplaza las notas, los merges y los formatos condicionales, igual que haría Excel.
- Marca el libro para recálculo completo al abrir.
- Es un `.xlsx` normal, sin macros, así que no lo bloquean las políticas de seguridad de Office.
- Los tests (`tests/test_excel.py`) verifican que los gráficos quedan idénticos salvo el rango de datos.

**Rangos con nombre** (*Fórmulas → Administrador de nombres*), recalculados en cada corrida:
- `ActEco_Datos`, `ActEco_Heatmap`, `ActEco_Fechas` y `ActEco_BRASIL`…
- `Ventas_*`, `ConfEmp_*`, `ConfCons_*`, `Infl_*`, `ANTAD_*`, `TPM_*`.

Sirven para tus propias fórmulas o gráficos.

---

## Trazabilidad

| Qué | Dónde |
|---|---|
| Estado de cada serie: fuente usada, último mes, validación e intentos fallidos | hoja `Log_Actualizacion` y `output/reportes/estado_AAAAMMDD_HHMM.csv` |
| Cada celda que cambió, con valor anterior y nuevo, si es dato nuevo o revisión, y su fuente | hoja `Log_Actualizacion` y `output/reportes/cambios_….csv` |
| Copia exacta de cada respuesta descargada (JSON, CSV, HTML, XLSX) con URL, hora y hash | `data/raw/AAAA-MM-DD/` y su `_index.jsonl` (los tokens se ocultan) |
| Series limpias en CSV, una por hoja | `data/store/` |
| Excel antes de cada corrida | `output/respaldos/` |
| Log técnico completo | `logs/corrida_….log` |

Con `run.bat --offline` se reconstruye el Excel desde la última descarga guardada, sin internet. Sirve para auditar una corrida pasada.

---

## Hoja Noticias

Por país muestra:
- Una tabla con el último dato de cada indicador, el mes previo, el cambio m/m y el dato de hace 12 meses, más una lectura automática ("sube/baja", "zona pesimista <50", rezago).
- Los titulares de los últimos 45 días que explican el movimiento, tomados de Google News RSS (sin API key) y de feeds oficiales configurables en `config/noticias.yaml`.
- Una columna **"Por qué"**, que extrae la causa del titular ("impulsada por…", "debido a…", "ante…").
- Un **párrafo de síntesis**, armado con reglas. Si pones `ANTHROPIC_API_KEY` en `.env`, Claude lo redacta usando solo esos datos y titulares. Es opcional; sin la clave todo funciona.

Para afinar las búsquedas, edita `consulta` y `filtro` en `config/noticias.yaml`.

---

## Datos manuales

Si una fuente falla, o para los segmentos de ANTAD, agrega filas a `data/manual_inputs.csv`:

```
hoja,columna,mes,valor,unidad,fuente,verificado,nota
6_ANTAD,AUTOSERVICIO,2026-09,1.8,pct,Comunicado ANTAD 9-oct-2026,si,
3_Confianza_Empresarial,COLOMBIA,2026-09,62.1,nivel,Fedesarrollo EOE (ya normalizado),si,
```

- Unidad `pct` significa que el valor viene en %; `nivel` es para índices.
- El dato manual tiene prioridad sobre el automático para ese mes.
- `verificado=si` pinta verde y `no` pinta naranja.

---

## Agregar o cambiar una serie

Todo se configura en `config/series.yaml`, sin tocar código. Ejemplo:

```yaml
MEXICO:
  nombre: Indicador de Confianza Empresarial (INEGI)
  fuentes:
    - {tipo: inegi, buscar: "indicador de confianza empresarial comercio",
       titulo_regex: "confianza empresarial.*comercio", transform: nivel, requiere: INEGI_TOKEN}
```

- Tipos disponibles: `bcb_sgs`, `ibge_sidra`, `bcrp`, `bis`, `fred`, `banxico`, `inegi`, `bcch_siete`, `bcch_api`, `archivo_web`, `texto_web`, `noticia_dato`.
- Transformaciones: `yoy`, `pct`, `auto_yoy`, `nivel`, `"formula:(x+100)/2"`.
- Tolerancias de validación por hoja: `validacion: {mad_max, rango}`.

---

## Solución de problemas

- **"El Excel está abierto":** ciérralo. Si no, la versión nueva se guarda como `Monitor_Retail_LATAM_actualizado_….xlsx`.
- **Errores SSL o de certificado en la red corporativa:** el código ya usa los certificados de Windows (`truststore`). Si tu red usa proxy, define `HTTPS_PROXY` en `.env` o en `config/settings.yaml`.
- **Una serie en "ERROR":** la columna *Detalle / intentos* del log dice por qué: token faltante, archivo no encontrado, o "no coincide con el histórico", que suele indicar que la fuente cambió de definición. Corre *Diagnóstico de fuentes* para ver la estructura del archivo descargado.
- **INEGI cambia una clave:** usa la tarea *Buscar id INEGI* y pega el id en `series.yaml`.

---

## Estructura

```
config/            series.yaml (fuentes) · noticias.yaml · settings.yaml
excel/             Monitor_Retail_LATAM.xlsx (maestro) · ..._original.xlsx (archivo recibido, referencia)
data/              manual_inputs.csv · raw/ (descargas) · store/ (CSV limpios)
src/monitor/       código: sources/ (una por fuente) · excel/ (edición XML) · news.py · pipeline.py
tests/             pruebas + respuestas reales de las APIs (fixtures)
output/, logs/     respaldos, reportes y logs de cada corrida
setup.bat, run.bat · .vscode/ (tareas, depuración, intérprete)
```

Al actualizar el código (un ZIP nuevo del repo), conserva tu carpeta `excel/`, tu `data/manual_inputs.csv` y tu `.env`.
