# Creacion Dataset VirusShare

Proyecto para descargar reportes de VirusShare y analizarlos como dataset local.
El flujo actual tiene dos pasos:

1. Descargar reportes JSON con `descarga_reportes_virusshare.py`.
2. Analizarlos con `analisis_de_reportes.py`, que genera un Excel y una base SQLite.

## Instalacion

```powershell
python -m venv .venv
.\.venv\Scripts\Activate.ps1
python -m pip install -r requirements.txt
```

Dependencias principales:

- `requests` y `tqdm` para descargar reportes.
- `openpyxl` para crear el Excel consolidado.
- `sqlite3` viene incluido con Python.

## Descargar reportes

Configura una API key de VirusShare:

```powershell
$env:VIRUSSHARE_API_KEY = "tu_api_key"
```

O varias keys:

```powershell
$env:VIRUSSHARE_API_KEYS = "key1,key2,key3"
```

Ejecutar con el lote por defecto:

```powershell
python descarga_reportes_virusshare.py
```

Ejecutar con un TXT especifico:

```powershell
python descarga_reportes_virusshare.py --input hashes\VirusShare_00499.txt
```

Opciones utiles:

```powershell
python descarga_reportes_virusshare.py `
  --input hashes\VirusShare_00499.txt `
  --mode file `
  --key-strategy sequential `
  --output-dir clasificacion
```

Salida esperada para `hashes\VirusShare_00499.txt`:

```text
clasificacion/
  VirusShare_00499/
    reportes/
      reporte/
        <hash>.json
    peticiones_exitosas_virusshare.txt
    peticiones_no_encontradas_virusshare.txt
    peticiones_benignas_virusshare.txt
    peticiones_errores_virusshare.txt
    estado_proceso_virusshare.json
    proceso_virusshare.log
```

Los errores de API se registran en `peticiones_errores_virusshare.txt`, pero no
crean un JSON de reporte. Asi, un fallo temporal de red, cuota o servicio no
bloquea futuros reintentos para ese hash.

## Analisis de Reportes

El script principal de analisis es:

```text
analisis_de_reportes.py
```

Genera dos artefactos:

```text
outputs/Analisis_de_Reportes_VirusShare_00499.xlsx
outputs/Analisis_de_Reportes_VirusShare_00499.db
```

El Excel sirve para revisar visualmente y filtrar. La base SQLite sirve para
consultas interactivas y analisis mas grandes sin las limitaciones de Excel.

### Prueba rapida

Antes de correr todo, prueba con pocos reportes:

```powershell
python analisis_de_reportes.py --limit 200 --name Prueba_Analisis_Reportes
```

### Ejecucion completa

```powershell
python analisis_de_reportes.py `
  --reports-dir clasificacion\VirusShare_00499\reportes\reporte `
  --output-dir outputs `
  --name Analisis_de_Reportes_VirusShare_00499
```

### Analisis consolidado de todos los lotes

Si existen varios conjuntos dentro de `clasificacion/`, por ejemplo
`VirusShare_00495`, `VirusShare_00499`, `VirusShare_00500`, etc., puedes
analizarlos en una sola salida:

```powershell
python analisis_de_reportes.py `
  --reports-root clasificacion `
  --output-dir outputs
```

Para acelerar el procesamiento en equipos con varios nucleos, usa `--workers`.
El valor `0` usa automaticamente los nucleos disponibles menos uno.

```powershell
python analisis_de_reportes.py `
  --reports-root clasificacion `
  --output-dir outputs `
  --workers 0 `
  --chunksize 200
```

Si el disco se satura, reduce `--workers`; si el CPU aun tiene margen, puedes
subirlo manualmente, por ejemplo `--workers 16`.

El script busca automaticamente carpetas con este patron:

```text
clasificacion/
  VirusShare_*/
    reportes/
      reporte/
        *.json
```

Cuando usas `--reports-root`, si no especificas `--name`, la salida queda como:

```text
outputs/Analisis_de_Reportes_Todos.xlsx
outputs/Analisis_de_Reportes_Todos.db
```

Cada muestra incluye la columna `lote_origen`, por ejemplo
`VirusShare_00499`, para saber de que conjunto proviene.

## Que incluye el Excel

- `Resumen`: conteos generales y rutas de salida.
- `Graficas`: graficas nativas de Excel con top tipos, top familias, muestras
  por dia y confianza de familia.
- `Lotes`: conteo de muestras incluidas por conjunto de origen.
- `Muestras`: detalle por hash.
- `Familia_por_Dia`: familias observadas por fecha de escaneo.
- `Tipo_por_Dia`: tipos de malware por fecha de escaneo.
- `Familia_por_Tipo`: cruce familia/tipo.
- `Familia_Tipo_Dia`: cruce fecha/familia/tipo.
- `Top_Familias`: familias mas frecuentes.
- `Top_Tipos`: tipos mas frecuentes.

Por defecto no se guardan todas las detecciones por motor, para que el Excel y
la base SQLite no crezcan demasiado. El script si lee esos motores internamente
para inferir familia/tipo, pero solo conserva resumenes como positivos, total y
porcentaje de deteccion.

Si necesitas auditar cada motor antivirus, ejecuta con:

```powershell
python analisis_de_reportes.py --include-engine-details
```

En ese modo se crean hojas `Detecciones_1`, `Detecciones_2`, etc. y se llena la
tabla `detections` en SQLite. Puede aumentar mucho el tamano de salida.

## Campos importantes

En `Muestras` se incluyen, entre otros:

- `fecha_escaneo_vt`: fecha de escaneo reportada por VirusTotal.
- `lote_origen`: lote VirusShare desde donde se obtuvo el reporte, por ejemplo
  `VirusShare_00499`.
- `fecha_agregado_virusshare`: fecha en que VirusShare agrego la muestra.
- `fecha_creacion_archivo`: fecha tomada de `exif.TimeStamp` cuando el reporte la
  incluye. En ejecutables PE suele corresponder al timestamp de compilacion; debe
  tratarse como metadato del archivo, no como fecha garantizada de creacion del
  malware.
- `timestamp_creacion_raw`: valor original de `exif.TimeStamp` para auditoria.
- `vt_positives`, `vt_total`, `detection_ratio`: indicadores de deteccion.
- `detection_percent`: porcentaje de motores que detectaron la muestra.
- `familia_probable`: familia inferida desde etiquetas AV.
- `familia_confianza`: `alta`, `media`, `baja` o `sin_inferir`.
- `familias_top`: candidatas de familia y sus votos ponderados.
- `tipo_probable`: tipo inferido, por ejemplo `trojan`, `phishing`, `worm`,
  `ransomware`, `stealer`, `backdoor`, etc.
- `tipo_confianza`: confianza de la inferencia del tipo.
- `detecciones_top`: etiquetas AV crudas para auditoria rapida.

La inferencia de familia/tipo es heuristica. El script usa votos ponderados por
motor antivirus; aun asi, para casos dudosos conviene revisar `familias_top`,
`tipos_top` y las hojas de detecciones crudas.

## Consultas interactivas con SQLite

La base `outputs/Analisis_de_Reportes_VirusShare_00499.db` contiene:

- `samples`
- `detections` solo se llena si usas `--include-engine-details`
- `family_day_counts`
- `type_day_counts`
- `family_type_counts`
- `family_type_day_counts`
- `lote_counts`

Ejemplos de consultas:

```sql
-- Familias mas comunes en una fecha
SELECT familia_probable, muestras
FROM family_day_counts
WHERE dia_escaneo_vt = '2021-11-12'
ORDER BY muestras DESC;

-- Tipos de malware por fecha
SELECT dia_escaneo_vt, tipo_probable, muestras
FROM type_day_counts
ORDER BY dia_escaneo_vt, muestras DESC;

-- Cantidad de muestras por lote
SELECT lote_origen, muestras
FROM lote_counts
ORDER BY lote_origen;

-- Familias mas comunes por lote
SELECT lote_origen, familia_probable, COUNT(*) AS muestras
FROM samples
GROUP BY lote_origen, familia_probable
ORDER BY lote_origen, muestras DESC;

-- Muestras de una familia concreta
SELECT hash_md5, lote_origen, fecha_escaneo_vt, fecha_creacion_archivo, tipo_probable, familia_confianza, vt_positives, vt_total
FROM samples
WHERE familia_probable = 'cryxos'
ORDER BY fecha_escaneo_vt;

-- Muestras por fecha de creacion/compilacion del archivo
SELECT hash_md5, familia_probable, tipo_probable, detection_percent, fecha_creacion_archivo
FROM samples
WHERE dia_creacion_archivo BETWEEN '2013-12-01' AND '2013-12-31'
ORDER BY fecha_creacion_archivo;

-- Detecciones crudas de una muestra
-- Requiere ejecutar con --include-engine-details
SELECT engine, result
FROM detections
WHERE hash_md5 = '000078146985524b12c3f5a727b831c0'
ORDER BY engine;
```

Puedes abrir el `.db` con DB Browser for SQLite, Power BI, Python, notebooks o
cualquier herramienta compatible con SQLite.

## Filtros opcionales

El analisis puede limitarse por fecha, familia, tipo o severidad:

```powershell
python analisis_de_reportes.py --date-from 2021-11-12 --date-to 2021-11-15
python analisis_de_reportes.py --family cryxos
python analisis_de_reportes.py --type phishing
python analisis_de_reportes.py --min-positives 10 --min-ratio 0.2
```

## Extraccion del dataset final

Cuando ya definas que familias, fechas y porcentajes de deteccion usaras, no
conviene mover los reportes originales. Usa `extraer_dataset_final.py` para crear
una copia curada de los reportes seleccionados y un manifest reproducible.

La configuracion se controla con:

```text
config_dataset_final.json
```

Ejecutar con la configuracion por defecto:

```powershell
python extraer_dataset_final.py --config config_dataset_final.json
```

Si quieres construir el dataset por bloques, por ejemplo hoy una familia y
manana otra, usa `--append`:

```powershell
python extraer_dataset_final.py --config configs_dataset\01_cryxos.json --append
python extraer_dataset_final.py --config configs_dataset\02_emotet.json --append
```

En modo acumulativo el script lee el manifest existente, agrega solo hashes
nuevos y evita duplicados. Cada muestra queda marcada con:

- `batch_id`: identificador del bloque de extraccion;
- `fecha_extraccion`: momento en que se agrego al manifest;
- `config_path`: JSON usado para esa extraccion;
- `config_hash`: huella SHA-256 de la configuracion aplicada.

Puedes fijar el `batch_id` en el JSON:

```json
{
  "batch_id": "01_cryxos",
  "copy_workers": 0,
  "filters": {
    "families": ["cryxos"]
  }
}
```

`copy_workers` controla cuantos hilos se usan para copiar reportes. El valor `0`
elige automaticamente un numero razonable para I/O; puedes fijarlo, por ejemplo
`8`, `16` o `32`, si quieres controlar la carga del disco.

El JSON indica la base de entrada, la carpeta de salida, filtros y limites de
seleccion. Ejemplo de criterios incluidos:

- excluir `sin_inferir`;
- usar solo familias con confianza `alta` o `media`;
- exigir un porcentaje minimo de deteccion;
- limitar la cantidad maxima por familia;
- balancear por rango de deteccion, fecha y lote de origen.

Salidas esperadas:

```text
outputs/dataset_final/
  manifest_dataset_final.csv
  manifest_dataset_final.xlsx
  seleccion_dataset_final.db
  config_usada_dataset_final.json
  reportes/
    <familia>/
      <lote_origen>/
        <hash>.json
```

El manifest conserva el hash, familia, tipo, porcentaje de deteccion, fechas,
lote de origen, ruta original del reporte y ruta copiada. Para probar sin copiar
reportes:

```powershell
python extraer_dataset_final.py --config config_dataset_final.json --dry-run
```

## Notas de Git

Las carpetas `clasificacion/` y `outputs/` pueden crecer mucho. `outputs/` esta
ignorada por `.gitignore`. Si decides ignorar tambien nuevos reportes de
`clasificacion/`, recuerda que los archivos que ya estan versionados seguiran en
Git hasta retirarlos explicitamente del indice.
