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

## Que incluye el Excel

- `Resumen`: conteos generales y rutas de salida.
- `Graficas`: graficas nativas de Excel con top tipos, top familias, muestras
  por dia y confianza de familia.
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
- `fecha_agregado_virusshare`: fecha en que VirusShare agrego la muestra.
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

-- Muestras de una familia concreta
SELECT hash_md5, fecha_escaneo_vt, tipo_probable, familia_confianza, vt_positives, vt_total
FROM samples
WHERE familia_probable = 'cryxos'
ORDER BY fecha_escaneo_vt;

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

## Notas de Git

Las carpetas `clasificacion/` y `outputs/` pueden crecer mucho. `outputs/` esta
ignorada por `.gitignore`. Si decides ignorar tambien nuevos reportes de
`clasificacion/`, recuerda que los archivos que ya estan versionados seguiran en
Git hasta retirarlos explicitamente del indice.
