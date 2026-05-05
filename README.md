# Creacion Dataset VirusShare

Script para crear un dataset local de reportes de VirusShare a partir de un TXT
con hashes. El archivo principal es `descarga_reportes_virusshare.py`.

## Instalacion

```powershell
python -m venv .venv
.\.venv\Scripts\Activate.ps1
python -m pip install -r requirements.txt
```

## Configuracion de API keys

Usa una de estas variables de entorno:

```powershell
$env:VIRUSSHARE_API_KEY = "tu_api_key"
$env:VIRUSSHARE_API_KEYS = "key1,key2,key3"
```

Tambien puedes llenar `API_KEYS` dentro del script, aunque las variables de
entorno son preferibles para no guardar secretos en Git.

## Uso

Ejecutar con el lote por defecto:

```powershell
python descarga_reportes_virusshare.py
```

Ejecutar con un TXT concreto:

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

Variables de entorno equivalentes:

- `VIRUSSHARE_INPUT_TXT`
- `VIRUSSHARE_MODE`
- `VIRUSSHARE_KEY_STRATEGY`
- `VIRUSSHARE_MAX_KEYS`
- `VIRUSSHARE_OUTPUT_DIR`
- `VIRUSSHARE_REQUESTS_PER_MIN`
- `VIRUSSHARE_DAILY_LIMIT`
- `VIRUSSHARE_CONTINUE_ON_PREVIOUS_QUOTA`

## Salida

Para `hashes\VirusShare_00499.txt`, el script crea:

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
crean JSON de reporte. Asi, un fallo temporal de red, cuota o servicio no deja
ese hash bloqueado para ejecuciones futuras.

## Notas de Git

La carpeta `clasificacion/` contiene datos generados y puede crecer mucho. Esta
ignorada en `.gitignore` para nuevas salidas. Si ya hay datos versionados, Git
los seguira rastreando hasta que se retiren explicitamente del indice.

## Filtrado y analisis de reportes

Una vez descargados los JSON, genera tablas CSV filtrables con:

```powershell
python filtrar_reportes_virusshare.py
```

Salida por defecto:

```text
dataset_filtrado/
  VirusShare_00499/
    muestras_filtradas.csv
    detecciones_por_motor.csv
    resumen_familias.csv
    resumen_tipos.csv
    resumen_dias.csv
```

Campos principales:

- `fecha_escaneo_vt`: fecha de escaneo reportada por VirusTotal.
- `fecha_agregado_virusshare`: fecha en que VirusShare agrego la muestra.
- `familia_probable`: familia inferida por votos desde las etiquetas AV.
- `tipo_probable`: tipo inferido por votos, por ejemplo `trojan`, `worm`,
  `backdoor`, `phishing`, `ransomware`, `stealer`, etc.
- `detecciones_top`: primeras etiquetas crudas para auditoria rapida.

Ejemplos de filtros:

```powershell
python filtrar_reportes_virusshare.py --min-positives 10
python filtrar_reportes_virusshare.py --date-from 2021-11-12 --date-to 2021-11-15
python filtrar_reportes_virusshare.py --type phishing
python filtrar_reportes_virusshare.py --family expiro
python filtrar_reportes_virusshare.py --family-score-threshold 0.20
```

La inferencia de familia/tipo es heuristica. Para investigacion fina, revisa
`detecciones_por_motor.csv`, que conserva cada motor y su etiqueta original.
Si `familia_probable` queda vacia, el resumen la agrupa como `sin_inferir`;
puedes bajar o subir `--family-score-threshold` segun que tan estricta quieras
la inferencia.

## Excel consolidado

Para crear un solo libro de Excel con todas las tablas:

```powershell
python crear_excel_virusshare.py
```

El archivo queda en:

```text
outputs/VirusShare_00499_dataset_filtrado.xlsx
```

Como `detecciones_por_motor.csv` supera el limite de filas por hoja de Excel, el
script divide esa tabla en hojas `Detecciones_1`, `Detecciones_2`, etc. dentro
del mismo libro.
