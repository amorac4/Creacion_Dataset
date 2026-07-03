# Prueba de etiquetado con AVClass

Esta carpeta contiene un experimento aislado para evaluar si AVClass produce
mejores etiquetas de familia que la heuristica local de `analisis_de_reportes.py`.

AVClass acepta reportes VirusTotal v2 en JSONL: un objeto JSON por linea. Los
reportes descargados en este repo vienen envueltos en JSON de VirusShare, asi que
primero se exporta una vista compatible con VT v2.

## 1. Instalar AVClass

```powershell
python -m pip install avclass-malicialab
```

Si `python` no esta en PATH, usa el Python del entorno que tengas activo.

## 2. Exportar reportes a JSONL compatible con AVClass

Para todos los lotes en `clasificacion/VirusShare_*/reportes/reporte`:

```powershell
python experimentos_avclass/exportar_vt_v2_jsonl.py `
  --reports-root clasificacion `
  --output experimentos_avclass/data/virusshare_vtv2.jsonl
```

Para probar rapido con pocos reportes:

```powershell
python experimentos_avclass/exportar_vt_v2_jsonl.py `
  --reports-root clasificacion `
  --output experimentos_avclass/data/prueba_1000_vtv2.jsonl `
  --limit 1000
```

Para evaluar solo hashes que ya estan en los CSV curados de `Dataset_V1/csv`:

```powershell
python experimentos_avclass/exportar_vt_v2_jsonl.py `
  --reports-root clasificacion `
  --dataset-csv-dir Dataset_V1/csv `
  --output experimentos_avclass/data/dataset_v1_prueba_1000_vtv2.jsonl `
  --limit 1000
```

## 3. Ejecutar AVClass

Familia probable por muestra:

```powershell
avclass -f experimentos_avclass/data/prueba_1000_vtv2.jsonl `
  -o experimentos_avclass/results/prueba_1000_avclass.labels
```

Tags completos por muestra:

```powershell
avclass -f experimentos_avclass/data/prueba_1000_vtv2.jsonl `
  -t `
  -o experimentos_avclass/results/prueba_1000_avclass.tags
```

AVClass usa `SINGLETON:<hash>` cuando no puede inferir una familia. Para este
dataset conviene mapear esos casos a `sin_inferir`.

## 4. Comparar contra etiquetas actuales

Contra los CSV curados de `Dataset_V1/csv`:

```powershell
python experimentos_avclass/comparar_avclass.py `
  --avclass-labels experimentos_avclass/results/prueba_1000_avclass.labels `
  --dataset-csv-dir Dataset_V1/csv `
  --output-csv experimentos_avclass/results/comparacion_prueba_1000.csv `
  --summary-json experimentos_avclass/results/comparacion_prueba_1000.json
```

El comparador genera:

- `comparacion_*.csv`: una fila por hash en comun.
- `comparacion_*.json`: resumen con coincidencias, desacuerdos y familias top.

## Criterio para decidir

AVClass vale la pena si:

- reduce etiquetas genericas como `malurl`, `static`, `badfile`, `camelot`;
- aumenta `sin_inferir` en casos sin familia real, en vez de inventarla;
- mantiene familias fuertes del dataset actual, como `salgorea`, `fragtor`,
  `injector`, `vbclone`, `zusy`;
- los desacuerdos importantes son explicables revisando `detecciones_top`.
