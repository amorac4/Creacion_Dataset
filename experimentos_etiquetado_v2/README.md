# Experimento de etiquetado v2

Espacio aislado para desarrollar y evaluar un sistema de etiquetado de familias
de malware con mayor precision, buena cobertura y abstencion explicita cuando la
evidencia sea insuficiente.

Este experimento no debe sobrescribir las etiquetas ni las bases actuales. Sus
resultados se guardaran en tablas y artefactos versionados como `label_v2` hasta
que hayan sido evaluados contra un conjunto de referencia revisado manualmente.

## Objetivos

- Separar familia, tipo, plataforma, comportamiento, packer y variante.
- Normalizar alias a una taxonomia canonica y auditable.
- Combinar evidencia local, AVClass y consenso entre motores independientes.
- Reducir `sin_inferir` sin convertir evidencia debil en etiquetas definitivas.
- Calibrar niveles de confianza con un conjunto de referencia manual.
- Producir una unica `familia_final` consumible por el dataset de imagenes.

## Estructura

```text
experimentos_etiquetado_v2/
  config/       Configuraciones, taxonomias y reglas versionadas.
  data/         Entradas generadas o temporales; no versionar datos pesados.
  results/      Metricas, comparaciones y salidas del experimento.
  src/          Implementacion del pipeline label_v2.
  tests/        Pruebas unitarias y casos de etiquetado conocidos.
```

## Principio de seguridad

Una muestra puede quedar como `sin_inferir` cuando no exista evidencia
suficiente. La cobertura debe medirse junto con precision, macro-F1, calibracion
de confianza y tasa de abstencion.

## Ejecutar la linea base

Desde la raiz del repositorio:

```powershell
python experimentos_etiquetado_v2/src/baseline.py
```

La ejecucion lee la base consolidada en modo consulta y escribe el resumen y el
inventario de familias en `results/baseline/`.
