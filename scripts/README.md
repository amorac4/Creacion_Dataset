# Scripts

Los scripts se agrupan por responsabilidad:

- `virusshare/`: entrada de datos desde VirusShare.
- `analysis/`: procesamiento de reportes, balance y auditorias.
- `dataset/`: extraccion y organizacion de datasets finales.

Los wrappers de la raiz siguen disponibles para compatibilidad. Por ejemplo,
estos dos comandos son equivalentes:

```powershell
python analisis_de_reportes.py --reports-root clasificacion --output-dir outputs
python scripts/analysis/analisis_de_reportes.py --reports-root clasificacion --output-dir outputs
```

