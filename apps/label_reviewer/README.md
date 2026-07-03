# Label Reviewer

Interfaz web local para revisar etiquetas hibridas.

No requiere dependencias externas: usa `http.server`, `sqlite3`, HTML, CSS y JS
del lado del navegador.

## Ejecutar

Desde la raiz del repositorio:

```powershell
python apps/label_reviewer/server.py --db outputs/Analisis_de_Reportes_Todos.db --port 8765
```

Abrir:

```text
http://127.0.0.1:8765
```

## Tablas que usa

Debe existir una DB enriquecida con:

- `samples`
- `label_hybrid`
- `label_pair_summary`
- `label_review_queue`

La app crea estas tablas auxiliares si no existen:

- `label_manual_reviews`: revisiones por hash.
- `label_pair_manual_reviews`: decisiones por par local/AVClass.

Tambien crea la vista:

- `label_final_current`: etiqueta final considerando overrides manuales.

## Flujo

1. Revisar pares conflictivos en "Pares".
2. Aplicar decision por par cuando sea clara.
3. Revisar hashes concretos en "Cola".
4. Exportar CSV final desde el dashboard.

