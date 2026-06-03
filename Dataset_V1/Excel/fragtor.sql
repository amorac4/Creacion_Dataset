SELECT hash_md5, lote_origen, familia_probable, size, extension, tipo_probable, detection_percent, fecha_escaneo_vt, dia_creacion_archivo
FROM samples
WHERE familia_probable = 'injector'
AND extension = 'exe'
ORDER BY dia_creacion_archivo;