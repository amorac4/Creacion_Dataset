#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
================================================================================
SCRIPT: descarga_reportes_virusshare.py
================================================================================

Lee hashes desde un TXT (uno por línea) y consulta VirusShare API v2.

Puede configurarse al inicio del script, con variables de entorno o con
argumentos por consola. Ejemplo:

    python descarga_reportes_virusshare.py --input hashes/VirusShare_00499.txt

ESTRUCTURA DE SALIDA POR LOTE
-----------------------------
Si INPUT_TXT = "VirusShare_00499.txt", se crea:

    clasificacion/
      VirusShare_00499/
        reportes/
          reporte/
            <hash>.json
        peticiones_exitosas_virusshare.txt
        peticiones_no_encontradas_virusshare.txt
        peticiones_benignas_virusshare.txt
        estado_proceso_virusshare.json
        proceso_virusshare.log

EVITA REPETICIONES DE 3 FORMAS
------------------------------
1) Deduplica el TXT de entrada.
2) Si ya existe el JSON del hash, no lo vuelve a pedir.
3) Si el hash ya está registrado como OK / NF / BN en los TXT del lote, no lo vuelve a pedir.

API KEYS
--------
Puedes ponerlas aquí en API_KEYS o usar variables de entorno:
- VIRUSSHARE_API_KEY
- VIRUSSHARE_API_KEYS="k1,k2,k3,..."
================================================================================
"""

from __future__ import annotations

import argparse
import datetime
import json
import logging
import os
from pathlib import Path
import re
import sys
import time
from typing import Any, Optional

import requests
from requests import Session
from requests.exceptions import RequestException
from tqdm import tqdm


# ==============================================================================
# [ --- CONFIGURACIÓN GENERAL --- ]
# ==============================================================================
INPUT_TXT = r"hashes\VirusShare_00499.txt"
MODE = "file"                  # "file" o "quick"
KEY_STRATEGY = "sequential"   # "sequential" o "roundrobin"
MAX_KEYS = 0                    # solo roundrobin: 0=todas, 1..N=usar N keys

# Si dejas API_KEYS vacío, intentará leer del entorno:
# - VIRUSSHARE_API_KEY
# - VIRUSSHARE_API_KEYS="k1,k2,k3,..."
API_KEYS: list[str] = []

BASE_URL = "https://virusshare.com/apiv2"
REQUESTS_PER_MIN_PER_KEY = 4
DAILY_LIMIT_PER_KEY = 5760
HOURS_TO_WAIT_ON_QUOTA = 24
BACKOFF_ON_204_SECONDS = 60
REQUEST_TIMEOUT = 45
CONTINUAR_SI_HAY_CUOTA_PREVIA = True

DIR_SALIDA_BASE = "clasificacion"
SUBCARPETAS_REPORTES = ("reportes", "reporte")
ESTADO_FILE_NAME = "estado_proceso_virusshare.json"
LOG_FILE_NAME = "proceso_virusshare.log"
LISTA_EXITOSOS_FILE_NAME = "peticiones_exitosas_virusshare.txt"
LISTA_NO_ENCONTRADOS_FILE_NAME = "peticiones_no_encontradas_virusshare.txt"
LISTA_BENIGNOS_FILE_NAME = "peticiones_benignas_virusshare.txt"
LISTA_ERRORES_FILE_NAME = "peticiones_errores_virusshare.txt"
# ==============================================================================


VALID_HASH_RE = re.compile(
    r"^[0-9a-fA-F]{32}$|^[0-9a-fA-F]{40}$|^[0-9a-fA-F]{56}$|^[0-9a-fA-F]{64}$|^[0-9a-fA-F]{96}$|^[0-9a-fA-F]{128}$"
)


class VirusShareAPIError(Exception):
    def __init__(self, code: str, message: str, status_code: Optional[int] = None) -> None:
        super().__init__(message)
        self.code = code
        self.message = message
        self.status_code = status_code


def env_bool(name: str, default: bool) -> bool:
    value = os.environ.get(name)
    if value is None:
        return default
    return value.strip().lower() in {"1", "true", "yes", "y", "si", "s"}


def env_int(name: str, default: int) -> int:
    value = os.environ.get(name, "").strip()
    if not value:
        return default
    try:
        return int(value)
    except ValueError:
        return default


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Descarga reportes de VirusShare para un TXT de hashes."
    )
    parser.add_argument(
        "-i",
        "--input",
        default=os.environ.get("VIRUSSHARE_INPUT_TXT", INPUT_TXT),
        help="TXT de hashes de entrada. Tambien puede usarse VIRUSSHARE_INPUT_TXT.",
    )
    parser.add_argument(
        "--mode",
        choices=("file", "quick"),
        default=os.environ.get("VIRUSSHARE_MODE", MODE),
        help="Endpoint de VirusShare a consultar.",
    )
    parser.add_argument(
        "--key-strategy",
        choices=("sequential", "roundrobin"),
        default=os.environ.get("VIRUSSHARE_KEY_STRATEGY", KEY_STRATEGY),
        help="Estrategia de uso de API keys.",
    )
    parser.add_argument(
        "--max-keys",
        type=int,
        default=env_int("VIRUSSHARE_MAX_KEYS", MAX_KEYS),
        help="Solo roundrobin: 0=todas, 1..N=usar N keys.",
    )
    parser.add_argument(
        "--output-dir",
        default=os.environ.get("VIRUSSHARE_OUTPUT_DIR", DIR_SALIDA_BASE),
        help="Directorio base de salida. Tambien puede usarse VIRUSSHARE_OUTPUT_DIR.",
    )
    parser.add_argument(
        "--requests-per-minute",
        type=int,
        default=env_int("VIRUSSHARE_REQUESTS_PER_MIN", REQUESTS_PER_MIN_PER_KEY),
        help="Limite de peticiones por minuto por key.",
    )
    parser.add_argument(
        "--daily-limit",
        type=int,
        default=env_int("VIRUSSHARE_DAILY_LIMIT", DAILY_LIMIT_PER_KEY),
        help="Limite diario por key para esta sesion.",
    )
    parser.add_argument(
        "--no-continue-on-quota",
        action="store_true",
        default=not env_bool("VIRUSSHARE_CONTINUE_ON_PREVIOUS_QUOTA", CONTINUAR_SI_HAY_CUOTA_PREVIA),
        help="Detiene la ejecucion si hay una pausa previa por cuota aun reciente.",
    )
    return parser.parse_args()


class VirusShareClient:
    def __init__(self, api_key: str, key_id: str, timeout: int = REQUEST_TIMEOUT) -> None:
        self.api_key = api_key.strip()
        self.key_id = key_id
        self.timeout = timeout
        self.session: Session = requests.Session()
        self.session.headers.update({"User-Agent": f"virusshare-client/1.0 ({key_id})"})

    def close(self) -> None:
        try:
            self.session.close()
        except Exception:
            pass

    def _get(self, endpoint: str, hash_value: str) -> dict[str, Any]:
        url = f"{BASE_URL}/{endpoint.lstrip('/')}"
        params = {"apikey": self.api_key, "hash": hash_value}

        try:
            resp = self.session.get(url, params=params, timeout=self.timeout)
        except RequestException as e:
            raise VirusShareAPIError("RequestException", str(e)) from e

        if resp.status_code == 200:
            try:
                return resp.json()
            except json.JSONDecodeError as e:
                raise VirusShareAPIError("InvalidJSON", f"Respuesta 200 pero no es JSON válido: {e}") from e

        if resp.status_code == 204:
            raise VirusShareAPIError("QuotaExceededError", "VirusShare devolvió HTTP 204 (rate limit/cuota).", 204)
        if resp.status_code == 400:
            raise VirusShareAPIError("BadRequestError", "VirusShare devolvió HTTP 400 (bad request).", 400)
        if resp.status_code == 403:
            raise VirusShareAPIError("WrongCredentialsError", "VirusShare devolvió HTTP 403 (API key inválida o sin permisos).", 403)
        if resp.status_code == 404:
            raise VirusShareAPIError("NotFoundHttpError", "VirusShare devolvió HTTP 404.", 404)
        if resp.status_code == 500:
            raise VirusShareAPIError("InternalServerError", "VirusShare devolvió HTTP 500.", 500)
        if resp.status_code == 503:
            raise VirusShareAPIError("ServiceUnavailableError", "VirusShare devolvió HTTP 503.", 503)

        raise VirusShareAPIError(
            "UnexpectedHttpError",
            f"HTTP {resp.status_code}: {resp.text[:300]}",
            resp.status_code,
        )

    def get_file_report(self, hash_value: str) -> dict[str, Any]:
        return self._get("file", hash_value)

    def get_quick_report(self, hash_value: str) -> dict[str, Any]:
        return self._get("quick", hash_value)


def key_id_from_key(key: str, idx: int) -> str:
    key = key.strip()
    if len(key) >= 8:
        return f"Key-VS-{idx + 1} ({key[:4]}...{key[-4:]})"
    return f"Key-VS-{idx + 1} (SHORT)"


def load_api_keys() -> list[str]:
    if API_KEYS:
        return [k.strip() for k in API_KEYS if k.strip()]

    env_one = os.environ.get("VIRUSSHARE_API_KEY", "").strip()
    env_many = os.environ.get("VIRUSSHARE_API_KEYS", "").strip()

    keys: list[str] = []
    if env_one:
        keys.append(env_one)
    if env_many:
        keys.extend([k.strip() for k in env_many.split(",") if k.strip()])

    # deduplicar preservando orden
    seen: set[str] = set()
    ordered: list[str] = []
    for key in keys:
        if key not in seen:
            seen.add(key)
            ordered.append(key)
    return ordered


def leer_hashes_txt(txt_path: Path) -> tuple[list[str], int]:
    if not txt_path.is_file():
        raise FileNotFoundError(f"No existe el TXT de hashes: {txt_path}")

    vistos: set[str] = set()
    hashes: list[str] = []
    skip_count = 0

    for line_no, raw in enumerate(txt_path.read_text(encoding="utf-8", errors="ignore").splitlines(), start=1):
        value = raw.strip().split()[0] if raw.strip() else ""
        value = value.strip().lower().replace('"', "").replace("'", "")
        if not value:
            continue
        if not VALID_HASH_RE.fullmatch(value):
            skip_count += 1
            logging.warning(f"[SKIP] Línea {line_no} sin hash válido: '{raw}'")
            continue
        if value not in vistos:
            vistos.add(value)
            hashes.append(value)

    return hashes, skip_count


def interpretar_response(response_value: Any) -> str:
    if response_value == 0:
        return "desconocido_o_no_encontrado"
    if response_value == 1:
        return "malware_o_detectado"
    if response_value == 2:
        return "benigno_o_sin_detecciones"
    return "respuesta_desconocida"


def leer_estado(estado_file_path: Path) -> dict[str, Any]:
    if not estado_file_path.is_file():
        return {}
    try:
        return json.loads(estado_file_path.read_text(encoding="utf-8"))
    except Exception:
        logging.warning("No se pudo leer el archivo de estado (puede estar corrupto).", exc_info=True)
        return {}


def escribir_estado(estado_file_path: Path, estado_str: str, extras: Optional[dict[str, Any]] = None) -> None:
    data: dict[str, Any] = {
        "ultimo_guardado": datetime.datetime.now(datetime.timezone.utc).isoformat(),
        "ultimo_estado": estado_str,
    }
    if extras:
        data.update(extras)
    try:
        estado_file_path.write_text(json.dumps(data, indent=4, ensure_ascii=False), encoding="utf-8")
        logging.info(f"Escribiendo nuevo estado: {estado_str} extras={extras or {}}")
    except Exception:
        logging.warning("No se pudo escribir el archivo de estado.", exc_info=True)


def leer_registro_hashes(registro_path: Path) -> set[str]:
    hashes: set[str] = set()
    if not registro_path.is_file():
        return hashes

    for line_no, raw in enumerate(registro_path.read_text(encoding="utf-8", errors="ignore").splitlines(), start=1):
        line = raw.strip()
        if not line or line.lower().startswith("hash"):
            continue
        hash_linea = line.split("\t")[0].strip().lower()
        if VALID_HASH_RE.fullmatch(hash_linea):
            hashes.add(hash_linea)
        else:
            logging.warning(f"[SKIP] Línea inválida en registro {registro_path.name}:{line_no}: '{raw}'")
    return hashes


def asegurar_header_registro(registro_path: Path) -> None:
    if not registro_path.exists():
        registro_path.write_text("Hash\n", encoding="utf-8")


def append_registro(registro_path: Path, hash_value: str) -> None:
    asegurar_header_registro(registro_path)
    with registro_path.open("a", encoding="utf-8") as f:
        f.write(f"{hash_value}\n")


def append_registro_error(registro_path: Path, hash_value: str, code: str, message: str, key_id: str) -> None:
    if not registro_path.exists():
        registro_path.write_text("Hash\tCode\tMessage\tKeyId\tUTC\n", encoding="utf-8")
    safe_message = message.replace("\t", " ").replace("\n", " ")[:500]
    timestamp = datetime.datetime.now(datetime.timezone.utc).isoformat()
    with registro_path.open("a", encoding="utf-8") as f:
        f.write(f"{hash_value}\t{code}\t{safe_message}\t{key_id}\t{timestamp}\n")


def construir_error_json(code: str, message: str, hash_val: str, key_id: str, lote: str, input_txt: Path) -> dict[str, Any]:
    return {
        "error_details": {
            "code": code,
            "message": message,
            "hash_consultado": hash_val,
            "key_id_usada": key_id,
        },
        "data_structure": {
            "lote": lote,
            "input_txt": str(input_txt),
            "hash": hash_val,
        },
    }


def main() -> None:
    args = parse_args()
    input_path = Path(args.input)
    lote_nombre = input_path.stem
    mode = args.mode
    key_strategy = args.key_strategy
    max_keys = max(args.max_keys, 0)
    output_dir = Path(args.output_dir)
    requests_per_minute = max(args.requests_per_minute, 1)
    daily_limit = max(args.daily_limit, 1)
    continuar_si_hay_cuota_previa = not args.no_continue_on_quota

    lote_dir = output_dir / lote_nombre
    reportes_dir = lote_dir.joinpath(*SUBCARPETAS_REPORTES)
    estado_file_path = lote_dir / ESTADO_FILE_NAME
    log_file_path = lote_dir / LOG_FILE_NAME
    lista_exitosos_file = lote_dir / LISTA_EXITOSOS_FILE_NAME
    lista_no_encontrados_file = lote_dir / LISTA_NO_ENCONTRADOS_FILE_NAME
    lista_benignos_file = lote_dir / LISTA_BENIGNOS_FILE_NAME
    lista_errores_file = lote_dir / LISTA_ERRORES_FILE_NAME

    lote_dir.mkdir(parents=True, exist_ok=True)
    reportes_dir.mkdir(parents=True, exist_ok=True)

    logging.basicConfig(
        level=logging.INFO,
        format="%(asctime)s [%(levelname)s] - %(message)s",
        datefmt="%Y-%m-%d %H:%M:%S",
        handlers=[logging.FileHandler(log_file_path, mode="a", encoding="utf-8")],
        force=True,
    )

    print(f"\nLote actual: '{lote_nombre}'")
    print(f"TXT de entrada: '{input_path.resolve()}'")
    print(f"Carpeta de reportes: '{reportes_dir.resolve()}'")
    logging.info(f"Lote: {lote_nombre}")
    logging.info(f"Input TXT: {input_path.resolve()}")
    logging.info(f"Modo: {mode}")

    keys = load_api_keys()
    if not keys:
        print("Error: No se encontró ninguna API key de VirusShare.")
        print("Configura API_KEYS en el script o usa VIRUSSHARE_API_KEY / VIRUSSHARE_API_KEYS.")
        sys.exit(1)

    if key_strategy not in {"sequential", "roundrobin"}:
        print("Error: KEY_STRATEGY debe ser 'sequential' o 'roundrobin'.")
        sys.exit(1)

    clients_pool: dict[str, VirusShareClient] = {}
    all_key_ids: list[str] = []
    for idx, key in enumerate(keys):
        kid = key_id_from_key(key, idx)
        all_key_ids.append(kid)
        clients_pool[kid] = VirusShareClient(key, kid, timeout=REQUEST_TIMEOUT)

    if key_strategy == "roundrobin":
        if max_keys == 0:
            active_key_ids = all_key_ids[:]
        else:
            active_key_ids = all_key_ids[:max_keys]
    else:
        active_key_ids = all_key_ids[:]

    espera_por_key = 60.0 / requests_per_minute
    if key_strategy == "roundrobin":
        tiempo_espera_api = espera_por_key / max(len(active_key_ids), 1)
    else:
        tiempo_espera_api = espera_por_key

    estado_previo = leer_estado(estado_file_path)
    if estado_previo.get("ultimo_estado") == "PAUSADO_POR_CUOTA":
        try:
            ultimo_guardado_dt = datetime.datetime.fromisoformat(estado_previo["ultimo_guardado"])
            tiempo_transcurrido = datetime.datetime.now(datetime.timezone.utc) - ultimo_guardado_dt
            horas_transcurridas = tiempo_transcurrido.total_seconds() / 3600
            if horas_transcurridas < HOURS_TO_WAIT_ON_QUOTA:
                horas_restantes = HOURS_TO_WAIT_ON_QUOTA - horas_transcurridas
                msg = (
                    f"Se detectó pausa previa por cuota. Han pasado {horas_transcurridas:.1f} h; "
                    f"faltan {horas_restantes:.1f} h para las {HOURS_TO_WAIT_ON_QUOTA} recomendadas."
                )
                print("[ADVERTENCIA] " + msg)
                logging.warning(msg)
                if not continuar_si_hay_cuota_previa:
                    print("Ejecución detenida por configuración (CONTINUAR_SI_HAY_CUOTA_PREVIA=False).")
                    sys.exit(0)
        except Exception:
            logging.warning("No se pudo interpretar la fecha del estado previo.", exc_info=True)

    escribir_estado(estado_file_path, "EJECUTANDO", extras={"input_txt": str(input_path)})

    try:
        hashes_totales, hashes_invalidos = leer_hashes_txt(input_path)
    except Exception as e:
        print(f"Error leyendo hashes: {e}")
        escribir_estado(estado_file_path, "ERROR_LEYENDO_TXT")
        sys.exit(1)

    if not hashes_totales:
        print("Error: No se encontraron hashes válidos en el TXT.")
        escribir_estado(estado_file_path, "ERROR_SIN_HASHES")
        sys.exit(1)

    for registro in (lista_exitosos_file, lista_no_encontrados_file, lista_benignos_file):
        asegurar_header_registro(registro)
    if not lista_errores_file.exists():
        lista_errores_file.write_text("Hash\tCode\tMessage\tKeyId\tUTC\n", encoding="utf-8")

    registrados_ok = leer_registro_hashes(lista_exitosos_file)
    registrados_nf = leer_registro_hashes(lista_no_encontrados_file)
    registrados_bn = leer_registro_hashes(lista_benignos_file)
    hashes_registrados = registrados_ok | registrados_nf | registrados_bn

    hashes_pendientes: list[str] = []
    skip_por_registro = 0
    skip_por_json = 0

    for hash_val in hashes_totales:
        ruta_json_archivo = reportes_dir / f"{hash_val}.json"
        if hash_val in hashes_registrados:
            skip_por_registro += 1
            logging.info(f"[SKIP] Ya registrado previamente: {hash_val}")
            continue
        if ruta_json_archivo.is_file():
            skip_por_json += 1
            logging.info(f"[SKIP] Ya existe JSON: {ruta_json_archivo.as_posix()}")
            continue
        hashes_pendientes.append(hash_val)

    print(f"\nHashes válidos únicos en TXT: {len(hashes_totales)}")
    print(f"Líneas inválidas/omitidas: {hashes_invalidos}")
    print(f"Saltados por registro previo: {skip_por_registro}")
    print(f"Saltados por JSON existente: {skip_por_json}")
    print(f"Pendientes reales: {len(hashes_pendientes)}")

    if not hashes_pendientes:
        print("No hay hashes pendientes. El lote ya está procesado.")
        escribir_estado(estado_file_path, "COMPLETADO")
        sys.exit(0)

    daily_budget = len(active_key_ids) * daily_limit
    hashes_a_procesar_hoy = hashes_pendientes[:daily_budget]
    if len(hashes_pendientes) > daily_budget:
        print(f"\nSe procesarán {len(hashes_a_procesar_hoy)} hashes de {len(hashes_pendientes)} por cuota de sesión.")
    else:
        print(f"\nSe procesarán los {len(hashes_a_procesar_hoy)} hashes pendientes.")

    exito_count = 0
    nf_count = 0
    benign_count = 0
    error_count = 0
    estado_final = "COMPLETADO"

    current_key_idx = 0
    per_key_used = {k: 0 for k in active_key_ids}
    next_key_from_state = estado_previo.get("next_free_key_id")
    if key_strategy == "sequential" and next_key_from_state in active_key_ids:
        current_key_idx = active_key_ids.index(next_key_from_state)
        logging.info(f"Reanudando sequential desde key guardada: {next_key_from_state}")

    def get_client_roundrobin(i: int) -> tuple[str, VirusShareClient]:
        kid = active_key_ids[i % len(active_key_ids)]
        return kid, clients_pool[kid]

    def get_client_sequential() -> Optional[tuple[str, VirusShareClient]]:
        nonlocal current_key_idx
        tries = 0
        while tries < len(active_key_ids):
            kid = active_key_ids[current_key_idx % len(active_key_ids)]
            if per_key_used.get(kid, 0) < daily_limit:
                return kid, clients_pool[kid]
            current_key_idx += 1
            tries += 1
        return None

    pbar = tqdm(total=len(hashes_a_procesar_hoy), desc="Procesando Hashes", unit="hash")

    try:
        for i, hash_val in enumerate(hashes_a_procesar_hoy):
            if key_strategy == "roundrobin":
                kid, client = get_client_roundrobin(i)
            else:
                elegido = get_client_sequential()
                if elegido is None:
                    estado_final = "PAUSADO_POR_CUOTA"
                    print("\n[ADVERTENCIA] Todas las keys activas alcanzaron el límite configurado.")
                    break
                kid, client = elegido

            pbar.set_postfix_str(f"OK:{exito_count} NF:{nf_count} BN:{benign_count} ERR:{error_count} Key:{kid}")
            ruta_json_archivo = reportes_dir / f"{hash_val}.json"

            logging.info(f"START hash={hash_val} lote={lote_nombre} key={kid} mode={mode}")

            try:
                data = client.get_file_report(hash_val) if mode == "file" else client.get_quick_report(hash_val)
                response_value = data.get("response")
                estado_resp = interpretar_response(response_value)

                if response_value == 1:
                    data["data_structure"] = {
                        "lote": lote_nombre,
                        "input_txt": str(input_path),
                        "hash": hash_val,
                        "mode": mode,
                    }
                    ruta_json_archivo.write_text(json.dumps(data, indent=4, ensure_ascii=False), encoding="utf-8")
                    append_registro(lista_exitosos_file, hash_val)
                    exito_count += 1
                    logging.info(f"OK   hash={hash_val} estado={estado_resp} saved='{ruta_json_archivo.as_posix()}'")

                elif response_value == 0:
                    error_json = construir_error_json(
                        "NotFoundError",
                        "VirusShare respondió response=0 (desconocido/no encontrado).",
                        hash_val,
                        kid,
                        lote_nombre,
                        input_path,
                    )
                    ruta_json_archivo.write_text(json.dumps(error_json, indent=4, ensure_ascii=False), encoding="utf-8")
                    append_registro(lista_no_encontrados_file, hash_val)
                    nf_count += 1
                    logging.info(f"NF   hash={hash_val} saved='{ruta_json_archivo.as_posix()}'")

                elif response_value == 2:
                    error_json = construir_error_json(
                        "BenignOrUndetected",
                        "VirusShare respondió response=2 (benigno o sin detecciones).",
                        hash_val,
                        kid,
                        lote_nombre,
                        input_path,
                    )
                    ruta_json_archivo.write_text(json.dumps(error_json, indent=4, ensure_ascii=False), encoding="utf-8")
                    append_registro(lista_benignos_file, hash_val)
                    benign_count += 1
                    logging.info(f"BN   hash={hash_val} saved='{ruta_json_archivo.as_posix()}'")

                else:
                    raise VirusShareAPIError(
                        "UnexpectedResponseValue",
                        f"VirusShare devolvió un valor inesperado en 'response': {response_value}"
                    )

                per_key_used[kid] = per_key_used.get(kid, 0) + 1

            except VirusShareAPIError as e:
                error_count += 1
                append_registro_error(lista_errores_file, hash_val, e.code, e.message, kid)
                logging.error(f"ERR  hash={hash_val} code={e.code} msg={e.message}")

                if e.code == "QuotaExceededError":
                    per_key_used[kid] = daily_limit
                    if key_strategy == "sequential":
                        current_key_idx += 1
                        if current_key_idx < 10**9 and active_key_ids:
                            next_k = active_key_ids[current_key_idx % len(active_key_ids)]
                            escribir_estado(estado_file_path, "EJECUTANDO", extras={"next_free_key_id": next_k})
                    time.sleep(BACKOFF_ON_204_SECONDS)

                elif e.code == "WrongCredentialsError":
                    per_key_used[kid] = daily_limit

            pbar.update(1)
            time.sleep(tiempo_espera_api)

    except KeyboardInterrupt:
        print("\nProceso interrumpido por el usuario.")
        estado_final = "PAUSADO_POR_USUARIO"

    finally:
        pbar.close()
        for c in clients_pool.values():
            c.close()

    extras: dict[str, Any] = {}
    if active_key_ids:
        extras["next_free_key_id"] = active_key_ids[current_key_idx % len(active_key_ids)]
    escribir_estado(estado_file_path, estado_final, extras=extras)

    print("\n--- Resumen Final ---")
    print(f"Lote: {lote_nombre}")
    print(f"Carpeta de reportes: {reportes_dir.resolve()}")
    print(f"OK: {exito_count}")
    print(f"No encontrados: {nf_count}")
    print(f"Benignos/sin detección: {benign_count}")
    print(f"Errores: {error_count}")
    print("Registros:")
    print(f"  - {lista_exitosos_file.resolve()}")
    print(f"  - {lista_no_encontrados_file.resolve()}")
    print(f"  - {lista_benignos_file.resolve()}")
    print(f"Log: {log_file_path.resolve()}")


if __name__ == "__main__":
    main()
