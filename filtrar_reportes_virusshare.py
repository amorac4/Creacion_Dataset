#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
Extrae campos utiles de reportes VirusShare/VirusTotal y genera CSV filtrables.

El script no consulta APIs. Solo procesa JSON ya descargados en:

    clasificacion/<lote>/reportes/reporte/*.json

La familia y el tipo probable se infieren por votos a partir de las etiquetas de
deteccion de los motores. Son heuristicas auditables, no verdad absoluta; por
eso tambien se exportan las detecciones crudas por motor.
"""

from __future__ import annotations

import argparse
import csv
import datetime as dt
import json
import re
from collections import Counter, defaultdict
from pathlib import Path
from typing import Any, Iterable


DEFAULT_REPORTS_DIR = r"clasificacion\VirusShare_00499\reportes\reporte"
DEFAULT_OUTPUT_DIR = r"dataset_filtrado\VirusShare_00499"

LABEL_SPLIT_RE = re.compile(r"[^a-zA-Z0-9]+")
HEX_RE = re.compile(r"^[0-9a-f]{8,}$", re.IGNORECASE)

TYPE_TOKEN_MAP = {
    "adware": "adware",
    "banker": "banker",
    "banking": "banker",
    "backdoor": "backdoor",
    "bot": "bot",
    "botnet": "bot",
    "coinminer": "miner",
    "crypto": "miner",
    "cryptominer": "miner",
    "downloader": "downloader",
    "dropper": "dropper",
    "exploit": "exploit",
    "hacktool": "hacktool",
    "infostealer": "stealer",
    "keylogger": "keylogger",
    "miner": "miner",
    "phish": "phishing",
    "phishing": "phishing",
    "pua": "pua",
    "pup": "pua",
    "ransom": "ransomware",
    "ransomware": "ransomware",
    "redirector": "redirector",
    "riskware": "riskware",
    "rootkit": "rootkit",
    "spyware": "spyware",
    "stealer": "stealer",
    "trojan": "trojan",
    "virus": "virus",
    "worm": "worm",
}

TYPE_PRIORITY = [
    "ransomware",
    "stealer",
    "banker",
    "keylogger",
    "rootkit",
    "backdoor",
    "bot",
    "worm",
    "virus",
    "downloader",
    "dropper",
    "exploit",
    "miner",
    "phishing",
    "redirector",
    "spyware",
    "adware",
    "hacktool",
    "riskware",
    "pua",
    "trojan",
]

NOISE_TOKENS = {
    "a",
    "ai",
    "agent",
    "agen",
    "android",
    "asm",
    "asmalw",
    "asmalwrg",
    "asmalws",
    "bat",
    "behaveslike",
    "clean",
    "classic",
    "cloud",
    "confidence",
    "cve",
    "detect",
    "detected",
    "doc",
    "docx",
    "eldorado",
    "file",
    "gen",
    "generic",
    "genericdet",
    "generickd",
    "generickdz",
    "genericrx",
    "heur",
    "heuristic",
    "high",
    "html",
    "infected",
    "js",
    "javascript",
    "linux",
    "low",
    "macos",
    "malicious",
    "malware",
    "ml",
    "msil",
    "other",
    "outbreak",
    "packed",
    "pdf",
    "possible",
    "powershell",
    "save",
    "script",
    "score",
    "spam",
    "suspicious",
    "tr",
    "troj",
    "unsafe",
    "variant",
    "vba",
    "vbs",
    "w32",
    "w64",
    "webpage",
    "win",
    "win32",
    "win64",
    "windows",
    "x64",
    "x86",
    "xml",
}

NOISE_TOKENS.update(TYPE_TOKEN_MAP.keys())

NOISE_PREFIXES = (
    "adware",
    "backdoor",
    "downloader",
    "dropper",
    "exploit",
    "generic",
    "heur",
    "malware",
    "phish",
    "ransom",
    "spyware",
    "trojan",
    "virus",
    "worm",
)


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Filtra y resume reportes JSON de VirusShare."
    )
    parser.add_argument(
        "--reports-dir",
        default=DEFAULT_REPORTS_DIR,
        help="Carpeta con los JSON de reportes.",
    )
    parser.add_argument(
        "--output-dir",
        default=DEFAULT_OUTPUT_DIR,
        help="Carpeta donde se escribiran los CSV.",
    )
    parser.add_argument(
        "--min-positives",
        type=int,
        default=1,
        help="Minimo de motores positivos para incluir una muestra.",
    )
    parser.add_argument(
        "--min-ratio",
        type=float,
        default=0.0,
        help="Minimo de positivos/total para incluir una muestra. Ejemplo: 0.2",
    )
    parser.add_argument(
        "--date-from",
        help="Fecha minima de escaneo VT en formato YYYY-MM-DD.",
    )
    parser.add_argument(
        "--date-to",
        help="Fecha maxima de escaneo VT en formato YYYY-MM-DD.",
    )
    parser.add_argument(
        "--family",
        help="Filtra por familia probable, comparacion parcial case-insensitive.",
    )
    parser.add_argument(
        "--type",
        dest="malware_type",
        help="Filtra por tipo probable, comparacion parcial case-insensitive.",
    )
    parser.add_argument(
        "--limit",
        type=int,
        default=0,
        help="Procesa solo N reportes. Util para pruebas.",
    )
    parser.add_argument(
        "--family-score-threshold",
        type=float,
        default=0.10,
        help="Score minimo para aceptar familia_probable. El top completo queda en familias_top.",
    )
    return parser.parse_args()


def parse_datetime(value: Any) -> dt.datetime | None:
    if not value:
        return None
    text = str(value).strip()
    for candidate in (text, text.replace("Z", "+00:00")):
        try:
            parsed = dt.datetime.fromisoformat(candidate)
            if parsed.tzinfo is None:
                parsed = parsed.replace(tzinfo=dt.timezone.utc)
            return parsed.astimezone(dt.timezone.utc)
        except ValueError:
            pass
    try:
        parsed = dt.datetime.strptime(text, "%Y-%m-%d %H:%M:%S")
        return parsed.replace(tzinfo=dt.timezone.utc)
    except ValueError:
        return None


def parse_date_filter(value: str | None) -> dt.date | None:
    if not value:
        return None
    return dt.date.fromisoformat(value)


def fmt_datetime(value: dt.datetime | None) -> str:
    return value.isoformat() if value else ""


def fmt_date(value: dt.datetime | None) -> str:
    return value.date().isoformat() if value else ""


def iter_report_paths(reports_dir: Path, limit: int = 0) -> Iterable[Path]:
    count = 0
    for path in sorted(reports_dir.glob("*.json")):
        yield path
        count += 1
        if limit and count >= limit:
            break


def label_tokens(label: str) -> list[str]:
    return [token.lower() for token in LABEL_SPLIT_RE.split(label) if token]


def is_hashish(token: str) -> bool:
    return any(ch.isdigit() for ch in token) or bool(HEX_RE.fullmatch(token))


def is_family_noise(token: str) -> bool:
    return token in NOISE_TOKENS or token.startswith(NOISE_PREFIXES)


def extract_detections(report: dict[str, Any]) -> list[dict[str, str]]:
    scans = ((report.get("virustotal") or {}).get("scans") or {})
    detections: list[dict[str, str]] = []
    if not isinstance(scans, dict):
        return detections

    for engine, details in scans.items():
        if not isinstance(details, dict):
            continue
        result = details.get("result")
        if details.get("detected") is True and result:
            detections.append({"engine": str(engine), "result": str(result)})
    return detections


def infer_type_and_family(detections: list[dict[str, str]], family_score_threshold: float) -> dict[str, Any]:
    type_votes: Counter[str] = Counter()
    family_votes: Counter[str] = Counter()

    for detection in detections:
        tokens = label_tokens(detection["result"])
        label_types = {TYPE_TOKEN_MAP[token] for token in tokens if token in TYPE_TOKEN_MAP}
        type_votes.update(label_types)

        for token in tokens:
            if len(token) < 4:
                continue
            if is_family_noise(token):
                continue
            if is_hashish(token):
                continue
            family_votes[token] += 1

    tipo = ""
    if type_votes:
        tipo = sorted(type_votes, key=lambda key: (-type_votes[key], TYPE_PRIORITY.index(key) if key in TYPE_PRIORITY else 999, key))[0]

    detecciones = max(len(detections), 1)
    familia = ""
    familia_score = 0.0
    if family_votes:
        familia = family_votes.most_common(1)[0][0]
        familia_score = family_votes[familia] / detecciones
        if familia_score < family_score_threshold:
            familia = ""
            familia_score = 0.0

    return {
        "tipo_probable": tipo,
        "tipo_score": round(type_votes[tipo] / detecciones, 4) if tipo else 0,
        "familia_probable": familia,
        "familia_score": round(familia_score, 4),
        "familias_top": ";".join(f"{name}:{count}" for name, count in family_votes.most_common(5)),
        "tipos_top": ";".join(f"{name}:{count}" for name, count in type_votes.most_common(5)),
    }


def detection_ratio(positives: Any, total: Any) -> float:
    try:
        positives_int = int(positives or 0)
        total_int = int(total or 0)
    except (TypeError, ValueError):
        return 0.0
    if total_int <= 0:
        return 0.0
    return positives_int / total_int


def row_from_report(path: Path, report: dict[str, Any], family_score_threshold: float) -> tuple[dict[str, Any], list[dict[str, str]]]:
    vt = report.get("virustotal") or {}
    detections = extract_detections(report)
    inferred = infer_type_and_family(detections, family_score_threshold)
    scan_dt = parse_datetime(vt.get("scan_date"))
    added_dt = parse_datetime(report.get("added_timestamp"))
    positives = int(vt.get("positives") or 0)
    total = int(vt.get("total") or 0)
    ratio = detection_ratio(positives, total)

    row = {
        "hash_md5": report.get("md5") or report.get("data_structure", {}).get("hash") or path.stem,
        "sha1": report.get("sha1", ""),
        "sha256": report.get("sha256", ""),
        "extension": report.get("extension", ""),
        "filetype": report.get("filetype", ""),
        "mimetype": report.get("mimetype", ""),
        "size": report.get("size", ""),
        "fecha_escaneo_vt": fmt_datetime(scan_dt),
        "dia_escaneo_vt": fmt_date(scan_dt),
        "fecha_agregado_virusshare": fmt_datetime(added_dt),
        "dia_agregado_virusshare": fmt_date(added_dt),
        "vt_positives": positives,
        "vt_total": total,
        "detection_ratio": round(ratio, 4),
        "motores_detectaron": len(detections),
        "familia_probable": inferred["familia_probable"],
        "familia_score": inferred["familia_score"],
        "familias_top": inferred["familias_top"],
        "tipo_probable": inferred["tipo_probable"],
        "tipo_score": inferred["tipo_score"],
        "tipos_top": inferred["tipos_top"],
        "detecciones_top": ";".join(d["result"] for d in detections[:10]),
        "permalink": vt.get("permalink", ""),
        "reporte_path": str(path),
    }
    return row, detections


def passes_filters(row: dict[str, Any], args: argparse.Namespace, date_from: dt.date | None, date_to: dt.date | None) -> bool:
    if row["vt_positives"] < args.min_positives:
        return False
    if row["detection_ratio"] < args.min_ratio:
        return False

    scan_day = dt.date.fromisoformat(row["dia_escaneo_vt"]) if row["dia_escaneo_vt"] else None
    if date_from and (scan_day is None or scan_day < date_from):
        return False
    if date_to and (scan_day is None or scan_day > date_to):
        return False

    if args.family and args.family.lower() not in row["familia_probable"].lower():
        return False
    if args.malware_type and args.malware_type.lower() not in row["tipo_probable"].lower():
        return False
    return True


def write_csv(path: Path, rows: list[dict[str, Any]], fieldnames: list[str]) -> None:
    with path.open("w", newline="", encoding="utf-8") as file:
        writer = csv.DictWriter(file, fieldnames=fieldnames)
        writer.writeheader()
        writer.writerows(rows)


def counter_rows(counter: Counter[str], key_name: str) -> list[dict[str, Any]]:
    return [{key_name: key or "sin_inferir", "muestras": value} for key, value in counter.most_common()]


def main() -> None:
    args = parse_args()
    reports_dir = Path(args.reports_dir)
    output_dir = Path(args.output_dir)
    output_dir.mkdir(parents=True, exist_ok=True)

    if not reports_dir.is_dir():
        raise SystemExit(f"No existe la carpeta de reportes: {reports_dir}")

    date_from = parse_date_filter(args.date_from)
    date_to = parse_date_filter(args.date_to)

    familias: Counter[str] = Counter()
    tipos: Counter[str] = Counter()
    dias: Counter[str] = Counter()
    row_count = 0
    detection_count = 0
    corrupt_count = 0
    main_fields = [
        "hash_md5",
        "sha1",
        "sha256",
        "extension",
        "filetype",
        "mimetype",
        "size",
        "fecha_escaneo_vt",
        "dia_escaneo_vt",
        "fecha_agregado_virusshare",
        "dia_agregado_virusshare",
        "vt_positives",
        "vt_total",
        "detection_ratio",
        "motores_detectaron",
        "familia_probable",
        "familia_score",
        "familias_top",
        "tipo_probable",
        "tipo_score",
        "tipos_top",
        "detecciones_top",
        "permalink",
        "reporte_path",
    ]
    detection_fields = ["hash_md5", "engine", "result", "familia_probable", "tipo_probable", "fecha_escaneo_vt"]
    corrupt_fields = ["reporte_path", "error"]

    with (
        (output_dir / "muestras_filtradas.csv").open("w", newline="", encoding="utf-8") as main_file,
        (output_dir / "detecciones_por_motor.csv").open("w", newline="", encoding="utf-8") as detection_file,
        (output_dir / "reportes_corruptos.csv").open("w", newline="", encoding="utf-8") as corrupt_file,
    ):
        main_writer = csv.DictWriter(main_file, fieldnames=main_fields)
        detection_writer = csv.DictWriter(detection_file, fieldnames=detection_fields)
        corrupt_writer = csv.DictWriter(corrupt_file, fieldnames=corrupt_fields)
        main_writer.writeheader()
        detection_writer.writeheader()
        corrupt_writer.writeheader()

        for path in iter_report_paths(reports_dir, args.limit):
            try:
                report = json.loads(path.read_text(encoding="utf-8"))
            except Exception as exc:
                corrupt_writer.writerow({"reporte_path": str(path), "error": str(exc)})
                corrupt_count += 1
                continue

            if "error_details" in report:
                continue

            row, detections = row_from_report(path, report, args.family_score_threshold)
            if not passes_filters(row, args, date_from, date_to):
                continue

            main_writer.writerow(row)
            row_count += 1
            familias[row["familia_probable"]] += 1
            tipos[row["tipo_probable"]] += 1
            dias[row["dia_escaneo_vt"]] += 1

            for detection in detections:
                detection_writer.writerow(
                    {
                        "hash_md5": row["hash_md5"],
                        "engine": detection["engine"],
                        "result": detection["result"],
                        "familia_probable": row["familia_probable"],
                        "tipo_probable": row["tipo_probable"],
                        "fecha_escaneo_vt": row["fecha_escaneo_vt"],
                    }
                )
                detection_count += 1

    write_csv(output_dir / "resumen_familias.csv", counter_rows(familias, "familia_probable"), ["familia_probable", "muestras"])
    write_csv(output_dir / "resumen_tipos.csv", counter_rows(tipos, "tipo_probable"), ["tipo_probable", "muestras"])
    write_csv(output_dir / "resumen_dias.csv", counter_rows(dias, "dia_escaneo_vt"), ["dia_escaneo_vt", "muestras"])

    print(f"Reportes filtrados: {row_count}")
    print(f"Detecciones por motor: {detection_count}")
    print(f"Reportes corruptos: {corrupt_count}")
    print(f"Salida: {output_dir.resolve()}")


if __name__ == "__main__":
    main()
