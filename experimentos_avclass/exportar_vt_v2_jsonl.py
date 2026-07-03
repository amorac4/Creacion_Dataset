#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
Exporta reportes VirusShare locales a JSONL compatible con AVClass.

AVClass entiende reportes VirusTotal v2 con un objeto JSON por linea. Los JSON
locales tienen el bloque VT dentro de la clave "virustotal"; este script extrae
ese bloque y conserva hashes/metadatos minimos en el nivel superior.
"""

from __future__ import annotations

import argparse
import csv
import json
from pathlib import Path
from typing import Any, Iterable


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Exporta reportes locales a JSONL VirusTotal v2 para AVClass.")
    group = parser.add_mutually_exclusive_group(required=True)
    group.add_argument("--reports-root", type=Path, help="Raiz con VirusShare_*/reportes/reporte.")
    group.add_argument("--reports-dir", type=Path, help="Carpeta concreta con JSON de reportes.")
    parser.add_argument(
        "--output",
        default=Path("experimentos_avclass/data/virusshare_vtv2.jsonl"),
        type=Path,
        help="JSONL de salida para avclass -f.",
    )
    parser.add_argument(
        "--dataset-csv-dir",
        type=Path,
        help="Si se indica, exporta solo hashes presentes en los CSV curados de Dataset_V1.",
    )
    parser.add_argument("--limit", type=int, default=0, help="Limita la cantidad de reportes exportados.")
    parser.add_argument("--include-empty", action="store_true", help="Incluye reportes sin scans VT.")
    return parser.parse_args()


def iter_report_paths(args: argparse.Namespace) -> Iterable[Path]:
    if args.reports_root:
        yield from sorted(args.reports_root.glob("VirusShare_*/reportes/reporte/*.json"))
    else:
        yield from sorted(args.reports_dir.glob("*.json"))


def load_dataset_hashes(csv_dir: Path | None) -> set[str] | None:
    if csv_dir is None:
        return None
    hashes: set[str] = set()
    for csv_path in sorted(csv_dir.glob("*.csv")):
        with csv_path.open("r", encoding="utf-8-sig", newline="") as handle:
            reader = csv.DictReader(handle)
            for row in reader:
                hash_md5 = str(row.get("hash_md5") or "").strip().lower()
                if hash_md5:
                    hashes.add(hash_md5)
    return hashes


def detected_count(scans: dict[str, Any]) -> int:
    return sum(1 for details in scans.values() if isinstance(details, dict) and details.get("detected") is True)


def normalize_scans(scans: dict[str, Any]) -> dict[str, Any]:
    normalized: dict[str, Any] = {}
    for engine, details in scans.items():
        if not isinstance(details, dict):
            continue
        clean_details = dict(details)
        result = clean_details.get("result")
        if result is not None and not isinstance(result, str):
            clean_details["result"] = str(result)
        normalized[str(engine)] = clean_details
    return normalized


def to_vt_v2_record(path: Path, report: dict[str, Any]) -> dict[str, Any] | None:
    vt = report.get("virustotal") or {}
    if not isinstance(vt, dict):
        vt = {}

    scans = vt.get("scans") or {}
    if not isinstance(scans, dict):
        scans = {}
    scans = normalize_scans(scans)

    hash_md5 = str(report.get("md5") or report.get("data_structure", {}).get("hash") or path.stem).lower()
    positives = vt.get("positives")
    total = vt.get("total")

    return {
        "response_code": vt.get("response_code", 1 if scans else 0),
        "resource": hash_md5,
        "md5": hash_md5,
        "sha1": report.get("sha1", ""),
        "sha256": report.get("sha256", ""),
        "scan_date": vt.get("scan_date", ""),
        "permalink": vt.get("permalink", ""),
        "positives": int(positives) if str(positives or "").isdigit() else detected_count(scans),
        "total": int(total) if str(total or "").isdigit() else len(scans),
        "scans": scans,
    }


def main() -> None:
    args = parse_args()
    args.output.parent.mkdir(parents=True, exist_ok=True)
    allowed_hashes = load_dataset_hashes(args.dataset_csv_dir)

    read_count = 0
    written_count = 0
    skipped_not_selected = 0
    skipped_empty = 0
    skipped_errors = 0

    with args.output.open("w", encoding="utf-8", newline="\n") as handle:
        for path in iter_report_paths(args):
            if args.limit and written_count >= args.limit:
                break
            read_count += 1
            if allowed_hashes is not None and path.stem.lower() not in allowed_hashes:
                skipped_not_selected += 1
                continue
            try:
                report = json.loads(path.read_text(encoding="utf-8-sig"))
            except (OSError, json.JSONDecodeError):
                skipped_errors += 1
                continue

            record = to_vt_v2_record(path, report)
            scans = record.get("scans") if record else {}
            if not args.include_empty and not scans:
                skipped_empty += 1
                continue

            handle.write(json.dumps(record, ensure_ascii=False, sort_keys=True))
            handle.write("\n")
            written_count += 1

    print(f"Reportes leidos: {read_count:,}")
    print(f"Registros escritos: {written_count:,}")
    print(f"Omitidos fuera de seleccion: {skipped_not_selected:,}")
    print(f"Omitidos sin scans: {skipped_empty:,}")
    print(f"Omitidos por error JSON/IO: {skipped_errors:,}")
    print(f"Salida: {args.output}")


if __name__ == "__main__":
    main()
