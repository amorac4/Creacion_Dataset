#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
Compara etiquetas de AVClass contra los CSV curados en Dataset_V1/csv.
"""

from __future__ import annotations

import argparse
import csv
import json
from collections import Counter
from pathlib import Path
from typing import Any


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Compara familias AVClass contra Dataset_V1/csv.")
    parser.add_argument("--avclass-labels", required=True, type=Path, help="Salida default de avclass.")
    parser.add_argument("--dataset-csv-dir", default=Path("Dataset_V1/csv"), type=Path)
    parser.add_argument(
        "--output-csv",
        default=Path("experimentos_avclass/results/comparacion_avclass_vs_dataset.csv"),
        type=Path,
    )
    parser.add_argument(
        "--summary-json",
        default=Path("experimentos_avclass/results/comparacion_avclass_vs_dataset.json"),
        type=Path,
    )
    return parser.parse_args()


def normalize_family(value: str) -> str:
    text = str(value or "").strip().lower()
    if not text or text.startswith("singleton:"):
        return "sin_inferir"
    return text


def load_avclass_labels(path: Path) -> dict[str, str]:
    labels: dict[str, str] = {}
    with path.open("r", encoding="utf-8-sig") as handle:
        for line in handle:
            line = line.strip()
            if not line or line.startswith("#"):
                continue
            parts = line.split()
            if len(parts) < 2:
                continue
            hash_md5 = parts[0].strip().lower()
            family = normalize_family(parts[1])
            labels[hash_md5] = family
    return labels


def load_dataset_labels(csv_dir: Path) -> dict[str, dict[str, Any]]:
    labels: dict[str, dict[str, Any]] = {}
    for csv_path in sorted(csv_dir.glob("*.csv")):
        fallback_family = csv_path.stem.lower()
        with csv_path.open("r", encoding="utf-8-sig", newline="") as handle:
            reader = csv.DictReader(handle)
            for row in reader:
                hash_md5 = str(row.get("hash_md5") or "").strip().lower()
                if not hash_md5:
                    continue
                family = normalize_family(str(row.get("familia_probable") or fallback_family))
                labels[hash_md5] = {
                    "familia_dataset": family,
                    "csv_origen": str(csv_path),
                    "lote_origen": row.get("lote_origen", ""),
                    "tipo_probable": row.get("tipo_probable", ""),
                    "detection_percent": row.get("detection_percent", ""),
                    "dia_creacion_archivo": row.get("dia_creacion_archivo", ""),
                }
    return labels


def main() -> None:
    args = parse_args()
    avclass_labels = load_avclass_labels(args.avclass_labels)
    dataset_labels = load_dataset_labels(args.dataset_csv_dir)

    args.output_csv.parent.mkdir(parents=True, exist_ok=True)
    args.summary_json.parent.mkdir(parents=True, exist_ok=True)

    common_hashes = sorted(set(avclass_labels) & set(dataset_labels))
    rows: list[dict[str, Any]] = []
    for hash_md5 in common_hashes:
        dataset_row = dataset_labels[hash_md5]
        dataset_family = dataset_row["familia_dataset"]
        avclass_family = avclass_labels[hash_md5]
        rows.append(
            {
                "hash_md5": hash_md5,
                "familia_dataset": dataset_family,
                "familia_avclass": avclass_family,
                "coincide": dataset_family == avclass_family,
                "avclass_sin_inferir": avclass_family == "sin_inferir",
                **dataset_row,
            }
        )

    fieldnames = [
        "hash_md5",
        "familia_dataset",
        "familia_avclass",
        "coincide",
        "avclass_sin_inferir",
        "csv_origen",
        "lote_origen",
        "tipo_probable",
        "detection_percent",
        "dia_creacion_archivo",
    ]
    with args.output_csv.open("w", encoding="utf-8", newline="") as handle:
        writer = csv.DictWriter(handle, fieldnames=fieldnames)
        writer.writeheader()
        writer.writerows(rows)

    total_common = len(rows)
    matches = sum(1 for row in rows if row["coincide"])
    summary = {
        "avclass_labels": len(avclass_labels),
        "dataset_labels": len(dataset_labels),
        "hashes_en_comun": total_common,
        "coincidencias": matches,
        "desacuerdos": total_common - matches,
        "porcentaje_coincidencia": round((matches / total_common) * 100, 2) if total_common else 0,
        "avclass_sin_inferir": sum(1 for row in rows if row["avclass_sin_inferir"]),
        "top_familias_dataset": Counter(row["familia_dataset"] for row in rows).most_common(20),
        "top_familias_avclass": Counter(row["familia_avclass"] for row in rows).most_common(20),
        "top_desacuerdos": Counter(
            f"{row['familia_dataset']} -> {row['familia_avclass']}"
            for row in rows
            if not row["coincide"]
        ).most_common(30),
    }
    args.summary_json.write_text(json.dumps(summary, ensure_ascii=False, indent=2), encoding="utf-8")

    print(f"Hashes AVClass: {len(avclass_labels):,}")
    print(f"Hashes Dataset_V1: {len(dataset_labels):,}")
    print(f"Hashes en comun: {total_common:,}")
    print(f"Coincidencias: {matches:,}")
    print(f"Desacuerdos: {total_common - matches:,}")
    print(f"CSV: {args.output_csv}")
    print(f"Resumen: {args.summary_json}")


if __name__ == "__main__":
    main()

