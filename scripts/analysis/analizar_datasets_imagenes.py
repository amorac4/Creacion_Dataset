#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
Analizar Datasets de Imagenes
=============================

Lee `manifest_imagenes.csv` generado por separar_dataset_imagenes.py y produce
un analisis detallado por cada dataset (balance + algoritmo):

- cobertura de imagenes esperadas/encontradas;
- conteos por split y familia;
- distribucion porcentual train/val/test;
- faltantes por split/familia/algoritmo;
- fechas min/max por split;
- validacion de archivos realmente presentes en disco.
"""

from __future__ import annotations

import argparse
import csv
import json
import statistics
from collections import Counter, defaultdict
from dataclasses import dataclass
from pathlib import Path
from typing import Any, Iterable


DEFAULT_MANIFEST = Path("outputs/datasets_imagenes/manifest_imagenes.csv")
DEFAULT_OUTPUT_DIR = Path("outputs/analisis_datasets_imagenes")
AVAILABLE_STATUSES = {"copiado", "ya_existia"}
SPLIT_ORDER = {"train": 0, "validacion": 1, "val": 1, "test": 2}


@dataclass(frozen=True)
class ManifestRow:
    dataset_folder: str
    actual_dataset_folder: str
    balance_sheet: str
    balance_name: str
    algorithm: str
    algorithm_source: str
    split: str
    family: str
    hash_md5: str
    creation_day: str
    source_image: str
    destination: str
    actual_destination: str
    status: str
    destination_exists: bool


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Genera analisis detallado por dataset de imagenes."
    )
    parser.add_argument("--manifest", type=Path, default=DEFAULT_MANIFEST, help="Manifest de imagenes.")
    parser.add_argument("--output-dir", type=Path, default=DEFAULT_OUTPUT_DIR, help="Directorio de reportes.")
    parser.add_argument(
        "--no-file-check",
        action="store_true",
        help="No valida si los archivos destino existen en disco.",
    )
    return parser.parse_args()


def dataset_folder_from_destination(destination: str) -> str:
    path = Path(destination)
    parts = path.parts
    if len(parts) >= 4:
        return parts[-4]
    return "sin_dataset"


def dataset_code(folder: str) -> str:
    text = str(folder or "")
    text = text.split("__", 1)[0]
    text = text.split("_", 1)[0]
    return text


def actual_folder_map(output_dir: Path) -> dict[str, str]:
    folders: dict[str, str] = {}
    if not output_dir.exists():
        return folders
    for path in sorted(output_dir.iterdir()):
        if path.is_dir():
            folders.setdefault(dataset_code(path.name), path.name)
    return folders


def resolve_destination(path_text: str, manifest_path: Path) -> Path:
    path = Path(path_text)
    if path.is_absolute():
        return path
    return (manifest_path.parent.parent.parent / path).resolve()


def resolve_actual_destination(path_text: str, manifest_path: Path, folders_by_code: dict[str, str]) -> tuple[str, Path]:
    path = Path(path_text)
    parts = path.parts
    if len(parts) < 4:
        resolved = resolve_destination(path_text, manifest_path)
        return dataset_folder_from_destination(path_text), resolved

    manifest_folder = parts[-4]
    actual_folder = folders_by_code.get(dataset_code(manifest_folder), manifest_folder)
    actual_path = manifest_path.parent / actual_folder / parts[-3] / parts[-2] / parts[-1]
    return actual_folder, actual_path.resolve()


def load_manifest(path: Path, check_files: bool = True) -> list[ManifestRow]:
    if not path.exists():
        raise SystemExit(f"No existe el manifest: {path}")

    rows: list[ManifestRow] = []
    folders_by_code = actual_folder_map(path.parent)
    with path.open("r", encoding="utf-8-sig", newline="") as handle:
        reader = csv.DictReader(handle)
        for raw in reader:
            destination = raw.get("destination", "")
            actual_folder, actual_destination = resolve_actual_destination(destination, path, folders_by_code)
            rows.append(
                ManifestRow(
                    dataset_folder=dataset_folder_from_destination(destination),
                    actual_dataset_folder=actual_folder,
                    balance_sheet=str(raw.get("balance_sheet", "")),
                    balance_name=str(raw.get("balance_name", "")),
                    algorithm=str(raw.get("algorithm", "")),
                    algorithm_source=str(raw.get("algorithm_source", "")),
                    split=str(raw.get("split", "")),
                    family=str(raw.get("family", "")),
                    hash_md5=str(raw.get("hash_md5", "")),
                    creation_day=str(raw.get("creation_day", "")),
                    source_image=str(raw.get("source_image", "")),
                    destination=destination,
                    actual_destination=str(actual_destination),
                    status=str(raw.get("status", "")),
                    destination_exists=actual_destination.exists() if check_files else False,
                )
            )
    return rows


def percent(part: int, total: int) -> float:
    return round((part / total) * 100, 4) if total else 0.0


def ratio(maximum: int, minimum: int) -> float:
    return round(maximum / minimum, 4) if minimum else 0.0


def is_available(row: ManifestRow) -> bool:
    return row.status in AVAILABLE_STATUSES


def key_sort(value: tuple[Any, ...]) -> tuple[Any, ...]:
    converted: list[Any] = []
    for item in value:
        text = str(item)
        converted.append(SPLIT_ORDER.get(text, text))
    return tuple(converted)


def write_csv(path: Path, rows: list[dict[str, Any]], fieldnames: list[str] | None = None) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    if fieldnames is None:
        fieldnames = []
        for row in rows:
            for key in row:
                if key not in fieldnames:
                    fieldnames.append(key)
    with path.open("w", encoding="utf-8", newline="") as handle:
        writer = csv.DictWriter(handle, fieldnames=fieldnames)
        writer.writeheader()
        writer.writerows(rows)


def count_values(rows: Iterable[ManifestRow], *fields: str) -> Counter[tuple[str, ...]]:
    counter: Counter[tuple[str, ...]] = Counter()
    for row in rows:
        counter[tuple(str(getattr(row, field)) for field in fields)] += 1
    return counter


def status_summary(rows: list[ManifestRow]) -> Counter[str]:
    return Counter(row.status for row in rows)


def summarize_dataset(rows: list[ManifestRow]) -> dict[str, Any]:
    expected = len(rows)
    available = sum(1 for row in rows if is_available(row))
    missing = sum(1 for row in rows if row.status == "imagen_no_encontrada")
    destination_exists = sum(1 for row in rows if row.destination_exists)
    unique_hashes = len({row.hash_md5 for row in rows})
    splits = sorted({row.split for row in rows}, key=lambda item: SPLIT_ORDER.get(item, 99))
    families = sorted({row.family for row in rows})
    family_totals = Counter(row.family for row in rows)
    split_totals = Counter(row.split for row in rows)
    family_values = list(family_totals.values())
    split_values = list(split_totals.values())
    dates = [row.creation_day for row in rows if row.creation_day]

    return {
        "dataset_folder": rows[0].dataset_folder if rows else "",
        "actual_dataset_folder": rows[0].actual_dataset_folder if rows else "",
        "balance_sheet": rows[0].balance_sheet if rows else "",
        "balance_name": rows[0].balance_name if rows else "",
        "algorithm": rows[0].algorithm if rows else "",
        "algorithm_source": rows[0].algorithm_source if rows else "",
        "expected_images": expected,
        "available_images": available,
        "missing_images": missing,
        "coverage_percent": percent(available, expected),
        "destination_exists": destination_exists,
        "destination_exists_percent": percent(destination_exists, expected),
        "unique_hashes": unique_hashes,
        "splits": ",".join(splits),
        "families": len(families),
        "family_min": min(family_values) if family_values else 0,
        "family_max": max(family_values) if family_values else 0,
        "family_mean": round(statistics.mean(family_values), 4) if family_values else 0,
        "family_max_min_ratio": ratio(max(family_values), min(family_values)) if family_values else 0,
        "split_min": min(split_values) if split_values else 0,
        "split_max": max(split_values) if split_values else 0,
        "date_min": min(dates) if dates else "",
        "date_max": max(dates) if dates else "",
        **{f"status_{status}": count for status, count in sorted(status_summary(rows).items())},
    }


def build_analysis(rows: list[ManifestRow]) -> dict[str, Any]:
    by_dataset: dict[str, list[ManifestRow]] = defaultdict(list)
    for row in rows:
        by_dataset[row.dataset_folder].append(row)

    dataset_rows = [summarize_dataset(dataset_rows) for _, dataset_rows in sorted(by_dataset.items())]

    split_family_rows: list[dict[str, Any]] = []
    missing_rows: list[dict[str, Any]] = []
    split_rows: list[dict[str, Any]] = []
    duplicate_rows: list[dict[str, Any]] = []

    for dataset_folder, dataset_rows_raw in sorted(by_dataset.items()):
        total_dataset = len(dataset_rows_raw)
        split_counts = count_values(dataset_rows_raw, "split")
        for (split,), count in sorted(split_counts.items(), key=lambda item: key_sort(item[0])):
            available = sum(1 for row in dataset_rows_raw if row.split == split and is_available(row))
            missing = sum(1 for row in dataset_rows_raw if row.split == split and row.status == "imagen_no_encontrada")
            split_rows.append(
                {
                    "dataset_folder": dataset_folder,
                    "split": split,
                    "expected_images": count,
                    "available_images": available,
                    "missing_images": missing,
                    "coverage_percent": percent(available, count),
                    "split_percent_dataset": percent(count, total_dataset),
                }
            )

        for (split, family), count in sorted(
            count_values(dataset_rows_raw, "split", "family").items(),
            key=lambda item: key_sort(item[0]),
        ):
            available = sum(
                1
                for row in dataset_rows_raw
                if row.split == split and row.family == family and is_available(row)
            )
            missing = sum(
                1
                for row in dataset_rows_raw
                if row.split == split and row.family == family and row.status == "imagen_no_encontrada"
            )
            dates = [
                row.creation_day
                for row in dataset_rows_raw
                if row.split == split and row.family == family and row.creation_day
            ]
            split_family_rows.append(
                {
                    "dataset_folder": dataset_folder,
                    "balance_sheet": dataset_rows_raw[0].balance_sheet,
                    "algorithm": dataset_rows_raw[0].algorithm,
                    "split": split,
                    "family": family,
                    "expected_images": count,
                    "available_images": available,
                    "missing_images": missing,
                    "coverage_percent": percent(available, count),
                    "date_min": min(dates) if dates else "",
                    "date_max": max(dates) if dates else "",
                }
            )

        for row in dataset_rows_raw:
            if row.status == "imagen_no_encontrada":
                missing_rows.append(
                    {
                        "dataset_folder": row.dataset_folder,
                        "balance_sheet": row.balance_sheet,
                        "algorithm": row.algorithm,
                        "algorithm_source": row.algorithm_source,
                        "split": row.split,
                        "family": row.family,
                        "hash_md5": row.hash_md5,
                        "creation_day": row.creation_day,
                        "destination": row.destination,
                    }
                )

        hash_splits: defaultdict[str, set[str]] = defaultdict(set)
        for row in dataset_rows_raw:
            hash_splits[row.hash_md5].add(row.split)
        for hash_md5, splits in sorted(hash_splits.items()):
            if len(splits) > 1:
                duplicate_rows.append(
                    {
                        "dataset_folder": dataset_folder,
                        "hash_md5": hash_md5,
                        "splits": ",".join(sorted(splits, key=lambda item: SPLIT_ORDER.get(item, 99))),
                    }
                )

    global_summary = {
        "datasets": len(by_dataset),
        "manifest_rows": len(rows),
        "expected_images": len(rows),
        "available_images": sum(1 for row in rows if is_available(row)),
        "missing_images": sum(1 for row in rows if row.status == "imagen_no_encontrada"),
        "destination_exists": sum(1 for row in rows if row.destination_exists),
        "coverage_percent": percent(sum(1 for row in rows if is_available(row)), len(rows)),
        "destination_exists_percent": percent(sum(1 for row in rows if row.destination_exists), len(rows)),
        "unique_hashes_global": len({row.hash_md5 for row in rows}),
        "balances": sorted({row.balance_sheet for row in rows}),
        "algorithms": sorted({row.algorithm for row in rows}),
        "families": sorted({row.family for row in rows}),
        "statuses": dict(status_summary(rows)),
    }

    return {
        "global": global_summary,
        "datasets": dataset_rows,
        "splits": split_rows,
        "split_family": split_family_rows,
        "missing": missing_rows,
        "duplicates": duplicate_rows,
    }


def markdown_table(headers: list[str], rows: list[dict[str, Any]], limit: int | None = None) -> str:
    selected = rows[:limit] if limit else rows
    lines = [
        "| " + " | ".join(headers) + " |",
        "| " + " | ".join("---" for _ in headers) + " |",
    ]
    for row in selected:
        lines.append("| " + " | ".join(str(row.get(header, "")) for header in headers) + " |")
    if limit and len(rows) > limit:
        lines.append(f"| ... | ... | ... | ... |")
    return "\n".join(lines)


def write_markdown(path: Path, analysis: dict[str, Any]) -> None:
    global_summary = analysis["global"]
    datasets = analysis["datasets"]
    split_family = analysis["split_family"]
    missing = analysis["missing"]

    lines: list[str] = []
    lines.append("# Analisis de datasets de imagenes")
    lines.append("")
    lines.append("## Resumen global")
    lines.append("")
    lines.append(f"- Datasets analizados: {global_summary['datasets']}")
    lines.append(f"- Operaciones/imagenes esperadas: {global_summary['expected_images']:,}")
    lines.append(f"- Imagenes disponibles segun manifest: {global_summary['available_images']:,}")
    lines.append(f"- Imagenes faltantes: {global_summary['missing_images']:,}")
    lines.append(f"- Cobertura: {global_summary['coverage_percent']}%")
    lines.append(f"- Archivos existentes en disco: {global_summary['destination_exists']:,}")
    lines.append(f"- Hashes unicos globales: {global_summary['unique_hashes_global']:,}")
    lines.append("")
    lines.append("## Resumen por dataset")
    lines.append("")
    lines.append(
        markdown_table(
            [
                "dataset_folder",
                "balance_sheet",
                "algorithm",
                "expected_images",
                "available_images",
                "missing_images",
                "coverage_percent",
                "unique_hashes",
                "family_max_min_ratio",
            ],
            datasets,
        )
    )

    lines.append("")
    lines.append("## Detalle por dataset")
    for dataset in datasets:
        dataset_folder = dataset["dataset_folder"]
        lines.append("")
        lines.append(f"### {dataset_folder}")
        lines.append("")
        lines.append(f"- Balance: {dataset['balance_sheet']} - {dataset['balance_name']}")
        lines.append(f"- Algoritmo: {dataset['algorithm']} ({dataset['algorithm_source']})")
        lines.append(f"- Esperadas: {dataset['expected_images']:,}")
        lines.append(f"- Disponibles: {dataset['available_images']:,}")
        lines.append(f"- Faltantes: {dataset['missing_images']:,}")
        lines.append(f"- Cobertura: {dataset['coverage_percent']}%")
        lines.append(f"- Hashes unicos: {dataset['unique_hashes']:,}")
        lines.append(f"- Rango temporal: {dataset['date_min']} a {dataset['date_max']}")
        lines.append(f"- Relacion familia mayor/menor: {dataset['family_max_min_ratio']}")

        rows_for_dataset = [row for row in split_family if row["dataset_folder"] == dataset_folder]
        lines.append("")
        lines.append(
            markdown_table(
                [
                    "split",
                    "family",
                    "expected_images",
                    "available_images",
                    "missing_images",
                    "coverage_percent",
                    "date_min",
                    "date_max",
                ],
                rows_for_dataset,
            )
        )

        missing_for_dataset = [row for row in missing if row["dataset_folder"] == dataset_folder]
        if missing_for_dataset:
            lines.append("")
            lines.append("Faltantes principales:")
            lines.append(
                markdown_table(
                    ["split", "family", "hash_md5", "creation_day"],
                    missing_for_dataset,
                    limit=20,
                )
            )

    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text("\n".join(lines), encoding="utf-8")


def write_outputs(output_dir: Path, analysis: dict[str, Any]) -> dict[str, Path]:
    output_dir.mkdir(parents=True, exist_ok=True)
    paths = {
        "global_json": output_dir / "resumen_global.json",
        "datasets_csv": output_dir / "resumen_por_dataset.csv",
        "splits_csv": output_dir / "resumen_por_split.csv",
        "split_family_csv": output_dir / "conteos_split_familia.csv",
        "missing_csv": output_dir / "faltantes.csv",
        "duplicates_csv": output_dir / "hashes_en_multiples_splits.csv",
        "markdown": output_dir / "reporte_detallado.md",
    }
    paths["global_json"].write_text(json.dumps(analysis["global"], ensure_ascii=False, indent=2), encoding="utf-8")
    write_csv(paths["datasets_csv"], analysis["datasets"])
    write_csv(paths["splits_csv"], analysis["splits"])
    write_csv(paths["split_family_csv"], analysis["split_family"])
    write_csv(paths["missing_csv"], analysis["missing"])
    write_csv(paths["duplicates_csv"], analysis["duplicates"])
    write_markdown(paths["markdown"], analysis)
    return paths


def main() -> None:
    args = parse_args()
    rows = load_manifest(args.manifest, check_files=not args.no_file_check)
    analysis = build_analysis(rows)
    paths = write_outputs(args.output_dir, analysis)

    print("Analisis de datasets de imagenes terminado")
    print(f"Datasets analizados: {analysis['global']['datasets']:,}")
    print(f"Imagenes esperadas: {analysis['global']['expected_images']:,}")
    print(f"Imagenes disponibles: {analysis['global']['available_images']:,}")
    print(f"Imagenes faltantes: {analysis['global']['missing_images']:,}")
    print(f"Cobertura: {analysis['global']['coverage_percent']}%")
    print(f"Reporte Markdown: {paths['markdown'].resolve()}")
    print(f"Resumen por dataset: {paths['datasets_csv'].resolve()}")


if __name__ == "__main__":
    main()
