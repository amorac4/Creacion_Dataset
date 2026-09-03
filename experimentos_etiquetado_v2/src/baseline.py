#!/usr/bin/env python3
"""Genera la linea base reproducible del etiquetado local actual."""

from __future__ import annotations

import argparse
import csv
import json
import sqlite3
from pathlib import Path
from typing import Any


ROOT = Path(__file__).resolve().parents[2]
DEFAULT_CONFIG = ROOT / "experimentos_etiquetado_v2" / "config" / "baseline.json"


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--config", type=Path, default=DEFAULT_CONFIG)
    return parser.parse_args()


def resolve_from_root(value: str) -> Path:
    path = Path(value)
    return path if path.is_absolute() else ROOT / path


def percentage(value: int, total: int) -> float:
    return round(100.0 * value / total, 4) if total else 0.0


def query_count(conn: sqlite3.Connection, sql: str, parameters: tuple[Any, ...] = ()) -> int:
    return int(conn.execute(sql, parameters).fetchone()[0])


def main() -> None:
    args = parse_args()
    config = json.loads(args.config.read_text(encoding="utf-8-sig"))
    database = resolve_from_root(config["database"])
    output_dir = resolve_from_root(config["output_dir"])
    output_dir.mkdir(parents=True, exist_ok=True)

    if not database.exists():
        raise SystemExit(f"No existe la base: {database}")

    generic = sorted({str(value).strip().lower() for value in config["generic_or_non_family_labels"]})
    conn = sqlite3.connect(database)
    try:
        total = query_count(conn, "SELECT COUNT(*) FROM samples")
        uninferred = query_count(
            conn,
            "SELECT COUNT(*) FROM samples WHERE COALESCE(familia_probable, 'sin_inferir') = 'sin_inferir'",
        )
        placeholders = ",".join("?" for _ in generic)
        generic_count = query_count(
            conn,
            f"SELECT COUNT(*) FROM samples WHERE LOWER(familia_probable) IN ({placeholders})",
            tuple(generic),
        )
        usable = total - uninferred - generic_count

        confidence_rows = [
            {"confianza": row[0] or "sin_valor", "muestras": row[1], "porcentaje": percentage(row[1], total)}
            for row in conn.execute(
                "SELECT familia_confianza, COUNT(*) FROM samples GROUP BY familia_confianza ORDER BY COUNT(*) DESC"
            )
        ]
        family_rows = [
            {
                "familia": row[0] or "sin_inferir",
                "muestras": row[1],
                "porcentaje": percentage(row[1], total),
                "score_promedio": round(float(row[2] or 0), 6),
                "es_generica_o_no_familia": int((row[0] or "sin_inferir").lower() in generic),
            }
            for row in conn.execute(
                """
                SELECT familia_probable, COUNT(*), AVG(familia_score)
                FROM samples
                GROUP BY familia_probable
                ORDER BY COUNT(*) DESC, familia_probable
                """
            )
        ]
    finally:
        conn.close()

    summary = {
        "experiment": "label_v2_baseline",
        "database": str(database),
        "total_samples": total,
        "inferred_raw": total - uninferred,
        "raw_coverage_percent": percentage(total - uninferred, total),
        "uninferred": uninferred,
        "uninferred_percent": percentage(uninferred, total),
        "generic_or_non_family": generic_count,
        "generic_or_non_family_percent": percentage(generic_count, total),
        "provisionally_usable": usable,
        "provisionally_usable_coverage_percent": percentage(usable, total),
        "confidence": confidence_rows,
        "generic_or_non_family_labels": generic,
        "notes": [
            "La cobertura util es una linea base, no una estimacion de precision.",
            "Las etiquetas genericas se contabilizan separadas de sin_inferir.",
            "La base no contiene columnas AVClass; esta ejecucion evalua solo la heuristica local."
        ]
    }

    summary_path = output_dir / "summary.json"
    families_path = output_dir / "families.csv"
    summary_path.write_text(json.dumps(summary, ensure_ascii=False, indent=2), encoding="utf-8")
    with families_path.open("w", encoding="utf-8", newline="") as handle:
        writer = csv.DictWriter(handle, fieldnames=list(family_rows[0]))
        writer.writeheader()
        writer.writerows(family_rows)

    print("Experimento label_v2 baseline terminado")
    print(f"Muestras: {total:,}")
    print(f"Cobertura cruda: {summary['raw_coverage_percent']:.2f}%")
    print(f"Genericas/no-familia: {generic_count:,} ({summary['generic_or_non_family_percent']:.2f}%)")
    print(f"Cobertura provisional util: {summary['provisionally_usable_coverage_percent']:.2f}%")
    print(f"Resumen: {summary_path}")
    print(f"Familias: {families_path}")


if __name__ == "__main__":
    main()
