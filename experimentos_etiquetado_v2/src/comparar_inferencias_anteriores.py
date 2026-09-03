#!/usr/bin/env python3
"""Compara Dataset_V1 historico contra la inferencia local consolidada actual."""

from __future__ import annotations

import csv
import json
import sqlite3
from collections import Counter, defaultdict
from pathlib import Path


ROOT = Path(__file__).resolve().parents[2]
CSV_DIR = ROOT / "Dataset_V1" / "csv"
DATABASE = ROOT / "outputs" / "Analisis_de_Reportes_Todos.db"
OUTPUT_DIR = ROOT / "experimentos_etiquetado_v2" / "results" / "comparacion_anteriores"


def main() -> None:
    previous_rows: list[dict[str, str]] = []
    labels_by_hash: dict[str, set[str]] = defaultdict(set)
    for path in sorted(CSV_DIR.glob("*.csv")):
        with path.open("r", encoding="utf-8-sig", newline="") as handle:
            for row in csv.DictReader(handle):
                hash_md5 = str(row.get("hash_md5") or "").strip().lower()
                label = str(row.get("familia_probable") or path.stem).strip().lower()
                if hash_md5:
                    previous_rows.append({"hash_md5": hash_md5, "familia_anterior": label, "archivo": path.name})
                    labels_by_hash[hash_md5].add(label)

    conn = sqlite3.connect(DATABASE)
    current = {
        row[0]: {"familia_actual": row[1], "confianza_actual": row[2], "score_actual": row[3]}
        for row in conn.execute(
            "SELECT LOWER(hash_md5), LOWER(familia_probable), familia_confianza, familia_score FROM samples"
        )
    }
    conn.close()

    comparison: list[dict[str, object]] = []
    for hash_md5, previous_labels in sorted(labels_by_hash.items()):
        prior = sorted(previous_labels)
        now = current.get(hash_md5)
        previous = "|".join(prior)
        if len(prior) > 1:
            status = "conflicto_en_dataset_v1"
        elif now is None:
            status = "ausente_en_base_actual"
        elif previous == now["familia_actual"]:
            status = "coincide"
        else:
            status = "cambio_de_familia"
        comparison.append(
            {
                "hash_md5": hash_md5,
                "familia_anterior": previous,
                "familia_actual": "" if now is None else now["familia_actual"],
                "confianza_actual": "" if now is None else now["confianza_actual"],
                "score_actual": "" if now is None else now["score_actual"],
                "estado": status,
            }
        )

    status_counts = Counter(str(row["estado"]) for row in comparison)
    transitions = Counter(
        (str(row["familia_anterior"]), str(row["familia_actual"]))
        for row in comparison
        if row["estado"] == "cambio_de_familia"
    )
    unique_hashes = len(labels_by_hash)
    summary = {
        "previous_csv_rows": len(previous_rows),
        "previous_unique_hashes": unique_hashes,
        "duplicate_rows": len(previous_rows) - unique_hashes,
        "hashes_with_conflicting_previous_labels": sum(len(v) > 1 for v in labels_by_hash.values()),
        "status_counts": dict(status_counts),
        "agreement_percent_over_unique_hashes": round(100 * status_counts["coincide"] / unique_hashes, 4),
        "top_family_changes": [
            {"familia_anterior": old, "familia_actual": new, "muestras": count}
            for (old, new), count in transitions.most_common(30)
        ],
        "limitations": [
            "Dataset_V1 contiene una seleccion de ocho familias, no todas las inferencias historicas.",
            "No existen resultados AVClass persistidos para compararlos por hash.",
            "La coincidencia entre versiones no demuestra que la etiqueta sea correcta."
        ],
    }

    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
    (OUTPUT_DIR / "summary.json").write_text(json.dumps(summary, ensure_ascii=False, indent=2), encoding="utf-8")
    with (OUTPUT_DIR / "comparacion_por_hash.csv").open("w", encoding="utf-8", newline="") as handle:
        writer = csv.DictWriter(handle, fieldnames=list(comparison[0]))
        writer.writeheader()
        writer.writerows(comparison)

    print(json.dumps(summary, ensure_ascii=False, indent=2))


if __name__ == "__main__":
    main()
