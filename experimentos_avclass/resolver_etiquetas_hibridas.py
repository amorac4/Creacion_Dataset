#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
Resuelve etiquetas finales combinando la heuristica local y AVClass.

El script crea tablas reproducibles dentro de la SQLite enriquecida:

- label_hybrid: decision por hash.
- label_pair_summary: resumen por par local/AVClass.
- label_review_queue: muestra priorizada de hashes para revisar manualmente.

No modifica las etiquetas originales.
"""

from __future__ import annotations

import argparse
import csv
import json
import sqlite3
from collections import Counter, defaultdict
from pathlib import Path
from typing import Any


DEFAULT_CONFIG = {
    "trusted_local_families_when_avclass_sin_inferir": [
        "eylr",
        "fragtor",
        "injector",
        "razy",
        "strictor",
        "vbclone",
        "zusy",
    ],
    "generic_or_low_trust_avclass_families": ["sin_inferir", "bomb", "sabsik", "wacatac"],
    "review_conflict_pairs": [],
    "manual_pair_overrides": [],
    "minimum_local_score_for_silver": 0.08,
    "minimum_avclass_tags_for_silver": 2,
    "sample_review_rows_per_pair": 50,
}


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Crea etiquetas hibridas y cola de revision en la DB.")
    parser.add_argument("--db", type=Path, default=Path("outputs/Analisis_de_Reportes_Todos.db"))
    parser.add_argument(
        "--config",
        type=Path,
        default=Path("experimentos_avclass/config_etiquetado_hibrido.json"),
        help="Reglas configurables para resolver etiquetas.",
    )
    parser.add_argument(
        "--output-dir",
        type=Path,
        default=Path("experimentos_avclass/results/etiquetado_hibrido"),
        help="Directorio para CSV/JSON de auditoria.",
    )
    return parser.parse_args()


def load_config(path: Path) -> dict[str, Any]:
    config = dict(DEFAULT_CONFIG)
    if path.exists():
        user_config = json.loads(path.read_text(encoding="utf-8-sig"))
        config.update(user_config)
    return config


def norm(value: Any) -> str:
    text = str(value or "").strip().lower()
    return text or "sin_inferir"


def pair_key(local: str, avclass: str) -> tuple[str, str]:
    return norm(local), norm(avclass)


def build_pair_map(items: list[dict[str, Any]]) -> dict[tuple[str, str], dict[str, Any]]:
    result: dict[tuple[str, str], dict[str, Any]] = {}
    for item in items:
        result[pair_key(item.get("familia_local"), item.get("familia_avclass"))] = item
    return result


def table_exists(conn: sqlite3.Connection, table: str) -> bool:
    row = conn.execute(
        "SELECT 1 FROM sqlite_master WHERE type = 'table' AND name = ?",
        (table,),
    ).fetchone()
    return row is not None


def recreate_schema(conn: sqlite3.Connection) -> None:
    conn.executescript(
        """
        DROP TABLE IF EXISTS label_hybrid;
        DROP TABLE IF EXISTS label_pair_summary;
        DROP TABLE IF EXISTS label_review_queue;

        CREATE TABLE label_hybrid (
            hash_md5 TEXT PRIMARY KEY,
            lote_origen TEXT,
            familia_local TEXT,
            familia_avclass TEXT,
            familia_final TEXT,
            confianza_final TEXT,
            decision_etiqueta TEXT,
            motivo_decision TEXT,
            requiere_revision INTEGER,
            prioridad_revision INTEGER,
            conflict_key TEXT,
            familia_confianza_local TEXT,
            familia_score_local REAL,
            avclass_tag_count INTEGER,
            clases_avclass TEXT,
            behaviors_avclass TEXT,
            detecciones_top TEXT
        );

        CREATE TABLE label_pair_summary (
            familia_local TEXT,
            familia_avclass TEXT,
            muestras INTEGER,
            familia_final_sugerida TEXT,
            confianza_sugerida TEXT,
            decision_sugerida TEXT,
            requiere_revision INTEGER,
            prioridad_revision INTEGER,
            motivo TEXT,
            PRIMARY KEY (familia_local, familia_avclass)
        );

        CREATE TABLE label_review_queue (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            hash_md5 TEXT,
            lote_origen TEXT,
            familia_local TEXT,
            familia_avclass TEXT,
            familia_final_sugerida TEXT,
            confianza_sugerida TEXT,
            decision_sugerida TEXT,
            prioridad_revision INTEGER,
            conflict_key TEXT,
            clases_avclass TEXT,
            behaviors_avclass TEXT,
            detecciones_top TEXT,
            nota_revision TEXT,
            familia_manual TEXT,
            estado_revision TEXT DEFAULT 'pendiente'
        );
        """
    )


def decide_label(row: sqlite3.Row, config: dict[str, Any]) -> dict[str, Any]:
    local = norm(row["familia_probable"])
    avclass = norm(row["familia_avclass"])
    local_score = float(row["familia_score"] or 0)
    local_confidence = norm(row["familia_confianza"])
    avclass_tag_count = int(row["avclass_tag_count"] or 0)

    trusted_local = set(config["trusted_local_families_when_avclass_sin_inferir"])
    generic_avclass = set(config["generic_or_low_trust_avclass_families"])
    review_pairs = build_pair_map(config.get("review_conflict_pairs") or [])
    overrides = build_pair_map(config.get("manual_pair_overrides") or [])
    min_local_score = float(config["minimum_local_score_for_silver"])
    min_avclass_tags = int(config["minimum_avclass_tags_for_silver"])

    key = pair_key(local, avclass)
    if key in overrides:
        override = overrides[key]
        return {
            "familia_final": norm(override.get("familia_final")),
            "confianza_final": norm(override.get("confianza_final") or "gold_override_pair"),
            "decision_etiqueta": norm(override.get("decision") or "override_pair"),
            "motivo_decision": str(override.get("motivo") or "Regla manual por par."),
            "requiere_revision": 0,
            "prioridad_revision": 0,
        }

    if local == "sin_inferir" and avclass == "sin_inferir":
        return {
            "familia_final": "sin_inferir",
            "confianza_final": "sin_inferir",
            "decision_etiqueta": "ambas_sin_inferir",
            "motivo_decision": "Ni la heuristica local ni AVClass infirieron familia.",
            "requiere_revision": 0,
            "prioridad_revision": 9,
        }

    if local == avclass:
        if local in generic_avclass:
            return {
                "familia_final": local,
                "confianza_final": "bronze_confirmada_generica",
                "decision_etiqueta": "coinciden_pero_generica",
                "motivo_decision": "Ambas fuentes coinciden, pero la familia es generica o de baja confianza.",
                "requiere_revision": 1,
                "prioridad_revision": 6,
            }
        return {
            "familia_final": local,
            "confianza_final": "gold_confirmada",
            "decision_etiqueta": "coinciden_local_y_avclass",
            "motivo_decision": "La heuristica local y AVClass coinciden.",
            "requiere_revision": 0,
            "prioridad_revision": 0,
        }

    if local != "sin_inferir" and avclass == "sin_inferir":
        if local in trusted_local or local_score >= min_local_score or local_confidence in {"alta", "media"}:
            return {
                "familia_final": local,
                "confianza_final": "silver_local_no_confirmada",
                "decision_etiqueta": "mantener_local_avclass_no_infiere",
                "motivo_decision": "AVClass no infiere; se conserva etiqueta local con confianza media.",
                "requiere_revision": 0,
                "prioridad_revision": 4,
            }
        return {
            "familia_final": local,
            "confianza_final": "bronze_local_debil",
            "decision_etiqueta": "mantener_local_debil_avclass_no_infiere",
            "motivo_decision": "Etiqueta local debil y AVClass no infiere; revisar si se usara para entrenamiento.",
            "requiere_revision": 1,
            "prioridad_revision": 5,
        }

    if local == "sin_inferir" and avclass != "sin_inferir":
        if avclass not in generic_avclass and avclass_tag_count >= min_avclass_tags:
            return {
                "familia_final": avclass,
                "confianza_final": "silver_avclass_aporta",
                "decision_etiqueta": "usar_avclass_local_no_infiere",
                "motivo_decision": "La heuristica local no infiere y AVClass propone familia no generica.",
                "requiere_revision": 0,
                "prioridad_revision": 4,
            }
        return {
            "familia_final": avclass,
            "confianza_final": "bronze_avclass_generica_o_debil",
            "decision_etiqueta": "avclass_propone_debil",
            "motivo_decision": "AVClass propone familia generica o con poca evidencia; revisar.",
            "requiere_revision": 1,
            "prioridad_revision": 3,
        }

    if key in review_pairs:
        motivo = str(review_pairs[key].get("motivo") or "Par conflictivo marcado para revision.")
        return {
            "familia_final": local,
            "confianza_final": "conflicto_prioritario",
            "decision_etiqueta": "mantener_local_conflicto_prioritario",
            "motivo_decision": motivo,
            "requiere_revision": 1,
            "prioridad_revision": 1,
        }

    if avclass in generic_avclass:
        return {
            "familia_final": local,
            "confianza_final": "silver_local_avclass_generica",
            "decision_etiqueta": "mantener_local_avclass_generica",
            "motivo_decision": "AVClass contradice con familia generica o baja confianza; se conserva local.",
            "requiere_revision": 1,
            "prioridad_revision": 2,
        }

    return {
        "familia_final": local,
        "confianza_final": "conflicto",
        "decision_etiqueta": "mantener_local_conflicto",
        "motivo_decision": "Ambas fuentes infieren familias distintas; requiere revision manual u otro algoritmo.",
        "requiere_revision": 1,
        "prioridad_revision": 2,
    }


def write_csv(path: Path, rows: list[dict[str, Any]], fieldnames: list[str]) -> None:
    with path.open("w", encoding="utf-8", newline="") as handle:
        writer = csv.DictWriter(handle, fieldnames=fieldnames, extrasaction="ignore")
        writer.writeheader()
        writer.writerows(rows)


def main() -> None:
    args = parse_args()
    if not args.db.exists():
        raise SystemExit(f"No existe la DB: {args.db}")
    config = load_config(args.config)
    args.output_dir.mkdir(parents=True, exist_ok=True)

    conn = sqlite3.connect(args.db)
    conn.row_factory = sqlite3.Row
    try:
        if not table_exists(conn, "samples"):
            raise SystemExit("La DB no tiene tabla samples.")
        recreate_schema(conn)

        query = """
            SELECT
                hash_md5, lote_origen, familia_probable, familia_confianza, familia_score,
                familia_avclass, avclass_tag_count, clases_avclass, behaviors_avclass, detecciones_top
            FROM samples
        """

        hybrid_rows = []
        pair_counter: Counter[tuple[str, str]] = Counter()
        pair_decision: dict[tuple[str, str], dict[str, Any]] = {}
        review_candidates: dict[tuple[str, str], list[dict[str, Any]]] = defaultdict(list)
        decision_counts = Counter()
        confidence_counts = Counter()

        for row in conn.execute(query):
            local = norm(row["familia_probable"])
            avclass = norm(row["familia_avclass"])
            decision = decide_label(row, config)
            conflict_key = f"{local} -> {avclass}"
            record = {
                "hash_md5": row["hash_md5"],
                "lote_origen": row["lote_origen"],
                "familia_local": local,
                "familia_avclass": avclass,
                "familia_final": decision["familia_final"],
                "confianza_final": decision["confianza_final"],
                "decision_etiqueta": decision["decision_etiqueta"],
                "motivo_decision": decision["motivo_decision"],
                "requiere_revision": decision["requiere_revision"],
                "prioridad_revision": decision["prioridad_revision"],
                "conflict_key": conflict_key,
                "familia_confianza_local": row["familia_confianza"],
                "familia_score_local": row["familia_score"],
                "avclass_tag_count": row["avclass_tag_count"],
                "clases_avclass": row["clases_avclass"],
                "behaviors_avclass": row["behaviors_avclass"],
                "detecciones_top": row["detecciones_top"],
            }
            hybrid_rows.append(record)
            pair = (local, avclass)
            pair_counter[pair] += 1
            pair_decision.setdefault(pair, decision)
            decision_counts[decision["decision_etiqueta"]] += 1
            confidence_counts[decision["confianza_final"]] += 1
            if decision["requiere_revision"]:
                review_candidates[pair].append(record)

        conn.executemany(
            """
            INSERT INTO label_hybrid (
                hash_md5, lote_origen, familia_local, familia_avclass, familia_final,
                confianza_final, decision_etiqueta, motivo_decision, requiere_revision,
                prioridad_revision, conflict_key, familia_confianza_local, familia_score_local,
                avclass_tag_count, clases_avclass, behaviors_avclass, detecciones_top
            ) VALUES (
                :hash_md5, :lote_origen, :familia_local, :familia_avclass, :familia_final,
                :confianza_final, :decision_etiqueta, :motivo_decision, :requiere_revision,
                :prioridad_revision, :conflict_key, :familia_confianza_local, :familia_score_local,
                :avclass_tag_count, :clases_avclass, :behaviors_avclass, :detecciones_top
            )
            """,
            hybrid_rows,
        )

        summary_rows = []
        for (local, avclass), count in pair_counter.most_common():
            decision = pair_decision[(local, avclass)]
            summary_rows.append(
                {
                    "familia_local": local,
                    "familia_avclass": avclass,
                    "muestras": count,
                    "familia_final_sugerida": decision["familia_final"],
                    "confianza_sugerida": decision["confianza_final"],
                    "decision_sugerida": decision["decision_etiqueta"],
                    "requiere_revision": decision["requiere_revision"],
                    "prioridad_revision": decision["prioridad_revision"],
                    "motivo": decision["motivo_decision"],
                }
            )

        conn.executemany(
            """
            INSERT INTO label_pair_summary (
                familia_local, familia_avclass, muestras, familia_final_sugerida,
                confianza_sugerida, decision_sugerida, requiere_revision,
                prioridad_revision, motivo
            ) VALUES (
                :familia_local, :familia_avclass, :muestras, :familia_final_sugerida,
                :confianza_sugerida, :decision_sugerida, :requiere_revision,
                :prioridad_revision, :motivo
            )
            """,
            summary_rows,
        )

        review_limit = int(config["sample_review_rows_per_pair"])
        review_rows = []
        for pair, records in review_candidates.items():
            records.sort(
                key=lambda item: (
                    int(item["prioridad_revision"]),
                    -float(item["avclass_tag_count"] or 0),
                    str(item["hash_md5"]),
                )
            )
            review_rows.extend(records[:review_limit])

        conn.executemany(
            """
            INSERT INTO label_review_queue (
                hash_md5, lote_origen, familia_local, familia_avclass, familia_final_sugerida,
                confianza_sugerida, decision_sugerida, prioridad_revision, conflict_key,
                clases_avclass, behaviors_avclass, detecciones_top
            ) VALUES (
                :hash_md5, :lote_origen, :familia_local, :familia_avclass, :familia_final,
                :confianza_final, :decision_etiqueta, :prioridad_revision, :conflict_key,
                :clases_avclass, :behaviors_avclass, :detecciones_top
            )
            """,
            review_rows,
        )

        conn.executescript(
            """
            CREATE INDEX IF NOT EXISTS idx_label_hybrid_final ON label_hybrid(familia_final);
            CREATE INDEX IF NOT EXISTS idx_label_hybrid_confianza ON label_hybrid(confianza_final);
            CREATE INDEX IF NOT EXISTS idx_label_hybrid_revision ON label_hybrid(requiere_revision, prioridad_revision);
            CREATE INDEX IF NOT EXISTS idx_label_pair_revision ON label_pair_summary(requiere_revision, prioridad_revision);
            CREATE INDEX IF NOT EXISTS idx_label_review_estado ON label_review_queue(estado_revision, prioridad_revision);
            """
        )
        conn.commit()

        write_csv(
            args.output_dir / "label_pair_summary.csv",
            summary_rows,
            [
                "familia_local",
                "familia_avclass",
                "muestras",
                "familia_final_sugerida",
                "confianza_sugerida",
                "decision_sugerida",
                "requiere_revision",
                "prioridad_revision",
                "motivo",
            ],
        )
        write_csv(
            args.output_dir / "label_review_queue.csv",
            review_rows,
            [
                "hash_md5",
                "lote_origen",
                "familia_local",
                "familia_avclass",
                "familia_final",
                "confianza_final",
                "decision_etiqueta",
                "prioridad_revision",
                "conflict_key",
                "clases_avclass",
                "behaviors_avclass",
                "detecciones_top",
            ],
        )

        summary = {
            "db": str(args.db),
            "total_muestras": len(hybrid_rows),
            "requieren_revision": sum(1 for row in hybrid_rows if row["requiere_revision"]),
            "pares_requieren_revision": sum(1 for row in summary_rows if row["requiere_revision"]),
            "decision_counts": decision_counts.most_common(),
            "confidence_counts": confidence_counts.most_common(),
            "top_pares_revision": [
                row
                for row in summary_rows
                if row["requiere_revision"]
            ][:30],
        }
        (args.output_dir / "resumen_etiquetado_hibrido.json").write_text(
            json.dumps(summary, ensure_ascii=False, indent=2),
            encoding="utf-8",
        )
        print(json.dumps(summary, ensure_ascii=False, indent=2)[:6000])
    finally:
        conn.close()


if __name__ == "__main__":
    main()
