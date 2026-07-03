#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
Agrega resultados AVClass a la SQLite generada por analisis_de_reportes.py.

Entrada esperada:
- outputs/Analisis_de_Reportes_Todos.db
- experimentos_avclass/results/virusshare_todos_avclass.labels
- experimentos_avclass/results/virusshare_todos_avclass.tags
"""

from __future__ import annotations

import argparse
import sqlite3
from collections import Counter, defaultdict
from pathlib import Path
from typing import Any


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Enriquece una DB de reportes con etiquetas AVClass.")
    parser.add_argument("--db", type=Path, default=Path("outputs/Analisis_de_Reportes_Todos.db"))
    parser.add_argument(
        "--labels",
        type=Path,
        default=Path("experimentos_avclass/results/virusshare_todos_avclass.labels"),
    )
    parser.add_argument(
        "--tags",
        type=Path,
        default=Path("experimentos_avclass/results/virusshare_todos_avclass.tags"),
    )
    return parser.parse_args()


def normalize_family(raw_family: str) -> tuple[str, int]:
    family = str(raw_family or "").strip().lower()
    if not family or family.startswith("singleton:"):
        return "sin_inferir", 1
    return family, 0


def parse_labels(path: Path) -> dict[str, dict[str, Any]]:
    labels: dict[str, dict[str, Any]] = {}
    with path.open("r", encoding="utf-8-sig") as handle:
        for line in handle:
            parts = line.strip().split()
            if len(parts) < 2:
                continue
            hash_md5 = parts[0].strip().lower()
            family_raw = parts[1].strip()
            family, is_singleton = normalize_family(family_raw)
            labels[hash_md5] = {
                "hash_md5": hash_md5,
                "familia_avclass_raw": family_raw,
                "familia_avclass": family,
                "avclass_sin_inferir": is_singleton,
            }
    return labels


def split_tag(tag: str) -> tuple[str, str]:
    if ":" not in tag:
        return "UNK", tag
    category, name = tag.split(":", 1)
    return category, name


def parse_tags(path: Path) -> tuple[dict[str, int], list[dict[str, Any]], dict[str, dict[str, Counter[str]]]]:
    tag_counts_by_hash: dict[str, int] = {}
    rows: list[dict[str, Any]] = []
    grouped: dict[str, dict[str, Counter[str]]] = defaultdict(lambda: defaultdict(Counter))

    with path.open("r", encoding="utf-8-sig") as handle:
        for line in handle:
            parts = line.strip().split("\t")
            if len(parts) < 3:
                continue
            hash_md5 = parts[0].strip().lower()
            try:
                tag_counts_by_hash[hash_md5] = int(parts[1])
            except ValueError:
                tag_counts_by_hash[hash_md5] = 0
            for item in parts[2].split(","):
                item = item.strip()
                if not item or "|" not in item:
                    continue
                tag_full, votes_text = item.rsplit("|", 1)
                try:
                    votes = int(votes_text)
                except ValueError:
                    votes = 1
                category, tag_name = split_tag(tag_full)
                rows.append(
                    {
                        "hash_md5": hash_md5,
                        "tag_category": category,
                        "tag_name": tag_name,
                        "tag_full": tag_full,
                        "votes": votes,
                    }
                )
                grouped[hash_md5][category][tag_name] += votes

    return tag_counts_by_hash, rows, grouped


def join_counter(counter: Counter[str], limit: int = 20) -> str:
    return ";".join(f"{name}:{votes}" for name, votes in counter.most_common(limit))


def ensure_column(conn: sqlite3.Connection, table: str, column: str, ddl_type: str) -> None:
    columns = {row[1] for row in conn.execute(f"PRAGMA table_info({table})")}
    if column not in columns:
        conn.execute(f"ALTER TABLE {table} ADD COLUMN {column} {ddl_type}")


def recreate_avclass_schema(conn: sqlite3.Connection) -> None:
    conn.executescript(
        """
        DROP TABLE IF EXISTS avclass_labels;
        DROP TABLE IF EXISTS avclass_tags;
        DROP TABLE IF EXISTS avclass_family_lote_counts;
        DROP TABLE IF EXISTS avclass_class_lote_counts;
        DROP TABLE IF EXISTS avclass_tag_lote_counts;

        CREATE TABLE avclass_labels (
            hash_md5 TEXT PRIMARY KEY,
            familia_avclass_raw TEXT,
            familia_avclass TEXT,
            avclass_sin_inferir INTEGER,
            avclass_tag_count INTEGER,
            clases_avclass TEXT,
            behaviors_avclass TEXT,
            file_tags_avclass TEXT,
            fam_tags_avclass TEXT
        );

        CREATE TABLE avclass_tags (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            hash_md5 TEXT,
            tag_category TEXT,
            tag_name TEXT,
            tag_full TEXT,
            votes INTEGER
        );

        CREATE TABLE avclass_family_lote_counts (
            lote_origen TEXT,
            familia_avclass TEXT,
            muestras INTEGER
        );

        CREATE TABLE avclass_class_lote_counts (
            lote_origen TEXT,
            class_avclass TEXT,
            muestras_con_clase INTEGER,
            votos INTEGER
        );

        CREATE TABLE avclass_tag_lote_counts (
            lote_origen TEXT,
            tag_category TEXT,
            tag_name TEXT,
            muestras_con_tag INTEGER,
            votos INTEGER
        );
        """
    )


def ensure_sample_columns(conn: sqlite3.Connection) -> None:
    columns = [
        ("familia_avclass_raw", "TEXT"),
        ("familia_avclass", "TEXT"),
        ("avclass_sin_inferir", "INTEGER"),
        ("avclass_tag_count", "INTEGER"),
        ("clases_avclass", "TEXT"),
        ("behaviors_avclass", "TEXT"),
        ("file_tags_avclass", "TEXT"),
        ("fam_tags_avclass", "TEXT"),
    ]
    for column, ddl_type in columns:
        ensure_column(conn, "samples", column, ddl_type)


def main() -> None:
    args = parse_args()
    if not args.db.exists():
        raise SystemExit(f"No existe la DB: {args.db}")
    if not args.labels.exists():
        raise SystemExit(f"No existe labels AVClass: {args.labels}")
    if not args.tags.exists():
        raise SystemExit(f"No existe tags AVClass: {args.tags}")

    labels = parse_labels(args.labels)
    tag_counts, tag_rows, grouped_tags = parse_tags(args.tags)

    conn = sqlite3.connect(args.db)
    try:
        conn.execute("PRAGMA journal_mode = WAL")
        conn.execute("PRAGMA synchronous = NORMAL")
        ensure_sample_columns(conn)
        recreate_avclass_schema(conn)

        label_rows = []
        for hash_md5, label in labels.items():
            categories = grouped_tags.get(hash_md5, {})
            label_rows.append(
                (
                    hash_md5,
                    label["familia_avclass_raw"],
                    label["familia_avclass"],
                    label["avclass_sin_inferir"],
                    tag_counts.get(hash_md5, 0),
                    join_counter(categories.get("CLASS", Counter())),
                    join_counter(categories.get("BEH", Counter())),
                    join_counter(categories.get("FILE", Counter())),
                    join_counter(categories.get("FAM", Counter())),
                )
            )

        conn.executemany(
            """
            INSERT INTO avclass_labels (
                hash_md5, familia_avclass_raw, familia_avclass, avclass_sin_inferir,
                avclass_tag_count, clases_avclass, behaviors_avclass, file_tags_avclass, fam_tags_avclass
            ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)
            """,
            label_rows,
        )

        conn.executemany(
            """
            INSERT INTO avclass_tags (hash_md5, tag_category, tag_name, tag_full, votes)
            VALUES (?, ?, ?, ?, ?)
            """,
            [
                (row["hash_md5"], row["tag_category"], row["tag_name"], row["tag_full"], row["votes"])
                for row in tag_rows
            ],
        )

        conn.execute(
            """
            UPDATE samples
            SET
                familia_avclass_raw = (
                    SELECT familia_avclass_raw FROM avclass_labels WHERE avclass_labels.hash_md5 = samples.hash_md5
                ),
                familia_avclass = (
                    SELECT familia_avclass FROM avclass_labels WHERE avclass_labels.hash_md5 = samples.hash_md5
                ),
                avclass_sin_inferir = (
                    SELECT avclass_sin_inferir FROM avclass_labels WHERE avclass_labels.hash_md5 = samples.hash_md5
                ),
                avclass_tag_count = (
                    SELECT avclass_tag_count FROM avclass_labels WHERE avclass_labels.hash_md5 = samples.hash_md5
                ),
                clases_avclass = (
                    SELECT clases_avclass FROM avclass_labels WHERE avclass_labels.hash_md5 = samples.hash_md5
                ),
                behaviors_avclass = (
                    SELECT behaviors_avclass FROM avclass_labels WHERE avclass_labels.hash_md5 = samples.hash_md5
                ),
                file_tags_avclass = (
                    SELECT file_tags_avclass FROM avclass_labels WHERE avclass_labels.hash_md5 = samples.hash_md5
                ),
                fam_tags_avclass = (
                    SELECT fam_tags_avclass FROM avclass_labels WHERE avclass_labels.hash_md5 = samples.hash_md5
                )
            WHERE EXISTS (
                SELECT 1 FROM avclass_labels WHERE avclass_labels.hash_md5 = samples.hash_md5
            )
            """
        )

        conn.executescript(
            """
            INSERT INTO avclass_family_lote_counts
            SELECT s.lote_origen, a.familia_avclass, COUNT(*) AS muestras
            FROM samples s
            JOIN avclass_labels a ON a.hash_md5 = s.hash_md5
            GROUP BY s.lote_origen, a.familia_avclass;

            INSERT INTO avclass_class_lote_counts
            SELECT s.lote_origen, t.tag_name AS class_avclass, COUNT(DISTINCT s.hash_md5) AS muestras_con_clase, SUM(t.votes) AS votos
            FROM samples s
            JOIN avclass_tags t ON t.hash_md5 = s.hash_md5
            WHERE t.tag_category = 'CLASS'
            GROUP BY s.lote_origen, t.tag_name;

            INSERT INTO avclass_tag_lote_counts
            SELECT s.lote_origen, t.tag_category, t.tag_name, COUNT(DISTINCT s.hash_md5) AS muestras_con_tag, SUM(t.votes) AS votos
            FROM samples s
            JOIN avclass_tags t ON t.hash_md5 = s.hash_md5
            GROUP BY s.lote_origen, t.tag_category, t.tag_name;

            CREATE INDEX IF NOT EXISTS idx_samples_avclass_family ON samples(familia_avclass);
            CREATE INDEX IF NOT EXISTS idx_samples_avclass_singleton ON samples(avclass_sin_inferir);
            CREATE INDEX IF NOT EXISTS idx_avclass_tags_hash ON avclass_tags(hash_md5);
            CREATE INDEX IF NOT EXISTS idx_avclass_tags_category ON avclass_tags(tag_category, tag_name);
            CREATE INDEX IF NOT EXISTS idx_avclass_labels_family ON avclass_labels(familia_avclass);
            """
        )
        conn.commit()

        samples = conn.execute("SELECT COUNT(*) FROM samples").fetchone()[0]
        enriched = conn.execute("SELECT COUNT(*) FROM samples WHERE familia_avclass IS NOT NULL").fetchone()[0]
        tags = conn.execute("SELECT COUNT(*) FROM avclass_tags").fetchone()[0]
        classes = conn.execute("SELECT COUNT(*) FROM avclass_tags WHERE tag_category = 'CLASS'").fetchone()[0]
        print(f"DB enriquecida: {args.db}")
        print(f"Muestras en samples: {samples:,}")
        print(f"Muestras con AVClass: {enriched:,}")
        print(f"Filas avclass_tags: {tags:,}")
        print(f"Filas CLASS: {classes:,}")
    finally:
        conn.close()


if __name__ == "__main__":
    main()
