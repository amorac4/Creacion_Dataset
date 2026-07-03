#!/usr/bin/env python3
# -*- coding: utf-8 -*-

from __future__ import annotations

import argparse
import csv
import datetime as dt
import json
import mimetypes
import sqlite3
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path
from urllib.parse import parse_qs, unquote, urlparse


APP_DIR = Path(__file__).resolve().parent
REPO_ROOT = APP_DIR.parents[1]
STATIC_DIR = APP_DIR / "static"
DEFAULT_DB = REPO_ROOT / "outputs" / "Analisis_de_Reportes_Todos.db"
DEFAULT_EXPORT = REPO_ROOT / "outputs" / "etiquetas_finales_hibridas.csv"


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Interfaz local para revisar etiquetas hibridas.")
    parser.add_argument("--db", type=Path, default=DEFAULT_DB)
    parser.add_argument("--host", default="127.0.0.1")
    parser.add_argument("--port", type=int, default=8765)
    return parser.parse_args()


def now_iso() -> str:
    return dt.datetime.now(dt.timezone.utc).isoformat()


def row_to_dict(row: sqlite3.Row | None) -> dict:
    return dict(row) if row is not None else {}


def json_response(handler: BaseHTTPRequestHandler, payload: object, status: int = 200) -> None:
    data = json.dumps(payload, ensure_ascii=False).encode("utf-8")
    handler.send_response(status)
    handler.send_header("Content-Type", "application/json; charset=utf-8")
    handler.send_header("Content-Length", str(len(data)))
    handler.end_headers()
    handler.wfile.write(data)


def text_response(handler: BaseHTTPRequestHandler, text: str, status: int = 200) -> None:
    data = text.encode("utf-8")
    handler.send_response(status)
    handler.send_header("Content-Type", "text/plain; charset=utf-8")
    handler.send_header("Content-Length", str(len(data)))
    handler.end_headers()
    handler.wfile.write(data)


def read_json_body(handler: BaseHTTPRequestHandler) -> dict:
    length = int(handler.headers.get("Content-Length") or "0")
    if length <= 0:
        return {}
    raw = handler.rfile.read(length).decode("utf-8")
    return json.loads(raw or "{}")


def norm(value: object) -> str:
    text = str(value or "").strip().lower()
    return text or "sin_inferir"


def compute_final(action: str, record: dict, familia_manual: str = "") -> tuple[str, str]:
    local = norm(record.get("familia_local") or record.get("familia_probable"))
    avclass = norm(record.get("familia_avclass"))
    suggested = norm(record.get("familia_final_sugerida") or record.get("familia_final"))
    action = norm(action)
    if action == "accept_local":
        return local, "aceptar_local"
    if action == "accept_avclass":
        return avclass, "aceptar_avclass"
    if action == "accept_suggested":
        return suggested, "aceptar_sugerida"
    if action == "sin_inferir":
        return "sin_inferir", "marcar_sin_inferir"
    if action == "manual":
        manual = norm(familia_manual)
        if manual == "sin_inferir":
            return "sin_inferir", "manual_sin_inferir"
        return manual, "manual"
    return suggested, "aceptar_sugerida"


class LabelReviewerApp:
    def __init__(self, db_path: Path):
        self.db_path = db_path.resolve()
        if not self.db_path.exists():
            raise FileNotFoundError(f"No existe la DB: {self.db_path}")
        self.init_db()

    def connect(self) -> sqlite3.Connection:
        conn = sqlite3.connect(self.db_path)
        conn.row_factory = sqlite3.Row
        return conn

    def init_db(self) -> None:
        with self.connect() as conn:
            missing = [
                table
                for table in ("samples", "label_hybrid", "label_pair_summary", "label_review_queue")
                if conn.execute(
                    "SELECT 1 FROM sqlite_master WHERE type='table' AND name=?",
                    (table,),
                ).fetchone()
                is None
            ]
            if missing:
                raise RuntimeError(f"La DB no tiene tablas requeridas: {', '.join(missing)}")

            conn.executescript(
                """
                CREATE TABLE IF NOT EXISTS label_manual_reviews (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    hash_md5 TEXT NOT NULL,
                    familia_local TEXT,
                    familia_avclass TEXT,
                    familia_final TEXT NOT NULL,
                    decision TEXT NOT NULL,
                    note TEXT,
                    reviewer TEXT,
                    reviewed_at TEXT NOT NULL
                );

                CREATE TABLE IF NOT EXISTS label_pair_manual_reviews (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    familia_local TEXT NOT NULL,
                    familia_avclass TEXT NOT NULL,
                    familia_final TEXT NOT NULL,
                    decision TEXT NOT NULL,
                    note TEXT,
                    reviewer TEXT,
                    reviewed_at TEXT NOT NULL
                );

                CREATE INDEX IF NOT EXISTS idx_label_manual_reviews_hash
                    ON label_manual_reviews(hash_md5, id);
                CREATE INDEX IF NOT EXISTS idx_label_pair_manual_reviews_pair
                    ON label_pair_manual_reviews(familia_local, familia_avclass, id);
                CREATE INDEX IF NOT EXISTS idx_label_hybrid_pair_lookup
                    ON label_hybrid(familia_local, familia_avclass, prioridad_revision, hash_md5);

                DROP VIEW IF EXISTS label_final_current;
                CREATE VIEW label_final_current AS
                WITH latest_hash AS (
                    SELECT r.*
                    FROM label_manual_reviews r
                    JOIN (
                        SELECT hash_md5, MAX(id) AS id
                        FROM label_manual_reviews
                        GROUP BY hash_md5
                    ) last ON last.hash_md5 = r.hash_md5 AND last.id = r.id
                ),
                latest_pair AS (
                    SELECT r.*
                    FROM label_pair_manual_reviews r
                    JOIN (
                        SELECT familia_local, familia_avclass, MAX(id) AS id
                        FROM label_pair_manual_reviews
                        GROUP BY familia_local, familia_avclass
                    ) last ON last.familia_local = r.familia_local
                        AND last.familia_avclass = r.familia_avclass
                        AND last.id = r.id
                )
                SELECT
                    lh.*,
                    COALESCE(h.familia_final, p.familia_final, lh.familia_final) AS familia_final_actual,
                    CASE
                        WHEN h.familia_final IS NOT NULL THEN 'manual_hash'
                        WHEN p.familia_final IS NOT NULL THEN 'manual_pair'
                        ELSE 'automatico'
                    END AS fuente_final_actual,
                    COALESCE(h.decision, p.decision, lh.decision_etiqueta) AS decision_actual,
                    COALESCE(h.note, p.note, lh.motivo_decision) AS nota_actual,
                    COALESCE(h.reviewed_at, p.reviewed_at, '') AS fecha_revision_actual
                FROM label_hybrid lh
                LEFT JOIN latest_hash h ON h.hash_md5 = lh.hash_md5
                LEFT JOIN latest_pair p ON p.familia_local = lh.familia_local
                    AND p.familia_avclass = lh.familia_avclass;
                """
            )
            conn.commit()

    def summary(self) -> dict:
        with self.connect() as conn:
            confidence = [
                row_to_dict(row)
                for row in conn.execute(
                    """
                    SELECT confianza_final AS name, COUNT(*) AS count
                    FROM label_hybrid
                    GROUP BY confianza_final
                    ORDER BY count DESC
                    """
                )
            ]
            final_counts = [
                row_to_dict(row)
                for row in conn.execute(
                    """
                    SELECT familia_final_actual AS name, COUNT(*) AS count
                    FROM label_final_current
                    GROUP BY familia_final_actual
                    ORDER BY count DESC
                    LIMIT 20
                    """
                )
            ]
            review_status = [
                row_to_dict(row)
                for row in conn.execute(
                    """
                    SELECT estado_revision AS name, COUNT(*) AS count
                    FROM label_review_queue
                    GROUP BY estado_revision
                    ORDER BY count DESC
                    """
                )
            ]
            top_pairs = [
                row_to_dict(row)
                for row in conn.execute(
                    """
                    SELECT
                        p.*,
                        (
                            SELECT h.hash_md5
                            FROM label_hybrid h
                            WHERE h.familia_local = p.familia_local
                              AND h.familia_avclass = p.familia_avclass
                            ORDER BY h.prioridad_revision, h.hash_md5
                            LIMIT 1
                        ) AS ejemplo_hash,
                        (
                            SELECT h.detecciones_top
                            FROM label_hybrid h
                            WHERE h.familia_local = p.familia_local
                              AND h.familia_avclass = p.familia_avclass
                              AND COALESCE(h.detecciones_top, '') <> ''
                            ORDER BY h.prioridad_revision, h.hash_md5
                            LIMIT 1
                        ) AS detecciones_top
                    FROM label_pair_summary p
                    WHERE requiere_revision = 1
                      AND NOT EXISTS (
                          SELECT 1
                          FROM label_pair_manual_reviews r
                          WHERE r.familia_local = p.familia_local
                            AND r.familia_avclass = p.familia_avclass
                      )
                    ORDER BY prioridad_revision, muestras DESC
                    LIMIT 12
                    """
                )
            ]
            totals = row_to_dict(
                conn.execute(
                    """
                    SELECT
                        (SELECT COUNT(*) FROM label_hybrid) AS muestras,
                        (SELECT COUNT(*) FROM label_hybrid WHERE requiere_revision = 1) AS requieren_revision,
                        (SELECT COUNT(*) FROM label_review_queue WHERE estado_revision = 'pendiente') AS pendientes,
                        (SELECT COUNT(*) FROM label_manual_reviews) AS revisiones_hash,
                        (SELECT COUNT(*) FROM label_pair_manual_reviews) AS revisiones_par
                    """
                ).fetchone()
            )
        return {
            "db": str(self.db_path),
            "totals": totals,
            "confidence": confidence,
            "final_counts": final_counts,
            "review_status": review_status,
            "top_pairs": top_pairs,
        }

    def pairs(self, params: dict[str, list[str]]) -> dict:
        only_review = params.get("review", ["1"])[0] != "0"
        search = params.get("q", [""])[0].strip().lower()
        limit = min(int(params.get("limit", ["200"])[0]), 1000)
        clauses = []
        values: list[object] = []
        if only_review:
            clauses.append("requiere_revision = 1")
            clauses.append(
                """
                NOT EXISTS (
                    SELECT 1
                    FROM label_pair_manual_reviews r
                    WHERE r.familia_local = p.familia_local
                      AND r.familia_avclass = p.familia_avclass
                )
                """
            )
        if search:
            clauses.append("(familia_local LIKE ? OR familia_avclass LIKE ?)")
            values.extend([f"%{search}%", f"%{search}%"])
        where = "WHERE " + " AND ".join(clauses) if clauses else ""
        query = f"""
            SELECT
                p.*,
                (
                    SELECT h.hash_md5
                    FROM label_hybrid h
                    WHERE h.familia_local = p.familia_local
                      AND h.familia_avclass = p.familia_avclass
                    ORDER BY h.prioridad_revision, h.hash_md5
                    LIMIT 1
                ) AS ejemplo_hash,
                (
                    SELECT h.detecciones_top
                    FROM label_hybrid h
                    WHERE h.familia_local = p.familia_local
                      AND h.familia_avclass = p.familia_avclass
                      AND COALESCE(h.detecciones_top, '') <> ''
                    ORDER BY h.prioridad_revision, h.hash_md5
                    LIMIT 1
                ) AS detecciones_top
            FROM label_pair_summary p
            {where}
            ORDER BY prioridad_revision, muestras DESC
            LIMIT ?
        """
        values.append(limit)
        with self.connect() as conn:
            rows = [row_to_dict(row) for row in conn.execute(query, values)]
        return {"rows": rows}

    def queue(self, params: dict[str, list[str]]) -> dict:
        status = params.get("status", ["pendiente"])[0]
        priority = params.get("priority", [""])[0]
        conflict = params.get("conflict_key", [""])[0]
        search = params.get("q", [""])[0].strip().lower()
        limit = min(int(params.get("limit", ["100"])[0]), 500)
        offset = max(int(params.get("offset", ["0"])[0]), 0)

        clauses = []
        values: list[object] = []
        if status != "all":
            clauses.append("estado_revision = ?")
            values.append(status)
        if priority:
            clauses.append("prioridad_revision = ?")
            values.append(int(priority))
        if conflict:
            clauses.append("conflict_key = ?")
            values.append(conflict)
        if search:
            clauses.append("(hash_md5 LIKE ? OR familia_local LIKE ? OR familia_avclass LIKE ?)")
            values.extend([f"%{search}%", f"%{search}%", f"%{search}%"])
        where = "WHERE " + " AND ".join(clauses) if clauses else ""
        with self.connect() as conn:
            total = conn.execute(f"SELECT COUNT(*) FROM label_review_queue {where}", values).fetchone()[0]
            rows = [
                row_to_dict(row)
                for row in conn.execute(
                    f"""
                    SELECT *
                    FROM label_review_queue
                    {where}
                    ORDER BY prioridad_revision, id
                    LIMIT ? OFFSET ?
                    """,
                    [*values, limit, offset],
                )
            ]
        return {"total": total, "limit": limit, "offset": offset, "rows": rows}

    def sample(self, hash_md5: str) -> dict:
        hash_md5 = hash_md5.strip().lower()
        with self.connect() as conn:
            sample = row_to_dict(
                conn.execute(
                    """
                    SELECT s.*, lf.familia_final_actual, lf.fuente_final_actual, lf.decision_actual, lf.nota_actual
                    FROM samples s
                    LEFT JOIN label_final_current lf ON lf.hash_md5 = s.hash_md5
                    WHERE s.hash_md5 = ?
                    """,
                    (hash_md5,),
                ).fetchone()
            )
            hybrid = row_to_dict(
                conn.execute("SELECT * FROM label_hybrid WHERE hash_md5 = ?", (hash_md5,)).fetchone()
            )
            tags = [
                row_to_dict(row)
                for row in conn.execute(
                    """
                    SELECT tag_category, tag_name, tag_full, votes
                    FROM avclass_tags
                    WHERE hash_md5 = ?
                    ORDER BY tag_category, votes DESC
                    LIMIT 200
                    """,
                    (hash_md5,),
                )
            ]
            reviews = [
                row_to_dict(row)
                for row in conn.execute(
                    """
                    SELECT *
                    FROM label_manual_reviews
                    WHERE hash_md5 = ?
                    ORDER BY id DESC
                    """,
                    (hash_md5,),
                )
            ]
        return {"sample": sample, "hybrid": hybrid, "tags": tags, "reviews": reviews}

    def review_hash(self, payload: dict) -> dict:
        hash_md5 = norm(payload.get("hash_md5"))
        action = str(payload.get("action") or "accept_suggested")
        note = str(payload.get("note") or "")
        reviewer = str(payload.get("reviewer") or "local")
        familia_manual = str(payload.get("familia_manual") or "")
        with self.connect() as conn:
            record = row_to_dict(
                conn.execute(
                    """
                    SELECT q.*, h.familia_final
                    FROM label_review_queue q
                    LEFT JOIN label_hybrid h ON h.hash_md5 = q.hash_md5
                    WHERE q.hash_md5 = ?
                    ORDER BY q.id DESC
                    LIMIT 1
                    """,
                    (hash_md5,),
                ).fetchone()
            )
            if not record:
                record = row_to_dict(conn.execute("SELECT * FROM label_hybrid WHERE hash_md5 = ?", (hash_md5,)).fetchone())
            if not record:
                raise ValueError(f"No existe hash en la DB: {hash_md5}")
            final, decision = compute_final(action, record, familia_manual)
            conn.execute(
                """
                INSERT INTO label_manual_reviews (
                    hash_md5, familia_local, familia_avclass, familia_final,
                    decision, note, reviewer, reviewed_at
                ) VALUES (?, ?, ?, ?, ?, ?, ?, ?)
                """,
                (
                    hash_md5,
                    norm(record.get("familia_local")),
                    norm(record.get("familia_avclass")),
                    final,
                    decision,
                    note,
                    reviewer,
                    now_iso(),
                ),
            )
            conn.execute(
                """
                UPDATE label_review_queue
                SET estado_revision = 'revisado', familia_manual = ?, nota_revision = ?
                WHERE hash_md5 = ?
                """,
                (final, note, hash_md5),
            )
            conn.commit()
        return {"ok": True, "hash_md5": hash_md5, "familia_final": final, "decision": decision}

    def review_pair(self, payload: dict) -> dict:
        local = norm(payload.get("familia_local"))
        avclass = norm(payload.get("familia_avclass"))
        action = str(payload.get("action") or "accept_suggested")
        note = str(payload.get("note") or "")
        reviewer = str(payload.get("reviewer") or "local")
        familia_manual = str(payload.get("familia_manual") or "")
        with self.connect() as conn:
            record = row_to_dict(
                conn.execute(
                    """
                    SELECT *
                    FROM label_pair_summary
                    WHERE familia_local = ? AND familia_avclass = ?
                    """,
                    (local, avclass),
                ).fetchone()
            )
            if not record:
                raise ValueError(f"No existe par: {local} -> {avclass}")
            final, decision = compute_final(action, record, familia_manual)
            conn.execute(
                """
                INSERT INTO label_pair_manual_reviews (
                    familia_local, familia_avclass, familia_final, decision, note, reviewer, reviewed_at
                ) VALUES (?, ?, ?, ?, ?, ?, ?)
                """,
                (local, avclass, final, decision, note, reviewer, now_iso()),
            )
            conn.execute(
                """
                UPDATE label_review_queue
                SET estado_revision = 'revisado_por_par', familia_manual = ?, nota_revision = ?
                WHERE familia_local = ? AND familia_avclass = ? AND estado_revision = 'pendiente'
                """,
                (final, note, local, avclass),
            )
            changed = conn.total_changes
            conn.commit()
        return {
            "ok": True,
            "familia_local": local,
            "familia_avclass": avclass,
            "familia_final": final,
            "decision": decision,
            "changes": changed,
        }

    def export_final(self, payload: dict) -> dict:
        output = Path(payload.get("output") or DEFAULT_EXPORT)
        if not output.is_absolute():
            output = REPO_ROOT / output
        output = output.resolve()
        if not str(output).startswith(str(REPO_ROOT.resolve())):
            raise ValueError("La ruta de export debe quedar dentro del repo.")
        output.parent.mkdir(parents=True, exist_ok=True)
        query = """
            SELECT
                hash_md5, lote_origen, familia_local, familia_avclass,
                familia_final_actual, fuente_final_actual, decision_actual,
                confianza_final, requiere_revision, prioridad_revision,
                clases_avclass, behaviors_avclass
            FROM label_final_current
            ORDER BY lote_origen, familia_final_actual, hash_md5
        """
        with self.connect() as conn, output.open("w", encoding="utf-8", newline="") as handle:
            rows = conn.execute(query)
            fieldnames = [description[0] for description in rows.description]
            writer = csv.writer(handle)
            writer.writerow(fieldnames)
            count = 0
            for row in rows:
                writer.writerow([row[name] for name in fieldnames])
                count += 1
        return {"ok": True, "output": str(output), "rows": count}


class Handler(BaseHTTPRequestHandler):
    app: LabelReviewerApp

    def do_GET(self) -> None:  # noqa: N802
        parsed = urlparse(self.path)
        path = parsed.path
        params = parse_qs(parsed.query)
        try:
            if path == "/":
                self.serve_static("index.html")
            elif path.startswith("/static/"):
                self.serve_static(path.removeprefix("/static/"))
            elif path == "/api/summary":
                json_response(self, self.app.summary())
            elif path == "/api/pairs":
                json_response(self, self.app.pairs(params))
            elif path == "/api/queue":
                json_response(self, self.app.queue(params))
            elif path.startswith("/api/sample/"):
                json_response(self, self.app.sample(unquote(path.removeprefix("/api/sample/"))))
            else:
                json_response(self, {"error": "not_found", "path": path}, status=404)
        except Exception as exc:  # pragma: no cover - UI safety net
            json_response(self, {"error": str(exc)}, status=500)

    def do_POST(self) -> None:  # noqa: N802
        parsed = urlparse(self.path)
        try:
            payload = read_json_body(self)
            if parsed.path == "/api/review/hash":
                json_response(self, self.app.review_hash(payload))
            elif parsed.path == "/api/review/pair":
                json_response(self, self.app.review_pair(payload))
            elif parsed.path == "/api/export/final":
                json_response(self, self.app.export_final(payload))
            else:
                json_response(self, {"error": "not_found", "path": parsed.path}, status=404)
        except Exception as exc:  # pragma: no cover - UI safety net
            json_response(self, {"error": str(exc)}, status=500)

    def serve_static(self, name: str) -> None:
        target = (STATIC_DIR / name).resolve()
        if not str(target).startswith(str(STATIC_DIR.resolve())) or not target.is_file():
            text_response(self, "Not found", status=404)
            return
        content_type = mimetypes.guess_type(target.name)[0] or "application/octet-stream"
        data = target.read_bytes()
        self.send_response(200)
        self.send_header("Content-Type", content_type)
        self.send_header("Content-Length", str(len(data)))
        self.end_headers()
        self.wfile.write(data)

    def log_message(self, fmt: str, *args: object) -> None:
        print(f"[label-reviewer] {self.address_string()} - {fmt % args}")


def main() -> int:
    args = parse_args()
    app = LabelReviewerApp(args.db)
    Handler.app = app
    server = ThreadingHTTPServer((args.host, args.port), Handler)
    print(f"Label Reviewer: http://{args.host}:{args.port}")
    print(f"DB: {app.db_path}")
    try:
        server.serve_forever()
    except KeyboardInterrupt:
        print("\nDeteniendo servidor.")
    finally:
        server.server_close()
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
