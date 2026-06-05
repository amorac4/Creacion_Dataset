#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
Recolecta hashes SHA-256 desde el buscador web autenticado de VirusShare.

VirusShare no documenta paginacion para las busquedas web. Para recolectar mas
resultados que el limite de una consulta, este script divide recursivamente el
rango temporal cuando una busqueda alcanza ese limite.

Ejemplo:

    python recolectar_hashes_virusshare.py \
      --family salgorea \
      --date-from 2010-01-01 \
      --date-to 2026-06-05 \
      --limit 1000

Credenciales:

    $env:VIRUSSHARE_USERNAME = "usuario"
    $env:VIRUSSHARE_PASSWORD = "contrasena"

Si no se define VIRUSSHARE_PASSWORD, se solicita sin mostrarla en pantalla.
"""

from __future__ import annotations

import argparse
from collections import deque
from dataclasses import dataclass
from datetime import datetime, timedelta, timezone
from getpass import getpass
from html import unescape
from html.parser import HTMLParser
import json
import os
from pathlib import Path
import re
import sys
import time
from typing import Any, Iterable
from urllib.parse import urljoin

import requests
from requests import Response, Session
from requests.exceptions import RequestException


BASE_URL = "https://virusshare.com/"
LOGIN_URL = urljoin(BASE_URL, "processlogin")
DEFAULT_EXTRA_QUERY = "extension:exe detgte:5"
DEFAULT_DATE_FROM = "2000-01-01"
DEFAULT_LIMIT = 1000
DEFAULT_DELAY = 2.0
DEFAULT_MIN_WINDOW_SECONDS = 60
DEFAULT_RETRIES = 5

SHA256_RE = re.compile(
    r"""href\s*=\s*["'][^"']*file\?(?:hash=)?([0-9a-fA-F]{64})[^"']*["']""",
    re.IGNORECASE,
)


@dataclass(frozen=True)
class SearchForm:
    action: str
    method: str
    query_field: str
    hidden_fields: dict[str, str]


@dataclass(frozen=True)
class SearchResults:
    hashes: set[str]
    timed_out: bool


@dataclass(frozen=True)
class TimeWindow:
    start: datetime
    end: datetime

    @property
    def seconds(self) -> float:
        return (self.end - self.start).total_seconds()

    def split(self) -> tuple["TimeWindow", "TimeWindow"]:
        midpoint = self.start + (self.end - self.start) / 2
        return TimeWindow(self.start, midpoint), TimeWindow(midpoint, self.end)

    def to_json(self) -> dict[str, str]:
        return {"start": format_timestamp(self.start), "end": format_timestamp(self.end)}

    @classmethod
    def from_json(cls, value: dict[str, str]) -> "TimeWindow":
        return cls(parse_datetime(value["start"]), parse_datetime(value["end"]))


class FormParser(HTMLParser):
    def __init__(self) -> None:
        super().__init__()
        self.forms: list[dict[str, Any]] = []
        self._current: dict[str, Any] | None = None

    def handle_starttag(self, tag: str, attrs: list[tuple[str, str | None]]) -> None:
        attributes = {key.lower(): value or "" for key, value in attrs}
        if tag.lower() == "form":
            self._current = {
                "action": attributes.get("action", ""),
                "method": attributes.get("method", "GET").upper(),
                "inputs": [],
            }
            return

        if tag.lower() == "input" and self._current is not None:
            self._current["inputs"].append(attributes)

    def handle_endtag(self, tag: str) -> None:
        if tag.lower() == "form" and self._current is not None:
            self.forms.append(self._current)
            self._current = None


class VirusShareSearchError(RuntimeError):
    pass


class VirusShareSearchClient:
    def __init__(
        self,
        username: str,
        password: str,
        delay: float,
        retries: int,
        timeout: int = 90,
    ) -> None:
        self.username = username
        self.password = password
        self.delay = max(0.0, delay)
        self.retries = max(1, retries)
        self.timeout = timeout
        self.session: Session = requests.Session()
        self.session.headers.update(
            {"User-Agent": "virusshare-family-hash-collector/1.0"}
        )
        self.search_form: SearchForm | None = None
        self._last_request_at = 0.0

    def close(self) -> None:
        self.session.close()

    def login(self) -> None:
        response = self._request(
            "POST",
            LOGIN_URL,
            data={
                "username": self.username,
                "password": self.password,
                "remember": "on",
            },
            apply_delay=False,
        )
        if is_login_page(response.text):
            raise VirusShareSearchError(
                "VirusShare no acepto las credenciales o la cuenta no tiene acceso."
            )

        search_page = self._request("GET", urljoin(BASE_URL, "search"))
        if is_login_page(search_page.text):
            raise VirusShareSearchError(
                "La sesion de VirusShare no quedo autenticada."
            )
        self.search_form = discover_search_form(search_page.text, search_page.url)

    def search(self, query: str) -> SearchResults:
        if self.search_form is None:
            raise VirusShareSearchError("Debes iniciar sesion antes de buscar.")

        payload = dict(self.search_form.hidden_fields)
        payload[self.search_form.query_field] = query
        kwargs = {"params": payload} if self.search_form.method == "GET" else {"data": payload}
        response = self._request(self.search_form.method, self.search_form.action, **kwargs)

        if is_login_page(response.text):
            raise VirusShareSearchError(
                "La sesion expiro durante la busqueda. Ejecuta nuevamente el script."
            )
        return SearchResults(
            hashes=extract_sha256_hashes(response.text),
            timed_out=search_timed_out(response.text),
        )

    def _request(
        self,
        method: str,
        url: str,
        *,
        apply_delay: bool = True,
        **kwargs: Any,
    ) -> Response:
        last_error: Exception | None = None
        for attempt in range(self.retries):
            if apply_delay:
                elapsed = time.monotonic() - self._last_request_at
                if elapsed < self.delay:
                    time.sleep(self.delay - elapsed)

            try:
                response = self.session.request(
                    method,
                    url,
                    timeout=self.timeout,
                    allow_redirects=True,
                    **kwargs,
                )
                self._last_request_at = time.monotonic()
                if response.status_code == 200:
                    return response
                if response.status_code not in {429, 500, 502, 503, 504}:
                    raise VirusShareSearchError(
                        f"VirusShare devolvio HTTP {response.status_code} en {response.url}"
                    )
                last_error = VirusShareSearchError(
                    f"VirusShare devolvio HTTP {response.status_code}"
                )
            except RequestException as exc:
                last_error = exc

            wait_seconds = min(120, 2 ** attempt * 5)
            print(
                f"Peticion fallida ({attempt + 1}/{self.retries}); "
                f"reintentando en {wait_seconds}s...",
                file=sys.stderr,
            )
            time.sleep(wait_seconds)

        raise VirusShareSearchError(f"No fue posible consultar VirusShare: {last_error}")


def parse_args() -> argparse.Namespace:
    tomorrow = (datetime.now(timezone.utc) + timedelta(days=1)).date().isoformat()
    parser = argparse.ArgumentParser(
        description=(
            "Recolecta hashes SHA-256 por familia desde la busqueda web de VirusShare, "
            "dividiendo rangos temporales para superar el limite por consulta."
        )
    )
    family_group = parser.add_mutually_exclusive_group(required=True)
    family_group.add_argument("--family", help="Nombre de una familia, por ejemplo salgorea.")
    family_group.add_argument(
        "--families-file",
        type=Path,
        help="TXT con una familia por linea. Se crea un TXT separado por familia.",
    )
    parser.add_argument(
        "--query",
        help="Consulta base completa. Solo puede usarse junto con --family.",
    )
    parser.add_argument(
        "--extra-query",
        default=DEFAULT_EXTRA_QUERY,
        help=f"Filtros agregados al nombre de familia. Por defecto: {DEFAULT_EXTRA_QUERY!r}.",
    )
    parser.add_argument("--date-from", default=DEFAULT_DATE_FROM, help="Inicio UTC inclusivo.")
    parser.add_argument(
        "--date-to",
        default=tomorrow,
        help=f"Fin UTC exclusivo. Por defecto: {tomorrow}.",
    )
    parser.add_argument(
        "--limit",
        type=int,
        default=DEFAULT_LIMIT,
        help=f"Limite por consulta. Por defecto: {DEFAULT_LIMIT}.",
    )
    parser.add_argument(
        "--min-window-seconds",
        type=int,
        default=DEFAULT_MIN_WINDOW_SECONDS,
        help=(
            "Intervalo minimo que se puede dividir. Si aun alcanza el limite, se registra "
            f"como saturado. Por defecto: {DEFAULT_MIN_WINDOW_SECONDS}."
        ),
    )
    parser.add_argument(
        "--delay",
        type=float,
        default=DEFAULT_DELAY,
        help=f"Segundos entre peticiones. Por defecto: {DEFAULT_DELAY}.",
    )
    parser.add_argument(
        "--retries",
        type=int,
        default=DEFAULT_RETRIES,
        help=f"Reintentos por peticion. Por defecto: {DEFAULT_RETRIES}.",
    )
    parser.add_argument(
        "--output-dir",
        type=Path,
        default=Path("hashes") / "familias",
        help="Directorio para TXT y estados. Por defecto: hashes/familias.",
    )
    parser.add_argument(
        "--fresh",
        action="store_true",
        help="Reinicia estado y sobrescribe el TXT de cada familia seleccionada.",
    )
    parser.add_argument(
        "--username",
        default=os.environ.get("VIRUSSHARE_USERNAME", ""),
        help="Usuario de VirusShare. Tambien acepta VIRUSSHARE_USERNAME.",
    )
    parser.add_argument(
        "--password",
        default=os.environ.get("VIRUSSHARE_PASSWORD", ""),
        help="Contrasena de VirusShare. Es preferible usar VIRUSSHARE_PASSWORD.",
    )
    args = parser.parse_args()

    if args.query and not args.family:
        parser.error("--query solo puede usarse con --family.")
    if args.limit < 1:
        parser.error("--limit debe ser mayor que cero.")
    if args.min_window_seconds < 1:
        parser.error("--min-window-seconds debe ser mayor que cero.")
    return args


def is_login_page(html_text: str) -> bool:
    lowered = html_text.lower()
    return (
        'action="processlogin"' in lowered
        or "please <a href=\"login\">login</a> to search" in lowered
    )


def search_timed_out(html_text: str) -> bool:
    lowered = html_text.lower()
    return any(
        phrase in lowered
        for phrase in (
            "query timed out",
            "query timeout",
            "search timed out",
            "search timeout",
            "query time limit",
        )
    )


def discover_search_form(html_text: str, page_url: str) -> SearchForm:
    parser = FormParser()
    parser.feed(html_text)

    for form in parser.forms:
        inputs = form["inputs"]
        candidates = [
            item
            for item in inputs
            if item.get("name")
            and item.get("type", "text").lower() in {"text", "search"}
            and item.get("name", "").lower() not in {"username", "password"}
        ]
        if not candidates:
            continue

        hidden = {
            item["name"]: item.get("value", "")
            for item in inputs
            if item.get("name") and item.get("type", "").lower() == "hidden"
        }
        return SearchForm(
            action=urljoin(page_url, form["action"] or page_url),
            method=form["method"] if form["method"] in {"GET", "POST"} else "GET",
            query_field=candidates[0]["name"],
            hidden_fields=hidden,
        )

    raise VirusShareSearchError(
        "No se encontro el formulario de busqueda. VirusShare pudo cambiar su interfaz."
    )


def extract_sha256_hashes(html_text: str) -> set[str]:
    return {match.lower() for match in SHA256_RE.findall(unescape(html_text))}


def parse_datetime(value: str) -> datetime:
    normalized = value.strip()
    if normalized.endswith("Z"):
        normalized = normalized[:-1] + "+00:00"
    try:
        parsed = datetime.fromisoformat(normalized)
    except ValueError as exc:
        raise ValueError(
            f"Fecha invalida {value!r}; usa YYYY-MM-DD o un timestamp ISO 8601."
        ) from exc
    if parsed.tzinfo is None:
        parsed = parsed.replace(tzinfo=timezone.utc)
    return parsed.astimezone(timezone.utc)


def format_timestamp(value: datetime) -> str:
    return value.astimezone(timezone.utc).isoformat(timespec="seconds").replace("+00:00", "Z")


def safe_name(value: str) -> str:
    cleaned = re.sub(r"[^a-zA-Z0-9._-]+", "_", value.strip()).strip("._")
    return cleaned or "familia"


def build_family_query(family: str, extra_query: str) -> str:
    family_term = family.strip()
    if re.search(r"\s", family_term):
        family_term = f'"{family_term}"'
    return " ".join(part for part in (family_term, extra_query.strip()) if part)


def build_window_query(base_query: str, window: TimeWindow, limit: int) -> str:
    return (
        f'{base_query} after:"{format_timestamp(window.start)}" '
        f'before:"{format_timestamp(window.end)}" limit:{limit}'
    )


def read_families(args: argparse.Namespace) -> list[tuple[str, str]]:
    if args.family:
        query = args.query or build_family_query(args.family, args.extra_query)
        validate_base_query(query)
        return [(args.family.strip(), query)]

    families: list[tuple[str, str]] = []
    seen: set[str] = set()
    for line in args.families_file.read_text(encoding="utf-8").splitlines():
        family = line.strip()
        if not family or family.startswith("#") or family.casefold() in seen:
            continue
        seen.add(family.casefold())
        query = build_family_query(family, args.extra_query)
        validate_base_query(query)
        families.append((family, query))
    if not families:
        raise ValueError(f"No se encontraron familias en {args.families_file}.")
    return families


def validate_base_query(query: str) -> None:
    if re.search(r"(?i)(?:^|\s)(?:after|before|limit):", query):
        raise ValueError(
            "La consulta base no debe incluir after:, before: ni limit:; "
            "el script agrega esos contextos automaticamente."
        )


def load_hashes(path: Path) -> set[str]:
    if not path.exists():
        return set()
    return {
        line.strip().lower()
        for line in path.read_text(encoding="utf-8").splitlines()
        if re.fullmatch(r"[0-9a-fA-F]{64}", line.strip())
    }


def append_hashes(path: Path, hashes: Iterable[str]) -> None:
    ordered = sorted(set(hashes))
    if not ordered:
        return
    with path.open("a", encoding="utf-8", newline="\n") as output:
        for hash_value in ordered:
            output.write(f"{hash_value}\n")
        output.flush()


def save_json(path: Path, value: dict[str, Any]) -> None:
    temporary = path.with_suffix(path.suffix + ".tmp")
    temporary.write_text(
        json.dumps(value, ensure_ascii=True, indent=2) + "\n",
        encoding="utf-8",
    )
    temporary.replace(path)


def initial_state(family: str, query: str, window: TimeWindow, limit: int) -> dict[str, Any]:
    return {
        "version": 1,
        "family": family,
        "query": query,
        "limit": limit,
        "date_from": format_timestamp(window.start),
        "date_to": format_timestamp(window.end),
        "pending_windows": [window.to_json()],
        "completed_windows": 0,
        "requests": 0,
        "unique_hashes": 0,
        "saturated_windows": [],
        "completed": False,
        "updated_at": format_timestamp(datetime.now(timezone.utc)),
    }


def load_or_create_state(
    state_path: Path,
    family: str,
    query: str,
    window: TimeWindow,
    limit: int,
    fresh: bool,
) -> dict[str, Any]:
    if fresh or not state_path.exists():
        return initial_state(family, query, window, limit)

    state = json.loads(state_path.read_text(encoding="utf-8"))
    expected = {
        "family": family,
        "query": query,
        "limit": limit,
        "date_from": format_timestamp(window.start),
        "date_to": format_timestamp(window.end),
    }
    actual = {key: state.get(key) for key in expected}
    if actual != expected:
        raise ValueError(
            f"El estado {state_path} pertenece a otra consulta. "
            "Usa --fresh o cambia --output-dir."
        )
    return state


def collect_family(
    client: VirusShareSearchClient,
    *,
    family: str,
    query: str,
    root_window: TimeWindow,
    limit: int,
    min_window_seconds: int,
    output_dir: Path,
    fresh: bool,
) -> None:
    name = safe_name(family)
    hashes_path = output_dir / f"{name}.txt"
    state_path = output_dir / f"{name}.estado.json"

    if fresh:
        hashes_path.write_text("", encoding="utf-8")

    known_hashes = load_hashes(hashes_path)
    state = load_or_create_state(state_path, family, query, root_window, limit, fresh)
    pending = deque(TimeWindow.from_json(item) for item in state["pending_windows"])
    state["unique_hashes"] = len(known_hashes)

    if state.get("completed") and not pending:
        print(f"[{family}] Ya completado: {len(known_hashes):,} hashes en {hashes_path}")
        return

    print(f"[{family}] Consulta base: {query}")
    print(f"[{family}] Reanudando con {len(pending):,} intervalos pendientes.")

    while pending:
        window = pending.popleft()
        full_query = build_window_query(query, window, limit)
        search_results = client.search(full_query)
        results = search_results.hashes
        state["requests"] += 1

        new_hashes = results - known_hashes
        append_hashes(hashes_path, new_hashes)
        known_hashes.update(new_hashes)

        saturated = len(results) >= limit or search_results.timed_out
        if saturated and window.seconds > min_window_seconds:
            left, right = window.split()
            pending.appendleft(right)
            pending.appendleft(left)
            outcome = "dividido"
        else:
            state["completed_windows"] += 1
            outcome = "completo"
            if saturated:
                state["saturated_windows"].append(window.to_json())
                outcome = "SATURADO"

        state["pending_windows"] = [item.to_json() for item in pending]
        state["unique_hashes"] = len(known_hashes)
        state["completed"] = not pending
        state["updated_at"] = format_timestamp(datetime.now(timezone.utc))
        save_json(state_path, state)

        print(
            f"[{family}] {format_timestamp(window.start)} -> "
            f"{format_timestamp(window.end)}: {len(results):,} resultados, "
            f"{len(new_hashes):,} nuevos, {len(known_hashes):,} unicos; {outcome}; "
            f"{len(pending):,} pendientes."
        )

    print(f"[{family}] Terminado: {len(known_hashes):,} hashes en {hashes_path}")
    if state["saturated_windows"]:
        print(
            f"[{family}] Advertencia: {len(state['saturated_windows'])} intervalos "
            f"continuaron saturados al llegar a {min_window_seconds}s. "
            "Repite con --min-window-seconds mas pequeno para intentar recuperar mas."
        )


def main() -> int:
    args = parse_args()
    try:
        start = parse_datetime(args.date_from)
        end = parse_datetime(args.date_to)
        if start >= end:
            raise ValueError("--date-from debe ser anterior a --date-to.")
        families = read_families(args)
    except (OSError, ValueError) as exc:
        print(f"Error de configuracion: {exc}", file=sys.stderr)
        return 2

    username = args.username.strip() or input("Usuario de VirusShare: ").strip()
    password = args.password or getpass("Contrasena de VirusShare: ")
    if not username or not password:
        print("Se requieren usuario y contrasena de VirusShare.", file=sys.stderr)
        return 2

    args.output_dir.mkdir(parents=True, exist_ok=True)
    client = VirusShareSearchClient(
        username=username,
        password=password,
        delay=args.delay,
        retries=args.retries,
    )
    try:
        print("Iniciando sesion en VirusShare...")
        client.login()
        print("Sesion iniciada.")
        root_window = TimeWindow(start, end)
        for family, query in families:
            collect_family(
                client,
                family=family,
                query=query,
                root_window=root_window,
                limit=args.limit,
                min_window_seconds=args.min_window_seconds,
                output_dir=args.output_dir,
                fresh=args.fresh,
            )
    except (OSError, ValueError, VirusShareSearchError) as exc:
        print(f"Error: {exc}", file=sys.stderr)
        return 1
    except KeyboardInterrupt:
        print("\nInterrumpido. El estado quedo guardado para reanudar.", file=sys.stderr)
        return 130
    finally:
        client.close()
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
