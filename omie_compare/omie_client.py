"""Cliente HTTP genérico para a API OMIE.

A API OMIE responde sempre com HTTP 200, mesmo em caso de erro: o erro vem
dentro do corpo JSON como `faultstring`/`faultcode`. Este cliente trata:

- Retry com backoff exponencial para timeouts, HTTP 425/429/500/502/503 e
  faults que indicam limite de requisições (rate limit).
- Throttle mínimo entre chamadas para não estourar o limite da OMIE.
- Paginação genérica para endpoints que seguem o padrão
  `pagina` / `registros_por_pagina` / `total_de_paginas`.
"""

from __future__ import annotations

import time
import unicodedata
from typing import Any, Callable, Iterator

import requests

BASE_URL = "https://app.omie.com.br/api/v1"

MAX_RETRIES = 5
BASE_BACKOFF_SECONDS = 2.0
THROTTLE_SECONDS = 0.35

RETRYABLE_HTTP_STATUS = {425, 429, 500, 502, 503}
RATE_LIMIT_MARKERS = (
    "limite de requisi",
    "too many requests",
    "rate limit",
)


class OmieError(Exception):
    """Erro de negócio retornado pela OMIE (faultstring/faultcode)."""

    def __init__(self, faultcode: str, faultstring: str):
        self.faultcode = faultcode
        self.faultstring = faultstring
        super().__init__(f"OMIE fault {faultcode}: {faultstring}")


def normalize_key(key: str) -> str:
    """Remove acentos e baixa a caixa de uma chave, para comparação tolerante."""
    nfkd = unicodedata.normalize("NFKD", str(key))
    ascii_only = "".join(c for c in nfkd if not unicodedata.combining(c))
    return ascii_only.lower().strip()


class OmieClient:
    """Cliente para um único app (empresa) da OMIE."""

    def __init__(self, app_key: str, app_secret: str, session: requests.Session | None = None):
        self.app_key = str(app_key)
        self.app_secret = str(app_secret)
        self.session = session or requests.Session()
        self._last_call_ts: float | None = None

    def _throttle(self) -> None:
        if self._last_call_ts is not None:
            elapsed = time.monotonic() - self._last_call_ts
            wait = THROTTLE_SECONDS - elapsed
            if wait > 0:
                time.sleep(wait)
        self._last_call_ts = time.monotonic()

    def call(self, resource: str, method: str, params: dict[str, Any]) -> dict[str, Any]:
        """Chama `<BASE_URL>/<resource>/` com o método `method` e um único
        parâmetro `params`, aplicando retry/backoff e throttle."""
        url = f"{BASE_URL}/{resource.strip('/')}/"
        payload = {
            "call": method,
            "app_key": self.app_key,
            "app_secret": self.app_secret,
            "param": [params],
        }

        last_exc: Exception | None = None
        for attempt in range(1, MAX_RETRIES + 1):
            self._throttle()
            try:
                resp = self.session.post(url, json=payload, timeout=60)
            except (requests.Timeout, requests.ConnectionError) as exc:
                last_exc = exc
                self._sleep_backoff(attempt)
                continue

            if resp.status_code in RETRYABLE_HTTP_STATUS:
                last_exc = OmieError(str(resp.status_code), f"HTTP {resp.status_code}")
                self._sleep_backoff(attempt)
                continue

            resp.raise_for_status()

            try:
                data = resp.json()
            except ValueError as exc:
                last_exc = exc
                self._sleep_backoff(attempt)
                continue

            fault = data.get("faultstring") if isinstance(data, dict) else None
            if fault:
                faultcode = str(data.get("faultcode", ""))
                fault_lower = fault.lower()
                if any(marker in fault_lower for marker in RATE_LIMIT_MARKERS):
                    last_exc = OmieError(faultcode, fault)
                    self._sleep_backoff(attempt)
                    continue
                raise OmieError(faultcode, fault)

            return data

        assert last_exc is not None
        raise last_exc

    @staticmethod
    def _sleep_backoff(attempt: int) -> None:
        delay = BASE_BACKOFF_SECONDS * (2 ** (attempt - 1))
        time.sleep(delay)


def find_records_list(data: dict[str, Any]) -> list[dict[str, Any]]:
    """Acha, dentro de uma resposta da OMIE, a primeira chave cujo valor é
    uma lista de dicts (a lista de registros), sem assumir um nome fixo."""
    for value in data.values():
        if isinstance(value, list) and value and all(isinstance(item, dict) for item in value):
            return value
    return []


def find_field(record: dict[str, Any], *keywords: str) -> Any:
    """Procura, entre as chaves de `record` (recursivamente no primeiro
    nível), a primeira cuja chave normalizada contenha algum de `keywords`."""
    for key, value in record.items():
        norm = normalize_key(key)
        if any(kw in norm for kw in keywords):
            return value
    return None


def paginate(
    fetch_page: Callable[[int], dict[str, Any]],
    per_page: int = 200,
) -> Iterator[dict[str, Any]]:
    """Percorre todas as páginas de um endpoint que segue o padrão
    `pagina` / `registros_por_pagina` / `total_de_paginas`, produzindo os
    registros (dicts) de cada página.

    `fetch_page(pagina)` deve retornar o corpo JSON bruto da chamada para
    aquela página; esta função descobre a lista de registros dentro dele.
    """
    pagina = 1
    total_de_paginas = 1
    while pagina <= total_de_paginas:
        data = fetch_page(pagina)
        for record in find_records_list(data):
            yield record

        total_de_paginas = int(
            data.get("total_de_paginas")
            or data.get("totalDePaginas")
            or 1
        )
        pagina += 1
