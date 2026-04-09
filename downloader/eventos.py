import base64
import binascii
import gzip
import json
import logging
import queue
import threading
import time
import xml.etree.ElementTree as ET
from dataclasses import dataclass
from typing import Callable, Optional

import requests

logger = logging.getLogger(__name__)
RETRY_STATUS_CODES = {429, 502, 503, 504}


@dataclass(slots=True)
class EventDownloadTask:
    """Representa uma consulta de eventos por chave da NFS-e."""

    chave: str
    dest_path: str
    nsu: int
    ano: str
    mes: str
    tipo_documento: str
    document_paths: tuple[str, ...] = ()
    canceladas_dir: str = ""


@dataclass(slots=True)
class EventFetchResult:
    """Resultado da busca de eventos para uma chave."""

    status: str
    reason: Optional[str] = None
    saved_count: int = 0
    cancelamento: bool = False


class NFSeEventDownloader:
    """Downloader para eventos da NFS-e por chave de acesso."""

    BASE_URL = "https://adn.nfse.gov.br/contribuintes/NFSe"

    def __init__(
        self,
        session,
        timeout: int = 30,
        delay_seconds: float = 2.0,
        max_retries: int = 3,
        retry_backoff: float = 2.0,
    ):
        self.session = session
        self.timeout = timeout
        self.delay_seconds = max(0.0, float(delay_seconds))
        self.max_retries = max(0, int(max_retries))
        self.retry_backoff = max(1.0, float(retry_backoff))
        self._last_request_at = 0.0

    def baixar(self, chave: str, dest_path: str) -> EventFetchResult:
        """Busca os eventos vinculados a ``chave`` e salva um XML consolidado."""
        url = f"{self.BASE_URL}/{chave}/Eventos"
        attempt = 0

        while True:
            self._respect_delay()

            try:
                with self.session.get(url, timeout=self.timeout) as resp:
                    self._mark_request()

                    if resp.status_code == 200:
                        xml_bytes = self._extract_event_xml(resp)
                        if not xml_bytes:
                            logger.debug("Nenhum evento localizado para a chave %s", chave)
                            return EventFetchResult(status="no_events")

                        self._salvar_arquivo(dest_path, xml_bytes)
                        return EventFetchResult(
                            status="saved",
                            saved_count=1,
                            cancelamento=self._contains_cancelamento(xml_bytes),
                        )

                    if resp.status_code in (204, 404):
                        logger.debug(
                            "Nenhum evento retornado para a chave %s (HTTP %s)",
                            chave,
                            resp.status_code,
                        )
                        return EventFetchResult(status="no_events")

                    if resp.status_code in RETRY_STATUS_CODES and attempt < self.max_retries:
                        wait_seconds = self.delay_seconds * (self.retry_backoff ** attempt)
                        logger.warning(
                            "Falha temporaria ao buscar eventos %s: HTTP %s. "
                            "Tentativa %s/%s em %.1fs.",
                            chave,
                            resp.status_code,
                            attempt + 1,
                            self.max_retries,
                            wait_seconds,
                        )
                        time.sleep(wait_seconds)
                        attempt += 1
                        continue

                    return EventFetchResult(
                        status="failure",
                        reason=f"HTTP {resp.status_code}",
                    )
            except Exception as exc:
                self._mark_request()

                if attempt < self.max_retries:
                    wait_seconds = self.delay_seconds * (self.retry_backoff ** attempt)
                    logger.warning(
                        "Erro temporario ao buscar eventos %s: %s. "
                        "Tentativa %s/%s em %.1fs.",
                        chave,
                        str(exc),
                        attempt + 1,
                        self.max_retries,
                        wait_seconds,
                    )
                    time.sleep(wait_seconds)
                    attempt += 1
                    continue

                return EventFetchResult(status="failure", reason=str(exc))

    def _extract_event_xml(self, resp: requests.Response) -> Optional[bytes]:
        """Extrai o XML de evento de respostas XML ou JSON da API."""
        content = resp.content or b""
        if not content.strip():
            return None

        if content.lstrip().startswith(b"<"):
            return content if self._looks_like_event_xml(content) else None

        try:
            payload = resp.json()
        except (json.JSONDecodeError, ValueError):
            return None

        documents = []
        for candidate in self._collect_payload_candidates(payload):
            xml_bytes = self._decode_payload(candidate)
            if xml_bytes and self._looks_like_event_xml(xml_bytes):
                documents.append(xml_bytes)

        if not documents:
            return None

        return self._merge_documents(documents)

    def _collect_payload_candidates(self, payload) -> list[str]:
        """Varre respostas JSON procurando campos que parecam conter XML."""
        candidates: list[str] = []

        def _walk(obj, key_hint: str = "") -> None:
            if isinstance(obj, dict):
                for key, value in obj.items():
                    _walk(value, str(key))
                return

            if isinstance(obj, list):
                for item in obj:
                    _walk(item, key_hint)
                return

            if not isinstance(obj, str):
                return

            stripped = obj.strip()
            key_lower = key_hint.lower()
            if not stripped:
                return

            if stripped.startswith("<"):
                candidates.append(stripped)
                return

            if any(token in key_lower for token in ("xml", "arquivo", "conteudo", "document", "evento")):
                candidates.append(stripped)

        _walk(payload)
        return candidates

    def _decode_payload(self, payload: str) -> Optional[bytes]:
        stripped = payload.strip()
        if not stripped:
            return None

        if stripped.startswith("<"):
            return stripped.encode("utf-8")

        try:
            decoded = base64.b64decode(stripped, validate=True)
        except (binascii.Error, ValueError):
            return None

        if not decoded:
            return None

        try:
            return gzip.decompress(decoded)
        except OSError:
            return decoded

    def _looks_like_event_xml(self, xml_bytes: bytes) -> bool:
        try:
            root = ET.fromstring(xml_bytes)
        except ET.ParseError:
            return False

        if any(
            root.find(path) is not None
            for path in (
                ".//{*}infEvento",
                ".//{*}Evento",
                ".//{*}evento",
                ".//{*}pedRegEvento",
                ".//{*}infPedReg",
            )
        ):
            return True

        root_tag = str(root.tag).lower()
        return "evento" in root_tag

    def _merge_documents(self, documents: list[bytes]) -> bytes:
        if len(documents) == 1:
            return documents[0]

        wrapper = ET.Element("eventos")
        appended = 0

        for document in documents:
            try:
                wrapper.append(ET.fromstring(document))
                appended += 1
            except ET.ParseError:
                logger.debug("Documento de evento ignorado por parse invalido durante consolidacao.")

        if appended == 0:
            return documents[0]

        if appended == 1:
            return ET.tostring(wrapper[0], encoding="utf-8", xml_declaration=True)

        return ET.tostring(wrapper, encoding="utf-8", xml_declaration=True)

    def _contains_cancelamento(self, xml_bytes: bytes) -> bool:
        try:
            root = ET.fromstring(xml_bytes)
        except ET.ParseError:
            return False

        for element in root.iter():
            tag = str(element.tag).lower()
            text = (element.text or "").strip().lower()
            if any(token in tag for token in ("cancel", "canc")):
                return True
            if any(token in text for token in ("cancelamento", "cancelado", "cancelada", "cancelar")):
                return True

        return False

    def _salvar_arquivo(self, dest_path: str, content: bytes) -> None:
        with open(dest_path, "wb") as file:
            file.write(content)

    def _respect_delay(self) -> None:
        if self.delay_seconds <= 0:
            return

        now = time.monotonic()
        elapsed = now - self._last_request_at
        wait_seconds = self.delay_seconds - elapsed

        if wait_seconds > 0:
            time.sleep(wait_seconds)

    def _mark_request(self) -> None:
        self._last_request_at = time.monotonic()


class ParallelEventDownloader:
    """Baixa eventos em background sem travar o loop principal de NSU/XML."""

    def __init__(
        self,
        pem_cert: str,
        timeout: int = 30,
        delay_seconds: float = 2.0,
        max_retries: int = 3,
        retry_backoff: float = 2.0,
        logger_instance: Optional[logging.Logger] = None,
        on_success: Optional[Callable[[EventDownloadTask, EventFetchResult], None]] = None,
        on_empty: Optional[Callable[[EventDownloadTask], None]] = None,
        on_failure: Optional[Callable[[EventDownloadTask, str], None]] = None,
    ) -> None:
        self.pem_cert = pem_cert
        self.timeout = timeout
        self.delay_seconds = max(0.0, float(delay_seconds))
        self.max_retries = max(0, int(max_retries))
        self.retry_backoff = max(1.0, float(retry_backoff))
        self.logger = logger_instance or logger
        self.on_success = on_success
        self.on_empty = on_empty
        self.on_failure = on_failure
        self._queue: queue.Queue[EventDownloadTask | None] = queue.Queue()
        self._thread: Optional[threading.Thread] = None
        self._worker_error: Optional[Exception] = None
        self.salvos = 0
        self.sem_evento = 0
        self.falhas = 0
        self.cancelamentos = 0

    def start(self) -> None:
        if self._thread and self._thread.is_alive():
            return

        self._thread = threading.Thread(
            target=self._worker,
            name="nfse-evento-worker",
            daemon=True,
        )
        self._thread.start()

    def enqueue(self, task: EventDownloadTask) -> None:
        self.start()
        self._queue.put(task)

    def finish(self) -> dict[str, object]:
        if self._thread is None:
            return {
                "salvos": 0,
                "sem_evento": 0,
                "falhas": 0,
                "cancelamentos": 0,
                "erro": None,
            }

        self._queue.put(None)
        self._thread.join()
        self._thread = None

        return {
            "salvos": self.salvos,
            "sem_evento": self.sem_evento,
            "falhas": self.falhas,
            "cancelamentos": self.cancelamentos,
            "erro": self._worker_error,
        }

    def _notify_success(self, task: EventDownloadTask, result: EventFetchResult) -> None:
        if not self.on_success:
            return

        try:
            self.on_success(task, result)
        except Exception as exc:
            self.logger.error("Erro no callback de sucesso do evento %s: %s", task.chave, exc)

    def _notify_empty(self, task: EventDownloadTask) -> None:
        if not self.on_empty:
            return

        try:
            self.on_empty(task)
        except Exception as exc:
            self.logger.error("Erro no callback de evento vazio %s: %s", task.chave, exc)

    def _notify_failure(self, task: EventDownloadTask, reason: str) -> None:
        if not self.on_failure:
            return

        try:
            self.on_failure(task, reason)
        except Exception as exc:
            self.logger.error("Erro no callback de falha do evento %s: %s", task.chave, exc)

    def _worker(self) -> None:
        session = requests.Session()
        session.cert = self.pem_cert
        session.verify = True
        downloader = NFSeEventDownloader(
            session,
            self.timeout,
            delay_seconds=self.delay_seconds,
            max_retries=self.max_retries,
            retry_backoff=self.retry_backoff,
        )

        self.logger.info(
            "Worker de eventos iniciado com delay=%.1fs, retries=%s, backoff=%.1f.",
            self.delay_seconds,
            self.max_retries,
            self.retry_backoff,
        )

        try:
            while True:
                task = self._queue.get()
                if task is None:
                    break

                try:
                    result = downloader.baixar(task.chave, task.dest_path)

                    if result.status == "saved":
                        self.salvos += result.saved_count
                        if result.cancelamento:
                            self.cancelamentos += 1
                        self._notify_success(task, result)
                    elif result.status == "no_events":
                        self.sem_evento += 1
                        self._notify_empty(task)
                    else:
                        self.falhas += 1
                        self._notify_failure(task, result.reason or "Falha na consulta")
                except Exception as exc:
                    self.falhas += 1
                    self._notify_failure(task, str(exc))
                finally:
                    self._queue.task_done()
        except Exception as exc:
            self._worker_error = exc
            self.logger.error("Worker de eventos encerrado com erro: %s", exc)
        finally:
            session.close()
