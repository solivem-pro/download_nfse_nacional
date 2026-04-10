import logging
import queue
import threading
import time
from dataclasses import dataclass
from typing import Callable, Optional

import requests

logger = logging.getLogger(__name__)
RETRY_STATUS_CODES = {429, 502, 503, 504}


@dataclass(slots=True)
class PDFDownloadTask:
    """Representa um PDF pendente para download por chave."""

    chave: str
    dest_path: str
    nsu: int
    ano: str
    mes: str
    tipo_documento: str


class NFSePDFDownloader:
    """Downloader para documentos PDF do portal nacional de NFS-e."""

    BASE_URL = "https://adn.nfse.gov.br/danfse"

    def __init__(
        self,
        session,
        timeout: int = 30,
        delay_seconds: float = 1.5,
        max_retries: int = 3,
        retry_backoff: float = 2.0,
    ):
        self.session = session
        self.timeout = timeout
        self.delay_seconds = max(0.0, float(delay_seconds))
        self.max_retries = max(0, int(max_retries))
        self.retry_backoff = max(1.0, float(retry_backoff))
        self._last_request_at = 0.0

    def baixar(self, chave: str, dest_path: str) -> bool:
        """
        Baixa um PDF da chave fornecida.

        Args:
            chave: Chave da NFS-e
            dest_path: Caminho de destino do arquivo

        Returns:
            True se download bem-sucedido, False caso contrario
        """
        url = f"{self.BASE_URL}/{chave}"

        attempt = 0

        while True:
            self._respect_delay()

            try:
                with self.session.get(url, timeout=self.timeout) as resp:
                    self._mark_request()

                    if resp.status_code == 200:
                        self._salvar_arquivo(dest_path, resp.content)
                        logger.info("PDF baixado com sucesso: %s", chave)
                        return True

                    if resp.status_code in RETRY_STATUS_CODES and attempt < self.max_retries:
                        wait_seconds = self.delay_seconds * (self.retry_backoff ** attempt)
                        logger.warning(
                            "Falha temporaria ao baixar PDF %s: HTTP %s. "
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

                    logger.error("Falha ao baixar PDF %s: HTTP %s", chave, resp.status_code)
                    return False

            except Exception as exc:
                self._mark_request()

                if attempt < self.max_retries:
                    wait_seconds = self.delay_seconds * (self.retry_backoff ** attempt)
                    logger.warning(
                        "Erro temporario ao baixar PDF %s: %s. "
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

                logger.error("Erro ao baixar PDF %s: %s", chave, str(exc))
                return False

    def _salvar_arquivo(self, dest_path: str, content: bytes) -> None:
        """Salva o conteudo no caminho especificado."""
        with open(dest_path, "wb") as file:
            file.write(content)

    def _respect_delay(self) -> None:
        """Mantem um intervalo minimo entre requests ao endpoint de PDF."""
        if self.delay_seconds <= 0:
            return

        now = time.monotonic()
        elapsed = now - self._last_request_at
        wait_seconds = self.delay_seconds - elapsed

        if wait_seconds > 0:
            time.sleep(wait_seconds)

    def _mark_request(self) -> None:
        """Marca o fim da ultima tentativa HTTP."""
        self._last_request_at = time.monotonic()

    def baixar_lote(self, chaves_destinos: list[tuple[str, str]]) -> tuple[int, int]:
        """
        Baixa multiplos PDFs em lote.

        Args:
            chaves_destinos: Lista de tuplas (chave, caminho_destino)

        Returns:
            Tupla (sucessos, falhas)
        """
        sucessos = 0
        falhas = 0

        for chave, destino in chaves_destinos:
            if self.baixar(chave, destino):
                sucessos += 1
            else:
                falhas += 1

        logger.info("Lote concluido: %d sucessos, %d falhas", sucessos, falhas)
        return sucessos, falhas


class ParallelPDFDownloader:
    """
    Baixa PDFs em background enquanto o loop principal continua no NSU/XML.

    Mantemos uma thread dedicada com sua propria sessao autenticada para o
    endpoint do DANFSe. Assim, o download do PDF por chave nao trava a varredura
    sequencial do endpoint de distribuicao por NSU.
    """

    def __init__(
        self,
        pem_cert: str,
        timeout: int = 30,
        delay_seconds: float = 1.5,
        max_retries: int = 3,
        retry_backoff: float = 2.0,
        logger_instance: Optional[logging.Logger] = None,
        on_success: Optional[Callable[[PDFDownloadTask], None]] = None,
        on_failure: Optional[Callable[[PDFDownloadTask, str], None]] = None,
    ) -> None:
        self.pem_cert = pem_cert
        self.timeout = timeout
        self.delay_seconds = max(0.0, float(delay_seconds))
        self.max_retries = max(0, int(max_retries))
        self.retry_backoff = max(1.0, float(retry_backoff))
        self.logger = logger_instance or logger
        self.on_success = on_success
        self.on_failure = on_failure
        self._queue: queue.Queue[PDFDownloadTask | None] = queue.Queue()
        self._thread: Optional[threading.Thread] = None
        self._worker_error: Optional[Exception] = None
        self.sucessos = 0
        self.falhas = 0

    def start(self) -> None:
        """Inicia a thread de trabalho, se necessario."""
        if self._thread and self._thread.is_alive():
            return

        self._thread = threading.Thread(
            target=self._worker,
            name="nfse-pdf-worker",
            daemon=True,
        )
        self._thread.start()

    def enqueue(self, task: PDFDownloadTask) -> None:
        """Enfileira um PDF para download."""
        self.start()
        self._queue.put(task)

    def finish(self) -> dict[str, object]:
        """
        Aguarda o worker concluir todos os PDFs ja enfileirados.

        O sentinel e inserido no fim da fila, entao tudo o que foi adicionado
        antes dele sera processado antes do encerramento.
        """
        if self._thread is None:
            return {"sucessos": 0, "falhas": 0, "erro": None}

        self._queue.put(None)
        self._thread.join()
        self._thread = None

        return {
            "sucessos": self.sucessos,
            "falhas": self.falhas,
            "erro": self._worker_error,
        }

    def _notify_success(self, task: PDFDownloadTask) -> None:
        if not self.on_success:
            return

        try:
            self.on_success(task)
        except Exception as exc:
            self.logger.error("Erro no callback de sucesso do PDF %s: %s", task.chave, exc)

    def _notify_failure(self, task: PDFDownloadTask, reason: str) -> None:
        if not self.on_failure:
            return

        try:
            self.on_failure(task, reason)
        except Exception as exc:
            self.logger.error("Erro no callback de falha do PDF %s: %s", task.chave, exc)

    def _worker(self) -> None:
        session = requests.Session()
        session.cert = self.pem_cert
        session.verify = True
        downloader = NFSePDFDownloader(
            session,
            self.timeout,
            delay_seconds=self.delay_seconds,
            max_retries=self.max_retries,
            retry_backoff=self.retry_backoff,
        )

        self.logger.info(
            "Worker de PDF iniciado com delay=%.1fs, retries=%s, backoff=%.1f.",
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
                    if downloader.baixar(task.chave, task.dest_path):
                        self.sucessos += 1
                        self._notify_success(task)
                    else:
                        self.falhas += 1
                        self._notify_failure(task, "Falha no download")
                except Exception as exc:
                    self.falhas += 1
                    self._notify_failure(task, str(exc))
                finally:
                    self._queue.task_done()
        except Exception as exc:
            self._worker_error = exc
            self.logger.error("Worker de PDF encerrado com erro: %s", exc)
        finally:
            session.close()
