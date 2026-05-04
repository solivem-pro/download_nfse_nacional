from __future__ import annotations

from PySide6 import QtCore

from ui_qt.services.download_service import run_download_batch


class DownloadWorker(QtCore.QObject):
    eventEmitted = QtCore.Signal(object)
    finished = QtCore.Signal(object)
    failed = QtCore.Signal(str)

    def __init__(self, companies: list[dict], year: str, month: str):
        super().__init__()
        self._companies = companies
        self._year = year
        self._month = month

    @QtCore.Slot()
    def run(self) -> None:
        try:
            results = run_download_batch(
                self._companies,
                self._year,
                self._month,
                on_event=lambda payload: self.eventEmitted.emit(payload),
            )
        except Exception as exc:
            self.failed.emit(str(exc))
            return

        self.finished.emit(results)
