from __future__ import annotations

from PySide6 import QtCore, QtGui, QtWidgets

from ui_qt.data_bridge import available_years, default_period, load_company_rows, summarize_company_rows
from ui_qt.download_worker import DownloadWorker
from ui_qt.services.download_service import export_company_archives, prepare_download_companies
from ui_qt.theme import COL_DANGER, COL_WARN, field_label, hint_label, make_action_button, make_card, set_card_padding


class DownloadPage(QtWidgets.QWidget):

    def __init__(self, parent=None):
        super().__init__(parent)
        self._rows: list[dict] = []
        self._thread: QtCore.QThread | None = None
        self._worker: DownloadWorker | None = None
        self._running = False
        self._pending_skipped: list[dict] = []
        self._pending_missing: list[str] = []
        self.setObjectName("DownloadPage")
        self.setStyleSheet(
            """
            QLabel#FieldLabel {
                color: #a2abb7;
                font-size: 10px;
                font-weight: 700;
                letter-spacing: 0.8px;
            }
            QLabel#MetaHint {
                color: #7d8591;
            }
            QTableWidget {
                alternate-background-color: #171b22;
            }
            QHeaderView::section {
                padding: 8px 8px;
            }
            QTableWidget::item {
                padding: 8px 8px;
            }
            QProgressBar {
                height: 14px;
            }
            """
        )

        root = QtWidgets.QVBoxLayout(self)
        root.setContentsMargins(0, 0, 0, 0)
        root.setSpacing(12)

        filter_card = make_card(self)
        filter_l = QtWidgets.QGridLayout(filter_card)
        set_card_padding(filter_l, 16, 12, 16, 12)
        filter_l.setHorizontalSpacing(10)
        filter_l.setVerticalSpacing(8)

        filter_l.addWidget(field_label("Ano"), 0, 0)
        self.combo_year = QtWidgets.QComboBox()
        self.combo_year.addItems(available_years())
        filter_l.addWidget(self.combo_year, 1, 0)

        filter_l.addWidget(field_label("M\u00eas"), 0, 1)
        self.combo_month = QtWidgets.QComboBox()
        self.combo_month.addItems([f"{month:02d}" for month in range(1, 13)])
        filter_l.addWidget(self.combo_month, 1, 1)

        filter_l.addWidget(field_label("Status"), 0, 2)
        self.combo_status = QtWidgets.QComboBox()
        self.combo_status.addItems(["Todos", "V\u00e1lidos", "Vencendo", "Vencidos"])
        filter_l.addWidget(self.combo_status, 1, 2)

        btn_row = QtWidgets.QHBoxLayout()
        btn_row.setSpacing(10)
        self.btn_select_all = make_action_button("Selec. Todos", "GhostAction", filter_card)
        self.btn_download = make_action_button("Baixar", "PrimaryAction", filter_card)
        self.btn_export = make_action_button("Exportar", "AccentAction", filter_card)
        for button in (self.btn_select_all, self.btn_download, self.btn_export):
            button.setMinimumHeight(34)
        btn_row.addWidget(self.btn_select_all)
        btn_row.addWidget(self.btn_download)
        btn_row.addWidget(self.btn_export)
        btn_row.addStretch(1)
        filter_l.addLayout(btn_row, 2, 0, 1, 3)
        root.addWidget(filter_card)

        table_card = make_card(self)
        table_l = QtWidgets.QVBoxLayout(table_card)
        set_card_padding(table_l, 16, 12, 16, 12)
        table_l.setSpacing(6)
        table_l.addWidget(field_label("EMPRESAS", table_card))
        self.table = QtWidgets.QTableWidget(0, 4)
        self.table.setHorizontalHeaderLabels(["C\u00f3digo", "Empresa", "CNPJ", "Certificado"])
        self.table.verticalHeader().setVisible(False)
        self.table.horizontalHeader().setStretchLastSection(True)
        self.table.setAlternatingRowColors(True)
        self.table.setSelectionBehavior(QtWidgets.QAbstractItemView.SelectRows)
        self.table.setSelectionMode(QtWidgets.QAbstractItemView.ExtendedSelection)
        self.table.setEditTriggers(QtWidgets.QAbstractItemView.NoEditTriggers)
        self.table.setShowGrid(False)
        self.table.setWordWrap(False)
        self.table.verticalHeader().setDefaultSectionSize(28)
        self.table.setColumnWidth(0, 90)
        self.table.setColumnWidth(2, 170)
        table_l.addWidget(self.table, 1)
        root.addWidget(table_card, 1)

        progress_card = make_card(self)
        progress_l = QtWidgets.QVBoxLayout(progress_card)
        set_card_padding(progress_l, 16, 12, 16, 12)
        progress_l.setSpacing(5)
        progress_l.addWidget(field_label("PROGRESSO", progress_card))
        self.lbl_progress = QtWidgets.QLabel("Nenhum processo em andamento.")
        self.lbl_progress.setObjectName("MetaHint")
        self.bar = QtWidgets.QProgressBar()
        self.bar.setRange(0, 100)
        self.bar.setValue(0)
        self.bar.setTextVisible(True)
        self.bar.setFixedHeight(16)
        progress_l.addWidget(self.lbl_progress)
        progress_l.addWidget(self.bar)
        progress_l.addWidget(
            hint_label("Acompanhe a empresa atual, o NSU em andamento e o total processado.", progress_card)
        )
        root.addWidget(progress_card)

        year, month = default_period()
        self.combo_year.setCurrentText(year)
        self.combo_month.setCurrentText(month)
        self.combo_status.currentTextChanged.connect(self.refresh)
        self.combo_year.currentTextChanged.connect(self.refresh)
        self.combo_month.currentTextChanged.connect(self.refresh)
        self.btn_select_all.clicked.connect(self._select_all_visible)
        self.btn_download.clicked.connect(self._start_download)
        self.btn_export.clicked.connect(self._export_selected)
        self.table.itemSelectionChanged.connect(self._update_action_state)
        self.refresh()

    def refresh(self) -> None:
        if self._running:
            return

        rows = load_company_rows()
        summary = summarize_company_rows(rows)
        selected_filter = self.combo_status.currentText()

        filtered = []
        for row in rows:
            if selected_filter == "Vencidos" and not row["cert_expired"]:
                continue
            if selected_filter == "Vencendo" and not row["cert_warning"]:
                continue
            if selected_filter == "V\u00e1lidos" and (row["cert_expired"] or row["cert_warning"]):
                continue
            filtered.append(row)

        self._rows = filtered
        self.table.setRowCount(len(filtered))
        for index, row in enumerate(filtered):
            values = [str(row["cod"]), row["empresa"], row["cnpj_formatado"], row["cert_status"]]
            for column, value in enumerate(values):
                item = QtWidgets.QTableWidgetItem(value)
                if column == 0:
                    item.setData(QtCore.Qt.UserRole, str(row["cod"]))
                self.table.setItem(index, column, item)
            self._apply_row_style(index, row)

        self.lbl_progress.setText(
            f"{len(filtered)} empresa(s) vis\u00edveis | per\u00edodo {self.combo_month.currentText()}/{self.combo_year.currentText()} | {summary['expired']} certificado(s) vencido(s)"
        )
        self._update_action_state()

    def _apply_row_style(self, row_index: int, row: dict) -> None:
        if row.get("cert_expired"):
            color = QtGui.QColor(COL_DANGER)
        elif row.get("cert_warning"):
            color = QtGui.QColor(COL_WARN)
        else:
            return

        for column in range(self.table.columnCount()):
            item = self.table.item(row_index, column)
            if item is None:
                continue
            item.setForeground(color)
            font = item.font()
            font.setBold(True)
            item.setFont(font)

    def _selected_codes(self) -> list[str]:
        rows = self.table.selectionModel().selectedRows()
        codes = []
        for row in rows:
            item = self.table.item(row.row(), 0)
            if item is not None:
                codes.append(str(item.data(QtCore.Qt.UserRole)))
        return codes

    def _select_all_visible(self) -> None:
        self.table.selectAll()
        self._update_action_state()

    def _update_action_state(self) -> None:
        has_selection = bool(self._selected_codes())
        self.btn_export.setEnabled(has_selection and not self._running)
        self.btn_select_all.setEnabled(bool(self._rows) and not self._running)

        if not has_selection or self._running:
            self.btn_download.setEnabled(False)
            return

        selected_set = set(self._selected_codes())
        has_valid = any(str(row["cod"]) in selected_set and not row.get("cert_expired") for row in self._rows)
        self.btn_download.setEnabled(has_valid)

    def _set_running_state(self, running: bool) -> None:
        self._running = running
        self.combo_year.setEnabled(not running)
        self.combo_month.setEnabled(not running)
        self.combo_status.setEnabled(not running)
        self.table.setEnabled(not running)
        self._update_action_state()

    def _export_selected(self) -> None:
        codes = self._selected_codes()
        if not codes:
            QtWidgets.QMessageBox.warning(self, "Sele\u00e7\u00e3o", "Selecione pelo menos uma empresa para exportar.")
            return

        destination = QtWidgets.QFileDialog.getExistingDirectory(self, "Selecionar pasta de destino")
        if not destination:
            return

        outcome = export_company_archives(codes, destination)
        message = self._build_export_summary(outcome)
        QtWidgets.QMessageBox.information(self, "Exporta\u00e7\u00e3o Concluida", message)

    def _start_download(self) -> None:
        codes = self._selected_codes()
        if not codes:
            QtWidgets.QMessageBox.warning(self, "Sele\u00e7\u00e3o", "Selecione pelo menos uma empresa para baixar.")
            return

        valid_companies, skipped_expired, missing = prepare_download_companies(codes)
        self._pending_skipped = skipped_expired
        self._pending_missing = missing

        if skipped_expired:
            lines = [f"[{company['cod']}] {company['nome']}" for company in skipped_expired]
            QtWidgets.QMessageBox.warning(
                self,
                "Certificados Vencidos",
                "As empresas abaixo foram ignoradas no download:\n\n" + "\n".join(lines[:15]),
            )

        if not valid_companies:
            QtWidgets.QMessageBox.warning(
                self,
                "Nenhuma Empresa Valida",
                "Nenhuma empresa valida ficou disponivel para o download.",
            )
            return

        self._set_running_state(True)
        self.bar.setRange(0, len(valid_companies))
        self.bar.setValue(0)
        self.lbl_progress.setText("Preparando ambiente de download...")

        self._thread = QtCore.QThread(self)
        self._worker = DownloadWorker(valid_companies, self.combo_year.currentText(), self.combo_month.currentText())
        self._worker.moveToThread(self._thread)
        self._thread.started.connect(self._worker.run)
        self._worker.eventEmitted.connect(self._handle_worker_event)
        self._worker.finished.connect(self._handle_worker_finished)
        self._worker.failed.connect(self._handle_worker_failed)
        self._worker.finished.connect(self._thread.quit)
        self._worker.failed.connect(self._thread.quit)
        self._thread.finished.connect(self._cleanup_worker)
        self._thread.start()

    def _handle_worker_event(self, event: dict) -> None:
        kind = str(event.get("kind", ""))
        index = int(event.get("index", 0) or 0)
        total = int(event.get("total", 0) or 0)
        code = event.get("code", "")
        name = event.get("name", "")
        nsu = int(event.get("nsu", 0) or 0)
        documents = int(event.get("documents", 0) or 0)
        message = str(event.get("message", "") or "")

        if kind == "company_start":
            self.lbl_progress.setText(
                f"Empresa {index}/{total}: [{code}] {name} | NSU {nsu} | {documents} documento(s)"
            )
            return

        if kind in {"progress", "document"}:
            self.lbl_progress.setText(
                f"Empresa {index}/{total}: [{code}] {name} | NSU {nsu} | {documents} documento(s)"
            )
            return

        if kind == "company_complete":
            self.bar.setValue(index)
            result = event.get("result", {}) or {}
            self.lbl_progress.setText(
                f"Conclu\u00edda empresa {index}/{total}: [{code}] {name} | {int(result.get('documentos', 0))} documento(s)"
            )
            return

    def _handle_worker_finished(self, results: list[dict]) -> None:
        self._set_running_state(False)
        total_results = len(results)
        total_docs = sum(int(result.get("documentos", 0)) for result in results)
        total_errors = sum(int(result.get("erros", 0)) for result in results)
        self.bar.setValue(self.bar.maximum())
        self.lbl_progress.setText(
            f"Download finalizado | {total_results} empresa(s) | {total_docs} documento(s) | {total_errors} erro(s)"
        )
        QtWidgets.QMessageBox.information(self, "Resumo do Download", self._build_download_summary(results))

    def _handle_worker_failed(self, message: str) -> None:
        self._set_running_state(False)
        self.lbl_progress.setText("Falha no processamento do download.")
        QtWidgets.QMessageBox.critical(self, "Erro no Download", message)

    def _cleanup_worker(self) -> None:
        if self._worker is not None:
            self._worker.deleteLater()
        if self._thread is not None:
            self._thread.deleteLater()
        self._worker = None
        self._thread = None
        self._update_action_state()

    def _build_export_summary(self, outcome: dict) -> str:
        exported = outcome.get("exported", [])
        missing = outcome.get("missing", [])
        errors = outcome.get("errors", [])
        lines = [f"Modo de salvamento: {outcome.get('save_mode', 'C\u00f3digo')}"]

        if exported:
            lines.append("")
            lines.append(f"{len(exported)} arquivo(s) exportado(s):")
            lines.extend(f"[{item['cod']}] {item['empresa']} -> {item['arquivo']}" for item in exported[:15])

        if missing:
            lines.append("")
            lines.append(f"{len(missing)} arquivo(s) sem pacote ZIP:")
            lines.extend(f"[{item['cod']}] {item['empresa']}" for item in missing[:10])

        if errors:
            lines.append("")
            lines.append(f"{len(errors)} erro(s) durante a exporta\u00e7\u00e3o:")
            lines.extend(f"[{item['cod']}] {item['empresa']} -> {item['erro']}" for item in errors[:10])

        if not exported and not missing and not errors:
            lines.append("")
            lines.append("Nenhum arquivo foi processado.")

        return "\n".join(lines)

    def _build_download_summary(self, results: list[dict]) -> str:
        total_docs = sum(int(result.get("documentos", 0)) for result in results)
        total_errors = sum(int(result.get("erros", 0)) for result in results)
        lines = [
            f"Empresas processadas: {len(results)}",
            f"Documentos baixados: {total_docs}",
            f"Erros: {total_errors}",
        ]

        if self._pending_skipped:
            lines.append("")
            lines.append(f"Empresas ignoradas por certificado vencido: {len(self._pending_skipped)}")
            lines.extend(f"[{item['cod']}] {item['nome']}" for item in self._pending_skipped[:10])

        if self._pending_missing:
            lines.append("")
            lines.append("C\u00f3digos sem cadastro encontrado: " + ", ".join(self._pending_missing))

        lines.append("")
        lines.append("Detalhes:")
        for result in results[:20]:
            status = "OK" if int(result.get("erros", 0)) == 0 else "ERRO"
            lines.append(
                f"[{result.get('cod', 'N/A')}] {result['empresa']} | {status} | "
                f"Notas: {int(result.get('documentos', 0))} | {result.get('mensagem', '')}"
            )
        if len(results) > 20:
            lines.append(f"... e mais {len(results) - 20} empresa(s).")

        return "\n".join(lines)
