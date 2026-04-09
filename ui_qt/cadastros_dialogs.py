from __future__ import annotations

from PySide6 import QtCore, QtGui, QtWidgets

from config.utils import formatar_cnpj, limpar_cnpj
from ui_qt.services.cadastros_service import (
    clear_nsu_records,
    copy_certificate_for_company,
    delete_nsu_record,
    list_company_forms,
    nsu_rows_for_company,
    upsert_company,
    upsert_nsu_record,
)
from ui_qt.theme import field_label, hint_label, make_action_button, make_card, set_card_padding, title_label


class CompanyDialog(QtWidgets.QDialog):
    def __init__(self, parent=None, company: dict | None = None):
        super().__init__(parent)
        self.company = company
        self.saved_company: dict | None = None
        self.setWindowTitle("Editar cadastro" if company else "Novo cadastro")
        self.setMinimumWidth(620)

        root = QtWidgets.QVBoxLayout(self)
        root.setContentsMargins(18, 18, 18, 18)
        root.setSpacing(12)

        header = make_card(self, "Hero")
        header_l = QtWidgets.QVBoxLayout(header)
        set_card_padding(header_l, 20, 18, 20, 18)
        header_l.addWidget(title_label("Cadastro de empresa", header))
        header_l.addWidget(
            hint_label(
                "Os dados salvos aqui usam a mesma base do programa e preparam a estrutura f\u00edsica da empresa.",
                header,
            )
        )
        root.addWidget(header)

        form_card = make_card(self)
        form_l = QtWidgets.QGridLayout(form_card)
        set_card_padding(form_l, 18, 16, 18, 16)
        form_l.setHorizontalSpacing(12)
        form_l.setVerticalSpacing(10)

        self.input_cod = QtWidgets.QLineEdit()
        self.input_cod.setValidator(QtGui.QIntValidator(1, 999999999, self))
        self.input_empresa = QtWidgets.QLineEdit()
        self.input_cnpj = QtWidgets.QLineEdit()
        self.input_cert_pass = QtWidgets.QLineEdit()
        self.input_cert_pass.setEchoMode(QtWidgets.QLineEdit.Password)
        self.input_cert_path = QtWidgets.QLineEdit()
        self.input_cert_path.setReadOnly(True)
        self.input_venc = QtWidgets.QLineEdit()
        self.input_venc.setReadOnly(True)

        browse_wrap = QtWidgets.QHBoxLayout()
        browse_wrap.setContentsMargins(0, 0, 0, 0)
        browse_wrap.setSpacing(8)
        browse_wrap.addWidget(self.input_cert_path, 1)
        self.btn_browse = make_action_button("Importar PFX", "GhostAction", form_card)
        browse_wrap.addWidget(self.btn_browse)

        fields = [
            ("C\u00f3digo", self.input_cod),
            ("Empresa", self.input_empresa),
            ("CNPJ", self.input_cnpj),
            ("Senha do certificado", self.input_cert_pass),
            ("Certificado (.pfx)", browse_wrap),
            ("Vencimento", self.input_venc),
        ]

        for row, (label, widget) in enumerate(fields):
            form_l.addWidget(field_label(label, form_card), row * 2, 0)
            if isinstance(widget, QtWidgets.QLayout):
                container = QtWidgets.QWidget(form_card)
                container.setLayout(widget)
                form_l.addWidget(container, row * 2 + 1, 0, 1, 2)
            else:
                form_l.addWidget(widget, row * 2 + 1, 0, 1, 2)

        form_l.addWidget(
            hint_label(
                "Informe a senha antes de importar o certificado para permitir a leitura autom\u00e1tica do vencimento.",
                form_card,
            ),
            len(fields) * 2,
            0,
            1,
            2,
        )
        root.addWidget(form_card)

        footer = QtWidgets.QHBoxLayout()
        footer.addStretch(1)
        self.btn_save = make_action_button("Salvar", "PrimaryAction", self)
        self.btn_cancel = make_action_button("Cancelar", "GhostAction", self)
        footer.addWidget(self.btn_save)
        footer.addWidget(self.btn_cancel)
        root.addLayout(footer)

        self.input_cnpj.textEdited.connect(self._format_cnpj_input)
        self.btn_browse.clicked.connect(self._browse_certificate)
        self.btn_save.clicked.connect(self._save)
        self.btn_cancel.clicked.connect(self.reject)

        if company:
            self._load_company()
        else:
            self.input_cod.setText(str(self._next_code()))

    def _load_company(self) -> None:
        if not self.company:
            return

        self.input_cod.setText(str(self.company["cod"]))
        self.input_cod.setReadOnly(True)
        self.input_empresa.setText(str(self.company.get("empresa", "")))
        self.input_cnpj.setText(str(self.company.get("cnpj_formatado") or formatar_cnpj(self.company.get("cnpj", ""))))
        self.input_cert_pass.setText(str(self.company.get("cert_pass", "")))
        self.input_cert_path.setText(str(self.company.get("cert_path", "")))
        self.input_venc.setText(str(self.company.get("venc", "")))

    def _format_cnpj_input(self, _text: str) -> None:
        current = self.input_cnpj.text()
        formatted = formatar_cnpj(limpar_cnpj(current))
        if formatted == current:
            return

        blocker = QtCore.QSignalBlocker(self.input_cnpj)
        self.input_cnpj.setText(formatted)
        del blocker
        self.input_cnpj.setCursorPosition(len(self.input_cnpj.text()))

    def _browse_certificate(self) -> None:
        raw_cod = self.input_cod.text().strip()
        cert_pass = self.input_cert_pass.text()

        if not raw_cod:
            QtWidgets.QMessageBox.warning(
                self, "C\u00f3digo necess\u00e1rio", "Informe o c\u00f3digo antes de importar o certificado."
            )
            return
        if not cert_pass:
            QtWidgets.QMessageBox.warning(
                self, "Senha necess\u00e1ria", "Informe a senha antes de importar o certificado."
            )
            return

        path, _ = QtWidgets.QFileDialog.getOpenFileName(self, "Selecionar certificado", "", "Certificados PFX (*.pfx)")
        if not path:
            return

        try:
            relative_path, vencimento = copy_certificate_for_company(path, raw_cod, cert_pass)
        except Exception as exc:
            QtWidgets.QMessageBox.critical(self, "Erro", f"Falha ao importar certificado:\n{exc}")
            return

        self.input_cert_path.setText(relative_path)
        self.input_venc.setText(vencimento)
        if vencimento:
            QtWidgets.QMessageBox.information(
                self,
                "Certificado importado",
                f"Certificado salvo com sucesso.\nVencimento: {vencimento}",
            )
        else:
            QtWidgets.QMessageBox.warning(
                self,
                "Certificado importado",
                "Certificado salvo com sucesso, mas n\u00e3o foi poss\u00edvel ler o vencimento.",
            )

    def _save(self) -> None:
        try:
            self.saved_company = upsert_company(
                {
                    "cod": self.input_cod.text().strip(),
                    "empresa": self.input_empresa.text().strip(),
                    "cnpj": self.input_cnpj.text().strip(),
                    "cert_pass": self.input_cert_pass.text(),
                    "cert_path": self.input_cert_path.text().strip(),
                    "venc": self.input_venc.text().strip(),
                },
                edit_code=int(self.company["cod"]) if self.company else None,
            )
        except ValueError as exc:
            QtWidgets.QMessageBox.warning(self, "Valida\u00e7\u00e3o", str(exc))
            return
        except Exception as exc:
            QtWidgets.QMessageBox.critical(self, "Erro", f"N\u00e3o foi poss\u00edvel salvar o cadastro:\n{exc}")
            return

        self.accept()

    def _next_code(self) -> int:
        companies = list_company_forms()
        if not companies:
            return 1
        return max(int(company["cod"]) for company in companies) + 1


class NSUEditorDialog(QtWidgets.QDialog):
    def __init__(self, company: dict[str, object], parent=None):
        super().__init__(parent)
        self.company = company
        self.setWindowTitle(f"Editar NSU - {company['empresa']}")
        self.resize(760, 620)

        root = QtWidgets.QVBoxLayout(self)
        root.setContentsMargins(18, 18, 18, 18)
        root.setSpacing(12)

        header = make_card(self, "Hero")
        header_l = QtWidgets.QVBoxLayout(header)
        set_card_padding(header_l, 20, 18, 20, 18)
        header_l.addWidget(title_label("Controle de NSU", header))
        header_l.addWidget(
            hint_label(
                "Os registros abaixo s\u00e3o gravados no arquivo nsu_competencia.json da empresa.",
                header,
            )
        )
        root.addWidget(header)

        form_card = make_card(self)
        form_l = QtWidgets.QGridLayout(form_card)
        set_card_padding(form_l, 18, 16, 18, 16)
        form_l.setHorizontalSpacing(12)
        form_l.setVerticalSpacing(10)

        self.input_year = QtWidgets.QSpinBox()
        self.input_year.setRange(2000, 2100)
        self.input_year.setValue(QtCore.QDate.currentDate().year())
        self.input_month = QtWidgets.QSpinBox()
        self.input_month.setRange(1, 12)
        self.input_month.setValue(QtCore.QDate.currentDate().month())
        self.input_start = QtWidgets.QLineEdit("0")
        self.input_end = QtWidgets.QLineEdit("0")
        validator = QtGui.QRegularExpressionValidator(QtCore.QRegularExpression(r"\d{0,20}"))
        self.input_start.setValidator(validator)
        self.input_end.setValidator(validator)

        form_l.addWidget(field_label("Ano", form_card), 0, 0)
        form_l.addWidget(self.input_year, 1, 0)
        form_l.addWidget(field_label("M\u00eas", form_card), 0, 1)
        form_l.addWidget(self.input_month, 1, 1)
        form_l.addWidget(field_label("NSU inicial", form_card), 0, 2)
        form_l.addWidget(self.input_start, 1, 2)
        form_l.addWidget(field_label("NSU final", form_card), 0, 3)
        form_l.addWidget(self.input_end, 1, 3)

        actions = QtWidgets.QHBoxLayout()
        actions.setSpacing(8)
        self.btn_add = make_action_button("Salvar registro", "PrimaryAction", form_card)
        self.btn_delete = make_action_button("Excluir selecionado", "Danger", form_card)
        self.btn_clear = make_action_button("Excluir Todos", "GhostAction", form_card)
        actions.addWidget(self.btn_add)
        actions.addWidget(self.btn_delete)
        actions.addWidget(self.btn_clear)
        actions.addStretch(1)
        form_l.addLayout(actions, 2, 0, 1, 4)
        root.addWidget(form_card)

        table_card = make_card(self)
        table_l = QtWidgets.QVBoxLayout(table_card)
        set_card_padding(table_l, 18, 14, 18, 14)
        table_l.addWidget(title_label("Registros salvos", table_card))

        self.table = QtWidgets.QTableWidget(0, 4)
        self.table.setHorizontalHeaderLabels(["Ano", "M\u00eas", "NSU Inicial", "NSU Final"])
        self.table.verticalHeader().setVisible(False)
        self.table.horizontalHeader().setStretchLastSection(True)
        self.table.setSelectionBehavior(QtWidgets.QAbstractItemView.SelectRows)
        self.table.setSelectionMode(QtWidgets.QAbstractItemView.SingleSelection)
        self.table.setEditTriggers(QtWidgets.QAbstractItemView.NoEditTriggers)
        self.table.setShowGrid(False)
        table_l.addWidget(self.table, 1)
        root.addWidget(table_card, 1)

        footer = QtWidgets.QHBoxLayout()
        footer.addStretch(1)
        self.btn_close = make_action_button("Fechar", "GhostAction", self)
        footer.addWidget(self.btn_close)
        root.addLayout(footer)

        self.btn_add.clicked.connect(self._save_record)
        self.btn_delete.clicked.connect(self._delete_selected)
        self.btn_clear.clicked.connect(self._clear_all)
        self.btn_close.clicked.connect(self.accept)
        self.table.itemSelectionChanged.connect(self._sync_inputs_from_selection)

        self._refresh_table()

    def _refresh_table(self) -> None:
        rows = nsu_rows_for_company(int(self.company["cod"]))
        self.table.setRowCount(len(rows))

        for index, row in enumerate(rows):
            self.table.setItem(index, 0, QtWidgets.QTableWidgetItem(row["ano"]))
            self.table.setItem(index, 1, QtWidgets.QTableWidgetItem(row["mes"]))
            self.table.setItem(index, 2, QtWidgets.QTableWidgetItem(str(row["nsu_inicial"])))
            self.table.setItem(index, 3, QtWidgets.QTableWidgetItem(str(row["nsu_final"])))

        self.btn_delete.setEnabled(bool(rows))

    def _sync_inputs_from_selection(self) -> None:
        selected = self.table.selectedItems()
        if not selected:
            self.btn_delete.setEnabled(False)
            return

        row = selected[0].row()
        self.input_year.setValue(int(self.table.item(row, 0).text()))
        self.input_month.setValue(int(self.table.item(row, 1).text()))
        self.input_start.setText(self.table.item(row, 2).text())
        self.input_end.setText(self.table.item(row, 3).text())
        self.btn_delete.setEnabled(True)

    def _save_record(self) -> None:
        try:
            upsert_nsu_record(
                int(self.company["cod"]),
                self.input_year.value(),
                self.input_month.value(),
                self.input_start.text().strip() or "0",
                self.input_end.text().strip() or "0",
            )
        except ValueError as exc:
            QtWidgets.QMessageBox.warning(self, "Valida\u00e7\u00e3o", str(exc))
            return
        except Exception as exc:
            QtWidgets.QMessageBox.critical(self, "Erro", f"N\u00e3o foi poss\u00edvel salvar o registro:\n{exc}")
            return

        self._refresh_table()

    def _delete_selected(self) -> None:
        selected = self.table.selectedItems()
        if not selected:
            return

        row = selected[0].row()
        ano = self.table.item(row, 0).text()
        mes = self.table.item(row, 1).text()
        if (
            QtWidgets.QMessageBox.question(
                self,
                "Confirma\u00e7\u00e3o",
                f"Excluir o registro {mes}/{ano}?",
            )
            != QtWidgets.QMessageBox.Yes
        ):
            return

        delete_nsu_record(int(self.company["cod"]), ano, mes)
        self._refresh_table()

    def _clear_all(self) -> None:
        if (
            QtWidgets.QMessageBox.question(
                self,
                "Confirma\u00e7\u00e3o",
                "Excluir todos os registros de NSU desta empresa?",
            )
            != QtWidgets.QMessageBox.Yes
        ):
            return

        clear_nsu_records(int(self.company["cod"]))
        self._refresh_table()
