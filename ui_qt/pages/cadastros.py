from __future__ import annotations

from PySide6 import QtCore, QtGui, QtWidgets

from ui_qt.cadastros_dialogs import CompanyDialog, NSUEditorDialog
from ui_qt.data_bridge import load_company_rows, summarize_company_rows
from ui_qt.services.cadastros_service import (
    delete_all_companies,
    delete_company,
    get_company_by_code,
    import_companies_from_xlsx,
    reset_all_nsu_files,
    suggested_import_spreadsheet_path,
)
from ui_qt.theme import COL_DANGER, COL_WARN, hint_label, make_action_button, make_card, set_card_padding, title_label


class CadastrosPage(QtWidgets.QWidget):
    backRequested = QtCore.Signal()
    companiesChanged = QtCore.Signal()

    def __init__(self, parent=None):
        super().__init__(parent)
        self._rows: list[dict] = []
        self.setStyleSheet(
            """
            QLabel {
                background: transparent;
            }
            QLabel#SectionTitle {
                color: #edf1f7;
            }
            QLabel#MetaHint {
                color: #8c94a0;
            }
            """
        )

        root = QtWidgets.QVBoxLayout(self)
        root.setContentsMargins(0, 0, 0, 0)
        root.setSpacing(11)

        header = make_card(self, "Hero")
        header_l = QtWidgets.QVBoxLayout(header)
        set_card_padding(header_l, 20, 16, 20, 16)
        header_l.setSpacing(5)
        header_l.addWidget(title_label("Cadastros", header))
        header_l.addWidget(hint_label("Gerencie os cadastros de empresas, certificados e controle de NSU.", header))
        root.addWidget(header)

        top = QtWidgets.QHBoxLayout()
        top.setSpacing(8)
        self.btn_add = make_action_button("Adicionar", "PrimaryOutlineAction", self)
        self.btn_import = make_action_button("Importar XLSX", "PrimaryOutlineAction", self)
        self.btn_edit = make_action_button("Editar", "AccentAction", self)
        self.btn_nsu = make_action_button("Editar NSU", "AccentAction", self)
        self.btn_delete = make_action_button("Excluir", "Danger", self)
        self.btn_reset_nsu = make_action_button("Resetar NSUs", "Danger", self)
        self.btn_delete_all = make_action_button("Excluir Todos", "Danger", self)
        for button in (
            self.btn_add,
            self.btn_import,
            self.btn_edit,
            self.btn_nsu,
            self.btn_delete,
            self.btn_reset_nsu,
            self.btn_delete_all,
        ):
            top.addWidget(button)
        top.addStretch(1)
        root.addLayout(top)

        table_card = make_card(self)
        table_l = QtWidgets.QVBoxLayout(table_card)
        set_card_padding(table_l, 18, 12, 18, 12)
        table_l.setSpacing(7)
        table_l.addWidget(title_label("Empresas cadastradas", table_card))

        self.table = QtWidgets.QTableWidget(0, 4)
        self.table.setHorizontalHeaderLabels(["C\u00f3digo", "Empresa", "CNPJ", "Venc. Cert."])
        self.table.verticalHeader().setVisible(False)
        self.table.horizontalHeader().setStretchLastSection(True)
        self.table.setAlternatingRowColors(True)
        self.table.setSelectionBehavior(QtWidgets.QAbstractItemView.SelectRows)
        self.table.setSelectionMode(QtWidgets.QAbstractItemView.SingleSelection)
        self.table.setEditTriggers(QtWidgets.QAbstractItemView.NoEditTriggers)
        self.table.setShowGrid(False)
        self.table.setColumnWidth(0, 90)
        self.table.setColumnWidth(2, 170)
        self.table.setColumnWidth(3, 130)
        table_l.addWidget(self.table, 1)
        root.addWidget(table_card, 1)

        notes = make_card(self)
        notes_l = QtWidgets.QVBoxLayout(notes)
        set_card_padding(notes_l, 18, 12, 18, 12)
        notes_l.setSpacing(7)
        notes_l.addWidget(title_label("Observações", notes))
        notes_l.addWidget(
            hint_label(
                "Adicionar e Editar atualizam a base principal. Editar NSU grava direto no arquivo "
                "nsu_competencia.json da empresa. Importar XLSX usa o formato Código | Empresa | CNPJ. "
                "Códigos com vírgula ou ponto são importados sem separadores.",
                notes,
            )
        )
        root.addWidget(notes)

        self.summary = QtWidgets.QLabel("")
        self.summary.setObjectName("MetaHint")
        root.addWidget(self.summary)

        self.table.itemSelectionChanged.connect(self._update_action_state)
        self.table.itemDoubleClicked.connect(lambda _item: self._edit_company())
        self.btn_add.clicked.connect(self._add_company)
        self.btn_import.clicked.connect(self._import_companies)
        self.btn_edit.clicked.connect(self._edit_company)
        self.btn_nsu.clicked.connect(self._edit_nsu)
        self.btn_delete.clicked.connect(self._delete_company)
        self.btn_reset_nsu.clicked.connect(self._reset_all_nsu)
        self.btn_delete_all.clicked.connect(self._delete_all)

        self.refresh()

    def refresh(self) -> None:
        rows = load_company_rows()
        summary = summarize_company_rows(rows)
        self._rows = rows
        self.table.setRowCount(len(rows))

        for index, row in enumerate(rows):
            values = [
                str(row["cod"]),
                row["empresa"],
                row["cnpj_formatado"],
                row["venc"] or "-",
            ]
            for column, value in enumerate(values):
                item = QtWidgets.QTableWidgetItem(value)
                if column == 0:
                    item.setData(QtCore.Qt.UserRole, int(row["cod"]))
                self.table.setItem(index, column, item)

            self._apply_row_style(index, row)

        self.summary.setText(
            f"{summary['total']} empresa(s) | {summary['expired']} vencida(s) | {summary['warning']} vencendo em breve"
        )
        if rows and self.table.currentRow() < 0:
            self.table.selectRow(0)
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
            if item is not None:
                item.setForeground(color)
                font = item.font()
                font.setBold(True)
                item.setFont(font)

    def _selected_code(self) -> int | None:
        selected = self.table.selectedItems()
        if not selected:
            return None
        item = self.table.item(selected[0].row(), 0)
        if item is None:
            return None
        return int(item.data(QtCore.Qt.UserRole))

    def _selected_company(self) -> dict | None:
        code = self._selected_code()
        if code is None:
            return None
        return get_company_by_code(code)

    def _update_action_state(self) -> None:
        has_selection = self._selected_code() is not None
        has_rows = bool(self._rows)
        self.btn_edit.setEnabled(has_selection)
        self.btn_nsu.setEnabled(has_selection)
        self.btn_delete.setEnabled(has_selection)
        self.btn_reset_nsu.setEnabled(has_rows)
        self.btn_delete_all.setEnabled(has_rows)

    def _add_company(self) -> None:
        dialog = CompanyDialog(self)
        if dialog.exec() == QtWidgets.QDialog.Accepted:
            self.refresh()
            self.companiesChanged.emit()
            QtWidgets.QMessageBox.information(self, "Sucesso", "Cadastro da empresa criado com sucesso!")

    def _import_companies(self) -> None:
        suggested = suggested_import_spreadsheet_path()
        initial_path = str(suggested) if suggested else ""
        file_path, _ = QtWidgets.QFileDialog.getOpenFileName(
            self,
            "Selecionar planilha de empresas",
            initial_path,
            "Planilhas Excel (*.xlsx)",
        )
        if not file_path:
            return

        try:
            outcome = import_companies_from_xlsx(file_path)
        except Exception as exc:
            QtWidgets.QMessageBox.critical(self, "Erro", f"Não foi possível importar a planilha:\n{exc}")
            return

        self.refresh()
        self.companiesChanged.emit()

        title = "Importação concluída"
        message = self._build_import_summary(outcome)
        if outcome["errors"]:
            QtWidgets.QMessageBox.warning(self, title, message)
        else:
            QtWidgets.QMessageBox.information(self, title, message)

    def _edit_company(self) -> None:
        company = self._selected_company()
        if not company:
            return

        dialog = CompanyDialog(self, company=company)
        if dialog.exec() == QtWidgets.QDialog.Accepted:
            self.refresh()
            self.companiesChanged.emit()
            QtWidgets.QMessageBox.information(self, "Sucesso", "Cadastro atualizado com sucesso!")

    def _edit_nsu(self) -> None:
        company = self._selected_company()
        if not company:
            return
        NSUEditorDialog(company, self).exec()

    def _delete_company(self) -> None:
        company = self._selected_company()
        if not company:
            return

        answer = QtWidgets.QMessageBox.question(
            self,
            "Confirma\u00e7\u00e3o",
            f"Deseja realmente excluir {company['empresa']}?",
        )
        if answer != QtWidgets.QMessageBox.Yes:
            return

        deleted = delete_company(company["cod"])
        if not deleted:
            QtWidgets.QMessageBox.warning(self, "Erro", "Não foi possível excluir a empresa selecionada.")
            return
        self.refresh()
        self.companiesChanged.emit()
        QtWidgets.QMessageBox.information(self, "Sucesso", "Cadastro excluído com sucesso!")

    def _reset_all_nsu(self) -> None:
        answer = QtWidgets.QMessageBox.question(
            self,
            "Confirma\u00e7\u00e3o",
            "Deseja realmente resetar os NSUs de TODOS os cadastros?",
        )
        if answer != QtWidgets.QMessageBox.Yes:
            return

        reset_all_nsu_files()
        QtWidgets.QMessageBox.information(self, "Sucesso", "NSUs de todos os cadastros foram resetados com sucesso!")

    def _delete_all(self) -> None:
        answer = QtWidgets.QMessageBox.question(
            self,
            "Confirma\u00e7\u00e3o",
            "Deseja realmente excluir TODOS os cadastros?",
        )
        if answer != QtWidgets.QMessageBox.Yes:
            return

        delete_all_companies()
        self.refresh()
        self.companiesChanged.emit()
        QtWidgets.QMessageBox.information(self, "Sucesso", "Todos os cadastros foram excluídos com sucesso!")

    def _build_import_summary(self, outcome: dict) -> str:
        created = outcome.get("created", [])
        updated = outcome.get("updated", [])
        unchanged = outcome.get("unchanged", [])
        errors = outcome.get("errors", [])

        lines = [
            f"Planilha: {QtCore.QFileInfo(str(outcome.get('path', ''))).fileName()}",
            f"Aba: {outcome.get('sheet', '-')}",
            "",
            f"Novas empresas: {len(created)}",
            f"Empresas atualizadas: {len(updated)}",
            f"Sem alterações: {len(unchanged)}",
            f"Erros: {len(errors)}",
        ]

        if created:
            lines.append("")
            lines.append("Criadas:")
            lines.extend(
                f"[{item['cod']}] {item['empresa']} | {item['cnpj_formatado']}" for item in created[:8]
            )

        if updated:
            lines.append("")
            lines.append("Atualizadas:")
            lines.extend(
                f"[{item['cod']}] {item['empresa']} | {item['cnpj_formatado']}" for item in updated[:8]
            )

        if errors:
            lines.append("")
            lines.append("Erros encontrados:")
            lines.extend(f"Linha {item['row']}: {item['message']}" for item in errors[:12])
            if len(errors) > 12:
                lines.append(f"... e mais {len(errors) - 12} erro(s).")

        if not created and not updated and not errors:
            lines.append("")
            lines.append("Nenhuma empresa nova foi importada.")

        return "\n".join(lines)
