from __future__ import annotations

from PySide6 import QtCore, QtWidgets

from ui_qt.data_bridge import (
    CONSULT_MODE_COMPETENCIA,
    CONSULT_MODE_EMISSAO,
    SAVE_MODE_CNPJ,
    SAVE_MODE_CODIGO,
    config_to_dict,
    load_app_config,
    save_app_config,
)
from ui_qt.theme import (
    field_label,
    hint_label,
    make_action_button,
    make_card,
    set_card_padding,
    title_label,
)


class ConfigPage(QtWidgets.QWidget):
    backRequested = QtCore.Signal()
    configSaved = QtCore.Signal()

    def __init__(self, parent=None):
        super().__init__(parent)
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
        root.setSpacing(10)

        header = make_card(self, "Hero")
        header_l = QtWidgets.QVBoxLayout(header)
        set_card_padding(header_l, 20, 16, 20, 16)
        header_l.setSpacing(5)
        header_l.addWidget(title_label("Configura\u00e7\u00f5es", header))
        header_l.addWidget(hint_label("Ajuste as configura\u00e7\u00f5es do programa.", header))
        root.addWidget(header)

        content = QtWidgets.QHBoxLayout()
        content.setContentsMargins(0, 0, 0, 0)
        content.setSpacing(10)

        form_card = make_card(self)
        form_card_l = QtWidgets.QVBoxLayout(form_card)
        set_card_padding(form_card_l, 18, 12, 18, 12)
        form_card_l.setSpacing(9)
        form_card_l.addWidget(title_label("Configura\u00e7\u00f5es", form_card))

        form = QtWidgets.QFormLayout()
        form.setLabelAlignment(QtCore.Qt.AlignLeft)
        form.setFormAlignment(QtCore.Qt.AlignTop)
        form.setHorizontalSpacing(12)
        form.setVerticalSpacing(9)
        form.setFieldGrowthPolicy(QtWidgets.QFormLayout.AllNonFixedFieldsGrow)
        form.setContentsMargins(0, 0, 0, 0)

        self.input_prefix = QtWidgets.QLineEdit()
        self.input_prefix.setPlaceholderText("NFSE")
        self.input_prefix.setClearButtonEnabled(True)
        self.input_delay = QtWidgets.QDoubleSpinBox()
        self.input_delay.setDecimals(1)
        self.input_delay.setSingleStep(0.5)
        self.input_delay.setRange(0.0, 999.0)
        self.input_delay.setSuffix(" s")

        self.input_timeout = QtWidgets.QSpinBox()
        self.input_timeout.setRange(1, 9999)
        self.input_timeout.setSuffix(" s")
        self.combo_mode = QtWidgets.QComboBox()
        self.combo_mode.addItems([CONSULT_MODE_COMPETENCIA, CONSULT_MODE_EMISSAO])
        self.combo_save = QtWidgets.QComboBox()
        self.combo_save.addItems([SAVE_MODE_CODIGO, SAVE_MODE_CNPJ])
        self.check_pdf = QtWidgets.QCheckBox("Baixar PDF")

        widgets = [
            ("Prefixo Arquivo", self.input_prefix),
            ("Delay (s)", self.input_delay),
            ("Timeout (s)", self.input_timeout),
            ("Modo Consulta", self.combo_mode),
            ("Modo Cadastros", self.combo_save),
        ]

        for label_text, widget in widgets:
            form.addRow(field_label(label_text, form_card), widget)

        form.addRow("", self.check_pdf)

        btn_row = QtWidgets.QHBoxLayout()
        btn_row.setSpacing(10)
        self.btn_save = make_action_button("Salvar", "PrimaryAction", form_card)
        self.btn_reload = make_action_button("Cancelar", "GhostAction", form_card)
        btn_row.addWidget(self.btn_save)
        btn_row.addWidget(self.btn_reload)
        btn_row.addStretch(1)

        form_card_l.addLayout(form)
        form_card_l.addLayout(btn_row)

        side_card = make_card(self)
        side_card_l = QtWidgets.QVBoxLayout(side_card)
        set_card_padding(side_card_l, 18, 12, 18, 12)
        side_card_l.setSpacing(9)
        side_card_l.addWidget(title_label("Ajuda", side_card))
        side_card_l.addWidget(
            hint_label(
                "Prefixo Arquivo define o in\u00edcio do nome do arquivo baixado.\n\n"
                "Modo Consulta permite consultar por Compet\u00eancia ou Emiss\u00e3o.\n"
                "Modo Cadastros define se o salvamento ser\u00e1 por C\u00f3digo ou CNPJ.\n"
                "Baixar PDF aumenta o tempo total de processamento.",
                side_card,
            )
        )
        self.summary = QtWidgets.QLabel("")
        self.summary.setObjectName("MetaValue")
        self.summary.setWordWrap(True)
        side_card_l.addWidget(self.summary)
        side_card_l.addStretch(1)

        content.addWidget(form_card, 3)
        content.addWidget(side_card, 2)
        root.addLayout(content)

        root.addStretch(1)

        self.btn_save.clicked.connect(self.save)
        self.btn_reload.clicked.connect(self.refresh)
        self.refresh()

    def refresh(self) -> None:
        config = config_to_dict(load_app_config())
        self.input_prefix.setText(str(config.get("file_prefix", "NFSE")))
        self.input_delay.setValue(float(config.get("delay_seconds", 0.5)))
        self.input_timeout.setValue(int(float(config.get("timeout", 60))))
        self.combo_mode.setCurrentText(str(config.get("consult_mode", CONSULT_MODE_COMPETENCIA)))
        self.combo_save.setCurrentText(str(config.get("save_mode", SAVE_MODE_CODIGO)))
        self.check_pdf.setChecked(bool(config.get("download_pdf", False)))
        self.summary.setText(
            f"Modo Consulta: {config.get('consult_mode', CONSULT_MODE_COMPETENCIA)}\n"
            f"Modo Cadastros: {config.get('save_mode', SAVE_MODE_CODIGO)}\n"
            f"Baixar PDF: {'Sim' if bool(config.get('download_pdf', False)) else 'N\u00e3o'}"
        )

    def save(self) -> None:
        try:
            config = save_app_config(
                {
                    "file_prefix": self.input_prefix.text().strip() or "NFSE",
                    "download_pdf": self.check_pdf.isChecked(),
                    "delay_seconds": float(self.input_delay.value()),
                    "timeout": int(self.input_timeout.value()),
                    "consult_mode": self.combo_mode.currentText(),
                    "save_mode": self.combo_save.currentText(),
                }
            )
        except ValueError:
            QtWidgets.QMessageBox.warning(
                self,
                "Erro",
                "Valores num\u00e9ricos inv\u00e1lidos!",
            )
            return

        QtWidgets.QMessageBox.information(
            self,
            "Configura\u00e7\u00f5es",
            "Configura\u00e7\u00f5es salvas com sucesso!",
        )
        self.configSaved.emit()
