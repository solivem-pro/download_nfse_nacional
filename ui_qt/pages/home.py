from __future__ import annotations

from PySide6 import QtCore, QtWidgets

from ui_qt.data_bridge import app_version, config_to_dict, load_app_config, load_company_rows, summarize_company_rows
from ui_qt.theme import (
    field_label,
    hint_label,
    make_action_button,
    make_card,
    pill_label,
    set_card_padding,
)


class HomePage(QtWidgets.QWidget):
    openDownloadRequested = QtCore.Signal()
    openCadastrosRequested = QtCore.Signal()
    openConfigRequested = QtCore.Signal()
    openDocsRequested = QtCore.Signal()
    aboutRequested = QtCore.Signal()

    def __init__(self, parent=None):
        super().__init__(parent)
        self.setStyleSheet(
            """
            QLabel {
                background: transparent;
            }
            QLabel#BrandTitle {
                color: #edf1f7;
            }
            QLabel#BrandSub {
                color: #8c94a0;
            }
            """
        )

        root = QtWidgets.QVBoxLayout(self)
        root.setContentsMargins(0, 0, 0, 0)
        root.setSpacing(10)

        hero = make_card(self, "Hero")
        hero_l = QtWidgets.QVBoxLayout(hero)
        set_card_padding(hero_l, 20, 16, 20, 16)
        hero_l.setSpacing(7)

        top_row = QtWidgets.QHBoxLayout()
        top_row.setSpacing(8)

        badge = pill_label("PORTAL NFSe", hero)
        badge.setFixedWidth(112)
        top_row.addWidget(badge)
        top_row.addStretch(1)

        top_status = pill_label(f"v{app_version()}", hero)
        top_status.setFixedWidth(76)
        top_row.addWidget(top_status)

        hero_l.addLayout(top_row)

        title = QtWidgets.QLabel("Download NFS-e Portal Nacional")
        title.setObjectName("BrandTitle")
        subtitle = QtWidgets.QLabel("Selecione uma op\u00e7\u00e3o abaixo para continuar.")
        subtitle.setObjectName("BrandSub")
        subtitle.setWordWrap(True)

        hero_l.addWidget(title)
        hero_l.addWidget(subtitle)

        action_row = QtWidgets.QHBoxLayout()
        action_row.setSpacing(7)

        btn_download = make_action_button("Baixar NFSe", "PrimaryAction", hero)
        btn_cadastros = make_action_button("Cadastros", "AccentAction", hero)
        btn_config = make_action_button("Configura\u00e7\u00f5es", "GhostAction", hero)
        btn_docs = make_action_button("Documenta\u00e7\u00e3o", "GhostAction", hero)
        btn_about = make_action_button("Sobre", "GhostAction", hero)

        btn_download.clicked.connect(self.openDownloadRequested.emit)
        btn_cadastros.clicked.connect(self.openCadastrosRequested.emit)
        btn_config.clicked.connect(self.openConfigRequested.emit)
        btn_docs.clicked.connect(self.openDocsRequested.emit)
        btn_about.clicked.connect(self.aboutRequested.emit)

        action_row.addWidget(btn_download)
        action_row.addWidget(btn_cadastros)
        action_row.addWidget(btn_config)
        action_row.addWidget(btn_docs)
        action_row.addWidget(btn_about)
        hero_l.addLayout(action_row)

        stats = QtWidgets.QGridLayout()
        stats.setContentsMargins(0, 0, 0, 0)
        stats.setHorizontalSpacing(9)
        stats.setVerticalSpacing(9)
        stats.setColumnStretch(0, 1)
        stats.setColumnStretch(1, 1)

        self._stat_value_labels: dict[str, QtWidgets.QLabel] = {}
        cards = [
            ("Empresas", "total", "Empresas cadastradas"),
            ("V\u00e1lidas", "valid", "Certificados sem vencimento"),
            ("Vencidas", "expired", "Certificados expirados"),
            ("Baixar PDF", "download_pdf", "Configura\u00e7\u00e3o atual do download"),
        ]

        for index, (name, key, helper_text) in enumerate(cards):
            frame = make_card(self)
            frame_l = QtWidgets.QVBoxLayout(frame)
            set_card_padding(frame_l, 16, 12, 16, 12)
            frame_l.setSpacing(5)
            frame.setMinimumHeight(96)

            lbl_name = field_label(name, frame)
            lbl_val = QtWidgets.QLabel("--")
            lbl_val.setObjectName("MetaValue")
            lbl_val.setWordWrap(True)
            lbl_hint = hint_label(helper_text, frame)
            frame_l.addWidget(lbl_name)
            frame_l.addWidget(lbl_val)
            frame_l.addWidget(lbl_hint)
            self._stat_value_labels[key] = lbl_val

            row = index // 2
            col = index % 2
            stats.addWidget(frame, row, col)

        root.addWidget(hero)
        root.addLayout(stats)
        root.addStretch(1)

        self.refresh()

    def refresh(self) -> None:
        rows = load_company_rows()
        summary = summarize_company_rows(rows)
        config = config_to_dict(load_app_config())

        self._stat_value_labels["total"].setText(str(summary["total"]))
        self._stat_value_labels["valid"].setText(str(summary["valid"]))
        self._stat_value_labels["expired"].setText(str(summary["expired"]))
        self._stat_value_labels["download_pdf"].setText(
            "Ativado" if bool(config.get("download_pdf", False)) else "Desativado"
        )
