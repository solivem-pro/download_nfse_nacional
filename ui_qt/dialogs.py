from __future__ import annotations

from PySide6 import QtCore, QtWidgets

from ui_qt.data_bridge import app_version, read_markdown_document
from ui_qt.theme import hint_label, make_action_button, make_card, set_card_padding, title_label


class DocsDialog(QtWidgets.QDialog):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Documenta\u00e7\u00e3o - Download NFSe Nacional")
        self.resize(920, 680)

        root = QtWidgets.QVBoxLayout(self)
        root.setContentsMargins(18, 18, 18, 18)
        root.setSpacing(12)

        header = make_card(self, "Hero")
        header_l = QtWidgets.QVBoxLayout(header)
        set_card_padding(header_l, 20, 18, 20, 18)
        header_l.setSpacing(6)
        header_l.addWidget(title_label("Documenta\u00e7\u00e3o", header))
        header_l.addWidget(
            hint_label(
                "Conte\u00fado do arquivo de documenta\u00e7\u00e3o do projeto, exibido diretamente nesta janela.",
                header,
            )
        )
        root.addWidget(header)

        viewer = QtWidgets.QTextBrowser(self)
        viewer.setOpenExternalLinks(True)
        try:
            viewer.setMarkdown(read_markdown_document())
        except Exception:
            viewer.setPlainText(read_markdown_document())
        root.addWidget(viewer, 1)

        footer = QtWidgets.QHBoxLayout()
        footer.addStretch(1)
        btn_close = make_action_button("Voltar", "PrimaryAction", self)
        btn_close.clicked.connect(self.accept)
        footer.addWidget(btn_close)
        root.addLayout(footer)


class AboutDialog(QtWidgets.QDialog):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Sobre - Download NFSe Nacional")
        self.setMinimumWidth(520)

        root = QtWidgets.QVBoxLayout(self)
        root.setContentsMargins(18, 18, 18, 18)
        root.setSpacing(12)

        card = make_card(self, "Hero")
        card_l = QtWidgets.QVBoxLayout(card)
        set_card_padding(card_l, 20, 18, 20, 18)
        card_l.setSpacing(8)

        title = QtWidgets.QLabel("Download NFS-e Portal Nacional")
        title.setObjectName("BrandTitle")
        subtitle = QtWidgets.QLabel(f"Vers\u00e3o {app_version()}")
        subtitle.setObjectName("BrandSub")
        subtitle.setAlignment(QtCore.Qt.AlignLeft | QtCore.Qt.AlignVCenter)
        body = QtWidgets.QLabel(
            "Aplica\u00e7\u00e3o para download de NFS-e no Portal Nacional, com acesso a download, "
            "cadastros, configura\u00e7\u00f5es e documenta\u00e7\u00e3o na mesma interface."
        )
        body.setWordWrap(True)
        body.setObjectName("MetaValue")

        card_l.addWidget(title)
        card_l.addWidget(subtitle)
        card_l.addWidget(body)
        root.addWidget(card)

        footer = QtWidgets.QHBoxLayout()
        footer.addStretch(1)
        btn_close = make_action_button("Voltar", "PrimaryAction", self)
        btn_close.clicked.connect(self.accept)
        footer.addWidget(btn_close)
        root.addLayout(footer)
