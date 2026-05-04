from __future__ import annotations

import sys
from pathlib import Path

from PySide6 import QtCore, QtGui, QtWidgets

from config.config import DIRETORIOS
from ui_qt.data_bridge import app_version
from ui_qt.dialogs import AboutDialog, DocsDialog
from ui_qt.pages import CadastrosPage, ConfigPage, DownloadPage, HomePage
from ui_qt.theme import apply_app_theme, make_card, make_nav_button


def _load_icon() -> QtGui.QIcon:
    try:
        icon_path = Path(DIRETORIOS["icone"])
        if icon_path.exists():
            return QtGui.QIcon(str(icon_path))
    except Exception:
        pass
    return QtGui.QIcon()


class MainWindow(QtWidgets.QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle(f"Download NFSe Nacional v{app_version()}")
        self.setMinimumSize(1180, 780)
        self.setWindowIcon(_load_icon())
        self.showMaximized()

        central = QtWidgets.QWidget()
        central.setObjectName("AppRoot")
        self.setCentralWidget(central)

        root = QtWidgets.QVBoxLayout(central)
        root.setContentsMargins(20, 16, 20, 16)
        root.setSpacing(12)

        header = make_card(central, "Hero")
        header_l = QtWidgets.QHBoxLayout(header)
        header_l.setContentsMargins(18, 15, 18, 15)
        header_l.setSpacing(10)

        badge = QtWidgets.QLabel("NFSe")
        badge.setObjectName("Badge")
        badge.setAlignment(QtCore.Qt.AlignCenter)
        badge.setFixedWidth(64)

        title_wrap = QtWidgets.QVBoxLayout()
        title_wrap.setContentsMargins(0, 0, 0, 0)
        title_wrap.setSpacing(2)

        lbl_title = QtWidgets.QLabel("Download NFS-e Portal Nacional")
        lbl_title.setObjectName("BrandTitle")
        lbl_sub = QtWidgets.QLabel("Selecione uma op\u00e7\u00e3o para continuar.")
        lbl_sub.setObjectName("BrandSub")

        title_wrap.addWidget(lbl_title)
        title_wrap.addWidget(lbl_sub)

        header_l.addWidget(badge)
        header_l.addLayout(title_wrap)
        header_l.addStretch(1)

        self.status = QtWidgets.QLabel(f"v{app_version()}")
        self.status.setObjectName("MetaValue")
        self.status.setAlignment(QtCore.Qt.AlignCenter)
        header_l.addWidget(self.status)
        root.addWidget(header)

        nav_wrap = make_card(central)
        nav_l = QtWidgets.QHBoxLayout(nav_wrap)
        nav_l.setContentsMargins(12, 9, 12, 9)
        nav_l.setSpacing(8)

        self.btn_home = make_nav_button("In\u00edcio", nav_wrap)
        self.btn_download = make_nav_button("Download", nav_wrap)
        self.btn_cadastros = make_nav_button("Cadastros", nav_wrap)
        self.btn_config = make_nav_button("Configura\u00e7\u00f5es", nav_wrap)

        self.btn_home.setChecked(True)
        for btn in (self.btn_home, self.btn_download, self.btn_cadastros, self.btn_config):
            nav_l.addWidget(btn)
        nav_l.addStretch(1)
        root.addWidget(nav_wrap)

        self.stack = QtWidgets.QStackedWidget()
        root.addWidget(self.stack, 1)

        self.page_home = HomePage()
        self.page_download = DownloadPage()
        self.page_cadastros = CadastrosPage()
        self.page_config = ConfigPage()

        self.stack.addWidget(self.page_home)
        self.stack.addWidget(self.page_download)
        self.stack.addWidget(self.page_cadastros)
        self.stack.addWidget(self.page_config)

        self.btn_home.clicked.connect(lambda: self.set_page(0))
        self.btn_download.clicked.connect(lambda: self.set_page(1))
        self.btn_cadastros.clicked.connect(lambda: self.set_page(2))
        self.btn_config.clicked.connect(lambda: self.set_page(3))

        self.page_home.openDownloadRequested.connect(lambda: self.set_page(1))
        self.page_home.openCadastrosRequested.connect(lambda: self.set_page(2))
        self.page_home.openConfigRequested.connect(lambda: self.set_page(3))
        self.page_home.openDocsRequested.connect(self._show_docs)
        self.page_home.aboutRequested.connect(self._show_about)
        self.page_config.configSaved.connect(self.refresh_shell)
        self.page_cadastros.companiesChanged.connect(self.refresh_shell)

        self.set_page(0)
        self.refresh_shell()

    def set_page(self, index: int) -> None:
        self.stack.setCurrentIndex(index)
        self.btn_home.setChecked(index == 0)
        self.btn_download.setChecked(index == 1)
        self.btn_cadastros.setChecked(index == 2)
        self.btn_config.setChecked(index == 3)
        self.refresh_shell()

    def refresh_shell(self) -> None:
        self.page_home.refresh()
        self.page_download.refresh()
        self.page_cadastros.refresh()
        self.page_config.refresh()

    def _show_docs(self) -> None:
        DocsDialog(self).exec()

    def _show_about(self) -> None:
        AboutDialog(self).exec()


PortalNFSeMainWindow = MainWindow


def run() -> int:
    app = QtWidgets.QApplication.instance() or QtWidgets.QApplication(sys.argv)
    apply_app_theme(app)
    win = MainWindow()
    win.show()
    app._portal_nfse_main_window = win  # type: ignore[attr-defined]
    return app.exec()


if __name__ == "__main__":
    raise SystemExit(run())
