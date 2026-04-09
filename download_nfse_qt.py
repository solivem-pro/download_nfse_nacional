from __future__ import annotations

import sys

from config.cadastro_db import initialize_company_database
from config.config import configurar_logging


def main() -> int:
    configurar_logging()
    initialize_company_database()

    try:
        from PySide6 import QtWidgets
    except ImportError as exc:
        raise SystemExit(
            "PySide6 n\u00e3o est\u00e1 instalado. Instale as depend\u00eancias de requirements.txt antes de abrir a interface."
        ) from exc

    from ui_qt.main_window import PortalNFSeMainWindow
    from ui_qt.theme import apply_portal_theme

    app = QtWidgets.QApplication.instance() or QtWidgets.QApplication(sys.argv)
    app.setApplicationName("Download NFSe Nacional")
    apply_portal_theme(app)

    window = PortalNFSeMainWindow()
    window.show()

    return app.exec()


if __name__ == "__main__":
    raise SystemExit(main())
