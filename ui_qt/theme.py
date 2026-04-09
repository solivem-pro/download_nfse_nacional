from __future__ import annotations

from PySide6 import QtCore, QtGui, QtWidgets

COL_BG = "#14171d"
COL_PANEL = "#181c25"
COL_CARD = "#1a1f29"
COL_BORDER = "#2d333c"
COL_INPUT = "#242a33"
COL_FG = "#e1e5ec"
COL_MUTED = "#818898"
COL_PRIMARY = "#33a2ff"
COL_PRIMARY_HOVER = "#1e8df0"
COL_ACCENT = "#22b2a1"
COL_ACCENT_HOVER = "#1aa392"
COL_DANGER = "#ea5b5b"
COL_WARN = "#f0a21e"
COL_SUCCESS = "#29ad58"


def apply_app_theme(app: QtWidgets.QApplication) -> None:
    app.setStyle("Fusion")

    palette = QtGui.QPalette()
    palette.setColor(QtGui.QPalette.Window, QtGui.QColor(COL_BG))
    palette.setColor(QtGui.QPalette.WindowText, QtGui.QColor(COL_FG))
    palette.setColor(QtGui.QPalette.Base, QtGui.QColor(COL_CARD))
    palette.setColor(QtGui.QPalette.AlternateBase, QtGui.QColor(COL_PANEL))
    palette.setColor(QtGui.QPalette.Text, QtGui.QColor(COL_FG))
    palette.setColor(QtGui.QPalette.Button, QtGui.QColor(COL_CARD))
    palette.setColor(QtGui.QPalette.ButtonText, QtGui.QColor(COL_FG))
    palette.setColor(QtGui.QPalette.Highlight, QtGui.QColor(COL_PRIMARY))
    palette.setColor(QtGui.QPalette.HighlightedText, QtGui.QColor("#081018"))
    palette.setColor(QtGui.QPalette.ToolTipBase, QtGui.QColor(COL_CARD))
    palette.setColor(QtGui.QPalette.ToolTipText, QtGui.QColor(COL_FG))
    app.setPalette(palette)

    app.setStyleSheet(
        f"""
        QMainWindow {{
            background: {COL_BG};
            color: {COL_FG};
            font-family: "Segoe UI";
            font-size: 10pt;
        }}

        QWidget {{
            background: transparent;
            color: {COL_FG};
            font-family: "Segoe UI";
            font-size: 10pt;
        }}

        QWidget#AppRoot {{
            background: qradialgradient(cx:0.18, cy:0.10, radius:1.1,
                fx:0.18, fy:0.10,
                stop:0 #1a202a, stop:0.38 {COL_BG}, stop:1 #0f1318);
        }}

        QFrame#Card, QFrame#Surface {{
            background: qlineargradient(x1:0, y1:0, x2:0, y2:1,
                stop:0 #1b2130, stop:1 {COL_CARD});
            border: 1px solid {COL_BORDER};
            border-radius: 14px;
        }}

        QFrame#Hero {{
            background: qlineargradient(x1:0, y1:0, x2:1, y2:1,
                stop:0 #1d2735, stop:0.45 #1a2230, stop:1 {COL_CARD});
            border: 1px solid {COL_BORDER};
            border-radius: 18px;
        }}

        QLabel#BrandTitle {{
            color: {COL_FG};
            background: transparent;
            font-size: 17px;
            font-weight: 700;
        }}

        QLabel#BrandSub, QLabel#MetaHint {{
            color: {COL_MUTED};
            background: transparent;
            font-size: 11px;
        }}

        QLabel#SectionTitle {{
            color: #f0f3f8;
            background: transparent;
            font-size: 14px;
            font-weight: 700;
        }}

        QLabel#FieldLabel {{
            color: #99a3b3;
            background: transparent;
            font-size: 10px;
            font-weight: 700;
            letter-spacing: 0.8px;
        }}

        QLabel#MetaValue {{
            color: {COL_FG};
            background: {COL_INPUT};
            border: 1px solid {COL_BORDER};
            border-radius: 8px;
            padding: 7px 10px;
            font-weight: 600;
        }}

        QLabel#Badge {{
            background: rgba(51,162,255,0.15);
            color: {COL_PRIMARY};
            border-radius: 10px;
            padding: 6px 10px;
            font-weight: 700;
        }}

        QStackedWidget, QStackedWidget > QWidget, QAbstractScrollArea {{
            background: transparent;
        }}

        QLineEdit, QTextEdit, QPlainTextEdit, QComboBox, QAbstractSpinBox {{
            background: {COL_INPUT};
            border: 1px solid {COL_BORDER};
            border-radius: 10px;
            padding: 9px 11px;
            color: {COL_FG};
            selection-background-color: rgba(51,162,255,0.26);
            selection-color: #f6f9fd;
        }}
        QTextBrowser {{
            background: {COL_CARD};
            border: 1px solid {COL_BORDER};
            border-radius: 12px;
            padding: 12px;
            color: {COL_FG};
            selection-background-color: rgba(51,162,255,0.26);
        }}
        QLineEdit:focus, QTextEdit:focus, QPlainTextEdit:focus, QComboBox:focus, QAbstractSpinBox:focus {{
            background: #29313c;
            border: 1px solid #67b8ff;
            outline: 0;
        }}

        QComboBox::drop-down {{
            subcontrol-origin: padding;
            subcontrol-position: top right;
            width: 22px;
            border: none;
        }}
        QComboBox QAbstractItemView {{
            background: {COL_CARD};
            color: {COL_FG};
            border: 1px solid {COL_BORDER};
            selection-background-color: rgba(51,162,255,0.24);
            selection-color: #f6f9fd;
        }}

        QMenu {{
            background: {COL_CARD};
            color: {COL_FG};
            border: 1px solid {COL_BORDER};
            padding: 6px;
        }}
        QMenu::item {{
            padding: 8px 12px;
            border-radius: 8px;
        }}
        QMenu::item:selected {{
            background: rgba(51,162,255,0.20);
        }}

        QCheckBox {{
            color: {COL_FG};
            spacing: 8px;
        }}
        QCheckBox::indicator {{
            width: 16px;
            height: 16px;
            border-radius: 4px;
            border: 1px solid {COL_BORDER};
            background: {COL_INPUT};
        }}
        QCheckBox::indicator:hover {{
            border-color: #4b5769;
        }}
        QCheckBox::indicator:checked {{
            background: {COL_PRIMARY};
            border-color: {COL_PRIMARY};
        }}
        QCheckBox:focus {{
            outline: 0;
        }}

        QPushButton {{
            background: {COL_INPUT};
            border: 1px solid {COL_BORDER};
            color: {COL_FG};
            padding: 8px 13px;
            border-radius: 10px;
            font-weight: 500;
        }}
        QPushButton:hover {{
            background: #282f3a;
            border: 1px solid #404958;
        }}
        QPushButton:pressed {{
            background: #1d232c;
            border-color: #364152;
        }}
        QPushButton:disabled {{
            color: #616775;
        }}
        QPushButton:focus {{
            border: 1px solid #67b8ff;
            outline: 0;
        }}

        QPushButton#PrimaryAction {{
            background: {COL_PRIMARY};
            border: 1px solid {COL_PRIMARY};
            color: #081018;
            font-weight: 700;
        }}
        QPushButton#PrimaryAction:hover {{
            background: {COL_PRIMARY_HOVER};
            border-color: {COL_PRIMARY_HOVER};
        }}
        QPushButton#PrimaryAction:pressed {{
            background: #1578cc;
            border-color: #1578cc;
        }}
        QPushButton#PrimaryAction:focus {{
            background: #59b5ff;
            border-color: #8ed0ff;
        }}

        QPushButton#PrimaryOutlineAction {{
            background: transparent;
            border: 1px solid {COL_PRIMARY};
            color: {COL_PRIMARY};
            font-weight: 700;
        }}
        QPushButton#PrimaryOutlineAction:hover {{
            background: rgba(51,162,255,0.12);
        }}
        QPushButton#PrimaryOutlineAction:focus {{
            background: rgba(51,162,255,0.14);
            border-color: #8ed0ff;
        }}
        QPushButton#PrimaryOutlineAction:pressed {{
            background: rgba(51,162,255,0.18);
            border-color: {COL_PRIMARY_HOVER};
        }}

        QPushButton#AccentAction {{
            background: transparent;
            border: 1px solid {COL_ACCENT};
            color: {COL_ACCENT};
            font-weight: 700;
        }}
        QPushButton#AccentAction:hover {{
            background: rgba(34,178,161,0.12);
        }}
        QPushButton#AccentAction:focus {{
            background: rgba(34,178,161,0.12);
            border-color: #39c1b1;
        }}
        QPushButton#AccentAction:pressed {{
            background: rgba(34,178,161,0.18);
            border-color: #2ab8a8;
        }}

        QPushButton#Danger {{
            background: transparent;
            border: 1px solid {COL_DANGER};
            color: {COL_DANGER};
            font-weight: 700;
        }}
        QPushButton#Danger:hover {{
            background: rgba(234,91,91,0.10);
        }}
        QPushButton#Danger:focus {{
            background: rgba(234,91,91,0.08);
            border-color: #f07b7b;
        }}
        QPushButton#Danger:pressed {{
            background: rgba(234,91,91,0.16);
            border-color: #ff8c8c;
        }}

        QPushButton#GhostAction {{
            background: transparent;
            border: 1px solid {COL_BORDER};
            color: {COL_MUTED};
        }}
        QPushButton#GhostAction:hover {{
            background: rgba(129,136,152,0.10);
            color: {COL_FG};
        }}
        QPushButton#GhostAction:focus {{
            background: rgba(129,136,152,0.10);
            border-color: #67b8ff;
            color: {COL_FG};
        }}

        QToolButton#NavButton {{
            background: transparent;
            border: 1px solid {COL_BORDER};
            border-radius: 10px;
            color: {COL_FG};
            padding: 9px 14px;
            font-weight: 600;
            min-width: 86px;
        }}
        QToolButton#NavButton:checked {{
            background: rgba(51,162,255,0.12);
            border-color: {COL_PRIMARY};
            color: {COL_FG};
        }}
        QToolButton#NavButton:hover {{
            background: rgba(255,255,255,0.03);
            border-color: #3a4250;
        }}
        QToolButton#NavButton:focus {{
            background: rgba(51,162,255,0.12);
            border-color: #67b8ff;
            outline: 0;
        }}
        QToolButton#NavButton:checked:focus {{
            background: rgba(51,162,255,0.16);
            border-color: #79c2ff;
        }}

        QTableWidget {{
            background: {COL_CARD};
            border: 1px solid {COL_BORDER};
            border-radius: 12px;
            gridline-color: {COL_BORDER};
            selection-background-color: rgba(51,162,255,0.24);
            alternate-background-color: #171c24;
        }}
        QTableWidget:focus {{
            border: 1px solid #67b8ff;
            outline: 0;
        }}
        QTableCornerButton::section {{
            background: {COL_CARD};
            border: none;
            border-bottom: 1px solid {COL_BORDER};
        }}
        QHeaderView::section {{
            background: {COL_CARD};
            color: {COL_MUTED};
            padding: 10px 8px;
            border: none;
            border-bottom: 1px solid {COL_BORDER};
            font-weight: 700;
        }}
        QTableWidget::item {{
            padding: 10px 8px;
        }}
        QTableWidget::item:selected {{
            color: {COL_FG};
            background: rgba(51,162,255,0.28);
        }}

        QScrollBar:vertical {{
            background: transparent;
            width: 12px;
            margin: 8px 2px 8px 2px;
        }}
        QScrollBar::handle:vertical {{
            background: #2d333c;
            min-height: 28px;
            border-radius: 6px;
        }}
        QScrollBar::handle:vertical:hover {{
            background: #3a4250;
        }}
        QScrollBar::add-line:vertical, QScrollBar::sub-line:vertical {{
            height: 0px;
        }}
        QScrollBar:horizontal {{
            background: transparent;
            height: 12px;
            margin: 2px 8px 2px 8px;
        }}
        QScrollBar::handle:horizontal {{
            background: #2d333c;
            min-width: 28px;
            border-radius: 6px;
        }}
        QScrollBar::handle:horizontal:hover {{
            background: #3a4250;
        }}
        QScrollBar::add-line:horizontal, QScrollBar::sub-line:horizontal {{
            width: 0px;
        }}

        QProgressBar {{
            background: {COL_INPUT};
            border: 1px solid {COL_BORDER};
            border-radius: 8px;
            text-align: center;
            color: {COL_FG};
            height: 18px;
        }}
        QProgressBar::chunk {{
            border-radius: 8px;
            background: {COL_PRIMARY};
        }}
        """
    )


def make_card(parent: QtWidgets.QWidget, object_name: str = "Card") -> QtWidgets.QFrame:
    frame = QtWidgets.QFrame(parent)
    frame.setObjectName(object_name)
    return frame


def set_card_padding(
    layout: QtWidgets.QLayout,
    left: int = 18,
    top: int = 14,
    right: int = 18,
    bottom: int = 14,
) -> None:
    layout.setContentsMargins(left, top, right, bottom)


def title_label(text: str, parent: QtWidgets.QWidget | None = None) -> QtWidgets.QLabel:
    label = QtWidgets.QLabel(text, parent)
    label.setObjectName("SectionTitle")
    return label


def hint_label(text: str, parent: QtWidgets.QWidget | None = None) -> QtWidgets.QLabel:
    label = QtWidgets.QLabel(text, parent)
    label.setObjectName("MetaHint")
    return label


def field_label(text: str, parent: QtWidgets.QWidget | None = None) -> QtWidgets.QLabel:
    label = QtWidgets.QLabel(text, parent)
    label.setObjectName("FieldLabel")
    return label


def pill_label(text: str, parent: QtWidgets.QWidget | None = None) -> QtWidgets.QLabel:
    label = QtWidgets.QLabel(text, parent)
    label.setObjectName("Badge")
    label.setAlignment(QtCore.Qt.AlignCenter)
    return label


def make_action_button(
    text: str,
    role: str = "GhostAction",
    parent: QtWidgets.QWidget | None = None,
) -> QtWidgets.QPushButton:
    button = QtWidgets.QPushButton(text, parent)
    button.setObjectName(role)
    button.setCursor(QtCore.Qt.PointingHandCursor)
    button.setMinimumHeight(38)
    return button


def make_nav_button(text: str, parent: QtWidgets.QWidget | None = None) -> QtWidgets.QToolButton:
    button = QtWidgets.QToolButton(parent)
    button.setText(text)
    button.setCheckable(True)
    button.setAutoExclusive(True)
    button.setObjectName("NavButton")
    button.setToolButtonStyle(QtCore.Qt.ToolButtonTextOnly)
    button.setCursor(QtCore.Qt.PointingHandCursor)
    return button


def apply_portal_theme(app: QtWidgets.QApplication) -> None:
    apply_app_theme(app)
