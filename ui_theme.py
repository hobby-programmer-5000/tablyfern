"""Workspace styling and locally bundled Lucide icons."""

import sys
from pathlib import Path

from PySide6.QtCore import QRectF, QSize, Qt
from PySide6.QtGui import QIcon, QPainter, QPalette, QPixmap
from PySide6.QtSvg import QSvgRenderer
from PySide6.QtWidgets import QApplication

APP_DIR = Path(sys.executable if getattr(sys, "frozen", False) else __file__).resolve().parent
ICON_DIR = APP_DIR / "assets" / "lucide"

LIGHT_STYLE = """
QWidget#workspace { background: #f6f7f9; color: #202124; }
QLabel { color: #202124; }
QLabel[muted="true"], QLabel#statusLine, QLabel#outputLabel { color: #5f6368; }
QFrame[panel="true"] { background: #ffffff; border: 1px solid #dadce0; border-radius: 8px; }
QPushButton[workspaceButton="true"], QComboBox[workspaceCombo="true"] {
    background: #ffffff; color: #303134; border: 1px solid #b7bbc2;
    border-radius: 4px; padding: 4px 12px;
}
QPushButton[workspaceButton="true"]:hover, QComboBox[workspaceCombo="true"]:hover {
    background: #f1f3f4; border-color: #80868b;
}
QPushButton[workspaceButton="true"]:pressed { background: #e8eaed; }
QPushButton[quiet="true"] { background: transparent; border-color: transparent; color: #5f6368; }
QPushButton[quiet="true"]:hover { background: #e8eaed; border-color: transparent; color: #202124; }
QPushButton#exportButton { background: #185abc; color: #ffffff; border-color: #185abc; }
QPushButton#exportButton:hover { background: #174ea6; border-color: #174ea6; }
QPushButton#exportButton:pressed { background: #123f87; border-color: #123f87; }
QPushButton[workspaceButton="true"]:disabled, QPushButton#exportButton:disabled,
QComboBox[workspaceCombo="true"]:disabled {
    background: #eceef1; color: #686c73; border-color: #dadce0;
}
QComboBox[workspaceCombo="true"] { padding-right: 25px; }
QComboBox[workspaceCombo="true"]::drop-down { border: none; width: 24px; }
QComboBox[workspaceCombo="true"]::down-arrow {
    image: url("${arrow}"); width: 14px; height: 14px;
}
QSplitter::handle { background: transparent; }
QSplitter::handle:hover { background: #dadce0; border-radius: 4px; }
QPlainTextEdit#sourceEditor, QTableWidget#previewGrid {
    background: #ffffff; border: 2px solid transparent; border-radius: 2px;
}
QStackedWidget#previewStack { background: #ffffff; border: none; }
QPlainTextEdit#sourceEditor {
    color: #303134; placeholder-text-color: #5f6368; padding: 4px 8px;
    selection-background-color: #d3e3fd; selection-color: #202124;
}
QTableWidget#previewGrid {
    alternate-background-color: #fafbfc; color: #202124;
    selection-background-color: #d3e3fd; selection-color: #202124;
}
QTableWidget#previewGrid::item { border-bottom: 1px solid #eef0f2; padding: 0 8px; }
QTableWidget#previewGrid QHeaderView::section {
    background: #f1f3f4; color: #5f6368; padding: 8px;
    border: none; border-bottom: 1px solid #dadce0;
}
QTableWidget#previewGrid QTableCornerButton::section {
    background: #f1f3f4; border: none; border-bottom: 1px solid #dadce0;
}
QPushButton[workspaceButton="true"]:focus, QPushButton#exportButton:focus,
QComboBox[workspaceCombo="true"]:focus, QPlainTextEdit#sourceEditor:focus,
QTableWidget#previewGrid:focus { border: 2px solid #185abc; }
QPushButton#exportButton:focus { border-color: #b8d4ff; }
QLabel#statusLine[error="true"] { color: #b3261e; }
QProgressBar { background: #e8eaed; border: none; border-radius: 2px; }
QProgressBar::chunk { background: #185abc; border-radius: 2px; }
"""


def use_light_style() -> bool:
    hints = QApplication.styleHints()
    return (
        hints.colorScheme() != Qt.ColorScheme.Dark
        and hints.accessibility().contrastPreference() != Qt.ContrastPreference.HighContrast
    )


def stock_icon(name: str, *, light: bool, inverse: bool = False) -> QIcon:
    artwork = (ICON_DIR / f"{name}.svg").read_bytes()
    palette = QApplication.palette()
    normal = "#ffffff" if inverse else "#5f6368"
    disabled = "#686c73"
    if not light:
        normal = palette.color(QPalette.ColorRole.ButtonText).name()
        disabled = palette.color(QPalette.ColorGroup.Disabled, QPalette.ColorRole.ButtonText).name()

    icon = QIcon()
    for color, mode in ((normal, QIcon.Mode.Normal), (disabled, QIcon.Mode.Disabled)):
        renderer = QSvgRenderer(artwork.replace(b"currentColor", color.encode("ascii")))
        for scale in (1, 2):
            pixmap = QPixmap(QSize(24, 24) * scale)
            pixmap.setDevicePixelRatio(scale)
            pixmap.fill(Qt.GlobalColor.transparent)
            painter = QPainter(pixmap)
            renderer.render(painter, QRectF(0, 0, 24, 24))
            painter.end()
            icon.addPixmap(pixmap, mode)
    return icon
