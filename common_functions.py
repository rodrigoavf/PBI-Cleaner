import os
import hashlib
import re
from PyQt6.QtCore import Qt
from PyQt6.QtGui import QFont, QPalette, QColor
from PyQt6.QtWidgets import QStyleFactory, QApplication

def count_files(root_dir):
    total = 0
    for _, _, files in os.walk(root_dir):
        total += len(files)
    return total
def code_editor_font(f_type="Consolas", f_size=10):
    return QFont(f_type, f_size)

# --- Application theme configuration ---
# Four curated themes to keep the UI cohesive.
THEME_PRESETS = {
    "tentacles_dark": "Tentacles Dark",
    "tentacles_light": "Tentacles Light",
    "tentacles_purple": "Tentacles Purple",
    "tentacles_green": "Tentacles Green",
}
APP_THEME = "tentacles_dark"


def apply_theme(app: QApplication | None, theme_name: str | None = None) -> str:
    """
    Apply one of the Tentacles themes to the QApplication instance.

    Parameters
    ----------
    app:
        The QApplication to style. If None, no changes are made.
    theme_name:
        Optional override for APP_THEME. Returns the resolved theme key.
    """
    if app is None:
        return "unknown"

    global APP_THEME

    chosen = (theme_name or APP_THEME or "").strip().lower()
    if chosen not in THEME_PRESETS:
        chosen = "tentacles_dark"

    palette = QPalette()
    style_name = "Fusion"

    if chosen == "tentacles_dark":
        _configure_dark_palette(palette)
    elif chosen == "tentacles_light":
        _configure_light_palette(palette)
    elif chosen == "tentacles_purple":
        _configure_purple_palette(palette)
    elif chosen == "tentacles_green":
        _configure_green_palette(palette)

    available = {name.lower(): name for name in QStyleFactory.keys()}
    style_key = available.get(style_name.lower())
    if style_key:
        style_obj = QStyleFactory.create(style_key)
        if style_obj is not None:
            app.setStyle(style_obj)

    APP_THEME = chosen
    app.setPalette(palette)
    app.setStyleSheet("")

    try:
        for widget in app.topLevelWidgets():
            widget.setPalette(palette)
            widget.update()
    except Exception:
        pass

    return chosen


def _configure_dark_palette(palette: QPalette):
    palette.setColor(QPalette.ColorRole.Window, QColor(32, 35, 39))
    palette.setColor(QPalette.ColorRole.WindowText, Qt.GlobalColor.white)
    palette.setColor(QPalette.ColorRole.Base, QColor(22, 25, 28))
    palette.setColor(QPalette.ColorRole.AlternateBase, QColor(40, 43, 47))
    palette.setColor(QPalette.ColorRole.ToolTipBase, QColor(53, 56, 61))
    palette.setColor(QPalette.ColorRole.ToolTipText, Qt.GlobalColor.white)
    palette.setColor(QPalette.ColorRole.Text, Qt.GlobalColor.white)
    palette.setColor(QPalette.ColorRole.Button, QColor(45, 48, 52))
    palette.setColor(QPalette.ColorRole.ButtonText, Qt.GlobalColor.white)
    palette.setColor(QPalette.ColorRole.BrightText, Qt.GlobalColor.red)
    palette.setColor(QPalette.ColorRole.Highlight, QColor(0, 120, 215))
    palette.setColor(QPalette.ColorRole.HighlightedText, Qt.GlobalColor.white)
    palette.setColor(QPalette.ColorRole.Link, QColor(100, 180, 255))
    palette.setColor(QPalette.ColorRole.LinkVisited, QColor(120, 140, 255))
    palette.setColor(QPalette.ColorRole.PlaceholderText, QColor(180, 180, 180))
    _apply_disabled_group(
        palette,
        text_color=QColor(150, 153, 158),
        highlight_color=QColor(55, 70, 95),
        base_color=QColor(28, 30, 34),
        button_color=QColor(38, 41, 45),
    )


def _configure_light_palette(palette: QPalette):
    palette.setColor(QPalette.ColorRole.Window, QColor(245, 246, 248))
    palette.setColor(QPalette.ColorRole.WindowText, QColor(30, 32, 34))
    palette.setColor(QPalette.ColorRole.Base, QColor(255, 255, 255))
    palette.setColor(QPalette.ColorRole.AlternateBase, QColor(236, 238, 241))
    palette.setColor(QPalette.ColorRole.ToolTipBase, QColor(255, 255, 255))
    palette.setColor(QPalette.ColorRole.ToolTipText, QColor(30, 32, 34))
    palette.setColor(QPalette.ColorRole.Text, QColor(25, 27, 29))
    palette.setColor(QPalette.ColorRole.Button, QColor(235, 237, 240))
    palette.setColor(QPalette.ColorRole.ButtonText, QColor(30, 32, 34))
    palette.setColor(QPalette.ColorRole.BrightText, Qt.GlobalColor.red)
    palette.setColor(QPalette.ColorRole.Highlight, QColor(0, 120, 215))
    palette.setColor(QPalette.ColorRole.HighlightedText, Qt.GlobalColor.white)
    palette.setColor(QPalette.ColorRole.Link, QColor(60, 120, 200))
    palette.setColor(QPalette.ColorRole.LinkVisited, QColor(90, 100, 200))
    palette.setColor(QPalette.ColorRole.PlaceholderText, QColor(120, 130, 140))
    _apply_disabled_group(
        palette,
        text_color=QColor(150, 155, 165),
        highlight_color=QColor(200, 210, 225),
        base_color=QColor(240, 242, 245),
        button_color=QColor(220, 224, 228),
    )


def _configure_purple_palette(palette: QPalette):
    palette.setColor(QPalette.ColorRole.Window, QColor(35, 30, 48))
    palette.setColor(QPalette.ColorRole.WindowText, Qt.GlobalColor.white)
    palette.setColor(QPalette.ColorRole.Base, QColor(28, 24, 38))
    palette.setColor(QPalette.ColorRole.AlternateBase, QColor(45, 38, 65))
    palette.setColor(QPalette.ColorRole.ToolTipBase, QColor(70, 60, 96))
    palette.setColor(QPalette.ColorRole.ToolTipText, Qt.GlobalColor.white)
    palette.setColor(QPalette.ColorRole.Text, Qt.GlobalColor.white)
    palette.setColor(QPalette.ColorRole.Button, QColor(60, 52, 82))
    palette.setColor(QPalette.ColorRole.ButtonText, Qt.GlobalColor.white)
    palette.setColor(QPalette.ColorRole.BrightText, Qt.GlobalColor.red)
    palette.setColor(QPalette.ColorRole.Highlight, QColor(155, 89, 182))
    palette.setColor(QPalette.ColorRole.HighlightedText, Qt.GlobalColor.white)
    palette.setColor(QPalette.ColorRole.Link, QColor(180, 140, 255))
    palette.setColor(QPalette.ColorRole.LinkVisited, QColor(200, 160, 255))
    palette.setColor(QPalette.ColorRole.PlaceholderText, QColor(180, 160, 200))
    _apply_disabled_group(
        palette,
        text_color=QColor(185, 175, 205),
        highlight_color=QColor(110, 80, 135),
        base_color=QColor(36, 32, 50),
        button_color=QColor(52, 46, 70),
    )


def _configure_green_palette(palette: QPalette):
    palette.setColor(QPalette.ColorRole.Window, QColor(30, 44, 38))
    palette.setColor(QPalette.ColorRole.WindowText, Qt.GlobalColor.white)
    palette.setColor(QPalette.ColorRole.Base, QColor(22, 34, 28))
    palette.setColor(QPalette.ColorRole.AlternateBase, QColor(40, 60, 50))
    palette.setColor(QPalette.ColorRole.ToolTipBase, QColor(70, 90, 80))
    palette.setColor(QPalette.ColorRole.ToolTipText, Qt.GlobalColor.white)
    palette.setColor(QPalette.ColorRole.Text, Qt.GlobalColor.white)
    palette.setColor(QPalette.ColorRole.Button, QColor(54, 74, 64))
    palette.setColor(QPalette.ColorRole.ButtonText, Qt.GlobalColor.white)
    palette.setColor(QPalette.ColorRole.BrightText, Qt.GlobalColor.red)
    palette.setColor(QPalette.ColorRole.Highlight, QColor(76, 175, 80))
    palette.setColor(QPalette.ColorRole.HighlightedText, Qt.GlobalColor.black)
    palette.setColor(QPalette.ColorRole.Link, QColor(130, 200, 150))
    palette.setColor(QPalette.ColorRole.LinkVisited, QColor(150, 210, 170))
    palette.setColor(QPalette.ColorRole.PlaceholderText, QColor(170, 190, 180))
    _apply_disabled_group(
        palette,
        text_color=QColor(175, 190, 180),
        highlight_color=QColor(90, 130, 95),
        base_color=QColor(34, 46, 40),
        button_color=QColor(58, 78, 68),
    )


def _apply_disabled_group(
    palette: QPalette,
    *,
    text_color: QColor,
    highlight_color: QColor,
    base_color: QColor,
    button_color: QColor,
) -> None:
    for role in (
        QPalette.ColorRole.WindowText,
        QPalette.ColorRole.Text,
        QPalette.ColorRole.ButtonText,
        QPalette.ColorRole.ToolTipText,
        QPalette.ColorRole.Link,
        QPalette.ColorRole.LinkVisited,
    ):
        palette.setColor(QPalette.ColorGroup.Disabled, role, text_color)

    palette.setColor(QPalette.ColorGroup.Disabled, QPalette.ColorRole.Highlight, highlight_color)
    palette.setColor(QPalette.ColorGroup.Disabled, QPalette.ColorRole.HighlightedText, text_color)
    palette.setColor(QPalette.ColorGroup.Disabled, QPalette.ColorRole.Base, base_color)
    palette.setColor(QPalette.ColorGroup.Disabled, QPalette.ColorRole.Button, button_color)

    placeholder = QColor(text_color)
    placeholder.setAlpha(180)
    palette.setColor(QPalette.ColorGroup.Disabled, QPalette.ColorRole.PlaceholderText, placeholder)

def simple_hash(value):
    # Convert anything to string then bytes
    s = str(value).encode("utf-8")
    # Create SHA-256 hash
    h = hashlib.sha256(s).hexdigest()
    # Keep only alphanumeric chars
    alnum = re.sub(r'[^A-Za-z0-9]', '', h)
    # Return first 19 chars
    return alnum[:19]