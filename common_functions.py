import ast
import copy
import hashlib
import json
import os
import re
from dataclasses import dataclass, field
from pathlib import Path
from typing import Any, Dict, List, Optional, Tuple

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


# --- PBIP project backend ----------------------------------------------------


@dataclass
class PowerQueryMetadata:
    tables: Dict[str, Dict[str, Any]] = field(default_factory=dict)
    query_order: List[str] = field(default_factory=list)
    query_groups: Dict[str, int] = field(default_factory=dict)
    error: Optional[str] = None

    def clone(self) -> "PowerQueryMetadata":
        return PowerQueryMetadata(
            tables=copy.deepcopy(self.tables),
            query_order=list(self.query_order),
            query_groups=dict(self.query_groups),
            error=self.error,
        )

    @property
    def available(self) -> bool:
        return self.error is None


@dataclass
class DaxQueriesMetadata:
    queries: Dict[str, str] = field(default_factory=dict)
    tab_order: List[str] = field(default_factory=list)
    default_tab: Optional[str] = None
    error: Optional[str] = None

    def clone(self) -> "DaxQueriesMetadata":
        return DaxQueriesMetadata(
            queries=dict(self.queries),
            tab_order=list(self.tab_order),
            default_tab=self.default_tab,
            error=self.error,
        )

    @property
    def available(self) -> bool:
        return self.error is None


@dataclass
class BookmarkMetadata:
    bookmarks: Dict[str, Dict[str, Any]] = field(default_factory=dict)
    folders: Dict[str, Dict[str, Any]] = field(default_factory=dict)
    items: List[Dict[str, Any]] = field(default_factory=list)
    structure: List[Dict[str, Any]] = field(default_factory=list)
    warnings: List[str] = field(default_factory=list)
    error: Optional[str] = None

    def clone(self) -> "BookmarkMetadata":
        return BookmarkMetadata(
            bookmarks=copy.deepcopy(self.bookmarks),
            folders=copy.deepcopy(self.folders),
            items=copy.deepcopy(self.items),
            structure=copy.deepcopy(self.structure),
            warnings=list(self.warnings),
            error=self.error,
        )

    @property
    def available(self) -> bool:
        return self.error is None


class PBIPProject:
    """Container that caches PBIP project assets loaded from disk."""

    def __init__(self, pbip_file: str | Path):
        self.pbip_path = _resolve_pbip_file(pbip_file)
        self._project_dir = self.pbip_path.parent
        self._stem = self.pbip_path.stem
        self._semantic_model_dir = self._project_dir / f"{self._stem}.SemanticModel"
        self._report_dir = self._project_dir / f"{self._stem}.Report"
        self._tables_metadata = PowerQueryMetadata()
        self._dax_queries_metadata = DaxQueriesMetadata()
        self._bookmarks_metadata = BookmarkMetadata()
        self.refresh_all()

    @property
    def project_dir(self) -> Path:
        return self._project_dir

    @property
    def semantic_model_dir(self) -> Path:
        return self._semantic_model_dir

    @property
    def report_dir(self) -> Path:
        return self._report_dir

    def refresh_all(self) -> None:
        """Reload every metadata bundle from disk."""
        self.reload_tables()
        self.reload_dax_queries()
        self.reload_bookmarks()

    def reload_tables(self) -> PowerQueryMetadata:
        self._tables_metadata = _load_power_query_metadata(self.pbip_path)
        return self.get_power_query_metadata()

    def reload_dax_queries(self) -> DaxQueriesMetadata:
        self._dax_queries_metadata = _load_dax_queries_metadata(self.pbip_path)
        return self.get_dax_queries_metadata()

    def reload_bookmarks(self) -> BookmarkMetadata:
        self._bookmarks_metadata = _load_bookmarks_metadata(self.pbip_path)
        return self.get_bookmarks_metadata()

    def get_power_query_metadata(self, *, clone: bool = True) -> PowerQueryMetadata:
        return self._tables_metadata.clone() if clone else self._tables_metadata

    def update_power_query_metadata(
        self,
        tables: Dict[str, Dict[str, Any]],
        query_order: List[str],
        query_groups: Dict[str, int],
    ) -> None:
        self._tables_metadata = PowerQueryMetadata(
            tables=copy.deepcopy(tables),
            query_order=list(query_order),
            query_groups=dict(query_groups),
            error=None,
        )

    def get_dax_queries_metadata(self, *, clone: bool = True) -> DaxQueriesMetadata:
        return self._dax_queries_metadata.clone() if clone else self._dax_queries_metadata

    def update_dax_queries_metadata(
        self,
        tab_order: List[str],
        queries: Dict[str, str],
        default_tab: Optional[str],
    ) -> None:
        self._dax_queries_metadata = DaxQueriesMetadata(
            tab_order=list(tab_order),
            queries=dict(queries),
            default_tab=default_tab,
            error=None,
        )

    def get_bookmarks_metadata(self, *, clone: bool = True) -> BookmarkMetadata:
        return self._bookmarks_metadata.clone() if clone else self._bookmarks_metadata

    def update_bookmarks_metadata(
        self,
        bookmarks: Dict[str, Dict[str, Any]],
        folders: Dict[str, Dict[str, Any]],
        items: List[Dict[str, Any]],
        warnings: Optional[List[str]] = None,
    ) -> None:
        self._bookmarks_metadata = BookmarkMetadata(
            bookmarks=copy.deepcopy(bookmarks),
            folders=copy.deepcopy(folders),
            items=copy.deepcopy(items),
            structure=[
                {"type": entry.get("type"), "id": entry.get("id")}
                for entry in items
                if isinstance(entry, dict) and entry.get("type") and entry.get("id")
            ],
            warnings=list(warnings or []),
            error=None,
        )


_PROJECT_CACHE: Dict[str, PBIPProject] = {}


def load_pbip_project(pbip_file: str | Path, *, force_reload: bool = False) -> PBIPProject:
    """Return a cached PBIPProject, loading metadata if necessary."""
    candidate = Path(pbip_file).expanduser().resolve()
    key = str(candidate)
    project = _PROJECT_CACHE.get(key)
    if project is None:
        project = PBIPProject(candidate)
        _PROJECT_CACHE[key] = project
    elif force_reload:
        project.refresh_all()
    return project


def clear_project_cache() -> None:
    """Drop all cached PBIPProject instances."""
    _PROJECT_CACHE.clear()


def _resolve_pbip_file(pbip_file: str | Path) -> Path:
    pbip_path = Path(pbip_file).expanduser().resolve()
    if pbip_path.suffix.lower() != ".pbip":
        raise ValueError(f"Expected a .pbip file, got '{pbip_path}'.")
    if not pbip_path.is_file():
        raise FileNotFoundError(f"PBIP file not found: {pbip_path}")
    return pbip_path


def _load_power_query_metadata(pbip_file: Path) -> PowerQueryMetadata:
    metadata = PowerQueryMetadata()
    try:
        semantic_root = pbip_file.parent / f"{pbip_file.stem}.SemanticModel"
        model_tmdl = semantic_root / "definition" / "model.tmdl"
        tables_dir = semantic_root / "definition" / "tables"

        if not model_tmdl.is_file():
            raise FileNotFoundError(f"model.tmdl not found under {semantic_root}")
        if not tables_dir.is_dir():
            raise FileNotFoundError(f"tables folder not found under {semantic_root}")

        model_text = model_tmdl.read_text(encoding="utf-8")
        metadata.query_order = _parse_query_order(model_text)
        metadata.query_groups = _parse_query_groups(model_text)

        tables: Dict[str, Dict[str, Any]] = {}
        for table_path in sorted(tables_dir.glob("*.tmdl")):
            try:
                content = table_path.read_text(encoding="utf-8")
            except OSError:
                continue
            tables[table_path.stem] = _parse_table_tmdl(table_path, content)

        metadata.tables = tables
    except Exception as exc:  # pragma: no cover - defensive
        metadata.error = str(exc)
    return metadata


def _parse_query_order(model_text: str) -> List[str]:
    match = re.search(r"annotation\s+PBI_QueryOrder\s*=\s*(\[.*?\])", model_text, re.DOTALL)
    if not match:
        return []

    try:
        value = ast.literal_eval(match.group(1))
    except (ValueError, SyntaxError):
        return []

    return [str(item) for item in value if isinstance(item, str)]


def _parse_query_groups(model_text: str) -> Dict[str, int]:
    pattern = re.compile(
        r"(?mi)^\s*queryGroup\s+(?P<name>'[^']+'|\"[^\"]+\"|[^\s\r\n]+)\s*\r?\n\s*annotation\s+PBI_QueryGroupOrder\s*=\s*(?P<order>\d+)"
    )
    groups: Dict[str, int] = {}

    for match in pattern.finditer(model_text):
        normalized = _normalize_group_path(match.group("name"))
        if not normalized:
            continue
        order_value = int(match.group("order"))
        existing = groups.get(normalized)
        if existing is None or order_value < existing:
            groups[normalized] = order_value

    return groups


def _normalize_group_path(raw: Optional[str]) -> Optional[str]:
    if not raw:
        return None
    text = raw.strip()
    if (text.startswith("'") and text.endswith("'")) or (text.startswith('"') and text.endswith('"')):
        text = text[1:-1]
    text = re.sub(r"[\\/]+", "/", text)
    parts = [part.strip() for part in text.split("/") if part.strip()]
    return "/".join(parts) if parts else None


def _parse_table_tmdl(table_path: Path, tmdl_text: str) -> Dict[str, Any]:
    column_pattern = re.compile(r'(?mi)^\s*column\s+(?:"([^"]+)"|([A-Za-z0-9_]+))\s*$')
    columns: List[str] = []
    for match in column_pattern.finditer(tmdl_text):
        column = match.group(1) or match.group(2) or ""
        column = column.strip()
        if column:
            columns.append(column)

    mode_match = re.search(r"(?mi)^\s*mode\s*:\s*([^\r\n]+)", tmdl_text)
    if mode_match:
        mode_value = mode_match.group(1).strip().lower()
    else:
        data_mode_match = re.search(
            r'(?mi)^\s*annotation\s+PBI_DataMode\s*=\s*"?(?P<value>.*?)"?\s*$',
            tmdl_text,
        )
        mode_value = data_mode_match.group("value").strip().lower() if data_mode_match else None

    table_type_match = re.search(r"(?mi)^\s*partition\s+[A-Za-z0-9_-]+\s*=\s*(m|calculated)\s*$", tmdl_text)
    table_type = table_type_match.group(1).lower() if table_type_match else "m"

    query_group_match = re.search(r"(?mi)^\s*queryGroup\s*:\s*([^\r\n]+)", tmdl_text) or re.search(
        r"(?mi)^\s*queryGroup\s+([^\r\n]+)", tmdl_text
    )
    query_group = _normalize_group_path(query_group_match.group(1)) if query_group_match else None

    code_text = _extract_table_code(tmdl_text) or ""
    code_language = "dax" if table_type == "calculated" else "m"

    return {
        "columns": columns,
        "import_mode": mode_value,
        "query_group": query_group,
        "code_text": code_text,
        "code_language": code_language,
        "table_type": table_type,
        "tmdl_path": str(table_path),
    }


def _unescape_quoted(text: str) -> str:
    try:
        return text.encode("utf-8").decode("unicode_escape")
    except UnicodeDecodeError:
        return text


def _strip_any_fence(text: str) -> str:
    if not text:
        return text
    stripped = text.strip()
    for fence in ("`", "~"):
        multiline = re.compile(rf"^\s*{fence}{{3}}[^\r\n]*\r?\n([\s\S]*?)\r?\n{fence}{{3}}\s*$")
        match = multiline.match(stripped)
        if match:
            return match.group(1).strip()
        inline = re.compile(rf"^\s*{fence}{{3}}[^\r\n]*\s*([\s\S]*?)\s*{fence}{{3}}\s*$")
        match = inline.match(stripped)
        if match:
            return match.group(1).strip()
    return stripped


def _extract_table_code(tmdl_text: str) -> Optional[str]:
    normalized = tmdl_text.replace("    ", "\t")

    quoted = re.search(r'(?ms)^\s*expression\s*=\s*"((?:[^"\\]|\\.)*)"', normalized)
    if quoted:
        return _strip_any_fence(_unescape_quoted(quoted.group(1)).strip())

    source_line = re.search(r"(?m)^\s*source\s*=\s*$", normalized)
    if not source_line:
        inline = re.search(r"(?ms)^\s*source\s*=\s*(.+?)(?=^\s*annotation\b|^\S|\Z)", normalized)
        if inline:
            result = inline.group(1).rstrip()
            result = re.sub(r"^(?:[ \t]{0,4})", "", result, flags=re.MULTILINE)
            return _strip_any_fence(result.strip())
        return None

    start = source_line.end()
    lines = normalized[start:].splitlines()

    index = 0
    while index < len(lines) and lines[index].strip() == "":
        index += 1
    if index >= len(lines):
        return None

    first_line = lines[index]
    base_indent = len(first_line) - len(first_line.lstrip())

    captured: List[str] = []
    for line in lines[index:]:
        stripped = line.lstrip()
        if stripped.startswith("annotation "):
            break
        if stripped and (len(line) - len(stripped)) < base_indent:
            break
        captured.append(line)

    while captured and captured[-1].strip() == "":
        captured.pop()

    result = "\n".join(captured)
    result = re.sub(r"^(?:[ \t]{0,4})", "", result, flags=re.MULTILINE)
    return _strip_any_fence(result.strip())


def _load_dax_queries_metadata(pbip_file: Path) -> DaxQueriesMetadata:
    metadata = DaxQueriesMetadata()
    try:
        dax_root = pbip_file.parent / f"{pbip_file.stem}.SemanticModel" / "DAXQueries"
        json_path = dax_root / ".pbi" / "daxQueries.json"
        if not json_path.is_file():
            raise FileNotFoundError(f"daxQueries.json not found under {dax_root}")

        data = json.loads(json_path.read_text(encoding="utf-8"))
        tab_order = data.get("tabOrder") or []
        if not isinstance(tab_order, list):
            raise ValueError("daxQueries.json missing a valid 'tabOrder' list.")

        metadata.tab_order = [str(entry) for entry in tab_order if isinstance(entry, str)]
        default_tab = data.get("defaultTab")
        metadata.default_tab = default_tab if isinstance(default_tab, str) else None

        queries: Dict[str, str] = {}
        for name in metadata.tab_order:
            dax_path = dax_root / f"{name}.dax"
            if not dax_path.is_file():
                continue
            try:
                queries[name] = dax_path.read_text(encoding="utf-8")
            except OSError:
                continue

        metadata.queries = queries
    except Exception as exc:  # pragma: no cover - defensive
        metadata.error = str(exc)
    return metadata


def _load_bookmarks_metadata(pbip_file: Path) -> BookmarkMetadata:
    metadata = BookmarkMetadata()
    try:
        base_dir = pbip_file.parent / f"{pbip_file.stem}.Report" / "definition" / "bookmarks"
        if not base_dir.is_dir():
            raise FileNotFoundError(f"Bookmarks folder not found under {base_dir.parent}")

        bookmarks_json = base_dir / "bookmarks.json"
        if not bookmarks_json.is_file():
            raise FileNotFoundError("bookmarks.json not found in bookmarks folder.")

        data = json.loads(bookmarks_json.read_text(encoding="utf-8"))
        items = data.get("items") or []
        if not isinstance(items, list):
            raise ValueError("bookmarks.json missing a valid 'items' list.")

        folders: Dict[str, Dict[str, Any]] = {}
        for entry in items:
            if not isinstance(entry, dict):
                continue
            name = entry.get("name")
            if not isinstance(name, str):
                continue
            children = entry.get("children")
            if isinstance(children, list):
                valid_children = [str(child) for child in children if isinstance(child, str)]
            else:
                valid_children = []
            if "children" in entry:
                display = entry.get("displayName") or name
                folders[name] = {"display": display, "children": valid_children}

        bookmarks: Dict[str, Dict[str, Any]] = {}
        warnings: List[str] = []

        for bookmark_path in base_dir.glob("*.bookmark.json"):
            stem = bookmark_path.name[: -len(".bookmark.json")]
            display_name = stem
            valid = True
            error_message: Optional[str] = None
            try:
                bookmark_data = json.loads(bookmark_path.read_text(encoding="utf-8"))
                display_name = bookmark_data.get("displayName") or display_name
            except json.JSONDecodeError as exc:
                display_name = f"{stem} (invalid)"
                valid = False
                error_message = f"Invalid JSON: {exc}"
                warnings.append(f"Bookmark '{stem}' has invalid JSON.")
            except Exception as exc:  # pragma: no cover - defensive
                display_name = f"{stem} (unreadable)"
                valid = False
                error_message = str(exc)
                warnings.append(f"Bookmark '{stem}' could not be read.")

            bookmarks[stem] = {
                "display_name": display_name,
                "path": str(bookmark_path),
                "valid": valid,
                "error": error_message,
                "used": False,
            }

        pages_dir = pbip_file.parent / f"{pbip_file.stem}.Report" / "definition" / "pages"
        _compute_bookmark_usage(bookmarks, pages_dir)

        metadata.bookmarks = bookmarks
        metadata.folders = folders
        metadata.items = [entry for entry in items if isinstance(entry, dict)]
        metadata.structure = [
            {"type": ("folder" if "children" in entry else "bookmark"), "id": entry.get("name")}
            for entry in metadata.items
            if isinstance(entry.get("name"), str)
        ]
        metadata.warnings = warnings
    except Exception as exc:  # pragma: no cover - defensive
        metadata.error = str(exc)
    return metadata


def _compute_bookmark_usage(bookmarks: Dict[str, Dict[str, Any]], pages_dir: Path) -> None:
    if not pages_dir.is_dir():
        return

    remaining = {name for name, meta in bookmarks.items() if meta.get("valid")}
    if not remaining:
        return

    for root, _, files in os.walk(pages_dir):
        for fname in files:
            if not fname.lower().endswith(".json"):
                continue
            path = Path(root) / fname
            try:
                content = path.read_text(encoding="utf-8", errors="ignore")
            except OSError:
                continue

            matched = {name for name in list(remaining) if name in content}
            if not matched:
                continue

            for name in matched:
                bookmarks[name]["used"] = True
            remaining.difference_update(matched)
            if not remaining:
                return
