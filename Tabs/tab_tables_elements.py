import html
import json
import os
import re
import threading
import urllib.parse
import urllib.request
import uuid
from collections import defaultdict
from pathlib import Path
from typing import Any, Callable, Dict, Optional, List, Tuple, Set

from PyQt6.QtCore import Qt, QObject, QRunnable, QThreadPool, pyqtSignal
from PyQt6.QtWidgets import (
    QApplication,
    QAbstractItemView,
    QButtonGroup,
    QFileDialog,
    QFormLayout,
    QHBoxLayout,
    QLabel,
    QMessageBox,
    QPushButton,
    QComboBox,
    QSplitter,
    QTreeWidget,
    QTreeWidgetItem,
    QVBoxLayout,
    QWidget,
    QMenu,
    QHeaderView,
)
from PyQt6.QtGui import (
    QFont,
    QIcon,
    QDragMoveEvent,
    QDropEvent,
    QShortcut,
    QKeySequence,
    QPixmap,
    QImage,
    QBrush,
    QColor,
    QSyntaxHighlighter,
    QTextCharFormat,
)
from Coding.code_editor import CodeEditor
from Coding.code_editor_support import DAXHighlighter, set_dax_model_identifiers


class _TableColumnHighlightMixin:
    """Shared highlight logic for table and column references."""

    def _init_table_column_support(self) -> None:
        self._table_patterns: List[re.Pattern] = []
        self._column_patterns: List[re.Pattern] = []

        self._table_format = QTextCharFormat()
        self._table_format.setForeground(QColor("#2c7be5"))
        self._table_format.setFontWeight(QFont.Weight.Medium)

        self._column_format = QTextCharFormat()
        self._column_format.setForeground(QColor("#00a676"))
        self._column_format.setFontWeight(QFont.Weight.Medium)

    def update_patterns(
        self,
        table_patterns: List[re.Pattern],
        column_patterns: List[re.Pattern],
    ) -> None:
        self._table_patterns = table_patterns
        self._column_patterns = column_patterns
        self.rehighlight()

    def _apply_table_column_highlighting(self, text: str) -> None:
        for pattern in self._table_patterns:
            for match in pattern.finditer(text):
                start, end = match.span()
                self.setFormat(start, end - start, self._table_format)

        for pattern in self._column_patterns:
            for match in pattern.finditer(text):
                start, end = match.span()
                self.setFormat(start, end - start, self._column_format)


class TableColumnHighlighter(_TableColumnHighlightMixin, QSyntaxHighlighter):
    """Syntax highlighter that colors table and column references."""

    def __init__(self, document):
        QSyntaxHighlighter.__init__(self, document)
        self._init_table_column_support()

    def highlightBlock(self, text: str) -> None:  # noqa: N802
        self._apply_table_column_highlighting(text)


class DAXTableColumnHighlighter(_TableColumnHighlightMixin, DAXHighlighter):
    """DAX highlighter augmented with table/column coloring."""

    def __init__(self, document):
        DAXHighlighter.__init__(self, document)
        self._init_table_column_support()

    def highlightBlock(self, text: str) -> None:  # noqa: N802
        super().highlightBlock(text)
        self._apply_table_column_highlighting(text)


class ChatRequestSignals(QObject):
    finished = pyqtSignal(str)
    error = pyqtSignal(str)


class ChatRequestWorker(QRunnable):
    def __init__(self, fn, *args, **kwargs):
        super().__init__()
        self.fn = fn
        self.args = args
        self.kwargs = kwargs
        self.signals = ChatRequestSignals()

    def run(self) -> None:  # pragma: no cover - background thread
        try:
            result = self.fn(*self.args, **self.kwargs)
        except Exception as exc:  # pragma: no cover - defensive
            self.signals.error.emit(str(exc))
        else:
            self.signals.finished.emit(result)


class ChatGPTFreeClient:
    CHAT_URL = "https://chatgptfree.ai/chat/"

    def __init__(self):
        self._opener = urllib.request.build_opener(urllib.request.HTTPCookieProcessor())
        self._config: Optional[Dict[str, Any]] = None
        self._session_id = str(uuid.uuid4())
        self._conversation_id = str(uuid.uuid4())
        self._lock = threading.Lock()

    def _ensure_config(self) -> None:
        if self._config is not None:
            return
        with self._lock:
            if self._config is not None:
                return
            request = urllib.request.Request(self.CHAT_URL)
            with self._opener.open(request, timeout=30) as response:
                html_text = response.read().decode("utf-8", errors="replace")
            data_prefix = "window.__NEXT_DATA__ = "
            idx = html_text.find(data_prefix)
            if idx == -1:
                raise RuntimeError("Unable to locate configuration payload.")
            start = idx + len(data_prefix)
            end = html_text.find("</script>", start)
            payload = html_text[start:end].rstrip(";\n ")
            obj = json.loads(payload)
            page_props = obj.get("props", {}).get("pageProps", {})
            self._config = {
                "model": page_props.get("defaultOpenAIModelName", "gpt-3.5-turbo"),
                "auth_token": page_props.get("authToken"),
                "csrf_token": page_props.get("csrfToken"),
            }

    def generate(self, prompt: str) -> str:
        self._ensure_config()
        if not self._config:
            raise RuntimeError("ChatGPT configuration unavailable.")
        params = {
            "conversation_id": self._conversation_id,
            "session_id": self._session_id,
            "action": "add",
            "messages": [
                {
                    "id": str(uuid.uuid4()),
                    "role": "user",
                    "content": {"content_type": "text", "parts": [prompt]},
                }
            ],
            "model": self._config.get("model"),
        }

        headers = {
            "Content-Type": "application/json",
            "Accept": "text/event-stream",
            "Referer": self.CHAT_URL,
        }
        auth_token = self._config.get("auth_token")
        csrf_token = self._config.get("csrf_token")
        if auth_token:
            headers["Authorization"] = f"Bearer {auth_token}"
        if csrf_token:
            headers["X-CSRF-Token"] = csrf_token

        data = json.dumps(params).encode("utf-8")
        request = urllib.request.Request(
            urllib.parse.urljoin(self.CHAT_URL, "/api/conversation"),
            data=data,
            headers=headers,
        )

        with self._opener.open(request, timeout=60) as response:
            buffer = ""
            while True:
                chunk = response.readline().decode("utf-8", errors="replace")
                if not chunk:
                    break
                buffer += chunk
                if "[DONE]" in chunk:
                    break

        result = ""
        for line in buffer.splitlines():
            if not line.startswith("data: "):
                continue
            payload = line[6:].strip()
            if payload in {"[DONE]", ""}:
                continue
            try:
                data = json.loads(payload)
            except json.JSONDecodeError:
                continue
            message = data.get("message", {})
            content = message.get("content", {})
            parts = content.get("parts") or []
            if parts:
                result = parts[0]

        if not result:
            raise RuntimeError("No response received from ChatGPT.")
        cleaned = html.unescape(result.strip())
        return cleaned
from common_functions import code_editor_font, APP_THEME, PBIPProject, load_pbip_project, _parse_table_measures


class HierarchyTree(QTreeWidget):
    """QTreeWidget that prevents tables from being nested while still allowing reordering."""

    structure_changed = pyqtSignal()

    def __init__(self, type_role: int, table_marker: str, parent: Optional[QWidget] = None):
        super().__init__(parent)
        self._type_role = type_role
        self._table_marker = table_marker
        self._drop_validator: Optional[
            Callable[[Optional[QTreeWidgetItem], List[QTreeWidgetItem], QAbstractItemView.DropIndicatorPosition], bool]
        ] = None

    def set_drop_validator(
        self,
        validator: Optional[
            Callable[[Optional[QTreeWidgetItem], List[QTreeWidgetItem], QAbstractItemView.DropIndicatorPosition], bool]
        ],
    ) -> None:
        self._drop_validator = validator

    def _validate_drop_event(self, event: QDropEvent) -> bool:
        if self._drop_validator is None:
            return True
        target = self.itemAt(event.position().toPoint())
        indicator = self.dropIndicatorPosition()
        items = self.selectedItems()
        if not items:
            return False
        return self._drop_validator(target, items, indicator)

    def _drag_contains_tables(self) -> bool:
        for item in self.selectedItems():
            if item and item.data(0, self._type_role) == self._table_marker:
                return True
        return False

    def _is_drop_on_table(self, event: QDropEvent) -> bool:
        if self.dropIndicatorPosition() != QAbstractItemView.DropIndicatorPosition.OnItem:
            return False
        target = self.itemAt(event.position().toPoint())
        if target is None:
            return False
        return target.data(0, self._type_role) == self._table_marker

    def _item_path(self, item: QTreeWidgetItem) -> Tuple[int, ...]:
        path: List[int] = []
        current = item
        while current is not None:
            parent = current.parent()
            if parent is None:
                idx = self.indexOfTopLevelItem(current)
            else:
                idx = parent.indexOfChild(current)
            path.insert(0, idx)
            current = parent
        return tuple(path)

    def _remove_item(self, item: QTreeWidgetItem) -> None:
        parent = item.parent()
        if parent is None:
            idx = self.indexOfTopLevelItem(item)
            if idx != -1:
                self.takeTopLevelItem(idx)
        else:
            idx = parent.indexOfChild(item)
            if idx != -1:
                parent.takeChild(idx)

    def _insert_item(
        self,
        container: Optional[QTreeWidgetItem],
        index: int,
        item: QTreeWidgetItem,
    ) -> None:
        if container is None:
            self.insertTopLevelItem(index, item)
        else:
            container.insertChild(index, item)

    def dragMoveEvent(self, event: QDragMoveEvent) -> None:
        if self._drag_contains_tables() and self._is_drop_on_table(event):
            event.accept()
            return
        super().dragMoveEvent(event)

    def dropEvent(self, event: QDropEvent) -> None:
        if not self._validate_drop_event(event):
            event.ignore()
            return

        if not (self._drag_contains_tables() and self._is_drop_on_table(event)):
            super().dropEvent(event)
            return

        target = self.itemAt(event.position().toPoint())
        if target is None:
            event.ignore()
            return

        tables = [
            item
            for item in self.selectedItems()
            if item.data(0, self._type_role) == self._table_marker and item is not target
        ]
        if not tables:
            event.ignore()
            return

        parent = target.parent()
        if parent is None and self.indexOfTopLevelItem(target) == -1:
            event.ignore()
            return

        container = parent
        index = (
            container.indexOfChild(target)
            if container is not None
            else self.indexOfTopLevelItem(target)
        )

        rect = self.visualItemRect(target)
        insert_after = event.position().toPoint().y() > rect.center().y()
        if insert_after:
            index += 1

        tables_sorted = sorted(tables, key=self._item_path)

        # Remove moving tables first (deepest items removed first to avoid index shifts)
        for table_item in sorted(tables_sorted, key=self._item_path, reverse=True):
            self._remove_item(table_item)

        insertion_order = tables_sorted if insert_after else list(reversed(tables_sorted))

        current_index = index
        for table_item in insertion_order:
            self._insert_item(container, current_index, table_item)
            current_index += 1

        # Restore selection on moved items
        self.clearSelection()
        for item in tables_sorted:
            item.setSelected(True)

        self.setCurrentItem(tables_sorted[-1])
        self.structure_changed.emit()
        event.accept()

    def edit(self, index, trigger, event):
        if index.column() != 0:
            return False
        return super().edit(index, trigger, event)


class PowerQueryTab(QWidget):
    """Tab that surfaces Power Query table metadata with editable details."""

    TYPE_ROLE = Qt.ItemDataRole.UserRole + 1
    KEY_ROLE = Qt.ItemDataRole.UserRole + 2
    ITEM_FOLDER = "folder"
    ITEM_TABLE = "table"
    ITEM_COLUMN = "column"
    ITEM_MEASURE_FOLDER = "measure_folder"
    ITEM_MEASURE = "measure"
    OTHER_QUERIES_NAME = "Other Queries"
    ALLOWED_CHILDREN: Dict[str, Set[str]] = {
        ITEM_FOLDER: frozenset({ITEM_FOLDER, ITEM_TABLE}),
        ITEM_TABLE: frozenset({ITEM_COLUMN, ITEM_MEASURE_FOLDER, ITEM_MEASURE}),
        ITEM_MEASURE_FOLDER: frozenset({ITEM_MEASURE_FOLDER, ITEM_MEASURE}),
        ITEM_MEASURE: frozenset(),
        ITEM_COLUMN: frozenset(),
    }
    ALLOWED_PARENTS: Dict[str, Set[Optional[str]]] = {
        ITEM_FOLDER: frozenset({None, ITEM_FOLDER}),
        ITEM_TABLE: frozenset({ITEM_FOLDER}),
        ITEM_COLUMN: frozenset({ITEM_TABLE}),
        ITEM_MEASURE_FOLDER: frozenset({ITEM_TABLE, ITEM_MEASURE_FOLDER}),
        ITEM_MEASURE: frozenset({ITEM_TABLE, ITEM_MEASURE_FOLDER}),
    }

    def __init__(self, project: Optional[PBIPProject] = None, pbip_file: Optional[str] = None):
        super().__init__()
        self.project = project
        self.pbip_file = str(project.pbip_path) if project else pbip_file
        self.tables_data: Dict[str, Dict[str, Any]] = {}
        self.query_order = []
        self.query_groups = {}
        self.current_table: Optional[str] = None
        self.current_measure_id: Optional[str] = None
        self.measure_view_mode = "expression"
        self.ignore_editor_changes = False
        self._ignore_tree_changes = False
        self._ignore_item_change = False
        self._loading_data = False
        self._table_usage_flags: Dict[str, bool] = {}
        self._column_usage_flags: Dict[Tuple[str, str], bool] = {}
        self._measure_usage_flags: Dict[str, bool] = {}
        self.dax_prompt_editor: Optional[CodeEditor] = None
        self.dax_output_editor: Optional[CodeEditor] = None
        self.dax_generate_button: Optional[QPushButton] = None
        self.dax_copy_button: Optional[QPushButton] = None
        self.dax_count_label: Optional[QLabel] = None
        self.dax_status_label: Optional[QLabel] = None
        self.dax_controls_widget: Optional[QWidget] = None
        self.dax_writer_container: Optional[QWidget] = None
        self.dax_prompt_highlighter: Optional[TableColumnHighlighter] = None
        self.dax_output_highlighter: Optional[DAXTableColumnHighlighter] = None
        self.dax_table_patterns: List[Tuple[re.Pattern, str]] = []
        self.dax_column_patterns: List[Tuple[List[re.Pattern], re.Pattern, Tuple[str, str]]] = []
        self._dax_column_highlight_patterns: List[re.Pattern] = []
        self.dax_thread_pool = QThreadPool.globalInstance()
        self.dax_api_client = ChatGPTFreeClient()
        self._dax_busy = False
        self._dax_prompt_shortcuts: List[QShortcut] = []
        self.is_dirty = False
        self.save_button: Optional[QPushButton] = None
        self.init_ui()
        if not self.project and self.pbip_file:
            try:
                self.project = load_pbip_project(self.pbip_file)
                self.pbip_file = str(self.project.pbip_path)
            except Exception:
                self.project = None
        if self.project or self.pbip_file:
            self.load_tables()

    def init_ui(self):
        """Build the user interface."""
        self.dax_prompt_editor = None
        self.dax_output_editor = None
        self.dax_controls_widget = None
        self.dax_writer_container = None
        self.dax_generate_button = None
        self.dax_copy_button = None
        self.dax_count_label = None
        self.dax_status_label = None
        self.dax_prompt_highlighter = None
        self.dax_output_highlighter = None
        self._dax_prompt_shortcuts = []
        self.dax_writer_button = None
        self.hotkey_hint_right = None
        self.measure_form_widget = None

        main_layout = QVBoxLayout(self)
        main_layout.setContentsMargins(8, 8, 8, 8)
        main_layout.setSpacing(8)

        top_bar = QHBoxLayout()
        self.save_button = QPushButton("💾 Save Changes")
        self.save_button.setEnabled(False)
        self.save_button.clicked.connect(self.save_changes)
        top_bar.addWidget(self.save_button)
        self.refresh_button = QPushButton("🔄 Reload Metadata")
        self.refresh_button.setToolTip("Reload Power Query metadata from disk")
        self.refresh_button.clicked.connect(self.refresh_tables)
        top_bar.addWidget(self.refresh_button)

        self.expand_button = QPushButton("⏬ Expand All")
        self.expand_button.clicked.connect(self.expand_all_groups)
        top_bar.addWidget(self.expand_button)

        self.collapse_button = QPushButton("⏫ Collapse All")
        self.collapse_button.clicked.connect(self.collapse_all_groups)
        top_bar.addWidget(self.collapse_button)

        self.sort_button = QPushButton("↑↓ Sort A-Z")
        self.sort_button.setToolTip("Sort folders and tables alphabetically (Other Queries stays last)")
        self.sort_button.clicked.connect(self.sort_folders_and_tables)
        top_bar.addWidget(self.sort_button)

        top_bar.addStretch()
        main_layout.addLayout(top_bar)

        splitter = QSplitter(Qt.Orientation.Horizontal)

        left_widget = QWidget()
        left_layout = QVBoxLayout(left_widget)
        left_layout.setContentsMargins(0, 0, 3, 0)

        left_layout.addWidget(QLabel("Tables"))

        self.table_tree = HierarchyTree(self.TYPE_ROLE, self.ITEM_TABLE)
        self.table_tree.set_drop_validator(self._validate_tree_drop)
        self.table_tree.setColumnCount(2)
        self.table_tree.setHeaderHidden(False)
        self.table_tree.setHeaderLabels(["Name", "Usage"])
        header = self.table_tree.header()
        header.setStretchLastSection(False)
        header.setSectionResizeMode(0, QHeaderView.ResizeMode.Stretch)
        header.setSectionResizeMode(1, QHeaderView.ResizeMode.ResizeToContents)
        self.table_tree.setSelectionMode(QAbstractItemView.SelectionMode.ExtendedSelection)
        self.table_tree.itemSelectionChanged.connect(self.on_tree_selection_changed)
        left_layout.addWidget(self.table_tree)
        self.table_tree.setDragEnabled(True)
        self.table_tree.setAcceptDrops(True)
        self.table_tree.setDropIndicatorShown(True)
        self.table_tree.setAlternatingRowColors(True)
        self.table_tree.setDefaultDropAction(Qt.DropAction.MoveAction)
        self.table_tree.setDragDropMode(QTreeWidget.DragDropMode.InternalMove)
        self.table_tree.setContextMenuPolicy(Qt.ContextMenuPolicy.CustomContextMenu)
        self.table_tree.customContextMenuRequested.connect(self.show_tree_context_menu)
        self.table_tree.structure_changed.connect(self.on_tree_structure_changed)
        self.table_tree.itemChanged.connect(self.on_tree_item_changed)
        try:
            self.table_tree.model().rowsMoved.connect(self.on_tree_structure_changed)
        except Exception:
            pass

        icon_dir = os.path.abspath(os.path.join(os.path.dirname(os.path.dirname(__file__)), "Images", "icons"))

        def _load_icon(name: str) -> QIcon:
            path = os.path.join(icon_dir, name)
            if not os.path.exists(path):
                return QIcon()
            if APP_THEME != "tentacles_light":
                pixmap = QPixmap(path)
                if not pixmap.isNull():
                    image = pixmap.toImage()
                    image.invertPixels(QImage.InvertMode.InvertRgb)
                    pixmap = QPixmap.fromImage(image)
                    return QIcon(pixmap)
            return QIcon(path)

        self.table_icons = {
            "m": _load_icon("Table.svg"),
            "calculated": _load_icon("Calculated-Table.svg"),
        }
        self.folder_icon = _load_icon("Folder.svg")
        self.measure_icon = _load_icon("Measure.svg")

        splitter.addWidget(left_widget)

        right_widget = QWidget()
        right_layout = QVBoxLayout(right_widget)
        right_layout.setContentsMargins(3, 0, 0, 0)

        self.query_label = QLabel("Query")
        right_layout.addWidget(self.query_label)

        self.measure_toggle_widget = QWidget()
        toggle_layout = QHBoxLayout(self.measure_toggle_widget)
        toggle_layout.setContentsMargins(0, 0, 0, 0)
        toggle_layout.setSpacing(6)

        self.measure_button_group = QButtonGroup(self.measure_toggle_widget)
        self.measure_button_group.setExclusive(True)
        self.measure_expression_button = QPushButton("Measure")
        self.measure_expression_button.setCheckable(True)
        self.measure_expression_button.setChecked(True)
        self.measure_expression_button.setToolTip("Show measure expression")
        self.measure_expression_button.clicked.connect(self.show_measure_expression)
        toggle_layout.addWidget(self.measure_expression_button)
        self.measure_button_group.addButton(self.measure_expression_button)

        self.measure_format_button = QPushButton("Format String")
        self.measure_format_button.setCheckable(True)
        self.measure_format_button.setToolTip("Show formatStringDefinition")
        self.measure_format_button.clicked.connect(self.show_measure_format)
        toggle_layout.addWidget(self.measure_format_button)
        self.measure_button_group.addButton(self.measure_format_button)

        self.dax_writer_button = QPushButton("DAX Writer")
        self.dax_writer_button.setCheckable(True)
        self.dax_writer_button.setToolTip("Generate a new measure using DAX Writer")
        self.dax_writer_button.clicked.connect(self.show_dax_writer)
        toggle_layout.addWidget(self.dax_writer_button)
        self.measure_button_group.addButton(self.dax_writer_button)

        toggle_layout.addStretch()
        self.measure_toggle_widget.hide()
        right_layout.addWidget(self.measure_toggle_widget)

        self.dax_controls_widget = QWidget()
        dax_controls_layout = QHBoxLayout(self.dax_controls_widget)
        dax_controls_layout.setContentsMargins(0, 0, 0, 0)
        dax_controls_layout.setSpacing(8)
        self.dax_generate_button = QPushButton("Generate DAX")
        self.dax_generate_button.clicked.connect(self.generate_dax_measure)
        dax_controls_layout.addWidget(self.dax_generate_button)
        self.dax_copy_button = QPushButton("Copy Output")
        self.dax_copy_button.clicked.connect(self.copy_dax_output)
        dax_controls_layout.addWidget(self.dax_copy_button)
        dax_controls_layout.addStretch()
        self.dax_status_label = QLabel("")
        self.dax_status_label.setStyleSheet("color: #888888;")
        dax_controls_layout.addWidget(self.dax_status_label)
        self.dax_controls_widget.hide()
        right_layout.addWidget(self.dax_controls_widget)

        self.dax_writer_container = QWidget()
        dax_writer_layout = QVBoxLayout(self.dax_writer_container)
        dax_writer_layout.setContentsMargins(0, 0, 0, 0)
        dax_writer_layout.setSpacing(6)
        dax_instruction = QLabel(
            "Describe the measure you need. Reference model objects as Table[Column] "
            "(for example: Sales[Amount])."
        )
        dax_instruction.setWordWrap(True)
        dax_writer_layout.addWidget(dax_instruction)

        self.dax_prompt_editor = CodeEditor()
        self.dax_prompt_editor.setFont(code_editor_font())
        self._prepare_dax_writer_editor(self.dax_prompt_editor, editable=True)
        self.dax_prompt_editor.setPlaceholderText(
            "Example: Create a measure that sums Sales[Amount] for the selected Calendar[Year]."
        )
        self.dax_prompt_editor.textChanged.connect(self._on_dax_prompt_changed)
        dax_writer_layout.addWidget(self.dax_prompt_editor)

        self.dax_count_label = QLabel("Tables mentioned: 0  Columns mentioned: 0")
        self.dax_count_label.setStyleSheet("color: #888888;")
        dax_writer_layout.addWidget(self.dax_count_label)

        dax_writer_layout.addWidget(QLabel("Generated DAX Measure:"))

        self.dax_output_editor = CodeEditor()
        self.dax_output_editor.setFont(code_editor_font())
        self._prepare_dax_writer_editor(self.dax_output_editor, editable=False)
        dax_writer_layout.addWidget(self.dax_output_editor)

        self._dax_prompt_shortcuts = [
            QShortcut(QKeySequence("Ctrl+Return"), self.dax_prompt_editor),
            QShortcut(QKeySequence("Ctrl+Enter"), self.dax_prompt_editor),
        ]
        for shortcut in self._dax_prompt_shortcuts:
            shortcut.activated.connect(self.generate_dax_measure)

        self.dax_writer_container.hide()
        right_layout.addWidget(self.dax_writer_container)

        self.query_editor = CodeEditor(language="m")
        self.query_editor.setFont(code_editor_font())
        self.query_editor.setEnabled(False)
        try:
            space_w = self.query_editor.fontMetrics().horizontalAdvance(" ")
            self.query_editor.setTabStopDistance(space_w * 4)
        except Exception:
            pass
        self.query_editor.textChanged.connect(self.on_query_text_changed)
        right_layout.addWidget(self.query_editor)

        self.hotkey_hint_right = QLabel(
            "Ctrl+K+C: Comment   |   Ctrl+K+U: Uncomment   |   Alt+Up: Move line up   |   Alt+Down: Move line down"
        )
        self.hotkey_hint_right.setStyleSheet("color: #666666; font-size: 10px;")
        self.hotkey_hint_right.setWordWrap(True)
        right_layout.addWidget(self.hotkey_hint_right)

        self.measure_form_widget = QWidget()
        form_layout = QFormLayout(self.measure_form_widget)
        form_layout.setContentsMargins(0, 0, 0, 0)
        form_layout.setSpacing(6)

        # Hotkey hints - left
        hotkey_hint_left = QLabel("F2: Rename\nDelete folder: Delete\nCtrl+N: New folder\nSpace: Set as default\nAlt+Up: Move up\nAlt+Down: Move down")
        hotkey_hint_left.setStyleSheet("color: #666666; font-size: 10px;")
        hotkey_hint_left.setWordWrap(True)
        left_layout.addWidget(hotkey_hint_left)

        self.import_mode_combo = QComboBox()
        self.import_mode_combo.setEnabled(False)
        self.import_mode_combo.addItem("Import", "import")
        self.import_mode_combo.addItem("DirectQuery", "directquery")
        self.import_mode_combo.currentIndexChanged.connect(self.on_import_mode_changed)
        form_layout.addRow("Data load mode", self.import_mode_combo)

        right_layout.addWidget(self.measure_form_widget)

        splitter.addWidget(right_widget)
        splitter.setStretchFactor(0, 1)
        splitter.setStretchFactor(1, 2)
        splitter.setSizes([300, 600])

        main_layout.addWidget(splitter)
        self.setup_shortcuts()

    def setup_shortcuts(self):
        self.save_shortcut = QShortcut(QKeySequence("Ctrl+S"), self)
        self.save_shortcut.setContext(Qt.ShortcutContext.WidgetShortcut)
        self.save_shortcut.activated.connect(self.save_changes)

        self.new_folder_shortcut = QShortcut(QKeySequence("Ctrl+N"), self.table_tree)
        self.new_folder_shortcut.setContext(Qt.ShortcutContext.WidgetShortcut)
        self.new_folder_shortcut.activated.connect(self._shortcut_create_folder)

        self.move_up_shortcut = QShortcut(QKeySequence("Alt+Up"), self.table_tree)
        self.move_up_shortcut.setContext(Qt.ShortcutContext.WidgetShortcut)
        self.move_up_shortcut.activated.connect(lambda: self.move_selected_items(-1))

        self.move_down_shortcut = QShortcut(QKeySequence("Alt+Down"), self.table_tree)
        self.move_down_shortcut.setContext(Qt.ShortcutContext.WidgetShortcut)
        self.move_down_shortcut.activated.connect(lambda: self.move_selected_items(1))

        self.delete_folder_shortcut = QShortcut(QKeySequence(Qt.Key.Key_Delete), self.table_tree)
        self.delete_folder_shortcut.setContext(Qt.ShortcutContext.WidgetShortcut)
        self.delete_folder_shortcut.activated.connect(self.delete_selected_items)

    def _set_dirty(self, value: bool):
        self.is_dirty = value
        if self.save_button:
            self.save_button.setEnabled(value)

    def mark_dirty(self):
        if self._loading_data:
            return
        self._set_dirty(True)

    def _normalize_group_path(self, raw: Optional[str]) -> Optional[str]:
        if not raw:
            return None
        text = raw.strip()
        if (text.startswith("'") and text.endswith("'")) or (text.startswith('"') and text.endswith('"')):
            text = text[1:-1]
        text = re.sub(r"[\\/]+", "/", text)
        parts = [part.strip() for part in text.split("/") if part.strip()]
        return "/".join(parts) if parts else None

    def _expand_group_path(self, path: str) -> List[str]:
        parts = [p for p in path.split("/") if p]
        expanded: List[str] = []
        for idx in range(len(parts)):
            expanded.append("/".join(parts[: idx + 1]))
        return expanded

    def _dax_table_identifiers(self, table_name: str) -> List[str]:
        if not table_name:
            return []
        name = table_name.strip()
        if not name:
            return []
        identifiers = [name]
        if not re.match(r"^[A-Za-z_][A-Za-z0-9_]*$", name):
            escaped = name.replace("'", "''")
            identifiers.append(f"'{escaped}'")
        return identifiers

    def _dax_column_identifiers(self, table_name: str, column_name: str) -> List[str]:
        if not column_name:
            return []
        column = column_name.replace("]", "]]").strip()
        if not column:
            return []
        identifiers = [f"[{column}]"]
        for table_identifier in self._dax_table_identifiers(table_name):
            identifiers.append(f"{table_identifier}[{column}]")
        return identifiers

    def _dax_measure_identifiers(self, table_name: str, measure_name: str) -> List[str]:
        if not measure_name:
            return []
        name = measure_name.replace("]", "]]").strip()
        if not name:
            return []
        identifiers = [f"[{name}]"]
        for table_identifier in self._dax_table_identifiers(table_name):
            identifiers.append(f"{table_identifier}[{name}]")
        return identifiers

    def _update_dax_model_identifiers(self):
        table_terms: List[str] = []
        column_terms: List[str] = []

        try:
            table_names = sorted(self.tables_data.keys(), key=str.casefold)
        except Exception:
            table_names = list(self.tables_data.keys())

        for table_name in table_names:
            table_terms.extend(self._dax_table_identifiers(table_name))
            info = self.tables_data.get(table_name, {}) or {}
            for column in info.get("columns", []) or []:
                column_terms.extend(self._dax_column_identifiers(table_name, column))
            for measure in info.get("measures", []) or []:
                column_terms.extend(self._dax_measure_identifiers(table_name, measure.get("name") or ""))

        set_dax_model_identifiers(table_terms, column_terms)

    def _semantic_model_root(self) -> Optional[Path]:
        if self.project and getattr(self.project, "semantic_model_dir", None):
            try:
                return Path(self.project.semantic_model_dir)
            except TypeError:
                return Path(str(self.project.semantic_model_dir))
        if not self.pbip_file:
            return None
        pbip_path = Path(self.pbip_file).expanduser()
        if pbip_path.suffix.lower() != ".pbip":
            return None
        return pbip_path.parent / f"{pbip_path.stem}.SemanticModel"

    def _gather_tmdl_texts(self, root: Optional[Path], *, recursive: bool = True) -> List[str]:
        texts: List[str] = []
        if not root:
            return texts
        try:
            base = Path(root)
        except TypeError:
            base = Path(str(root))
        if not base.exists():
            return texts
        iterator = base.rglob("*.tmdl") if recursive else base.glob("*.tmdl")
        for file_path in iterator:
            if not file_path.is_file():
                continue
            try:
                texts.append(file_path.read_text(encoding="utf-8"))
            except UnicodeDecodeError:
                try:
                    texts.append(file_path.read_text(encoding="utf-8-sig"))
                except Exception:
                    continue
            except OSError:
                continue
        return texts

    @staticmethod
    def _is_simple_identifier(name: str) -> bool:
        return bool(name) and bool(re.match(r"^[A-Za-z_][A-Za-z0-9_]*$", name))

    @staticmethod
    def _quote_identifier(name: str) -> str:
        escaped = name.replace("'", "''") if name else ""
        return f"'{escaped}'"

    @staticmethod
    def _bracket_identifier(name: str) -> str:
        return name.replace("]", "]]")

    def _identifier_regex_forms(self, name: str) -> List[str]:
        if not name:
            return []
        forms = [self._quote_identifier(name.strip())]
        if self._is_simple_identifier(name.strip()):
            forms.append(name.strip())
        unique = list(dict.fromkeys(forms))
        return [re.escape(form) for form in unique if form]

    @staticmethod
    def _pattern_occurrences(pattern: re.Pattern[str], texts: List[str]) -> int:
        total = 0
        for text in texts:
            total += len(pattern.findall(text))
        return total

    def _build_table_reference_pattern(self, table_name: str) -> Optional[re.Pattern[str]]:
        tokens = self._identifier_regex_forms(table_name)
        if not tokens:
            return None
        group = "|".join(tokens)
        return re.compile(rf"(?<![\w])(?:{group})(?![\w])", re.IGNORECASE)

    def _build_column_reference_pattern(self, table_name: str, column_name: str) -> Optional[re.Pattern[str]]:
        del table_name  # unused, kept for signature compatibility
        if not column_name:
            return None
        column_token = re.escape(self._bracket_identifier(column_name.strip()))
        if not column_token:
            return None
        return re.compile(rf"(?<![\w\]])\[\s*{column_token}\s*\](?![\w])", re.IGNORECASE)

    def _build_measure_reference_patterns(
        self,
        table_name: str,
        measure_name: str,
    ) -> Tuple[Optional[re.Pattern[str]], Optional[re.Pattern[str]]]:
        del table_name  # unused, kept for signature compatibility
        if not measure_name:
            return None, None
        measure_token = re.escape(self._bracket_identifier(measure_name.strip()))
        if not measure_token:
            return None, None
        bracket_pattern = re.compile(rf"(?<![\w\]])\[\s*{measure_token}\s*\](?![\w])", re.IGNORECASE)
        return bracket_pattern, None

    def _refresh_usage_statistics(self) -> None:
        self._table_usage_flags.clear()
        self._column_usage_flags.clear()
        self._measure_usage_flags.clear()

        root = self._semantic_model_root()
        if root is None:
            return

        definition_dir = root / "definition"
        tables_dir = definition_dir / "tables"
        pages_dir = definition_dir / "pages"

        table_texts = self._gather_tmdl_texts(tables_dir, recursive=False)
        pages_texts = self._gather_tmdl_texts(pages_dir, recursive=True)
        combined_texts = table_texts + pages_texts

        if not self.tables_data:
            return

        for table_name, data in self.tables_data.items():
            table_pattern = self._build_table_reference_pattern(table_name)
            if table_pattern is not None and combined_texts:
                count = self._pattern_occurrences(table_pattern, combined_texts)
                self._table_usage_flags[table_name] = count > 1
            else:
                self._table_usage_flags[table_name] = False

            columns = data.get("columns") or []
            for column_name in columns:
                column_pattern = self._build_column_reference_pattern(table_name, column_name)
                if column_pattern is not None and combined_texts:
                    occurrences = self._pattern_occurrences(column_pattern, combined_texts)
                    self._column_usage_flags[(table_name, column_name)] = occurrences >= 1
                else:
                    self._column_usage_flags[(table_name, column_name)] = False

            measures = data.get("measures") or []
            for measure in measures:
                measure_id = measure.get("id")
                measure_name = measure.get("name")
                if not measure_id or not measure_name:
                    continue
                with_table_pattern, standalone_pattern = self._build_measure_reference_patterns(
                    table_name,
                    measure_name,
                )
                count = 0
                if with_table_pattern is not None and combined_texts:
                    count += self._pattern_occurrences(with_table_pattern, combined_texts)
                if standalone_pattern is not None and combined_texts:
                    count += self._pattern_occurrences(standalone_pattern, combined_texts)
                self._measure_usage_flags[measure_id] = count > 1

    def _prepare_dax_writer_editor(self, editor: CodeEditor, *, editable: bool) -> None:
        if editable and self.dax_prompt_highlighter:
            self.dax_prompt_highlighter.setDocument(None)
            self.dax_prompt_highlighter = None
        if not editable and self.dax_output_highlighter:
            self.dax_output_highlighter.setDocument(None)
            self.dax_output_highlighter = None

        existing = getattr(editor, "_highlighter", None)
        if existing:
            try:
                existing.setDocument(None)
            except Exception:
                pass
            editor._highlighter = None

        try:
            editor.set_language("dax", force=True, enable_highlighter=False)
        except Exception:
            pass

        try:
            space_w = editor.fontMetrics().horizontalAdvance(" ")
            editor.setTabStopDistance(space_w * 4)
        except Exception:
            pass

        try:
            editor.setLineWrapMode(CodeEditor.LineWrapMode.NoWrap)
        except Exception:
            pass
        editor.setReadOnly(not editable)

        if editable:
            self.dax_prompt_highlighter = TableColumnHighlighter(editor.document())
        else:
            self.dax_output_highlighter = DAXTableColumnHighlighter(editor.document())

        self._update_dax_writer_highlighters()

    def _update_dax_writer_highlighters(self) -> None:
        def _deduplicate(patterns: List[re.Pattern]) -> List[re.Pattern]:
            unique: List[re.Pattern] = []
            seen: Set[Tuple[str, int]] = set()
            for pattern in patterns:
                key = (pattern.pattern, pattern.flags)
                if key in seen:
                    continue
                seen.add(key)
                unique.append(pattern)
            return unique

        table_regexes = _deduplicate([pattern for pattern, _ in self.dax_table_patterns])
        column_regexes = _deduplicate(self._dax_column_highlight_patterns)

        if self.dax_prompt_highlighter:
            self.dax_prompt_highlighter.update_patterns(table_regexes, column_regexes)
        if self.dax_output_highlighter:
            self.dax_output_highlighter.update_patterns(table_regexes, column_regexes)

    def _refresh_dax_writer_metadata(self) -> None:
        tables: Dict[str, List[str]] = {}
        for name, info in (self.tables_data or {}).items():
            columns = [col for col in (info.get("columns") or []) if isinstance(col, str) and col.strip()]
            if columns:
                tables[name] = sorted(dict.fromkeys(columns), key=str.casefold)

        self.dax_table_patterns = []
        self.dax_column_patterns = []
        self._dax_column_highlight_patterns = []

        if not tables:
            self._update_dax_writer_highlighters()
            self._on_dax_prompt_changed()
            return

        for table in tables:
            for form in self._table_autocomplete_forms(table):
                if not form:
                    continue
                pattern = re.compile(rf"(?<![\w\]]){re.escape(form)}(?![\w\[])", re.IGNORECASE)
                self.dax_table_patterns.append((pattern, table))

        for table, columns in tables.items():
            for column in columns:
                escaped_col = column.replace("]", "]]")
                bracket_pattern = re.compile(
                    rf"\[\s*{re.escape(escaped_col)}\s*\]",
                    re.IGNORECASE,
                )
                bound_patterns: List[re.Pattern] = []
                for table_form in self._table_autocomplete_forms(table):
                    if not table_form:
                        continue
                    bound_patterns.append(
                        re.compile(
                            rf"{re.escape(table_form)}\s*\[\s*{re.escape(escaped_col)}\s*\]",
                            re.IGNORECASE,
                        )
                    )
                self.dax_column_patterns.append((bound_patterns, bracket_pattern, (table, column)))
                self._dax_column_highlight_patterns.extend(bound_patterns)
                self._dax_column_highlight_patterns.append(bracket_pattern)

        self._update_dax_writer_highlighters()
        self._on_dax_prompt_changed()

    @staticmethod
    def _table_autocomplete_forms(table: str) -> List[str]:
        if not table:
            return []
        forms = [table]
        if re.search(r"[^\w]", table):
            escaped = table.replace("'", "''")
            forms.append(f"'{escaped}'")
        return forms

    def _on_dax_prompt_changed(self) -> None:
        if not self.dax_prompt_editor or not self.dax_count_label:
            return
        text = self.dax_prompt_editor.toPlainText()
        tables_found, columns_found = self._count_dax_mentions(text)
        self.dax_count_label.setText(
            f"Tables mentioned: {len(tables_found)}  Columns mentioned: {len(columns_found)}"
        )

    def _count_dax_mentions(self, text: str) -> Tuple[Set[str], Set[Tuple[str, str]]]:
        mentioned_tables: Set[str] = set()
        mentioned_columns: Set[Tuple[str, str]] = set()

        for pattern, table in self.dax_table_patterns:
            if pattern.search(text):
                mentioned_tables.add(table)

        for bound_patterns, bracket_pattern, (table, column) in self.dax_column_patterns:
            if any(p.search(text) for p in bound_patterns) or bracket_pattern.search(text):
                mentioned_columns.add((table, column))
                mentioned_tables.add(table)

        return mentioned_tables, mentioned_columns

    def generate_dax_measure(self) -> None:
        if self._dax_busy:
            return
        if not self.dax_prompt_editor or not self.dax_output_editor or not self.dax_generate_button:
            return
        if not self.project:
            QMessageBox.warning(self, "No PBIP", "Select a PBIP file to continue.")
            return

        text = self.dax_prompt_editor.toPlainText().strip()
        if not text:
            QMessageBox.information(self, "Empty Prompt", "Please describe the measure you need.")
            return

        if not self.dax_table_patterns:
            self._refresh_dax_writer_metadata()

        tables_found, columns_found = self._count_dax_mentions(text)
        if not tables_found:
            QMessageBox.warning(
                self,
                "Missing Table",
                "Reference at least one table from the model (e.g., Sales or Sales[Amount]).",
            )
            return
        if not columns_found:
            QMessageBox.warning(
                self,
                "Missing Column",
                "Reference at least one column using Table[Column] syntax in your description.",
            )
            return

        final_prompt = (
            "Follow this instructions to the letter.\n"
            "Give your answer with a DAX code for Power BI and absolutely nothing else.\n"
            "The DAX code should be properly formatted and indented to be readable and easy to copy and paste.\n"
            "The DAX code should contain simple comments. The very first line of the code must not be a comment.\n"
            "Build a DAX measure that:\n"
            f"{text}"
        )

        self._set_dax_busy(True, "Generating DAX measure...")
        worker = ChatRequestWorker(self.dax_api_client.generate, final_prompt)
        worker.signals.finished.connect(self._on_dax_generation_success)
        worker.signals.error.connect(self._on_dax_generation_error)
        self.dax_thread_pool.start(worker)

    def _on_dax_generation_success(self, response: str) -> None:
        self._set_dax_busy(False, "Generation complete.")
        if self.dax_output_editor:
            self.dax_output_editor.setPlainText(response)

    def _on_dax_generation_error(self, message: str) -> None:
        self._set_dax_busy(False, f"Generation failed: {message}")
        QMessageBox.critical(
            self,
            "Generation Failed",
            f"ChatGPT request failed:\n{message}",
        )

    def _set_dax_busy(self, busy: bool, status: str = "") -> None:
        self._dax_busy = busy
        if self.dax_generate_button:
            self.dax_generate_button.setEnabled(not busy)
        if self.dax_copy_button:
            self.dax_copy_button.setEnabled(not busy)
        if self.dax_prompt_editor:
            self.dax_prompt_editor.setEnabled(not busy)
        if self.dax_status_label is not None:
            self.dax_status_label.setText(status)

    def copy_dax_output(self) -> None:
        if not self.dax_output_editor:
            return
        text = self.dax_output_editor.toPlainText().strip()
        if not text:
            return
        clipboard = QApplication.clipboard()
        clipboard.setText(text, mode=clipboard.Mode.Clipboard)
        if self.dax_status_label is not None:
            self.dax_status_label.setText("Copied generated DAX to clipboard.")

    def _show_measure_editor_widgets(self) -> None:
        if self.dax_controls_widget:
            self.dax_controls_widget.hide()
        if self.dax_writer_container:
            self.dax_writer_container.hide()
        if self.query_editor:
            self.query_editor.show()
        if self.hotkey_hint_right:
            self.hotkey_hint_right.show()
        if self.measure_form_widget:
            self.measure_form_widget.show()

    def _show_dax_writer_widgets(self) -> None:
        if self.query_editor:
            self.query_editor.hide()
        if self.hotkey_hint_right:
            self.hotkey_hint_right.hide()
        if self.measure_form_widget:
            self.measure_form_widget.hide()
        if self.dax_controls_widget:
            self.dax_controls_widget.show()
        if self.dax_writer_container:
            self.dax_writer_container.show()

    def _apply_usage_state(self, item: QTreeWidgetItem, used: Optional[bool]) -> None:
        column = 1
        if item.columnCount() <= column:
            return
        if used is None:
            item.setText(column, "")
            item.setForeground(column, QBrush())
            return
        if used:
            item.setText(column, "Used")
            item.setForeground(column, QBrush(QColor("#2e7d32")))
        else:
            item.setText(column, "Unused")
            item.setForeground(column, QBrush(QColor("#c62828")))

    def _shortcut_create_folder(self):
        current = self.table_tree.currentItem()
        parent_folder = None
        if current:
            current_type = current.data(0, self.TYPE_ROLE)
            if current_type == self.ITEM_FOLDER:
                parent_folder = current
            else:
                ancestor = current.parent()
                if ancestor and ancestor.data(0, self.TYPE_ROLE) == self.ITEM_FOLDER:
                    parent_folder = ancestor
        self.create_new_folder(parent_folder)

    def move_selected_items(self, direction: int):
        if direction not in (-1, 1):
            return
        selectable_types = {self.ITEM_FOLDER, self.ITEM_TABLE}
        selected = [item for item in self.table_tree.selectedItems() if item.data(0, self.TYPE_ROLE) in selectable_types]
        if not selected:
            return

        parent_groups: dict[int, list[QTreeWidgetItem]] = defaultdict(list)
        parent_refs: dict[int, Optional[QTreeWidgetItem]] = {}
        for item in selected:
            parent = item.parent()
            key = id(parent) if parent is not None else -1
            parent_refs[key] = parent
            parent_groups[key].append(item)

        moved = False
        self._ignore_tree_changes = True
        try:
            for key, items in parent_groups.items():
                parent = parent_refs.get(key)
                items.sort(key=lambda itm: self._get_parent_and_index(itm)[1], reverse=(direction > 0))
                for item in items:
                    parent_ref, index = self._get_parent_and_index(item)
                    if index < 0:
                        continue
                    sibling_count = parent_ref.childCount() if parent_ref else self.table_tree.topLevelItemCount()
                    if direction == -1:
                        if index == 0:
                            continue
                        self._take_item(parent_ref, index)
                        self._insert_item(parent_ref, index - 1, item)
                        moved = True
                    else:
                        if index >= sibling_count - 1:
                            continue
                        self._take_item(parent_ref, index)
                        self._insert_item(parent_ref, index + 1, item)
                        moved = True
        finally:
            self._ignore_tree_changes = False

        if moved:
            self.table_tree.clearSelection()
            for item in selected:
                item.setSelected(True)
            if selected:
                self.table_tree.setCurrentItem(selected[0])
                self.table_tree.scrollToItem(selected[0], QAbstractItemView.ScrollHint.EnsureVisible)
            self.on_tree_structure_changed()

    def delete_selected_items(self):
        selected = self.table_tree.selectedItems()
        if not selected:
            return

        measure_items = []
        measure_folders = []
        query_folders = []
        for item in selected:
            item_type = item.data(0, self.TYPE_ROLE)
            if item_type == self.ITEM_MEASURE:
                measure_items.append(item)
            elif item_type == self.ITEM_MEASURE_FOLDER:
                measure_folders.append(item)
            elif item_type == self.ITEM_FOLDER and item.data(0, self.KEY_ROLE) is not None:
                query_folders.append(item)

        for measure in measure_items:
            self.delete_measure_item(measure)

        measure_folders.sort(key=self._item_depth, reverse=True)
        for folder in measure_folders:
            self.delete_measure_folder(folder)

        query_folders.sort(key=self._item_depth, reverse=True)
        for folder in query_folders:
            self.delete_folder(folder)

    def _get_parent_and_index(self, item: QTreeWidgetItem):
        parent = item.parent()
        if parent:
            return parent, parent.indexOfChild(item)
        return None, self.table_tree.indexOfTopLevelItem(item)

    def _take_item(self, parent: Optional[QTreeWidgetItem], index: int):
        if parent is None:
            return self.table_tree.takeTopLevelItem(index)
        return parent.takeChild(index)

    def _insert_item(self, parent: Optional[QTreeWidgetItem], index: int, item: QTreeWidgetItem):
        if parent is None:
            self.table_tree.insertTopLevelItem(index, item)
        else:
            parent.insertChild(index, item)

    def _item_depth(self, item: QTreeWidgetItem) -> int:
        depth = 0
        current = item.parent()
        while current is not None:
            depth += 1
            current = current.parent()
        return depth

    def choose_pbip_file(self):
        """Prompt the user to pick a PBIP file and load it."""
        chosen, _ = QFileDialog.getOpenFileName(
            self,
            "Select PBIP file",
            "",
            "Power BI Project (*.pbip)",
        )
        if not chosen:
            return
        try:
            self.project = load_pbip_project(chosen, force_reload=True)
        except Exception as exc:
            QMessageBox.critical(self, "Failed to Load", f"Could not load PBIP project:\n{exc}")
            return
        self.pbip_file = str(self.project.pbip_path)
        self.load_tables()

    def refresh_tables(self):
        """Reload tables using the current PBIP file."""
        if self.project:
            self.project.reload_tables()
        elif self.pbip_file:
            try:
                self.project = load_pbip_project(self.pbip_file, force_reload=True)
                self.pbip_file = str(self.project.pbip_path)
            except Exception as exc:
                QMessageBox.warning(self, "Reload Failed", f"Could not reload PBIP project:\n{exc}")
                return
        self.load_tables()

    def load_tables(self):
        """Extract tables, columns, and related metadata from the PBIP definition."""
        self.clear_details()

        if not self.project and self.pbip_file:
            try:
                self.project = load_pbip_project(self.pbip_file)
                self.pbip_file = str(self.project.pbip_path)
            except Exception as exc:
                QMessageBox.warning(self, "Load Failed", f"Unable to load PBIP project:\n{exc}")
                return

        if not self.project:
            QMessageBox.information(self, "No PBIP", "Select a PBIP file to load Power Query metadata.")
            return

        self._loading_data = True
        try:
            metadata = self.project.get_power_query_metadata()
            if metadata.error:
                QMessageBox.warning(
                    self,
                    "Metadata Error",
                    f"Could not load Power Query metadata:\n{metadata.error}",
                )
                self.tables_data = {}
                self.query_order = []
                self.query_groups = {}
                self.populate_tree()
                self._refresh_dax_writer_metadata()
                self._set_dirty(False)
                return

            self.tables_data = metadata.tables
            self.query_order = metadata.query_order
            self.query_groups = metadata.query_groups

            self._update_dax_model_identifiers()
            self._refresh_dax_writer_metadata()
            self._refresh_usage_statistics()
            self.populate_tree()
            self._set_dirty(False)
        finally:
            self._loading_data = False

    def save_changes(self):
        """Persist current tree layout back to the PBIP model files."""
        if not self.pbip_file:
            QMessageBox.information(self, "No PBIP", "Select a PBIP file to save changes.")
            return

        pbip_path = os.path.abspath(self.pbip_file)
        if not os.path.isfile(pbip_path):
            QMessageBox.warning(self, "Missing PBIP", f"PBIP file not found:\n{pbip_path}")
            return

        semantic_root = os.path.splitext(pbip_path)[0] + ".SemanticModel"
        model_tmdl = os.path.join(semantic_root, "definition", "model.tmdl")
        tables_dir = os.path.join(semantic_root, "definition", "tables")

        if not os.path.isfile(model_tmdl) or not os.path.isdir(tables_dir):
            QMessageBox.warning(
                self,
                "Missing Metadata",
                "Could not locate model.tmdl or tables metadata for the selected PBIP file.",
            )
            return

        try:
            folder_paths, table_order, table_groups = self._collect_tree_layout()
            self.query_groups = {path: idx for idx, path in enumerate(folder_paths)}
            self.query_order = table_order

            self._update_model_tmdl(model_tmdl, folder_paths, table_order)

            for table_name, group_path in table_groups.items():
                info = self.tables_data.get(table_name)
                if not info:
                    continue
                table_path = os.path.join(tables_dir, f"{table_name}.tmdl")
                if not os.path.isfile(table_path):
                    continue
                if (info.get("table_type") or "").lower() != "calculated":
                    self._update_table_definition(table_path, group_path)
                self._write_table_measures(table_name, table_path)

            if self.project:
                self.project.update_power_query_metadata(self.tables_data, self.query_order, self.query_groups)
            self._set_dirty(False)
            QMessageBox.information(self, "Success", "Changes saved successfully!")
        except Exception as exc:
            QMessageBox.critical(self, "Error", f"Failed to save changes:\n{exc}")

    def populate_tree(self):
        """Populate the tree view with tables and their columns."""
        self._ignore_item_change = True
        self._ignore_tree_changes = True
        self.table_tree.blockSignals(True)
        self.table_tree.clear()

        if not self.tables_data:
            self.table_tree.blockSignals(False)
            self._ignore_tree_changes = False
            self._ignore_item_change = False
            return

        if self.query_order:
            sorted_tables = [name for name in self.query_order if name in self.tables_data]
            extras = sorted(name for name in self.tables_data if name not in self.query_order)
            sorted_tables.extend(extras)
        else:
            sorted_tables = sorted(self.tables_data.keys(), key=str.casefold)

        group_map: dict[Optional[str], list[str]] = {}
        all_group_paths: set[str] = set()
        path_orders: Dict[str, float] = {}

        table_order = {name.lower(): idx for idx, name in enumerate(self.query_order)}

        def table_sort_key(name: str):
            return (table_order.get(name.lower(), len(table_order)), name.lower())

        for table_name in sorted_tables:
            info = self.tables_data.get(table_name, {})
            group_key = info.get("query_group") or None
            if group_key:
                group_map.setdefault(group_key, []).append(table_name)
                for expanded in self._expand_group_path(group_key):
                    all_group_paths.add(expanded)
                    if expanded not in path_orders:
                        path_orders[expanded] = float("inf")
            else:
                group_map.setdefault(None, []).append(table_name)

        for path, order in self.query_groups.items():
            if not path:
                continue
            for expanded in self._expand_group_path(path):
                all_group_paths.add(expanded)
                if expanded in path_orders:
                    path_orders[expanded] = min(path_orders[expanded], order)
                else:
                    path_orders[expanded] = order

        folder_items: Dict[str, QTreeWidgetItem] = {}

        def ensure_folder(path: Optional[str]) -> Optional[QTreeWidgetItem]:
            if not path:
                return None
            if path in folder_items:
                return folder_items[path]
            parts = [p for p in path.split("/") if p]
            if not parts:
                return None
            parent_path = "/".join(parts[:-1]) if len(parts) > 1 else None
            parent_item = ensure_folder(parent_path) if parent_path else None
            display_name = parts[-1]
            item = self._create_folder_item(path, display_name)
            folder_items[path] = item
            if parent_item:
                parent_item.addChild(item)
            return item

        top_level_paths = {path.split("/")[0] for path in all_group_paths if path}
        top_level_orders: Dict[str, float] = {}
        for path, order in path_orders.items():
            if not path:
                continue
            top_segment = path.split("/")[0]
            current = top_level_orders.get(top_segment)
            if current is None or order < current:
                top_level_orders[top_segment] = order

        ordered_top_level = sorted(
            top_level_paths,
            key=lambda name: (top_level_orders.get(name, float("inf")), name.lower()),
        )

        for path in sorted(all_group_paths, key=lambda p: (path_orders.get(p, float("inf")), p.lower())):
            ensure_folder(path)

        for top_path in ordered_top_level:
            folder_item = ensure_folder(top_path)
            if folder_item:
                self.table_tree.addTopLevelItem(folder_item)

        for group_key, tables in group_map.items():
            if group_key is None:
                continue
            folder_item = ensure_folder(group_key)
            if not folder_item:
                continue
            tables_in_group = sorted(tables, key=table_sort_key)
            for table_name in tables_in_group:
                table_item = self._create_table_item(table_name, group_key)
                folder_item.addChild(table_item)

        other_tables = sorted(group_map.get(None, []), key=table_sort_key)
        if other_tables:
            other_item = self.ensure_other_queries_folder()
            for table_name in other_tables:
                table_item = self._create_table_item(table_name, None)
                other_item.addChild(table_item)
        else:
            self.ensure_other_queries_folder()

        self.table_tree.blockSignals(False)
        self.ensure_other_queries_last()
        self.table_tree.collapseAll()
        self.table_tree.clearSelection()
        self.display_table_details(None)
        self._ignore_tree_changes = False
        self._ignore_item_change = False
        self.on_tree_structure_changed()

    def on_tree_selection_changed(self):
        """Update detail pane when the selection changes."""
        current_item = self.table_tree.currentItem()
        if not current_item:
            self.display_table_details(None)
            return

        item_type = current_item.data(0, self.TYPE_ROLE)
        if item_type == self.ITEM_FOLDER:
            self.display_table_details(None)
            return

        if item_type == self.ITEM_MEASURE_FOLDER:
            self.display_table_details(None)
            return

        if item_type == self.ITEM_MEASURE:
            table_name = current_item.data(0, Qt.ItemDataRole.UserRole)
            measure_id = current_item.data(0, self.KEY_ROLE)
            self.display_measure_details(table_name, measure_id)
            return

        if item_type == self.ITEM_COLUMN:
            parent = current_item.parent()
            if parent:
                table_name = parent.data(0, Qt.ItemDataRole.UserRole)
                self.display_table_details(table_name)
            else:
                self.display_table_details(None)
            return

        table_name = current_item.data(0, Qt.ItemDataRole.UserRole)
        self.display_table_details(table_name)

    def display_table_details(self, table_name: Optional[str]):
        """Show M query, import mode, and query group for the selected table."""
        self.current_table = table_name
        self.current_measure_id = None
        self.measure_view_mode = "expression"
        self.ignore_editor_changes = True
        self._show_measure_editor_widgets()
        if self.dax_controls_widget:
            self.dax_controls_widget.hide()
        if self.dax_writer_container:
            self.dax_writer_container.hide()
        if self.dax_writer_button:
            self.dax_writer_button.setChecked(False)
        if self.dax_status_label is not None:
            self.dax_status_label.setText("")
        if hasattr(self, "measure_toggle_widget"):
            self.measure_toggle_widget.hide()
        if hasattr(self, "measure_expression_button"):
            self.measure_expression_button.setChecked(True)
        if hasattr(self, "measure_format_button"):
            self.measure_format_button.setChecked(False)

        if not table_name or table_name not in self.tables_data:
            self.query_label.setText("Query")
            self.query_editor.setPlainText("")
            self.query_editor.setEnabled(False)
            self.import_mode_combo.blockSignals(True)
            self.import_mode_combo.setCurrentIndex(0)
            self.import_mode_combo.blockSignals(False)
            self.import_mode_combo.setEnabled(False)
            self.ignore_editor_changes = False
            return

        table_info = self.tables_data[table_name]

        language = (table_info.get("code_language") or "").lower()
        table_type = (table_info.get("table_type") or "").lower()
        if not language:
            language = "dax" if table_type == "calculated" else "m"
        label_text = "Power Query (M)" if language == "m" else "Calculated Table (DAX)"
        self.query_label.setText(label_text)

        self.query_editor.setEnabled(True)
        if self.query_editor.language() != language:
            self.query_editor.set_language(language)
        self.query_editor.setPlainText(table_info.get("code_text") or "")

        mode_value = (table_info.get("import_mode") or "").lower()
        idx = self.import_mode_combo.findData(mode_value)
        if idx == -1:
            idx = 0
        self.import_mode_combo.blockSignals(True)
        self.import_mode_combo.setCurrentIndex(idx)
        self.import_mode_combo.blockSignals(False)
        self.import_mode_combo.setEnabled(True)

        self.ignore_editor_changes = False

    def display_measure_details(self, table_name: Optional[str], measure_id: Optional[str]):
        """Show measure expression or format string in the editor."""
        self.current_table = table_name
        self.current_measure_id = measure_id
        self.ignore_editor_changes = True
        self._show_measure_editor_widgets()
        if self.dax_controls_widget:
            self.dax_controls_widget.hide()
        if self.dax_writer_container:
            self.dax_writer_container.hide()
        if self.dax_writer_button:
            self.dax_writer_button.setChecked(False)
        if self.dax_status_label is not None:
            self.dax_status_label.setText("")

        measure = self._get_measure_data(table_name, measure_id)
        if not measure:
            self.query_label.setText("Measure")
            self.query_editor.clear()
            self.query_editor.setEnabled(False)
            if hasattr(self, "measure_toggle_widget"):
                self.measure_toggle_widget.hide()
            if self.dax_controls_widget:
                self.dax_controls_widget.hide()
            if self.dax_writer_container:
                self.dax_writer_container.hide()
            if self.dax_writer_button:
                self.dax_writer_button.setChecked(False)
            self.import_mode_combo.blockSignals(True)
            self.import_mode_combo.setCurrentIndex(0)
            self.import_mode_combo.blockSignals(False)
            self.import_mode_combo.setEnabled(False)
            self.ignore_editor_changes = False
            return

        if hasattr(self, "measure_toggle_widget"):
            self.measure_toggle_widget.show()
        if hasattr(self, "measure_expression_button"):
            self.measure_expression_button.setChecked(True)
        if hasattr(self, "measure_format_button"):
            self.measure_format_button.setChecked(False)
        if self.dax_writer_button:
            self.dax_writer_button.setChecked(False)

        self.measure_view_mode = "expression"
        self.query_label.setText(f"Measure: {measure.get('name', '')}")
        self.query_editor.setEnabled(True)
        if self.query_editor.language() != "dax":
            self.query_editor.set_language("dax")
        self.query_editor.setPlainText(measure.get("expression") or "")

        self.import_mode_combo.blockSignals(True)
        self.import_mode_combo.setCurrentIndex(0)
        self.import_mode_combo.blockSignals(False)
        self.import_mode_combo.setEnabled(False)

        self.ignore_editor_changes = False

    def show_measure_expression(self):
        if self.ignore_editor_changes:
            return
        measure = self._get_measure_data(self.current_table, self.current_measure_id)
        if not measure:
            return
        self.measure_view_mode = "expression"
        self._show_measure_editor_widgets()
        if self.measure_expression_button:
            self.measure_expression_button.setChecked(True)
        if self.measure_format_button:
            self.measure_format_button.setChecked(False)
        if self.dax_writer_button:
            self.dax_writer_button.setChecked(False)
        if self.dax_status_label is not None:
            self.dax_status_label.setText("")
        self.ignore_editor_changes = True
        if self.query_editor.language() != "dax":
            self.query_editor.set_language("dax")
        self.query_editor.setPlainText(measure.get("expression") or "")
        self.query_label.setText(f"Measure: {measure.get('name', '')}")
        self.ignore_editor_changes = False

    def show_measure_format(self):
        if self.ignore_editor_changes:
            return
        measure = self._get_measure_data(self.current_table, self.current_measure_id)
        if not measure:
            return
        self.measure_view_mode = "format"
        self._show_measure_editor_widgets()
        if self.measure_expression_button:
            self.measure_expression_button.setChecked(False)
        if self.measure_format_button:
            self.measure_format_button.setChecked(True)
        if self.dax_writer_button:
            self.dax_writer_button.setChecked(False)
        if self.dax_status_label is not None:
            self.dax_status_label.setText("")
        self.ignore_editor_changes = True
        if self.query_editor.language() != "dax":
            self.query_editor.set_language("dax")
        format_text = measure.get("format_string") or ""
        self.query_editor.setPlainText(format_text)
        self.query_label.setText(f"Format String: {measure.get('name', '')}")
        self.ignore_editor_changes = False

    def show_dax_writer(self):
        if self.ignore_editor_changes:
            return
        if self.measure_view_mode == "writer":
            return
        self.measure_view_mode = "writer"
        self._show_dax_writer_widgets()
        if self.measure_expression_button:
            self.measure_expression_button.setChecked(False)
        if self.measure_format_button:
            self.measure_format_button.setChecked(False)
        if self.dax_writer_button:
            self.dax_writer_button.setChecked(True)
        self.query_label.setText("DAX Writer")
        if self.dax_status_label is not None and not self._dax_busy:
            self.dax_status_label.setText("")
        if not self.dax_table_patterns:
            self._refresh_dax_writer_metadata()
        self._on_dax_prompt_changed()
        if self.dax_prompt_editor:
            self.dax_prompt_editor.setFocus()

    def on_query_text_changed(self):
        """Persist query/expression changes to in-memory state."""
        if self.ignore_editor_changes:
            return
        if self.measure_view_mode == "writer":
            return
        if self.current_measure_id:
            measure = self._get_measure_data(self.current_table, self.current_measure_id)
            if not measure:
                return
            text = self.query_editor.toPlainText()
            if self.measure_view_mode == "format":
                measure["format_string"] = text if text.strip() else None
            else:
                measure["expression"] = text
            self.mark_dirty()
            return
        if not self.current_table:
            return
        self.tables_data[self.current_table]["code_text"] = self.query_editor.toPlainText()
        self.tables_data[self.current_table]["code_language"] = (self.query_editor.language() or "").lower()
        self.mark_dirty()

    def on_import_mode_changed(self, index: int):
        """Persist import mode changes to in-memory state."""
        if self.ignore_editor_changes or not self.current_table:
            return
        value = self.import_mode_combo.itemData(index)
        self.tables_data[self.current_table]["import_mode"] = value if value else None
        self.mark_dirty()

    def clear_details(self):
        """Reset detail pane and selection state."""
        self._ignore_tree_changes = True
        self.table_tree.clear()
        self._ignore_tree_changes = False
        self.current_table = None
        self.current_measure_id = None
        self.measure_view_mode = "expression"
        self.ignore_editor_changes = True
        self._show_measure_editor_widgets()
        if self.dax_controls_widget:
            self.dax_controls_widget.hide()
        if self.dax_writer_container:
            self.dax_writer_container.hide()
        if self.dax_writer_button:
            self.dax_writer_button.setChecked(False)
        if self.dax_status_label is not None:
            self.dax_status_label.setText("")
        if self.measure_expression_button:
            self.measure_expression_button.setChecked(True)
        if self.measure_format_button:
            self.measure_format_button.setChecked(False)
        self.query_label.setText("Query")
        self.query_editor.clear()
        self.query_editor.setEnabled(False)
        self.import_mode_combo.blockSignals(True)
        self.import_mode_combo.setCurrentIndex(0)
        self.import_mode_combo.blockSignals(False)
        self.import_mode_combo.setEnabled(False)
        if hasattr(self, "measure_toggle_widget"):
            self.measure_toggle_widget.hide()
        self.ignore_editor_changes = False
        set_dax_model_identifiers([], [])

    def expand_all_groups(self):
        self.table_tree.expandAll()

    def collapse_all_groups(self):
        self.table_tree.collapseAll()

    def sort_folders_and_tables(self):
        if self.table_tree.topLevelItemCount() == 0:
            return
        self._ignore_tree_changes = True
        current_item = self.table_tree.currentItem()
        folders = []
        other_item = None
        while self.table_tree.topLevelItemCount():
            item = self.table_tree.takeTopLevelItem(0)
            key = item.data(0, self.KEY_ROLE)
            if key is None:
                other_item = item
            else:
                folders.append(item)
        folders.sort(key=lambda itm: itm.text(0).lower())
        for folder in folders:
            self.sort_tables_in_folder(folder)
            self.table_tree.addTopLevelItem(folder)
        if other_item:
            self.sort_tables_in_folder(other_item)
            self.table_tree.addTopLevelItem(other_item)
        self.ensure_other_queries_last()
        self._ignore_tree_changes = False
        self.on_tree_structure_changed()
        if current_item:
            self.table_tree.setCurrentItem(current_item)

    def sort_tables_in_folder(self, folder_item: QTreeWidgetItem):
        tables = []
        for i in range(folder_item.childCount()):
            child = folder_item.child(i)
            if child.data(0, self.TYPE_ROLE) == self.ITEM_TABLE:
                tables.append(child)
        if not tables:
            return
        for child in tables:
            folder_item.removeChild(child)
        tables.sort(key=lambda itm: itm.text(0).lower())
        for child in tables:
            self.ensure_columns_sorted(child)
            folder_item.addChild(child)

    def ensure_columns_sorted(self, table_item: QTreeWidgetItem):
        columns = []
        for i in range(table_item.childCount()):
            column_item = table_item.child(i)
            if column_item.data(0, self.TYPE_ROLE) == self.ITEM_COLUMN:
                columns.append(column_item)
        if not columns:
            return
        for column_item in columns:
            table_item.removeChild(column_item)
        columns.sort(key=lambda itm: itm.text(0).lower())
        for column_item in columns:
            table_item.addChild(column_item)

    def ensure_measure_items_sorted(self, parent_item: Optional[QTreeWidgetItem]):
        if not parent_item:
            return
        previous_flag = self._ignore_tree_changes
        self._ignore_tree_changes = True
        try:
            folders: List[QTreeWidgetItem] = []
            measures: List[QTreeWidgetItem] = []
            for idx in range(parent_item.childCount() - 1, -1, -1):
                child = parent_item.child(idx)
                child_type = child.data(0, self.TYPE_ROLE)
                if child_type == self.ITEM_MEASURE_FOLDER:
                    parent_item.takeChild(idx)
                    folders.append(child)
                elif child_type == self.ITEM_MEASURE:
                    parent_item.takeChild(idx)
                    measures.append(child)
            folders.sort(key=lambda itm: itm.text(0).strip().lower())
            measures.sort(key=lambda itm: itm.text(0).strip().lower())
            for folder in folders:
                parent_item.addChild(folder)
            for measure in measures:
                parent_item.addChild(measure)
            for folder in folders:
                self.ensure_measure_items_sorted(folder)
        finally:
            self._ignore_tree_changes = previous_flag

    def _determine_drop_parent(
        self,
        target_item: Optional[QTreeWidgetItem],
        items: List[QTreeWidgetItem],
        indicator: QAbstractItemView.DropIndicatorPosition,
    ) -> Optional[QTreeWidgetItem]:
        if indicator == QAbstractItemView.DropIndicatorPosition.OnItem:
            if target_item is None:
                return None
            target_type = target_item.data(0, self.TYPE_ROLE)
            if target_type == self.ITEM_TABLE:
                for dragged in items:
                    if dragged.data(0, self.TYPE_ROLE) == self.ITEM_TABLE:
                        return target_item.parent()
            return target_item
        if indicator in (
            QAbstractItemView.DropIndicatorPosition.AboveItem,
            QAbstractItemView.DropIndicatorPosition.BelowItem,
        ):
            return target_item.parent() if target_item else None
        return None

    def _is_descendant_item(
        self,
        node: Optional[QTreeWidgetItem],
        potential_ancestor: Optional[QTreeWidgetItem],
    ) -> bool:
        if node is None or potential_ancestor is None:
            return False
        current = node
        while current is not None:
            if current is potential_ancestor:
                return True
            current = current.parent()
        return False

    def _folder_contains_tables(self, folder_item: QTreeWidgetItem) -> bool:
        stack = [folder_item]
        while stack:
            node = stack.pop()
            for idx in range(node.childCount()):
                child = node.child(idx)
                child_type = child.data(0, self.TYPE_ROLE)
                if child_type == self.ITEM_TABLE:
                    return True
                if child_type == self.ITEM_FOLDER:
                    stack.append(child)
        return False

    def _validate_tree_drop(
        self,
        target_item: Optional[QTreeWidgetItem],
        items: List[QTreeWidgetItem],
        indicator: QAbstractItemView.DropIndicatorPosition,
    ) -> bool:
        if not items:
            return False

        target_type = target_item.data(0, self.TYPE_ROLE) if target_item else None

        if (
            indicator == QAbstractItemView.DropIndicatorPosition.OnItem
            and target_type == self.ITEM_TABLE
        ):
            for dragged in items:
                if dragged.data(0, self.TYPE_ROLE) == self.ITEM_TABLE:
                    return False

        if (
            indicator == QAbstractItemView.DropIndicatorPosition.OnItem
            and target_type == self.ITEM_FOLDER
        ):
            for dragged in items:
                if dragged.data(0, self.TYPE_ROLE) == self.ITEM_FOLDER and self._folder_contains_tables(dragged):
                    return False

        parent_item = self._determine_drop_parent(target_item, items, indicator)
        parent_type = parent_item.data(0, self.TYPE_ROLE) if parent_item else None

        for dragged in items:
            if dragged is None:
                return False
            if parent_item is not None and self._is_descendant_item(parent_item, dragged):
                return False

            child_type = dragged.data(0, self.TYPE_ROLE)
            if child_type is None:
                return False

            if child_type == self.ITEM_FOLDER and dragged.data(0, self.KEY_ROLE) is None and parent_item is not None:
                return False

            allowed_parents = self.ALLOWED_PARENTS.get(child_type, frozenset())
            if parent_item is None:
                if None not in allowed_parents:
                    return False
                continue

            if parent_type not in allowed_parents:
                return False

            allowed_children = self.ALLOWED_CHILDREN.get(parent_type, frozenset())
            if child_type not in allowed_children:
                return False

            if (
                child_type == self.ITEM_FOLDER
                and self._folder_contains_tables(dragged)
                and parent_item is not dragged.parent()
                and parent_type == self.ITEM_FOLDER
            ):
                return False

            if child_type in {self.ITEM_MEASURE, self.ITEM_MEASURE_FOLDER, self.ITEM_COLUMN}:
                target_table = self._table_item_for(parent_item)
                if target_table is None:
                    return False

        return True

    def show_tree_context_menu(self, position):
        menu = QMenu(self.table_tree)
        item = self.table_tree.itemAt(position)
        item_type = item.data(0, self.TYPE_ROLE) if item else None

        new_query_folder_action = None
        new_measure_action = None
        new_measure_folder_action = None
        rename_action = None
        delete_action = None

        if item is None:
            new_query_folder_action = menu.addAction("New Folder")
        elif item_type == self.ITEM_FOLDER:
            new_query_folder_action = menu.addAction("New Folder")
            if item.data(0, self.KEY_ROLE) is not None:
                rename_action = menu.addAction("Rename Folder")
                delete_action = menu.addAction("Delete Folder")
        elif item_type == self.ITEM_TABLE:
            new_measure_action = menu.addAction("New Measure")
            new_measure_folder_action = menu.addAction("New Display Folder")
        elif item_type == self.ITEM_MEASURE_FOLDER:
            new_measure_action = menu.addAction("New Measure")
            rename_action = menu.addAction("Rename Folder")
            delete_action = menu.addAction("Delete Folder")
        elif item_type == self.ITEM_MEASURE:
            rename_action = menu.addAction("Rename Measure")
            delete_action = menu.addAction("Delete Measure")
        else:
            new_query_folder_action = menu.addAction("New Folder")

        global_pos = self.table_tree.viewport().mapToGlobal(position)
        chosen = menu.exec(global_pos)

        if chosen == new_query_folder_action:
            parent_folder = item if (item and item.data(0, self.TYPE_ROLE) == self.ITEM_FOLDER) else None
            self.create_new_folder(parent_folder)
        elif new_measure_action and chosen == new_measure_action and item is not None:
            self.create_new_measure(item)
        elif new_measure_folder_action and chosen == new_measure_folder_action and item is not None:
            target = item
            if item_type == self.ITEM_MEASURE:
                target = item.parent() or self._table_item_for(item)
            if target is not None:
                self.create_new_measure_folder(target)
        elif rename_action and chosen == rename_action:
            if item_type == self.ITEM_FOLDER:
                self.rename_folder(item)
            else:
                self.table_tree.editItem(item)
        elif delete_action and chosen == delete_action:
            if item_type == self.ITEM_FOLDER:
                self.delete_folder(item)
            elif item_type == self.ITEM_MEASURE_FOLDER:
                self.delete_measure_folder(item)
            elif item_type == self.ITEM_MEASURE:
                self.delete_measure_item(item)

    def create_new_folder(self, parent_folder: Optional[QTreeWidgetItem] = None):
        name = self.generate_unique_folder_name("New Folder", parent_folder)
        parent_path = parent_folder.data(0, self.KEY_ROLE) if parent_folder else None
        full_path = f"{parent_path}/{name}" if parent_path else name
        folder_item = self._create_folder_item(full_path, name)
        # If a parent folder is provided, add as a child (nested folder).
        if parent_folder is not None:
            parent_folder.addChild(folder_item)
        else:
            # Otherwise, add at top-level just before Other Queries.
            other_item = self.find_other_queries_item()
            if other_item:
                idx = self.table_tree.indexOfTopLevelItem(other_item)
                self.table_tree.insertTopLevelItem(idx, folder_item)
            else:
                self.table_tree.addTopLevelItem(folder_item)
        self.table_tree.setCurrentItem(folder_item)
        self.table_tree.editItem(folder_item)
        self.on_tree_structure_changed()

    def create_new_measure(self, parent_item: QTreeWidgetItem):
        table_item = self._table_item_for(parent_item)
        if not table_item:
            return
        table_name = table_item.data(0, Qt.ItemDataRole.UserRole)
        if not table_name:
            return

        base_parent = parent_item if parent_item.data(0, self.TYPE_ROLE) == self.ITEM_MEASURE_FOLDER else table_item
        measure_name = self.generate_unique_measure_name("New Measure", table_name)

        folder_path = None
        if parent_item.data(0, self.TYPE_ROLE) == self.ITEM_MEASURE_FOLDER:
            folder_path = parent_item.data(0, self.KEY_ROLE) or self._normalize_display_path(parent_item.text(0))

        measure_id = uuid.uuid4().hex
        measure_entry = {
            "id": measure_id,
            "name": measure_name,
            "expression": "",
            "expression_indent": "    ",
            "indent": "    ",
            "display_folder": self._display_folder_from_path(folder_path),
            "lineage_tag": None,
            "format_string": None,
            "format_indent": "    ",
            "other_metadata": [],
            "quoted_name": False,
            "original_name_token": measure_name,
        }

        self.tables_data.setdefault(table_name, {}).setdefault("measures", []).append(measure_entry)
        self._measure_usage_flags[measure_id] = False
        measure_item = self._create_measure_item(table_name, measure_entry)
        base_parent.addChild(measure_item)
        self.ensure_measure_items_sorted(base_parent)
        self.table_tree.expandItem(base_parent)
        self.table_tree.setCurrentItem(measure_item)
        self.table_tree.editItem(measure_item)
        measure_item.setSelected(True)
        self._sync_measures_from_tree()
        self.display_measure_details(table_name, measure_id)
        self.mark_dirty()

    def create_new_measure_folder(self, parent_item: QTreeWidgetItem):
        table_item = self._table_item_for(parent_item)
        if not table_item:
            return
        table_name = table_item.data(0, Qt.ItemDataRole.UserRole)
        if not table_name:
            return

        folder_name = self.generate_unique_measure_folder_name("New Folder", table_item)
        folder_item = QTreeWidgetItem([folder_name, ""])
        folder_item.setData(0, self.TYPE_ROLE, self.ITEM_MEASURE_FOLDER)
        folder_item.setData(0, Qt.ItemDataRole.UserRole, table_name)
        flags = (
            Qt.ItemFlag.ItemIsEnabled
            | Qt.ItemFlag.ItemIsSelectable
            | Qt.ItemFlag.ItemIsEditable
            | Qt.ItemFlag.ItemIsDragEnabled
            | Qt.ItemFlag.ItemIsDropEnabled
        )
        folder_item.setFlags(flags)
        if self.folder_icon and not self.folder_icon.isNull():
            folder_item.setIcon(0, self.folder_icon)
        self._apply_usage_state(folder_item, None)

        container = (
            parent_item
            if parent_item.data(0, self.TYPE_ROLE) == self.ITEM_MEASURE_FOLDER
            else table_item
        )
        container.addChild(folder_item)

        if container.data(0, self.TYPE_ROLE) == self.ITEM_MEASURE_FOLDER:
            base_path = container.data(0, self.KEY_ROLE)
        else:
            base_path = None
        self._reset_measure_folder_paths(folder_item, base_path)
        self.ensure_measure_items_sorted(container)
        self.table_tree.expandItem(container)
        if container is not table_item:
            self.table_tree.expandItem(table_item)
        self.table_tree.setCurrentItem(folder_item)
        self.table_tree.editItem(folder_item)
        self.mark_dirty()

    def delete_measure_item(self, measure_item: QTreeWidgetItem):
        measure_id = measure_item.data(0, self.KEY_ROLE)
        table_name = measure_item.data(0, Qt.ItemDataRole.UserRole)
        table_item = self._table_item_for(measure_item)
        parent = measure_item.parent()
        if parent:
            parent.removeChild(measure_item)
            self.ensure_measure_items_sorted(parent)
        else:
            idx = self.table_tree.indexOfTopLevelItem(measure_item)
            if idx != -1:
                self.table_tree.takeTopLevelItem(idx)
            if table_item:
                self.ensure_measure_items_sorted(table_item)

        if measure_id:
            self._measure_usage_flags.pop(measure_id, None)

        if table_name and measure_id:
            measures = self.tables_data.get(table_name, {}).get("measures") or []
            self.tables_data[table_name]["measures"] = [m for m in measures if m.get("id") != measure_id]

        if self.current_measure_id == measure_id:
            self.display_table_details(table_name)

        self._sync_measures_from_tree()
        self.mark_dirty()

    def delete_measure_folder(self, folder_item: QTreeWidgetItem):
        table_item = self._table_item_for(folder_item)
        if not table_item:
            return
        parent = folder_item.parent()
        if parent is None:
            parent = table_item
        parent_path = parent.data(0, self.KEY_ROLE) if parent.data(0, self.TYPE_ROLE) == self.ITEM_MEASURE_FOLDER else None

        # Move all children up one level while preserving order
        while folder_item.childCount():
            child = folder_item.takeChild(0)
            parent.addChild(child)
            if child.data(0, self.TYPE_ROLE) == self.ITEM_MEASURE_FOLDER:
                self._reset_measure_folder_paths(child, parent_path)

        if folder_item.parent():
            folder_item.parent().removeChild(folder_item)
        else:
            idx = self.table_tree.indexOfTopLevelItem(folder_item)
            if idx != -1:
                self.table_tree.takeTopLevelItem(idx)

        self.ensure_measure_items_sorted(parent)
        self._sync_measures_from_tree()
        self.mark_dirty()

    def _reset_measure_folder_paths(self, folder_item: QTreeWidgetItem, parent_path: Optional[str]) -> None:
        segment = folder_item.text(0).strip()
        canonical = segment if not parent_path else f"{parent_path}/{segment}"
        folder_item.setData(0, self.KEY_ROLE, canonical)
        table_item = self._table_item_for(folder_item)
        if table_item:
            folder_item.setData(0, Qt.ItemDataRole.UserRole, table_item.data(0, Qt.ItemDataRole.UserRole))
        for idx in range(folder_item.childCount()):
            child = folder_item.child(idx)
            child_type = child.data(0, self.TYPE_ROLE)
            if child_type == self.ITEM_MEASURE_FOLDER:
                self._reset_measure_folder_paths(child, canonical)

    def rename_folder(self, item: QTreeWidgetItem):
        if not item or item.data(0, self.TYPE_ROLE) != self.ITEM_FOLDER:
            return
        if item.data(0, self.KEY_ROLE) is None:
            return
        self.table_tree.editItem(item)

    def delete_folder(self, item: QTreeWidgetItem):
        if not item or item.data(0, self.TYPE_ROLE) != self.ITEM_FOLDER:
            return
        folder_key = item.data(0, self.KEY_ROLE)
        if folder_key is None:
            return
        other_item = self.ensure_other_queries_folder()
        while item.childCount():
            child = item.takeChild(0)
            if child.data(0, self.TYPE_ROLE) == self.ITEM_TABLE:
                other_item.addChild(child)
        self.sort_tables_in_folder(other_item)
        parent = item.parent()
        if parent is None:
            idx = self.table_tree.indexOfTopLevelItem(item)
            if idx != -1:
                self.table_tree.takeTopLevelItem(idx)
        else:
            parent.removeChild(item)
        self.query_groups.pop(folder_key, None)
        self.ensure_other_queries_last()
        self.on_tree_structure_changed()

    def generate_unique_folder_name(self, base: str, parent: Optional[QTreeWidgetItem] = None) -> str:
        existing: set[str] = set()
        if parent is None:
            for i in range(self.table_tree.topLevelItemCount()):
                item = self.table_tree.topLevelItem(i)
                if item.data(0, self.TYPE_ROLE) != self.ITEM_FOLDER:
                    continue
                key = item.data(0, self.KEY_ROLE)
                if key is None:
                    continue
                existing.add(item.text(0).strip().lower())
            existing.add(self.OTHER_QUERIES_NAME.lower())
        else:
            for i in range(parent.childCount()):
                child = parent.child(i)
                if child.data(0, self.TYPE_ROLE) != self.ITEM_FOLDER:
                    continue
                key = child.data(0, self.KEY_ROLE)
                if key is None:
                    continue
                existing.add(child.text(0).strip().lower())
        candidate = base
        suffix = 1
        while candidate.lower() in existing:
            candidate = f"{base} {suffix}"
            suffix += 1
        return candidate

    def on_tree_item_changed(self, item: QTreeWidgetItem, column: int):
        if self._ignore_item_change or not item:
            return
        item_type = item.data(0, self.TYPE_ROLE)

        if item_type == self.ITEM_FOLDER:
            folder_key = item.data(0, self.KEY_ROLE)
            new_name = item.text(0).strip()
            if folder_key is None:
                self._ignore_item_change = True
                item.setText(0, self.OTHER_QUERIES_NAME)
                self._ignore_item_change = False
                return
            if not new_name:
                self._ignore_item_change = True
                item.setText(0, folder_key)
                self._ignore_item_change = False
                return
            if new_name == (folder_key.split("/")[-1] if folder_key else ""):
                return
            parent_item = item.parent()
            parent_key = parent_item.data(0, self.KEY_ROLE) if parent_item else None
            siblings = (
                [parent_item.child(i) for i in range(parent_item.childCount())] if parent_item
                else [self.table_tree.topLevelItem(i) for i in range(self.table_tree.topLevelItemCount())]
            )
            for sibling in siblings:
                if sibling is item:
                    continue
                if sibling.data(0, self.TYPE_ROLE) != self.ITEM_FOLDER:
                    continue
                sibling_key = sibling.data(0, self.KEY_ROLE)
                if sibling_key is None:
                    continue
                if sibling.text(0).strip().lower() == new_name.lower():
                    self._ignore_item_change = True
                    item.setText(0, folder_key.split("/")[-1] if folder_key else new_name)
                    self._ignore_item_change = False
                    return
            new_key = f"{parent_key}/{new_name}" if parent_key else new_name
            self._ignore_item_change = True
            item.setText(0, new_name)
            item.setData(0, self.KEY_ROLE, new_key)
            self._ignore_item_change = False
            self.on_tree_structure_changed()
        elif item_type == self.ITEM_MEASURE_FOLDER:
            old_path = item.data(0, self.KEY_ROLE)
            new_name = item.text(0).strip()
            if not new_name:
                self._ignore_item_change = True
                item.setText(0, (old_path.split("/")[-1] if old_path else "Folder"))
                self._ignore_item_change = False
                return
            parent_item = item.parent() or self._table_item_for(item)
            parent_path = parent_item.data(0, self.KEY_ROLE) if parent_item and parent_item.data(0, self.TYPE_ROLE) == self.ITEM_MEASURE_FOLDER else None
            siblings = [parent_item.child(i) for i in range(parent_item.childCount())] if parent_item else []
            for sibling in siblings:
                if sibling is item:
                    continue
                if sibling.data(0, self.TYPE_ROLE) != self.ITEM_MEASURE_FOLDER:
                    continue
                if sibling.text(0).strip().lower() == new_name.lower():
                    self._ignore_item_change = True
                    item.setText(0, old_path.split("/")[-1] if old_path else new_name)
                    self._ignore_item_change = False
                    return
            self._ignore_item_change = True
            self._reset_measure_folder_paths(item, parent_path)
            self._ignore_item_change = False
            if parent_item:
                self.ensure_measure_items_sorted(parent_item)
            self._sync_measures_from_tree()
            self.mark_dirty()
        elif item_type == self.ITEM_MEASURE:
            measure_id = item.data(0, self.KEY_ROLE)
            table_name = item.data(0, Qt.ItemDataRole.UserRole)
            new_name = item.text(0).strip()
            measure = self._get_measure_data(table_name, measure_id)
            previous_name = measure.get("name") if measure else ""

            if not new_name:
                self._ignore_item_change = True
                item.setText(0, previous_name)
                self._ignore_item_change = False
                return

            if table_name:
                for existing in self.tables_data.get(table_name, {}).get("measures") or []:
                    if existing.get("id") == measure_id:
                        continue
                    if (existing.get("name") or "").strip().lower() == new_name.lower():
                        self._ignore_item_change = True
                        item.setText(0, previous_name)
                        self._ignore_item_change = False
                        return

            if measure:
                measure["name"] = new_name
                measure["quoted_name"] = None
            container = item.parent() or self._table_item_for(item)
            if container:
                self.ensure_measure_items_sorted(container)
            self._sync_measures_from_tree()
            if self.current_measure_id == measure_id:
                if self.measure_view_mode == "format":
                    self.query_label.setText(f"Format String: {new_name}")
                else:
                    self.query_label.setText(f"Measure: {new_name}")
            self.mark_dirty()

    def _collect_tree_layout(self) -> Tuple[List[str], List[str], Dict[str, Optional[str]]]:
        """Traverse the tree to gather folder paths, table order, and group assignments."""
        folder_paths: List[str] = []
        table_order: List[str] = []
        table_groups: Dict[str, Optional[str]] = {}

        def traverse_folder(folder_item: QTreeWidgetItem, parent_parts: List[str]) -> None:
            name = folder_item.text(0).strip()
            is_other_queries = folder_item.data(0, self.KEY_ROLE) is None
            working_parts = parent_parts
            if is_other_queries:
                folder_item.setData(0, self.KEY_ROLE, None)
                working_parts = []
            else:
                working_parts = parent_parts + [name] if name else parent_parts.copy()
                path_value = "/".join(working_parts) if working_parts else name
                if path_value:
                    folder_paths.append(path_value)
                    folder_item.setData(0, self.KEY_ROLE, path_value)
                else:
                    folder_item.setData(0, self.KEY_ROLE, None)

            for child_index in range(folder_item.childCount()):
                child = folder_item.child(child_index)
                child_type = child.data(0, self.TYPE_ROLE)
                if child_type == self.ITEM_FOLDER:
                    traverse_folder(child, working_parts)
                elif child_type == self.ITEM_TABLE:
                    table_name = child.data(0, Qt.ItemDataRole.UserRole) or child.text(0).strip()
                    group_path = "/".join(working_parts) if working_parts else None
                    table_groups[table_name] = group_path
                    child.setData(0, self.KEY_ROLE, group_path)
                    info = self.tables_data.get(table_name)
                    table_type = ""
                    if info is not None:
                        info["query_group"] = group_path
                        table_type = (info.get("table_type") or "").lower()
                    self.ensure_columns_sorted(child)
                    for cc in range(child.childCount()):
                        col = child.child(cc)
                        if col.data(0, self.TYPE_ROLE) == self.ITEM_COLUMN:
                            col.setData(0, Qt.ItemDataRole.UserRole, table_name)
                    if table_type != "calculated":
                        table_order.append(table_name)

        for top_index in range(self.table_tree.topLevelItemCount()):
            top_item = self.table_tree.topLevelItem(top_index)
            item_type = top_item.data(0, self.TYPE_ROLE)
            if item_type == self.ITEM_FOLDER:
                traverse_folder(top_item, [])
            elif item_type == self.ITEM_TABLE:
                table_name = top_item.data(0, Qt.ItemDataRole.UserRole) or top_item.text(0).strip()
                table_groups[table_name] = None
                info = self.tables_data.get(table_name)
                table_type = ""
                if info is not None:
                    info["query_group"] = None
                    table_type = (info.get("table_type") or "").lower()
                self.ensure_columns_sorted(top_item)
                for cc in range(top_item.childCount()):
                    col = top_item.child(cc)
                    if col.data(0, self.TYPE_ROLE) == self.ITEM_COLUMN:
                        col.setData(0, Qt.ItemDataRole.UserRole, table_name)
                if table_type != "calculated":
                    table_order.append(table_name)

        return folder_paths, table_order, table_groups

    def on_tree_structure_changed(self, *args, **kwargs):
        if self._ignore_tree_changes:
            return
        self._ignore_tree_changes = True
        changed = False
        try:
            current_item = self.table_tree.currentItem()
            self.ensure_other_queries_last()

            def _item_type(node: Optional[QTreeWidgetItem]) -> Optional[str]:
                return node.data(0, self.TYPE_ROLE) if node is not None else None

            def _detach(item: QTreeWidgetItem) -> None:
                parent = item.parent()
                if parent is not None:
                    parent.removeChild(item)
                else:
                    idx_local = self.table_tree.indexOfTopLevelItem(item)
                    if idx_local != -1:
                        self.table_tree.takeTopLevelItem(idx_local)

            def _insert_at_root(item: QTreeWidgetItem) -> None:
                nonlocal changed
                other = self.ensure_other_queries_folder()
                idx_other = self.table_tree.indexOfTopLevelItem(other)
                if idx_other == -1:
                    self.table_tree.addTopLevelItem(item)
                else:
                    self.table_tree.insertTopLevelItem(idx_other, item)
                changed = True

            def _nearest_folder_ancestor(item: QTreeWidgetItem) -> Optional[QTreeWidgetItem]:
                current = item.parent()
                while current is not None:
                    if _item_type(current) == self.ITEM_FOLDER:
                        return current
                    current = current.parent()
                return None

            def _closest_table(item: QTreeWidgetItem) -> Optional[QTreeWidgetItem]:
                table_item = self._table_item_for(item)
                if table_item:
                    return table_item
                table_name = item.data(0, Qt.ItemDataRole.UserRole)
                if table_name:
                    return self._find_table_item(table_name)
                return None

            def _repair_parent(item: QTreeWidgetItem) -> bool:
                nonlocal changed
                node_type = _item_type(item)
                parent_item = item.parent()
                parent_type = _item_type(parent_item)

                allowed = self.ALLOWED_PARENTS.get(node_type, frozenset())
                if parent_type in allowed:
                    if node_type == self.ITEM_MEASURE_FOLDER:
                        table_item = self._table_item_for(item)
                        if table_item is None:
                            table_item = _closest_table(item)
                            if table_item is not None and table_item is not parent_item:
                                _detach(item)
                                table_item.addChild(item)
                                self._reset_measure_folder_paths(item, None)
                                self.ensure_measure_items_sorted(table_item)
                                changed = True
                                return True
                    return False

                if node_type == self.ITEM_FOLDER:
                    destination = _nearest_folder_ancestor(item)
                    _detach(item)
                    if destination is not None:
                        destination.addChild(item)
                    else:
                        _insert_at_root(item)
                    changed = True
                    return True

                if node_type == self.ITEM_TABLE:
                    destination_folder = _nearest_folder_ancestor(item)
                    if destination_folder is None:
                        destination_folder = self.ensure_other_queries_folder()
                    _detach(item)
                    destination_folder.addChild(item)
                    self.sort_tables_in_folder(destination_folder)
                    self.ensure_columns_sorted(item)
                    self.ensure_measure_items_sorted(item)
                    changed = True
                    return True

                if node_type == self.ITEM_COLUMN:
                    table_item = _closest_table(item)
                    if table_item is None:
                        return False
                    _detach(item)
                    table_item.addChild(item)
                    item.setData(0, Qt.ItemDataRole.UserRole, table_item.data(0, Qt.ItemDataRole.UserRole))
                    self.ensure_columns_sorted(table_item)
                    changed = True
                    return True

                if node_type == self.ITEM_MEASURE_FOLDER:
                    table_item = _closest_table(item)
                    if table_item is None:
                        return False
                    _detach(item)
                    table_item.addChild(item)
                    self._reset_measure_folder_paths(item, None)
                    self.ensure_measure_items_sorted(table_item)
                    changed = True
                    return True

                if node_type == self.ITEM_MEASURE:
                    table_item = _closest_table(item)
                    if table_item is None:
                        return False
                    _detach(item)
                    table_item.addChild(item)
                    item.setData(0, Qt.ItemDataRole.UserRole, table_item.data(0, Qt.ItemDataRole.UserRole))
                    self.ensure_measure_items_sorted(table_item)
                    changed = True
                    return True

                return False

            idx = 0
            while idx < self.table_tree.topLevelItemCount():
                top_item = self.table_tree.topLevelItem(idx)
                if top_item is None:
                    idx += 1
                    continue
                if _item_type(top_item) != self.ITEM_FOLDER:
                    if not _repair_parent(top_item):
                        idx += 1
                else:
                    idx += 1

            queue: List[QTreeWidgetItem] = []
            for i in range(self.table_tree.topLevelItemCount()):
                queue.append(self.table_tree.topLevelItem(i))

            while queue:
                node = queue.pop(0)
                if node is None:
                    continue

                if _repair_parent(node):
                    queue.append(node)
                    continue

                node_type = _item_type(node)
                allowed_children = self.ALLOWED_CHILDREN.get(node_type, frozenset())
                child_index = 0
                while child_index < node.childCount():
                    child = node.child(child_index)
                    if child is None:
                        child_index += 1
                        continue
                    child_type = _item_type(child)
                    if child_type not in allowed_children:
                        if not _repair_parent(child):
                            child_index += 1
                        else:
                            continue
                    else:
                        queue.append(child)
                        child_index += 1

                if node_type == self.ITEM_MEASURE_FOLDER:
                    parent_item = node.parent()
                    parent_path = (
                        parent_item.data(0, self.KEY_ROLE)
                        if parent_item and _item_type(parent_item) == self.ITEM_MEASURE_FOLDER
                        else None
                    )
                    self._reset_measure_folder_paths(node, parent_path)
                elif node_type == self.ITEM_TABLE:
                    self.ensure_columns_sorted(node)
                    self.ensure_measure_items_sorted(node)

            self.ensure_other_queries_last()

            def _ensure_tables_sorted_recursive(node: QTreeWidgetItem):
                if node.data(0, self.TYPE_ROLE) == self.ITEM_FOLDER:
                    self.sort_tables_in_folder(node)
                for idx in range(node.childCount()):
                    _ensure_tables_sorted_recursive(node.child(idx))

            for i in range(self.table_tree.topLevelItemCount()):
                _ensure_tables_sorted_recursive(self.table_tree.topLevelItem(i))

            self.ensure_other_queries_last()

            previous_order = list(self.query_order)
            previous_groups = dict(self.query_groups)
            previous_table_groups = {name: data.get("query_group") for name, data in self.tables_data.items()}

            folder_paths, table_order, table_groups = self._collect_tree_layout()
            new_query_groups = {path: idx for idx, path in enumerate(folder_paths)}

            table_group_changed = any(
                previous_table_groups.get(name) != table_groups.get(name)
                for name in set(previous_table_groups.keys()) | set(table_groups.keys())
            )

            changed = (
                table_order != previous_order
                or new_query_groups != previous_groups
                or table_group_changed
            )

            self.query_groups = new_query_groups
            self.query_order = table_order

            measure_changed = self._sync_measures_from_tree()
            changed = changed or measure_changed

            if current_item:
                self.table_tree.setCurrentItem(current_item)
        finally:
            self._ignore_tree_changes = False
        if changed:
            self._refresh_dax_writer_metadata()
            self.mark_dirty()

    def _update_model_tmdl(self, model_path: str, folder_paths: List[str], table_order: List[str]) -> None:
        with open(model_path, "r", encoding="utf-8", newline="") as f:
            original_text = f.read()

        newline = "\r\n" if "\r\n" in original_text else "\n"

        group_pattern = re.compile(
            r'(?m)^(?P<indent>[ \t]*)queryGroup\s+(?:\'[^\']*\'|"[^"]*"|[^\s\r\n]+)\s*\r?\n(?P<anno_indent>[ \t]*)annotation\s+PBI_QueryGroupOrder\s*=\s*\d+\s*\r?\n?'
        )
        first_group = group_pattern.search(original_text)
        group_indent = first_group.group("indent") if first_group else None
        annotation_indent = first_group.group("anno_indent") if first_group else None

        text = group_pattern.sub("", original_text)

        order_pattern = re.compile(r"(annotation\s+PBI_QueryOrder\s*=\s*)(\[.*?\])", re.DOTALL)
        order_match = order_pattern.search(text)
        if not order_match:
            raise ValueError("annotation PBI_QueryOrder not found in model.tmdl")

        order_start = order_match.start()
        order_end = order_match.end()
        order_prefix = order_match.group(1)

        line_start = text.rfind("\n", 0, order_start)
        if line_start == -1:
            line_start = 0
        else:
            line_start += 1
        order_indent = text[line_start:order_start]
        indent_unit = "\t" if "\t" in order_indent else "    "

        if group_indent is None:
            group_indent = order_indent
        if annotation_indent is None:
            annotation_indent = group_indent + indent_unit

        if table_order:
            new_list_repr = "[" + ",".join(f'"{name}"' for name in table_order) + "]"
        else:
            new_list_repr = "[]"
        updated_order_block = order_prefix + new_list_repr

        before_order = text[:order_start]
        after_order = text[order_end:]

        if folder_paths:
            group_lines = []
            for idx, path in enumerate(folder_paths):
                group_lines.append(f"{group_indent}queryGroup '{path}'")
                group_lines.append(f"{annotation_indent}annotation PBI_QueryGroupOrder = {idx}")
            group_block = newline.join(group_lines) + newline
        else:
            group_block = ""

        if group_block and before_order and not before_order.endswith(("\n", "\r")):
            before_order += newline

        new_text = before_order + group_block + updated_order_block + after_order

        if new_text != original_text:
            with open(model_path, "w", encoding="utf-8", newline="") as f:
                f.write(new_text)

    def _update_table_definition(self, table_path: str, group_path: Optional[str]) -> None:
        with open(table_path, "r", encoding="utf-8", newline="") as f:
            lines = f.readlines()

        if not lines:
            return

        line_ending = "\n"
        for line in lines:
            if line.endswith("\r\n"):
                line_ending = "\r\n"
                break
            if line.endswith("\n"):
                line_ending = "\n"
                break

        mode_idx = None
        mode_indent = ""
        query_idx = None
        query_indent = ""

        for idx, line in enumerate(lines):
            stripped = line.lstrip()
            if mode_idx is None and re.match(r"mode\s*:", stripped, re.IGNORECASE):
                mode_idx = idx
                mode_indent = line[: len(line) - len(stripped)]
            if query_idx is None and re.match(r"querygroup\b", stripped, re.IGNORECASE):
                query_idx = idx
                query_indent = line[: len(line) - len(stripped)]

        if mode_idx is None:
            return

        changed = False

        if group_path is None:
            if query_idx is not None:
                del lines[query_idx]
                changed = True
                if query_idx < len(lines) and not lines[query_idx].strip():
                    del lines[query_idx]
        else:
            target_indent = query_indent or mode_indent
            new_line = f"{target_indent}queryGroup: '{group_path}'{line_ending}"
            if query_idx is not None:
                if lines[query_idx] != new_line:
                    lines[query_idx] = new_line
                    changed = True
            else:
                insert_idx = mode_idx + 1
                lines.insert(insert_idx, new_line)
                changed = True

        if changed:
            with open(table_path, "w", encoding="utf-8", newline="") as f:
                f.writelines(lines)

    def _format_measure_name(self, measure: Dict[str, Any]) -> str:
        name = (measure.get("name") or "").strip()
        if not name:
            return "Measure"
        quoted = measure.get("quoted_name")
        if quoted is None:
            quoted = not re.match(r"^[A-Za-z_][A-Za-z0-9_]*$", name)
        if quoted:
            escaped = name.replace("'", "''")
            return f"'{escaped}'"
        return name

    def _render_measure_block(self, measure: Dict[str, Any], newline: str) -> str:
        indent = measure.get("indent") or "    "
        expression_indent = measure.get("expression_indent") or "    "
        lines: List[str] = []
        name_token = self._format_measure_name(measure)
        lines.append(f"{indent}measure {name_token} =")

        expression_text = (measure.get("expression") or "").splitlines()
        if not expression_text:
            expression_text = ["0"]
        for expr_line in expression_text:
            lines.append(f"{indent}{expression_indent}{expr_line}")

        display_folder = measure.get("display_folder")
        if display_folder:
            lines.append(f"{indent}displayFolder: {display_folder}")

        lineage_tag = measure.get("lineage_tag")
        if lineage_tag:
            lines.append(f"{indent}lineageTag: {lineage_tag}")

        for meta_line in measure.get("other_metadata") or []:
            lines.append(meta_line)

        format_text = measure.get("format_string")
        if format_text is not None:
            stripped = format_text.strip()
            if stripped:
                format_indent = measure.get("format_indent") or expression_indent
                lines.append(f"{indent}formatStringDefinition =")
                for fmt_line in format_text.splitlines():
                    lines.append(f"{indent}{format_indent}{fmt_line}")

        return newline.join(lines)

    def _render_measure_section(self, measures: List[Dict[str, Any]], newline: str) -> str:
        if not measures:
            return ""
        blocks = [self._render_measure_block(measure, newline) for measure in measures]
        section = (newline * 2).join(blocks)
        if not section.endswith(newline):
            section += newline
        section += newline
        return section

    def _write_table_measures(self, table_name: str, table_path: str) -> None:
        table_info = self.tables_data.get(table_name)
        if not table_info:
            return

        try:
            with open(table_path, "r", encoding="utf-8", newline="") as f:
                current_text = f.read()
        except OSError:
            return

        parse_info = _parse_table_measures(current_text)
        newline = parse_info.get("newline", "\n")
        new_section = self._render_measure_section(table_info.get("measures") or [], newline)
        section_range = parse_info.get("measures_section")

        if section_range:
            start, end = section_range
            if new_section:
                new_text = current_text[:start] + new_section + current_text[end:]
            else:
                new_text = current_text[:start] + current_text[end:]
        else:
            if not new_section:
                # No measures to insert and none existed.
                return
            insert_pos = parse_info.get("measure_insert_pos", len(current_text))
            before = current_text[:insert_pos]
            after = current_text[insert_pos:]
            if before and not before.endswith(newline):
                before += newline
            new_text = before + new_section
            if after and not after.startswith(newline):
                if not new_text.endswith(newline):
                    new_text += newline
            new_text += after

        with open(table_path, "w", encoding="utf-8", newline="") as f:
            f.write(new_text)

        updated_parse = _parse_table_measures(new_text)
        table_info["tmdl_text"] = new_text
        table_info["measure_section"] = updated_parse.get("measures_section")
        table_info["measure_insert_pos"] = updated_parse.get("measure_insert_pos")
        table_info["line_ending"] = updated_parse.get("newline", newline)

    def _create_folder_item(self, group_key: Optional[str], display_name: str) -> QTreeWidgetItem:
        folder_item = QTreeWidgetItem([display_name, ""])
        folder_item.setData(0, self.TYPE_ROLE, self.ITEM_FOLDER)
        folder_item.setData(0, self.KEY_ROLE, group_key)
        flags = Qt.ItemFlag.ItemIsEnabled | Qt.ItemFlag.ItemIsSelectable | Qt.ItemFlag.ItemIsDropEnabled
        if group_key is not None:
            flags |= Qt.ItemFlag.ItemIsDragEnabled | Qt.ItemFlag.ItemIsEditable
        folder_item.setFlags(flags)
        if self.folder_icon and not self.folder_icon.isNull():
            folder_item.setIcon(0, self.folder_icon)
        folder_item.setExpanded(False)
        self._apply_usage_state(folder_item, None)
        return folder_item

    def _create_table_item(self, table_name: str, group_key: Optional[str]) -> QTreeWidgetItem:
        table_item = QTreeWidgetItem([table_name, ""])
        table_item.setData(0, Qt.ItemDataRole.UserRole, table_name)
        table_item.setData(0, self.TYPE_ROLE, self.ITEM_TABLE)
        table_item.setData(0, self.KEY_ROLE, group_key)
        # Do not allow dropping into a table; only allow selecting/dragging it.
        flags = (
            Qt.ItemFlag.ItemIsSelectable
            | Qt.ItemFlag.ItemIsEnabled
            | Qt.ItemFlag.ItemIsDragEnabled
            | Qt.ItemFlag.ItemIsDropEnabled
        )
        table_item.setFlags(flags)
        table_type = (self.tables_data.get(table_name, {}).get("table_type") or "").lower()
        icon = self.table_icons.get(table_type)
        if icon and not icon.isNull():
            table_item.setIcon(0, icon)
        self._apply_usage_state(table_item, self._table_usage_flags.get(table_name))

        columns = sorted(self.tables_data.get(table_name, {}).get("columns", []), key=str.casefold)
        for column in columns:
            column_item = QTreeWidgetItem([column, ""])
            column_item.setData(0, Qt.ItemDataRole.UserRole, table_name)
            column_item.setData(0, self.TYPE_ROLE, self.ITEM_COLUMN)
            column_item.setFlags(Qt.ItemFlag.ItemIsSelectable | Qt.ItemFlag.ItemIsEnabled)
            self._apply_usage_state(column_item, self._column_usage_flags.get((table_name, column)))
            table_item.addChild(column_item)

        measures = list(self.tables_data.get(table_name, {}).get("measures") or [])
        for idx, measure in enumerate(measures):
            measure.setdefault("order", idx)
        measures.sort(key=lambda entry: entry.get("order", 0))
        folder_cache: Dict[str, QTreeWidgetItem] = {}
        for measure in measures:
            measure_id = measure.get("id")
            if not measure_id:
                measure_id = uuid.uuid4().hex
                measure["id"] = measure_id
            folder_path = self._normalize_display_path(measure.get("display_folder"))
            if folder_path:
                folder_item = self._ensure_measure_folder(table_item, table_name, folder_path, folder_cache)
                folder_item.addChild(self._create_measure_item(table_name, measure))
            else:
                table_item.addChild(self._create_measure_item(table_name, measure))
        self.ensure_measure_items_sorted(table_item)
        return table_item

    def _normalize_display_path(self, raw: Optional[str]) -> Optional[str]:
        if not raw:
            return None
        parts = [part.strip() for part in re.split(r"[\\/]", str(raw)) if part.strip()]
        return "/".join(parts) if parts else None

    def _display_folder_from_path(self, path: Optional[str]) -> Optional[str]:
        if not path:
            return None
        parts = [part.strip() for part in path.split("/") if part.strip()]
        return " / ".join(parts) if parts else None

    def _ensure_measure_folder(
        self,
        table_item: QTreeWidgetItem,
        table_name: str,
        folder_path: str,
        cache: Dict[str, QTreeWidgetItem],
    ) -> QTreeWidgetItem:
        normalized = self._normalize_display_path(folder_path)
        if not normalized:
            return table_item

        if normalized in cache:
            return cache[normalized]

        parts = normalized.split("/")
        current_path_parts: List[str] = []
        parent_item = table_item

        for segment in parts:
            current_path_parts.append(segment)
            current_path = "/".join(current_path_parts)
            existing = cache.get(current_path)
            if existing:
                parent_item = existing
                continue

            folder_item = QTreeWidgetItem([segment, ""])
            folder_item.setData(0, self.TYPE_ROLE, self.ITEM_MEASURE_FOLDER)
            folder_item.setData(0, self.KEY_ROLE, current_path)
            folder_item.setData(0, Qt.ItemDataRole.UserRole, table_name)
            flags = (
                Qt.ItemFlag.ItemIsEnabled
                | Qt.ItemFlag.ItemIsSelectable
                | Qt.ItemFlag.ItemIsEditable
                | Qt.ItemFlag.ItemIsDragEnabled
                | Qt.ItemFlag.ItemIsDropEnabled
            )
            folder_item.setFlags(flags)
            if self.folder_icon and not self.folder_icon.isNull():
                folder_item.setIcon(0, self.folder_icon)
            self._apply_usage_state(folder_item, None)
            parent_item.addChild(folder_item)
            cache[current_path] = folder_item
            parent_item = folder_item

        return cache[normalized]

    def _create_measure_item(self, table_name: str, measure: Dict[str, Any]) -> QTreeWidgetItem:
        display_name = measure.get("name") or ""
        item = QTreeWidgetItem([display_name, ""])
        item.setData(0, Qt.ItemDataRole.UserRole, table_name)
        item.setData(0, self.TYPE_ROLE, self.ITEM_MEASURE)
        item.setData(0, self.KEY_ROLE, measure.get("id"))
        flags = (
            Qt.ItemFlag.ItemIsEnabled
            | Qt.ItemFlag.ItemIsSelectable
            | Qt.ItemFlag.ItemIsEditable
            | Qt.ItemFlag.ItemIsDragEnabled
        )
        item.setFlags(flags)
        if hasattr(self, "measure_icon") and self.measure_icon and not self.measure_icon.isNull():
            item.setIcon(0, self.measure_icon)
        usage_flag = None
        measure_id = measure.get("id")
        if measure_id:
            usage_flag = self._measure_usage_flags.get(measure_id)
        self._apply_usage_state(item, usage_flag)
        return item

    def _get_measure_data(self, table_name: Optional[str], measure_id: Optional[str]) -> Optional[Dict[str, Any]]:
        if not table_name or not measure_id:
            return None
        table_entry = self.tables_data.get(table_name) or {}
        for measure in table_entry.get("measures") or []:
            if measure.get("id") == measure_id:
                return measure
        return None

    def _table_item_for(self, item: Optional[QTreeWidgetItem]) -> Optional[QTreeWidgetItem]:
        current = item
        while current is not None and current.data(0, self.TYPE_ROLE) != self.ITEM_TABLE:
            current = current.parent()
        return current

    def _find_table_item(self, table_name: Optional[str]) -> Optional[QTreeWidgetItem]:
        if not table_name:
            return None

        def search(node: QTreeWidgetItem) -> Optional[QTreeWidgetItem]:
            if node.data(0, self.TYPE_ROLE) == self.ITEM_TABLE and node.data(0, Qt.ItemDataRole.UserRole) == table_name:
                return node
            for idx in range(node.childCount()):
                result = search(node.child(idx))
                if result:
                    return result
            return None

        for idx in range(self.table_tree.topLevelItemCount()):
            found = search(self.table_tree.topLevelItem(idx))
            if found:
                return found
        return None

    def generate_unique_measure_name(self, base: str, table_name: str) -> str:
        existing = {
            (measure.get("name") or "").strip().lower()
            for measure in (self.tables_data.get(table_name, {}).get("measures") or [])
        }
        candidate = base
        suffix = 1
        while candidate.strip().lower() in existing:
            candidate = f"{base} {suffix}"
            suffix += 1
        return candidate

    def generate_unique_measure_folder_name(self, base: str, parent: QTreeWidgetItem) -> str:
        existing: set[str] = set()
        for idx in range(parent.childCount()):
            child = parent.child(idx)
            if child.data(0, self.TYPE_ROLE) == self.ITEM_MEASURE_FOLDER:
                existing.add(child.text(0).strip().lower())
        candidate = base
        suffix = 1
        while candidate.strip().lower() in existing:
            candidate = f"{base} {suffix}"
            suffix += 1
        return candidate

    def _sync_measures_from_tree(self) -> bool:
        """
        Rebuild measure ordering and display folders from the current tree layout.

        Returns True if any measure ordering or assignment changed.
        """
        prev_state = {
            table: [
                (idx, measure.get("id"), measure.get("display_folder"), measure.get("name"))
                for idx, measure in enumerate(data.get("measures") or [])
            ]
            for table, data in self.tables_data.items()
        }

        measure_lookup: Dict[str, Dict[str, Any]] = {}
        for table_name, data in self.tables_data.items():
            for measure in data.get("measures") or []:
                identifier = measure.get("id")
                if identifier:
                    measure_lookup[identifier] = measure

        updated: Dict[str, List[Dict[str, Any]]] = {}

        def traverse(node: QTreeWidgetItem, table_name: str, parent_path: Optional[str]) -> None:
            for idx in range(node.childCount()):
                child = node.child(idx)
                child_type = child.data(0, self.TYPE_ROLE)
                if child_type == self.ITEM_MEASURE_FOLDER:
                    segment = child.text(0).strip()
                    canonical = segment if not parent_path else f"{parent_path}/{segment}"
                    child.setData(0, self.KEY_ROLE, canonical)
                    child.setData(0, Qt.ItemDataRole.UserRole, table_name)
                    traverse(child, table_name, canonical)
                elif child_type == self.ITEM_MEASURE:
                    measure_id = child.data(0, self.KEY_ROLE)
                    if not measure_id:
                        continue
                    measure = measure_lookup.get(measure_id)
                    if measure is None:
                        measure = {
                            "id": measure_id,
                            "name": child.text(0),
                            "expression": "",
                            "expression_indent": "    ",
                            "indent": "    ",
                            "display_folder": None,
                            "lineage_tag": None,
                            "format_string": None,
                            "format_indent": "    ",
                            "other_metadata": [],
                        }
                        measure_lookup[measure_id] = measure
                    measure["name"] = child.text(0)
                    measure["display_folder"] = self._display_folder_from_path(parent_path)
                    measure["order"] = len(updated.setdefault(table_name, []))
                    updated.setdefault(table_name, []).append(measure)
                    child.setData(0, Qt.ItemDataRole.UserRole, table_name)
                else:
                    # Recurse into child containers
                    traverse(child, table_name, parent_path)

        def iter_table_items() -> List[QTreeWidgetItem]:
            result: List[QTreeWidgetItem] = []

            def collect(item: QTreeWidgetItem):
                if item.data(0, self.TYPE_ROLE) == self.ITEM_TABLE:
                    result.append(item)
                for i in range(item.childCount()):
                    collect(item.child(i))

            for i in range(self.table_tree.topLevelItemCount()):
                collect(self.table_tree.topLevelItem(i))
            return result

        table_items = iter_table_items()
        for table_item in table_items:
            self.ensure_measure_items_sorted(table_item)

        for table_item in table_items:
            table_name = table_item.data(0, Qt.ItemDataRole.UserRole)
            if not table_name:
                continue
            updated.setdefault(table_name, [])
            traverse(table_item, table_name, None)

        changed = False
        for table_name, data in self.tables_data.items():
            if table_name in updated:
                new_list = updated[table_name]
            else:
                new_list = data.get("measures") or []
            data["measures"] = new_list
            summary = [
                (idx, m.get("id"), m.get("display_folder"), m.get("name"))
                for idx, m in enumerate(new_list)
            ]
            if summary != prev_state.get(table_name, []):
                changed = True

        return changed

    def ensure_other_queries_folder(self) -> QTreeWidgetItem:
        item = self.find_other_queries_item()
        if item:
            return item
        item = self._create_folder_item(None, self.OTHER_QUERIES_NAME)
        self.table_tree.addTopLevelItem(item)
        self.ensure_other_queries_last()
        return item

    def ensure_other_queries_last(self):
        other = self.find_other_queries_item()
        if not other:
            return
        idx = self.table_tree.indexOfTopLevelItem(other)
        if idx == -1 or idx == self.table_tree.topLevelItemCount() - 1:
            return
        current_item = self.table_tree.currentItem()
        other = self.table_tree.takeTopLevelItem(idx)
        self.table_tree.addTopLevelItem(other)
        if current_item:
            self.table_tree.setCurrentItem(current_item)

    def find_other_queries_item(self) -> Optional[QTreeWidgetItem]:
        for i in range(self.table_tree.topLevelItemCount()):
            item = self.table_tree.topLevelItem(i)
            if item.data(0, self.KEY_ROLE) is None:
                return item
        return None


