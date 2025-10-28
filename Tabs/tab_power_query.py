import ast
import os
import re
from collections import defaultdict
from typing import Dict, Optional, List, Tuple

from PyQt6.QtCore import Qt
from PyQt6.QtWidgets import (
    QAbstractItemView,
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
)
from PyQt6.QtGui import (
    QFont, QIcon, QDragEnterEvent, QDropEvent, QShortcut, QKeySequence, QIcon, QPixmap, QImage
)
from Coding.code_editor import CodeEditor
from common_functions import code_editor_font, APP_THEME


class PowerQueryTab(QWidget):
    """Tab that surfaces Power Query table metadata with editable details."""

    TYPE_ROLE = Qt.ItemDataRole.UserRole + 1
    KEY_ROLE = Qt.ItemDataRole.UserRole + 2
    ITEM_FOLDER = "folder"
    ITEM_TABLE = "table"
    ITEM_COLUMN = "column"
    OTHER_QUERIES_NAME = "Other Queries"

    def __init__(self, pbip_file: Optional[str] = None):
        super().__init__()
        self.pbip_file = pbip_file
        self.tables_data: Dict[str, Dict[str, Optional[str]]] = {}
        self.query_order = []
        self.query_groups = {}
        self.current_table: Optional[str] = None
        self.ignore_editor_changes = False
        self._ignore_tree_changes = False
        self._ignore_item_change = False
        self._loading_data = False
        self.is_dirty = False
        self.save_button: Optional[QPushButton] = None
        self.init_ui()
        if pbip_file:
            self.load_tables()

    def init_ui(self):
        """Build the user interface."""
        main_layout = QVBoxLayout(self)
        main_layout.setContentsMargins(8, 8, 8, 8)
        main_layout.setSpacing(8)

        top_bar = QHBoxLayout()
        self.save_button = QPushButton("Save Changes")
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

        self.table_tree = QTreeWidget()
        self.table_tree.setHeaderHidden(True)
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

        splitter.addWidget(left_widget)

        right_widget = QWidget()
        right_layout = QVBoxLayout(right_widget)
        right_layout.setContentsMargins(3, 0, 0, 0)

        self.query_label = QLabel("Query")
        right_layout.addWidget(self.query_label)

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

        form_layout = QFormLayout()
        form_layout.setContentsMargins(0, 0, 0, 0)
        form_layout.setSpacing(6)

        # Hotkey hints - right
        hotkey_hint_right = QLabel("Ctrl+K+C: Comment   |   Ctrl+K+U: Uncomment   |   Alt+Up: Move line up   |   Alt+Down: Move line down")
        hotkey_hint_right.setStyleSheet("color: #666666; font-size: 10px;")
        hotkey_hint_right.setWordWrap(True)
        right_layout.addWidget(hotkey_hint_right)

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

        right_layout.addLayout(form_layout)

        splitter.addWidget(right_widget)
        splitter.setSizes([250, 750])

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
        self.delete_folder_shortcut.activated.connect(self.delete_selected_folders)

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

    def delete_selected_folders(self):
        selected = [item for item in self.table_tree.selectedItems() if item.data(0, self.TYPE_ROLE) == self.ITEM_FOLDER and item.data(0, self.KEY_ROLE) is not None]
        if not selected:
            return
        # Delete deeper folders first so parents still exist when needed
        selected.sort(key=self._item_depth, reverse=True)
        for folder in selected:
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
        self.pbip_file = chosen
        self.load_tables()

    def refresh_tables(self):
        """Reload tables using the current PBIP file."""
        self.load_tables()

    def load_tables(self):
        """Extract tables, columns, and related metadata from the PBIP definition."""
        self.clear_details()

        if not self.pbip_file:
            QMessageBox.information(self, "No PBIP", "Select a PBIP file to load Power Query metadata.")
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

        self._loading_data = True
        try:
            with open(model_tmdl, "r", encoding="utf-8") as f:
                model_tmdl_data = f.read()

            order_match = re.search(r"annotation\s+PBI_QueryOrder\s*=\s*(\[.*?\])", model_tmdl_data, re.DOTALL)
            if order_match:
                self.query_order = ast.literal_eval(order_match.group(1))
            else:
                self.query_order = []

            group_pattern = re.compile(
                r"(?mi)^\s*queryGroup\s+(?P<name>'[^']+'|\"[^\"]+\"|[^\s\r\n]+)\s*\r?\n\s*annotation\s+PBI_QueryGroupOrder\s*=\s*(?P<order>\d+)"
            )
            query_groups: Dict[str, int] = {}
            for match in group_pattern.finditer(model_tmdl_data):
                normalized = self._normalize_group_path(match.group("name"))
                if not normalized:
                    continue
                order_value = int(match.group("order"))
                if normalized in query_groups:
                    query_groups[normalized] = min(query_groups[normalized], order_value)
                else:
                    query_groups[normalized] = order_value
            self.query_groups = query_groups

            tables_data = {}
            for filename in os.listdir(tables_dir):
                if not filename.endswith(".tmdl"):
                    continue

                table_name = os.path.splitext(filename)[0]
                table_path = os.path.join(tables_dir, filename)
                with open(table_path, "r", encoding="utf-8") as table_file:
                    tmdl_data = table_file.read()

                columns = re.findall(r'(?mi)^\s*column\s+([A-Za-z0-9_]+)\s*$', tmdl_data)

                mode_match = re.search(r'(?mi)^\s*mode\s*:\s*([^\r\n]+)', tmdl_data)
                mode = mode_match.group(1).strip() if mode_match else (
                    re.search(r'(?mi)^\s*annotation\s+PBI_DataMode\s*=\s*"?(.*?)"?\s*$', tmdl_data).group(1).strip()
                    if re.search(r'(?mi)^\s*annotation\s+PBI_DataMode', tmdl_data) else None
                )
                mode = mode.lower() if mode else None

                table_type_match = r'(?mi)^\s*partition\s+([A-Za-z0-9_-]+)\s*=\s*(m|calculated)\s*$'
                table_type = re.search(table_type_match, tmdl_data)
                table_type_value = table_type.group(2).lower() if table_type else "m"

                query_group_match = re.search(r'(?mi)^\s*queryGroup\s*:\s*([^\r\n]+)', tmdl_data) \
                    or re.search(r'(?mi)^\s*queryGroup\s+([^\r\n]+)', tmdl_data)
                query_group = self._normalize_group_path(query_group_match.group(1)) if query_group_match else None

                def unescape_quoted(s: str) -> str:
                    """Unescape quoted M code from TMDL text."""
                    return s.encode('utf-8').decode('unicode_escape')

                def _strip_any_fence(text: str) -> str:
                    """Remove ``` or ´´´ fences (with or without language specifier)."""
                    if not text:
                        return text

                    t = text.strip()

                    # Multiline fences: ```lang\n ... \n```  or  ´´´lang\n ... \n´´´
                    m = re.match(r'^\s*([`´])\1\1[^\r\n]*\r?\n([\s\S]*?)\r?\n\1\1\1\s*$', t)
                    if m:
                        return m.group(2).strip()

                    # Single-line fences: ```code```  or  ´´´code´´´
                    m = re.match(r'^\s*([`´])\1\1[^\r\n]*\s*([\s\S]*?)\s*\1\1\1\s*$', t)
                    if m:
                        return m.group(2).strip()

                    return t
                
                def extract_table_code(tmdl_text: str):
                    """Extract the M (Power Query) code block from a .tmdl section."""

                    # --- Step 1: normalize indentation before extraction ---
                    # Convert every 4 spaces to a single tab for consistency
                    tmdl_text = tmdl_text.replace("    ", "\t")

                    # --- Step 2: normal extraction logic ---
                    # A) Quoted expression form: expression = "let\n...."
                    mq = re.search(r'(?ms)^\s*expression\s*=\s*"((?:[^"\\]|\\.)*)"', tmdl_text)
                    if mq:
                        return unescape_quoted(mq.group(1)).strip()

                    # B) Indented block after a standalone "source =" line
                    src = re.search(r'(?m)^\s*source\s*=\s*$', tmdl_text)
                    if not src:
                        # Fallback: inline source = <content> ... up to next annotation or EOF
                        m_inline = re.search(r'(?ms)^\s*source\s*=\s*(.+?)(?=^\s*annotation\b|^\S|\Z)', tmdl_text)
                        result = m_inline.group(1).rstrip() if m_inline else None
                        if result:
                            # Remove first 4 spaces/tabs from each line
                            result = re.sub(r'^(?:[ \t]{0,4})', '', result, flags=re.MULTILINE)
                            result = _strip_any_fence(result.strip())
                        return result

                    start = src.end()
                    lines = tmdl_text[start:].splitlines()

                    # find first non-empty line to set base indentation
                    i = 0
                    while i < len(lines) and lines[i].strip() == "":
                        i += 1
                    if i == len(lines):
                        return None

                    first = lines[i]
                    base_indent = len(first) - len(first.lstrip())

                    out = []
                    for line in lines[i:]:
                        stripped = line.lstrip()
                        # stop if an annotation starts (even if indented)
                        if stripped.startswith("annotation "):
                            break
                        # stop if indentation drops below the first M line
                        if stripped and (len(line) - len(stripped)) < base_indent:
                            break
                        out.append(line)

                    # trim trailing blank lines
                    while out and out[-1].strip() == "":
                        out.pop()

                    # --- Step 3: remove up to 4 leading spaces/tabs per line ---
                    result = "\n".join(out)
                    result = re.sub(r'^(?:[ \t]{0,4})', '', result, flags=re.MULTILINE)

                    # --- Step 4: clean code fences ---
                    result = _strip_any_fence(result.strip())

                    return result

                table_code = extract_table_code(tmdl_data) or ""
                if table_type_value == "calculated":
                    code_text = table_code
                    code_language = "dax"
                    table_type_value = "calculated"
                else:
                    code_text = table_code
                    code_language = "m"

                tables_data[table_name] = {
                    "columns": columns,
                    "import_mode": mode,
                    "query_group": query_group,
                    "code_text": code_text,
                    "code_language": code_language,
                    "table_type": table_type_value,
                }

            self.tables_data = tables_data
            self.populate_tree()
            self._set_dirty(False)

        except Exception as exc:
            QMessageBox.critical(self, "Error", f"Failed to load tables:\n{exc}")
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
                if (info.get("table_type") or "").lower() == "calculated":
                    continue
                table_path = os.path.join(tables_dir, f"{table_name}.tmdl")
                if not os.path.isfile(table_path):
                    continue
                self._update_table_definition(table_path, group_path)

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
        self.ignore_editor_changes = True

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

    def on_query_text_changed(self):
        """Persist query/expression changes to in-memory state."""
        if self.ignore_editor_changes or not self.current_table:
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
        self.ignore_editor_changes = True
        self.query_label.setText("Query")
        self.query_editor.clear()
        self.query_editor.setEnabled(False)
        self.import_mode_combo.blockSignals(True)
        self.import_mode_combo.setCurrentIndex(0)
        self.import_mode_combo.blockSignals(False)
        self.import_mode_combo.setEnabled(False)
        self.ignore_editor_changes = False

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

    def show_tree_context_menu(self, position):
        menu = QMenu(self.table_tree)
        new_folder_action = menu.addAction("New Folder")
        rename_action = None
        delete_action = None

        item = self.table_tree.itemAt(position)
        if item and item.data(0, self.TYPE_ROLE) == self.ITEM_FOLDER and item.data(0, self.KEY_ROLE) is not None:
            rename_action = menu.addAction("Rename Folder")
            delete_action = menu.addAction("Delete Folder")

        global_pos = self.table_tree.viewport().mapToGlobal(position)
        chosen = menu.exec(global_pos)

        if chosen == new_folder_action:
            # If a folder was right-clicked, create the new folder inside it
            parent_folder = item if (item and item.data(0, self.TYPE_ROLE) == self.ITEM_FOLDER) else None
            self.create_new_folder(parent_folder)
        elif rename_action and chosen == rename_action:
            self.rename_folder(item)
        elif delete_action and chosen == delete_action:
            self.delete_folder(item)

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
        if item.data(0, self.TYPE_ROLE) != self.ITEM_FOLDER:
            return
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

            # Enforce structure rules:
            # - No items (tables or folders) are allowed inside a table.
            # - Tables may not exist at top-level; move them to Other Queries.
            # - Folders can be nested (allowed), so do not force-flatten them.

            # 1) Move any top-level tables to Other Queries (keeps structure sane)
            idx = self.table_tree.topLevelItemCount() - 1
            while idx >= 0:
                top_item = self.table_tree.topLevelItem(idx)
                if top_item and top_item.data(0, self.TYPE_ROLE) == self.ITEM_TABLE:
                    taken = self.table_tree.takeTopLevelItem(idx)
                    other = self.ensure_other_queries_folder()
                    other.addChild(taken)
                idx -= 1

            # 2) For each table anywhere in the tree, ensure its children are columns only.
            def _fix_table_children(item: QTreeWidgetItem):
                if item.data(0, self.TYPE_ROLE) == self.ITEM_TABLE:
                    parent_folder = item.parent()
                    move_index = item.childCount() - 1
                    while move_index >= 0:
                        ch = item.child(move_index)
                        ch_type = ch.data(0, self.TYPE_ROLE)
                        if ch_type == self.ITEM_TABLE:
                            # Table inside table -> move as sibling under the parent folder
                            item.takeChild(move_index)
                            if parent_folder is not None:
                                parent_folder.addChild(ch)
                            else:
                                self.ensure_other_queries_folder().addChild(ch)
                        elif ch_type == self.ITEM_FOLDER:
                            # Folder inside table -> move to top-level
                            item.takeChild(move_index)
                            self.table_tree.addTopLevelItem(ch)
                        move_index -= 1
                # Recurse into children (folders may contain tables)
                for i in range(item.childCount()):
                    _fix_table_children(item.child(i))

            for i in range(self.table_tree.topLevelItemCount()):
                _fix_table_children(self.table_tree.topLevelItem(i))

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

            if current_item:
                self.table_tree.setCurrentItem(current_item)
        finally:
            self._ignore_tree_changes = False
        if changed:
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
            item_indent = order_indent + indent_unit
            items = [f'{item_indent}"{name}"' for name in table_order]
            new_list_repr = "[" + newline + (newline.join(items)) + newline + order_indent + "]"
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

    def _create_folder_item(self, group_key: Optional[str], display_name: str) -> QTreeWidgetItem:
        folder_item = QTreeWidgetItem([display_name])
        folder_item.setData(0, self.TYPE_ROLE, self.ITEM_FOLDER)
        folder_item.setData(0, self.KEY_ROLE, group_key)
        flags = Qt.ItemFlag.ItemIsEnabled | Qt.ItemFlag.ItemIsSelectable | Qt.ItemFlag.ItemIsDropEnabled
        if group_key is not None:
            flags |= Qt.ItemFlag.ItemIsDragEnabled | Qt.ItemFlag.ItemIsEditable
        folder_item.setFlags(flags)
        if self.folder_icon and not self.folder_icon.isNull():
            folder_item.setIcon(0, self.folder_icon)
        folder_item.setExpanded(False)
        return folder_item

    def _create_table_item(self, table_name: str, group_key: Optional[str]) -> QTreeWidgetItem:
        table_item = QTreeWidgetItem([table_name])
        table_item.setData(0, Qt.ItemDataRole.UserRole, table_name)
        table_item.setData(0, self.TYPE_ROLE, self.ITEM_TABLE)
        table_item.setData(0, self.KEY_ROLE, group_key)
        # Do not allow dropping into a table; only allow selecting/dragging it.
        flags = Qt.ItemFlag.ItemIsSelectable | Qt.ItemFlag.ItemIsEnabled | Qt.ItemFlag.ItemIsDragEnabled
        table_item.setFlags(flags)
        table_type = (self.tables_data.get(table_name, {}).get("table_type") or "").lower()
        icon = self.table_icons.get(table_type)
        if icon and not icon.isNull():
            table_item.setIcon(0, icon)

        columns = sorted(self.tables_data.get(table_name, {}).get("columns", []), key=str.casefold)
        for column in columns:
            column_item = QTreeWidgetItem([column])
            column_item.setData(0, Qt.ItemDataRole.UserRole, table_name)
            column_item.setData(0, self.TYPE_ROLE, self.ITEM_COLUMN)
            column_item.setFlags(Qt.ItemFlag.ItemIsSelectable | Qt.ItemFlag.ItemIsEnabled)
            table_item.addChild(column_item)
        return table_item

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
