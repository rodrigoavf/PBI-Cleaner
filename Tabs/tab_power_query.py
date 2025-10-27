import ast
import os
import re
from typing import Dict, Optional

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
        self.init_ui()
        if pbip_file:
            self.load_tables()

    def init_ui(self):
        """Build the user interface."""
        main_layout = QVBoxLayout(self)
        main_layout.setContentsMargins(8, 8, 8, 8)
        main_layout.setSpacing(8)

        top_bar = QHBoxLayout()
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
        self.table_tree.setSelectionMode(QAbstractItemView.SelectionMode.MultiSelection)
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
        hotkey_hint_left = QLabel("F2: Rename\nDelete: Delete\nCtrl+N: New folder\nSpace: Set as default\nAlt+Up: Move up\nAlt+Down: Move down")
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

    def setup_shortcuts(self):
        self.new_folder_shortcut = QShortcut(QKeySequence("Ctrl+N"), self)
        self.new_folder_shortcut.activated.connect(self.create_new_folder)

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

        try:
            with open(model_tmdl, "r", encoding="utf-8") as f:
                model_tmdl_data = f.read()

            order_match = re.search(r"annotation\s+PBI_QueryOrder\s*=\s*(\[.*?\])", model_tmdl_data, re.DOTALL)
            if order_match:
                self.query_order = ast.literal_eval(order_match.group(1))
            else:
                self.query_order = []

            group_matches = re.findall(
                r"queryGroup\s+(\w+)\s*\n\s*annotation\s+PBI_QueryGroupOrder\s*=\s*(\d+)",
                model_tmdl_data,
            )
            self.query_groups = {name: int(order) for name, order in group_matches}

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
                query_group = query_group_match.group(1).strip() if query_group_match else None

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

        except Exception as exc:
            QMessageBox.critical(self, "Error", f"Failed to load tables:\n{exc}")

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
        for table_name in sorted_tables:
            info = self.tables_data.get(table_name, {})
            group_key = info.get("query_group") or None
            group_map.setdefault(group_key, []).append(table_name)

        if None not in group_map:
            group_map[None] = []

        table_order = {name.lower(): idx for idx, name in enumerate(self.query_order)}

        def table_sort_key(name: str):
            return (table_order.get(name.lower(), len(table_order)), name.lower())

        group_order_list = []
        for group_key in group_map:
            if group_key is None:
                display_name = self.OTHER_QUERIES_NAME
                group_order_list.append((1, float('inf'), display_name.lower(), group_key, display_name))
            else:
                order_val = self.query_groups.get(group_key, float('inf'))
                display_name = group_key
                group_order_list.append((0, order_val, display_name.lower(), group_key, display_name))
        group_order_list.sort()

        for _, _, _, group_key, display_name in group_order_list:
            group_item = self._create_folder_item(group_key, display_name)
            self.table_tree.addTopLevelItem(group_item)

            tables_in_group = sorted(group_map.get(group_key, []), key=table_sort_key)
            for table_name in tables_in_group:
                table_item = self._create_table_item(table_name, group_key)
                group_item.addChild(table_item)

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

    def on_import_mode_changed(self, index: int):
        """Persist import mode changes to in-memory state."""
        if self.ignore_editor_changes or not self.current_table:
            return
        value = self.import_mode_combo.itemData(index)
        self.tables_data[self.current_table]["import_mode"] = value if value else None

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
            self.create_new_folder()
        elif rename_action and chosen == rename_action:
            self.rename_folder(item)
        elif delete_action and chosen == delete_action:
            self.delete_folder(item)

    def create_new_folder(self):
        name = self.generate_unique_folder_name("New Folder")
        folder_item = self._create_folder_item(name, name)
        other_item = self.find_other_queries_item()
        if other_item:
            idx = self.table_tree.indexOfTopLevelItem(other_item)
            self.table_tree.insertTopLevelItem(idx, folder_item)
        else:
            self.table_tree.addTopLevelItem(folder_item)
        max_order = max(self.query_groups.values(), default=-1)
        self.query_groups[name] = max_order + 1
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
        idx = self.table_tree.indexOfTopLevelItem(item)
        self.table_tree.takeTopLevelItem(idx)
        self.query_groups.pop(folder_key, None)
        self.ensure_other_queries_last()
        self.on_tree_structure_changed()

    def generate_unique_folder_name(self, base: str) -> str:
        existing = {self.table_tree.topLevelItem(i).text(0).lower() for i in range(self.table_tree.topLevelItemCount())}
        existing.add(self.OTHER_QUERIES_NAME.lower())
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
        # Ensure uniqueness
        for i in range(self.table_tree.topLevelItemCount()):
            other = self.table_tree.topLevelItem(i)
            if other is item:
                continue
            if other.text(0).strip().lower() == new_name.lower():
                self._ignore_item_change = True
                item.setText(0, folder_key)
                self._ignore_item_change = False
                return
        if new_name == folder_key:
            return
        self._ignore_item_change = True
        item.setText(0, new_name)
        item.setData(0, self.KEY_ROLE, new_name)
        self._ignore_item_change = False
        order_val = self.query_groups.pop(folder_key, len(self.query_groups))
        self.query_groups[new_name] = order_val
        self.on_tree_structure_changed()

    def on_tree_structure_changed(self, *args, **kwargs):
        if self._ignore_tree_changes:
            return
        self._ignore_tree_changes = True
        try:
            current_item = self.table_tree.currentItem()
            self.ensure_other_queries_last()

            for idx in range(self.table_tree.topLevelItemCount()):
                folder = self.table_tree.topLevelItem(idx)
                child_idx = folder.childCount() - 1
                while child_idx >= 0:
                    child = folder.child(child_idx)
                    if child.data(0, self.TYPE_ROLE) == self.ITEM_FOLDER:
                        folder.takeChild(child_idx)
                        self.table_tree.addTopLevelItem(child)
                    child_idx -= 1
            self.ensure_other_queries_last()

            new_group_order: dict[str, int] = {}
            order = 0
            for idx in range(self.table_tree.topLevelItemCount()):
                folder = self.table_tree.topLevelItem(idx)
                folder_key = folder.data(0, self.KEY_ROLE)
                if folder_key is not None:
                    new_group_order[folder_key] = order
                    order += 1
                for child_idx in range(folder.childCount()):
                    child = folder.child(child_idx)
                    item_type = child.data(0, self.TYPE_ROLE)
                    if item_type == self.ITEM_TABLE:
                        table_name = child.data(0, Qt.ItemDataRole.UserRole)
                        self.tables_data[table_name]["query_group"] = folder_key
                        child.setData(0, self.KEY_ROLE, folder_key)
                        self.ensure_columns_sorted(child)
                    elif item_type == self.ITEM_COLUMN:
                        parent = child.parent()
                        if parent:
                            table_name = parent.data(0, Qt.ItemDataRole.UserRole)
                            child.setData(0, Qt.ItemDataRole.UserRole, table_name)
            self.query_groups = new_group_order
            if current_item:
                self.table_tree.setCurrentItem(current_item)
        finally:
            self._ignore_tree_changes = False

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
        flags = Qt.ItemFlag.ItemIsSelectable | Qt.ItemFlag.ItemIsEnabled | Qt.ItemFlag.ItemIsDragEnabled | Qt.ItemFlag.ItemIsDropEnabled
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
