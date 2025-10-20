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
)
from PyQt6.QtGui import QFont, QIcon
from Coding.code_editor import CodeEditor
from common_functions import code_editor_font


class PowerQueryTab(QWidget):
    """Tab that surfaces Power Query table metadata with editable details."""

    def __init__(self, pbip_file: Optional[str] = None):
        super().__init__()
        self.pbip_file = pbip_file
        self.tables_data: Dict[str, Dict[str, Optional[str]]] = {}
        self.query_order = []
        self.query_groups = {}
        self.current_table: Optional[str] = None
        self.ignore_editor_changes = False
        self.init_ui()
        if pbip_file:
            self.load_tables()

    def init_ui(self):
        """Build the user interface."""
        main_layout = QVBoxLayout(self)
        main_layout.setContentsMargins(8, 8, 8, 8)
        main_layout.setSpacing(8)

        top_bar = QHBoxLayout()
        self.refresh_button = QPushButton("Reload Metadata")
        self.refresh_button.setToolTip("Reload Power Query metadata from disk")
        self.refresh_button.clicked.connect(self.refresh_tables)
        top_bar.addWidget(self.refresh_button)

        self.open_pbip_button = QPushButton("Open PBIP...")
        self.open_pbip_button.setToolTip("Select a PBIP file to load Power Query metadata")
        self.open_pbip_button.clicked.connect(self.choose_pbip_file)
        top_bar.addWidget(self.open_pbip_button)

        top_bar.addStretch()
        main_layout.addLayout(top_bar)

        splitter = QSplitter(Qt.Orientation.Horizontal)

        left_widget = QWidget()
        left_layout = QVBoxLayout(left_widget)
        left_layout.setContentsMargins(0, 0, 6, 0)

        left_layout.addWidget(QLabel("Tables"))

        self.table_tree = QTreeWidget()
        self.table_tree.setHeaderHidden(True)
        self.table_tree.setSelectionMode(QAbstractItemView.SelectionMode.SingleSelection)
        self.table_tree.itemSelectionChanged.connect(self.on_tree_selection_changed)
        left_layout.addWidget(self.table_tree)

        icon_dir = os.path.abspath(os.path.join(os.path.dirname(os.path.dirname(__file__)), "Images", "icons"))

        def _load_icon(name: str) -> QIcon:
            path = os.path.join(icon_dir, name)
            return QIcon(path) if os.path.exists(path) else QIcon()

        self.table_icons = {
            "m": _load_icon("Table.svg"),
            "calculated": _load_icon("Calculated-Table.svg"),
        }

        splitter.addWidget(left_widget)

        right_widget = QWidget()
        right_layout = QVBoxLayout(right_widget)
        right_layout.setContentsMargins(6, 0, 0, 0)

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
        form_layout.setContentsMargins(0, 8, 0, 0)
        form_layout.setSpacing(6)

        self.import_mode_combo = QComboBox()
        self.import_mode_combo.setEnabled(False)
        self.import_mode_combo.addItem("Import", "import")
        self.import_mode_combo.addItem("DirectQuery", "directquery")
        self.import_mode_combo.currentIndexChanged.connect(self.on_import_mode_changed)
        form_layout.addRow("Import mode", self.import_mode_combo)

        right_layout.addLayout(form_layout)

        splitter.addWidget(right_widget)
        splitter.setSizes([250, 750])

        main_layout.addWidget(splitter)

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
        self.table_tree.blockSignals(True)
        self.table_tree.clear()

        if not self.tables_data:
            self.table_tree.blockSignals(False)
            return

        if self.query_order:
            sorted_tables = [name for name in self.query_order if name in self.tables_data]
            extras = sorted(name for name in self.tables_data if name not in self.query_order)
            sorted_tables.extend(extras)
        else:
            sorted_tables = sorted(self.tables_data.keys(), key=str.casefold)

        for table_name in sorted_tables:
            table_item = QTreeWidgetItem([table_name])
            table_item.setData(0, Qt.ItemDataRole.UserRole, table_name)
            table_type = (self.tables_data.get(table_name, {}).get("table_type") or "").lower()
            icon = self.table_icons.get(table_type)
            if icon and not icon.isNull():
                table_item.setIcon(0, icon)

            for column in self.tables_data.get(table_name, {}).get("columns", []):
                column_item = QTreeWidgetItem([column])
                column_item.setData(0, Qt.ItemDataRole.UserRole, table_name)
                table_item.addChild(column_item)

            table_item.setExpanded(False)
            self.table_tree.addTopLevelItem(table_item)

        self.table_tree.blockSignals(False)

        # Select the first table by default.
        first_item = self.table_tree.topLevelItem(0)
        if first_item:
            self.table_tree.setCurrentItem(first_item)
            self.display_table_details(first_item.data(0, Qt.ItemDataRole.UserRole))

    def on_tree_selection_changed(self):
        """Update detail pane when the selection changes."""
        current_item = self.table_tree.currentItem()
        if not current_item:
            self.display_table_details(None)
            return

        table_name = current_item.data(0, Qt.ItemDataRole.UserRole)
        if current_item.parent() is not None:
            # Ensure the parent table item holds the detail context.
            table_name = current_item.parent().data(0, Qt.ItemDataRole.UserRole)
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
        self.table_tree.clear()
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
