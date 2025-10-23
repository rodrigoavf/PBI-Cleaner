import os
import sys
import subprocess
from PyQt6.QtWidgets import (
    QApplication, QWidget, QLabel, QLineEdit, QPushButton, QVBoxLayout, QHBoxLayout,
    QMessageBox, QTableWidget, QTableWidgetItem, QHeaderView, QAbstractItemView, QProgressBar,
    QSizePolicy, QScrollArea, QCheckBox
)
from PyQt6.QtCore import Qt, QTimer
from PyQt6.QtGui import (
    QFont, QIcon, QDragEnterEvent, QDropEvent, QShortcut, QKeySequence
)
from common_functions import count_files

valid_extensions = {".json", ".tmdl"}

def find_files_with_target(root_dir, target, before=20, after=20, update_callback=None,
                           case_sensitive=False, full_match=False):
    import re
    matches = []
    processed = 0
    total_files = count_files(root_dir)

    for subdir, _, files in os.walk(root_dir):
        for fname in files:
            processed += 1
            ext = os.path.splitext(fname)[1].lower()
            if ext in valid_extensions:
                full_path = os.path.join(subdir, fname)
                try:
                    with open(full_path, "r", encoding="utf-8") as f:
                        text = f.read()
                        flags = 0 if case_sensitive else re.IGNORECASE
                        if full_match:
                            pattern = r"(?<![A-Za-z0-9_])" + re.escape(target) + r"(?![A-Za-z0-9_])"
                        else:
                            pattern = re.escape(target)

                        it = list(re.finditer(pattern, text, flags))
                        if it:
                            first = it[0]
                            idx = first.start()
                            start = max(0, idx - before)
                            end = min(len(text), idx + len(target) + after)
                            context_snippet = text[start:end].replace("\n", " ").replace("\r", " ")
                            count = len(it)
                            matches.append((os.path.basename(full_path), full_path, context_snippet, count))
                except Exception:
                    pass

            if update_callback and processed % 25 == 0:
                update_callback(processed, total_files)

    return matches

class FileSearchApp(QWidget):
    def __init__(self, pbip_file: str = None):
        super().__init__()
        self.setMinimumWidth(300)
        self.sort_states = [None, None, None, None]
        self.last_clicked_row = None
        self.pbip_file = pbip_file
        self.init_ui()

    def init_ui(self):
        self.target_label = QLabel("Target String:")
        self.target_input = QLineEdit("KPI-")

        target_layout = QHBoxLayout()
        target_layout.addWidget(self.target_label)
        target_layout.addWidget(self.target_input)

        # Search options
        self.case_sensitive_cb = QCheckBox("Case sensitive")
        self.exact_match_cb = QCheckBox("Full match only")
        self.exact_match_cb.setToolTip("If enabled, match whole tokens only (not substrings)")

        # options will be placed inline with context inputs

        self.before_label = QLabel("Chars before:")
        self.before_input = QLineEdit("20")
        self.before_input.setFixedWidth(60)
        self.after_label = QLabel("Chars after:")
        self.after_input = QLineEdit("20")
        self.after_input.setFixedWidth(60)

        context_layout = QHBoxLayout()
        context_layout.addWidget(self.before_label)
        context_layout.addWidget(self.before_input)
        context_layout.addWidget(self.after_label)
        context_layout.addWidget(self.after_input)

        self.search_button = QPushButton("🔍 Search")
        self.search_button.clicked.connect(self.start_search)

        # Place options and search button inline with context
        self.search_button.setText("🔍 Search")
        context_layout.addSpacing(12)
        context_layout.addWidget(self.case_sensitive_cb)
        context_layout.addWidget(self.exact_match_cb)
        context_layout.addSpacing(12)
        context_layout.addWidget(self.search_button)
        context_layout.addStretch()

        self.table = QTableWidget()
        self.table.setColumnCount(4)
        self.table.setSelectionBehavior(QAbstractItemView.SelectionBehavior.SelectRows)
        self.table.setSelectionMode(QAbstractItemView.SelectionMode.SingleSelection)
        self.table.setEditTriggers(QAbstractItemView.EditTrigger.NoEditTriggers)
        self.table.setHorizontalScrollMode(QAbstractItemView.ScrollMode.ScrollPerPixel)
        self.table.setWordWrap(False)
        self.table.setMinimumHeight(0)
        self.table.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.MinimumExpanding)
        self.table.horizontalHeader().sectionClicked.connect(self.sort_by_column)
        self.table.itemSelectionChanged.connect(self.update_buttons_state)
        self.table.viewport().installEventFilter(self)

        header = self.table.horizontalHeader()
        header.setSectionResizeMode(0, QHeaderView.ResizeMode.Interactive)
        header.setSectionResizeMode(1, QHeaderView.ResizeMode.Stretch)
        header.setSectionResizeMode(2, QHeaderView.ResizeMode.Interactive)
        header.setStretchLastSection(False)

        self.open_button = QPushButton("📂 Open Folder")
        self.delete_button = QPushButton("❌ Delete File")
        self.open_button.setEnabled(False)
        self.delete_button.setEnabled(False)
        self.open_button.clicked.connect(self.open_selected_folder)
        self.delete_button.clicked.connect(self.delete_selected_file)
        # New: open file action
        self.open_file_button = QPushButton("📄 Open File")
        self.open_file_button.setEnabled(False)
        self.open_file_button.clicked.connect(self.open_selected_file)

        actions_layout = QHBoxLayout()
        actions_layout.addWidget(self.open_file_button)
        actions_layout.addWidget(self.open_button)
        actions_layout.addWidget(self.delete_button)
        actions_layout.addStretch()

        self.status = QLabel("No results yet.")
        self.status.setMinimumHeight(24)
        self.status.setSizePolicy(QSizePolicy.Policy.Preferred, QSizePolicy.Policy.Fixed)
        self.status.setAlignment(Qt.AlignmentFlag.AlignLeft | Qt.AlignmentFlag.AlignVCenter)
        self.progress_bar = QProgressBar()
        self.progress_bar.setVisible(False)
        self.progress_bar.setMaximum(100)
        self.progress_bar.setFixedHeight(16)

        footer_widget = QWidget()
        footer_widget.setSizePolicy(QSizePolicy.Policy.Preferred, QSizePolicy.Policy.Fixed)
        footer_widget.setFixedHeight(30)

        progress_layout = QHBoxLayout(footer_widget)
        progress_layout.setContentsMargins(0, 4, 0, 4)
        progress_layout.setSpacing(8)
        progress_layout.addWidget(self.status)
        progress_layout.addWidget(self.progress_bar)

        # Build scrollable content section (everything above footer)
        content_widget = QWidget()
        content_layout = QVBoxLayout(content_widget)
        content_layout.setContentsMargins(0, 0, 0, 0)
        content_layout.setSpacing(8)
        content_layout.addLayout(target_layout)
        content_layout.addLayout(context_layout)
        content_layout.addLayout(actions_layout)
        content_layout.addWidget(self.table, 1)

        scroll_area = QScrollArea()
        scroll_area.setWidgetResizable(True)
        scroll_area.setWidget(content_widget)

        main_layout = QVBoxLayout()
        main_layout.setContentsMargins(8, 8, 8, 8)
        main_layout.setSpacing(8)
        main_layout.addWidget(scroll_area, 1)
        main_layout.addWidget(footer_widget, alignment=Qt.AlignmentFlag.AlignBottom)

        self.setLayout(main_layout)
        self.setup_shortcuts()

    def start_search(self):
        if not self.pbip_file or not os.path.isfile(self.pbip_file):
            QMessageBox.warning(self, "Missing File", "No .pbip file selected.")
            return

        target = self.target_input.text().strip()
        if not target:
            QMessageBox.warning(self, "Missing Target", "Please enter a target string.")
            return

        try:
            chars_before = int(self.before_input.text())
            chars_after = int(self.after_input.text())
        except ValueError:
            QMessageBox.warning(self, "Invalid Input", "Before/After values must be integers.")
            return

        root_dir = os.path.dirname(self.pbip_file)

        self.search_button.setEnabled(False)
        self.open_button.setEnabled(False)
        self.open_file_button.setEnabled(False)
        self.delete_button.setEnabled(False)
        self.status.setText("🔍 Searching... please wait.")
        self.progress_bar.setVisible(True)
        self.progress_bar.setValue(0)
        self.table.setRowCount(0)
        QApplication.processEvents()

        def update_progress(processed, total):
            pct = int(processed / total * 100)
            self.progress_bar.setValue(pct)
            self.status.setText(f"Searching... {processed}/{total} files scanned")
            QApplication.processEvents()

        hits = find_files_with_target(
            root_dir,
            target,
            chars_before,
            chars_after,
            update_callback=update_progress,
            case_sensitive=self.case_sensitive_cb.isChecked(),
            full_match=self.exact_match_cb.isChecked(),
        )

        self.progress_bar.setVisible(False)
        self.search_button.setEnabled(True)

        if not hits:
            self.status.setText("❌ No results found.")
            QMessageBox.information(self, "No Results", "No files found containing the target string.")
            return

        self.populate_table(hits)

    def setup_shortcuts(self):
        # Line edits: use their native signal
        self.target_input.returnPressed.connect(self.start_search)
        self.before_input.returnPressed.connect(self.start_search)
        self.after_input.returnPressed.connect(self.start_search)

        # Delete anywhere in the window
        self.delete_button_shortcut = QShortcut(QKeySequence(Qt.Key.Key_Delete), self.table)
        self.delete_button_shortcut.setContext(Qt.ShortcutContext.WindowShortcut)
        self.delete_button_shortcut.activated.connect(self.delete_selected_file)

        # Enter in the table opens the selected file
        self.open_button_shortcut = QShortcut(QKeySequence(Qt.Key.Key_Return), self.table)
        self.open_button_shortcut.setContext(Qt.ShortcutContext.WidgetShortcut)
        self.open_button_shortcut.activated.connect(self.open_selected_file)
        # Support keypad Enter as well
        self.open_button_shortcut2 = QShortcut(QKeySequence(Qt.Key.Key_Enter), self.table)
        self.open_button_shortcut2.setContext(Qt.ShortcutContext.WidgetShortcut)
        self.open_button_shortcut2.activated.connect(self.open_selected_file)

    def populate_table(self, hits):
        self.table.setRowCount(len(hits))
        total_refs = sum(int(x[3]) for x in hits)
        total_files = len(hits)

        for row, (fname, path, context, count) in enumerate(hits):
            name_item = QTableWidgetItem(fname)
            path_item = QTableWidgetItem(path)
            path_item.setTextAlignment(Qt.AlignmentFlag.AlignLeft | Qt.AlignmentFlag.AlignVCenter)
            path_item.setToolTip(path)
            context_item = QTableWidgetItem(context)
            context_item.setToolTip(context)
            count_item = QTableWidgetItem(str(count))
            count_item.setTextAlignment(Qt.AlignmentFlag.AlignCenter)

            self.table.setItem(row, 0, name_item)
            self.table.setItem(row, 1, path_item)
            self.table.setItem(row, 2, context_item)
            self.table.setItem(row, 3, count_item)

        self.auto_adjust_columns()
        self.status.setText(f"✅ Found {total_refs} references across {total_files} files.")
        self.update_buttons_state()

    def auto_adjust_columns(self):
        QTimer.singleShot(50, self._resize_columns_now)

    def _resize_columns_now(self):
        header = self.table.horizontalHeader()
        if self.table.columnCount() == 4:
            self.table.setHorizontalHeaderLabels(["File Name", "Full Path", "Context", "Occurrences"])
        self.table.resizeColumnsToContents()

        name_width = max(200, self.table.columnWidth(0))
        count_width = max(100, self.table.columnWidth(3))
        vp_width = self.table.viewport().width() - self.table.verticalScrollBar().width() - 10
        remaining = max(0, vp_width - (name_width + count_width))

        hint_path = max(1, self.table.sizeHintForColumn(1))
        hint_context = max(1, self.table.sizeHintForColumn(2))
        ratio_path = hint_path * 1.5
        ratio_context = hint_context
        total = ratio_path + ratio_context

        path_width = int(remaining * (ratio_path / total))
        context_width = remaining - path_width

        header.resizeSection(0, name_width)
        header.resizeSection(1, max(200, path_width))
        header.resizeSection(2, max(150, context_width))
        header.resizeSection(3, count_width)

        for i in range(4):
            header.setSectionResizeMode(i, QHeaderView.ResizeMode.Interactive)

        header.setStretchLastSection(False)

    def resizeEvent(self, event):
        super().resizeEvent(event)
        QTimer.singleShot(0, self._resize_columns_now)

    def get_selected_file(self):
        selected = self.table.currentRow()
        if selected == -1:
            return None
        return self.table.item(selected, 1).text()

    def update_buttons_state(self):
        has_selection = self.table.currentRow() != -1
        self.open_file_button.setEnabled(has_selection)
        self.open_button.setEnabled(has_selection)
        self.delete_button.setEnabled(has_selection)

    def open_selected_folder(self):
        path = self.get_selected_file()
        if not path:
            return

        folder = os.path.dirname(path)
        try:
            if sys.platform.startswith("win"):
                os.startfile(folder)
            elif sys.platform.startswith("darwin"):
                subprocess.Popen(["open", folder])
            else:
                subprocess.Popen(["xdg-open", folder])
        except Exception as e:
            QMessageBox.critical(self, "Error", f"Failed to open folder:\n{e}")

    def open_selected_file(self):
        path = self.get_selected_file()
        if not path:
            return

        try:
            if sys.platform.startswith("win"):
                os.startfile(path)
            elif sys.platform.startswith("darwin"):
                subprocess.Popen(["open", path])
            else:
                subprocess.Popen(["xdg-open", path])
        except Exception as e:
            QMessageBox.critical(self, "Error", f"Failed to open file:\n{e}")

    def delete_selected_file(self):
        path = self.get_selected_file()
        if not path:
            return

        reply = QMessageBox.question(
            self,
            "Delete Confirmation",
            f"Are you sure you want to delete:\n{path}?\n\n⚠️ This action cannot be undone. Use at your own risk!",
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
        )

        if reply == QMessageBox.StandardButton.Yes:
            try:
                os.remove(path)
                self.table.removeRow(self.table.currentRow())
                QMessageBox.information(self, "Deleted", f"File deleted:\n{path}")
            except Exception as e:
                QMessageBox.critical(self, "Error", f"Failed to delete file:\n{e}")

    def sort_by_column(self, column):
        if len(self.sort_states) < self.table.columnCount():
            self.sort_states = [None] * self.table.columnCount()

        if self.sort_states[column] is None:
            order = Qt.SortOrder.AscendingOrder
            self.sort_states[column] = "asc"
        elif self.sort_states[column] == "asc":
            order = Qt.SortOrder.DescendingOrder
            self.sort_states[column] = "desc"
        else:
            order = None
            self.sort_states[column] = None

        for i in range(self.table.columnCount()):
            if i != column:
                self.sort_states[i] = None

        if order is None:
            self.table.setSortingEnabled(False)
        else:
            self.table.setSortingEnabled(True)
            self.table.sortItems(column, order)

    def eventFilter(self, source, event):
        # Handle double-click explicitly to open file
        if source is self.table.viewport() and event.type() == event.Type.MouseButtonDblClick:
            index = self.table.indexAt(event.pos())
            if index.isValid():
                self.open_selected_file()
                return True

        if source is self.table.viewport() and event.type() == event.Type.MouseButtonPress:
            index = self.table.indexAt(event.pos())
            if index.isValid():
                row = index.row()
                if self.last_clicked_row == row:
                    self.table.clearSelection()
                    self.last_clicked_row = None
                    return True
                self.last_clicked_row = row
        return super().eventFilter(source, event)
