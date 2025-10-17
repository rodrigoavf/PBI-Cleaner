import os
import sys
import subprocess
from PyQt6.QtWidgets import (
    QApplication, QWidget, QLabel, QLineEdit, QPushButton,
    QFileDialog, QVBoxLayout, QHBoxLayout, QMessageBox, QTableWidget,
    QTableWidgetItem, QHeaderView, QAbstractItemView, QProgressBar, QHeaderView
)
from PyQt6.QtCore import (Qt, QTimer, QUrl)

valid_extensions = {".json", ".tmdl"}


def file_contains_target(path, target):
    try:
        with open(path, "r", encoding="utf-8") as f:
            content = f.read()
            return target in content, content.count(target)
    except Exception:
        return False, 0


def count_files(root_dir):
    total = 0
    for _, _, files in os.walk(root_dir):
        total += len(files)
    return total


def find_files_with_target(root_dir, target, before=20, after=20, update_callback=None):
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
                        idx = text.find(target)
                        if idx != -1:
                            start = max(0, idx - before)
                            end = min(len(text), idx + len(target) + after)
                            context_snippet = text[start:end].replace("\n", " ").replace("\r", " ")
                            count = text.count(target)
                            matches.append((os.path.basename(full_path), full_path, context_snippet, count))
                except Exception:
                    pass

            if update_callback and processed % 25 == 0:
                update_callback(processed, total_files)

    return matches

class FileSearchApp(QWidget):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("File Search Tool")
        self.setMinimumSize(1000, 600)
        self.sort_states = [None, None, None]
        self.last_clicked_row = None
        self.init_ui()

    def init_ui(self):
        # --- Directory selector (.pbip) ---
        self.dir_label = QLabel("Select .pbip file:")
        self.dir_input = QLineEdit("C:/Users/rodrigo.ferreira/Desktop/Devoteam/Operations.pbip")
        self.dir_input.setToolTip("Select a .pbip (Power BI Project) file")
        self.dir_button = QPushButton("Browse")
        self.dir_button.clicked.connect(self.browse_pbip_file)

        dir_layout = QHBoxLayout()
        dir_layout.addWidget(self.dir_label)
        dir_layout.addWidget(self.dir_input)
        dir_layout.addWidget(self.dir_button)

        # --- Target string ---
        self.target_label = QLabel("Target String:")
        self.target_input = QLineEdit("KPI-")
        self.target_input.setToolTip("Text to search for inside files")

        target_layout = QHBoxLayout()
        target_layout.addWidget(self.target_label)
        target_layout.addWidget(self.target_input)

        # --- Context range controls ---
        self.before_label = QLabel("Chars before:")
        self.before_input = QLineEdit("20")
        self.before_input.setFixedWidth(60)
        self.before_input.setToolTip("Number of characters before the target string to show in context preview")

        self.after_label = QLabel("Chars after:")
        self.after_input = QLineEdit("20")
        self.after_input.setFixedWidth(60)
        self.after_input.setToolTip("Number of characters after the target string to show in context preview")

        context_layout = QHBoxLayout()
        context_layout.addWidget(self.before_label)
        context_layout.addWidget(self.before_input)
        context_layout.addWidget(self.after_label)
        context_layout.addWidget(self.after_input)
        context_layout.addStretch()

        # --- Search button ---
        self.search_button = QPushButton("Search")
        self.search_button.clicked.connect(self.start_search)

        # --- Results table ---
        self.table = QTableWidget()
        self.table.setColumnCount(4)
        #self.table.setHorizontalHeaderLabels(["File Name", "Full Path", "Context", "Occurrences"])
        self.sort_states = [None] * self.table.columnCount()
        self.table.setSelectionBehavior(QAbstractItemView.SelectionBehavior.SelectRows)
        self.table.setEditTriggers(QAbstractItemView.EditTrigger.NoEditTriggers)
        self.table.setStyleSheet("QTableWidget::item:selected { background-color: #0078d4; }")
        self.table.horizontalHeader().sectionClicked.connect(self.sort_by_column)
        self.table.itemSelectionChanged.connect(self.update_buttons_state)
        self.table.viewport().installEventFilter(self)
        self.table.setHorizontalScrollMode(QAbstractItemView.ScrollMode.ScrollPerPixel)
        self.table.setWordWrap(False)
        #self.table.setStyleSheet("")

        header = self.table.horizontalHeader()
        header.setSectionResizeMode(0, QHeaderView.ResizeMode.Interactive)
        header.setSectionResizeMode(1, QHeaderView.ResizeMode.Stretch)
        header.setSectionResizeMode(2, QHeaderView.ResizeMode.Interactive)
        header.setStretchLastSection(False)

        # --- Action buttons ---
        self.open_button = QPushButton("📂 Open Folder")
        self.delete_button = QPushButton("❌ Delete File")
        self.open_button.setEnabled(False)
        self.delete_button.setEnabled(False)
        self.open_button.clicked.connect(self.open_selected_folder)
        self.delete_button.clicked.connect(self.delete_selected_file)

        actions_layout = QHBoxLayout()
        actions_layout.addWidget(self.open_button)
        actions_layout.addWidget(self.delete_button)
        actions_layout.addStretch()

        # --- Status + progress ---
        self.status = QLabel("No results yet.")
        self.progress_bar = QProgressBar()
        self.progress_bar.setVisible(False)
        self.progress_bar.setMaximum(100)

        progress_layout = QHBoxLayout()
        progress_layout.addWidget(self.status)
        progress_layout.addWidget(self.progress_bar)

        credit_label = QLabel('<a href="https://www.linkedin.com/in/rodrigoavf/">Created by Rodrigo Ferreira</a>')
        credit_label.setTextFormat(Qt.TextFormat.RichText)
        credit_label.setTextInteractionFlags(Qt.TextInteractionFlag.TextBrowserInteraction)
        credit_label.setOpenExternalLinks(True)
        credit_label.setAlignment(Qt.AlignmentFlag.AlignRight)
        credit_label.setStyleSheet("color: gray; font-size: 10pt; margin-right: 8px;")

        # --- Layout setup ---
        layout = QVBoxLayout()
        layout.addLayout(dir_layout)
        layout.addLayout(target_layout)
        layout.addLayout(context_layout)
        layout.addWidget(self.search_button)
        layout.addLayout(actions_layout)
        layout.addWidget(self.table)
        layout.addLayout(progress_layout)
        layout.addWidget(credit_label)
        self.setLayout(layout)

    # ---------- File selection ----------
    def browse_pbip_file(self):
        file_path, _ = QFileDialog.getOpenFileName(
            self,
            "Select Power BI Project File (.pbip)",
            "",
            "Power BI Project (*.pbip);;All Files (*.*)",
        )

        # user canceled dialog → do nothing
        if not file_path:
            return

        if not file_path.lower().endswith(".pbip"):
            QMessageBox.critical(self, "Invalid File", "The selected file is not a .pbip file.")
            self.dir_input.clear()
            return

        self.dir_input.setText(file_path)

    # ---------- Search ----------
    def start_search(self):
        file_path = self.dir_input.text().strip()
        target = self.target_input.text().strip()

        try:
            chars_before = int(self.before_input.text())
            chars_after = int(self.after_input.text())
        except ValueError:
            QMessageBox.warning(self, "Invalid Input", "Before/After values must be integers.")
            return

        if not file_path or not os.path.isfile(file_path):
            QMessageBox.warning(self, "Missing File", "Please select a valid .pbip file first.")
            return

        if not file_path.lower().endswith(".pbip"):
            QMessageBox.critical(self, "Invalid File", "You must select a .pbip file.")
            return

        root_dir = os.path.dirname(file_path)

        if not target:
            QMessageBox.warning(self, "Missing Target", "Please enter a target string.")
            return

        # Disable UI while searching
        self.search_button.setEnabled(False)
        self.open_button.setEnabled(False)
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

        hits = find_files_with_target(root_dir, target, chars_before, chars_after, update_callback=update_progress)

        self.progress_bar.setVisible(False)
        self.search_button.setEnabled(True)

        if not hits:
            self.status.setText("❌ No results found.")
            QMessageBox.information(self, "No Results", "No files found containing the target string.")
            return

        self.populate_table(hits)

    def populate_table(self, hits):
        self.table.setRowCount(len(hits))

        # x[2] is now context (string), so total_refs should sum x[3]
        total_refs = sum(int(x[3]) for x in hits)
        total_files = len(hits)

        for row, (fname, path, context, count) in enumerate(hits):
            name_item = QTableWidgetItem(fname)
            path_item = QTableWidgetItem(path)
            #path_item.setTextAlignment(Qt.AlignmentFlag.AlignLeft)
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
        """Ensure columns fit perfectly after table is populated."""
        QTimer.singleShot(50, self._resize_columns_now)  # small delay for layout to settle

    def _resize_columns_now(self):
        header = self.table.horizontalHeader()

        # Make sure headers are correctly labeled
        if self.table.columnCount() == 4:
            self.table.setHorizontalHeaderLabels(["File Name", "Full Path", "Context", "Occurrences"])

        # Let Qt calculate minimum content widths
        self.table.resizeColumnsToContents()

        # Fixed columns
        name_width = max(200, self.table.columnWidth(0))
        count_width = max(100, self.table.columnWidth(3))

        # Available space for the 2 middle columns (minus small margin)
        vp_width = self.table.viewport().width() - self.table.verticalScrollBar().width() - 10
        remaining = max(0, vp_width - (name_width + count_width))

        # Ratio for Full Path vs Context (Full Path gets more space)
        hint_path = max(1, self.table.sizeHintForColumn(1))
        hint_context = max(1, self.table.sizeHintForColumn(2))
        ratio_path = hint_path * 1.5
        ratio_context = hint_context
        total = ratio_path + ratio_context

        path_width = int(remaining * (ratio_path / total))
        context_width = remaining - path_width

        # Apply final widths (slightly shrink total to prevent scroll)
        header.resizeSection(0, name_width)
        header.resizeSection(1, max(200, path_width))
        header.resizeSection(2, max(150, context_width))
        header.resizeSection(3, count_width)

        # Allow user resizing for all
        for i in range(4):
            header.setSectionResizeMode(i, QHeaderView.ResizeMode.Interactive)

        header.setStretchLastSection(False)

    def resizeEvent(self, event):
        """Recalculate widths whenever the window is resized."""
        super().resizeEvent(event)
        QTimer.singleShot(0, self._resize_columns_now)

    def get_selected_file(self):
        selected = self.table.currentRow()
        if selected == -1:
            return None
        return self.table.item(selected, 1).text()

    def update_buttons_state(self):
        has_selection = self.table.currentRow() != -1
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

    def delete_selected_file(self):
        path = self.get_selected_file()
        if not path:
            return

        reply = QMessageBox.question(
            self,
            "Delete Confirmation",
            f"Are you sure you want to delete:\n{path}?",
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
        )

        if reply == QMessageBox.StandardButton.Yes:
            try:
                os.remove(path)
                self.table.removeRow(self.table.currentRow())
                QMessageBox.information(self, "Deleted", f"File deleted:\n{path}")
                self.update_status_after_deletion()
            except Exception as e:
                QMessageBox.critical(self, "Error", f"Failed to delete file:\n{e}")

    def update_status_after_deletion(self):
        rows = self.table.rowCount()
        refs = sum(int(self.table.item(r, 2).text()) for r in range(rows)) if rows else 0
        self.status.setText(f"✅ Remaining {refs} references across {rows} files." if rows else "No results left.")

    def sort_by_column(self, column):
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

    # ---------- Custom row deselection ----------
    def eventFilter(self, source, event):
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


if __name__ == "__main__":
    app = QApplication([])
    window = FileSearchApp()
    window.show()
    app.exec()
