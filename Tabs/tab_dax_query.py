import os
import json
from PyQt6.QtWidgets import (
    QWidget, QVBoxLayout, QHBoxLayout, QPushButton, QListWidget,
    QListWidgetItem, QTextEdit, QSplitter, QMessageBox, QLabel
)
from PyQt6.QtCore import Qt, QMimeData
from PyQt6.QtGui import QFont, QIcon, QDragEnterEvent, QDropEvent
from PyQt6.QtCore import QStringListModel
from PyQt6.QtWidgets import QCompleter
from DAX.dax_editor_support import DAXHighlighter, DAX_KEYWORDS, DAX_FUNCTIONS
from DAX.qcode_editor import QCodeEditor

class DAXQueryTab(QWidget):
    def __init__(self, pbip_file: str = None):
        super().__init__()
        self.pbip_file = pbip_file
        self.default_query = None
        self.queries = {}
        self.init_ui()
        if pbip_file:
            self.load_queries()

    def init_ui(self):
        # Main layout
        main_layout = QVBoxLayout(self)
        main_layout.setContentsMargins(8, 8, 8, 8)
        main_layout.setSpacing(8)

        # Save button at the top
        top_bar = QHBoxLayout()
        self.save_button = QPushButton("💾 Save Changes")
        self.save_button.clicked.connect(self.save_changes)
        self.save_button.setEnabled(False)
        top_bar.addWidget(self.save_button)
        # Refresh button to discard unsaved changes by reloading from disk
        self.refresh_button = QPushButton("⟳ Undo All Changes")
        self.refresh_button.setToolTip("Reload queries from disk (undo unsaved changes)")
        self.refresh_button.clicked.connect(self.refresh_queries)
        top_bar.addWidget(self.refresh_button)
        top_bar.addStretch()
        main_layout.addLayout(top_bar)

        # Create splitter for left and right sides
        splitter = QSplitter(Qt.Orientation.Horizontal)
        
        # Left side - Query list
        left_widget = QWidget()
        left_layout = QVBoxLayout(left_widget)
        left_layout.setContentsMargins(0, 0, 0, 0)

        # Query list with drag-drop support
        self.query_list = QListWidget()
        self.query_list.setDragEnabled(True)
        self.query_list.setAcceptDrops(True)
        self.query_list.setDropIndicatorShown(True)
        self.query_list.setDragDropMode(QListWidget.DragDropMode.InternalMove)
        self.query_list.currentItemChanged.connect(self.on_query_selected)
        # Double-click to set as default
        self.query_list.itemDoubleClicked.connect(self.on_query_double_clicked)
        # Enable Save when items are reordered via drag-and-drop
        try:
            self.query_list.model().rowsMoved.connect(self.on_query_order_changed)
        except Exception:
            pass
        left_layout.addWidget(QLabel("DAX Queries:"))
        left_layout.addWidget(self.query_list)

        # Make Default button
        self.default_btn = QPushButton("⭐ Set as default")
        self.default_btn.clicked.connect(self.make_default)
        self.default_btn.setEnabled(False)
        left_layout.addWidget(self.default_btn)

        # Right side - Query editor
        right_widget = QWidget()
        right_layout = QVBoxLayout(right_widget)
        right_layout.setContentsMargins(0, 0, 0, 0)

        # Query editor
        right_layout.addWidget(QLabel("Query DAX Code:"))
        self.query_editor = QCodeEditor()
        self.query_editor.setFont(QFont("Consolas", 10))
        # Make tab width smaller (2 spaces worth)
        try:
            space_w = self.query_editor.fontMetrics().horizontalAdvance(' ')
            self.query_editor.setTabStopDistance(space_w * 4)
        except Exception:
            pass
        self.query_editor.textChanged.connect(self.on_text_changed)
        right_layout.addWidget(self.query_editor)

        # Attach DAX syntax highlighting and IntelliSense
        try:
            self._dax_highlighter = DAXHighlighter(self.query_editor.document())
            self._dax_highlighter.rehighlight()
        except Exception:
            self._dax_highlighter = None
        try:
            # Build a dedicated DAX completer
            completions = sorted(set([*DAX_KEYWORDS, *DAX_FUNCTIONS]))
            # Parent the model to the editor to keep it alive
            model = QStringListModel(completions, self.query_editor)
            completer = QCompleter(model, self.query_editor)
            completer.setCaseSensitivity(Qt.CaseSensitivity.CaseInsensitive)
            completer.setFilterMode(Qt.MatchFlag.MatchContains)
            completer.setCompletionMode(QCompleter.CompletionMode.PopupCompletion)
            try:
                completer.setModelSorting(QCompleter.ModelSorting.CaseInsensitivelySortedModel)
            except Exception:
                pass
            self.query_editor.setCompleter(completer)
            self._dax_completer = completer
            self._dax_model = model
        except Exception:
            self._dax_completer = None

        # Add widgets to splitter
        splitter.addWidget(left_widget)
        splitter.addWidget(right_widget)
        
        # Set initial sizes
        splitter.setSizes([200, 800])
        
        main_layout.addWidget(splitter)

    def refresh_queries(self):
        """Reload queries from disk, discarding unsaved changes."""
        self.load_queries()
        self.save_button.setEnabled(False)

    def on_query_double_clicked(self, item: QListWidgetItem):
        """When a query is double-clicked, set it as default."""
        if item is not None:
            self.query_list.setCurrentItem(item)
        self.make_default()

    def load_queries(self):
        """Load DAX queries from the PBIP file"""
        if not self.pbip_file or not os.path.isfile(self.pbip_file):
            return

        root_dir = os.path.splitext(self.pbip_file)[0] + ".SemanticModel/DAXQueries"
        json_path = root_dir + "/.pbi/daxQueries.json"
        
        try:
            with open(json_path, "r", encoding="utf-8") as f:
                data = json.load(f)

                queries = {}
                for query in data["tabOrder"]:
                    with open(root_dir + "/" + query + ".dax", "r", encoding="utf-8") as f:
                        dax_code = f.read()
                    queries[query] = dax_code

                self.queries = queries
                self.default_query = data["defaultTab"]
                
                # Populate the list widget
                self.query_list.clear()
                for name in self.queries.keys():
                    item = QListWidgetItem(name)
                    if name == self.default_query:
                        item.setText(f"{name} ⭐")
                    self.query_list.addItem(item)
                
                if self.query_list.count() > 0:
                    self.query_list.setCurrentRow(0)

        except Exception as e:
            QMessageBox.warning(self, "Error", f"Failed to load queries:\n{str(e)}")

    def on_query_selected(self, current, previous):
        """Handle query selection change"""
        if not current:
            self.query_editor.clear()
            self.query_editor.setEnabled(False)
            self.default_btn.setEnabled(False)
            return

        query_name = current.text().replace(" ⭐", "")
        self.query_editor.setEnabled(True)
        self.query_editor.setPlainText(self.queries.get(query_name, ""))
        
        # Enable/disable Make Default button
        is_default = query_name == self.default_query
        self.default_btn.setEnabled(not is_default)

    def on_text_changed(self):
        """Handle query text changes"""
        current = self.query_list.currentItem()
        if not current:
            return

        query_name = current.text().replace(" ⭐", "")
        new_text = self.query_editor.toPlainText()
        
        if new_text != self.queries.get(query_name):
            self.queries[query_name] = new_text
            self.save_button.setEnabled(True)

    def on_query_order_changed(self, *args, **kwargs):
        """Enable Save when the list order changes."""
        self.save_button.setEnabled(True)

    def make_default(self):
        """Make the selected query the default"""
        current = self.query_list.currentItem()
        if not current:
            return

        # Update the icons and text
        for i in range(self.query_list.count()):
            item = self.query_list.item(i)
            name = item.text().replace(" ⭐", "")
            item.setIcon(QIcon())
            item.setText(name)

        # Set the new default
        query_name = current.text()
        self.default_query = query_name
        current.setIcon(QIcon.fromTheme("star"))
        current.setText(f"{query_name} ⭐")
        
        self.default_btn.setEnabled(False)
        self.save_button.setEnabled(True)

    def save_changes(self):
        """Save changes to the queries.json file"""
        if not self.pbip_file:
            return

        root_dir = os.path.splitext(self.pbip_file)[0] + ".SemanticModel/DAXQueries"
        json_path = root_dir + "/.pbi/daxQueries.json"
        
        try:
            new_query_order = []
            new_default_query = None

            for i in range(self.query_list.count()):
                item_text = self.query_list.item(i).text().replace(" ⭐", "")
                new_query_order.append(item_text)
                if self.query_list.item(i).text().endswith("⭐"):
                    new_default_query = item_text

            # Read the existing JSON data
            with open(json_path, "r", encoding="utf-8") as f:
                data = json.load(f)

            # Update only the relevant keys
            data["tabOrder"] = new_query_order
            data["defaultTab"] = new_default_query

            # Write the updated JSON back
            with open(json_path, "w", encoding="utf-8") as f:
                json.dump(data, f, indent=4, ensure_ascii=False)

            # Write back individual .dax files
            for name, code in self.queries.items():
                with open(root_dir + "/" + name + ".dax", "w", encoding="utf-8") as f:
                    f.write(code)
            
            self.save_button.setEnabled(False)
            QMessageBox.information(self, "Success", "Changes saved successfully!")
            
        except Exception as e:
            QMessageBox.critical(self, "Error", f"Failed to save changes:\n{str(e)}")
