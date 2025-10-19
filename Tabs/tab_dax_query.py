import os
import json
from PyQt6.QtWidgets import (
    QWidget, QVBoxLayout, QHBoxLayout, QPushButton, QListWidget,
    QListWidgetItem, QSplitter, QMessageBox, QLabel, QMenu, QAbstractItemView
)
from PyQt6.QtCore import Qt, QPoint
from PyQt6.QtGui import (
    QFont, QIcon, QDragEnterEvent, QDropEvent, QShortcut, QKeySequence
)
from DAX.qcode_editor import QCodeEditor


INVALID_FILENAME_CHARS = set('<>:"/\\|?*')


class DAXQueryTab(QWidget):
    def __init__(self, pbip_file: str = None):
        super().__init__()
        self.pbip_file = pbip_file
        self.default_query = None
        self.queries = {}
        self.ignore_item_changes = False
        self.renaming_item = None
        self.renaming_original_name = ""
        self.ignore_editor_changes = False
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
        self.refresh_button = QPushButton("🔄 Reload queries")
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
        left_layout.setContentsMargins(0, 0, 6, 0)

        # Query list with drag-drop support
        self.query_list = QListWidget()
        self.query_list.setSelectionMode(QAbstractItemView.SelectionMode.ExtendedSelection)
        self.query_list.setDragEnabled(True)
        self.query_list.setAcceptDrops(True)
        self.query_list.setDropIndicatorShown(True)
        self.query_list.setDragDropMode(QListWidget.DragDropMode.InternalMove)
        self.query_list.setEditTriggers(QAbstractItemView.EditTrigger.NoEditTriggers)
        self.query_list.setContextMenuPolicy(Qt.ContextMenuPolicy.CustomContextMenu)
        self.query_list.customContextMenuRequested.connect(self.show_query_context_menu)
        self.query_list.itemSelectionChanged.connect(self.on_selection_changed)
        self.query_list.itemChanged.connect(self.on_query_item_changed)
        try:
            self.query_list.itemDelegate().closeEditor.connect(self.on_item_editor_closed)
        except Exception:
            pass
        # Double-click to set as default
        self.query_list.itemDoubleClicked.connect(self.on_query_double_clicked)
        # Enable Save when items are reordered via drag-and-drop
        try:
            self.query_list.model().rowsMoved.connect(self.on_query_order_changed)
        except Exception:
            pass
        left_layout.addWidget(QLabel("DAX Queries"))
        left_layout.addWidget(self.query_list)

        # Right side - Query editor
        right_widget = QWidget()
        right_layout = QVBoxLayout(right_widget)
        right_layout.setContentsMargins(6, 0, 0, 0)

        # Query editor
        right_layout.addWidget(QLabel("DAX Query Code"))
        self.query_editor = QCodeEditor(language='dax')
        self.query_editor.setFont(QFont("Consolas", 10))
        self.query_editor.setEnabled(False)
        try:
            space_w = self.query_editor.fontMetrics().horizontalAdvance(' ')
            self.query_editor.setTabStopDistance(space_w * 4)
        except Exception:
            pass
        self.query_editor.textChanged.connect(self.on_text_changed)
        right_layout.addWidget(self.query_editor)

        self.multi_selection_label = QLabel("Multiple queries selected. Editor is disabled.")
        self.multi_selection_label.setStyleSheet("color: #aa0000; font-style: italic;")
        self.multi_selection_label.setWordWrap(True)
        self.multi_selection_label.setVisible(False)
        right_layout.addWidget(self.multi_selection_label)

        # Hotkey hints - right
        hotkey_hint_right = QLabel("Ctrl+K+C: Comment   |   Ctrl+K+U: Uncomment   |   Alt+Up: Move line up   |   Alt+Down: Move line down")
        hotkey_hint_right.setStyleSheet("color: #666666; font-size: 10px;")
        hotkey_hint_right.setWordWrap(True)
        right_layout.addWidget(hotkey_hint_right)

        # Hotkey hints - left
        hotkey_hint_left = QLabel("F2: Rename\nDelete: Delete selected\nCtrl+N: New query\nAlt+Up: Move up\nAlt+Down: Move down")
        hotkey_hint_left.setStyleSheet("color: #666666; font-size: 10px;")
        hotkey_hint_left.setWordWrap(True)
        left_layout.addWidget(hotkey_hint_left)

        # Add widgets to splitter
        splitter.addWidget(left_widget)
        splitter.addWidget(right_widget)

        # Set initial sizes
        splitter.setSizes([200, 800])

        main_layout.addWidget(splitter)

        self.setup_shortcuts()

    def setup_shortcuts(self):
        """Register keyboard shortcuts for query list actions."""
        self.set_default_shortcut = QShortcut(QKeySequence(Qt.Key.Key_Space), self.query_list)
        self.set_default_shortcut.activated.connect(self.make_default)

        self.rename_shortcut = QShortcut(QKeySequence(Qt.Key.Key_F2), self)
        self.rename_shortcut.activated.connect(self.rename_selected_query)

        self.delete_shortcut = QShortcut(QKeySequence(Qt.Key.Key_Delete), self.query_list)
        self.delete_shortcut.activated.connect(self.delete_selected_queries)

        self.new_query_shortcut = QShortcut(QKeySequence("Ctrl+N"), self)
        self.new_query_shortcut.activated.connect(self.add_new_query)

        self.move_up_shortcut = QShortcut(QKeySequence("Alt+Up"), self.query_list)
        self.move_up_shortcut.activated.connect(lambda: self.move_selected_queries(-1))

        self.move_down_shortcut = QShortcut(QKeySequence("Alt+Down"), self.query_list)
        self.move_down_shortcut.activated.connect(lambda: self.move_selected_queries(1))

    def refresh_queries(self):
        """Reload queries from disk, discarding unsaved changes."""
        self.load_queries()
        self.save_button.setEnabled(False)

    def show_query_context_menu(self, position: QPoint):
        """Show context menu for the query list."""
        menu = QMenu(self.query_list)

        set_default_action = menu.addAction("Set as Default")
        rename_action = menu.addAction("Rename")
        delete_action = menu.addAction("Delete")
        menu.addSeparator()
        new_action = menu.addAction("New Query")

        selected_items = self.query_list.selectedItems()
        selected_count = len(selected_items)

        can_set_default = selected_count == 1 and self.get_item_name(selected_items[0]) != self.default_query
        set_default_action.setEnabled(can_set_default)
        rename_action.setEnabled(selected_count == 1 or (selected_count == 0 and self.query_list.currentItem() is not None))
        delete_action.setEnabled(selected_count > 0)

        chosen_action = menu.exec(self.query_list.mapToGlobal(position))
        if chosen_action == set_default_action:
            self.make_default()
        elif chosen_action == rename_action:
            self.rename_selected_query()
        elif chosen_action == delete_action:
            self.delete_selected_queries()
        elif chosen_action == new_action:
            self.add_new_query()

    def on_query_double_clicked(self, item: QListWidgetItem):
        """When a query is double-clicked, set it as default."""
        if item is None:
            return
        self.query_list.clearSelection()
        item.setSelected(True)
        self.query_list.setCurrentItem(item)
        self.make_default()

    def load_queries(self):
        """Load DAX queries from the PBIP file."""
        if not self.pbip_file or not os.path.isfile(self.pbip_file):
            return

        root_dir = os.path.splitext(self.pbip_file)[0] + ".SemanticModel/DAXQueries"
        json_path = os.path.join(root_dir, ".pbi", "daxQueries.json")

        try:
            with open(json_path, "r", encoding="utf-8") as f:
                data = json.load(f)

            queries = {}
            for query in data["tabOrder"]:
                with open(os.path.join(root_dir, f"{query}.dax"), "r", encoding="utf-8") as f:
                    dax_code = f.read()
                queries[query] = dax_code

            self.queries = queries
            self.default_query = data.get("defaultTab")

            self.ignore_item_changes = True
            self.query_list.clear()
            for name in self.queries.keys():
                self.query_list.addItem(self.create_query_item(name))
            self.ignore_item_changes = False

            if self.query_list.count() > 0:
                self.query_list.setCurrentRow(0)
            else:
                self.on_selection_changed()

        except Exception as e:
            QMessageBox.warning(self, "Error", f"Failed to load queries:\n{str(e)}")

    def on_selection_changed(self):
        """Handle query selection changes."""
        selected_items = self.query_list.selectedItems()

        if len(selected_items) == 1:
            item = selected_items[0]
            query_name = self.get_item_name(item)
            self.ignore_editor_changes = True
            self.query_editor.setPlainText(self.queries.get(query_name, ""))
            self.ignore_editor_changes = False
            self.query_editor.setEnabled(True)
            self.multi_selection_label.setVisible(False)
            #self.default_btn.setEnabled(query_name != self.default_query)
        elif len(selected_items) > 1:
            self.ignore_editor_changes = True
            self.query_editor.clear()
            self.ignore_editor_changes = False
            self.query_editor.setEnabled(False)
            self.multi_selection_label.setVisible(True)
            #self.default_btn.setEnabled(False)
        else:
            self.ignore_editor_changes = True
            self.query_editor.clear()
            self.ignore_editor_changes = False
            self.query_editor.setEnabled(False)
            self.multi_selection_label.setVisible(False)
            #self.default_btn.setEnabled(False)

    def on_text_changed(self):
        """Handle query text changes."""
        if self.ignore_editor_changes:
            return

        current = self.query_list.currentItem()
        if not current or len(self.query_list.selectedItems()) != 1:
            return

        query_name = self.get_item_name(current)
        new_text = self.query_editor.toPlainText()

        if new_text != self.queries.get(query_name):
            self.queries[query_name] = new_text
            self.save_button.setEnabled(True)

    def on_query_order_changed(self, *args, **kwargs):
        """Enable Save when the list order changes."""
        self.save_button.setEnabled(True)

    def create_query_item(self, name: str) -> QListWidgetItem:
        """Create a QListWidgetItem configured for the query list."""
        item = QListWidgetItem()
        item.setFlags(
            Qt.ItemFlag.ItemIsSelectable
            | Qt.ItemFlag.ItemIsEnabled
            | Qt.ItemFlag.ItemIsDragEnabled
            | Qt.ItemFlag.ItemIsEditable
        )
        self.update_item_display(item, name)
        return item

    def update_item_display(self, item: QListWidgetItem, name: str):
        """Update the visual representation of a list item."""
        previous_ignore_state = self.ignore_item_changes
        self.ignore_item_changes = True
        item.setData(Qt.ItemDataRole.UserRole, name)
        item.setData(Qt.ItemDataRole.EditRole, name)
        if name == self.default_query:
            item.setIcon(QIcon.fromTheme("star"))
            item.setText(f"{name} ⭐")
        else:
            item.setIcon(QIcon())
            item.setText(name)
        self.ignore_item_changes = previous_ignore_state

    def refresh_item_displays(self):
        """Refresh display for all list items."""
        previous_ignore_state = self.ignore_item_changes
        self.ignore_item_changes = True
        for i in range(self.query_list.count()):
            item = self.query_list.item(i)
            name = self.get_item_name(item)
            self.update_item_display(item, name)
        self.ignore_item_changes = previous_ignore_state

    def get_item_name(self, item: QListWidgetItem) -> str:
        """Return the underlying query name for the given list item."""
        stored = item.data(Qt.ItemDataRole.UserRole)
        if isinstance(stored, str) and stored:
            return stored
        return item.text().replace(" ⭐", "")

    def rename_selected_query(self):
        """Begin in-place renaming of the selected query."""
        selected_items = self.query_list.selectedItems()
        if len(selected_items) == 0 and self.query_list.currentItem() is not None:
            selected_items = [self.query_list.currentItem()]

        if len(selected_items) != 1:
            return

        item = selected_items[0]
        self.query_list.setCurrentItem(item)
        self.renaming_item = item
        self.renaming_original_name = self.get_item_name(item)
        self.ignore_item_changes = True
        item.setText(self.renaming_original_name)
        self.ignore_item_changes = False
        self.query_list.editItem(item)

    def on_query_item_changed(self, item: QListWidgetItem):
        """Handle the completion of an in-place rename."""
        if self.ignore_item_changes or item is None or item is not self.renaming_item:
            return

        new_name = item.text().strip()
        if not new_name:
            QMessageBox.warning(self, "Invalid Name", "Query name cannot be empty.")
            self.update_item_display(item, self.renaming_original_name)
            self.cancel_rename()
            return

        if any(char in INVALID_FILENAME_CHARS for char in new_name):
            QMessageBox.warning(
                self,
                "Invalid Name",
                f"Query name cannot contain any of the following characters:\n{''.join(sorted(INVALID_FILENAME_CHARS))}",
            )
            self.update_item_display(item, self.renaming_original_name)
            self.cancel_rename()
            return

        new_name_lower = new_name.lower()
        original_lower = self.renaming_original_name.lower()
        existing_names_lower = {name.lower() for name in self.queries if name.lower() != original_lower}
        if new_name_lower in existing_names_lower:
            QMessageBox.warning(self, "Duplicate Name", f"A query named '{new_name}' already exists.")
            self.update_item_display(item, self.renaming_original_name)
            self.cancel_rename()
            return

        if new_name == self.renaming_original_name:
            self.update_item_display(item, self.renaming_original_name)
            self.cancel_rename()
            return

        # Preserve insertion order while renaming the key
        reordered_queries = {}
        for name, code in self.queries.items():
            if name == self.renaming_original_name:
                reordered_queries[new_name] = code
            else:
                reordered_queries[name] = code
        self.queries = reordered_queries

        if self.default_query == self.renaming_original_name:
            self.default_query = new_name

        self.update_item_display(item, new_name)
        self.save_button.setEnabled(True)
        self.cancel_rename()

        # Keep the renamed item selected and update the editor
        item.setSelected(True)
        self.query_list.setCurrentItem(item)
        self.on_selection_changed()

    def on_item_editor_closed(self, *_):
        """Restore item display if rename was cancelled."""
        if self.renaming_item is not None:
            self.update_item_display(self.renaming_item, self.renaming_original_name)
            self.cancel_rename()

    def cancel_rename(self):
        """Clear rename tracking state."""
        self.renaming_item = None
        self.renaming_original_name = ""

    def add_new_query(self):
        """Add a new query with a unique default name."""
        new_name = self.generate_unique_name("New_Query_")
        self.queries[new_name] = ""

        item = self.create_query_item(new_name)
        self.query_list.addItem(item)
        self.query_list.clearSelection()
        item.setSelected(True)
        self.query_list.setCurrentItem(item)
        self.query_list.scrollToItem(item)

        self.save_button.setEnabled(True)
        self.on_selection_changed()

    def delete_selected_queries(self):
        """Delete selected queries after confirmation."""
        selected_items = self.query_list.selectedItems()
        if not selected_items:
            return

        if len(selected_items) == 1:
            name = self.get_item_name(selected_items[0])
            prompt = f"Delete query '{name}'?"
        else:
            prompt = f"Delete {len(selected_items)} queries?"

        confirm = QMessageBox.question(
            self,
            "Delete Queries",
            prompt,
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
            QMessageBox.StandardButton.No,
        )
        if confirm != QMessageBox.StandardButton.Yes:
            return

        rows = sorted((self.query_list.row(item), item) for item in selected_items)
        for row, item in reversed(rows):
            name = self.get_item_name(item)
            self.queries.pop(name, None)
            self.query_list.takeItem(row)

        self.ensure_default_query()
        self.save_button.setEnabled(True)

        if self.query_list.count() > 0:
            self.query_list.setCurrentRow(min(rows[0][0], self.query_list.count() - 1))
        else:
            self.on_selection_changed()

    def move_selected_queries(self, direction: int):
        """Move selected queries up or down."""
        if direction not in (-1, 1):
            return

        selected_items = self.query_list.selectedItems()
        if not selected_items:
            return

        rows = [self.query_list.row(item) for item in selected_items]
        if direction < 0 and min(rows) == 0:
            return
        if direction > 0 and max(rows) == self.query_list.count() - 1:
            return

        current_item = self.query_list.currentItem()
        ordered_items = sorted(selected_items, key=lambda item: self.query_list.row(item))
        if direction > 0:
            ordered_items.reverse()

        for item in ordered_items:
            row = self.query_list.row(item)
            self.query_list.takeItem(row)
            self.query_list.insertItem(row + direction, item)

        self.query_list.clearSelection()
        for item in selected_items:
            item.setSelected(True)

        if current_item in selected_items:
            self.query_list.setCurrentItem(current_item)
        elif selected_items:
            self.query_list.setCurrentItem(selected_items[0])

        self.save_button.setEnabled(True)

    def ensure_default_query(self):
        """Ensure there is a valid default query after deletions."""
        if self.query_list.count() == 0:
            self.default_query = None
            self.on_selection_changed()
            return

        existing_names = {self.get_item_name(self.query_list.item(i)) for i in range(self.query_list.count())}
        if self.default_query not in existing_names:
            self.default_query = self.get_item_name(self.query_list.item(0))
        self.refresh_item_displays()
        self.on_selection_changed()

    def generate_unique_name(self, base: str) -> str:
        """Generate a unique query name using the provided base string."""
        existing_lower = {name.lower() for name in self.queries}
        index = 1
        while True:
            candidate = f"{base}{index}"
            if candidate.lower() not in existing_lower:
                return candidate
            index += 1

    def make_default(self):
        """Make the selected query the default."""
        selected_items = self.query_list.selectedItems()
        if len(selected_items) != 1:
            return

        item = selected_items[0]
        query_name = self.get_item_name(item)
        if query_name == self.default_query:
            return

        self.default_query = query_name
        self.refresh_item_displays()
        self.default_btn.setEnabled(False)
        self.save_button.setEnabled(True)

    def save_changes(self):
        """Save changes to the queries."""
        if not self.pbip_file:
            return

        root_dir = os.path.splitext(self.pbip_file)[0] + ".SemanticModel/DAXQueries"
        json_path = os.path.join(root_dir, ".pbi", "daxQueries.json")

        try:
            new_query_order = []
            for i in range(self.query_list.count()):
                item_name = self.get_item_name(self.query_list.item(i))
                new_query_order.append(item_name)

            if not new_query_order:
                QMessageBox.warning(self, "No Queries", "There are no queries to save.")
                return

            if not self.default_query or self.default_query not in new_query_order:
                self.default_query = new_query_order[0]
                self.refresh_item_displays()

            with open(json_path, "r", encoding="utf-8") as f:
                data = json.load(f)

            data["tabOrder"] = new_query_order
            data["defaultTab"] = self.default_query

            with open(json_path, "w", encoding="utf-8") as f:
                json.dump(data, f, indent=4, ensure_ascii=False)

            os.makedirs(root_dir, exist_ok=True)

            # Remove existing .dax files before writing the updated set
            for filename in os.listdir(root_dir):
                if filename.lower().endswith(".dax"):
                    os.remove(os.path.join(root_dir, filename))

            for name in new_query_order:
                code = self.queries.get(name, "")
                with open(os.path.join(root_dir, f"{name}.dax"), "w", encoding="utf-8") as f:
                    f.write(code)

            self.save_button.setEnabled(False)
            QMessageBox.information(self, "Success", "Changes saved successfully!")

        except Exception as e:
            QMessageBox.critical(self, "Error", f"Failed to save changes:\n{str(e)}")
