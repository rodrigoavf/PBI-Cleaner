import json
import os
from dataclasses import dataclass
from typing import Dict, List, Optional, Set

from PyQt6.QtCore import Qt, QPoint
from PyQt6.QtGui import QBrush, QColor, QIcon, QImage, QKeySequence, QPixmap, QShortcut
from PyQt6.QtWidgets import (
    QAbstractItemView,
    QHeaderView,
    QHBoxLayout,
    QLabel,
    QLineEdit,
    QMenu,
    QMessageBox,
    QPushButton,
    QTreeWidget,
    QTreeWidgetItem,
    QVBoxLayout,
    QWidget,
)

from common_functions import APP_THEME, simple_hash
import random

@dataclass
class BookmarkMeta:
    display_name: str
    path: Optional[str]
    valid: bool
    used: bool = False
    error: Optional[str] = None


class BookmarkTreeWidget(QTreeWidget):
    """Tree widget with custom drag/drop constraints for bookmark management."""

    def __init__(self, parent_tab: "TabBookmarks"):
        super().__init__(parent_tab)
        self._tab = parent_tab

    def dropEvent(self, event):
        super().dropEvent(event)
        self._tab.on_tree_structure_changed()

    def dragMoveEvent(self, event):
        target = self.itemAt(event.position().toPoint())
        indicator = self.dropIndicatorPosition()
        dragged_types = {item.data(0, TabBookmarks.ITEM_TYPE_ROLE) for item in self.selectedItems()}

        if target is not None:
            target_type = target.data(0, TabBookmarks.ITEM_TYPE_ROLE)

            if indicator == QAbstractItemView.DropIndicatorPosition.OnItem:
                if target_type == TabBookmarks.ITEM_BOOKMARK:
                    event.ignore()
                    return
                if target_type == TabBookmarks.ITEM_FOLDER and TabBookmarks.ITEM_FOLDER in dragged_types:
                    # prevent nesting folders
                    event.ignore()
                    return

        super().dragMoveEvent(event)

    def edit(self, index, trigger, event):
        if index.column() != 0:
            return False
        return super().edit(index, trigger, event)


class TabBookmarks(QWidget):
    """Bookmarks tab that surfaces bookmark metadata in a tree view."""

    ITEM_TYPE_ROLE = Qt.ItemDataRole.UserRole + 1
    ITEM_ID_ROLE = Qt.ItemDataRole.UserRole + 2
    ITEM_BOOKMARK = "bookmark"
    ITEM_FOLDER = "folder"

    def __init__(self, pbip_file: Optional[str] = None):
        super().__init__()
        self.pbip_file = pbip_file
        self.bookmarks: Dict[str, BookmarkMeta] = {}
        self.folders: Dict[str, dict] = {}
        self.structure: List[dict] = []
        self.dirty = False
        self._loading = False
        self._suppress_dirty = False
        self._ignore_item_changed = False

        self.folder_icon = self._load_icon("Folder.svg")

        self.init_ui()
        if self.pbip_file:
            self.load_bookmarks()
        else:
            self.update_status("Select a PBIP file to view bookmarks.", warning=False)

    def init_ui(self):
        layout = QVBoxLayout(self)
        layout.setContentsMargins(8, 8, 8, 8)
        layout.setSpacing(8)

        toolbar = QVBoxLayout()
        toolbar.setContentsMargins(0, 0, 0, 0)
        toolbar.setSpacing(4)

        primary_row = QHBoxLayout()
        self.save_button = QPushButton("💾 Save")
        self.save_button.setEnabled(False)
        self.save_button.clicked.connect(self.on_save_clicked)
        primary_row.addWidget(self.save_button)

        self.delete_button = QPushButton("❌ Delete")
        self.delete_button.setEnabled(False)
        self.delete_button.clicked.connect(self.delete_selected_item)
        primary_row.addWidget(self.delete_button)

        self.reload_button = QPushButton("🔄 Reload")
        self.reload_button.clicked.connect(self.reload_from_disk)
        primary_row.addWidget(self.reload_button)

        self.sort_button = QPushButton("↑↓ Sort A-Z")
        self.sort_button.setToolTip("Sort the current folder (or root) alphabetically.")
        self.sort_button.clicked.connect(self.sort_current_scope)
        primary_row.addWidget(self.sort_button)

        primary_row.addStretch()

        self.dirty_label = QLabel("Unsaved changes")
        self.dirty_label.setStyleSheet("color: #d8902f; font-style: italic;")
        self.dirty_label.setVisible(False)
        primary_row.addWidget(self.dirty_label)

        secondary_row = QHBoxLayout()
        self.expand_button = QPushButton("⏬ Expand All")
        self.expand_button.setToolTip("Expand all folders and bookmarks.")
        self.expand_button.clicked.connect(self.expand_all_items)
        secondary_row.addWidget(self.expand_button)

        self.collapse_button = QPushButton("⏫ Collapse All")
        self.collapse_button.setToolTip("Collapse all folders.")
        self.collapse_button.clicked.connect(self.collapse_all_items)
        secondary_row.addWidget(self.collapse_button)

        self.select_all_button = QPushButton("Select All")
        self.select_all_button.setToolTip("Select every bookmark and folder.")
        self.select_all_button.clicked.connect(self.select_all_items)
        secondary_row.addWidget(self.select_all_button)

        self.unselect_all_button = QPushButton("Unselect All")
        self.unselect_all_button.setToolTip("Clear all selections.")
        self.unselect_all_button.clicked.connect(self.unselect_all_items)
        secondary_row.addWidget(self.unselect_all_button)

        self.select_not_used_button = QPushButton("Select Not Used")
        self.select_not_used_button.setToolTip("Select all bookmarks not referenced in any report page.")
        self.select_not_used_button.clicked.connect(self.select_not_used_items)
        secondary_row.addWidget(self.select_not_used_button)

        secondary_row.addStretch()

        self.filter_input = QLineEdit()
        self.filter_input.setPlaceholderText("Filter bookmarks...")
        self.filter_input.setClearButtonEnabled(False)
        self.filter_input.textChanged.connect(self.apply_filter)
        self.filter_input.setFixedWidth(220)
        secondary_row.addWidget(self.filter_input)

        self.clear_filter_button = QPushButton("🧹 Clear")
        self.clear_filter_button.setEnabled(False)
        self.clear_filter_button.clicked.connect(self.clear_filter)
        secondary_row.addWidget(self.clear_filter_button)

        toolbar.addLayout(primary_row)
        toolbar.addLayout(secondary_row)

        layout.addLayout(toolbar)

        self.tree = BookmarkTreeWidget(self)
        self.tree.setColumnCount(2)
        self.tree.setHeaderHidden(False)
        self.tree.setHeaderLabels(["Bookmark", "Usage"])
        header = self.tree.header()
        header.setStretchLastSection(False)
        header.setSectionResizeMode(0, QHeaderView.ResizeMode.Stretch)
        header.setSectionResizeMode(1, QHeaderView.ResizeMode.ResizeToContents)
        self.tree.setAlternatingRowColors(True)
        self.tree.setRootIsDecorated(True)
        self.tree.setSelectionMode(QAbstractItemView.SelectionMode.ExtendedSelection)
        self.tree.setEditTriggers(
            QAbstractItemView.EditTrigger.EditKeyPressed | QAbstractItemView.EditTrigger.SelectedClicked
        )
        self.tree.setDragEnabled(True)
        self.tree.setAcceptDrops(True)
        self.tree.setDropIndicatorShown(True)
        self.tree.setDefaultDropAction(Qt.DropAction.MoveAction)
        self.tree.setDragDropMode(QTreeWidget.DragDropMode.InternalMove)
        self.tree.itemSelectionChanged.connect(self.update_actions_state)
        self.tree.itemChanged.connect(self.on_item_changed)
        try:
            self.tree.model().rowsMoved.connect(self.on_tree_structure_changed)
        except Exception:
            pass
        self.tree.setContextMenuPolicy(Qt.ContextMenuPolicy.CustomContextMenu)
        self.tree.customContextMenuRequested.connect(self.show_context_menu)
        layout.addWidget(self.tree)

        self.status_label = QLabel()
        self.status_label.setStyleSheet("color: #d9534f;")
        self.status_label.setWordWrap(True)
        layout.addWidget(self.status_label)

        # Shortcuts
        QShortcut(QKeySequence("F2"), self, activated=self.rename_selected_item)
        QShortcut(QKeySequence("Ctrl+N"), self, activated=self.new_folder_shortcut)
        QShortcut(QKeySequence(Qt.Key.Key_Delete), self, activated=self.delete_selected_item)
        QShortcut(QKeySequence("Alt+Up"), self, activated=lambda: self.move_item(-1))
        QShortcut(QKeySequence("Alt+Down"), self, activated=lambda: self.move_item(1))

    # --- Icon helpers -----------------------------------------------------
    def _load_icon(self, name: str) -> QIcon:
        icon_dir = os.path.abspath(os.path.join(os.path.dirname(os.path.dirname(__file__)), "Images", "icons"))
        path = os.path.join(icon_dir, name)
        if not os.path.exists(path):
            return QIcon()

        if APP_THEME != "tentacles_light":
            pixmap = QPixmap(path)
            if pixmap.isNull():
                return QIcon(path)
            image = pixmap.toImage()
            image.invertPixels(QImage.InvertMode.InvertRgb)
            return QIcon(QPixmap.fromImage(image))
        return QIcon(path)

    # --- Data loading -----------------------------------------------------
    def bookmarks_base_dir(self) -> Optional[str]:
        if not self.pbip_file:
            return None
        base = os.path.splitext(self.pbip_file)[0] + ".Report"
        return os.path.join(base, "definition", "bookmarks")

    def pages_definition_dir(self) -> Optional[str]:
        if not self.pbip_file:
            return None
        base = os.path.splitext(self.pbip_file)[0] + ".Report"
        return os.path.join(base, "definition", "pages")

    def load_bookmarks(self):
        base_dir = self.bookmarks_base_dir()
        self._loading = True
        self.tree.clear()
        self.bookmarks.clear()
        self.folders.clear()
        self.structure.clear()
        warnings: List[str] = []
        self.tree.setEnabled(True)

        if not base_dir or not os.path.isdir(base_dir):
            self.tree.setEnabled(False)
            self.update_status("Bookmarks folder not found for this PBIP.", warning=True)
            self.mark_dirty(False, force=True)
            self._loading = False
            self.update_actions_state()
            self.apply_filter()
            return

        bookmarks_json_path = os.path.join(base_dir, "bookmarks.json")
        try:
            with open(bookmarks_json_path, "r", encoding="utf-8") as fh:
                data = json.load(fh)
        except FileNotFoundError:
            self.tree.setEnabled(False)
            self.update_status("bookmarks.json not found in the project.", warning=True)
            self.mark_dirty(False, force=True)
            self._loading = False
            self.update_actions_state()
            self.apply_filter()
            return
        except Exception as exc:
            self.tree.setEnabled(False)
            self.update_status(f"Failed to parse bookmarks.json: {exc}", warning=True)
            self.mark_dirty(False, force=True)
            self._loading = False
            self.update_actions_state()
            self.apply_filter()
            return

        raw_items = data.get("items")
        if not isinstance(raw_items, list):
            self.tree.setEnabled(False)
            self.update_status("bookmarks.json missing a valid 'items' list.", warning=True)
            self.mark_dirty(False, force=True)
            self._loading = False
            self.update_actions_state()
            self.apply_filter()
            return

        child_lookup: Set[str] = set()
        for entry in raw_items:
            if isinstance(entry, dict):
                children = entry.get("children")
                if isinstance(children, list):
                    child_lookup.update(str(child) for child in children if isinstance(child, str))

        # Read folder definitions
        for entry in raw_items:
            if not isinstance(entry, dict):
                continue
            name = entry.get("name")
            if not isinstance(name, str):
                continue
            if "children" in entry:
                children = entry.get("children")
                if isinstance(children, list):
                    valid_children = [child for child in children if isinstance(child, str)]
                else:
                    valid_children = []
                display_name = entry.get("displayName") or name
                self.folders[name] = {"display": display_name, "children": valid_children}

        # Read bookmark metadata from files
        for fname in os.listdir(base_dir):
            if not fname.lower().endswith(".bookmark.json"):
                continue
            stem = fname[:-len(".bookmark.json")]
            path = os.path.join(base_dir, fname)
            display = stem
            valid = True
            error_message = None
            try:
                with open(path, "r", encoding="utf-8") as fh:
                    bookmark_data = json.load(fh)
                display = bookmark_data.get("displayName") or display
            except json.JSONDecodeError as exc:
                display = f"{stem} (invalid)"
                valid = False
                error_message = f"Invalid JSON: {exc}"
                warnings.append(f"Bookmark '{stem}' has invalid JSON.")
            except Exception as exc:
                display = f"{stem} (unreadable)"
                valid = False
                error_message = str(exc)
                warnings.append(f"Bookmark '{stem}' could not be read.")

            self.bookmarks[stem] = BookmarkMeta(display_name=display, path=path, valid=valid, error=error_message)

        self._compute_bookmark_usage()

        # Build tree according to items order
        bookmarks_in_tree: Set[str] = set()
        bookmarks_in_folders = {child for folder in self.folders.values() for child in folder["children"]}

        for entry in raw_items:
            if not isinstance(entry, dict):
                continue
            name = entry.get("name")
            if not isinstance(name, str):
                continue

            if "children" in entry:
                folder_item = self.create_folder_item(name)
                self.tree.addTopLevelItem(folder_item)
                for child_name in self.folders.get(name, {}).get("children", []):
                    self.add_bookmark_item(child_name, folder_item)
                    bookmarks_in_tree.add(child_name)
                folder_item.setExpanded(True)
            else:
                if name in bookmarks_in_folders:
                    continue
                self.add_bookmark_item(name, None)
                bookmarks_in_tree.add(name)

        # Add any bookmarks not referenced in items (sorted for stability)
        for name in sorted(set(self.bookmarks.keys()) - bookmarks_in_tree, key=lambda n: self.bookmarks[n].display_name.casefold()):
            self.add_bookmark_item(name, None)

        if warnings:
            self.update_status("; ".join(warnings), warning=True)
        else:
            self.update_status("", warning=False)

        self.tree.expandToDepth(0)
        self.mark_dirty(False, force=True)
        self._loading = False
        self._suppress_dirty = True
        self.on_tree_structure_changed()
        self.apply_filter()
        self.update_actions_state()
        if not self._loading:
            self._adjust_usage_column_width()

    def _compute_bookmark_usage(self):
        for meta in self.bookmarks.values():
            meta.used = False

        pages_dir = self.pages_definition_dir()
        if not pages_dir or not os.path.isdir(pages_dir):
            return

        remaining: Set[str] = set(self.bookmarks.keys())
        if not remaining:
            return

        for root, _, files in os.walk(pages_dir):
            for fname in files:
                if not fname.lower().endswith(".json"):
                    continue
                path = os.path.join(root, fname)
                try:
                    with open(path, "r", encoding="utf-8", errors="ignore") as fh:
                        content = fh.read()
                except Exception:
                    continue

                matched = {name for name in remaining if name in content}
                if matched:
                    for name in matched:
                        self.bookmarks[name].used = True
                    remaining.difference_update(matched)
                    if not remaining:
                        return

    # --- UI helpers -------------------------------------------------------
    def update_status(self, text: str, *, warning: bool):
        if not text:
            self.status_label.clear()
            return
        color = "#d9534f" if warning else "#6c757d"
        self.status_label.setStyleSheet(f"color: {color};")
        self.status_label.setText(text)

    def mark_dirty(self, dirty: bool, *, force: bool = False):
        if self._loading and not force:
            return
        if not force and self.dirty == dirty:
            return
        self.dirty = dirty
        self.save_button.setEnabled(dirty)
        self.dirty_label.setVisible(dirty)

    def update_actions_state(self):
        has_selection = bool(self.tree.selectedItems())
        self.delete_button.setEnabled(has_selection)

    def create_folder_item(self, folder_id: str) -> QTreeWidgetItem:
        display = self.folders.get(folder_id, {}).get("display", folder_id)
        item = QTreeWidgetItem([display, ""])
        item.setData(0, self.ITEM_TYPE_ROLE, self.ITEM_FOLDER)
        item.setData(0, self.ITEM_ID_ROLE, folder_id)
        flags = (
            Qt.ItemFlag.ItemIsEnabled
            | Qt.ItemFlag.ItemIsSelectable
            | Qt.ItemFlag.ItemIsEditable
            | Qt.ItemFlag.ItemIsDragEnabled
            | Qt.ItemFlag.ItemIsDropEnabled
        )
        item.setFlags(flags)
        if self.folder_icon and not self.folder_icon.isNull():
            item.setIcon(0, self.folder_icon)
        return item

    def add_bookmark_item(self, bookmark_id: str, parent: Optional[QTreeWidgetItem]):
        meta = self.bookmarks.get(bookmark_id)
        if meta is None:
            display = f"{bookmark_id} (missing)"
            item = QTreeWidgetItem([display, ""])
            item.setFlags(Qt.ItemFlag.ItemIsSelectable)
            item.setToolTip(0, "Bookmark file not found.")
            meta = BookmarkMeta(
                display_name=display,
                path=None,
                valid=False,
                used=False,
                error="Missing bookmark file.",
            )
            self.bookmarks[bookmark_id] = meta
        else:
            display = meta.display_name
            item = QTreeWidgetItem([display, ""])
            flags = Qt.ItemFlag.ItemIsEnabled | Qt.ItemFlag.ItemIsSelectable | Qt.ItemFlag.ItemIsDragEnabled
            if meta.valid:
                flags |= Qt.ItemFlag.ItemIsEditable
                item.setToolTip(0, meta.path or "")
            else:
                item.setToolTip(0, meta.error or "Bookmark has issues.")
                flags &= ~Qt.ItemFlag.ItemIsEnabled
            item.setFlags(flags)

        item.setData(0, self.ITEM_TYPE_ROLE, self.ITEM_BOOKMARK)
        item.setData(0, self.ITEM_ID_ROLE, bookmark_id)

        self._apply_usage_style(item, meta)

        if parent is None:
            self.tree.addTopLevelItem(item)
        else:
            parent.addChild(item)

        if not self._loading:
            self._adjust_usage_column_width()

    def _apply_usage_style(self, item: QTreeWidgetItem, meta: Optional[BookmarkMeta]):
        if meta is None:
            item.setText(1, "")
            item.setForeground(1, QBrush())
            return

        if meta.used:
            item.setText(1, "✔ Used")
            item.setForeground(1, QBrush(QColor("#2e7d32")))
        else:
            item.setText(1, "✘ Not Used")
            item.setForeground(1, QBrush(QColor("#c62828")))

    def _adjust_usage_column_width(self):
        if self.tree.columnCount() <= 1:
            return
        column = 1
        header = self.tree.header()
        header.setSectionResizeMode(column, QHeaderView.ResizeMode.ResizeToContents)
        self.tree.doItemsLayout()
        width = max(self.tree.sizeHintForColumn(column), header.sectionSize(column))
        if width <= 0:
            width = 80
        header.setSectionResizeMode(column, QHeaderView.ResizeMode.Interactive)
        header.resizeSection(column, int(width * 2))


    # --- Actions ----------------------------------------------------------
    def select_all_items(self):
        self.tree.selectAll()

    def unselect_all_items(self):
        self.tree.clearSelection()
        self.update_actions_state()

    def select_not_used_items(self):
        def traverse(item: QTreeWidgetItem):
            if item.data(0, self.ITEM_TYPE_ROLE) == self.ITEM_BOOKMARK:
                bookmark_id = item.data(0, self.ITEM_ID_ROLE)
                meta = self.bookmarks.get(bookmark_id)
                if meta and not meta.used:
                    item.setSelected(True)
            for idx in range(item.childCount()):
                traverse(item.child(idx))

        self.tree.blockSignals(True)
        try:
            self.tree.clearSelection()
            for i in range(self.tree.topLevelItemCount()):
                traverse(self.tree.topLevelItem(i))
        finally:
            self.tree.blockSignals(False)
        self.update_actions_state()

    def on_save_clicked(self):
        if not self.pbip_file:
            QMessageBox.warning(self, "No Project", "Select a PBIP project before saving bookmarks.")
            return

        base_dir = self.bookmarks_base_dir()
        if not base_dir or not os.path.isdir(base_dir):
            QMessageBox.critical(self, "Save Failed", "Bookmarks folder for this PBIP project could not be located.")
            return

        try:
            snapshot, folder_displays, bookmark_display, seen_bookmarks, seen_folders = self._collect_tree_snapshot()
        except ValueError as exc:
            QMessageBox.warning(self, "Cannot Save", str(exc))
            return

        bookmarks_json_path = os.path.join(base_dir, "bookmarks.json")
        try:
            with open(bookmarks_json_path, "r", encoding="utf-8") as fh:
                bookmarks_json = json.load(fh)
            if not isinstance(bookmarks_json, dict):
                bookmarks_json = {}
        except FileNotFoundError:
            bookmarks_json = {}
        except Exception as exc:
            QMessageBox.critical(self, "Save Failed", f"Failed to read bookmarks.json:\n{exc}")
            return

        new_items: List[dict] = []
        for entry in snapshot:
            if entry["type"] == self.ITEM_FOLDER:
                new_items.append({
                    "name": entry["id"],
                    "displayName": entry["display"],
                    "children": entry["children"],
                })
            else:
                new_items.append({"name": entry["id"]})

        bookmarks_json["items"] = new_items

        warnings: List[str] = []

        try:
            with open(bookmarks_json_path, "w", encoding="utf-8") as fh:
                json.dump(bookmarks_json, fh, indent=4, ensure_ascii=False)
                fh.write("\n")
        except Exception as exc:
            QMessageBox.critical(self, "Save Failed", f"Failed to write bookmarks.json:\n{exc}")
            return

        for bookmark_id, display in bookmark_display.items():
            meta = self.bookmarks.get(bookmark_id)
            bookmark_path = meta.path if meta and meta.path else os.path.join(base_dir, f"{bookmark_id}.bookmark.json")

            try:
                with open(bookmark_path, "r", encoding="utf-8") as fh:
                    bookmark_json = json.load(fh)
                if not isinstance(bookmark_json, dict):
                    raise ValueError("Unexpected JSON structure.")
            except FileNotFoundError:
                warnings.append(f"Bookmark file not found: {bookmark_id}.bookmark.json")
                continue
            except Exception as exc:
                warnings.append(f"Failed to read {bookmark_id}.bookmark.json: {exc}")
                continue

            bookmark_json["displayName"] = display

            try:
                with open(bookmark_path, "w", encoding="utf-8") as fh:
                    json.dump(bookmark_json, fh, indent=4, ensure_ascii=False)
                    fh.write("\n")
            except Exception as exc:
                warnings.append(f"Failed to write {bookmark_id}.bookmark.json: {exc}")
                continue

            if meta:
                meta.display_name = display
                meta.path = bookmark_path
                meta.valid = True
                meta.error = None

        existing_ids = set(self.bookmarks.keys())
        removed_bookmarks = existing_ids - seen_bookmarks

        for bookmark_id in removed_bookmarks:
            meta = self.bookmarks.get(bookmark_id)
            bookmark_path = meta.path if meta and meta.path else os.path.join(base_dir, f"{bookmark_id}.bookmark.json")
            if bookmark_path and os.path.exists(bookmark_path):
                try:
                    os.remove(bookmark_path)
                except Exception as exc:
                    warnings.append(f"Failed to delete {bookmark_id}.bookmark.json: {exc}")
                    continue
            self.bookmarks.pop(bookmark_id, None)

        for folder_id in list(self.folders.keys()):
            if folder_id not in seen_folders:
                self.folders.pop(folder_id, None)

        for entry in snapshot:
            if entry["type"] == self.ITEM_FOLDER:
                self.folders[entry["id"]] = {
                    "display": entry["display"],
                    "children": entry["children"],
                }

        self.structure = [{"type": entry["type"], "id": entry["id"]} for entry in snapshot]

        self.apply_filter()
        self.update_actions_state()

        if warnings:
            warning_text = "\n".join(f"- {msg}" for msg in warnings)
            self.mark_dirty(True, force=True)
            QMessageBox.warning(self, "Saved with Warnings", f"Bookmarks saved, but some issues occurred:\n{warning_text}")
        else:
            self.mark_dirty(False, force=True)
            QMessageBox.information(self, "Bookmarks Saved", "Bookmarks have been saved successfully.")

    def reload_from_disk(self):
        if self.dirty:
            reply = QMessageBox.question(
                self,
                "Discard changes?",
                "Reloading will discard your unsaved changes. Continue?",
                QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
            )
            if reply != QMessageBox.StandardButton.Yes:
                return
        self.load_bookmarks()

    def sort_current_scope(self):
        selected_keys, current_key = self._capture_selection()
        self._loading = True
        try:
            folders: List[QTreeWidgetItem] = []
            bookmarks: List[QTreeWidgetItem] = []

            while self.tree.topLevelItemCount():
                item = self.tree.takeTopLevelItem(0)
                item_type = item.data(0, self.ITEM_TYPE_ROLE)
                if item_type == self.ITEM_FOLDER:
                    self._sort_child_items(item)
                    folders.append(item)
                else:
                    bookmarks.append(item)

            key_func = lambda node: node.text(0).casefold()
            folders.sort(key=key_func)
            bookmarks.sort(key=key_func)

            for item in folders + bookmarks:
                self.tree.addTopLevelItem(item)
        finally:
            self._loading = False

        self._restore_selection(selected_keys, current_key)
        self.on_tree_structure_changed()
        self.apply_filter()

    def expand_all_items(self):
        self.tree.expandAll()

    def collapse_all_items(self):
        self.tree.collapseAll()

    def clear_filter(self):
        self.filter_input.clear()
        self.filter_input.setFocus()

    def apply_filter(self):
        pattern_raw = self.filter_input.text() or ""
        pattern = pattern_raw.strip().casefold()
        self.clear_filter_button.setEnabled(bool(pattern))

        selected_keys, current_key = self._capture_selection()
        self.tree.setUpdatesEnabled(False)
        try:
            if not pattern:
                for item in self._iter_items():
                    item.setHidden(False)
            else:
                for i in range(self.tree.topLevelItemCount()):
                    item = self.tree.topLevelItem(i)
                    self._apply_filter_to_item(item, pattern)
        finally:
            self.tree.setUpdatesEnabled(True)

        if pattern:
            visible_keys = set()
            for item in self._iter_items():
                if not item.isHidden():
                    key = (item.data(0, self.ITEM_TYPE_ROLE), item.data(0, self.ITEM_ID_ROLE))
                    if key[1] is not None:
                        visible_keys.add(key)
            selected_keys = [key for key in selected_keys if key in visible_keys]
            if current_key and current_key not in visible_keys:
                current_key = selected_keys[0] if selected_keys else None

        self._restore_selection(selected_keys, current_key)
        self.update_actions_state()

    def _sort_child_items(self, parent_item: QTreeWidgetItem):
        if parent_item.childCount() <= 1:
            return
        children = [parent_item.child(i) for i in range(parent_item.childCount())]
        children.sort(key=lambda node: node.text(0).casefold())
        for index in range(parent_item.childCount() - 1, -1, -1):
            parent_item.takeChild(index)
        for child in children:
            parent_item.addChild(child)

    def _collect_tree_snapshot(self):
        root = self.tree.invisibleRootItem()
        snapshot: List[dict] = []
        folder_displays: Dict[str, str] = {}
        bookmark_displays: Dict[str, str] = {}
        seen_folders: Set[str] = set()
        seen_bookmarks: Set[str] = set()

        for idx in range(root.childCount()):
            item = root.child(idx)
            if item is None:
                continue
            item_type = item.data(0, self.ITEM_TYPE_ROLE)
            item_id = item.data(0, self.ITEM_ID_ROLE)
            if not item_id:
                continue

            if item_type == self.ITEM_FOLDER:
                if item_id in seen_folders:
                    raise ValueError(f"Duplicate folder identifier '{item_id}' detected. Please rename one of the folders before saving.")
                seen_folders.add(item_id)

                display_name = item.text(0)
                folder_displays[item_id] = display_name

                children_ids: List[str] = []
                seen_children: Set[str] = set()
                for child_idx in range(item.childCount()):
                    child = item.child(child_idx)
                    child_type = child.data(0, self.ITEM_TYPE_ROLE)
                    child_id = child.data(0, self.ITEM_ID_ROLE)
                    if child_type != self.ITEM_BOOKMARK or not child_id:
                        continue
                    if child_id in seen_children:
                        raise ValueError(
                            f"Folder '{display_name}' contains duplicate bookmark '{child_id}'. Adjust the order before saving."
                        )
                    if child_id in seen_bookmarks:
                        raise ValueError(
                            f"Duplicate bookmark identifier '{child_id}' detected. Bookmarks must be unique before saving."
                        )
                    seen_children.add(child_id)
                    seen_bookmarks.add(child_id)
                    children_ids.append(child_id)
                    bookmark_displays[child_id] = child.text(0)

                snapshot.append({
                    "type": self.ITEM_FOLDER,
                    "id": item_id,
                    "display": display_name,
                    "children": children_ids,
                })
            elif item_type == self.ITEM_BOOKMARK:
                if item_id in seen_bookmarks:
                    raise ValueError(
                        f"Duplicate bookmark identifier '{item_id}' detected. Bookmarks must be unique before saving."
                    )
                seen_bookmarks.add(item_id)
                bookmark_displays[item_id] = item.text(0)
                snapshot.append({
                    "type": self.ITEM_BOOKMARK,
                    "id": item_id,
                })

        return snapshot, folder_displays, bookmark_displays, seen_bookmarks, seen_folders

    def _apply_filter_to_item(self, item: QTreeWidgetItem, pattern: str) -> bool:
        matches_self = pattern in item.text(0).casefold()
        any_child_visible = False
        for idx in range(item.childCount()):
            child = item.child(idx)
            if self._apply_filter_to_item(child, pattern):
                any_child_visible = True

        is_folder = item.data(0, self.ITEM_TYPE_ROLE) == self.ITEM_FOLDER
        visible = matches_self or any_child_visible
        item.setHidden(not visible)
        if is_folder:
            item.setExpanded(bool(pattern) and (matches_self or any_child_visible))
        return visible

    def _capture_selection(self):
        selected_keys = []
        seen = set()
        for item in self.tree.selectedItems():
            key = (item.data(0, self.ITEM_TYPE_ROLE), item.data(0, self.ITEM_ID_ROLE))
            if key[1] is None or key in seen:
                continue
            seen.add(key)
            selected_keys.append(key)
        current_item = self.tree.currentItem()
        current_key = None
        if current_item is not None:
            current_key = (
                current_item.data(0, self.ITEM_TYPE_ROLE),
                current_item.data(0, self.ITEM_ID_ROLE),
            )
        return selected_keys, current_key

    def _restore_selection(self, selected_keys, current_key):
        selected_set = {key for key in selected_keys if key[1] is not None}
        target_current = None
        self.tree.blockSignals(True)
        try:
            self.tree.clearSelection()
            for item in self._iter_items():
                key = (item.data(0, self.ITEM_TYPE_ROLE), item.data(0, self.ITEM_ID_ROLE))
                if key in selected_set:
                    item.setSelected(True)
                if current_key and key == current_key:
                    target_current = item
        finally:
            self.tree.blockSignals(False)

        if target_current:
            self.tree.setCurrentItem(target_current)
        else:
            selection = self.tree.selectedItems()
            if selection:
                self.tree.setCurrentItem(selection[0])
        self.update_actions_state()

    def _iter_items(self, parent: Optional[QTreeWidgetItem] = None):
        if parent is None:
            iterator = (self.tree.topLevelItem(i) for i in range(self.tree.topLevelItemCount()))
        else:
            iterator = (parent.child(i) for i in range(parent.childCount()))
        for item in iterator:
            yield item
            yield from self._iter_items(item)

    def show_context_menu(self, point: QPoint):
        item = self.tree.itemAt(point)
        menu = QMenu(self)
        selected_items = self.tree.selectedItems()
        multi_selection = len(selected_items) > 1

        if multi_selection:
            menu.addAction("Delete", self.delete_selected_item)
        elif item is None:
            menu.addAction("New Folder", lambda: self.create_new_folder(None))
        else:
            item_type = item.data(0, self.ITEM_TYPE_ROLE)
            if item_type == self.ITEM_FOLDER:
                menu.addAction("Rename", self.rename_selected_item)
                menu.addAction("New Folder", lambda: self.create_new_folder(item))
                menu.addAction("Delete", self.delete_selected_item)
            elif item_type == self.ITEM_BOOKMARK:
                if item.flags() & Qt.ItemFlag.ItemIsEditable:
                    menu.addAction("Rename", self.rename_selected_item)
                menu.addAction("Delete", self.delete_selected_item)

        if menu.actions():
            menu.exec(self.tree.viewport().mapToGlobal(point))

    def rename_selected_item(self):
        item = self.tree.currentItem()
        if not item:
            return
        if not item.flags() & Qt.ItemFlag.ItemIsEditable:
            return
        self.tree.editItem(item, 0)

    def new_folder_shortcut(self):
        current = self.tree.currentItem()
        self.create_new_folder(current if current and current.data(0, self.ITEM_TYPE_ROLE) == self.ITEM_FOLDER else None)

    def create_new_folder(self, reference_item: Optional[QTreeWidgetItem]):
        folder_id = self.generate_folder_id()
        display_name = self.generate_folder_name()
        self.folders[folder_id] = {"display": display_name, "children": []}
        item = self.create_folder_item(folder_id)

        root = self.tree.invisibleRootItem()
        insert_index = root.childCount()
        if reference_item is not None:
            insert_index = root.indexOfChild(reference_item) + 1
        root.insertChild(insert_index, item)
        self.tree.setCurrentItem(item)
        self.mark_dirty(True)
        self.tree.editItem(item, 0)
        self.on_tree_structure_changed()

    def generate_folder_id(self) -> str:
        existing_ids = set(self.folders.keys()) | set(self.bookmarks.keys())
        while True:
            candidate = f"Tentacles_{simple_hash(f"Tentacles_{random.random()}")}"
            if candidate not in existing_ids:
                return candidate

    def generate_folder_name(self) -> str:
        base = "New Folder"
        existing = {self.folders[f]["display"] for f in self.folders}
        if base not in existing:
            return base
        index = 2
        while True:
            candidate = f"{base} {index}"
            if candidate not in existing:
                return candidate
            index += 1

    def delete_selected_item(self):
        selected_items = list(self.tree.selectedItems())
        if not selected_items:
            return

        selected_keys, current_key = self._capture_selection()
        folders = [item for item in selected_items if item.data(0, self.ITEM_TYPE_ROLE) == self.ITEM_FOLDER]
        bookmarks = [item for item in selected_items if item.data(0, self.ITEM_TYPE_ROLE) == self.ITEM_BOOKMARK]

        if not folders and not bookmarks:
            return

        deleted_keys = {(self.ITEM_FOLDER, item.data(0, self.ITEM_ID_ROLE)) for item in folders}
        deleted_keys |= {(self.ITEM_BOOKMARK, item.data(0, self.ITEM_ID_ROLE)) for item in bookmarks}

        if len(selected_items) == 1:
            item = selected_items[0]
            item_type = item.data(0, self.ITEM_TYPE_ROLE)
            if item_type == self.ITEM_FOLDER:
                folder_id = item.data(0, self.ITEM_ID_ROLE)
                reply = QMessageBox.question(
                    self,
                    "Delete folder?",
                    "Delete this folder? Bookmarks inside will move to the root level.",
                    QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
                )
                if reply != QMessageBox.StandardButton.Yes:
                    return

                children_to_restore = []
                while item.childCount():
                    child = item.takeChild(0)
                    child.setExpanded(False)
                    children_to_restore.append((child, child.isSelected()))
                for child, was_selected in children_to_restore:
                    self.tree.addTopLevelItem(child)
                    child.setSelected(was_selected)

                parent = item.parent()
                if parent:
                    parent.removeChild(item)
                else:
                    index = self.tree.indexOfTopLevelItem(item)
                    if index >= 0:
                        self.tree.takeTopLevelItem(index)
                self.folders.pop(folder_id, None)
            elif item_type == self.ITEM_BOOKMARK:
                reply = QMessageBox.question(
                    self,
                    "Remove bookmark?",
                    "Remove this bookmark from the current view? (Changes are not saved yet.)",
                    QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
                )
                if reply != QMessageBox.StandardButton.Yes:
                    return
                parent = item.parent()
                if parent:
                    parent.removeChild(item)
                else:
                    index = self.tree.indexOfTopLevelItem(item)
                    if index >= 0:
                        self.tree.takeTopLevelItem(index)
            else:
                return
        else:
            parts = []
            if folders:
                count = len(folders)
                parts.append(f"{count} folder{'s' if count != 1 else ''}")
            if bookmarks:
                count = len(bookmarks)
                parts.append(f"{count} bookmark{'s' if count != 1 else ''}")
            detail = " and ".join(parts) if len(parts) == 2 else parts[0]
            note = " Bookmarks inside folders will move to the root level." if folders else ""
            reply = QMessageBox.question(
                self,
                "Delete selection?",
                f"Delete the selected {detail}?{note}",
                QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
            )
            if reply != QMessageBox.StandardButton.Yes:
                return

            for item in folders:
                folder_id = item.data(0, self.ITEM_ID_ROLE)
                children_to_restore = []
                while item.childCount():
                    child = item.takeChild(0)
                    child.setExpanded(False)
                    children_to_restore.append((child, child.isSelected()))
                for child, was_selected in children_to_restore:
                    self.tree.addTopLevelItem(child)
                    child.setSelected(was_selected)

                parent = item.parent()
                if parent:
                    parent.removeChild(item)
                else:
                    index = self.tree.indexOfTopLevelItem(item)
                    if index >= 0:
                        self.tree.takeTopLevelItem(index)
                self.folders.pop(folder_id, None)

            for item in bookmarks:
                parent = item.parent()
                if parent:
                    parent.removeChild(item)
                else:
                    index = self.tree.indexOfTopLevelItem(item)
                    if index >= 0:
                        self.tree.takeTopLevelItem(index)

        selected_keys = [key for key in selected_keys if key not in deleted_keys]
        if current_key in deleted_keys:
            current_key = None

        self._restore_selection(selected_keys, current_key)
        self.mark_dirty(True)
        self.on_tree_structure_changed()
        self.apply_filter()

    def move_item(self, direction: int):
        item = self.tree.currentItem()
        if item is None:
            return
        parent = item.parent()
        container = parent if parent is not None else self.tree.invisibleRootItem()
        index = container.indexOfChild(item)
        target = index + direction
        if target < 0 or target >= container.childCount():
            return

        self._loading = True
        container.takeChild(index)
        container.insertChild(target, item)
        self._loading = False
        self.tree.setCurrentItem(item)
        self.on_tree_structure_changed()

    # --- Tree synchronization ---------------------------------------------
    def on_item_changed(self, item: QTreeWidgetItem, column: int):
        if column != 0:
            return
        if self._loading or self._ignore_item_changed:
            return
        text = item.text(column).strip()
        if not text:
            # restore previous text
            self._ignore_item_changed = True
            if item.data(0, self.ITEM_TYPE_ROLE) == self.ITEM_FOLDER:
                folder_id = item.data(0, self.ITEM_ID_ROLE)
                previous = self.folders.get(folder_id, {}).get("display", folder_id)
                item.setText(0, previous)
            else:
                bookmark_id = item.data(0, self.ITEM_ID_ROLE)
                previous = self.bookmarks.get(
                    bookmark_id,
                    BookmarkMeta(display_name=bookmark_id, path=None, valid=False),
                ).display_name
                item.setText(0, previous)
            self._ignore_item_changed = False
            return

        item_type = item.data(0, self.ITEM_TYPE_ROLE)
        if item_type == self.ITEM_FOLDER:
            folder_id = item.data(0, self.ITEM_ID_ROLE)
            if folder_id in self.folders:
                self.folders[folder_id]["display"] = text
            else:
                self.folders[folder_id] = {"display": text, "children": []}
        elif item_type == self.ITEM_BOOKMARK:
            bookmark_id = item.data(0, self.ITEM_ID_ROLE)
            meta = self.bookmarks.get(bookmark_id)
            if meta:
                meta.display_name = text
        self.mark_dirty(True)
        self.on_tree_structure_changed()
        self.apply_filter()

    def on_tree_structure_changed(self, *args, **kwargs):
        if self._loading:
            return

        self._loading = True
        should_mark_dirty = not self._suppress_dirty
        try:
            root = self.tree.invisibleRootItem()
            new_structure: List[dict] = []
            new_folders_children: Dict[str, List[str]] = {}

            for i in range(root.childCount()):
                item = root.child(i)
                item_type = item.data(0, self.ITEM_TYPE_ROLE)
                item_id = item.data(0, self.ITEM_ID_ROLE)

                if item_type == self.ITEM_FOLDER:
                    # Ensure folders remain top-level
                    if item.parent():
                        parent = item.parent()
                        parent.removeChild(item)
                        root.insertChild(i, item)
                    folder_display = item.text(0)
                    folder_data = self.folders.setdefault(item_id, {"display": folder_display, "children": []})
                    folder_data["display"] = folder_display

                    children_ids: List[str] = []
                    for c_idx in range(item.childCount()):
                        child = item.child(c_idx)
                        child_id = child.data(0, self.ITEM_ID_ROLE)
                        child_type = child.data(0, self.ITEM_TYPE_ROLE)
                        if child_type == self.ITEM_BOOKMARK and child_id:
                            children_ids.append(child_id)
                    new_folders_children[item_id] = children_ids
                    new_structure.append({"type": self.ITEM_FOLDER, "id": item_id})
                elif item_type == self.ITEM_BOOKMARK:
                    bookmark_id = item_id
                    meta = self.bookmarks.get(bookmark_id)
                    if meta:
                        meta.display_name = item.text(0)
                    new_structure.append({"type": self.ITEM_BOOKMARK, "id": bookmark_id})

            # Update folder membership
            for folder_id, children in new_folders_children.items():
                self.folders.setdefault(folder_id, {"display": folder_id, "children": []})
                self.folders[folder_id]["children"] = children

            self.structure = new_structure
        finally:
            self._loading = False
            self._suppress_dirty = False
        if should_mark_dirty:
            self.mark_dirty(True)

    # --- External API -----------------------------------------------------
    def set_pbip_file(self, pbip_file: Optional[str]):
        self.pbip_file = pbip_file
        if pbip_file:
            self.load_bookmarks()
        else:
            self.tree.clear()
            self.tree.setEnabled(False)
            self.update_status("Select a PBIP file to view bookmarks.", warning=False)
            self.mark_dirty(False, force=True)
            self.apply_filter()
