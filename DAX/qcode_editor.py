from PyQt6.QtCore import Qt, QStringListModel
from PyQt6.QtGui import QTextCursor
from PyQt6.QtWidgets import QPlainTextEdit, QCompleter

try:
    from DAX.dax_editor_support import get_language_definition
except Exception:  # Fallback for environments without support module
    def get_language_definition(language):
        return None


class QCodeEditor(QPlainTextEdit):
    """Lightweight code editor with QCompleter-based autocompletion.

    - Ctrl+Space triggers the popup
    - Typing word characters updates and filters suggestions
    - Popup hides on Enter/Tab/Escape and non-word characters
    """

    def __init__(self, parent=None, language: str | None = None):
        super().__init__(parent)
        self._completer: QCompleter | None = None
        self._completer_model: QStringListModel | None = None
        self._function_names: set[str] = set()
        self._highlighter = None
        self._language: str | None = None
        self._line_comment: str | None = None
        self._ctrl_k_sequence: bool = False

        if language:
            self.set_language(language)

    def setCompleter(self, completer: QCompleter | None):
        if self._completer:
            try:
                self._completer.activated[str].disconnect(self.insertCompletion)
            except Exception:
                pass

        self._completer_model = None
        self._completer = completer
        if self._completer:
            # Anchor popup to the editor widget
            self._completer.setWidget(self)
            try:
                # PyQt6 may not support the old [str] signature form
                self._completer.activated.connect(self.insertCompletion)
            except Exception:
                # Fallback for environments exposing typed overloads
                try:
                    self._completer.activated[str].connect(self.insertCompletion)
                except Exception:
                    pass
            try:
                model = self._completer.model()
                if isinstance(model, QStringListModel):
                    self._completer_model = model
            except Exception:
                pass

    def completer(self) -> QCompleter | None:
        return self._completer

    def language(self) -> str | None:
        return self._language

    def set_language(self, language: str | None):
        """Configure syntax support for the requested language."""
        requested = language.lower() if language else None
        if requested == self._language:
            return

        # Tear down existing highlighter
        if self._highlighter:
            try:
                self._highlighter.setDocument(None)
            except Exception:
                pass
            self._highlighter = None

        definition = get_language_definition(language)
        if not definition:
            self._language = None
            self._function_names.clear()
            self.setCompleter(None)
            self._line_comment = None
            return

        self._language = definition.name.lower()
        self._function_names = {str(name).upper() for name in definition.functions}

        completions = definition.completions()
        if completions:
            model = QStringListModel(completions, self)
            completer = QCompleter(model, self)
            completer.setCaseSensitivity(Qt.CaseSensitivity.CaseInsensitive)
            completer.setFilterMode(Qt.MatchFlag.MatchContains)
            completer.setCompletionMode(QCompleter.CompletionMode.PopupCompletion)
            try:
                completer.setModelSorting(QCompleter.ModelSorting.CaseInsensitivelySortedModel)
            except Exception:
                pass
            self.setCompleter(completer)
            self._completer_model = model
        else:
            self.setCompleter(None)
            self._completer_model = None

        if definition.highlighter_cls:
            try:
                self._highlighter = definition.highlighter_cls(self.document())
                try:
                    self._highlighter.rehighlight()
                except Exception:
                    pass
            except Exception:
                self._highlighter = None
        self._line_comment = definition.line_comment

    def set_function_names(self, names):
        """Override function names used for () auto-insertion."""
        self._function_names = {str(name).upper() for name in (names or [])}

    def insertCompletion(self, *args):
        # Accept either (str) or (QModelIndex)
        if not args:
            return
        if isinstance(args[0], str):
            completion = args[0]
        else:
            idx = args[0]
            c = self._completer
            try:
                completion = c.completionModel().data(idx, c.completionRole())
            except Exception:
                return
        if completion is None:
            return
        completion_text = str(completion)

        tc = self.textCursor()
        tc.beginEditBlock()
        # replace current word with the completion
        tc.select(QTextCursor.SelectionType.WordUnderCursor)
        tc.removeSelectedText()
        tc.insertText(completion_text)
        tc.endEditBlock()
        self.setTextCursor(tc)

        # If it's a known function, add parentheses and place cursor inside
        try:
            if completion_text.upper() in self._function_names:
                # Avoid duplicating if the next char is already '('
                next_char = self._next_char_after_cursor()
                if next_char != '(':
                    tc = self.textCursor()
                    tc.beginEditBlock()
                    tc.insertText('()')
                    tc.movePosition(QTextCursor.MoveOperation.Left)
                    tc.endEditBlock()
                    self.setTextCursor(tc)
        except Exception:
            pass

    def _next_char_after_cursor(self) -> str:
        tc = self.textCursor()
        temp = QTextCursor(tc)
        temp.movePosition(QTextCursor.MoveOperation.Right, QTextCursor.MoveMode.KeepAnchor, 1)
        return temp.selectedText()

    def wordUnderCursor(self) -> str:
        tc = self.textCursor()
        block_text = tc.block().text()
        pos = tc.positionInBlock()
        l = pos
        while l > 0 and (block_text[l-1].isalnum() or block_text[l-1] == '_'):
            l -= 1
        r = pos
        while r < len(block_text) and (block_text[r].isalnum() or block_text[r] == '_'):
            r += 1
        return block_text[l:r]

    def focusInEvent(self, e):
        if self._completer:
            self._completer.setWidget(self)
        super().focusInEvent(e)

    def focusOutEvent(self, e):
        self._ctrl_k_sequence = False
        if self._completer:
            self._completer.popup().hide()
        super().focusOutEvent(e)

    def keyPressEvent(self, e):
        key = e.key()
        mods = e.modifiers()

        # Alt+Up / Alt+Down to move lines
        if mods == Qt.KeyboardModifier.AltModifier and key in (Qt.Key.Key_Up, Qt.Key.Key_Down):
            direction = -1 if key == Qt.Key.Key_Up else 1
            if self._move_selected_lines(direction):
                e.accept()
                return

        # Handle second chord after Ctrl+K
        if self._ctrl_k_sequence:
            handled = False
            if mods & Qt.KeyboardModifier.ControlModifier:
                if key == Qt.Key.Key_C:
                    handled = self._comment_selection()
                elif key == Qt.Key.Key_U:
                    handled = self._uncomment_selection()
            self._ctrl_k_sequence = False
            if handled:
                e.accept()
                return

        # Detect Ctrl+K chord start
        if (mods & Qt.KeyboardModifier.ControlModifier) and key == Qt.Key.Key_K:
            self._ctrl_k_sequence = True
            e.accept()
            return
        else:
            self._ctrl_k_sequence = False

        c = self._completer

        # If completer visible, let it handle navigation/acceptance keys
        if c and c.popup().isVisible():
            if e.key() in (
                Qt.Key.Key_Return,
                Qt.Key.Key_Enter,
                Qt.Key.Key_Escape,
                Qt.Key.Key_Tab,
                Qt.Key.Key_Backtab,
                Qt.Key.Key_Up,
                Qt.Key.Key_Down,
                Qt.Key.Key_PageUp,
                Qt.Key.Key_PageDown,
            ):
                e.ignore()
                return

        # Ctrl+Space: trigger before base handler to avoid side effects
        if c and (e.modifiers() & Qt.KeyboardModifier.ControlModifier) and e.key() == Qt.Key.Key_Space:
            prefix = self.wordUnderCursor()
            self._showCompleter(c, prefix, allow_empty=True)
            return

        # Let base class insert the character first
        super().keyPressEvent(e)

        if not c:
            return

        typed = e.text()

        # Update as user types word chars or deletes
        if (typed and (typed.isalnum() or typed == '_')) or e.key() in (Qt.Key.Key_Backspace, Qt.Key.Key_Delete):
            prefix = self.wordUnderCursor()
            if prefix:
                self._showCompleter(c, prefix)
            else:
                c.popup().hide()
        else:
            # Non-word character typed -> hide
            c.popup().hide()

    def _showCompleter(self, c: QCompleter, prefix: str, allow_empty: bool = False):
        c.setCompletionPrefix(prefix)
        if not prefix and not allow_empty:
            c.popup().hide()
            return
        cr = self.cursorRect()
        try:
            width = c.popup().sizeHintForColumn(0) + c.popup().verticalScrollBar().sizeHint().width() + 12
            cr.setWidth(width)
        except Exception:
            pass
        model = c.completionModel()
        try:
            if model.rowCount() > 0:
                c.popup().setCurrentIndex(model.index(0, 0))
        except AttributeError:
            pass
        try:
            c.complete(cr)
        except Exception:
            # Fallback: let completer place itself
            c.complete()

    def _move_selected_lines(self, direction: int) -> bool:
        cursor = self.textCursor()
        doc = self.document()
        start_block, end_block = self._selected_block_range(cursor)
        if not start_block.isValid() or not end_block.isValid():
            return False

        start_pos = start_block.position()

        if direction < 0:
            previous_block = start_block.previous()
            if not previous_block.isValid():
                return False

            selection_end = end_block.position() + end_block.length()
            selected_text = self._get_plain_text(start_pos, selection_end)
            selected_length = len(selected_text)

            combined_start = previous_block.position()
            combined_end = selection_end
            preceding_text = self._get_plain_text(combined_start, start_pos)
            new_text = selected_text + preceding_text
            new_selection_start = combined_start
        else:
            next_block = end_block.next()
            if not next_block.isValid():
                return False
            if not next_block.next().isValid() and not next_block.text() and next_block.length() <= 1:
                return False

            selection_end = next_block.position()
            selected_text = self._get_plain_text(start_pos, selection_end)
            selected_length = len(selected_text)

            combined_start = start_pos
            combined_end = next_block.position() + next_block.length()
            following_text = self._get_plain_text(selection_end, combined_end)
            new_text = following_text + selected_text
            new_selection_start = combined_start + len(following_text)

        edit_cursor = QTextCursor(doc)
        edit_cursor.setPosition(combined_start)
        edit_cursor.setPosition(combined_end, QTextCursor.MoveMode.KeepAnchor)
        edit_cursor.beginEditBlock()
        edit_cursor.insertText(new_text)
        edit_cursor.endEditBlock()

        new_cursor = self.textCursor()
        new_cursor.setPosition(new_selection_start)
        new_cursor.setPosition(new_selection_start + selected_length, QTextCursor.MoveMode.KeepAnchor)
        self.setTextCursor(new_cursor)
        return True

    def _comment_selection(self) -> bool:
        token = self._line_comment
        if not token:
            return False

        cursor = self.textCursor()
        original_anchor = cursor.anchor()
        original_position = cursor.position()
        had_selection = cursor.hasSelection()

        doc = self.document()
        start_block, end_block = self._selected_block_range(cursor)
        if not start_block.isValid() or not end_block.isValid():
            return False

        start_num = start_block.blockNumber()
        end_num = end_block.blockNumber()

        changes: list[tuple[int, int, int, str]] = []  # (position, delta, removed_len, text)
        for num in range(start_num, end_num + 1):
            block = doc.findBlockByNumber(num)
            if not block.isValid():
                continue
            text = block.text()
            stripped = text.lstrip()
            if not stripped:
                continue
            insert_pos = block.position()
            insert_text = token + ' '
            changes.append((insert_pos, len(insert_text), 0, insert_text))

        if not changes:
            return True

        edit_cursor = QTextCursor(doc)
        edit_cursor.beginEditBlock()
        for pos, _, _, text in sorted(changes, key=lambda item: item[0], reverse=True):
            edit_cursor.setPosition(pos)
            edit_cursor.insertText(text)
        edit_cursor.endEditBlock()

        deltas = sorted([(pos, delta, removed) for pos, delta, removed, _ in changes], key=lambda item: item[0])
        self._restore_selection(original_anchor, original_position, had_selection, deltas)
        return True

    def _uncomment_selection(self) -> bool:
        token = self._line_comment
        if not token:
            return False

        cursor = self.textCursor()
        original_anchor = cursor.anchor()
        original_position = cursor.position()
        had_selection = cursor.hasSelection()

        doc = self.document()
        start_block, end_block = self._selected_block_range(cursor)
        if not start_block.isValid() or not end_block.isValid():
            return False

        start_num = start_block.blockNumber()
        end_num = end_block.blockNumber()

        changes: list[tuple[int, int, int]] = []  # (position, delta, removed_len)
        edit_ranges: list[tuple[int, int]] = []

        for num in range(start_num, end_num + 1):
            block = doc.findBlockByNumber(num)
            if not block.isValid():
                continue
            text = block.text()
            if not text.strip():
                continue
            if text.startswith(token):
                remove_pos = block.position()
                remove_len = len(token)
                remainder = text[len(token):]
                if remainder.startswith(' '):
                    remove_len += 1
                changes.append((remove_pos, -remove_len, remove_len))
                edit_ranges.append((remove_pos, remove_pos + remove_len))
            else:
                leading_len = len(text) - len(text.lstrip())
                stripped = text[leading_len:]
                if stripped.startswith(token):
                    remove_pos = block.position() + leading_len
                    remove_len = len(token)
                    remainder = stripped[len(token):]
                    if remainder.startswith(' '):
                        remove_len += 1
                    changes.append((remove_pos, -remove_len, remove_len))
                    edit_ranges.append((remove_pos, remove_pos + remove_len))

        if not edit_ranges:
            return True

        edit_cursor = QTextCursor(doc)
        edit_cursor.beginEditBlock()
        for start, end in sorted(edit_ranges, key=lambda item: item[0], reverse=True):
            edit_cursor.setPosition(start)
            edit_cursor.setPosition(end, QTextCursor.MoveMode.KeepAnchor)
            edit_cursor.removeSelectedText()
        edit_cursor.endEditBlock()

        self._restore_selection(original_anchor, original_position, had_selection, sorted(changes, key=lambda item: item[0]))
        return True

    def _selected_block_range(self, cursor: QTextCursor):
        doc = self.document()
        if cursor.hasSelection():
            start = cursor.selectionStart()
            end = cursor.selectionEnd()
            if end > start:
                end_index = end - 1
            else:
                end_index = end
            end_index = max(0, min(end_index, doc.characterCount() - 1))
            start_block = doc.findBlock(start)
            end_block = doc.findBlock(end_index)
        else:
            pos = cursor.position()
            start_block = doc.findBlock(pos)
            end_block = start_block
        return start_block, end_block

    def _get_plain_text(self, start: int, end: int) -> str:
        selection_cursor = QTextCursor(self.document())
        selection_cursor.setPosition(start)
        selection_cursor.setPosition(end, QTextCursor.MoveMode.KeepAnchor)
        return selection_cursor.selection().toPlainText()

    def _restore_selection(self, anchor: int, position: int, had_selection: bool, changes: list[tuple[int, int, int]]):
        new_anchor = anchor
        new_position = position

        if changes:
            new_anchor = self._adjust_position(anchor, changes)
            new_position = self._adjust_position(position, changes)

        max_pos = max(0, self.document().characterCount() - 1)
        new_anchor = max(0, min(new_anchor, max_pos))
        new_position = max(0, min(new_position, max_pos))

        new_cursor = self.textCursor()
        new_cursor.setPosition(new_anchor)
        if had_selection:
            new_cursor.setPosition(new_position, QTextCursor.MoveMode.KeepAnchor)
        else:
            new_cursor.setPosition(new_position)
        self.setTextCursor(new_cursor)

    @staticmethod
    def _adjust_position(pos: int, changes: list[tuple[int, int, int]]) -> int:
        new_pos = pos
        for change_pos, delta, removed_len in changes:
            if delta >= 0:
                if pos >= change_pos:
                    new_pos += delta
            else:
                removal_end = change_pos - delta  # delta is negative
                if pos < change_pos:
                    continue
                if pos <= removal_end:
                    new_pos = change_pos
                else:
                    new_pos += delta
        return max(0, new_pos)
