from PyQt6.QtCore import Qt
from PyQt6.QtGui import QTextCursor
from PyQt6.QtWidgets import QPlainTextEdit, QCompleter

# Optional: known function names to auto-insert parentheses for
# Prefer the local DAX module; fall back to Tabs if present
try:
    from DAX.dax_editor_support import DAX_FUNCTIONS as _DAX_FUNCTIONS
except Exception:
    try:
        from DAX.dax_editor_support import DAX_FUNCTIONS as _DAX_FUNCTIONS
    except Exception:
        _DAX_FUNCTIONS = []

_FUNC_SET = {str(name).upper() for name in _DAX_FUNCTIONS}


class QCodeEditor(QPlainTextEdit):
    """Lightweight code editor with QCompleter-based autocompletion.

    - Ctrl+Space triggers the popup
    - Typing word characters updates and filters suggestions
    - Popup hides on Enter/Tab/Escape and non-word characters
    """

    def __init__(self, parent=None):
        super().__init__(parent)
        self._completer: QCompleter | None = None

    def setCompleter(self, completer: QCompleter | None):
        if self._completer:
            try:
                self._completer.activated[str].disconnect(self.insertCompletion)
            except Exception:
                pass

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

    def completer(self) -> QCompleter | None:
        return self._completer

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
            if completion_text.upper() in _FUNC_SET:
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
        if self._completer:
            self._completer.popup().hide()
        super().focusOutEvent(e)

    def keyPressEvent(self, e):
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
