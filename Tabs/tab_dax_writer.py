import html
import json
import os
import re
import threading
import uuid
import urllib.parse
import urllib.request
from typing import Dict, List, Optional, Set, Tuple

from PyQt6.QtCore import Qt, QObject, QRunnable, QThreadPool, pyqtSignal
from PyQt6.QtGui import (
    QColor,
    QFont,
    QKeySequence,
    QShortcut,
    QSyntaxHighlighter,
    QTextCharFormat,
)
from PyQt6.QtWidgets import (
    QApplication,
    QAbstractItemView,
    QHBoxLayout,
    QLabel,
    QMessageBox,
    QPushButton,
    QSplitter,
    QTreeWidget,
    QTreeWidgetItem,
    QVBoxLayout,
    QWidget,
)

from Coding.code_editor import CodeEditor
from Coding.code_editor_support import DAXHighlighter, set_dax_model_identifiers
from common_functions import code_editor_font


class _TableColumnHighlightMixin:
    """Mixin providing shared formatting for table and column references."""

    def _init_table_column_support(self):
        self._table_patterns: List[re.Pattern] = []
        self._column_patterns: List[re.Pattern] = []

        self._table_format = QTextCharFormat()
        self._table_format.setForeground(QColor("#2c7be5"))
        self._table_format.setFontWeight(QFont.Weight.Medium)

        self._column_format = QTextCharFormat()
        self._column_format.setForeground(QColor("#00a676"))
        self._column_format.setFontWeight(QFont.Weight.Medium)

    def update_patterns(
        self,
        table_patterns: List[re.Pattern],
        column_patterns: List[re.Pattern],
    ):
        self._table_patterns = table_patterns
        self._column_patterns = column_patterns
        self.rehighlight()

    def _apply_table_column_highlighting(self, text: str) -> None:
        for pattern in self._table_patterns:
            for match in pattern.finditer(text):
                start, end = match.span()
                self.setFormat(start, end - start, self._table_format)

        for pattern in self._column_patterns:
            for match in pattern.finditer(text):
                start, end = match.span()
                self.setFormat(start, end - start, self._column_format)


class TableColumnHighlighter(_TableColumnHighlightMixin, QSyntaxHighlighter):
    """Custom highlighter that only colors table and column references."""

    def __init__(self, document):
        QSyntaxHighlighter.__init__(self, document)
        self._init_table_column_support()

    def highlightBlock(self, text: str) -> None:
        self._apply_table_column_highlighting(text)


class DAXTableColumnHighlighter(_TableColumnHighlightMixin, DAXHighlighter):
    """DAX syntax highlighter extended with table/column coloring."""

    def __init__(self, document):
        DAXHighlighter.__init__(self, document)
        self._init_table_column_support()

    def highlightBlock(self, text: str) -> None:
        super().highlightBlock(text)
        self._apply_table_column_highlighting(text)


class ChatRequestSignals(QObject):
    finished = pyqtSignal(str)
    error = pyqtSignal(str)


class ChatRequestWorker(QRunnable):
    def __init__(self, fn, *args, **kwargs):
        super().__init__()
        self.fn = fn
        self.args = args
        self.kwargs = kwargs
        self.signals = ChatRequestSignals()

    def run(self):
        try:
            result = self.fn(*self.args, **self.kwargs)
        except Exception as exc:  # pragma: no cover - defensive
            self.signals.error.emit(str(exc))
        else:
            self.signals.finished.emit(result)


class ChatGPTFreeClient:
    CHAT_URL = "https://chatgptfree.ai/chat/"

    def __init__(self):
        self._opener = urllib.request.build_opener(urllib.request.HTTPCookieProcessor())
        self._config: Optional[Dict] = None
        self._session_id = str(uuid.uuid4())
        self._conversation_id = str(uuid.uuid4())
        self._lock = threading.Lock()

    def _fetch_page_config(self) -> Dict:
        headers = {
            "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64)",
            "Referer": self.CHAT_URL,
        }
        request = urllib.request.Request(self.CHAT_URL, headers=headers)
        with self._opener.open(request, timeout=30) as resp:
            html_text = resp.read().decode("utf-8", "ignore")

        match = re.search(r"data-config='(.*?)'", html_text)
        if not match:
            raise RuntimeError("Unable to locate chatbot configuration on page.")

        config = json.loads(html.unescape(match.group(1)))
        required = ("ajaxUrl", "nonce", "botId")
        if not all(key in config for key in required):
            raise RuntimeError("Incomplete chatbot configuration received.")
        return config

    def _ensure_config(self):
        if self._config is None:
            self._config = self._fetch_page_config()

    def _refresh_nonce(self):
        if not self._config:
            return
        ajax_url = self._config.get("ajaxUrl")
        headers = {
            "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64)",
            "Referer": self.CHAT_URL,
        }
        data = {
            "action": self._config.get("nonceRefreshAction", "aipkit_get_frontend_chat_nonce"),
            "bot_id": str(self._config.get("botId", "")),
        }
        encoded = urllib.parse.urlencode(data).encode("utf-8")
        request = urllib.request.Request(ajax_url, data=encoded, headers=headers)
        with self._opener.open(request, timeout=30) as resp:
            payload = json.loads(resp.read().decode("utf-8", "ignore"))
        nonce = payload.get("data", {}).get("nonce")
        if not nonce:
            raise RuntimeError("Nonce refresh failed.")
        self._config["nonce"] = nonce

    def generate(self, prompt: str) -> str:
        with self._lock:
            self._ensure_config()
            if not self._config:
                raise RuntimeError("Missing chatbot configuration.")

            ajax_url = self._config["ajaxUrl"]
            headers = {
                "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64)",
                "Referer": self.CHAT_URL,
                "Content-Type": "application/x-www-form-urlencoded",
            }
            data = {
                "action": "aipkit_frontend_chat_message",
                "_ajax_nonce": self._config["nonce"],
                "bot_id": str(self._config["botId"]),
                "session_id": self._session_id,
                "conversation_uuid": self._conversation_id,
                "post_id": str(self._config.get("postId", "")),
                "message": prompt,
            }

            encoded = urllib.parse.urlencode(data).encode("utf-8")
            request = urllib.request.Request(ajax_url, data=encoded, headers=headers)

            try:
                response = self._opener.open(request, timeout=60)
                payload_text = response.read().decode("utf-8", "ignore")
            except urllib.error.HTTPError as http_err:
                if http_err.code in (403, 401):
                    self._refresh_nonce()
                    data["_ajax_nonce"] = self._config["nonce"]
                    encoded = urllib.parse.urlencode(data).encode("utf-8")
                    request = urllib.request.Request(ajax_url, data=encoded, headers=headers)
                    response = self._opener.open(request, timeout=60)
                    payload_text = response.read().decode("utf-8", "ignore")
                else:
                    raise

        payload = json.loads(payload_text)
        if not payload.get("success"):
            message = payload.get("data", {}).get("message") or "ChatGPT request failed."
            raise RuntimeError(message)

        reply = payload.get("data", {}).get("reply", "").strip()
        if not reply:
            raise RuntimeError("ChatGPT returned an empty response.")

        return self._clean_reply(reply)

    @staticmethod
    def _clean_reply(text: str) -> str:
        fence_match = re.search(r"```(?:dax)?\s*(.*?)```", text, re.IGNORECASE | re.DOTALL)
        if fence_match:
            return fence_match.group(1).strip()
        return text.strip()


class DAXWriterTab(QWidget):
    """Tab that helps create DAX measures via ChatGPT."""

    def __init__(self, pbip_file: Optional[str] = None):
        super().__init__()
        self.pbip_file = pbip_file
        self.tables_data: Dict[str, List[str]] = {}
        self.table_patterns: List[Tuple[re.Pattern, str]] = []
        self.column_patterns: List[Tuple[List[re.Pattern], re.Pattern, Tuple[str, str]]] = []
        self.api_client = ChatGPTFreeClient()
        self.thread_pool = QThreadPool.globalInstance()

        self.prompt_editor: Optional[CodeEditor] = None
        self.output_editor: Optional[CodeEditor] = None
        self.table_tree: Optional[QTreeWidget] = None
        self.generate_button: Optional[QPushButton] = None
        self.copy_button: Optional[QPushButton] = None
        self.count_label: Optional[QLabel] = None
        self.status_label: Optional[QLabel] = None
        self.prompt_highlighter: Optional[TableColumnHighlighter] = None
        self.output_highlighter: Optional[TableColumnHighlighter] = None
        self._column_highlight_patterns: List[re.Pattern] = []
        self._selected_table: Optional[str] = None

        self._building_metadata = False

        self._init_ui()
        if pbip_file:
            self.load_metadata()

    # ----- UI -----------------------------------------------------------------
    def _init_ui(self):
        main_layout = QVBoxLayout(self)
        main_layout.setContentsMargins(8, 8, 8, 8)
        main_layout.setSpacing(8)

        top_bar = QHBoxLayout()
        self.generate_button = QPushButton("Generate DAX")
        self.generate_button.clicked.connect(self.generate_measure)
        top_bar.addWidget(self.generate_button)

        top_bar.addStretch()
        main_layout.addLayout(top_bar)

        prompt_label = QLabel(
            "Describe the measure you need. Reference model objects as Table[Column] "
            "(for example: Sales[Amount])."
        )
        main_layout.addWidget(prompt_label)

        self.prompt_editor = CodeEditor()
        self.prompt_editor.setFont(code_editor_font())
        self._prepare_editor(self.prompt_editor, editable=True)
        self.prompt_editor.setPlaceholderText(
            "Example: Create a measure that sums Sales[Amount] for the selected Calendar[Year]."
        )
        self.prompt_editor.textChanged.connect(self._on_prompt_changed)
        main_layout.addWidget(self.prompt_editor)

        self.count_label = QLabel("Tables mentioned: 0  Columns mentioned: 0")
        self.count_label.setStyleSheet("color: #888888;")
        main_layout.addWidget(self.count_label)

        output_header = QHBoxLayout()
        output_header.addWidget(QLabel("Generated DAX Measure:"))
        output_header.addStretch()
        self.copy_button = QPushButton("Copy")
        self.copy_button.clicked.connect(self.copy_output)
        output_header.addWidget(self.copy_button)
        main_layout.addLayout(output_header)

        bottom_splitter = QSplitter(Qt.Orientation.Horizontal)
        bottom_splitter.setChildrenCollapsible(False)

        tables_container = QWidget()
        tables_layout = QVBoxLayout(tables_container)
        tables_layout.setContentsMargins(0, 0, 6, 0)
        tables_layout.setSpacing(4)

        tables_layout.addWidget(QLabel("Tables"))

        self.table_tree = QTreeWidget()
        self.table_tree.setHeaderHidden(True)
        self.table_tree.setRootIsDecorated(False)
        self.table_tree.setSelectionMode(QAbstractItemView.SelectionMode.SingleSelection)
        self.table_tree.setEditTriggers(QAbstractItemView.EditTrigger.NoEditTriggers)
        self.table_tree.itemSelectionChanged.connect(self._on_table_selection_changed)
        tables_layout.addWidget(self.table_tree)

        bottom_splitter.addWidget(tables_container)

        editor_container = QWidget()
        editor_layout = QVBoxLayout(editor_container)
        editor_layout.setContentsMargins(0, 0, 0, 0)
        editor_layout.setSpacing(0)

        self.output_editor = CodeEditor()
        self.output_editor.setFont(code_editor_font())
        self._prepare_editor(self.output_editor, editable=False)
        editor_layout.addWidget(self.output_editor)

        bottom_splitter.addWidget(editor_container)
        bottom_splitter.setStretchFactor(0, 0)
        bottom_splitter.setStretchFactor(1, 1)
        bottom_splitter.setSizes([220, 580])

        main_layout.addWidget(bottom_splitter)

        self.status_label = QLabel("")
        self.status_label.setStyleSheet("color: #888888;")
        main_layout.addWidget(self.status_label)

        shortcut = QShortcut(QKeySequence("Ctrl+Return"), self.prompt_editor)
        shortcut.activated.connect(self.generate_measure)

        shortcut_alt = QShortcut(QKeySequence("Ctrl+Enter"), self.prompt_editor)
        shortcut_alt.activated.connect(self.generate_measure)

    def _prepare_editor(self, editor: CodeEditor, *, editable: bool):
        if editable and self.prompt_highlighter:
            self.prompt_highlighter.setDocument(None)
            self.prompt_highlighter = None
        if not editable and self.output_highlighter:
            self.output_highlighter.setDocument(None)
            self.output_highlighter = None

        existing = getattr(editor, "_highlighter", None)
        if existing:
            try:
                existing.setDocument(None)
            except Exception:
                pass
            editor._highlighter = None

        if editable:
            highlighter_cls = TableColumnHighlighter
        else:
            try:
                editor.set_language("dax", force=True, enable_highlighter=False)
            except Exception:
                pass
            highlighter_cls = DAXTableColumnHighlighter

        highlighter = highlighter_cls(editor.document())
        if editable:
            self.prompt_highlighter = highlighter
        else:
            self.output_highlighter = highlighter

        self._update_highlighters()

    # ----- Metadata -----------------------------------------------------------
    def load_metadata(self):
        if not self.pbip_file:
            return
        if self._building_metadata:
            return

        self._building_metadata = True
        try:
            tables = self._extract_model_metadata()
            self.tables_data = tables
            self._build_patterns()
            self._update_highlighters()
            self._update_autocomplete()
            self._update_table_tree()
            self._on_prompt_changed()

            table_count = len(tables)
            column_count = sum(len(cols) for cols in tables.values())
            self.status_label.setText(
                f"Loaded {table_count} tables and {column_count} columns from the model."
            )
        except Exception as exc:
            self.tables_data = {}
            self.table_patterns = []
            self.column_patterns = []
            self._update_highlighters()
            self._update_table_tree()
            self.status_label.setText(f"Metadata load failed: {exc}")
        finally:
            self._building_metadata = False

    def _extract_model_metadata(self) -> Dict[str, List[str]]:
        if not self.pbip_file:
            raise RuntimeError("No PBIP file selected.")

        pbip_path = os.path.abspath(self.pbip_file)
        if not os.path.isfile(pbip_path):
            raise RuntimeError(f"PBIP file not found: {pbip_path}")

        semantic_root = os.path.splitext(pbip_path)[0] + ".SemanticModel"
        tables_dir = os.path.join(semantic_root, "definition", "tables")
        if not os.path.isdir(tables_dir):
            raise RuntimeError("Model tables directory not found.")

        table_columns: Dict[str, List[str]] = {}

        for filename in os.listdir(tables_dir):
            if not filename.lower().endswith(".tmdl"):
                continue
            table_name = os.path.splitext(filename)[0]
            table_path = os.path.join(tables_dir, filename)
            try:
                with open(table_path, "r", encoding="utf-8") as handle:
                    content = handle.read()
            except OSError:
                continue

            columns = self._parse_columns_from_tmdl(content)
            table_columns[table_name] = columns

        if not table_columns:
            raise RuntimeError("No table metadata found in model.")

        return table_columns

    @staticmethod
    def _parse_columns_from_tmdl(content: str) -> List[str]:
        pattern = re.compile(
            r'(?mi)^\s*column\s+(?:"([^"]+)"|([A-Za-z0-9_]+))\s*$'
        )
        columns: List[str] = []
        for match in pattern.finditer(content):
            column = match.group(1) or match.group(2) or ""
            column = column.strip()
            if column:
                columns.append(column)
        return columns

    def _update_autocomplete(self):
        tables_terms: List[str] = []
        columns_terms: List[str] = []

        for table, columns in self.tables_data.items():
            tables_terms.extend(self._table_autocomplete_forms(table))
            for column in columns:
                columns_terms.extend(self._column_autocomplete_forms(table, column))

        set_dax_model_identifiers(tables_terms, columns_terms)
        if self.prompt_editor:
            self.prompt_editor.set_language("dax", force=True, enable_highlighter=False)

    def _update_highlighters(self):
        def _deduplicate(patterns: List[re.Pattern]) -> List[re.Pattern]:
            unique: List[re.Pattern] = []
            seen = set()
            for pattern in patterns:
                key = (pattern.pattern, pattern.flags)
                if key in seen:
                    continue
                seen.add(key)
                unique.append(pattern)
            return unique

        table_regexes = _deduplicate([pattern for pattern, _ in self.table_patterns])
        column_regexes = _deduplicate(self._column_highlight_patterns)

        if self.prompt_highlighter:
            self.prompt_highlighter.update_patterns(table_regexes, column_regexes)
        if self.output_highlighter:
            self.output_highlighter.update_patterns(table_regexes, column_regexes)

    def _update_table_tree(self):
        if not self.table_tree:
            return

        previous = self._selected_table
        self.table_tree.blockSignals(True)
        self.table_tree.clear()

        tables = sorted(self.tables_data.keys(), key=str.casefold)
        selected_item: Optional[QTreeWidgetItem] = None

        for table in tables:
            item = QTreeWidgetItem([table])
            self.table_tree.addTopLevelItem(item)
            if table == previous:
                selected_item = item

        if selected_item:
            self.table_tree.setCurrentItem(selected_item)
            self._selected_table = selected_item.text(0)
        else:
            self._selected_table = None

        if self.table_tree.topLevelItemCount():
            self.table_tree.resizeColumnToContents(0)

        self.table_tree.blockSignals(False)

    @staticmethod
    def _table_autocomplete_forms(table: str) -> List[str]:
        forms = [table]
        if re.search(r"[^\w]", table):
            escaped = table.replace("'", "''")
            forms.append(f"'{escaped}'")
        return forms

    @staticmethod
    def _column_autocomplete_forms(table: str, column: str) -> List[str]:
        escaped_col = column.replace("]", "]]")
        forms = [f"[{escaped_col}]"]
        for table_form in DAXWriterTab._table_autocomplete_forms(table):
            forms.append(f"{table_form}[{escaped_col}]")
        return forms

    def _build_patterns(self):
        self.table_patterns = []
        self.column_patterns = []
        self._column_highlight_patterns = []

        for table in self.tables_data:
            for form in self._table_autocomplete_forms(table):
                if not form:
                    continue
                pattern = re.compile(
                    rf"(?<![\w\]]){re.escape(form)}(?![\w\[])",
                    re.IGNORECASE,
                )
                self.table_patterns.append((pattern, table))

        for table, columns in self.tables_data.items():
            for column in columns:
                escaped_col = column.replace("]", "]]")
                bracket_pattern = re.compile(
                    rf"\[\s*{re.escape(escaped_col)}\s*\]", re.IGNORECASE
                )
                bound_patterns: List[re.Pattern] = []
                for table_form in self._table_autocomplete_forms(table):
                    if not table_form:
                        continue
                    bound_patterns.append(
                        re.compile(
                            rf"{re.escape(table_form)}\s*\[\s*{re.escape(escaped_col)}\s*\]",
                            re.IGNORECASE,
                        )
                    )
                self.column_patterns.append((bound_patterns, bracket_pattern, (table, column)))
                self._column_highlight_patterns.extend(bound_patterns)
                self._column_highlight_patterns.append(bracket_pattern)

    # ----- Prompt helpers -----------------------------------------------------
    def _on_prompt_changed(self):
        if not self.prompt_editor:
            return
        text = self.prompt_editor.toPlainText()
        tables_found, columns_found = self._count_mentions(text)
        self.count_label.setText(
            f"Tables mentioned: {len(tables_found)}  Columns mentioned: {len(columns_found)}"
        )

    def _on_table_selection_changed(self):
        if not self.table_tree:
            return
        selected = self.table_tree.selectedItems()
        if selected:
            table_name = selected[0].text(0)
            self._selected_table = table_name
        else:
            self._selected_table = None

    def _count_mentions(self, text: str) -> Tuple[Set[str], Set[Tuple[str, str]]]:
        mentioned_tables: Set[str] = set()
        mentioned_columns: Set[Tuple[str, str]] = set()

        for pattern, table in self.table_patterns:
            if pattern.search(text):
                mentioned_tables.add(table)

        for bound_patterns, _, (table, column) in self.column_patterns:
            if bound_patterns and any(p.search(text) for p in bound_patterns):
                mentioned_columns.add((table, column))
                mentioned_tables.add(table)

        return mentioned_tables, mentioned_columns

    # ----- Actions ------------------------------------------------------------
    def generate_measure(self):
        if not self.prompt_editor or not self.output_editor or not self.generate_button:
            return
        if not self.pbip_file:
            QMessageBox.warning(self, "No PBIP", "Select a PBIP file to continue.")
            return

        text = self.prompt_editor.toPlainText().strip()
        if not text:
            QMessageBox.information(self, "Empty Prompt", "Please describe the measure you need.")
            return

        tables_found, columns_found = self._count_mentions(text)
        if not tables_found:
            QMessageBox.warning(
                self,
                "Missing Table",
                "Reference at least one table from the model (e.g., Sales or Sales[Amount]).",
            )
            return
        if not columns_found:
            QMessageBox.warning(
                self,
                "Missing Column",
                "Reference at least one column using Table[Column] syntax in your description.",
            )
            return

        final_prompt = (
            "Follow this instructions to the letter.\n"
            "Give your answer with a DAX code for Power BI and absolutely nothing else.\n"
            "The DAX code should be properly formatted and indented to be readable and easy to copy and paste.\n"
            "The DAX code should contain simple comments. The very first line of the code must not be a comment.\n"
            "Build a DAX measure that:\n"
            f"{text}"
        )

        self._set_busy(True, "Generating DAX measure...")
        worker = ChatRequestWorker(self.api_client.generate, final_prompt)
        worker.signals.finished.connect(self._on_generation_success)
        worker.signals.error.connect(self._on_generation_error)
        self.thread_pool.start(worker)

    def _on_generation_success(self, response: str):
        self._set_busy(False, "Generation complete.")
        if self.output_editor:
            self.output_editor.setPlainText(response)

    def _on_generation_error(self, message: str):
        self._set_busy(False, f"Generation failed: {message}")
        QMessageBox.critical(
            self,
            "Generation Failed",
            f"ChatGPT request failed:\n{message}",
        )

    def _set_busy(self, busy: bool, status: str):
        if self.generate_button:
            self.generate_button.setEnabled(not busy)
        if self.copy_button:
            self.copy_button.setEnabled(not busy)
        if self.prompt_editor:
            self.prompt_editor.setEnabled(not busy)
        if self.status_label:
            self.status_label.setText(status)

    def copy_output(self):
        if not self.output_editor:
            return
        text = self.output_editor.toPlainText().strip()
        if not text:
            return
        clipboard = QApplication.clipboard()
        clipboard.setText(text, mode=clipboard.Mode.Clipboard)
        self.status_label.setText("Copied generated DAX to clipboard.")
