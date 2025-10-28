import html
import json
import os
import re
import threading
import uuid
import urllib.parse
import urllib.request
from dataclasses import dataclass
from typing import Dict, List, Optional, Set, Tuple

from PyQt6.QtCore import Qt, QObject, QRunnable, QThreadPool, pyqtSignal
from PyQt6.QtGui import QKeySequence, QShortcut
from PyQt6.QtWidgets import (
    QApplication,
    QHBoxLayout,
    QLabel,
    QMessageBox,
    QPushButton,
    QVBoxLayout,
    QWidget,
)

from Coding.code_editor import CodeEditor
from Coding.code_editor_support import set_dax_model_identifiers
from common_functions import code_editor_font


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
        self.column_patterns: List[Tuple[List[re.Pattern], Tuple[str, str]]] = []
        self.api_client = ChatGPTFreeClient()
        self.thread_pool = QThreadPool.globalInstance()

        self.prompt_editor: Optional[CodeEditor] = None
        self.output_editor: Optional[CodeEditor] = None
        self.generate_button: Optional[QPushButton] = None
        self.copy_button: Optional[QPushButton] = None
        self.count_label: Optional[QLabel] = None
        self.status_label: Optional[QLabel] = None

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

        prompt_label = QLabel("Describe the measure you need (mention at least one table and column):")
        main_layout.addWidget(prompt_label)

        self.prompt_editor = CodeEditor(language="dax")
        self.prompt_editor.setFont(code_editor_font())
        self.prompt_editor.textChanged.connect(self._on_prompt_changed)
        main_layout.addWidget(self.prompt_editor)

        self.count_label = QLabel("Tables mentioned: 0 • Columns mentioned: 0")
        self.count_label.setStyleSheet("color: #888888;")
        main_layout.addWidget(self.count_label)

        output_header = QHBoxLayout()
        output_header.addWidget(QLabel("Generated DAX Measure:"))
        output_header.addStretch()
        self.copy_button = QPushButton("Copy")
        self.copy_button.clicked.connect(self.copy_output)
        output_header.addWidget(self.copy_button)
        main_layout.addLayout(output_header)

        self.output_editor = CodeEditor(language="dax", edit=False)
        self.output_editor.setFont(code_editor_font())
        self.output_editor.setReadOnly(True)
        self.output_editor.setUndoRedoEnabled(False)
        main_layout.addWidget(self.output_editor)

        self.status_label = QLabel("")
        self.status_label.setStyleSheet("color: #888888;")
        main_layout.addWidget(self.status_label)

        shortcut = QShortcut(QKeySequence("Ctrl+Return"), self.prompt_editor)
        shortcut.activated.connect(self.generate_measure)

        shortcut_alt = QShortcut(QKeySequence("Ctrl+Enter"), self.prompt_editor)
        shortcut_alt.activated.connect(self.generate_measure)

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
            self._update_autocomplete()
            self._build_patterns()
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
        if self.prompt_editor and self.prompt_editor.language() != "dax":
            self.prompt_editor.set_language("dax")

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

        for table in self.tables_data:
            for form in self._table_autocomplete_forms(table):
                pattern = re.compile(
                    rf"(?<![\w\]]){re.escape(form)}(?![\w\[])",
                    re.IGNORECASE,
                )
                self.table_patterns.append((pattern, table))

        for table, columns in self.tables_data.items():
            for column in columns:
                escaped_col = column.replace("]", "]]")
                forms = [f"[{escaped_col}]"]
                for table_form in self._table_autocomplete_forms(table):
                    forms.append(f"{table_form}[{escaped_col}]")
                patterns = [re.compile(re.escape(form), re.IGNORECASE) for form in forms]
                self.column_patterns.append((patterns, (table, column)))

    # ----- Prompt helpers -----------------------------------------------------
    def _on_prompt_changed(self):
        if not self.prompt_editor:
            return
        text = self.prompt_editor.toPlainText()
        tables_found, columns_found = self._count_mentions(text)
        self.count_label.setText(
            f"Tables mentioned: {len(tables_found)} • Columns mentioned: {len(columns_found)}"
        )

    def _count_mentions(self, text: str) -> Tuple[Set[str], Set[Tuple[str, str]]]:
        text_lower = text.lower()
        mentioned_tables: Set[str] = set()
        mentioned_columns: Set[Tuple[str, str]] = set()

        for pattern, table in self.table_patterns:
            if pattern.search(text_lower):
                mentioned_tables.add(table)

        for patterns, key in self.column_patterns:
            if any(p.search(text_lower) for p in patterns):
                mentioned_columns.add(key)

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
                "Mention at least one table from the model in your description.",
            )
            return
        if not columns_found:
            QMessageBox.warning(
                self,
                "Missing Column",
                "Mention at least one column from the model in your description.",
            )
            return

        final_prompt = (
            "I am a robot and I expect you to answer with a DAX code and absolutely nothing else, "
            "I repeat ABSOLUTELY NOTHING BUT THE DAX CODE for Power BI.\n"
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
