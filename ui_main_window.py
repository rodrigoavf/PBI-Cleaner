import os
import sys
import subprocess
import webbrowser
from PyQt6.QtWidgets import (
    QApplication, QMainWindow, QWidget, QTabWidget, QVBoxLayout, QLabel, QPushButton,
    QHBoxLayout, QFileDialog, QLineEdit, QMessageBox, QStyleFactory
)
from PyQt6.QtCore import Qt
from PyQt6.QtGui import QPixmap, QAction, QActionGroup
from Tabs.tab_search import FileSearchApp
from Tabs.tab_dax_query import DAXQueryTab
from Tabs.tab_power_query import PowerQueryTab
from common_functions import apply_theme, THEME_PRESETS

class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.current_theme = apply_theme(QApplication.instance())
        self.setWindowTitle("Tentacles")
        self.setMinimumSize(800, 300)

        self.pbip_path = None
        self.theme_actions: dict[str, QAction] = {}
        self.theme_action_group: QActionGroup | None = None
        self.reload_project_action: QAction | None = None
        self.open_project_folder_action: QAction | None = None

        self.init_ui()
        self.setup_menu()
        self.refresh_menu_state()

    def init_ui(self):
        # --- Starting screen for PBIP selection ---
        start_widget = QWidget()
        start_layout = QVBoxLayout()
        start_layout.setContentsMargins(60, 40, 60, 40)
        start_layout.setSpacing(20)
        
        # Add logo image
        logo_label = QLabel()
        logo = QPixmap(os.path.join(os.path.dirname(__file__), "Images", "Full_Logo.png"))
        target_height = 150
        logo = logo.scaledToHeight(target_height, Qt.TransformationMode.SmoothTransformation)
        logo_label.setPixmap(logo)
        logo_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        logo_label.setMinimumHeight(target_height + 20)

        # --- Description area ---
        description = QLabel(
            "Tentacles helps you explore and organize your Power BI Project files "
            "in an intuitive way. Select your <b>.pbip</b> project to get started — "
            "then use the tools to search, organize, and clean your data model."
        )
        description.setWordWrap(True)
        description.setAlignment(Qt.AlignmentFlag.AlignCenter)
        description.setStyleSheet("font-size: 11pt; margin-bottom: 25px;")

        # --- File selector ---
        file_prompt = QLabel("Select your Power BI Project (.pbip) file to begin:")
        file_prompt.setAlignment(Qt.AlignmentFlag.AlignCenter)
        file_prompt.setStyleSheet("font-size: 12pt; margin-top: 20px;")

        self.file_input = QLineEdit()
        self.file_input.setPlaceholderText("Select your .pbip file...")
        self.file_input.setText("C:/Users/rodrigo.ferreira/Desktop/Devoteam/Supply & Purchasing.pbip")
        self.file_input.setReadOnly(False)

        browse_btn = QPushButton("Browse")
        browse_btn.clicked.connect(self.select_pbip_file)

        file_layout = QHBoxLayout()
        file_layout.addWidget(self.file_input)
        file_layout.addWidget(browse_btn)

        confirm_btn = QPushButton("Continue")
        confirm_btn.clicked.connect(self.load_main_tabs)
        confirm_btn.setFixedWidth(180)
        confirm_btn.setStyleSheet("font-size: 12pt; padding: 6px 12px;")

        # --- Footer credit ---
        credit_label = QLabel('<a href="https://www.linkedin.com/in/rodrigoavf/">Created by Rodrigo Ferreira</a>')
        credit_label.setTextFormat(Qt.TextFormat.RichText)
        credit_label.setTextInteractionFlags(Qt.TextInteractionFlag.TextBrowserInteraction)
        credit_label.setOpenExternalLinks(True)
        credit_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        credit_label.setStyleSheet("color: gray; font-size: 10pt; margin-top: 40px;")

        # --- Add widgets to layout ---
        start_layout.addWidget(logo_label)
        start_layout.addWidget(description)
        start_layout.addWidget(file_prompt)
        start_layout.addLayout(file_layout)
        start_layout.addWidget(confirm_btn, alignment=Qt.AlignmentFlag.AlignCenter)
        start_layout.addStretch()
        start_layout.addWidget(credit_label)

        start_widget.setLayout(start_layout)
        self.setCentralWidget(start_widget)
        self.setup_shortcuts()

    def setup_menu(self):
        """Create the application menu bar."""
        menu_bar = self.menuBar()
        menu_bar.clear()

        # --- File menu ---
        file_menu = menu_bar.addMenu("&File")

        new_project_action = QAction("New Project", self)
        # new_project_action.setShortcut("Ctrl+N")
        new_project_action.triggered.connect(self.change_file)
        file_menu.addAction(new_project_action)

        open_action = QAction("Open PBIP...", self)
        open_action.setShortcut("Ctrl+O")
        open_action.triggered.connect(self.open_pbip_via_menu)
        file_menu.addAction(open_action)

        self.reload_project_action = QAction("Reload Project", self)
        self.reload_project_action.setShortcut("Ctrl+R")
        self.reload_project_action.triggered.connect(self.reload_current_project)
        file_menu.addAction(self.reload_project_action)

        self.open_project_folder_action = QAction("Open Project Folder", self)
        self.open_project_folder_action.setShortcut("Ctrl+Shift+O")
        self.open_project_folder_action.triggered.connect(self.open_project_folder)
        file_menu.addAction(self.open_project_folder_action)

        file_menu.addSeparator()

        exit_action = QAction("Exit", self)
        exit_action.setShortcut("Ctrl+Q")
        exit_action.triggered.connect(self.close)
        file_menu.addAction(exit_action)

        # --- Settings menu ---
        settings_menu = menu_bar.addMenu("&Settings")

        theme_menu = settings_menu.addMenu("Theme")
        self.theme_actions = {}
        self.theme_action_group = QActionGroup(self)
        self.theme_action_group.setExclusive(True)

        available_styles = {name.lower() for name in QStyleFactory.keys()}

        for key, label in THEME_PRESETS.items():
            if key not in {"default"}:
                if key.startswith("fusion"):
                    if "fusion" not in available_styles:
                        continue
                elif key == "windowsvista":
                    if "windowsvista" not in available_styles:
                        continue
                elif key == "windows":
                    if "windows" not in available_styles and "windowsvista" not in available_styles:
                        continue
                elif key == "macintosh":
                    if "macintosh" not in available_styles:
                        continue

            action = QAction(label, self)
            action.setCheckable(True)
            action.triggered.connect(lambda checked, k=key: self.change_theme(k) if checked else None)
            self.theme_action_group.addAction(action)
            theme_menu.addAction(action)
            self.theme_actions[key] = action

        settings_menu.addSeparator()

        reset_size_action = QAction("Reset Window Size", self)
        reset_size_action.triggered.connect(self.reset_window_size)
        settings_menu.addAction(reset_size_action)

        # --- About menu ---
        about_menu = menu_bar.addMenu("&About")
        about_action = QAction("About Tentacles", self)
        about_action.triggered.connect(self.show_about_dialog)
        about_menu.addAction(about_action)

        author_action = QAction("Visit Author Profile", self)
        author_action.triggered.connect(self.visit_author_profile)
        about_menu.addAction(author_action)

        self.update_theme_checks()

    def setup_shortcuts(self):
        """Register keyboard shortcuts for main window actions."""
        self.file_input.returnPressed.connect(self.load_main_tabs)

    def refresh_menu_state(self):
        """Enable or disable menu actions based on current state."""
        if self.reload_project_action is None or self.open_project_folder_action is None:
            return

        has_project = bool(self.pbip_path)
        self.reload_project_action.setEnabled(has_project)
        self.open_project_folder_action.setEnabled(has_project)
        self.update_theme_checks()

    def update_theme_checks(self):
        """Reflect the active theme in the Settings menu."""
        if not self.theme_actions:
            return
        target = (self.current_theme or "").lower()
        matched = False
        for key, action in self.theme_actions.items():
            should_check = key.lower() == target
            action.blockSignals(True)
            action.setChecked(should_check)
            action.blockSignals(False)
            if should_check:
                matched = True
        if not matched:
            default_action = self.theme_actions.get("fusion_light") or self.theme_actions.get("default")
            if default_action:
                default_action.blockSignals(True)
                default_action.setChecked(True)
                default_action.blockSignals(False)

    def open_pbip_via_menu(self):
        """Open a PBIP file using the File > Open menu."""
        previous_path = self.pbip_path
        self.select_pbip_file()
        if self.pbip_path and self.pbip_path != previous_path:
            self.load_main_tabs()

    def reload_current_project(self):
        """Reload the active PBIP project."""
        if not self.pbip_path:
            QMessageBox.information(self, "No Project", "Select a .pbip file first.")
            return
        self.load_main_tabs()

    def open_project_folder(self):
        """Open the directory that contains the current project file."""
        if not self.pbip_path:
            QMessageBox.information(self, "No Project", "Select a .pbip file first.")
            return

        folder = os.path.dirname(self.pbip_path)
        if not folder or not os.path.isdir(folder):
            QMessageBox.warning(self, "Folder Missing", "Could not locate the project folder on disk.")
            return

        try:
            if sys.platform.startswith("win"):
                os.startfile(folder)
            elif sys.platform == "darwin":
                subprocess.check_call(["open", folder])
            else:
                subprocess.check_call(["xdg-open", folder])
        except Exception as exc:
            QMessageBox.warning(self, "Open Folder Failed", f"Unable to open the project folder:\n{exc}")

    def reset_window_size(self):
        """Restore the window to its default size."""
        self.resize(1000, 500)

    def change_theme(self, theme_key: str):
        """Apply a theme selection from the Settings menu."""
        resolved = apply_theme(QApplication.instance(), theme_key)
        self.current_theme = resolved
        self.update_theme_checks()

    def show_about_dialog(self):
        """Display application and author information."""
        about_text = (
            "<b>Tentacles</b><br>"
            "A companion tool to explore, clean, and organise Power BI project assets.<br><br>"
            "Developed by Rodrigo Ferreira to streamline working with PBIP files, "
            "offering quick search, query editing, and model insights.<br><br>"
            "Powered by PyQt6 and open-source contributions."
        )
        QMessageBox.about(self, "About Tentacles", about_text)

    def visit_author_profile(self):
        """Open the author's public profile in the default browser."""
        webbrowser.open("https://www.linkedin.com/in/rodrigoavf/")

    def select_pbip_file(self):
        file_path, _ = QFileDialog.getOpenFileName(
            self,
            "Select Power BI Project File (.pbip)",
            "",
            "Power BI Project (*.pbip);;All Files (*.*)",
        )

        if not file_path:
            return

        if not file_path.lower().endswith(".pbip"):
            QMessageBox.critical(self, "Invalid File", "Please select a valid .pbip file.")
            return

        self.pbip_path = file_path
        self.setWindowTitle(f"Tentacles{" - " + os.path.basename(self.pbip_path) if self.pbip_path else ""}")
        self.file_input.setText(file_path)
        self.refresh_menu_state()

    def load_main_tabs(self):
        # If pbip_path wasn't set via Browse, try to use whatever is in the text field.
        if not self.pbip_path:
            candidate = (self.file_input.text() or "").strip()
            if candidate and candidate.lower().endswith('.pbip') and os.path.isfile(candidate):
                self.pbip_path = candidate
                self.setWindowTitle(f"Tentacles{" - " + os.path.basename(self.pbip_path) if self.pbip_path else ""}")
            else:
                QMessageBox.warning(self, "Missing File", "Please select a .pbip file first.")
                return

        tabs = QTabWidget()
        dax_queries_tab = DAXQueryTab(self.pbip_path)
        bookmarks_tab = QWidget()
        measures_tab = QWidget()
        tables_tab = QWidget()
        columns_tab = QWidget()
        power_query_tab = PowerQueryTab(self.pbip_path)
        search_tab = FileSearchApp(self.pbip_path)

        # Set up placeholder content for unimplemented tabs
        for w, label in [
            (bookmarks_tab, "Bookmarks"),
            (measures_tab, "Measures"),
            (tables_tab, "Tables"),
            (columns_tab, "Columns"),
            (power_query_tab, "Power Query"),
        ]:
            layout = QVBoxLayout()
            layout.addWidget(QLabel(f"Coming soon: {label} section"))
            w.setLayout(layout)

        tabs.addTab(dax_queries_tab, "DAX Queries")
        tabs.addTab(bookmarks_tab, "Bookmarks")
        tabs.addTab(measures_tab, "Measures")
        tabs.addTab(tables_tab, "Tables")
        tabs.addTab(columns_tab, "Columns")
        tabs.addTab(power_query_tab, "Power Query")
        tabs.addTab(search_tab, "Search Files")

        info_widget = QWidget()
        info_layout = QHBoxLayout()

        file_label = QLabel(f"Current file: {self.pbip_path}")
        change_button = QPushButton("Change File")
        change_button.clicked.connect(self.change_file)

        info_layout.addWidget(file_label)
        info_layout.addStretch()
        info_layout.addWidget(change_button)
        info_widget.setLayout(info_layout)

        main_widget = QWidget()
        main_layout = QVBoxLayout()
        main_layout.addWidget(info_widget)
        main_layout.addWidget(tabs)

        # --- Footer credit label (below all tabs) ---
        credit_label = QLabel('<a href="https://www.linkedin.com/in/rodrigoavf/">Created by Rodrigo Ferreira</a>')
        credit_label.setTextFormat(Qt.TextFormat.RichText)
        credit_label.setTextInteractionFlags(Qt.TextInteractionFlag.TextBrowserInteraction)
        credit_label.setOpenExternalLinks(True)
        credit_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        credit_label.setStyleSheet("color: gray; font-size: 10pt; margin-top: 12px; margin-bottom: 8px;")

        main_layout.addWidget(credit_label)

        main_widget.setLayout(main_layout)
        self.setCentralWidget(main_widget)
        self.refresh_menu_state()

    def change_file(self):
        self.pbip_path = None
        self.init_ui()
        self.refresh_menu_state()
