import os
from PyQt6.QtWidgets import (
    QMainWindow, QWidget, QTabWidget, QVBoxLayout, QLabel, QPushButton,
    QHBoxLayout, QFileDialog, QLineEdit, QMessageBox
)
from PyQt6.QtCore import Qt
from PyQt6.QtGui import QPixmap
from Tabs.tab_search import FileSearchApp
from Tabs.tab_dax_query import DAXQueryTab

class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Tentacles")
        self.setMinimumSize(1000, 500)

        self.pbip_path = None

        self.init_ui()

    def init_ui(self):
        # --- Starting screen for PBIP selection ---
        start_widget = QWidget()
        start_layout = QVBoxLayout()
        start_layout.setContentsMargins(60, 40, 60, 40)
        start_layout.setSpacing(20)
        
        # Add logo image
        logo_label = QLabel()
        logo = QPixmap(os.path.join(os.path.dirname(__file__), "Images", "Full_Logo.png"))
        target_height = 100
        logo = logo.scaledToHeight(target_height, Qt.TransformationMode.SmoothTransformation)
        logo_label.setPixmap(logo)
        logo_label.setAlignment(Qt.AlignmentFlag.AlignCenter)

        # --- Description area ---
        description = QLabel(
            "Tentacles helps you explore and organize your Power BI Project files "
            "in an intuitive way. Select your <b>.pbip</b> project to get started — "
            "then use the tools to search, organize, and clean your data model."
        )
        description.setWordWrap(True)
        description.setAlignment(Qt.AlignmentFlag.AlignCenter)
        description.setStyleSheet("font-size: 11pt; margin-bottom: 25px;")

        # --- Feature list ---
        features_title = QLabel("Main Features")
        features_title.setAlignment(Qt.AlignmentFlag.AlignLeft)
        features_title.setStyleSheet("font-size: 14pt; font-weight: bold; margin-top: 15px;")

        features_text = QLabel(
            """
            <ul style="font-size:12pt; line-height: 170%; color: white;">
            <li><b>Search Files</b> – Find and inspect files containing specific text.</li>
            <li><b>Bookmarks</b> – Reorganize, delete, or rename your bookmarks.</li>
            <li><b>Measures</b> – Identify unused measures, delete, group, or rename them.</li>
            <li><b>Tables</b> – Detect unused tables, delete or rename them easily.</li>
            <li><b>Columns</b> – Find unused columns, clean up and rename them.</li>
            </ul>
            """
        )
        features_text.setAlignment(Qt.AlignmentFlag.AlignLeft)
        features_text.setStyleSheet("margin-left: 60px; margin-right: 60px; color: white;")


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
        start_layout.addWidget(features_title)
        start_layout.addWidget(features_text)
        start_layout.addWidget(file_prompt)
        start_layout.addLayout(file_layout)
        start_layout.addWidget(confirm_btn, alignment=Qt.AlignmentFlag.AlignCenter)
        start_layout.addStretch()
        start_layout.addWidget(credit_label)

        start_widget.setLayout(start_layout)
        self.setCentralWidget(start_widget)

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
        self.file_input.setText(file_path)

    def load_main_tabs(self):
        # If pbip_path wasn't set via Browse, try to use whatever is in the text field.
        if not self.pbip_path:
            candidate = (self.file_input.text() or "").strip()
            if candidate and candidate.lower().endswith('.pbip') and os.path.isfile(candidate):
                self.pbip_path = candidate
            else:
                QMessageBox.warning(self, "Missing File", "Please select a .pbip file first.")
                return

        tabs = QTabWidget()
        dax_queries_tab = DAXQueryTab(self.pbip_path)
        bookmarks_tab = QWidget()
        measures_tab = QWidget()
        tables_tab = QWidget()
        columns_tab = QWidget()
        search_tab = FileSearchApp(self.pbip_path)

        # Set up placeholder content for unimplemented tabs
        for w, label in [
            (bookmarks_tab, "Bookmarks"),
            (measures_tab, "Measures"),
            (tables_tab, "Tables"),
            (columns_tab, "Columns"),
        ]:
            layout = QVBoxLayout()
            layout.addWidget(QLabel(f"Coming soon: {label} section"))
            w.setLayout(layout)

        tabs.addTab(dax_queries_tab, "DAX Queries")
        tabs.addTab(bookmarks_tab, "Bookmarks")
        tabs.addTab(measures_tab, "Measures")
        tabs.addTab(tables_tab, "Tables")
        tabs.addTab(columns_tab, "Columns")
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

    def change_file(self):
        self.pbip_path = None
        self.init_ui()
