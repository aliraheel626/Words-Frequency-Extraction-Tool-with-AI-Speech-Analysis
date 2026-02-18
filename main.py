"""
Words Frequency Extraction Tool
================================
A professional PySide6 desktop application that extracts word frequencies
from DOCX files and generates Excel reports with charts.
Includes AI-powered speech analysis using Cerebras.

Author: Speech Analyser Project
"""

import sys
import os
import re

from PySide6.QtWidgets import (
    QApplication,
    QMainWindow,
    QWidget,
    QVBoxLayout,
    QHBoxLayout,
    QLabel,
    QPushButton,
    QFileDialog,
    QProgressBar,
    QFrame,
    QGraphicsDropShadowEffect,
    QSizePolicy,
    QSpacerItem,
    QLineEdit,
)
from PySide6.QtCore import (
    Qt,
    QThread,
    Signal,
    QPropertyAnimation,
    QEasingCurve,
)
from PySide6.QtGui import (
    QFont,
    QColor,
    QPalette,
    QLinearGradient,
    QPainter,
    QIcon,
    QPixmap,
    QBrush,
    QPen,
)
from docx import Document
from openpyxl import Workbook
from openpyxl.chart import BarChart, Reference
from openpyxl.chart.series import SeriesLabel
from openpyxl.styles import Font as XlFont, PatternFill, Alignment, Border, Side
from spacy.lang.en.stop_words import STOP_WORDS
from langchain_cerebras import ChatCerebras
from langchain_core.messages import HumanMessage


# ─────────────────────────────────────────────────────────────────────────────
# Worker Thread for Background Processing
# ─────────────────────────────────────────────────────────────────────────────
class ExtractionWorker(QThread):
    """Background worker that extracts words and creates the Excel report."""

    progress = Signal(int)
    status_message = Signal(str)
    finished_signal = Signal(str)  # output path
    error_signal = Signal(str)

    def __init__(self, input_path, output_path, api_key=""):
        super().__init__()
        self.input_path = input_path
        self.output_path = output_path
        self.api_key = api_key
        print(
            f"DEBUG: Worker initialized with API Key: '{self.api_key[:5]}...' (Length: {len(self.api_key)})"
        )

    def run(self):
        try:
            # ── Step 1: Read the DOCX file ───────────────────────────────
            self.status_message.emit("Reading document...")
            self.progress.emit(10)

            doc = Document(self.input_path)
            full_text = []
            for para in doc.paragraphs:
                full_text.append(para.text)
            text = "\n".join(full_text)
            self.progress.emit(25)

            # ── Step 2: Extract words ────────────────────────────────────
            self.status_message.emit("Extracting words...")

            # Split into words - keep alphanumeric and apostrophes in words
            raw_words = re.findall(r"[A-Za-z0-9']+(?:'[A-Za-z]+)*", text)
            # Convert all words to lowercase for frequency counting
            words_lower = [w.lower() for w in raw_words]
            self.progress.emit(40)

            # ── Step 3: Calculate frequencies ────────────────────────────
            self.status_message.emit("Calculating frequencies...")
            word_freq = {}
            for w in words_lower:
                word_freq[w] = word_freq.get(w, 0) + 1
            # Sort by frequency descending, then alphabetically
            sorted_freq = sorted(word_freq.items(), key=lambda x: (-x[1], x[0]))
            self.progress.emit(55)

            # ── Step 4: Create Excel workbook ────────────────────────────
            self.status_message.emit("Generating Excel report...")
            wb = Workbook()

            # ── Sheet 1: All words (one per row) ─────────────────────────
            ws1 = wb.active
            ws1.title = "Sheet"

            # Row 1 is empty (matching expected output), words start from row 2
            for idx, word in enumerate(raw_words):
                ws1.cell(row=idx + 2, column=1, value=word)
                # Emit progress updates periodically
                if idx % 500 == 0:
                    pct = 55 + int((idx / max(len(raw_words), 1)) * 15)
                    self.progress.emit(min(pct, 70))

            # Adjust column width for Sheet 1
            ws1.column_dimensions["A"].width = 25
            self.progress.emit(72)

            # ── Sheet 2: Word Frequencies with chart ─────────────────────
            ws2 = wb.create_sheet(title="Word Frequencies")

            # Header row styling
            header_font = XlFont(name="Calibri", size=11, bold=True, color="FFFFFF")
            header_fill = PatternFill(
                start_color="2F5496", end_color="2F5496", fill_type="solid"
            )
            header_alignment = Alignment(horizontal="center", vertical="center")
            thin_border = Border(
                left=Side(style="thin"),
                right=Side(style="thin"),
                top=Side(style="thin"),
                bottom=Side(style="thin"),
            )

            # Write headers
            ws2.cell(row=1, column=1, value="WORD")
            ws2.cell(row=1, column=2, value="FREQUENCY")

            for col in [1, 2]:
                cell = ws2.cell(row=1, column=col)
                cell.font = header_font
                cell.fill = header_fill
                cell.alignment = header_alignment
                cell.border = thin_border

            self.progress.emit(75)

            # Write frequency data
            data_font = XlFont(name="Calibri", size=11)
            word_fill = PatternFill(
                start_color="FFFF00", end_color="FFFF00", fill_type="solid"
            )

            # Use spaCy's comprehensive stop words list
            stop_words = STOP_WORDS

            # Separate content words from stop words
            content_words = [
                (w, f) for w, f in sorted_freq if w.lower() not in stop_words
            ]
            stop_word_entries = [
                (w, f) for w, f in sorted_freq if w.lower() in stop_words
            ]

            # Write content words first (yellow highlighted), then stop words
            row_num = 2
            for idx, (word, freq) in enumerate(content_words):
                word_cell = ws2.cell(row=row_num, column=1, value=word)
                word_cell.font = data_font
                word_cell.fill = word_fill
                word_cell.border = thin_border

                freq_cell = ws2.cell(row=row_num, column=2, value=freq)
                freq_cell.font = data_font
                freq_cell.alignment = Alignment(horizontal="center")
                freq_cell.border = thin_border

                row_num += 1
                if idx % 200 == 0:
                    pct = 75 + int((idx / max(len(sorted_freq), 1)) * 10)
                    self.progress.emit(min(pct, 82))

            # Track where content words end (for chart range)
            content_end_row = row_num - 1

            # Write stop words after (no yellow highlight)
            for idx, (word, freq) in enumerate(stop_word_entries):
                word_cell = ws2.cell(row=row_num, column=1, value=word)
                word_cell.font = data_font
                word_cell.border = thin_border

                freq_cell = ws2.cell(row=row_num, column=2, value=freq)
                freq_cell.font = data_font
                freq_cell.alignment = Alignment(horizontal="center")
                freq_cell.border = thin_border

                row_num += 1

            # Column widths
            ws2.column_dimensions["A"].width = 20
            ws2.column_dimensions["B"].width = 15
            self.progress.emit(87)

            # ── Create Bar Chart (content words only) ────────────────────
            self.status_message.emit("Creating chart...")

            chart = BarChart()
            chart.type = "col"
            chart.title = "Word Frequencies Chart"
            chart.y_axis.title = "Frequency"
            chart.x_axis.title = "Word"
            chart.width = 15
            chart.height = 7.5

            # Chart covers only content words (rows 2 to content_end_row)
            data_ref = Reference(ws2, min_col=2, min_row=2, max_row=content_end_row)
            cats_ref = Reference(ws2, min_col=1, min_row=2, max_row=content_end_row)

            chart.add_data(data_ref, titles_from_data=False)
            chart.set_categories(cats_ref)

            # Name the series explicitly
            chart.series[0].title = SeriesLabel(v="FREQUENCY")

            ws2.add_chart(chart, "D1")
            self.progress.emit(88)

            # ── Sheet 3: AI Speech Analysis (Cerebras) ─────────────────────
            if self.api_key:
                self.status_message.emit("Analysing speech with Cerebras AI...")
                try:
                    llm = ChatCerebras(
                        model="gpt-oss-120b",
                        api_key=self.api_key,
                        temperature=0.3,
                    )

                    # Truncate text if very long (API limit)
                    analysis_text = text[:100000] if len(text) > 100000 else text

                    prompt = (
                        "You are an expert political speech analyst. "
                        "Analyse the following presidential speech and provide:\n\n"
                        "1. CONTENT: A comprehensive summary of this speech "
                        "(3-5 sentences covering the main topics, arguments, "
                        "and key points discussed).\n\n"
                        "2. PURPOSE: What is the purpose of this speech? "
                        "(2-3 sentences explaining the speaker's intent, "
                        "the target audience, and the desired outcome).\n\n"
                        f"SPEECH TEXT:\n{analysis_text}\n\n"
                        "Respond in this exact format:\n"
                        "CONTENT: <your summary>\n"
                        "PURPOSE: <your analysis>"
                    )

                    response = llm.invoke([HumanMessage(content=prompt)])
                    ai_text = response.content.strip()

                    # Parse response
                    content_summary = ""
                    purpose_text = ""
                    if "PURPOSE:" in ai_text:
                        parts = ai_text.split("PURPOSE:", 1)
                        content_summary = parts[0].replace("CONTENT:", "").strip()
                        purpose_text = parts[1].strip()
                    else:
                        content_summary = ai_text
                        purpose_text = "Could not determine purpose."

                    self.progress.emit(95)

                    # Create the Speech Analysis sheet
                    ws3 = wb.create_sheet(title="Speech Analysis")

                    # Style definitions
                    analysis_header_font = XlFont(
                        name="Calibri", size=12, bold=True, color="FFFFFF"
                    )
                    analysis_fill = PatternFill(
                        start_color="1F4E79", end_color="1F4E79", fill_type="solid"
                    )
                    analysis_data_font = XlFont(name="Calibri", size=11)
                    wrap_alignment = Alignment(
                        horizontal="left", vertical="top", wrap_text=True
                    )
                    center_alignment = Alignment(horizontal="center", vertical="center")
                    thin_border_ai = Border(
                        left=Side(style="thin"),
                        right=Side(style="thin"),
                        top=Side(style="thin"),
                        bottom=Side(style="thin"),
                    )

                    # Headers
                    for col_idx, header_text in enumerate(
                        ["Category", "Description"], start=1
                    ):
                        cell = ws3.cell(row=1, column=col_idx, value=header_text)
                        cell.font = analysis_header_font
                        cell.fill = analysis_fill
                        cell.alignment = center_alignment
                        cell.border = thin_border_ai

                    # Row 2: Content
                    cat_cell = ws3.cell(row=2, column=1, value="Content")
                    cat_cell.font = XlFont(name="Calibri", size=11, bold=True)
                    cat_cell.alignment = center_alignment
                    cat_cell.border = thin_border_ai

                    desc_cell = ws3.cell(row=2, column=2, value=content_summary)
                    desc_cell.font = analysis_data_font
                    desc_cell.alignment = wrap_alignment
                    desc_cell.border = thin_border_ai

                    # Row 3: Purpose
                    cat_cell2 = ws3.cell(row=3, column=1, value="Purpose")
                    cat_cell2.font = XlFont(name="Calibri", size=11, bold=True)
                    cat_cell2.alignment = center_alignment
                    cat_cell2.border = thin_border_ai

                    desc_cell2 = ws3.cell(row=3, column=2, value=purpose_text)
                    desc_cell2.font = analysis_data_font
                    desc_cell2.alignment = wrap_alignment
                    desc_cell2.border = thin_border_ai

                    # Column widths
                    ws3.column_dimensions["A"].width = 15
                    ws3.column_dimensions["B"].width = 80
                    ws3.row_dimensions[2].height = 80
                    ws3.row_dimensions[3].height = 60

                except Exception as ai_err:
                    # AI analysis is optional — don't fail the whole export
                    err_msg = f"AI analysis skipped: {str(ai_err)}"
                    print(f"DEBUG: {err_msg}")
                    self.status_message.emit(err_msg[:60])
            else:
                print("DEBUG: No API key provided — skipping AI analysis.")
                self.status_message.emit("No API key — skipping AI analysis.")

            self.progress.emit(98)

            # ── Final Step: Save workbook ────────────────────────────────
            self.status_message.emit("Saving file...")
            wb.save(self.output_path)
            self.progress.emit(100)

            self.finished_signal.emit(self.output_path)

        except Exception as e:
            self.error_signal.emit(str(e))


# ─────────────────────────────────────────────────────────────────────────────
# Custom Animated Progress Bar
# ─────────────────────────────────────────────────────────────────────────────
class AnimatedProgressBar(QProgressBar):
    """A visually enhanced progress bar with smooth gradient and glow."""

    def __init__(self, parent=None):
        super().__init__(parent)
        self.setTextVisible(False)
        self.setFixedHeight(28)
        self.setStyleSheet("""
            QProgressBar {
                border: 1px solid #3a3a5c;
                border-radius: 14px;
                background-color: #1e1e2f;
                padding: 2px;
            }
            QProgressBar::chunk {
                border-radius: 12px;
                background: qlineargradient(
                    x1:0, y1:0, x2:1, y2:0,
                    stop:0 #00c853,
                    stop:0.5 #69f0ae,
                    stop:1 #00e676
                );
            }
        """)


# ─────────────────────────────────────────────────────────────────────────────
# Main Application Window
# ─────────────────────────────────────────────────────────────────────────────
class MainWindow(QMainWindow):
    """Premium Words Frequency Extraction Tool main window."""

    def __init__(self):
        super().__init__()
        self.worker = None
        self.input_path = None
        self.setup_ui()
        self.setup_animations()

    # ── UI Setup ─────────────────────────────────────────────────────────
    def setup_ui(self):
        self.setWindowTitle("Words Frequency Extraction Tool")
        self.setFixedSize(720, 660)
        self.setWindowIcon(self._create_icon())

        # Central widget with gradient background
        central = QWidget()
        self.setCentralWidget(central)

        # Main gradient background
        central.setStyleSheet("""
            QWidget#centralWidget {
                background: qlineargradient(
                    x1:0, y1:0, x2:1, y2:1,
                    stop:0 #0f0c29,
                    stop:0.5 #302b63,
                    stop:1 #24243e
                );
            }
        """)
        central.setObjectName("centralWidget")

        main_layout = QVBoxLayout(central)
        main_layout.setContentsMargins(30, 25, 30, 25)
        main_layout.setSpacing(0)

        # ── Glass Card Container ─────────────────────────────────────────
        card = QFrame()
        card.setObjectName("glassCard")
        card.setStyleSheet("""
            QFrame#glassCard {
                background: rgba(255, 255, 255, 0.07);
                border: 1px solid rgba(255, 255, 255, 0.15);
                border-radius: 20px;
            }
        """)

        # Drop shadow for card
        shadow = QGraphicsDropShadowEffect()
        shadow.setBlurRadius(40)
        shadow.setColor(QColor(0, 0, 0, 80))
        shadow.setOffset(0, 8)
        card.setGraphicsEffect(shadow)

        card_layout = QVBoxLayout(card)
        card_layout.setContentsMargins(40, 35, 40, 35)
        card_layout.setSpacing(18)

        # ── Feather Icon + Title ─────────────────────────────────────────
        title_layout = QHBoxLayout()
        title_layout.setAlignment(Qt.AlignCenter)

        # Feather emoji as icon
        icon_label = QLabel("🪶")
        icon_label.setStyleSheet("font-size: 32px; background: transparent;")
        title_layout.addWidget(icon_label)

        title_label = QLabel("Words Frequency Extraction Tool")
        title_label.setAlignment(Qt.AlignCenter)
        title_label.setStyleSheet("""
            font-family: 'Segoe UI', 'Inter', sans-serif;
            font-size: 26px;
            font-weight: 700;
            color: #ffffff;
            background: transparent;
            letter-spacing: 0.5px;
        """)
        title_layout.addWidget(title_label)
        card_layout.addLayout(title_layout)

        # ── Separator ───────────────────────────────────────────────────
        separator = QFrame()
        separator.setFixedHeight(1)
        separator.setStyleSheet("""
            background: qlineargradient(
                x1:0, y1:0, x2:1, y2:0,
                stop:0 transparent,
                stop:0.3 rgba(255,255,255,0.3),
                stop:0.7 rgba(255,255,255,0.3),
                stop:1 transparent
            );
        """)
        card_layout.addWidget(separator)

        # ── Instructions ─────────────────────────────────────────────────
        instructions = QLabel(
            "1. Enter your Cerebras API key for AI speech analysis.\n"
            "2. Select a DOCX file to extract word frequencies.\n"
            "3. Save the output Excel file where you want."
        )
        instructions.setAlignment(Qt.AlignCenter)
        instructions.setStyleSheet("""
            font-family: 'Segoe UI', sans-serif;
            font-size: 14px;
            color: rgba(255, 255, 255, 0.65);
            background: transparent;
            line-height: 1.6;
            padding: 5px 0;
        """)
        card_layout.addWidget(instructions)

        # ── Cerebras API Key Input ──────────────────────────────────────────
        api_key_label = QLabel("🔑  Cerebras API Key")
        api_key_label.setStyleSheet("""
            font-family: 'Segoe UI', sans-serif;
            font-size: 13px;
            font-weight: 600;
            color: rgba(255, 255, 255, 0.55);
            background: transparent;
            padding: 0 5px;
        """)
        card_layout.addWidget(api_key_label)

        self.api_key_input = QLineEdit()
        self.api_key_input.setPlaceholderText(
            "Paste your Cerebras API key here..."
        )
        self.api_key_input.setEchoMode(QLineEdit.Password)
        self.api_key_input.setFixedHeight(38)
        self.api_key_input.setStyleSheet("""
            QLineEdit {
                font-family: 'Segoe UI', monospace;
                font-size: 13px;
                color: #e0e0e0;
                background: rgba(255, 255, 255, 0.06);
                border: 1px solid rgba(255, 255, 255, 0.15);
                border-radius: 10px;
                padding: 6px 14px;
                selection-background-color: #6c63ff;
            }
            QLineEdit:focus {
                border: 1px solid rgba(108, 99, 255, 0.6);
                background: rgba(255, 255, 255, 0.09);
            }
            QLineEdit::placeholder {
                color: rgba(255, 255, 255, 0.3);
            }
        """)
        # Pre-fill from environment if available
        env_key = os.environ.get("CEREBRAS_API_KEY", "")
        if env_key:
            self.api_key_input.setText(env_key)
        card_layout.addWidget(self.api_key_input)

        card_layout.addSpacerItem(
            QSpacerItem(10, 8, QSizePolicy.Minimum, QSizePolicy.Fixed)
        )

        # ── File Status Label ────────────────────────────────────────────
        self.file_label = QLabel("No file selected.")
        self.file_label.setAlignment(Qt.AlignCenter)
        self.file_label.setWordWrap(True)
        self.file_label.setStyleSheet("""
            font-family: 'Segoe UI', sans-serif;
            font-size: 14px;
            font-weight: 600;
            color: #b0bec5;
            background: transparent;
            padding: 4px 10px;
        """)
        card_layout.addWidget(self.file_label)

        # ── Progress Bar ─────────────────────────────────────────────────
        self.progress_bar = AnimatedProgressBar()
        self.progress_bar.setValue(0)
        card_layout.addWidget(self.progress_bar)

        # ── Select File Button ───────────────────────────────────────────
        btn_layout = QHBoxLayout()
        btn_layout.setAlignment(Qt.AlignCenter)

        self.select_btn = QPushButton("  Select File")
        self.select_btn.setFixedSize(200, 48)
        self.select_btn.setCursor(Qt.PointingHandCursor)
        self.select_btn.setStyleSheet("""
            QPushButton {
                font-family: 'Segoe UI', sans-serif;
                font-size: 15px;
                font-weight: 600;
                color: #ffffff;
                background: qlineargradient(
                    x1:0, y1:0, x2:1, y2:0,
                    stop:0 #6c63ff,
                    stop:1 #9c27b0
                );
                border: none;
                border-radius: 24px;
                padding: 0 20px;
                letter-spacing: 0.5px;
            }
            QPushButton:hover {
                background: qlineargradient(
                    x1:0, y1:0, x2:1, y2:0,
                    stop:0 #7c73ff,
                    stop:1 #ab47bc
                );
            }
            QPushButton:pressed {
                background: qlineargradient(
                    x1:0, y1:0, x2:1, y2:0,
                    stop:0 #5c53ef,
                    stop:1 #8e24aa
                );
            }
            QPushButton:disabled {
                background: rgba(108, 99, 255, 0.3);
                color: rgba(255,255,255,0.4);
            }
        """)
        self.select_btn.clicked.connect(self.on_select_file)

        # Button glow effect
        btn_shadow = QGraphicsDropShadowEffect()
        btn_shadow.setBlurRadius(25)
        btn_shadow.setColor(QColor(108, 99, 255, 120))
        btn_shadow.setOffset(0, 4)
        self.select_btn.setGraphicsEffect(btn_shadow)

        btn_layout.addWidget(self.select_btn)
        card_layout.addLayout(btn_layout)

        card_layout.addSpacerItem(
            QSpacerItem(10, 10, QSizePolicy.Minimum, QSizePolicy.Fixed)
        )

        # ── Status Message Label ─────────────────────────────────────────
        self.status_label = QLabel("")
        self.status_label.setAlignment(Qt.AlignCenter)
        self.status_label.setWordWrap(True)
        self.status_label.setStyleSheet("""
            font-family: 'Segoe UI', sans-serif;
            font-size: 14px;
            font-style: italic;
            font-weight: 600;
            color: #69f0ae;
            background: transparent;
            padding: 4px 10px;
        """)
        card_layout.addWidget(self.status_label)

        # ── Output Path Label ────────────────────────────────────────────
        self.output_label = QLabel("")
        self.output_label.setAlignment(Qt.AlignCenter)
        self.output_label.setWordWrap(True)
        self.output_label.setStyleSheet("""
            font-family: 'Segoe UI', sans-serif;
            font-size: 13px;
            font-weight: 500;
            color: rgba(255, 255, 255, 0.8);
            background: transparent;
            padding: 4px 10px;
        """)
        card_layout.addWidget(self.output_label)

        # Add stretch to bottom
        card_layout.addStretch()

        main_layout.addWidget(card)

        # ── Footer ───────────────────────────────────────────────────────
        footer = QLabel("Words Frequency Extraction Tool  •  v1.0")
        footer.setAlignment(Qt.AlignCenter)
        footer.setStyleSheet("""
            font-family: 'Segoe UI', sans-serif;
            font-size: 11px;
            color: rgba(255,255,255,0.25);
            background: transparent;
            padding-top: 10px;
        """)
        main_layout.addWidget(footer)

    # ── Animations ───────────────────────────────────────────────────────
    def setup_animations(self):
        """Set up entrance animations for the window."""
        # Fade-in effect via opacity
        self.setWindowOpacity(0.0)
        self._fade_anim = QPropertyAnimation(self, b"windowOpacity")
        self._fade_anim.setDuration(600)
        self._fade_anim.setStartValue(0.0)
        self._fade_anim.setEndValue(1.0)
        self._fade_anim.setEasingCurve(QEasingCurve.OutCubic)
        self._fade_anim.start()

    # ── Create App Icon ──────────────────────────────────────────────────
    def _create_icon(self):
        """Generate a simple app icon programmatically."""
        pixmap = QPixmap(64, 64)
        pixmap.fill(Qt.transparent)
        painter = QPainter(pixmap)
        painter.setRenderHint(QPainter.Antialiasing)

        # Gradient circle
        gradient = QLinearGradient(0, 0, 64, 64)
        gradient.setColorAt(0, QColor("#6c63ff"))
        gradient.setColorAt(1, QColor("#9c27b0"))
        painter.setBrush(QBrush(gradient))
        painter.setPen(Qt.NoPen)
        painter.drawRoundedRect(4, 4, 56, 56, 16, 16)

        # "W" letter
        painter.setPen(QPen(QColor("#ffffff"), 2))
        font = QFont("Segoe UI", 26, QFont.Bold)
        painter.setFont(font)
        painter.drawText(pixmap.rect(), Qt.AlignCenter, "W")
        painter.end()

        return QIcon(pixmap)

    # ── Event Handlers ───────────────────────────────────────────────────
    def on_select_file(self):
        """Handle file selection and initiate processing."""
        file_path, _ = QFileDialog.getOpenFileName(
            self,
            "Select a DOCX Speech File",
            "",
            "Word Documents (*.docx);;All Files (*)",
        )

        if not file_path:
            return

        self.input_path = file_path

        # Update UI to show selected file
        display_path = file_path.replace("/", os.sep).replace("\\", "/")
        self.file_label.setText(f"Processing file: {display_path}")
        self.file_label.setStyleSheet("""
            font-family: 'Segoe UI', sans-serif;
            font-size: 14px;
            font-weight: 600;
            color: #e0e0e0;
            background: transparent;
            padding: 4px 10px;
        """)
        self.status_label.setText("")
        self.output_label.setText("")
        self.progress_bar.setValue(0)

        # Ask for output location
        default_name = os.path.splitext(os.path.basename(file_path))[0] + "_output.xlsx"
        default_dir = os.path.dirname(file_path)
        save_path, _ = QFileDialog.getSaveFileName(
            self,
            "Save Output Excel File",
            os.path.join(default_dir, default_name),
            "Excel Files (*.xlsx);;All Files (*)",
        )

        if not save_path:
            self.file_label.setText("No file selected.")
            self.file_label.setStyleSheet("""
                font-family: 'Segoe UI', sans-serif;
                font-size: 14px;
                font-weight: 600;
                color: #b0bec5;
                background: transparent;
                padding: 4px 10px;
            """)
            return

        # Ensure .xlsx extension
        if not save_path.lower().endswith(".xlsx"):
            save_path += ".xlsx"

        # Disable button during processing
        self.select_btn.setEnabled(False)
        self.select_btn.setText("  Processing...")

        # Start worker thread
        api_key = self.api_key_input.text().strip()
        self.worker = ExtractionWorker(file_path, save_path, api_key=api_key)
        self.worker.progress.connect(self.on_progress)
        self.worker.status_message.connect(self.on_status)
        self.worker.finished_signal.connect(self.on_finished)
        self.worker.error_signal.connect(self.on_error)
        self.worker.start()

    def on_progress(self, value):
        """Update progress bar with smooth animation."""
        self.progress_bar.setValue(value)

    def on_status(self, message):
        """Update status message during processing."""
        self.status_label.setText(message)
        self.status_label.setStyleSheet("""
            font-family: 'Segoe UI', sans-serif;
            font-size: 14px;
            font-style: italic;
            font-weight: 600;
            color: #ffab40;
            background: transparent;
            padding: 4px 10px;
        """)

    def on_finished(self, output_path):
        """Handle successful completion."""
        self.status_label.setText("Extraction completed.")
        self.status_label.setStyleSheet("""
            font-family: 'Segoe UI', sans-serif;
            font-size: 15px;
            font-style: italic;
            font-weight: 700;
            color: #69f0ae;
            background: transparent;
            padding: 4px 10px;
        """)

        display_path = output_path.replace("\\", "/")
        self.output_label.setText(f"The output file is saved at: {display_path}")
        self.output_label.setStyleSheet("""
            font-family: 'Segoe UI', sans-serif;
            font-size: 13px;
            font-weight: 600;
            color: rgba(255, 255, 255, 0.85);
            background: transparent;
            padding: 4px 10px;
        """)

        # Re-enable button
        self.select_btn.setEnabled(True)
        self.select_btn.setText("  Select File")

    def on_error(self, error_msg):
        """Handle errors during processing."""
        self.status_label.setText(f"Error: {error_msg}")
        self.status_label.setStyleSheet("""
            font-family: 'Segoe UI', sans-serif;
            font-size: 14px;
            font-style: italic;
            font-weight: 600;
            color: #ff5252;
            background: transparent;
            padding: 4px 10px;
        """)
        self.progress_bar.setValue(0)

        # Re-enable button
        self.select_btn.setEnabled(True)
        self.select_btn.setText("  Select File")


# ─────────────────────────────────────────────────────────────────────────────
# Application Entry Point
# ─────────────────────────────────────────────────────────────────────────────
def main():
    # High DPI scaling
    QApplication.setHighDpiScaleFactorRoundingPolicy(
        Qt.HighDpiScaleFactorRoundingPolicy.PassThrough
    )

    app = QApplication(sys.argv)
    app.setStyle("Fusion")

    # Dark palette for the entire application
    palette = QPalette()
    palette.setColor(QPalette.Window, QColor("#0f0c29"))
    palette.setColor(QPalette.WindowText, QColor("#ffffff"))
    palette.setColor(QPalette.Base, QColor("#1a1a2e"))
    palette.setColor(QPalette.AlternateBase, QColor("#16213e"))
    palette.setColor(QPalette.ToolTipBase, QColor("#1a1a2e"))
    palette.setColor(QPalette.ToolTipText, QColor("#ffffff"))
    palette.setColor(QPalette.Text, QColor("#ffffff"))
    palette.setColor(QPalette.Button, QColor("#1a1a2e"))
    palette.setColor(QPalette.ButtonText, QColor("#ffffff"))
    palette.setColor(QPalette.BrightText, QColor("#ff5252"))
    palette.setColor(QPalette.Highlight, QColor("#6c63ff"))
    palette.setColor(QPalette.HighlightedText, QColor("#ffffff"))
    app.setPalette(palette)

    window = MainWindow()
    window.show()

    sys.exit(app.exec())


if __name__ == "__main__":
    main()
