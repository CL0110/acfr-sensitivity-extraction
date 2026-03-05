import sys
import json
import re
from pathlib import Path
from PyQt5.QtWidgets import (QApplication, QMainWindow, QWidget, QVBoxLayout, 
                             QHBoxLayout, QComboBox, QPushButton, QLabel, 
                             QScrollArea, QGridLayout, QMessageBox, QFrame,
                             QSlider, QLineEdit, QFileDialog)
from PyQt5.QtCore import Qt
from PyQt5.QtGui import QPixmap, QImage, QFont
try:
    import fitz
except ImportError:
    import pymupdf as fitz
import subprocess
import platform
import os


def find_pdf_page_by_printed_number(pdf_path, printed_page_num):
    """Scan the PDF to find which physical page has the given printed page number.
    
    ACFRs typically print page numbers in headers/footers. This searches each page's
    text for the printed number in the header/footer region and returns the physical
    1-based page index.
    
    Returns the physical page number (1-based) or None if not found.
    """
    try:
        doc = fitz.open(str(pdf_path))
        target = str(int(printed_page_num))
        
        for page_idx in range(len(doc)):
            page = doc[page_idx]
            blocks = page.get_text("blocks")
            page_height = page.rect.height
            
            for block in blocks:
                # block layout: (x0, y0, x1, y1, text, block_no, block_type)
                block_y = block[1]
                block_text = str(block[4]) if len(block) > 4 else ""
                
                # Only check header (top 12%) and footer (bottom 15%) regions
                is_footer = block_y > page_height * 0.85
                is_header = block_y < page_height * 0.12
                
                if (is_footer or is_header) and re.search(rf'\b{target}\b', block_text):
                    doc.close()
                    return page_idx + 1  # Return 1-based
        
        doc.close()
        return None
        
    except Exception:
        return None


class SensitivityCheckerApp(QMainWindow):
    """QA tool for reviewing sensitivity analysis data extracted from ACFR PDFs.
    
    Loads sensitivity_cache.json (or any JSON with the sensitivity extractor's format)
    and displays the extracted discount-rate sensitivity data alongside the source PDF pages.
    
    Handles the printed-vs-physical page number mismatch common in ACFRs by scanning
    for printed page numbers in headers/footers and caching the resolved physical page.
    """

    def __init__(self, json_file_path=None, pdf_directory=None):
        super().__init__()
        self.json_file_path = json_file_path
        self.pdf_directory = pdf_directory or 'Final to Extract'
        self.pdf_data = {}
        self.zoom_level = 170
        self.current_pages = []
        self.highlight_terms = [
            "sensitivity", "discount rate", "net pension liability",
            "1% decrease", "1% increase", "net pension asset",
        ]
        self.field_editors = {}

        if json_file_path and Path(json_file_path).exists():
            self.pdf_data = self.load_json_data()

        self.init_ui()

    # ── Data I/O ──────────────────────────────────

    def load_json_data(self):
        try:
            with open(self.json_file_path, 'r') as f:
                return json.load(f)
        except Exception as e:
            QMessageBox.critical(None, "Error", f"Failed to load JSON file: {str(e)}")
            return {}

    def save_json_data(self):
        try:
            with open(self.json_file_path, 'w') as f:
                json.dump(self.pdf_data, f, indent=2)
            return True
        except Exception as e:
            QMessageBox.critical(self, "Error", f"Failed to save JSON file: {str(e)}")
            return False

    # ── File / folder browsers ────────────────────

    def browse_json_file(self):
        file_path, _ = QFileDialog.getOpenFileName(
            self, "Select JSON File", "", "JSON Files (*.json);;All Files (*)")
        if file_path:
            self.json_file_path = file_path
            self.json_path_label.setText(f"JSON: {os.path.basename(file_path)}")
            self.json_path_label.setToolTip(file_path)
            self.pdf_data = self.load_json_data()
            self.pdf_dropdown.clear()
            self.pdf_dropdown.addItems(list(self.pdf_data.keys()))
            self.update_pdf_counter()
            self.update_data_display()

    def browse_pdf_folder(self):
        folder_path = QFileDialog.getExistingDirectory(
            self, "Select PDF Folder", self.pdf_directory)
        if folder_path:
            self.pdf_directory = folder_path
            self.pdf_folder_label.setText(f"PDF Folder: {os.path.basename(folder_path)}")
            self.pdf_folder_label.setToolTip(folder_path)

    # ── Navigation ────────────────────────────────

    def select_previous_pdf(self):
        idx = self.pdf_dropdown.currentIndex()
        if idx > 0:
            self.pdf_dropdown.setCurrentIndex(idx - 1)

    def select_next_pdf(self):
        idx = self.pdf_dropdown.currentIndex()
        if idx < self.pdf_dropdown.count() - 1:
            self.pdf_dropdown.setCurrentIndex(idx + 1)

    def update_pdf_counter(self):
        current = self.pdf_dropdown.currentIndex()
        total = self.pdf_dropdown.count()
        self.pdf_counter_label.setText(f"({current + 1} of {total})" if total else "(0 of 0)")

    # ── UI setup ──────────────────────────────────

    def init_ui(self):
        self.setWindowTitle("Sensitivity Data Checker")
        self.setGeometry(100, 100, 1600, 900)

        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        main_layout = QVBoxLayout()
        central_widget.setLayout(main_layout)

        # ── Top bar ──
        top_bar = QHBoxLayout()
        top_bar.setSpacing(10)

        json_browse_btn = QPushButton("Browse JSON")
        json_browse_btn.setStyleSheet("padding: 8px; font-size: 12px; background-color: #607D8B; color: white; font-weight: bold;")
        json_browse_btn.clicked.connect(self.browse_json_file)

        self.json_path_label = QLabel("No JSON")
        self.json_path_label.setStyleSheet("font-size: 11px; color: #555;")
        if self.json_file_path:
            self.json_path_label.setText(os.path.basename(self.json_file_path))
            self.json_path_label.setToolTip(self.json_file_path)

        pdf_folder_btn = QPushButton("Browse Folder")
        pdf_folder_btn.setStyleSheet("padding: 8px; font-size: 12px; background-color: #607D8B; color: white; font-weight: bold;")
        pdf_folder_btn.clicked.connect(self.browse_pdf_folder)

        self.pdf_folder_label = QLabel(os.path.basename(self.pdf_directory))
        self.pdf_folder_label.setStyleSheet("font-size: 11px; color: #555;")
        self.pdf_folder_label.setToolTip(self.pdf_directory)

        dropdown_label = QLabel("PDF:")
        dropdown_label.setStyleSheet("font-weight: bold; font-size: 14px;")
        self.pdf_dropdown = QComboBox()
        self.pdf_dropdown.setStyleSheet("font-size: 14px; padding: 5px;")
        self.pdf_dropdown.setMinimumWidth(300)
        if self.pdf_data:
            self.pdf_dropdown.addItems(list(self.pdf_data.keys()))
        self.pdf_dropdown.currentTextChanged.connect(self.update_data_display)

        self.pdf_counter_label = QLabel()
        self.pdf_counter_label.setStyleSheet("font-weight: bold; font-size: 14px; color: #2196F3; padding: 5px;")

        prev_btn = QPushButton("← Prev")
        prev_btn.setStyleSheet("padding: 8px; font-size: 12px; background-color: #9C27B0; color: white; font-weight: bold;")
        prev_btn.clicked.connect(self.select_previous_pdf)

        next_btn = QPushButton("Next →")
        next_btn.setStyleSheet("padding: 8px; font-size: 12px; background-color: #FF9800; color: white; font-weight: bold;")
        next_btn.clicked.connect(self.select_next_pdf)

        for w in [json_browse_btn, self.json_path_label, QLabel("|"),
                   pdf_folder_btn, self.pdf_folder_label, QLabel("|"),
                   dropdown_label, self.pdf_dropdown, self.pdf_counter_label,
                   prev_btn, next_btn]:
            top_bar.addWidget(w)
        top_bar.addStretch()
        main_layout.addLayout(top_bar)

        # Separator
        sep = QFrame(); sep.setFrameShape(QFrame.HLine); sep.setFrameShadow(QFrame.Sunken)
        main_layout.addWidget(sep)

        # ── Data grid ──
        self.data_frame = QFrame()
        self.data_frame.setFrameStyle(QFrame.Box | QFrame.Raised)
        self.data_frame.setLineWidth(2)
        self.data_layout = QGridLayout()
        self.data_layout.setVerticalSpacing(5)
        self.data_layout.setHorizontalSpacing(10)
        self.data_frame.setLayout(self.data_layout)
        main_layout.addWidget(self.data_frame)

        # ── Page controls row ──
        page_control_layout = QHBoxLayout()

        open_pdf_btn = QPushButton("Open PDF in External Viewer")
        open_pdf_btn.setStyleSheet("padding: 12px; font-size: 16px; background-color: #4CAF50; color: white; font-weight: bold;")
        open_pdf_btn.clicked.connect(self.open_pdf_external)

        view_page_btn = QPushButton("View Source Page (In App)")
        view_page_btn.setStyleSheet("padding: 12px; font-size: 16px; background-color: #2196F3; color: white; font-weight: bold;")
        view_page_btn.clicked.connect(self.view_source_page)

        # Manual page navigation
        goto_label = QLabel("Go to PDF page:")
        goto_label.setStyleSheet("font-weight: bold; font-size: 14px; padding-left: 20px;")
        self.goto_input = QLineEdit()
        self.goto_input.setStyleSheet("font-size: 14px; padding: 5px;")
        self.goto_input.setFixedWidth(80)
        self.goto_input.setPlaceholderText("#")
        self.goto_input.returnPressed.connect(self.goto_page)

        goto_btn = QPushButton("Go")
        goto_btn.setStyleSheet("padding: 10px 18px; font-size: 14px; font-weight: bold; background-color: #FF5722; color: white;")
        goto_btn.clicked.connect(self.goto_page)

        page_control_layout.addWidget(open_pdf_btn)
        page_control_layout.addWidget(view_page_btn)
        page_control_layout.addWidget(goto_label)
        page_control_layout.addWidget(self.goto_input)
        page_control_layout.addWidget(goto_btn)
        page_control_layout.addStretch()
        main_layout.addLayout(page_control_layout)

        # ── Zoom controls ──
        zoom_layout = QHBoxLayout()

        zoom_label = QLabel("Zoom:")
        zoom_label.setStyleSheet("font-weight: bold; font-size: 16px;")

        self.zoom_out_btn = QPushButton("−")
        self.zoom_out_btn.setStyleSheet("padding: 8px 15px; font-size: 20px; font-weight: bold;")
        self.zoom_out_btn.clicked.connect(lambda: self.zoom_slider.setValue(max(self.zoom_slider.value() - 10, 50)))

        self.zoom_slider = QSlider(Qt.Horizontal)
        self.zoom_slider.setMinimum(50)
        self.zoom_slider.setMaximum(200)
        self.zoom_slider.setValue(self.zoom_level)
        self.zoom_slider.setTickPosition(QSlider.TicksBelow)
        self.zoom_slider.setTickInterval(25)
        self.zoom_slider.valueChanged.connect(self.on_zoom_changed)

        self.zoom_in_btn = QPushButton("+")
        self.zoom_in_btn.setStyleSheet("padding: 8px 15px; font-size: 20px; font-weight: bold;")
        self.zoom_in_btn.clicked.connect(lambda: self.zoom_slider.setValue(min(self.zoom_slider.value() + 10, 200)))

        self.zoom_value_label = QLabel(f"{self.zoom_level}%")
        self.zoom_value_label.setStyleSheet("font-size: 16px; font-weight: bold; min-width: 60px;")

        reset_btn = QPushButton("Reset")
        reset_btn.setStyleSheet("padding: 8px; font-size: 14px;")
        reset_btn.clicked.connect(lambda: self.zoom_slider.setValue(100))

        for w in [zoom_label, self.zoom_out_btn, self.zoom_slider,
                   self.zoom_in_btn, self.zoom_value_label, reset_btn]:
            zoom_layout.addWidget(w)
        zoom_layout.addStretch()
        main_layout.addLayout(zoom_layout)

        # ── Page viewer ──
        self.page_scroll_area = QScrollArea()
        self.page_scroll_area.setWidgetResizable(True)
        self.page_content = QWidget()
        self.page_layout = QVBoxLayout()
        self.page_content.setLayout(self.page_layout)
        self.page_scroll_area.setWidget(self.page_content)
        main_layout.addWidget(self.page_scroll_area)

        # Initial display
        if self.pdf_data:
            self.update_pdf_counter()
            self.update_data_display()

    # ── Data display ──────────────────────────────

    def update_data_display(self):
        self.update_pdf_counter()

        # Clear previous widgets
        for i in reversed(range(self.data_layout.count())):
            self.data_layout.itemAt(i).widget().setParent(None)
        for i in reversed(range(self.page_layout.count())):
            self.page_layout.itemAt(i).widget().setParent(None)

        self.current_pages = []
        self.field_editors = {}

        selected_pdf = self.pdf_dropdown.currentText()
        if not selected_pdf or not self.pdf_data:
            return

        data = self.pdf_data[selected_pdf]

        # Check for error entries
        if "error" in data and not data.get("plans"):
            error_label = QLabel(f"Error: {data['error']}")
            error_label.setStyleSheet("color: red; font-size: 15px; font-weight: bold; padding: 10px;")
            self.data_layout.addWidget(error_label, 0, 0, 1, 9)
            return

        header_style = "font-weight: bold; font-size: 15px; background-color: #E0E0E0; padding: 5px;"

        # ── Row 0: Source page, actual PDF page, dollar unit (all editable) ──
        src_label = QLabel("Doc Page (printed):")
        src_label.setStyleSheet("font-weight: bold; font-size: 14px; padding: 5px; color: #1565C0;")
        self.data_layout.addWidget(src_label, 0, 0)

        self.source_page_editor = QLineEdit(str(data.get("source_page", "")))
        self.source_page_editor.setStyleSheet("font-size: 14px; padding: 5px; max-width: 80px;")
        self.source_page_editor.setToolTip("Printed page number from the document (what Gemini reported)")
        self.data_layout.addWidget(self.source_page_editor, 0, 1)

        pdf_page_label = QLabel("Actual PDF Page:")
        pdf_page_label.setStyleSheet("font-weight: bold; font-size: 14px; padding: 5px; color: #C62828;")
        self.data_layout.addWidget(pdf_page_label, 0, 2)

        self.actual_pdf_page_editor = QLineEdit(str(data.get("actual_pdf_page", "")))
        self.actual_pdf_page_editor.setStyleSheet("font-size: 14px; padding: 5px; max-width: 80px;")
        self.actual_pdf_page_editor.setPlaceholderText("auto")
        self.actual_pdf_page_editor.setToolTip(
            "Physical PDF page number. Leave blank to auto-detect from printed page.")
        self.data_layout.addWidget(self.actual_pdf_page_editor, 0, 3)

        unit_label = QLabel("Dollar Unit:")
        unit_label.setStyleSheet("font-weight: bold; font-size: 14px; padding: 5px; color: #1565C0;")
        self.data_layout.addWidget(unit_label, 0, 4)

        self.dollar_unit_editor = QLineEdit(str(data.get("dollar_unit", "")))
        self.dollar_unit_editor.setStyleSheet("font-size: 14px; padding: 5px; max-width: 120px;")
        self.data_layout.addWidget(self.dollar_unit_editor, 0, 5)

        save_meta_btn = QPushButton("Save Page/Unit")
        save_meta_btn.setStyleSheet("padding: 6px 12px; font-size: 13px; font-weight: bold; background-color: #FF9800; color: white;")
        save_meta_btn.clicked.connect(self.update_meta)
        self.data_layout.addWidget(save_meta_btn, 0, 6, 1, 2)

        # ── Row 1: Column headers ──
        headers = ["Plan Name", "Rate −1%", "Current Rate", "Rate +1%",
                    "NPL (−1%)", "NPL (Current)", "NPL (+1%)", "Actions"]
        col_stretches = [3, 1, 1, 1, 2, 2, 2, 2]

        for col, (header, stretch) in enumerate(zip(headers, col_stretches)):
            lbl = QLabel(header)
            lbl.setStyleSheet(header_style)
            self.data_layout.addWidget(lbl, 1, col)
            self.data_layout.setColumnStretch(col, stretch)

        # ── Row 2+: One row per plan ──
        plans = data.get("plans", [])
        field_keys = ["plan_name", "rateminus1", "current_discount_rate", "rateplus1",
                      "nplminus1", "npl_current", "nplplus1"]

        for plan_idx, plan in enumerate(plans):
            grid_row = plan_idx + 2

            for col, key in enumerate(field_keys):
                value = plan.get(key, "")
                editor = QLineEdit(str(value) if value is not None else "")
                editor.setStyleSheet("font-size: 15px; padding: 5px;")
                self.field_editors[(plan_idx, key)] = editor
                self.data_layout.addWidget(editor, grid_row, col)

            update_btn = QPushButton("Update")
            update_btn.setStyleSheet("padding: 6px 12px; font-size: 13px; font-weight: bold; background-color: #4CAF50; color: white;")
            update_btn.clicked.connect(lambda checked, pi=plan_idx: self.update_plan(pi))
            self.data_layout.addWidget(update_btn, grid_row, 7)

    def update_meta(self):
        """Save changes to source_page, actual_pdf_page, and dollar_unit."""
        selected_pdf = self.pdf_dropdown.currentText()
        if not selected_pdf:
            return

        data = self.pdf_data[selected_pdf]
        changes = []

        new_source = self.source_page_editor.text().strip()
        old_source = str(data.get("source_page", ""))
        if new_source != old_source:
            try:
                data["source_page"] = int(new_source)
            except ValueError:
                data["source_page"] = new_source
            changes.append(f"Doc Page: {old_source} → {new_source}")

        new_actual = self.actual_pdf_page_editor.text().strip()
        old_actual = str(data.get("actual_pdf_page", ""))
        if new_actual != old_actual:
            if new_actual:
                try:
                    data["actual_pdf_page"] = int(new_actual)
                except ValueError:
                    data["actual_pdf_page"] = new_actual
            else:
                data.pop("actual_pdf_page", None)
            changes.append(f"PDF Page: {old_actual} → {new_actual or '(auto)'}")

        new_unit = self.dollar_unit_editor.text().strip()
        old_unit = str(data.get("dollar_unit", ""))
        if new_unit != old_unit:
            data["dollar_unit"] = new_unit
            changes.append(f"Dollar Unit: {old_unit} → {new_unit}")

        if not changes:
            QMessageBox.information(self, "No Changes", "No values were changed.")
            return

        if self.save_json_data():
            QMessageBox.information(self, "Updated", "\n".join(changes))

    def update_plan(self, plan_idx):
        selected_pdf = self.pdf_dropdown.currentText()
        if not selected_pdf:
            return

        plans = self.pdf_data[selected_pdf].get("plans", [])
        if plan_idx >= len(plans):
            return

        plan = plans[plan_idx]
        field_keys = ["plan_name", "rateminus1", "current_discount_rate", "rateplus1",
                      "nplminus1", "npl_current", "nplplus1"]
        numeric_keys = {"rateminus1", "current_discount_rate", "rateplus1",
                        "nplminus1", "npl_current", "nplplus1"}

        changes = []
        for key in field_keys:
            editor = self.field_editors.get((plan_idx, key))
            if not editor:
                continue
            new_val = editor.text().strip()
            old_val = plan.get(key, "")

            if key in numeric_keys and new_val:
                try:
                    new_val = float(new_val)
                    if new_val == int(new_val) and '.' not in editor.text():
                        new_val = int(new_val)
                except ValueError:
                    pass

            if str(new_val) != str(old_val):
                changes.append(f"{key}: {old_val} → {new_val}")
                plan[key] = new_val

        if not changes:
            QMessageBox.information(self, "No Changes", "No values were changed.")
            return

        if self.save_json_data():
            QMessageBox.information(self, "Updated",
                f"Plan '{plan.get('plan_name', plan_idx)}':\n\n" + "\n".join(changes))

    # ── PDF viewing ───────────────────────────────

    def get_pdf_path(self):
        selected_pdf = self.pdf_dropdown.currentText()
        pdf_path = os.path.join(self.pdf_directory, selected_pdf)
        if not os.path.exists(pdf_path):
            QMessageBox.warning(self, "File Not Found",
                f"PDF file not found: {pdf_path}\n\n"
                f"Please ensure the PDF is in the correct directory.")
            return None
        return pdf_path

    def open_pdf_external(self):
        pdf_path = self.get_pdf_path()
        if not pdf_path:
            return
        try:
            if platform.system() == 'Windows':
                subprocess.Popen(['start', '', str(pdf_path)], shell=True)
            elif platform.system() == 'Darwin':
                subprocess.Popen(['open', str(pdf_path)])
            else:
                subprocess.Popen(['xdg-open', str(pdf_path)])
        except Exception as e:
            QMessageBox.critical(self, "Error", f"Failed to open PDF: {str(e)}")

    def _resolve_pdf_page(self, pdf_path):
        """Figure out which physical PDF page to display.
        
        Priority:
        1. actual_pdf_page if manually set in JSON
        2. Auto-detect by scanning PDF for printed page number in headers/footers
        3. Fall back to source_page as physical page (original behavior, may be wrong)
        """
        selected_pdf = self.pdf_dropdown.currentText()
        data = self.pdf_data.get(selected_pdf, {})

        # 1. Use manually set actual_pdf_page
        actual = data.get("actual_pdf_page")
        if actual:
            try:
                return int(actual), "actual_pdf_page"
            except (ValueError, TypeError):
                pass

        # 2. Auto-detect from printed page number
        source_page = data.get("source_page")
        if source_page:
            try:
                printed_num = int(source_page)
                found = find_pdf_page_by_printed_number(pdf_path, printed_num)
                if found:
                    # Cache it so we don't scan every time
                    data["actual_pdf_page"] = found
                    self.save_json_data()
                    self.actual_pdf_page_editor.setText(str(found))
                    return found, "auto-detected"
            except (ValueError, TypeError):
                pass

            # 3. Fall back to source_page as physical page
            try:
                return int(source_page), "source_page fallback — may be wrong, check manually"
            except (ValueError, TypeError):
                pass

        return None, None

    def view_source_page(self):
        """View the source page, resolving printed → physical page number."""
        selected_pdf = self.pdf_dropdown.currentText()
        if not selected_pdf:
            return

        pdf_path = self.get_pdf_path()
        if not pdf_path:
            return

        page_num, method = self._resolve_pdf_page(pdf_path)

        if not page_num:
            QMessageBox.information(self, "No Page",
                "No source page recorded for this entry.\n\n"
                "Use 'Go to PDF page' to navigate manually.")
            return

        self._clear_and_render(pdf_path, page_num, method)

    def goto_page(self):
        """Navigate to a manually entered physical PDF page number."""
        page_text = self.goto_input.text().strip()
        if not page_text:
            return

        try:
            page_num = int(page_text)
        except ValueError:
            QMessageBox.warning(self, "Invalid Page", "Please enter a valid page number.")
            return

        pdf_path = self.get_pdf_path()
        if not pdf_path:
            return

        self._clear_and_render(pdf_path, page_num, "manual")

    def _clear_and_render(self, pdf_path, page_num, method_note=None):
        """Clear the viewer and render a single page."""
        for i in reversed(range(self.page_layout.count())):
            self.page_layout.itemAt(i).widget().setParent(None)

        try:
            self.current_pages = [page_num]
            self.render_page(pdf_path, page_num - 1, method_note=method_note)
        except Exception as e:
            QMessageBox.critical(self, "Error", f"Failed to render page: {str(e)}")

    def render_page(self, pdf_path, page_index, method_note=None):
        try:
            doc = fitz.open(str(pdf_path))
            if page_index < 0 or page_index >= len(doc):
                raise ValueError(f"Page {page_index + 1} does not exist "
                                 f"(PDF has {len(doc)} pages)")

            page = doc[page_index]

            # Highlight sensitivity-related terms
            if self.highlight_terms:
                for term in self.highlight_terms:
                    instances = page.search_for(term,
                        flags=fitz.TEXT_PRESERVE_WHITESPACE | fitz.TEXT_PRESERVE_LIGATURES)
                    for inst in instances:
                        hl = page.add_highlight_annot(inst)
                        hl.set_colors(stroke=(1, 1, 0))
                        hl.update()

            zoom_factor = 2 * (self.zoom_level / 100)
            mat = fitz.Matrix(zoom_factor, zoom_factor)
            pix = page.get_pixmap(matrix=mat)

            img = QImage(pix.samples, pix.width, pix.height, pix.stride, QImage.Format_RGB888)

            # Title showing both physical and context
            page_title_text = f"PDF Page {page_index + 1} of {len(doc)}"
            if method_note:
                page_title_text += f"  [{method_note}]"

            title = QLabel(page_title_text)
            title.setStyleSheet("font-weight: bold; font-size: 18px; margin-top: 10px;")
            title.setAlignment(Qt.AlignCenter)
            self.page_layout.addWidget(title)

            label = QLabel()
            pixmap = QPixmap.fromImage(img)
            max_width = int(800 * (self.zoom_level / 100))
            label.setPixmap(pixmap.scaledToWidth(max_width, Qt.SmoothTransformation))
            label.setAlignment(Qt.AlignCenter)
            self.page_layout.addWidget(label)

            separator = QFrame()
            separator.setFrameShape(QFrame.HLine)
            separator.setFrameShadow(QFrame.Sunken)
            self.page_layout.addWidget(separator)

            doc.close()

        except Exception as e:
            err = QLabel(f"Error rendering page {page_index + 1}: {str(e)}")
            err.setStyleSheet("color: red; font-size: 16px;")
            self.page_layout.addWidget(err)

    # ── Zoom ──────────────────────────────────────

    def on_zoom_changed(self, value):
        self.zoom_level = value
        self.zoom_value_label.setText(f"{value}%")
        if self.current_pages:
            pdf_path = self.get_pdf_path()
            if pdf_path:
                for i in reversed(range(self.page_layout.count())):
                    self.page_layout.itemAt(i).widget().setParent(None)
                for page_num in self.current_pages:
                    self.render_page(pdf_path, page_num - 1)


def main():
    app = QApplication(sys.argv)

    json_file = "sensitivity_cache.json" if Path("sensitivity_cache.json").exists() else None
    window = SensitivityCheckerApp(json_file)
    window.showMaximized()
    sys.exit(app.exec_())


if __name__ == "__main__":
    main()
