"""
Sensitivity Extraction GUI
PyQt5 GUI wrapper for sensitivity_extractor.py — mirrors Final.py's interface.

Usage: python sensitivity_gui.py
"""

import os
import sys
import json
import time
from pathlib import Path

from PyQt5.QtWidgets import (QApplication, QMainWindow, QWidget, QVBoxLayout,
                             QHBoxLayout, QPushButton, QLabel, QFileDialog,
                             QLineEdit, QTextEdit, QProgressBar, QMessageBox,
                             QCheckBox)
from PyQt5.QtCore import QThread, pyqtSignal

# Import pipeline components from the existing extractor
from sensitivity_extractor import (
    SensitivityPageExtractor,
    GeminiExtractor,
    PlanMatcher,
    validate_plan,
    write_to_excel,
    parse_filename,
    API_KEY,
)


class ExtractionThread(QThread):
    """Background thread for the full sensitivity extraction pipeline."""

    progress_update = pyqtSignal(str)
    progress_value = pyqtSignal(int)
    finished = pyqtSignal(bool, str)

    def __init__(self, pdf_folder, output_xlsx, api_key, resume, plan_list_path):
        super().__init__()
        self.pdf_folder = pdf_folder
        self.output_xlsx = output_xlsx
        self.api_key = api_key
        self.resume = resume
        self.plan_list_path = plan_list_path

        self.consecutive_errors = 0
        self.max_consecutive_errors = 3

    def run(self):
        try:
            pdf_folder = os.path.abspath(self.pdf_folder)
            folder_name = os.path.basename(pdf_folder)
            trimmed_folder = os.path.join(os.path.dirname(pdf_folder),
                                          f"Sensitivity_Trimmed_{folder_name}")
            json_cache = os.path.join(os.path.dirname(pdf_folder),
                                      "sensitivity_cache.json")
            Path(trimmed_folder).mkdir(parents=True, exist_ok=True)

            # Load plan list
            plan_matcher = PlanMatcher(self.plan_list_path)
            if plan_matcher.plans:
                self.progress_update.emit(
                    f"Loaded {len(plan_matcher.plans)} plans from master list")
            else:
                self.progress_update.emit(
                    "No plan list loaded — output will not include matched plan names")

            # Load cache if resuming
            if self.resume and os.path.exists(json_cache):
                with open(json_cache, "r", encoding="utf-8") as f:
                    results = json.load(f)
                self.progress_update.emit(
                    f"Resumed from cache — {len(results)} PDFs already processed")
            else:
                results = {}

            # Discover PDFs
            pdf_files = sorted([f for f in os.listdir(pdf_folder)
                                if f.lower().endswith(".pdf")])

            if not pdf_files:
                self.finished.emit(False, "No PDF files found in selected folder")
                return

            total = len(pdf_files)
            self.progress_update.emit(f"Found {total} PDFs in {pdf_folder}\n")

            extractor_pages = SensitivityPageExtractor()
            gemini = GeminiExtractor(api_key=self.api_key)

            for i, pdf_file in enumerate(pdf_files, 1):
                # Check consecutive error limit
                if self.consecutive_errors >= self.max_consecutive_errors:
                    msg = (f"Stopping — {self.max_consecutive_errors} "
                           f"consecutive API errors encountered")
                    self.progress_update.emit(f"\n✗ {msg}")
                    self._save_cache(results, json_cache)
                    self.finished.emit(False, msg)
                    return

                # Skip if cached
                if pdf_file in results and "error" not in results[pdf_file]:
                    self.progress_update.emit(
                        f"[{i}/{total}] Skipping (cached): {pdf_file}")
                    self.progress_value.emit(int(i / total * 90))
                    continue

                self.progress_update.emit(f"[{i}/{total}] Processing: {pdf_file}")

                input_path = os.path.join(pdf_folder, pdf_file)
                trimmed_path = os.path.join(trimmed_folder, pdf_file)

                # Step 1: Identify and extract sensitivity pages
                try:
                    pages_found, total_pages = extractor_pages.extract_pages(
                        input_path, trimmed_path)

                    if not pages_found:
                        self.progress_update.emit(
                            "  ⚠ No sensitivity table pages found — skipping")
                        results[pdf_file] = {
                            "error": "No sensitivity keywords found on any page",
                            "plans": []
                        }
                        self._save_cache(results, json_cache)
                        self.progress_value.emit(int(i / total * 90))
                        continue

                    page_nums = [p + 1 for p in pages_found]
                    self.progress_update.emit(
                        f"  Found sensitivity content on pages {page_nums} "
                        f"of {total_pages}")

                except Exception as e:
                    self.progress_update.emit(f"  ✗ PDF reading error: {e}")
                    results[pdf_file] = {
                        "error": f"PDF read error: {e}", "plans": []}
                    self._save_cache(results, json_cache)
                    self.progress_value.emit(int(i / total * 90))
                    continue

                # Step 2: Send to Gemini
                try:
                    extracted = gemini.extract(trimmed_path)
                    results[pdf_file] = extracted
                    self.consecutive_errors = 0

                    plan_count = len(extracted.get("plans", []))
                    self.progress_update.emit(
                        f"  ✓ Extracted {plan_count} plan(s)")

                    # Inline validation warnings
                    for plan in extracted.get("plans", []):
                        warnings = validate_plan(plan)
                        for w in warnings:
                            self.progress_update.emit(
                                f"    ⚠ {plan.get('plan_name', '?')}: {w}")

                except Exception as e:
                    self.consecutive_errors += 1
                    self.progress_update.emit(
                        f"  ✗ API error ({self.consecutive_errors}/"
                        f"{self.max_consecutive_errors}): {e}")
                    results[pdf_file] = {
                        "error": f"API error: {e}", "plans": []}

                # Save cache after each PDF
                self._save_cache(results, json_cache)
                self.progress_value.emit(int(i / total * 90))

                # Rate limiting
                time.sleep(4)

            # Step 3: Write to Excel
            self.progress_update.emit("\nWriting results to Excel...")
            write_to_excel(results, self.output_xlsx, plan_matcher=plan_matcher)
            self.progress_value.emit(100)

            # Summary
            success = sum(1 for r in results.values() if r.get("plans"))
            errors = len(results) - success
            total_plans = sum(len(r.get("plans", []))
                              for r in results.values())

            summary = (f"Done! {success}/{len(results)} PDFs extracted "
                       f"({total_plans} plan rows), {errors} errors")
            self.progress_update.emit(f"\n✓ {summary}")
            self.progress_update.emit(f"Results → {self.output_xlsx}")
            self.progress_update.emit(f"Cache   → {json_cache}")

            self.finished.emit(True, self.output_xlsx)

        except Exception as e:
            self.finished.emit(False, str(e))

    def _save_cache(self, results, path):
        with open(path, "w", encoding="utf-8") as f:
            json.dump(results, f, indent=2, ensure_ascii=False)


class MainWindow(QMainWindow):
    """Main GUI window for sensitivity extraction."""

    def __init__(self):
        super().__init__()
        self.pdf_folder = ""
        self.plan_list_path = ""
        self.json_file = ""
        self.excel_file = ""
        self.api_key = API_KEY
        self.init_ui()

    def init_ui(self):
        self.setWindowTitle("Sensitivity Analysis Extractor")
        self.setGeometry(100, 100, 850, 700)

        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        layout = QVBoxLayout()

        # ── Step 1: PDF Folder ──
        layout.addWidget(QLabel("Step 1: Select PDF Folder (ACFRs)"))

        pdf_row = QHBoxLayout()
        self.pdf_label = QLabel("No folder selected")
        self.pdf_btn = QPushButton("Browse PDF Folder")
        self.pdf_btn.clicked.connect(self.browse_pdf_folder)
        pdf_row.addWidget(self.pdf_label)
        pdf_row.addWidget(self.pdf_btn)
        layout.addLayout(pdf_row)

        # ── Step 2: Plan List (optional) ──
        layout.addWidget(QLabel("\nStep 2: Select Master Plan List (optional)"))

        plan_row = QHBoxLayout()
        self.plan_label = QLabel("No plan list selected")
        self.plan_btn = QPushButton("Browse Plan List")
        self.plan_btn.clicked.connect(self.browse_plan_list)
        plan_row.addWidget(self.plan_label)
        plan_row.addWidget(self.plan_btn)
        layout.addLayout(plan_row)

        # ── Step 3: Output filename ──
        layout.addWidget(QLabel("\nStep 3: Set Output Excel Filename"))

        output_row = QHBoxLayout()
        output_row.addWidget(QLabel("Filename:"))
        self.output_input = QLineEdit("sensitivity_results.xlsx")
        output_row.addWidget(self.output_input)
        layout.addLayout(output_row)

        # ── Resume checkbox ──
        self.resume_checkbox = QCheckBox(
            "Resume from previous run (use cached results)")
        layout.addWidget(self.resume_checkbox)

        # ── Process Button ──
        self.process_btn = QPushButton("Extract Sensitivity Data")
        self.process_btn.setStyleSheet(
            "padding: 14px; font-size: 16px; font-weight: bold; "
            "background-color: #4CAF50; color: white;")
        self.process_btn.clicked.connect(self.start_extraction)
        self.process_btn.setEnabled(False)
        layout.addWidget(self.process_btn)

        # ── Progress Bar ──
        self.progress_bar = QProgressBar()
        layout.addWidget(self.progress_bar)

        # ── Log Area ──
        layout.addWidget(QLabel("\nProcessing Log:"))
        self.log_text = QTextEdit()
        self.log_text.setReadOnly(True)
        layout.addWidget(self.log_text)

        # ── Divider ──
        layout.addWidget(QLabel("\n" + "=" * 50))
        layout.addWidget(QLabel("Step 4: Write Results to Excel"))

        # JSON file selection for merge
        json_select_row = QHBoxLayout()
        self.json_file_label = QLabel("No JSON file selected")
        self.json_file_btn = QPushButton("Browse JSON File")
        self.json_file_btn.clicked.connect(self.browse_json_file)
        json_select_row.addWidget(self.json_file_label)
        json_select_row.addWidget(self.json_file_btn)
        layout.addLayout(json_select_row)

        # Excel output file selection for merge
        excel_row = QHBoxLayout()
        self.excel_label = QLabel("No Excel file selected")
        self.excel_btn = QPushButton("Browse Output Excel")
        self.excel_btn.clicked.connect(self.browse_excel_file)
        excel_row.addWidget(self.excel_label)
        excel_row.addWidget(self.excel_btn)
        layout.addLayout(excel_row)

        # Plan list for merge (reuses the one from extraction if set)
        merge_plan_row = QHBoxLayout()
        self.merge_plan_label = QLabel("Plan list: (uses Step 2 selection)")
        self.merge_plan_label.setStyleSheet("font-size: 11px; color: #555;")
        merge_plan_row.addWidget(self.merge_plan_label)
        layout.addLayout(merge_plan_row)

        # Write button
        self.write_btn = QPushButton("Write JSON Data to Excel")
        self.write_btn.setStyleSheet(
            "padding: 14px; font-size: 16px; font-weight: bold; "
            "background-color: #2196F3; color: white;")
        self.write_btn.clicked.connect(self.write_to_excel)
        self.write_btn.setEnabled(False)
        layout.addWidget(self.write_btn)

        central_widget.setLayout(layout)

    # ── Browse handlers ──

    def browse_pdf_folder(self):
        folder = QFileDialog.getExistingDirectory(self, "Select PDF Folder")
        if folder:
            self.pdf_folder = folder
            self.pdf_label.setText(os.path.basename(folder))
            self.pdf_label.setToolTip(folder)
            self.process_btn.setEnabled(True)
            self.log_text.append(f"Selected PDF folder: {folder}\n")

    def browse_plan_list(self):
        file, _ = QFileDialog.getOpenFileName(
            self, "Select Plan List",
            "", "Excel/CSV Files (*.xlsx *.xls *.csv);;All Files (*)")
        if file:
            self.plan_list_path = file
            self.plan_label.setText(os.path.basename(file))
            self.plan_label.setToolTip(file)
            self.merge_plan_label.setText(f"Plan list: {os.path.basename(file)}")
            self.log_text.append(f"Selected plan list: {file}\n")

    def browse_json_file(self):
        file, _ = QFileDialog.getOpenFileName(
            self, "Select JSON File", "",
            "JSON Files (*.json);;All Files (*)")
        if file:
            self.json_file = file
            self.json_file_label.setText(os.path.basename(file))
            self.json_file_label.setToolTip(file)
            self.check_write_button()
            self.log_text.append(f"Selected JSON file: {file}\n")

    def browse_excel_file(self):
        file, _ = QFileDialog.getSaveFileName(
            self, "Select Output Excel File", "sensitivity_results.xlsx",
            "Excel Files (*.xlsx);;All Files (*)")
        if file:
            self.excel_file = file
            self.excel_label.setText(os.path.basename(file))
            self.excel_label.setToolTip(file)
            self.check_write_button()
            self.log_text.append(f"Selected Excel output: {file}\n")

    def check_write_button(self):
        if hasattr(self, 'json_file') and self.json_file and \
           hasattr(self, 'excel_file') and self.excel_file:
            self.write_btn.setEnabled(True)

    def write_to_excel(self):
        if not hasattr(self, 'json_file') or not self.json_file:
            QMessageBox.warning(self, "Warning",
                                "Please select a JSON file first!")
            return
        if not hasattr(self, 'excel_file') or not self.excel_file:
            QMessageBox.warning(self, "Warning",
                                "Please select an output Excel file first!")
            return

        try:
            self.write_btn.setEnabled(False)
            self.log_text.append("\n" + "=" * 50)
            self.log_text.append("Writing JSON data to Excel...\n")

            # Load JSON
            with open(self.json_file, 'r', encoding='utf-8') as f:
                results = json.load(f)

            # Load plan matcher if available
            plan_matcher = None
            if self.plan_list_path:
                plan_matcher = PlanMatcher(self.plan_list_path)
                if plan_matcher.plans:
                    self.log_text.append(
                        f"Using plan list with {len(plan_matcher.plans)} plans")

            # Write to Excel
            write_to_excel(results, self.excel_file, plan_matcher=plan_matcher)

            total = len(results)
            success = sum(1 for r in results.values() if r.get("plans"))
            total_plans = sum(len(r.get("plans", []))
                              for r in results.values())

            self.log_text.append(
                f"\n✓ Written {success} PDFs ({total_plans} plan rows) "
                f"to {self.excel_file}")

            QMessageBox.information(
                self, "Success",
                f"Excel written!\n\n{success}/{total} PDFs "
                f"({total_plans} plan rows)\n\n{self.excel_file}")

        except Exception as e:
            QMessageBox.critical(
                self, "Error", f"Failed to write Excel: {str(e)}")
            self.log_text.append(f"\n✗ Error: {str(e)}")

        finally:
            self.write_btn.setEnabled(True)

    # ── Extraction ──

    def start_extraction(self):
        if not self.pdf_folder:
            QMessageBox.warning(self, "Warning",
                                "Please select a PDF folder first!")
            return

        output_filename = self.output_input.text().strip()
        if not output_filename:
            output_filename = "sensitivity_results.xlsx"

        # Put output in the same parent directory as the PDF folder
        output_path = os.path.join(
            os.path.dirname(self.pdf_folder), output_filename)

        self.process_btn.setEnabled(False)
        self.pdf_btn.setEnabled(False)
        self.plan_btn.setEnabled(False)
        self.progress_bar.setValue(0)
        self.log_text.clear()
        self.log_text.append("Starting sensitivity extraction...\n")

        self.thread = ExtractionThread(
            pdf_folder=self.pdf_folder,
            output_xlsx=output_path,
            api_key=self.api_key,
            resume=self.resume_checkbox.isChecked(),
            plan_list_path=self.plan_list_path or None,
        )
        self.thread.progress_update.connect(self.update_log)
        self.thread.progress_value.connect(self.progress_bar.setValue)
        self.thread.finished.connect(self.extraction_finished)
        self.thread.start()

    def extraction_finished(self, success, message):
        self.process_btn.setEnabled(True)
        self.pdf_btn.setEnabled(True)
        self.plan_btn.setEnabled(True)

        if success:
            # Auto-populate the JSON file for Step 4 (cache file)
            cache_path = os.path.join(
                os.path.dirname(self.pdf_folder), "sensitivity_cache.json")
            if os.path.exists(cache_path):
                self.json_file = cache_path
                self.json_file_label.setText(os.path.basename(cache_path))
                self.json_file_label.setToolTip(cache_path)
                self.check_write_button()

            QMessageBox.information(
                self, "Success",
                f"Extraction completed!\n\nResults saved to:\n{message}")
        else:
            QMessageBox.critical(
                self, "Error", f"Extraction failed: {message}")

    def update_log(self, message):
        self.log_text.append(message)
        self.log_text.verticalScrollBar().setValue(
            self.log_text.verticalScrollBar().maximum())


def main():
    app = QApplication(sys.argv)
    window = MainWindow()
    window.showMaximized()
    sys.exit(app.exec_())


if __name__ == "__main__":
    main()
