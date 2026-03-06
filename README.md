# ACFR Sensitivity Analysis Extraction Pipeline

Automated extraction of discount-rate sensitivity data from state and local government Annual Comprehensive Financial Reports (ACFRs) using AI-powered document processing.

## Overview

Public pension plans are required to disclose the sensitivity of their net pension liability (NPL) to changes in the discount rate. This information is buried across hundreds of ACFRs — lengthy PDF financial reports that vary widely in format and structure. Manually locating and transcribing this data from 200+ reports (50,000+ pages) took weeks, if not months.

This pipeline compresses that process into a single day by:

1. Scanning each ACFR PDF to identify pages containing sensitivity analysis tables

2. Extracting structured data (plan names, discount rates, NPL at ±1%) using Google's Gemini API

3. Matching extracted plan names against a master list of known pension plans

4. Validating results with automated checks and a side-by-side QA review tool

5. Outputting clean, analysis-ready data in Excel

## Repo Structure
acfr-sensitivity-extraction/
│
├── sensitivity_extractor.py     # Core extraction logic: PDF page detection,
│                                  # Gemini API integration, plan matching,
│                                  # validation, and Excel output
│
├── sensitivity_gui.py           # PyQt5 GUI for running the full extraction
│                                  # pipeline with progress tracking and caching
│
├── sensitivity_checking.py      # QA review tool: displays extracted data
│                                  # alongside source PDF pages for verification
│
├── start_sensitivity_extraction_v2.bat   # Windows launcher (extraction GUI)
├── sensitivity_checking.bat              # Windows launcher (QA tool)
│
├── requirements.txt
└── README.md

## How It Works

### Extraction Pipeline (`sensitivity_gui.py`)

The GUI walks through a four-step workflow:

1. **Select a folder** of ACFR PDFs
2. **Load a master plan list** (optional) for fuzzy name matching
3. **Run extraction** — the pipeline processes each PDF sequentially:
   - Identifies pages with sensitivity-related keywords
   - Trims the PDF to only relevant pages
   - Sends trimmed pages to the Gemini API with a structured prompt
   - Caches results to JSON after each file (supports resume on failure)
4. **Write results to Excel** — merges extracted data with matched plan names

Key design decisions:
- **JSON caching with resume**: API calls are expensive and rate-limited. The pipeline caches after every PDF so a failure at PDF #150 doesn't lose the first 149 results.
- **Consecutive error detection**: Stops automatically after 3 consecutive API errors to avoid burning through quota on a systemic issue.
- **Rate limiting**: Built-in 4-second delay between API calls to stay within Gemini's rate limits.

### QA Review Tool (`sensitivity_checking.py`)

A separate GUI for validating extracted data against the source documents:

- Displays extracted discount rates and NPL values in editable fields alongside the rendered PDF page
- Handles the printed-vs-physical page number mismatch common in ACFRs (e.g., the document says "page 87" but it's actually the 103rd page of the PDF) by scanning headers/footers for printed page numbers
- Highlights sensitivity-related terms on rendered pages for quick visual confirmation
- Supports inline editing and saves corrections back to the JSON cache
- Zoom controls and manual page navigation for hard-to-find tables

## Tech Stack

- **Python 3.10+**
- **Google Gemini API** — structured data extraction from PDF pages
- **PyQt5** — desktop GUI for both extraction and QA workflows
- **PyMuPDF (fitz)** — PDF page rendering, text extraction, and keyword highlighting
- **openpyxl** — Excel output generation
- **fuzzywuzzy** — fuzzy string matching for plan name reconciliation

## Setup
```bash
pip install -r requirements.txt
```


Automated extraction of discount-rate sensitivity data from state and local government Annual Comprehensive Financial Reports (ACFRs) using AI-powered document processing.

## Overview

Public pension plans are required to disclose the sensitivity of their net pension liability (NPL) to changes in the discount rate. This information is buried across hundreds of ACFRs — lengthy PDF financial reports that vary widely in format and structure. Manually locating and transcribing this data from 200+ reports (50,000+ pages) took weeks.

This pipeline compresses that process into a single day by:

1. Scanning each ACFR PDF to identify pages containing sensitivity analysis tables

2. Extracting structured data (plan names, discount rates, NPL at ±1%) using Google's Gemini API

3. Matching extracted plan names against a master list of known pension plans

4. Validating results with automated checks and a side-by-side QA review tool

5. Outputting clean, analysis-ready data in Excel

## Repo Structure

```
acfr-sensitivity-extraction/
│
├── sensitivity_extractor.py       # Core extraction logic: PDF page detection,
│                                  # Gemini API integration, plan matching,
│                                  # validation, and Excel output
│
├── sensitivity_gui.py             # PyQt5 GUI for running the full extraction
│                                  # pipeline with progress tracking and caching
│
├── sensitivity_checking.py        # QA review tool: displays extracted data
│                                  # alongside source PDF pages for verification
│
├── start_sensitivity_extraction_v2.bat  # Windows launcher (extraction GUI)
├── sensitivity_checking.bat             # Windows launcher (QA tool)
│
├── requirements.txt
└── README.md
```
