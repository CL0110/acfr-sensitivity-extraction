# ACFR Sensitivity Analysis Extraction Pipeline

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
