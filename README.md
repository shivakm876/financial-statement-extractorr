# Financial Statement Extractor  
### Research Tool Implementation – Option A

---

## Overview

This project implements a minimal research portal slice as part of the Tech Intern L2 Assignment.

The system allows a researcher to:

1. Upload one or more financial statement PDFs
2. Automatically extract structured Income Statement data
3. Standardize extracted line items into a predefined financial schema
4. Generate an analyst-ready Excel file for further analysis

The focus of this implementation is reliability, structured output, and prevention of hallucinated financial data.

---

## 🌐 Live Deployment

Public URL:  
[Insert your Render deployment link here]

---

## ✅ Implemented Research Tool

### Option A – Financial Statement Extraction to Excel

### Input
- Annual reports
- Quarterly financial statements
- Income statement PDFs (text-based)

### Output
- Structured Excel file (.xlsx)
- Standardized 19-line-item income statement
- Multi-year data extraction
- Confidence-level indicators
- Calculated derived metrics (EBITDA, EBIT, Gross Profit, Margin)

---

## 🔍 Design Philosophy

This tool is designed as a deterministic financial extractor rather than an open-ended AI chatbot.

Key principles:

- No hallucinated financial numbers
- Structured schema output
- Analyst transparency
- Explicit handling of ambiguity
- Confidence tagging for traceability

---

## 📊 Schema (Standardized Financial Model)

The system normalizes extracted data into 19 standardized line items including:

- Revenue  
- Other Income  
- Cost of Materials  
- Purchases  
- Employee Expenses  
- Finance Costs  
- Depreciation  
- Total Expenses  
- Profit Before Tax (PBT)  
- Tax Expense  
- Profit After Tax (PAT)  
- EBITDA (calculated if missing)  
- EBIT (calculated if missing)  
- Gross Profit (calculated)  
- Net Profit Margin (%)  
- And additional core metrics  

---

## 🧠 Handling Key Judgment Calls

### 1. Different Line Item Names

Handles variations such as:
- "Sales" → Revenue
- "Revenue from Operations" → Revenue
- "PBT" → Profit Before Tax
- "PAT" → Profit After Tax
- "COGS" → Cost of Materials

Uses deterministic mapping logic with confidence scoring.

---

### 2. Missing Line Items

If a line item is not found:
- Marked as "—" in Excel
- Confidence set to Low
- Analyst Notes indicate "Not found in document"

No values are fabricated.

---

### 3. Numeric Extraction Reliability

- Cleans formatting noise
- Handles commas
- Handles negative values in parentheses
- Prevents hallucinated numeric generation
- Keeps only explicitly detected numbers

---

### 4. Currency & Units Detection

Automatically extracts:
- Currency (INR / USD)
- Units (Crores / Lakhs / Millions)

Displayed clearly in Excel subtitle.

---

### 5. Multi-Year Extraction

The system:
- Detects year columns dynamically
- Handles multiple reporting periods
- Extracts values across all detected years

---

### 6. Confidence Scoring

Each standardized line item is tagged as:
- High confidence (direct match)
- Medium confidence (pattern-mapped)
- Low confidence (not found)

Excel uses color coding for visual clarity.

---

## 🏗 Architecture

### Backend
- FastAPI
- Deterministic rule-based extraction
- Pattern-based label normalization

### PDF Processing
- pdfplumber for table extraction
- Structured filtering to isolate Income Statement

### Output Generation
- openpyxl for Excel generation
- Confidence-based formatting
- Analyst Notes column

---

## 🔌 API Endpoints

### POST `/extract`
Uploads PDF(s) and returns downloadable Excel file.

### POST `/preview`
Returns:
- Company name
- Currency
- Units
- Years detected
- Confidence summary
- Number of extracted rows

### GET `/`
Serves upload interface.

---

## ⚙ Reliability & Limitations

- Designed primarily for text-based PDFs.
- OCR-dependent scanned PDFs may require system-level dependencies.
- Free hosting may impose file size or runtime limits.
- Focused on correctness and structured output over performance optimization.

---

## 📈 Why Rule-Based Instead of LLM-Based?

This implementation intentionally avoids LLM-based extraction to:

- Eliminate hallucination risk
- Ensure numeric precision
- Maintain reproducibility
- Provide traceable mapping logic
- Improve analyst trust

---

## 🚀 Future Improvements

Potential enhancements:

- Raw extraction sheet for audit transparency
- Confidence scoring as percentage instead of categorical
- Optional hybrid LLM fallback mode
- Expanded financial schema coverage
- Support for scanned PDFs via OCR-enabled deployment

---

## 👤 Author

sivaganaga km
Tech Intern L2 Assignment Submission
