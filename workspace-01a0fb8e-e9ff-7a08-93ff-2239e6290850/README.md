# Offline Intelligent PDF-to-Excel Agent 🤖📊

An autonomous, 100% offline agent system designed to extract complex structured tables, metadata, and key-value fields from PDF documents (invoices, purchase orders, statements, bills) and generate formatted Microsoft Excel (`.xlsx`) workbooks.

The agent continuously **learns from past extractions, chat instructions, and historical feedback** using an offline memory store to guarantee **exact mathematical precision and schema consistency**.

---

## 🌟 Key Capabilities

1. **100% Offline Execution**: Runs completely on-device without cloud API keys, external network requests, or data privacy risks.
2. **History & Chat Learning Loop**:
   - Remembers previous layout fingerprints, column mappings, and user instructions.
   - Learns field synonyms and user corrections in real-time (e.g. *"Rename 'Particulars' to 'Item Description'"*).
   - Automatically matches recurring templates (e.g. Vendor Invoices, Bank Statements).
3. **Exact Result Engine**:
   - Multi-layer mathematical checksums: asserts `Quantity × Unit Price == Line Total`.
   - Cross-reconciles `Subtotal + Tax - Discount == Grand Total`.
   - Auto-detects and self-heals OCR discrepancies, missing totals, or shifted columns.
4. **Multi-Engine PDF Extraction**:
   - Lattice & Stream vector table boundary detection (`pdfplumber`).
   - Positional text layout and reading order extraction (`pymupdf`).
   - Fallback high-DPI OCR (`pytesseract` + `poppler-utils`) for scanned paper receipts.
   - Multi-page continuous table stitching.
5. **Formatted Excel Exporter (`openpyxl`)**:
   - Professional styling with visual hierarchy (Classic Blue / Slate Charcoal themes).
   - Native Excel formulas for totals (`=SUM(...)`).
   - Exact cell types: Currency (`$#,##0.00` or `₹#,##0.00`), Integers (`#,##0`), Decimals, ISO Dates (`YYYY-MM-DD`).
   - Multi-tab workbooks: `Extracted Data`, `Document Summary`, and `Agent Audit Log`.
   - Auto-adjusted column widths with padding.
   - Bank-statement exports contain only a `Transactions` sheet with `Txn Date`, `Value Date`, `Particulars`, `Ref. No.`, `Debit`, `Credit`, and `Balance`; INR amounts are formatted using `₹`.

---

## 🏗️ Architecture Overview

```
/home/user/
  ├── agent/
  │   ├── core_agent.py          # Central agent orchestrator & execution pipeline
  │   ├── parser.py              # Multi-engine PDF extractor (pymupdf, pdfplumber, OCR)
  │   ├── memory.py              # SQLite historical memory, template store & synonym engine
  │   ├── schema_matcher.py      # Fuzzy & historical schema alignment
  │   ├── validator.py           # Mathematical verification & accuracy scoring
  │   └── excel_exporter.py      # Professional Excel generator with openpyxl
  ├── web/
  │   ├── app.py                 # Interactive Flask web application (Port 5000)
  │   ├── templates/index.html   # Web UI dashboard & spreadsheet viewer
  │   └── static/                # CSS styling & JS frontend controller
  ├── samples/                   # Sample test PDFs (Invoices, Multi-page POs, Statements)
  ├── outputs/                   # Generated .xlsx Excel files
  ├── history/                   # Persistent SQLite DB (agent_memory.db)
  ├── cli.py                     # Rich interactive command-line interface
  └── test_agent.py              # Automated test suite
```

---

## 🚀 Quick Start Guide

### 1. Interactive Terminal Shell
Run the agent in interactive shell mode:
```bash
python3 /home/user/cli.py interactive
```

### 2. Extract a Single PDF Document
```bash
python3 /home/user/cli.py process /home/user/samples/invoice_standard.pdf -o /home/user/outputs/my_invoice.xlsx
```

### 3. Guide Extraction with Natural Language Chat
```bash
python3 /home/user/cli.py process /home/user/samples/noisy_vendor_bill.pdf --chat "Rename Particulars to Description and Qnty to Quantity"
```

### 4. Teach the Agent New Rules (Offline Learning)
```bash
python3 /home/user/cli.py learn "Rename 'Rate/Unit' to 'Unit Price'"
python3 /home/user/cli.py learn "Set date format to YYYY-MM-DD"
```

### 5. Inspect Historical Runs & Accuracy
```bash
python3 /home/user/cli.py history
```

### 6. View Learned Document Templates
```bash
python3 /home/user/cli.py templates
```

### 7. Batch Process an Entire Directory
```bash
python3 /home/user/cli.py batch /home/user/samples/ -o /home/user/outputs/
```

---

## 💻 Python SDK Usage

You can embed the agent in your own scripts:

```python
from agent import PDFToExcelAgent

# Initialize the offline agent
agent = PDFToExcelAgent()

# 1. Process a PDF and get structured output
result = agent.process_pdf(
    pdf_path="samples/invoice_standard.pdf",
    output_excel_path="outputs/invoice_result.xlsx",
    chat_instruction="Ensure Line Total is calculated from Quantity * Unit Price"
)

print(f"Excel generated at: {result['excel_path']}")
print(f"Accuracy Score: {result['accuracy_score'] * 100:.1f}%")
print(f"Extracted Headers: {result['headers']}")
print(f"Extracted Rows: {result['row_count']}")

# 2. Teach the agent rules for future documents
agent.learn_from_chat("Rename 'Particulars' to 'Description'")
```

---

## 🌐 Web Dashboard & Live Preview

The built-in web dashboard runs on `http://0.0.0.0:5000` and offers:
- **Drag-and-Drop Ingestion**: Upload any PDF or pick from one-click test samples.
- **Interactive Spreadsheet Editor**: Edit cells directly in the browser; modifications update agent memory and re-export to Excel.
- **Agent Chat Assistant**: Direct conversational interface to instruct the agent on layout rules, column renames, and date formats.
- **Audit & Math Verification Tab**: Visual breakdown of checked arithmetic equations, grand totals, and confidence metrics.
- **One-Click Download**: Instantly download the formatted `.xlsx` workbook.

---

## 🧪 Running Automated Tests

Run the complete test suite:
```bash
python3 /home/user/test_agent.py
```
