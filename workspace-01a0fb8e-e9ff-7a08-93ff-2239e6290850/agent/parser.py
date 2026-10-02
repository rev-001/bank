"""
Multi-Engine PDF Extractor for Offline PDF-to-Excel Agent.
Integrates PyMuPDF (fitz), PDFPlumber, and Tesseract OCR with intelligent
table boundary detection, multi-page stitching, and key-value extraction.
"""

import re
import os
from pathlib import Path
from typing import Dict, List, Any, Optional, Tuple
import pymupdf  # PyMuPDF
import pdfplumber
import pytesseract
from PIL import Image


class PDFParser:
    def __init__(self, tesseract_cmd: Optional[str] = None):
        if tesseract_cmd:
            pytesseract.pytesseract.tesseract_cmd = tesseract_cmd

    def parse(self, pdf_path: str) -> Dict[str, Any]:
        """
        Main parsing pipeline.
        Returns a rich structure containing:
        - raw_text: full document string
        - pages_text: list of strings per page
        - tables: list of extracted tables with headers and clean rows
        - key_values: dictionary of detected metadata / header fields
        - is_scanned: boolean flag indicating if OCR was required
        - page_count: total pages
        """
        pdf_path = str(Path(pdf_path).resolve())
        if not os.path.exists(pdf_path):
            raise FileNotFoundError(f"PDF file not found: {pdf_path}")

        # Step 1: Digital text extraction with PyMuPDF
        pages_text, full_text, is_scanned = self._extract_text_pymupdf(pdf_path)

        # Step 2: Fallback to OCR if scanned or minimal text
        if is_scanned or len(full_text.strip()) < 40:
            pages_text, full_text = self._extract_ocr(pdf_path)
            is_scanned = True

        # Step 3: Extract structured tables with PDFPlumber + PyMuPDF
        tables = self._extract_tables_pdfplumber(pdf_path)
        pymupdf_tables = self._extract_tables_pymupdf(pdf_path)
        tables = self._select_richest_tables(tables, pymupdf_tables)

        # Step 4: Extract Key-Value Metadata Pairs
        key_values = self._extract_key_values(full_text)
        statement_table = self._extract_positioned_bank_statement(pdf_path)
        if not statement_table:
            statement_table = self._normalize_bank_statement_table(tables, full_text)
        if statement_table:
            tables = [statement_table]
            key_values = self._extract_bank_statement_key_values(full_text)

        return {
            "file_path": pdf_path,
            "filename": os.path.basename(pdf_path),
            "page_count": len(pages_text),
            "is_scanned": is_scanned,
            "full_text": full_text,
            "pages_text": pages_text,
            "tables": tables,
            "key_values": key_values,
            "is_bank_statement": statement_table is not None
        }

    def _extract_positioned_bank_statement(self, pdf_path: str) -> Optional[Dict[str, Any]]:
        """Extract HDFC-style statements using the transaction columns' page positions."""
        date_pattern = re.compile(r"\d{2}/\d{2}/\d{2}")
        rows: List[List[str]] = []
        headers = [
            "Txn Date",
            "Value Date",
            "Particulars",
            "Ref. No.",
            "Debit",
            "Credit",
            "Balance",
        ]

        with pymupdf.open(pdf_path) as doc:
            document_text = "\n".join(page.get_text("text") for page in doc).lower()
            if not all(term in document_text for term in (
                "statement of account",
                "narration",
                "chq./ref.no.",
                "withdrawal",
                "deposit",
                "closing balance",
            )):
                return None

            for page in doc:
                words = page.get_text("words")
                starts = sorted(
                    (float(y0), str(word))
                    for x0, y0, _x1, _y1, word, *_rest in words
                    if x0 < 70 and y0 >= 220 and date_pattern.fullmatch(str(word))
                )

                if not starts:
                    if rows:
                        for x0, y0, _x1, _y1, word, *_rest in words:
                            if not 220 <= y0 < 270:
                                continue
                            if 65 <= x0 < 270:
                                rows[-1][2] = (rows[-1][2] + " " + str(word)).strip()
                            elif 270 <= x0 < 360:
                                rows[-1][3] += str(word)
                    continue

                first_start = starts[0][0]
                if rows:
                    for x0, y0, _x1, _y1, word, *_rest in words:
                        if not 220 <= y0 < first_start:
                            continue
                        if 65 <= x0 < 270:
                            rows[-1][2] = (rows[-1][2] + " " + str(word)).strip()
                        elif 270 <= x0 < 360:
                            rows[-1][3] += str(word)

                for index, (start_y, transaction_date) in enumerate(starts):
                    end_y = starts[index + 1][0] if index + 1 < len(starts) else 780
                    columns: List[List[Tuple[float, float, str]]] = [[] for _ in headers]

                    for x0, y0, _x1, _y1, word, *_rest in words:
                        if not start_y <= y0 < end_y:
                            continue
                        if x0 < 70:
                            column = 0
                        elif x0 < 270:
                            column = 1
                        elif x0 < 360:
                            column = 2
                        elif x0 < 405:
                            column = 3
                        elif x0 < 490:
                            column = 4
                        elif x0 < 565:
                            column = 5
                        else:
                            column = 6
                        columns[column].append((float(y0), float(x0), str(word)))

                    for column in columns:
                        column.sort(key=lambda item: (item[0], item[1]))

                    def text_at(column_index: int) -> str:
                        return " ".join(word for _y, _x, word in columns[column_index]).strip()

                    def compact_at(column_index: int) -> str:
                        return "".join(word for _y, _x, word in columns[column_index]).strip()

                    rows.append([
                        transaction_date,
                        compact_at(3),
                        text_at(1),
                        compact_at(2),
                        compact_at(4),
                        compact_at(5),
                        compact_at(6),
                    ])

        if not rows:
            raise ValueError("Bank statement layout was detected, but no transactions could be extracted.")

        return {"page": 1, "headers": headers, "rows": rows}

    def _normalize_bank_statement_table(
        self, tables: List[Dict[str, Any]], text: str
    ) -> Optional[Dict[str, Any]]:
        """Map a detected tabular bank statement to the standard transaction columns."""
        text_lower = text.lower()
        if "statement" not in text_lower or not any(term in text_lower for term in ("account", "bank")):
            return None

        aliases = {
            "Txn Date": {"date", "txndate", "transactiondate", "transdate", "postingdate"},
            "Value Date": {"valuedate", "valuedt"},
            "Particulars": {"particulars", "description", "narration", "details"},
            "Ref. No.": {"ref", "refno", "referenceno", "chqrefno", "chequeno"},
            "Debit": {"debit", "debitamount", "withdrawal", "withdrawals"},
            "Credit": {"credit", "creditamount", "deposit", "deposits"},
            "Balance": {"balance", "closingbalance", "closingbal"},
        }
        headers = list(aliases)

        for table in tables:
            source_headers = table.get("headers") or []
            normalized_headers = {
                re.sub(r"[^a-z0-9]", "", str(header).lower()): index
                for index, header in enumerate(source_headers)
            }
            indexes = {}
            for output_header, names in aliases.items():
                indexes[output_header] = next(
                    (normalized_headers[name] for name in names if name in normalized_headers),
                    None,
                )

            required = ("Txn Date", "Particulars", "Debit", "Credit", "Balance")
            if any(indexes[name] is None for name in required):
                continue

            rows = []
            for source_row in table.get("rows") or []:
                def value_for(header: str) -> Any:
                    index = indexes[header]
                    if index is None or index >= len(source_row):
                        return ""
                    return source_row[index]

                transaction_date = value_for("Txn Date")
                rows.append([
                    transaction_date,
                    value_for("Value Date"),
                    value_for("Particulars"),
                    value_for("Ref. No."),
                    value_for("Debit"),
                    value_for("Credit"),
                    value_for("Balance"),
                ])

            if rows:
                return {
                    "page": table.get("page", 1),
                    "headers": headers,
                    "rows": rows,
                }

        return None

    def _extract_bank_statement_key_values(self, text: str) -> Dict[str, str]:
        """Extract account and statement-summary fields from a bank statement."""
        key_values: Dict[str, str] = {}
        patterns = {
            "Account Number": r"Account\s+No\s*:?\s*(\d+)",
            "Statement From": r"From\s*:\s*(\d{2}/\d{2}/\d{4})",
            "Statement To": r"To\s*:\s*(\d{2}/\d{2}/\d{4})",
            "Opening Balance": r"Opening\s+Balance\s+Dr\s+Count\s+Cr\s+Count\s+Debits\s+Credits\s+Closing\s+Bal\s+([\d,]+\.\d{2})",
            "Debit Transactions": r"Opening\s+Balance\s+Dr\s+Count\s+Cr\s+Count\s+Debits\s+Credits\s+Closing\s+Bal\s+[\d,]+\.\d{2}\s+(\d+)",
            "Credit Transactions": r"Opening\s+Balance\s+Dr\s+Count\s+Cr\s+Count\s+Debits\s+Credits\s+Closing\s+Bal\s+[\d,]+\.\d{2}\s+\d+\s+(\d+)",
            "Total Debits": r"Opening\s+Balance\s+Dr\s+Count\s+Cr\s+Count\s+Debits\s+Credits\s+Closing\s+Bal\s+[\d,]+\.\d{2}\s+\d+\s+\d+\s+([\d,]+\.\d{2})",
            "Total Credits": r"Opening\s+Balance\s+Dr\s+Count\s+Cr\s+Count\s+Debits\s+Credits\s+Closing\s+Bal\s+[\d,]+\.\d{2}\s+\d+\s+\d+\s+[\d,]+\.\d{2}\s+([\d,]+\.\d{2})",
            "Closing Balance": r"Opening\s+Balance\s+Dr\s+Count\s+Cr\s+Count\s+Debits\s+Credits\s+Closing\s+Bal\s+[\d,]+\.\d{2}\s+\d+\s+\d+\s+[\d,]+\.\d{2}\s+[\d,]+\.\d{2}\s+([\d,]+\.\d{2})",
        }
        for label, pattern in patterns.items():
            match = re.search(pattern, text, re.IGNORECASE)
            if match:
                key_values[label] = match.group(1)

        if "Account Number" in key_values:
            key_values["Bank"] = "HDFC Bank"
            key_values["Currency"] = "INR"

        return key_values

    def _extract_text_pymupdf(self, pdf_path: str) -> Tuple[List[str], str, bool]:
        """Extract text and determine if document is digital or scanned."""
        pages_text = []
        doc = pymupdf.open(pdf_path)
        total_chars = 0

        for page in doc:
            text = page.get_text("text")
            pages_text.append(text)
            total_chars += len(text.strip())

        doc.close()
        full_text = "\n\n".join(pages_text)
        is_scanned = (total_chars < 50)  # If very few chars, likely scanned image
        return pages_text, full_text, is_scanned

    def _extract_ocr(self, pdf_path: str) -> Tuple[List[str], str]:
        """Render PDF pages to images and execute OCR with Tesseract."""
        pages_text = []
        doc = pymupdf.open(pdf_path)

        for page_num in range(len(doc)):
            page = doc[page_num]
            # Render at 300 DPI for high-accuracy OCR
            pix = page.get_pixmap(dpi=300)
            img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
            ocr_text = pytesseract.image_to_string(img)
            pages_text.append(ocr_text)

        doc.close()
        full_text = "\n\n".join(pages_text)
        return pages_text, full_text

    def _extract_tables_pdfplumber(self, pdf_path: str) -> List[Dict[str, Any]]:
        """
        Extract tables across all pages using pdfplumber with multi-strategy detection,
        line merging, and cross-page stitching.
        """
        all_tables: List[Dict[str, Any]] = []

        try:
            with pdfplumber.open(pdf_path) as pdf:
                for page_idx, page in enumerate(pdf.pages):
                    page_tables = []
                    
                    # Strategy 1: Explicit lines (Lattice)
                    extracted = page.extract_tables(table_settings={
                        "vertical_strategy": "lines",
                        "horizontal_strategy": "lines",
                        "snap_tolerance": 4,
                        "join_tolerance": 4,
                    })

                    # Strategy 2: Text alignment (Stream) if lattice found nothing
                    if not extracted:
                        extracted = page.extract_tables(table_settings={
                            "vertical_strategy": "text",
                            "horizontal_strategy": "text",
                            "snap_tolerance": 3,
                            "min_words_vertical": 2
                        })

                    # Strategy 3: Default auto extract
                    if not extracted:
                        extracted = page.extract_tables()

                    for tbl in extracted:
                        cleaned = self._clean_table_matrix(tbl)
                        if cleaned and len(cleaned["rows"]) > 0:
                            cleaned["page"] = page_idx + 1
                            page_tables.append(cleaned)

                    all_tables.extend(page_tables)

        except Exception as e:
            # Fallback text table parser if pdfplumber encounters errors
            all_tables = []

        # If pdfplumber didn't find clear tables, try text-based tabular parsing
        if not all_tables:
            all_tables = self._fallback_text_table_parser(pdf_path)

        # Merge continuous multi-page tables
        merged_tables = self._merge_multipage_tables(all_tables)
        return merged_tables

    def _clean_table_matrix(self, raw_table: List[List[Optional[str]]]) -> Optional[Dict[str, Any]]:
        """Clean messy table rows, extract headers, merge wrapped multi-line cells."""
        if not raw_table or len(raw_table) < 2:
            return None

        # Filter out completely empty rows
        filtered_rows = []
        for row in raw_table:
            cleaned_row = [re.sub(r'\s+', ' ', str(cell).strip()) if cell is not None else "" for cell in row]
            if any(cell != "" for cell in cleaned_row):
                filtered_rows.append(cleaned_row)

        if len(filtered_rows) < 2:
            return None

        # Determine header row (usually row 0, but sometimes row 1 if title row precedes it)
        header_idx = 0
        for i, row in enumerate(filtered_rows[:3]):
            non_empty = [c for c in row if c]
            # If row has mostly alpha keywords and multiple columns, it's likely the header
            if len(non_empty) >= 2:
                header_idx = i
                break

        headers = filtered_rows[header_idx]
        # Replace empty headers with Column_1, Column_2 etc.
        headers = [h if h else f"Column_{j+1}" for j, h in enumerate(headers)]

        data_rows = filtered_rows[header_idx + 1:]

        # Merge wrapped lines (e.g. multi-line descriptions where subsequent row has only 1 non-empty cell)
        consolidated_rows = []
        for row in data_rows:
            # Check if this row is an exact duplicate of headers (repeating page header)
            if [c.lower() for c in row] == [h.lower() for h in headers]:
                continue

            non_empty_indices = [idx for idx, c in enumerate(row) if c]
            if not non_empty_indices:
                continue

            # Merge wrapped description lines only (no new date/amount — those are new transactions)
            if len(non_empty_indices) == 1 and consolidated_rows:
                idx = non_empty_indices[0]
                cell_text = row[idx]
                looks_like_new_txn = bool(
                    re.search(r'\d{1,2}[/\-\.]\d{1,2}([/\-\.]\d{2,4})?', cell_text)
                    or re.search(r'[\$€£₹]\s*[\d,]+', cell_text)
                    or re.match(r'^-?\d[\d,]*\.?\d*$', cell_text.replace(" ", ""))
                )
                if not looks_like_new_txn:
                    prev_row = consolidated_rows[-1]
                    prev_row[idx] = (prev_row[idx] + " " + cell_text).strip()
                    continue

            if len(row) < len(headers):
                row.extend([""] * (len(headers) - len(row)))
            elif len(row) > len(headers):
                row = row[:len(headers)]
            consolidated_rows.append(row)

        return {
            "headers": headers,
            "rows": consolidated_rows
        }

    def _fallback_text_table_parser(self, pdf_path: str) -> List[Dict[str, Any]]:
        """Regex and whitespace-based columnar parsing when vector tables are absent."""
        doc = pymupdf.open(pdf_path)
        tables = []

        for page_idx, page in enumerate(doc):
            lines = page.get_text("text").splitlines()
            potential_rows = []
            
            for line in lines:
                line_str = line.strip()
                if not line_str:
                    continue
                # Split by 2 or more spaces or tabs
                parts = re.split(r'\s{2,}|\t', line_str)
                if len(parts) >= 2:
                    potential_rows.append(parts)

            if len(potential_rows) >= 3:
                # Normalize column count to most frequent count
                counts = [len(r) for r in potential_rows]
                most_common_col_count = max(set(counts), key=counts.count)
                
                padded = []
                for r in potential_rows:
                    row = list(r)
                    if len(row) < most_common_col_count:
                        row.extend([""] * (most_common_col_count - len(row)))
                    elif len(row) > most_common_col_count:
                        extra = " ".join(row[most_common_col_count - 1:])
                        row = row[:most_common_col_count - 1] + [extra]
                    padded.append(row)
                if len(padded) >= 2:
                    headers = padded[0]
                    headers = [h if h else f"Col_{j+1}" for j, h in enumerate(headers)]
                    tables.append({
                        "page": page_idx + 1,
                        "headers": headers,
                        "rows": padded[1:]
                    })
        doc.close()
        return tables

    def _merge_multipage_tables(self, tables: List[Dict[str, Any]]) -> List[Dict[str, Any]]:
        """If multiple pages have identical or very similar headers, stitch them into one continuous table."""
        if len(tables) <= 1:
            return self._dedupe_table_rows(tables)

        merged: List[Dict[str, Any]] = []
        current = tables[0]

        for next_tbl in tables[1:]:
            curr_headers = [h.lower().strip() for h in current["headers"]]
            next_headers = [h.lower().strip() for h in next_tbl["headers"]]

            # If headers are identical or matching column count with 80%+ similarity
            if curr_headers == next_headers or (len(curr_headers) == len(next_headers) and len(set(curr_headers).intersection(set(next_headers))) >= len(curr_headers) * 0.7):
                current["rows"].extend(next_tbl["rows"])
            else:
                merged.append(current)
                current = next_tbl

        merged.append(current)
        return self._dedupe_table_rows(merged)

    def _extract_tables_pymupdf(self, pdf_path: str) -> List[Dict[str, Any]]:
        """Extract tables with PyMuPDF find_tables as a second engine."""
        tables: List[Dict[str, Any]] = []
        try:
            doc = pymupdf.open(pdf_path)
            for page_idx, page in enumerate(doc):
                finder = page.find_tables()
                if not finder or not finder.tables:
                    continue
                for tab in finder.tables:
                    matrix = tab.extract()
                    cleaned = self._clean_table_matrix(matrix)
                    if cleaned and cleaned.get("rows"):
                        cleaned["page"] = page_idx + 1
                        tables.append(cleaned)
            doc.close()
        except Exception:
            return []
        return tables

    def _select_richest_tables(
        self, plumber_tables: List[Dict[str, Any]], pymupdf_tables: List[Dict[str, Any]]
    ) -> List[Dict[str, Any]]:
        """Keep the engine that captured more transaction rows, then stitch pages."""
        def row_count(tables: List[Dict[str, Any]]) -> int:
            return sum(len(t.get("rows") or []) for t in tables)

        if not plumber_tables:
            chosen = pymupdf_tables
        elif not pymupdf_tables:
            chosen = plumber_tables
        elif row_count(pymupdf_tables) > row_count(plumber_tables):
            chosen = pymupdf_tables
        else:
            chosen = plumber_tables

        if not chosen:
            return []
        return self._merge_multipage_tables(chosen)

    def _dedupe_table_rows(self, tables: List[Dict[str, Any]]) -> List[Dict[str, Any]]:
        """Drop exact duplicate transaction rows (repeating page headers already stripped)."""
        cleaned = []
        for tbl in tables:
            seen = set()
            unique_rows = []
            for row in tbl.get("rows") or []:
                key = tuple(str(c).strip().lower() for c in row)
                if key in seen:
                    continue
                seen.add(key)
                unique_rows.append(row)
            tbl = dict(tbl)
            tbl["rows"] = unique_rows
            if unique_rows:
                cleaned.append(tbl)
        return cleaned

    def _extract_key_values(self, text: str) -> Dict[str, str]:
        """Extract key invoice/document header metadata using comprehensive regex patterns."""
        kv = {}

        patterns = {
            "Invoice Number": [
                r"(?:invoice\s*(?:no|number|#|num|code)|inv\s*#?)[:\s]+([A-Z0-9\-\_/]+)",
                r"(?:bill\s*no|bill\s*#)[:\s]+([A-Z0-9\-\_/]+)"
            ],
            "Invoice Date": [
                r"(?:invoice\s*date|bill\s*date|issue\s*date|date)[:\s]+([0-9]{1,4}[/\-\.][0-9]{1,2}[/\-\.][0-9]{1,4})",
                r"(?:invoice\s*date|bill\s*date|date)[:\s]+([A-Za-z]{3,9}\s+[0-9]{1,2},?\s+[0-9]{4})",
                r"(?:invoice\s*date|bill\s*date|date)[:\s]+([0-9]{1,2}\s+[A-Za-z]{3,9}\s+[0-9]{4})"
            ],
            "Due Date": [
                r"(?:due\s*date|payment\s*due)[:\s]+([0-9]{1,4}[/\-\.][0-9]{1,2}[/\-\.][0-9]{1,4})",
                r"(?:due\s*date|payment\s*due)[:\s]+([A-Za-z]{3,9}\s+[0-9]{1,2},?\s+[0-9]{4})"
            ],
            "Purchase Order": [
                r"(?:po\s*#|p\.o\.\s*#|po\s*no|purchase\s*order(?:\s*no)?)[:\s]+([A-Z0-9\-\_/]+)"
            ],
            "Customer / Bill To": [
                r"(?:bill\s*to|billed\s*to|customer\s*name|client)[:\s]+([^\n\r]+)",
                r"(?:to:\s*)([A-Za-z0-9\s,\.\&]+?)(?=\s*(?:invoice|date|ship|attn|\n{2}))"
            ],
            "Vendor / Company": [
                r"(?:from|vendor|supplier|company\s*name|issued\s*by)[:\s]+([^\n\r]+)"
            ],
            "Tax ID / GSTIN / VAT": [
                r"(?:gstin|gst\s*no|vat\s*no|tax\s*id|ein|abn)[:\s]+([A-Z0-9\-\_/]+)"
            ],
            "Subtotal": [
                r"(?:subtotal|sub\s*total|net\s*amount)[:\s]+([\$€£₹]?\s*[\d,]+\.\d{2})"
            ],
            "Tax Amount": [
                r"(?:tax(?:\s*amount)?|vat|gst|sales\s*tax|cgst\+sgst)(?:\s*\([0-9\.\s%]+\))?[:\s]+([\$€£₹]?\s*[\d,]+\.\d{2})",
                r"(?:tax|vat|gst|sales\s*tax)[:\s]+([\$€£₹]?\s*[\d,]+\.\d{2})"
            ],
            "Discount": [
                r"(?:discount|less:\s*discount|disc)[:\s]+([\$€£₹]?\s*[\d,]+\.\d{2})"
            ],
            "Grand Total": [
                r"(?:grand\s*total|total\s*due|amount\s*due|total\s*payable|total\s*amount|balance\s*due)[:\s]+([\$€£₹]?\s*[\d,]+\.\d{2})",
                r"(?:total)[:\s]+([\$€£₹]?\s*[\d,]+\.\d{2})"
            ],
            "Payment Terms": [
                r"(?:payment\s*terms|terms)[:\s]+([^\n\r]+)"
            ],
            "Currency": [
                r"(\bUSD\b|\bEUR\b|\bGBP\b|\bINR\b|\bCAD\b|\bAUD\b|\$|€|£|₹)"
            ]
        }

        for key, regex_list in patterns.items():
            for pat in regex_list:
                match = re.search(pat, text, re.IGNORECASE)
                if match:
                    val = match.group(1).strip()
                    # Clean up value
                    val = re.sub(r'[\r\n]+', ' ', val)
                    val = re.sub(r'\s{2,}', ' ', val)
                    if val and len(val) < 100:
                        kv[key] = val
                        break

        # If Currency not detected but symbols found in totals
        if "Currency" not in kv:
            if "$" in text:
                kv["Currency"] = "USD ($)"
            elif "€" in text:
                kv["Currency"] = "EUR (€)"
            elif "£" in text:
                kv["Currency"] = "GBP (£)"
            elif "₹" in text:
                kv["Currency"] = "INR (₹)"

        return kv
