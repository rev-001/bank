"""
Schema Normalization & Column Alignment Engine.
Maps noisy extracted headers to clean, standardized target schemas
using history-based templates, fuzzy matching, and data-type inference.
"""

import difflib
import re
from typing import Dict, List, Any, Optional, Tuple
from datetime import datetime


class SchemaMatcher:
    def __init__(self, memory_engine):
        self.memory = memory_engine

    def standardize_table(self, table: Dict[str, Any], template: Optional[Dict[str, Any]] = None,
                          chat_rules: Optional[List[Dict[str, Any]]] = None) -> Dict[str, Any]:
        """
        Takes raw extracted table with headers & rows, and aligns headers
        using historical templates, synonyms, fuzzy matching, and chat rules.
        """
        raw_headers = table.get("headers", [])
        raw_rows = table.get("rows", [])

        if not raw_headers or not raw_rows:
            return table

        # Step 1: Base mappings from synonyms database
        synonym_dict = self.memory.get_all_synonyms()
        
        # Step 2: Override with Template mappings if available
        template_mappings = {}
        if template and "column_mappings" in template:
            template_mappings = {k.lower().strip(): v for k, v in template["column_mappings"].items()}

        # Step 3: Check chat rules for column renames
        rule_renames = {}
        if chat_rules:
            for rule in chat_rules:
                if rule.get("action_type") == "rename_col":
                    payload = rule.get("rule_payload", {})
                    src = payload.get("source", "").lower().strip()
                    tgt = payload.get("target", "").strip()
                    if src and tgt:
                        rule_renames[src] = tgt

        mapped_headers = []
        column_types = []

        for idx, h in enumerate(raw_headers):
            h_clean = h.strip()
            h_lower = h_clean.lower()
            
            # Check 1: Chat rule exact match
            if h_lower in rule_renames:
                standard_name = rule_renames[h_lower]
            # Check 2: Template mapping exact match
            elif h_lower in template_mappings:
                standard_name = template_mappings[h_lower]
            # Check 3: Synonyms dictionary match
            elif h_lower in synonym_dict:
                standard_name = synonym_dict[h_lower]
            else:
                # Check 4: Fuzzy match against known synonyms
                fuzzy_match = self._fuzzy_match_header(h_lower, synonym_dict)
                if fuzzy_match:
                    standard_name = fuzzy_match
                else:
                    # Clean up the raw header string into Title Case
                    standard_name = self._beautify_header(h_clean)

            mapped_headers.append(standard_name)

        # Step 4: Infer and sanitize column data types across all rows
        sanitized_rows, col_types = self._sanitize_and_type_rows(raw_rows, mapped_headers)

        return {
            "page": table.get("page", 1),
            "original_headers": raw_headers,
            "headers": mapped_headers,
            "rows": sanitized_rows,
            "column_types": col_types
        }

    def _fuzzy_match_header(self, raw_header: str, synonym_dict: Dict[str, str]) -> Optional[str]:
        """Fuzzy match raw header against all known aliases."""
        clean_raw = re.sub(r'[^a-z0-9\s]', '', raw_header)
        if not clean_raw:
            return None

        # Check for substring matches
        for alias, std_name in synonym_dict.items():
            if alias in clean_raw or clean_raw in alias:
                if len(clean_raw) >= 3 and len(alias) >= 3:
                    return std_name

        # Difflib close matches
        matches = difflib.get_close_matches(clean_raw, list(synonym_dict.keys()), n=1, cutoff=0.78)
        if matches:
            return synonym_dict[matches[0]]

        return None

    def _beautify_header(self, text: str) -> str:
        """Turn 'item_description_1' or 'QTY (PCS)' into clean title 'Item Description 1'."""
        text = re.sub(r'[_\-]+', ' ', text)
        words = text.strip().split()
        return ' '.join(w.capitalize() if not w.isupper() else w for w in words)

    def _sanitize_and_type_rows(self, rows: List[List[str]], headers: List[str]) -> Tuple[List[List[Any]], List[str]]:
        """
        Converts text cells to clean typed values (floats, ints, dates, clean strings),
        and infers overall column types (currency, integer, date, text).
        """
        if not rows:
            return rows, ["text"] * len(headers)

        num_cols = len(headers)
        column_samples: List[List[Any]] = [[] for _ in range(num_cols)]
        sanitized_rows = []

        for row in rows:
            # Pad row if too short
            padded_row = list(row) + [""] * (num_cols - len(row))
            clean_row = []
            
            for col_idx in range(num_cols):
                raw_val = padded_row[col_idx]
                parsed_val = self._parse_cell_value(raw_val, headers[col_idx])
                clean_row.append(parsed_val)
                if parsed_val is not None and str(parsed_val).strip() != "":
                    column_samples[col_idx].append(parsed_val)

            sanitized_rows.append(clean_row)

        # Infer column types
        col_types = []
        for col_idx in range(num_cols):
            samples = column_samples[col_idx]
            col_name = headers[col_idx].lower()
            
            if not samples:
                col_types.append("text")
                continue

            float_count = sum(1 for s in samples if isinstance(s, (int, float)))
            date_count = sum(1 for s in samples if isinstance(s, datetime) or (isinstance(s, str) and re.match(r'^\d{4}-\d{2}-\d{2}', s)))
            
            if "price" in col_name or "amount" in col_name or "total" in col_name or "rate" in col_name or "cost" in col_name or "tax" in col_name or "discount" in col_name or "debit" in col_name or "credit" in col_name or "balance" in col_name or "withdrawal" in col_name or "deposit" in col_name:
                col_types.append("currency")
            elif "qty" in col_name or "quantity" in col_name or "hours" in col_name or "units" in col_name:
                col_types.append("integer" if all(isinstance(s, int) or (isinstance(s, float) and s.is_integer()) for s in samples if isinstance(s, (int, float))) else "decimal")
            elif date_count >= len(samples) * 0.7 or "date" in col_name:
                col_types.append("date")
            elif float_count >= len(samples) * 0.8:
                col_types.append("number")
            else:
                col_types.append("text")

        return sanitized_rows, col_types

    def _parse_cell_value(self, cell_str: Any, col_header: str) -> Any:
        """Parse raw cell string into typed Python float/int/str/date."""
        if cell_str is None:
            return ""
        
        s = str(cell_str).strip()
        if not s or s == "-":
            return ""

        # Normalize currency / number formats (e.g. '$ 1,250.50' or '1.250,50 €')
        # Check parenthesized negative: (125.00) -> -125.00
        is_negative = False
        if re.match(r'^\s*\(\s*[\$€£₹]?\s*[\d,.]+\s*\)\s*$', s):
            is_negative = True
            s = re.sub(r'[\(\)]', '', s).strip()

        # Check if it matches number pattern
        num_clean = re.sub(r'[\$€£₹,\s]', '', s)
        
        # Check European comma decimal: e.g. '1250,50' -> '1250.50'
        if re.match(r'^-?\d+,\d{2}$', num_clean):
            num_clean = num_clean.replace(',', '.')

        # Try parsing integer
        if re.match(r'^-?\d+$', num_clean):
            try:
                val = int(num_clean)
                return -val if is_negative else val
            except ValueError:
                pass

        # Try parsing float
        if re.match(r'^-?\d+\.\d+$', num_clean):
            try:
                val = float(num_clean)
                return -val if is_negative else val
            except ValueError:
                pass

        # Check Date formats
        # YYYY-MM-DD or DD/MM/YYYY or MM/DD/YYYY
        date_patterns = [
            (r'^(\d{4})[-/.](\d{1,2})[-/.](\d{1,2})$', '%Y-%m-%d'),
            (r'^(\d{1,2})[-/.](\d{1,2})[-/.](\d{4})$', '%d/%m/%Y'),
            (r'^(\d{1,2})\s+([A-Za-z]{3,9})\s+(\d{4})$', '%d %b %Y')
        ]
        for pat, fmt in date_patterns:
            if re.match(pat, s):
                try:
                    # Return normalized ISO string YYYY-MM-DD
                    dt = datetime.strptime(s, fmt)
                    return dt.strftime('%Y-%m-%d')
                except Exception:
                    pass

        return s
