"""
Memory & History Engine for Offline PDF-to-Excel Agent.
Stores and learns from past extraction sessions, user chat instructions,
template fingerprints, and manual cell corrections.
"""

import json
import sqlite3
import hashlib
import re
from datetime import datetime
from pathlib import Path
from typing import Dict, List, Optional, Any, Tuple


_PROJECT_ROOT = Path(__file__).resolve().parent.parent


class AgentMemory:
    def __init__(self, db_path: str = str(_PROJECT_ROOT / "history" / "agent_memory.db")):
        self.db_path = Path(db_path)
        self.db_path.parent.mkdir(parents=True, exist_ok=True)
        self._init_db()

    def _get_connection(self) -> sqlite3.Connection:
        conn = sqlite3.connect(str(self.db_path))
        conn.row_factory = sqlite3.Row
        return conn

    def _init_db(self):
        conn = self._get_connection()
        try:
            cursor = conn.cursor()
            
            # Templates table: stores learned document structures
            cursor.execute("""
            CREATE TABLE IF NOT EXISTS templates (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                name TEXT NOT NULL UNIQUE,
                fingerprint TEXT NOT NULL,
                document_type TEXT DEFAULT 'generic',
                header_patterns TEXT NOT NULL, -- JSON list
                column_mappings TEXT NOT NULL, -- JSON dict {raw_header: standard_header}
                key_value_rules TEXT NOT NULL, -- JSON dict
                validation_rules TEXT NOT NULL, -- JSON dict
                sample_text_snippet TEXT,
                usage_count INTEGER DEFAULT 1,
                last_accuracy_score REAL DEFAULT 1.0,
                created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
                updated_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
            )
            """)

            # Chat Instructions / Learning Rules table
            cursor.execute("""
            CREATE TABLE IF NOT EXISTS chat_instructions (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                template_id INTEGER,
                user_prompt TEXT NOT NULL,
                action_type TEXT NOT NULL, -- 'rename_col', 'date_format', 'ignore_row', 'math_rule', 'custom_regex'
                rule_payload TEXT NOT NULL, -- JSON
                is_global INTEGER DEFAULT 0,
                created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
                FOREIGN KEY (template_id) REFERENCES templates(id) ON DELETE CASCADE
            )
            """)

            # Extraction History Runs table
            cursor.execute("""
            CREATE TABLE IF NOT EXISTS extraction_runs (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                pdf_filename TEXT NOT NULL,
                file_hash TEXT NOT NULL,
                template_name TEXT,
                table_row_count INTEGER DEFAULT 0,
                key_value_count INTEGER DEFAULT 0,
                accuracy_score REAL DEFAULT 1.0,
                excel_output_path TEXT,
                raw_extracted_json TEXT,
                final_data_json TEXT,
                user_corrections_json TEXT,
                notes TEXT,
                created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
            )
            """)

            # Global Synonyms / Dictionary table
            cursor.execute("""
            CREATE TABLE IF NOT EXISTS column_synonyms (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                standard_name TEXT NOT NULL,
                alias TEXT NOT NULL UNIQUE
            )
            """)

            conn.commit()
            self._seed_default_synonyms(cursor, conn)
        finally:
            conn.close()

    def _seed_default_synonyms(self, cursor: sqlite3.Cursor, conn: sqlite3.Connection):
        """Populate initial standard field aliases so the agent knows common synonyms immediately."""
        defaults = {
            "Description": ["desc", "description", "item description", "particulars", "item", "product", "details", "service", "item_desc", "designation"],
            "Quantity": ["qty", "quantity", "units", "count", "qnty", "vol", "volume", "hours", "hrs", "pcs", "pieces"],
            "Unit Price": ["unit price", "rate", "price", "unit cost", "cost/unit", "price/unit", "unit_rate", "rate/unit", "unit amount"],
            "Total Price": ["amount", "total", "total amount", "extended price", "total price", "net amount", "line total", "cost", "sum", "subtotal", "amt", "line amt"],
            "Item Code": ["item code", "sku", "part #", "part no", "product code", "code", "item #", "hsn", "hsn/sac", "article"],
            "Tax": ["tax", "tax amount", "vat", "gst", "cgst", "sgst", "igst", "sales tax", "tax %", "tax rate"],
            "Discount": ["discount", "disc", "disc %", "discount amount", "rebate", "less discount"],
            "Date": ["date", "invoice date", "bill date", "trans date", "issue date", "posting date", "txn date"],
            "Reference": ["ref", "reference", "ref no", "po number", "order no", "po #", "voucher no", "chq no"]
        }
        for std, aliases in defaults.items():
            for alias in aliases:
                cursor.execute("""
                INSERT OR IGNORE INTO column_synonyms (standard_name, alias)
                VALUES (?, ?)
                """, (std, alias.lower().strip()))
        conn.commit()

    @staticmethod
    def generate_fingerprint(text: str, headers: List[str] = None) -> str:
        """Create a resilient structural hash from text anchor words and column patterns."""
        normalized_text = re.sub(r'\s+', ' ', text.lower()[:2000])
        # Extract prominent keywords/tokens
        tokens = re.findall(r'[a-z]{4,}', normalized_text)
        token_sample = sorted(list(set(tokens[:40])))
        header_sample = sorted([h.lower().strip() for h in (headers or [])])
        
        fingerprint_basis = f"TOKENS:{'_'.join(token_sample)}|HEADERS:{'_'.join(header_sample)}"
        return hashlib.sha256(fingerprint_basis.encode('utf-8')).hexdigest()[:16]

    def find_matching_template(self, text: str, headers: List[str] = None) -> Optional[Dict[str, Any]]:
        """Look through historical templates to find the best matching layout."""
        conn = self._get_connection()
        try:
            cursor = conn.cursor()
            cursor.execute("SELECT * FROM templates ORDER BY usage_count DESC")
            templates = cursor.fetchall()
            
            if not templates:
                return None

            best_match = None
            highest_score = 0.0

            text_lower = text.lower()
            header_lower_set = set(h.lower().strip() for h in (headers or []))

            for row in templates:
                score = 0.0
                t_headers = set(json.loads(row["header_patterns"]))
                
                # Check header overlap
                if header_lower_set and t_headers:
                    overlap = len(header_lower_set.intersection(t_headers))
                    total = max(len(header_lower_set), len(t_headers))
                    header_similarity = overlap / total if total > 0 else 0
                    score += header_similarity * 0.6

                # Check snippet keywords
                snippet = (row["sample_text_snippet"] or "").lower()
                if snippet:
                    snippet_words = set(re.findall(r'[a-z]{4,}', snippet))
                    text_words = set(re.findall(r'[a-z]{4,}', text_lower[:2000]))
                    if snippet_words:
                        overlap_words = len(snippet_words.intersection(text_words))
                        word_sim = overlap_words / len(snippet_words)
                        score += word_sim * 0.4

                if score > highest_score and score >= 0.35:
                    highest_score = score
                    best_match = dict(row)
                    best_match["header_patterns"] = json.loads(row["header_patterns"])
                    best_match["column_mappings"] = json.loads(row["column_mappings"])
                    best_match["key_value_rules"] = json.loads(row["key_value_rules"])
                    best_match["validation_rules"] = json.loads(row["validation_rules"])
                    best_match["match_confidence"] = round(highest_score, 3)

            return best_match
        finally:
            conn.close()

    def save_template(self, name: str, fingerprint: str, header_patterns: List[str],
                      column_mappings: Dict[str, str], key_value_rules: Dict[str, Any] = None,
                      validation_rules: Dict[str, Any] = None, sample_text: str = "",
                      document_type: str = "generic") -> int:
        """Create or update a learned template based on user experience / chat."""
        key_value_rules = key_value_rules or {}
        validation_rules = validation_rules or {}
        
        conn = self._get_connection()
        try:
            cursor = conn.cursor()
            cursor.execute("""
            INSERT INTO templates (
                name, fingerprint, document_type, header_patterns, column_mappings,
                key_value_rules, validation_rules, sample_text_snippet, usage_count, updated_at
            ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, 1, CURRENT_TIMESTAMP)
            ON CONFLICT(name) DO UPDATE SET
                fingerprint = excluded.fingerprint,
                header_patterns = excluded.header_patterns,
                column_mappings = excluded.column_mappings,
                key_value_rules = excluded.key_value_rules,
                validation_rules = excluded.validation_rules,
                sample_text_snippet = excluded.sample_text_snippet,
                usage_count = templates.usage_count + 1,
                updated_at = CURRENT_TIMESTAMP
            """, (
                name,
                fingerprint,
                document_type,
                json.dumps([h.lower().strip() for h in header_patterns]),
                json.dumps(column_mappings),
                json.dumps(key_value_rules),
                json.dumps(validation_rules),
                sample_text[:1500]
            ))
            conn.commit()
            return cursor.lastrowid
        finally:
            conn.close()

    def add_chat_instruction(self, user_prompt: str, action_type: str, rule_payload: Dict[str, Any],
                             template_id: Optional[int] = None, is_global: bool = True) -> int:
        """Store instructions given in conversation (e.g. 'Always convert unit rate to Float')."""
        conn = self._get_connection()
        try:
            cursor = conn.cursor()
            cursor.execute("""
            INSERT INTO chat_instructions (template_id, user_prompt, action_type, rule_payload, is_global)
            VALUES (?, ?, ?, ?, ?)
            """, (template_id, user_prompt, action_type, json.dumps(rule_payload), 1 if is_global else 0))
            conn.commit()
            return cursor.lastrowid
        finally:
            conn.close()

    def get_learned_rules(self, template_id: Optional[int] = None) -> List[Dict[str, Any]]:
        """Retrieve all active chat instructions (both global and template-specific)."""
        conn = self._get_connection()
        try:
            cursor = conn.cursor()
            if template_id:
                cursor.execute("""
                SELECT * FROM chat_instructions 
                WHERE is_global = 1 OR template_id = ?
                ORDER BY created_at ASC
                """, (template_id,))
            else:
                cursor.execute("SELECT * FROM chat_instructions WHERE is_global = 1 ORDER BY created_at ASC")
            
            rows = cursor.fetchall()
            results = []
            for r in rows:
                item = dict(r)
                item["rule_payload"] = json.loads(r["rule_payload"])
                results.append(item)
            return results
        finally:
            conn.close()

    def add_synonym(self, standard_name: str, alias: str):
        conn = self._get_connection()
        try:
            cursor = conn.cursor()
            cursor.execute("""
            INSERT OR REPLACE INTO column_synonyms (standard_name, alias)
            VALUES (?, ?)
            """, (standard_name.strip(), alias.lower().strip()))
            conn.commit()
        finally:
            conn.close()

    def get_all_synonyms(self) -> Dict[str, str]:
        """Returns mapping of alias -> standard_name."""
        conn = self._get_connection()
        try:
            cursor = conn.cursor()
            cursor.execute("SELECT standard_name, alias FROM column_synonyms")
            return {row["alias"].lower(): row["standard_name"] for row in cursor.fetchall()}
        finally:
            conn.close()

    def record_extraction_run(self, pdf_filename: str, file_hash: str, template_name: Optional[str],
                              table_rows: int, kv_count: int, accuracy_score: float,
                              excel_path: str, raw_data: Any, final_data: Any, notes: str = "") -> int:
        """Record the complete run history to allow historical auditing and iterative learning."""
        conn = self._get_connection()
        try:
            cursor = conn.cursor()
            cursor.execute("""
            INSERT INTO extraction_runs (
                pdf_filename, file_hash, template_name, table_row_count, key_value_count,
                accuracy_score, excel_output_path, raw_extracted_json, final_data_json, notes
            ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
            """, (
                pdf_filename,
                file_hash,
                template_name,
                table_rows,
                kv_count,
                accuracy_score,
                excel_path,
                json.dumps(raw_data, default=str),
                json.dumps(final_data, default=str),
                notes
            ))
            conn.commit()
            return cursor.lastrowid
        finally:
            conn.close()

    def update_run_corrections(self, run_id: int, user_corrections: Dict[str, Any], feedback_notes: str = ""):
        """When a user corrects cells in Excel or Web UI, log the correction so future runs are exact."""
        conn = self._get_connection()
        try:
            cursor = conn.cursor()
            cursor.execute("""
            UPDATE extraction_runs
            SET user_corrections_json = ?, notes = notes || ' | Feedback: ' || ?
            WHERE id = ?
            """, (json.dumps(user_corrections), feedback_notes, run_id))
            conn.commit()
        finally:
            conn.close()

    def get_history(self, limit: int = 50) -> List[Dict[str, Any]]:
        conn = self._get_connection()
        try:
            cursor = conn.cursor()
            cursor.execute("""
            SELECT id, pdf_filename, template_name, table_row_count, key_value_count,
                   accuracy_score, excel_output_path, created_at, notes
            FROM extraction_runs
            ORDER BY created_at DESC LIMIT ?
            """, (limit,))
            return [dict(row) for row in cursor.fetchall()]
        finally:
            conn.close()

    def get_templates(self) -> List[Dict[str, Any]]:
        conn = self._get_connection()
        try:
            cursor = conn.cursor()
            cursor.execute("SELECT * FROM templates ORDER BY usage_count DESC")
            templates = []
            for row in cursor.fetchall():
                t = dict(row)
                t["header_patterns"] = json.loads(row["header_patterns"])
                t["column_mappings"] = json.loads(row["column_mappings"])
                t["key_value_rules"] = json.loads(row["key_value_rules"])
                t["validation_rules"] = json.loads(row["validation_rules"])
                templates.append(t)
            return templates
        finally:
            conn.close()
