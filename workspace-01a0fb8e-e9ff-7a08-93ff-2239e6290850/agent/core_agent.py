"""
Core Offline Agent Orchestrator.
Coordinates PDF parsing, historical memory retrieval, schema normalization,
exact mathematical verification, and Excel workbook synthesis.
"""

import os
import re
import hashlib
from pathlib import Path
from typing import Dict, List, Any, Optional

from agent.memory import AgentMemory
from agent.parser import PDFParser
from agent.schema_matcher import SchemaMatcher
from agent.validator import ExactResultValidator
from agent.excel_exporter import ExcelExporter


_PROJECT_ROOT = Path(__file__).resolve().parent.parent


class PDFToExcelAgent:
    def __init__(self, db_path: str = str(_PROJECT_ROOT / "history" / "agent_memory.db"),
                 default_output_dir: str = str(_PROJECT_ROOT / "outputs")):
        self.memory = AgentMemory(db_path=db_path)
        self.parser = PDFParser()
        self.schema_matcher = SchemaMatcher(self.memory)
        self.validator = ExactResultValidator()
        self.exporter = ExcelExporter()
        self.default_output_dir = Path(default_output_dir)
        self.default_output_dir.mkdir(parents=True, exist_ok=True)

    def process_pdf(self, pdf_path: str, output_excel_path: Optional[str] = None,
                    template_name: Optional[str] = None,
                    chat_instruction: Optional[str] = None) -> Dict[str, Any]:
        """
        End-to-End Extraction Pipeline:
        1. Parse PDF layout, tables & key-values
        2. Query Historical Memory for matching templates & past instructions
        3. Interpret & apply any new chat instructions
        4. Normalize table schema to standard types
        5. Exact mathematical verification & reconciliation
        6. Export styled Excel workbook
        7. Save session history for continuous offline learning
        """
        pdf_path = str(Path(pdf_path).resolve())
        if not os.path.exists(pdf_path):
            raise FileNotFoundError(f"PDF not found: {pdf_path}")

        # Compute file hash
        with open(pdf_path, 'rb') as f:
            file_hash = hashlib.sha256(f.read()).hexdigest()[:16]

        # 1. Multi-Engine Extraction — keep every table/transaction from the PDF
        parsed_doc = self.parser.parse(pdf_path)
        full_text = parsed_doc["full_text"]
        raw_tables = parsed_doc["tables"]
        key_values = parsed_doc["key_values"]
        source_filename = os.path.basename(pdf_path)

        is_bank_statement = parsed_doc.get("is_bank_statement", False)
        if is_bank_statement:
            primary_table = raw_tables[0]
        elif raw_tables:
            primary_table = self._combine_all_tables(raw_tables, source_filename)
        else:
            primary_table = {
                "page": 1,
                "headers": ["Source PDF", "Field", "Extracted Value"],
                "rows": [[source_filename, k, str(v)] for k, v in key_values.items()]
            }

        # 2. Historical Memory Lookup
        matched_template = None
        if template_name:
            # Look up specific template
            templates = self.memory.get_templates()
            for t in templates:
                if t["name"].lower() == template_name.lower():
                    matched_template = t
                    break
        else:
            # Auto-find matching template from past works and history
            matched_template = self.memory.find_matching_template(
                text=full_text,
                headers=primary_table.get("headers", [])
            )

        # 3. Retrieve Historical Chat Rules
        template_id = matched_template["id"] if matched_template else None
        learned_rules = self.memory.get_learned_rules(template_id=template_id)

        # 4. Process new ad-hoc chat instruction if provided
        if chat_instruction:
            rule = self._interpret_chat_instruction(chat_instruction)
            if rule:
                self.memory.add_chat_instruction(
                    user_prompt=chat_instruction,
                    action_type=rule["action_type"],
                    rule_payload=rule["rule_payload"],
                    template_id=template_id,
                    is_global=rule.get("is_global", True)
                )
                learned_rules.append(rule)

        # 5. Schema Normalization & Column Mapping
        standardized_table = self.schema_matcher.standardize_table(
            table=primary_table,
            template=matched_template,
            chat_rules=learned_rules
        )
        if is_bank_statement:
            standardized_table = {
                **standardized_table,
                "headers": [
                    "Txn Date",
                    "Value Date",
                    "Particulars",
                    "Ref. No.",
                    "Debit",
                    "Credit",
                    "Balance",
                ],
                "column_types": [
                    "date",
                    "date",
                    "text",
                    "text",
                    "currency",
                    "currency",
                    "currency",
                ],
            }

        # 6. Exact Result Verification & Reconciliations
        val_result = self.validator.validate_and_reconcile(
            table=standardized_table,
            key_values=key_values,
            custom_validation_rules=matched_template.get("validation_rules") if matched_template else None
        )
        final_table = val_result["table"]
        accuracy_score = val_result["accuracy_score"]

        # 7. Export statements separately; merge other documents into the master workbook.
        master_path = str(self.default_output_dir / "all_transactions.xlsx")

        meta_info = {
            "filename": source_filename,
            "template_name": matched_template["name"] if matched_template else "Auto-Discovered Schema",
            "accuracy_score": accuracy_score,
            "is_scanned": parsed_doc["is_scanned"],
            "currency": "INR" if is_bank_statement else key_values.get("Currency"),
            "is_bank_statement": is_bank_statement,
        }

        if is_bank_statement:
            output_excel_path = output_excel_path or str(
                self.default_output_dir / f"{Path(source_filename).stem}_extracted.xlsx"
            )
            self.exporter.export(
                output_path=output_excel_path,
                table_data=final_table,
                key_values=key_values,
                validation_info=val_result,
                metadata=meta_info
            )
            excel_file = str(Path(output_excel_path).resolve())
            master_path = excel_file
        else:
            if not output_excel_path:
                output_excel_path = master_path
            excel_file = self.exporter.merge_into_master(
                master_path=master_path,
                table_data=final_table,
                key_values=key_values,
                validation_info=val_result,
                metadata=meta_info
            )

            if str(Path(output_excel_path).resolve()) != str(Path(master_path).resolve()):
                self.exporter.export(
                    output_path=output_excel_path,
                    table_data=final_table,
                    key_values=key_values,
                    validation_info=val_result,
                    metadata=meta_info
                )
                excel_file = str(Path(output_excel_path).resolve())

        # 8. Record in Extraction History
        run_id = self.memory.record_extraction_run(
            pdf_filename=os.path.basename(pdf_path),
            file_hash=file_hash,
            template_name=matched_template["name"] if matched_template else "Standard Auto-Mapped",
            table_rows=len(final_table.get("rows", [])),
            kv_count=len(key_values),
            accuracy_score=accuracy_score,
            excel_path=excel_file,
            raw_data=parsed_doc,
            final_data={"table": final_table, "key_values": key_values},
            notes=f"Processed with {len(learned_rules)} learned rules from history"
        )

        return {
            "run_id": run_id,
            "pdf_path": pdf_path,
            "excel_path": excel_file,
            "master_excel_path": master_path,
            "template_used": matched_template["name"] if matched_template else "Standard Auto-Mapped",
            "template_confidence": matched_template.get("match_confidence", 0.95) if matched_template else 0.90,
            "accuracy_score": accuracy_score,
            "headers": final_table["headers"],
            "rows": final_table["rows"],
            "row_count": len(final_table["rows"]),
            "key_values": key_values,
            "is_bank_statement": is_bank_statement,
            "validation_audit": val_result["audit_log"],
            "rules_applied_count": len(learned_rules)
        }

    def learn_from_chat(self, chat_prompt: str, template_name: Optional[str] = None) -> Dict[str, Any]:
        """
        Interprets natural language instructions offline and updates agent memory.
        Examples:
        - "Rename particulars to Item Description"
        - "Map qty to Quantity and rate to Unit Price"
        - "Save this template as Acme Invoices"
        """
        rule = self._interpret_chat_instruction(chat_prompt)
        if not rule:
            return {"status": "unrecognized", "message": "Could not identify a specific rule from prompt."}

        # Check if user wants to save a template
        if rule["action_type"] == "save_template_name":
            t_name = rule["rule_payload"].get("name")
            return {
                "status": "success",
                "action": "save_template",
                "template_name": t_name,
                "message": f"Recorded template alias '{t_name}'. It will be matched to similar documents."
            }

        # Save general rule
        inst_id = self.memory.add_chat_instruction(
            user_prompt=chat_prompt,
            action_type=rule["action_type"],
            rule_payload=rule["rule_payload"],
            template_id=None,
            is_global=True
        )

        # If it's a synonym/rename, also store directly in synonym dictionary
        if rule["action_type"] == "rename_col":
            src = rule["rule_payload"].get("source")
            tgt = rule["rule_payload"].get("target")
            if src and tgt:
                self.memory.add_synonym(tgt, src)

        return {
            "status": "success",
            "instruction_id": inst_id,
            "action": rule["action_type"],
            "payload": rule["rule_payload"],
            "message": f"Learned new rule from chat: {chat_prompt}"
        }

    def _interpret_chat_instruction(self, prompt: str) -> Optional[Dict[str, Any]]:
        """Offline rule parser that maps natural language statements to structured actions."""
        p = prompt.strip()
        p_lower = p.lower()

        # Pattern 1: Rename Column (e.g. "Rename Particulars to Description" or "Change col 1 to Item")
        rename_match = re.search(r'(?:rename|change|map|set)\s+[\'"]?([a-zA-Z0-9_\s#]+)[\'"]?\s+(?:to|as)\s+[\'"]?([a-zA-Z0-9_\s#]+)[\'"]?', p, re.IGNORECASE)
        if rename_match:
            source = rename_match.group(1).strip()
            target = rename_match.group(2).strip()
            return {
                "action_type": "rename_col",
                "rule_payload": {"source": source, "target": target},
                "is_global": True
            }

        # Pattern 2: Template naming (e.g. "Save template as Acme Corp" or "Remember format as Medical Bills")
        template_match = re.search(r'(?:save|remember|store)\s+(?:this\s+)?(?:template|format|layout)\s+as\s+[\'"]?([a-zA-Z0-9_\s]+)[\'"]?', p, re.IGNORECASE)
        if template_match:
            t_name = template_match.group(1).strip()
            return {
                "action_type": "save_template_name",
                "rule_payload": {"name": t_name},
                "is_global": False
            }

        # Pattern 3: Date formatting (e.g. "Use date format DD/MM/YYYY" or "Set date to ISO")
        date_match = re.search(r'(?:date\s*format|format\s*date)\s+(?:to|as|is)\s+([a-zA-Z0-9/\-\.]+)', p, re.IGNORECASE)
        if date_match:
            fmt = date_match.group(1).strip()
            return {
                "action_type": "date_format",
                "rule_payload": {"format": fmt},
                "is_global": True
            }

        # Pattern 4: Ignore row / header filter
        ignore_match = re.search(r'ignore\s+rows?\s+(?:containing|with)\s+[\'"]?([a-zA-Z0-9_\s]+)[\'"]?', p, re.IGNORECASE)
        if ignore_match:
            keyword = ignore_match.group(1).strip()
            return {
                "action_type": "ignore_row",
                "rule_payload": {"keyword": keyword},
                "is_global": True
            }

        # Fallback synonym pattern: "X means Y" or "X is Y"
        means_match = re.search(r'[\'"]?([a-zA-Z0-9_\s]+)[\'"]?\s+(?:means|refers to|equals|stands for)\s+[\'"]?([a-zA-Z0-9_\s]+)[\'"]?', p, re.IGNORECASE)
        if means_match:
            source = means_match.group(1).strip()
            target = means_match.group(2).strip()
            return {
                "action_type": "rename_col",
                "rule_payload": {"source": source, "target": target},
                "is_global": True
            }

        return None

    def _combine_all_tables(self, tables: List[Dict[str, Any]], source_filename: str) -> Dict[str, Any]:
        """Union every extracted table so all PDF transactions land in one sheet."""
        source_col = "Source PDF"

        def norm(header: str) -> str:
            return re.sub(r"\s+", " ", str(header).strip().lower())

        ordered_headers: List[str] = [source_col]
        seen = {norm(source_col)}
        for tbl in tables:
            for header in tbl.get("headers") or []:
                key = norm(header)
                if not key or key in seen:
                    continue
                seen.add(key)
                ordered_headers.append(header)

        combined_rows: List[List[Any]] = []
        for tbl in tables:
            headers = tbl.get("headers") or []
            index = {norm(h): i for i, h in enumerate(headers)}
            for row in tbl.get("rows") or []:
                mapped = []
                for header in ordered_headers:
                    if norm(header) == norm(source_col):
                        mapped.append(source_filename)
                        continue
                    col_i = index.get(norm(header))
                    mapped.append(row[col_i] if col_i is not None and col_i < len(row) else "")
                if any(str(c).strip() for c in mapped[1:]):
                    combined_rows.append(mapped)

        return {
            "page": tables[0].get("page", 1) if tables else 1,
            "headers": ordered_headers,
            "rows": combined_rows
        }

    def save_current_as_template(self, template_name: str, pdf_path: str,
                                 column_mappings: Optional[Dict[str, str]] = None) -> int:
        """Explicitly register a PDF format as a named template for future auto-detection."""
        parsed_doc = self.parser.parse(pdf_path)
        tables = parsed_doc["tables"]
        headers = tables[0]["headers"] if tables else []
        fingerprint = AgentMemory.generate_fingerprint(parsed_doc["full_text"], headers)

        if not column_mappings and tables:
            synonyms = self.memory.get_all_synonyms()
            column_mappings = {}
            for h in headers:
                h_clean = h.strip().lower()
                column_mappings[h] = synonyms.get(h_clean, h)

        return self.memory.save_template(
            name=template_name,
            fingerprint=fingerprint,
            header_patterns=headers,
            column_mappings=column_mappings or {},
            key_value_rules=parsed_doc["key_values"],
            sample_text=parsed_doc["full_text"][:1000]
        )

    def batch_process(self, input_dir: str, output_dir: Optional[str] = None) -> List[Dict[str, Any]]:
        """Batch process all PDF documents in a directory."""
        in_path = Path(input_dir)
        out_path = Path(output_dir) if output_dir else self.default_output_dir
        out_path.mkdir(parents=True, exist_ok=True)

        results = []
        pdf_files = list(in_path.glob("*.pdf")) + list(in_path.glob("*.PDF"))
        original_out = self.default_output_dir
        self.default_output_dir = out_path
        out_path.mkdir(parents=True, exist_ok=True)

        try:
            for pdf_file in pdf_files:
                try:
                    res = self.process_pdf(str(pdf_file))
                    results.append(res)
                except Exception as e:
                    results.append({
                        "pdf_path": str(pdf_file),
                        "status": "error",
                        "error": str(e)
                    })
        finally:
            self.default_output_dir = original_out

        return results

    def get_history(self, limit: int = 50) -> List[Dict[str, Any]]:
        return self.memory.get_history(limit=limit)

    def get_templates(self) -> List[Dict[str, Any]]:
        return self.memory.get_templates()
