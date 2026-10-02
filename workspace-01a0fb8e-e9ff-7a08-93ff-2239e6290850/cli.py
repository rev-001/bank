#!/usr/bin/env python3
"""
CLI Tool for Offline PDF-to-Excel Agent.
Supports single-file extraction, batch conversion, interactive chat learning,
history inspection, and template management.
"""

import argparse
import sys
import os
from pathlib import Path
from rich.console import Console
from rich.table import Table
from rich.panel import Panel
from rich.prompt import Prompt

from agent.core_agent import PDFToExcelAgent

console = Console()


def print_banner():
    console.print(Panel.fit(
        "[bold cyan]OFFLINE PDF-TO-EXCEL INTELLIGENT AGENT[/bold cyan]\n"
        "[dim]Learns from previous chats, past works, and historical runs to deliver exact results[/dim]",
        border_style="cyan"
    ))


def cmd_process(agent: PDFToExcelAgent, args):
    pdf_path = args.pdf_path
    if not os.path.exists(pdf_path):
        console.print(f"[bold red]Error:[/bold red] PDF file not found: {pdf_path}")
        return

    console.print(f"\n[cyan]Processing PDF:[/cyan] {pdf_path}")
    if args.chat:
        console.print(f"[yellow]Applying Chat Instruction:[/yellow] {args.chat}")

    with console.status("[bold green]Extracting layout, consulting history & calculating exact results..."):
        res = agent.process_pdf(
            pdf_path=pdf_path,
            output_excel_path=args.output,
            template_name=args.template,
            chat_instruction=args.chat
        )

    console.print("\n[bold green]✓ Extraction & Excel Generation Complete![/bold green]")
    console.print(f"[bold]Excel File Saved:[/bold] [green]{res['excel_path']}[/green]")
    console.print(f"[bold]Template Used:[/bold] {res['template_used']} (Confidence: {res['template_confidence']*100:.1f}%)")
    console.print(f"[bold]Accuracy Score:[/bold] {res['accuracy_score']*100:.1f}%")
    console.print(f"[bold]Total Rows Extracted:[/bold] {res['row_count']}")

    # Display Extracted Table Preview
    if res["headers"] and res["rows"]:
        table = Table(title="Extracted Table Preview", show_header=True, header_style="bold magenta")
        for h in res["headers"]:
            table.add_column(str(h))
        for row in res["rows"][:10]:
            table.add_row(*[str(c) for c in row])
        console.print(table)
        if len(res["rows"]) > 10:
            console.print(f"[dim]... and {len(res['rows']) - 10} more rows[/dim]")

    # Display Key-Values
    if res["key_values"]:
        kv_table = Table(title="Extracted Metadata & Key-Values", show_header=True, header_style="bold blue")
        kv_table.add_column("Field")
        kv_table.add_column("Value")
        for k, v in res["key_values"].items():
            kv_table.add_row(k, str(v))
        console.print(kv_table)


def cmd_batch(agent: PDFToExcelAgent, args):
    in_dir = args.input_dir
    if not os.path.isdir(in_dir):
        console.print(f"[bold red]Error:[/bold red] Directory not found: {in_dir}")
        return

    console.print(f"\n[cyan]Batch Processing Directory:[/cyan] {in_dir}")
    results = agent.batch_process(in_dir, args.output)
    
    table = Table(title="Batch Extraction Summary", show_header=True, header_style="bold green")
    table.add_column("File")
    table.add_column("Template")
    table.add_column("Rows", justify="right")
    table.add_column("Accuracy", justify="right")
    table.add_column("Excel Output")

    for r in results:
        if r.get("status") == "error":
            table.add_row(os.path.basename(r["pdf_path"]), "N/A", "0", "[red]Failed[/red]", r.get("error"))
        else:
            table.add_row(
                os.path.basename(r["pdf_path"]),
                r["template_used"],
                str(r["row_count"]),
                f"{r['accuracy_score']*100:.1f}%",
                os.path.basename(r["excel_path"])
            )
    console.print(table)


def cmd_history(agent: PDFToExcelAgent, args):
    history = agent.get_history(limit=args.limit)
    if not history:
        console.print("[yellow]No extraction history recorded yet.[/yellow]")
        return

    table = Table(title="Agent Extraction History (Learned from Past Runs)", show_header=True, header_style="bold cyan")
    table.add_column("Run ID", justify="right")
    table.add_column("PDF File")
    table.add_column("Template")
    table.add_column("Rows", justify="right")
    table.add_column("Accuracy", justify="right")
    table.add_column("Created At")
    table.add_column("Notes")

    for h in history:
        acc_color = "green" if h["accuracy_score"] >= 0.9 else "yellow"
        table.add_row(
            str(h["id"]),
            h["pdf_filename"],
            str(h["template_name"] or "Auto"),
            str(h["table_row_count"]),
            f"[{acc_color}]{h['accuracy_score']*100:.1f}%[/{acc_color}]",
            str(h["created_at"])[:16],
            str(h["notes"] or "")
        )
    console.print(table)


def cmd_templates(agent: PDFToExcelAgent, args):
    templates = agent.get_templates()
    if not templates:
        console.print("[yellow]No custom templates saved yet. Templates are auto-learned during extraction.[/yellow]")
        return

    table = Table(title="Learned Document Templates", show_header=True, header_style="bold magenta")
    table.add_column("ID", justify="right")
    table.add_column("Template Name")
    table.add_column("Headers Recognized")
    table.add_column("Usage Count", justify="right")
    table.add_column("Last Updated")

    for t in templates:
        table.add_row(
            str(t["id"]),
            t["name"],
            ", ".join(t["header_patterns"][:4]) + ("..." if len(t["header_patterns"]) > 4 else ""),
            str(t["usage_count"]),
            str(t["updated_at"])[:16]
        )
    console.print(table)


def cmd_learn(agent: PDFToExcelAgent, args):
    prompt = args.prompt
    console.print(f"\n[cyan]Learning from User Instruction:[/cyan] \"{prompt}\"")
    res = agent.learn_from_chat(prompt)
    if res["status"] == "success":
        console.print(f"[bold green]✓ Rule Recorded Successfully:[/bold green] {res['message']}")
    else:
        console.print(f"[yellow]Note:[/yellow] {res['message']}")


def cmd_interactive(agent: PDFToExcelAgent):
    print_banner()
    console.print("[dim]Type commands like:[/dim]")
    console.print("  • [green]extract path/to/doc.pdf[/green]")
    console.print("  • [green]rename 'particulars' to 'Description'[/green]")
    console.print("  • [green]history[/green] or [green]templates[/green]")
    console.print("  • [green]exit[/green]\n")

    while True:
        try:
            user_input = Prompt.ask("[bold cyan]Agent[/bold cyan]").strip()
            if not user_input:
                continue
            if user_input.lower() in ["exit", "quit", "q"]:
                console.print("[green]Goodbye![/green]")
                break
            
            if user_input.lower() == "history":
                cmd_history(agent, argparse.Namespace(limit=20))
            elif user_input.lower() == "templates":
                cmd_templates(agent, argparse.Namespace())
            elif user_input.lower().startswith("extract ") or user_input.lower().startswith("process "):
                parts = user_input.split(maxsplit=1)
                pdf_p = parts[1].strip().strip('"').strip("'")
                cmd_process(agent, argparse.Namespace(pdf_path=pdf_p, output=None, template=None, chat=None))
            else:
                # Treat as chat instruction learning
                cmd_learn(agent, argparse.Namespace(prompt=user_input))

        except (KeyboardInterrupt, EOFError):
            console.print("\n[green]Exiting agent.[/green]")
            break


def main():
    parser = argparse.ArgumentParser(description="Offline PDF-to-Excel Agent")
    subparsers = parser.add_subparsers(dest="command", help="Available Commands")

    # Process command
    p_proc = subparsers.add_parser("process", help="Extract single PDF into Excel")
    p_proc.add_argument("pdf_path", help="Path to input PDF file")
    p_proc.add_argument("-o", "--output", help="Path to output .xlsx file")
    p_proc.add_argument("-t", "--template", help="Specific template name to apply")
    p_proc.add_argument("-c", "--chat", help="Ad-hoc chat instruction to guide extraction")

    # Batch command
    p_batch = subparsers.add_parser("batch", help="Batch extract directory of PDFs")
    p_batch.add_argument("input_dir", help="Directory containing PDF files")
    p_batch.add_argument("-o", "--output", help="Directory for generated Excel files")

    # History command
    p_hist = subparsers.add_parser("history", help="View past extraction runs & accuracy")
    p_hist.add_argument("-n", "--limit", type=int, default=20, help="Number of records to show")

    # Templates command
    p_tpl = subparsers.add_parser("templates", help="View learned document templates")

    # Learn command
    p_learn = subparsers.add_parser("learn", help="Teach the agent a new rule via chat prompt")
    p_learn.add_argument("prompt", help="Natural language instruction (e.g. 'Rename Rate to Unit Price')")

    # Interactive command
    subparsers.add_parser("interactive", help="Start interactive agent terminal shell")

    args = parser.parse_args()
    agent = PDFToExcelAgent()

    if args.command == "process":
        cmd_process(agent, args)
    elif args.command == "batch":
        cmd_batch(agent, args)
    elif args.command == "history":
        cmd_history(agent, args)
    elif args.command == "templates":
        cmd_templates(agent, args)
    elif args.command == "learn":
        cmd_learn(agent, args)
    elif args.command == "interactive" or args.command is None:
        cmd_interactive(agent)


if __name__ == "__main__":
    main()
