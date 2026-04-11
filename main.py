#!/usr/bin/env python3
"""
md2pptx — AI-Powered Markdown to Professional PPTX Generator
Hackathon: Accenture EZ | April 9-12, 2026

Usage:
    python main.py <input.md> [output.pptx] [--slides N]

Examples:
    python main.py test_cases/accenture_tech.md
    python main.py test_cases/accenture_tech.md outputs/my_deck.pptx --slides 13
"""

import argparse
import os
import sys
import time
from pathlib import Path

# Load .env before importing modules that need API keys
try:
    from dotenv import load_dotenv
    load_dotenv()
except ImportError:
    pass

from md2pptx.src.parser import parse_markdown
from md2pptx.src.planner import plan_slides
from md2pptx.src.builder import generate_pptx


def _banner():
    print("""
+------------------------------------------------------+
|          md2pptx  •  EZ Hackathon 2026        |
|    AI-Powered Markdown -> Professional PPTX          |
+------------------------------------------------------+
""")


def main():
    _banner()
    parser = argparse.ArgumentParser(
        description="Convert Markdown to Professional PPTX using OpenRouter AI"
    )
    parser.add_argument("input", help="Path to input .md file")
    parser.add_argument("output", nargs="?", help="Output .pptx path (optional)")
    parser.add_argument("--slides", type=int, default=13,
                        help="Target slide count (10-15, default: 13)")
    parser.add_argument("--provider", choices=["openrouter", "gemini", "huggingface"], default="openrouter",
                        help="API provider for slide planning (default: openrouter)")
    parser.add_argument("--template", default="template.pptx",
                        help="Base template filename in assets/ (e.g. template1.pptx)")
    args = parser.parse_args()

    # ── Validate input ──────────────────────────────────────────────────────
    input_path = Path(args.input)
    if not input_path.exists():
        print(f"ERROR: File not found: {input_path}")
        sys.exit(1)
    if input_path.suffix.lower() != ".md":
        print(f"ERROR: Input must be a .md file (got: {input_path.suffix})")
        sys.exit(1)
    if args.provider == "openrouter" and not os.getenv("OPENROUTER_API_KEY"):
        print("WARN: OPENROUTER_API_KEY not set - will use rule-based fallback planner")
    elif args.provider == "gemini" and not os.getenv("GEMINI_API_KEY"):
        print("WARN: GEMINI_API_KEY not set - will use rule-based fallback planner")
    elif args.provider == "huggingface" and not os.getenv("HUGGINGFACE_API_KEY"):
        print("WARN: HUGGINGFACE_API_KEY not set - will use rule-based fallback planner")

    # ── Check File Size (Constraint: Max 5MB) ───────────────────────────────
    MAX_SIZE = 5 * 1024 * 1024  # 5 MB
    file_size = input_path.stat().st_size
    if file_size > MAX_SIZE:
        print(f"ERROR: Input file is too large ({file_size / (1024*1024):.1f} MB). "
              f"Max allowed is 5 MB.")
        sys.exit(1)

    # ── Output path ─────────────────────────────────────────────────────────
    if args.output:
        output_path = os.path.abspath(args.output)
    else:
        out_dir = os.path.join("md2pptx", "outputs")
        os.makedirs(out_dir, exist_ok=True)
        output_path = os.path.join(out_dir, f"{input_path.stem}.pptx")
        output_path = os.path.abspath(output_path)

    target_slides = max(10, min(15, args.slides))

    # ── Read markdown ────────────────────────────────────────────────────────
    print(f"Reading:  {input_path}")
    with open(input_path, "r", encoding="utf-8") as f:
        md_text = f.read()
    print(f"   {len(md_text):,} characters | "
          f"{md_text.count(chr(10)):,} lines")

    # ── Parse ────────────────────────────────────────────────────────────────
    t0 = time.time()
    print("\nParsing markdown structure...")
    parsed = parse_markdown(md_text)
    print(f"   Title:    {parsed['title'][:60]}")
    print(f"   Sections: {len(parsed['sections'])}")
    print(f"   Tables:   {len(parsed['all_tables'])}")
    print(f"   Has numerical data: {parsed['has_numerical_data']}")
    print(f"   ({time.time() - t0:.1f}s)")

    # ── Plan slides ───────────────────────────────────────────────────────────
    print(f"\nPlanning slides with {args.provider.capitalize()} AI...")
    t1 = time.time()
    slide_plan = plan_slides(parsed, md_text, target_slides=target_slides, provider=args.provider)
    print(f"   Planned {len(slide_plan)} slides  ({time.time() - t1:.1f}s)")

    # Print plan summary
    for sp in slide_plan:
        icon = {
            "title": "[TITLE] ", "agenda": "[AGENDA]", "exec_summary": "[EXEC  ]",
            "section_divider": "[DIVIDE]", "content": "[CONT  ]", "two_column": "[2-COL ]",
            "data_chart": "[CHART ]", "data_table": "[TABLE ]", "conclusion": "[CONCL ]",
            "thankyou": "[THANKS]",
        }.get(sp.get("type", ""), "  ")
        # Sanitize title for terminal output to avoid UnicodeEncodeError on Windows
        safe_title = sp.get('title', '').encode('utf-8', errors='ignore').decode('utf-8', errors='ignore')[:50]
        try:
            print(f"   Slide {sp['slide_number']:02d} {icon} [{sp.get('type','?'):<15}] {safe_title}")
        except UnicodeEncodeError:
            # Fallback for environments with restricted encoding
            print(f"   Slide {sp['slide_number']:02d} {icon} [{sp.get('type','?'):<15}] [Title contains special characters]")

    # ── Build PPTX ────────────────────────────────────────────────────────────
    print(f"\nBuilding PPTX...")
    t2 = time.time()
    try:
        generate_pptx(parsed, slide_plan, output_path, args.template)
        print(f"   Built in {time.time() - t2:.1f}s")
    except PermissionError:
        print(f"\nERROR: Permission Denied while saving to: {output_path}")
        print("   This usually happens because the file is OPEN in PowerPoint.")
        print("   Please close the file and try again.")
        sys.exit(1)
    except Exception as e:
        print(f"\nERROR during build: {e}")
        sys.exit(1)

    total = time.time() - t0
    print(f"\n{'-'*54}")
    print(f"DONE in {total:.1f}s")
    print(f"Output: {os.path.abspath(output_path)}")
    print(f"{'-'*54}")


if __name__ == "__main__":
    main()