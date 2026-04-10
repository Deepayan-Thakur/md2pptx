"""
planner.py — Gemini AI-powered slide planning.
Takes the parsed markdown structure and returns an ordered slide plan as JSON.
"""
import os
import json
import re
import time
import google.generativeai as genai
from typing import Dict, Any, List, Optional
try:
    from huggingface_hub import InferenceClient
except ImportError:
    InferenceClient = None


def _setup_gemini():
    api_key = os.getenv("GEMINI_API_KEY")
    if not api_key:
        raise EnvironmentError("GEMINI_API_KEY not set in environment / .env file")
    genai.configure(api_key=api_key)
    return genai.GenerativeModel("gemini-2.0-flash")


def _setup_huggingface():
    api_key = os.getenv("HUGGINGFACE_API_KEY")
    if not api_key:
        raise EnvironmentError("HUGGINGFACE_API_KEY not set in environment / .env file")
    if InferenceClient is None:
        raise ImportError("huggingface_hub package not found. Please install it.")
    return InferenceClient("deepseek-ai/DeepSeek-R1-Distill-Qwen-32B", token=api_key)


SYSTEM_CONTEXT = """You are a senior McKinsey/Accenture presentation designer.
Your task: create a professional 12-14 slide plan from a structured content brief and full markdown text.

Return ONLY a valid JSON array — no markdown, no explanation, no code fences.

Each slide object MUST have these exact keys:
{
  "slide_number": <int>,
  "type": "<title|agenda|exec_summary|section_divider|two_column|content|data_table|data_chart|conclusion|thankyou>",
  "layout": "<cover|divider|blank|title_only>",
  "title": "<string, max 8 words>",
  "subtitle": "<string or ''>",
  "bullets": ["<max 6 items, complete informative sentences synthesis>"],
  "left_bullets": ["<for two_column>"],
  "right_bullets": ["<for two_column>"],
  "has_chart": <true|false>,
  "has_table": <true|false>,
  "table_index": <int or -1>,
  "section_ref": "<which H2 section>"
}

Slide flow RULES:
1. Slide 1: type=title, layout=cover (title + subtitle from document)
2. Slide 2: type=agenda, layout=blank (list main H2 sections as bullets)
3. Slide 3: type=exec_summary, layout=blank (3-4 key insight bullets)
4. Slides 4-11: section content — vary layouts! Use two_column for comparisons, data_chart for numerical slides, data_table for tables
5. Second-to-last: type=conclusion, layout=blank (3-5 key takeaways as numbered bullets)
6. Last: type=thankyou, layout=cover

VISUAL VARIETY & CONTENT rules:
- SYNTHESIZE VALUABLE INSIGHTS: Do not just copy/paste text. You must reason over the raw Markdown text provided and synthesize high-value, comprehensive bullets that are practically useful for an executive audience.
- GRAPHICAL DATA ENFORCEMENT: When numerical values, statistics or tables are present, you MUST construct `data_chart` slides. Set `has_chart: true` and specify `table_index` corresponding to the table mapped in the JSON. DO NOT invent numerical values for charts. Let the graph engine do it dynamically.
- Use two_column when section has 4+ points
- Use section_divider before each major theme group
- Max 2 consecutive slides of same type
"""



def plan_slides(parsed: Dict[str, Any], md_text: str, target_slides: int = 13, provider: str = "huggingface") -> List[Dict]:
    """Call AI to produce a slide plan. Falls back to rule-based plan on failure."""
    brief = _build_brief(parsed, target_slides)
    try:
        if provider == "gemini":
            model = _setup_gemini()
            prompt = f"{SYSTEM_CONTEXT}\n\nHere is the full Markdown document content:\n{md_text}\n\nExtracted structural brief (use for section hierarchy):\n{json.dumps(brief, indent=2)}\n\nReturn the JSON array now:"
            response = model.generate_content(prompt)
            text = response.text.strip()
        elif provider == "huggingface":
            client = _setup_huggingface()
            # Truncate md_text to avoid exceeding the strict context window limit 
            # (typically 8k) on the free Hugging Face API serverless inference
            safe_md_text = md_text[:12000] 
            messages = [
                {"role": "system", "content": SYSTEM_CONTEXT},
                {"role": "user", "content": f"Here is the full Markdown document content:\n{safe_md_text}\n\nExtracted structural brief (use for section hierarchy):\n{json.dumps(brief, indent=2)}\n\nReturn the JSON array now:"}
            ]
            response = client.chat_completion(
                messages,
                max_tokens=4000,
                temperature=0.3
            )
            text = response.choices[0].message.content.strip()
            
            # Strip reasoning tags output by DeepSeek-R1 models before parsing
            text = re.sub(r'<think>.*?</think>', '', text, flags=re.DOTALL).strip()
            
        text = re.sub(r'^```(?:json)?', '', text, flags=re.MULTILINE)
        text = re.sub(r'```$', '', text, flags=re.MULTILINE).strip()
        
        # Robust extraction: find the first '[' and last ']' to isolate the array
        start_idx = text.find('[')
        end_idx = text.rfind(']')
        if start_idx != -1 and end_idx != -1:
            text = text[start_idx:end_idx+1]
            
        # Clean up common JSON syntax errors like trailing commas
        text = re.sub(r',\s*([\]}])', r'\1', text)
        
        plan = json.loads(text)
        if isinstance(plan, list) and len(plan) >= 5:
            print(f"   {provider.capitalize()} planned {len(plan)} slides")
            return _validate_plan(plan, parsed)
    except Exception as exc:
        print(f"   {provider.capitalize()} planning failed ({exc}), using rule-based fallback")
    return _fallback_plan(parsed)


def _build_brief(parsed: Dict[str, Any], target: int) -> Dict:
    sections_info = []
    for sec in parsed["sections"][:10]:
        sub_titles = [s.title for s in sec.subsections[:4]]
        bullets = sec.short_bullets(6) if hasattr(sec, 'short_bullets') else []
        sections_info.append({
            "title": sec.title,
            "has_numbers": sec.has_numbers,
            "subsections": sub_titles,
            "preview": sec.content[:1500] if sec.content else ' '.join(bullets)[:1500],
            "has_table": bool(sec.tables),
        })
    tables_info = []
    for idx, tbl in enumerate(parsed["all_tables"][:6]):
        tables_info.append({
            "index": idx,
            "title": tbl.title,
            "headers": tbl.headers,
            "rows": tbl.rows, # Pass full table rows to guarantee graph creation capability
            "is_numerical": tbl.is_numerical,
        })
    return {
        "doc_title": parsed["title"],
        "doc_subtitle": parsed["subtitle"],
        "exec_summary_preview": parsed["executive_summary"][:400],
        "sections": sections_info,
        "tables": tables_info,
        "has_numerical_data": parsed["has_numerical_data"],
        "target_slide_count": target,
    }


def _validate_plan(plan: List[Dict], parsed: Dict) -> List[Dict]:
    """Ensure required fields exist with defaults."""
    defaults = {
        "subtitle": "", "bullets": [], "left_bullets": [], "right_bullets": [],
        "has_chart": False, "chart": None, "has_table": False, "table_index": -1,
        "section_ref": "",
    }
    valid = []
    for i, slide in enumerate(plan):
        s = {**defaults, **slide}
        s["slide_number"] = i + 1
        # Ensure lists are lists
        for key in ("bullets", "left_bullets", "right_bullets"):
            if not isinstance(s[key], list):
                s[key] = [str(s[key])] if s[key] else []
        valid.append(s)
    return valid


def _fallback_plan(parsed: Dict[str, Any]) -> List[Dict]:
    """Rule-based slide plan when Gemini is unavailable."""
    slides = []

    def _slide(type_, layout, title, subtitle="", bullets=None, **kwargs):
        s = {
            "slide_number": len(slides) + 1,
            "type": type_,
            "layout": layout,
            "title": title,
            "subtitle": subtitle,
            "bullets": bullets or [],
            "left_bullets": [],
            "right_bullets": [],
            "has_chart": False,
            "chart": None,
            "has_table": False,
            "table_index": -1,
            "section_ref": "",
        }
        s.update(kwargs)
        slides.append(s)

    # Title
    _slide("title", "cover", parsed["title"], parsed.get("subtitle", ""))

    # Agenda
    agenda_bullets = [sec.title for sec in parsed["sections"][:7]]
    _slide("agenda", "blank", "Agenda", bullets=agenda_bullets)

    # Exec summary
    if parsed.get("executive_summary"):
        from .parser import clean_text
        import re
        es = parsed["executive_summary"]
        bullets = re.split(r'(?<=[.!?])\s+', es)
        bullets = [b for b in bullets if len(b) > 20][:6]
        _slide("exec_summary", "blank", "Executive Summary", bullets=bullets)

    # Sections
    table_idx = 0
    for sec in parsed["sections"][:8]:
        if len(slides) >= 12:
            break
        bs = sec.short_bullets(5) if hasattr(sec, 'short_bullets') else []
        if not bs:
            bs = [sub.title for sub in sec.subsections[:5]]
        if sec.has_numbers and sec.tables:
            tbl = sec.tables[0]
            cd = tbl.chart_data()
            if cd:
                _slide("data_chart", "blank", sec.title,
                       has_chart=True, chart=cd, section_ref=sec.title)
                for global_tbl in parsed["all_tables"]:
                    if global_tbl.title == tbl.title:
                        ti = parsed["all_tables"].index(global_tbl)
                        slides[-1]["has_table"] = True
                        slides[-1]["table_index"] = ti
                        break
            else:
                _slide("data_table", "blank", sec.title,
                       has_table=True, table_index=table_idx, section_ref=sec.title)
            table_idx += 1
        elif len(bs) >= 4:
            mid = len(bs) // 2
            _slide("two_column", "blank", sec.title,
                   left_bullets=bs[:mid], right_bullets=bs[mid:],
                   section_ref=sec.title)
        else:
            _slide("content", "blank", sec.title, bullets=bs, section_ref=sec.title)

    # Conclusion
    conclusion_bullets = [
        "Strategic acquisitions drive AI-powered capability growth",
        "Reinvention Services model delivers integrated multi-domain solutions",
        "Strong financial performance validates acquisition-led strategy",
        "Future focus: agentic AI, sovereign cybersecurity, advanced learning",
    ]
    _slide("conclusion", "blank", "Key Takeaways", bullets=conclusion_bullets)

    # Thank you
    _slide("thankyou", "cover", "Thank You",
           subtitle="Questions & Discussion")

    return slides