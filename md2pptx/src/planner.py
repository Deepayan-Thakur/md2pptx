"""
planner.py — Gemini AI-powered slide planning.
Takes the parsed markdown structure and returns an ordered slide plan as JSON.
"""
import os
import json
import re
from typing import Dict, Any, List, Optional
from openai import OpenAI


def _get_openrouter_client():
    api_key = os.getenv("OPENROUTER_API_KEY")
    if not api_key:
        raise EnvironmentError("OPENROUTER_API_KEY not set in environment / .env file")
    return OpenAI(
        base_url="https://openrouter.ai/api/v1",
        api_key=api_key
    )


SYSTEM_CONTEXT = """You are a senior McKinsey/Accenture presentation designer.
Your task: create a professional 12-14 slide plan from a structured content brief and full markdown text.

Return ONLY a valid JSON array — no markdown, no explanation, no code fences.

Each slide object MUST have these exact keys:
{
  "slide_number": <int>,
  "type": "<title|agenda|exec_summary|section_divider|two_column|monitoring|content|data_table|data_chart|infographic|conclusion|thankyou>",
  "infographic_type": "<timeline|process|comparison|none>",
  "layout": "<cover|divider|blank|title_only>",
  "title": "<string, max 8 words>",
  "subtitle": "<string or ''>",
  "bullets": ["<highly detailed, comprehensive long-form sentences (20-40 words each). MINIMUM 4 items>"],
  "left_bullets": ["<detailed explanatory long-form bullet points>"],
  "right_bullets": ["<detailed explanatory long-form bullet points>"],
  "has_chart": <true|false>,
  "has_table": <true|false>,
  "table_index": <int or -1>,
  "section_ref": "<which H2 section>"
}

Slide flow MANDATED RULES:
1. Slide 1: type=title, layout=cover
2. Slide 2: type=agenda, layout=blank (list main H2 sections)
3. Slide 3: type=exec_summary, layout=blank (3-4 key synthesized insights)
4. Slide 4: type=agentic_logic, layout=blank (Describe the system architecture: Parser -> Planner -> Builder)
5. Slide 5-12: Section content. USE VARIETY:
   - Use `type=infographic, infographic_type=process` for steps, workflows, or sequences.
   - Use `type=infographic, infographic_type=timeline` for dates, history, or schedules.
   - Use `type=monitoring` for slides discussing thresholds, metrics, risk levels, or KPIs.
   - Use `type=two_column` for comparisons or balanced lists.
   - Use `type=data_chart` whenever numerical data is present in the brief.
6. Slide 13: type=conclusion, layout=blank (3-5 key takeaways)
7. Slide 14: type=thankyou, layout=cover

VISUAL VARIETY & CONTENT rules:
- DENSE CONTENT: The presentation must look densely populated. Provide highly detailed, comprehensive explanations. Do not use short fragments. Each bullet MUST be a full, detailed paragraph. Use heavily detailed professional language.
- DATA ENFORCEMENT: If statistics exist, you MUST use `data_chart`. DO NOT invent numbers.
- INFOGRAPHICS: Actively use `infographic` type for process-oriented or time-oriented sections.
"""



def plan_slides(parsed: Dict[str, Any], md_text: str, target_slides: int = 13, provider: str = "openrouter") -> List[Dict]:
    """Call OpenRouter AI to produce a slide plan. Falls back to rule-based plan on failure."""
    brief = _build_brief(parsed, target_slides)
    try:
        client = _get_openrouter_client()
        prompt = (
            f"{SYSTEM_CONTEXT}\n\n"
            f"Here is the full Markdown document content:\n{md_text[:14000]}\n\n"
            f"Extracted structural brief (use for section hierarchy):\n{json.dumps(brief, indent=2)}\n\n"
            f"Return the JSON array now:"
        )
        model = os.getenv("OPENROUTER_MODEL", "deepseek/deepseek-r1")
        response = client.chat.completions.create(
            model=model,
            messages=[{"role": "user", "content": prompt}],
            temperature=0.3,
            max_tokens=4000,
        )
        text = response.choices[0].message.content.strip()
        # Remove deepseek <think> tags if present
        text = re.sub(r'<think>.*?</think>', '', text, flags=re.IGNORECASE | re.DOTALL).strip()
        text = re.sub(r'^```(?:json)?', '', text, flags=re.MULTILINE)
        text = re.sub(r'```$', '', text, flags=re.MULTILINE).strip()
        start_idx = text.find('[')
        end_idx = text.rfind(']')
        if start_idx != -1 and end_idx != -1:
            text = text[start_idx:end_idx+1]
        text = re.sub(r',\s*([\]}])', r'\1', text)
        plan = json.loads(text)
        if isinstance(plan, list) and len(plan) >= 5:
            print(f"   OpenRouter ({model}) planned {len(plan)} slides")
            return _validate_plan(plan, parsed)
    except Exception as exc:
        print(f"   OpenRouter planning failed ({exc}), using rule-based fallback")
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
        "infographic_type": "none", "section_ref": "",
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