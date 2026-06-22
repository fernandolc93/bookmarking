import pdfplumber
import pikepdf
import pandas as pd
from pathlib import Path
import re

# ── Configuration ─────────────────────────────────────────────────────────────
TITLE_Y_MIN     = 40
TITLE_Y_MAX     = 80
TITLE_X_COL_END = 244

TITLE_FIXES = {
    "D isease History - ALCL_Initial diagnosis" : "Disease History - ALCL_Initial diagnosis",
    "D isease History - ALCL_More events"        : "Disease History - ALCL_More events",
    "D isease History - IMT_Initial diagnosis"   : "Disease History - IMT_Initial diagnosis",
    "D isease History - IMT_More events"         : "Disease History - IMT_More events",
}

_STOP_WORDS = {"question", "controlname", "values"}
_SEP_RE     = re.compile(r"^_{40,}\s*$")
_HEADER_LINE = "Question ControlName Values"
# ─────────────────────────────────────────────────────────────────────────────

def detect_source_type(path: Path) -> str:
    """Returns 'pdf' or 'text'."""
    with open(path, "rb") as fh:
        magic = fh.read(5)
    if magic == b"%PDF-":
        return "pdf"
    try:
        with pikepdf.open(path):
            pass
        return "pdf"
    except Exception:
        return "text"

def extract_title_pdf(page) -> str | None:
    words = page.extract_words()
    
    title_words = [
        w for w in words
        if TITLE_Y_MIN <= w["top"] <= TITLE_Y_MAX
        and w["x0"] < TITLE_X_COL_END
        and w["text"].lower() not in _STOP_WORDS
    ]
    if title_words:
        title_words.sort(key=lambda w: (round(w["top"] / 4) * 4, w["x0"]))
        title = " ".join(w["text"] for w in title_words).strip()
        return TITLE_FIXES.get(title, title)

    chars = page.chars
    if not chars:
        return None
    from statistics import median
    body_size = median(c["size"] for c in chars)
    lines = {}
    for c in chars:
        lines.setdefault(round(c["top"] / 4) * 4, []).append(c)
    for y in sorted(lines):
        lc   = lines[y]
        bold = any("bold" in (c.get("fontname") or "").lower() for c in lc)
        big  = any(c["size"] > body_size * 1.1 for c in lc)
        if bold or big:
            lc.sort(key=lambda c: c["x0"])
            text = "".join(c["text"] for c in lc).strip()
            if len(text) > 3:
                return TITLE_FIXES.get(text, text)
    return None

def extract_titles_from_text(path: Path) -> tuple[list[tuple[int, str | None]], int]:
    text  = path.read_text(encoding="utf-8", errors="replace")
    lines = text.split("\n")
    page_titles = []
    section_num = 0
    for i, line in enumerate(lines):
        if _SEP_RE.match(line.rstrip()):
            section_num += 1
            title = None
            for j in range(i + 1, min(i + 6, len(lines))):
                candidate = lines[j].strip()
                if not candidate or candidate.startswith("_") or candidate == _HEADER_LINE:
                    continue
                title = candidate
                break
            if title:
                title = TITLE_FIXES.get(title, title)
            page_titles.append((section_num, title))
    return page_titles, section_num

def extract_forms_from_pdf(pdf_file_path: str | Path) -> pd.DataFrame:
    path = Path(pdf_file_path)
    source_type = detect_source_type(path)
    
    if source_type == "pdf":
        page_titles = []
        with pdfplumber.open(path) as pdf:
            total_pages = len(pdf.pages)
            for i, page in enumerate(pdf.pages, start=1):
                title = extract_title_pdf(page)
                page_titles.append((i, title))
    else:
        page_titles, total_pages = extract_titles_from_text(path)
        
    forms = []
    prev_title = None
    for page, title in page_titles:
        if title != prev_title:
            if title is not None:
                forms.append({"Form Name": title, "Start Page": page})
            prev_title = title
            
    df = pd.DataFrame(forms)
    if not df.empty:
        df["End Page"] = (df["Start Page"].shift(-1).fillna(total_pages + 1).astype(int) - 1)
    else:
        df = pd.DataFrame(columns=["Form Name", "Start Page", "End Page"])
        
    return df
