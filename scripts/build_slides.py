#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
Build a lesson slide deck (.pptx) and a video script (.md) from one slides.json file.

Usage:
  python build_slides.py <path/to/slides.json> [output folder]
  python build_slides.py <path/to/slides.json> --check-only     (run every check, write nothing)

Output (default: the folder of slides.json):
  <file_name>.pptx              16:9 deck, the narration of each slide in its speaker notes
  <file_name>.video-script.md   one block per slide: on-screen text, production notes,
                                narration written for text-to-speech, estimated seconds

The script enforces the slide standards in references/slide-standards.md and STOPS (exit code 1)
on hard violations: more than 6 lines, text that does not fit at 24 pt, a single orphan word on
the last line, a missing icon or image, an unknown slide type, banned dashes or stock AI phrases.
Softer problems (long lines, digits in narration, the same slide type 3 times in a row, a thin
video) are printed as warnings.
"""
import io
import json
import math
import os
import re
import shutil
import subprocess
import sys
import tempfile

from PIL import Image
from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
from pptx.oxml.ns import qn
from pptx.util import Inches, Pt

if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    sys.stderr.reconfigure(encoding="utf-8", errors="replace")

SKILL_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
ICON_DIR = os.path.join(SKILL_DIR, "assets", "icons")

# ----------------------------------------------------------------------------
# Design constants
# ----------------------------------------------------------------------------
W, H = 13.333, 7.5          # 16:9 canvas in inches
MARGIN = 0.6                # safe area from every edge
FONT = "Arial"              # sans serif, full Latin and Vietnamese coverage, on every machine
SZ_TITLE = 36               # slide title
SZ_BODY = 24                # the smallest size allowed anywhere on a projected slide
SZ_CAPTION = 24             # labels, captions, sources: same floor, no exceptions
SZ_FOOT = 24                # footer and slide number: same floor
MAX_LINES = 6               # 6x6 rule: at most 6 lines
MAX_WORDS = 8               # 8x8 rule: warn above 8 words per line (x2 for full sentences)
WPM = 150                   # narration speed used for timing estimates

THEMES = {
    # three colours: background, text, accent. "card" and "muted" are tints of background and text.
    "navy":   dict(bg="14213D", text="FFFFFF", accent="FCA311", card="1F2F55", muted="CBD5E1"),
    "light":  dict(bg="FFFFFF", text="1E2761", accent="E05A2B", card="EEF2FA", muted="4A5878"),
    "forest": dict(bg="0F3D2E", text="F5F5F5", accent="97BC62", card="19523E", muted="DDEBE2"),
    "plum":   dict(bg="2D1B4E", text="FFFFFF", accent="F9C846", card="3C2A63", muted="D5CBEA"),
}

DEFAULT_TEXT = {
    "en": {"objectives_title": "Learning objectives", "objectives_label": "By the end of this lesson you can",
           "agenda_title": "What this lesson covers", "section_label": "Part", "summary_title": "Three things to remember",
           "summary_label": "Summary", "question_label": "Pause the video and answer", "closing_label": "Do this next",
           "next_lesson": "Next lesson: ", "source": "Source: ", "step": "Step", "script_title": "Video script",
           "pause_phrase": r"pause the video|pause here|hit pause"},
    "vi": {"objectives_title": "Mục tiêu bài học", "objectives_label": "Học xong bài này, bạn làm được",
           "agenda_title": "Nội dung bài học", "section_label": "Phần", "summary_title": "Ba từ khóa cần nhớ",
           "summary_label": "Tóm tắt", "question_label": "Tạm dừng video và trả lời", "closing_label": "Việc làm ngay sau bài này",
           "next_lesson": "Bài tiếp theo: ", "source": "Nguồn: ", "step": "Bước", "script_title": "Kịch bản video",
           "pause_phrase": r"tạm dừng"},
}

ERRORS, WARNINGS = [], []


def err(s, msg):
    ERRORS.append(f"S{s:02d}: {msg}")


def warn(s, msg):
    WARNINGS.append(f"S{s:02d}: {msg}")


def rgb(hexstr):
    return RGBColor.from_string(hexstr.upper())


# ----------------------------------------------------------------------------
# Plain human voice: banned dashes and stock AI phrases (references/plain-human-voice.md)
# ----------------------------------------------------------------------------
BANNED_DASHES = "\u2014\u2013\u2012\u2015\u2010\u2011\u2212\u2043\ufe58\ufe63\uff0d"
AI_PHRASES = {
    "en": ["delve", "in today's fast-paced", "in today's world", "ever-evolving", "fast-paced world", "tapestry",
           "game-changer", "game changer", "unlock the power", "unleash", "harness the power", "seamless",
           "seamlessly", "cutting-edge", "state-of-the-art", "it is worth noting", "it's worth noting",
           "plays a crucial role", "plays a vital role", "plays a pivotal role", "a testament to", "in conclusion",
           "embark on", "let's dive in", "dive into", "navigate the complexities", "elevate your", "holistic",
           "synergy", "paradigm shift", "robust", "leverage"],
    "vi": ["trong bối cảnh", "đóng vai trò quan trọng", "đóng vai trò then chốt", "góp phần quan trọng",
           "hành trình", "bứt phá", "chìa khoá", "chìa khóa", "đa dạng và phong phú", "hãy cùng", "có thể nói",
           "nhìn chung", "điều này cho thấy", "một cách hiệu quả", "nâng tầm", "kỷ nguyên số", "không ngừng thay đổi",
           "mở ra cơ hội", "ngày càng cao", "tóm lại", "đột phá", "độc đáo", "sẽ được trang bị", "toàn diện",
           "tối ưu hoá", "tối ưu hóa", "trong thế giới ngày nay", "không thể phủ nhận", "đó là lý do tại sao",
           "khám phá", "thế giới không ngừng"],
}
NOT_ONLY = {"en": r"not only\b.{0,80}\bbut also", "vi": r"không chỉ(?![a-zà-ỹ]).{0,80}mà còn"}
LETTER = r"a-zà-ỹ"


def check_voice(idx, text, where, lang):
    t = str(text)
    for ch in BANNED_DASHES:
        if ch in t:
            err(idx, f"{where}: banned dash '{ch}' (use '-', a comma, a colon or a new sentence): \"{t[:50]}\"")
            break
    low = t.lower()
    for code in {lang, "en"}:
        pat = NOT_ONLY.get(code)
        if pat and re.search(pat, low, re.S):
            err(idx, f"{where}: stock phrase \"not only... but also\": \"{t[:50]}\"")
        for phrase in AI_PHRASES.get(code, []):
            if re.search(r"(?<![" + LETTER + r"])" + re.escape(phrase) + r"(?![" + LETTER + r"])", low):
                err(idx, f"{where}: stock AI phrase \"{phrase}\": \"{t[:50]}\"")


# ----------------------------------------------------------------------------
# Numbers spoken as words (English and Vietnamese) for timing and on-screen checks
# ----------------------------------------------------------------------------
EN_NUM = {"zero": 0, "one": 1, "two": 2, "three": 3, "four": 4, "five": 5, "six": 6, "seven": 7, "eight": 8,
          "nine": 9, "ten": 10, "eleven": 11, "twelve": 12, "thirteen": 13, "fourteen": 14, "fifteen": 15,
          "sixteen": 16, "seventeen": 17, "eighteen": 18, "nineteen": 19, "twenty": 20, "thirty": 30,
          "forty": 40, "fifty": 50, "sixty": 60, "ninety": 90}
VI_NUM = {"không": 0, "một": 1, "hai": 2, "ba": 3, "bốn": 4, "tư": 4, "năm": 5, "sáu": 6, "bảy": 7, "tám": 8,
          "chín": 9, "mười": 10, "lăm": 5}


def _vi_tokens_to_num(tokens):
    if not tokens:
        return None
    if len(tokens) == 1 and tokens[0] in VI_NUM:
        return VI_NUM[tokens[0]]
    if len(tokens) == 2 and tokens[0] == "mười" and tokens[1] in VI_NUM:
        return 10 + VI_NUM[tokens[1]]
    if len(tokens) >= 2 and tokens[0] in VI_NUM and tokens[1] == "mươi":
        v = VI_NUM[tokens[0]] * 10
        if len(tokens) == 3:
            v += 1 if tokens[2] == "mốt" else VI_NUM.get(tokens[2], 0)
        return v
    return None


def _en_tokens_to_num(tokens):
    if not tokens:
        return None
    toks = [x for t in tokens for x in t.split("-")]
    if all(t in EN_NUM for t in toks):
        return sum(EN_NUM[t] for t in toks)
    return None


def seconds_in_text(text):
    """Every duration said or written in a text, in seconds: '60 seconds', 'two minutes', 'hai phút'."""
    out = set()
    low = str(text).lower()
    for m in re.finditer(r"(\d+(?:[,.]\d+)?)\s*(seconds?|secs?|minutes?|mins?|giây|phút)\b", low):
        v = float(m.group(1).replace(",", "."))
        out.add(round(v * (60 if m.group(2).startswith(("min", "phút")) else 1)))
    for m in re.finditer(r"((?:[a-zà-ỹ\-]+\s){1,3}?)(seconds?|minutes?|giây|phút)", low + " "):
        toks = m.group(1).split()
        for n in (3, 2, 1):
            if len(toks) < n:
                continue
            v = _en_tokens_to_num(toks[-n:])
            if v is None:
                v = _vi_tokens_to_num(toks[-n:])
            if v is not None:
                out.add(v * (60 if m.group(2).startswith(("min", "phút")) else 1))
                break
    return out


def pause_seconds_in_notes(notes):
    """'Silence 120 seconds' / 'Pause 2 minutes' / 'Lặng 40 giây' in production notes -> seconds."""
    total = 0
    for m in re.finditer(r"(?:silence|pause|hold|lặng)\s+(\d+)\s*(seconds?|secs?|minutes?|giây|phút)", str(notes), re.I):
        total += int(m.group(1)) * (60 if m.group(2).lower().startswith(("min", "phút")) else 1)
    return total


# ----------------------------------------------------------------------------
# Colour contrast: text and muted text must stay readable even when video fades them
# ----------------------------------------------------------------------------
def _lum(h):
    c = [int(h[i:i + 2], 16) / 255 for i in (0, 2, 4)]
    c = [x / 12.92 if x <= .03928 else ((x + .055) / 1.055) ** 2.4 for x in c]
    return .2126 * c[0] + .7152 * c[1] + .0722 * c[2]


def contrast(a, b):
    x, y = sorted([_lum(a), _lum(b)], reverse=True)
    return (x + .05) / (y + .05)


def check_contrast(theme, name):
    for fg in ("text", "muted"):
        for bg, need in (("bg", 7.0), ("card", 6.0)):
            r = contrast(theme[fg], theme[bg])
            if r < need:
                ERRORS.append(f"Theme '{name}': '{fg}' on '{bg}' has contrast {r:.1f}:1, below {need}:1. Lighten '{fg}'.")


# ----------------------------------------------------------------------------
# Text measurement with the real Arial metrics (or a metric-compatible substitute)
# ----------------------------------------------------------------------------
FONT_FILES = {
    False: ["arial.ttf", "Arial.ttf", "/Library/Fonts/Arial.ttf", "/System/Library/Fonts/Supplemental/Arial.ttf",
            "LiberationSans-Regular.ttf", "/usr/share/fonts/truetype/liberation/LiberationSans-Regular.ttf",
            "/usr/share/fonts/liberation/LiberationSans-Regular.ttf"],
    True: ["arialbd.ttf", "Arial Bold.ttf", "/Library/Fonts/Arial Bold.ttf",
           "/System/Library/Fonts/Supplemental/Arial Bold.ttf", "LiberationSans-Bold.ttf",
           "/usr/share/fonts/truetype/liberation/LiberationSans-Bold.ttf",
           "/usr/share/fonts/liberation/LiberationSans-Bold.ttf"],
}
_fonts = {}
_warned_font = []


def _font(pt, bold):
    key = (pt, bold)
    if key not in _fonts:
        _fonts[key] = None
        from PIL import ImageFont
        for name in FONT_FILES[bold]:
            try:
                _fonts[key] = ImageFont.truetype(name, int(pt * 4))   # measure at 4x for precision
                break
            except Exception:
                continue
        if _fonts[key] is None and not _warned_font:
            _warned_font.append(1)
            WARNINGS.append("S00: Arial (or Liberation Sans) not found: text fit is estimated, check the rendered images carefully")
    return _fonts[key]


def text_width_in(text, pt, bold=False):
    f = _font(pt, bold)
    if f is None:
        return len(text) * pt * 0.52 / 72
    return f.getlength(text) / 4 / 72


LINE_HEIGHT = 1.15   # one Arial line is about 1.15 em tall; paragraph spacing multiplies it
NBSP = " "


def _toks(para):
    """Split on ordinary whitespace; a non-breaking space keeps two words together."""
    return [w for w in re.split(r"[ \t\r\n]+", str(para)) if w]


def wrap_lines(para, width_in, pt, bold=False):
    out, cur = [], ""
    for w in _toks(para):
        trial = (cur + " " + w) if cur else w
        if text_width_in(trial, pt, bold) > width_in and cur:
            out.append(cur)
            cur = w
        else:
            cur = trial
    out.append(cur)
    return out


def count_lines(text, width_in, pt, bold=False):
    return sum(len(wrap_lines(p, width_in, pt, bold)) for p in str(text).split("\n"))


def orphan(text, width_in, pt, bold=False):
    """A paragraph that wraps and leaves one lonely word on its last line. Returns that paragraph or None."""
    for para in str(text).split("\n"):
        lines = wrap_lines(para, width_in, pt, bold)
        if len(lines) > 1 and len(_toks(lines[-1])) == 1 and NBSP not in lines[-1]:
            return para
    return None


def glue_last_two(s):
    """Join the last two words of each paragraph with a non-breaking space so the last line never holds one word."""
    out = []
    for para in str(s).split("\n"):
        tk = _toks(para)
        out.append(" ".join(tk[:-2] + [NBSP.join(tk[-2:])]) if len(tk) >= 3 else para)
    return "\n".join(out)


# Vietnamese compounds that must not be split across two lines (joined with a non-breaking space)
VI_COMPOUNDS = ["sơ đồ", "người học", "học viên", "mục tiêu", "kế hoạch", "bảng kiểm", "thành phần", "học liệu",
                "bài tập", "đồ thị", "hàm số", "hai mươi", "ba mươi", "bốn mươi", "năm mươi", "chín mươi"]
_RE_VI_COMPOUNDS = re.compile("|".join(re.escape(x) for x in sorted(VI_COMPOUNDS, key=len, reverse=True)), re.I)


def keep_compounds(s, lang):
    if lang != "vi" or not isinstance(s, str):
        return s
    return _RE_VI_COMPOUNDS.sub(lambda m: m.group(0).replace(" ", NBSP), s)


def fits(text, width_in, height_in, pt, spacing=1.2, bold=False):
    return count_lines(text, width_in, pt, bold) * pt * LINE_HEIGHT * spacing / 72 <= height_in + 0.02


def fit_size(text, width_in, height_in, pt, min_pt, spacing=1.2, bold=False):
    size = pt
    while size >= min_pt:
        if fits(text, width_in, height_in, size, spacing, bold):
            return size
        size -= 2
    return None


def words(text):
    return len(re.findall(r"\S+", str(text)))


# ----------------------------------------------------------------------------
# Lucide icons -> coloured PNG (cairosvg, sharp-cli, rsvg-convert or Inkscape), cached per deck
# ----------------------------------------------------------------------------
_icon_cache = {}


def _rasterize(svg_path, png_path):
    try:
        import cairosvg
        cairosvg.svg2png(url=svg_path, write_to=png_path, output_width=512, output_height=512)
        return True
    except Exception:
        pass
    for exe, args in (("sharp", ["-i", svg_path, "-o", png_path, "resize", "512", "512"]),
                      ("sharp.cmd", ["-i", svg_path, "-o", png_path, "resize", "512", "512"]),
                      ("rsvg-convert", ["-w", "512", "-h", "512", "-o", png_path, svg_path]),
                      ("inkscape", [svg_path, "--export-type=png", f"--export-filename={png_path}", "-w", "512", "-h", "512"])):
        path = shutil.which(exe)
        if path:
            r = subprocess.run([path] + args, capture_output=True)
            if r.returncode == 0 and os.path.isfile(png_path):
                return True
    return False


def icon_png(name, color, out_dir):
    key = (name, color)
    if key in _icon_cache:
        return _icon_cache[key]
    svg = os.path.join(ICON_DIR, f"{name}.svg")
    if not os.path.isfile(svg):
        return None
    cache_dir = os.path.join(out_dir, "_icons")
    os.makedirs(cache_dir, exist_ok=True)
    png = os.path.join(cache_dir, f"{name}-{color}.png")
    if not os.path.isfile(png):
        with open(svg, encoding="utf-8") as f:
            data = f.read().replace("currentColor", f"#{color}")
        tmp_svg = os.path.join(cache_dir, f"{name}-{color}.svg")
        with open(tmp_svg, "w", encoding="utf-8") as f:
            f.write(data)
        ok = _rasterize(tmp_svg, png)
        os.remove(tmp_svg)
        if not ok:
            raise SystemExit("No SVG rasterizer found. Install one: pip install cairosvg  (or npm i -g sharp-cli, "
                             "or rsvg-convert / Inkscape).")
    _icon_cache[key] = png
    return png


# ----------------------------------------------------------------------------
# The deck
# ----------------------------------------------------------------------------
class Deck:
    def __init__(self, spec, out_dir):
        self.spec = spec
        self.out_dir = out_dir
        self.lang = spec.get("language", "en")
        self.L = DEFAULT_TEXT.get(self.lang, DEFAULT_TEXT["en"])
        self.theme_name = spec.get("theme", "navy")
        self.theme = THEMES[self.theme_name]
        check_contrast(self.theme, self.theme_name)
        self.upper = spec.get("title_case", "upper") == "upper"
        self.prs = Presentation()
        self.prs.slide_width = Inches(W)
        self.prs.slide_height = Inches(H)
        self.blank = self.prs.slide_layouts[6]
        self.n_total = len(spec["slides"])
        self.idx = 0
        self.blocks = []
        self._said = ""

    def T(self, s):
        """Titles, course name and lesson name in capitals when title_case is 'upper' (the JSON stays as written)."""
        return str(s).upper() if self.upper else str(s)

    # --- primitives ----------------------------------------------------------
    def new_slide(self, bg=None):
        s = self.prs.slides.add_slide(self.blank)
        fill = s.background.fill
        fill.solid()
        fill.fore_color.rgb = rgb(bg or self.theme["bg"])
        self.idx += 1
        return s

    def rect(self, s, x, y, w, h, color, rounded=False, radius=0.08):
        shp = s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE if rounded else MSO_SHAPE.RECTANGLE,
                                 Inches(x), Inches(y), Inches(w), Inches(h))
        shp.fill.solid()
        shp.fill.fore_color.rgb = rgb(color)
        shp.line.fill.background()
        shp.shadow.inherit = False
        if rounded:
            shp.adjustments[0] = min(0.5, radius / min(w, h))
        return shp

    def circle(self, s, x, y, d, color):
        shp = s.shapes.add_shape(MSO_SHAPE.OVAL, Inches(x), Inches(y), Inches(d), Inches(d))
        shp.fill.solid()
        shp.fill.fore_color.rgb = rgb(color)
        shp.line.fill.background()
        shp.shadow.inherit = False
        return shp

    def text(self, s, x, y, w, h, content, size, color=None, bold=False, align="l", anchor="t", italic=False,
             bullets=False, spacing=1.15, min_size=None, para_after=0):
        """content: a string (new lines with \\n) or a list of lines (bullets=True)."""
        lines = content if isinstance(content, list) else str(content).split("\n")
        lines = [keep_compounds(x, self.lang) for x in lines]
        joined = "\n".join(lines)
        if min_size is not None:
            box_w = w
            w = w - (size * 1.1 / 72 if bullets else 0)   # a bullet's hanging indent narrows the text column
            ok = fit_size(joined, w, h, size, min_size, spacing, bold)
            if ok is None:
                err(self.idx, f"Text does not fit a {w:.1f} x {h:.1f} in box even at {min_size} pt. Shorten it or split the slide: \"{joined[:60]}...\"")
                ok = min_size
            size = ok
            bad = orphan(joined, w, size, bold)
            if bad:   # standard typographic fix first: pull the last two words onto one line
                glued = glue_last_two(joined)
                if fits(glued, w, h, size, spacing, bold) and not orphan(glued, w, size, bold):
                    lines = glued.split("\n")
                    joined, bad = glued, None
            if bad:
                err(self.idx, f"Orphan word: one word alone on the last line at {size} pt in a {w:.1f} in box: \"{bad[:60]}\". Rewrite to fit one line or balance two lines")
            w = box_w
        tb = s.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
        tf = tb.text_frame
        tf.word_wrap = True
        tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = 0
        tf.vertical_anchor = {"t": MSO_ANCHOR.TOP, "m": MSO_ANCHOR.MIDDLE, "b": MSO_ANCHOR.BOTTOM}[anchor]
        for i, line in enumerate(lines):
            p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
            p.alignment = {"l": PP_ALIGN.LEFT, "c": PP_ALIGN.CENTER, "r": PP_ALIGN.RIGHT}[align]
            p.line_spacing = spacing
            if para_after:
                p.space_after = Pt(para_after)
            run = p.add_run()
            run.text = line
            f = run.font
            f.name = FONT
            f.size = Pt(size)
            f.bold = bold
            f.italic = italic
            f.color.rgb = rgb(color or self.theme["text"])
            if bullets:
                pPr = p._p.get_or_add_pPr()
                indent = int(Pt(size * 1.1))
                pPr.set("marL", str(indent))
                pPr.set("indent", str(-indent))
                buclr = pPr.makeelement(qn("a:buClr"), {})
                buclr.append(buclr.makeelement(qn("a:srgbClr"), {"val": self.theme["accent"]}))
                pPr.append(buclr)
                pPr.append(pPr.makeelement(qn("a:buChar"), {"char": "•"}))
        return tb

    def icon(self, s, name, x, y, size, color=None, bg=None):
        png = icon_png(name, color or self.theme["accent"], self.out_dir)
        if not png:
            err(self.idx, f"No icon \"{name}\" in assets/icons (see the list in references/slide-types.md)")
            return
        if bg:
            self.circle(s, x, y, size, bg)
            pad = size * 0.22
            s.shapes.add_picture(png, Inches(x + pad), Inches(y + pad), Inches(size - 2 * pad), Inches(size - 2 * pad))
        else:
            s.shapes.add_picture(png, Inches(x), Inches(y), Inches(size), Inches(size))

    def picture(self, s, path, x, y, w, h):
        if not os.path.isabs(path):
            path = os.path.join(os.path.dirname(self.spec["_path"]), path)
        if not os.path.isfile(path):
            err(self.idx, f"Image not found: {path}")
            return
        with Image.open(path) as im:
            iw, ih = im.size
        scale = min(w / iw, h / ih)
        pw, ph = iw * scale, ih * scale
        s.shapes.add_picture(path, Inches(x + (w - pw) / 2), Inches(y + (h - ph) / 2), Inches(pw), Inches(ph))

    # --- shared frame ----------------------------------------------------------
    def frame(self, s, title=None, label=None):
        t = self.theme
        y = 0.32
        if label:
            self.text(s, MARGIN, y, W - 2 * MARGIN, 0.42, label, SZ_CAPTION, color=t["accent"], bold=True)
            y += 0.5
        else:
            y = 0.45
        if title:
            title = self.T(title)
            size = fit_size(title, W - 2 * MARGIN, 1.0, SZ_TITLE, 28, bold=True)
            if size is None:
                err(self.idx, f"Title too long for one line: \"{title}\"")
                size = 28
            if orphan(title, W - 2 * MARGIN, size, True):
                title = glue_last_two(title)
                if orphan(title, W - 2 * MARGIN, size, True) or not fits(title, W - 2 * MARGIN, 1.0, size, 1.15, True):
                    err(self.idx, f"Orphan word in the title: \"{title}\"")
            self.text(s, MARGIN, y, W - 2 * MARGIN, 1.0, title, size, bold=True, anchor="t")
        self.footer(s)

    def credit(self):
        return " · ".join(v for v in [self.spec.get("instructor"), self.spec.get("contact")] if v)

    def footer(self, s, color=None, bar=None):
        t = self.theme
        self.text(s, MARGIN, H - 0.58, 8, 0.45, self.credit(), SZ_FOOT, color=color or t["muted"])
        self.text(s, W - MARGIN - 1.5, H - 0.58, 1.5, 0.45, f"{self.idx}/{self.n_total}", SZ_FOOT,
                  color=color or t["muted"], align="r")
        self.rect(s, 0, H - 0.06, W * self.idx / self.n_total, 0.06, bar or t["accent"])   # progress bar

    def content_top(self, has_label=True):
        return 1.95 if has_label else 1.6

    def desc_color(self):
        t = self.theme
        return t["muted"] if t["bg"] != "FFFFFF" else t["text"]

    # --- slide types ------------------------------------------------------------
    def s_title(self, d):
        t, sp = self.theme, self.spec
        s = self.new_slide()
        self.text(s, MARGIN, 0.55, W - 2 * MARGIN, 0.5, sp.get("organization", ""), SZ_CAPTION, color=t["muted"], align="c")
        self.text(s, MARGIN, 1.05, W - 2 * MARGIN, 0.5, self.T(sp.get("course", "")), SZ_CAPTION, align="c")
        self.text(s, MARGIN, 2.0, W - 2 * MARGIN, 0.5, d.get("label", sp.get("lesson_code", "")), SZ_CAPTION,
                  color=t["accent"], bold=True, align="c")
        title = self.T(d.get("title", sp.get("lesson_title", "")))
        self.text(s, 1.2, 2.55, W - 2.4, 1.9, title, 44, bold=True, align="c", anchor="m", min_size=32)
        if d.get("image"):
            self.picture(s, d["image"], W / 2 - 1.3, 4.55, 2.6, 1.6)
        else:
            self.icon(s, d.get("icon", "graduation-cap"), W / 2 - 0.7, 4.6, 1.4, color=t["bg"], bg=t["accent"])
        self.text(s, MARGIN, 6.3, W - 2 * MARGIN, 0.5, self.credit(), SZ_CAPTION, align="c")
        return "centered", [title]

    def s_objectives(self, d):
        t = self.theme
        s = self.new_slide()
        title = d.get("title", self.L["objectives_title"])
        self.frame(s, title, d.get("label", self.L["objectives_label"]))
        items = d["items"]
        self.check_list(items)
        icons = d.get("icons") or ["check-circle"] * len(items)
        y0, avail = self.content_top(), 6.75 - self.content_top()
        gap = 0.18
        hh = (avail - gap * (len(items) - 1)) / len(items)
        for i, it in enumerate(items):
            y = y0 + i * (hh + gap)
            self.rect(s, MARGIN, y, W - 2 * MARGIN, hh, t["card"], rounded=True, radius=0.15)
            ic_d = min(0.75, hh - 0.3)
            self.icon(s, icons[i % len(icons)], MARGIN + 0.3, y + (hh - ic_d) / 2, ic_d, color=t["bg"], bg=t["accent"])
            self.text(s, MARGIN + 1.35, y + 0.12, W - 2 * MARGIN - 1.65, hh - 0.24, it, SZ_BODY, anchor="m", min_size=SZ_BODY)
        return "vertical grid", [title] + items

    def s_agenda(self, d):
        t = self.theme
        s = self.new_slide()
        title = d.get("title", self.L["agenda_title"])
        self.frame(s, title, d.get("label", self.spec.get("lesson_code", "")))
        items = d["items"]
        self.check_list(items)
        cur = d.get("current")
        first = int(d.get("start", 1))
        y0, avail = self.content_top(), 6.75 - self.content_top()
        gap = 0.16
        hh = (avail - gap * (len(items) - 1)) / len(items)
        for i, it in enumerate(items):
            y = y0 + i * (hh + gap)
            active = (cur == i + 1)
            self.rect(s, MARGIN, y, W - 2 * MARGIN, hh, t["accent"] if active else t["card"], rounded=True, radius=0.15)
            self.text(s, MARGIN + 0.3, y, 1.1, hh, f"{first + i:02d}", 30, color=t["bg"] if active else t["accent"], bold=True, anchor="m")
            self.text(s, MARGIN + 1.4, y + 0.1, W - 2 * MARGIN - 1.7, hh - 0.2, it, SZ_BODY,
                      color=t["bg"] if active else t["text"], bold=active, anchor="m", min_size=SZ_BODY)
        return "vertical grid", [title] + [f"{first + i}. {x}" for i, x in enumerate(items)]

    def s_section(self, d):
        t = self.theme
        s = self.new_slide()
        no = str(d.get("number", "")).zfill(2) if d.get("number") is not None else ""
        if no:   # big number, last digit and the dot in the accent colour
            tb = self.text(s, MARGIN, 1.6, 4, 2.4, no[:-1], 150, bold=True, anchor="m")
            r = tb.text_frame.paragraphs[0].add_run()
            r.text = no[-1] + "."
            r.font.name, r.font.size, r.font.bold = FONT, Pt(150), True
            r.font.color.rgb = rgb(t["accent"])
        self.text(s, MARGIN, 4.05, W - 2 * MARGIN, 0.45, d.get("label", self.L["section_label"]), SZ_CAPTION, color=t["accent"], bold=True)
        self.text(s, MARGIN, 4.5, W - 2 * MARGIN, 1.2, self.T(d["title"]), 40, bold=True, anchor="t", min_size=30)
        if d.get("subtitle"):
            self.text(s, MARGIN, 5.75, W - 2 * MARGIN, 0.9, d["subtitle"], SZ_BODY, color=t["muted"], min_size=SZ_BODY)
        if d.get("icon"):
            self.icon(s, d["icon"], W - MARGIN - 2.2, 1.7, 2.2, color=t["bg"], bg=t["accent"])
        self.footer(s)
        return "centered", [no, d["title"], d.get("subtitle", "")]

    def s_cards(self, d):
        t = self.theme
        s = self.new_slide()
        self.frame(s, d["title"], d.get("label"))
        cards = d["cards"]
        n = len(cards)
        if not 2 <= n <= 6:
            err(self.idx, f"'cards' needs 2 to 6 cards, got {n}")
        cols = n if n <= 3 else (2 if n == 4 else 3)
        rows = math.ceil(n / cols)
        gap = 0.25
        top = self.content_top(bool(d.get("label")))
        cw = (W - 2 * MARGIN - gap * (cols - 1)) / cols
        ch = (6.7 - top - gap * (rows - 1)) / rows
        stacked = rows == 2 and cols == 3   # 5-6 cards: icon on top, one-line card title
        for i, c in enumerate(cards):
            r, col = divmod(i, cols)
            x, y = MARGIN + col * (cw + gap), top + r * (ch + gap)
            self.rect(s, x, y, cw, ch, t["card"], rounded=True, radius=0.18)
            if stacked:
                ic = 0.5
                self.icon(s, c.get("icon", "star"), x + 0.25, y + 0.15, ic, color=t["bg"], bg=t["accent"])
                self.text(s, x + 0.25, y + 0.7, cw - 0.5, 0.42, c["title"], SZ_BODY, bold=True, anchor="m", min_size=SZ_BODY, spacing=1.0)
                if c.get("text"):
                    self.text(s, x + 0.25, y + 1.15, cw - 0.5, ch - 1.3, c["text"], SZ_BODY, color=self.desc_color(), min_size=SZ_BODY, spacing=1.05)
                continue
            ic = 0.6
            self.icon(s, c.get("icon", "star"), x + 0.25, y + 0.25, ic, color=t["bg"], bg=t["accent"])
            self.text(s, x + 0.25 + ic + 0.2, y + 0.15, cw - ic - 0.7, 0.85, c["title"], SZ_BODY, bold=True, anchor="m", min_size=SZ_BODY, spacing=1.1)
            if c.get("text"):
                self.text(s, x + 0.25, y + 1.15, cw - 0.5, ch - 1.35, c["text"], SZ_BODY, color=self.desc_color(), min_size=SZ_BODY, spacing=1.1)
        return "grid", [d["title"]] + [f"{c['title']}: {c.get('text', '')}" for c in cards]

    def s_two_column(self, d):
        t = self.theme
        s = self.new_slide()
        self.frame(s, d["title"], d.get("label"))
        tw = 6.6
        items = [keep_compounds(x, self.lang) for x in d["items"]]
        self.check_list(items)
        items = [glue_last_two(x) if (orphan(x, tw - 0.5, SZ_BODY) or orphan(x, tw, SZ_BODY)) else x for x in items]
        for it in items:
            if orphan(it, tw - 0.5, SZ_BODY):
                err(self.idx, f"Orphan word in a bullet: \"{it}\". Rewrite to fit one line or balance two lines")
        top = self.content_top(bool(d.get("label")))
        left_visual = d.get("visual_side", "right") == "left"
        tx = MARGIN + (W - 2 * MARGIN - tw) if left_visual else MARGIN
        gx = MARGIN if left_visual else MARGIN + tw + 0.4
        gw = W - 2 * MARGIN - tw - 0.4
        gh = 6.7 - top
        self.text(s, tx, top, tw, gh, items, SZ_BODY, bullets=True, min_size=SZ_BODY, para_after=10, spacing=1.2)
        if d.get("image"):
            self.picture(s, d["image"], gx, top, gw, gh)
            if d.get("caption"):
                self.text(s, gx, 6.75 - 0.45, gw, 0.45, d["caption"], SZ_CAPTION, color=t["muted"], italic=True, align="c")
        else:
            self.rect(s, gx, top, gw, gh, t["card"], rounded=True, radius=0.2)
            di = min(gw, gh) * 0.55
            self.icon(s, d.get("icon", "lightbulb"), gx + (gw - di) / 2, top + (gh - di) / 2, di, color=t["bg"], bg=t["accent"])
        return "half and half", [d["title"]] + items

    def s_process(self, d):
        t = self.theme
        s = self.new_slide()
        self.frame(s, d["title"], d.get("label"))
        steps = d["steps"]
        n = len(steps)
        if not 2 <= n <= 6:
            err(self.idx, f"'process' needs 2 to 6 steps, got {n}")
        top = self.content_top(bool(d.get("label")))
        if n <= 4:   # horizontal: step boxes with arrows
            gap = 0.45
            cw = (W - 2 * MARGIN - gap * (n - 1)) / n
            ch = 6.7 - top
            for i, st in enumerate(steps):
                x = MARGIN + i * (cw + gap)
                self.rect(s, x, top, cw, ch, t["card"], rounded=True, radius=0.18)
                self.circle(s, x + 0.25, top + 0.25, 0.7, t["accent"])
                self.text(s, x + 0.25, top + 0.25, 0.7, 0.7, str(i + 1), 26, bold=True, color=t["bg"], align="c", anchor="m")
                if st.get("icon"):
                    self.icon(s, st["icon"], x + cw - 0.95, top + 0.3, 0.6)
                self.text(s, x + 0.25, top + 1.15, cw - 0.5, 1.3, st["title"], SZ_BODY, bold=True, min_size=SZ_BODY, spacing=1.1)
                if st.get("text"):
                    self.text(s, x + 0.25, top + 2.5, cw - 0.5, ch - 2.75, st["text"], SZ_BODY, color=self.desc_color(), min_size=SZ_BODY, spacing=1.1)
                if i < n - 1:
                    self.icon(s, "arrow-right", x + cw + gap / 2 - 0.18, top + ch / 2 - 0.18, 0.36)
            layout = "horizontal grid"
        else:        # vertical timeline: number, title and description on one row
            gap = 0.12
            rh = (6.7 - top - gap * (n - 1)) / n
            self.rect(s, MARGIN + 0.33, top + 0.3, 0.06, 6.7 - top - 0.6, t["card"])
            for i, st in enumerate(steps):
                y = top + i * (rh + gap)
                dc = min(0.62, rh - 0.1)
                self.circle(s, MARGIN + 0.36 - dc / 2, y + (rh - dc) / 2, dc, t["accent"])
                self.text(s, MARGIN + 0.36 - dc / 2, y + (rh - dc) / 2, dc, dc, str(i + 1), SZ_BODY, bold=True, color=t["bg"], align="c", anchor="m")
                self.rect(s, MARGIN + 0.95, y, W - 2 * MARGIN - 0.95, rh, t["card"], rounded=True, radius=0.12)
                tw = 4.6
                self.text(s, MARGIN + 1.2, y, tw, rh, st["title"], SZ_BODY, bold=True, anchor="m", min_size=SZ_BODY, spacing=1.05)
                if st.get("text"):
                    self.text(s, MARGIN + 1.4 + tw, y, W - 2 * MARGIN - 1.85 - tw, rh, st["text"], SZ_BODY,
                              color=self.desc_color(), anchor="m", min_size=SZ_BODY, spacing=1.05)
            layout = "vertical grid"
        return layout, [d["title"]] + [f"{self.L['step']} {i + 1}: {b['title']}" for i, b in enumerate(steps)]

    def s_compare(self, d):
        t = self.theme
        s = self.new_slide()
        self.frame(s, d["title"], d.get("label"))
        top = self.content_top(bool(d.get("label")))
        gap = 0.4
        cw = (W - 2 * MARGIN - gap) / 2
        ch = 6.7 - top
        for i, side in enumerate([d["left"], d["right"]]):
            x = MARGIN + i * (cw + gap)
            self.rect(s, x, top, cw, ch, t["card"], rounded=True, radius=0.18)
            self.icon(s, side.get("icon", "x-circle" if i == 0 else "check-circle"), x + 0.3, top + 0.3, 0.6, color=t["bg"], bg=t["accent"])
            self.text(s, x + 1.05, top + 0.3, cw - 1.35, 0.6, side["title"], 26, bold=True, anchor="m", color=t["accent"])
            self.check_list(side["items"])
            self.text(s, x + 0.3, top + 1.1, cw - 0.6, ch - 1.35, side["items"], SZ_BODY, bullets=True, min_size=SZ_BODY, para_after=8)
        return "half and half", [d["title"], f"{d['left']['title']}: " + "; ".join(d["left"]["items"]),
                                 f"{d['right']['title']}: " + "; ".join(d["right"]["items"])]

    def s_big_numbers(self, d):
        t = self.theme
        s = self.new_slide()
        self.frame(s, d["title"], d.get("label"))
        nums = d["numbers"]
        n = len(nums)
        top = self.content_top(bool(d.get("label")))
        gap = 0.4
        cw = (W - 2 * MARGIN - gap * (n - 1)) / n
        ch = 6.7 - top
        for i, it in enumerate(nums):
            x = MARGIN + i * (cw + gap)
            self.rect(s, x, top, cw, ch, t["card"], rounded=True, radius=0.2)
            if it.get("icon"):
                self.icon(s, it["icon"], x + cw / 2 - 0.35, top + 0.35, 0.7)
            self.text(s, x + 0.2, top + ch * 0.28, cw - 0.4, ch * 0.4, str(it["value"]), 72, bold=True, color=t["accent"], align="c", anchor="m", min_size=48)
            self.text(s, x + 0.3, top + ch * 0.7, cw - 0.6, ch * 0.28, it["label"], SZ_BODY, align="c", anchor="t", min_size=SZ_BODY)
        return "horizontal grid", [d["title"]] + [f"{x['value']} {x['label']}" for x in nums]

    def s_question(self, d):
        t = self.theme
        s = self.new_slide(bg=t["accent"])   # a different background so learners spot the pause at once
        self.text(s, MARGIN, 0.5, W - 2 * MARGIN, 0.42, d.get("label", self.L["question_label"]), SZ_CAPTION, color=t["bg"], bold=True)
        self.icon(s, d.get("icon", "pause-circle"), W - MARGIN - 1.0, 0.45, 1.0, color=t["accent"], bg=t["bg"])
        self.text(s, MARGIN, 1.1, W - 2.6, 1.7, d["question"], 34, bold=True, color=t["bg"], anchor="t", min_size=26)
        opts = d.get("options", [])
        y = 3.0
        if opts:
            self.check_list(opts)
            hh = min(0.8, (6.1 - y - 0.15 * (len(opts) - 1)) / len(opts))
            for i, o in enumerate(opts):
                yy = y + i * (hh + 0.15)
                self.rect(s, MARGIN, yy, W - 2 * MARGIN, hh, t["bg"], rounded=True, radius=0.15)
                self.text(s, MARGIN + 0.3, yy, 0.7, hh, chr(65 + i), SZ_BODY, bold=True, color=t["accent"], anchor="m")
                self.text(s, MARGIN + 1.0, yy + 0.08, W - 2 * MARGIN - 1.3, hh - 0.16, o, SZ_BODY, anchor="m", min_size=SZ_BODY)
        elif d.get("instructions"):
            lines = count_lines(d["instructions"], W - 2 * MARGIN - 0.8, SZ_BODY)
            bh = min(3.1, max(1.2, lines * SZ_BODY * LINE_HEIGHT * 1.2 / 72 + 0.6))
            self.rect(s, MARGIN, y, W - 2 * MARGIN, bh, t["bg"], rounded=True, radius=0.2)
            self.text(s, MARGIN + 0.4, y + 0.3, W - 2 * MARGIN - 0.8, bh - 0.6, d["instructions"], SZ_BODY, min_size=SZ_BODY)
        if d.get("timer"):
            self.icon(s, "timer", MARGIN, 6.35, 0.36, color=t["bg"])
            self.text(s, MARGIN + 0.5, 6.3, 6, 0.42, d["timer"], SZ_CAPTION, color=t["bg"], bold=True, anchor="m")
        self.footer(s, color=t["bg"], bar=t["bg"])
        return "centered", [d["question"]] + opts + ([d["instructions"]] if d.get("instructions") else [])

    def s_key_message(self, d):
        t = self.theme
        s = self.new_slide()
        self.icon(s, d.get("icon", "quote"), W / 2 - 0.55, 1.2, 1.1, color=t["bg"], bg=t["accent"])
        if d.get("label"):
            self.text(s, MARGIN, 2.4, W - 2 * MARGIN, 0.45, d["label"], SZ_CAPTION, color=t["accent"], bold=True, align="c")
        self.text(s, 1.2, 2.95, W - 2.4, 2.6, d["message"], 40, bold=True, align="c", anchor="m", min_size=28)
        if d.get("subtitle"):
            self.text(s, 1.5, 5.65, W - 3, 0.9, d["subtitle"], SZ_BODY, color=t["muted"], align="c", min_size=SZ_BODY)
        self.footer(s)
        return "centered", [d["message"], d.get("subtitle", "")]

    def s_image(self, d):
        t = self.theme
        s = self.new_slide()
        self.frame(s, d["title"], d.get("label"))
        top = self.content_top(bool(d.get("label")))
        gh = 6.7 - top - (0.5 if d.get("caption") else 0)
        self.picture(s, d["image"], MARGIN, top, W - 2 * MARGIN, gh)
        if d.get("caption"):
            self.text(s, MARGIN, 6.75 - 0.45, W - 2 * MARGIN, 0.45, d["caption"], SZ_CAPTION, color=t["muted"], italic=True, align="c")
        return "centered", [d["title"], d.get("caption", "")]

    def s_table(self, d):
        t = self.theme
        s = self.new_slide()
        self.frame(s, d["title"], d.get("label"))
        cols, rows = d["columns"], d["rows"]
        if len(rows) > 5:
            err(self.idx, f"A table holds at most 5 data rows, got {len(rows)}. Use 'cards' or split the slide")
        top = self.content_top(bool(d.get("label")))
        nrow, ncol = len(rows) + 1, len(cols)
        tw, th = W - 2 * MARGIN, 6.7 - top
        tbl = s.shapes.add_table(nrow, ncol, Inches(MARGIN), Inches(top), Inches(tw), Inches(th)).table
        widths = d.get("widths") or [1] * ncol
        tot = sum(widths)
        for j, wgt in enumerate(widths):
            tbl.columns[j].width = Inches(tw * wgt / tot)
        rh = th / nrow
        for i in range(nrow):
            tbl.rows[i].height = Inches(rh)
        for i in range(nrow):
            for j in range(ncol):
                cell = tbl.cell(i, j)
                val = keep_compounds(str(cols[j] if i == 0 else rows[i - 1][j]), self.lang)
                cell.fill.solid()
                cell.fill.fore_color.rgb = rgb(t["accent"] if i == 0 else t["card"])
                cell.margin_left = cell.margin_right = Inches(0.15)
                cell.margin_top = cell.margin_bottom = Inches(0.08)
                cell.vertical_anchor = MSO_ANCHOR.MIDDLE
                tf = cell.text_frame
                tf.word_wrap = True
                r = tf.paragraphs[0].add_run()
                cw_in = tw * widths[j] / tot - 0.3
                if orphan(val, cw_in, SZ_BODY, i == 0):
                    val = glue_last_two(val)
                    if orphan(val, cw_in, SZ_BODY, i == 0):
                        err(self.idx, f"Orphan word in a table cell: \"{val}\". Rewrite or change the column widths")
                r.text = val
                r.font.name, r.font.bold = FONT, (i == 0)
                if fit_size(val, cw_in, rh - 0.16, SZ_BODY, SZ_BODY) is None:
                    err(self.idx, f"Table cell does not fit at 24 pt. Shorten it or drop a column: \"{val}\"")
                r.font.size = Pt(SZ_BODY)
                r.font.color.rgb = rgb(t["bg"] if i == 0 else t["text"])
        # The default table style draws WHITE borders, which show up as empty frames when a video reveals
        # rows one by one. Redraw all four borders of every cell in the background colour.
        from lxml import etree
        for i in range(nrow):
            for j in range(ncol):
                tcPr = tbl.cell(i, j)._tc.get_or_add_tcPr()
                for tag in ("a:lnL", "a:lnR", "a:lnT", "a:lnB"):
                    for old in tcPr.findall(qn(tag)):
                        tcPr.remove(old)
                lns = []
                for tag in ("a:lnL", "a:lnR", "a:lnT", "a:lnB"):
                    ln = etree.SubElement(tcPr, qn(tag), w="25400", cap="flat", cmpd="sng", algn="ctr")
                    sf = etree.SubElement(ln, qn("a:solidFill"))
                    etree.SubElement(sf, qn("a:srgbClr"), val=t["bg"])
                    etree.SubElement(ln, qn("a:prstDash"), val="solid")
                    lns.append(ln)
                for ln in reversed(lns):   # borders must come before the fill (schema order)
                    tcPr.remove(ln)
                    tcPr.insert(0, ln)
        tbl.first_row = True
        tbl.horz_banding = False
        return "grid", [d["title"]] + [" | ".join(map(str, r)) for r in [cols] + rows]

    def s_summary(self, d):
        t = self.theme
        s = self.new_slide()
        self.frame(s, d.get("title", self.L["summary_title"]), d.get("label", self.L["summary_label"]))
        kws = d["keywords"]
        top = self.content_top()
        n = len(kws)
        gap = 0.35
        cw = (W - 2 * MARGIN - gap * (n - 1)) / n
        ch = 2.0 if d.get("items") else 6.7 - top
        for i, k in enumerate(kws):
            x = MARGIN + i * (cw + gap)
            self.rect(s, x, top, cw, ch, t["accent"], rounded=True, radius=0.2)
            self.text(s, x + 0.2, top, cw - 0.4, ch, k, 32, bold=True, color=t["bg"], align="c", anchor="m", min_size=24)
        if d.get("items"):
            self.check_list(d["items"])
            self.text(s, MARGIN, top + ch + 0.35, W - 2 * MARGIN, 6.7 - top - ch - 0.35, d["items"], SZ_BODY, bullets=True, min_size=SZ_BODY, para_after=8)
        return "horizontal grid", kws + d.get("items", [])

    def s_closing(self, d):
        t = self.theme
        s = self.new_slide()
        self.icon(s, d.get("icon", "rocket"), W / 2 - 0.65, 0.9, 1.3, color=t["bg"], bg=t["accent"])
        self.text(s, MARGIN, 2.4, W - 2 * MARGIN, 0.45, d.get("label", self.L["closing_label"]), SZ_CAPTION, color=t["accent"], bold=True, align="c")
        self.text(s, 1.2, 2.9, W - 2.4, 2.0, d["action"], 34, bold=True, align="c", anchor="m", min_size=26)
        if d.get("next_lesson"):
            self.rect(s, 2.0, 5.1, W - 4, 1.1, t["card"], rounded=True, radius=0.2)
            self.text(s, 2.2, 5.1, W - 4.4, 1.1, self.L["next_lesson"] + d["next_lesson"], SZ_BODY, align="c", anchor="m", min_size=SZ_BODY)
        self.footer(s)
        return "centered", [d["action"], d.get("next_lesson", "")]

    def s_chart(self, d):
        """A native PowerPoint chart (editable) and one takeaway sentence beside it."""
        from pptx.chart.data import CategoryChartData
        from pptx.enum.chart import XL_CHART_TYPE, XL_LABEL_POSITION
        t = self.theme
        s = self.new_slide()
        self.frame(s, d["title"], d.get("label"))
        top = self.content_top(bool(d.get("label")))
        cats, vals = d["categories"], d["values"]
        if not 2 <= len(cats) <= 6 or len(cats) != len(vals):
            err(self.idx, "'chart' needs 2 to 6 categories and one value per category")
        kind_name = d.get("chart_type", "column")
        kind = {"column": XL_CHART_TYPE.COLUMN_CLUSTERED, "bar": XL_CHART_TYPE.BAR_CLUSTERED,
                "pie": XL_CHART_TYPE.PIE, "line": XL_CHART_TYPE.LINE_MARKERS}.get(kind_name)
        if kind is None:
            err(self.idx, f"chart_type must be column, bar, pie or line (got {kind_name})")
            kind = XL_CHART_TYPE.COLUMN_CLUSTERED
        cd = CategoryChartData()
        cd.categories = cats
        cd.add_series(d.get("series_name", d["title"]), vals)
        cw = 7.6
        ch = s.shapes.add_chart(kind, Inches(MARGIN), Inches(top), Inches(cw), Inches(6.7 - top), cd).chart
        ch.has_legend = False
        ch.has_title = False
        ch.font.name = FONT
        ch.font.size = Pt(SZ_CAPTION)
        ch.font.color.rgb = rgb(t["text"])
        plot = ch.plots[0]
        plot.has_data_labels = True
        dl = plot.data_labels
        dl.font.size = Pt(SZ_CAPTION)
        dl.font.bold = True
        dl.font.color.rgb = rgb(t["text"])
        unit = d.get("unit", "")
        dl.number_format = '0"' + unit + '"' if unit else "0"
        dl.number_format_is_linked = False
        hi = d.get("highlight")
        ser = plot.series[0]
        if kind_name == "pie":
            dl.position = XL_LABEL_POSITION.OUTSIDE_END
            dl.show_category_name = True   # category name and value on the label, no legend
            dl.show_value = True
            dl.show_percentage = False
            from lxml import etree
            dLbls = dl._element
            for sep in dLbls.findall(qn("c:separator")):
                dLbls.remove(sep)
            sep = etree.SubElement(dLbls, qn("c:separator"))
            sep.text = "\n"
            bubble = dLbls.find(qn("c:showBubbleSize"))
            if bubble is not None:
                dLbls.remove(sep)
                bubble.addnext(sep)
            shades = [t["accent"], t["muted"], t["text"]]   # 'card' is too close to the background
            for i in range(len(cats)):
                p = ser.points[i]
                p.format.fill.solid()
                p.format.fill.fore_color.rgb = rgb(shades[i % 3])
        else:
            dl.position = XL_LABEL_POSITION.OUTSIDE_END if kind_name != "line" else XL_LABEL_POSITION.ABOVE
            ser.format.fill.solid()
            ser.format.fill.fore_color.rgb = rgb(t["muted"] if hi else t["accent"])
            if kind_name == "line":
                ser.format.line.color.rgb = rgb(t["accent"])
                ser.format.line.width = Pt(4)
                ser.smooth = False
            if hi:
                p = ser.points[int(hi) - 1]
                p.format.fill.solid()
                p.format.fill.fore_color.rgb = rgb(t["accent"])
            try:
                plot.gap_width = 60
            except Exception:
                pass
            ca, va = ch.category_axis, ch.value_axis
            ca.tick_labels.font.size = Pt(SZ_CAPTION)
            ca.tick_labels.font.color.rgb = rgb(t["text"])
            ca.format.line.color.rgb = rgb(t["muted"])
            ca.has_major_gridlines = False
            va.has_major_gridlines = False
            va.visible = False
        rx = MARGIN + cw + 0.4
        rw = W - MARGIN - rx
        self.rect(s, rx, top, rw, 6.7 - top, t["card"], rounded=True, radius=0.2)
        if d.get("icon"):
            self.icon(s, d["icon"], rx + 0.3, top + 0.3, 0.7, color=t["bg"], bg=t["accent"])
        src_h = 1.1 if d.get("source") else 0
        self.text(s, rx + 0.3, top + 1.2, rw - 0.6, 6.7 - top - 1.5 - src_h, d["takeaway"], 26, bold=True, min_size=SZ_BODY)
        if d.get("source"):
            self.text(s, rx + 0.3, 6.7 - src_h, rw - 0.6, src_h - 0.2, self.L["source"] + d["source"], SZ_CAPTION,
                      color=t["muted"], italic=True, anchor="b", min_size=SZ_BODY)
        return "half and half", [d["title"]] + [f"{c}: {v}{unit}" for c, v in zip(cats, vals)] + [d["takeaway"]]

    # --- checks -------------------------------------------------------------------
    def check_list(self, items):
        if len(items) > MAX_LINES:
            err(self.idx, f"More than {MAX_LINES} lines on one slide ({len(items)}): split the slide")
        for it in items:
            if words(it) > MAX_WORDS * 2:
                warn(self.idx, f"Long line, {words(it)} words (rule: {MAX_WORDS} or fewer): \"{it[:50]}...\"")

    KINDS = {
        "title": "s_title", "objectives": "s_objectives", "agenda": "s_agenda", "section": "s_section",
        "cards": "s_cards", "two-column": "s_two_column", "process": "s_process", "compare": "s_compare",
        "big-numbers": "s_big_numbers", "question": "s_question", "key-message": "s_key_message",
        "image": "s_image", "table": "s_table", "summary": "s_summary", "closing": "s_closing", "chart": "s_chart",
    }
    RHYTHM_BREAKERS = {"question", "key-message", "image", "big-numbers", "chart", "section"}
    READ_EVERY_ITEM = ("summary", "two-column")
    STOP = {"en": set("the a an and or of to in on for with is are be this that it as at by from your you we our "
                      "they their its then so but not do does did can will".split()),
            "vi": set("và là của cho có các những một này đó khi thì mà với trong để được không cả như hay hoặc "
                      "cũng đã sẽ ở từ tới đến về bị rồi nên vì nếu".split())}

    def build(self):
        prev, streak, used = None, 0, set()
        for d in self.spec["slides"]:
            k = d.get("type")
            if k not in self.KINDS:
                self.idx += 1
                err(self.idx, f"Unknown slide type: {k}. Types: {', '.join(self.KINDS)}")
                self.prs.slides.add_slide(self.blank)
                continue

            def scan(v, where):
                if isinstance(v, str):
                    check_voice(self.idx + 1, v, where, self.lang)
                elif isinstance(v, list):
                    for x in v:
                        scan(x, where)
                elif isinstance(v, dict):
                    for kk, x in v.items():
                        scan(x, "narration" if kk == "narration" else where)
            scan({kk: v for kk, v in d.items() if kk not in ("type", "icon", "icons", "image")}, "on-screen text")
            layout, visible = getattr(self, self.KINDS[k])(d)
            used.add(k)
            streak = streak + 1 if k == prev else 1
            if streak >= 3 and k != "image":
                warn(self.idx, f"Three '{k}' slides in a row: vary the layout")
            prev = k
            narr = d.get("narration", "").strip()
            notes = narr + (("\n\n[Production] " + d["notes"]) if d.get("notes") else "")
            self.prs.slides[self.idx - 1].notes_slide.notes_text_frame.text = notes.strip()
            self.check_narration(narr, d)
            dur = d.get("duration")
            if not dur:
                dur = max(4, round(words(narr) / WPM * 60)) + pause_seconds_in_notes(d.get("notes", ""))
            self.blocks.append(dict(idx=self.idx, type=k, layout=layout, visible=[v for v in visible if v],
                                    narration=narr, notes=d.get("notes", ""), dur=dur, video=d.get("video")))
        self.check_items_read()
        if self.n_total >= 9:
            kinds = [b["type"] for b in self.blocks]
            for mark in (self.n_total / 3, 2 * self.n_total / 3):
                zone = range(max(0, int(mark) - 2), min(len(kinds), int(mark) + 2))
                if not any(kinds[i] in self.RHYTHM_BREAKERS for i in zone):
                    warn(0, f"Around slide {int(mark) + 1} (the 1/3 or 2/3 mark) there is no rhythm breaker "
                            "(question, key message, big number, chart, image, section): attention drops here")
        plan = self.spec.get("videos")
        if plan:
            tot = {}
            for b in self.blocks:
                if b.get("video") is not None:
                    tot[b["video"]] = tot.get(b["video"], 0) + b["dur"]
            for v in plan:
                got = tot.get(v["number"], 0) / 60.0
                if got < v["minutes"] * 0.6:
                    warn(0, f"Video {v['number']} ({v.get('title', '')}): slides, narration and pauses add up to {got:.1f} min, "
                            f"the lesson plan says {v['minutes']} min (below 60%). Compare with the lesson source: is a section, "
                            "example or pause question missing? If the source really is shorter, fix the plan. Do not pad.")
                elif got > v["minutes"] * 1.35:
                    warn(0, f"Video {v['number']}: {got:.1f} min, more than 35% over the planned {v['minutes']} min. Consider splitting it")
            loose = [b["idx"] for b in self.blocks if b.get("video") is None]
            if loose:
                warn(0, f"Slides not assigned to a video: {loose}")
            self.video_tot = tot
        agendas = [d for d in self.spec["slides"] if d.get("type") == "agenda"]
        if agendas and not self.spec.get("video_number"):
            n_parts = len(agendas[0].get("items", []))
            n_sec = sum(1 for d in self.spec["slides"] if d.get("type") == "section")
            # A part divider is a 'section' slide or the agenda repeated with "current" on part 2 or later
            # (SKILL.md allows either). The first agenda, on part 1, opens the lesson and does not count.
            again = set()
            for d in agendas:
                try:
                    if int(d.get("current") or 0) >= 2:
                        again.add(int(d["current"]))
                except (TypeError, ValueError):
                    pass
            n_div = n_sec + len(again)
            if n_parts >= 2 and n_div < n_parts - 1:
                warn(0, f"The agenda lists {n_parts} parts but there are only {n_div} part dividers "
                        f"({n_sec} 'section' slides, {len(again)} repeated agendas with \"current\" 2 or more); "
                        f"need {n_parts - 1}, one before each part from part 2 on")
        if len(used) < 4 and self.n_total >= 8:
            warn(0, f"The whole deck uses only {len(used)} slide types; use at least 4")

    def _content_words(self, text):
        stop = self.STOP.get(self.lang, self.STOP["en"])
        return [w for w in re.findall(r"[0-9a-zà-ỹ]+", str(text).lower()) if w not in stop and (len(w) > 1 or w.isdigit())]

    def check_items_read(self):
        """On slides whose items appear one by one with the voice, the narration of THAT slide must read every item."""
        sl = self.spec["slides"]

        def cover(line, said):
            ws = self._content_words(line)
            return (sum(1 for w in ws if w in said) / len(ws)) if ws else 1.0
        for i, d in enumerate(sl):
            if d.get("type") not in self.READ_EVERY_ITEM or not d.get("items"):
                continue
            said = set(self._content_words(d.get("narration", "")))
            nxt = set(self._content_words(sl[i + 1].get("narration", ""))) if i + 1 < len(sl) else set()
            for line in d["items"]:
                if not isinstance(line, str) or cover(line, said) >= 0.5:
                    continue
                if nxt and cover(line, nxt) >= 0.5:
                    err(i + 1, f"Item \"{line[:50]}\" is shown on this '{d['type']}' slide but read on the NEXT slide. Read each item on its own slide")
                else:
                    err(i + 1, f"Item \"{line[:50]}\" is never read by this slide's narration (under 50% of its content words). Read every item as it appears")

    def check_narration(self, narr, d):
        if not narr:
            warn(self.idx, "No narration")
            return
        low = narr.lower()
        pause = re.search(self.L["pause_phrase"], low)
        if d.get("type") == "question":
            if not pause and not d.get("options"):
                warn(self.idx, "A 'question' slide whose narration never says to pause the video: the video will stand still without telling anyone why")
            if d.get("timer") and pause:
                on_screen = seconds_in_text(d["timer"])
                spoken = seconds_in_text(narr)
                if on_screen and not (on_screen & spoken):
                    err(self.idx, f"Timer label \"{d['timer']}\" does not match the time the narration gives ({sorted(spoken) or 'none'} s): learners see one number and hear another")
        if self.lang == "vi":
            # "bốn chữ này" only after "bốn chữ" has been named (earlier narration or on screen)
            on_screen = json.dumps({k: v for k, v in d.items() if k not in ("narration", "notes")}, ensure_ascii=False).lower()
            for m in re.finditer(r"(một|hai|ba|bốn|năm|sáu|bảy|tám|chín|mười) (chữ|từ khóa|từ|mô hình|thành tố|bước|tiêu chí|ý|phiếu|kiểu) (này|đó|kia|ấy)(?![a-zà-ỹ])", low):
                phrase = f"{m.group(1)} {m.group(2)}"
                if phrase not in self._said + " " + low[:m.start()] + " " + on_screen:
                    err(self.idx, f"\"{m.group(0)}\" points back at \"{phrase}\" before it was named: listeners cannot tell what it refers to")
        else:
            on_screen = json.dumps({k: v for k, v in d.items() if k not in ("narration", "notes")}, ensure_ascii=False).lower()
            for m in re.finditer(r"\b(these|those) (two|three|four|five|six|seven|eight|nine|ten) (words|terms|models|steps|criteria|ideas|parts|questions)\b", low):
                phrase = f"{m.group(2)} {m.group(3)}"
                if phrase not in self._said + " " + low[:m.start()] + " " + on_screen:
                    err(self.idx, f"\"{m.group(0)}\" points back at \"{phrase}\" before they were named: listeners cannot tell what it refers to")
        self._said += " " + low
        for line in narr.splitlines():
            line = line.strip()
            if line and line[-1] not in ".,?!":
                warn(self.idx, f"Narration line does not end with . , ? or !: \"{line[:50]}\"")
            if re.search(r"\d", line):
                warn(self.idx, f"Digits in the narration; write numbers the way they are spoken: \"{line[:50]}\"")

    def write_script(self, path, pptx_name):
        sp = self.spec
        total = sum(b["dur"] for b in self.blocks)
        out = [f"# {self.L['script_title']}: {sp.get('lesson_code', '')} {self.T(sp.get('lesson_title', ''))}".strip(), "",
               "| | |", "|---|---|",
               f"| Course | {sp.get('course', '')} |",
               f"| Instructor | {sp.get('instructor', '')} |",
               f"| Slide file | `{pptx_name}` (the narration is also in each slide's speaker notes) |",
               "| Slide images | `_preview/slide-NN.png` after `render_slides.py` (1920x1080) |",
               f"| Slides | {self.n_total} |",
               f"| Estimated length | {total // 60} min {total % 60} s (at {WPM} words per minute, plus pauses) |",
               f"| Voice | {sp.get('voice', 'any text-to-speech voice or a human narrator; one line = one breath')} |"]
        if sp.get("videos") and getattr(self, "video_tot", None):
            out += ["", "## Coverage per video", "", "| Video | Title | Planned (min) | Slides + narration (min) |", "|---|---|---|---|"]
            for v in sp["videos"]:
                out.append(f"| {v['number']} | {v.get('title', '')} | {v['minutes']} | {self.video_tot.get(v['number'], 0) / 60:.1f} |")
        out += ["", "## Overview", "", "| # | Type | Layout | Seconds | On screen |", "|---|---|---|---|---|"]
        for b in self.blocks:
            first = (b["visible"][0] if b["visible"] else "").replace("|", "/")
            out.append(f"| S{b['idx']:02d} | {b['type']} | {b['layout']} | {b['dur']} | {first[:70]} |")
        out += ["", "## Slide by slide"]
        for b in self.blocks:
            out += ["", f"### S{b['idx']:02d} · {b['type']} · {b['dur']} s", "", "**On screen:**"]
            out += [f"- {v}" for v in b["visible"]]
            if b["notes"]:
                out += ["", f"**Production notes:** {b['notes']}"]
            out += ["", "**Narration:**", "", "```", b["narration"] or "(none)", "```"]
        with open(path, "w", encoding="utf-8") as f:
            f.write("\n".join(out) + "\n")
        return total


def main():
    if len(sys.argv) < 2 or sys.argv[1] in ("-h", "--help"):
        print(__doc__)
        sys.exit(2)
    check_only = "--check-only" in sys.argv
    argv = [a for a in sys.argv[1:] if a != "--check-only"]
    src = os.path.abspath(argv[0])
    out_dir = tempfile.mkdtemp(prefix="slides_check_") if check_only else (
        os.path.abspath(argv[1]) if len(argv) > 1 else os.path.dirname(src))
    os.makedirs(out_dir, exist_ok=True)
    with open(src, encoding="utf-8") as f:
        spec = json.load(f)
    spec["_path"] = src
    if spec.get("theme", "navy") not in THEMES:
        raise SystemExit(f"theme must be one of: {', '.join(THEMES)}")
    name = spec.get("file_name") or os.path.splitext(os.path.basename(src))[0]
    deck = Deck(spec, out_dir)
    deck.build()
    # Final gate: no text run anywhere below 24 pt, measured on the deck about to be saved.
    for si, sl in enumerate(deck.prs.slides, 1):
        for sh in sl.shapes:
            frames = [sh.text_frame] if sh.has_text_frame else (
                [c.text_frame for r in sh.table.rows for c in r.cells] if getattr(sh, "has_table", False) and sh.has_table else [])
            for tf in frames:
                for para in tf.paragraphs:
                    for r in para.runs:
                        if r.text.strip() and r.font.size is not None and r.font.size.pt < SZ_BODY:
                            err(si, f"Text at {r.font.size.pt:g} pt is below the 24 pt floor: \"{r.text[:40]}\"")
    if ERRORS:
        print("ERRORS (nothing written):")
        for e in ERRORS:
            print("  -", e)
        if WARNINGS:
            print("Warnings:")
            for w in WARNINGS:
                print("  -", w)
        sys.exit(1)
    if check_only:
        print("CHECK ONLY: no blocking errors (nothing written).")
        for w in WARNINGS:
            print("  - warning", w)
        shutil.rmtree(out_dir, ignore_errors=True)
        return
    pptx_path = os.path.join(out_dir, name + ".pptx")
    md_path = os.path.join(out_dir, name + ".video-script.md")
    deck.prs.save(pptx_path)
    total = deck.write_script(md_path, os.path.basename(pptx_path))
    print(f"Wrote: {pptx_path}")
    print(f"Wrote: {md_path}")
    print(f"Slides: {deck.n_total}; estimated length: {total // 60} min {total % 60} s")
    if WARNINGS:
        print("Warnings:")
        for w in WARNINGS:
            print("  -", w)


if __name__ == "__main__":
    main()
