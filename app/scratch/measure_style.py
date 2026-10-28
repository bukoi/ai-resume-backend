from __future__ import annotations

import re
from collections import Counter
from statistics import median

from docx import Document
from docx.oxml.ns import qn
from lxml import etree

from app.schema.resume import Resume
from app.schema.roles import RoleMap
from app.schema.template import (
    Bullets,
    Colors,
    Dates,
    HeaderLayout,
    Links,
    Page,
    RoleStyle,
    Rule,
    SectionConfig,
    SectionLayouts,
    SkillsLayout,
    Spacing,
    Structure,
    Template,
    Typography,
)
from app.scratch.dump_docx import Block

JC = {"left": "left", "start": "left", "center": "center", "right": "right",
      "end": "right", "both": "justify", "distribute": "justify"}
PAGES = {"A4": (8.27, 11.69), "Letter": (8.5, 11.0), "Legal": (8.5, 14.0)}

# May is valid in both long and short format:
LONG = r"\b(January|February|March|April|May|June|July|August|September|October|November|December)\b\s+\d{4}"
SHORT = r"\b(Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Sept|Oct|Nov|Dec)\b\.?\s+\d{4}"
DATE_PATTERNS = [("long_name", LONG), ("short_name", SHORT),
                 ("MM.YYYY", r"\b\d{1,2}\.\d{4}\b"), ("MM/YYYY", r"\b\d{1,2}/\d{4}\b"),
                 ("YYYY", r"^\D*\d{4}\D*$")]


def _chain(style):
    while style is not None:
        yield style
        style = style.base_style


def _top(counter: Counter):
    return counter.most_common(1)[0][0] if counter else None


def _med(values):
    values = [v for v in values if v is not None]
    return round(median(values), 2) if values else None


class Ctx:
    """Inherited formatting lookup: direct -> style chain -> document defaults."""

    def __init__(self, doc):
        defaults = doc.styles.element.find(qn("w:docDefaults"))
        self.def_rpr = self.def_ppr = None
        if defaults is not None:
            r, p = defaults.find(qn("w:rPrDefault")), defaults.find(qn("w:pPrDefault"))
            self.def_rpr = r.find(qn("w:rPr")) if r is not None else None
            self.def_ppr = p.find(qn("w:pPr")) if p is not None else None
        self.theme = self._theme_fonts(doc)

    @staticmethod
    def _theme_fonts(doc) -> dict:
        out = {}
        ns = {"a": "http://schemas.openxmlformats.org/drawingml/2006/main"}
        for part in doc.part.package.iter_parts():
            if str(part.partname).endswith("/theme/theme1.xml"):
                root = etree.fromstring(part.blob)
                for kind in ("major", "minor"):
                    el = root.find(f".//a:{kind}Font/a:latin", ns)
                    if el is not None:
                        out[kind] = el.get("typeface")
        return out

    def _rprs(self, run, p):
        yield run._r.rPr
        for s in _chain(p.style):
            yield s.element.rPr
        yield self.def_rpr

    def _pprs(self, p):
        yield p._p.pPr
        for s in _chain(p.style):
            yield s.element.pPr
        yield self.def_ppr

    def font(self, run, p):
        for rpr in self._rprs(run, p):
            rf = rpr.find(qn("w:rFonts")) if rpr is not None else None
            if rf is None:
                continue
            name = rf.get(qn("w:ascii")) or rf.get(qn("w:hAnsi"))
            if name:
                return name
            theme = rf.get(qn("w:asciiTheme")) or rf.get(qn("w:hAnsiTheme"))
            if theme:
                return self.theme.get("major" if theme.startswith("major") else "minor")
        return None

    def size(self, run, p):
        for rpr in self._rprs(run, p):
            sz = rpr.find(qn("w:sz")) if rpr is not None else None
            if sz is not None:
                return int(sz.get(qn("w:val"))) / 2
        return None

    def toggle(self, run, p, tag):
        for rpr in self._rprs(run, p):
            el = rpr.find(qn(tag)) if rpr is not None else None
            if el is not None:
                return el.get(qn("w:val"), "true") not in ("0", "false", "off")
        return None

    def color(self, run, p):
        for rpr in self._rprs(run, p):
            el = rpr.find(qn("w:color")) if rpr is not None else None
            if el is not None:
                val = el.get(qn("w:val"))
                return None if val in (None, "auto") else val.upper()
        return None

    def align(self, p):
        for ppr in self._pprs(p):
            jc = ppr.find(qn("w:jc")) if ppr is not None else None
            if jc is not None:
                return JC.get(jc.get(qn("w:val")))
        return "left"

    def spacing_pt(self, p, attr):          # attr: "w:before" or "w:after"
        for ppr in self._pprs(p):
            sp = ppr.find(qn("w:spacing")) if ppr is not None else None
            if sp is not None and sp.get(qn(attr)) is not None:
                return int(sp.get(qn(attr))) / 20
        return None

    def line_height(self, p):
        for ppr in self._pprs(p):
            sp = ppr.find(qn("w:spacing")) if ppr is not None else None
            if sp is not None and sp.get(qn("w:line")) is not None:
                if sp.get(qn("w:lineRule"), "auto") == "auto":
                    return round(int(sp.get(qn("w:line"))) / 240, 2)
                return None
        return None

    def indent_in(self, p):
        for ppr in self._pprs(p):
            ind = ppr.find(qn("w:ind")) if ppr is not None else None
            if ind is not None:
                left = ind.get(qn("w:left"))
                hanging = ind.get(qn("w:hanging"))
                left_in = round(int(left) / 1440, 2) if left else None
                hang_in = round(int(hanging) / 1440, 2) if hanging else None
                return left_in, hang_in
        return None, None

    def rule_below(self, p):
        for ppr in self._pprs(p):
            bdr = ppr.find(qn("w:pBdr")) if ppr is not None else None
            b = bdr.find(qn("w:bottom")) if bdr is not None else None
            if b is not None:
                if b.get(qn("w:val")) in ("nil", "none"):
                    return None
                val = b.get(qn("w:val"))
                color = b.get(qn("w:color"))
                sz = b.get(qn("w:sz"))
                space = b.get(qn("w:space"))
                return (
                    val or "single",
                    None if color in (None, "auto") else color.upper(),
                    int(sz) / 8 if sz else None,
                    int(space) if space else None,
                )
        return None


def clean_roles(roles: RoleMap, n: int) -> RoleMap:
    """Drop invented or out-of-range indices from the LLM output."""
    return RoleMap(**{k: sorted({i for i in v if 0 <= i < n}) for k, v in roles.model_dump().items()})


def measure_role(ctx: Ctx, blocks: list[Block], idxs: list[int]) -> RoleStyle:
    fonts, sizes, bold, italic, caps, colors, aligns = (Counter() for _ in range(7))
    for i in idxs:
        b = blocks[i]
        p = b.para
        if p is None:
            continue
        aligns[ctx.align(p)] += 1
        for r in p.runs:
            n = len(r.text.strip())
            if not n:
                continue
            for counter, value in (
                (fonts, ctx.font(r, p)), (sizes, ctx.size(r, p)),
                (bold, bool(ctx.toggle(r, p, "w:b"))),
                (italic, bool(ctx.toggle(r, p, "w:i"))),
                (caps, bool(ctx.toggle(r, p, "w:caps")) or b.caps),
                (colors, ctx.color(r, p)),
            ):
                if value is not None:
                    counter[value] += n
    return RoleStyle(font=_top(fonts), size_pt=_top(sizes), bold=_top(bold), italic=_top(italic),
                     caps=_top(caps), color=_top(colors), align=_top(aligns))


def measure_link_style(doc) -> tuple[RoleStyle, str | None, bool]:
    fonts, sizes, colors, underlines = (Counter() for _ in range(4))
    for p in doc.paragraphs:
        for hl in p._p.iter(qn("w:hyperlink")):
            for r in hl.iter(qn("w:r")):
                rPr = r.find(qn("w:rPr"))
                if rPr is not None:
                    c = rPr.find(qn("w:color"))
                    if c is not None and c.get(qn("w:val")) not in (None, "auto"):
                        colors[c.get(qn("w:val")).upper()] += 1
                    u = rPr.find(qn("w:u"))
                    if u is not None and u.get(qn("w:val"), "single") not in ("none", "nil"):
                        underlines[True] += 1
                    sz = rPr.find(qn("w:sz"))
                    if sz is not None:
                        sizes[int(sz.get(qn("w:val"))) / 2] += 1
                    rf = rPr.find(qn("w:rFonts"))
                    if rf is not None:
                        name = rf.get(qn("w:ascii")) or rf.get(qn("w:hAnsi"))
                        if name:
                            fonts[name] += 1

    # Check Hyperlink style in document styles if not found directly
    if not colors:
        try:
            for style in doc.styles:
                if style.name.lower() == "hyperlink":
                    if style.font.color and style.font.color.rgb:
                        colors[str(style.font.color.rgb)] += 1
                    if style.font.underline:
                        underlines[True] += 1
                    if style.font.size:
                        sizes[style.font.size.pt] += 1
                    if style.font.name:
                        fonts[style.font.name] += 1
        except Exception:
            pass

    link_color = _top(colors)
    underline = bool(_top(underlines)) if underlines else False
    link_style = RoleStyle(font=_top(fonts), size_pt=_top(sizes), color=link_color)
    return link_style, link_color, underline


def measure_page(doc) -> Page:
    s = doc.sections[0]
    w, h = s.page_width.inches, s.page_height.inches
    fmt = next((n for n, (pw, ph) in PAGES.items() if abs(w - pw) < 0.05 and abs(h - ph) < 0.05), "Other")
    cols = s._sectPr.find(qn("w:cols"))
    return Page(
        format=fmt, width_in=round(w, 2), height_in=round(h, 2),
        margin_top_in=round(s.top_margin.inches, 2), margin_bottom_in=round(s.bottom_margin.inches, 2),
        margin_left_in=round(s.left_margin.inches, 2), margin_right_in=round(s.right_margin.inches, 2),
        header_footer_in=round(s.header_distance.inches, 2),
        columns=int(cols.get(qn("w:num"), 1)) if cols is not None else 1,
    )


def measure_bullets(ctx: Ctx, blocks: list[Block]) -> Bullets:
    bullet_blocks = [b for b in blocks if b.bullet and b.para is not None]
    chars = Counter()
    lefts = []
    hangs = []
    for b in bullet_blocks:
        t = b.text.lstrip()
        if t and t[0] in "•-–▪*●":
            chars[t[0]] += 1
        l, h = ctx.indent_in(b.para)
        if l is not None:
            lefts.append(l)
        if h is not None:
            hangs.append(h)

    char = _top(chars) or "•"
    if char not in ("•", "-", "–", "▪"):
        char = "•"
    return Bullets(
        char=char,
        indent_in=_med(lefts) if lefts else 0.1,
        hanging_in=_med(hangs) if hangs else 0.15,
    )


def measure_spacing(ctx: Ctx, blocks: list[Block], roles: RoleMap) -> Spacing:
    def of(idxs, attr, fallback=None):
        vals = [ctx.spacing_pt(blocks[i].para, attr) for i in idxs if blocks[i].para is not None]
        m = _med(vals)
        return m if m is not None else fallback

    bullets = [i for i, b in enumerate(blocks) if b.bullet]
    sample = (roles.body or bullets or [None])[0]
    p0 = blocks[sample].para if sample is not None else None

    lh = ctx.line_height(p0) if p0 is not None else None
    if lh is None:
        lh = 1.15

    return Spacing(
        line_height=lh,
        between_sections_pt=of(roles.section_titles, "w:before", fallback=6.0),
        title_to_content_pt=of(roles.section_titles, "w:after", fallback=2.0),
        between_blocks_pt=of(roles.entry_titles, "w:before", fallback=5.0),
        inside_block_pt=of(bullets, "w:after", fallback=1.0),
    )


def detect_date_format(dates: list[str]):
    # Distinguish long vs short month formats even when "May" is present
    long_count = sum(1 for d in dates if re.search(r"\b(January|February|March|April|June|July|August|September|October|November|December)\b\s+\d{4}", d, re.I))
    short_count = sum(1 for d in dates if re.search(r"\b(Jan|Feb|Mar|Apr|Jun|Jul|Aug|Sep|Sept|Oct|Nov|Dec)\b\.?\s+\d{4}", d, re.I))

    votes = Counter()
    for d in dates:
        for name, pattern in DATE_PATTERNS:
            if re.search(pattern, d, re.I):
                if re.search(r"\bMay\b\s+\d{4}", d, re.I):
                    name = "long_name" if long_count > short_count else "short_name"
                votes[name] += 1
                break
    return _top(votes)


def detect_range_separator(blocks: list[Block]) -> str:
    patterns = [
        ("–", r"\b\d{4}\s*–\s*(\d{4}|Present|Current|Ongoing)"),
        ("-", r"\b\d{4}\s*-\s*(\d{4}|Present|Current|Ongoing)"),
        ("to", r"\b\d{4}\s+to\s+(\d{4}|Present|Current|Ongoing)"),
    ]
    for b in blocks:
        for sep, pat in patterns:
            if re.search(pat, b.text, re.I):
                return sep
    return "–"


def detect_location_detail(loc: str | None) -> str:
    if not loc:
        return "city_region"
    parts = [p.strip() for p in loc.split(",") if p.strip()]
    if len(parts) <= 1:
        return "city"
    elif len(parts) == 2:
        return "city_region"
    else:
        return "full"


def trailing_alignment(blocks: list[Block], needle: str | None):
    """'right' if the text sits after a tab or in a right-aligned line, else 'left'."""
    if not needle:
        return None
    for b in blocks:
        if needle in b.text:
            if b.align == "right":
                return "right"
            tail = b.text.rsplit("⇥", 1)[-1] if "⇥" in b.text else ""
            return "right" if needle in tail else "left"
    return None


def measure_skills(blocks: list[Block], idxs: list[int]):
    cells = [blocks[i].where for i in idxs if blocks[i].where.startswith("table")]
    columns = {w.split("c")[-1] for w in cells}
    layout = "columns" if len(columns) > 1 else "inline"
    texts = [blocks[i].text for i in idxs]
    categorized = "category_per_line" if len(texts) > 1 and all(":" in t for t in texts) else None
    return layout, (len(columns) if layout == "columns" else None), categorized


def measure_structure(doc, blocks) -> Structure:
    return Structure(
        uses_tables=any(b.where.startswith("table") for b in blocks),
        uses_text_boxes=any(b.where == "textbox" for b in blocks),
        has_images=len(doc.inline_shapes) > 0,          # floating images are not counted
        contact_in_page_header=any(b.where == "header" for b in blocks),
    )


def measure_template(path: str, blocks: list[Block], roles: RoleMap, resume: Resume) -> Template:
    doc = Document(path)
    ctx = Ctx(doc)
    roles = clean_roles(roles, len(blocks))
    bullets_idx = [i for i, b in enumerate(blocks) if b.bullet]

    # Measure link styles (Fix B)
    link_style, link_color, link_underline = measure_link_style(doc)

    typo = Typography(
        name=measure_role(ctx, blocks, roles.name),
        contact=measure_role(ctx, blocks, roles.contact),
        section_title=measure_role(ctx, blocks, roles.section_titles),
        entry_title=measure_role(ctx, blocks, roles.entry_titles),
        entry_subtitle=measure_role(ctx, blocks, roles.entry_subtitles),
        body=measure_role(ctx, blocks, roles.body),
        bullet=measure_role(ctx, blocks, bullets_idx),
        date=measure_role(ctx, blocks, roles.dates),
        link=link_style,
    )
    typo.primary_font = typo.name.font or typo.section_title.font
    typo.secondary_font = typo.body.font or typo.bullet.font

    dated = [*resume.experience, *resume.projects, *resume.education]
    all_dates = [d for x in dated for d in (x.start, x.end) if d] + [c.date for c in resume.certifications if c.date]
    first_start = next((x.start for x in dated if x.start), None)
    first_loc = next((x.location for x in [*resume.experience, *resume.education] if x.location), None)

    # Section rule
    sec_rule_raw = next((r for i in roles.section_titles
                         if blocks[i].para is not None and (r := ctx.rule_below(blocks[i].para))), None)
    sec_rule = Rule(
        enabled=True,
        style=sec_rule_raw[0],
        color=sec_rule_raw[1],
        width_pt=sec_rule_raw[2],
        space_pt=sec_rule_raw[3],
    ) if sec_rule_raw else Rule()

    # Header divider rule (Fix D)
    header_rule_raw = next((r for i in (roles.name + roles.contact)
                           if blocks[i].para is not None and (r := ctx.rule_below(blocks[i].para))), None)
    header_divider = Rule(
        enabled=True,
        style=header_rule_raw[0],
        color=header_rule_raw[1],
        width_pt=header_rule_raw[2],
        space_pt=header_rule_raw[3],
    ) if header_rule_raw else Rule()

    skills_layout, skills_cols, categorized = measure_skills(blocks, roles.skills)

    header = HeaderLayout(
        alignment=typo.name.align if typo.name.align in ("left", "center", "right") else "center",
        location_detail=detect_location_detail(first_loc),
        divider=header_divider,
    )

    dates = Dates(
        format=detect_date_format(all_dates) or "short_name",
        range_separator=detect_range_separator(blocks),
        position="right_same_line" if trailing_alignment(blocks, first_start) == "right" else "own_line",
    )

    links = Links(
        display="label",
        underline=link_underline,
    )

    bullets = measure_bullets(ctx, blocks)

    colors = Colors(
        text=typo.body.color or "000000",
        accent=typo.name.color or (sec_rule.color if sec_rule.enabled else None),
        muted=typo.entry_subtitle.color or typo.contact.color,
        link=link_color,
    )

    sections = [
        SectionConfig(key=s.key, heading=s.heading, custom_title=s.custom_title)
        for s in resume.section_order
    ]
    section_layouts = SectionLayouts(
        skills=SkillsLayout(
            mode="columns" if skills_layout == "columns" else "category_lines",
            columns=skills_cols,
        )
    )

    return Template(
        page=measure_page(doc),
        colors=colors,
        typography=typo,
        spacing=measure_spacing(ctx, blocks, roles),
        bullets=bullets,
        dates=dates,
        links=links,
        header=header,
        sections=sections,
        section_layouts=section_layouts,
        section_rule=sec_rule,
        structure=measure_structure(doc, blocks),
    )