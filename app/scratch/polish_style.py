"""
polish_style.py
---------------
Takes a Template that was measured from an imported DOCX and asks Gemini to
suggest targeted styling improvements:
  - Fix ATS violations (font sizes, margins, etc.)
  - Fill in nulls with good defaults that match the document''s visual identity
  - Improve spacing, layout choices, and date format
  - Keeps the original font family, color palette, and structure

Returns a *new* Template with only the AI-suggested fields overridden.
Everything not mentioned by the AI stays exactly as measured.
"""

from __future__ import annotations

from typing import Any

from google import genai
from pydantic import BaseModel

from app.config import settings
from app.schema.resume import Resume
from app.schema.template import (
    ATS_SAFE_FONTS,
    AtsRules,
    Bullets,
    Colors,
    CertificationLayout,
    AchievementLayout,
    CustomLayout,
    Dates,
    EducationLayout,
    ExperienceLayout,
    HeaderLayout,
    LanguageLayout,
    Links,
    Page,
    ProjectLayout,
    Rule,
    RoleStyle,
    SectionLayouts,
    SkillsLayout,
    Spacing,
    Structure,
    SummaryLayout,
    Template,
    Typography,
)


# ---------------------------------------------------------------------------
# TemplateOverride - same shape as Template but every field is Optional
# Gemini fills only the fields it wants to change.
# ---------------------------------------------------------------------------

class _RoleStyleOverride(BaseModel):
    font: str | None = None
    size_pt: float | None = None
    bold: bool | None = None
    italic: bool | None = None
    caps: bool | None = None
    color: str | None = None
    align: str | None = None

class _ColorsOverride(BaseModel):
    text: str | None = None
    muted: str | None = None
    accent: str | None = None
    link: str | None = None
    background: str | None = None

class _TypographyOverride(BaseModel):
    primary_font: str | None = None
    secondary_font: str | None = None
    name: _RoleStyleOverride | None = None
    contact: _RoleStyleOverride | None = None
    section_title: _RoleStyleOverride | None = None
    entry_title: _RoleStyleOverride | None = None
    entry_subtitle: _RoleStyleOverride | None = None
    body: _RoleStyleOverride | None = None
    bullet: _RoleStyleOverride | None = None
    date: _RoleStyleOverride | None = None
    link: _RoleStyleOverride | None = None

class _SpacingOverride(BaseModel):
    line_height: float | None = None
    between_sections_pt: float | None = None
    title_to_content_pt: float | None = None
    between_blocks_pt: float | None = None
    inside_block_pt: float | None = None

class _RuleOverride(BaseModel):
    enabled: bool | None = None
    color: str | None = None
    width_pt: float | None = None
    style: str | None = None
    space_pt: float | None = None

class _BulletsOverride(BaseModel):
    char: str | None = None
    indent_in: float | None = None
    hanging_in: float | None = None
    max_per_entry: int | None = None

class _DatesOverride(BaseModel):
    format: str | None = None
    range_separator: str | None = None
    present_label: str | None = None
    position: str | None = None

class _PageOverride(BaseModel):
    margin_top_in: float | None = None
    margin_bottom_in: float | None = None
    margin_left_in: float | None = None
    margin_right_in: float | None = None
    max_pages: int | None = None

class _HeaderOverride(BaseModel):
    alignment: str | None = None
    contact_layout: str | None = None
    contact_separator: str | None = None
    location_detail: str | None = None
    divider: _RuleOverride | None = None

class _SectionLayoutsOverride(BaseModel):
    summary: SummaryLayout | None = None
    skills: SkillsLayout | None = None
    experience: ExperienceLayout | None = None
    projects: ProjectLayout | None = None
    education: EducationLayout | None = None
    certifications: CertificationLayout | None = None
    achievements: AchievementLayout | None = None
    languages: LanguageLayout | None = None
    custom: CustomLayout | None = None

class _LinksOverride(BaseModel):
    display: str | None = None
    underline: bool | None = None

class TemplateOverride(BaseModel):
    """Partial template - only the fields the AI wants to change."""
    reasoning: str | None = None          # AI explains why it made each change
    page: _PageOverride | None = None
    colors: _ColorsOverride | None = None
    typography: _TypographyOverride | None = None
    spacing: _SpacingOverride | None = None
    bullets: _BulletsOverride | None = None
    dates: _DatesOverride | None = None
    links: _LinksOverride | None = None
    header: _HeaderOverride | None = None
    section_layouts: _SectionLayoutsOverride | None = None
    section_rule: _RuleOverride | None = None


# ---------------------------------------------------------------------------
# Helpers: human-readable template summary and ATS issue list
# ---------------------------------------------------------------------------

def _fmt(val: Any, unit: str = "") -> str:
    return f"{val}{unit}" if val is not None else "null (unmeasured)"


def build_style_summary(t: Template) -> str:
    ty = t.typography
    sp = t.spacing
    pg = t.page

    lines = [
        "=== PAGE ===",
        f"  Format: {pg.format}  Columns: {pg.columns}",
        f"  Margins (in): top={_fmt(pg.margin_top_in)} bottom={_fmt(pg.margin_bottom_in)}"
        f"  left={_fmt(pg.margin_left_in)} right={_fmt(pg.margin_right_in)}",
        f"  max_pages: {pg.max_pages}",
        "",
        "=== COLORS ===",
        f"  text={t.colors.text}  muted={t.colors.muted}  accent={t.colors.accent}"
        f"  link={t.colors.link}  bg={t.colors.background}",
        "",
        "=== TYPOGRAPHY ===",
        f"  primary_font: {ty.primary_font}  secondary_font: {ty.secondary_font}",
    ]
    for role_name in ["name", "contact", "section_title", "entry_title",
                      "entry_subtitle", "body", "bullet", "date", "link"]:
        rs: RoleStyle = getattr(ty, role_name)
        lines.append(
            f"  {role_name:15}: font={rs.font or '(inherit)'}  size={_fmt(rs.size_pt, 'pt')}"
            f"  bold={rs.bold}  italic={rs.italic}  caps={rs.caps}"
            f"  color={rs.color}  align={rs.align}"
        )
    lines += [
        "",
        "=== SPACING ===",
        f"  line_height: {_fmt(sp.line_height)}",
        f"  between_sections: {_fmt(sp.between_sections_pt, 'pt')}",
        f"  title_to_content: {_fmt(sp.title_to_content_pt, 'pt')}",
        f"  between_blocks: {_fmt(sp.between_blocks_pt, 'pt')}",
        f"  inside_block: {_fmt(sp.inside_block_pt, 'pt')}",
        "",
        "=== BULLETS ===",
        f"  char={t.bullets.char}  indent={_fmt(t.bullets.indent_in, 'in')}"
        f"  hanging={_fmt(t.bullets.hanging_in, 'in')}  max_per_entry={t.bullets.max_per_entry}",
        "",
        "=== DATES ===",
        f"  format={t.dates.format}  position={t.dates.position}"
        f"  separator={t.dates.range_separator}  present_label={t.dates.present_label}",
        "",
        "=== HEADER ===",
        f"  alignment={t.header.alignment}  contact_layout={t.header.contact_layout}"
        f"  contact_separator={t.header.contact_separator}  location_detail={t.header.location_detail}",
        f"  divider: enabled={t.header.divider.enabled}  color={t.header.divider.color}"
        f"  width={_fmt(t.header.divider.width_pt, 'pt')}",
        "",
        "=== SECTION RULE ===",
        f"  enabled={t.section_rule.enabled}  color={t.section_rule.color}"
        f"  width={_fmt(t.section_rule.width_pt, 'pt')}  style={t.section_rule.style}"
        f"  space={_fmt(t.section_rule.space_pt, 'pt')}",
        "",
        "=== SECTION ORDER ===",
        "  " + " -> ".join(s.key for s in t.sections),
        "",
        "=== SECTION LAYOUTS ===",
        f"  skills.mode: {t.section_layouts.skills.mode}",
        f"  experience.title_order: {t.section_layouts.experience.title_order}"
        f"  title_separator: {t.section_layouts.experience.title_separator}"
        f"  location_position: {t.section_layouts.experience.location_position}",
        f"  projects.tech_position: {t.section_layouts.projects.tech_position}"
        f"  link_position: {t.section_layouts.projects.link_position}",
        f"  education.institution_position: {t.section_layouts.education.institution_position}",
    ]
    return "\n".join(lines)


def ats_issues_list(t: Template) -> str:
    """Return a plain-text list of ATS issues detected in the template."""
    issues = []
    ty = t.typography
    pg = t.page
    ats = t.ats

    # Font sizes
    for role_name in ["body", "bullet"]:
        rs: RoleStyle = getattr(ty, role_name)
        if rs.size_pt is not None and rs.size_pt < ats.min_body_pt:
            issues.append(f"  - {role_name} font size {rs.size_pt}pt < min {ats.min_body_pt}pt")
        if rs.size_pt is not None and rs.size_pt > ats.max_body_pt:
            issues.append(f"  - {role_name} font size {rs.size_pt}pt > max {ats.max_body_pt}pt")

    # Unsafe fonts
    if ats.safe_fonts_only:
        for role_name in ["body", "bullet", "section_title", "entry_title"]:
            rs: RoleStyle = getattr(ty, role_name)
            if rs.font and rs.font not in ATS_SAFE_FONTS:
                issues.append(f"  - {role_name} uses non-ATS-safe font ''{rs.font}''")

    # Margins
    for side, val in [("top", pg.margin_top_in), ("bottom", pg.margin_bottom_in),
                      ("left", pg.margin_left_in), ("right", pg.margin_right_in)]:
        if val is not None and val < ats.min_margin_in:
            issues.append(f"  - margin_{side} = {val:.2f}in < min {ats.min_margin_in}in")

    # Date format
    if t.dates.format == "YYYY":
        issues.append("  - dates.format=''YYYY'' loses month info; prefer ''short_name'' or ''MM/YYYY''")

    # Structure issues
    if t.structure.uses_tables:
        issues.append("  - Document uses tables (ATS may mis-parse columns)")
    if t.structure.uses_text_boxes:
        issues.append("  - Document uses text boxes (invisible to most ATS parsers)")
    if t.structure.has_images:
        issues.append("  - Document contains images")
    if t.structure.contact_in_page_header:
        issues.append("  - Contact info is in the page header (hidden from ATS)")

    # Null spacing
    null_fields = []
    for fname in ["line_height", "between_sections_pt", "title_to_content_pt",
                  "between_blocks_pt", "inside_block_pt"]:
        if getattr(t.spacing, fname) is None:
            null_fields.append(fname)
    if null_fields:
        issues.append(f"  - spacing has unmeasured nulls: {', '.join(null_fields)}")

    return "\n".join(issues) if issues else "None detected."


# ---------------------------------------------------------------------------
# Deep merge: apply TemplateOverride onto Template
# ---------------------------------------------------------------------------

def _merge_rule(base: Rule, ov: _RuleOverride | None) -> Rule:
    if ov is None:
        return base
    d = base.model_dump()
    for k, v in ov.model_dump(exclude_none=True).items():
        d[k] = v
    return Rule(**d)


def _merge_role_style(base: RoleStyle, ov: _RoleStyleOverride | None) -> RoleStyle:
    if ov is None:
        return base
    d = base.model_dump()
    for k, v in ov.model_dump(exclude_none=True).items():
        d[k] = v
    return RoleStyle(**d)


def apply_overrides(base: Template, ov: TemplateOverride) -> Template:
    """Deep-merge ov onto base. Only non-None override fields replace base values."""
    t = base.model_copy(deep=True)

    if ov.page:
        pg = t.page.model_dump()
        for k, v in ov.page.model_dump(exclude_none=True).items():
            pg[k] = v
        t.page = Page(**pg)

    if ov.colors:
        col = t.colors.model_dump()
        for k, v in ov.colors.model_dump(exclude_none=True).items():
            col[k] = v
        t.colors = Colors(**col)

    if ov.typography:
        for role_name in ["name", "contact", "section_title", "entry_title",
                          "entry_subtitle", "body", "bullet", "date", "link"]:
            role_ov: _RoleStyleOverride | None = getattr(ov.typography, role_name)
            if role_ov:
                setattr(t.typography, role_name,
                        _merge_role_style(getattr(t.typography, role_name), role_ov))
        if ov.typography.primary_font is not None:
            t.typography.primary_font = ov.typography.primary_font
        if ov.typography.secondary_font is not None:
            t.typography.secondary_font = ov.typography.secondary_font

    if ov.spacing:
        sp = t.spacing.model_dump()
        for k, v in ov.spacing.model_dump(exclude_none=True).items():
            sp[k] = v
        t.spacing = Spacing(**sp)

    if ov.bullets:
        bl = t.bullets.model_dump()
        for k, v in ov.bullets.model_dump(exclude_none=True).items():
            bl[k] = v
        t.bullets = Bullets(**bl)

    if ov.dates:
        dt = t.dates.model_dump()
        for k, v in ov.dates.model_dump(exclude_none=True).items():
            dt[k] = v
        t.dates = Dates(**dt)

    if ov.links:
        lk = t.links.model_dump()
        for k, v in ov.links.model_dump(exclude_none=True).items():
            lk[k] = v
        t.links = Links(**lk)

    if ov.header:
        hd = t.header.model_dump()
        for k, v in ov.header.model_dump(exclude={"divider"}, exclude_none=True).items():
            hd[k] = v
        t.header = HeaderLayout(**hd)
        if ov.header.divider:
            t.header.divider = _merge_rule(t.header.divider, ov.header.divider)

    if ov.section_layouts:
        sl = ov.section_layouts
        cur = t.section_layouts
        for sec_name in ["summary", "skills", "experience", "projects",
                         "education", "certifications", "achievements", "languages", "custom"]:
            override_sec = getattr(sl, sec_name)
            if override_sec is not None:
                setattr(cur, sec_name, override_sec)
        t.section_layouts = cur

    if ov.section_rule:
        t.section_rule = _merge_rule(t.section_rule, ov.section_rule)

    return t


# ---------------------------------------------------------------------------
# Main function
# ---------------------------------------------------------------------------

POLISH_PROMPT = """\
You are an expert resume designer and ATS consultant.

Below is the current template measured from an uploaded DOCX resume.
Your job is to suggest targeted improvements to make this resume:
1. ATS-safe (fonts, sizes, margins within standard ranges)
2. Visually polished (good spacing, readable hierarchy, consistent layout)
3. Better structured (correct date format, clear role/company separation, proper section order)

=== CURRENT TEMPLATE ===
{summary}

=== ATS ISSUES DETECTED ===
{issues}

=== RULES ===
- Preserve the original font family, color palette, and visual identity.
- Do NOT change any text content (that is handled separately).
- Fill null spacing values with sensible defaults matching the document density.
- If a section_rule is not enabled but matches a detected border color, enable it.
- Keep margins at or above 0.5in for ATS safety.
- Set dates.format to "short_name" or "MM/YYYY" unless the original clearly shows another.
- ALWAYS set dates.position to "right_same_line" (right-aligned tab stop on the same line as the role/title).
- Output ONLY the fields you want to change. Leave unchanged fields absent.
- Add a brief "reasoning" string explaining each change you made.

Output a JSON object matching the TemplateOverride schema."""


def polish_template(template: Template, resume: Resume | None = None) -> tuple[Template, TemplateOverride]:
    """
    Ask Gemini to improve the template styling.

    Returns:
        (polished_template, override) - the override shows what changed and why.
    """
    summary = build_style_summary(template)
    issues = ats_issues_list(template)

    prompt = POLISH_PROMPT.format(summary=summary, issues=issues)

    client = genai.Client(api_key=settings.gemini_api_key)
    resp = client.models.generate_content(
        model=settings.gemini_model,
        contents=prompt,
        config={
            "response_mime_type": "application/json",
            "response_schema": TemplateOverride,
            "temperature": 0.3,
        },
    )
    override = TemplateOverride.model_validate_json(resp.text)
    polished = apply_overrides(template, override)

    # Hard-enforce user preference: dates always right-aligned on same line
    polished.dates.position = "right_same_line"

    return polished, override
