from typing import Literal

from pydantic import BaseModel

Align = Literal["left", "center", "right", "justify"]
SectionKey = Literal["summary", "skills", "experience", "projects", "certifications",
                     "achievements", "education", "languages", "custom"]
Position = Literal["right_same_line", "inline_same_line", "own_line", "hidden"]

# Commonly recommended defaults, not official vendor rules.
ATS_SAFE_FONTS = {"Arial", "Calibri", "Cambria", "Garamond", "Georgia", "Helvetica",
                  "Times New Roman", "Verdana", "Tahoma"}
STANDARD_HEADINGS = {
    "summary": {"summary", "professional summary", "profile"},
    "skills": {"skills", "technical skills"},
    "experience": {"experience", "work experience", "professional experience"},
    "projects": {"projects"},
    "certifications": {"certifications", "certifications & achievements"},
    "achievements": {"achievements", "awards", "certifications & achievements"},
    "education": {"education"},
    "languages": {"languages"},
}


# ---------- LOOK ----------

class Page(BaseModel):
    format: Literal["A4", "Letter", "Legal", "Other"] = "A4"
    width_in: float | None = None
    height_in: float | None = None
    margin_top_in: float | None = None
    margin_bottom_in: float | None = None
    margin_left_in: float | None = None
    margin_right_in: float | None = None
    header_footer_in: float | None = None
    columns: int = 1
    max_pages: int | None = None          # soft limit: the editor warns, it never cuts content


class Colors(BaseModel):
    text: str = "000000"
    muted: str | None = None              # dates, locations, tech stack lines
    accent: str | None = None             # name, section rule
    link: str | None = None
    background: str = "FFFFFF"


class RoleStyle(BaseModel):
    font: str | None = None
    size_pt: float | None = None
    bold: bool | None = None
    italic: bool | None = None
    caps: bool | None = None
    color: str | None = None              # hex, e.g. "1F2937"
    align: Align | None = None


class Typography(BaseModel):
    primary_font: str | None = None       # headings
    secondary_font: str | None = None     # body
    fallback_font: str = "Arial"          # used when a template font is missing
    name: RoleStyle = RoleStyle()
    contact: RoleStyle = RoleStyle()
    section_title: RoleStyle = RoleStyle()
    entry_title: RoleStyle = RoleStyle()
    entry_subtitle: RoleStyle = RoleStyle()
    body: RoleStyle = RoleStyle()
    bullet: RoleStyle = RoleStyle()
    date: RoleStyle = RoleStyle()
    link: RoleStyle = RoleStyle()


class Spacing(BaseModel):
    line_height: float | None = None          # 1.0 = 100%
    between_sections_pt: float | None = None
    title_to_content_pt: float | None = None
    between_blocks_pt: float | None = None
    inside_block_pt: float | None = None


class Rule(BaseModel):                         # the line under section titles or headers
    enabled: bool = False
    color: str | None = None
    width_pt: float | None = None
    style: str = "single"
    space_pt: float | None = None


class Bullets(BaseModel):
    char: Literal["•", "-", "–", "▪"] = "•"
    indent_in: float | None = None             # where the bullet sits
    hanging_in: float | None = None            # gap between bullet and text
    max_per_entry: int | None = None           # soft limit for the editor and the AI


class Dates(BaseModel):
    format: Literal["long_name", "short_name", "MM/YYYY", "MM.YYYY", "YYYY"] | None = "short_name"
    range_separator: Literal["–", "-", "to"] = "–"
    present_label: Literal["Present", "Current", "Ongoing"] = "Present"
    position: Literal["right_same_line", "inline_same_line", "own_line"] = "right_same_line"


class Links(BaseModel):
    display: Literal["label", "url", "label_and_url"] = "label"
    underline: bool = False


# ---------- HEADER ----------

class HeaderLayout(BaseModel):
    alignment: Literal["left", "center", "right"] = "center"
    contact_layout: Literal["one_line", "two_lines", "stacked"] = "one_line"
    contact_separator: Literal["|", "•", "·", "–", ","] = "|"
    contact_order: list[Literal["location", "email", "phone", "links"]] = ["location", "email", "phone", "links"]
    location_detail: Literal["hidden", "city", "city_region", "full"] = "city_region"
    show_photo: bool = False
    divider: Rule = Rule()


# ---------- PER-SECTION LAYOUT (one block per section of your resume schema) ----------

class SummaryLayout(BaseModel):
    max_lines: int | None = 4                   # editor warning

class SkillsLayout(BaseModel):
    mode: Literal["category_lines", "inline", "columns"] = "category_lines"
    columns: int | None = None
    category_bold: bool = True
    category_separator: Literal[":", "-", "|"] = ":"
    item_separator: Literal[", ", " | ", " • "] = ", "

class ExperienceLayout(BaseModel):
    title_order: Literal["role_first", "company_first"] = "role_first"
    title_separator: Literal["|", ",", "–", "at"] = "|"
    location_position: Position = "own_line"

class ProjectLayout(BaseModel):
    subtitle_position: Literal["inline", "own_line", "hidden"] = "own_line"   # name vs description
    tech_position: Literal["inline", "own_line", "after_bullets", "hidden"] = "own_line"
    tech_label: str | None = None               # e.g. "Tech:"; None = no label
    tech_separator: Literal[", ", " | "] = ", "
    link_position: Literal["title_line_right", "after_title", "own_line", "hidden"] = "after_title"

class EducationLayout(BaseModel):
    institution_position: Literal["inline", "own_line"] = "own_line"
    show_grade: bool = True
    show_location: bool = True
    notes_as_bullets: bool = False              # e.g. a "Final Year Project" line

class CertificationLayout(BaseModel):
    issuer_separator: Literal["–", "-", "|", ","] = "–"
    date_position: Literal["inline", "right_same_line"] = "inline"
    as_bullets: bool = True

class AchievementLayout(BaseModel):
    detail_separator: Literal["–", "-", ":", "|"] = "-"
    bold_title: bool = True
    as_bullets: bool = True

class LanguageLayout(BaseModel):
    mode: Literal["inline", "list"] = "inline"
    proficiency_style: Literal["parentheses", "dash", "colon"] = "parentheses"

class CustomLayout(BaseModel):
    mode: Literal["bullets", "paragraph", "inline_list"] = "bullets"


class SectionLayouts(BaseModel):
    summary: SummaryLayout = SummaryLayout()
    skills: SkillsLayout = SkillsLayout()
    experience: ExperienceLayout = ExperienceLayout()
    projects: ProjectLayout = ProjectLayout()
    education: EducationLayout = EducationLayout()
    certifications: CertificationLayout = CertificationLayout()
    achievements: AchievementLayout = AchievementLayout()
    languages: LanguageLayout = LanguageLayout()
    custom: CustomLayout = CustomLayout()


class SectionConfig(BaseModel):                # list order = order in the document
    key: SectionKey
    heading: str                               # text shown as the section title
    custom_title: str | None = None
    visible: bool = True
    also: list[SectionKey] = []                # extra sections rendered under this same heading


# ---------- ATS ----------

class AtsRules(BaseModel):
    mode: Literal["enforce", "warn", "off"] = "warn"   # site templates: enforce; imported: warn
    single_column: bool = True
    no_tables: bool = True
    no_text_boxes: bool = True
    no_images: bool = True
    contact_in_body: bool = True               # not inside the Word page header/footer
    standard_headings: bool = True
    safe_fonts_only: bool = True
    min_body_pt: float = 10.0
    max_body_pt: float = 12.0
    min_margin_in: float = 0.5


class Structure(BaseModel):                    # facts measured from an imported file
    uses_tables: bool = False
    uses_text_boxes: bool = False
    has_images: bool = False
    contact_in_page_header: bool = False


class Template(BaseModel):
    name: str | None = None
    source: Literal["site", "imported"] = "site"
    page: Page = Page()
    colors: Colors = Colors()
    typography: Typography = Typography()
    spacing: Spacing = Spacing()
    bullets: Bullets = Bullets()
    dates: Dates = Dates()
    links: Links = Links()
    header: HeaderLayout = HeaderLayout()
    sections: list[SectionConfig] = []
    section_layouts: SectionLayouts = SectionLayouts()
    section_rule: Rule = Rule()
    ats: AtsRules = AtsRules()
    structure: Structure = Structure()