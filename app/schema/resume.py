from pydantic import AliasChoices, BaseModel, Field

from app.schema.template import SectionKey


class SectionRef(BaseModel):
    key: SectionKey
    heading: str                      # text as written in the file
    custom_title: str | None = None   # which custom section, when key == "custom"



    
class Link(BaseModel):
    label: str
    url: str | None = None


class Header(BaseModel):
    name: str
    location: str | None = None
    email: str | None = None
    phone: str | None = None
    links: list[Link] = []


class SkillGroup(BaseModel):
    category: str
    items: list[str]


class Experience(BaseModel):
    role: str
    company: str | None = None
    location: str | None = None
    start: str | None = None  # keep dates exactly as written
    end: str | None = None
    bullets: list[str] = []


class Project(BaseModel):
    name: str
    subtitle: str | None = None
    links: list[Link] = []
    tech: list[str] = []
    start: str | None = None
    end: str | None = None
    bullets: list[str] = []


class Certification(BaseModel):
    title: str
    issuer: str | None = None
    date: str | None = None
    details: str | None = None
    

class Education(BaseModel):
    degree: str
    institution: str | None = None
    location: str | None = None
    grade: str | None = None
    start: str | None = None
    end: str | None = None
    notes: list[str] = []


class CustomSection(BaseModel):
    title: str
    lines: list[str] = []
    links: list[Link] = []

class Achievement(BaseModel):           # was Achivement
    title: str
    detail: str | None = None           # required detail forces the LLM to invent text

class LanguageSkill(BaseModel):         # was Languages (a single item, not a list)
    language: str
    proficiency: str | None = None

class Resume(BaseModel):
    header: Header
    summary: str | None = None
    skills: list[SkillGroup] = []
    experience: list[Experience] = []
    projects: list[Project] = []
    certifications: list[Certification] = []
    achievements: list[Achievement] = Field(default_factory=list, validation_alias=AliasChoices("achievements", "achivements"))
    education: list[Education] = []
    languages: list[LanguageSkill] = []
    custom_sections: list[CustomSection] = []  # anything that doesn't fit above
    section_order: list[SectionRef] = []