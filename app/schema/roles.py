from pydantic import BaseModel

from app.schema.resume import Resume


class RoleMap(BaseModel):
    name: list[int] = []
    contact: list[int] = []
    section_titles: list[int] = []
    entry_titles: list[int] = []      # job role, project name, degree lines
    entry_subtitles: list[int] = []   # company, tech stack, institution lines
    body: list[int] = []              # summary and plain paragraph text
    dates: list[int] = []             # any line that contains a date
    skills: list[int] = []


class ParsedResume(BaseModel):
    resume: Resume
    roles: RoleMap