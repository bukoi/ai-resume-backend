import json
import sys
from pathlib import Path

from app.config import settings
from app.export.docx_renderer import render_docx
from app.export.pdf_preview import docx_to_pdf
from app.schema.resume import Resume
from app.schema.template import Template
from app.scratch.extract import extract_resume_and_template
from app.scratch.polish_style import polish_template

here = Path(__file__).parent
out_dir = Path(settings.export_dir)
out_dir.mkdir(parents=True, exist_ok=True)

resume_json_path = out_dir / "resume.json"
template_json_path = out_dir / "template.json"

# Check if JSON files exist in scratch fallback
if not resume_json_path.exists() and (here / "resume.json").exists():
    resume_json_path = here / "resume.json"
if not template_json_path.exists() and (here / "template.json").exists():
    template_json_path = here / "template.json"

force_extract = "--extract" in sys.argv
do_polish   = "--polish"  in sys.argv
docx_in = here / "resume.docx"

# Parse custom docx path if provided
for arg in sys.argv[1:]:
    if not arg.startswith("--") and arg.endswith(".docx"):
        docx_in = Path(arg)

if force_extract or not (resume_json_path.exists() and template_json_path.exists()):
    print(f"[1/2] Running AI extraction from '{docx_in}'...")
    parsed, template = extract_resume_and_template(docx_in)
    resume = parsed.resume

    # Save to export_dir
    (out_dir / "resume.json").write_text(resume.model_dump_json(indent=2), encoding="utf-8")
    (out_dir / "template.json").write_text(template.model_dump_json(indent=2), encoding="utf-8")
    print(f"  -> Saved extracted JSONs to {out_dir}")
else:
    print(f"[1/2] Loading saved JSONs (use --extract to re-run AI extraction)...")
    print(f"  -> Resume:   {resume_json_path}")
    print(f"  -> Template: {template_json_path}")
    resume_data = json.loads(resume_json_path.read_text(encoding="utf-8"))
    resume = Resume.model_validate(resume_data)
    template = Template.model_validate_json(template_json_path.read_text(encoding="utf-8"))

# Optionally AI-polish the template (keep content, improve styling)
if do_polish:
    print("[1b/2] Asking AI to polish the template styling...")
    template, override = polish_template(template, resume)
    polished_json = out_dir / "template_polished.json"
    override_json = out_dir / "template_override.json"
    polished_json.write_text(template.model_dump_json(indent=2), encoding="utf-8")
    override_json.write_text(override.model_dump_json(indent=2), encoding="utf-8")
    print(f"  -> Polished template: {polished_json}")
    print(f"  -> Override (what changed): {override_json}")
    if override.reasoning:
        print(f"  -> AI reasoning: {override.reasoning}")

# 2. Render to DOCX & PDF
print(f"[2/2] Rendering DOCX & PDF preview...")
out_docx = out_dir / "rendered_resume.docx"

try:
    docx_file = render_docx(resume, template, out_docx)
    print(f"  -> DOCX: {docx_file}")
except PermissionError:
    print(f"\n[ERROR] Permission denied: '{out_docx}' is open in Word. Please close it and re-run.\n")
    sys.exit(1)

try:
    pdf_file = docx_to_pdf(docx_file, out_dir)
    print(f"  -> PDF:  {pdf_file}")
except PermissionError:
    print(f"\n[ERROR] Permission denied: PDF viewer has the file open. Please close it and re-run.\n")
    sys.exit(1)

print("\nFinished! Output files are ready in backend/exports/")