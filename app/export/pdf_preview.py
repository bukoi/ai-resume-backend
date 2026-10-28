import os
import shutil
import subprocess
import sys
import tempfile
from pathlib import Path


def docx_to_pdf(docx_path: str | Path, out_dir: str | Path | None = None, timeout: int = 90, out_name: str | None = None) -> Path:
    docx_path = Path(docx_path).resolve()
    out_dir = Path(out_dir or docx_path.parent).resolve()
    out_dir.mkdir(parents=True, exist_ok=True)

    soffice = shutil.which("soffice") or shutil.which("libreoffice")
    if not soffice:
        raise RuntimeError("LibreOffice not found. Install libreoffice-writer in the Docker image.")

    # A throwaway profile folder avoids lock errors and makes parallel conversions safe.
    with tempfile.TemporaryDirectory() as profile:
        cmd = [soffice, f"-env:UserInstallation={Path(profile).as_uri()}", "--headless",
               "--norestore", "--convert-to", "pdf", "--outdir", str(out_dir), str(docx_path)]
        try:
            subprocess.run(cmd, check=True, capture_output=True, text=True, timeout=timeout,
                           env={**os.environ, "HOME": profile})
        except subprocess.CalledProcessError as e:
            raise RuntimeError(f"LibreOffice failed: {e.stderr or e.stdout}") from e

    pdf = out_dir / f"{docx_path.stem}.pdf"
    if not pdf.exists():
        raise RuntimeError("LibreOffice finished but produced no PDF.")

    if out_name and out_name != pdf.name:
        target_pdf = out_dir / out_name
        pdf.replace(target_pdf)
        return target_pdf

    return pdf


if __name__ == "__main__":
    print(docx_to_pdf(sys.argv[1], sys.argv[2] if len(sys.argv) > 2 else None))