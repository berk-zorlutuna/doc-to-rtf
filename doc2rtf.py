#!/usr/bin/env python3
"""Convert DOC and DOCX files to RTF using LibreOffice in headless mode.

Usage:
    python doc2rtf.py input.docx
    python doc2rtf.py file1.doc file2.docx -o out_dir
    python doc2rtf.py some_folder -r -o out_dir
"""
from __future__ import annotations

import argparse
import os
import shutil
import subprocess
import sys
import tempfile
from pathlib import Path

SUPPORTED_EXTENSIONS = {".doc", ".docx"}


class ConversionError(RuntimeError):
    """Raised when a file cannot be converted."""


def find_soffice() -> str:
    """Locate the LibreOffice executable."""
    env = os.environ.get("SOFFICE_PATH")
    if env and Path(env).is_file():
        return env
    for name in ("soffice", "libreoffice"):
        found = shutil.which(name)
        if found:
            return found
    candidates = [
        r"C:\Program Files\LibreOffice\program\soffice.exe",
        r"C:\Program Files (x86)\LibreOffice\program\soffice.exe",
        "/Applications/LibreOffice.app/Contents/MacOS/soffice",
    ]
    for c in candidates:
        if Path(c).is_file():
            return c
    raise ConversionError(
        "LibreOffice was not found. Install it or set SOFFICE_PATH to the soffice executable."
    )


def convert_file(src: str | Path, out_dir: str | Path | None = None,
                 overwrite: bool = True, timeout: int = 120) -> Path:
    """Convert one .doc/.docx file to .rtf and return the output path."""
    src = Path(src).resolve()
    if not src.is_file():
        raise ConversionError(f"Input file not found: {src}")
    if src.suffix.lower() not in SUPPORTED_EXTENSIONS:
        raise ConversionError(f"Unsupported file type (expected .doc/.docx): {src}")

    out_dir = Path(out_dir).resolve() if out_dir else src.parent
    out_dir.mkdir(parents=True, exist_ok=True)
    target = out_dir / (src.stem + ".rtf")
    if target.exists() and not overwrite:
        raise ConversionError(f"Output already exists (use --overwrite): {target}")

    soffice = find_soffice()
    # A private profile dir lets several conversions run at once and avoids
    # clashing with a LibreOffice instance the user may already have open.
    with tempfile.TemporaryDirectory(prefix="doc2rtf_") as tmp:
        profile = Path(tmp, "profile").as_uri()
        work_out = Path(tmp, "out")
        work_out.mkdir()
        cmd = [
            soffice, f"-env:UserInstallation={profile}",
            "--headless", "--norestore", "--nolockcheck",
            "--convert-to", "rtf", "--outdir", str(work_out), str(src),
        ]
        try:
            proc = subprocess.run(cmd, capture_output=True, text=True, timeout=timeout)
        except subprocess.TimeoutExpired as exc:
            raise ConversionError(f"Timed out converting {src}") from exc

        produced = work_out / (src.stem + ".rtf")
        if proc.returncode != 0 or not produced.is_file():
            detail = (proc.stderr or proc.stdout).strip()
            raise ConversionError(f"LibreOffice failed to convert {src}: {detail}")
        shutil.move(str(produced), str(target))
    return target


def collect_inputs(paths: list[str], recursive: bool) -> list[Path]:
    files: list[Path] = []
    for p in map(Path, paths):
        if p.is_dir():
            it = p.rglob("*") if recursive else p.iterdir()
            files.extend(sorted(f for f in it
                                if f.is_file() and f.suffix.lower() in SUPPORTED_EXTENSIONS
                                and not f.name.startswith("~$")))
        else:
            files.append(p)
    return files


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description="Convert DOC/DOCX files to RTF.")
    parser.add_argument("inputs", nargs="+", help="Files or folders to convert")
    parser.add_argument("-o", "--output-dir", help="Output folder (default: next to each input)")
    parser.add_argument("-r", "--recursive", action="store_true", help="Search folders recursively")
    parser.add_argument("--no-overwrite", action="store_true", help="Skip files whose RTF already exists")
    parser.add_argument("--timeout", type=int, default=120, help="Seconds allowed per file")
    args = parser.parse_args(argv)

    files = collect_inputs(args.inputs, args.recursive)
    if not files:
        print("No .doc/.docx files found.", file=sys.stderr)
        return 1

    failures = 0
    for f in files:
        try:
            out = convert_file(f, args.output_dir, overwrite=not args.no_overwrite,
                               timeout=args.timeout)
            print(f"OK    {f} -> {out}")
        except ConversionError as e:
            failures += 1
            print(f"FAIL  {e}", file=sys.stderr)
    print(f"\n{len(files) - failures} converted, {failures} failed.")
    return 1 if failures else 0


if __name__ == "__main__":
    sys.exit(main())
