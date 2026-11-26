import subprocess
import sys
from pathlib import Path

import pytest
from docx import Document

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))
import doc2rtf  # noqa: E402


def make_docx(path: Path, text: str) -> Path:
    d = Document()
    d.add_heading("Test Heading", level=1)
    p = d.add_paragraph(text + " ")
    p.add_run("bold part").bold = True
    t = d.add_table(rows=1, cols=2)
    t.cell(0, 0).text, t.cell(0, 1).text = "CellA", "CellB"
    d.save(path)
    return path


def make_doc(tmp: Path, text: str) -> Path:
    """Create a legacy .doc by converting a .docx with LibreOffice."""
    src = make_docx(tmp / "legacy_src.docx", text)
    outdir = tmp / "docgen"
    subprocess.run([doc2rtf.find_soffice(), f"-env:UserInstallation={(tmp / 'p').as_uri()}",
                    "--headless", "--convert-to", "doc:MS Word 97", "--outdir", str(outdir),
                    str(src)], check=True, capture_output=True, timeout=120)
    doc = outdir / "legacy_src.doc"
    final = tmp / "legacy.doc"
    doc.rename(final)
    return final


def assert_rtf(path: Path, *needles: str):
    data = path.read_text(encoding="latin-1")
    assert data.startswith("{\\rtf1")
    for n in needles:
        assert n in data


def test_docx_to_rtf(tmp_path):
    src = make_docx(tmp_path / "sample.docx", "Hello DOCX world")
    out = doc2rtf.convert_file(src)
    assert out == tmp_path / "sample.rtf"
    assert_rtf(out, "Hello DOCX world", "bold part", "CellA", "\\b")


def test_doc_to_rtf(tmp_path):
    src = make_doc(tmp_path, "Hello DOC world")
    out = doc2rtf.convert_file(src, tmp_path / "out")
    assert out == tmp_path / "out" / "legacy.rtf"
    assert_rtf(out, "Hello DOC world", "CellB")


def test_unsupported_extension(tmp_path):
    f = tmp_path / "x.txt"
    f.write_text("nope")
    with pytest.raises(doc2rtf.ConversionError):
        doc2rtf.convert_file(f)


def test_missing_file(tmp_path):
    with pytest.raises(doc2rtf.ConversionError):
        doc2rtf.convert_file(tmp_path / "missing.docx")


def test_no_overwrite(tmp_path):
    src = make_docx(tmp_path / "a.docx", "x")
    (tmp_path / "a.rtf").write_text("existing")
    with pytest.raises(doc2rtf.ConversionError):
        doc2rtf.convert_file(src, overwrite=False)


def test_cli_folder_recursive(tmp_path):
    (tmp_path / "in" / "sub").mkdir(parents=True)
    make_docx(tmp_path / "in" / "one.docx", "first")
    make_docx(tmp_path / "in" / "sub" / "two.docx", "second")
    (tmp_path / "in" / "ignore.txt").write_text("skip")
    rc = doc2rtf.main([str(tmp_path / "in"), "-r", "-o", str(tmp_path / "out")])
    assert rc == 0
    assert sorted(p.name for p in (tmp_path / "out").iterdir()) == ["one.rtf", "two.rtf"]
