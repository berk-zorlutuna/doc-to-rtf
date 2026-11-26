import gc
import sys
import time
from pathlib import Path
from unittest import mock

import pytest

tk = pytest.importorskip("tkinter")
ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))

from test_doc2rtf import make_doc, make_docx  # noqa: E402


@pytest.fixture
def app():
    try:
        root = tk.Tk()
    except tk.TclError:
        pytest.skip("no display available")
    import doc2rtf_gui
    a = doc2rtf_gui.Doc2RtfApp(root)
    yield a
    root.destroy()
    del a, root
    # Collect the destroyed window here, on the main thread. Otherwise Python's
    # cyclic GC may free its Tk objects later from a conversion worker thread,
    # which makes Tcl abort the process.
    gc.collect()


def run_until_done(app, timeout=120):
    end = time.time() + timeout
    while time.time() < end:
        app.root.update()
        if str(app.convert_btn["state"]) == "normal":
            return
        time.sleep(0.05)
    raise TimeoutError


def test_convert_button_disabled_until_file_chosen(app):
    assert str(app.convert_btn["state"]) == "disabled"


@pytest.mark.parametrize("kind", ["docx", "doc"])
def test_browse_and_convert_saves_next_to_input(app, tmp_path, kind):
    src = make_docx(tmp_path / "My Report.docx", "GUI text") if kind == "docx" \
        else make_doc(tmp_path, "GUI text")
    with mock.patch("doc2rtf_gui.filedialog.askopenfilename", return_value=str(src)), \
         mock.patch("doc2rtf_gui.messagebox.showinfo") as info:
        app.browse()
        assert app.output_var.get() == str(src.with_suffix(".rtf"))
        app.convert()
        run_until_done(app)
    out = src.with_suffix(".rtf")
    assert out.parent == src.parent and out.stem == src.stem
    assert out.read_text(encoding="latin-1").startswith("{\\rtf1")
    assert "GUI text" in out.read_text(encoding="latin-1")
    info.assert_called_once()


def test_cancel_browse_keeps_state(app):
    with mock.patch("doc2rtf_gui.filedialog.askopenfilename", return_value=""):
        app.browse()
    assert app.input_path is None
    assert str(app.convert_btn["state"]) == "disabled"


def test_existing_output_declined_is_not_overwritten(app, tmp_path):
    src = make_docx(tmp_path / "a.docx", "x")
    (tmp_path / "a.rtf").write_text("keep me")
    app.set_input(str(src))
    with mock.patch("doc2rtf_gui.messagebox.askyesno", return_value=False):
        app.convert()
    assert (tmp_path / "a.rtf").read_text() == "keep me"
    assert "cancelled" in app.status_var.get()


def test_error_is_reported(app, tmp_path):
    bad = tmp_path / "broken.docx"
    bad.write_bytes(b"not really a word file")
    app.set_input(str(bad))
    with mock.patch("doc2rtf_gui.convert_file",
                    side_effect=__import__("doc2rtf").ConversionError("boom")), \
         mock.patch("doc2rtf_gui.messagebox.showerror") as err:
        app.convert()
        run_until_done(app)
    err.assert_called_once()
    assert "failed" in app.status_var.get()
