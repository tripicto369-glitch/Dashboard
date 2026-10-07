"""LibreOffice end-to-end test for the generated vbaProject.bin.

Builds a small .xlsm with xlsxwriter + ``build_vba_project`` and opens it in a
headless LibreOffice via UNO (``tests/lo_harness.py``) to prove that:

* every module is imported into the document's Basic library with its source
  (Cyrillic text intact);
* standard-module macros run and write cells (``WriteOk`` -> ``OK-П``);
* functions return values, class modules instantiate, sheet code names resolve;
* the ``ThisWorkbook`` document module is bound (``Workbook_Open`` fires);
* the harness supports concurrent instances and PDF export, can restart a
  stopped instance, and leaves no soffice/profile behind after SIGTERM or
  Ctrl-C during startup.

Run with ``python3 -I tests/test_vba_lo.py``.  When LibreOffice / ``uno`` is
missing the test prints SKIP and exits 0, unless ``LHOS_REQUIRE_LO=1``.
"""

from __future__ import annotations

import atexit
import importlib.util
import os
import shutil
import signal
import subprocess
import sys
import tempfile
import time
import traceback
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "src"))
sys.path.insert(0, str(ROOT / "tests"))

HAVE_LO = importlib.util.find_spec("uno") is not None and shutil.which("soffice") is not None

if not HAVE_LO and "pytest" in sys.modules:  # pragma: no cover
    import pytest

    pytest.skip("LibreOffice/uno not available", allow_module_level=True)

import xlsxwriter  # noqa: E402

from lhos.vbaproject import VbaModule, build_vba_project  # noqa: E402

STANDARD = """Option Explicit

' Модуль проверки: строки на кириллице
Public Const APP_TITLE As String = "Содержание ЛХОС по СИКН"

Sub WriteOk()
    ThisWorkbook.Worksheets(1).Range("A1").Value = "OK-" & ChrW(1055)
End Sub

Function GetTitle() As String
    GetTitle = APP_TITLE
End Function

Function AddUp(ByVal a As Double, ByVal b As Double) As Double
    AddUp = a + b
End Function

Sub ClassToCell()
    Dim c As clsCounter
    Set c = New clsCounter
    c.Add 2
    c.Add 40
    ThisWorkbook.Worksheets(2).Range("B2").Value = c.Total
End Sub

Sub CodeNameToCell()
    shData.Range("B3").Value = "Данные"
End Sub
"""

CLASS = """Option Explicit
Private m_total As Double

Public Sub Add(ByVal v As Double)
    m_total = m_total + v
End Sub

Public Property Get Total() As Double
    Total = m_total
End Property
"""

WORKBOOK = """Option Explicit

Private Sub Workbook_Open()
    ThisWorkbook.Worksheets(1).Range("C1").Value = "opened"
End Sub
"""

MODULES = [
    VbaModule("ThisWorkbook", "workbook", WORKBOOK),
    VbaModule("shDash", "worksheet", "Option Explicit\n"),
    VbaModule("shData", "worksheet", ""),
    VbaModule("modTest", "standard", STANDARD),
    VbaModule("clsCounter", "class", CLASS),
]
SHEETS = [("Дашборд", "shDash"), ("Данные", "shData")]

_STATE: dict = {}


def build_xlsm(directory: Path) -> Path:
    """Write ``vba_e2e.xlsm`` (and its vbaProject.bin) into ``directory``."""
    bin_path = directory / "vbaProject.bin"
    bin_path.write_bytes(build_vba_project(MODULES))
    xlsm = directory / "vba_e2e.xlsm"
    wb = xlsxwriter.Workbook(str(xlsm))
    for title, code_name in SHEETS:
        ws = wb.add_worksheet(title)
        ws.set_vba_name(code_name)
        ws.write("B1", title)
    wb.add_vba_project(str(bin_path))
    wb.set_vba_name("ThisWorkbook")
    wb.close()
    return xlsm


def _env():
    """Lazily create the scratch dir, workbook and a shared LibreOffice."""
    if not _STATE:
        from lo_harness import LibreOffice, default_base_dir

        work = Path(tempfile.mkdtemp(prefix="vba_lo_", dir=default_base_dir()))
        _STATE["work"] = work
        _STATE["xlsm"] = build_xlsm(work)
        _STATE["lo"] = LibreOffice(base_dir=work).start()
        atexit.register(teardown)  # pytest runs never reach main()
    return _STATE


def teardown() -> None:
    if _STATE:
        _STATE["lo"].stop()
        shutil.rmtree(_STATE["work"], ignore_errors=True)
        _STATE.clear()


def _open(**kwargs):
    env = _env()
    return env["lo"].open(env["xlsm"], **kwargs)


def _code_lines(code: str) -> list:
    return [line for line in code.replace("\r\n", "\n").split("\n") if line.strip()]


# ---------------------------------------------------------------------------


def test_modules_imported_with_source() -> None:
    doc = _open()
    try:
        library = doc.find_library("modTest")
        modules = doc.basic_modules()[library]
        assert set(modules) == {m.name for m in MODULES}, modules.keys()
        for m in MODULES:
            lo_lines = [line.rstrip("\r") for line in modules[m.name].split("\n")]
            pos = 0
            for line in _code_lines(m.code):  # our lines appear, in order
                pos = lo_lines.index(line, pos) + 1
        assert 'Public Const APP_TITLE As String = "Содержание ЛХОС по СИКН"' in modules["modTest"]
    finally:
        doc.close()


def test_macro_writes_cell() -> None:
    doc = _open()
    try:
        assert doc.get("Дашборд", "A1") is None
        doc.run_macro("modTest", "WriteOk")
        assert doc.get_string("Дашборд", "A1") == "OK-П"
        assert doc.get_type("Дашборд", "A1") == "TEXT"
    finally:
        doc.close()


def test_functions_classes_and_code_names() -> None:
    doc = _open()
    try:
        assert doc.run_macro("modTest", "GetTitle") == "Содержание ЛХОС по СИКН"
        assert doc.run_macro("modTest", "AddUp", (2.5, 40)) == 42.5
        doc.run_macro("modTest", "ClassToCell")
        doc.run_macro("modTest", "CodeNameToCell")
        doc.recalc()
        assert doc.get("Данные", "B2") == 42.0
        assert doc.get("Данные", "B3") == "Данные"
        names = {s.Name: s.CodeName for s in doc.doc.Sheets}
        assert names == dict(SHEETS), names
    finally:
        doc.close()


def test_workbook_open_event_fires() -> None:
    # LibreOffice only dispatches VBA document events for non-hidden loads
    # (still headless, nothing is displayed).
    doc = _open(hidden=False)
    try:
        assert doc.get_string("Дашборд", "C1") == "opened"
    finally:
        doc.close()


def test_harness_parallel_instances_and_pdf() -> None:
    from lo_harness import LibreOffice

    env = _env()
    other = LibreOffice(base_dir=env["work"]).start()
    try:
        assert other.port != env["lo"].port
        assert other.profile_dir != env["lo"].profile_dir
        doc = other.open(env["xlsm"])
        try:
            doc.run_macro("modTest", "WriteOk")
            pdf = doc.export_pdf(env["work"] / "out.pdf")
            assert pdf.read_bytes()[:5] == b"%PDF-"
        finally:
            doc.close()
    finally:
        profile, proc = other.profile_dir, other.proc
        other.stop()
    assert proc.poll() is not None, "soffice still running after stop()"
    assert not profile.exists(), "profile not removed"


def _procs_using(profile: Path) -> list:
    """PIDs whose command line mentions ``profile`` (independent of the harness)."""
    needle = str(profile).encode()
    pids = []
    for entry in Path("/proc").iterdir():
        if entry.name.isdigit() and int(entry.name) != os.getpid():
            try:
                if needle in (entry / "cmdline").read_bytes():
                    pids.append(int(entry.name))
            except OSError:
                pass
    return pids


def _run_child(body: str, sig: int, delay: float, wait_for_up: bool) -> tuple:
    """Run ``body`` in a child that owns a harness instance ``lo``, then signal it.

    The child prints its profile dir first.  With ``wait_for_up`` the parent
    also waits for a line printed by ``body`` (soffice is running); the
    signal is sent ``delay`` seconds later.  Returns ``(returncode, profile)``
    once no process refers to the profile any more (or after 10 s).
    """
    work = _env()["work"]
    code = (
        "import sys, time\n"
        f"sys.path.insert(0, {str(ROOT / 'tests')!r})\n"
        "from lo_harness import LibreOffice\n"
        f"lo = LibreOffice(base_dir={str(work)!r})\n"
        "print(lo.profile_dir, flush=True)\n" + body
    )
    child = subprocess.Popen([sys.executable, "-I", "-c", code], stdout=subprocess.PIPE, text=True)
    try:
        profile = Path(child.stdout.readline().strip())
        if wait_for_up:
            child.stdout.readline()
        time.sleep(delay)
        child.send_signal(sig)
        rc = child.wait(timeout=60)
    finally:
        if child.poll() is None:
            child.kill()
            child.wait()
    deadline = time.monotonic() + 10
    while _procs_using(profile) and time.monotonic() < deadline:
        time.sleep(0.1)
    return rc, profile


def test_harness_cleans_up_on_sigterm() -> None:
    body = "lo.start()\nprint('up', flush=True)\ntime.sleep(120)\n"
    rc, profile = _run_child(body, signal.SIGTERM, 0.5, wait_for_up=True)
    assert rc == 128 + signal.SIGTERM, rc
    assert not _procs_using(profile), "soffice orphaned after SIGTERM"
    assert not profile.exists(), "profile left after SIGTERM"


def test_harness_cleans_up_on_interrupt_during_start() -> None:
    # SIGINT 0.3 s after the child begins start(), i.e. while it still waits
    # for the UNO socket (soffice needs well over a second to come up).
    body = "try:\n    lo.start()\nexcept KeyboardInterrupt:\n    sys.exit(3)\ntime.sleep(120)\n"
    rc, profile = _run_child(body, signal.SIGINT, 0.3, wait_for_up=False)
    assert rc == 3, f"interrupt did not land inside start() (exit code {rc})"
    assert not _procs_using(profile), "soffice orphaned after Ctrl-C during start()"
    assert not profile.exists(), "profile left after Ctrl-C during start()"


def test_harness_restart_after_stop() -> None:
    from lo_harness import LibreOffice

    lo = LibreOffice(base_dir=_env()["work"])
    try:
        lo.start()
        lo.stop()
        lo.start()  # profile was removed by stop(); start() must recreate it
        doc = lo.open(_env()["xlsm"])
        try:
            assert doc.run_macro("modTest", "AddUp", (1, 2)) == 3.0
        finally:
            doc.close()
    finally:
        lo.stop()
    assert not lo.profile_dir.exists()


# ---------------------------------------------------------------------------


def main() -> int:
    if not HAVE_LO:
        if os.environ.get("LHOS_REQUIRE_LO") == "1":
            print("FAIL LibreOffice/uno required but not available")
            return 1
        print("SKIP LibreOffice/uno not available")
        return 0
    tests = [(n, f) for n, f in globals().items() if n.startswith("test_") and callable(f)]
    failed = 0
    try:
        for name, func in tests:
            try:
                func()
                print(f"PASS {name}")
            except Exception:
                failed += 1
                print(f"FAIL {name}")
                traceback.print_exc()
    finally:
        teardown()
    print(f"{len(tests) - failed}/{len(tests)} passed")
    return 1 if failed else 0


if __name__ == "__main__":
    sys.exit(main())
