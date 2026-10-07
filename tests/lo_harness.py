"""Reusable LibreOffice (headless, UNO) plumbing for end-to-end tests.

Each :class:`LibreOffice` instance starts its own ``soffice`` process with a
unique socket port and a throw-away user profile, so several instances (or
repeated test runs) never interfere.  Processes are started in their own
process group and are killed on ``stop()``, on context-manager exit and at
interpreter exit (including Ctrl-C during startup and SIGTERM, which is turned
into ``SystemExit`` unless the program installed its own handler).  Only a
SIGKILL of the Python process can orphan soffice.

Example::

    from lo_harness import LibreOffice

    with LibreOffice() as lo:
        doc = lo.open("book.xlsm")                 # hidden, macros allowed
        print(doc.basic_modules())                 # {"VBAProject": {"Module1": "..."}}
        doc.run_macro("Module1", "WriteOk")        # library auto-discovered
        doc.recalc()
        print(doc.get_string("Sheet1", "A1"))
        doc.export_pdf("book.pdf")
        doc.close()

Requires the system ``python3`` with the ``uno`` module (python3-uno).
"""

from __future__ import annotations

import atexit
import os
import shutil
import signal
import socket
import subprocess
import tempfile
import threading
import time
import uuid
from pathlib import Path
from typing import Any, Optional, Sequence

import uno  # noqa: E402  (provided by python3-uno)
from com.sun.star.beans import PropertyValue  # type: ignore
from com.sun.star.connection import NoConnectException  # type: ignore

__all__ = ["LibreOffice", "LoDocument", "LoError", "prop", "props", "default_base_dir"]

# Cell content types of com.sun.star.table.CellContentType
CELL_EMPTY, CELL_VALUE, CELL_TEXT, CELL_FORMULA = "EMPTY", "VALUE", "TEXT", "FORMULA"

# com.sun.star.document.MacroExecMode.ALWAYS_EXECUTE_NO_WARN
MACRO_ALWAYS_EXECUTE_NO_WARN = 4

# Strong references: an instance dropped without stop() is still killed at exit.
_LIVE: set = set()


class LoError(RuntimeError):
    """LibreOffice failed to start, load a document or run a macro."""


def prop(name: str, value: Any) -> PropertyValue:
    """Build a ``com.sun.star.beans.PropertyValue``."""
    p = PropertyValue()
    p.Name = name
    p.Value = value
    return p


def props(**kwargs: Any) -> tuple:
    """``props(Hidden=True)`` -> tuple of PropertyValue (UNO media descriptor)."""
    return tuple(prop(k, v) for k, v in kwargs.items())


def default_base_dir() -> Path:
    """Where profiles go: ``$LHOS_SCRATCH`` if set, else the system temp dir."""
    return Path(os.environ.get("LHOS_SCRATCH") or tempfile.gettempdir())


def _free_port() -> int:
    with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as s:
        s.bind(("127.0.0.1", 0))
        return s.getsockname()[1]


def _url(path: os.PathLike | str) -> str:
    return uno.systemPathToFileUrl(str(Path(path).resolve()))


def _group_members(pgid: int) -> list:
    """PIDs of live (non-zombie) processes in process group ``pgid`` (Linux)."""
    members = []
    for entry in Path("/proc").iterdir():
        if not entry.name.isdigit():
            continue
        try:
            stat = (entry / "stat").read_text()
        except OSError:
            continue
        # Fields after the parenthesised command: state ppid pgrp ...
        fields = stat[stat.rfind(")") + 2 :].split()
        if len(fields) > 2 and int(fields[2]) == pgid and fields[0] != "Z":
            members.append(int(entry.name))
    return members


@atexit.register
def _kill_all() -> None:
    for lo in list(_LIVE):
        try:
            lo.stop(graceful=False)
        except Exception:
            pass


def _exit_on_sigterm(signum: int, _frame: Any) -> None:
    raise SystemExit(128 + signum)


def _install_sigterm_handler() -> None:
    """Turn SIGTERM into ``SystemExit`` so the atexit cleanup above still runs.

    Python's default SIGTERM action ends the process *without* running
    atexit handlers, and soffice lives in its own session, so a test run
    stopped by ``timeout``/CI cancellation would orphan soffice and its
    profile.  Only installed from the main thread and only when nobody else
    handles SIGTERM.
    """
    if threading.current_thread() is not threading.main_thread():
        return
    if signal.getsignal(signal.SIGTERM) is signal.SIG_DFL:
        signal.signal(signal.SIGTERM, _exit_on_sigterm)


class LibreOffice:
    """One headless ``soffice`` process reachable over a UNO socket."""

    def __init__(
        self,
        *,
        base_dir: Optional[os.PathLike | str] = None,
        soffice: str = "soffice",
        start_timeout: float = 90.0,
        keep_profile: bool = False,
        vba_executable: bool = True,
    ) -> None:
        self.soffice = soffice
        self.start_timeout = start_timeout
        self.keep_profile = keep_profile
        self.vba_executable = vba_executable
        base = Path(base_dir) if base_dir else default_base_dir()
        base.mkdir(parents=True, exist_ok=True)
        self.profile_dir = Path(tempfile.mkdtemp(prefix="lo_profile_", dir=base))
        self.port: Optional[int] = None
        self.proc: Optional[subprocess.Popen] = None
        self.ctx = None
        self.smgr = None
        self.desktop = None
        self._log_path = self.profile_dir / "soffice.log"
        self._log = None

    # -- lifecycle ---------------------------------------------------------

    def __enter__(self) -> "LibreOffice":
        self.start()
        return self

    def __exit__(self, *exc: Any) -> None:
        self.stop()

    def start(self) -> "LibreOffice":
        if self.proc is not None:
            return self
        _install_sigterm_handler()
        # stop() removes the profile, so a restarted instance needs it again.
        self.profile_dir.mkdir(parents=True, exist_ok=True)
        # Track the instance *before* spawning: an interrupt (Ctrl-C) while
        # waiting for the socket must not leave soffice running.
        _LIVE.add(self)
        last_error: Optional[Exception] = None
        try:
            for _attempt in range(3):  # a port can be grabbed between probe and bind
                self.port = _free_port()
                try:
                    self._spawn_and_connect()
                    self._configure()
                    return self
                except LoError as exc:
                    last_error = exc
                    self._kill_process()
        except BaseException:
            self.stop(graceful=False)  # any other failure: kill + drop profile
            raise
        self.stop(graceful=False)
        raise LoError(f"could not start LibreOffice: {last_error}")

    def _spawn_and_connect(self) -> None:
        accept = f"socket,host=127.0.0.1,port={self.port};urp;StarOffice.ComponentContext"
        cmd = [
            self.soffice,
            f"-env:UserInstallation={self.profile_dir.as_uri()}",
            "--headless",
            "--invisible",
            "--nologo",
            "--nodefault",
            "--norestore",
            "--nofirststartwizard",
            "--nolockcheck",
            f"--accept={accept}",
        ]
        self._log = open(self._log_path, "ab")
        self.proc = subprocess.Popen(
            cmd,
            stdin=subprocess.DEVNULL,
            stdout=self._log,
            stderr=subprocess.STDOUT,
            start_new_session=True,  # own process group -> killpg reaches soffice.bin
        )
        local = uno.getComponentContext()
        resolver = local.ServiceManager.createInstanceWithContext(
            "com.sun.star.bridge.UnoUrlResolver", local
        )
        url = f"uno:socket,host=127.0.0.1,port={self.port};urp;StarOffice.ComponentContext"
        deadline = time.monotonic() + self.start_timeout
        while True:
            if self.proc.poll() is not None:
                raise LoError(f"soffice exited with {self.proc.returncode}: {self._tail_log()}")
            try:
                self.ctx = resolver.resolve(url)
                break
            except NoConnectException:
                if time.monotonic() > deadline:
                    raise LoError(f"timeout connecting to soffice on port {self.port}")
                time.sleep(0.25)
        self.smgr = self.ctx.ServiceManager
        self.desktop = self.smgr.createInstanceWithContext("com.sun.star.frame.Desktop", self.ctx)

    def _configure(self) -> None:
        """Make imported VBA executable (Tools > Options > Load/Save > VBA)."""
        self.set_config(
            "/org.openoffice.Office.Calc/Filter/Import/VBA",
            Load=True,
            Executable=self.vba_executable,
            Save=True,
        )

    def set_config(self, node: str, **values: Any) -> None:
        """Set configuration values under ``node`` and commit them."""
        provider = self.smgr.createInstanceWithContext(
            "com.sun.star.configuration.ConfigurationProvider", self.ctx
        )
        access = provider.createInstanceWithArguments(
            "com.sun.star.configuration.ConfigurationUpdateAccess", (prop("nodepath", node),)
        )
        for key, value in values.items():
            access.setPropertyValue(key, value)
        access.commitChanges()

    def _tail_log(self, size: int = 2000) -> str:
        try:
            if self._log:
                self._log.flush()
            return self._log_path.read_bytes()[-size:].decode("utf-8", "replace")
        except OSError:
            return ""

    def _kill_process(self) -> None:
        """Force-kill soffice and wait until every process of it is gone.

        ``soffice`` runs as ``oosplash`` (our child, process-group leader)
        plus ``soffice.bin`` (its child).  Killing soffice.bin first lets
        oosplash reap it and exit, so no zombie is left behind for init; the
        whole group is then SIGKILLed as a fallback.
        """
        if self.proc is not None:
            pgid = self.proc.pid  # start_new_session=True -> pgid == pid
            linux = Path("/proc").is_dir()
            if linux and self.proc.poll() is None:
                for pid in _group_members(pgid):
                    if pid != pgid:
                        try:
                            os.kill(pid, signal.SIGKILL)
                        except ProcessLookupError:
                            pass
                try:
                    self.proc.wait(timeout=5)
                except subprocess.TimeoutExpired:
                    pass
            try:
                os.killpg(pgid, signal.SIGKILL)
            except (ProcessLookupError, PermissionError):
                pass
            try:
                self.proc.wait(timeout=10)
            except subprocess.TimeoutExpired:
                pass
            deadline = time.monotonic() + 10
            while linux and _group_members(pgid) and time.monotonic() < deadline:
                time.sleep(0.05)
        self.proc = None
        self.ctx = self.smgr = self.desktop = None
        if self._log:
            self._log.close()
            self._log = None

    def stop(self, graceful: bool = True, timeout: float = 20.0) -> None:
        """Terminate soffice (gracefully if possible) and remove the profile."""
        if self.proc is not None and graceful and self.desktop is not None:
            try:
                self.desktop.terminate()
            except Exception:
                pass  # the bridge drops when soffice exits; that's expected
            try:
                self.proc.wait(timeout=timeout)
            except subprocess.TimeoutExpired:
                pass
        self._kill_process()  # whatever is left of oosplash + soffice.bin
        _LIVE.discard(self)
        if not self.keep_profile:
            shutil.rmtree(self.profile_dir, ignore_errors=True)

    # -- documents ---------------------------------------------------------

    def open(
        self,
        path: os.PathLike | str,
        *,
        hidden: bool = True,
        read_only: bool = False,
        macros: bool = True,
        filter_name: Optional[str] = None,
    ) -> "LoDocument":
        """Load a document; macros run without prompts unless ``macros=False``.

        Note: LibreOffice dispatches VBA document events (``Workbook_Open``,
        ``Worksheet_Activate`` ...) only for non-hidden loads; pass
        ``hidden=False`` to test them (still headless, nothing is shown).
        """
        if self.desktop is None:
            self.start()
        path = Path(path)
        if not path.exists():
            raise LoError(f"no such file: {path}")
        args = dict(Hidden=hidden, ReadOnly=read_only)
        if macros:
            args["MacroExecutionMode"] = MACRO_ALWAYS_EXECUTE_NO_WARN
        if filter_name:
            args["FilterName"] = filter_name
        comp = self.desktop.loadComponentFromURL(_url(path), "_blank", 0, props(**args))
        if comp is None:
            raise LoError(f"LibreOffice could not load {path}")
        return LoDocument(self, comp, path)

    def new_calc(self, hidden: bool = True) -> "LoDocument":
        comp = self.desktop.loadComponentFromURL(
            "private:factory/scalc", "_blank", 0, props(Hidden=hidden)
        )
        return LoDocument(self, comp, None)

    def convert(
        self, src: os.PathLike | str, dst: os.PathLike | str, filter_name: str, **filter_opts: Any
    ) -> Path:
        """Open ``src`` hidden, store it as ``dst`` with ``filter_name``, close it."""
        doc = self.open(src, read_only=True, macros=False)
        try:
            return doc.store_as(dst, filter_name, **filter_opts)
        finally:
            doc.close()


class LoDocument:
    """Thin wrapper over a loaded spreadsheet component."""

    def __init__(self, lo: LibreOffice, component: Any, path: Optional[Path]) -> None:
        self.lo = lo
        self.doc = component
        self.path = path

    # -- sheets / cells ----------------------------------------------------

    def sheet_names(self) -> list:
        return list(self.doc.Sheets.getElementNames())

    def sheet(self, name_or_index: str | int) -> Any:
        sheets = self.doc.Sheets
        if isinstance(name_or_index, int):
            return sheets.getByIndex(name_or_index)
        if not sheets.hasByName(name_or_index):
            raise LoError(f"no sheet {name_or_index!r}; have {self.sheet_names()}")
        return sheets.getByName(name_or_index)

    def cell(self, sheet: str | int, address: str) -> Any:
        return self.sheet(sheet).getCellRangeByName(address)

    def get_value(self, sheet: str | int, address: str) -> float:
        return self.cell(sheet, address).getValue()

    def get_string(self, sheet: str | int, address: str) -> str:
        return self.cell(sheet, address).getString()

    def get_formula(self, sheet: str | int, address: str) -> str:
        return self.cell(sheet, address).getFormula()

    def get_type(self, sheet: str | int, address: str) -> str:
        """One of ``EMPTY``, ``VALUE``, ``TEXT``, ``FORMULA``."""
        return self.cell(sheet, address).Type.value

    def get(self, sheet: str | int, address: str) -> Any:
        """Python value of one cell: None, float or str (formulas -> result)."""
        cell = self.cell(sheet, address)
        kind = cell.Type.value
        if kind == CELL_EMPTY:
            return None
        if kind == CELL_VALUE:
            return cell.getValue()
        if kind == CELL_TEXT:
            return cell.getString()
        # Formula: numeric result unless it evaluates to text.
        result_type = cell.FormulaResultType2  # com.sun.star.sheet.FormulaResult
        return cell.getString() if result_type == 2 else cell.getValue()

    def get_range(self, sheet: str | int, address: str) -> list:
        """Rows of values (floats / strings; empty cells are '') for a range."""
        return [list(row) for row in self.sheet(sheet).getCellRangeByName(address).getDataArray()]

    def set_value(self, sheet: str | int, address: str, value: Any) -> None:
        cell = self.cell(sheet, address)
        if isinstance(value, str):
            cell.setString(value)
        else:
            cell.setValue(value)

    def recalc(self) -> None:
        self.doc.calculateAll()

    # -- macros ------------------------------------------------------------

    def basic_libraries(self) -> list:
        return list(self.doc.BasicLibraries.getElementNames())

    def basic_modules(self) -> dict:
        """``{library: {module: source}}`` for the document's Basic libraries."""
        libs = self.doc.BasicLibraries
        result = {}
        for name in libs.getElementNames():
            libs.loadLibrary(name)
            lib = libs.getByName(name)
            result[name] = {m: lib.getByName(m) for m in lib.getElementNames()}
        return result

    def find_library(self, module: str) -> str:
        """Name of the Basic library containing ``module``."""
        for lib, modules in self.basic_modules().items():
            if module in modules:
                return lib
        raise LoError(f"module {module!r} not found; libraries: {self.basic_libraries()}")

    def run_macro(
        self, module: str, sub: str, args: Sequence[Any] = (), *, library: Optional[str] = None
    ) -> Any:
        """Run ``library.module.sub(*args)`` from the document; returns its result."""
        library = library or self.find_library(module)
        uri = (
            f"vnd.sun.star.script:{library}.{module}.{sub}"
            "?language=Basic&location=document"
        )
        try:
            script = self.doc.getScriptProvider().getScript(uri)
            result, _out_idx, _out = script.invoke(tuple(args), (), ())
        except Exception as exc:  # uno exceptions are not subclasses of LoError
            raise LoError(f"macro {uri} failed: {exc}") from exc
        return result

    # -- output ------------------------------------------------------------

    def store_as(self, dst: os.PathLike | str, filter_name: str, **filter_opts: Any) -> Path:
        """``storeToURL`` with a filter, e.g. ``calc_pdf_Export``, ``calc_png_Export``,
        ``Calc MS Excel 2007 XML``.  ``filter_opts`` become FilterData entries."""
        dst = Path(dst)
        dst.parent.mkdir(parents=True, exist_ok=True)
        args = {"FilterName": filter_name, "Overwrite": True}
        if filter_opts:
            args["FilterData"] = uno.Any(
                "[]com.sun.star.beans.PropertyValue", props(**filter_opts)
            )
        uno.invoke(self.doc, "storeToURL", (_url(dst), props(**args)))
        return dst

    def export_pdf(self, dst: os.PathLike | str, **filter_opts: Any) -> Path:
        return self.store_as(dst, "calc_pdf_Export", **filter_opts)

    def export_png(self, dst: os.PathLike | str, **filter_opts: Any) -> Path:
        """PNG of the active sheet (``PixelWidth``/``PixelHeight`` optional)."""
        return self.store_as(dst, "calc_png_Export", **filter_opts)

    def close(self) -> None:
        if self.doc is not None:
            try:
                self.doc.close(True)
            except Exception:
                try:
                    self.doc.dispose()
                except Exception:
                    pass
            self.doc = None


def unique_name(prefix: str = "lhos") -> str:
    """Helper for tests that want unique file names in a shared scratch dir."""
    return f"{prefix}_{uuid.uuid4().hex[:10]}"
