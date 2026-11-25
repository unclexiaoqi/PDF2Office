"""PDF → Word conversion engine.

Two engines, chosen by the user in the UI:

``word``
    Drives Microsoft Word's built-in PDF import ("PDF Reflow") through a small
    VBScript + COM helper.  Word rebuilds paragraphs, real tables and pictures
    and writes a .docx — the text stays editable and the layout is
    reconstructed as faithfully as Word can manage.

``image``
    Pure Python: every page is rendered with PyMuPDF and wrapped into a .docx
    as one full-page picture.  The layout is then pixel-exact, the text is not
    editable, and **no Office installation is required at all**.

The batch/reporting protocol (START / ENGINE / BEGIN / OK / ERR / DONE) is
shared by both engines so the GUI can treat them identically.
"""

from __future__ import annotations

import io
import os
import shutil
import subprocess
import tempfile
import time
import uuid
from dataclasses import dataclass
from pathlib import Path
from typing import Callable, Optional

import pymupdf
from docx import Document
from docx.enum.text import WD_BREAK
from docx.shared import Emu, Pt
from PIL import Image

MODES = {
    "image300": {
        "name": "原版式",
        "short": "300 DPI",
        "desc": "每页转成高清图，版式和原 PDF 完全一致",
        "tag": "版式 100% 保留 · 文字不可编辑",
        "badge": "推荐",
        "engine": "image",
        "dpi": 300,
        "quality": 92,
        "stage": "正在还原版面…",
    },
    "editable": {
        "name": "可编辑",
        "short": "Word 重排",
        "desc": "文字可编辑、表格还原；版式尽量接近",
        "tag": "复杂排版可能有偏移 · 需要 Office",
        "engine": "word",
        "stage": "正在重排版式…",
    },
    "image200": {
        "name": "版式还原",
        "short": "200 DPI",
        "desc": "每页转成图片，体积更小",
        "tag": "适合邮件 / 微信发送",
        "engine": "image",
        "dpi": 200,
        "quality": 92,
        "stage": "正在还原版面…",
    },
    "image120": {
        "name": "小体积",
        "short": "120 DPI",
        "desc": "每页转成图片，优先减小体积",
        "tag": "适合快速预览",
        "engine": "image",
        "dpi": 120,
        "quality": 85,
        "stage": "正在还原版面…",
    },
}

MODE_ORDER = ["image300", "editable", "image200", "image120"]
DEFAULT_MODE = "image300"
SUPPORTED_EXTS = {".pdf"}

POINTS_PER_EMU = 12700.0


class EngineError(RuntimeError):
    pass


# ---------------------------------------------------------------------------
# Running the console helper without flashing a console window
# ---------------------------------------------------------------------------
_NO_WINDOW_FLAGS = getattr(subprocess, "CREATE_NO_WINDOW", 0x08000000)
_HIDDEN_STARTUP = None
if os.name == "nt":
    _HIDDEN_STARTUP = subprocess.STARTUPINFO()
    _HIDDEN_STARTUP.dwFlags |= subprocess.STARTF_USESHOWWINDOW
    _HIDDEN_STARTUP.wShowWindow = getattr(subprocess, "SW_HIDE", 0)


@dataclass
class FileResult:
    src: str
    dst: Optional[str] = None
    ok: bool = False
    error: Optional[str] = None
    pages: int = 0
    size_before: int = 0
    size_after: int = 0
    seconds: float = 0.0
    cancelled: bool = False
    note: str = ""


def human_size(n: int) -> str:
    if n <= 0:
        return "0 B"
    value = float(n)
    for unit in ("B", "KB", "MB", "GB"):
        if value < 1024 or unit == "GB":
            return f"{value:.0f} {unit}" if unit == "B" else f"{value:.1f} {unit}"
        value /= 1024.0
    return f"{value:.1f} GB"


# ---------------------------------------------------------------------------
# VBScript helper: Word's PDF import
# ---------------------------------------------------------------------------
VBS_WORD = r"""
Option Explicit

Dim fso, listPath, logPath, cancelPath, app, doc, line, parts, src, dst, idx, msg
Dim preRunning, oldAlerts, okCount, failCount, cancelledFlag

Set fso = CreateObject("Scripting.FileSystemObject")
listPath = WScript.Arguments(0)
logPath = WScript.Arguments(1)
cancelPath = WScript.Arguments(2)

Sub LogLine(text)
    Dim f
    Set f = fso.OpenTextFile(logPath, 8, True)
    f.WriteLine text
    f.Close
End Sub

LogLine "START" & vbTab & "0"
Set app = Nothing
preRunning = False
On Error Resume Next
Set app = GetObject("", "Word.Application")
If Err.Number = 0 Then
    preRunning = True
    LogLine "ENGINE" & vbTab & "Word"
Else
    Err.Clear
    Set app = CreateObject("Word.Application")
    If Err.Number = 0 Then
        LogLine "ENGINE" & vbTab & "Word"
    Else
        Err.Clear
        Set app = CreateObject("KWPS.Application")
        If Err.Number = 0 Then
            LogLine "ENGINE" & vbTab & "WPS"
        Else
            Err.Clear
        End If
    End If
    If Not (app Is Nothing) Then
        app.Visible = False
    End If
End If
On Error GoTo 0

If app Is Nothing Then
    LogLine "FATAL" & vbTab & "APP"
    WScript.Quit 2
End If

On Error Resume Next
oldAlerts = app.DisplayAlerts
app.DisplayAlerts = 0
On Error GoTo 0

okCount = 0
failCount = 0
cancelledFlag = False
idx = 0

Dim listStream
On Error Resume Next
Set listStream = fso.OpenTextFile(listPath, 1, False, -1)
If Err.Number <> 0 Then
    LogLine "FATAL" & vbTab & "LIST"
    WScript.Quit 2
End If
On Error GoTo 0

Do Until listStream.AtEndOfStream
    line = listStream.ReadLine
    If Len(line) > 0 Then
        If fso.FileExists(cancelPath) Then
            cancelledFlag = True
            Exit Do
        End If
        parts = Split(line, vbTab)
        If UBound(parts) >= 1 Then
            src = parts(0)
            dst = parts(1)
            LogLine "BEGIN" & vbTab & idx

            Set doc = Nothing
            On Error Resume Next
            Set doc = app.Documents.Open(src, False, True, False)
            If Err.Number <> 0 Or doc Is Nothing Then
                msg = Err.Number
                Err.Clear
                If msg = 0 Then msg = "OPEN_FAILED"
                LogLine idx & vbTab & "ERR" & vbTab & msg
                failCount = failCount + 1
            Else
                On Error Resume Next
                doc.SaveAs2 dst, 16
                If Err.Number <> 0 Then
                    msg = Err.Number
                    Err.Clear
                    doc.SaveAs dst, 16
                End If
                If Err.Number <> 0 Then
                    msg = Err.Number
                    Err.Clear
                    LogLine idx & vbTab & "ERR" & vbTab & msg
                    failCount = failCount + 1
                ElseIf fso.FileExists(dst) Then
                    LogLine idx & vbTab & "OK"
                    okCount = okCount + 1
                Else
                    LogLine idx & vbTab & "ERR" & vbTab & "OUTPUT_MISSING"
                    failCount = failCount + 1
                End If
                Err.Clear
                On Error Resume Next
                doc.Close 0
                Err.Clear
            End If
            On Error GoTo 0
        End If
    End If
    idx = idx + 1
Loop
listStream.Close

On Error Resume Next
If Not preRunning Then
    app.DisplayAlerts = oldAlerts
    app.Quit
Else
    app.DisplayAlerts = oldAlerts
End If
Err.Clear
On Error GoTo 0

If cancelledFlag Then
    LogLine "CANCELLED"
Else
    LogLine "DONE" & vbTab & okCount & vbTab & failCount
End If
WScript.Quit 0
"""


def _com_error_text(text: str) -> str:
    lowered = (text or "").lower()
    if "OUTPUT_MISSING" in text:
        return "Word 没有生成输出文件，请检查源 PDF 是否完整。"
    if "OPEN_FAILED" in text:
        return "Word 无法打开这个 PDF。"
    if "4198" in text:
        return "Word 在执行中被一个弹窗卡住了（常见于复杂 PDF）。请再试一次，或改用「版式还原」档。"
    if "password" in lowered or "密码" in text:
        return "这个 PDF 受密码保护，无法转换。"
    if "-2147023174" in text or "-2147467259" in text:
        return "Office 无法打开这个 PDF（可能已损坏、正在被占用，或是扫描件格式异常）。"
    if "发生意外" in text or "意外" in text or "unexpected" in lowered:
        return "Word 无法处理这个 PDF（可能已损坏、被占用，或格式不受支持）。"
    if text.strip().replace("-", "").replace(" ", "").isdigit():
        return "Word 无法处理这个 PDF（可能已损坏、被占用，或格式不受支持）。"
    if "无法打开" in text or "could not open" in lowered or "not found" in lowered:
        return "Word 无法打开这个 PDF（可能已损坏、被占用，或格式不受支持）。"
    return text or "Word 转换失败"


# ---------------------------------------------------------------------------
# Word's PDF import pops up a modal "NUIDialog" that blocks automation
# ---------------------------------------------------------------------------
_DIALOG_CLASS = "NUIDialog"
_DIALOG_TITLE = "Microsoft Word"
_WM_COMMAND = 0x0111
_IDOK = 1
_BM_CLICK = 0x00F5
_OK_BUTTONS = {"确定", "是", "继续", "OK", "Yes", "Continue"}


def _dismiss_blocking_dialog() -> Optional[str]:
    """Press the accept button on Word's blocking dialogs, if one is showing.

    Word raises modal dialogs while importing a PDF; the automation call then
    fails (error 4198) or freezes until somebody clicks it.  Two shapes are
    handled:

    * ``NUIDialog``   — Office's own frame (single ``NetUIHWND`` child).
    * ``#32770``      — classic Win32 dialog with a 确定 / 取消 button pair;
                        only the accept-side button is ever clicked.

    Ordinary Word prompts are left alone: both shapes must be titled
    "Microsoft Word", and for the classic dialog a button whose caption is an
    accept caption must exist.
    """
    if os.name != "nt":
        return None
    try:
        import ctypes

        user32 = ctypes.windll.user32

        def class_of(hwnd) -> str:
            buffer = ctypes.create_unicode_buffer(256)
            user32.GetClassNameW(hwnd, buffer, 256)
            return buffer.value

        def title_of(hwnd) -> str:
            buffer = ctypes.create_unicode_buffer(256)
            user32.GetWindowTextW(hwnd, buffer, 256)
            return buffer.value.strip()

        def children_of(hwnd) -> list[tuple[int, str, str]]:
            found: list[tuple[int, str, str]] = []

            def enum_child(child, _param):
                found.append((child, class_of(child), title_of(child)))
                return True

            user32.EnumChildWindows(hwnd, _PROC(enum_child), 0)
            return found

        dismissed = None

        def enum_top(hwnd, _param):
            nonlocal dismissed
            kind = class_of(hwnd)
            if kind not in ("NUIDialog", "#32770"):
                return True
            if title_of(hwnd) != _DIALOG_TITLE:
                return True
            kids = children_of(hwnd)
            if kind == "NUIDialog":
                if len(kids) == 1 and kids[0][1] == "NetUIHWND":
                    user32.PostMessageW(hwnd, _WM_COMMAND, _IDOK, 0)
                    dismissed = kind
            else:
                for child, child_class, child_text in kids:
                    if child_class != "Button":
                        continue
                    caption = child_text.replace("&", "").strip()
                    if caption in _OK_BUTTONS:
                        user32.SendMessageW(child, _BM_CLICK, 0, 0)
                        dismissed = kind
                        break
            return True

        user32.EnumWindows(_PROC(enum_top), 0)
        return dismissed
    except Exception:
        return None


if os.name == "nt":
    import ctypes as _ctypes

    _PROC = _ctypes.WINFUNCTYPE(_ctypes.c_bool, _ctypes.c_void_p, _ctypes.c_void_p)
else:  # pragma: no cover - keeps the module importable off Windows
    _PROC = None


# ---------------------------------------------------------------------------
# Page-image engine (no Office required)
# ---------------------------------------------------------------------------
def _render_page(page: pymupdf.Page, dpi: float, quality: int) -> bytes:
    pix = page.get_pixmap(dpi=int(round(dpi)), alpha=False)
    mode = "L" if pix.n == 1 else "RGB"
    image = Image.frombytes(mode, (pix.width, pix.height), pix.samples)
    if image.mode not in ("RGB", "L"):
        image = image.convert("RGB")
    buffer = io.BytesIO()
    image.save(buffer, "JPEG", quality=quality, optimize=True, subsampling=0)
    return buffer.getvalue()


def _pdf_page_count(path) -> int:
    try:
        with pymupdf.open(str(path)) as doc:
            return doc.page_count
    except Exception:
        return 0


def build_docx_from_pages(pdf_path: str | Path, dst: str | Path, dpi: float,
                          quality: int = 92,
                          on_progress: Optional[Callable[[float], None]] = None,
                          cancel_event=None) -> int:
    """Render every page of *pdf_path* into a one-picture-per-page .docx."""
    _ensure_pdf(pdf_path)
    document = Document()
    source = pymupdf.open(str(pdf_path))
    try:
        first = source[0].rect
        section = document.sections[0]
        section.page_width = Emu(int(round(first.width * POINTS_PER_EMU)))
        section.page_height = Emu(int(round(first.height * POINTS_PER_EMU)))
        for attribute in ("left_margin", "right_margin", "top_margin", "bottom_margin"):
            setattr(section, attribute, Emu(0))
        section.header_distance = Emu(0)
        section.footer_distance = Emu(0)

        total = source.page_count
        for index, page in enumerate(source, 1):
            if cancel_event is not None and cancel_event.is_set():
                raise EngineError("已取消")
            data = _render_page(page, dpi, quality)
            paragraph = document.add_paragraph()
            fmt = paragraph.paragraph_format
            fmt.space_before = Pt(0)
            fmt.space_after = Pt(0)
            fmt.line_spacing = 1.0
            run = paragraph.add_run()
            run.add_picture(io.BytesIO(data),
                            width=Emu(int(round(page.rect.width * POINTS_PER_EMU))),
                            height=Emu(int(round(page.rect.height * POINTS_PER_EMU))))
            if index < total:
                run.add_break(WD_BREAK.PAGE)
            if on_progress is not None:
                try:
                    on_progress(index / total)
                except Exception:
                    pass
    finally:
        source.close()
    target = Path(dst)
    tmp = target.with_name(f".~pdf2office-{uuid.uuid4().hex}.docx")
    try:
        document.save(str(tmp))
        os.replace(tmp, target)
    finally:
        try:
            tmp.unlink(missing_ok=True)
        except Exception:
            pass
    return total


def _looks_scanned(path: str | Path) -> bool:
    """True when a PDF has (almost) no extractable text — a scan/photo PDF."""
    try:
        with pymupdf.open(str(path)) as doc:
            pages = min(5, doc.page_count)
            if pages == 0:
                return False
            chars = sum(len(doc[i].get_text().strip()) for i in range(pages))
            return chars < 12 * pages
    except Exception:
        return False


def _ensure_pdf(path: str | Path) -> None:
    """Reject anything that is not really a PDF.

    PyMuPDF can also open Office documents (MuPDF renders docx/xlsx/pptx), and
    it would happily "convert" a .pptx into a one-page Word file.  Checking the
    extension *and* the document type keeps that from happening silently.
    """
    suffix = Path(path).suffix.lower()
    if suffix != ".pdf":
        raise EngineError(f"这不是 PDF 文件（扩展名 {suffix or '无'}）。")
    try:
        with pymupdf.open(str(path)) as doc:
            if not doc.is_pdf:
                raise EngineError("这个文件不是 PDF。")
            if doc.needs_pass:
                raise EngineError("这个 PDF 有密码保护，请先解密。")
    except EngineError:
        raise
    except Exception as exc:
        raise EngineError(f"无法打开这个 PDF：{exc}") from exc


# ---------------------------------------------------------------------------
# Batch driver
# ---------------------------------------------------------------------------
def _unique_path(folder: Path, stem: str, overwrite: bool) -> Path:
    target = folder / f"{stem}.docx"
    if overwrite or not target.exists():
        return target
    index = 1
    while True:
        candidate = folder / f"{stem} ({index}).docx"
        if not candidate.exists():
            return candidate
        index += 1


class BatchConverter:
    """Runs a batch of PDF → DOCX conversions and reports through ``emit``."""

    def __init__(
        self,
        files: list[str],
        mode: str,
        outdir: Optional[str],
        overwrite: bool,
        cancel_event,
        emit: Callable[[str, dict], None],
    ) -> None:
        self.files = list(files)
        self.mode = mode if mode in MODES else DEFAULT_MODE
        self.preset = MODES[self.mode]
        self.engine = self.preset.get("engine", "word")
        self.outdir = outdir
        self.overwrite = overwrite
        self.cancel_event = cancel_event
        self.emit = emit
        self.engine_name = ""

    # -- helpers ---------------------------------------------------------
    def _targets(self) -> list[tuple[Path, Path]]:
        targets = []
        for src in self.files:
            src_path = Path(src)
            folder = Path(self.outdir) if self.outdir else src_path.parent
            folder.mkdir(parents=True, exist_ok=True)
            dst = _unique_path(folder, src_path.stem, self.overwrite)
            # never write over the file we are reading (e.g. a mis-typed input)
            try:
                if os.path.normcase(str(dst.resolve())) == os.path.normcase(str(src_path.resolve())):
                    dst = _unique_path(folder, f"{src_path.stem} (转换)", self.overwrite)
            except Exception:
                pass
            targets.append((src_path, dst))
        return targets

    def _progress(self, index: int, fraction: float) -> None:
        self.emit("progress", {"index": index, "fraction": fraction})

    def _finish_image(self, index: int, src: Path, dst: Path, started: float) -> FileResult:
        result = FileResult(src=str(src))
        try:
            pages = build_docx_from_pages(
                src, dst, float(self.preset.get("dpi", 300)),
                int(self.preset.get("quality", 92)),
                on_progress=lambda fraction: self._progress(index, fraction),
                cancel_event=self.cancel_event,
            )
            result.pages = pages
            result.dst = str(dst)
            result.ok = pages > 0
        except EngineError as exc:
            result.error = str(exc)
            result.cancelled = "取消" in str(exc)
        except Exception as exc:
            result.error = f"{exc.__class__.__name__}: {exc}"
        finally:
            result.seconds = time.time() - started
            try:
                result.size_before = src.stat().st_size
                if dst.exists():
                    result.size_after = dst.stat().st_size
            except Exception:
                pass
            if result.ok and _looks_scanned(src):
                result.note = "扫描件：得到整页图片"
        return result

    def run(self) -> None:
        if self.engine == "image":
            self._run_image()
        else:
            self._run_word()

    # -- page-image engine ----------------------------------------------
    def _run_image(self) -> None:
        stage_text = self.preset.get("stage") or "正在还原版面…"
        try:
            targets = self._targets()
            results: list[FileResult] = []
            cancelled = False
            started_at = time.time()
            for index, (src, dst) in enumerate(targets):
                if self.cancel_event.is_set():
                    cancelled = True
                    results.append(FileResult(src=str(src), cancelled=True, error="已取消"))
                    self.emit("file_cancelled", {"index": index,
                                                 "result": results[-1]})
                    continue
                self.emit("start", {"index": index, "name": src.name})
                self.emit("stage", {"text": stage_text})
                result = self._finish_image(index, src, dst, started_at)
                results.append(result)
                if result.ok:
                    self.emit("file_ok", {"index": index, "result": result})
                elif result.cancelled:
                    cancelled = True
                    self.emit("file_cancelled", {"index": index, "result": result})
                else:
                    self.emit("file_error", {"index": index, "result": result})
            self.emit("done", {"results": results, "cancelled": cancelled,
                               "engine": "PyMuPDF"})
        except Exception as exc:
            self.emit("fatal", {"text": f"发生意外错误：{exc}"})

    # -- Word engine -----------------------------------------------------
    def _run_word(self) -> None:
        work_dir = Path(tempfile.mkdtemp(prefix="pdf2office-"))
        log_path = work_dir / "log.txt"
        list_path = work_dir / "list.txt"
        cancel_path = work_dir / "cancel.flag"
        vbs_path = work_dir / "convert.vbs"
        proc = None
        stage_text = self.preset.get("stage") or "正在重排版式…"

        try:
            self.emit("stage", {"text": "正在启动 Word 引擎…"})
            targets = self._targets()
            slots: list[Optional[FileResult]] = [None] * len(targets)
            # Word would happily open a .pptx/.docx and "convert" it, so bad
            # inputs are rejected here — with the original file index kept for
            # the progress events the GUI relies on.
            sent: list[int] = []
            for index, (src, dst) in enumerate(targets):
                try:
                    _ensure_pdf(src)
                except EngineError as exc:
                    try:
                        size = src.stat().st_size
                    except Exception:
                        size = 0
                    outcome = FileResult(src=str(src), error=str(exc),
                                         size_before=size, seconds=0.0)
                    slots[index] = outcome
                    self.emit("file_error", {"index": index, "result": outcome})
                    continue
                sent.append(index)

            if not sent:
                self.emit("done", {"results": [s for s in slots if s],
                                   "cancelled": False, "engine": self.engine_name})
                return

            vbs_path.write_text(VBS_WORD, encoding="ascii")
            with list_path.open("w", encoding="utf-16") as fh:
                for index in sent:
                    src, dst = targets[index]
                    fh.write(f"{src}\t{dst}\n")

            cscript = os.path.join(os.environ.get("SystemRoot", r"C:\Windows"),
                                   "System32", "cscript.exe")
            if not os.path.exists(cscript):
                cscript = "cscript.exe"
            child_env = os.environ.copy()
            child_env["PATH"] = ";".join(
                part for part in child_env.get("PATH", "").split(";")
                if part and "_internal" not in part.lower()
            )
            for key in ("PYTHONHOME", "PYTHONPATH", "_MEIPASS", "_MEIPASS2"):
                child_env.pop(key, None)
            err_path = work_dir / "helper.err"
            helper_err = open(err_path, "wb")
            try:
                proc = subprocess.Popen(
                    [cscript, "//nologo", "//B", str(vbs_path), str(list_path),
                     str(log_path), str(cancel_path)],
                    stdout=subprocess.DEVNULL, stderr=helper_err,
                    creationflags=_NO_WINDOW_FLAGS, startupinfo=_HIDDEN_STARTUP,
                    env=child_env,
                )
            finally:
                helper_err.close()

            handled: set[int] = set()
            started_at = time.time()
            line_cursor = 0
            last_progress = time.time()
            finished = False
            cancelled = False
            fatal = ""
            exit_checks = 0
            dialog_check = 0.0
            dialog_noticed = False

            while True:
                if self.cancel_event.is_set():
                    try:
                        cancel_path.touch()
                    except Exception:
                        pass

                lines: list[str] = []
                try:
                    if log_path.exists():
                        content = log_path.read_text(encoding="ascii", errors="replace")
                        content = content.replace("\r\n", "\n").replace("\r", "\n")
                        lines = content.split("\n")
                        if lines and lines[-1] != "":
                            lines.pop()
                except Exception:
                    lines = []

                if line_cursor < len(lines):
                    last_progress = time.time()
                while line_cursor < len(lines):
                    parts = lines[line_cursor].split("\t")
                    line_cursor += 1
                    if not parts:
                        continue
                    head = parts[0]
                    if head == "START":
                        continue
                    if head == "ENGINE":
                        self.engine_name = parts[1] if len(parts) > 1 else ""
                        self.emit("stage", {"text": f"正在使用 {self.engine_name} 引擎…"})
                        continue
                    if head == "FATAL":
                        code = parts[1] if len(parts) > 1 else ""
                        fatal = {
                            "APP": "没有找到可用的 Microsoft Word / WPS，"
                                   "请安装 Office 或改用「版式还原」档。",
                            "LIST": "无法读取转换任务列表，请重试。",
                        }.get(code, "Word 启动失败")
                        continue
                    if head == "CANCELLED":
                        cancelled = True
                        continue
                    if head == "DONE":
                        finished = True
                        continue
                    if head == "BEGIN":
                        if len(parts) > 1:
                            try:
                                vbs_index = int(parts[1])
                            except ValueError:
                                continue
                            if vbs_index < len(sent):
                                file_index = sent[vbs_index]
                                self.emit("start", {"index": file_index,
                                                    "name": targets[file_index][0].name})
                        continue
                    try:
                        vbs_index = int(head)
                    except ValueError:
                        continue
                    if vbs_index in handled or vbs_index >= len(sent):
                        continue
                    handled.add(vbs_index)
                    file_index = sent[vbs_index]
                    src, dst = targets[file_index]
                    if len(parts) > 1 and parts[1] == "OK":
                        self.emit("compress", {"index": file_index, "text": stage_text})
                        result = FileResult(
                            src=str(src), dst=str(dst), ok=True,
                            pages=_pdf_page_count(src),
                            size_before=src.stat().st_size if src.exists() else 0,
                            size_after=dst.stat().st_size if dst.exists() else 0,
                            seconds=time.time() - started_at,
                        )
                        if _looks_scanned(src):
                            result.note = "扫描件：只得到图片"
                        slots[file_index] = result
                        self.emit("file_ok", {"index": file_index, "result": result})
                    else:
                        message = parts[2] if len(parts) > 2 else "转换失败"
                        result = FileResult(
                            src=str(src),
                            error=_com_error_text(message),
                            seconds=time.time() - started_at,
                        )
                        try:
                            if dst.exists() and dst.stat().st_size == 0:
                                dst.unlink(missing_ok=True)
                        except Exception:
                            pass
                        slots[file_index] = result
                        self.emit("file_error", {"index": file_index, "result": result})

                if finished or cancelled or fatal:
                    break
                # Word's PDF import raises a modal prompt that freezes the
                # automation call; answer it for the user.
                if time.time() - dialog_check > 0.8:
                    dialog_check = time.time()
                    if _dismiss_blocking_dialog():
                        last_progress = time.time()
                        if not dialog_noticed:
                            dialog_noticed = True
                            self.emit("stage",
                                      {"text": "已自动确认 Word 的导入提示…"})
                if proc.poll() is not None:
                    exit_checks += 1
                    if exit_checks >= 4:
                        break
                    time.sleep(0.3)
                    continue
                # Word gives no progress while it reflows a document, and a
                # complex 100-page PDF can take many minutes.  Only treat it as
                # hung after a very long silence; the 停止 button still works
                # between files.
                if time.time() - last_progress > 1800:
                    fatal = "Word 长时间没有响应，转换已停止。"
                    try:
                        proc.terminate()
                    except Exception:
                        pass
                    break
                time.sleep(0.15)

            if proc is not None:
                try:
                    proc.wait(timeout=10)
                except Exception:
                    try:
                        proc.terminate()
                    except Exception:
                        pass

            if fatal:
                raise EngineError(fatal)

            # Fallback: a real .docx on disk means the conversion succeeded.
            for idx, (src, dst) in enumerate(targets):
                if slots[idx] is not None:
                    continue
                try:
                    if dst.exists() and dst.stat().st_size > 0:
                        result = FileResult(
                            src=str(src), dst=str(dst), ok=True,
                            pages=_pdf_page_count(src),
                            size_before=src.stat().st_size if src.exists() else 0,
                            size_after=dst.stat().st_size,
                            seconds=time.time() - started_at,
                        )
                        slots[idx] = result
                        self.emit("file_ok", {"index": idx, "result": result})
                except Exception:
                    pass

            detail = ""
            if not cancelled:
                try:
                    code = proc.returncode if proc is not None else None
                except Exception:
                    code = None
                tail = ""
                for candidate in (work_dir / "helper.err", log_path):
                    try:
                        if candidate.exists():
                            text = candidate.read_text(encoding="ascii",
                                                       errors="replace").strip()
                            if text:
                                tail += f" {text[-200:]}"
                    except Exception:
                        pass
                detail = (f"（引擎退出码 {code}"
                          f"{('；' + tail.strip()) if tail.strip() else ''}）")

            for idx, (src, dst) in enumerate(targets):
                if slots[idx] is not None:
                    continue
                outcome = FileResult(
                    src=str(src), cancelled=cancelled,
                    error=("已取消" if cancelled
                           else "转换进程提前结束，文件未完成转换。" + detail),
                    seconds=time.time() - started_at,
                )
                slots[idx] = outcome
                if cancelled:
                    self.emit("file_cancelled", {"index": idx, "result": outcome})
                else:
                    self.emit("file_error", {"index": idx, "result": outcome})

            ordered = [item for item in slots if item is not None]
            self.emit("done", {"results": ordered, "cancelled": cancelled,
                               "engine": self.engine_name})
        except EngineError as exc:
            self.emit("fatal", {"text": str(exc)})
        except Exception as exc:
            self.emit("fatal", {"text": f"发生意外错误：{exc}"})
        finally:
            if proc is not None and proc.poll() is None:
                try:
                    proc.terminate()
                except Exception:
                    pass
            shutil.rmtree(work_dir, ignore_errors=True)
