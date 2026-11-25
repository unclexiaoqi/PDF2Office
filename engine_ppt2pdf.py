"""PPT → PDF conversion engine.

Pipeline
--------
1. The slide export is delegated to a small VBScript driven by ``cscript``.  It
   uses the installed Microsoft PowerPoint (COM) for maximum layout fidelity
   and keeps PowerPoint automation out of the GUI process.
2. ``optimize_pdf`` post-processes the exported PDF according to the chosen
   quality tier.

Why post-processing exists — measured behaviour of PowerPoint
-------------------------------------------------------------
PowerPoint's own PDF writer downsamples **every** embedded raster image to
200 DPI relative to the size it occupies on the slide, regardless of the export
route: ``SaveAs(ppSaveAsPDF)``, ``ExportAsFixedFormat(Intent=Print)`` and
``ExportAsFixedFormat(Intent=Screen)`` were all measured (only the screen
intent differs, at 96 DPI).  A 378 MB deck whose pictures are 6000×4500 px
therefore leaves PowerPoint as a ~12 MB PDF.  The loss happens *inside*
PowerPoint, not here — no amount of downstream care can undo it.

That is why the "ultra" tier does something different: it reads the original
pictures straight out of the ``.pptx`` package and writes them back into the
PDF at the very same placement, undoing PowerPoint's downsampling while text,
tables and charts stay vector.  Each replacement is guarded by three
independent checks (slot geometry, cropped aspect ratio, thumbnail
correlation); anything that cannot be matched with confidence is left exactly
as PowerPoint exported it.

Measuring the exported PDF showed PowerPoint has a second, quieter failure
mode: for some slides it keeps the full pixel count but re-encodes the picture
at quality ~75 (real decks: 4:2:0 chroma, quantisation tables ~20 points below
the rest of the document).  Nothing is downsampled, so a "restore the missing
pixels" rule never fires and the file silently stays blurry.  The tier
therefore also reads the quantisation tables out of every embedded JPEG and
re-encodes the ones PowerPoint squeezed harder than we would.  When the deck's
picture is a JPEG that can be dropped in untouched, it is embedded byte for
byte instead — zero generation loss, and smaller than re-encoding it.
"""

from __future__ import annotations

import io
import os
import shutil
import subprocess
import sys
import tempfile
import time
import uuid
import zipfile
from dataclasses import dataclass
from pathlib import Path
from typing import Callable, Optional
from xml.etree import ElementTree

import pikepdf
import pymupdf
from PIL import Image, ImageChops, ImageStat

MODES = {
    "native": {
        "name": "原图无损",
        "short": "原始像素",
        "desc": "逐图实测：画质不低于 WPS / PowerPoint 类导出；能直嵌原图字节就不重压",
        "tag": "图片 100% + 文字矢量",
        "badge": "推荐",
        "stage": "正在写入原图…",
        "mode": "vector",
        "cap": None,          # never shrink
        "upgrade": True,      # restore original picture resolution
        "max_dpi": None,      # no ceiling at all — 1:1 original pixels
        "quality": 95,
        "subsampling": 0,     # 4:4:4 — no chroma smear
    },
    "render": {
        "name": "整页直出",
        "short": "400 DPI",
        "desc": "完全不经过 PPT 的 PDF 导出，逐页渲染成图",
        "tag": "屏幕效果完全还原",
        "stage": "正在逐页渲染…",
        "mode": "render",
        "render_dpi": 400,
        "quality": 95,
        "subsampling": 0,
    },
    "high": {
        "name": "高清",
        "short": "PPT 直出",
        "desc": "PowerPoint 原始输出，不做二次压缩（约 200 DPI）",
        "tag": "转换最快 · 不经重编码",
        "stage": "正在校验…",
        "mode": "vector",
        "cap": None,
        "upgrade": False,
    },
    "medium": {
        "name": "标准",
        "short": "200 DPI",
        "desc": "图片上限 200 DPI，等同 PowerPoint 最高画质",
        "tag": "适合邮件 / 微信发送",
        "stage": "正在压缩体积…",
        "mode": "vector",
        "cap": 200,
        "upgrade": False,
        "quality": 90,
        "subsampling": 0,
    },
    "low": {
        "name": "小体积",
        "short": "80 DPI",
        "desc": "高压缩，优先减小体积",
        "tag": "适合快速预览",
        "stage": "正在压缩体积…",
        "mode": "vector",
        "cap": 80,
        "upgrade": False,
        "quality": 70,
        "subsampling": 2,
    },
}

MODE_ORDER = ["native", "render", "medium", "low"]
DEFAULT_MODE = "native"
SUPPORTED_EXTS = {".ppt", ".pptx", ".pptm", ".pps", ".ppsx", ".potx", ".potm"}
_ZIP_BASED = {".pptx", ".pptm", ".ppsx", ".potx", ".potm"}

# Open XML namespaces (stdlib ElementTree keeps this dependency-free).
_NS_P = "{http://schemas.openxmlformats.org/presentationml/2006/main}"
_NS_A = "{http://schemas.openxmlformats.org/drawingml/2006/main}"
_NS_R = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}"
EMU_PER_PT = 12700.0

#: Geometry tolerance (points) when matching a PDF image slot to a PPT picture.
_SLOT_TOLERANCE_PT = 2.0
#: A source picture must offer at least this much more detail to be worth it.
_UPGRADE_MIN_GAIN = 1.15
#: Maximum mean grey difference (0-255) between the PDF image and its source.
_THUMB_TOLERANCE = 12.0
#: ITU T.81 Annex K tables — needed to read the quality back out of a JPEG.
_STD_LUMA = (
    16, 11, 10, 16, 24, 40, 51, 61, 12, 12, 14, 19, 26, 58, 60, 55,
    14, 13, 16, 24, 40, 57, 69, 56, 14, 17, 22, 29, 51, 87, 80, 62,
    18, 22, 37, 56, 68, 109, 103, 77, 24, 35, 55, 64, 81, 104, 113, 92,
    49, 64, 78, 87, 103, 121, 120, 101, 72, 92, 95, 98, 112, 100, 103, 99,
)

#: Set ``PPT2PDF_DEBUG=1`` to explain every picture decision on stderr.
_DEBUG = bool(os.environ.get("PPT2PDF_DEBUG"))


def _trace(message: str) -> None:
    if _DEBUG:
        try:
            print(f"[ppt2pdf] {message}", file=sys.stderr)
        except Exception:
            pass


class EngineError(RuntimeError):
    pass


# ---------------------------------------------------------------------------
# Running the console helper without flashing a console window
# ---------------------------------------------------------------------------
#: ``cscript.exe`` is a console program.  Started from a windowed (GUI)
#: application it has no console to attach to, so Windows creates a new one —
#: the black box that used to flash on every conversion.  CREATE_NO_WINDOW
#: suppresses that console entirely; the STARTUPINFO below is the belt-and-
#: braces fallback for shells that still honour SW_HIDE.
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
# VBScript export helper (ASCII only — paths are passed as arguments)
# ---------------------------------------------------------------------------
VBS_EXPORT = r"""
Option Explicit

Dim fso, listPath, logPath, cancelPath, app, pres, line, parts, src, dst, idx, msg
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
Set app = GetObject("", "PowerPoint.Application")
If Err.Number = 0 Then
    preRunning = True
    LogLine "ENGINE" & vbTab & "PowerPoint"
Else
    Err.Clear
    Set app = CreateObject("PowerPoint.Application")
    If Err.Number = 0 Then
        LogLine "ENGINE" & vbTab & "PowerPoint"
    Else
        Err.Clear
        Set app = CreateObject("KWPP.Application")
        If Err.Number = 0 Then
            LogLine "ENGINE" & vbTab & "WPS"
        Else
            Err.Clear
        End If
    End If
    If Not (app Is Nothing) Then
        app.Visible = True
        app.WindowState = 2
    End If
End If
On Error GoTo 0

If app Is Nothing Then
    LogLine "FATAL" & vbTab & "APP"
    WScript.Quit 2
End If

On Error Resume Next
oldAlerts = app.DisplayAlerts
app.DisplayAlerts = 1
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

            Set pres = Nothing
            On Error Resume Next
            Set pres = app.Presentations.Open(src, True, False, False)
            If Err.Number <> 0 Then
                msg = Err.Number
                Err.Clear
                LogLine idx & vbTab & "ERR" & vbTab & msg
                failCount = failCount + 1
            Else
                On Error Resume Next
                pres.SaveAs dst, 32
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
                pres.Close
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


VBS_RENDER = r"""
Option Explicit

Dim fso, listPath, logPath, cancelPath, app, pres, line, parts, src, dst, imgdir
Dim idx, msg, preRunning, oldAlerts, okCount, failCount, cancelledFlag, dpi
Dim wPx, hPx, i, slideCount, png, ratio, tbl, line2

Set fso = CreateObject("Scripting.FileSystemObject")
listPath = WScript.Arguments(0)
logPath = WScript.Arguments(1)
cancelPath = WScript.Arguments(2)
dpi = CDbl(WScript.Arguments(3))

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
Set app = GetObject("", "PowerPoint.Application")
If Err.Number = 0 Then
    preRunning = True
    LogLine "ENGINE" & vbTab & "PowerPoint"
Else
    Err.Clear
    Set app = CreateObject("PowerPoint.Application")
    If Err.Number = 0 Then
        LogLine "ENGINE" & vbTab & "PowerPoint"
    Else
        Err.Clear
        Set app = CreateObject("KWPP.Application")
        If Err.Number = 0 Then
            LogLine "ENGINE" & vbTab & "WPS"
        Else
            Err.Clear
        End If
    End If
    If Not (app Is Nothing) Then
        app.Visible = True
        app.WindowState = 2
    End If
End If
On Error GoTo 0

If app Is Nothing Then
    LogLine "FATAL" & vbTab & "APP"
    WScript.Quit 2
End If

On Error Resume Next
oldAlerts = app.DisplayAlerts
app.DisplayAlerts = 1
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
        If UBound(parts) >= 2 Then
            src = parts(0)
            dst = parts(1)
            imgdir = parts(2)
            LogLine "BEGIN" & vbTab & idx

            Set pres = Nothing
            On Error Resume Next
            Set pres = app.Presentations.Open(src, True, False, False)
            If Err.Number <> 0 Then
                msg = Err.Number
                Err.Clear
                LogLine idx & vbTab & "ERR" & vbTab & msg
                failCount = failCount + 1
            Else
                On Error Resume Next
                wPx = CLng(pres.PageSetup.SlideWidth / 72 * dpi)
                hPx = CLng(pres.PageSetup.SlideHeight / 72 * dpi)
                If wPx < 1 Then wPx = 1
                If hPx < 1 Then hPx = 1
                If wPx > 14000 Then
                    ratio = 14000 / wPx
                    wPx = 14000
                    hPx = CLng(hPx * ratio)
                End If
                If hPx > 14000 Then hPx = 14000
                LogLine idx & vbTab & "SIZE" & vbTab & pres.PageSetup.SlideWidth _
                    & vbTab & pres.PageSetup.SlideHeight
                slideCount = pres.Slides.Count
                Err.Clear

                For i = 1 To slideCount
                    If fso.FileExists(cancelPath) Then
                        cancelledFlag = True
                        Exit For
                    End If
                    png = imgdir & "\p" & Right("0000" & CStr(i), 4) & ".png"
                    pres.Slides(i).Export png, "PNG", wPx, hPx
                    If Err.Number <> 0 Then
                        msg = Err.Number
                        Err.Clear
                        Exit For
                    End If
                    LogLine idx & vbTab & "PAGE" & vbTab & i & vbTab & slideCount
                Next

                If Err.Number <> 0 Then
                    msg = Err.Number
                    Err.Clear
                    LogLine idx & vbTab & "ERR" & vbTab & msg
                    failCount = failCount + 1
                ElseIf cancelledFlag Then
                    LogLine idx & vbTab & "ERR" & vbTab & "CANCELLED"
                    failCount = failCount + 1
                Else
                    LogLine idx & vbTab & "OK"
                    okCount = okCount + 1
                End If
                Err.Clear
                On Error Resume Next
                pres.Close
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
        return "PowerPoint 没有生成输出文件，请检查源文件是否完整。"
    if "LIST_READ_ERROR" in text:
        return "无法读取转换任务列表，请重试。"
    if "password" in lowered or "密码" in text:
        return "演示文稿受密码保护，无法直接转换。"
    if "-2134163454" in text:
        return "PowerPoint 无法打开这个文件（可能已损坏，或不是有效的演示文稿）。"
    if "-2147023174" in text:
        return "PowerPoint 没有正常启动，请稍后重试或重启电脑。"
    if "发生意外" in text or "意外" in text or "unexpected" in lowered:
        return "PowerPoint 无法打开这个文件（可能已损坏、被占用，或格式不受支持）。"
    if text.strip().replace("-", "").replace(" ", "").isdigit():
        return "PowerPoint 无法打开这个文件（可能已损坏、被占用，或格式不受支持）。"
    if "无法打开" in text or "could not open" in lowered or "not found" in lowered:
        return "PowerPoint 无法打开这个文件（可能已损坏、被占用，或格式不受支持）。"
    return text or "PowerPoint 转换失败"


# ---------------------------------------------------------------------------
# Source picture index (.pptx → full-resolution pixels)
# ---------------------------------------------------------------------------
@dataclass(frozen=True)
class _Picture:
    """One picture placement on one slide."""

    page: int
    rect: tuple[float, float, float, float]      # x, y, width, height in points
    part: str                                    # zip member of the image part
    crop: tuple[float, float, float, float]      # l, t, r, b as 0..1 fractions


def _read_rels(archive: zipfile.ZipFile, part: str) -> dict[str, str]:
    folder, name = part.rsplit("/", 1) if "/" in part else ("", part)
    rels_name = f"{folder}/_rels/{name}.rels" if folder else f"_rels/{name}.rels"
    try:
        root = ElementTree.fromstring(archive.read(rels_name))
    except Exception:
        return {}
    targets: dict[str, str] = {}
    for rel in root:
        rid = rel.get("Id")
        target = rel.get("Target")
        if not rid or not target:
            continue
        if target.startswith("/"):
            resolved = target.lstrip("/")
        else:
            resolved = f"{folder}/{target}" if folder else target
        # normalise ../media/x.png style targets
        parts: list[str] = []
        for chunk in resolved.split("/"):
            if chunk in ("", "."):
                continue
            if chunk == "..":
                if parts:
                    parts.pop()
                continue
            parts.append(chunk)
        targets[rid] = "/".join(parts)
    return targets


def _slide_parts(archive: zipfile.ZipFile) -> list[str]:
    """Slide parts in presentation order (not filename order — slides get reordered)."""
    try:
        pres = ElementTree.fromstring(archive.read("ppt/presentation.xml"))
        rels = _read_rels(archive, "ppt/presentation.xml")
        listing = pres.find(f"{_NS_P}sldIdLst")
        ordered: list[str] = []
        if listing is not None:
            for slide in listing.findall(f"{_NS_P}sldId"):
                rid = slide.get(f"{_NS_R}id")
                if rid and rid in rels:
                    ordered.append(rels[rid])
        if ordered:
            return ordered
    except Exception:
        pass

    def order(name: str) -> int:
        digits = "".join(ch for ch in name.rsplit("/", 1)[-1] if ch.isdigit())
        return int(digits) if digits else 0

    return sorted(
        (item for item in archive.namelist()
         if item.startswith("ppt/slides/slide") and item.endswith(".xml")),
        key=order,
    )


def _walk_shapes(element, base_x, base_y, kx, ky, rels, page, out) -> None:
    for child in element:
        tag = child.tag
        if tag == f"{_NS_P}grpSp":
            next_base_x, next_base_y = base_x, base_y
            next_kx, next_ky = kx, ky
            xfrm = child.find(f".//{_NS_A}xfrm")
            if xfrm is not None:
                off = xfrm.find(f"{_NS_A}off")
                ext = xfrm.find(f"{_NS_A}ext")
                ch_off = xfrm.find(f"{_NS_A}chOff")
                ch_ext = xfrm.find(f"{_NS_A}chExt")
                if None not in (off, ext, ch_off, ch_ext):
                    try:
                        group_w, group_h = int(ch_ext.get("cx")), int(ch_ext.get("cy"))
                        box_w, box_h = int(ext.get("cx")), int(ext.get("cy"))
                        if group_w and group_h:
                            origin_x = base_x + int(off.get("x")) * kx
                            origin_y = base_y + int(off.get("y")) * ky
                            next_kx = kx * box_w / group_w
                            next_ky = ky * box_h / group_h
                            next_base_x = origin_x - int(ch_off.get("x")) * next_kx
                            next_base_y = origin_y - int(ch_off.get("y")) * next_ky
                    except (TypeError, ValueError):
                        pass
            _walk_shapes(child, next_base_x, next_base_y, next_kx, next_ky,
                         rels, page, out)
            continue

        if tag not in (f"{_NS_P}sp", f"{_NS_P}pic", f"{_NS_P}cxnSp"):
            continue

        xfrm = child.find(f".//{_NS_A}xfrm")
        if xfrm is None:
            continue
        off, ext = xfrm.find(f"{_NS_A}off"), xfrm.find(f"{_NS_A}ext")
        if off is None or ext is None:
            continue
        try:
            x = base_x + int(off.get("x")) * kx
            y = base_y + int(off.get("y")) * ky
            width = int(ext.get("cx")) * kx
            height = int(ext.get("cy")) * ky
        except (TypeError, ValueError):
            continue
        if width <= 0 or height <= 0:
            continue

        # Slides use p:blipFill for <p:pic>, a:blipFill for picture-filled shapes.
        for fill in child.iter():
            if not fill.tag.endswith("}blipFill"):
                continue
            blip = fill.find(f"{_NS_A}blip")
            if blip is None:
                continue
            rid = blip.get(f"{_NS_R}embed")
            if not rid or rid not in rels:
                continue
            crop = (0.0, 0.0, 0.0, 0.0)
            src_rect = fill.find(f"{_NS_A}srcRect")
            if src_rect is not None:
                try:
                    crop = tuple(
                        min(0.95, max(0.0, int(src_rect.get(key, 0)) / 100000.0))
                        for key in ("l", "t", "r", "b")
                    )
                except (TypeError, ValueError):
                    crop = (0.0, 0.0, 0.0, 0.0)
            out.append(_Picture(
                page=page,
                rect=(x / EMU_PER_PT, y / EMU_PER_PT,
                      width / EMU_PER_PT, height / EMU_PER_PT),
                part=rels[rid],
                crop=crop,  # type: ignore[arg-type]
            ))


def _crop_image(image: Image.Image, crop) -> Image.Image:
    left, top, right, bottom = crop
    if not any(crop):
        return image
    width, height = image.size
    x0 = int(round(width * left))
    y0 = int(round(height * top))
    x1 = int(round(width * (1.0 - right)))
    y1 = int(round(height * (1.0 - bottom)))
    x1 = min(width, max(x0 + 1, x1))
    y1 = min(height, max(y0 + 1, y1))
    return image.crop((x0, y0, x1, y1))


def _flatten_cropped(raw: Image.Image, crop) -> Optional[Image.Image]:
    """Crop *raw* and return RGB pixels, or None when real transparency is used.

    Fully opaque alpha channels (extremely common in render exports) are dropped
    safely; pictures that really rely on transparency are left to PowerPoint,
    because the colour PowerPoint composited them against is unknown.
    """
    cropped = _crop_image(raw, crop)
    if cropped.mode == "P":
        cropped = cropped.convert("RGBA")
    if cropped.mode in ("RGBA", "LA", "PA"):
        try:
            if cropped.getchannel("A").getextrema()[0] < 255:
                return None
        except Exception:
            return None
    return cropped.convert("RGB")


class _PptxPictures:
    """Picture placements per slide plus access to the original image pixels."""

    def __init__(self, path: str | Path) -> None:
        self._archive = zipfile.ZipFile(str(path))
        self._pages: dict[int, list[_Picture]] = {}
        self._cache: dict[tuple, Optional[Image.Image]] = {}
        self._raw: Optional[tuple[str, bytes]] = None
        try:
            for page, part in enumerate(_slide_parts(self._archive)):
                try:
                    root = ElementTree.fromstring(self._archive.read(part))
                except Exception:
                    continue
                tree = root.find(f".//{_NS_P}spTree")
                if tree is None:
                    continue
                rels = _read_rels(self._archive, part)
                found: list[_Picture] = []
                _walk_shapes(tree, 0, 0, 1.0, 1.0, rels, page, found)
                if found:
                    self._pages[page] = found
        except Exception:
            self._pages = {}

    def candidates(self, page: int) -> list[_Picture]:
        return self._pages.get(page, [])

    def load(self, picture: _Picture) -> Optional[Image.Image]:
        """Full-resolution RGB pixels of *picture*, already cropped like PowerPoint."""
        key = (picture.part, picture.crop)
        if key in self._cache:
            return self._cache[key]
        image: Optional[Image.Image] = None
        try:
            data = self.source_bytes(picture)
            if data is None:
                raise ValueError("image part missing")
            with Image.open(io.BytesIO(data)) as raw:
                raw.load()
                image = _flatten_cropped(raw, picture.crop)
        except Exception:
            image = None
        if len(self._cache) >= 2:
            self._cache.clear()      # decoded pictures are huge; keep it tiny
        self._cache[key] = image
        return image

    def source_bytes(self, picture: _Picture) -> Optional[bytes]:
        """The untouched bytes of the image part — never decoded, never re-encoded.

        Handing these straight to the PDF is the only way to guarantee zero
        generation loss for pictures PowerPoint already stores lossily.
        """
        try:
            if self._raw is not None and self._raw[0] == picture.part:
                return self._raw[1]
            with self._archive.open(picture.part) as handle:
                data = handle.read()
            self._raw = (picture.part, data)
            return data
        except Exception:
            return None

    def close(self) -> None:
        try:
            self._archive.close()
        except Exception:
            pass


# ---------------------------------------------------------------------------
# PDF image helpers
# ---------------------------------------------------------------------------
def _filter_names(obj) -> list[str]:
    value = obj.get("/Filter")
    if value is None:
        return []
    if isinstance(value, pikepdf.Array):
        return [str(item) for item in value]
    return [str(value)]


def _page_images(doc: pymupdf.Document) -> dict[int, list[tuple[int, pymupdf.Rect]]]:
    """xref → [(page number, placement rect), …] for every image on every page."""
    found: dict[int, list[tuple[int, pymupdf.Rect]]] = {}
    for page_no, page in enumerate(doc):
        try:
            images = page.get_images(full=True)
        except Exception:
            continue
        for image in images:
            xref = image[0]
            try:
                rects = page.get_image_rects(xref)
            except Exception:
                continue
            for rect in rects:
                found.setdefault(xref, []).append((page_no, rect))
    return found


def _decode_flate(obj, doc: pymupdf.Document, xref: int) -> Optional[Image.Image]:
    try:
        pix = pymupdf.Pixmap(doc, xref)
    except Exception:
        return None
    if pix.alpha:
        return None
    if pix.n == 4:
        try:
            pix = pymupdf.Pixmap(pymupdf.csRGB, pix)
        except Exception:
            return None
    if pix.n not in (1, 3):
        return None
    mode = "L" if pix.n == 1 else "RGB"
    try:
        return Image.frombytes(mode, (pix.width, pix.height), pix.samples)
    except Exception:
        return None


def _thumb(image: Image.Image, size: int = 48) -> bytes:
    return image.convert("L").resize((size, size), Image.BILINEAR).tobytes()


def _thumb_distance(left: bytes, right: bytes) -> float:
    if not left or not right or len(left) != len(right):
        return 255.0
    return sum(abs(a - b) for a, b in zip(left, right)) / len(left)


def _pdf_thumb(doc: pymupdf.Document, xref: int, obj) -> Optional[bytes]:
    """Cheap fingerprint of the image PowerPoint embedded (JPEG draft decoding)."""
    try:
        if any("DCTDecode" in name for name in _filter_names(obj)):
            with Image.open(io.BytesIO(obj.read_raw_bytes())) as image:
                image.draft("L", (192, 192))
                image.load()
                return _thumb(image)
        image = _decode_flate(obj, doc, xref)
        if image is None:
            return None
        return _thumb(image)
    except Exception:
        return None


def _find_picture(pictures: _PptxPictures, page: int, rect) -> Optional[_Picture]:
    """Nearest picture placement on *page* for the PDF image slot *rect*."""
    best: Optional[_Picture] = None
    best_distance = _SLOT_TOLERANCE_PT
    for candidate in pictures.candidates(page):
        x, y, width, height = candidate.rect
        distance = (abs(x - rect.x0) + abs(y - rect.y0)
                    + abs(width - rect.width) + abs(height - rect.height))
        if distance < best_distance:
            best, best_distance = candidate, distance
    return best


def _resize_to_dpi(image: Image.Image, width_pt: float, max_dpi: float) -> Image.Image:
    target_width = max(1, int(round(width_pt / 72.0 * max_dpi)))
    if target_width >= image.width:
        return image
    target_height = max(1, int(round(image.height * target_width / image.width)))
    return image.resize((target_width, target_height), Image.LANCZOS)


def _encode_jpeg(image: Image.Image, quality: int, subsampling: int) -> bytes:
    if image.mode not in ("RGB", "L"):
        image = image.convert("RGB")
    buffer = io.BytesIO()
    image.save(buffer, "JPEG", quality=quality, optimize=True,
               subsampling=subsampling)
    return buffer.getvalue()


def _replace_stream(obj, image: Image.Image, data: bytes) -> None:
    obj.write(data, filter=pikepdf.Name.DCTDecode)
    obj.Width = image.width
    obj.Height = image.height
    obj.BitsPerComponent = 8
    obj.ColorSpace = (pikepdf.Name.DeviceRGB if image.mode == "RGB"
                      else pikepdf.Name.DeviceGray)
    for key in ("/DecodeParms", "/Decode", "/Intent", "/Interpolate"):
        if key in obj:
            del obj[key]


# ---------------------------------------------------------------------------
# Reading the compression back out of what PowerPoint wrote
# ---------------------------------------------------------------------------
def _jpeg_quality(data: bytes) -> Optional[tuple[int, int]]:
    """Estimate ``(quality, chroma_subsampling)`` of a JPEG from its headers.

    ``chroma_subsampling`` is 0 for 4:4:4, 1 for 4:2:2, 2 for 4:2:0 and -1 when
    the sampling factors are anything unusual.  The quality figure comes from
    the luma quantisation table (ITU T.81 Annex K), which is exactly how a JPEG
    encoder such as PowerPoint's records its own setting — no pixels are
    decoded, so this costs microseconds even for a 22-megapixel picture.
    """
    if len(data) < 4 or data[:2] != b"\xff\xd8":
        return None
    tables: dict[int, list[int]] = {}
    subsampling = -1
    index, size = 2, len(data)
    try:
        while index < size - 3:
            if data[index] != 0xFF:
                index += 1
                continue
            marker = data[index + 1]
            if marker in (0xD8, 0xD9) or 0xD0 <= marker <= 0xD7:
                index += 2
                continue
            if marker == 0xFF:
                index += 1
                continue
            length = (data[index + 2] << 8) | data[index + 3]
            segment = data[index + 4:index + 2 + length]
            if marker == 0xDB:
                cursor = 0
                while cursor < len(segment):
                    precision, table_id = segment[cursor] >> 4, segment[cursor] & 0x0F
                    cursor += 1
                    span = 128 if precision else 64
                    if cursor + span > len(segment):
                        return None
                    if precision:
                        values = [(segment[cursor + 2 * k] << 8)
                                  | segment[cursor + 2 * k + 1] for k in range(64)]
                    else:
                        values = list(segment[cursor:cursor + 64])
                    tables[table_id] = values
                    cursor += span
            elif marker in (0xC0, 0xC1, 0xC2):
                components = segment[5] if len(segment) > 5 else 0
                if components >= 3 and len(segment) >= 10:
                    luma = segment[7] >> 4, segment[7] & 0x0F
                    subsampling = {(1, 1): 0, (2, 1): 1, (2, 2): 2}.get(luma, -1)
                elif components:
                    subsampling = 0
            elif marker == 0xDA:      # entropy data follows — stop before it
                break
            index += 2 + length
    except Exception:
        return None

    table = tables.get(0)
    if not table:
        return None
    best_quality, best_error = 100, None
    for candidate in range(1, 101):
        scale = 5000 // candidate if candidate < 50 else 200 - candidate * 2
        error = 0
        for standard, actual in zip(_STD_LUMA, table):
            target = min(255, max(1, (standard * scale + 50) // 100))
            error += abs(target - actual)
        if best_error is None or error < best_error:
            best_quality, best_error = candidate, error
    return best_quality, subsampling


def _passthrough_bytes(pictures: _PptxPictures,
                       picture: _Picture) -> Optional[bytes]:
    """Original JPEG bytes, when the whole picture can be replaced as-is."""
    if any(picture.crop):
        return None                     # a crop forces a re-encode anyway
    if Path(picture.part).suffix.lower() not in (".jpg", ".jpeg"):
        return None
    data = pictures.source_bytes(picture)
    if not data:
        return None
    try:
        with Image.open(io.BytesIO(data)) as probe:
            if probe.format != "JPEG" or probe.mode != "RGB":
                return None             # CMYK / greyscale would need /ColorSpace care
    except Exception:
        return None
    return data


# ---------------------------------------------------------------------------
# The "no worse than the other exporter" floor
# ---------------------------------------------------------------------------
#: Mean squared error of a picture, used as the yardstick below.
def _mse(reference: Image.Image, image: Image.Image) -> float:
    if reference.size != image.size or reference.mode != image.mode:
        return float("inf")
    try:
        rms = ImageStat.Stat(ImageChops.difference(reference, image)).rms
    except Exception:
        return float("inf")
    if not rms:
        return float("inf")
    return sum(value * value for value in rms) / len(rms)


def _decode_jpeg(data: bytes) -> Optional[Image.Image]:
    try:
        with Image.open(io.BytesIO(data)) as probe:
            probe.load()
            return probe.convert("RGB")
    except Exception:
        return None


def _exporter_floor(reference: Image.Image) -> float:
    """Error of the encode a WPS-class exporter writes: quality 100, 4:2:0.

    Measured on real decks: WPS's PDF pictures decode back to exactly what
    libjpeg produces for ``quality=100, subsampling=2`` (five pictures matched
    to 0.00 dB).  That makes the floor reproducible here without running WPS —
    whatever we write has to be at least this close to the original.
    """
    data = _encode_jpeg(reference, 100, 2)
    if not data:
        return float("inf")             # no floor: fall back to the tier quality
    decoded = _decode_jpeg(data)
    if decoded is None:
        return float("inf")
    return _mse(reference, decoded)


def _quality_ladder(start: int) -> list[int]:
    """Tier quality first, then the steps that can clear the export floor."""
    values = [max(1, min(100, int(start)))]
    for step in (96, 97, 98, 99, 100):
        if step > values[0]:
            values.append(step)
    return values


def _floor_payload(pictures: _PptxPictures, picture: _Picture, source: Image.Image,
                   quality: int, subsampling: int, max_dpi, width_pt: float,
                   floor: float) -> Optional[tuple[Image.Image, bytes, str, float]]:
    """Smallest payload whose error is no worse than *floor*.

    Candidates are the deck's own JPEG (error 0) and our encodes, tried at the
    tier quality first and escalated only while they still miss the floor.
    """
    options: list[tuple[int, float, str, Image.Image, bytes]] = []

    original = _passthrough_bytes(pictures, picture)
    if original:
        options.append((len(original), 0.0, "verbatim", source, original))

    image = source
    if max_dpi is not None:
        image = _resize_to_dpi(source, width_pt, float(max_dpi))

    best: Optional[tuple[Image.Image, bytes, str, float]] = None
    for level in _quality_ladder(quality):
        data = _encode_jpeg(image, level, subsampling)
        if not data:
            continue
        decoded = _decode_jpeg(data)
        if decoded is None:
            continue
        error = _mse(image, decoded)
        best = (image, data, level, error)
        if error <= floor:
            break
    if best is None:
        return None
    options.append((len(best[1]), best[3], "recode", best[0], best[1]))

    _, error, kind, image, data = min(options, key=lambda item: item[0])
    return image, data, kind, error


# ---------------------------------------------------------------------------
# Post-processing
# ---------------------------------------------------------------------------
def optimize_pdf(src: str | Path, dst: str | Path, preset: dict,
                 pptx_path: str | Path | None = None,
                 on_progress: Optional[Callable[[float], None]] = None) -> dict[str, int]:
    """Write *src* to *dst* applying *preset*.

    ``upgrade``   restores the original picture pixels from *pptx_path*.
    ``cap``       downsamples any image above that DPI (None = never shrink).
    With neither set the file is copied untouched.

    Returns how many pictures were rewritten and how: ``verbatim`` (the
    original file was embedded byte for byte), ``upgraded`` (PowerPoint had
    thrown pixels away and they were put back) and ``repaired`` (the pixels
    were already there, but PowerPoint had compressed them much harder than we
    would).
    """
    src, dst = str(src), str(dst)
    stats = {"verbatim": 0, "upgraded": 0, "repaired": 0}
    cap = preset.get("cap")
    upgrade = bool(preset.get("upgrade"))
    if cap is None and not upgrade:
        shutil.copyfile(src, dst)
        return stats

    quality = int(preset.get("quality", 90))
    subsampling = int(preset.get("subsampling", 0))
    max_dpi = preset.get("max_dpi")

    pictures: Optional[_PptxPictures] = None
    if upgrade and pptx_path and Path(pptx_path).suffix.lower() in _ZIP_BASED:
        try:
            pictures = _PptxPictures(pptx_path)
        except Exception:
            pictures = None

    doc = pymupdf.open(src)
    try:
        layout = _page_images(doc)
        total = max(1, len(layout))
        pdf = pikepdf.open(src)
        try:
            for done, (xref, uses) in enumerate(layout.items(), 1):
                try:
                    _rewrite_image(pdf, doc, xref, uses, cap, upgrade, max_dpi,
                                   quality, subsampling, pictures, stats)
                except Exception:
                    pass
                if on_progress is not None and (done % 4 == 0 or done == total):
                    try:
                        on_progress(done / total)
                    except Exception:
                        pass
            pdf.save(dst, compress_streams=True, recompress_flate=True)
        finally:
            pdf.close()
    finally:
        if pictures is not None:
            pictures.close()
        doc.close()

    # Never hand back a "compressed" file that is not actually smaller.
    if not upgrade and (not os.path.exists(dst) or os.path.getsize(dst) >= os.path.getsize(src)):
        shutil.copyfile(src, dst)
    return stats


def _rewrite_image(pdf, doc, xref: int, uses, cap, upgrade: bool, max_dpi,
                   quality: int, subsampling: int,
                   pictures: Optional[_PptxPictures],
                   stats: Optional[dict[str, int]] = None) -> None:
    try:
        obj = pdf.get_object((xref, 0))
    except Exception:
        return
    if not isinstance(obj, pikepdf.Stream):
        return
    if obj.get("/Subtype") != pikepdf.Name.Image:
        return
    if "/SMask" in obj or "/Mask" in obj:
        return
    try:
        width, height = int(obj.Width), int(obj.Height)
    except Exception:
        return
    if width < 16 or height < 16:
        return

    filters = _filter_names(obj)
    if not any("DCTDecode" in name or "FlateDecode" in name for name in filters):
        return
    stats = stats if stats is not None else {}

    slot = max(uses, key=lambda item: item[1].width * item[1].height)
    page_no, rect = slot
    width_pt = rect.width
    if width_pt <= 1:
        return
    dpi_now = width / (width_pt / 72.0)

    # -- 1. try to restore the original picture -------------------------
    if upgrade and pictures is not None:
        if max_dpi is None or dpi_now < float(max_dpi) - 0.5:
            verdict = _restore_picture(doc, xref, obj, pictures, page_no, rect,
                                       width, height, width_pt,
                                       bool(any("DCTDecode" in name for name in filters)),
                                       quality, subsampling, max_dpi)
            if verdict:
                stats[verdict] = stats.get(verdict, 0) + 1
                return

    # -- 2. otherwise downsample when above the cap ---------------------
    if cap is None:
        return
    cap = float(cap)
    if dpi_now <= cap + 0.5:
        return
    try:
        original = obj.read_raw_bytes()
    except Exception:
        return
    if len(original) < 16 * 1024:
        return

    image = _pdf_image(doc, xref, obj)
    if image is None:
        return
    shrunk = _resize_to_dpi(image, width_pt, cap)
    if shrunk.width >= image.width:
        return
    data = _encode_jpeg(shrunk, quality, subsampling)
    if len(data) >= len(original) * 0.98:
        return
    _replace_stream(obj, shrunk, data)


def _restore_picture(doc, xref: int, obj, pictures: _PptxPictures,
                     page_no: int, rect, width: int, height: int, width_pt: float,
                     stream_is_jpeg: bool, quality: int, subsampling: int,
                     max_dpi=None) -> Optional[str]:
    """Rewrite one PDF image from the picture stored in the deck.

    Nothing in the output may be worse than what a WPS-class exporter writes
    for the same slide, so every picture is judged against a measured floor:
    the error of the ``quality 100, 4:2:0`` encode such an exporter produces.
    The replacement is then the smallest payload that clears it:

    ``verbatim``  the deck's own JPEG, embedded byte for byte — error 0, so no
                  amount of re-encoding can beat it, and it is usually the
                  smallest candidate as well;
    ``upgraded``  our re-encode at full resolution, when PowerPoint had
                  downsampled the picture;
    ``repaired``  pixels were already full-size, but PowerPoint's bytes were
                  below the floor.

    Returns the treatment label, or ``None`` when the stream was left alone.
    """
    picture = _find_picture(pictures, page_no, rect)
    if picture is None:
        _trace(f"p{page_no + 1}: no picture placement matched the {width}x{height} slot")
        return None
    source = pictures.load(picture)
    if source is None:
        return None

    cropped = bool(any(picture.crop))
    if cropped and abs(source.width / source.height - width / height) > 0.02:
        return None
    if source.width < width or source.height < height:
        return None                     # the deck's copy is no larger: nothing to win

    fingerprint = _pdf_thumb(doc, xref, obj)
    if (fingerprint is None
            or _thumb_distance(fingerprint, _thumb(source)) > _THUMB_TOLERANCE):
        _trace(f"p{page_no + 1}: thumbnail mismatch, left untouched")
        return None

    downsampled = (source.width >= width * _UPGRADE_MIN_GAIN
                   and source.height >= height * _UPGRADE_MIN_GAIN)
    if not stream_is_jpeg and not downsampled:
        return None                     # already lossless at full size

    floor = _exporter_floor(source)
    current_error = float("inf")
    if not downsampled:
        if (source.width, source.height) != (width, height):
            return None
        if not stream_is_jpeg:
            return None                 # lossless beats anything we could write
        current = _pdf_image(doc, xref, obj)
        if current is not None:
            current_error = _mse(source, current)
        if current_error <= floor:
            if _DEBUG:
                probe = _jpeg_quality(_stream_bytes(obj))
                _trace(f"p{page_no + 1}: PowerPoint's encode already clears the floor"
                       f" (Q{probe[0] if probe else '?'}"
                       f"{', 4:2:0' if probe and probe[1] > 0 else ''}), kept")
            return None

    pick = _floor_payload(pictures, picture, source, quality, subsampling,
                          max_dpi, width_pt, floor)
    if pick is None:
        return None
    image, data, kind, error = pick

    if not downsampled and current_error <= error * 1.11:
        _trace(f"p{page_no + 1}: replacement would not improve on PowerPoint's, kept")
        return None

    verdict = "verbatim" if kind == "verbatim" else ("upgraded" if downsampled else "repaired")
    _trace(f"p{page_no + 1}: {width}x{height} -> {image.width}x{image.height} "
           f"{verdict} (mse {error:.1f} vs {current_error:.1f}, floor {floor:.1f}, "
           f"{len(data) / 1048576:.2f}MB)")
    _replace_stream(obj, image, data)
    return verdict


def _stream_bytes(obj) -> bytes:
    try:
        return obj.read_raw_bytes() or b""
    except Exception:
        return b""


def _image_note(stats: dict[str, int]) -> str:
    """Short, honest summary of what happened to the pictures."""
    parts: list[str] = []
    if stats.get("verbatim"):
        parts.append(f"原图直嵌 {stats['verbatim']} 张")
    if stats.get("repaired"):
        parts.append(f"修复压缩 {stats['repaired']} 张")
    if stats.get("upgraded"):
        parts.append(f"补分辨率 {stats['upgraded']} 张")
    return "、".join(parts)


def _pdf_image(doc, xref: int, obj) -> Optional[Image.Image]:
    try:
        if any("DCTDecode" in name for name in _filter_names(obj)):
            with Image.open(io.BytesIO(obj.read_raw_bytes())) as image:
                image.load()
                if image.mode not in ("RGB", "L"):
                    return None
                return image.copy()
        return _decode_flate(obj, doc, xref)
    except Exception:
        return None


def assemble_pages(page_dir: Path, dst: str | Path, slide_pt: tuple[float, float] | None,
                   dpi: float, quality: int, subsampling: int,
                   on_progress: Optional[Callable[[float], None]] = None) -> int:
    """Wrap rendered slide images into a PDF, one image per page.

    The page box is the real slide size (in points), so the document is
    physically identical to the presentation no matter what resolution the
    pages were rendered at.
    """
    pages = sorted(Path(page_dir).glob("p*.png"))
    if not pages:
        raise EngineError("PowerPoint 没有渲染出任何页面图像。")
    target = Path(dst)
    tmp = target.parent / f".~ppt2pdf-{uuid.uuid4().hex}.pdf"
    document = pymupdf.open()
    try:
        for done, path in enumerate(pages, 1):
            with Image.open(path) as raw:
                raw.load()
                image = raw.convert("RGB")
            if slide_pt:
                width_pt, height_pt = float(slide_pt[0]), float(slide_pt[1])
            else:
                width_pt = image.width / dpi * 72.0
                height_pt = image.height / dpi * 72.0
            buffer = io.BytesIO()
            image.save(buffer, "JPEG", quality=int(quality), optimize=True,
                       subsampling=int(subsampling))
            page = document.new_page(width=width_pt, height=height_pt)
            page.insert_image(pymupdf.Rect(0, 0, width_pt, height_pt),
                              stream=buffer.getvalue())
            del image
            if on_progress is not None:
                try:
                    on_progress(done / len(pages))
                except Exception:
                    pass
        document.save(str(tmp), deflate=True, garbage=3)
    finally:
        document.close()
    os.replace(tmp, target)
    return len(pages)


def _unique_path(folder: Path, stem: str, overwrite: bool) -> Path:
    target = folder / f"{stem}.pdf"
    if overwrite or not target.exists():
        return target
    index = 1
    while True:
        candidate = folder / f"{stem} ({index}).pdf"
        if not candidate.exists():
            return candidate
        index += 1


class BatchConverter:
    """Runs a batch and reports through ``emit``."""

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
        self.outdir = outdir
        self.overwrite = overwrite
        self.cancel_event = cancel_event
        self.emit = emit
        # "vector" (PowerPoint export + image restore) or "render" (page raster)
        self.engine_mode = self.preset.get("mode", "vector")
        self.render_dpi = float(self.preset.get("render_dpi", 400))
        self._work_dir: Optional[Path] = None
        self._slide_sizes: dict[int, tuple[float, float]] = {}
        self.engine = ""

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

    def _pages_dir(self, index: int) -> Path:
        assert self._work_dir is not None
        return self._work_dir / f"p{index:04d}"

    def _progress(self, index: int, fraction: float) -> None:
        self.emit("progress", {"index": index, "fraction": fraction})

    def _finish_one(self, index: int, src: Path, dst: Path, started: float) -> FileResult:
        result = FileResult(src=str(src))
        try:
            if self.engine_mode == "render":
                pages = assemble_pages(
                    self._pages_dir(index), dst, self._slide_sizes.get(index),
                    self.render_dpi, int(self.preset.get("quality", 95)),
                    int(self.preset.get("subsampling", 0)),
                    on_progress=lambda fraction: self._progress(index, fraction),
                )
                result.pages = pages
                result.dst = str(dst)
                result.size_before = 0
                result.size_after = dst.stat().st_size
                result.ok = pages > 0
                return result

            raw_size = dst.stat().st_size
            if raw_size <= 0:
                raise EngineError("PowerPoint 导出的 PDF 是空文件。")
            needs_work = bool(self.preset.get("upgrade")) or self.preset.get("cap") is not None
            pages = 0
            stats: dict[str, int] = {}
            if needs_work:
                tmp = dst.parent / f".~ppt2pdf-{uuid.uuid4().hex}.pdf"
                try:
                    stats = optimize_pdf(
                        dst, tmp, self.preset, pptx_path=src,
                        on_progress=lambda fraction: self._progress(index, fraction),
                    ) or {}
                    if tmp.exists() and tmp.stat().st_size > 0:
                        # Validate before swapping in: on any problem the
                        # untouched PowerPoint PDF stays as the result.
                        with pymupdf.open(tmp) as check:
                            pages = check.page_count
                        os.replace(tmp, dst)
                except Exception:
                    pages = 0
                finally:
                    try:
                        tmp.unlink(missing_ok=True)
                    except Exception:
                        pass
            if not pages:
                with pymupdf.open(dst) as check:
                    pages = check.page_count
            result.pages = pages
            result.dst = str(dst)
            result.size_before = raw_size
            result.size_after = dst.stat().st_size
            result.note = _image_note(stats)
            result.ok = True
        except Exception as exc:
            result.error = f"{exc.__class__.__name__}: {exc}"
        finally:
            result.seconds = time.time() - started
        return result

    # -- main ------------------------------------------------------------
    def run(self) -> None:
        work_dir = Path(tempfile.mkdtemp(prefix="ppt2pdf-"))
        self._work_dir = work_dir
        render = self.engine_mode == "render"
        log_path = work_dir / "log.txt"
        list_path = work_dir / "list.txt"
        cancel_path = work_dir / "cancel.flag"
        vbs_path = work_dir / "export.vbs"
        proc = None
        stage_text = self.preset.get("stage") or "正在优化画质…"

        try:
            self.emit("stage", {"text": "正在启动 PowerPoint 引擎…"})
            targets = self._targets()
            slots: list[Optional[FileResult]] = [None] * len(targets)
            vbs_path.write_text(VBS_RENDER if render else VBS_EXPORT,
                                encoding="ascii")
            with list_path.open("w", encoding="utf-16") as fh:
                for index, (src, dst) in enumerate(targets):
                    if render:
                        pages_dir = self._pages_dir(index)
                        pages_dir.mkdir(parents=True, exist_ok=True)
                        fh.write(f"{src}\t{dst}\t{pages_dir}\n")
                    else:
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
            command = [cscript, "//nologo", "//B", str(vbs_path), str(list_path),
                       str(log_path), str(cancel_path)]
            if render:
                command.append(f"{self.render_dpi:g}")
            err_path = work_dir / "helper.err"
            helper_err = open(err_path, "wb")
            try:
                proc = subprocess.Popen(
                    command, stdout=subprocess.DEVNULL, stderr=helper_err,
                    creationflags=_NO_WINDOW_FLAGS,
                    startupinfo=_HIDDEN_STARTUP,
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
                            lines.pop()  # ignore a partially written last line
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
                        engine = parts[1] if len(parts) > 1 else ""
                        self.engine = engine
                        self.emit("stage", {"text": f"正在使用 {engine} 引擎…"})
                        continue
                    if head == "FATAL":
                        code = parts[1] if len(parts) > 1 else ""
                        fatal = {
                            "APP": "没有找到可用的 PowerPoint / WPS 演示，"
                                   "请先安装 Microsoft Office 或 WPS Office。",
                            "LIST": "无法读取转换任务列表，请重试。",
                        }.get(code, "PowerPoint 启动失败")
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
                                idx = int(parts[1])
                            except ValueError:
                                continue
                            if idx < len(self.files):
                                self.emit("start", {"index": idx,
                                                    "name": Path(self.files[idx]).name})
                        continue
                    try:
                        idx = int(head)
                    except ValueError:
                        continue
                    if len(parts) > 1 and parts[1] == "PAGE":
                        try:
                            page_no, page_total = int(parts[2]), int(parts[3])
                        except (IndexError, ValueError):
                            continue
                        if page_total > 0:
                            self._progress(idx, page_no / page_total)
                        continue
                    if len(parts) > 1 and parts[1] == "SIZE":
                        try:
                            self._slide_sizes[idx] = (float(parts[2]), float(parts[3]))
                        except (IndexError, ValueError):
                            pass
                        continue
                    if idx in handled or idx >= len(targets):
                        continue
                    handled.add(idx)
                    src, dst = targets[idx]
                    if len(parts) > 1 and parts[1] == "OK":
                        self.emit("compress", {"index": idx, "text": stage_text})
                        result = self._finish_one(idx, src, dst, started_at)
                        slots[idx] = result
                        if result.ok:
                            self.emit("file_ok", {"index": idx, "result": result})
                        else:
                            self.emit("file_error", {"index": idx, "result": result})
                    else:
                        message = " ".join(parts[2:]) if len(parts) > 2 else "转换失败"
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
                        slots[idx] = result
                        self.emit("file_error", {"index": idx, "result": result})

                if finished or cancelled or fatal:
                    break
                if proc.poll() is not None:
                    exit_checks += 1
                    if exit_checks >= 4:
                        break
                    time.sleep(0.3)
                    continue
                if time.time() - last_progress > 300:
                    fatal = "PowerPoint 长时间没有响应，转换已停止。"
                    try:
                        proc.terminate()
                    except Exception:
                        pass
                    break
                time.sleep(0.15)

            # Wait briefly for the helper to exit so the log is complete.
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

            # Fallback: even if the helper log could not be parsed, usable
            # output on disk means the export actually succeeded.
            for idx, (src, dst) in enumerate(targets):
                if slots[idx] is not None:
                    continue
                try:
                    if render:
                        if not any(self._pages_dir(idx).glob("p*.png")):
                            continue
                    elif not (dst.exists() and dst.stat().st_size > 0):
                        continue
                    if not render:
                        self.emit("compress", {"index": idx, "text": stage_text})
                    result = self._finish_one(idx, src, dst, started_at)
                    slots[idx] = result
                    if result.ok:
                        self.emit("file_ok", {"index": idx, "result": result})
                    else:
                        self.emit("file_error", {"index": idx, "result": result})
                except Exception:
                    pass

            # Any file without an outcome was interrupted / cancelled.
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
                detail = f"（引擎退出码 {code}{('；' + tail.strip()) if tail.strip() else ''}）"
            for idx, (src, dst) in enumerate(targets):
                if slots[idx] is not None:
                    continue
                cancelled_result = FileResult(
                    src=str(src), cancelled=cancelled,
                    error="已取消" if cancelled else "转换进程提前结束，文件未完成转换。" + detail,
                    seconds=time.time() - started_at,
                )
                slots[idx] = cancelled_result
                if cancelled:
                    self.emit("file_cancelled", {"index": idx, "result": cancelled_result})
                else:
                    self.emit("file_error", {"index": idx, "result": cancelled_result})

            ordered = [item for item in slots if item is not None]
            self.emit("done", {"results": ordered, "cancelled": cancelled,
                               "engine": self.engine})
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
