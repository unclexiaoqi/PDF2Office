"""PDF → PowerPoint conversion engine.

``image`` modes
    Every PDF page is rendered with PyMuPDF and placed on its own slide, sized
    to the original page.  The layout is pixel-exact and no Office is needed.

``editable`` mode
    Rebuilds the page from its own content: pictures are placed where they
    were, text blocks become real PowerPoint text boxes (with font size and
    colour), so the text can be edited.  Fidelity is approximate — vector
    lines, table rules and effects are not reproduced.
"""

from __future__ import annotations

import io
import os
import time
import uuid
from dataclasses import dataclass
from pathlib import Path
from typing import Optional

import pymupdf
from PIL import Image
from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.util import Emu, Pt

MODES = {
    "image300": {
        "name": "原版式",
        "short": "300 DPI",
        "desc": "每页做成一张幻灯片，版式和原 PDF 完全一致",
        "tag": "版式 100% 保留 · 文字不可编辑",
        "badge": "推荐",
        "engine": "image",
        "dpi": 300,
        "quality": 92,
        "stage": "正在还原版面…",
    },
    "editable": {
        "name": "可编辑",
        "short": "重建文字",
        "desc": "文字变成可选可改的文本框，图片按原位放置",
        "tag": "版式近似 · 矢量线条不保留",
        "engine": "editable",
        "stage": "正在重建页面…",
    },
    "image200": {
        "name": "版式还原",
        "short": "200 DPI",
        "desc": "每页做成一张幻灯片，体积更小",
        "tag": "适合邮件 / 微信发送",
        "engine": "image",
        "dpi": 200,
        "quality": 92,
        "stage": "正在还原版面…",
    },
    "image120": {
        "name": "小体积",
        "short": "120 DPI",
        "desc": "每页做成一张幻灯片，优先减小体积",
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

EMU_PER_PT = 12700.0
#: PowerPoint refuses slides bigger than 56 inches.
_MAX_SLIDE_PT = 56 * 72


class EngineError(RuntimeError):
    pass


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
# helpers
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


def _ensure_pdf(path: str | Path) -> None:
    """Reject anything that is not really a PDF.

    PyMuPDF also opens Office documents (MuPDF renders docx/xlsx/pptx), so
    without this check a .pptx would be turned into a one-page slide deck
    instead of being reported as the wrong input type.
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


def _slide_size(page: pymupdf.Page) -> tuple[float, float]:
    """Page size in points, clamped to what PowerPoint accepts."""
    width, height = page.rect.width, page.rect.height
    scale = min(1.0, _MAX_SLIDE_PT / max(width, height, 1.0))
    return width * scale, height * scale


def _clean_font(name: str) -> str:
    """Drop the subset prefix Word/PDF embeds, e.g. 'BCDEEE+Calibri'."""
    if name and "+" in name and len(name.split("+", 1)[0]) == 6:
        return name.split("+", 1)[1]
    return name or ""


def _set_east_asian_font(run, name: str) -> None:
    """python-pptx only sets the latin typeface; CJK needs <a:ea> too."""
    if not name:
        return
    try:
        from pptx.oxml.ns import qn

        rpr = run.font._rPr
        ea = rpr.find(qn("a:ea"))
        if ea is None:
            ea = rpr.makeelement(qn("a:ea"), {})
            rpr.append(ea)
        ea.set("typeface", name)
    except Exception:
        pass


def _add_text_block(slide, block: dict, scale_x: float, scale_y: float) -> None:
    lines = block.get("lines") or []
    if not lines:
        return
    x0, y0, x1, y1 = block["bbox"]
    left = Emu(int(round(x0 * scale_x * EMU_PER_PT)))
    top = Emu(int(round(y0 * scale_y * EMU_PER_PT)))
    width = Emu(max(int(round((x1 - x0) * scale_x * EMU_PER_PT)), 9144))
    height = Emu(max(int(round((y1 - y0) * scale_y * EMU_PER_PT)), 9144))
    box = slide.shapes.add_textbox(left, top, width, height)
    frame = box.text_frame
    frame.word_wrap = False
    try:
        frame.margin_left = 0
        frame.margin_right = 0
        frame.margin_top = 0
        frame.margin_bottom = 0
    except Exception:
        pass
    used = False
    for line in lines:
        paragraph = frame.paragraphs[0] if not used else frame.add_paragraph()
        used = True
        try:
            paragraph.space_before = Pt(0)
            paragraph.space_after = Pt(0)
            paragraph.line_spacing = 1.0
        except Exception:
            pass
        for span in line.get("spans", []):
            text = span.get("text", "")
            if not text:
                continue
            run = paragraph.add_run()
            run.text = text
            try:
                run.font.size = Pt(round(float(span.get("size", 12)), 1))
            except Exception:
                pass
            colour = span.get("color", 0)
            try:
                run.font.color.rgb = RGBColor(
                    (colour >> 16) & 0xFF, (colour >> 8) & 0xFF, colour & 0xFF)
            except Exception:
                pass
            font_name = _clean_font(str(span.get("font", "")))
            if font_name:
                try:
                    run.font.name = font_name
                except Exception:
                    pass
                _set_east_asian_font(run, font_name)


# ---------------------------------------------------------------------------
# builders
# ---------------------------------------------------------------------------
def build_image_pptx(pdf_path, dst, dpi: float, quality: int,
                     on_progress=None, cancel_event=None) -> int:
    """One slide per page, each filled with the rendered page."""
    _ensure_pdf(pdf_path)
    source = pymupdf.open(str(pdf_path))
    presentation = Presentation()
    try:
        width_pt, height_pt = _slide_size(source[0])
        presentation.slide_width = Emu(int(round(width_pt * EMU_PER_PT)))
        presentation.slide_height = Emu(int(round(height_pt * EMU_PER_PT)))
        blank = presentation.slide_layouts[6]
        total = source.page_count
        for index, page in enumerate(source, 1):
            if cancel_event is not None and cancel_event.is_set():
                raise EngineError("已取消")
            slide = presentation.slides.add_slide(blank)
            data = _render_page(page, dpi, quality)
            page_width, page_height = _slide_size(page)
            if abs(page_width - width_pt) < 0.5 and abs(page_height - height_pt) < 0.5:
                slide.shapes.add_picture(io.BytesIO(data), 0, 0,
                                         width=presentation.slide_width,
                                         height=presentation.slide_height)
            else:
                # mixed page sizes: keep the aspect, centre on the slide
                scale = min(width_pt / page_width, height_pt / page_height)
                draw_w = int(round(page_width * scale * EMU_PER_PT))
                draw_h = int(round(page_height * scale * EMU_PER_PT))
                slide.shapes.add_picture(
                    io.BytesIO(data),
                    Emu(int((presentation.slide_width - draw_w) / 2)),
                    Emu(int((presentation.slide_height - draw_h) / 2)),
                    width=Emu(draw_w), height=Emu(draw_h))
            if on_progress is not None:
                try:
                    on_progress(index / total)
                except Exception:
                    pass
    finally:
        source.close()
    _save(presentation, dst)
    return total


def build_editable_pptx(pdf_path, dst, on_progress=None, cancel_event=None) -> int:
    """Rebuild each page from pictures + real text boxes."""
    _ensure_pdf(pdf_path)
    source = pymupdf.open(str(pdf_path))
    presentation = Presentation()
    try:
        width_pt, height_pt = _slide_size(source[0])
        presentation.slide_width = Emu(int(round(width_pt * EMU_PER_PT)))
        presentation.slide_height = Emu(int(round(height_pt * EMU_PER_PT)))
        blank = presentation.slide_layouts[6]
        total = source.page_count
        for index, page in enumerate(source, 1):
            if cancel_event is not None and cancel_event.is_set():
                raise EngineError("已取消")
            slide = presentation.slides.add_slide(blank)
            page_w, page_h = _slide_size(page)
            scale_x = page_w / max(page.rect.width, 1.0)
            scale_y = page_h / max(page.rect.height, 1.0)

            # pictures first so text lands on top
            seen: set[int] = set()
            for item in page.get_images(full=True):
                xref = item[0]
                if xref in seen:
                    continue
                seen.add(xref)
                try:
                    rects = page.get_image_rects(xref)
                except Exception:
                    rects = []
                if not rects:
                    continue
                rect = max(rects, key=lambda r: r.width * r.height)
                try:
                    info = source.extract_image(xref)
                    data = info["image"]
                except Exception:
                    continue
                if rect.width < 4 or rect.height < 4:
                    continue
                try:
                    slide.shapes.add_picture(
                        io.BytesIO(data),
                        Emu(int(round(rect.x0 * scale_x * EMU_PER_PT))),
                        Emu(int(round(rect.y0 * scale_y * EMU_PER_PT))),
                        width=Emu(int(round(rect.width * scale_x * EMU_PER_PT))),
                        height=Emu(int(round(rect.height * scale_y * EMU_PER_PT))))
                except Exception:
                    # exotic image encodings (JPX, CCITT, raw) are not supported
                    # by python-pptx; skipping one picture beats failing the page
                    continue

            try:
                blocks = page.get_text("dict").get("blocks", [])
            except Exception:
                blocks = []
            for block in blocks:
                if block.get("type") != 0:
                    continue
                _add_text_block(slide, block, scale_x, scale_y)

            if on_progress is not None:
                try:
                    on_progress(index / total)
                except Exception:
                    pass
    finally:
        source.close()
    _save(presentation, dst)
    return total


def _save(presentation: Presentation, dst) -> None:
    target = Path(dst)
    tmp = target.with_name(f".~pdf2office-{uuid.uuid4().hex}.pptx")
    try:
        presentation.save(str(tmp))
        os.replace(tmp, target)
    finally:
        try:
            tmp.unlink(missing_ok=True)
        except Exception:
            pass


# ---------------------------------------------------------------------------
# batch driver
# ---------------------------------------------------------------------------
def _unique_path(folder: Path, stem: str, overwrite: bool) -> Path:
    target = folder / f"{stem}.pptx"
    if overwrite or not target.exists():
        return target
    index = 1
    while True:
        candidate = folder / f"{stem} ({index}).pptx"
        if not candidate.exists():
            return candidate
        index += 1


class BatchConverter:
    """Runs a batch of PDF → PPTX conversions."""

    def __init__(self, files, mode, outdir, overwrite, cancel_event, emit) -> None:
        self.files = list(files)
        self.mode = mode if mode in MODES else DEFAULT_MODE
        self.preset = MODES[self.mode]
        self.engine = self.preset.get("engine", "image")
        self.outdir = outdir
        self.overwrite = overwrite
        self.cancel_event = cancel_event
        self.emit = emit

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

    def run(self) -> None:
        stage_text = self.preset.get("stage") or "正在转换…"
        results: list[FileResult] = []
        cancelled = False
        try:
            targets = self._targets()
            for index, (src, dst) in enumerate(targets):
                if self.cancel_event.is_set():
                    cancelled = True
                    outcome = FileResult(src=str(src), cancelled=True, error="已取消")
                    results.append(outcome)
                    self.emit("file_cancelled", {"index": index, "result": outcome})
                    continue
                self.emit("start", {"index": index, "name": src.name})
                self.emit("stage", {"text": stage_text})
                started = time.time()
                result = FileResult(src=str(src), dst=str(dst))
                try:
                    if self.engine == "editable":
                        pages = build_editable_pptx(
                            src, dst,
                            on_progress=lambda fraction, i=index: self._progress(i, fraction),
                            cancel_event=self.cancel_event)
                    else:
                        pages = build_image_pptx(
                            src, dst, float(self.preset.get("dpi", 300)),
                            int(self.preset.get("quality", 92)),
                            on_progress=lambda fraction, i=index: self._progress(i, fraction),
                            cancel_event=self.cancel_event)
                    result.pages = pages
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
                results.append(result)
                if result.ok:
                    self.emit("file_ok", {"index": index, "result": result})
                elif result.cancelled:
                    cancelled = True
                    self.emit("file_cancelled", {"index": index, "result": result})
                else:
                    self.emit("file_error", {"index": index, "result": result})
            self.emit("done", {"results": results, "cancelled": cancelled,
                               "engine": "PyMuPDF + python-pptx"})
        except Exception as exc:
            self.emit("fatal", {"text": f"发生意外错误：{exc}"})
