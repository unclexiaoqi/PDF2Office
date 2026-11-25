"""PDF2Office — one tool, three conversions.

    pdf2word   PDF  → Word  (.docx)
    ppt2pdf    PPT  → PDF   (.pdf)
    pdf2ppt    PDF  → PPT   (.pptx)

Each direction lives in its own engine module and exposes the same
``MODES`` / ``MODE_ORDER`` / ``DEFAULT_MODE`` / ``BatchConverter`` API; this
module is the thin facade the GUI talks to.
"""

from __future__ import annotations

from typing import Callable, Optional

import engine_pdf2ppt as _pdf2ppt
import engine_pdf2word as _pdf2word
import engine_ppt2pdf as _ppt2pdf
from engine_pdf2word import FileResult, human_size

__all__ = [
    "CONVERSIONS",
    "CONVERSION_ORDER",
    "DEFAULT_CONVERSION",
    "BatchConverter",
    "FileResult",
    "human_size",
]

PDF_EXTS = _pdf2word.SUPPORTED_EXTS
PPT_EXTS = _ppt2pdf.SUPPORTED_EXTS

CONVERSIONS = {
    "pdf2word": {
        "name": "PDF 转 Word",
        "short": "PDF → Word",
        "engine": _pdf2word.BatchConverter,
        "modes": _pdf2word.MODES,
        "order": _pdf2word.MODE_ORDER,
        "default_mode": _pdf2word.DEFAULT_MODE,
        "accept": PDF_EXTS,
        "suffix": ".docx",
        "hint": "把 PDF 拖到这里，或点击上方按钮添加（支持 .pdf，可批量）",
    },
    "ppt2pdf": {
        "name": "PPT 转 PDF",
        "short": "PPT → PDF",
        "engine": _ppt2pdf.BatchConverter,
        "modes": _ppt2pdf.MODES,
        "order": _ppt2pdf.MODE_ORDER,
        "default_mode": _ppt2pdf.DEFAULT_MODE,
        "accept": PPT_EXTS,
        "suffix": ".pdf",
        "hint": "把 PPT 拖到这里，或点击上方按钮添加（支持 .ppt / .pptx / .pptm / .ppsx）",
    },
    "pdf2ppt": {
        "name": "PDF 转 PPT",
        "short": "PDF → PPT",
        "engine": _pdf2ppt.BatchConverter,
        "modes": _pdf2ppt.MODES,
        "order": _pdf2ppt.MODE_ORDER,
        "default_mode": _pdf2ppt.DEFAULT_MODE,
        "accept": PDF_EXTS,
        "suffix": ".pptx",
        "hint": "把 PDF 拖到这里，或点击上方按钮添加（支持 .pdf，可批量）",
    },
}

CONVERSION_ORDER = ["pdf2word", "ppt2pdf", "pdf2ppt"]
DEFAULT_CONVERSION = "pdf2word"


class BatchConverter:
    """Dispatches to the engine that matches ``conversion``."""

    def __init__(
        self,
        files: list[str],
        conversion: str,
        mode: str,
        outdir: Optional[str],
        overwrite: bool,
        cancel_event,
        emit: Callable[[str, dict], None],
    ) -> None:
        spec = CONVERSIONS.get(conversion) or CONVERSIONS[DEFAULT_CONVERSION]
        self.conversion = conversion if conversion in CONVERSIONS else DEFAULT_CONVERSION
        self.mode = mode if mode in spec["modes"] else spec["default_mode"]
        self._impl = spec["engine"](files, self.mode, outdir, overwrite,
                                    cancel_event, emit)

    def run(self) -> None:
        self._impl.run()
