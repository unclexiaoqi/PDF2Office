"""PDF2Office self-test — exercises every conversion, mode and edge case.

Run from the project folder:

    python selftest.py            # full suite
    python selftest.py --quick    # skip the Office-dependent modes

It builds its own inputs under %TEMP%\\pdf2office-selftest and prints a
PASS/FAIL line per check, exiting non-zero if anything fails.
"""

from __future__ import annotations

import ctypes
import os
import shutil
import sys
import threading
import time

import pymupdf

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE)

import converter  # noqa: E402

WORK = os.path.join(os.environ.get("TEMP", r"C:\Temp"), "pdf2office-selftest")
RESULTS: list[tuple[str, bool, str]] = []
QUICK = "--quick" in sys.argv


# ---------------------------------------------------------------------------
# helpers
# ---------------------------------------------------------------------------
def check(name: str, condition: bool, detail: str = "") -> bool:
    RESULTS.append((name, bool(condition), detail))
    mark = "PASS" if condition else "FAIL"
    line = f"[{mark}] {name}"
    if detail and not condition:
        line += f"  -> {detail}"
    print(line, flush=True)
    return bool(condition)


def build_inputs() -> dict[str, str]:
    os.makedirs(WORK, exist_ok=True)
    paths = {}

    image = os.path.join(WORK, "pic.jpg")
    if not os.path.exists(image):
        from PIL import Image, ImageDraw

        im = Image.new("RGB", (1200, 800))
        draw = ImageDraw.Draw(im)
        for y in range(0, 800, 2):
            draw.rectangle([0, y, 1200, y + 2], fill=(40 + y // 5, 110, 190 - y // 8))
        im.save(image, "JPEG", quality=90)

    # 2-page PDF with text, table and picture
    doc = pymupdf.open()
    page = doc.new_page(width=595, height=842)
    page.insert_text((60, 90), "PDF2Office 自检文档", fontsize=20, fontname="china-s")
    page.insert_text((60, 130), "正文内容，用于验证版式。", fontsize=12, fontname="china-s")
    for r in range(3):
        for c in range(3):
            page.insert_text((66 + c * 150, 222 + r * 30), f"R{r + 1}C{c + 1}", fontsize=11)
    page.insert_image(pymupdf.Rect(60, 320, 400, 560), filename=image)
    page = doc.new_page(width=595, height=842)
    for i in range(8):
        page.insert_text((60, 100 + i * 26), f"要点 {i + 1}", fontsize=13, fontname="china-s")
    paths["pdf"] = os.path.join(WORK, "sample.pdf")
    doc.save(paths["pdf"])
    doc.close()

    # odd-size PDF (wide) to check page-size preservation
    doc = pymupdf.open()
    page = doc.new_page(width=1200, height=500)
    page.insert_text((60, 100), "wide page", fontsize=24)
    paths["wide"] = os.path.join(WORK, "wide.pdf")
    doc.save(paths["wide"])
    doc.close()

    # encrypted PDF
    doc = pymupdf.open()
    page = doc.new_page()
    page.insert_text((72, 100), "secret", fontsize=14)
    paths["locked"] = os.path.join(WORK, "locked.pdf")
    doc.save(paths["locked"], encryption=pymupdf.PDF_ENCRYPT_AES_256,
             user_pw="user", owner_pw="owner")
    doc.close()

    # image-only (scanned) PDF — must be flagged in the result note
    doc = pymupdf.open()
    page = doc.new_page(width=595, height=842)
    page.insert_image(pymupdf.Rect(0, 0, 595, 842), filename=image)
    paths["scanned"] = os.path.join(WORK, "scanned.pdf")
    doc.save(paths["scanned"])
    doc.close()

    # corrupt file
    paths["broken"] = os.path.join(WORK, "broken.pdf")
    with open(paths["broken"], "wb") as handle:
        handle.write(b"%PDF-1.7\nthis is not a real pdf\n")

    # empty file
    paths["empty"] = os.path.join(WORK, "empty.pdf")
    open(paths["empty"], "wb").close()

    # 2-slide deck
    from pptx import Presentation
    from pptx.util import Emu, Pt

    prs = Presentation()
    prs.slide_width = Emu(int(420 / 25.4 * 914400))
    prs.slide_height = Emu(int(297 / 25.4 * 914400))
    for title in ("自检页一", "自检页二"):
        slide = prs.slides.add_slide(prs.slide_layouts[6])
        slide.shapes.add_picture(image, 0, 0, width=prs.slide_width,
                                 height=prs.slide_height)
        box = slide.shapes.add_textbox(Emu(500000), Emu(300000), Emu(8000000), Emu(1000000))
        run = box.text_frame.paragraphs[0].add_run()
        run.text = title
        run.font.size = Pt(40)
    paths["pptx"] = os.path.join(WORK, "deck.pptx")
    prs.save(paths["pptx"])

    return paths


def run_batch(conversion: str, mode: str, files: list[str],
              overwrite: bool = True, cancel_after: float | None = None,
              fresh: bool = True):
    outdir = os.path.join(WORK, "out", f"{conversion}_{mode}")
    if fresh:
        shutil.rmtree(outdir, ignore_errors=True)
    os.makedirs(outdir, exist_ok=True)
    events: list[tuple[str, dict]] = []
    cancel = threading.Event()
    worker = converter.BatchConverter(
        files, conversion, mode, outdir, overwrite, cancel,
        lambda kind, payload: events.append((kind, payload)),
    )
    if cancel_after is None:
        started = time.time()
        worker.run()
        return events, outdir, time.time() - started
    thread = threading.Thread(target=worker.run, daemon=True)
    thread.start()
    time.sleep(cancel_after)
    cancel.set()
    thread.join(120)
    return events, outdir, cancel_after


def outputs(outdir: str, suffix: str) -> list[str]:
    if not os.path.isdir(outdir):
        return []
    return sorted(f for f in os.listdir(outdir) if f.lower().endswith(suffix))


def pages_of(path: str) -> int:
    ext = os.path.splitext(path)[1].lower()
    if ext == ".pdf":
        with pymupdf.open(path) as doc:
            return doc.page_count
    if ext == ".docx":
        from docx import Document

        doc = Document(path)
        return len(doc.inline_shapes) or len(doc.tables) or 1
    if ext == ".pptx":
        from pptx import Presentation

        return len(Presentation(path).slides)
    return 0


# ---------------------------------------------------------------------------
# checks
# ---------------------------------------------------------------------------
def check_conversion_matrix(paths: dict[str, str]) -> None:
    """Every conversion × every mode produces a valid file of the right kind."""
    matrix = {
        "pdf2word": (paths["pdf"], ".docx"),
        "ppt2pdf": (paths["pptx"], ".pdf"),
        "pdf2ppt": (paths["pdf"], ".pptx"),
    }
    for conversion, (source, suffix) in matrix.items():
        spec = converter.CONVERSIONS[conversion]
        for mode in spec["order"]:
            if QUICK and spec["modes"][mode].get("engine") in ("word", "editable"):
                check(f"{conversion}/{mode} (skipped, needs Office)", True, "")
                continue
            events, outdir, elapsed = run_batch(conversion, mode, [source])
            fatal = [p for kind, p in events if kind == "fatal"]
            done = [p for kind, p in events if kind == "done"]
            files = outputs(outdir, suffix)
            ok = not fatal and bool(done) and done[0]["results"][0].ok and len(files) == 1
            detail = fatal[0] if fatal else (done[0]["results"][0].error if done else "no done")
            size = os.path.getsize(os.path.join(outdir, files[0])) // 1024 if files else 0
            check(f"{conversion}/{mode} -> {suffix} ({elapsed:.1f}s, {size}KB)",
                  ok, detail or "missing output")


def check_page_counts(paths: dict[str, str]) -> None:
    """Page/slide counts must match the source."""
    with pymupdf.open(paths["pdf"]) as doc:
        source_pages = doc.page_count
    _events, outdir, _ = run_batch("pdf2ppt", "image300", [paths["pdf"]])
    files = outputs(outdir, ".pptx")
    if files:
        count = pages_of(os.path.join(outdir, files[0]))
        check(f"pdf2ppt slide count == {source_pages}", count == source_pages,
              f"got {count}")
    _events, outdir, _ = run_batch("pdf2word", "image300", [paths["pdf"]])
    files = outputs(outdir, ".docx")
    if files:
        from docx import Document

        doc = Document(os.path.join(outdir, files[0]))
        check(f"pdf2word image count == {source_pages}",
              len(doc.inline_shapes) == source_pages,
              f"got {len(doc.inline_shapes)}")


def check_page_size_preserved(paths: dict[str, str]) -> None:
    """Physical page/slide size must follow the PDF page size."""
    with pymupdf.open(paths["wide"]) as doc:
        width_pt = doc[0].rect.width
    events, outdir, _ = run_batch("pdf2ppt", "image300", [paths["wide"]])
    files = outputs(outdir, ".pptx")
    if files:
        from pptx import Presentation

        prs = Presentation(os.path.join(outdir, files[0]))
        ok = abs(prs.slide_width.pt - width_pt) < 2
        check(f"pdf2ppt keeps page width ({width_pt:.0f}pt)", ok,
              f"slide={prs.slide_width.pt:.0f}pt")


def check_batch(paths: dict[str, str]) -> None:
    files = [paths["pdf"], paths["wide"]]
    events, outdir, elapsed = run_batch("pdf2word", "image200", files)
    done = [p for kind, p in events if kind == "done"]
    count = len(outputs(outdir, ".docx"))
    ok = count == 2 and bool(done) and all(r.ok for r in done[0]["results"])
    check(f"batch of 2 -> 2 docx ({elapsed:.1f}s)", ok, f"got {count}")


def check_overwrite(paths: dict[str, str]) -> None:
    run_batch("pdf2word", "image120", [paths["pdf"]], overwrite=False)
    first = outputs(os.path.join(WORK, "out", "pdf2word_image120"), ".docx")
    run_batch("pdf2word", "image120", [paths["pdf"]], overwrite=False, fresh=False)
    outdir = os.path.join(WORK, "out", "pdf2word_image120")
    second = outputs(outdir, ".docx")
    check("second run auto-renames (no overwrite)", len(first) == 1 and len(second) == 2,
          f"{first} -> {second}")

    run_batch("pdf2word", "image120", [paths["pdf"]], overwrite=True)
    outdir = os.path.join(WORK, "out", "pdf2word_image120")
    check("overwrite=True keeps one file", len(outputs(outdir, ".docx")) == 1,
          str(outputs(outdir, ".docx")))


def check_target_guard(paths: dict[str, str]) -> None:
    """The output path must never equal the input path (would clobber it)."""
    folder = os.path.join(WORK, "guard")
    shutil.rmtree(folder, ignore_errors=True)
    os.makedirs(folder, exist_ok=True)
    victim = os.path.join(folder, "deck.pptx")
    shutil.copyfile(paths["pptx"], victim)

    engine = converter.CONVERSIONS["pdf2ppt"]["engine"]
    worker = engine([victim], "image300", folder, True, threading.Event(),
                    lambda kind, payload: None)
    targets = worker._targets()
    same = os.path.normcase(str(targets[0][1])) == os.path.normcase(victim)
    check("output path never equals input path", not same,
          f"target={targets[0][1].name}")
    check("source file still intact", os.path.getsize(victim) == os.path.getsize(paths["pptx"]),
          f"{os.path.getsize(victim)} vs {os.path.getsize(paths['pptx'])}")


def check_cancel(paths: dict[str, str]) -> None:
    events, outdir, _ = run_batch("pdf2word", "image300", [paths["pdf"]],
                                  cancel_after=0.0)
    kinds = [kind for kind, _ in events]
    done = [p for kind, p in events if kind == "done"]
    cancelled = any(k in ("file_cancelled",) for k in kinds)
    ok = bool(done) and (cancelled or done[0].get("cancelled"))
    check("cancel stops the batch cleanly", ok, f"kinds={sorted(set(kinds))}")


def check_bad_inputs(paths: dict[str, str]) -> None:
    for name, path in (("broken pdf", paths["broken"]),
                       ("empty file", paths["empty"]),
                       ("encrypted pdf", paths["locked"])):
        events, outdir, _ = run_batch("pdf2word", "image300", [path])
        fatal = [p for kind, p in events if kind == "fatal"]
        done = [p for kind, p in events if kind == "done"]
        ok = not fatal and bool(done) and not done[0]["results"][0].ok
        detail = ""
        if fatal:
            detail = "fatal: " + str(fatal[0])[:80]
        elif done:
            detail = str(done[0]["results"][0].error)[:80]
        check(f"{name} fails gracefully", ok, detail)


def check_wrong_extension(paths: dict[str, str]) -> None:
    events, outdir, _ = run_batch("pdf2ppt", "image300", [paths["pptx"]])
    fatal = [p for kind, p in events if kind == "fatal"]
    done = [p for kind, p in events if kind == "done"]
    ok = not fatal and bool(done) and not done[0]["results"][0].ok
    check("wrong extension fails without crashing", ok,
          str(fatal[0])[:80] if fatal else "")
    if done:
        error = str(done[0]["results"][0].error or "")
        check("wrong extension says why", "不是 PDF" in error, error[:80])


def check_scanned_note(paths: dict[str, str]) -> None:
    events, outdir, _ = run_batch("pdf2word", "image300", [paths["scanned"]])
    done = [p for kind, p in events if kind == "done"]
    note = done[0]["results"][0].note if done else ""
    check("scanned PDF is flagged", "扫描件" in (note or ""), f"note={note!r}")


def check_picture_fidelity(paths: dict[str, str]) -> None:
    """native mode must be at least as faithful as a WPS-class export.

    The tier's promise is a measured one: every picture in the output has to be
    no further from the deck's original than the ``quality 100, 4:2:0`` encode
    that other exporters write.  Comparing mean squared error against exactly
    that encode is the only way to know the promise still holds.
    """
    import zipfile

    import pikepdf

    import engine_ppt2pdf as engine

    events, outdir, elapsed = run_batch("ppt2pdf", "native", [paths["pptx"]])
    done = [p for kind, p in events if kind == "done"]
    files = outputs(outdir, ".pdf")
    if not (done and files):
        check("native is no worse than a WPS-class export", False, "no output")
        return

    result = done[0]["results"][0]
    with zipfile.ZipFile(paths["pptx"]) as archive:
        media = [n for n in archive.namelist() if n.startswith("ppt/media/")]
        original = archive.read(media[0]) if media else b""
    source = engine._decode_jpeg(original)
    if source is None:
        check("native is no worse than a WPS-class export", False, "no source picture")
        return

    path = os.path.join(outdir, files[0])
    with pymupdf.open(path) as doc, pikepdf.open(path) as pdf:
        found = doc[0].get_images(full=True) if doc.page_count else []
        error = None
        if found:
            obj = pdf.get_object((found[0][0], 0))
            image = engine._pdf_image(doc, found[0][0], obj)
            if image is not None:
                error = engine._mse(source, image)
    floor = engine._exporter_floor(source)
    ok = error is not None and error <= floor
    check(f"native is no worse than a WPS-class export ({elapsed:.1f}s)", ok,
          f"error={error} floor={floor:.2f} note={result.note!r}")


def check_jpeg_quality_probe() -> None:
    """The probe decides which pictures get repaired — it must read Q back."""
    import io as _io

    import engine_ppt2pdf as engine
    from PIL import Image, ImageDraw

    image = Image.new("RGB", (240, 180), (30, 120, 200))
    draw = ImageDraw.Draw(image)
    for x in range(0, 240, 3):
        draw.line([(x, 0), (x, 180)], fill=(x, 255 - x, 90))
    results = []
    for quality, sampling in ((95, 0), (75, 2), (60, 2)):
        buffer = _io.BytesIO()
        image.save(buffer, "JPEG", quality=quality, subsampling=sampling)
        results.append((quality, sampling, engine._jpeg_quality(buffer.getvalue())))
    ok = all(probe is not None and abs(probe[0] - quality) <= 3
             and probe[1] == sampling
             for quality, sampling, probe in results)
    check("jpeg quality probe reads Q and chroma back", ok, str(results))


def check_no_console_window(paths: dict[str, str]) -> None:
    """The Word helper must not flash a console window."""
    if QUICK:
        check("no console window (skipped)", True, "")
        return
    found: list[int] = []

    def callback(hwnd, _param):
        buffer = ctypes.create_unicode_buffer(256)
        ctypes.windll.user32.GetClassNameW(hwnd, buffer, 256)
        if buffer.value in ("ConsoleWindowClass", "PseudoConsoleWindow"):
            if ctypes.windll.user32.IsWindowVisible(hwnd):
                found.append(hwnd)
        return True

    proc = ctypes.WINFUNCTYPE(ctypes.c_bool, ctypes.c_void_p, ctypes.c_void_p)(callback)
    ctypes.windll.user32.EnumWindows(proc, 0)
    before = set(found)

    outdir = os.path.join(WORK, "out", "console_probe")
    shutil.rmtree(outdir, ignore_errors=True)
    os.makedirs(outdir, exist_ok=True)
    cancel = threading.Event()
    worker = converter.BatchConverter([paths["pdf"]], "pdf2word", "editable", outdir,
                                      True, cancel, lambda kind, payload: None)
    thread = threading.Thread(target=worker.run, daemon=True)
    thread.start()
    seen: set[int] = set()
    while thread.is_alive():
        found.clear()
        ctypes.windll.user32.EnumWindows(proc, 0)
        seen |= (set(found) - before)
        time.sleep(0.1)
    thread.join(30)
    check("no console window during Word conversion", not seen, f"seen={len(seen)}")


def check_settings_roundtrip(paths: dict[str, str]) -> None:
    """GUI state: conversion switch, mode memory and settings persistence."""
    try:
        import tkinter as tk

        import app as gui
    except Exception as exc:  # pragma: no cover
        check("GUI import", False, str(exc))
        return

    config_dir = gui.config_dir()
    settings_path = config_dir / "settings.json"
    backup = settings_path.read_text("utf-8") if settings_path.exists() else None
    try:
        # a clean profile: the default conversion must not depend on leftovers
        settings_path.unlink(missing_ok=True)
        root = tk.Tk()
        root.withdraw()
        instance = gui.App(root)

        check("default conversion is pdf2word",
              instance.conversion == "pdf2word", instance.conversion)
        check("mode cards built for pdf2word",
              set(instance.cards) == set(converter.CONVERSIONS["pdf2word"]["order"]),
              str(list(instance.cards)))

        instance._select_conversion("ppt2pdf")
        check("switch to ppt2pdf rebuilds cards",
              set(instance.cards) == set(converter.CONVERSIONS["ppt2pdf"]["order"]),
              str(list(instance.cards)))
        check("ppt2pdf default mode selected",
              instance.mode == converter.CONVERSIONS["ppt2pdf"]["default_mode"],
              instance.mode)

        instance._select_mode("low")
        instance._select_conversion("pdf2ppt")
        instance._select_conversion("ppt2pdf")
        check("mode memory per conversion", instance.mode == "low", instance.mode)

        # pretend the user is on another type, then switch: mismatched files go
        instance.conversion = "pdf2ppt"
        instance.mode = instance.modes["pdf2ppt"]
        instance.files = [paths["pdf"], paths["pptx"]]
        instance._select_conversion("ppt2pdf")
        check("switching drops mismatched files",
              [os.path.basename(p) for p in instance.files] == ["deck.pptx"],
              str([os.path.basename(p) for p in instance.files]))

        instance._save_settings()
        import json

        saved = json.loads((config_dir / "settings.json").read_text("utf-8"))
        check("settings persist conversion+modes",
              saved.get("conversion") == "ppt2pdf" and saved.get("modes", {}).get("ppt2pdf") == "low",
              str(saved)[:120])
        check("disclaimer state starts empty", saved.get("disclaimer", "") == "",
              repr(saved.get("disclaimer")))
        instance.disclaimer = gui.DISCLAIMER_VERSION
        instance._save_settings()
        check("accepted disclaimer is remembered", gui.disclaimer_accepted(),
              repr(gui._read_settings_dict().get("disclaimer")))
        instance._save_settings()
        check("disclaimer survives later settings saves", gui.disclaimer_accepted(), "")
        root.destroy()
    except Exception as exc:
        check("GUI checks", False, f"{exc.__class__.__name__}: {exc}")
    finally:
        try:
            if backup is None:
                settings_path.unlink(missing_ok=True)
            else:
                settings_path.write_text(backup, "utf-8")
        except Exception:
            pass


def check_disclaimer_phrase() -> None:
    """The first-run gate must accept only the real sentence (tolerating layout)."""
    import app as gui

    phrase = gui.CONSENT_PHRASE
    loosely = "  " + phrase.replace("，", ",").replace("。", ".") + "\n"
    ok = (gui._consent_matches(phrase)
          and gui._consent_matches(loosely)
          and not gui._consent_matches(phrase[:-4])
          and not gui._consent_matches("")
          and not gui._consent_matches("同意"))
    check("disclaimer gate needs the exact sentence", ok, phrase)


def check_outward_disclaimers() -> None:
    """对外文档必须带三处声明：文首 7 条特别说明、DISCLAIMER.md 11 节、文末声明行 + 安装包提示。"""
    root = os.path.dirname(os.path.abspath(__file__))
    required = ["README.md", "CHANGELOG.md", "RELEASING.md",
                os.path.join(".github", "RELEASE_BODY.md")]
    missing = []
    for name in required:
        try:
            text = open(os.path.join(root, name), encoding="utf-8").read()
        except Exception:
            missing.append(f"{name}(读不到)")
            continue
        if "特别说明（请先阅读）" not in text or "DISCLAIMER.md" not in text:
            missing.append(name)
    check("每个对外文档都带「特别说明」块", not missing, str(missing))

    no_footer = []
    for name in required:
        try:
            text = open(os.path.join(root, name), encoding="utf-8").read()
        except Exception:
            no_footer.append(f"{name}(读不到)")
            continue
        if "完整免责声明见" not in text:
            no_footer.append(name)
    check("每个对外文档结尾都有声明行", not no_footer, str(no_footer))

    try:
        disclaimer = open(os.path.join(root, "DISCLAIMER.md"), encoding="utf-8").read()
    except Exception:
        disclaimer = ""
    wanted = [f"## {index} " for index in range(1, 12)]
    absent = [item.strip() for item in wanted if item not in disclaimer]
    check("DISCLAIMER.md 是 11 节", not absent, f"缺: {absent}")

    try:
        readme = open(os.path.join(root, "README.md"), encoding="utf-8").read()
    except Exception:
        readme = ""
    bullets = sum(1 for line in readme.splitlines() if line.startswith("> - "))
    check("README 文首「特别说明」正好 7 条", bullets == 7, f"实际 {bullets} 条")

    notice = os.path.join(root, "installer", "DISCLAIMER.txt")
    try:
        with open(notice, encoding="utf-8-sig") as handle:
            ok = "仅供学习" in handle.read()
    except Exception:
        ok = False
    check("安装包带免责声明提示", ok, notice)


def check_cli_validation() -> None:
    """Invalid CLI combinations must return 2 and not touch the disk."""
    import app as gui

    code = gui._cli(["--convert", "--conversion", "nope", "--mode", "nope", "X:\\missing.pdf"])
    check("cli rejects unknown conversion/empty input", code == 2, f"exit={code}")

    code = gui._cli(["--convert", "--conversion", "pdf2word", "missing.pdf"])
    check("cli reports missing file", code == 2, f"exit={code}")


def check_cli_success(paths: dict[str, str]) -> None:
    """A real CLI run must exit 0 — this is what the .exe is driven by."""
    import app as gui

    outdir = os.path.join(WORK, "out", "cli")
    shutil.rmtree(outdir, ignore_errors=True)
    os.makedirs(outdir, exist_ok=True)
    for conversion, mode, source, suffix in (
        ("pdf2word", "image120", paths["pdf"], ".docx"),
        ("pdf2ppt", "image120", paths["pdf"], ".pptx"),
        ("ppt2pdf", "native", paths["pptx"], ".pdf"),
    ):
        code = gui._cli(["--convert", "--conversion", conversion, "--mode", mode,
                         "--out", outdir, source])
        produced = [f for f in os.listdir(outdir) if f.endswith(suffix)]
        check(f"cli {conversion}/{mode} exits 0", code == 0 and bool(produced),
              f"exit={code} files={os.listdir(outdir)}")


def main() -> int:
    print(f"PDF2Office self-test  (python {sys.version.split()[0]})")
    print(f"work dir: {WORK}\n")
    paths = build_inputs()

    check_conversion_matrix(paths)
    check_page_counts(paths)
    check_page_size_preserved(paths)
    check_batch(paths)
    check_overwrite(paths)
    check_target_guard(paths)
    check_cancel(paths)
    check_bad_inputs(paths)
    check_wrong_extension(paths)
    check_scanned_note(paths)
    check_picture_fidelity(paths)
    check_jpeg_quality_probe()
    check_disclaimer_phrase()
    check_outward_disclaimers()
    check_settings_roundtrip(paths)
    check_cli_validation()
    check_cli_success(paths)
    check_no_console_window(paths)

    passed = sum(1 for _, ok, _ in RESULTS if ok)
    total = len(RESULTS)
    print(f"\n{passed}/{total} checks passed")
    if passed != total:
        print("failures:")
        for name, ok, detail in RESULTS:
            if not ok:
                print(f"  - {name}: {detail}")
        return 1
    return 0


if __name__ == "__main__":
    sys.exit(main())
