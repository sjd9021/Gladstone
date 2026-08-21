"""Geometry check for generated photo sheets, measured through Word itself.

Builds a sheet in each layout, exports it to PDF with Microsoft Word via
AppleScript, and measures where each photo and its caption actually land.
Every test photo carries a yellow vertical centre line, so a photo's true
centre is read straight off the page rather than inferred from its edges.

Checks, for both layouts:
  * the photo block meets the page margin it is supposed to (right margin for
    WK Webster, page centre for Gladstone)
  * every caption is centred under its own photo

Renderer notes
--------------
Word is the only renderer trusted here; the alternatives each lie in a way
that matters. LibreOffice inserts ~0.31 cm of its own after every inline
image, so right-aligned rows look 0.6 cm out when they are not. QuickLook
lays images out correctly but ignores tblLayout/gridCol/jc, rendering the
caption table at the left margin with shrink-to-fit cells. docx-preview
handles both but placed the real reports' right edge 0.3 cm off.

Requires macOS with Word installed, and Automation permission for the terminal
running it. Word must not already be sitting on a modal dialog, or the export
silently produces nothing.

Working files go in ~/Documents/photo-sheet-verify rather than a system temp
directory. Word is sandboxed, and opening a file in /var/folders or /tmp makes
it raise a "Grant File Access" sheet that nobody can see and that blocks
AppleScript until it times out - which looks exactly like Word hanging. The
first run against a new folder still shows that sheet once; click Grant Access
and Word remembers it, which is why the folder is stable and is not deleted
between runs.

Usage:
    python verify_layout.py
"""

import io
import subprocess
import sys
from pathlib import Path

import numpy as np
from PIL import Image, ImageDraw
from pdf2image import convert_from_path

DPI = 300
PPC = DPI / 2.54                 # pixels per cm
PAGE_W_CM = 21.0
TOLERANCE_CM = 0.15              # 1.5 mm; measured error is ~0.1 mm

EXPECTED = {
    # layout: (description, target value in cm)
    "wkw": ("photo block right edge", 19.10),        # A4 less 1.90 cm margin
    "gladstone": ("photo block midpoint", 10.50),    # page centre
}


def word_export_pdf(docx_path, pdf_path):
    """Export a .docx to PDF using Word's own layout engine."""
    script = f'''
    tell application "Microsoft Word"
        activate
        close every document saving no
        delay 1
        open POSIX file "{docx_path}"
        delay 3
        save as document 1 file name (POSIX file "{pdf_path}" as string) file format format PDF
        delay 2
        close every document saving no
    end tell
    '''
    subprocess.run(["osascript", "-e", script], capture_output=True)
    if not Path(pdf_path).exists():
        raise RuntimeError(f"Word did not produce {pdf_path}")


def measure(pdf_path):
    """Yield (photo_centres, caption_centres, block_left, block_right) per row."""
    img = convert_from_path(pdf_path, dpi=DPI, first_page=1, last_page=1)[0]
    a = np.asarray(img.convert("RGB")).astype(int)
    width = img.width

    saturated = (a.max(axis=2) - a.min(axis=2)) > 40      # the test photos
    yellow = (a[:, :, 0] > 170) & (a[:, :, 1] > 170) & (a[:, :, 2] < 130)
    # Caption text only: black-ish pixels that are not part of a photo. The
    # rows pack tightly enough that the caption search strip below one photo
    # row can clip the top of the next row's photos; without this exclusion
    # those photo pixels get averaged into the "caption" centre.
    ink = (np.asarray(img.convert("L")) < 180) & ~saturated

    rows = np.flatnonzero(saturated.sum(axis=1) > 0.2 * width)
    if rows.size == 0:
        return
    brk = np.flatnonzero(np.diff(rows) > 20)
    starts = np.r_[rows[0], rows[brk + 1]]
    ends = np.r_[rows[brk], rows[-1]] + 1

    for idx, (y0, y1) in enumerate(zip(starts, ends)):
        band = saturated[y0:y1]
        cols = np.flatnonzero(band.sum(axis=0) > 0.5 * (y1 - y0))
        if cols.size == 0:
            continue
        cbrk = np.flatnonzero(np.diff(cols) > 10)
        cs = np.r_[cols[0], cols[cbrk + 1]]
        ce = np.r_[cols[cbrk], cols[-1]] + 1
        photos = [(s, e) for s, e in zip(cs, ce) if (e - s) / PPC > 3]
        if not photos:
            continue

        # The caption strip runs from this photo row's bottom to just above
        # the next photo row (or a fixed 1.6 cm on the last row). A fixed
        # height bled into the following row once the spacing fix packed the
        # rows tighter, and stray dark pixels there skewed the caption centre.
        strip_end = (int(starts[idx + 1]) - 2 if idx + 1 < len(starts)
                     else y1 + int(1.6 * PPC))
        caption_band = ink[y1:strip_end]
        centres, captions = [], []
        for s, e in photos:
            strip = yellow[y0:y1, s:e]
            marks = np.flatnonzero(strip.sum(axis=0) > 0.5 * (y1 - y0))
            centres.append((s + marks.mean()) / PPC if marks.size
                           else (s + e) / 2 / PPC)
            lo, hi = max(0, s - int(0.15 * PPC)), min(width, e + int(0.15 * PPC))
            found = np.flatnonzero(caption_band[:, lo:hi].any(axis=0))
            captions.append((lo + (found[0] + found[-1]) / 2) / PPC
                            if found.size else None)

        yield centres, captions, photos[0][0] / PPC, photos[-1][1] / PPC


def build_test_sheet(app, layout_key, tmp):
    items = []
    for i in range(5):
        im = Image.new("RGB", (1200, 880), (35 + 38 * i, 90, 140))
        ImageDraw.Draw(im).line((600, 0, 600, 880), fill=(255, 255, 0), width=8)
        buf = io.BytesIO()
        im.save(buf, "JPEG")
        items.append((f"Caption {i + 1} describing the damage found", buf.getvalue()))
    path = tmp / f"verify_{layout_key}.docx"
    path.write_bytes(app.build_docx(items, layout_key))
    return path


def check(app, layout_key, tmp):
    docx_path = build_test_sheet(app, layout_key, tmp)
    pdf_path = tmp / f"verify_{layout_key}.pdf"
    word_export_pdf(docx_path, pdf_path)

    what, target = EXPECTED[layout_key]
    print(f"\n=== {app.LAYOUTS[layout_key]['label']} ===")

    worst_caption, worst_block, rows = 0.0, 0.0, 0
    for centres, captions, left, right in measure(pdf_path):
        rows += 1
        actual = right if layout_key == "wkw" else (left + right) / 2
        worst_block = max(worst_block, abs(actual - target))
        print(f"  row {rows}: block {left:.3f}..{right:.3f} cm   "
              f"{what} {actual:.3f} (want {target:.2f})")
        for i, (pc, cc) in enumerate(zip(centres, captions)):
            if cc is None:
                continue
            off = cc - pc
            worst_caption = max(worst_caption, abs(off))
            flag = "OK " if abs(off) < TOLERANCE_CM else "OFF"
            print(f"    {flag} photo {i + 1}: centre {pc:.3f}  "
                  f"caption {cc:.3f}  offset {off:+.3f} cm")

    ok = rows > 0 and worst_caption < TOLERANCE_CM and worst_block < TOLERANCE_CM
    print(f"  worst caption offset {worst_caption:.3f} cm, "
          f"worst block error {worst_block:.3f} cm -> {'PASS' if ok else 'FAIL'}")
    return ok


def main():
    sys.path.insert(0, str(Path(__file__).parent))
    import app

    # A stable, non-hidden folder: Word grants file access per folder, and
    # deleting and recreating it makes Word re-prompt on every run.
    tmp = Path.home() / "Documents" / "photo-sheet-verify"
    tmp.mkdir(parents=True, exist_ok=True)
    try:
        ok = all(check(app, key, tmp) for key in ("wkw", "gladstone"))
    finally:
        for leftover in tmp.glob("verify_*"):
            leftover.unlink(missing_ok=True)
    print("\nRESULT:", "PASS" if ok else "FAIL")
    return 0 if ok else 1


if __name__ == "__main__":
    raise SystemExit(main())
