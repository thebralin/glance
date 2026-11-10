"""Generates per-file-type icons for Explorer (src-tauri/icons/filetypes/*.ico).

The installer hook (src-tauri/windows/hooks.nsh) points each ProgId's DefaultIcon
at these files, so .pdf/.docx/.pptx/.xlsx opened with Glance look different.

Run once after changing the design:  python scripts/make-filetype-icons.py
Only some types:                      python scripts/make-filetype-icons.py markdown text
Requires Pillow (pip install pillow).
"""
import sys
from pathlib import Path

from PIL import Image, ImageDraw, ImageFont

ROOT = Path(__file__).resolve().parent.parent
OUT = ROOT / "src-tauri" / "icons" / "filetypes"

TYPES = {
    # file name: (label, band color, darker shade for the fold)
    "pdf": ("PDF", "#E5252A", "#A8181C"),
    "word": ("DOC", "#2B6CD8", "#1A4A9C"),
    "powerpoint": ("PPT", "#EA6A1F", "#B04A10"),
    "excel": ("XLS", "#1E8A4C", "#136034"),
    "markdown": ("MD", "#6E56CF", "#4C3A9E"),
    "text": ("TXT", "#64748B", "#475263"),
    "json": ("{ }", "#C27C0E", "#8A5608"),
}
BRAND = ("#5B6CFF", "#8E4DFF")  # app icon gradient
SIZES = [16, 20, 24, 32, 40, 48, 64, 256]
SS = 4  # supersampling factor

FONT_CANDIDATES = [
    "C:/Windows/Fonts/segoeuib.ttf",
    "C:/Windows/Fonts/arialbd.ttf",
    "DejaVuSans-Bold.ttf",
]


def load_font(px: int) -> ImageFont.FreeTypeFont:
    for name in FONT_CANDIDATES:
        try:
            return ImageFont.truetype(name, px)
        except OSError:
            continue
    return ImageFont.load_default()


def render(size: int, label: str, color: str, shade: str) -> Image.Image:
    s = size * SS
    img = Image.new("RGBA", (s, s), (0, 0, 0, 0))
    d = ImageDraw.Draw(img)
    u = s / 32  # design grid: 32x32 units

    small = size < 32
    # page with a folded top-right corner
    left, top, right, bottom = 5 * u, 1.5 * u, 27 * u, 30.5 * u
    fold = 7 * u
    r = 2 * u
    outline = "#9AA0B4"
    stroke = max(1, round(u * (0.9 if small else 0.6)))

    d.rounded_rectangle([left, top, right, bottom], radius=r, fill="white", outline=outline, width=stroke)
    # cut the corner and draw the fold
    d.polygon([(right - fold - u, top - u), (right + u, top - u), (right + u, top + fold + u)], fill=(0, 0, 0, 0))
    d.line([(right - fold, top), (right, top + fold)], fill=outline, width=stroke)
    d.polygon([(right - fold, top), (right - fold, top + fold), (right, top + fold)], fill=shade)

    # colored band with the type label; at tiny sizes the color alone does the job
    band_top = 14 * u if not small else 13 * u
    band_bottom = 23 * u if not small else 27 * u
    d.rounded_rectangle([2 * u, band_top, 23 * u, band_bottom], radius=1.5 * u, fill=color)

    if not small:
        font = load_font(round(7.4 * u))
        cx, cy = 12.5 * u, (band_top + band_bottom) / 2
        d.text((cx, cy), label, font=font, fill="white", anchor="mm")

        # text lines on the page
        for i, w in enumerate((12, 15, 10)):
            y = 5 * u + i * 2.6 * u
            d.rounded_rectangle([8.5 * u, y, (8.5 + w) * u, y + 1.2 * u], radius=0.6 * u, fill="#C9CEFF")

        # Glance magnifier badge in the brand color, bottom-right
        bx, by, br = 25.5 * u, 26.5 * u, 4.6 * u
        d.ellipse([bx - br, by - br, bx + br, by + br], fill=BRAND[1], outline="white", width=round(0.9 * u))
        lr = 2.1 * u
        lx, ly = bx - 0.6 * u, by - 0.6 * u
        d.line([(lx + 1.2 * u, ly + 1.2 * u), (bx + 2.6 * u, by + 2.6 * u)], fill="white", width=round(1.3 * u))
        d.ellipse([lx - lr, ly - lr, lx + lr, ly + lr], fill=BRAND[1], outline="white", width=round(0.9 * u))

    return img.resize((size, size), Image.LANCZOS)


def main() -> None:
    OUT.mkdir(parents=True, exist_ok=True)
    only = sys.argv[1:]
    unknown = set(only) - TYPES.keys()
    if unknown:
        sys.exit(f"unknown types: {', '.join(sorted(unknown))}; known: {', '.join(TYPES)}")
    for name, (label, color, shade) in TYPES.items():
        if only and name not in only:
            continue
        frames = [render(sz, label, color, shade) for sz in SIZES]
        biggest = frames[-1]
        biggest.save(
            OUT / f"{name}.ico",
            format="ICO",
            sizes=[(sz, sz) for sz in SIZES],
            append_images=frames[:-1],
        )
        print("wrote", OUT / f"{name}.ico")


if __name__ == "__main__":
    main()
