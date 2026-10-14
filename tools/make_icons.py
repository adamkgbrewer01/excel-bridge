"""Draw the Excel Bridge ribbon icons: three bars, one highlighted, on a blue tile.

Run:  python3 tools/make_icons.py   (from the Add-in folder)
"""
from pathlib import Path

from PIL import Image, ImageDraw

OUT = Path(__file__).resolve().parent.parent / "assets"
BLUE = (42, 120, 214, 255)
PALE = (205, 226, 251, 255)
WHITE = (255, 255, 255, 255)


def draw(size):
    scale = 8
    big = size * scale
    img = Image.new("RGBA", (big, big), (0, 0, 0, 0))
    d = ImageDraw.Draw(img)
    d.rounded_rectangle([0, 0, big - 1, big - 1], radius=big * 0.22, fill=BLUE)
    pad = big * 0.2
    gap = big * 0.06
    width = (big - 2 * pad - 2 * gap) / 3
    base = big - pad
    for i, (h, colour) in enumerate([(0.45, PALE), (0.75, WHITE), (0.3, PALE)]):
        x0 = pad + i * (width + gap)
        d.rounded_rectangle([x0, base - h * (big - 2 * pad), x0 + width, base], radius=width * 0.25, fill=colour)
    return img.resize((size, size), Image.LANCZOS)


if __name__ == "__main__":
    OUT.mkdir(exist_ok=True)
    for size in (16, 32, 64, 80, 128):
        draw(size).save(OUT / f"icon-{size}.png")
    print("icons written to", OUT)
