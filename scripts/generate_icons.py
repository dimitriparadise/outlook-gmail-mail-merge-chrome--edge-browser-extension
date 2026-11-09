"""Draws the extension icons (icons/icon-16.png ... icon-128.png). Requires Pillow."""
from pathlib import Path
from PIL import Image, ImageDraw


ROOT = Path(__file__).resolve().parents[1]
ICON_DIR = ROOT / "icons"

BLUE = "#2563eb"
BLUE_DARK = "#1d4ed8"


def rounded(draw, box, radius, fill, outline=None, width=1):
    draw.rounded_rectangle(box, radius=radius, fill=fill, outline=outline, width=width)


def make_icon(size):
    scale = size / 128
    img = Image.new("RGBA", (size, size), (0, 0, 0, 0))
    d = ImageDraw.Draw(img)
    rounded(d, (0, 0, size - 1, size - 1), int(26 * scale), BLUE, BLUE_DARK, max(1, int(2 * scale)))

    pad = int(26 * scale)
    envelope = (pad, int(34 * scale), size - pad, int(88 * scale))
    rounded(d, envelope, int(9 * scale), "white")
    d.line((envelope[0] + int(6 * scale), envelope[1] + int(8 * scale), size // 2, envelope[1] + int(34 * scale), envelope[2] - int(6 * scale), envelope[1] + int(8 * scale)), fill=BLUE, width=max(1, int(5 * scale)))
    d.line((envelope[0] + int(6 * scale), envelope[3] - int(8 * scale), size // 2, envelope[1] + int(36 * scale), envelope[2] - int(6 * scale), envelope[3] - int(8 * scale)), fill="#93c5fd", width=max(1, int(3 * scale)))

    check = [
        (int(48 * scale), int(96 * scale)),
        (int(60 * scale), int(108 * scale)),
        (int(84 * scale), int(78 * scale)),
    ]
    d.line(check, fill="white", width=max(2, int(9 * scale)), joint="curve")
    return img


def main():
    ICON_DIR.mkdir(exist_ok=True)
    for size in (16, 32, 48, 128):
        make_icon(size).save(ICON_DIR / f"icon-{size}.png")


if __name__ == "__main__":
    main()
