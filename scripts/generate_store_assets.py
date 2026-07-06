from pathlib import Path
from PIL import Image, ImageDraw, ImageFont


ROOT = Path(__file__).resolve().parents[1]
ICON_DIR = ROOT / "icons"
ASSET_DIR = ROOT / "store-assets"
SCREENSHOT_DIR = ASSET_DIR / "screenshots"
PROMO_DIR = ASSET_DIR / "promo"
RAW_POPUP_DIR = ASSET_DIR / "raw-popup"


BLUE = "#2563eb"
BLUE_DARK = "#1d4ed8"
SLATE = "#1f2937"
MUTED = "#475569"
LIGHT = "#f8fafc"
BORDER = "#dbe4f0"
GREEN = "#166534"
RED = "#b91c1c"


def font(size, bold=False):
    candidates = [
        "/System/Library/Fonts/Supplemental/Arial Bold.ttf" if bold else "/System/Library/Fonts/Supplemental/Arial.ttf",
        "/System/Library/Fonts/Supplemental/Helvetica Bold.ttf" if bold else "/System/Library/Fonts/Supplemental/Helvetica.ttf",
        "/Library/Fonts/Arial.ttf",
    ]
    for candidate in candidates:
        try:
            return ImageFont.truetype(candidate, size)
        except OSError:
            continue
    return ImageFont.load_default()


def rounded(draw, box, radius, fill, outline=None, width=1):
    draw.rounded_rectangle(box, radius=radius, fill=fill, outline=outline, width=width)


def draw_text(draw, xy, text, size=28, fill=SLATE, bold=False, max_width=None, line_gap=8):
    f = font(size, bold)
    if not max_width:
        draw.multiline_text(xy, text, font=f, fill=fill, spacing=line_gap)
        bbox = draw.multiline_textbbox(xy, text, font=f, spacing=line_gap)
        return bbox[3]

    lines = []
    for paragraph in str(text).split("\n"):
        if paragraph == "":
            lines.append("")
            continue
        words = paragraph.split()
        current = ""
        for word in words:
            test = f"{current} {word}".strip()
            if draw.textbbox((0, 0), test, font=f)[2] <= max_width or not current:
                current = test
            else:
                lines.append(current)
                current = word
        if current:
            lines.append(current)
    if not lines:
        lines.append("")

    y = xy[1]
    for line in lines:
        if line:
            draw.text((xy[0], y), line, font=f, fill=fill)
        else:
            y += size // 2
            continue
        y += size + line_gap
    return y


def paste_icon(base, size, xy):
    icon = make_icon(size)
    base.paste(icon, xy, icon)


def popup_crop(name, top=0, height=860, scale=0.78):
    source = Image.open(RAW_POPUP_DIR / name).convert("RGB")
    width = min(430, source.width)
    bottom = min(source.height, top + height)
    crop = source.crop((0, top, width, bottom))
    return crop.resize((int(crop.width * scale), int(crop.height * scale)), Image.LANCZOS)


def paste_popup(canvas, popup, xy):
    shadow = Image.new("RGBA", (popup.width + 28, popup.height + 28), (0, 0, 0, 0))
    sd = ImageDraw.Draw(shadow)
    rounded(sd, (14, 14, shadow.width - 14, shadow.height - 14), 18, (15, 23, 42, 24))
    canvas.paste(shadow.convert("RGB"), (xy[0] - 14, xy[1] - 14), shadow)
    canvas.paste(popup, xy)


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


def save_icons():
    ICON_DIR.mkdir(exist_ok=True)
    for size in (16, 32, 48, 128):
        make_icon(size).save(ICON_DIR / f"icon-{size}.png")


def base_canvas():
    img = Image.new("RGB", (1280, 800), LIGHT)
    d = ImageDraw.Draw(img)
    d.rectangle((0, 0, 1280, 92), fill="white")
    d.line((0, 92, 1280, 92), fill=BORDER, width=2)
    paste_icon(img, 54, (48, 19))
    draw_text(d, (120, 24), "Mail Merge Draft Helper", 32, SLATE, True)
    draw_text(d, (120, 58), "Personalized Gmail and Outlook drafts from CSV", 18, MUTED)
    return img, d


def draw_popup(d, x, y, w=480):
    rounded(d, (x, y, x + w, y + 650), 20, "white", BORDER, 2)
    draw_text(d, (x + 24, y + 22), "Mail Merge Drafts", 26, SLATE, True)
    draw_text(d, (x + 24, y + 58), "CSV columns can be used as variables like {{Course}}.", 15, MUTED, max_width=w - 48)

    y0 = y + 104
    fields = [
        ("1. Student CSV", "Name,Email,Course,Section,DueDate\nJohn,john@example.com,ISOM 210,A,Friday"),
        ("3. Subject template", "Reminder for {{Course}}"),
        ("6. Body template", "Hi {{Name}},\n\nThis is a quick reminder about {{Course}}."),
    ]
    for label, value in fields:
        draw_text(d, (x + 24, y0), label, 16, SLATE, True)
        box_h = 90 if "\n" in value else 50
        rounded(d, (x + 24, y0 + 26, x + w - 24, y0 + 26 + box_h), 10, "#ffffff", "#cbd5e1", 2)
        draw_text(d, (x + 38, y0 + 40), value, 16, SLATE, max_width=w - 76, line_gap=3)
        y0 += box_h + 70

    rounded(d, (x + 24, y0, x + w - 24, y0 + 52), 12, BLUE)
    draw_text(d, (x + 144, y0 + 13), "Generate Drafts", 20, "white", True)
    return y + 650


def screenshot_overview():
    img, d = base_canvas()
    popup = popup_crop("01-initial-popup.png", top=0, height=850, scale=0.78)
    paste_popup(img, popup, (82, 116))
    draw_text(d, (620, 174), "Turn CSV rows into reviewed email drafts", 46, SLATE, True, 560)
    draw_text(d, (622, 310), "Import a list, write one reusable template, and preview every personalized Gmail or Outlook draft before opening it.", 26, MUTED, max_width=540, line_gap=10)
    bullets = [
        "Template variables from your CSV headers",
        "Gmail and Outlook Web compose support",
        "Local saved progress with one-click reset",
    ]
    y = 470
    for bullet in bullets:
        d.ellipse((622, y + 7, 638, y + 23), fill=BLUE)
        draw_text(d, (654, y), bullet, 24, SLATE)
        y += 58
    img.save(SCREENSHOT_DIR / "01-overview.png")


def screenshot_preview():
    img, d = base_canvas()
    popup = popup_crop("03-second-preview.png", top=570, height=830, scale=0.78)
    paste_popup(img, popup, (782, 112))
    draw_text(d, (90, 176), "Review the real generated draft", 46, SLATE, True, 610)
    draw_text(d, (92, 312), "The store screenshot now uses the actual popup UI rendered from popup.html, popup.css, and popup.js.", 26, MUTED, max_width=600, line_gap=10)
    draw_text(d, (92, 468), "Use Previous / Next Preview to inspect recipients, then open exactly the current preview or a selected range.", 26, MUTED, max_width=610, line_gap=10)
    img.save(SCREENSHOT_DIR / "02-preview-and-range.png")


def screenshot_privacy():
    img, d = base_canvas()
    popup = popup_crop("04-auto-send-labels.png", top=420, height=840, scale=0.78)
    paste_popup(img, popup, (806, 112))
    draw_text(d, (80, 150), "Built for careful sending", 46, SLATE, True)
    draw_text(d, (82, 220), "Preflight checks help catch common mistakes before drafts are opened.", 25, MUTED, max_width=600)

    checks = [
        ("Unknown template variables", "Find typos like {{Coursee}} before they create blank content."),
        ("Recipient warnings", "Catch invalid addresses, duplicates, and To/CC/BCC overlap."),
        ("Local-only state", "CSV text and generated drafts stay in Chrome extension storage."),
        ("Clear Saved Data", "Remove saved CSV, templates, generated drafts, and progress anytime."),
    ]
    y = 340
    for title, body in checks:
        rounded(d, (80, y, 594, y + 84), 14, "white", BORDER, 2)
        d.ellipse((108, y + 28, 134, y + 54), fill="#dcfce7")
        d.line((114, y + 41, 123, y + 50, 138, y + 29), fill=GREEN, width=4)
        draw_text(d, (154, y + 16), title, 22, SLATE, True)
        draw_text(d, (154, y + 46), body, 16, MUTED, max_width=395)
        y += 102

    img.save(SCREENSHOT_DIR / "03-safety-and-privacy.png")


def promo_tile():
    img = Image.new("RGB", (440, 280), BLUE)
    d = ImageDraw.Draw(img)
    paste_icon(img, 74, (36, 44))
    draw_text(d, (130, 54), "Mail Merge", 32, "white", True)
    draw_text(d, (130, 92), "Draft Helper", 32, "white", True)
    draw_text(d, (38, 160), "Personalized Gmail and Outlook drafts from CSV lists.", 22, "#e0f2fe", max_width=360, line_gap=6)
    img.save(PROMO_DIR / "small-promo-tile-440x280.png")


def main():
    ICON_DIR.mkdir(exist_ok=True)
    SCREENSHOT_DIR.mkdir(parents=True, exist_ok=True)
    PROMO_DIR.mkdir(parents=True, exist_ok=True)
    save_icons()
    screenshot_overview()
    screenshot_preview()
    screenshot_privacy()
    promo_tile()


if __name__ == "__main__":
    main()
