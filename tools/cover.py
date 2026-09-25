"""Simple listing cover: red header with Swiss cross, subtitle and bullet list."""
from PIL import Image, ImageDraw, ImageFont

BOLD = "/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf"
REG = "/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf"


def make_cover(lines, subtitle, bullets, out):
    w, h = 1600, 1200
    img = Image.new("RGB", (w, h), "#FFFFFF")
    d = ImageDraw.Draw(img)
    d.rectangle([0, 0, w, 420], fill="#D52B1E")
    cx, cy, s = 170, 210, 46
    d.rectangle([cx - s * 2, cy - s * 2, cx + s * 2, cy + s * 2], fill="#FFFFFF")
    d.rectangle([cx - s * 0.6, cy - s * 1.6, cx + s * 0.6, cy + s * 1.6], fill="#D52B1E")
    d.rectangle([cx - s * 1.6, cy - s * 0.6, cx + s * 1.6, cy + s * 0.6], fill="#D52B1E")
    big = ImageFont.truetype(BOLD, 84)
    for i, line in enumerate(lines[:2]):
        d.text((330, 115 + i * 110), line, font=big, fill="#FFFFFF")
    d.text((80, 480), subtitle, font=ImageFont.truetype(BOLD, 50), fill="#1F2937")
    small = ImageFont.truetype(REG, 38)
    y = 590
    for b in bullets[:5]:
        d.ellipse([90, y + 12, 112, y + 34], fill="#D52B1E")
        d.text((135, y), b, font=small, fill="#374151")
        y += 95
    img.save(out)
