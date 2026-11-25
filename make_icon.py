"""Generate the PDF2Office icon (rounded tile + page + DOC)."""
import os
from pathlib import Path

from PIL import Image, ImageDraw, ImageFont

OUT = Path(__file__).resolve().parent / "assets"
OUT.mkdir(parents=True, exist_ok=True)
S = 1024
img = Image.new("RGBA", (S, S), (0, 0, 0, 0))
d = ImageDraw.Draw(img)

# rounded blue tile
d.rounded_rectangle([0, 0, S - 1, S - 1], radius=210, fill=(37, 99, 235, 255))
d.rounded_rectangle([36, 36, S - 37, S - 37], radius=185, fill=(59, 130, 246, 255))
d.rounded_rectangle([36, 36, S - 37, S // 2], radius=185, fill=(59, 130, 246, 255))

# white page
page = [250, 170, 774, 854]
d.rounded_rectangle(page, radius=36, fill=(255, 255, 255, 255))
fold = [(620, 170), (774, 324), (620, 324)]
d.polygon(fold, fill=(219, 234, 254, 255))
d.line([(620, 170), (620, 324), (774, 324)], fill=(191, 219, 254, 255), width=6)

font = None
for path in (r"C:\Windows\Fonts\msyhbd.ttc", r"C:\Windows\Fonts\arialbd.ttf",
             r"C:\Windows\Fonts\segoeuib.ttf"):
    if os.path.exists(path):
        try:
            font = ImageFont.truetype(path, 190)
            break
        except Exception:
            pass
if font is None:
    font = ImageFont.load_default()

text = "DOC"
bbox = d.textbbox((0, 0), text, font=font)
tw, th = bbox[2] - bbox[0], bbox[3] - bbox[1]
tx = (S - tw) / 2 - bbox[0]
ty = 520 - bbox[1]
d.text((tx, ty), text, font=font, fill=(37, 99, 235, 255))

# red "source" accent (PDF -> Word)
d.rounded_rectangle([400, 660, 624, 700], radius=20, fill=(239, 68, 68, 255))

sizes = [256, 128, 64, 48, 32, 24, 16]
img.save(OUT / "icon.png")
img.save(OUT / "icon.ico", sizes=[(s, s) for s in sizes])
print("saved", OUT / "icon.ico", os.path.getsize(OUT / "icon.ico"), "bytes")
