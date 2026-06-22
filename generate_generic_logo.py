from pathlib import Path
from PIL import Image, ImageDraw, ImageFont

ROOT = Path(__file__).resolve().parent
size = 512
image = Image.new("RGBA", (size, size), "#0B4F8A")
draw = ImageDraw.Draw(image)
draw.rounded_rectangle((42, 42, 470, 470), radius=96, fill="#1267A8", outline="#D8ECFA", width=16)
draw.ellipse((132, 116, 380, 364), fill="#F8FBFE")
draw.rectangle((245, 160, 267, 346), fill="#1267A8")
draw.rectangle((174, 242, 338, 264), fill="#1267A8")
try:
    font = ImageFont.truetype("arialbd.ttf", 72)
except Exception:
    font = ImageFont.load_default()
text = "LSP"
box = draw.textbbox((0, 0), text, font=font)
draw.text(((size - (box[2] - box[0])) / 2, 382), text, font=font, fill="white")
image.save(ROOT / "generic_logo.png")
image.save(ROOT / "generic_logo.ico", sizes=[(16, 16), (32, 32), (48, 48), (64, 64), (128, 128), (256, 256)])
