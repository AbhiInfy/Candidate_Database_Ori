from pathlib import Path

from PIL import Image, ImageDraw, ImageFont


ROOT = Path(__file__).resolve().parent
OUT = ROOT / "Hiring_Automation_Tool.ico"


def font(size, bold=False):
    candidates = [
        "C:/Windows/Fonts/segoeuib.ttf" if bold else "C:/Windows/Fonts/segoeui.ttf",
        "C:/Windows/Fonts/arialbd.ttf" if bold else "C:/Windows/Fonts/arial.ttf",
    ]
    for candidate in candidates:
        if Path(candidate).exists():
            return ImageFont.truetype(candidate, size)
    return ImageFont.load_default()


def main():
    img = Image.new("RGBA", (256, 256), (0, 0, 0, 0))
    draw = ImageDraw.Draw(img)
    draw.rounded_rectangle((16, 16, 240, 240), radius=46, fill="#12324a")
    draw.rounded_rectangle((52, 54, 204, 210), radius=18, fill="#ffffff")
    draw.rectangle((74, 80, 182, 102), fill="#1f77b4")
    draw.rectangle((74, 120, 182, 138), fill="#d9eaf7")
    draw.rectangle((74, 154, 154, 172), fill="#d9eaf7")
    draw.ellipse((152, 142, 220, 210), fill="#1f8a5b")
    draw.line((170, 177, 185, 192, 205, 160), fill="#ffffff", width=10)
    draw.text((83, 174), "AI", font=font(32, True), fill="#12324a")
    img.save(OUT, sizes=[(256, 256), (128, 128), (64, 64), (32, 32), (16, 16)])
    print(OUT)


if __name__ == "__main__":
    main()
