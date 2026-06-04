from pathlib import Path

from PIL import Image, ImageDraw, ImageFont


ROOT = Path(__file__).resolve().parent
PREVIEW_DIR = ROOT / "Hiring Automation Tool" / "combined_video_assets" / "final_preview"
OUTPUT = PREVIEW_DIR / "preview_sheet.png"


def font(size: int, bold: bool = False):
    path = r"C:\Windows\Fonts\arialbd.ttf" if bold else r"C:\Windows\Fonts\arial.ttf"
    return ImageFont.truetype(path, size) if Path(path).exists() else ImageFont.load_default()


def main() -> None:
    files = sorted(PREVIEW_DIR.glob("preview_*.png"), key=lambda p: int(p.stem.split("_")[1]))
    thumb_w, thumb_h = 360, 202
    label_h = 38
    cols = 3
    rows = (len(files) + cols - 1) // cols
    sheet = Image.new("RGB", (cols * thumb_w, rows * (thumb_h + label_h)), "white")
    draw = ImageDraw.Draw(sheet)
    for idx, path in enumerate(files):
        second = path.stem.split("_")[1]
        img = Image.open(path).convert("RGB")
        img.thumbnail((thumb_w, thumb_h), Image.Resampling.LANCZOS)
        x = (idx % cols) * thumb_w
        y = (idx // cols) * (thumb_h + label_h)
        sheet.paste(img, (x + (thumb_w - img.width) // 2, y))
        draw.text((x + 8, y + thumb_h + 8), f"{second}s", fill=(16, 58, 81), font=font(20, True))
    sheet.save(OUTPUT)
    print(OUTPUT)


if __name__ == "__main__":
    main()
