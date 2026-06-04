from pathlib import Path
import subprocess

import imageio_ffmpeg
from PIL import Image, ImageDraw, ImageFont


ROOT = Path(__file__).resolve().parent
OUT_DIR = ROOT / "Hiring Automation Tool" / "video_review_frames"
OUT_DIR.mkdir(parents=True, exist_ok=True)

VIDEOS = [
    ("new_live_run", Path(r"C:\Users\Nitin Maheshwari\Videos\2026-06-01 10-38-00.mp4"), [2, 10, 22, 33]),
    ("tool_video", Path(r"C:\Users\Nitin Maheshwari\Downloads\Hiring_Automation_Tool.mp4"), [5, 45, 90, 160, 250, 360, 440]),
    ("explainer", Path(r"C:\Users\Nitin Maheshwari\Downloads\Hiring_Automation_Explainer.mp4"), [5, 35, 80, 140, 220, 300, 340]),
]


def ffmpeg_extract(video: Path, second: int, output: Path) -> None:
    subprocess.run(
        [
            imageio_ffmpeg.get_ffmpeg_exe(),
            "-y",
            "-ss",
            str(second),
            "-i",
            str(video),
            "-frames:v",
            "1",
            str(output),
        ],
        check=True,
        stdout=subprocess.DEVNULL,
        stderr=subprocess.DEVNULL,
    )


def font(size: int, bold: bool = False):
    path = r"C:\Windows\Fonts\arialbd.ttf" if bold else r"C:\Windows\Fonts\arial.ttf"
    return ImageFont.truetype(path, size) if Path(path).exists() else ImageFont.load_default()


def main() -> None:
    thumbs: list[tuple[str, Path]] = []
    for label, video, seconds in VIDEOS:
        for second in seconds:
            out = OUT_DIR / f"{label}_{second:03d}.png"
            ffmpeg_extract(video, second, out)
            thumbs.append((f"{label} {second}s", out))

    thumb_w, thumb_h = 320, 180
    label_h = 38
    cols = 4
    rows = (len(thumbs) + cols - 1) // cols
    sheet = Image.new("RGB", (cols * thumb_w, rows * (thumb_h + label_h)), "white")
    draw = ImageDraw.Draw(sheet)
    for idx, (label, path) in enumerate(thumbs):
        img = Image.open(path).convert("RGB")
        img.thumbnail((thumb_w, thumb_h), Image.Resampling.LANCZOS)
        x = (idx % cols) * thumb_w
        y = (idx // cols) * (thumb_h + label_h)
        sheet.paste(img, (x + (thumb_w - img.width) // 2, y))
        draw.text((x + 8, y + thumb_h + 8), label, fill=(16, 58, 81), font=font(18, True))
    sheet_path = OUT_DIR / "contact_sheet.png"
    sheet.save(sheet_path)
    print(sheet_path)


if __name__ == "__main__":
    main()
