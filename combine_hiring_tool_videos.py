from pathlib import Path
import subprocess

import imageio_ffmpeg
from PIL import Image, ImageDraw, ImageFont


ROOT = Path(__file__).resolve().parent
OUT_DIR = ROOT / "Hiring Automation Tool"
ASSET_DIR = OUT_DIR / "combined_video_assets"
TEMP_DIR = OUT_DIR / "_combined_video_temp"

VIDEO_LIVE = Path(r"C:\Users\Nitin Maheshwari\Videos\2026-06-01 10-38-00.mp4")
VIDEO_TOOL = Path(r"C:\Users\Nitin Maheshwari\Downloads\Hiring_Automation_Tool.mp4")
VIDEO_EXPLAINER = Path(r"C:\Users\Nitin Maheshwari\Downloads\Hiring_Automation_Explainer.mp4")
LOGO = Path(r"C:\Users\Nitin Maheshwari\Downloads\Logo\Logo\motifzone_020919_OP-01.jpg")

OUTPUT = OUT_DIR / "Hiring Automation Tool - Combined Client Video.mp4"
WIDTH = 1280
HEIGHT = 720
FPS = 24


def ffmpeg() -> str:
    return imageio_ffmpeg.get_ffmpeg_exe()


def run_ffmpeg(args: list[str]) -> None:
    subprocess.run([ffmpeg(), *args], check=True)


def font(size: int, bold: bool = False):
    path = r"C:\Windows\Fonts\arialbd.ttf" if bold else r"C:\Windows\Fonts\arial.ttf"
    return ImageFont.truetype(path, size) if Path(path).exists() else ImageFont.load_default()


def logo_image() -> Image.Image:
    image = Image.open(LOGO).convert("RGB")
    gray = image.convert("L")
    mask = gray.point(lambda p: 0 if p > 246 else 255)
    bbox = mask.getbbox()
    if bbox:
        pad = 55
        image = image.crop(
            (
                max(bbox[0] - pad, 0),
                max(bbox[1] - pad, 0),
                min(bbox[2] + pad, image.width),
                min(bbox[3] + pad, image.height),
            )
        )
    image.thumbnail((330, 135), Image.Resampling.LANCZOS)
    return image


def make_title_card() -> Path:
    ASSET_DIR.mkdir(parents=True, exist_ok=True)
    path = ASSET_DIR / "title_card.png"
    image = Image.new("RGB", (WIDTH, HEIGHT), "#103a51")
    draw = ImageDraw.Draw(image)

    logo = logo_image()
    card = Image.new("RGB", (380, 165), "white")
    card.paste(logo, ((card.width - logo.width) // 2, (card.height - logo.height) // 2))
    image.paste(card, (72, 70))

    draw.text((74, 285), "Hiring Automation Tool", fill="white", font=font(58, True))
    draw.text(
        (78, 368),
        "Resume email processing, CV download, and Excel candidate tracking",
        fill="#d8f2ef",
        font=font(27),
    )
    draw.rounded_rectangle((78, 470, 318, 522), radius=12, fill="#ef3125")
    draw.text((105, 484), "Client Demo", fill="white", font=font(23, True))
    draw.rounded_rectangle((342, 470, 660, 522), radius=12, fill="#272f78")
    draw.text((370, 484), "Motifzone Pvt Ltd", fill="white", font=font(23, True))
    image.save(path)
    return path


def make_section_card(filename: str, title: str, subtitle: str) -> Path:
    path = ASSET_DIR / filename
    image = Image.new("RGB", (WIDTH, HEIGHT), "#f6f8fb")
    draw = ImageDraw.Draw(image)
    draw.rectangle((0, 0, WIDTH, 96), fill="#103a51")
    small_logo = logo_image()
    small_logo.thumbnail((170, 70), Image.Resampling.LANCZOS)
    logo_card = Image.new("RGB", (210, 78), "white")
    logo_card.paste(small_logo, ((210 - small_logo.width) // 2, (78 - small_logo.height) // 2))
    image.paste(logo_card, (1000, 10))
    draw.text((70, 32), "Hiring Automation Tool", fill="white", font=font(30, True))
    draw.text((90, 240), title, fill="#103a51", font=font(56, True))
    draw.text((94, 325), subtitle, fill="#355266", font=font(28))
    image.save(path)
    return path


def render_card(path: Path, duration: int, output: Path) -> None:
    run_ffmpeg(
        [
            "-y",
            "-loop",
            "1",
            "-t",
            str(duration),
            "-i",
            str(path),
            "-f",
            "lavfi",
            "-t",
            str(duration),
            "-i",
            "anullsrc=channel_layout=stereo:sample_rate=44100",
            "-vf",
            f"scale={WIDTH}:{HEIGHT},fps={FPS},format=yuv420p",
            "-shortest",
            "-c:v",
            "libx264",
            "-preset",
            "veryfast",
            "-crf",
            "22",
            "-c:a",
            "aac",
            "-ar",
            "44100",
            "-ac",
            "2",
            str(output),
        ]
    )


def render_segment(source: Path, start: float, duration: float, output: Path) -> None:
    run_ffmpeg(
        [
            "-y",
            "-ss",
            str(start),
            "-t",
            str(duration),
            "-i",
            str(source),
            "-vf",
            f"scale={WIDTH}:{HEIGHT}:force_original_aspect_ratio=decrease,pad={WIDTH}:{HEIGHT}:(ow-iw)/2:(oh-ih)/2:color=white,fps={FPS},format=yuv420p",
            "-af",
            "aresample=44100",
            "-c:v",
            "libx264",
            "-preset",
            "veryfast",
            "-crf",
            "22",
            "-c:a",
            "aac",
            "-ar",
            "44100",
            "-ac",
            "2",
            str(output),
        ]
    )


def concat(segments: list[Path]) -> None:
    list_file = TEMP_DIR / "segments.txt"
    list_file.write_text("\n".join(f"file '{p.as_posix()}'" for p in segments), encoding="utf-8")
    run_ffmpeg(["-y", "-f", "concat", "-safe", "0", "-i", str(list_file), "-c", "copy", str(OUTPUT)])


def main() -> None:
    OUT_DIR.mkdir(exist_ok=True)
    TEMP_DIR.mkdir(exist_ok=True)

    title = make_title_card()
    problem = make_section_card("section_problem.png", "Why Automation?", "Reduce manual resume handling and repeated Excel entry.")
    workflow = make_section_card("section_workflow.png", "How It Works", "Scan mailbox, filter CVs, extract data, and update Excel.")
    live = make_section_card("section_live.png", "Live Tool Run", "The actual desktop tool processing the hiring mailbox.")

    segments = [
        TEMP_DIR / "01_title.mp4",
        TEMP_DIR / "02_problem_card.mp4",
        TEMP_DIR / "03_problem.mp4",
        TEMP_DIR / "04_workflow_card.mp4",
        TEMP_DIR / "05_process.mp4",
        TEMP_DIR / "06_modes.mp4",
        TEMP_DIR / "07_live_card.mp4",
        TEMP_DIR / "08_live_run.mp4",
        TEMP_DIR / "09_clean_tracking.mp4",
        TEMP_DIR / "10_closing.mp4",
    ]

    render_card(title, 4, segments[0])
    render_card(problem, 2, segments[1])
    # From the main tool video: manual effort / business requirement. Avoids the explainer title.
    render_segment(VIDEO_TOOL, 38, 70, segments[2])
    render_card(workflow, 2, segments[3])
    # From the main tool video: process modes and automation flow.
    render_segment(VIDEO_TOOL, 145, 105, segments[4])
    render_segment(VIDEO_TOOL, 430, 32, segments[5])
    render_card(live, 2, segments[6])
    # Actual latest run, kept once so the same UI content is not repeated.
    render_segment(VIDEO_LIVE, 0, 35.5, segments[7])
    # From the explainer: benefits/clean tracking section only, skipping the "Automation Explainer" title.
    render_segment(VIDEO_EXPLAINER, 28, 48, segments[8])
    render_segment(VIDEO_EXPLAINER, 300, 42, segments[9])

    concat(segments)
    print(OUTPUT)


if __name__ == "__main__":
    main()
