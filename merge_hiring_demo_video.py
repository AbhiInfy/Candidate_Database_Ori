import subprocess
from pathlib import Path

import imageio_ffmpeg
from PIL import Image, ImageDraw, ImageFont


PROJECT_DIR = Path(__file__).resolve().parent
DELIVERABLE_DIR = PROJECT_DIR / "Hiring Automation Tool"
ASSET_DIR = DELIVERABLE_DIR / "presentation_assets"

EARLIER_VIDEO = Path(
    r"C:\Users\Nitin Maheshwari\Documents\Codex\2026-05-21\files-mentioned-by-the-user-2026\2026-05-21 10-58-07_highlighted-window.mp4"
)
LATEST_VIDEO = Path(r"C:\Users\Nitin Maheshwari\Videos\2026-05-21 12-10-09.mp4")
OUTPUT_VIDEO = DELIVERABLE_DIR / "Hiring Automation Tool - Merged Demo Video.mp4"
TEMP_DIR = DELIVERABLE_DIR / "_video_merge_temp"

WIDTH = 1280
HEIGHT = 720
FPS = 30


def font(size: int, bold: bool = False) -> ImageFont.FreeTypeFont:
    candidates = [
        r"C:\Windows\Fonts\arialbd.ttf" if bold else r"C:\Windows\Fonts\arial.ttf",
        r"C:\Windows\Fonts\calibrib.ttf" if bold else r"C:\Windows\Fonts\calibri.ttf",
    ]
    for candidate in candidates:
        if Path(candidate).exists():
            return ImageFont.truetype(candidate, size)
    return ImageFont.load_default()


def wrapped_lines(draw: ImageDraw.ImageDraw, text: str, text_font: ImageFont.FreeTypeFont, max_width: int) -> list[str]:
    lines: list[str] = []
    for paragraph in text.split("\n"):
        words = paragraph.split()
        current = ""
        for word in words:
            trial = f"{current} {word}".strip()
            if draw.textbbox((0, 0), trial, font=text_font)[2] <= max_width:
                current = trial
            else:
                if current:
                    lines.append(current)
                current = word
        if current:
            lines.append(current)
    return lines


def make_card(filename: str, title: str, subtitle: str, bullets: list[str] | None = None) -> Path:
    DELIVERABLE_DIR.mkdir(exist_ok=True)
    path = DELIVERABLE_DIR / filename

    image = Image.new("RGB", (WIDTH, HEIGHT), "#f6f8fb")
    draw = ImageDraw.Draw(image)

    draw.rectangle((0, 0, WIDTH, 120), fill="#103a51")
    draw.text((70, 34), "Motifzone Private Limited", fill="white", font=font(32, True))
    draw.rounded_rectangle((70, 164, WIDTH - 70, HEIGHT - 70), radius=16, fill="white", outline="#d7e0e8", width=2)

    title_font = font(54, True)
    subtitle_font = font(26)
    body_font = font(30)

    draw.text((110, 210), title, fill="#103a51", font=title_font)
    y = 300
    for line in wrapped_lines(draw, subtitle, subtitle_font, WIDTH - 220):
        draw.text((112, y), line, fill="#355266", font=subtitle_font)
        y += 38

    if bullets:
        y += 24
        for bullet in bullets:
            draw.ellipse((116, y + 12, 130, y + 26), fill="#24a19c")
            for line in wrapped_lines(draw, bullet, body_font, WIDTH - 270):
                draw.text((150, y), line, fill="#142b3c", font=body_font)
                y += 42
            y += 16

    image.save(path)
    return path


def run_ffmpeg(args: list[str]) -> None:
    ffmpeg = imageio_ffmpeg.get_ffmpeg_exe()
    subprocess.run([ffmpeg, *args], check=True)


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
            f"scale={WIDTH}:{HEIGHT}:force_original_aspect_ratio=decrease,pad={WIDTH}:{HEIGHT}:(ow-iw)/2:(oh-ih)/2:color=F6F8FB,fps={FPS},format=yuv420p",
            "-shortest",
            "-c:v",
            "libx264",
            "-preset",
            "ultrafast",
            "-crf",
            "23",
            "-c:a",
            "aac",
            "-ar",
            "44100",
            "-ac",
            "2",
            str(output),
        ]
    )


def render_video(path: Path, output: Path) -> None:
    run_ffmpeg(
        [
            "-y",
            "-i",
            str(path),
            "-vf",
            f"scale={WIDTH}:{HEIGHT}:force_original_aspect_ratio=decrease,pad={WIDTH}:{HEIGHT}:(ow-iw)/2:(oh-ih)/2:color=F6F8FB,fps={FPS},format=yuv420p",
            "-c:v",
            "libx264",
            "-preset",
            "ultrafast",
            "-crf",
            "23",
            "-c:a",
            "aac",
            "-ar",
            "44100",
            "-ac",
            "2",
            str(output),
        ]
    )


def concat_segments(segments: list[Path]) -> None:
    list_file = TEMP_DIR / "segments.txt"
    lines = [f"file '{segment.as_posix()}'" for segment in segments]
    list_file.write_text("\n".join(lines), encoding="utf-8")
    run_ffmpeg(
        [
            "-y",
            "-f",
            "concat",
            "-safe",
            "0",
            "-i",
            str(list_file),
            "-c",
            "copy",
            str(OUTPUT_VIDEO),
        ]
    )


def main() -> None:
    if not EARLIER_VIDEO.exists():
        raise FileNotFoundError(f"Earlier video was not found: {EARLIER_VIDEO}")
    if not LATEST_VIDEO.exists():
        raise FileNotFoundError(f"Latest video was not found: {LATEST_VIDEO}")

    title_card = make_card(
        "merged_video_title.png",
        "Hiring Automation Tool",
        "Client demonstration video showing the resume email automation flow.",
        [
            "Reads resume emails from the hiring mailbox",
            "Downloads only candidate CV files",
            "Updates the Excel tracker with candidate details",
        ],
    )
    flow_card = ASSET_DIR / "slide_03.png"
    transition_card = make_card(
        "merged_video_transition.png",
        "Live Run",
        "The next section shows the tool running on the system and processing mailbox attachments.",
        ["Skipped JD files are visible", "Candidate CVs are downloaded automatically"],
    )
    end_card = make_card(
        "merged_video_end.png",
        "Outcome",
        "A clean hiring tracker is ready for recruiter review.",
        [
            "Duplicate candidates are controlled",
            "Downloaded CVs are organized",
            "Excel output is ready for sharing",
        ],
    )

    TEMP_DIR.mkdir(exist_ok=True)
    segments = [
        TEMP_DIR / "01_title.mp4",
        TEMP_DIR / "02_flow.mp4",
        TEMP_DIR / "03_earlier_demo.mp4",
        TEMP_DIR / "04_transition.mp4",
        TEMP_DIR / "05_latest_demo.mp4",
        TEMP_DIR / "06_end.mp4",
    ]

    render_card(title_card, 5, segments[0])
    render_card(flow_card if flow_card.exists() else title_card, 4, segments[1])
    render_video(EARLIER_VIDEO, segments[2])
    render_card(transition_card, 4, segments[3])
    render_video(LATEST_VIDEO, segments[4])
    render_card(end_card, 5, segments[5])
    concat_segments(segments)

    print(OUTPUT_VIDEO)


if __name__ == "__main__":
    main()
