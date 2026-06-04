import html
import os
import zipfile
from pathlib import Path
from xml.sax.saxutils import escape

from PIL import Image, ImageDraw, ImageFont


ROOT = Path(__file__).resolve().parent
OUT = ROOT / "Hiring Automation Tool"
ASSETS = OUT / "presentation_assets"
PPTX = OUT / "Hiring Automation Tool - Client Presentation.pptx"
HTML = OUT / "Hiring Automation Tool - Video Presentation.html"
SCRIPT = OUT / "Hiring Automation Tool - Voiceover Script.txt"

W, H = 1600, 900
BG = "#f5f7fb"
NAVY = "#12324a"
BLUE = "#1f77b4"
GREEN = "#1f8a5b"
ORANGE = "#f59e0b"
RED = "#c2410c"
DARK = "#203040"
MUTED = "#5f7080"


def font(size, bold=False):
    candidates = [
        "C:/Windows/Fonts/segoeuib.ttf" if bold else "C:/Windows/Fonts/segoeui.ttf",
        "C:/Windows/Fonts/arialbd.ttf" if bold else "C:/Windows/Fonts/arial.ttf",
    ]
    for candidate in candidates:
        if Path(candidate).exists():
            return ImageFont.truetype(candidate, size)
    return ImageFont.load_default()


F_TITLE = font(58, True)
F_SUB = font(30)
F_HEAD = font(36, True)
F_BODY = font(27)
F_SMALL = font(21)
F_LABEL = font(22, True)


slides = [
    {
        "title": "Hiring Automation Tool",
        "subtitle": "AI-enabled resume processing and candidate tracking for Motifzone Pvt Ltd",
        "bullets": [
            "Automatically reads resume emails from the hiring mailbox",
            "Downloads PDF/DOCX CV attachments",
            "Extracts candidate name, email, contact number, experience and skills",
            "Maintains a formatted Excel tracker with duplicate checks",
        ],
        "voice": "Introducing the Hiring Automation Tool, developed for Motifzone Private Limited. This tool automates resume collection from email, extracts candidate information, and updates a professional Excel tracker.",
    },
    {
        "title": "Business Requirement",
        "subtitle": "Why this tool was needed",
        "bullets": [
            "Recruitment mailbox receives many resumes and job descriptions daily",
            "Manual downloading and data entry takes time",
            "Duplicate candidate records create confusion",
            "Recruiters need a clean, latest candidate tracker",
        ],
        "voice": "The requirement came from a practical recruitment challenge. Resume emails arrive continuously, and manually downloading attachments, reading each CV, and entering details into Excel consumes valuable recruiter time.",
    },
    {
        "title": "Automated Process Flow",
        "subtitle": "From mailbox to structured Excel",
        "flow": ["Careers Mailbox", "CV Download", "AI Parsing", "Duplicate Check", "Excel Tracker"],
        "voice": "The process starts from the careers mailbox. The tool reads new emails, downloads CV attachments, extracts useful data, checks duplicates, and updates the Excel tracker automatically.",
    },
    {
        "title": "Professional User Interface",
        "subtitle": "Simple buttons for daily use",
        "ui": True,
        "voice": "The client-facing interface is simple. The user enters the email password or app password, then clicks Run New Emails for daily incremental processing, Full Scan for historical processing, or Start From Now to create a fresh checkpoint.",
    },
    {
        "title": "Live Demo Flow",
        "subtitle": "What the client sees while the automation runs",
        "demo": True,
        "voice": "During the live demo, the user opens the Hiring Automation Tool, clicks Run New Emails, watches useful progress messages such as skipped job descriptions and newly added candidates, and then opens the updated Excel tracker.",
    },
    {
        "title": "Automatic CV Download",
        "subtitle": "No manual attachment saving required",
        "bullets": [
            "Connects securely using IMAP SSL",
            "Reads the configured recruitment mailbox",
            "Downloads supported resume attachments",
            "Supports PDF and DOCX resumes",
            "Keeps JD and non-CV files separate through skip detection",
        ],
        "voice": "One of the key features is automatic CV download. Recruiters do not need to manually save attachments. The tool connects securely to the mailbox and downloads supported resumes.",
    },
    {
        "title": "AI Resume Data Extraction",
        "subtitle": "Candidate details are identified from CV content",
        "bullets": [
            "Candidate Name",
            "Email ID",
            "Contact Number",
            "Total Experience",
            "Current Company where available",
            "Relevant Skills",
        ],
        "voice": "The tool extracts important candidate information such as name, email, contact number, total experience, current company when available, and relevant skills.",
    },
    {
        "title": "Excel Candidate Tracker",
        "subtitle": "Clean, formatted and sorted output",
        "excel": True,
        "voice": "The final output is a structured Excel tracker. The headings are formatted, rows are sorted by latest received email, and duplicate candidates are controlled using name, email, and phone number.",
    },
    {
        "title": "Incremental Email Check",
        "subtitle": "Daily runs process only new emails",
        "bullets": [
            "Run New Emails checks only emails after the last checkpoint",
            "Start From Now creates a fresh checkpoint",
            "Full Scan is available only when old emails must be rescanned",
            "This keeps daily processing fast and efficient",
        ],
        "voice": "For daily usage, the tool performs incremental checking. After the checkpoint is created, Run New Emails checks only new emails received after the previous run.",
    },
    {
        "title": "Error And Skip Visibility",
        "subtitle": "Useful messages without technical noise",
        "bullets": [
            "Shows skipped JD or non-CV attachments",
            "Shows incomplete CV skip messages",
            "Hides low-level PDF font warnings",
            "Keeps the process client-friendly",
        ],
        "voice": "The tool shows useful messages such as skipped job descriptions and incomplete CVs, while hiding confusing low-level technical warnings. This keeps the experience professional for clients.",
    },
    {
        "title": "Benefits For Motifzone Pvt Ltd",
        "subtitle": "Achievement and business value",
        "bullets": [
            "Reduces manual recruitment operations effort",
            "Improves candidate data quality",
            "Speeds up shortlisting and tracking",
            "Creates a repeatable automation foundation",
            "Demonstrates practical AI adoption by Motifzone Pvt Ltd",
        ],
        "voice": "This is a practical automation achievement for Motifzone Private Limited. It reduces manual effort, improves data quality, speeds up recruitment tracking, and demonstrates real AI adoption in business operations.",
    },
]


def wrap(draw, text, fnt, width):
    words = text.split()
    lines = []
    line = ""
    for word in words:
        test = f"{line} {word}".strip()
        if draw.textbbox((0, 0), test, font=fnt)[2] <= width:
            line = test
        else:
            if line:
                lines.append(line)
            line = word
    if line:
        lines.append(line)
    return lines


def card(draw, xy, fill="#ffffff", outline="#dbe3ea"):
    draw.rounded_rectangle(xy, radius=14, fill=fill, outline=outline, width=2)


def draw_header(draw, title, subtitle):
    draw.rectangle((0, 0, W, 112), fill=NAVY)
    draw.text((70, 24), title, font=F_TITLE if len(title) < 28 else F_HEAD, fill="white")
    draw.text((74, 82), subtitle, font=F_SMALL, fill="#cfe4f3")


def draw_icon(draw, x, y, kind, color=BLUE):
    draw.ellipse((x, y, x + 72, y + 72), fill=color)
    if kind == "mail":
        draw.rectangle((x + 18, y + 23, x + 54, y + 49), outline="white", width=4)
        draw.line((x + 18, y + 23, x + 36, y + 39, x + 54, y + 23), fill="white", width=4)
    elif kind == "cv":
        draw.rectangle((x + 22, y + 14, x + 52, y + 58), fill="white")
        draw.line((x + 28, y + 28, x + 46, y + 28), fill=color, width=3)
        draw.line((x + 28, y + 39, x + 46, y + 39), fill=color, width=3)
    elif kind == "ai":
        draw.rounded_rectangle((x + 18, y + 18, x + 54, y + 54), radius=8, outline="white", width=4)
        draw.ellipse((x + 28, y + 29, x + 34, y + 35), fill="white")
        draw.ellipse((x + 40, y + 29, x + 46, y + 35), fill="white")
    elif kind == "excel":
        draw.rectangle((x + 18, y + 18, x + 56, y + 56), fill="white")
        draw.line((x + 18, y + 31, x + 56, y + 31), fill=color, width=3)
        draw.line((x + 31, y + 18, x + 31, y + 56), fill=color, width=3)
    elif kind == "shield":
        draw.polygon([(x + 36, y + 12), (x + 56, y + 22), (x + 50, y + 55), (x + 36, y + 62), (x + 22, y + 55), (x + 16, y + 22)], fill="white")
        draw.line((x + 27, y + 38, x + 35, y + 47, x + 49, y + 27), fill=color, width=5)


def draw_visual_panel(draw):
    card(draw, (1020, 190, 1485, 720), fill="#ffffff")
    draw.text((1060, 220), "Automation Snapshot", font=F_LABEL, fill=NAVY)
    steps = [
        ("mail", "Reads new mailbox items", BLUE),
        ("cv", "Downloads resume files", GREEN),
        ("ai", "Extracts candidate details", ORANGE),
        ("excel", "Updates Excel tracker", BLUE),
        ("shield", "Masks sensitive demo data", GREEN),
    ]
    y = 275
    for kind, label, color in steps:
        draw_icon(draw, 1060, y, kind, color)
        for i, line in enumerate(wrap(draw, label, F_SMALL, 290)):
            draw.text((1150, y + 14 + i * 25), line, font=F_SMALL, fill=DARK)
        y += 82


def draw_bullets(draw, bullets):
    y = 190
    for bullet in bullets:
        draw.ellipse((92, y + 10, 112, y + 30), fill=BLUE)
        for i, line in enumerate(wrap(draw, bullet, F_BODY, 800)):
            draw.text((135, y + i * 34), line, font=F_BODY, fill=DARK)
        y += 82
    draw_visual_panel(draw)


def draw_flow(draw, items):
    x = 95
    y = 310
    box_w = 250
    gap = 60
    icons = ["mail", "cv", "ai", "shield", "excel"]
    colors = [BLUE, GREEN, ORANGE, NAVY, GREEN]
    for i, item in enumerate(items):
        card(draw, (x, y, x + box_w, y + 180), fill="#ffffff")
        draw_icon(draw, x + 88, y + 24, icons[i], colors[i])
        draw.text((x + 28, y + 112), f"{i + 1}.", font=F_LABEL, fill=ORANGE)
        for line_no, line in enumerate(wrap(draw, item, F_LABEL, 165)):
            draw.text((x + 62, y + 112 + line_no * 27), line, font=F_LABEL, fill=NAVY)
        if i < len(items) - 1:
            ax = x + box_w + 12
            draw.line((ax, y + 90, ax + gap - 22, y + 90), fill=ORANGE, width=7)
            draw.polygon([(ax + gap - 22, y + 90), (ax + gap - 42, y + 78), (ax + gap - 42, y + 102)], fill=ORANGE)
        x += box_w + gap


def draw_ui(draw):
    card(draw, (145, 185, 1455, 780), fill="#ffffff")
    draw.rectangle((145, 185, 1455, 290), fill=NAVY)
    draw.text((190, 215), "Hiring Automation Tool", font=F_HEAD, fill="white")
    draw.text((193, 260), "Automatically reads resume emails and updates the candidate Excel tracker", font=F_SMALL, fill="#cfe4f3")
    draw.text((210, 335), "Email password / app password", font=F_LABEL, fill=DARK)
    draw.rounded_rectangle((210, 375, 760, 430), radius=8, outline="#b9c6d3", width=2, fill="#f8fbfd")
    draw.text((230, 388), "••••••••••••", font=F_BODY, fill=MUTED)
    buttons = [("Run New Emails", GREEN), ("Full Scan", BLUE), ("Start From Now", ORANGE)]
    x = 210
    for label, color in buttons:
        draw.rounded_rectangle((x, 490, x + 250, 558), radius=10, fill=color)
        draw.text((x + 30, 510), label, font=F_LABEL, fill="white")
        x += 285
    draw.rounded_rectangle((210, 610, 1260, 700), radius=8, fill="#f6f8fb", outline="#dce5ee")
    draw.text((235, 632), "Skipped non-CV attachment: Job Description RPA Lead.docx", font=F_SMALL, fill=MUTED)
    draw.text((235, 662), "New candidates added: 5    Excel file updated successfully", font=F_SMALL, fill=GREEN)


def draw_demo_flow(draw):
    card(draw, (95, 180, 1505, 765), fill="#ffffff")
    panels = [
        ("1", "Open Tool", "Client opens the Hiring Automation Tool from a desktop icon."),
        ("2", "Run New Emails", "Tool checks only new mailbox items after the checkpoint."),
        ("3", "Visible Progress", "JD skips and incomplete CV messages are shown clearly."),
        ("4", "Excel Updated", "Candidate tracker is updated with masked demo data."),
    ]
    x = 135
    for number, title, text in panels:
        draw.rounded_rectangle((x, 245, x + 310, 610), radius=18, fill="#f8fbfd", outline="#dbe3ea", width=2)
        draw.ellipse((x + 25, 270, x + 85, 330), fill=ORANGE)
        draw.text((x + 45, 284), number, font=F_HEAD, fill="white")
        draw.text((x + 30, 360), title, font=F_LABEL, fill=NAVY)
        yy = 405
        for line in wrap(draw, text, F_SMALL, 250):
            draw.text((x + 30, yy), line, font=F_SMALL, fill=DARK)
            yy += 28
        x += 345
    draw.text((135, 665), "Actual running screen can be recorded and placed after this slide for the final client video.", font=F_BODY, fill=GREEN)


def draw_excel(draw):
    card(draw, (110, 190, 1490, 760), fill="#ffffff")
    headers = ["Candidate Name", "Email", "Contact", "Experience", "Skills", "Received"]
    xs = [130, 380, 690, 900, 1100, 1320]
    widths = [230, 290, 190, 170, 200, 140]
    y = 220
    for x, w, header in zip(xs, widths, headers):
        draw.rectangle((x, y, x + w, y + 48), fill="#d9eaf7", outline="#b7cad9")
        draw.text((x + 10, y + 13), header, font=F_SMALL, fill=NAVY)
    rows = [
        ["Dalia Ammar", "d***@gmail.com", "+20 *** ***", "15 years", "Oracle CX", "Latest"],
        ["Naveed Sarwar", "r***@hotmail.com", "+92 *** ***", "SAP", "ABAP, Fiori", "Latest"],
        ["Kranthi", "k***@gmail.com", "98******31", "Data", "Azure, DBT", "Latest"],
        ["Madhavi Latha", "m***@gmail.com", "97******89", "Finance", "Oracle Fusion", "Older"],
    ]
    y = 268
    for row in rows:
        for x, w, value in zip(xs, widths, row):
            draw.rectangle((x, y, x + w, y + 58), fill="#ffffff", outline="#dce5ee")
            draw.text((x + 10, y + 17), value, font=F_SMALL, fill=DARK)
        y += 58
    draw.text((130, 560), "Duplicate checks: Name + Email + Contact Number", font=F_BODY, fill=GREEN)
    draw.text((130, 620), "Sorted by newest received email first", font=F_BODY, fill=BLUE)


def render_slide(index, slide):
    img = Image.new("RGB", (W, H), BG)
    draw = ImageDraw.Draw(img)
    draw_header(draw, slide["title"], slide["subtitle"])
    if slide.get("flow"):
        draw_flow(draw, slide["flow"])
    elif slide.get("ui"):
        draw_ui(draw)
    elif slide.get("demo"):
        draw_demo_flow(draw)
    elif slide.get("excel"):
        draw_excel(draw)
    else:
        draw_bullets(draw, slide["bullets"])
    draw.text((70, 842), f"Motifzone Pvt Ltd | Hiring Automation Tool | {index + 1}/{len(slides)}", font=F_SMALL, fill=MUTED)
    path = ASSETS / f"slide_{index + 1:02d}.png"
    img.save(path)
    return path


def pptx_package(slide_paths):
    def rels_for_slide(i):
        return f'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/slide_{i:02d}.png"/>
</Relationships>'''

    content_types = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
<Default Extension="xml" ContentType="application/xml"/>
<Default Extension="png" ContentType="image/png"/>
<Override PartName="/ppt/presentation.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"/>
''' + "\n".join(
        f'<Override PartName="/ppt/slides/slide{i}.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>'
        for i in range(1, len(slide_paths) + 1)
    ) + "\n</Types>"

    root_rels = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="ppt/presentation.xml"/>
</Relationships>'''

    pres_rels = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
''' + "\n".join(
        f'<Relationship Id="rId{i}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide{i}.xml"/>'
        for i in range(1, len(slide_paths) + 1)
    ) + "\n</Relationships>"

    sld_ids = "\n".join(
        f'<p:sldId id="{255 + i}" r:id="rId{i}"/>' for i in range(1, len(slide_paths) + 1)
    )
    presentation = f'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:presentation xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
<p:sldSz cx="12192000" cy="6858000" type="wide"/>
<p:sldIdLst>{sld_ids}</p:sldIdLst>
</p:presentation>'''

    slide_xml = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sld xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
<p:cSld><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/>
<p:pic><p:nvPicPr><p:cNvPr id="2" name="Slide Image"/><p:cNvPicPr/><p:nvPr/></p:nvPicPr>
<p:blipFill><a:blip r:embed="rId1"/><a:stretch><a:fillRect/></a:stretch></p:blipFill>
<p:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="12192000" cy="6858000"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr>
</p:pic></p:spTree></p:cSld><p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:sld>'''

    with zipfile.ZipFile(PPTX, "w", zipfile.ZIP_DEFLATED) as z:
        z.writestr("[Content_Types].xml", content_types)
        z.writestr("_rels/.rels", root_rels)
        z.writestr("ppt/presentation.xml", presentation)
        z.writestr("ppt/_rels/presentation.xml.rels", pres_rels)
        for i, path in enumerate(slide_paths, start=1):
            z.writestr(f"ppt/slides/slide{i}.xml", slide_xml)
            z.writestr(f"ppt/slides/_rels/slide{i}.xml.rels", rels_for_slide(i))
            z.write(path, f"ppt/media/slide_{i:02d}.png")


def html_package(slide_paths):
    slide_data = [
        {
            "img": f"presentation_assets/{path.name}",
            "title": slide["title"],
            "voice": slide["voice"],
        }
        for path, slide in zip(slide_paths, slides)
    ]
    items = ",\n".join(
        "{img:%r,title:%r,voice:%r}" % (item["img"], item["title"], item["voice"]) for item in slide_data
    )
    doc = f'''<!doctype html>
<html>
<head>
<meta charset="utf-8">
<title>Hiring Automation Tool - Video Presentation</title>
<style>
body{{margin:0;background:#0f1720;color:#fff;font-family:Segoe UI,Arial,sans-serif;}}
.stage{{height:100vh;display:flex;flex-direction:column;align-items:center;justify-content:center;gap:14px;}}
img{{max-width:94vw;max-height:82vh;border-radius:10px;box-shadow:0 20px 70px #0008;}}
.bar{{width:94vw;display:flex;justify-content:space-between;align-items:center;}}
button{{background:#1f77b4;color:white;border:0;border-radius:8px;padding:12px 18px;font-size:16px;cursor:pointer;}}
button.alt{{background:#1f8a5b;}}
#caption{{color:#cfe4f3;font-size:18px;}}
</style>
</head>
<body>
<div class="stage">
  <img id="slide" src="">
  <div class="bar">
    <div id="caption"></div>
    <div>
      <button onclick="prevSlide()">Previous</button>
      <button class="alt" onclick="playDeck()">Play With Voice</button>
      <button onclick="nextSlide()">Next</button>
    </div>
  </div>
</div>
<script>
const slides=[{items}];
let index=0;
function show(){{document.getElementById('slide').src=slides[index].img;document.getElementById('caption').textContent=(index+1)+' / '+slides.length+' - '+slides[index].title;}}
function speak(text, done){{speechSynthesis.cancel(); const u=new SpeechSynthesisUtterance(text); u.rate=0.92; u.onend=done; speechSynthesis.speak(u);}}
function playDeck(){{show(); speak(slides[index].voice,()=>{{ if(index<slides.length-1){{index++; setTimeout(playDeck,700);}} }});}}
function nextSlide(){{index=Math.min(slides.length-1,index+1);show();}}
function prevSlide(){{index=Math.max(0,index-1);show();}}
show();
</script>
</body>
</html>'''
    HTML.write_text(doc, encoding="utf-8")


def write_script():
    lines = []
    for i, slide in enumerate(slides, start=1):
        lines.append(f"Slide {i}: {slide['title']}")
        lines.append(slide["voice"])
        lines.append("")
    SCRIPT.write_text("\n".join(lines), encoding="utf-8")


def main():
    ASSETS.mkdir(parents=True, exist_ok=True)
    slide_paths = [render_slide(i, slide) for i, slide in enumerate(slides)]
    pptx_package(slide_paths)
    html_package(slide_paths)
    write_script()
    print(PPTX)
    print(HTML)
    print(SCRIPT)


if __name__ == "__main__":
    main()
