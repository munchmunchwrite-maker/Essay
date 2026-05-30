"""Build the TUM Deutschlandstipendium Word document."""

from docx import Document
from docx.shared import Pt, RGBColor, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH
import os

ESSAY_DIR = os.path.dirname(os.path.abspath(__file__))

DARK_BLUE = RGBColor(0x00, 0x28, 0x6E)   # TUM-ish blue
ACCENT = RGBColor(0x8C, 0x16, 0x1C)       # TUM red
DARK_GRAY = RGBColor(0x33, 0x33, 0x33)
MID_GRAY = RGBColor(0x55, 0x55, 0x55)
RED = RGBColor(0xAA, 0x00, 0x00)
GREEN = RGBColor(0x1A, 0x6B, 0x2A)
FILL = RGBColor(0xC0, 0x55, 0x00)         # orange for [FILL IN] items


def sf(run, name="Calibri", size=11, bold=False, italic=False, color=None):
    run.font.name = name
    run.font.size = Pt(size)
    run.bold = bold
    run.italic = italic
    if color:
        run.font.color.rgb = color


def heading(doc, text, level=1):
    sizes = {0: 18, 1: 14, 2: 12, 3: 11}
    colors = {0: DARK_BLUE, 1: DARK_BLUE, 2: ACCENT, 3: DARK_GRAY}
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(14 if level <= 1 else 8)
    p.paragraph_format.space_after = Pt(4)
    r = p.add_run(text)
    sf(r, size=sizes.get(level, 11), bold=True, color=colors.get(level, DARK_BLUE))
    return p


def body(doc, text, size=10.5, color=None, space_before=0, space_after=5,
         bold=False, italic=False):
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(space_before)
    p.paragraph_format.space_after = Pt(space_after)
    r = p.add_run(text)
    sf(r, size=size, bold=bold, italic=italic, color=color or DARK_GRAY)
    return p


def mixed(doc, parts, space_before=1, space_after=4):
    """parts: list of (text, bold, italic, color)"""
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(space_before)
    p.paragraph_format.space_after = Pt(space_after)
    for text, bold, italic, color in parts:
        r = p.add_run(text)
        sf(r, size=10.5, bold=bold, italic=italic, color=color or DARK_GRAY)
    return p


def bullet(doc, text, indent=0.3, color=None):
    p = doc.add_paragraph(style="List Bullet")
    p.paragraph_format.left_indent = Inches(indent)
    p.paragraph_format.space_before = Pt(1)
    p.paragraph_format.space_after = Pt(2)
    r = p.add_run(text)
    sf(r, size=10.5, color=color or DARK_GRAY)
    return p


def fill_field(doc, label, answer, note=None):
    """A labeled field: label in bold blue, answer in gray, optional note in orange."""
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(6)
    p.paragraph_format.space_after = Pt(2)
    lr = p.add_run(label + ": ")
    sf(lr, size=10.5, bold=True, color=DARK_BLUE)
    if "[FILL IN" in answer:
        ar = p.add_run(answer)
        sf(ar, size=10.5, bold=False, color=FILL, italic=True)
    else:
        ar = p.add_run(answer)
        sf(ar, size=10.5, color=DARK_GRAY)
    if note:
        p2 = doc.add_paragraph()
        p2.paragraph_format.space_before = Pt(0)
        p2.paragraph_format.space_after = Pt(4)
        p2.paragraph_format.left_indent = Inches(0.25)
        nr = p2.add_run(note)
        sf(nr, size=9.5, italic=True, color=MID_GRAY)


def divider(doc):
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(4)
    p.paragraph_format.space_after = Pt(4)
    r = p.add_run("_" * 80)
    sf(r, size=7, color=RGBColor(0xCC, 0xCC, 0xCC))


def read_cover_letter_body(filepath):
    """Extract just the letter body from a markdown cover letter file."""
    with open(filepath, encoding="utf-8") as f:
        lines = f.read().splitlines()
    body_lines = []
    in_body = False
    for line in lines:
        stripped = line.strip()
        if stripped == "---" and not in_body:
            in_body = True
            continue
        if in_body:
            if stripped.startswith("---") and body_lines:
                break
            body_lines.append(line)
    # Group into paragraphs
    paragraphs = []
    current = []
    for line in body_lines:
        if line.strip() == "":
            if current:
                paragraphs.append(" ".join(c.strip() for c in current))
                current = []
        else:
            current.append(line.strip())
    if current:
        paragraphs.append(" ".join(c.strip() for c in current))
    return [p for p in paragraphs if p and not p.startswith("**Word count")]


def build_tum_doc(output_path):
    doc = Document()

    for section in doc.sections:
        section.top_margin = Inches(1)
        section.bottom_margin = Inches(1)
        section.left_margin = Inches(1.15)
        section.right_margin = Inches(1.15)

    # Cover page
    t = doc.add_heading("TUM Deutschlandstipendium — Application Materials", level=0)
    t.alignment = WD_ALIGN_PARAGRAPH.CENTER
    for run in t.runs:
        run.font.name = "Calibri"
        run.font.color.rgb = DARK_BLUE

    s = doc.add_paragraph("MSc Finance and Information Management (FIM) | Technical University of Munich")
    s.alignment = WD_ALIGN_PARAGRAPH.CENTER
    for r in s.runs:
        sf(r, size=11, italic=True, color=MID_GRAY)

    s2 = doc.add_paragraph("Applicant: Lauren  |  Prepared: May 30, 2026")
    s2.alignment = WD_ALIGN_PARAGRAPH.CENTER
    for r in s2.runs:
        sf(r, size=10, italic=True, color=MID_GRAY)

    body(doc, "This document contains: (1) a guided fill-in for every application form field, "
         "with suggested answers based on Lauren's resume, LinkedIn, and prior essay materials; "
         "and (2) two cover letter drafts for the Deutschlandstipendium. "
         "Fields highlighted in orange require information that must be confirmed or provided by Lauren.",
         color=DARK_GRAY, space_before=8, space_after=4)

    # Evaluation criteria reminder
    heading(doc, "Selection Criteria (for reference while reading cover letters)", level=2)
    bullet(doc, "Academic performance: 60% of the evaluation weight")
    bullet(doc, "Social and university engagement: 20%")
    bullet(doc, "Personal circumstances (financial hardship, first-gen, health, etc.): 20%")
    body(doc, "Lauren's 3.912 GPA converts to approximately 1.1 on the German scale (exceptional). "
         "Her engagement and personal circumstances both add meaningful points.",
         size=10, italic=True, color=MID_GRAY)

    doc.add_page_break()

    # ── APPLICATION FORM ──────────────────────────────────────────────────────
    heading(doc, "Part 1: Application Form — Guided Answers", level=1)

    heading(doc, "University Entrance Qualification", level=2)
    fill_field(doc, "Acquired in Germany", "No")
    fill_field(doc, "Overall grade (university entrance qualification)",
               "[FILL IN — your high school GPA, converted with Bavarian Formula or from VPD]",
               "Example: if your U.S. high school GPA was 3.8/4.0, the Bavarian Formula gives: "
               "1 + 3 × ((4.0 − 3.8) / (4.0 − 2.0)) = 1.3. Enter with a dot, e.g. 1.3")
    fill_field(doc, "Date of issue", "[FILL IN — your high school graduation date, DD.MM.YYYY]")
    fill_field(doc, "Document to upload", "U.S. high school diploma or final transcript (PDF). "
               "If you have a uni-assist VPD, upload that instead.")

    heading(doc, "Academic Achievements for Current Degree Program (TUM)", level=2)
    fill_field(doc, "Credits", "Select: 'I do not have a transcript from winter semester 2025/2026'",
               "You have not yet begun study at TUM.")
    fill_field(doc, "Weighted average grade", "Select: 'I do not have a transcript from winter semester 2025/2026'")

    heading(doc, "Completed Degree (University of Utah)", level=2)
    fill_field(doc, "Overall grade (German scale)", "1.1",
               "Bavarian Formula: 1 + 3 × ((4.0 − 3.912) / (4.0 − 2.0)) = 1.132 → enter as 1.1 "
               "(truncate, do not round). If you receive a VPD from uni-assist, use that grade instead.")
    fill_field(doc, "Name of university", "University of Utah, David Eccles School of Business")
    fill_field(doc, "Document to upload", "University of Utah official transcript and degree certificate. "
               "If you have a uni-assist VPD, upload that. VPD deadline: August 9, 2026.")

    heading(doc, "University Commitment (July 2023 – May 2026)", level=2)
    fill_field(doc, "Have you been committed to a university", "Yes")

    body(doc, "Short description (max 500 characters):", bold=True, color=DARK_BLUE, space_after=2)
    desc_box = doc.add_paragraph()
    desc_box.paragraph_format.left_indent = Inches(0.2)
    desc_box.paragraph_format.space_after = Pt(4)
    r = desc_box.add_run(
        '"IS Ambassador and Eccles Scholar Ambassador, U. of Utah (2021-2024): organized events '
        'with Salesforce and Adobe; mentored honors business students. Marketing Co-Lead, Business '
        'Student Government (2021-2024): advocated for students to deans. Campus Life Mentor '
        '(2021-2024): mentored first-year and transfer students. Film Director, ASUU '
        '(2021-2023): managed $30K events budget."'
    )
    sf(r, size=10.5, italic=True, color=DARK_GRAY)
    body(doc, "~385 characters — within the 500 character limit.", size=9.5, italic=True, color=MID_GRAY)

    body(doc, "Confirmation documents to request:", bold=True, color=DARK_BLUE, space_before=4, space_after=2)
    bullet(doc, "ASUU (Associated Students of the University of Utah) — Film Director role")
    bullet(doc, "David Eccles School of Business student affairs — IS Ambassador, BSG Marketing Co-Lead, Eccles Scholar Ambassador")
    bullet(doc, "University of Utah — Campus Life Mentor role")
    body(doc, "Each letter must include: date of issue, type of work, time period, average hours per month. "
         "Combine all into one PDF.", size=9.5, italic=True, color=MID_GRAY)

    heading(doc, "Social Commitment (July 2023 – May 2026)", level=2)
    fill_field(doc, "Have you been socially committed", "Yes")

    body(doc, "Short description (max 500 characters):", bold=True, color=DARK_BLUE, space_after=2)
    desc_box2 = doc.add_paragraph()
    desc_box2.paragraph_format.left_indent = Inches(0.2)
    desc_box2.paragraph_format.space_after = Pt(4)
    r2 = desc_box2.add_run(
        '"Disaster Services Coordinator and International Services Outreach Team, American Red Cross '
        '(Dec 2025-present). Emerging Leaders Board Member, National MS Society. Volunteer, '
        'Umsonstladen Greifswald, Germany (Jan-Jun 2025). Graphic Design Specialist, Bennion Center, '
        'U. of Utah (2019-present)."'
    )
    sf(r2, size=10.5, italic=True, color=DARK_GRAY)
    body(doc, "~330 characters — within the 500 character limit.", size=9.5, italic=True, color=MID_GRAY)

    body(doc, "Confirmation documents to request:", bold=True, color=DARK_BLUE, space_before=4, space_after=2)
    bullet(doc, "American Red Cross — Disaster Services Coordinator / International Services Outreach")
    bullet(doc, "National MS Society — Emerging Leaders Board Member")
    bullet(doc, "Umsonstladen Greifswald, Germany — volunteer (request in German or English)")
    bullet(doc, "Bennion Center, University of Utah — Graphic Design Specialist")

    heading(doc, "Personal Circumstances", level=2)
    fill_field(doc, "Physical or mental illness",
               "Potentially: ADHD (documented in prior scholarship essays as affecting early education). "
               "Include only if you have a medical certificate from a licensed doctor and it has affected your studies.",
               "Your 3.912 GPA already demonstrates exceptional performance. Only include this if you feel "
               "it meaningfully contextualizes any aspects of your academic history.")
    fill_field(doc, "Child care", "Not applicable")
    fill_field(doc, "Care of a close relative", "Not applicable")
    fill_field(doc, "Non-academic family background",
               "[FILL IN — select if your mother does not hold a university degree]",
               "Lauren is described as the first in her family to pursue graduate study abroad. "
               "If her mother does not hold a university degree, this category applies. "
               "Upload: a signed statement from your parent(s) confirming they do not hold a university degree.")
    fill_field(doc, "Migration background", "Not applicable")
    fill_field(doc, "War or refugee experience", "Not applicable")

    heading(doc, "Employment (July 2023 – May 2026)", level=2)
    fill_field(doc, "Have you been employed alongside your studies", "Yes")
    fill_field(doc, "Average hours per week", "[FILL IN — estimate based on your roles, typically 15–20 hrs/week]")
    body(doc, "Employment to document:", bold=True, color=DARK_BLUE, space_before=4, space_after=2)
    bullet(doc, "Maverik, Inc. — IT Analyst: Dec 2022–Aug 2023 (July–Aug 2023 within window)")
    bullet(doc, "Breeze Airways — Tax Intern: Aug 2025–Mar 2026")
    bullet(doc, "American Red Cross — Disaster Services Coordinator: Dec 2025–May 2026")
    bullet(doc, "Congress-Bundestag Fellowship (U.S. Dept. of State): Jul 2024–Jun 2025 — include if stipend-based")
    body(doc, "Upload employment contracts combined into one PDF. Each must include: date of issue, "
         "employee name, contract period, weekly hours, and both signatures.", size=9.5, italic=True, color=MID_GRAY)

    heading(doc, "Special Achievements, Awards, and Prizes", level=2)
    body(doc, "Qualifying achievements (last 36 months):", bold=True, color=DARK_BLUE, space_before=4, space_after=2)
    bullet(doc, "Congress-Bundestag Fellowship: selected as one of 75 U.S. fellows from a competitive national applicant pool")
    bullet(doc, "English Sterling Scholar Runner-Up (if within the 36-month window — verify date)")
    bullet(doc, "Dean's List: may not qualify under TUM's competitive award definition — check with the admissions office")
    body(doc, "Upload official confirmation documents for each achievement you list.",
         size=9.5, italic=True, color=MID_GRAY)

    heading(doc, "Remaining Fields", level=2)
    fill_field(doc, "BAföG", "No, I do not receive BAföG")
    fill_field(doc, "Former Deutschlandstipendium funding", "Not applicable — leave blank")
    fill_field(doc, "How did you find out", "[FILL IN — Internet / TUM website / recommendation / DAAD / other]")

    doc.add_page_break()

    # ── COVER LETTER 1 ────────────────────────────────────────────────────────
    heading(doc, "Part 2: Cover Letter — Draft 1", level=1)
    heading(doc, "Angle: Academic Excellence + Leadership + Personal Circumstances", level=2)
    body(doc, "Approximately 555 words | Max 1.5 pages | English", italic=True, color=MID_GRAY)

    divider(doc)

    cl1_intro = doc.add_paragraph()
    cl1_intro.paragraph_format.space_before = Pt(8)
    cl1_intro.paragraph_format.space_after = Pt(4)
    for r_text, bold in [
        ("Cover Letter — Deutschlandstipendium Application\n", True),
        ("MSc Finance and Information Management\n", False),
        ("Technical University of Munich\n", False),
        ("\nDear Selection Committee,\n", False),
    ]:
        r = cl1_intro.add_run(r_text)
        sf(r, size=11, bold=bold, color=DARK_GRAY)

    cl1_file = os.path.join(ESSAY_DIR, "TUM_cover_letter_draft1.md")
    paragraphs = read_cover_letter_body(cl1_file)
    for para in paragraphs:
        p = doc.add_paragraph(para)
        p.paragraph_format.space_before = Pt(0)
        p.paragraph_format.space_after = Pt(8)
        for r in p.runs:
            sf(r, size=11, color=DARK_GRAY)

    divider(doc)
    doc.add_page_break()

    # ── COVER LETTER 2 ────────────────────────────────────────────────────────
    heading(doc, "Part 3: Cover Letter — Draft 2", level=1)
    heading(doc, "Angle: Germany Connection + Digital Finance Career Path", level=2)
    body(doc, "Approximately 530 words | Max 1.5 pages | English", italic=True, color=MID_GRAY)

    divider(doc)

    cl2_intro = doc.add_paragraph()
    cl2_intro.paragraph_format.space_before = Pt(8)
    cl2_intro.paragraph_format.space_after = Pt(4)
    for r_text, bold in [
        ("Cover Letter — Deutschlandstipendium Application\n", True),
        ("MSc Finance and Information Management\n", False),
        ("Technical University of Munich\n", False),
        ("\nDear Selection Committee,\n", False),
    ]:
        r = cl2_intro.add_run(r_text)
        sf(r, size=11, bold=bold, color=DARK_GRAY)

    cl2_file = os.path.join(ESSAY_DIR, "TUM_cover_letter_draft2.md")
    paragraphs2 = read_cover_letter_body(cl2_file)
    for para in paragraphs2:
        p = doc.add_paragraph(para)
        p.paragraph_format.space_before = Pt(0)
        p.paragraph_format.space_after = Pt(8)
        for r in p.runs:
            sf(r, size=11, color=DARK_GRAY)

    divider(doc)

    body(doc,
         "\n\nAll information is based on Lauren's provided resume, LinkedIn, and prior essay materials as of May 30, 2026. "
         "Fields marked in orange must be confirmed and filled in by Lauren before submission.",
         size=9, italic=True, color=MID_GRAY, space_before=8)

    doc.save(output_path)
    print(f"Saved: {output_path}")


if __name__ == "__main__":
    build_tum_doc(os.path.join(ESSAY_DIR, "TUM_Deutschlandstipendium.docx"))
