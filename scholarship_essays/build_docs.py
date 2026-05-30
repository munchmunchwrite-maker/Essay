"""Build Word documents from the scholarship essay markdown files."""

from docx import Document
from docx.shared import Pt, RGBColor, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH
import re
import os

ESSAY_DIR = os.path.dirname(os.path.abspath(__file__))

# Essay files in order, with human-readable labels
ALL_ESSAYS = [
    (
        "essay_CL_draft1_technical_professional.md",
        "MSc Computational Linguistics — Draft 1",
        "Technical / Professional Angle",
    ),
    (
        "essay_CL_draft2_personal_narrative.md",
        "MSc Computational Linguistics — Draft 2",
        "Personal Narrative / Language Journey Angle",
    ),
    (
        "essay_CL_draft3_research_focused.md",
        "MSc Computational Linguistics — Draft 3",
        "Research-Focused / Intellectual Question Angle",
    ),
    (
        "essay_TS_draft1_cultural_bridge.md",
        "MA Transcultural Studies — Draft 4",
        "Cultural Bridge-Builder Angle",
    ),
    (
        "essay_TS_draft2_personal_transformation.md",
        "MA Transcultural Studies — Draft 5",
        "Personal Transformation Angle",
    ),
    (
        "essay_TS_draft3_professional_policy.md",
        "MA Transcultural Studies — Draft 6",
        "Professional / International Policy Angle",
    ),
]

# The two best essays (by filename) and why
BEST_ESSAYS = [
    (
        "essay_CL_draft2_personal_narrative.md",
        "MSc Computational Linguistics — Draft 2",
        "Personal Narrative / Language Journey Angle",
        "Selected as best CL essay: the literacy-to-language-to-NLP arc is the most emotionally "
        "resonant and cohesive narrative across the three CL drafts. It leads with a specific, "
        "memorable hook, builds organically through the Gastmutter experience, names Dr. Letitia "
        "Pârcălăbescu as a concrete Heidelberg draw, and closes the financial need section with "
        "genuine personal stakes.",
    ),
    (
        "essay_TS_draft5_note",  # placeholder key — actual file below
        "MA Transcultural Studies — Draft 5",
        "Personal Transformation Angle",
        "Selected as best TS essay: the opening on late literacy and language as access is "
        "unusually compelling and directly parallels the field (how language shapes belonging). "
        "The 'Us and Them' podcast reference feels organic rather than name-dropped, and the "
        "Congress-Bundestag disorientation anecdote — realizing one's assumptions are culturally "
        "specific — is both vivid and intellectually relevant to Transcultural Studies.",
    ),
]

BEST_ESSAY_FILES = [
    "essay_CL_draft2_personal_narrative.md",
    "essay_TS_draft2_personal_transformation.md",
]


def read_essay_body(filepath):
    """Read a markdown essay file and return just the essay body text as paragraphs."""
    with open(filepath, encoding="utf-8") as f:
        raw = f.read()

    # Strip markdown metadata header (lines starting with #, ##, or the --- word count footer)
    lines = raw.splitlines()
    body_lines = []
    in_header = True
    for line in lines:
        stripped = line.strip()
        # Skip the header block (everything up to and including the first blank line after ---'s)
        if in_header:
            if stripped.startswith("#") or stripped == "---" or stripped == "":
                continue
            else:
                in_header = False
        # Stop at the word count footer
        if stripped.startswith("---") or stripped.startswith("**Word count"):
            break
        body_lines.append(line)

    # Collapse multiple blank lines into one, strip trailing whitespace
    paragraphs = []
    current = []
    for line in body_lines:
        if line.strip() == "":
            if current:
                paragraphs.append(" ".join(current).strip())
                current = []
        else:
            current.append(line.strip())
    if current:
        paragraphs.append(" ".join(current).strip())

    return [p for p in paragraphs if p]


def set_body_font(paragraph, size=11):
    for run in paragraph.runs:
        run.font.size = Pt(size)
        run.font.name = "Calibri"


def add_heading(doc, text, level=1, color=None):
    p = doc.add_heading(text, level=level)
    p.alignment = WD_ALIGN_PARAGRAPH.LEFT
    for run in p.runs:
        run.font.name = "Calibri"
        if color:
            run.font.color.rgb = RGBColor(*color)
    return p


def add_essay_to_doc(doc, filename, title, subtitle, selection_note=None):
    filepath = os.path.join(ESSAY_DIR, filename)

    # Section title
    add_heading(doc, title, level=1, color=(0x1F, 0x3A, 0x6B))  # dark blue

    # Subtitle (angle)
    sub = doc.add_paragraph(subtitle)
    sub.alignment = WD_ALIGN_PARAGRAPH.LEFT
    for run in sub.runs:
        run.font.name = "Calibri"
        run.font.size = Pt(10)
        run.font.italic = True
        run.font.color.rgb = RGBColor(0x55, 0x55, 0x55)

    # Optional selection note
    if selection_note:
        note_para = doc.add_paragraph()
        note_run = note_para.add_run("Why this essay was selected: ")
        note_run.bold = True
        note_run.font.size = Pt(10)
        note_run.font.name = "Calibri"
        note_run.font.color.rgb = RGBColor(0x1F, 0x3A, 0x6B)
        body_run = note_para.add_run(selection_note)
        body_run.font.size = Pt(10)
        body_run.font.name = "Calibri"
        body_run.font.color.rgb = RGBColor(0x33, 0x33, 0x33)
        note_para.paragraph_format.space_after = Pt(8)

    # Essay body
    paragraphs = read_essay_body(filepath)
    for i, para_text in enumerate(paragraphs):
        p = doc.add_paragraph(para_text)
        p.paragraph_format.space_before = Pt(0)
        p.paragraph_format.space_after = Pt(8)
        p.paragraph_format.first_line_indent = Inches(0)
        set_body_font(p, size=11)

    # Divider
    doc.add_paragraph("")


def build_all_essays_doc(output_path):
    doc = Document()

    # Page margins
    for section in doc.sections:
        section.top_margin = Inches(1)
        section.bottom_margin = Inches(1)
        section.left_margin = Inches(1.15)
        section.right_margin = Inches(1.15)

    # Cover heading
    title = doc.add_heading("HAUS Study Scholarship — All Essay Drafts", level=0)
    title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    for run in title.runs:
        run.font.name = "Calibri"
        run.font.color.rgb = RGBColor(0x1F, 0x3A, 0x6B)

    sub = doc.add_paragraph("Applicant: Lauren  |  Application Year: 2026–27")
    sub.alignment = WD_ALIGN_PARAGRAPH.CENTER
    for run in sub.runs:
        run.font.name = "Calibri"
        run.font.size = Pt(11)
        run.font.italic = True
        run.font.color.rgb = RGBColor(0x55, 0x55, 0x55)

    note = doc.add_paragraph(
        "Six drafts are included below: Drafts 1–3 target the MSc in Computational Linguistics "
        "at Heidelberg University; Drafts 4–6 target the MA in Transcultural Studies. Each draft "
        "addresses all four prompt questions (why Heidelberg, why HAUS, study proposal and future "
        "plans, HAUS Ambassador role) within the 500–600 word limit."
    )
    note.alignment = WD_ALIGN_PARAGRAPH.LEFT
    for run in note.runs:
        run.font.name = "Calibri"
        run.font.size = Pt(10)
        run.font.color.rgb = RGBColor(0x33, 0x33, 0x33)
    note.paragraph_format.space_after = Pt(16)

    doc.add_page_break()

    for filename, title_text, subtitle in ALL_ESSAYS:
        add_essay_to_doc(doc, filename, title_text, subtitle)
        doc.add_page_break()

    doc.save(output_path)
    print(f"Saved: {output_path}")


def build_best_essays_doc(output_path):
    doc = Document()

    for section in doc.sections:
        section.top_margin = Inches(1)
        section.bottom_margin = Inches(1)
        section.left_margin = Inches(1.15)
        section.right_margin = Inches(1.15)

    title = doc.add_heading("HAUS Study Scholarship — Top 2 Essays", level=0)
    title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    for run in title.runs:
        run.font.name = "Calibri"
        run.font.color.rgb = RGBColor(0x1F, 0x3A, 0x6B)

    sub = doc.add_paragraph("Applicant: Lauren  |  Application Year: 2026–27")
    sub.alignment = WD_ALIGN_PARAGRAPH.CENTER
    for run in sub.runs:
        run.font.name = "Calibri"
        run.font.size = Pt(11)
        run.font.italic = True
        run.font.color.rgb = RGBColor(0x55, 0x55, 0x55)

    note = doc.add_paragraph(
        "These two essays were selected from the full set of six drafts as the strongest "
        "candidates for submission. Essay 1 is written for the MSc in Computational Linguistics; "
        "Essay 2 is written for the MA in Transcultural Studies. Selection rationale is noted "
        "beneath each title."
    )
    note.alignment = WD_ALIGN_PARAGRAPH.LEFT
    for run in note.runs:
        run.font.name = "Calibri"
        run.font.size = Pt(10)
        run.font.color.rgb = RGBColor(0x33, 0x33, 0x33)
    note.paragraph_format.space_after = Pt(16)

    doc.add_page_break()

    best = [
        (
            "essay_CL_draft2_personal_narrative.md",
            "MSc Computational Linguistics — Draft 2 (Top Pick)",
            "Personal Narrative / Language Journey Angle",
            "The literacy-to-language-to-NLP arc is the most emotionally resonant and cohesive "
            "narrative across the three CL drafts. It leads with a specific, memorable hook, "
            "builds organically through the Gastmutter experience, names Dr. Letitia "
            "Pârcălăbescu as a concrete Heidelberg draw, and closes the financial need section "
            "with genuine personal stakes that connect to Lauren's mother's unfulfilled dream.",
        ),
        (
            "essay_TS_draft2_personal_transformation.md",
            "MA Transcultural Studies — Draft 5 (Top Pick)",
            "Personal Transformation Angle",
            "The opening on late literacy and language as access is unusually compelling and "
            "mirrors the field's core concerns. The HGGS 'Us and Them' podcast reference feels "
            "organic rather than name-dropped, and the Congress-Bundestag disorientation "
            "anecdote — realizing one's assumptions are culturally specific — is both vivid and "
            "directly relevant to Transcultural Studies as a discipline.",
        ),
    ]

    for filename, title_text, subtitle, selection_note in best:
        add_essay_to_doc(doc, filename, title_text, subtitle, selection_note=selection_note)
        doc.add_page_break()

    doc.save(output_path)
    print(f"Saved: {output_path}")


if __name__ == "__main__":
    build_all_essays_doc(os.path.join(ESSAY_DIR, "HAUS_All_Six_Essay_Drafts.docx"))
    build_best_essays_doc(os.path.join(ESSAY_DIR, "HAUS_Top_2_Essays.docx"))
