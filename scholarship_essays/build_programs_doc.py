"""Build a Word document from the German programs research markdown file."""

from docx import Document
from docx.shared import Pt, RGBColor, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
import os

ESSAY_DIR = os.path.dirname(os.path.abspath(__file__))

DARK_BLUE = RGBColor(0x1F, 0x3A, 0x6B)
ORANGE = RGBColor(0xC0, 0x55, 0x00)
DARK_GRAY = RGBColor(0x33, 0x33, 0x33)
MID_GRAY = RGBColor(0x55, 0x55, 0x55)
RED = RGBColor(0xAA, 0x00, 0x00)
GREEN = RGBColor(0x1A, 0x6B, 0x2A)


def set_font(run, name="Calibri", size=11, bold=False, italic=False, color=None):
    run.font.name = name
    run.font.size = Pt(size)
    run.bold = bold
    run.italic = italic
    if color:
        run.font.color.rgb = color


def add_para(doc, text, size=11, bold=False, italic=False, color=None,
             space_before=0, space_after=6, align=WD_ALIGN_PARAGRAPH.LEFT):
    p = doc.add_paragraph()
    p.alignment = align
    p.paragraph_format.space_before = Pt(space_before)
    p.paragraph_format.space_after = Pt(space_after)
    run = p.add_run(text)
    set_font(run, size=size, bold=bold, italic=italic, color=color)
    return p


def add_heading(doc, text, level=1):
    colors = {1: DARK_BLUE, 2: DARK_BLUE, 3: ORANGE}
    sizes = {1: 16, 2: 13, 3: 11}
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(12 if level == 1 else 8)
    p.paragraph_format.space_after = Pt(4)
    run = p.add_run(text)
    set_font(run, size=sizes.get(level, 11), bold=True,
             color=colors.get(level, DARK_BLUE))
    return p


def add_bullet(doc, text, indent=0.3):
    p = doc.add_paragraph(style="List Bullet")
    p.paragraph_format.left_indent = Inches(indent)
    p.paragraph_format.space_before = Pt(1)
    p.paragraph_format.space_after = Pt(2)
    run = p.add_run(text)
    set_font(run, size=10.5, color=DARK_GRAY)
    return p


def add_warning(doc, text):
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(4)
    p.paragraph_format.space_after = Pt(6)
    run = p.add_run(text)
    set_font(run, size=10.5, bold=True, color=RED)
    return p


def add_url_line(doc, label, url):
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(1)
    p.paragraph_format.space_after = Pt(2)
    label_run = p.add_run(f"{label}: ")
    set_font(label_run, size=10, bold=True, color=DARK_BLUE)
    url_run = p.add_run(url)
    set_font(url_run, size=10, color=MID_GRAY, italic=True)
    return p


def add_divider(doc):
    p = doc.add_paragraph("─" * 80)
    p.paragraph_format.space_before = Pt(6)
    p.paragraph_format.space_after = Pt(6)
    for run in p.runs:
        run.font.color.rgb = RGBColor(0xCC, 0xCC, 0xCC)
        run.font.size = Pt(8)


def add_program(doc, rank, university, program, deadline, urgency, fit_text,
                documents, language_req, essay_req, urls, notes=None):

    add_heading(doc, f"#{rank}  {university} — {program}", level=2)

    # Rank label and urgency
    row_p = doc.add_paragraph()
    row_p.paragraph_format.space_before = Pt(0)
    row_p.paragraph_format.space_after = Pt(4)
    rank_run = row_p.add_run(f"Deadline: ")
    set_font(rank_run, size=10.5, bold=True, color=DARK_GRAY)
    dl_run = row_p.add_run(deadline)
    urgent = "URGENT" in urgency or "⚠️" in urgency
    set_font(dl_run, size=10.5, bold=urgent, color=RED if urgent else DARK_BLUE)
    if urgency:
        urg_run = row_p.add_run(f"  {urgency}")
        set_font(urg_run, size=10.5, bold=True, color=RED if urgent else GREEN)

    # Why Lauren fits
    add_heading(doc, "Why This Program Fits Your Profile", level=3)
    add_para(doc, fit_text, size=10.5, color=DARK_GRAY, space_after=4)

    # Documents
    add_heading(doc, "Required Application Documents", level=3)
    for doc_item in documents:
        add_bullet(doc, doc_item)

    # Language
    add_heading(doc, "Language Requirements", level=3)
    for lang_item in language_req:
        add_bullet(doc, lang_item)

    # Essay
    add_heading(doc, "Essay / Writing Requirement", level=3)
    add_para(doc, essay_req, size=10.5, color=DARK_GRAY, space_after=4)

    # Notes
    if notes:
        add_heading(doc, "Important Notes", level=3)
        add_para(doc, notes, size=10.5, color=ORANGE, space_after=4)

    # URLs
    add_heading(doc, "Key URLs", level=3)
    for label, url in urls:
        add_url_line(doc, label, url)

    add_divider(doc)


def build_programs_doc(output_path):
    doc = Document()

    for section in doc.sections:
        section.top_margin = Inches(1)
        section.bottom_margin = Inches(1)
        section.left_margin = Inches(1.1)
        section.right_margin = Inches(1.1)

    # Cover
    title = doc.add_heading("German Master's Programs — Recommended for Lauren", level=0)
    title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    for run in title.runs:
        run.font.name = "Calibri"
        run.font.color.rgb = DARK_BLUE

    sub = doc.add_paragraph("English-Taught Programs | Ranked by Profile Compatibility | Deadlines May–July 2026")
    sub.alignment = WD_ALIGN_PARAGRAPH.CENTER
    for run in sub.runs:
        run.font.name = "Calibri"
        run.font.size = Pt(11)
        run.font.italic = True
        run.font.color.rgb = MID_GRAY

    add_warning(doc, "⚠️  URGENT: University of Göttingen MA Digital Humanities — Deadline June 1, 2026 (2 days away). "
                     "No essay or letter of recommendation required. Apply immediately.")
    add_warning(doc, "⚠️  URGENT: FU Berlin MA North American Studies — Deadline May 31, 2026 (tomorrow). "
                     "Non-German degree holders must apply via uni-assist (VPD takes 4–6 weeks). "
                     "Contact JFKI admissions directly about feasibility.")

    # Profile summary
    add_heading(doc, "Your Profile (Used for Ranking)", level=2)
    profile_items = [
        "BS Information Systems and Accounting, University of Utah — GPA 3.912",
        "Technical skills: Python, SQL, database systems, business analytics, AI for Business Processes",
        "Congress-Bundestag State Department fellowship, Berlin (intensive German training, 2024–25)",
        "Study abroad: University of Waikato, New Zealand",
        "U.S. Dept. of Commerce internship (policy data analysis, Census datasets)",
        "YWCA volunteer: tutoring immigrants and refugees; German teaching (elementary)",
        "German: intermediate (B1–B2, from fellowship training and self-study)",
        "English: native speaker (C2 automatic exemption on all English requirements)",
        "Career goals: computational linguistics / NLP, transcultural studies, international policy, cultural diplomacy",
    ]
    for item in profile_items:
        add_bullet(doc, item)

    # Summary table (text version)
    add_heading(doc, "At a Glance: All 7 Programs", level=2)
    table_data = [
        ("Rank", "University", "Program", "Deadline", "Essay?"),
        ("1", "Heidelberg", "MA Transcultural Studies", "Mid-June 2026", "Yes — 900 words"),
        ("2 ⚠️", "Göttingen", "MA Digital Humanities", "June 1, 2026 URGENT", "No"),
        ("3", "Tübingen", "MA Computational Linguistics", "July 15, 2026", "Check portal"),
        ("4", "Münster", "MA National & Transnational Studies", "July 15, 2026", "Yes — 2,000-word scholarly essay"),
        ("5", "Stuttgart", "MSc Computational Linguistics", "July 15 (non-EU: inquire)", "Check portal"),
        ("6", "Bremen", "MA Transcultural Studies", "June 15, 2026", "Yes — 2-page letter"),
        ("7 ⚠️", "FU Berlin", "MA North American Studies", "May 31, 2026 (tomorrow)", "No"),
    ]

    tbl = doc.add_table(rows=len(table_data), cols=5)
    tbl.style = "Light Grid Accent 1"
    for i, row_data in enumerate(table_data):
        row = tbl.rows[i]
        for j, cell_text in enumerate(row_data):
            cell = row.cells[j]
            cell.text = cell_text
            for paragraph in cell.paragraphs:
                for run in paragraph.runs:
                    run.font.name = "Calibri"
                    run.font.size = Pt(9)
                    if i == 0:
                        run.bold = True
                        run.font.color.rgb = DARK_BLUE
                    elif "URGENT" in cell_text or "tomorrow" in cell_text:
                        run.font.color.rgb = RED
    doc.add_paragraph("")

    doc.add_page_break()

    # Program 1
    add_program(
        doc,
        rank=1,
        university="Heidelberg University",
        program="MA Transcultural Studies (MATS)",
        deadline="Mid-June 2026 (recommended for non-EU) — no hard deadline, mid-September absolute latest",
        urgency="Apply now",
        fit_text=(
            "This is the program Lauren is already writing essays for, and it's the strongest overall fit. "
            "Her Congress-Bundestag fellowship year in Berlin, YWCA work with immigrants and refugees, "
            "self-taught German through poetry, study abroad in New Zealand, and stated career interest in "
            "cultural diplomacy all map directly onto the program's focus on cross-cultural encounter, "
            "transcultural exchange, and belonging. The motivation letter essays already drafted for the HAUS "
            "scholarship can be directly adapted (adjust length to 900 words). "
            "Note: The BA requirement is humanities/cultural/social sciences. Lauren's IS/Accounting "
            "background should be framed through its international and social science dimensions "
            "(Understanding the Global Economy, international accounting, policy-oriented data analysis at DoC)."
        ),
        documents=[
            "Certified copy of BA degree (or equivalent)",
            "Certified copy of transcript of records",
            "Certified copies of language certificates: English C1 minimum + TWO additional languages",
            "Letter of motivation in English — maximum 900 words",
            "CV in English — maximum 2 pages",
            "Formal signed statement confirming you wrote the motivation letter yourself",
        ],
        language_req=[
            "English: minimum C1 — TOEFL iBT 100, IELTS 7.0, TOEIC 785, or CAE. U.S. degree holders typically satisfy this via passport/degree.",
            "Two additional languages required (German intermediate from fellowship counts; consider French or Spanish if applicable).",
        ],
        essay_req="Letter of motivation: 900 words maximum, in English. Free-form statement on why Transcultural Studies, why Heidelberg, your academic background, language skills, and scholarly or professional interest. Use and adapt the HAUS scholarship essay drafts 4–6 already written.",
        urls=[
            ("Application portal", "https://www.hcts.uni-heidelberg.de/en/studies/masters-program-mats/application"),
            ("Program page", "https://www.uni-heidelberg.de/en/study/all-subjects/transcultural-studies/transcultural-studies-master"),
            ("DAAD listing", "https://www2.daad.de/deutschland/studienangebote/international-programmes/en/detail/4273/"),
        ],
        notes="Semester fee: €161.10 + €1,500/semester non-EU tuition (Baden-Württemberg policy). Budget approximately €3,300/semester in fees alone.",
    )

    # Program 2
    add_program(
        doc,
        rank=2,
        university="University of Göttingen",
        program="MA Digital Humanities",
        deadline="June 1, 2026 — applies equally to EU and non-EU applicants",
        urgency="⚠️ APPLY TODAY OR TOMORROW — 2 DAYS REMAINING",
        fit_text=(
            "This is the strongest bridge between Lauren's existing technical skills and her humanistic interests. "
            "Her IS/Accounting degree with Python, database systems, business analytics, and AI for Business "
            "Processes maps directly onto Path B of the admission requirements (42+ credits in humanities/social "
            "sciences + 18+ credits in Python and Data Science). IS and Accounting are social science-adjacent "
            "fields with significant quantitative/analytical content. Admission is open — no competitive "
            "restriction. As a native English speaker with a U.S. university degree, Lauren is automatically "
            "exempt from the English C1 requirement. No essay, no letter of recommendation — just an online "
            "application with transcripts. This is the fastest application to submit on this list."
        ),
        documents=[
            "Online application only — no paper documents required",
            "Transcripts and degree certificate (upload as PDF)",
            "English C1 proof — EXEMPT as a native English speaker with a U.S. degree",
        ],
        language_req=[
            "English: C1 required, but native speakers from English-speaking countries are automatically exempt.",
            "German: not required at admission. Students with less than DSH-1 German must complete German language modules (12 credits) as part of soft skills — Lauren's intermediate German from her fellowship likely satisfies or reduces this requirement.",
        ],
        essay_req="No essay or letter of motivation required. Online application with transcripts only.",
        urls=[
            ("Apply here (online portal)", "https://masterbewerbung.phil.uni-goettingen.de/de/master_applications/new?master_programme_id=56"),
            ("Program page", "https://www.uni-goettingen.de/en/digital+humanities+%28m.a.%29/652252.html"),
            ("Faculty MA overview", "https://www5.uni-goettingen.de/en/107958.html"),
        ],
        notes="Decisions sent mid–late July 2026. Additional application period opens mid–late September but non-EU applicants needing a visa MUST apply in the June 1 regular window. No tuition fees beyond standard semester contribution (~€350/semester).",
    )

    # Program 3
    add_program(
        doc,
        rank=3,
        university="University of Tübingen",
        program="MA Computational Linguistics",
        deadline="July 15, 2026 for non-EU applicants (winter semester 2026/27)",
        urgency="Standard — apply by mid-June to allow processing time",
        fit_text=(
            "Tübingen is one of Germany's leading CL programs with particular strengths in NLP, "
            "computational semantics, and corpus linguistics — precisely the research areas Lauren has "
            "identified in her application essays. Admission is open (not numerus clausus), meaning "
            "her IS/Accounting background with Python and data coursework gives her a realistic path. "
            "Foreign students without German at admission must acquire German during their first year — "
            "Lauren already has intermediate German from her fellowship, giving her an advantage. "
            "The essays already drafted for the HAUS scholarship (CL Drafts 1–3) can be directly adapted "
            "for any required motivation letter."
        ),
        documents=[
            "Online application via the University of Tübingen portal",
            "Bachelor's degree certificate or equivalent",
            "Transcript of records",
            "English language proof (U.S. graduates are typically exempt — verify with the portal)",
            "Any additional documents specified in admission regulations (check the portal)",
        ],
        language_req=[
            "English: proof of proficiency required. Native English speakers from English-speaking countries are typically exempt.",
            "German: not required at admission. Students without German proof must acquire it during the first year of study. Lauren's intermediate German from her fellowship counts in her favor.",
        ],
        essay_req="No formal essay listed as a general requirement. Check the program's admission portal for any supplemental motivation letter or personal statement requirement.",
        urls=[
            ("Program page and application", "https://uni-tuebingen.de/en/study/finding-a-course/degree-programs-available/detail/course/computational-linguistics-master/"),
            ("DAAD listing", "https://www.daad.de/en/studying-in-germany/universities/all-degree-programmes/detail/eberhard-karls-university-tuebingen-computer-linguistics-computational-linguistics-w6834/"),
            ("International applicant info", "https://uni-tuebingen.de/en/study/applying-to-tuebingen/international-applicants/"),
        ],
        notes="Tuition: €1,500/semester for non-EU students (Baden-Württemberg). This is in addition to the standard semester fee.",
    )

    # Program 4
    add_program(
        doc,
        rank=4,
        university="University of Münster",
        program="MA National and Transnational Studies (NTS)",
        deadline="Phase 2: May 16 – July 15, 2026. Notifications by end of July.",
        urgency="Non-EU: apply early in Phase 2 for maximum visa processing time",
        fit_text=(
            "The MA NTS focuses on national and transnational identities through literature, culture, "
            "language, and social science — directly aligned with Lauren's interest in how cultures "
            "interact and her Congress-Bundestag experience navigating U.S.–Germany institutional "
            "differences. The program specifically values stays abroad and internships, both of which "
            "Lauren has extensively. As a native English speaker, Lauren automatically meets the C2 "
            "English requirement (TOEFL 109+ equivalent) without needing a language test. "
            "The main risk: the BA background requirement asks for literature, linguistics, cultural studies, "
            "political science, sociology, etc. Lauren's IS/Accounting degree doesn't fall neatly into "
            "these categories. However, her coursework in global economics, international accounting, and "
            "Understanding the Global Economy — combined with extensive intercultural experience — provides "
            "a compelling non-traditional entry point."
        ),
        documents=[
            "Online application form (Bewerbungsportal — apply at the university portal, NOT via uni-assist)",
            "Proof of school-leaving qualification (high school diploma or Abitur equivalent)",
            "Undergraduate diploma(s)",
            "Transcript(s) of records",
            "Scholarly essay (~2,000 words / approx. 5 pages, in English) — on the year's specific topic posted on the admissions page",
            "CV in tabular form",
            "Proof of language skills for non-native speakers only (Lauren as a U.S. native speaker is EXEMPT)",
            "Proof of stay abroad (Congress-Bundestag fellowship in Berlin; University of Waikato study abroad)",
            "Proof of internships (U.S. Dept. of Commerce, Maverik, Breeze Airways, American Red Cross)",
        ],
        language_req=[
            "English: C2 level required. As a native English speaker with a U.S. degree, Lauren is automatically exempt from submitting a language test.",
            "German: desirable but not required for admission.",
        ],
        essay_req="Scholarly essay of approximately 2,000 words (5 pages) on the specific topic set for the 2026/27 cycle. This is an academic writing sample demonstrating scholarly argumentation — not a personal statement. The essay topic is posted on the admissions page each year. No letter of recommendation is accepted.",
        urls=[
            ("Admissions and essay topic", "https://www.uni-muenster.de/MA_transnational/Admissions/index.html"),
            ("Program home", "https://www.sos.uni-muenster.de/MA_transnational/"),
            ("FAQ", "https://www.sos.uni-muenster.de/MA_transnational/faq.html"),
        ],
    )

    # Program 5
    add_program(
        doc,
        rank=5,
        university="University of Stuttgart",
        program="MSc Computational Linguistics",
        deadline="July 15, 2026 for Germans and EU applicants. Non-EU: contact the international office to confirm.",
        urgency="Non-EU: inquire immediately about your specific deadline",
        fit_text=(
            "Stuttgart's MSc CL is 100% English-taught with approximately 90% international students — "
            "one of the most internationally oriented CL programs in Germany. It has a strong applied NLP "
            "profile with close industry connections in the Stuttgart region and optional integrated "
            "internships. Lauren's Python, data analysis, AI coursework, and two years of applied data "
            "work at the U.S. Department of Commerce give her a strong computational entry point. "
            "No German language requirement makes this the most linguistically accessible CL program on this list."
        ),
        documents=[
            "Online application via the Stuttgart international degree-seeking student portal",
            "Bachelor's degree certificate",
            "Transcripts",
            "English C1 language proof (TOEFL iBT 79, IELTS 6.5, or equivalent) — Lauren as a native English speaker is exempt",
            "CV",
            "Letter of motivation (verify specific requirements at the portal)",
        ],
        language_req=[
            "English: C1 CEFR minimum (TOEFL iBT 79 / IELTS 6.5). Lauren as a native U.S. English speaker is exempt.",
            "German: not required for this program.",
        ],
        essay_req="Check the application portal for letter of motivation requirements. The HAUS CL essays already drafted (Drafts 1–3) can be directly adapted.",
        urls=[
            ("Application portal (international students)", "https://www.uni-stuttgart.de/studium/bewerbung/international-degree/"),
            ("DAAD program listing", "https://www2.daad.de/deutschland/studienangebote/international-programmes/en/detail/4239/"),
            ("Program overview (DAAD)", "https://www.daad.de/en/studying-in-germany/universities/all-degree-programmes/detail/university-of-stuttgart-computational-linguistics-w39397/"),
        ],
        notes="Non-EU students: the application system shows the non-EU deadline as 'please enquire.' Contact the Stuttgart international admissions office as soon as possible to confirm whether you can still apply for winter 2026. Tuition: €1,500/semester non-EU (Baden-Württemberg).",
    )

    # Program 6
    add_program(
        doc,
        rank=6,
        university="University of Bremen",
        program="MA Transcultural Studies (MATS)",
        deadline="June 15, 2026 (application period: May 1 – June 15)",
        urgency="Apply by June 15 — but verify German proficiency status first",
        fit_text=(
            "Bremen's MATS covers anthropology, cultural research, religious studies, and literature "
            "from a transdisciplinary, postcolonial, and global perspective — thematically similar to "
            "Heidelberg's program. Lauren's YWCA work, Germany fellowship, and international career "
            "interest all align. However, this program is ranked lower than Heidelberg's TS because "
            "it primarily teaches in German (with some English courses), and it requires German C1 "
            "certification for enrollment. Lauren's intermediate German from her fellowship "
            "is approximately B1–B2; she may need to sit a Goethe-Institut or DaF examination "
            "to certify her level before she can enroll."
        ),
        documents=[
            "Online application via Bremen portal",
            "Bachelor's degree certificate (minimum 130 credit points)",
            "Transcript of records",
            "Proof of German C1 proficiency (Goethe-Zertifikat C1 or DaF score of 16+ points)",
            "English B2 proof (Lauren's U.S. degree satisfies this)",
            "Letter of motivation (maximum 2 A4 pages)",
            "Language certificates",
        ],
        language_req=[
            "German: C1 REQUIRED — Goethe-Zertifikat C1 or DaF exam score of 16+ points. This is the main barrier for Lauren.",
            "English: B2 minimum — Lauren far exceeds this as a native speaker.",
        ],
        essay_req="Letter of motivation: maximum 2 A4 pages. Can be adapted from the HAUS TS scholarship essay drafts (4–6), adjusted for Bremen's specific program focus.",
        urls=[
            ("Application page", "https://www.uni-bremen.de/en/kultur/study/ma-transcultural-studies/the-application"),
            ("Program home", "https://www.uni-bremen.de/en/kultur/study/ma-transcultural-studies"),
            ("Bremen master programs list", "https://www.uni-bremen.de/en/master/master-programs"),
        ],
        notes="The program is taught primarily in German (with some English courses). Lauren should take a Goethe-Institut C1 exam or DaF test before applying to confirm she meets the language requirement. If her German is assessed as B2, she should target Heidelberg's TS program instead (no German requirement).",
    )

    # Program 7
    add_program(
        doc,
        rank=7,
        university="Freie Universität Berlin",
        program="MA North American Studies (John F. Kennedy Institute)",
        deadline="May 31, 2026 — TOMORROW",
        urgency="⚠️ Contact JFKI admissions TODAY about late application options",
        fit_text=(
            "The JFKI MA covers North American culture, history, literature, political science, sociology, "
            "and economics — all areas relevant to Lauren's background as a U.S. citizen with government, "
            "international policy, and intercultural experience. Her DoC internship, Congress-Bundestag "
            "fellowship (placing her as a U.S. representative in Germany), and interest in U.S.–Germany "
            "relations make her a natural candidate. No motivation letter or recommendation letter required — "
            "just an online form and CV. However, the practical barrier is the uni-assist VPD requirement "
            "(4–6 weeks to process) for non-German degree holders, combined with tomorrow's deadline. "
            "This program is best targeted for the 2027 application cycle."
        ),
        documents=[
            "Online application form (filled out and submitted online)",
            "Curriculum vitae (uploaded to the portal)",
            "English C1 proof — Lauren's U.S. passport automatically satisfies this",
            "VPD (Vorprüfungsdokumentation) from uni-assist — required for non-German degree holders, takes 4–6 weeks",
            "Bachelor's degree in: North American Studies, History, Cultural Studies, Literature, Political Science, Sociology, Economics, or interdisciplinary with 90+ ECTS in those subjects",
            "NO motivation letter or recommendation letter required or accepted",
        ],
        language_req=[
            "English: C1 required. As a U.S. citizen with a U.S. university degree, Lauren automatically qualifies.",
            "German: not required for this program.",
        ],
        essay_req="No essay, motivation letter, or recommendation letter required. Application consists of the online form and CV only.",
        urls=[
            ("Applicant information", "https://www.jfki.fu-berlin.de/en/academics/ma/Bewerber/index.html"),
            ("Application FAQ", "https://www.jfki.fu-berlin.de/en/academics/ma/FAQ/index.html"),
            ("FU Berlin application deadlines", "https://www.fu-berlin.de/en/studium/bewerbung/bewerbungsfristen/index.html"),
            ("Program page", "https://www.fu-berlin.de/en/studium/studienangebot/master/nordamerikastudien/index.html"),
        ],
        notes="RECOMMENDED ACTION: Call or email the JFKI admissions office today (May 30) to ask about late application options or direct enrollment. If not possible this cycle, add this to your 2027 application list — it would be a strong application given your profile.",
    )

    # Also worth noting
    add_heading(doc, "Also Worth Noting: Past Deadline for 2026 (Apply in 2027)", level=2)
    add_para(doc, "Saarland University — MSc Language Science and Technology (LST)", size=11, bold=True, color=DARK_BLUE)
    add_para(doc,
             "One of Europe's most prestigious programs in language technology and computational linguistics, "
             "with world-class NLP faculty and research labs. The February 15 deadline for winter 2026 has "
             "already passed, but this program should be at the top of Lauren's list for the 2027 application cycle.",
             size=10.5, color=DARK_GRAY)
    add_url_line(doc, "Program page", "https://www.uni-saarland.de/en/study/programmes/master/lst.html")
    add_url_line(doc, "DAAD listing", "https://www2.daad.de/deutschland/studienangebote/international-programmes/en/detail/3739/")
    add_para(doc, "Deadline: February 15 each year (for October start)", size=10, italic=True, color=ORANGE)

    add_para(doc,
             "\n\nAll information verified from official university and DAAD sources as of May 30, 2026. "
             "Deadlines and requirements can change — always confirm directly with each university's "
             "admissions office before submitting an application.",
             size=9, italic=True, color=MID_GRAY)

    doc.save(output_path)
    print(f"Saved: {output_path}")


if __name__ == "__main__":
    build_programs_doc(os.path.join(ESSAY_DIR, "HAUS_German_Programs_Ranked.docx"))
