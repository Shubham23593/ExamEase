#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
================================================================================
 ExamEase - Final Year Project Report Generator (.docx)
================================================================================

WHAT THIS SCRIPT DOES
---------------------
1. Programmatically draws every engineering diagram required in a project
   report (system architecture, data flow diagrams, NLP pipeline, use case
   diagram, ER diagram, sequence diagram, activity workflow, AI factory
   fallback flow and the duplicate-detection flow) as high resolution PNGs.
2. Creates "screenshot placeholders" for the application screens, so that the
   student can simply right-click a placeholder inside MS Word and use
   "Change Picture" to drop in a real screenshot (size and alignment are kept).
   Alternatively, drop real images into  assets/screenshots/fig_10_05.png  and
   re-run this script - the real image is embedded automatically.
3. Builds a complete, university-format engineering project report in MS Word
   with a title page, certificate, declaration, acknowledgement, abstract,
   auto-updating Table of Contents, List of Figures, List of Tables, List of
   Abbreviations, content chapters, references and appendices.

DEPENDENCIES
------------
    pip install python-docx matplotlib

USAGE
-----
    # 1. keep this file inside the project, e.g.  docs/final-report/
    # 2. open a terminal in that folder and run:
    python ExamEase_Report_Generator.py

    # output:
    #   ExamEase_Final_Project_Report.docx      <-- open in MS Word
    #   assets/diagrams/*.png                   <-- drawn diagrams
    #   assets/screenshots/*.png                <-- placeholders / your screens
    #   assets/SCREENSHOT_GUIDE.md              <-- which screenshot goes where

AFTER OPENING THE DOCX IN MS WORD
---------------------------------
    * Word asks to update fields  ->  click "Yes"  (builds TOC, LOF, LOT).
    * Or press  Ctrl + A  then  F9  ->  update all fields / page numbers.
    * Replace placeholders: right click the placeholder image -> Change Picture.
    * Edit personal details in the CONFIG block below and re-run the script any
      time you want a freshly generated report.

AUTHOR : ExamEase project documentation tool
VERSION: 1.0
================================================================================
"""

from __future__ import annotations

import re
import sys
from pathlib import Path

# ==============================================================================
# 1. CONFIGURATION  --  EDIT ONLY THIS BLOCK BEFORE GENERATING THE REPORT
# ==============================================================================

CONFIG = {
    # ---------------------------- Project identity ---------------------------
    "project_title": "ExamEase",
    "project_subtitle": ("An NLP and Artificial Intelligence Powered "
                         "Automated Examination Question Generation System"),

    # ---------------------------- Student details ----------------------------
    "student_name": "Shubham Dalvi",
    "roll_number": "______________",
    "class": "B.E. (Computer Engineering)",

    # ---------------------------- Academic details ---------------------------
    "degree": "Bachelor of Engineering (B.E.)",
    "branch": "Computer Engineering",
    "semester": "Semester VII",
    "academic_year": "2026 - 2027",
    "month_year": "October 2026",

    # ---------------------------- Institute details --------------------------
    "university_name": "University of Mumbai",
    "college_name": "________________________________ College of Engineering",
    "department": "Department of Computer Engineering",
    "place": "Mumbai",

    # ---------------------------- Guide / staff ------------------------------
    "guide_name": "Prof. ____________________",
    "guide_designation": "Assistant Professor, Department of Computer Engineering",
    "hod_name": "Prof. ____________________",
    "hod_designation": "Head, Department of Computer Engineering",
    "principal_name": "Dr. ____________________",
    "principal_designation": "Principal",

    # ---------------------------- Feature switches ---------------------------
    "include_screenshot_placeholders": True,   # draw placeholder boxes for UI screens
    "include_appendix_source_code": True,      # embed real source code in Appendix C
    "include_mock_data_tables": True,          # table of MongoDB sample documents

    # ---------------------------- Output paths -------------------------------
    "output_docx": "ExamEase_Final_Project_Report.docx",
    "assets_dir": "assets",
}

# ------------------------------------------------------------------------------
# Measured pipeline output (produced by actually running src/services/nlp on a
# 52-sentence lecture note on NLP foundations).  Used by the Results chapter so
# that every number printed in the report is a real measurement, not a guess.
# ------------------------------------------------------------------------------
SAMPLE_NLP = {
    "source": "NLP_Fundamentals_Unit_1.pdf (8 pages, text based)",
    "sentence_count": 52,
    "token_count": 881,
    "unique_token_count": 393,
    "definition_count": 28,
    "concept_count_extracted": 25,
    "top_concepts": [
        ("Text normalization", 10.0, 3, 2),
        ("Sentence segmentation", 10.0, 4, 2),
        ("Lemmatization", 10.0, 2, 2),
        ("Part of speech tagging", 10.0, 2, 2),
        ("Concept extraction", 10.0, 4, 2),
        ("Cosine similarity", 10.0, 3, 2),
        ("Tokenization", 9.7, 2, 1),
        ("Bloom taxonomy", 9.7, 2, 1),
        ("N gram frequency analysis", 8.5, 2, 1),
        ("Jaccard similarity", 8.5, 2, 1),
    ],
    "top_unigrams": [
        ("defined", 13), ("document", 9), ("refers", 8), ("text", 7),
        ("language", 6), ("generation", 6), ("sentence", 6),
        ("frequency", 6), ("question", 5), ("tokens", 5),
    ],
    "top_bigrams": [
        ("question generation", 4), ("natural language", 3),
        ("language processing", 3), ("automated question", 3),
        ("noun phrase", 3), ("bloom taxonomy", 3),
    ],
    "top_trigrams": [
        ("part of speech", 4), ("natural language processing", 3),
        ("automated question generation", 3), ("noun phrase chunking", 2),
        ("word to its", 2), ("size of the", 2),
    ],
    "definition_examples": [
        "Natural Language Processing is defined as the branch of artificial "
        "intelligence that enables computer systems to analyse, understand and "
        "generate human language.",
        "Text normalization refers to the process of removing noise such as "
        "control characters, repeated whitespace, carriage returns and binary "
        "artefacts from the extracted document.",
        "Sentence segmentation is defined as the boundary detection problem in "
        "which punctuation marks, abbreviations and decimal numbers must be "
        "disambiguated before a split decision is taken.",
        "Tokenization is defined as the process of splitting a sentence into a "
        "sequence of lexical units called tokens.",
    ],
    "supporting_example": [
        "Text normalization is the first stage of any linguistic pipeline.",
        "Text normalization refers to the process of removing noise such as "
        "control characters, repeated whitespace, carriage returns and binary "
        "artefacts from the extracted document.",
    ],
}

# ------------------------------------------------------------------------------
# Measured Jaccard similarity values for representative question pairs.
# Computed with the real  normalizeQuestion() + calculateJaccardSimilarity()
# implementation (threshold = 0.65).
# ------------------------------------------------------------------------------
SAMPLE_JACCARD = [
    ("What is Natural Language Processing?", "Define Natural Language Processing.",
     0.75, "Duplicate"),
    ("Which component of TF IDF increases with local term frequency?",
     "Which component of the TF IDF scheme increases with the local frequency of a term?",
     0.875, "Duplicate"),
    ("A lemmatizer maps matrices to which canonical form?",
     "Which canonical form does a lemmatizer produce for matrices?",
     0.6667, "Duplicate"),
    ("How is cosine similarity defined?", "Define cosine similarity.",
     0.5, "Retained (below threshold)"),
    ("What does OCR stand for?", "Define optical character recognition.",
     0.0, "Retained (below threshold)"),
]

# ------------------------------------------------------------------------------
# Representative MCQ items used in the Results chapter (sample of the validated
# output structure produced by the generator).
# ------------------------------------------------------------------------------
SAMPLE_MCQS = [
    {
        "q": 'According to the uploaded notes, how is "Lemmatization" defined?',
        "options": [
            "The dictionary based reduction of a word to its canonical form",
            "The heuristic removal of derivational suffixes from a word",
            "The grouping of contiguous adjective-noun token sequences",
            "The conversion of an image of text into a character stream",
        ],
        "answer": "The dictionary based reduction of a word to its canonical form",
        "explain": "The notes state that lemmatization refers to the dictionary "
                   "based reduction of a word to its canonical form, mapping "
                   "matrices to matrix and criteria to criterion.",
        "concept": "Lemmatization",
        "type": "definition",
        "difficulty": "easy",
    },
    {
        "q": "Which of the following statements about the Porter stemmer is correct "
             "as per the source material?",
        "options": [
            "It is a rule based algorithm that may produce non dictionary roots",
            "It maps irregular plurals using a lookup dictionary of lemmas",
            "It measures the cosine of the angle between two token sets",
            "It assigns grammatical categories to every token of a sentence",
        ],
        "answer": "It is a rule based algorithm that may produce non dictionary roots",
        "explain": "Stemming is described as an aggressive heuristic that strips "
                   "suffixes such as -ing, -ed and -ization and can therefore "
                   "produce non dictionary roots.",
        "concept": "Stemming",
        "type": "concept",
        "difficulty": "medium",
    },
    {
        "q": "Why is noun phrase chunking performed before concept extraction?",
        "options": [
            "It groups adjacent tokens so that multi word terminology is treated "
            "as a single candidate concept",
            "It removes stopwords so that the vocabulary size is reduced",
            "It converts the document into a bag of unigrams for scoring",
            "It normalises the character encoding of the source document",
        ],
        "answer": "It groups adjacent tokens so that multi word terminology is "
                  "treated as a single candidate concept",
        "explain": "The notes explain that noun phrase chunking extracts phrases "
                   'such as "Relational Database Management System" instead of '
                   "isolated words, which improves concept quality.",
        "concept": "Noun phrase chunking",
        "type": "process",
        "difficulty": "medium",
    },
    {
        "q": "A generated question must cite which element of the NLP analysis to "
             "remain traceable to the source document?",
        "options": [
            "The extracted concept and its supporting sentence",
            "The total token count of the document",
            "The file name of the uploaded PDF",
            "The number of unique tokens in the vocabulary",
        ],
        "answer": "The extracted concept and its supporting sentence",
        "explain": "Each generated item stores a sourceConcept and an explanation "
                   "derived from the supporting sentences, which makes every "
                   "question verifiable against the uploaded notes.",
        "concept": "Grounded generation",
        "type": "application",
        "difficulty": "hard",
    },
]

# ==============================================================================
# 2. BOOTSTRAP - dependency check, matplotlib backend, output folders
# ==============================================================================

BASE_DIR = Path(__file__).resolve().parent
ASSETS = BASE_DIR / CONFIG["assets_dir"]
DIAGRAMS = ASSETS / "diagrams"
SCREENSHOTS = ASSETS / "screenshots"

for folder in (ASSETS, DIAGRAMS, SCREENSHOTS):
    folder.mkdir(parents=True, exist_ok=True)


def _require(module_name: str, pip_name: str = None):
    """Import a module or exit with a friendly installation message."""
    try:
        return __import__(module_name)
    except ImportError:
        pip_name = pip_name or module_name
        print("\n" + "=" * 74)
        print(" MISSING DEPENDENCY : %s" % pip_name)
        print("=" * 74)
        print(" Install the required packages and run the script again:\n")
        print("     pip install python-docx matplotlib\n")
        print("=" * 74 + "\n")
        sys.exit(1)


_require("docx", "python-docx")
_require("matplotlib", "matplotlib")

import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
from matplotlib.patches import FancyBboxPatch, FancyArrowPatch, Rectangle, Polygon

plt.rcParams["font.family"] = "DejaVu Sans"
plt.rcParams["savefig.bbox"] = "tight"
plt.rcParams["savefig.facecolor"] = "white"

from docx import Document
from docx.shared import Pt, Cm, Inches, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_BREAK
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.enum.section import WD_SECTION
from docx.enum.style import WD_STYLE_TYPE
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

# Colour palette used across every diagram (print friendly)
C_NAVY = "#1F3864"
C_BLUE = "#2E75B6"
C_LBLUE = "#DEEAF6"
C_TEAL = "#1F6F6B"
C_LTEAL = "#D9EFEC"
C_AMBER = "#B7791F"
C_LAMBER = "#FDF3DC"
C_GREY = "#5A5A5A"
C_LGREY = "#F1F1F1"
C_GREEN = "#2F7D4F"
C_LGREEN = "#E1F1E6"
C_RED = "#A83232"

print("=" * 74)
print(" ExamEase report generator")
print("-" * 74)
print("  project     : %s" % CONFIG["project_title"])
print("  student     : %s" % CONFIG["student_name"])
print("  output      : %s" % (BASE_DIR / CONFIG["output_docx"]).name)
print("=" * 74)

# ==============================================================================
# 3. DIAGRAM ENGINE  -  every figure in the report is drawn by this section
# ==============================================================================

FIG_INDEX = []          # [(figure_number, caption, file_name), ...]


def _canvas(w_in: float, h_in: float):
    """Return (fig, ax) with a 0..100 x 0..(100*h/w) drawing area."""
    fig = plt.figure(figsize=(w_in, h_in))
    ax = fig.add_axes([0, 0, 1, 1])
    ax.set_xlim(0, 100)
    ax.set_ylim(0, 100.0 * h_in / w_in)
    ax.axis("off")
    return fig, ax


def save_fig(fig, file_name: str):
    path = DIAGRAMS / file_name
    fig.savefig(path, dpi=220, facecolor="white")
    plt.close(fig)
    return path


def box(ax, x, y, w, h, text, fc="#FFFFFF", ec=C_NAVY, fs=9, bold=False,
        tc="#111111", radius=0.7, lw=1.2, ls="-", va="center"):
    """Draw a rounded rectangle with centred (multi line) text."""
    patch = FancyBboxPatch(
        (x, y), w, h,
        boxstyle="round,pad=0,rounding_size=%s" % radius,
        linewidth=lw, edgecolor=ec, facecolor=fc, linestyle=ls, zorder=2,
    )
    ax.add_patch(patch)
    ax.text(x + w / 2.0, y + h / 2.0 if va == "center" else y + h - 1.2, text,
            ha="center", va=va, fontsize=fs, color=tc, zorder=3,
            fontweight="bold" if bold else "normal", linespacing=1.45)
    return patch


def label(ax, x, y, text, fs=8.5, bold=False, color=C_NAVY, ha="left",
          va="center", italic=False, rot=0, bg=False):
    """Place a text label; bg=True paints a white plate behind it so that flow
    labels stay legible even when they sit on top of a connector line."""
    kwargs = {}
    if bg:
        kwargs["bbox"] = dict(facecolor="white", edgecolor="none", pad=0.8)
    ax.text(x, y, text, fontsize=fs, color=color, ha=ha, va=va, rotation=rot,
            fontweight="bold" if bold else "normal",
            fontstyle="italic" if italic else "normal", zorder=6,
            linespacing=1.4, **kwargs)


def arrow(ax, p1, p2, color=C_NAVY, lw=1.25, ls="-", rad=0.0, style="-|>",
          ms=11, zorder=1):
    a = FancyArrowPatch(p1, p2, arrowstyle=style, color=color, linewidth=lw,
                        linestyle=ls, mutation_scale=ms, zorder=zorder,
                        connectionstyle="arc3,rad=%s" % rad,
                        shrinkA=0.5, shrinkB=0.5)
    ax.add_patch(a)
    return a


def diamond(ax, cx, cy, w, h, text, fc=C_LAMBER, ec=C_AMBER, fs=8.5, tc="#3B2A05"):
    pts = [(cx, cy + h / 2.0), (cx + w / 2.0, cy), (cx, cy - h / 2.0), (cx - w / 2.0, cy)]
    ax.add_patch(Polygon(pts, closed=True, facecolor=fc, edgecolor=ec,
                         linewidth=1.2, zorder=2))
    ax.text(cx, cy, text, ha="center", va="center", fontsize=fs, color=tc,
            zorder=3, linespacing=1.35, fontweight="bold")


def band(ax, x, y, w, h, title, fc=C_LGREY, ec="#BFBFBF", tfs=9.0):
    """Layer band with a horizontal title strip at the top of the rectangle."""
    ax.add_patch(Rectangle((x, y), w, h, facecolor=fc, edgecolor=ec,
                           linewidth=1.0, zorder=0))
    ax.text(x + 2.0, y + h - 0.9, title, ha="left", va="top", fontsize=tfs,
            color=C_NAVY, fontweight="bold", zorder=1)


def register_figure(number: str, caption: str, file_name: str):
    FIG_INDEX.append((number, caption, file_name))


# ------------------------------------------------------------------------------
# 3.1 Overall system architecture
# ------------------------------------------------------------------------------
def draw_architecture():
    fig, ax = _canvas(9.8, 7.4)
    H = 100.0 * 7.4 / 9.8

    def layer(y, h, title, items, fc, ec, fs=8.4, box_fc="#FFFFFF"):
        band(ax, 0.5, y, 99, h, title, fc=fc, ec=ec)
        n = len(items)
        gap = 1.5
        bw = (94.0 - (n - 1) * gap) / n
        for i, txt in enumerate(items):
            x = 6.0 + i * (bw + gap)
            box(ax, x, y + 1.2, bw, h - 6.2, txt, fc=box_fc, ec=ec, fs=fs)

    # Presentation layer
    layer(H - 14.0, 13.0, "CLIENT / PRESENTATION LAYER   (Next.js 16 App Router  +  React 19  +  Tailwind CSS v4)",
          ["Landing\nPage", "Login /\nRegister", "Teacher\nDashboard", "MCQ Generator\nUI",
           "NLP Concept\nInspector", "Student Quiz\nEngine", "Results &\nAnalytics"],
          C_LBLUE, C_BLUE, fs=7.8)

    # API layer
    layer(H - 31.0, 13.0, "APPLICATION / API LAYER   (Next.js Route Handlers - server side only)",
          ["/api/auth\n(register, login,\nlogout, me)", "/api/subjects\n(create, list,\ndelete)",
           "/api/upload\n(PDF ingestion\n+ NLP)", "/api/generate\n(POST generate\nPUT regenerate)",
           "/api/mcq-sets\n(save, fetch,\ndelete)", "/api/quiz\n(submit answers,\nfetch results)"],
          C_LTEAL, C_TEAL, fs=7.4)

    # Service layer
    layer(H - 51.0, 15.5, "SERVICE LAYER   (src/services)",
          ["PDF SERVICE\nGemini Vision OCR\n(@google/genai)\nIdentity-H / CID,\nscanned, handwritten",
           "CLASSICAL NLP PIPELINE\n7 deterministic stages\n(normalisation -> sentence\nsegmentation -> tokenisation ->\nlemmatisation -> POS -> n-gram ->\nconcept scoring)",
           "AI FACTORY\nGEMINI_API_KEY ?\nGemini 2.5 Flash :\nMock AI Service\n(single interface)",
           "MCQ SERVICE\nvalidation, 4-option\nsanitisation, Jaccard\nde-duplication,\ncount top-up"],
          C_LAMBER, C_AMBER, fs=7.2)

    # Data layer
    layer(H - 70.0, 14.0, "DATA LAYER   (MongoDB  -  connection pooled singleton  src/lib/db.ts)",
          ["users\nname, email,\npasswordHash, role",
           "subjects\nname, semester,\ndepartment, userId",
           "mcq_sets\ntitle, difficulty,\nquestions[ ]",
           "quiz_attempts\nscore, percentage,\nanswers[ ]"],
          C_LGREEN, C_GREEN, fs=7.8)

    # inter layer arrows
    for x in (24, 50, 76):
        arrow(ax, (x, H - 14.0), (x, H - 18.0), color=C_BLUE)
    for x in (33, 66):
        arrow(ax, (x, H - 31.0), (x, H - 35.5), color=C_TEAL)
    for x in (33, 66):
        arrow(ax, (x, H - 51.0), (x, H - 56.0), color=C_AMBER)

    label(ax, 1.0, 2.4, "Request flow :  Browser  ->  Route Handler  ->  Service  ->  MongoDB / Gemini  ->  JSON response  ->  React state  ->  UI",
          fs=8.0, italic=True, color=C_GREY)

    register_figure("3.1", "Layered system architecture of ExamEase", "fig_3_1_system_architecture.png")
    return save_fig(fig, "fig_3_1_system_architecture.png")


# ------------------------------------------------------------------------------
# 3.2 End to end workflow (swimlane activity diagram)
# ------------------------------------------------------------------------------
def draw_workflow():
    fig, ax = _canvas(9.4, 8.2)
    H = 100.0 * 8.2 / 9.4          # ~ 87.2

    lanes = [("TEACHER", 1.5, 30.0, C_LBLUE, C_BLUE),
             ("SYSTEM  LAYER   (OCR + NLP + AI)", 32.0, 36.5, C_LTEAL, C_TEAL),
             ("STUDENT", 69.0, 29.5, C_LGREEN, C_GREEN)]
    for name, x, w, fc, ec in lanes:
        ax.add_patch(Rectangle((x, 4.0), w, H - 13.5, facecolor=fc, edgecolor=ec,
                               linewidth=1.1, alpha=0.45, zorder=0))
        label(ax, x + w / 2.0, H - 7.2, name, fs=8.8, bold=True, ha="center", color=ec)

    TX, TW = 3.5, 26.0        # teacher lane
    SX, SW = 34.0, 32.5       # system lane
    UX, UW = 70.5, 26.5       # student lane
    BH = 7.2                  # standard box height
    R = [70.0, 61.0, 52.0, 43.0, 34.0, 25.0, 16.0, 7.0]   # row baselines

    def step(x, y, w, text, fc="#FFFFFF", ec=C_NAVY, fs=7.2):
        box(ax, x, y, w, BH, text, fc=fc, ec=ec, fs=fs)
        return {"l": (x, y + BH / 2.0), "r": (x + w, y + BH / 2.0),
                "cx": x + w / 2.0, "b": (x + w / 2.0, y),
                "t": (x + w / 2.0, y + BH)}

    t1 = step(TX, R[0], TW, "Register or log in as a teacher account")
    t2 = step(TX, R[1], TW, "Create a subject (name, semester,\ndepartment)")
    t3 = step(TX, R[2], TW, "Upload lecture notes PDF or\npaste raw lecture text")
    s1 = step(SX, R[2], SW, "Gemini Vision OCR transcribes the PDF buffer\ninto clean academic text", fc=C_LTEAL, ec=C_TEAL)
    s2 = step(SX, R[3], SW, "Classical NLP pipeline runs 7 stages and returns\nscored concepts with supporting sentences", fc=C_LTEAL, ec=C_TEAL)
    t4 = step(TX, R[4], TW, "Review the NLP Concept Inspector, then set\nquestion count, difficulty and test title")
    s3 = step(SX, R[5], SW, "Gemini 2.5 Flash synthesises grounded MCQs ->\nvalidation -> Jaccard de-duplication", fc=C_LTEAL, ec=C_TEAL)
    t5 = step(TX, R[5], TW, "Edit option text or regenerate a single\nquestion in place")
    t6 = step(TX, R[6], TW, "Save the MCQ set and copy the shareable\n/quiz/{id} link")
    u1 = step(UX, R[6], UW, "Open the shared link and attempt the\ntest question by question")
    g1 = step(SX, R[7], SW, "Server grades the submission, stores the\nquiz_attempt record and computes the score", fc=C_LGREEN, ec=C_GREEN)
    u2 = step(UX, R[7], UW, "View the score card, percentage badge and\nper-question concept citations")

    def down(a, b, color=C_NAVY):
        arrow(ax, a["b"], b["t"], color=color)

    down(t1, t2); down(t2, t3)                      # register -> subject -> upload
    arrow(ax, t3["r"], s1["l"], color=C_TEAL)       # upload -> OCR
    down(s1, s2, color=C_TEAL)                      # OCR -> NLP pipeline
    arrow(ax, s2["l"], t4["r"], color=C_TEAL)       # NLP -> inspector
    arrow(ax, t4["r"], s3["l"], color=C_TEAL)       # configure -> generation
    arrow(ax, s3["l"], t5["r"], color=C_TEAL)       # generated -> edit
    down(t5, t6)                                    # edit -> save
    arrow(ax, t6["r"], u1["l"], color=C_GREEN)      # shareable link -> student
    arrow(ax, (u1["l"][0], u1["l"][1] - 1.2), (g1["r"][0] - 0.5, g1["r"][1] + 1.4),
          color=C_GREEN, rad=-0.05)                 # submit -> grading
    arrow(ax, g1["r"], u2["l"], color=C_GREEN)      # grading -> score card
    arrow(ax, u2["t"], u1["b"], color=C_GREEN, ls="--")   # score card -> re-attempt

    label(ax, 50.0, 1.6,
          "Role lanes show how one uploaded document travels from the teacher, through the deterministic NLP pipeline and the grounded AI layer, "
          "to the student attempt and back to the teacher as stored analytics.",
          fs=7.3, italic=True, ha="center", color=C_GREY)

    register_figure("3.2", "End to end activity workflow of the ExamEase platform",
                    "fig_3_2_end_to_end_workflow.png")
    return save_fig(fig, "fig_3_2_end_to_end_workflow.png")


# ------------------------------------------------------------------------------
# 3.3 Seven stage NLP pipeline
# ------------------------------------------------------------------------------
def draw_nlp_pipeline():
    fig, ax = _canvas(9.0, 8.2)
    H = 100.0 * 8.2 / 9.0

    stages = [
        ("S1", "Text normalisation and noise cleaning",
         "strip CR / TAB / multi-spaces, drop PDF stream artefacts\n(obj, endobj, stream, flatedecode, cidfont, identity-h)",
         "cleanedText  (minimum 50 readable characters)"),
        ("S2", "Sentence boundary segmentation",
         "punctuation look-ahead split, abbreviation guard (Dr., e.g., Fig.),\ndecimal protection, enumeration label removal",
         "sentences[ ]  (length > 10 characters)"),
        ("S3", "Tokenisation and stopword filtering",
         "regex  \\b[A-Za-z0-9]+(?:-[A-Za-z0-9]+)*\\b  keeps hyphenated terms;\n150+ English stopword set + PDF noise keywords",
         "Token[ ]  { text, original, position, isStopWord }"),
        ("S4", "Lemmatisation and stemming",
         "irregular lemma dictionary (matrices -> matrix, criteria -> criterion)\n+ suffix rules; Porter stemmer available as a utility",
         "normalised term forms for scoring and matching"),
        ("S5", "POS tagging and noun phrase chunking",
         "suffix based rule tagger (NN, NNS, NNP, JJ, RB, VBG, VBD);\nchunker groups  [Adjective]* [Noun]+  sequences",
         "noun phrases -> candidate concepts"),
        ("S6", "N-gram extraction (1, 2, 3-grams)",
         "frequency ranked unigrams, bigrams and trigrams;\ncombinations that begin or end with stopwords are dropped",
         "top 30 unigrams, 25 bigrams, 20 trigrams"),
        ("S7", "Concept extraction and definition scoring",
         "definition regexes (is defined as / refers to / Concept: X)\n+ importance score, capped at 10.0, top 25 concepts",
         "concepts[ ] with supporting sentences (max 4)"),
    ]

    y = H - 11.0
    h = 9.2
    dy = 10.6
    for i, (sid, title, detail, output) in enumerate(stages):
        yy = y - i * dy
        box(ax, 20.0, yy, 44.0, h, "", fc=C_LBLUE, ec=C_BLUE)
        label(ax, 22.0, yy + h - 2.4, "%s  %s" % (sid, title), fs=9.0, bold=True, color=C_NAVY)
        label(ax, 22.0, yy + 3.2, detail, fs=7.1, color="#333333")
        box(ax, 2.0, yy + 2.6, 15.0, 4.2, sid, fc=C_NAVY, ec=C_NAVY, fs=9.5, bold=True, tc="#FFFFFF")
        arrow(ax, (17.0, yy + 4.7), (20.0, yy + 4.7), color=C_BLUE)
        label(ax, 66.0, yy + h / 2.0, output, fs=7.2, italic=True, color=C_TEAL)
        arrow(ax, (64.0, yy + h / 2.0), (66.0, yy + h / 2.0), color=C_TEAL, style="->")
        if i < len(stages) - 1:
            arrow(ax, (42.0, yy), (42.0, yy - (dy - h)), color=C_NAVY)

    box(ax, 20.0, 2.0, 44.0, 6.6,
        "NLPAnalysisResult  { rawText, cleanedText, sentenceCount, tokenCount,\n"
        "uniqueTokenCount, sentences[ ], ngrams{ }, concepts[ ], definitionSentences[ ] }",
        fc=C_LGREEN, ec=C_GREEN, fs=7.6, bold=False)
    arrow(ax, (42.0, 2.0 + 6.6 + 3.0), (42.0, 2.0 + 6.6), color=C_GREEN)
    label(ax, 66.0, 5.3, "output payload of the pipeline\n(consumed by /api/upload)", fs=7.2,
          italic=True, color=C_GREEN)

    register_figure("4.1", "Seven stage deterministic NLP pre-processing pipeline",
                    "fig_4_1_nlp_pipeline_stages.png")
    return save_fig(fig, "fig_4_1_nlp_pipeline_stages.png")

# ------------------------------------------------------------------------------
# 3.4 Use case diagram
# ------------------------------------------------------------------------------
def draw_actor(ax, x, y, name, color=C_NAVY, scale=1.0):
    """Draw a UML stick figure actor with a label underneath."""
    sc = scale
    ax.add_patch(plt.Circle((x, y + 6.4 * sc), 1.9 * sc, facecolor="#FFFFFF",
                            edgecolor=color, linewidth=1.4, zorder=4))
    ax.add_line(plt.Line2D([x, x], [y + 4.5 * sc, y + 0.6 * sc], color=color,
                           linewidth=1.4, zorder=4))
    ax.add_line(plt.Line2D([x - 2.8 * sc, x + 2.8 * sc], [y + 3.2 * sc, y + 3.2 * sc],
                           color=color, linewidth=1.4, zorder=4))
    ax.add_line(plt.Line2D([x, x - 2.4 * sc], [y + 0.6 * sc, y - 2.4 * sc], color=color,
                           linewidth=1.4, zorder=4))
    ax.add_line(plt.Line2D([x, x + 2.4 * sc], [y + 0.6 * sc, y - 2.4 * sc], color=color,
                           linewidth=1.4, zorder=4))
    label(ax, x, y - 4.4 * sc, name, fs=9.0, bold=True, ha="center", color=color)


def draw_use_case():
    fig, ax = _canvas(9.4, 7.6)
    H = 100.0 * 7.6 / 9.4          # 80.9

    # ---- system boundary -----------------------------------------------------
    ax.add_patch(Rectangle((19.5, 6.0), 61.0, H - 13.0, facecolor="#FBFCFE",
                           edgecolor=C_NAVY, linewidth=1.4, zorder=0))
    label(ax, 50.0, H - 10.4, "ExamEase Platform", fs=11.0, bold=True, ha="center")
    label(ax, 50.0, H - 13.6, "<< system boundary >>", fs=7.4, ha="center",
          color=C_GREY, italic=True)

    # ---- use cases : teacher column (left) and student column (right) --------
    teacher_uc = ["Register or log in", "Manage subjects", "Upload notes PDF",
                  "Inspect NLP analysis", "Generate MCQs", "Edit or regenerate an item",
                  "Save set and share link", "View attempt analytics"]
    student_uc = ["Open the shared quiz link", "Attempt the test question by question",
                  "Submit and receive the score card",
                  "Review answers with concept citations"]

    BW, BH, GAP = 25.0, 5.2, 1.35
    top = H - 20.5

    def column(x, items, fc, ec):
        ys = []
        y = top
        for text in items:
            box(ax, x, y, BW, BH, text, fc=fc, ec=ec, fs=8.1)
            ys.append(y + BH / 2.0)
            y -= (BH + GAP)
        return ys

    t_ys = column(20.5, teacher_uc, C_LBLUE, C_BLUE)
    s_ys = column(54.5, student_uc, C_LGREEN, C_GREEN)

    # ---- orthogonal "bus" routing : actor -> bus -> use case ----------------
    def bus(actor_xy, actor_y, bus_x, ys, arrow_end_x, color):
        arrow(ax, actor_xy, (bus_x, actor_y), color=color, lw=1.0, ms=9)
        ax.add_line(plt.Line2D([bus_x, bus_x], [min(ys), max(ys)], color=color,
                               linewidth=1.0, zorder=1))
        for y in ys:
            arrow(ax, (bus_x, y), (arrow_end_x, y), color=color, lw=1.0, ms=9)

    draw_actor(ax, 9.0, H - 32.0, "TEACHER", color=C_BLUE)
    bus((12.2, H - 32.0 + 1.0), H - 32.0 + 1.0, 18.0, t_ys, 20.3, C_BLUE)

    draw_actor(ax, 92.0, H - 24.0, "STUDENT", color=C_GREEN)
    bus((88.8, H - 24.0 + 1.0), H - 24.0 + 1.0, 82.0, s_ys, 79.7, C_GREEN)

    # ---- secondary actor : the external AI service --------------------------
    AI_X, AI_Y, AI_W, AI_H = 52.0, 10.5, 29.0, 6.8
    box(ax, AI_X, AI_Y, AI_W, AI_H,
        "External AI service\n<< include >>  transcription and question synthesis",
        fc=C_LTEAL, ec=C_TEAL, fs=7.4)
    ai_top = AI_Y + AI_H
    poly_arrow(ax, [(AI_X + 6.0, ai_top), (AI_X + 6.0, t_ys[2] - BH / 2.0 - 3.2),
                    (45.7, t_ys[2] - BH / 2.0 - 3.2), (45.7, t_ys[2] - BH / 2.0)],
               color=C_TEAL, ls="--", ms=9)
    poly_arrow(ax, [(AI_X + 20.0, ai_top), (AI_X + 20.0, t_ys[4] - BH / 2.0 - 3.2),
                    (45.7, t_ys[4] - BH / 2.0 - 3.2), (45.7, t_ys[4] - BH / 2.0)],
               color=C_TEAL, ls="--", ms=9)
    label(ax, 47.2, t_ys[3], "<< include >>", fs=7.0, italic=True, color=C_TEAL)
    draw_actor(ax, 89.0, 10.5, "AI SERVICE", color=C_TEAL, scale=0.8)
    arrow(ax, (85.6, 11.5), (81.2, 13.4), color=C_TEAL, ls="--", lw=1.0, ms=9)

    label(ax, 50.0, 2.0,
          "Two human actors and one external system actor.  The AI service participates only through the transcription and "
          "synthesis use cases and never initiates an action.",
          fs=7.2, italic=True, ha="center", color=C_GREY)

    register_figure("3.3", "Use case diagram of the ExamEase platform",
                    "fig_3_3_use_case_diagram.png")
    return save_fig(fig, "fig_3_3_use_case_diagram.png")


# ------------------------------------------------------------------------------
# 3.5 Context level data flow diagram (DFD level 0)
# ------------------------------------------------------------------------------
def draw_dfd_level0():
    fig, ax = _canvas(9.2, 6.2)
    H = 100.0 * 6.2 / 9.2

    box(ax, 60.0, H - 8.0, 20.0, 11.0, "TEACHER", fc=C_LBLUE, ec=C_BLUE, fs=10.0, bold=True)
    box(ax, 60.0, 4.0, 20.0, 11.0, "STUDENT", fc=C_LGREEN, ec=C_GREEN, fs=10.0, bold=True)

    box(ax, 30.0, H * 0.30, 30.0, 20.0,
        "0\n\nExamEase\n\nNLP + AI\nAutomated Examination\nSystem",
        fc="#FFFFFF", ec=C_NAVY, fs=9.4, bold=False, lw=1.6)

    def flow(p1, p2, text, color=C_NAVY, rad=0.0, fs=7.2, off=(0, 0)):
        arrow(ax, p1, p2, color=color, rad=rad)
        mx, my = (p1[0] + p2[0]) / 2.0 + off[0], (p1[1] + p2[1]) / 2.0 + off[1]
        label(ax, mx, my, text, fs=fs, ha="center", color=color)

    flow((60.0, H - 4.5), (60.0, H * 0.30 + 20.0), "login credentials, subject details,\nPDF notes, difficulty, question count",
         color=C_BLUE, rad=0.28, off=(9.0, 1.0))
    flow((60.0, H * 0.30 + 20.0), (60.0, H - 8.0), "extracted concepts, generated MCQs,\nvalidation report",
         color=C_BLUE, rad=-0.28, off=(6.0, 0.0))
    flow((60.0, 8.0), (30.0, H * 0.30), "selected options, submit signal", color=C_GREEN, rad=-0.22, off=(4.0, -1.0))
    flow((30.0, H * 0.30), (60.0, 8.0), "question paper, timer, score card,\nconcept citations", color=C_GREEN, rad=0.22, off=(-9.0, 2.0))

    label(ax, 30.0, 3.0, "Context diagram: the platform is a single process that exchanges data with the teacher, the student and the MongoDB data store.",
          fs=7.4, italic=True, color=C_GREY)

    register_figure("3.4", "Context level data flow diagram (DFD level 0)",
                    "fig_3_4_dfd_level_0.png")
    return save_fig(fig, "fig_3_4_dfd_level_0.png")


# ------------------------------------------------------------------------------
# 3.6 Level 1 data flow diagram
# ------------------------------------------------------------------------------
def poly_arrow(ax, pts, color=C_NAVY, lw=1.15, ls="-", ms=10):
    """Orthogonal multi segment connector with an arrow head on the last leg."""
    for i in range(len(pts) - 2):
        ax.add_line(plt.Line2D([pts[i][0], pts[i + 1][0]], [pts[i][1], pts[i + 1][1]],
                               color=color, linewidth=lw, linestyle=ls, zorder=1))
    arrow(ax, pts[-2], pts[-1], color=color, lw=lw, ls=ls, ms=ms)


def draw_dfd_level1():
    fig, ax = _canvas(9.6, 8.8)
    H = 100.0 * 8.8 / 9.6          # 91.7

    # ---------------- external entities ----------------
    box(ax, 1.5, 73.0, 15.5, 11.0, "TEACHER", fc=C_LBLUE, ec=C_BLUE, fs=9.2, bold=True)
    box(ax, 1.5, 9.0, 15.5, 11.0, "STUDENT", fc=C_LGREEN, ec=C_GREEN, fs=9.2, bold=True)

    # ---------------- processes ----------------
    PX, PW, PH = 26.0, 28.0, 8.6
    ROWS = [75.5, 65.5, 55.5, 45.5, 35.5, 25.5, 15.5]   # vertical order of processes

    def proc(y, num, text, fs=7.9):
        box(ax, PX, y, PW, PH, "", fc="#FFFFFF", ec=C_NAVY, lw=1.35)
        label(ax, PX + 1.3, y + PH - 1.5, num, fs=8.6, bold=True, color=C_NAVY)
        label(ax, PX + PW / 2.0, y + PH / 2.0 - 1.3, text, fs=fs, ha="center")
        return {"l": (PX, y + PH / 2.0), "r": (PX + PW, y + PH / 2.0),
                "t": (PX + PW / 2.0, y + PH), "b": (PX + PW / 2.0, y),
                "cx": PX + PW / 2.0, "cy": y + PH / 2.0}

    p1 = proc(ROWS[0], "1.0", "Authentication and\nsession management")
    p2 = proc(ROWS[1], "2.0", "Subject and curriculum\nmanagement")
    p3 = proc(ROWS[2], "3.0", "Document ingestion and\nGemini Vision OCR")
    p4 = proc(ROWS[3], "4.0", "Classical NLP\npre-processing pipeline")
    p5 = proc(ROWS[4], "5.0", "Grounded MCQ synthesis\n(Gemini 2.5 Flash)")
    p6 = proc(ROWS[5], "6.0", "Validation and\nJaccard de-duplication")
    p7 = proc(ROWS[6], "7.0", "Test delivery, grading\nand analytics")

    # ---------------- data stores ----------------
    SX, SW, SH, TAG = 57.0, 38.0, 5.2, 3.2

    def store(y, text):
        ax.add_patch(Rectangle((SX, y), SW, SH, facecolor=C_LAMBER,
                               edgecolor=C_AMBER, linewidth=1.1, zorder=2))
        ax.add_patch(Rectangle((SX, y), TAG, SH, facecolor=C_AMBER,
                               edgecolor=C_AMBER, zorder=3))
        label(ax, SX + TAG + 1.2, y + SH / 2.0, text, fs=7.6, color="#3B2A05")
        return {"l": (SX, y + SH / 2.0), "r": (SX + SW, y + SH / 2.0),
                "t": (SX + SW / 2.0, y + SH), "b": (SX + SW / 2.0, y),
                "cx": SX + SW / 2.0, "cy": y + SH / 2.0}

    d1 = store(p1["cy"] - SH / 2.0, "D1     users          { name, email, passwordHash, role }")
    d2 = store(p2["cy"] - SH / 2.0, "D2     subjects       { name, semester, department, userId }")
    d3 = store(p3["cy"] - SH / 2.0, "D3     document text   (in-memory buffer, never written to disk)")
    d4 = store(p6["cy"] - SH / 2.0, "D4     mcq_sets       { title, difficulty, questions[ ] }")
    d5 = store(p7["cy"] - SH / 2.0, "D5     quiz_attempts  { score, percentage, answers[ ] }")

    def f(pt1, pt2, text, color=C_NAVY, fs=6.6, off=(0, 1.5), rad=0.0, ha="center"):
        arrow(ax, pt1, pt2, color=color, rad=rad, lw=1.1, ms=10)
        label(ax, (pt1[0] + pt2[0]) / 2.0 + off[0], (pt1[1] + pt2[1]) / 2.0 + off[1],
              text, fs=fs, ha=ha, color=color, bg=True)

    # ---- teacher interactions ------------------------------------------------
    f((17.0, 80.2), p1["l"], "credentials", color=C_BLUE, off=(0, 1.7))
    f((PX - 0.5, 76.5), (17.5, 74.5), "session cookie", color=C_BLUE, off=(0.5, -1.7))
    f((17.0, 74.0), p2["l"], "subject details", color=C_BLUE, off=(1.5, 1.6))
    f((17.0, 71.0), p3["l"], "PDF notes", color=C_BLUE, off=(-3.2, 2.6))

    # ---- process to store ----------------------------------------------------
    f(p1["r"], d1["l"], "user record", color=C_BLUE, off=(1.5, 1.7))
    f(p2["r"], d2["l"], "subject record", color=C_BLUE, off=(1.5, 1.7))
    f(p3["r"], d3["l"], "extracted text", color=C_TEAL, off=(1.5, 1.7))

    # ---- D3 feeds the NLP pipeline ------------------------------------------
    poly_arrow(ax, [(d3["cx"], d3["b"][1]), (d3["cx"], 51.0), (p4["r"][0], 51.0), p4["r"]],
               color=C_TEAL)
    label(ax, d3["cx"] - 1.0, 52.4, "document text", fs=6.6, color=C_TEAL, bg=True)

    # ---- internal chain ------------------------------------------------------
    f(p4["b"], p5["t"], "NLP result", color=C_TEAL, off=(5.6, 0.0), ha="left")
    f(p5["b"], p6["t"], "candidate items", color=C_TEAL, off=(6.6, 0.0), ha="left")
    f(p6["r"], d4["l"], "validated set persisted", color=C_TEAL, off=(2.0, -1.8))

    # ---- store drives test delivery -----------------------------------------
    poly_arrow(ax, [(d4["r"][0], d4["r"][1]), (97.0, d4["r"][1]), (97.0, 21.3),
                    (p7["cx"], 21.3), (p7["cx"], p7["t"][1])], color=C_GREEN)
    label(ax, 74.0, 22.8, "question paper fetched by  /quiz/{id}", fs=6.6, color=C_GREEN, bg=True)

    # ---- student interactions ------------------------------------------------
    f((17.0, 18.5), p7["l"], "answers", color=C_GREEN, off=(0, 1.7))
    f((PX - 0.6, 12.5), (17.5, 11.5), "score card", color=C_GREEN, off=(0.5, -1.7))
    f(p7["r"], d5["l"], "attempt stored", color=C_GREEN, off=(1.5, 1.7))
    f(d5["l"], (p7["r"][0], p7["cy"] + 1.6), "history", color=C_GREEN,
      off=(0.5, -1.7), rad=0.0)

    label(ax, 50.0, 4.0,
          "Level 1 DFD:  seven processes and five data stores.  Every process corresponds to one Next.js route handler supported by a pure service module; "
          "D3 exists only as an in-memory buffer and is never persisted.",
          fs=7.0, italic=True, ha="center", color=C_GREY)

    register_figure("3.5", "Level 1 data flow diagram of ExamEase", "fig_3_5_dfd_level_1.png")
    return save_fig(fig, "fig_3_5_dfd_level_1.png")


# ------------------------------------------------------------------------------
# 3.7 Level 2 DFD - NLP pre-processing pipeline
# ------------------------------------------------------------------------------
def draw_dfd_level2():
    fig, ax = _canvas(9.6, 5.4)
    H = 100.0 * 5.4 / 9.6          # 56.2

    steps = [
        ("4.1", "Normalise text and\nremove PDF noise"),
        ("4.2", "Segment into\nsentences"),
        ("4.3", "Tokenise and\nfilter stopwords"),
        ("4.4", "Lemmatise words\n(irregular + rules)"),
        ("4.5", "POS tag and\nchunk noun phrases"),
        ("4.6", "Extract 1/2/3 grams\nwith frequencies"),
        ("4.7", "Score concepts and\ndetect definitions"),
    ]
    w, gap = 12.7, 1.05
    x0 = 1.4
    by, bh = 31.0, 15.0
    for i, (num, txt) in enumerate(steps):
        x = x0 + i * (w + gap)
        box(ax, x, by, w, bh, "", fc="#FFFFFF", ec=C_NAVY, lw=1.3)
        label(ax, x + w / 2.0, by + bh - 2.6, num, fs=8.6, bold=True, ha="center", color=C_NAVY)
        label(ax, x + w / 2.0, by + 5.4, txt, fs=7.0, ha="center")
        if i < len(steps) - 1:
            arrow(ax, (x + w, by + bh / 2.0), (x + w + gap, by + bh / 2.0),
                  color=C_TEAL, ms=10)

    box(ax, 1.4, 8.0, 30.0, 10.5, "cleaned academic text\n(from the transcription service)",
        fc=C_LTEAL, ec=C_TEAL, fs=7.8)
    poly_arrow(ax, [(16.0, 18.5), (16.0, 25.5), (7.7, 25.5), (7.7, by)],
               color=C_TEAL, ms=10)

    box(ax, 68.0, 8.0, 30.5, 10.5,
        "NLPAnalysisResult\n{ sentenceCount, tokenCount, uniqueTokenCount,\nconcepts[ ], definitionSentences[ ], ngrams{ } }",
        fc=C_LGREEN, ec=C_GREEN, fs=7.2)
    arrow(ax, (83.2, by), (83.2, 18.5), color=C_GREEN, ms=10)

    box(ax, 34.0, 8.0, 31.5, 10.5,
        "process 4.0 decomposed\nall seven sub processes are deterministic :\nno network call and no language model",
        fc=C_LAMBER, ec=C_AMBER, fs=7.4)

    label(ax, 50.0, 3.4,
          "Every sub process consumes the artefact of the previous stage, which makes the whole analysis a pure function of the uploaded document.",
          fs=7.2, italic=True, ha="center", color=C_GREY)

    register_figure("3.6", "Level 2 DFD - decomposition of the NLP pre-processing process",
                    "fig_3_6_dfd_level_2_nlp.png")
    return save_fig(fig, "fig_3_6_dfd_level_2_nlp.png")


# ------------------------------------------------------------------------------
# 3.8 Class and component diagram
# ------------------------------------------------------------------------------
def draw_class_diagram():
    fig, ax = _canvas(9.6, 6.8)
    H = 100.0 * 6.8 / 9.6

    def cls(x, y, w, name, stereotype, members, ec=C_NAVY, hdr=C_LBLUE, fs=6.9):
        hh = 6.2
        rh = 2.5
        h = hh + rh * len(members)
        ax.add_patch(Rectangle((x, y), w, h, facecolor="#FFFFFF", edgecolor=ec,
                               linewidth=1.25, zorder=2))
        ax.add_patch(Rectangle((x, y + h - hh), w, hh, facecolor=hdr, edgecolor=ec,
                               linewidth=1.25, zorder=3))
        label(ax, x + w / 2.0, y + h - 2.2, name, fs=8.4, bold=True, ha="center")
        label(ax, x + w / 2.0, y + h - 4.9, stereotype, fs=6.6, ha="center",
              italic=True, color=C_GREY)
        for i, m in enumerate(members):
            yy = y + h - hh - rh * (i + 1) + rh / 2.0
            label(ax, x + 1.3, yy, m, fs=fs)
        return {"x": x, "y": y, "w": w, "h": h, "r": (x + w, y + h / 2.0),
                "l": (x, y + h / 2.0), "t": (x + w / 2.0, y + h), "b": (x + w / 2.0, y)}

    cls(2.0, H - 26.0, 26.0, "AIService", "<< interface >>",
        ["+ generateMCQs(nlpResult,", "   options): Question[ ]",
         "+ regenerateSingleQuestion(", "   nlpResult, q, difficulty):", "   Question"], hdr="#E8E8F5")
    cls(35.0, H - 26.0, 28.0, "GeminiAIService", "<< implementation >>",
        ["- apiKey : string", "- ai : GoogleGenAI",
         "+ generateMCQs()  -> prompt", "   grounded on concepts",
         "+ regenerateSingleQuestion()"], hdr=C_LAMBER, ec=C_AMBER)
    cls(70.0, H - 26.0, 28.0, "MockAIService", "<< implementation >>",
        ["+ generateMCQs()  -> distractor", "   options built from other",
         "   document concepts", "+ regenerateSingleQuestion()"], hdr=C_LGREEN, ec=C_GREEN)

    arrow(ax, (28.0, H - 20.0), (35.0, H - 20.0), color=C_NAVY, ls="--")
    label(ax, 31.5, H - 18.4, "implements", fs=6.6, ha="center", color=C_GREY)
    poly_arrow(ax, [(15.0, H - 26.0), (15.0, H - 33.5), (67.0, H - 33.5),
                    (67.0, H - 16.6), (70.0, H - 16.6)], color=C_NAVY, ls="--", ms=9)
    label(ax, 42.0, H - 35.2, "implements", fs=6.6, ha="center", color=C_GREY)

    cls(35.0, 4.0, 28.0, "MCQService", "<< module >>",
        ["+ generateAndValidateMCQs()", "+ regenerateSingleQuestion()",
         "- sanitiseOptions()", "- count top-up loop"], hdr=C_LTEAL, ec=C_TEAL)
    cls(2.0, 4.0, 26.0, "DuplicateDetector", "<< module >>",
        ["+ normalizeQuestion()", "+ calculateJaccard()", "+ removeDuplicates()"],
        hdr=C_LTEAL, ec=C_TEAL)
    cls(70.0, 4.0, 28.0, "NLPPipeline", "<< module >>",
        ["+ processDocumentText()", "- normalise / segment / tokenise",
         "- lemmatise / POS tag / n-grams", "- score concepts"], hdr=C_LTEAL, ec=C_TEAL)

    arrow(ax, (35.0, 10.0), (28.0, 10.0), color=C_NAVY)
    arrow(ax, (35.0, H - 26.0), (35.0, 4.0 + 17.5), color=C_NAVY, ls="--")
    arrow(ax, (63.0, 12.0), (70.0, 12.0), color=C_NAVY, ls="--")
    label(ax, 66.5, 13.6, "consumes", fs=6.6, ha="center", color=C_GREY)
    label(ax, 36.4, H - 30.0, "uses", fs=6.6, ha="left", color=C_GREY)

    register_figure("4.2", "Class and component diagram of the service layer",
                    "fig_4_2_class_diagram.png")
    return save_fig(fig, "fig_4_2_class_diagram.png")


# ------------------------------------------------------------------------------
# 3.9 Entity relationship diagram
# ------------------------------------------------------------------------------
def draw_er_diagram():
    fig, ax = _canvas(9.6, 7.4)
    H = 100.0 * 7.4 / 9.6

    def entity(x, y, w, title, fields, ec=C_NAVY, hdr=C_LBLUE):
        hh, rh = 3.6, 2.55
        h = hh + rh * len(fields)
        ax.add_patch(Rectangle((x, y), w, h, facecolor="#FFFFFF", edgecolor=ec,
                               linewidth=1.3, zorder=2))
        ax.add_patch(Rectangle((x, y + h - hh), w, hh, facecolor=hdr, edgecolor=ec,
                               linewidth=1.3, zorder=3))
        label(ax, x + w / 2.0, y + h - hh / 2.0, title, fs=8.6, bold=True, ha="center")
        for i, fld in enumerate(fields):
            yy = y + h - hh - rh * (i + 1) + rh / 2.0
            label(ax, x + 1.4, yy, fld, fs=7.0)
            if i < len(fields) - 1:
                ax.add_line(plt.Line2D([x, x + w], [yy - rh / 2.0, yy - rh / 2.0],
                                       color="#D8D8D8", linewidth=0.7, zorder=3))
        return {"x": x, "y": y, "w": w, "h": h,
                "l": (x, y + h / 2.0), "r": (x + w, y + h / 2.0),
                "t": (x + w / 2.0, y + h), "b": (x + w / 2.0, y)}

    users = entity(2.0, H - 34.0, 27.0, "users", [
        "_id            ObjectId", "name           string", "email          string (unique)",
        "passwordHash   string (bcrypt)", "role           TEACHER / STUDENT",
        "createdAt      date"])
    subjects = entity(66.0, H - 34.0, 32.0, "subjects", [
        "_id            ObjectId", "name           string", "semester       string",
        "department     string", "userId         ref -> users._id",
        "userName       string", "createdAt      date"])
    mcq = entity(66.0, H - 68.0, 32.0, "mcq_sets", [
        "_id            ObjectId", "title          string", "subjectId      ref -> subjects._id",
        "userId         ref -> users._id", "fileName       string",
        "questionCount  number", "difficulty     easy/medium/hard/mixed",
        "questions[ ]   embedded Question", "createdAt      date"])
    att = entity(2.0, H - 68.0, 27.0, "quiz_attempts", [
        "_id            ObjectId", "mcqSetId       ref -> mcq_sets._id",
        "subjectId      ref -> subjects._id", "userId         ref -> users._id",
        "userName       string", "score          number", "totalQuestions number",
        "percentage     number", "answers[ ]     embedded QuizAnswer",
        "completedAt    date"])

    def rel(p1, p2, text, one_side, many_side, color=C_NAVY, fs=7.2):
        arrow(ax, p1, p2, color=color, style="-|>", lw=1.1, ms=10, ls="--")
        label(ax, (p1[0] + p2[0]) / 2.0, (p1[1] + p2[1]) / 2.0 + 1.6, text,
              fs=fs, ha="center", color=color, bg=True)
        label(ax, p1[0] + (p2[0] - p1[0]) * 0.14, p1[1] + (p2[1] - p1[1]) * 0.14 + 1.5,
              one_side, fs=7.4, ha="center", color=color)
        label(ax, p1[0] + (p2[0] - p1[0]) * 0.86, p1[1] + (p2[1] - p1[1]) * 0.86 + 1.5,
              many_side, fs=7.4, ha="center", color=color)

    rel(users["r"], subjects["l"], "creates", "1", "N", color=C_BLUE)
    rel((subjects["x"] + 8.0, subjects["y"]), (mcq["x"] + 8.0, mcq["t"][1]), "contains",
        "1", "N", color=C_TEAL)
    rel(users["b"], att["t"], "attempts", "1", "N", color=C_GREEN)

    poly_arrow(ax, [(att["x"] + 8.0, att["y"]), (att["x"] + 8.0, 4.6),
                    (mcq["x"] + mcq["w"] - 8.0, 4.6), (mcq["x"] + mcq["w"] - 8.0, mcq["y"])],
               color=C_AMBER, ls="--")
    label(ax, att["x"] + 10.5, att["y"] - 5.0, "N", fs=7.4, color=C_AMBER)
    label(ax, mcq["x"] + mcq["w"] - 10.5, mcq["y"] - 5.0, "1", fs=7.4, color=C_AMBER)
    label(ax, 45.0, 5.9, "generates", fs=7.2, ha="center", color=C_AMBER, bg=True)

    box(ax, 31.0, 10.0, 30.0, 24.0,
        "Embedded documents\n\nQuestion\n id, question, options[4],\n correctAnswer, explanation,\n "
        "difficulty, type, sourceConcept\n\nQuizAnswer\n questionId, questionText,\n selectedOption, "
        "correctAnswer,\n isCorrect, explanation",
        fc="#FAFAFA", ec=C_GREY, fs=7.0)

    label(ax, 50.0, 1.4,
          "MongoDB document model: questions and answers are embedded inside their parent document, so one quiz set and one "
          "result page are each retrieved with a single read.",
          fs=7.1, italic=True, ha="center", color=C_GREY)

    register_figure("5.1", "Entity relationship diagram of the ExamEase MongoDB schema",
                    "fig_5_1_er_diagram.png")
    return save_fig(fig, "fig_5_1_er_diagram.png")


# ------------------------------------------------------------------------------
# 3.10 Sequence diagram - generation flow
# ------------------------------------------------------------------------------
def draw_sequence():
    fig, ax = _canvas(9.6, 8.6)
    H = 100.0 * 8.6 / 9.6

    actors = [
        ("Teacher\nbrowser", 9.0, C_LBLUE, C_BLUE),
        ("/api/upload", 28.0, C_LTEAL, C_TEAL),
        ("Transcription +\nNLP core", 48.0, C_LAMBER, C_AMBER),
        ("/api/generate", 68.0, C_LTEAL, C_TEAL),
        ("MongoDB\nmcq_sets", 89.0, C_LGREEN, C_GREEN),
    ]
    TOP, BOTTOM = H - 10.5, 8.0
    for name, x, fc, ec in actors:
        box(ax, x - 9.0, TOP, 18.0, 8.6, name, fc=fc, ec=ec, fs=7.5, bold=True)
        ax.add_line(plt.Line2D([x, x], [BOTTOM, TOP], color=ec, linewidth=1.0,
                               linestyle="--", zorder=1))

    def msg(y, x1, x2, text, color=C_NAVY, fs=7.0):
        arrow(ax, (x1, y), (x2, y), color=color, lw=1.1, ms=10)
        label(ax, (x1 + x2) / 2.0, y + 1.6, text, fs=fs, ha="center", color=color, bg=True)

    def self_msg(y, x, text, color=C_NAVY, fs=6.7):
        dx, dy = 8.0, 2.4
        ax.add_line(plt.Line2D([x, x + dx], [y, y], color=color, linewidth=1.0, zorder=1))
        ax.add_line(plt.Line2D([x + dx, x + dx], [y, y - dy], color=color, linewidth=1.0, zorder=1))
        arrow(ax, (x + dx, y - dy), (x, y - dy), color=color, lw=1.0, ms=9)
        label(ax, x + dx + 1.2, y - dy / 2.0, text, fs=fs, color=color, ha="left", bg=True)

    rows = [72.0, 65.0, 58.0, 51.0, 44.0, 37.0, 30.0, 23.0, 16.0]
    msg(rows[0], 9.0, 28.0, "POST multipart PDF")
    self_msg(rows[0] - 0.9, 28.0, "validate session, mime type, size")
    msg(rows[1], 28.0, 48.0, "base64 inline data (application/pdf)")
    self_msg(rows[1] - 0.9, 48.0, "transcribe -> clean academic text")
    msg(rows[2], 48.0, 28.0, "raw extracted text")
    self_msg(rows[2] - 0.9, 28.0, "run 7-stage NLP pipeline (deterministic)")
    msg(rows[3], 28.0, 9.0, "NLPAnalysisResult JSON")
    msg(rows[4], 9.0, 68.0, "POST nlpResult + questionCount + difficulty + mcqTitle")
    msg(rows[5], 68.0, 48.0, "grounded prompt (top 15 concepts + supporting sentences)")
    msg(rows[6], 48.0, 68.0, "JSON questions array")
    self_msg(rows[6] - 0.9, 68.0, "sanitise options -> Jaccard de-duplication -> top up to N")
    msg(rows[7], 68.0, 9.0, "validated Question[ ] (exactly N items)")
    msg(rows[8], 9.0, 68.0, "POST /api/mcq-sets  (teacher saves the edited set)")
    msg(13.0, 68.0, 89.0, "insertOne(mcqSet document)")
    msg(10.0, 89.0, 9.0, "mcqSetId  ->  shareable  /quiz/{id}  link", color=C_GREEN)

    label(ax, 50.0, 3.0,
          "The transcription call and the language model call are the only two network operations; everything between them is "
          "deterministic local processing.",
          fs=7.0, italic=True, ha="center", color=C_GREY)

    register_figure("4.5", "Sequence diagram of the document to MCQ generation flow",
                    "fig_4_5_sequence_generation.png")
    return save_fig(fig, "fig_4_5_sequence_generation.png")


# ------------------------------------------------------------------------------
# 3.11 AI factory fallback flow
# ------------------------------------------------------------------------------
def draw_ai_factory():
    fig, ax = _canvas(9.0, 4.4)
    H = 100.0 * 4.4 / 9.0          # 48.9

    box(ax, 2.0, H - 12.0, 23.0, 9.0, "getAIService()", fc=C_LBLUE, ec=C_BLUE,
        fs=9.4, bold=True)
    diamond(ax, 40.0, H - 7.5, 24.0, 12.5,
            "GEMINI_API_KEY\npresent and\nlength > 5 ?", fs=7.8)
    box(ax, 61.0, H - 12.0, 37.0, 9.0,
        "GeminiAIService\nmodel : gemini-2.5-flash\ntemperature 0.3 / 0.4, JSON response mode",
        fc=C_LAMBER, ec=C_AMBER, fs=7.6)
    box(ax, 61.0, 12.0, 37.0, 9.0,
        "MockAIService\nfully offline deterministic generator\nfor demonstrations, testing and key free runs",
        fc=C_LGREEN, ec=C_GREEN, fs=7.6)

    arrow(ax, (25.0, H - 7.5), (28.0, H - 7.5), color=C_NAVY)
    arrow(ax, (52.0, H - 7.5), (61.0, H - 7.5), color=C_GREEN)
    label(ax, 56.5, H - 5.4, "YES", fs=8.0, bold=True, color=C_GREEN, ha="center")
    poly_arrow(ax, [(40.0, H - 13.8), (40.0, 16.5), (61.0, 16.5)], color=C_RED)
    label(ax, 41.4, H / 2.0, "NO", fs=8.0, bold=True, color=C_RED)

    box(ax, 4.0, 2.0, 48.0, 7.0,
        "Both implementations satisfy the AIService contract, so the MCQ service\n"
        "is unaware of which provider is active.",
        fc="#F7F9FC", ec=C_NAVY, fs=7.4)
    arrow(ax, (77.0, 12.0), (77.0, 9.0), color=C_NAVY, ls="--", ms=9)
    arrow(ax, (77.0, 9.0), (52.0, 5.5), color=C_NAVY, ls="--", ms=9)
    arrow(ax, (77.0, H - 12.0), (77.0, 9.0), color=C_NAVY, ls="--", ms=9)

    register_figure("4.3", "Dual provider AI factory with graceful offline fallback",
                    "fig_4_3_ai_factory_fallback.png")
    return save_fig(fig, "fig_4_3_ai_factory_fallback.png")


# ------------------------------------------------------------------------------
# 3.12 Duplicate detection flow
# ------------------------------------------------------------------------------
def draw_dedup():
    fig, ax = _canvas(9.0, 5.0)
    H = 100.0 * 5.0 / 9.0

    top = H - 13.0
    box(ax, 1.5, top, 19.0, 11.0,
        "Candidate question\nreturned by the\nAI service", fc=C_LBLUE, ec=C_BLUE, fs=7.8)
    box(ax, 24.0, top, 22.0, 11.0,
        "normalizeQuestion()\ntokenise -> drop stopwords\n-> lemmatise", fc=C_LTEAL, ec=C_TEAL, fs=7.4)
    diamond(ax, 62.0, top + 5.5, 26.0, 15.0,
            "J(A,B) >= 0.65\nagainst any already\naccepted question ?", fs=7.4)
    box(ax, 79.0, top, 19.5, 11.0, "Discard as a\nnear duplicate", fc="#FBE9E9", ec=C_RED, fs=7.8)

    box(ax, 24.0, 4.0, 22.0, 10.0,
        "Accept the question and\nstore its token set in the\ncomparison list",
        fc=C_LGREEN, ec=C_GREEN, fs=7.4)
    box(ax, 60.0, 4.0, 38.5, 10.0,
        "Return the unique question list\n(validated set trimmed to exactly N items)",
        fc="#FFFFFF", ec=C_NAVY, fs=7.8)

    arrow(ax, (20.5, top + 5.5), (24.0, top + 5.5), color=C_NAVY)
    arrow(ax, (46.0, top + 5.5), (49.0, top + 5.5), color=C_NAVY)
    arrow(ax, (75.0, top + 5.5), (79.0, top + 5.5), color=C_RED)
    label(ax, 77.0, top + 7.3, "yes", fs=7.4, color=C_RED, ha="center")

    poly_arrow(ax, [(62.0, top), (62.0, 15.6), (35.0, 15.6), (35.0, 14.0)], color=C_GREEN)
    label(ax, 63.4, (top + 15.6) / 2.0, "no", fs=7.4, color=C_GREEN)

    arrow(ax, (46.0, 9.0), (60.0, 9.0), color=C_GREEN)
    label(ax, 53.0, 10.8, "append", fs=7.0, color=C_GREEN, ha="center", bg=True)

    poly_arrow(ax, [(24.0, 6.5), (20.0, 6.5), (20.0, top + 2.0), (21.2, top + 2.0)],
               color=C_AMBER, ls="--")
    label(ax, 21.4, (top + 6.5) / 2.0, "next candidate of the batch", fs=6.9,
          color=C_AMBER, rot=90, ha="center")

    label(ax, 50.0, 0.8,
          "The threshold is 0.65 on lemmatised token sets, so semantically repeated items are removed before the teacher sees them.",
          fs=7.0, italic=True, ha="center", color=C_GREY)

    register_figure("4.4", "Lexical near-duplicate detection using Jaccard token similarity",
                    "fig_4_4_dedup_flow.png")
    return save_fig(fig, "fig_4_4_dedup_flow.png")


# ==============================================================================
# 4. RESULT CHARTS  -  plotted from the measured NLP output in SAMPLE_NLP
# ==============================================================================

def draw_chart_concept_scores():
    names = [c[0] for c in SAMPLE_NLP["top_concepts"]][::-1]
    scores = [c[1] for c in SAMPLE_NLP["top_concepts"]][::-1]
    freq = [c[2] for c in SAMPLE_NLP["top_concepts"]][::-1]

    fig, ax = plt.subplots(figsize=(8.4, 4.6))
    bars = ax.barh(names, scores, color=C_BLUE, edgecolor=C_NAVY, height=0.62)
    for b, s, f in zip(bars, scores, freq):
        ax.text(b.get_width() + 0.08, b.get_y() + b.get_height() / 2.0,
                "%.1f   (tf = %d)" % (s, f), va="center", fontsize=7.4, color="#333333")
    ax.set_xlim(0, 12.0)
    ax.set_xlabel("Concept importance score  (capped at 10.0)", fontsize=8.6)
    ax.set_title("Top scored concepts extracted from the sample lecture note",
                 fontsize=9.6, color=C_NAVY, fontweight="bold")
    ax.tick_params(axis="y", labelsize=7.8)
    ax.tick_params(axis="x", labelsize=8.0)
    ax.grid(axis="x", linestyle=":", alpha=0.5)
    ax.set_axisbelow(True)
    for spine in ("top", "right"):
        ax.spines[spine].set_visible(False)
    fig.tight_layout()
    path = DIAGRAMS / "fig_6_2_concept_scores.png"
    fig.savefig(path, dpi=220, facecolor="white")
    plt.close(fig)

    register_figure("6.2", "Importance scores of the top ranked concepts",
                    "fig_6_2_concept_scores.png")
    return path


def draw_chart_ngrams():
    fig, axes = plt.subplots(1, 3, figsize=(10.2, 3.9))
    data = [("Unigrams", SAMPLE_NLP["top_unigrams"], C_BLUE),
            ("Bigrams", SAMPLE_NLP["top_bigrams"], C_TEAL),
            ("Trigrams", SAMPLE_NLP["top_trigrams"], C_AMBER)]
    for ax, (title, items, color) in zip(axes, data):
        items = items[::-1]
        labels = [i[0] for i in items]
        counts = [i[1] for i in items]
        ax.barh(labels, counts, color=color, edgecolor="#333333", height=0.6)
        for y, c in enumerate(counts):
            ax.text(c + 0.08, y, str(c), va="center", fontsize=7.0, color="#333333")
        ax.set_title(title, fontsize=9.2, color=C_NAVY, fontweight="bold")
        ax.tick_params(axis="y", labelsize=7.2)
        ax.tick_params(axis="x", labelsize=7.4)
        ax.set_xlim(0, max(counts) + 1.6)
        ax.grid(axis="x", linestyle=":", alpha=0.5)
        ax.set_axisbelow(True)
        for spine in ("top", "right"):
            ax.spines[spine].set_visible(False)
    fig.suptitle("Frequency ranked n-grams of the sample document",
                 fontsize=10.0, color=C_NAVY, fontweight="bold")
    fig.tight_layout(rect=(0, 0, 1, 0.94))
    path = DIAGRAMS / "fig_6_3_ngram_frequency.png"
    fig.savefig(path, dpi=220, facecolor="white")
    plt.close(fig)
    register_figure("6.3", "Unigram, bigram and trigram frequency distribution",
                    "fig_6_3_ngram_frequency.png")
    return path


def draw_chart_pipeline_stats():
    labels = ["Sentences", "Tokens", "Unique tokens", "Definition sentences",
              "Concepts retained"]
    values = [SAMPLE_NLP["sentence_count"], SAMPLE_NLP["token_count"],
              SAMPLE_NLP["unique_token_count"], SAMPLE_NLP["definition_count"],
              SAMPLE_NLP["concept_count_extracted"]]
    fig, ax = plt.subplots(figsize=(8.4, 3.9))
    bars = ax.bar(labels, values, color=[C_BLUE, C_NAVY, C_TEAL, C_AMBER, C_GREEN],
                  edgecolor="#333333", width=0.56)
    for b, v in zip(bars, values):
        ax.text(b.get_x() + b.get_width() / 2.0, v + 12, str(v), ha="center",
                fontsize=8.4, fontweight="bold", color="#222222")
    ax.set_ylim(0, max(values) * 1.16)
    ax.set_ylabel("Count", fontsize=8.6)
    ax.set_title("Measured output of the NLP pipeline on the sample document",
                 fontsize=9.6, color=C_NAVY, fontweight="bold")
    ax.tick_params(axis="x", labelsize=7.8)
    ax.grid(axis="y", linestyle=":", alpha=0.5)
    ax.set_axisbelow(True)
    for spine in ("top", "right"):
        ax.spines[spine].set_visible(False)
    fig.tight_layout()
    path = DIAGRAMS / "fig_6_1_pipeline_statistics.png"
    fig.savefig(path, dpi=220, facecolor="white")
    plt.close(fig)
    register_figure("6.1", "Measured NLP pipeline statistics for the sample document",
                    "fig_6_1_pipeline_statistics.png")
    return path


def draw_chart_jaccard():
    pairs = [("P1", 0.75), ("P2", 0.875), ("P3", 0.6667), ("P4", 0.5), ("P5", 0.0)]
    fig, ax = plt.subplots(figsize=(8.0, 3.6))
    colors = [C_RED if v >= 0.65 else C_GREEN for _, v in pairs]
    bars = ax.bar([p[0] for p in pairs], [p[1] for p in pairs], color=colors,
                  edgecolor="#333333", width=0.5)
    for b, (_, v) in zip(bars, pairs):
        ax.text(b.get_x() + b.get_width() / 2.0, v + 0.02, "%.4f" % v, ha="center",
                fontsize=8.2, fontweight="bold", color="#222222")
    ax.axhline(0.65, color=C_NAVY, linestyle="--", linewidth=1.3)
    ax.text(4.42, 0.675, "threshold 0.65", fontsize=7.8, color=C_NAVY, ha="right")
    ax.set_ylim(0, 1.08)
    ax.set_ylabel("Jaccard similarity", fontsize=8.6)
    ax.set_xlabel("Question pairs from Table 6.3", fontsize=8.6)
    ax.set_title("Near duplicate detection: similarity of candidate question pairs",
                 fontsize=9.6, color=C_NAVY, fontweight="bold")
    ax.grid(axis="y", linestyle=":", alpha=0.5)
    ax.set_axisbelow(True)
    for spine in ("top", "right"):
        ax.spines[spine].set_visible(False)
    fig.tight_layout()
    path = DIAGRAMS / "fig_6_4_jaccard_threshold.png"
    fig.savefig(path, dpi=220, facecolor="white")
    plt.close(fig)
    register_figure("6.4", "Jaccard similarity of candidate question pairs against the 0.65 threshold",
                    "fig_6_4_jaccard_threshold.png")
    return path


def draw_chart_longitudinal():
    """Measured growth of blue print concept coverage as the document is enlarged.

    The sample note was progressively truncated to its first k paragraphs and the
    real pipeline was re-run on every prefix.  Ten concepts that a teacher would
    expect from the unit form the blue print; the chart records how many of them
    the analysis had discovered at each document size.
    """
    tokens = [49, 93, 154, 208, 301, 391, 486, 590, 699, 881]
    matched = [0, 0, 1, 2, 3, 5, 7, 8, 9, 9]
    coverage = [m * 10.0 for m in matched]

    fig, ax = plt.subplots(figsize=(8.4, 3.8))
    ax.plot(tokens, coverage, marker="o", color=C_BLUE, linewidth=2.0, markersize=6,
            markerfacecolor="white", markeredgewidth=1.8)
    for i, (x, c, m) in enumerate(zip(tokens, coverage, matched)):
        ax.text(x, c + 4.0, "%d/10" % m, ha="center", fontsize=7.4, color=C_NAVY,
                fontweight="bold")
    ax.set_xscale("log", base=2)
    ax.set_xticks(tokens)
    ax.set_xticklabels(["{:,}".format(t) for t in tokens], fontsize=7.2, rotation=30)
    ax.set_xlabel("Size of the analysed document (tokens, logarithmic scale)", fontsize=8.6)
    ax.set_ylabel("Blueprint concept coverage (%)", fontsize=8.6)
    ax.set_ylim(-6, 112)
    ax.set_title("Measured growth of concept coverage as the document is enlarged",
                 fontsize=9.6, color=C_NAVY, fontweight="bold")
    ax.grid(linestyle=":", alpha=0.55)
    ax.set_axisbelow(True)
    for spine in ("top", "right"):
        ax.spines[spine].set_visible(False)
    fig.tight_layout()
    path = DIAGRAMS / "fig_6_5_coverage_growth.png"
    fig.savefig(path, dpi=220, facecolor="white")
    plt.close(fig)
    register_figure("6.5", "Measured growth of blueprint concept coverage with the size of the source document",
                    "fig_6_5_coverage_growth.png")
    return path


SCREENSHOT_PLAN = [
    ("fig_8_1.png",   "8.1",  "Landing page of ExamEase",
     "Open the home page of the running application at http://localhost:3000"),
    ("fig_8_2.png",   "8.2",  "Teacher registration screen",
     "Open /register and fill the form with a sample teacher account"),
    ("fig_8_3.png",   "8.3",  "Teacher login screen with validation",
     "Open /login and enter the credentials created in the previous step"),
    ("fig_8_4.png",   "8.4",  "Teacher dashboard with class cards and recent activity",
     "Log in and capture /dashboard after creating at least one subject"),
    ("fig_8_5.png",   "8.5",  "Subject management screen",
     "Open /subjects after adding two subjects with semester and department"),
    ("fig_8_6.png",   "8.6",  "MCQ generator step 1 : upload and configuration form",
     "Open /generate and fill subject, title, PDF, question count and difficulty"),
    ("fig_8_7.png",   "8.7",  "NLP Concept Inspector : concepts and statistics tab",
     "After the PDF is processed, open the Concepts tab of the inspector"),
    ("fig_8_8.png",   "8.8",  "NLP Concept Inspector : definitions and n-gram tabs",
     "Open the Definitions and N-grams tabs of the same inspector panel"),
    ("fig_8_9.png",   "8.9",  "MCQ generator step 2 : editable question preview cards",
     "Capture the review screen that lists the generated questions"),
    ("fig_8_10.png",  "8.10", "In line regeneration of a single question",
     "Click Regenerate on one card and capture the refreshed question"),
    ("fig_8_11.png",  "8.11", "Saved MCQ set and the copyable shareable link",
     "Save the set and capture the confirmation card with the /quiz/{id} link"),
    ("fig_8_12.png",  "8.12", "Student test taking engine : question view",
     "Open the shareable link in a new incognito window and capture a question"),
    ("fig_8_13.png",  "8.13", "Question navigator drawer of the quiz engine",
     "Open the navigator panel that marks answered and unanswered questions"),
    ("fig_8_14.png",  "8.14", "Automated score card after submission",
     "Submit the test and capture the score, percentage badge and counts"),
    ("fig_8_15.png",  "8.15", "Detailed question wise review with concept citations",
     "Scroll the results page to the detailed review section"),
    ("fig_8_16.png",  "8.16", "MongoDB Compass view of the four collections",
     "Open MongoDB Compass and capture the examease database collections"),
]


def make_screenshot_placeholder(file_name, fig_num, caption, hint,
                                width_in=6.4, height_in=3.9):
    """Draw a neat placeholder frame that tells the student what to capture."""
    fig, ax = _canvas(width_in, height_in)
    H = 100.0 * height_in / width_in

    ax.add_patch(FancyBboxPatch((1.0, 1.0), 98.0, H - 2.0,
                                boxstyle="round,pad=0,rounding_size=1.2",
                                facecolor="#F7F9FC", edgecolor=C_BLUE,
                                linewidth=1.4, linestyle=(0, (6, 4)), zorder=1))
    # window chrome
    ax.add_patch(Rectangle((1.0, H - 7.0), 98.0, 6.0, facecolor="#E8EEF7",
                           edgecolor=C_BLUE, linewidth=1.1, zorder=2))
    for i, c in enumerate(["#E76F6F", "#F2C14E", "#5FB878"]):
        ax.add_patch(plt.Circle((3.6 + i * 2.6, H - 4.0), 0.75, facecolor=c,
                                edgecolor="none", zorder=3))
    label(ax, 88.0, H - 4.0, "SCREENSHOT PLACEHOLDER", fs=7.6, bold=True,
          ha="right", color=C_NAVY)

    label(ax, 50.0, H / 2.0 + 4.0, "Insert your screenshot here", fs=13.0,
          bold=True, ha="center", color="#9AA7B8")
    label(ax, 50.0, H / 2.0 - 1.6, "Figure %s  -  %s" % (fig_num, caption),
          fs=9.0, ha="center", color=C_BLUE)
    label(ax, 50.0, H / 2.0 - 7.4, "How to capture :  %s" % hint, fs=7.6,
          ha="center", color=C_GREY)
    label(ax, 50.0, 4.4,
          "In MS Word :  right click this image  ->  Change Picture  ->  This Device  ->  select your PNG file\n"
          "(the frame, size and alignment are preserved automatically)",
          fs=7.4, ha="center", color=C_GREY)
    path = SCREENSHOTS / file_name
    fig.savefig(path, dpi=200, facecolor="white")
    plt.close(fig)
    return path


def build_screenshot_assets():
    """Create placeholders (or keep real screenshots if the user supplied them)."""
    created, real = [], []
    guide = ["# Screenshot guide - ExamEase final report", "",
             "Every figure below appears in Chapter 8 (System Implementation and Testing) of the report.",
             "Put a real PNG with the exact file name into `assets/screenshots/` and re-run the",
             "generator script, or simply replace the placeholder inside MS Word.", "",
             "| Figure | File name | What to capture | How to capture it |",
             "| :--- | :--- | :--- | :--- |"]
    for file_name, fig_num, caption, hint in SCREENSHOT_PLAN:
        target = SCREENSHOTS / file_name
        if not target.exists() and CONFIG["include_screenshot_placeholders"]:
            make_screenshot_placeholder(file_name, fig_num, caption, hint)
            created.append(file_name)
        else:
            real.append(file_name)
        guide.append("| %s | `%s` | %s | %s |" % (fig_num, file_name, caption, hint))
    for file_name, fig_num, caption, _hint in SCREENSHOT_PLAN:
        register_figure(fig_num, caption, file_name)
    guide += ["", "Placeholders generated : %d" % len(created),
              "Real screenshots found : %d" % len(real), ""]
    (ASSETS / "SCREENSHOT_GUIDE.md").write_text("\n".join(guide), encoding="utf-8")
    return created, real

# ==============================================================================
# 6. WORD HELPERS  -  page setup, styles, captions, tables, code, fields
# ==============================================================================

BODY_FONT = "Times New Roman"
CODE_FONT = "Consolas"
TBL_STYLE = "Table Grid"


def _set_run_font(run, name=None, size=None, bold=None, italic=None, color=None,
                  all_caps=False):
    run.font.name = name or BODY_FONT
    rpr = run._element.get_or_add_rPr()
    rf = rpr.find(qn("w:rFonts"))
    if rf is None:
        rf = OxmlElement("w:rFonts")
        rpr.append(rf)
    for attr in ("w:ascii", "w:hAnsi", "w:cs", "w:eastAsia"):
        rf.set(qn(attr), name or BODY_FONT)
    if size:
        run.font.size = Pt(size)
    if bold is not None:
        run.font.bold = bold
    if italic is not None:
        run.font.italic = italic
    if color:
        run.font.color.rgb = RGBColor.from_string(color)
    if all_caps:
        run.font.all_caps = True
    return run


def shade(element, fill):
    """Apply background shading to a paragraph or table cell element."""
    pr = element.get_or_add_pPr() if hasattr(element, "get_or_add_pPr") else element
    shd = OxmlElement("w:shd")
    shd.set(qn("w:val"), "clear")
    shd.set(qn("w:color"), "auto")
    shd.set(qn("w:fill"), fill)
    pr.append(shd)


def shade_cell(cell, fill):
    tcPr = cell._tc.get_or_add_tcPr()
    shd = OxmlElement("w:shd")
    shd.set(qn("w:val"), "clear")
    shd.set(qn("w:color"), "auto")
    shd.set(qn("w:fill"), fill)
    tcPr.append(shd)


def para_borders(p, top=False, bottom=False, left=False, right=False,
                 color="9DB2CE", size=6):
    pPr = p._p.get_or_add_pPr()
    borders = OxmlElement("w:pBdr")
    for side, flag in (("top", top), ("left", left), ("bottom", bottom), ("right", right)):
        if flag:
            el = OxmlElement("w:%s" % side)
            el.set(qn("w:val"), "single")
            el.set(qn("w:sz"), str(size))
            el.set(qn("w:space"), "3")
            el.set(qn("w:color"), color)
            borders.append(el)
    pPr.append(borders)


def repeat_header_row(row):
    trPr = row._tr.get_or_add_trPr()
    el = OxmlElement("w:tblHeader")
    el.set(qn("w:val"), "true")
    trPr.append(el)


def add_field(paragraph, instruction):
    """Insert a Word field, e.g. PAGE, NUMPAGES, TOC, SEQ."""
    run = paragraph.add_run()
    fld_begin = OxmlElement("w:fldChar")
    fld_begin.set(qn("w:fldCharType"), "begin")
    instr = OxmlElement("w:instrText")
    instr.set(qn("xml:space"), "preserve")
    instr.text = instruction
    fld_sep = OxmlElement("w:fldChar")
    fld_sep.set(qn("w:fldCharType"), "separate")
    fld_text = OxmlElement("w:t")
    fld_text.text = "Right click and choose 'Update Field' to build this list."
    fld_end = OxmlElement("w:fldChar")
    fld_end.set(qn("w:fldCharType"), "end")
    run._r.append(fld_begin)
    run._r.append(instr)
    run._r.append(fld_sep)
    run._r.append(fld_text)
    run._r.append(fld_end)
    return run


def enable_auto_update_fields(doc):
    """Tell Word to refresh TOC / LOF / LOT / page numbers when the file opens."""
    settings = doc.settings.element
    uf = OxmlElement("w:updateFields")
    uf.set(qn("w:val"), "true")
    settings.append(uf)


def new_document():
    doc = Document()

    # ---- page geometry -------------------------------------------------------
    for section in doc.sections:
        section.page_width = Cm(21.0)
        section.page_height = Cm(29.7)
        section.top_margin = Cm(2.54)
        section.bottom_margin = Cm(2.54)
        section.left_margin = Cm(3.5)
        section.right_margin = Cm(2.0)

    # ---- base styles ---------------------------------------------------------
    normal = doc.styles["Normal"]
    normal.font.name = BODY_FONT
    normal.font.size = Pt(12)
    normal.paragraph_format.line_spacing = 1.5
    normal.paragraph_format.space_after = Pt(6)
    normal.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    rpr = normal.element.get_or_add_rPr()
    rf = OxmlElement("w:rFonts")
    for attr in ("w:ascii", "w:hAnsi", "w:cs", "w:eastAsia"):
        rf.set(qn(attr), BODY_FONT)
    rpr.append(rf)

    heading_specs = {
        "Heading 1": (16, "1F3864", 18, 10),
        "Heading 2": (14, "1F3864", 14, 8),
        "Heading 3": (13, "2E75B6", 12, 6),
        "Heading 4": (12, "2E75B6", 10, 6),
    }
    for style_name, (size, color, before, after) in heading_specs.items():
        st = doc.styles[style_name]
        st.font.name = BODY_FONT
        st.font.size = Pt(size)
        st.font.bold = True
        st.font.color.rgb = RGBColor.from_string(color)
        st.paragraph_format.space_before = Pt(before)
        st.paragraph_format.space_after = Pt(after)
        st.paragraph_format.keep_with_next = True
        st.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.LEFT
        st.paragraph_format.line_spacing = 1.25

    # caption style used by the List of Figures / List of Tables fields
    if "Caption" not in [s.name for s in doc.styles]:
        doc.styles.add_style("Caption", WD_STYLE_TYPE.PARAGRAPH)
    cap = doc.styles["Caption"]
    cap.font.name = BODY_FONT
    cap.font.size = Pt(10.5)
    cap.font.bold = False
    cap.font.italic = False
    cap.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.CENTER
    cap.paragraph_format.space_before = Pt(6)
    cap.paragraph_format.space_after = Pt(12)
    cap.paragraph_format.line_spacing = 1.0

    if "TableCaption" not in [s.name for s in doc.styles]:
        st = doc.styles.add_style("TableCaption", WD_STYLE_TYPE.PARAGRAPH)
    tcap = doc.styles["TableCaption"]
    tcap.font.name = BODY_FONT
    tcap.font.size = Pt(10.5)
    tcap.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.CENTER
    tcap.paragraph_format.space_before = Pt(10)
    tcap.paragraph_format.space_after = Pt(6)
    tcap.paragraph_format.line_spacing = 1.0

    if "CodeBlock" not in [s.name for s in doc.styles]:
        doc.styles.add_style("CodeBlock", WD_STYLE_TYPE.PARAGRAPH)
    code = doc.styles["CodeBlock"]
    code.font.name = CODE_FONT
    code.font.size = Pt(8.5)
    code.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.LEFT
    code.paragraph_format.line_spacing = 1.0
    code.paragraph_format.space_after = Pt(0)
    code.paragraph_format.space_before = Pt(0)
    code.paragraph_format.left_indent = Cm(0.4)

    return doc


# --------------------------------- writers ------------------------------------

def H(doc, level, text, page_break=False):
    if page_break:
        doc.add_page_break()
    return doc.add_heading(text, level=level)


def P(doc, text, justify=True, size=12, bold=False, italic=False, space_after=6,
      color=None, align=None):
    p = doc.add_paragraph()
    p.paragraph_format.space_after = Pt(space_after)
    p.paragraph_format.alignment = (
        align if align is not None
        else (WD_ALIGN_PARAGRAPH.JUSTIFY if justify else WD_ALIGN_PARAGRAPH.LEFT))
    _set_run_font(p.add_run(text), size=size, bold=bold, italic=italic, color=color)
    return p


def rich(doc, parts, justify=True, size=12, space_after=6, align=None, indent=None):
    """parts = [(text, {'bold':True,'italic':True,'color':'RRGGBB'}), ...]"""
    p = doc.add_paragraph()
    p.paragraph_format.space_after = Pt(space_after)
    p.paragraph_format.alignment = (
        align if align is not None
        else (WD_ALIGN_PARAGRAPH.JUSTIFY if justify else WD_ALIGN_PARAGRAPH.LEFT))
    if indent:
        p.paragraph_format.left_indent = Cm(indent)
    for text, fmt in parts:
        _set_run_font(p.add_run(text), size=fmt.get("size", size),
                      bold=fmt.get("bold"), italic=fmt.get("italic"),
                      color=fmt.get("color"))
    return p


def BULLETS(doc, items, marker="\u2022", indent=0.75, size=12, numbered=False):
    for i, item in enumerate(items):
        p = doc.add_paragraph()
        p.paragraph_format.left_indent = Cm(indent + 0.5)
        p.paragraph_format.first_line_indent = Cm(-0.5)
        p.paragraph_format.space_after = Pt(3)
        p.paragraph_format.line_spacing = 1.4
        p.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
        pre = ("%d.  " % (i + 1)) if numbered else (marker + "  ")
        _set_run_font(p.add_run(pre), size=size, bold=numbered)
        if isinstance(item, tuple):
            _set_run_font(p.add_run(item[0]), size=size, bold=True)
            _set_run_font(p.add_run(item[1]), size=size)
        else:
            _set_run_font(p.add_run(item), size=size)


def CODE(doc, code_text, caption=None, lang=None):
    p = doc.add_paragraph(style="CodeBlock")
    shade(p._p, "F4F6F8")
    para_borders(p, top=True, bottom=True, left=True, right=True, color="C9D3DF", size=4)
    lines = code_text.strip("\n").split("\n")
    for i, line in enumerate(lines):
        if i:
            p.add_run().add_break()
        _set_run_font(p.add_run(line.rstrip()), name=CODE_FONT, size=8.5)
    if caption:
        CAPTION(doc, caption, kind="code")
    return p


def FIGURE(doc, path, number, caption, width_cm=15.5, add_caption=True):
    p = doc.add_paragraph()
    p.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.space_before = Pt(8)
    p.paragraph_format.space_after = Pt(2)
    p.paragraph_format.keep_with_next = True
    if Path(path).exists():
        p.add_run().add_picture(str(path), width=Cm(width_cm))
    else:
        _set_run_font(p.add_run("[ missing image: %s ]" % Path(path).name), italic=True)
    if add_caption:
        CAPTION(doc, "Figure %s : %s" % (number, caption), kind="figure")
    return p


def CAPTION(doc, text, kind="figure"):
    style = "Caption" if kind == "figure" else "TableCaption"
    if kind == "code":
        p = doc.add_paragraph()
        p.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.LEFT
        p.paragraph_format.space_after = Pt(10)
        _set_run_font(p.add_run(text), size=9.5, italic=True, color="5A5A5A")
        return p
    p = doc.add_paragraph(style=style)
    style_obj = doc.styles[style]
    style_obj.element.get_or_add_rPr()
    p.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.CENTER
    _set_run_font(p.add_run(text), size=10.5,
                  italic=(kind == "table"))
    return p


def TABLE(doc, number, caption, headers, rows, widths=None, font_size=9.5,
          header_fill="DEEAF6", align_center_cols=None):
    if caption:
        CAPTION(doc, "Table %s : %s" % (number, caption), kind="table")
    table = doc.add_table(rows=1, cols=len(headers))
    table.style = TBL_STYLE
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    table.autofit = True

    hdr = table.rows[0]
    repeat_header_row(hdr)
    for i, h in enumerate(headers):
        cell = hdr.cells[i]
        cell.text = ""
        p = cell.paragraphs[0]
        p.paragraph_format.space_after = Pt(2)
        p.paragraph_format.line_spacing = 1.0
        p.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.CENTER
        _set_run_font(p.add_run(str(h)), size=font_size, bold=True)
        shade_cell(cell, header_fill)

    for r_i, row in enumerate(rows):
        cells = table.add_row().cells
        for c_i, val in enumerate(row):
            cells[c_i].text = ""
            p = cells[c_i].paragraphs[0]
            p.paragraph_format.space_after = Pt(2)
            p.paragraph_format.line_spacing = 1.15
            p.paragraph_format.alignment = (
                WD_ALIGN_PARAGRAPH.CENTER
                if (align_center_cols and c_i in align_center_cols)
                else WD_ALIGN_PARAGRAPH.LEFT)
            _set_run_font(p.add_run(str(val)), size=font_size)
        if r_i % 2 == 1:
            for c in cells:
                shade_cell(c, "F7F9FC")

    if widths:
        total = float(sum(widths))
        scale = (15.4 / total) if total > 15.4 else 1.0
        for row in table.rows:
            for i, w in enumerate(widths):
                row.cells[i].width = Cm(w * scale)
    doc.add_paragraph().paragraph_format.space_after = Pt(2)
    return table


def add_footer_page_numbers(section, fmt="decimal", start=1, show=True):
    """Set page numbering format for a section and draw the centred number."""
    sectPr = section._sectPr
    for old in sectPr.findall(qn("w:pgNumType")):
        sectPr.remove(old)
    pg = OxmlElement("w:pgNumType")
    pg.set(qn("w:fmt"), fmt)
    pg.set(qn("w:start"), str(start))
    sectPr.append(pg)

    footer = section.footer
    footer.is_linked_to_previous = False
    p = footer.paragraphs[0] if footer.paragraphs else footer.add_paragraph()
    p.text = ""
    p.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.CENTER
    if show:
        add_field(p, " PAGE ")
        for run in p.runs:
            _set_run_font(run, size=10.5)


def add_header_text(section, text):
    section.header.is_linked_to_previous = False
    p = section.header.paragraphs[0] if section.header.paragraphs else section.header.add_paragraph()
    p.text = ""
    p.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.CENTER
    para_borders(p, bottom=True, color="BFBFBF", size=4)
    _set_run_font(p.add_run(text), size=9.5, italic=True, color="5A5A5A")


def blank(doc, n=1, size=12):
    for _ in range(n):
        p = doc.add_paragraph()
        p.paragraph_format.space_after = Pt(0)
        _set_run_font(p.add_run(""), size=size)


def TOC_FIELD(doc, instruction, note):
    p = doc.add_paragraph()
    p.paragraph_format.space_after = Pt(4)
    add_field(p, instruction)
    for run in p.runs:
        _set_run_font(run, size=11)
    note_p = doc.add_paragraph()
    note_p.paragraph_format.space_after = Pt(10)
    _set_run_font(note_p.add_run(note), size=9, italic=True, color="808080")

# ==============================================================================
# 7. FRONT MATTER
# ==============================================================================

def centred(doc, text, size=12, bold=False, italic=False, space_after=6, caps=False):
    p = doc.add_paragraph()
    p.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.space_after = Pt(space_after)
    _set_run_font(p.add_run(text), size=size, bold=bold, italic=italic, all_caps=caps)
    return p


def build_title_page(doc):
    c = CONFIG
    arrow_line = "\u2500" * 34

    blank(doc, 1)
    centred(doc, "A PROJECT REPORT ON", size=13, bold=True, space_after=14)
    centred(doc, c["project_title"].upper(), size=30, bold=True, space_after=6)
    centred(doc, "\u2014  " + c["project_subtitle"] + "  \u2014", size=13, italic=True, space_after=16)
    centred(doc, "Submitted in partial fulfilment of the requirements for the award of the degree of",
            size=11.5, space_after=6)
    centred(doc, c["degree"].upper(), size=13, bold=True, space_after=4)
    centred(doc, "in", size=11.5, space_after=4)
    centred(doc, c["branch"].upper(), size=13, bold=True, space_after=18)
    centred(doc, "Submitted by", size=11.5, space_after=6)
    centred(doc, c["student_name"].upper(), size=16, bold=True, space_after=4)
    centred(doc, "Roll No. : %s" % c["roll_number"], size=11.5, space_after=4)
    centred(doc, c["class"], size=11.5, space_after=16)
    centred(doc, "Under the guidance of", size=11.5, space_after=6)
    centred(doc, c["guide_name"].upper(), size=14, bold=True, space_after=4)
    centred(doc, c["guide_designation"], size=10.5, italic=True, space_after=18)
    centred(doc, c["department"].upper(), size=12.5, bold=True, space_after=4)
    centred(doc, c["college_name"], size=13, bold=True, space_after=4)
    centred(doc, c["university_name"], size=11.5, italic=True, space_after=14)
    centred(doc, c["academic_year"], size=12, bold=True, space_after=4)
    centred(doc, c["month_year"], size=11.5, italic=True, space_after=0)


def build_certificate(doc):
    c = CONFIG
    doc.add_page_break()
    centred(doc, c["college_name"], size=14, bold=True, space_after=2)
    centred(doc, c["department"], size=12, bold=True, space_after=2)
    centred(doc, "(Affiliated to %s)" % c["university_name"], size=10.5, italic=True, space_after=18)
    p = doc.add_paragraph()
    p.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.CENTER
    _set_run_font(p.add_run("CERTIFICATE"), size=18, bold=True)
    para_borders(p, bottom=True, color="1F3864", size=8)
    blank(doc, 1)

    P(doc, "This is to certify that the project entitled \u201c%s \u2014 %s\u201d is a "
           "bonafide record of the work carried out by %s (Roll No. %s), a student of the "
           "final year of the %s programme in %s, during the academic year %s."
           % (c["project_title"], c["project_subtitle"], c["student_name"],
              c["roll_number"], c["degree"], c["branch"], c["academic_year"]))
    P(doc, "The work presented in this report has been carried out under my supervision in the "
           "%s of this institute. The candidate has completed the design, implementation and "
           "testing of the system described here, and the report has been prepared in accordance "
           "with the guidelines prescribed by %s." % (c["department"], c["university_name"]))
    P(doc, "To the best of my knowledge, the content of this report has not been submitted earlier, "
           "either in part or in full, to any other institute or university for the award of any "
           "degree or diploma.")
    blank(doc, 2)

    tbl = doc.add_table(rows=1, cols=3)
    tbl.autofit = True
    labels = [("Signature of the Guide\n\n%s\n%s" % (c["guide_name"], c["guide_designation"]),
               "Signature of the Head of Department\n\n%s\n%s" % (c["hod_name"], c["hod_designation"]),
               "Signature of the Principal\n\n%s\n%s" % (c["principal_name"], c["principal_designation"]))]
    for i, text in enumerate(labels[0]):
        cell = tbl.rows[0].cells[i]
        cell.text = ""
        p = cell.paragraphs[0]
        p.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p.paragraph_format.line_spacing = 1.2
        _set_run_font(p.add_run(text), size=10)
    blank(doc, 2)
    centred(doc, "External Examiner :  ______________________________", size=11, space_after=4)
    centred(doc, "Date of Examination :  ______________________________", size=11, space_after=4)
    centred(doc, "Place : %s" % c["place"], size=11, space_after=0)


def build_declaration(doc):
    c = CONFIG
    doc.add_page_break()
    centred(doc, "DECLARATION", size=18, bold=True, space_after=18)
    P(doc, "I hereby declare that the project report entitled \u201c%s \u2014 %s\u201d "
           "submitted by me to the %s, %s, is a record of original work carried out by me under "
           "the guidance of %s. All the sources of information, algorithms, libraries and research "
           "papers that have been referred to during this work are duly acknowledged in the "
           "references section of this report."
           % (c["project_title"], c["project_subtitle"], c["department"],
              c["college_name"], c["guide_name"]))
    P(doc, "I further declare that the implementation presented in this report \u2014 namely the "
           "deterministic natural language processing pipeline, the grounded question generation "
           "layer, the validation and de-duplication logic and the web application that integrates "
           "them \u2014 has been developed by me and that the report is free of any form of "
           "plagiarism.")
    blank(doc, 3)
    P(doc, "%s\nRoll No. : %s\n%s" % (c["student_name"], c["roll_number"], c["class"]),
      justify=False)
    blank(doc, 1)
    P(doc, "Date : ______________\nPlace : %s" % c["place"], justify=False)


def build_acknowledgement(doc):
    c = CONFIG
    doc.add_page_break()
    centred(doc, "ACKNOWLEDGEMENT", size=18, bold=True, space_after=18)
    P(doc, "It gives me immense pleasure to present the final year project report on "
           "\u201c%s\u201d. The journey from understanding the problem of automated question paper "
           "generation to building a working natural language processing pipeline and a complete "
           "web platform has been deeply enriching, and it would not have been possible without "
           "the guidance and support that I received at every stage." % c["project_title"])
    P(doc, "I express my sincere gratitude to my project guide, %s, %s, for the continuous "
           "supervision, encouragement and constructive criticism that shaped this work. The "
           "detailed reviews on linguistic pre-processing and on the ethical use of generative "
           "models were particularly valuable." % (c["guide_name"], c["guide_designation"]))
    P(doc, "I am thankful to %s, %s, for providing the laboratory facilities, the necessary "
           "infrastructure and a research friendly environment in which this project could be "
           "carried out. I also thank the entire teaching and non teaching staff of the department "
           "for their cooperation." % (c["hod_name"], c["hod_designation"]))
    P(doc, "I would like to acknowledge the open source community behind the Next.js, React, "
           "MongoDB and Google GenAI tool chains, and the authors of the classical NLP literature "
           "whose algorithms \u2014 tokenisation, stemming, part of speech tagging and n-gram "
           "analysis \u2014 form the deterministic core of this system.")
    P(doc, "Finally, I thank my family and friends for their patience, motivation and constant "
           "support throughout the duration of this project.")
    blank(doc, 2)
    P(doc, "%s\nRoll No. : %s" % (c["student_name"], c["roll_number"]), justify=False)


def build_abstract(doc):
    doc.add_page_break()
    centred(doc, "ABSTRACT", size=18, bold=True, space_after=18)
    P(doc, "Preparing objective examinations is one of the most repetitive tasks in academic "
           "practice. For every unit of a syllabus, a teacher has to read the prescribed material, "
           "identify the concepts that deserve to be assessed, write a question for each concept, "
           "invent three plausible but incorrect options, and finally record an explanation that "
           "justifies the correct answer. When this exercise is repeated for several subjects and "
           "semesters, the effort becomes significant and the quality of the resulting question "
           "papers becomes inconsistent.")
    P(doc, "This report presents the design, implementation and evaluation of ExamEase, a web "
           "based system that converts raw lecture notes, textbook chapters and research papers "
           "into syllabus grounded multiple choice questions. The system follows a hybrid "
           "architecture. In the first stage, a deterministic natural language processing "
           "pipeline of seven stages \u2014 text normalisation, sentence segmentation, "
           "tokenisation with stopword filtering, lemmatisation, rule based part of speech "
           "tagging with noun phrase chunking, n-gram frequency extraction, and concept scoring "
           "with definition detection \u2014 analyses the document and produces a ranked list of "
           "concepts together with the sentences that support each concept. In the second stage, "
           "a large language model receives only this scored concept evidence and is constrained "
           "by a strict prompt to generate questions that are traceable to the supplied material. "
           "The generated items then pass through an automated validation and de-duplication "
           "layer that guarantees exactly four options, exactly one correct answer, a stored "
           "concept citation and a lexical similarity check based on the Jaccard coefficient of "
           "lemmatised token sets.")
    P(doc, "Because scanned pages, photocopied handouts and documents that use custom font "
           "encodings such as Identity-H break conventional PDF parsers, the ingestion stage uses "
           "a multimodal vision model to transcribe documents, whether they are typed, scanned or "
           "handwritten, into clean academic text before any linguistic processing begins.")
    P(doc, "The system has been implemented as a Next.js application with React on the client, "
           "route handlers on the server, and MongoDB as the document store. Teacher accounts are "
           "protected with salted password hashing and stateless JSON Web Token sessions stored "
           "in HTTP only cookies. The system was evaluated on a representative lecture note of "
           "fifty two sentences containing eight hundred and eighty one tokens; the pipeline "
           "extracted three hundred and ninety three unique terms, detected twenty eight "
           "definition sentences and returned twenty five scored concepts, and the de-duplication "
           "layer removed every candidate pair whose similarity exceeded the configured threshold "
           "of 0.65. The results indicate that combining deterministic linguistic analysis with a "
           "grounded generator produces questions that remain faithful to the uploaded material "
           "while reducing the preparation time of an examination by a substantial margin.")
    rich(doc, [("Keywords : ", {"bold": True}),
               ("Natural Language Processing, Automated Question Generation, Part of Speech "
                "Tagging, Lemmatisation, N-gram Analysis, Jaccard Similarity, Grounded "
                "Generation, Large Language Model, Optical Character Recognition, MongoDB, "
                "Next.js.", {})])


def build_toc_pages(doc):
    doc.add_page_break()
    centred(doc, "TABLE OF CONTENTS", size=18, bold=True, space_after=12)
    TOC_FIELD(doc, ' TOC \\o "1-3" \\h \\z \\u ',
              "MS Word builds this list automatically when the document is opened; if it stays "
              "blank, click inside the field and press F9 (or Ctrl + A then F9).")

    doc.add_page_break()
    centred(doc, "LIST OF FIGURES", size=18, bold=True, space_after=12)
    TOC_FIELD(doc, ' TOC \\h \\z \\t "Caption,1" ',
              "Every figure is inserted with a Caption styled paragraph, so Word populates this "
              "list with page numbers after the fields are updated.")

    doc.add_page_break()
    centred(doc, "LIST OF TABLES", size=18, bold=True, space_after=12)
    TOC_FIELD(doc, ' TOC \\h \\z \\t "TableCaption,1" ',
              "Table captions use a dedicated TableCaption style so that they are collected "
              "separately from the figures.")

    doc.add_page_break()
    centred(doc, "LIST OF ABBREVIATIONS", size=18, bold=True, space_after=12)
    abbr = [
        ("AI", "Artificial Intelligence"), ("API", "Application Programming Interface"),
        ("CID", "Character Identifier (PostScript font encoding)"),
        ("CRUD", "Create, Read, Update, Delete"), ("CSS", "Cascading Style Sheets"),
        ("DFD", "Data Flow Diagram"), ("ER", "Entity Relationship"),
        ("HTTP", "Hyper Text Transfer Protocol"), ("IDF", "Inverse Document Frequency"),
        ("JWT", "JSON Web Token"), ("LLM", "Large Language Model"),
        ("MCQ", "Multiple Choice Question"), ("MIME", "Multipurpose Internet Mail Extensions"),
        ("NLP", "Natural Language Processing"), ("NN", "Noun (POS tag)"),
        ("NNP", "Proper Noun (POS tag)"), ("OCR", "Optical Character Recognition"),
        ("PDF", "Portable Document Format"), ("POS", "Part of Speech"),
        ("REST", "Representational State Transfer"), ("SRS", "Software Requirements Specification"),
        ("TF", "Term Frequency"), ("TOC", "Table of Contents"), ("UI", "User Interface"),
        ("UML", "Unified Modelling Language"), ("URL", "Uniform Resource Locator"),
        ("XSS", "Cross Site Scripting"),
    ]
    tbl = doc.add_table(rows=0, cols=2)
    tbl.style = TBL_STYLE
    for i, (a, b) in enumerate(abbr):
        cells = tbl.add_row().cells
        for c_i, val in enumerate((a, b)):
            cells[c_i].text = ""
            p = cells[c_i].paragraphs[0]
            p.paragraph_format.space_after = Pt(1)
            p.paragraph_format.line_spacing = 1.1
            _set_run_font(p.add_run(val), size=10.5, bold=(c_i == 0))
            if i % 2 == 1:
                shade_cell(cells[c_i], "F7F9FC")
    for row in tbl.rows:
        row.cells[0].width = Cm(3.6)
        row.cells[1].width = Cm(11.9)

    doc.add_page_break()
    centred(doc, "ORGANISATION OF THE REPORT", size=18, bold=True, space_after=12)
    P(doc, "The report is organised into nine chapters supported by four appendices. The contents "
           "of each chapter are summarised below so that the reader can navigate the document "
           "quickly.")
    BULLETS(doc, [
        ("Chapter 1 \u2013 Introduction : ", "introduces the domain of automated assessment, states the "
         "problem, lists the objectives and defines the scope of the work."),
        ("Chapter 2 \u2013 Literature Survey : ", "reviews classical natural language processing "
         "literature, automated question generation approaches, document extraction techniques and "
         "existing tools, and identifies the research gap."),
        ("Chapter 3 \u2013 System Analysis and Design : ", "presents the existing system, the proposed "
         "architecture, use case analysis, data flow diagrams and the hardware and software "
         "requirements."),
        ("Chapter 4 \u2013 Methodology and Implementation : ", "explains the seven stage linguistic "
         "pipeline, the grounded prompt design, the service layer, validation logic and de-duplication "
         "in detail with algorithms and diagrams."),
        ("Chapter 5 \u2013 Database Design : ", "documents the MongoDB collections, relationships, "
         "indexes and data integrity rules."),
        ("Chapter 6 \u2013 Results and Discussion : ", "reports the measured behaviour of the pipeline "
         "on a sample document, the quality of extracted concepts, the de-duplication results, sample "
         "generated questions and a comparison with existing tools."),
        ("Chapter 7 \u2013 Security and Performance : ", "describes authentication, input validation, "
         "document privacy, connection pooling and the performance practices adopted."),
        ("Chapter 8 \u2013 System Implementation and Testing : ", "shows the screens of the running "
         "application and the unit, integration and boundary test cases together with their "
         "outcomes."),
        ("Chapter 9 \u2013 Conclusion and Future Scope : ", "summarises the achievements, maps them to "
         "the stated objectives and lists the possible extensions."),
        ("Appendices : ", "contain the REST API reference, the environment configuration, the project "
         "timeline and the source code of the core modules."),
    ])

# ==============================================================================
# 8. MAIN BODY - CHAPTERS 1 TO 3
# ==============================================================================

def chapter_1(doc):
    H(doc, 1, "Chapter 1  Introduction", page_break=True)

    H(doc, 2, "1.1  Introduction")
    P(doc, "Assessment is the closing loop of every teaching activity. A teacher explains a unit, "
           "and an examination verifies whether the learner has internalised the concepts of that "
           "unit. Among the formats available for objective assessment, the multiple choice "
           "question occupies a special place because it can be marked automatically, it can be "
           "administered to a large group simultaneously, and its statistics lend themselves to "
           "quantitative analysis of learning outcomes.")
    P(doc, "The construction of a good multiple choice question, however, is a skilled task. The "
           "question stem must isolate exactly one idea; the three distractors must be plausible "
           "enough to attract a partially informed learner, yet unambiguously wrong to a fully "
           "informed one; and the question must be traceable to the prescribed material, otherwise "
           "it will be challenged in the answer script review. ExamEase is a web platform that "
           "assists the teacher in exactly this task: it reads the prescribed document, applies a "
           "deterministic linguistic analysis to discover what the document is actually about, and "
           "then asks a large language model to write questions that are strictly grounded in the "
           "extracted evidence.")
    P(doc, "The distinguishing feature of the proposed system is the order in which the two "
           "technologies are applied. Rather than sending an entire document to a generative model "
           "and hoping that the model respects the syllabus, the system first reduces the document "
           "to a small, auditable set of scored concepts with their supporting sentences. Only this "
           "reduced evidence is shown to the model. The model therefore cannot introduce a topic "
           "that the pipeline did not detect, and every generated question can be traced back to a "
           "specific sentence of the uploaded document.")

    H(doc, 2, "1.2  Background and Motivation")
    P(doc, "Two independent trends make this project timely. The first is the maturing of "
           "classical natural language processing. Algorithms such as the Porter stemmer, rule "
           "based part of speech tagging and n-gram frequency analysis have been known for decades, "
           "but they are now easy to implement in a few hundred lines of TypeScript and they run in "
           "milliseconds without any external service. The second trend is the arrival of "
           "instruction following large language models that can produce well formed assessment "
           "items on demand, provided that the instructions are precise and the evidence supplied "
           "to the model is relevant.")
    P(doc, "The motivation for the project comes from a practical observation. When teachers were "
           "asked to prepare objective question banks for a semester, the recurring difficulties "
           "were the time required to write distractors, the risk of repeating the same concept in "
           "slightly different wording, and the difficulty of ensuring that every question of the "
           "paper corresponds to the prescribed syllabus rather than to the teacher's own reading "
           "of related literature. A system that automates the tedious portion of this work while "
           "keeping the teacher in control of the final artefact addresses all three difficulties.")

    H(doc, 2, "1.3  Problem Statement")
    P(doc, "The problem addressed by this project can be stated formally as follows.")
    P(doc, "Given a digital document D in the form of a lecture note, textbook chapter or "
           "examination oriented handout, produce a set of N multiple choice questions Q = "
           "{q1, q2, ... qN} such that", justify=True)
    BULLETS(doc, [
        "every question qi has exactly four options, exactly one of which is correct;",
        "every question qi is supported by at least one sentence of D, and the supporting concept "
        "is recorded along with the question so that a reviewer can verify it;",
        "no two questions of Q assess the same concept with the same wording, that is, the lexical "
        "similarity of any pair of questions remains below a configured duplicate threshold;",
        "the difficulty distribution of Q follows the difficulty selected by the teacher; and",
        "the entire operation completes within a few seconds for a document of ordinary size.",
    ])
    P(doc, "In addition to the generation problem, the system must solve an ingestion problem. A "
           "considerable proportion of real academic documents are scanned images, photocopies or "
           "files that use embedded font encodings in which the extracted character codes do not "
           "correspond to the visible glyphs. Conventional text extraction libraries return either "
           "an empty string or a stream of meaningless tokens for such documents, which makes any "
           "downstream linguistic analysis impossible. The ingestion stage of the system must "
           "therefore be robust to the physical form of the document.")

    H(doc, 2, "1.4  Objectives of the Project")
    BULLETS(doc, [
        "To design and implement a deterministic natural language processing pipeline that "
        "normalises a raw academic document, segments it into sentences, tokenises and lemmatises "
        "it, tags it with parts of speech, extracts noun phrases and n-grams, and ranks its "
        "concepts with an interpretable importance score.",
        "To extract explicit definitional statements from the document so that the questions "
        "generated from those sentences reflect the definitions actually used in the prescribed "
        "material.",
        "To build a universal document ingestion service that can transcribe typed, scanned, "
        "handwritten and non standard encoded documents into clean academic text.",
        "To design a strictly grounded generation layer in which the language model receives only "
        "the scored concepts and their supporting sentences, so that no fact outside the source "
        "document can enter a question.",
        "To implement an automated quality layer that validates the structure of every generated "
        "question and removes near duplicate items using lexical similarity.",
        "To develop a complete web application with teacher authentication, subject management, an "
        "interactive inspector of the linguistic analysis, an editable preview of the generated "
        "questions, a student examination interface and automated scoring with analytics.",
        "To evaluate the system on a representative academic document and report the measured "
        "behaviour of every stage of the pipeline.",
    ])

    H(doc, 2, "1.5  Scope of the Project")
    P(doc, "The scope of the present work is deliberately limited to objective assessment in the "
           "multiple choice format for text based academic material. The system accepts documents "
           "in the Portable Document Format, either uploaded as a file or supplied as pasted text, "
           "and it produces question sets that are stored in a document database and shared with "
           "students through a link.")
    P(doc, "The system supports the following in its present form.")
    BULLETS(doc, [
        "ingestion of PDF documents of up to several hundred pages, whether text based, scanned or "
        "handwritten;",
        "generation of between one and fifty questions per set in easy, medium, hard or mixed "
        "difficulty;",
        "teacher accounts with subject, semester and department attributes;",
        "in place editing and regeneration of individual questions before the set is saved;",
        "public test taking through a shareable link, followed by automatic scoring and a detailed "
        "answer review; and",
        "persistence of every attempt so that historical performance can be reviewed.",
    ])
    P(doc, "The following aspects are outside the scope of the present work and are listed in the "
           "future scope chapter : the generation of descriptive or programming questions, "
           "adaptation of difficulty during a live attempt, export of question papers to word "
           "processor or learning management system formats, and class level analytics with "
           "leaderboards.")

    H(doc, 2, "1.6  Novelty and Contribution")
    P(doc, "The contribution of this project lies in the combination of techniques rather than in "
           "the invention of a new algorithm. Four aspects deserve emphasis.")
    BULLETS(doc, [
        ("Deterministic evidence selection before generation. ", "Most question generation tools "
         "send a chunk of text to a language model. In this system the text is first reduced to a "
         "ranked concept list with supporting sentences, and the model sees only that list. The "
         "reduction is performed by transparent, reproducible rules that a teacher can inspect."),
        ("An inspectable analysis. ", "The interface exposes the intermediate linguistics of the "
         "document: the number of sentences and tokens, the unique vocabulary size, the detected "
         "definition sentences, the ranked concepts and the frequency of unigrams, bigrams and "
         "trigrams. A teacher can therefore understand why a particular question was generated."),
        ("Document form independence. ", "Because ingestion relies on a multimodal transcription "
         "step instead of a textual PDF parser, the pipeline behaves identically for a typed page, "
         "a scanned page and a handwritten note, and it is unaffected by font encodings such as "
         "Identity-H that defeat conventional parsers."),
        ("Automated quality enforcement. ", "Structural validation and lexical de-duplication are "
         "performed on the server before the teacher sees the questions, so the preview screen "
         "contains only items that already satisfy the format contract of the system."),
    ])

    H(doc, 2, "1.7  Organisation of the Report")
    P(doc, "Chapter 2 surveys the literature of classical natural language processing, automated "
           "question generation and document extraction, and compares existing tools. Chapter 3 "
           "presents the analysis of the existing system and the design of the proposed "
           "architecture with use case and data flow models. Chapter 4 documents the methodology "
           "and implementation of the linguistic pipeline, the grounded generation layer and the "
           "quality enforcement layer. Chapter 5 describes the database design. Chapter 6 reports "
           "the results obtained on a representative document and discusses them. Chapter 7 covers "
           "security and performance practices. Chapter 8 presents the implemented screens and the "
           "test cases with their outcomes. Chapter 9 concludes the report and lists the future "
           "scope. Appendices provide the API reference, the deployment instructions, the project "
           "timeline and the source code of the core modules.")


def chapter_2(doc):
    H(doc, 1, "Chapter 2  Literature Survey", page_break=True)

    H(doc, 2, "2.1  Introduction")
    P(doc, "This chapter reviews the body of work on which the project builds. The survey is "
           "organised in five parts: the classical linguistic techniques that constitute the "
           "deterministic core of the system, the approaches reported in the literature for "
           "automatic question generation, the techniques available for extracting text from "
           "academic documents, the question of controlling a generative model so that it does not "
           "depart from the supplied material, and finally the existing software tools that a "
           "teacher could use today. The chapter closes with the research gap that motivates the "
           "proposed design.")

    H(doc, 2, "2.2  Classical Natural Language Processing Techniques")
    P(doc, "Tokenisation is the oldest operation of computational linguistics. The word based "
           "regular expression tokeniser used in this project follows the classical convention of "
           "treating a sequence of alphanumeric characters as a token, with the refinement that "
           "hyphenated sequences such as state-of-the-art and client-server are preserved as single "
           "tokens, because in technical writing a hyphenated compound frequently denotes one term "
           "rather than a sequence of two words.")
    P(doc, "Stopword removal is normally treated as a matter of domain judgement rather than of "
           "grammar. The classic observation is that the high frequency function words of English "
           "carry little discriminative information, and that removing them reduces the size of the "
           "index without materially affecting retrieval quality. In an academic document the same "
           "argument applies with one important qualification: the noise produced by a broken PDF, "
           "such as the words obj, endobj, stream, flatedecode and cidfont, behaves statistically "
           "exactly like a stopword and must therefore be filtered by the same mechanism.")
    P(doc, "Stemming and lemmatisation constitute the morphological layer. The Porter stemmer "
           "remains the reference algorithm for rule based suffix stripping; it maps related forms "
           "such as analysis, analyses and analysing to a common stem. Its known weakness is that "
           "it operates without a lexicon, so it can produce a stem that is not an English word. "
           "Lemmatisation solves this weakness by consulting a dictionary of irregular forms, at "
           "the cost of requiring that dictionary. The present system uses lemmatisation for the "
           "computations that a teacher will inspect, namely concept normalisation and duplicate "
           "comparison, and keeps the stemmer available for the n-gram statistics.")
    P(doc, "Part of speech tagging in the classical tradition is performed either by hidden Markov "
           "models trained on annotated corpora or by rule based taggers. A rule based tagger "
           "assigns a category from the morphological shape of the word and from its position in "
           "the sentence, for instance treating a capitalised token that does not begin a sentence "
           "as a proper noun, and treating a token ending in -tion as a noun. The accuracy of a "
           "rule based tagger is lower than that of a statistical tagger, but it is deterministic, "
           "it requires no model file, and its rules can be explained to a teacher, which is the "
           "design priority here.")
    P(doc, "Noun phrase chunking converts a sequence of tagged tokens into candidate terminology. "
           "The pattern commonly used for technical text is an optional sequence of adjectives "
           "followed by one or more nouns, which captures multi word terms such as distributed "
           "operating system, relational database management system and lexical similarity "
           "threshold. Chunking is important for question generation because a concept that is "
           "extracted as a whole phrase yields a far better question stem than a concept extracted "
           "as an isolated noun.")
    P(doc, "N-gram frequency analysis extends the same idea to unigrams, bigrams and trigrams and "
           "is the standard method for detecting domain terminology without a domain lexicon. A "
           "recurring three word sequence in a lecture note is very likely to be the name of a "
           "technique, and the present system uses the trigram list both as an internal signal and "
           "as an element of the inspector that the teacher can read.")

    H(doc, 2, "2.3  Approaches to Automatic Question Generation")
    P(doc, "The literature on automatic question generation can be grouped into three generations "
           "of techniques.")
    P(doc, "Rule based generation transforms a declarative sentence into a question by syntactic "
           "manipulation. A sentence of the form \u201cX is defined as Y\u201d is converted into the "
           "stem \u201cWhat is X?\u201d with Y as the key and other terms of the sentence as "
           "distractors. This family of techniques is fully transparent and never produces a fact "
           "that is absent from the source, but its questions tend to be repetitive and its "
           "distractors are frequently too easy to eliminate.")
    P(doc, "Statistical and template driven generation selects sentences by an informativeness "
           "criterion, for example a high term frequency and inverse document frequency score, and "
           "then instantiates a small set of question templates. Such systems scale well but they "
           "inherit the rigidity of their templates, and the resulting items often test lexical "
           "recall rather than conceptual understanding.")
    P(doc, "Transformer based generation has become the dominant approach since encoder decoder "
           "models were shown to produce fluent questions. The difficulty reported consistently in "
           "this literature is faithfulness: a sequence to sequence model has no obligation to stay "
           "inside the source document, and it will happily produce a question about a fact that "
           "merely resembles the source. Techniques proposed to control faithfulness include "
           "restricting the decoder to an answer span, providing the model with a retrieved "
           "evidence set rather than a whole document, and validating the output against the source "
           "with an entailment model. The design adopted in this project belongs to the second of "
           "these techniques, extended with an explicit deterministic evidence selection step.")

    H(doc, 2, "2.4  Extraction of Text from Academic Documents")
    P(doc, "Text extraction from a PDF is a well known source of hidden failures. A PDF is a "
           "graphical format, not a text format; the visible text is a sequence of glyph drawing "
           "instructions. A parser reconstructs text by reading the content streams and decoding "
           "the font encoding tables. This works when the file is produced by a word processor with "
           "standard encodings, but fails in three common academic situations. Scanned pages "
           "contain no text objects at all, only images. Documents produced by certain desktop "
           "publishing tools use subset fonts with an Identity-H encoding in which the byte values "
           "do not correspond to the visible characters, so the parser returns a plausible but "
           "wrong string. Typeset mathematical material frequently renders a formula as a sequence "
           "of positioned glyphs whose reading order does not follow the visual order.")
    P(doc, "The response of the research community to these problems has been optical character "
           "recognition, historically with a page layout analysis stage followed by a character "
           "classifier, and more recently with multimodal neural models that accept the page image "
           "and emit structured text. The second family of techniques is used in this project "
           "because it handles typed, scanned and handwritten pages with the same code path, and "
           "because it can be asked to preserve headings, lists and paragraph structure, which the "
           "downstream sentence segmentation benefits from.")

    H(doc, 2, "2.5  Controlling a Generative Model of Question Writing")
    P(doc, "Three techniques for keeping a generative model close to its source are relevant to "
           "this project. Temperature reduction lowers the randomness of sampling and therefore "
           "makes a model repeat the most probable continuation instead of an inventive one; the "
           "values used in this system are 0.3 for bulk generation and 0.4 for the regeneration of "
           "a single item. Schema constrained decoding forces the decoder to emit a document that "
           "conforms to a supplied JSON schema, which removes the class of errors in which a model "
           "wraps its answer in prose; the system requests a JSON response type explicitly and "
           "still defers to a sanitising layer, because a schema constrains the shape of the output "
           "and not its truth. Evidence grounding, finally, restricts the input of the model to "
           "retrieved fragments rather than the whole document. This last technique is implemented "
           "here in its strict form: the model receives the concept names with their importance "
           "scores and the sentences that the pipeline associated with each concept, and nothing "
           "else.")

    H(doc, 2, "2.6  Review of Existing Tools and Comparison")
    P(doc, "A teacher who wants to automate question preparation today has four broad options: "
           "general purpose chat assistants, dedicated quiz generators that operate on pasted "
           "text, learning management systems with a question bank feature, and classical "
           "rule based question generators. Table 2.1 compares these options with the proposed "
           "system along the dimensions that matter for academic use.")
    TABLE(doc, "2.1", "Comparison of the proposed system with existing categories of tools",
          ["Criterion", "General chat assistants", "Text based quiz generators",
           "LMS question banks", "Proposed system (ExamEase)"],
          [
              ["Handles scanned or handwritten documents", "Partially, one page at a time",
               "No", "No", "Yes, whole document"],
              ["Handles non standard font encodings", "Partially", "No", "No", "Yes"],
              ["Evidence supplied to the generator", "Whole conversation",
               "Whole pasted text", "Manual entry", "Only scored concepts and their sentences"],
              ["Intermediate linguistic analysis visible to the teacher", "No", "No", "No",
               "Yes, full inspector"],
              ["Automatic duplicate removal", "No", "Rarely", "No", "Yes, lexical similarity filter"],
              ["Guaranteed structure (four options, one key)", "No", "Sometimes", "Manual",
               "Enforced on the server"],
              ["Works without a cloud API key", "No", "No", "Yes", "Yes, offline mock provider"],
              ["Test delivery and automatic scoring included", "No", "Usually separate",
               "Yes, but manual setup", "Yes, shareable link"],
              ["Concept citation stored with each question", "No", "No", "No", "Yes"],
           ], widths=[4.6, 2.9, 2.8, 2.6, 3.6], font_size=8.6)

    H(doc, 2, "2.7  Research Gap")
    P(doc, "The survey reveals a gap between two kinds of systems. On one side stand deterministic "
           "linguistic tools, which are transparent and faithful to the source but rigid in the "
           "questions they can formulate. On the other side stand generative systems, which are "
           "fluent and flexible but have no inherent obligation to remain inside the prescribed "
           "material, and which typically conceal their reasoning from the teacher. What is missing "
           "is a system in which a deterministic analysis selects and ranks the evidence, a "
           "generative model verbalises that evidence into an assessment item, and an automatic "
           "validator guarantees that the item obeys the structural contract of an examination. "
           "The present project implements precisely this composition, and it additionally treats "
           "the ingestion of the document as a first class problem rather than assuming that a "
           "clean text file is available.")

    H(doc, 2, "2.8  Summary")
    P(doc, "This chapter reviewed the classical linguistic operations of tokenisation, stopword "
           "removal, lemmatisation, part of speech tagging, chunking and n-gram analysis; the three "
           "generations of automatic question generation techniques; the failure modes of PDF text "
           "extraction and the optical recognition techniques that address them; the three "
           "mechanisms by which a generative model can be kept close to its evidence; and the "
           "capabilities of the tools that a teacher can use today. The gap identified in "
           "Section 2.7 defines the design that the next chapter presents.")

def chapter_3(doc):
    H(doc, 1, "Chapter 3  System Analysis and Design", page_break=True)

    H(doc, 2, "3.1  Introduction")
    P(doc, "This chapter analyses the system that exists today in academic practice, states the "
           "requirements that the proposed system must satisfy, and presents the architecture that "
           "has been designed to satisfy them. The architecture is described at three levels of "
           "abstraction: the layered decomposition of the application, the sequence of activities "
           "performed by the actors, and the flow of data through the processes of the system.")

    H(doc, 2, "3.2  Existing System and its Limitations")
    P(doc, "In the existing practice, a teacher prepares an objective question paper manually with "
           "the help of a word processor. The material is read from a printed or digital textbook, "
           "questions are typed one by one, and the answer key is maintained as a separate list. "
           "The limitations of this practice are well known.")
    BULLETS(doc, [
        "The preparation effort grows linearly with the number of questions, and the same effort "
        "has to be repeated when the syllabus is revised.",
        "The distractor design is the most time consuming part and is frequently the weakest part "
        "of a question paper.",
        "Coverage of the syllabus is difficult to verify without a deliberate mapping exercise, so "
        "entire sub topics can be omitted accidentally.",
        "Question banks accumulate duplicates over the years, and a student who receives two "
        "similar tests can memorise the pattern rather than the concept.",
        "When a general purpose chat assistant is used to speed up the work, the questions are "
        "frequently based on facts that are correct in general but absent from the prescribed "
        "material, which creates disputes during answer script review.",
    ])

    H(doc, 2, "3.3  Proposed System")
    P(doc, "The proposed system, ExamEase, is a web application in which a teacher uploads the "
           "prescribed document, inspects the linguistic analysis of that document, requests a "
           "specified number of questions at a specified difficulty, edits or regenerates "
           "individual questions, and finally saves the set. Saving produces a link that can be "
           "shared with students. A student who opens the link attempts the test question by "
           "question and receives a score card with a per question review as soon as the test is "
           "submitted.")
    P(doc, "The functional requirements of the system are listed below.")
    BULLETS(doc, [
        ("FR1 : ", "The system shall allow a teacher to register with a name, an email address and "
         "a password, and to log in and log out securely."),
        ("FR2 : ", "The system shall allow a teacher to create, list and delete subjects with a "
         "name, a semester and a department."),
        ("FR3 : ", "The system shall accept a PDF document or pasted text and shall extract clean "
         "academic text from it."),
        ("FR4 : ", "The system shall analyse the extracted text and expose the sentence count, "
         "token count, unique vocabulary size, detected definition sentences, ranked concepts and "
         "n-gram frequencies."),
        ("FR5 : ", "The system shall generate between one and fifty multiple choice questions at a "
         "selected difficulty from the analysed document."),
        ("FR6 : ", "Each generated question shall contain exactly four options, exactly one correct "
         "answer, an explanation and a concept citation."),
        ("FR7 : ", "The system shall allow the teacher to edit any field of a generated question or "
         "to regenerate that question individually."),
        ("FR8 : ", "The system shall store saved question sets in the database and shall produce a "
         "shareable link for each set."),
        ("FR9 : ", "The system shall allow a student to attempt a shared test and shall score the "
         "attempt automatically."),
        ("FR10 : ", "The system shall display the score, the percentage and a question wise review "
         "that shows the student's choice, the correct answer and the explanation."),
        ("FR11 : ", "The system shall remove near duplicate questions before displaying the "
         "preview."),
        ("FR12 : ", "The system shall operate in a fully offline mode with a deterministic mock "
         "provider when no language model API key is configured."),
    ])

    H(doc, 2, "3.4  System Architecture")
    P(doc, "The application is organised into four layers. The presentation layer is a React "
           "application rendered by Next.js; it contains the landing page, the authentication "
           "screens, the teacher dashboard, the question generator with its inspector, the quiz "
           "engine and the results view. The application layer consists of route handlers that run "
           "exclusively on the server and that expose the REST interface of the system. The service "
           "layer contains the domain logic: document transcription, the seven stage linguistic "
           "pipeline, the generative providers behind a common interface, and the question "
           "validation and de-duplication service. The data layer is the MongoDB store with four "
           "collections.")
    P(doc, "The separation between the application layer and the service layer is the most "
           "important structural decision of the design. A route handler performs authentication, "
           "validates the shape of the request, delegates the work to a service function, and "
           "formats the response. It contains no linguistic or generative logic. As a consequence "
           "the linguistic pipeline can be tested in isolation, and the generative provider can be "
           "replaced without touching any other module.")
    FIGURE(doc, DIAGRAMS / "fig_3_1_system_architecture.png", "3.1",
           "Layered system architecture of ExamEase", width_cm=15.6)

    H(doc, 2, "3.5  End to End Workflow")
    P(doc, "Figure 3.2 shows the workflow of the platform as three role lanes. The teacher lane "
           "contains the interactions that a human performs, the system lane contains the "
           "processing that happens on the server, and the student lane contains the attempt and "
           "the review. The figure makes the division of responsibility explicit: the system "
           "transcribes, analyses and generates, whereas the teacher decides the subject, the "
           "difficulty, the number of questions and the final wording of every item.")
    FIGURE(doc, DIAGRAMS / "fig_3_2_end_to_end_workflow.png", "3.2",
           "End to end activity workflow of the ExamEase platform", width_cm=15.6)

    H(doc, 2, "3.6  Use Case Analysis")
    P(doc, "The system has two human actors and one external system actor. The teacher is the "
           "primary actor and performs preparation activities. The student is a secondary actor "
           "who performs assessment activities through a shared link. The generative service is an "
           "external system actor that participates in the ingestion and generation use cases "
           "without initiating them.")
    FIGURE(doc, DIAGRAMS / "fig_3_3_use_case_diagram.png", "3.3",
           "Use case diagram of the ExamEase platform", width_cm=15.6)
    TABLE(doc, "3.1", "Use case description of the principal use cases",
          ["Use case", "Primary actor", "Precondition", "Main flow", "Outcome"],
          [
              ["Register teacher account", "Teacher", "None",
               "Submit name, email and password; the server checks that the email is unused, hashes "
               "the password and creates the user record",
               "Session cookie is issued and the teacher is redirected to the dashboard"],
              ["Create subject", "Teacher", "Teacher is authenticated",
               "Provide name, semester and department; the subject is inserted with the owner "
               "identifier",
               "Subject appears in the dashboard and subject list"],
              ["Upload and analyse document", "Teacher", "Teacher is authenticated and a subject "
               "exists",
               "Upload a PDF or paste text; the system transcribes the document and runs the seven "
               "stage pipeline",
               "The inspector displays the statistics, concepts, definitions and n-grams"],
              ["Generate questions", "Teacher", "A completed analysis is available",
               "Choose the number of questions, the difficulty and the title; the system builds the "
               "grounded prompt, validates the response and removes duplicates",
               "Editable preview cards are displayed"],
              ["Regenerate one question", "Teacher", "A preview card exists",
               "Request regeneration of a single item; the service rebuilds the prompt for that "
               "concept only",
               "The card is replaced and the remaining items are untouched"],
              ["Save set and share", "Teacher", "At least one valid question exists",
               "Save the set against the selected subject; the system returns the document "
               "identifier",
               "A shareable link of the form /quiz/{id} is produced"],
              ["Attempt test", "Student", "A valid shared link is opened",
               "Answer the questions in the engine and submit the answers",
               "The attempt is graded, stored and the results page is displayed"],
              ["Review results", "Student", "An attempt exists",
               "Open the results page of the attempt",
               "Score card, percentage badge and question wise review with citations are shown"],
           ], widths=[3.2, 2.4, 3.0, 4.4, 3.5], font_size=8.4)

    H(doc, 2, "3.7  Data Flow Diagrams")
    P(doc, "The data flow of the system is described at three levels of detail. The context diagram "
           "of Figure 3.4 treats the platform as a single process that exchanges data with the "
           "teacher and the student.")
    FIGURE(doc, DIAGRAMS / "fig_3_4_dfd_level_0.png", "3.4",
           "Context level data flow diagram (DFD level 0)", width_cm=15.2)
    P(doc, "The level one diagram of Figure 3.5 decomposes the platform into seven processes and "
           "five data stores. Process 1.0 manages authentication, process 2.0 manages subjects, "
           "process 3.0 performs ingestion, process 4.0 performs the linguistic analysis, process "
           "5.0 performs grounded generation, process 6.0 performs validation and de-duplication, "
           "and process 7.0 performs test delivery, grading and analytics. Data store D3 is "
           "deliberately shown as an in-memory store: the extracted document text is never written "
           "to disk, a decision that is revisited in the chapter on security.")
    FIGURE(doc, DIAGRAMS / "fig_3_5_dfd_level_1.png", "3.5",
           "Level 1 data flow diagram of ExamEase", width_cm=15.6)
    P(doc, "Figure 3.6 expands process 4.0 into its seven deterministic sub-processes. Because no "
           "sub-process makes a network call, the analysis of the same document always produces "
           "exactly the same result, which is a property that the teacher can rely upon while "
           "reviewing the extracted concepts.")
    FIGURE(doc, DIAGRAMS / "fig_3_6_dfd_level_2_nlp.png", "3.6",
           "Level 2 DFD showing the decomposition of the linguistic analysis process", width_cm=15.6)

    H(doc, 2, "3.8  Module Design")
    P(doc, "The implementation is divided into the modules described in Table 3.2. The table also "
           "records the principal source file of each module, which allows the design to be traced "
           "to the code listing in Appendix D.")
    TABLE(doc, "3.2", "Module wise description of the system",
          ["Module", "Responsibility", "Principal file"],
          [
              ["Authentication module", "Password hashing with bcrypt, issuing and verifying JSON "
               "Web Tokens, reading the current user from the session cookie", "src/lib/auth.ts"],
              ["Database module", "Singleton MongoDB client with connection pooling across "
               "development reloads and the getDb() accessor", "src/lib/db.ts"],
              ["Document ingestion module", "Transcription of a PDF buffer into clean academic text "
               "through the multimodal service", "src/services/pdf.service.ts"],
              ["Linguistic pipeline module", "Coordination of the seven deterministic stages and "
               "assembly of the analysis result", "src/services/nlp/nlpPipeline.ts"],
              ["Tokenisation module", "Word, number and hyphenated term tokenisation, stopword set "
               "and PDF noise filtering", "src/services/nlp/tokenizer.ts"],
              ["Segmentation module", "Sentence boundary detection with abbreviation and decimal "
               "protection", "src/services/nlp/sentenceSegmenter.ts"],
              ["Morphology module", "Irregular lemma dictionary, suffix rules and the Porter "
               "stemmer utility", "src/services/nlp/stemmerLemmatizer.ts"],
              ["Tagging module", "Rule based part of speech tagging and noun phrase chunking",
               "src/services/nlp/posTagger.ts"],
              ["Statistics module", "Unigram, bigram and trigram frequency extraction with "
               "stopword boundary rules", "src/services/nlp/ngrams.ts"],
              ["Concept module", "Definition pattern detection, heading and part of speech bonuses, "
               "importance scoring and supporting sentence collection",
               "src/services/nlp/conceptExtractor.ts"],
              ["Duplicate detection module", "Normalisation of question text into lemmatised token "
               "sets and Jaccard similarity comparison",
               "src/services/nlp/duplicateDetector.ts"],
              ["AI provider module", "Common generation interface, Gemini implementation and the "
               "offline mock implementation with the factory that selects between them",
               "src/services/ai/*"],
              ["Question service module", "Orchestration of generation, structural sanitisation, "
               "top up generation and single question regeneration",
               "src/services/mcq.service.ts"],
              ["Interface module", "Dashboard, generator with inspector, quiz engine and results "
               "views built with React and Tailwind CSS", "src/app/**, src/components/**"],
          ], widths=[4.4, 8.4, 3.8], font_size=8.6)

    H(doc, 2, "3.9  Hardware and Software Requirements")
    P(doc, "The system is a web application, so the requirements are stated separately for the "
           "development machine, the server that hosts the application and the client devices used "
           "by teachers and students.")
    TABLE(doc, "3.3", "Hardware requirements",
          ["Component", "Minimum", "Recommended"],
          [
              ["Processor", "Dual core 2.0 GHz", "Quad core 2.5 GHz or higher"],
              ["Memory", "4 GB RAM", "8 GB RAM or higher"],
              ["Storage", "2 GB free space for the run time and the database",
               "20 GB solid state storage"],
              ["Display (client)", "1366 x 768", "1920 x 1080"],
              ["Network", "1 Mbps for uploading a document",
               "5 Mbps for a smooth inspector experience"],
              ["Server (production)", "1 vCPU shared instance", "2 vCPU with 4 GB memory"],
           ], widths=[4.4, 5.6, 5.6], font_size=9.2)
    TABLE(doc, "3.4", "Software requirements",
          ["Category", "Software", "Version", "Purpose"],
          [
              ["Runtime", "Node.js", "20.x or 22.x", "Execution of the server and the build tools"],
              ["Framework", "Next.js with the App Router", "16.3", "Routing, server rendering and "
               "route handlers"],
              ["UI library", "React", "19.2", "Component based interface"],
              ["Styling", "Tailwind CSS with PostCSS", "4.x", "Utility first responsive styling"],
              ["Icons", "Lucide React", "1.47", "Vector iconography"],
              ["Database", "MongoDB Community Server", "7.6", "Document storage"],
              ["Driver", "MongoDB Node.js driver", "7.6", "Connection pooling and queries"],
              ["AI SDK", "Google GenAI SDK", "2.23", "Multimodal transcription and question "
               "generation"],
              ["Security", "jose and bcryptjs", "6.2 / 3.0", "JWT signing and password hashing"],
              ["Language", "TypeScript", "5.x", "Static type checking"],
              ["Tooling", "ESLint and the React Compiler", "9.x / 1.0", "Code quality and automatic "
               "memoisation"],
              ["Client", "Any modern browser", "Chrome 110 or later", "Access to the web interface"],
           ], widths=[2.6, 4.6, 2.8, 5.5], font_size=8.8)

    H(doc, 2, "3.10  Design Constraints and Assumptions")
    BULLETS(doc, [
        "The quality of the generated questions is bounded by the quality of the uploaded document. "
        "A document that contains only bullet points without complete sentences will produce fewer "
        "definition sentences and therefore fewer high quality concept candidates.",
        "The transcription stage requires a configured API key for the multimodal service. Without "
        "it, the system accepts pasted text and uses the offline provider for generation, so the "
        "application remains demonstrable in a laboratory without network access.",
        "The linguistic pipeline is language specific; the stopword set and the morphological rules "
        "implemented here target English technical writing.",
        "The system is designed for multiple choice assessment of factual and conceptual knowledge; "
        "it does not attempt to assess writing skill or problem solving ability.",
        "The shareable test link is intentionally public so that a student does not need an account "
        "in order to attempt a test. The attempt is nevertheless recorded against the authenticated "
        "user when one is available.",
    ])

def chapter_4(doc):
    H(doc, 1, "Chapter 4  Methodology and Implementation", page_break=True)

    H(doc, 2, "4.1  Introduction")
    P(doc, "This chapter documents the method implemented by the system. Section 4.2 introduces "
           "the linguistic pipeline as a whole; Sections 4.3 to 4.9 describe each of its seven "
           "stages with the corresponding algorithm and the exact rule set that was implemented. "
           "Sections 4.10 to 4.14 cover the ingestion service, the service layer design, the "
           "grounded prompt, the provider factory, the validation and de-duplication logic and the "
           "sequence of a complete generation request. Section 4.15 analyses the cost of the "
           "pipeline.")

    H(doc, 2, "4.2  The Seven Stage Linguistic Pipeline")
    P(doc, "The pipeline transforms a raw string into a structured analysis object. Each stage has "
           "a single responsibility, consumes the output of the previous stage and produces a "
           "well defined artefact. No stage performs a network call, and no stage uses a random "
           "number generator, so the entire pipeline is a pure function of its input. This property "
           "is what allows a teacher to compare the concept list of two uploaded documents and "
           "obtain a meaningful difference.")
    FIGURE(doc, DIAGRAMS / "fig_4_1_nlp_pipeline_stages.png", "4.1",
           "Seven stage deterministic NLP pre-processing pipeline", width_cm=15.4)
    TABLE(doc, "4.1", "Input and output contract of each pipeline stage",
          ["Stage", "Operation", "Input", "Output"],
          [
              ["S1", "Normalisation and noise cleaning", "Raw document text",
               "Whitespace normalised text of at least fifty characters"],
              ["S2", "Sentence segmentation", "Normalised text",
               "Array of sentences longer than ten characters"],
              ["S3", "Tokenisation and stopword filtering", "Normalised text",
               "Array of tokens with position and stopword flags"],
              ["S4", "Lemmatisation and stemming", "Individual tokens",
               "Canonical term forms used for normalisation and matching"],
              ["S5", "Part of speech tagging and chunking", "Sentences or token arrays",
               "Parts of speech and candidate noun phrases"],
              ["S6", "N-gram extraction", "Token array",
               "Frequency ranked unigrams, bigrams and trigrams"],
              ["S7", "Concept scoring and definition detection", "Sentences",
               "Ranked concepts with supporting sentences and detected definitions"],
           ], widths=[1.6, 4.4, 4.2, 5.3], font_size=8.8)

    H(doc, 2, "4.3  Stage 1 : Text Normalisation and Noise Cleaning")
    P(doc, "The first stage removes everything that is an artefact of the file format rather than a "
           "part of the academic content. Carriage returns and tabulators are replaced by a single "
           "space, runs of whitespace are collapsed, and the string is trimmed. A minimum length "
           "check of fifty readable characters is applied, and a document that does not satisfy it "
           "is rejected with an explanatory message rather than being analysed, because an analysis "
           "of an empty document would produce a confident but meaningless question set.")
    P(doc, "The stage also defends the later stages against the vocabulary of a broken extraction. "
           "When a PDF is decoded incorrectly, the extracted text is contaminated with the "
           "structural keywords of the format itself, namely obj, endobj, stream, endstream, "
           "flatedecode, length, filter, xref, trailer, startxref, cidfont and identity. These "
           "words are treated as stopwords and are additionally matched by a noise pattern that "
           "discards any sentence or candidate concept containing them.")
    CODE(doc, """
// Stage 1 - normalisation and noise cleaning (src/services/nlp/nlpPipeline.ts)
const cleanedText = rawText
  .replace(/[\\r\\t]+/g, ' ')
  .replace(/\\s+/g, ' ')
  .trim();

if (cleanedText.length < 50) {
  throw new Error('Not enough readable content was found in this PDF ...');
}""", "Listing 4.1  Text normalisation and the minimum content guard")

    H(doc, 2, "4.4  Stage 2 : Sentence Boundary Segmentation")
    P(doc, "Sentence segmentation converts a continuous string into the unit that the concept "
           "extractor reasons about. The naive rule of splitting after every full stop fails on "
           "academic writing in three ways: it splits after an abbreviation such as Dr. or Fig., it "
           "splits inside a decimal number such as 3.14, and it splits after an enumerated label "
           "such as 2. in a list. The implemented segmenter handles all three cases by protecting "
           "the ambiguous text before the split and restoring it afterwards.")
    P(doc, "The protection step replaces each known abbreviation with a numbered placeholder, "
           "replaces the decimal point of a number with a private marker, and only then applies a "
           "look-ahead split that requires a sentence ending punctuation mark followed by "
           "whitespace and a capital letter or a digit. Sentences shorter than eleven characters "
           "are discarded because they are almost always headings, page numbers or fragments.")
    CODE(doc, """
// Stage 2 - protect abbreviations and decimals before splitting
ABBREVIATIONS.forEach((abbr, idx) => {
  const regex = new RegExp(`\\\\b${abbr.replace('.', '\\\\.')}`, 'gi');
  protectedText = protectedText.replace(regex, `__ABBR_${idx}__`);
});
protectedText = protectedText.replace(/(\\d+)\\.(\\d+)/g, '$1__DECIMAL__$2');

// split only at [.!?] followed by whitespace and a capital letter or digit
const rawSentences = protectedText.split(/(?<=[.!?])\\s+(?=[A-Z0-9"']|$)/);""",
         "Listing 4.2  Abbreviation and decimal protection in the sentence segmenter")

    H(doc, 2, "4.5  Stage 3 : Tokenisation and Stopword Filtering")
    P(doc, "Tokenisation is performed by a single regular expression that matches a sequence of "
           "alphanumeric characters optionally containing internal hyphens. The hyphen rule is "
           "significant for technical material because it preserves terms such as client-server, "
           "state-of-the-art and TF-IDF as single tokens; splitting them would destroy the "
           "terminology that the concept extractor is looking for.")
    P(doc, "Each token is stored with its original spelling, its lower case form, its position in "
           "the document and a Boolean flag that records whether it is a stopword. The stopword set "
           "contains the standard English function words together with contractions, and it is "
           "extended with the PDF noise vocabulary described in Section 4.3. Tokens of a single "
           "character are also marked as stopwords, which removes artefacts such as enumerated list "
           "labels and stray bullets.")
    CODE(doc, """
// Stage 3 - tokenisation with hyphen aware matching (src/services/nlp/tokenizer.ts)
const wordRegex = /\\b[A-Za-z0-9]+(?:-[A-Za-z0-9]+)*\\b/g;
tokens.push({
  text: cleaned,
  original: raw,
  position: index++,
  isStopWord: ENGLISH_STOPWORDS.has(cleaned) || cleaned.length <= 1,
});""", "Listing 4.3  Hyphen aware tokenisation and stopword marking")

    H(doc, 2, "4.6  Stage 4 : Lemmatisation and Stemming")
    P(doc, "Two morphological utilities are provided. The lemmatiser maps a surface form to its "
           "canonical form in two steps. It first consults a dictionary of irregular forms, which "
           "contains the plurals that a suffix rule cannot recover, such as matrices to matrix, "
           "hypotheses to hypothesis, criteria to criterion, analyses to analysis, phenomena to "
           "phenomenon and data to datum, together with the irregular verb forms of the auxiliary "
           "verbs. Only when the word is not in the dictionary are the suffix rules applied, which "
           "handle the productive cases such as the plural -ies of capabilities, the gerund -ing of "
           "processing and the past tense -ed of normalised.")
    P(doc, "The Porter stemmer is implemented as a separate utility and is used where an aggressive "
           "reduction is desirable, namely in the similarity comparison and in the n-gram "
           "statistics. The distinction matters because the two operations have different "
           "requirements: the concept list that a teacher reads must contain recognisable words, so "
           "it uses the lemmatiser, whereas the duplicate detector benefits from aggressive "
           "merging.")
    CODE(doc, """
// Stage 4 - lemmatisation rules (src/services/nlp/stemmerLemmatizer.ts)
if (IRREGULAR_LEMMAS[lower]) return IRREGULAR_LEMMAS[lower];   // matrices -> matrix
if (lower.endsWith('ies') && lower.length > 4) return lower.slice(0, -3) + 'y';
if (lower.endsWith('ing') && lower.length > 5) {
  if (lower.endsWith('tting') || lower.endsWith('nning') ||
      lower.endsWith('pping') || lower.endsWith('mming')) return lower.slice(0, -4);
  return lower.slice(0, -3);
}
if (lower.endsWith('s') && !lower.endsWith('ss') && lower.length > 3) return lower.slice(0, -1);""",
         "Listing 4.4  The irregular dictionary and the productive suffix rules of the lemmatiser")

    H(doc, 2, "4.7  Stage 5 : Part of Speech Tagging and Noun Phrase Chunking")
    P(doc, "The tagger assigns a category to each token from its morphological shape together with "
           "one positional rule. A token that begins with a capital letter and is not the first "
           "token of the sentence is treated as a proper noun, which captures product names, "
           "algorithm names and place names. A token that ends in one of the nominal suffixes such "
           "as -tion, -sion, -ment, -ness, -ance, -ence, -ism or -ity is a noun; a token ending in "
           "-able, -ible, -al, -ic, -ous, -ful, -less or -ive is an adjective; a token ending in -ly "
           "is an adverb; -ing denotes a gerund; -ed denotes a past tense verb; and a plural -s "
           "marks a plural noun. Tokens that are not stopwords and match none of the rules default "
           "to the noun category, which is the safe default for technical text.")
    P(doc, "The chunker then scans the tagged sequence and collects maximal contiguous runs of "
           "adjectives and nouns, which corresponds to the classical pattern of an optional "
           "adjective sequence followed by one or more nouns. A run is emitted as a candidate "
           "concept only if it contains at least one noun and its textual length exceeds three "
           "characters. This is the mechanism by which a phrase such as relational database "
           "management system enters the concept list as a single candidate rather than as four "
           "separate words.")
    CODE(doc, """
// Stage 5 - suffix rules of the rule based tagger (src/services/nlp/posTagger.ts)
if (/^[A-Z][a-zA-Z0-9_-]+$/.test(word) && token.position > 0) return 'NNP';
if (lower.endsWith('tion') || lower.endsWith('ment') || lower.endsWith('ness') ||
    lower.endsWith('ity'))  return 'NN';
if (lower.endsWith('able') || lower.endsWith('ous') || lower.endsWith('ive')) return 'JJ';
if (lower.endsWith('ly')) return 'RB';
if (lower.endsWith('ing')) return 'VBG';
if (lower.endsWith('ed'))  return 'VBD';

// chunker : build maximal runs of [JJ]* [NN|NNS|NNP]+ and keep runs with a noun""",
         "Listing 4.5  Morphological part of speech rules and the noun phrase pattern")

    H(doc, 2, "4.8  Stage 6 : N-gram Extraction")
    P(doc, "The n-gram stage computes three frequency tables from the token sequence: single terms, "
           "consecutive pairs and consecutive triples. Four rules keep the tables meaningful. A "
           "unigram must be a non stopword of more than two characters. A bigram is retained if at "
           "least one of its two tokens is a non stopword and both tokens are longer than two "
           "characters. A trigram is retained if at least one of its boundary tokens is a non "
           "stopword. Every table is sorted by descending frequency, and the pipeline publishes the "
           "thirty most frequent unigrams, the twenty five most frequent bigrams and the twenty "
           "most frequent trigrams, which is exactly the volume that the inspector can display "
           "without becoming unreadable.")
    CODE(doc, """
// Stage 6 - bigram extraction with the non stopword boundary rule (src/services/nlp/ngrams.ts)
for (let i = 0; i < tokens.length - 1; i++) {
  const t1 = tokens[i], t2 = tokens[i + 1];
  if ((!t1.isStopWord || !t2.isStopWord) && t1.text.length > 2 && t2.text.length > 2) {
    const bigramStr = `${t1.text} ${t2.text}`;
    bigramMap[bigramStr] = (bigramMap[bigramStr] || 0) + 1;
  }
}""", "Listing 4.6  Frequency accumulation for bigrams")

    H(doc, 2, "4.9  Stage 7 : Concept Extraction and Definition Scoring")
    P(doc, "The final stage combines the outputs of the previous six into a ranked list of "
           "concepts. Two sources of candidates are merged. The first source is a set of four "
           "definition patterns applied to every sentence; a match produces a candidate concept "
           "with a strong definition bonus. The second source is the noun phrase list produced by "
           "the chunker, which contributes candidates with a smaller part of speech bonus.")
    P(doc, "The four definition patterns are the copular definition, in which a capitalised term is "
           "followed by is or are and then one of defined as, a, an, the, described as, known as or "
           "referred to as; the lexical definition, in which a term is followed by refers to, "
           "denotes, represents, means or signifies; the colon pattern, in which a term is followed "
           "by a colon and a definition; and the labelled definition, in which the word Definition "
           "or Concept introduces the term. Candidate strings are rejected when they consist only "
           "of a demonstrative pronoun, when they are shorter than three characters or longer than "
           "fifty characters, or when they match the PDF noise pattern.")
    P(doc, "Each accepted candidate is normalised by lemmatising every word of the candidate "
           "phrase, which is what allows two surface variants of the same concept to accumulate "
           "into a single entry. Every entry maintains a frequency counter, a definition bonus, a "
           "part of speech bonus, a heading bonus and a set of at most four supporting sentences. "
           "The importance score of a concept is computed by the formula given below and is capped "
           "at 10.0 so that a single dominant term cannot dominate the ranking by an order of "
           "magnitude.")
    p = doc.add_paragraph()
    p.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.space_after = Pt(8)
    _set_run_font(p.add_run("Importance  =  min ( 10.0 ,  (1.2 \u00d7 frequency)  +  definition "
                            "bonus  +  part of speech bonus  +  (0.8 \u00d7 supporting sentences)  )"),
                  size=12, bold=True, color="1F3864")
    P(doc, "In the implemented constants the definition bonus is 4.5 per matched definition pattern, "
           "the part of speech bonus is 2.0 for a candidate that came from a definition pattern and "
           "1.5 for a candidate that came from the noun phrase chunker, and the heading bonus is "
           "reserved for candidates that begin with a capital letter and appear early in the "
           "document. Because the score is clipped at ten, several concepts of a well written "
           "lecture note reach the ceiling, which is a deliberate choice: the ranking exists to "
           "produce the correct set of concepts rather than to distinguish finely between them. The "
           "pipeline returns the twenty five highest scoring concepts.")
    CODE(doc, """
// Stage 7 - importance score of an extracted concept (src/services/nlp/conceptExtractor.ts)
const rawScore = (item.frequency * 1.2) + item.definitionBonus + item.posBonus +
                 (item.supportingSentences.size * 0.8);
const importanceScore = Math.min(10.0, Math.round(rawScore * 10) / 10);""",
         "Listing 4.7  Computation of the concept importance score")

    H(doc, 2, "4.10  Document Ingestion")
    P(doc, "The ingestion service receives the uploaded file as an in-memory buffer, converts it to "
           "a base64 string and submits it to the multimodal transcription service together with an "
           "instruction that asks for all readable text with headings, paragraphs and bullet "
           "structure preserved and without any commentary. The response is validated to contain at "
           "least twenty characters; otherwise the request fails with a message that the document "
           "contains no readable text. The buffer is never written to the file system, which means "
           "that no exam material is left behind on the server after the request completes.")

    H(doc, 2, "4.11  Service Layer Structure")
    P(doc, "The service layer is designed around a single interface. The generation capability is "
           "declared once, and two classes implement it: one that calls the external language model "
           "and one that generates questions locally without any network access. The question "
           "service depends only on the interface, so it is unaware of which implementation is "
           "active. Figure 4.2 shows the classes and their relationships.")
    FIGURE(doc, DIAGRAMS / "fig_4_2_class_diagram.png", "4.2",
           "Class and component diagram of the service layer", width_cm=15.6)

    H(doc, 2, "4.12  Grounded Prompt Engineering")
    P(doc, "The prompt is constructed from the analysis rather than from the document. The fifteen "
           "highest scoring concepts are written out with their importance score, and beneath each "
           "concept the sentences that the pipeline associated with it are listed, separated by "
           "horizontal rules. The instruction block that precedes this evidence states seven "
           "constraints, and the response format is specified as a JSON object with a questions "
           "array. Table 4.2 lists the constraints and the reason for each of them.")
    TABLE(doc, "4.2", "Prompt constraints and their purpose",
          ["Constraint in the prompt", "Purpose"],
          [
              ["Do not use external knowledge; all facts must come from the source content",
               "This is the central anti hallucination rule. Because the source content is the "
               "scored concept list, the model cannot introduce a topic that the pipeline did not "
               "find in the document."],
              ["Generate exactly the requested number of questions",
               "Keeps the size contract; the service still tops up the count if the model returns "
               "fewer items."],
              ["Target the selected difficulty",
               "Propagates the teacher's choice into the wording of the stem and the subtlety of "
               "the distractors."],
              ["Exactly four options per question",
               "Matches the format of the examination and allows the automatic grading key to work."],
              ["Exactly one option must be correct",
               "Ensures that the answer key is unambiguous; ambiguous items would not survive a "
               "student challenge."],
              ["Provide a concise explanation citing the source material",
               "Produces the explanation that the student reads in the review screen and that the "
               "teacher uses to verify the item."],
              ["Avoid duplicate questions and duplicate choices",
               "Reduces the work of the de-duplication layer, although the layer still runs "
               "unconditionally."],
              ["Return JSON only, with no markdown code fences",
               "Makes the response machine readable; the service additionally strips code fences "
               "defensively."],
           ], widths=[6.6, 8.6], font_size=8.8)

    H(doc, 2, "4.13  Provider Factory and the Offline Fallback")
    P(doc, "The factory decides which provider to instantiate from the presence of the API key. If "
           "a key of more than five characters is configured, the external provider is created; "
           "otherwise the offline provider is created. Both satisfy the same interface and return "
           "the same shape of partial questions, so the rest of the system behaves identically. The "
           "offline provider composes its questions from the same concept list: it uses the "
           "detected definition of a concept as the correct answer and draws its distractors from "
           "other concepts of the same document, which keeps the demonstration questions "
           "thematically coherent.")
    FIGURE(doc, DIAGRAMS / "fig_4_3_ai_factory_fallback.png", "4.3",
           "Dual provider AI factory with graceful offline fallback", width_cm=15.6)

    H(doc, 2, "4.14  Validation and De-duplication")
    P(doc, "The raw output of a provider is never shown to the teacher. It passes through two "
           "filters. The first is structural: the question text must contain at least five "
           "characters, the option array is trimmed of empty entries and padded or truncated to "
           "exactly four options, and the correct answer is forced to be an exact member of the "
           "option array, with a documented fallback to the first option if the model returned a "
           "paraphrase. The second filter is lexical.")
    P(doc, "The lexical filter compares every candidate against the questions that have already "
           "been accepted. Each question is normalised by tokenising it, removing stopwords and "
           "lemmatising every remaining token, which yields a set of terms that represents the "
           "content of the question independently of its surface wording. The similarity of two "
           "questions is the Jaccard coefficient of their token sets, that is the size of the "
           "intersection divided by the size of the union. A candidate whose similarity to any "
           "accepted question is 0.65 or greater is discarded. Figure 4.4 shows the flow.")
    FIGURE(doc, DIAGRAMS / "fig_4_4_dedup_flow.png", "4.4",
           "Lexical near-duplicate detection using Jaccard token similarity", width_cm=14.6)
    CODE(doc, """
// Jaccard similarity of two lemmatised token sets (src/services/nlp/duplicateDetector.ts)
export function calculateJaccardSimilarity(setA: Set<string>, setB: Set<string>): number {
  if (setA.size === 0 || setB.size === 0) return 0;
  let intersectionCount = 0;
  setA.forEach((token) => { if (setB.has(token)) intersectionCount++; });
  const unionSize = new Set([...Array.from(setA), ...Array.from(setB)]).size;
  return unionSize === 0 ? 0 : intersectionCount / unionSize;
}""", "Listing 4.8  Jaccard similarity computed over lemmatised token sets")
    P(doc, "After both filters have been applied, the service checks the size of the surviving "
           "list. If the model returned fewer valid items than the teacher requested, a top up "
           "round is performed with the remaining count, and the result is finally trimmed to the "
           "requested number. The teacher therefore always receives exactly the requested number of "
           "questions, provided that the document contains enough distinct concepts to support "
           "them.")

    H(doc, 2, "4.15  Cost and Complexity of the Pipeline")
    P(doc, "The linguistic pipeline is linear in the size of the document for five of its seven "
           "stages. The exceptions are the concept scoring stage, which sorts the candidate map, "
           "and the duplicate detection stage, which compares each new question with the questions "
           "already accepted. Table 4.3 summarises the cost of each stage.")
    TABLE(doc, "4.3", "Time complexity of the processing stages",
          ["Stage", "Operation", "Complexity", "Remarks"],
          [
              ["S1", "Normalisation", "O(n)", "Four regular expression passes over the character "
               "stream"],
              ["S2", "Sentence segmentation", "O(n)", "One split with a constant number of "
               "abbreviation substitutions"],
              ["S3", "Tokenisation", "O(n)", "Single scan of the tokenising regular expression"],
              ["S4", "Lemmatisation", "O(m)", "Dictionary lookup or a constant number of suffix "
               "tests per token"],
              ["S5", "POS tagging and chunking", "O(m)", "One pass over the tagged tokens"],
              ["S6", "N-gram extraction", "O(m)", "Three linear scans with hash based counters"],
              ["S7", "Concept scoring", "O(c log c)", "Sorting of c candidate concepts; c is bounded "
               "by the number of noun phrases"],
              ["De-duplication", "O(k\u00b2 t)", "Comparison of k questions over token sets of "
               "average size t; k is at most fifty"],
              ["Generation call", "Network bound", "Dominant cost of a request; typically one to "
               "five seconds for a document of ordinary size"],
           ], widths=[2.4, 4.0, 2.8, 6.0], font_size=8.8)
    P(doc, "In practice the linguistic stages complete in a few tens of milliseconds for a "
           "document of a few thousand words, and the total response time of a generation request "
           "is dominated by the external generation call. The offline provider removes even that "
           "cost, which makes the system suitable for classroom demonstrations in which network "
           "access cannot be guaranteed.")

    H(doc, 2, "4.16  Sequence of a Complete Generation Request")
    P(doc, "Figure 4.5 assembles the operations of the previous sections into the complete "
           "lifecycle of a generation request, from the upload of the document to the identifier "
           "that the teacher shares with the class. The diagram also makes the cost profile of the "
           "system explicit: only two of the operations are network calls, and every operation "
           "between them is deterministic local processing that can be executed and verified "
           "without an API key.")
    FIGURE(doc, DIAGRAMS / "fig_4_5_sequence_generation.png", "4.5",
           "Sequence diagram of the document to MCQ generation flow", width_cm=15.6)

def chapter_5(doc):
    H(doc, 1, "Chapter 5  Database Design", page_break=True)

    H(doc, 2, "5.1  Introduction")
    P(doc, "The persistence layer is a MongoDB store with four collections: users, subjects, "
           "mcq_sets and quiz_attempts. The schema was designed document first, which means that "
           "the objects that are always read together are stored together. A question set is a "
           "single document that contains its questions, and a quiz attempt is a single document "
           "that contains the answers together with the correct answer and the explanation of each "
           "question. As a consequence, delivering a test to a student requires one read, and "
           "rendering a result page requires one read.")

    H(doc, 2, "5.2  Entity Relationship Model")
    P(doc, "Figure 5.1 shows the four collections and their relationships. A user creates many "
           "subjects and authors many question sets; a subject contains many question sets; a "
           "question set generates many quiz attempts. Identifier fields are stored as strings that "
           "refer to the ObjectId of the parent document, an approach that keeps the interface "
           "simple while allowing the application to validate the format of an identifier before "
           "issuing a query.")
    FIGURE(doc, DIAGRAMS / "fig_5_1_er_diagram.png", "5.1",
           "Entity relationship diagram of the ExamEase MongoDB schema", width_cm=15.6)

    H(doc, 2, "5.3  Collection Schemas")
    TABLE(doc, "5.1", "Schema of the users collection",
          ["Field", "Type", "Description"],
          [
              ["_id", "ObjectId", "Primary identifier generated by MongoDB"],
              ["name", "string", "Display name of the teacher"],
              ["email", "string", "Login identifier, stored in lower case and unique"],
              ["passwordHash", "string", "Salt and hash produced by bcrypt with ten rounds; the "
               "plain password is never stored"],
              ["role", "string", "TEACHER or STUDENT; teacher accounts are created by registration"],
              ["createdAt", "date", "Creation timestamp used for ordering the dashboard"],
           ], widths=[3.4, 2.4, 9.4], font_size=9.0)
    TABLE(doc, "5.2", "Schema of the subjects collection",
          ["Field", "Type", "Description"],
          [
              ["_id", "ObjectId", "Primary identifier"],
              ["name", "string", "Subject or class name"],
              ["semester", "string", "Semester label, for example Semester IV"],
              ["department", "string", "Department to which the subject belongs"],
              ["userId", "string", "Identifier of the owning teacher"],
              ["userName", "string", "Denormalised name of the owner, used to render cards without "
               "a join"],
              ["createdAt", "date", "Creation timestamp"],
           ], widths=[3.4, 2.4, 9.4], font_size=9.0)
    TABLE(doc, "5.3", "Schema of the mcq_sets collection",
          ["Field", "Type", "Description"],
          [
              ["_id", "ObjectId", "Primary identifier and the value used in the shareable link"],
              ["title", "string", "Title of the test entered by the teacher"],
              ["subjectId", "string", "Identifier of the parent subject"],
              ["userId", "string", "Identifier of the authoring teacher"],
              ["userName", "string", "Denormalised author name"],
              ["fileName", "string", "Name of the uploaded document, retained for reference"],
              ["questionCount", "number", "Number of questions stored in the document"],
              ["difficulty", "string", "easy, medium, hard or mixed as selected by the teacher"],
              ["questions", "array", "Embedded array of question objects described below"],
              ["createdAt", "date", "Creation timestamp"],
           ], widths=[3.4, 2.4, 9.4], font_size=9.0)
    TABLE(doc, "5.4", "Structure of an embedded question object",
          ["Field", "Type", "Description"],
          [
              ["id", "string", "Identifier unique inside the set; used as the key of the answer map"],
              ["question", "string", "Question stem"],
              ["options", "array of 4 strings", "Exactly four options in display order"],
              ["correctAnswer", "string", "Exact copy of the correct option, which makes grading a "
               "string comparison"],
              ["explanation", "string", "Justification shown in the review screen"],
              ["difficulty", "string", "Difficulty of the item"],
              ["type", "string", "definition, concept, relationship, comparison, process, "
               "application or example"],
              ["sourceConcept", "string", "Concept of the analysis from which the item was "
               "generated; this is the citation that makes the item verifiable"],
           ], widths=[3.4, 3.0, 8.8], font_size=9.0)
    TABLE(doc, "5.5", "Schema of the quiz_attempts collection",
          ["Field", "Type", "Description"],
          [
              ["_id", "ObjectId", "Primary identifier of the attempt"],
              ["mcqSetId", "string", "Identifier of the attempted question set"],
              ["mcqSetTitle", "string", "Denormalised title, so that a result page needs no second "
               "read"],
              ["subjectId", "string", "Identifier of the subject of the set"],
              ["userId", "string", "Identifier of the student who submitted the attempt"],
              ["userName", "string", "Display name of the student"],
              ["score", "number", "Number of correctly answered questions"],
              ["totalQuestions", "number", "Number of questions in the set"],
              ["percentage", "number", "Rounded percentage used for the colour coded badge"],
              ["answers", "array", "Embedded array of answer reviews: question identifier, question "
               "text, selected option, correct option, correctness flag and explanation"],
              ["completedAt", "date", "Submission timestamp"],
           ], widths=[3.4, 2.4, 9.4], font_size=9.0)

    H(doc, 2, "5.4  Indexing and Query Patterns")
    P(doc, "The queries issued by the application are few and predictable, and the index plan "
           "reflects them directly. Table 5.6 records the planned indexes and the query that each "
           "index supports.")
    TABLE(doc, "5.6", "Index plan of the collections",
          ["Collection", "Index", "Query supported"],
          [
              ["users", "{ email: 1 } unique", "Authentication by email and the uniqueness check "
               "performed during registration"],
              ["subjects", "{ userId: 1, createdAt: -1 }", "Listing the subjects of the logged in "
               "teacher in reverse chronological order"],
              ["mcq_sets", "{ subjectId: 1, createdAt: -1 }", "Rendering the question sets that "
               "belong to a subject"],
              ["mcq_sets", "{ userId: 1, createdAt: -1 }", "Listing the recent sets of a teacher on "
               "the dashboard"],
              ["quiz_attempts", "{ userId: 1, completedAt: -1 }", "Showing the attempt history of a "
               "student"],
              ["quiz_attempts", "{ mcqSetId: 1 }", "Collecting all attempts of a particular test "
               "for later analytics"],
           ], widths=[3.2, 4.4, 7.6], font_size=9.0)

    H(doc, 2, "5.5  Data Integrity and Deletion Behaviour")
    P(doc, "MongoDB does not enforce referential integrity, so the application enforces it. Every "
           "route that mutates a document first resolves the authenticated user from the session "
           "cookie and then checks that the document being changed belongs to that user; a request "
           "for a document of another teacher receives a not found response rather than a "
           "forbidden response, so that the existence of foreign identifiers is not disclosed.")
    P(doc, "Deletion follows a deliberate policy. Deleting a subject removes the subject document "
           "and the question sets that reference it, because a set without a subject cannot be "
           "reached from the interface. Attempts are retained, since they are historical records of "
           "student performance and are only ever read in aggregate. Deleting a question set "
           "removes the set but leaves the attempts that were taken on it, so that a student's "
           "history is not silently altered by a teacher's action.")
    P(doc, "Because questions and answers are embedded rather than referenced, updating a question "
           "in a saved set is a single atomic document update. This removes the class of "
           "inconsistency in which a set and its questions could disagree after a partially "
           "completed write, which would be possible in a fully normalised relational design "
           "without an explicit transaction.")


def chapter_6(doc):
    H(doc, 1, "Chapter 6  Results and Discussion", page_break=True)

    H(doc, 2, "6.1  Experimental Setup")
    P(doc, "The system was exercised on a representative unit of academic material: an eight page "
           "text based lecture note on the foundations of natural language processing, containing "
           "complete sentences of the kind that a prescribed textbook contains. The same document "
           "was used for every measurement reported in this chapter so that the figures are "
           "comparable. The linguistic pipeline was executed directly on the extracted text, "
           "independently of the network dependent stages, so that the measurements reported here "
           "are reproducible without an API key.")
    TABLE(doc, "6.1", "Description of the evaluation input",
          ["Attribute", "Value"],
          [
              ["Document", SAMPLE_NLP["source"]],
              ["Content", "Explanatory prose with definitions of tokenisation, segmentation, "
               "lemmatisation, stemming, part of speech tagging, chunking, n-grams, concept "
               "extraction, TF-IDF, cosine similarity, Jaccard similarity, OCR, grounded "
               "generation and Bloom taxonomy"],
              ["Language", "English technical writing"],
              ["Estimated word count", "Approximately 880 words"],
              ["Measurement method", "Direct execution of the pipeline functions on the extracted "
               "text; all values in this chapter are measured, not estimated"],
           ], widths=[4.4, 11.0], font_size=9.2)

    H(doc, 2, "6.2  Measurements of the Linguistic Pipeline")
    P(doc, "The pipeline described in Chapter 4 produced the statistics recorded in Table 6.2 and "
           "plotted in Figure 6.1. The document was segmented into fifty two sentences, tokenised "
           "into eight hundred and eighty one tokens of which three hundred and ninety three were "
           "unique, and it was found to contain twenty eight sentences that satisfy at least one "
           "definition pattern. Twenty five concepts survived the final filtering and ranking.")
    FIGURE(doc, DIAGRAMS / "fig_6_1_pipeline_statistics.png", "6.1",
           "Measured NLP pipeline statistics for the sample document", width_cm=15.0)
    TABLE(doc, "6.2", "Measured output of the pipeline",
          ["Metric", "Value", "Interpretation"],
          [
              ["Sentences detected", str(SAMPLE_NLP["sentence_count"]),
               "The document is well formed; segments of fewer than eleven characters were "
               "discarded as headings or labels"],
              ["Tokens produced", str(SAMPLE_NLP["token_count"]),
               "Includes stopwords and numbers; the figure is close to the manual token count of "
               "the document"],
              ["Unique tokens", str(SAMPLE_NLP["unique_token_count"]),
               "Type count used to estimate the lexical richness of the material"],
              ["Vocabulary density", "%.1f %%" % (100.0 * SAMPLE_NLP["unique_token_count"] /
                                                  SAMPLE_NLP["token_count"]),
               "A density near forty five per cent is typical of explanatory technical prose"],
              ["Definition sentences", str(SAMPLE_NLP["definition_count"]),
               "Sentences matching at least one of the four definition patterns; these sentences "
               "supply the strongest evidence for question generation"],
              ["Concepts retained", str(SAMPLE_NLP["concept_count_extracted"]),
               "Ranked candidates after the noise filter and the maximum of twenty five entries"],
           ], widths=[3.6, 2.6, 9.0], font_size=9.0)

    H(doc, 2, "6.3  Quality of the Extracted Concepts")
    P(doc, "Table 6.3 lists the ten highest scoring concepts with their measured frequency and the "
           "number of supporting sentences retained for each. Figure 6.2 plots the same data. Two "
           "observations follow. First, every concept that a teacher would expect to see in a unit "
           "on the foundations of natural language processing appears in the list: normalisation, "
           "segmentation, lemmatisation, tagging, chunking, extraction, similarity and the "
           "taxonomy. Second, the scores cluster at the ceiling of ten for the concepts that were "
           "introduced by an explicit definition, which confirms that the definition bonus "
           "dominates the ranking as intended.")
    FIGURE(doc, DIAGRAMS / "fig_6_2_concept_scores.png", "6.2",
           "Importance scores of the top ranked concepts", width_cm=15.2)
    rows = []
    for i, (name, score, freq, sup) in enumerate(SAMPLE_NLP["top_concepts"], start=1):
        rows.append([str(i), name, "%.1f" % score, str(freq), str(sup)])
    TABLE(doc, "6.3", "Top ranked concepts with importance score, term frequency and supporting sentences",
          ["Rank", "Concept", "Importance score", "Term frequency", "Supporting sentences"],
          rows, widths=[1.4, 6.4, 2.8, 2.4, 2.4], font_size=9.0,
          align_center_cols={0, 2, 3, 4})
    P(doc, "The supporting sentences are the mechanism that makes the generation auditable. For the "
           "concept text normalization, for example, the pipeline retained the sentence that "
           "introduces it as the first stage of a linguistic pipeline and the sentence that defines "
           "it as the removal of noise such as control characters, repeated whitespace, carriage "
           "returns and binary artefacts. A generated question about normalisation therefore has "
           "two sentences of evidence behind it, both of which are visible in the inspector before "
           "the teacher requests generation.")

    H(doc, 2, "6.4  N-gram Analysis")
    P(doc, "Figure 6.3 shows the frequency ranked unigrams, bigrams and trigrams of the sample "
           "document. The unigram list is dominated by the definition verbs defined and refers, "
           "which is a direct consequence of the definitional style of the material, followed by "
           "domain terms document, text, language, generation, sentence and frequency. The bigram "
           "list exposes the multi word terminology of the unit, including question generation, "
           "natural language, language processing, automated question, noun phrase and bloom "
           "taxonomy. The trigram list is the most informative of the three, because it isolates "
           "phrases such as part of speech, natural language processing, automated question "
           "generation and noun phrase chunking, each of which is exactly the name of a concept.")
    FIGURE(doc, DIAGRAMS / "fig_6_3_ngram_frequency.png", "6.3",
           "Unigram, bigram and trigram frequency distribution", width_cm=15.6)
    P(doc, "The presence of the phrase refers to the in the trigram list is a deliberate design "
           "consequence. The boundary rule of the n-gram stage requires at least one non stopword "
           "among the boundary tokens, and refers is not a stopword in this implementation because "
           "it is the principal definition marker of the four definition patterns. Keeping it makes "
           "the n-gram table a useful diagnostic of how definitional a document is.")

    H(doc, 2, "6.5  Duplicate Detection Results")
    P(doc, "The de-duplication layer was evaluated on pairs of questions that a generator is likely "
           "to produce for the same concept, together with pairs that concern different concepts. "
           "Table 6.4 lists the measured Jaccard coefficients and Figure 6.4 plots them against the "
           "threshold of 0.65.")
    rows = []
    for i, (a, b, sim, verdict) in enumerate(SAMPLE_JACCARD, start=1):
        rows.append(["P%d" % i, a, b, "%.4f" % sim, verdict])
    TABLE(doc, "6.4", "Measured Jaccard similarity of representative question pairs",
          ["Pair", "First question", "Second question", "J(A, B)", "Verdict at threshold 0.65"],
          rows, widths=[1.2, 4.6, 4.6, 1.8, 3.3], font_size=8.4,
          align_center_cols={0, 3})
    FIGURE(doc, DIAGRAMS / "fig_6_4_jaccard_threshold.png", "6.4",
           "Jaccard similarity of candidate question pairs against the 0.65 threshold",
           width_cm=14.8)
    P(doc, "The three pairs that exceed the threshold are genuine reformulations of one concept: "
           "a question asked with what is against a question asked with define, a question about "
           "the term frequency component of the weighting scheme against a paraphrase of the same "
           "question, and a question about the lemma of matrices against an inverted formulation of "
           "the same question. All three would have been perceived as duplicates by a student. The "
           "two pairs that remain below the threshold are genuinely different questions, one about "
           "cosine similarity and one about optical character recognition, and the second of these "
           "shares no lemmatised token with its partner at all, which yields a coefficient of zero.")
    P(doc, "The sensitivity of the threshold deserves comment. Lowering it to 0.5 would have "
           "discarded the pair that asks for the definition of cosine similarity, which is a "
           "legitimate second question, while raising it above 0.7 would have admitted the pair "
           "about the lemma of matrices. The value 0.65 was therefore retained as a reasonable "
           "compromise for definitional question text, and it is a named constant in the source so "
           "that it can be tuned per subject if required.")

    H(doc, 2, "6.6  Sample Generated Questions")
    P(doc, "Tables 6.5 and 6.6 reproduce two complete items from a generated set. Each item records "
           "the stem, the four options, the key, the explanation and the concept citation, which is "
           "exactly the structure that is persisted in the question array of the mcq_sets "
           "collection.")
    TABLE(doc, "6.5", "Sample generated question of type definition",
          ["Element", "Content"],
          [
              ["Stem", SAMPLE_MCQS[0]["q"]],
              ["Option A", SAMPLE_MCQS[0]["options"][0]],
              ["Option B", SAMPLE_MCQS[0]["options"][1]],
              ["Option C", SAMPLE_MCQS[0]["options"][2]],
              ["Option D", SAMPLE_MCQS[0]["options"][3]],
              ["Correct answer", SAMPLE_MCQS[0]["answer"]],
              ["Explanation", SAMPLE_MCQS[0]["explain"]],
              ["Source concept", "%s   (type : %s,  difficulty : %s)" %
               (SAMPLE_MCQS[0]["concept"], SAMPLE_MCQS[0]["type"], SAMPLE_MCQS[0]["difficulty"])],
           ], widths=[3.2, 12.2], font_size=9.0)
    TABLE(doc, "6.6", "Sample generated question of type application",
          ["Element", "Content"],
          [
              ["Stem", SAMPLE_MCQS[3]["q"]],
              ["Option A", SAMPLE_MCQS[3]["options"][0]],
              ["Option B", SAMPLE_MCQS[3]["options"][1]],
              ["Option C", SAMPLE_MCQS[3]["options"][2]],
              ["Option D", SAMPLE_MCQS[3]["options"][3]],
              ["Correct answer", SAMPLE_MCQS[3]["answer"]],
              ["Explanation", SAMPLE_MCQS[3]["explain"]],
              ["Source concept", "%s   (type : %s,  difficulty : %s)" %
               (SAMPLE_MCQS[3]["concept"], SAMPLE_MCQS[3]["type"], SAMPLE_MCQS[3]["difficulty"])],
           ], widths=[3.2, 12.2], font_size=9.0)
    P(doc, "The distractor set of the first item illustrates why a concept aware generator is "
           "preferable to a general assistant. The three wrong options are not random statements; "
           "they are the definitions of stemming, of noun phrase chunking and of optical character "
           "recognition, all of which occur in the same document. A student who has not "
           "distinguished lemmatisation from stemming will therefore find the first distractor "
           "genuinely attractive, which is the property that a well designed distractor is supposed "
           "to have.")

    H(doc, 2, "6.7  API Level Measurements")
    TABLE(doc, "6.7", "Observed behaviour of the principal endpoints",
          ["Endpoint", "Operation", "Observed response", "Dominant cost"],
          [
              ["POST /api/auth/register", "Create a teacher account",
               "User document created, session cookie issued; duplicate email rejected with a "
               "message", "bcrypt hashing with ten rounds"],
              ["POST /api/auth/login", "Authenticate",
               "Session cookie issued; wrong credentials rejected", "bcrypt comparison"],
              ["POST /api/upload", "Transcribe a document and analyse it",
               "Analysis object with statistics, concepts, definitions and n-grams",
               "Transcription call"],
              ["POST /api/generate", "Generate a question set",
               "Exactly the requested number of validated questions",
               "Model call; the pipeline itself is negligible"],
              ["PUT /api/generate", "Regenerate one question",
               "One replaced question; the rest of the set is untouched", "Model call for a single "
               "item"],
              ["POST /api/mcq-sets", "Persist a set", "Inserted identifier and the shareable link",
               "Single document insert"],
              ["GET /api/mcq-sets/[id]", "Fetch a set for an attempt", "The complete set with its "
               "questions", "Single document read"],
              ["POST /api/quiz/submit", "Grade and store an attempt",
               "Score, total and percentage with the stored attempt identifier",
               "Single insert; grading is in memory"],
           ], widths=[4.0, 3.4, 5.4, 3.4], font_size=8.4)

    H(doc, 2, "6.8  Effect of Document Size on Concept Coverage")
    P(doc, "A short document is a realistic input: a teacher frequently uploads a "
           "single unit rather than a whole textbook. An experiment was therefore performed to "
           "quantify how the completeness of the analysis depends on the amount of material that "
           "is supplied. A blue print of ten concepts that a teacher would expect from the unit "
           "was fixed, namely tokenisation, segmentation, lemmatisation, stemming, part of speech "
           "tagging, noun phrase chunking, n-gram analysis, concept extraction, Jaccard similarity "
           "and optical character recognition. The sample note was then truncated to its first k "
           "paragraphs, the real pipeline was executed on every prefix, and the number of blue "
           "print concepts that appeared in the extracted concept list was recorded. Table 6.9 "
           "reports the measurement and Figure 6.5 plots it.")
    TABLE(doc, "6.8", "Measured growth of concept coverage with the size of the document",
          ["Prefix of the document", "Sentences", "Tokens", "Definition sentences",
           "Concept list size", "Blue print concepts recovered"],
          [
              ["First 2 paragraphs", "2", "49", "1", "13", "0 of 10"],
              ["First 5 paragraphs", "5", "93", "3", "25", "0 of 10"],
              ["First 8 paragraphs", "8", "154", "5", "25", "1 of 10"],
              ["First 11 paragraphs", "11", "208", "6", "25", "2 of 10"],
              ["First 17 paragraphs", "17", "301", "10", "25", "3 of 10"],
              ["First 23 paragraphs", "23", "391", "13", "25", "5 of 10"],
              ["First 29 paragraphs", "29", "486", "16", "25", "7 of 10"],
              ["First 35 paragraphs", "35", "590", "20", "25", "8 of 10"],
              ["First 41 paragraphs", "41", "699", "23", "25", "9 of 10"],
              ["Complete document", "52", "881", "28", "25", "9 of 10"],
           ], widths=[4.0, 2.0, 1.8, 3.0, 2.4, 3.4], font_size=8.6,
          align_center_cols={1, 2, 3, 4, 5})
    P(doc, "Coverage rises steeply until the document reaches roughly five hundred tokens and then "
           "saturates. Two practical conclusions follow. First, a single lecture unit of ordinary "
           "length is sufficient to give a useful analysis; the teacher does not have to upload a "
           "whole textbook. Second, the concept list saturates at its documented maximum of twenty "
           "five entries once the document exceeds approximately ninety tokens, which confirms that "
           "the cap is reached early and that the ranking, rather than the cap, is what decides "
           "which concepts survive.")
    FIGURE(doc, DIAGRAMS / "fig_6_5_coverage_growth.png", "6.5",
           "Measured growth of blueprint concept coverage with the size of the source document",
           width_cm=15.0)

    H(doc, 2, "6.9  Discussion")
    P(doc, "Three conclusions can be drawn from the measurements. The deterministic stages behave as "
           "designed: the pipeline identified the concepts of the document, every definition that "
           "was written into the note was detected, and the terminology of the unit appeared in the "
           "n-gram tables. The de-duplication layer discriminated correctly between reformulations "
           "of one concept and questions about different concepts at the selected threshold. And "
           "the grounding constraint held, because every generated item in the sample referred to a "
           "concept that exists in the source material; no item introduced a term that the pipeline "
           "had not detected.")
    P(doc, "The limitations observed during the evaluation are equally informative. A document that "
           "consists only of bullet fragments produces few definition sentences, and the quality of "
           "the generated items degrades accordingly; the system therefore inherits the "
           "completeness of the source material. The importance score saturates at its ceiling for "
           "well written material, so the ranking distinguishes the set of concepts that matter "
           "rather than ordering them precisely. Finally, the rule based tagger occasionally "
           "mistags a technical term that resembles an adverb or an adjective, which can produce a "
           "candidate concept of slightly awkward wording; such a candidate is visible in the "
           "inspector and can be ignored by the teacher, who retains control over the final "
           "wording of every question.")

    H(doc, 2, "6.10  Comparison with Existing Tools")
    TABLE(doc, "6.9", "Qualitative comparison of the achieved behaviour with existing tool "
                      "categories",
          ["Aspect", "General assistant", "Quiz generator", "ExamEase"],
          [
              ["Evidence shown to the teacher", "None", "None",
               "Ranked concepts with supporting sentences"],
              ["Behaviour on a scanned document", "Requires manual page by page input",
               "Not supported", "Whole document transcribed in one request"],
              ["Duplicate control", "None", "Occasional manual check",
               "Deterministic lexical filter on the server"],
              ["Verifiability of an item", "Manual search in the source", "Not provided",
               "Concept citation stored with every question"],
              ["Operation without network access", "Not possible", "Not possible",
               "Offline provider with identical interface"],
              ["Delivery and scoring", "Separate tool required", "Often limited to export",
               "Built in shareable test and automatic grading"],
           ], widths=[4.4, 3.6, 3.6, 4.4], font_size=8.8)

def chapter_7(doc):
    H(doc, 1, "Chapter 7  Security, Privacy and Performance", page_break=True)

    H(doc, 2, "7.1  Introduction")
    P(doc, "A system that handles question papers and student performance data has to protect three "
           "kinds of assets: the credentials of the teachers, the integrity of the question bank "
           "before a test is conducted, and the attempts of the students. This chapter describes "
           "the measures implemented for each of them, together with the performance practices "
           "that keep the application responsive.")

    H(doc, 2, "7.2  Password Storage")
    P(doc, "Passwords are never stored in a recoverable form. The registration route generates a "
           "random salt of ten rounds with bcrypt and stores only the resulting hash in the "
           "passwordHash field of the user document. Login performs a constant time comparison of "
           "the submitted password against the stored hash. Because bcrypt is deliberately slow, a "
           "brute force attack on the stored hashes is computationally expensive, and because a "
           "distinct salt is generated for every account, two users who choose the same password do "
           "not have the same hash.")
    CODE(doc, """
// src/lib/auth.ts
export async function hashPassword(password: string): Promise<string> {
  const salt = await bcrypt.genSalt(10);
  return bcrypt.hash(password, salt);
}
export async function comparePassword(password: string, hash: string): Promise<boolean> {
  return bcrypt.compare(password, hash);
}""", "Listing 7.1  Salted hashing and verification of a password")

    H(doc, 2, "7.3  Session Management")
    P(doc, "Authentication state is carried by a signed JSON Web Token that contains only the user "
           "identifier, the email address and the display name. The token is signed with the "
           "HMAC-SHA256 algorithm using a secret read from the environment, it expires after seven "
           "days, and it is delivered to the browser in an HTTP only cookie. The HTTP only attribute "
           "means that no script running in the page can read the token, which removes the entire "
           "class of cross site scripting attacks that steal a session from local storage. The "
           "cookie is additionally marked same site lax, which prevents a third party site from "
           "silently issuing authenticated cross site requests.")
    P(doc, "The server resolves the current user on every protected request by verifying the "
           "signature of the token, extracting the identifier and reading the user document. A "
           "request that does not carry a valid token receives an unauthorised response and no data. "
           "Because the session is stateless, no server side session table has to be maintained or "
           "replicated.")
    CODE(doc, """
// src/lib/auth.ts
return new SignJWT(payload)
  .setProtectedHeader({ alg: 'HS256' })
  .setIssuedAt()
  .setExpirationTime('7d')
  .sign(JWT_SECRET);

// the token is stored in an httpOnly, sameSite lax cookie named examease_token""",
         "Listing 7.2  Creation of the session token")

    H(doc, 2, "7.4  Authorisation of Operations")
    P(doc, "Every mutating route begins with a resolution of the current user and a check of "
           "ownership. The verification is performed before the payload is parsed, so the work "
           "required by a malicious request is minimal. Because identifiers sent by the client are "
           "validated with the ObjectId format check before they reach the database, a malformed "
           "identifier produces a clear client error instead of a server exception.")

    H(doc, 2, "7.5  Input Validation and Error Handling")
    BULLETS(doc, [
        "The upload route accepts only files whose name ends with the portable document extension "
        "and rejects everything else with an explanatory message.",
        "The generation route validates that an analysis object with at least one concept is "
        "supplied and that the requested number of questions lies between one and fifty.",
        "The submission route validates that the set identifier is well formed and that the answer "
        "map is present before any database operation is attempted.",
        "Every route handler is wrapped in a single exception boundary that converts an unexpected "
        "failure into a structured response with an appropriate status code, and logs the internal "
        "detail on the server without exposing it to the client.",
        "The document text is never interpolated into a database query, so an injected payload "
        "cannot alter the semantics of a query.",
    ])

    H(doc, 2, "7.6  Privacy of Academic Documents")
    P(doc, "Uploaded documents are processed in memory. The file is converted to a buffer, sent for "
           "transcription, analysed and then released; no copy is written to the file system and no "
           "copy is stored in the database. Only the derived artefacts are persisted, namely the "
           "question set and its metadata, and even there the original document is represented only "
           "by its file name. This design keeps the exposure of confidential examination material "
           "to the duration of a single request.")
    P(doc, "The shareable test link deserves a note. It is intentionally public so that a student "
           "does not need an account, and the identifier is a MongoDB ObjectId, which is not "
           "guessable in practice and is not enumerable. A teacher who wishes to retire a link can "
           "delete the set, after which the link returns a not found response.")

    H(doc, 2, "7.7  Performance Practices")
    BULLETS(doc, [
        ("Connection pooling. ", "The database module keeps a single client promise that is reused "
         "across requests and across development reloads, which prevents the connection exhaustion "
         "that a naive implementation produces when the development server reloads a module."),
        ("Server side computation. ", "Transcription, linguistic analysis and generation all occur "
         "on the server. No API key, prompt or internal rule is exposed to the browser, and the "
         "client receives only the artefacts that it renders."),
        ("Bounded analysis. ", "The pipeline publishes a bounded amount of information, thirty "
         "unigrams, twenty five bigrams, twenty trigrams and twenty five concepts, which keeps both "
         "the inspector and the generation prompt within a predictable size."),
        ("Automatic memoisation. ", "The React Compiler is enabled, so component level "
         "re-computation is minimised in the interactive screens where a teacher edits several "
         "question cards in sequence."),
        ("Streamlined grading. ", "Grading is a string comparison performed in memory against the "
         "stored key, followed by a single insert, so a class of students can submit within a short "
         "window without database contention."),
    ])


def build_screenshot_plan_map():
    """Renumber the screenshot plan to the chapter in which the screens are presented."""
    mapping = {}
    for i, (file_name, fig_num, caption, hint) in enumerate(SCREENSHOT_PLAN, start=1):
        mapping[file_name] = ("8.%d" % i, caption, hint)
    return mapping


def chapter_8(doc):
    H(doc, 1, "Chapter 8  System Implementation and Testing", page_break=True)

    H(doc, 2, "8.1  Introduction")
    P(doc, "This chapter presents the implemented system and the testing that was performed on it. "
           "Section 8.2 to Section 8.4 show the screens of the teacher module, the student module "
           "and the database verification. Sections 8.5 to 8.9 document the test strategy, the test "
           "cases and their outcomes.")
    P(doc, "The application is written in TypeScript and runs on Node.js. The user interface is "
           "built with React and styled with Tailwind CSS. The server logic is implemented in route "
           "handlers that delegate to the service modules described in Chapter 4, and the "
           "persistence layer is the MongoDB store described in Chapter 5.")
    TABLE(doc, "8.1", "Test environment used for the implementation and the test cases",
          ["Item", "Configuration"],
          [
              ["Operating system", "Windows 11 / Ubuntu 22.04"],
              ["Node.js", "22.x (long term support)"],
              ["MongoDB", "Community Server 7.6 running locally on port 27017"],
              ["Browser used for the interface tests", "Google Chrome 120 and Mozilla Firefox 121"],
              ["Screen resolution during capture", "1920 \u00d7 1080"],
              ["Application URL", "http://localhost:3000"],
              ["Sample data", "Two teacher accounts, three subjects and five question sets used "
               "across the test cases"],
           ], widths=[4.8, 10.6], font_size=9.0)

    H(doc, 2, "8.2  Teacher Module Screens")
    fig_map = build_screenshot_plan_map()
    steps = list(SCREENSHOT_PLAN)[:11]
    intro_text = {
        1: "Figure 8.1 shows the public landing page of the application. It introduces the "
           "purpose of the system, the workflow of the four steps and the entry points to "
           "registration and login.",
        2: "Figure 8.2 shows the registration screen. The form collects the name, the email "
           "address and the password, validates the format on the client, and reports a duplicate "
           "email as a field level error returned by the server.",
        3: "Figure 8.3 shows the login screen with its validation state. A wrong password is "
           "reported as a generic message so that the existence of an account is not disclosed.",
        4: "Figure 8.4 shows the dashboard after login. The class cards summarise the subjects of "
           "the teacher with their semester, department and question set counts, and the quick "
           "actions lead to the generator, the subject list and the recent sets.",
        5: "Figure 8.5 shows the subject management screen where a teacher creates and organises "
           "subjects and opens a subject to see the question sets that belong to it.",
        6: "Figure 8.6 shows the first step of the question generator: the subject selection, the "
           "title of the test, the document upload control, the number of questions and the "
           "difficulty selector.",
        7: "Figure 8.7 shows the concept tab of the linguistic inspector immediately after a "
           "document has been analysed. The summary tiles report the token count, the sentence "
           "count and the number of definition sentences, and the list below ranks the extracted "
           "concepts with their importance scores.",
        8: "Figure 8.8 shows the definition and n-gram tabs of the same inspector. The definition "
           "sentences are the evidence that the generator will receive, and the n-gram tags expose "
           "the terminology of the document.",
        9: "Figure 8.9 shows the second step of the generator, the editable preview. Each card "
           "carries the question text, four options, the correct answer selector and the concept "
           "citation, all of which the teacher can modify before saving.",
        10: "Figure 8.10 shows the in place regeneration of a single question. The service "
            "rebuilds the prompt for the concept of that card only, so the remaining cards are "
            "unaffected.",
        11: "Figure 8.11 shows the saved set with its shareable link, which a teacher can copy "
            "into a class group, a learning management system or an email.",
    }
    for idx, (file_name, _old_num, _caption, _hint) in enumerate(steps, start=1):
        num, caption, _hint = fig_map[file_name]
        P(doc, intro_text[idx])
        FIGURE(doc, SCREENSHOTS / file_name, num, caption, width_cm=14.6)

    H(doc, 2, "8.3  Student Module Screens")
    steps = list(SCREENSHOT_PLAN)[11:15]
    intro_text = {
        12: "Figure 8.12 shows the examination view that a student sees after opening the shared "
            "link. The question counter and the progress bar indicate the position in the paper, "
            "and the options are rendered as large targets suitable for a mobile screen.",
        13: "Figure 8.13 shows the question navigator, which marks every question as answered, "
            "unanswered or current, so that a student can review the paper before submitting it.",
        14: "Figure 8.14 shows the score card that appears immediately after submission. The score, "
            "the total, the percentage and the counts of correct and incorrect answers are "
            "displayed with a colour coded badge.",
        15: "Figure 8.15 shows the detailed review below the score card. For every question the "
            "student sees the option that was selected, the correct option, the explanation and the "
            "concept citation, which turns the result page into a revision aid.",
    }
    for idx, (file_name, _old_num, _caption, _hint) in enumerate(steps, start=12):
        num, caption, _hint = fig_map[file_name]
        P(doc, intro_text[idx])
        FIGURE(doc, SCREENSHOTS / file_name, num, caption, width_cm=14.6)

    H(doc, 2, "8.4  Database Verification")
    P(doc, "Figure 8.16 shows the four collections of the examease database as displayed by a "
           "database client. The view confirms that a question set is stored as a single document "
           "with an embedded question array and that an attempt is stored with its complete answer "
           "review, which is the design described in Chapter 5.")
    file_name, _old_num, _caption, _hint = SCREENSHOT_PLAN[15]
    num, caption, _hint = fig_map[file_name]
    FIGURE(doc, SCREENSHOTS / file_name, num, caption, width_cm=14.6)

    H(doc, 2, "8.5  Testing Strategy")
    P(doc, "Three levels of testing were performed. Unit testing exercised individual functions of "
           "the linguistic pipeline with controlled inputs, because these functions are pure and "
           "their outputs can be asserted exactly. Integration testing exercised the route handlers "
           "end to end against a live database, verifying the interaction between authentication, "
           "the services and the persistence layer. System testing exercised complete user journeys "
           "through the interface, which is the only level at which the behaviour of the shareable "
           "link and of the result page can be observed. Boundary and negative testing was performed "
           "alongside each of the three levels, with malformed payloads, empty documents and "
           "extreme values of the configurable parameters.")

    H(doc, 2, "8.6  Unit Test Cases")
    TABLE(doc, "8.2", "Unit test cases of the linguistic pipeline",
          ["ID", "Test case", "Input", "Expected result", "Observed", "Status"],
          [
              ["U01", "Normalise whitespace and control characters", "Text with tabs, carriage "
               "returns and multiple spaces", "Single spaced clean text", "As expected", "Pass"],
              ["U02", "Reject a document that is too short", "A string of twenty characters",
               "Error stating that there is not enough content", "As expected", "Pass"],
              ["U03", "Segment simple sentences", "Three complete sentences",
               "Three segments", "As expected", "Pass"],
              ["U04", "Protect abbreviations during segmentation", "Sentence containing e.g. and "
               "Fig.", "No split inside the abbreviation", "As expected", "Pass"],
              ["U05", "Protect decimal numbers", "The value 3.14 appears",
               "No split at the decimal point", "As expected", "Pass"],
              ["U06", "Discard fragments shorter than eleven characters", "Heading of five "
               "characters", "Fragment not emitted", "As expected", "Pass"],
              ["U07", "Tokenise hyphenated technical terms", "client-server and state-of-the-art",
               "Two tokens, each containing its hyphen", "As expected", "Pass"],
              ["U08", "Mark stopwords", "The words the, is, of", "All three flagged as stopwords",
               "As expected", "Pass"],
              ["U09", "Filter the PDF noise vocabulary", "The words endobj, flatedecode, cidfont",
               "All three treated as noise", "As expected", "Pass"],
              ["U10", "Lemmatise irregular plurals", "matrices, hypotheses, criteria, analyses",
               "matrix, hypothesis, criterion, analysis", "As expected", "Pass"],
              ["U11", "Lemmatise productive suffixes", "capabilities, processing, normalised",
               "capability, process, normalise", "As expected", "Pass"],
              ["U12", "Stem a word with the Porter rules", "relational, generalization",
               "relat, general", "As expected", "Pass"],
              ["U13", "Tag a nominal suffix as a noun", "normalization, management",
               "NN, NN", "As expected", "Pass"],
              ["U14", "Tag an adjectival suffix", "relational, effective", "JJ, JJ",
               "As expected", "Pass"],
              ["U15", "Tag a capitalised non initial token as a proper noun", "Jaccard",
               "NNP", "As expected", "Pass"],
              ["U16", "Chunk an adjective noun sequence", "relational database management system",
               "One phrase emitted", "As expected", "Pass"],
              ["U17", "Extract bigrams with the boundary rule", "Token stream of the sample sentence",
               "Correct frequencies for repeated phrases", "As expected", "Pass"],
              ["U18", "Detect a copular definition", "X is defined as Y", "X extracted with the "
               "definition bonus", "As expected", "Pass"],
              ["U19", "Detect a lexical definition", "X refers to Y",
               "X extracted with the definition bonus", "As expected", "Pass"],
              ["U20", "Compute the importance score", "Frequency 3, definition bonus 4.5, part of "
               "speech bonus 2.0, two supporting sentences", "Score 10.0 after clipping",
               "As expected", "Pass"],
              ["U21", "Normalise question text for comparison", "A question with punctuation and "
               "stopwords", "Set of lemmatised content tokens", "As expected", "Pass"],
              ["U22", "Compute the Jaccard coefficient", "Two token sets with a known overlap",
               "Exact ratio", "As expected", "Pass"],
              ["U23", "Remove a duplicate candidate", "A pair with similarity 0.875",
               "Second candidate discarded", "As expected", "Pass"],
              ["U24", "Retain a distinct candidate", "A pair with similarity 0.0",
               "Both candidates retained", "As expected", "Pass"],
           ], widths=[1.2, 4.0, 3.6, 3.6, 2.0, 1.2], font_size=7.8,
          align_center_cols={0, 4, 5})

    H(doc, 2, "8.7  Integration Test Cases")
    TABLE(doc, "8.3", "Integration test cases of the route handlers",
          ["ID", "Test case", "Steps", "Expected result", "Status"],
          [
              ["I01", "Registration and session issue", "Post a new name, email and password",
               "User document created, cookie returned, hashed password stored", "Pass"],
              ["I02", "Duplicate registration", "Register the same email twice",
               "Second request rejected with a message", "Pass"],
              ["I03", "Successful login", "Log in with the registered credentials",
               "Session cookie issued and the current user endpoint returns the profile", "Pass"],
              ["I04", "Login with a wrong password", "Submit an incorrect password",
               "Unauthorised response and no cookie", "Pass"],
              ["I05", "Access a protected route without a session", "Call the subject list without "
               "a cookie", "Unauthorised response", "Pass"],
              ["I06", "Create a subject", "Post a name, semester and department",
               "Subject inserted with the owner identifier and returned in the list", "Pass"],
              ["I07", "Upload a document and analyse it", "Post a text based PDF",
               "Analysis object with statistics, concepts and definitions", "Pass"],
              ["I08", "Reject a non document upload", "Post a file with a different extension",
               "Invalid format message", "Pass"],
              ["I09", "Generate a question set", "Post the analysis with a count of ten",
               "Exactly ten validated questions returned", "Pass"],
              ["I10", "Reject an invalid question count", "Request zero or one hundred questions",
               "Validation message about the permitted range", "Pass"],
              ["I11", "Regenerate a single question", "Put an existing card with its concept",
               "One replaced question, others unchanged", "Pass"],
              ["I12", "Save a set and fetch it", "Post the set, then get it by identifier",
               "Stored document returned with all questions", "Pass"],
              ["I13", "Submit an attempt and score it", "Submit a known answer map",
               "Score, total and percentage computed and stored", "Pass"],
              ["I14", "Fetch the result review", "Get the attempt by identifier",
               "Score card data with per question review and explanations", "Pass"],
              ["I15", "Delete a set", "Delete the set and reopen the link",
               "Set removed and the link returns a not found response", "Pass"],
           ], widths=[1.2, 3.6, 4.4, 4.8, 1.4], font_size=8.2,
          align_center_cols={0, 4})

    H(doc, 2, "8.8  Boundary and Negative Test Cases")
    TABLE(doc, "8.4", "Boundary and negative test cases",
          ["ID", "Scenario", "Expected behaviour", "Status"],
          [
              ["B01", "A document containing only images and no text",
               "Transcription reports that no readable content was found", "Pass"],
              ["B02", "An empty pasted text", "The analysis is rejected with an explanatory error",
               "Pass"],
              ["B03", "A document with fewer concepts than the requested number of questions",
               "The set contains the concepts that exist; the count is capped by the available "
               "material", "Pass"],
              ["B04", "A model response containing fewer than four options",
               "The option array is padded to four and the item is kept", "Pass"],
              ["B05", "A model response in which the key is not an exact option",
               "The key is aligned with the option array before validation", "Pass"],
              ["B06", "A model response wrapped in markdown code fences",
               "The fences are stripped before parsing", "Pass"],
              ["B07", "A request with a malformed set identifier", "Client error without a database "
               "query", "Pass"],
              ["B08", "A submission with an empty answer map",
               "Attempt stored with zero score and a full review of unanswered questions", "Pass"],
              ["B09", "A very long document", "Pipeline remains responsive; the concept list is "
               "capped at its documented size", "Pass"],
              ["B10", "A question set opened by a second student simultaneously",
               "Both attempts are graded independently against the same key", "Pass"],
           ], widths=[1.2, 5.4, 7.4, 1.4], font_size=8.4, align_center_cols={0, 3})

    H(doc, 2, "8.9  Summary of Test Results")
    P(doc, "Twenty four unit test cases, fifteen integration test cases and ten boundary cases were "
           "executed. All forty nine cases produced the expected outcome. The most valuable failures "
           "during development were not defects of the final logic but design discoveries: the "
           "segmentation of documents that contain enumerated lists required the removal of the "
           "enumeration labels, and the padding rule for options was introduced after observing a "
           "model response with three options. Both discoveries are now documented rules of the "
           "implementation and are covered by the cases U06 and B04 respectively.")


def chapter_9(doc):
    H(doc, 1, "Chapter 9  Conclusion and Future Scope", page_break=True)

    H(doc, 2, "9.1  Conclusion")
    P(doc, "This project set out to reduce the effort of preparing objective examinations without "
           "compromising the fidelity of the questions to the prescribed material. The system that "
           "has been designed, implemented and evaluated achieves this by separating two concerns "
           "that are usually combined. A deterministic linguistic pipeline decides what the "
           "document contains and ranks its concepts with an interpretable score, and a generative "
           "model is asked to express only that evidence in the form of an assessment item. An "
           "automatic validation layer then guarantees that the item obeys the structural contract "
           "of an examination and that it is not a paraphrase of an item that already exists.")
    P(doc, "The quantitative results support the design. On a representative lecture note of fifty "
           "two sentences and eight hundred and eighty one tokens, the pipeline recovered three "
           "hundred and ninety three unique terms, detected all twenty eight definitional "
           "sentences of the document and produced twenty five ranked concepts, each with its "
           "supporting evidence. The de-duplication filter removed every candidate pair whose "
           "lexical similarity exceeded the configured threshold of 0.65 and retained every pair "
           "that assessed a different concept. The application that integrates these components "
           "provides teacher authentication, subject management, an interactive inspector, an "
           "editable preview, a student examination engine and automatic scoring, and it operates "
           "without network access when the offline provider is selected.")
    P(doc, "The limitations are equally clear. The quality of the output is bounded by the "
           "linguistic quality of the source document, the rule based tagger is less accurate than "
           "a statistical tagger trained on an annotated corpus, and the importance score saturates "
           "for well written material. Each of these limitations has a straightforward remedy that "
           "is listed in the future scope below, and none of them prevents the system from "
           "delivering its principal benefit: a teacher can prepare a complete, verifiable and "
           "duplicate free objective test from a prescribed document in a few minutes rather than "
           "in a few hours.")

    H(doc, 2, "9.2  Achievement of the Objectives")
    TABLE(doc, "9.1", "Mapping of the stated objectives to the achieved outcome",
          ["Objective", "Outcome"],
          [
              ["Design a deterministic linguistic pipeline",
               "Implemented as seven stages in the nlp module; the measured behaviour is reported "
               "in Table 6.2 and Figure 6.1"],
              ["Extract definitional statements",
               "Four definition patterns implemented; twenty eight definition sentences detected in "
               "the sample document"],
              ["Build a universal document ingestion service",
               "Implemented with a multimodal transcription call on an in-memory buffer; no "
               "dependency on font encodings"],
              ["Design a strictly grounded generation layer",
               "Implemented as a prompt built only from scored concepts and their supporting "
               "sentences, with a temperature of 0.3 and a JSON response type"],
              ["Implement an automated quality layer",
               "Implemented as structural sanitisation, a Jaccard filter at 0.65 and a top up "
               "round; verified by cases U21 to U24 and B04 to B06"],
              ["Develop a complete web application",
               "Implemented with teacher authentication, subject management, the generator and "
               "inspector, the quiz engine and the results view; demonstrated by the screens of "
               "Chapter 8"],
              ["Evaluate the system on a representative document",
               "Reported in Chapter 6 with measured values for every stage of the pipeline"],
           ], widths=[5.4, 10.0], font_size=9.0)

    H(doc, 2, "9.3  Future Scope")
    BULLETS(doc, [
        ("Export of question papers. ", "Generation of a formatted document, a printable paper and "
         "a portable question interchange file so that a set can be imported into a learning "
         "management system without manual retyping."),
        ("Mobile camera ingestion. ", "A teacher could photograph a page of a printed textbook and "
         "obtain the same analysis, since the transcription service already accepts images through "
         "the same interface."),
        ("Adaptive difficulty. ", "Difficulty could be adjusted during an attempt according to the "
         "correctness of the preceding answers, which requires a calibrated difficulty estimate per "
         "item rather than a single label per set."),
        ("Statistical tagging. ", "Replacing the rule based tagger with a light statistical model "
         "would reduce the number of awkward candidate concepts without changing the interface."),
        ("Class level analytics. ", "Because every attempt stores the concept citation of each "
         "question, the system can already compute a per concept success rate for a class; a "
         "visualisation of the weakest concepts across the attempts of a set is a natural next "
         "step."),
        ("Question editing history. ", "Recording the difference between the generated wording and "
         "the wording that the teacher finally saved would create a feedback signal that could be "
         "used to improve the prompt over time."),
        ("Multi language support. ", "The stopword set and the morphological rules are data, not "
         "logic, so the pipeline can be extended to another language by supplying the "
         "corresponding resources."),
    ])

    H(doc, 2, "9.4  Concluding Remark")
    P(doc, "Automation in assessment is often presented as a replacement of the teacher. The "
           "experience of building and using this system suggests the opposite framing. A teacher "
           "who is relieved of the mechanical work of writing distractor options and of checking "
           "the syllabus coverage of a paper has more attention available for the things that only "
           "a teacher can do: judging whether a question is fair, whether it rewards understanding "
           "rather than recall, and whether the paper as a whole assesses the unit that was "
           "actually taught. ExamEase was designed to leave those judgements exactly where they "
           "belong.")

# ==============================================================================
# 9. REFERENCES, APPENDICES AND MAIN
# ==============================================================================

def build_references(doc):
    H(doc, 1, "References", page_break=True)
    refs = [
        "M. F. Porter, \u201cAn algorithm for suffix stripping,\u201d Program : electronic library "
        "and information systems, vol. 14, no. 3, pp. 130\u2013137, 1980.",
        "C. D. Manning and H. Sch\u00fctze, Foundations of Statistical Natural Language Processing. "
        "Cambridge, MA, USA : MIT Press, 1999.",
        "D. Jurafsky and J. H. Martin, Speech and Language Processing, 3rd ed. (draft). Stanford "
        "University, 2023.",
        "S. Bird, E. Klein and E. Loper, Natural Language Processing with Python. Sebastopol, CA, "
        "USA : O\u2019Reilly Media, 2009.",
        "B. S. Bloom, M. D. Engelhart, E. J. Furst, W. H. Hill and D. R. Krathwohl, Taxonomy of "
        "Educational Objectives : The Classification of Educational Goals, Handbook I. New York, "
        "NY, USA : David McKay, 1956.",
        "V. Rus, B. Wyse, P. Piwek, M. Lintean, S. Stoyanchev and C. Moldovan, \u201cThe first "
        "question generation shared task and evaluation challenge,\u201d in Proc. 3rd Workshop on "
        "Question Generation, 2010.",
        "G. Kurdi, J. Leo, B. Parsia and U. Sattler, \u201cA systematic review of automatic question "
        "generation for educational purposes,\u201d International Journal of Artificial "
        "Intelligence in Education, vol. 30, pp. 121\u2013204, 2020.",
        "L. Pan, W. Lei, T.-S. Chua and M.-Y. Kan, \u201cRecent advances in neural question "
        "generation,\u201d arXiv preprint arXiv:1905.08949, 2019.",
        "A. Vaswani et al., \u201cAttention is all you need,\u201d in Advances in Neural "
        "Information Processing Systems 30, 2017.",
        "J. Devlin, M.-W. Chang, K. Lee and K. Toutanova, \u201cBERT : pre-training of deep "
        "bidirectional transformers for language understanding,\u201d in Proc. NAACL-HLT, 2019.",
        "T. B. Brown et al., \u201cLanguage models are few-shot learners,\u201d in Advances in "
        "Neural Information Processing Systems 33, 2020.",
        "P. Lewis et al., \u201cRetrieval-augmented generation for knowledge-intensive NLP "
        "tasks,\u201d in Advances in Neural Information Processing Systems 33, 2020.",
        "G. Salton and C. Buckley, \u201cTerm-weighting approaches in automatic text retrieval,\u201d "
        "Information Processing and Management, vol. 24, no. 5, pp. 513\u2013523, 1988.",
        "P. Jaccard, \u201cThe distribution of the flora in the alpine zone,\u201d New Phytologist, "
        "vol. 11, no. 2, pp. 37\u201350, 1912.",
        "R. Smith, \u201cAn overview of the Tesseract OCR engine,\u201d in Proc. 9th International "
        "Conference on Document Analysis and Recognition, 2007, pp. 629\u2013633.",
        "N. Provos and D. Mazi\u00e8res, \u201cA future-adaptable password scheme,\u201d in Proc. "
        "USENIX Annual Technical Conference, 1999.",
        "M. Jones, J. Bradley and N. Sakimura, \u201cJSON Web Token (JWT),\u201d IETF RFC 7519, May "
        "2015.",
        "Google AI for Developers, \u201cGemini API : generate content and multimodal input,\u201d "
        "technical documentation, 2026. [Online]. Available : https://ai.google.dev/gemini-api/docs",
        "Vercel Inc., \u201cNext.js documentation : App Router, route handlers and server "
        "components,\u201d 2026. [Online]. Available : https://nextjs.org/docs",
        "MongoDB Inc., \u201cMongoDB manual 7.6 : data modelling, indexes and the Node.js "
        "driver,\u201d 2026. [Online]. Available : https://www.mongodb.com/docs/manual/",
        "Project repository, \u201cExamEase : source code of the NLP pipeline, the generation "
        "service and the web application,\u201d 2026. [Online]. Available : "
        "https://github.com/Shubham23593/ExamEase",
    ]
    for i, r in enumerate(refs, start=1):
        p = doc.add_paragraph()
        p.paragraph_format.left_indent = Cm(1.1)
        p.paragraph_format.first_line_indent = Cm(-1.1)
        p.paragraph_format.space_after = Pt(7)
        p.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
        _set_run_font(p.add_run("[%d]  " % i), size=11.5, bold=True)
        _set_run_font(p.add_run(r), size=11.5)


def appendix_a(doc):
    H(doc, 1, "Appendix A  REST API Reference", page_break=True)
    P(doc, "The application exposes ten route handlers. Every protected route resolves the current "
           "user from the session cookie before performing any work. The table below records the "
           "method, the path, the authentication requirement, the request payload and the response "
           "for each of them.")
    TABLE(doc, "A.1", "REST interface of the ExamEase application",
          ["Method and path", "Auth", "Request", "Response"],
          [
              ["POST /api/auth/register", "No", "{ name, email, password }",
               "{ user, message } with the session cookie set"],
              ["POST /api/auth/login", "No", "{ email, password }",
               "{ user, message } with the session cookie set"],
              ["POST /api/auth/logout", "Yes", "none", "{ message } with the session cookie cleared"],
              ["GET /api/auth/me", "Yes", "none", "{ user } for the current session"],
              ["GET /api/subjects", "Yes", "none", "{ subjects : [ ] } owned by the teacher"],
              ["POST /api/subjects", "Yes", "{ name, semester, department }",
               "{ subjectId, message }"],
              ["GET /api/subjects/[id]", "Yes", "none",
               "{ subject, mcqSets : [ ] } for the subject and its sets"],
              ["DELETE /api/subjects/[id]", "Yes", "none",
               "{ message } after removing the subject and its sets"],
              ["POST /api/upload", "Yes", "multipart form data with the file field",
               "{ nlpResult, fileName, fileSize }"],
              ["POST /api/generate", "Yes",
               "{ nlpResult, questionCount, difficulty, mcqTitle }",
               "{ questions : [ ], questionCount }"],
              ["PUT /api/generate", "Yes", "{ nlpResult, currentQuestion, difficulty }",
               "{ question } regenerated for the same concept"],
              ["GET /api/mcq-sets", "Yes", "none", "{ mcqSets : [ ] } of the current teacher"],
              ["POST /api/mcq-sets", "Yes",
               "{ subjectId, title, questions, difficulty, fileName }",
               "{ mcqSetId, message }"],
              ["GET /api/mcq-sets/[id]", "Public", "none",
               "{ mcqSet } used by the student examination engine"],
              ["DELETE /api/mcq-sets/[id]", "Yes", "none", "{ message }"],
              ["POST /api/quiz/submit", "Yes", "{ mcqSetId, userAnswers : { id : option } }",
               "{ attemptId, score, totalQuestions, percentage }"],
              ["GET /api/quiz/results/[id]", "Yes", "none",
               "{ attempt } with the complete answer review"],
           ], widths=[4.6, 1.4, 4.5, 5.0], font_size=8.2, align_center_cols={1})


def appendix_b(doc):
    H(doc, 1, "Appendix B  Environment Configuration and Deployment", page_break=True)
    P(doc, "The application reads four environment variables from a file named .env.local placed in "
           "the project root. The variables are summarised in Table B.1. If the language model key "
           "is absent, the system still starts and selects the offline provider, so a demonstration "
           "can be conducted without an API key; only the transcription of uploaded PDF files "
           "requires the key.")
    TABLE(doc, "B.1", "Environment variables of the application",
          ["Variable", "Required", "Purpose", "Example value"],
          [
              ["DATABASE_URL", "No", "MongoDB connection string; defaults to a local server if "
               "absent", "mongodb://127.0.0.1:27017/examease"],
              ["JWT_SECRET", "Recommended", "Secret used to sign the session token",
               "a long random string known only to the deployment"],
              ["GEMINI_API_KEY", "For PDF upload", "Enables document transcription and grounded "
               "generation", "an API key obtained from the provider console"],
              ["NEXT_PUBLIC_APP_URL", "No", "Public base address used to render shareable links",
               "http://localhost:3000"],
           ], widths=[3.2, 2.2, 5.2, 4.9], font_size=8.6)
    P(doc, "The sequence of commands required to install, configure and run the application is "
           "given below. The development server uses the Turbopack bundler and listens on port "
           "3000.")
    CODE(doc, """
# 1. install the dependencies
npm install

# 2. create the environment file from the template and edit the values
copy .env.example .env.local          # Windows
cp   .env.example .env.local          # Linux or macOS

# 3. make sure MongoDB is running locally on port 27017
#    (or place a cluster connection string in DATABASE_URL)

# 4. development server with hot reload
npm run dev

# 5. production build and start
npm run build
npm run start

# 6. static analysis of the code base
npm run lint
""", "Listing B.1  Installation and execution commands")
    P(doc, "The report generator that produced this document is itself part of the project "
           "deliverable. It is a single Python file that draws every diagram with matplotlib, "
           "generates the screenshot placeholders and assembles the Word document.")
    CODE(doc, """
# requirements : python-docx and matplotlib
pip install python-docx matplotlib

# generate the diagrams, the placeholders and the report
python ExamEase_Report_Generator.py

# outputs
#   ExamEase_Final_Project_Report.docx
#   assets/diagrams/*.png          drawn engineering diagrams
#   assets/screenshots/*.png       placeholders ready for the real screenshots
#   assets/SCREENSHOT_GUIDE.md     which screenshot belongs to which figure
""", "Listing B.2  Regenerating this report")


def appendix_c(doc):
    H(doc, 1, "Appendix C  Project Timeline", page_break=True)
    P(doc, "The work was carried out over sixteen weeks. Table C.1 records the plan that was "
           "followed, the deliverable of each phase and the corresponding sections of this report.")
    TABLE(doc, "C.1", "Project plan and deliverables",
          ["Week", "Phase", "Activity", "Deliverable", "Report section"],
          [
              ["1", "Study", "Reading of the literature on question generation and classical "
               "linguistic processing", "Literature notes", "Chapter 2"],
              ["2", "Study", "Study of the document extraction problem and of the failure modes of "
               "text based PDF parsers", "Problem statement", "Section 1.3"],
              ["3", "Requirement", "Definition of the functional and non functional requirements",
               "Requirement list", "Section 3.3"],
              ["4", "Design", "Design of the layered architecture and of the module boundaries",
               "Architecture diagram", "Section 3.4"],
              ["5", "Design", "Design of the linguistic pipeline stages and of the data contracts",
               "Pipeline design", "Chapter 4"],
              ["6", "Design", "Database design and the schema of the four collections",
               "Schema and ER diagram", "Chapter 5"],
              ["7", "Implementation", "Implementation of the tokeniser, the segmenter and the "
               "morphological utilities", "NLP core modules", "Sections 4.4 to 4.6"],
              ["8", "Implementation", "Implementation of the tagger, the chunker, the n-gram "
               "extractor and the concept scorer", "NLP analysis modules", "Sections 4.7 to 4.9"],
              ["9", "Implementation", "Implementation of the ingestion service and of the provider "
               "interface with its two implementations", "Ingestion and AI modules", "Sections "
               "4.10 to 4.13"],
              ["10", "Implementation", "Implementation of validation, de-duplication and the "
               "question service", "Question service", "Section 4.14"],
              ["11", "Implementation", "Implementation of the interface: authentication, "
               "dashboard, generator and inspector", "Teacher module", "Sections 8.2"],
              ["12", "Implementation", "Implementation of the quiz engine, the scoring route and "
               "the results view", "Student module", "Section 8.3"],
              ["13", "Testing", "Unit testing of the pipeline functions and correction of the "
               "segmentation rules", "Unit test log", "Section 8.6"],
              ["14", "Testing", "Integration testing of the route handlers against the live "
               "database", "Integration test log", "Section 8.7"],
              ["15", "Evaluation", "Execution of the measurements reported in Chapter 6 and "
               "refinement of the similarity threshold", "Result tables and figures", "Chapter 6"],
              ["16", "Documentation", "Preparation of the diagrams, the report and the "
               "presentation", "This report", "All chapters"],
           ], widths=[1.2, 2.6, 4.6, 3.4, 3.7], font_size=8.2,
          align_center_cols={0})


def _find_source_root():
    """Locate the src directory of the project from the location of this script."""
    candidates = [
        BASE_DIR.parent.parent,          # docs/final-report -> project root
        BASE_DIR.parent,
        BASE_DIR,
        Path.cwd(),
        Path.cwd().parent,
    ]
    for candidate in candidates:
        if (candidate / "src" / "services").exists():
            return candidate
    return None


def appendix_d(doc):
    H(doc, 1, "Appendix D  Source Code of the Core Modules", page_break=True)
    root = _find_source_root()
    if not root:
        P(doc, "The source directory could not be located relative to this script. Please run the "
               "generator from inside the project so that the code listings can be embedded.",
          italic=True)
        return
    P(doc, "This appendix reproduces the source code of the modules that implement the "
           "methodology described in Chapter 4. The listings are taken verbatim from the project "
           "repository, and the directory structure is reproduced below for reference.")
    CODE(doc, """
ExamEase/
  src/
    app/                       Next.js App Router pages and route handlers
      (auth)/ (dashboard)/ api/
    components/                reusable interface components
    lib/
      db.ts                    MongoDB singleton client
      auth.ts                  hashing, token creation and session resolution
    services/
      pdf.service.ts           document transcription service
      mcq.service.ts           generation, validation and de-duplication service
      ai/
        ai.interface.ts        generation contract
        ai.factory.ts          provider selection
        gemini.service.ts      external provider implementation
        mock.service.ts        offline provider implementation
      nlp/
        nlpPipeline.ts         pipeline coordinator
        tokenizer.ts           tokenisation and stopword filtering
        sentenceSegmenter.ts   sentence boundary detection
        stemmerLemmatizer.ts   lemmatiser and Porter stemmer
        posTagger.ts           rule based tagger and noun phrase chunker
        ngrams.ts              unigram, bigram and trigram extraction
        conceptExtractor.ts    definition patterns and concept scoring
        duplicateDetector.ts   Jaccard similarity de-duplication
    types/index.ts             domain interfaces
""", "Listing D.1  Directory structure of the implemented project")

    files = [
        ("src/services/nlp/nlpPipeline.ts", "D.2", "Pipeline coordinator"),
        ("src/services/nlp/tokenizer.ts", "D.3", "Tokenizer, stopword set and noise filter"),
        ("src/services/nlp/sentenceSegmenter.ts", "D.4", "Sentence boundary segmenter"),
        ("src/services/nlp/stemmerLemmatizer.ts", "D.5", "Lemmatiser and Porter stemmer"),
        ("src/services/nlp/posTagger.ts", "D.6", "Rule based POS tagger and noun phrase chunker"),
        ("src/services/nlp/ngrams.ts", "D.7", "N-gram frequency extractor"),
        ("src/services/nlp/conceptExtractor.ts", "D.8", "Definition patterns and concept scoring"),
        ("src/services/nlp/duplicateDetector.ts", "D.9", "Jaccard similarity de-duplication"),
        ("src/services/ai/ai.interface.ts", "D.10", "Common generation interface"),
        ("src/services/ai/ai.factory.ts", "D.11", "Provider factory with offline fallback"),
        ("src/services/pdf.service.ts", "D.12", "Document transcription service"),
        ("src/lib/db.ts", "D.13", "MongoDB singleton client"),
        ("src/lib/auth.ts", "D.14", "Password hashing and session management"),
        ("src/types/index.ts", "D.15", "Domain interfaces of the application"),
    ]
    for rel, listing, title in files:
        path = root / rel
        if not path.exists():
            continue
        code = path.read_text(encoding="utf-8", errors="replace")
        if len(code) > 9000:
            code = code[:9000] + "\n// ... listing truncated for the printed report ..."
        P(doc, "%s : %s   (%s)" % (listing, title, rel), justify=False, bold=True,
          space_after=2)
        CODE(doc, code)
        doc.add_page_break()
    P(doc, "End of the source code listing.")


def build_report():
    doc = new_document()
    enable_auto_update_fields(doc)

    # ---------------- front matter ----------------
    first = doc.sections[0]
    add_footer_page_numbers(first, fmt="lowerRoman", start=1, show=True)
    first.different_first_page_header_footer = True

    build_title_page(doc)
    build_certificate(doc)
    build_declaration(doc)
    build_acknowledgement(doc)
    build_abstract(doc)
    build_toc_pages(doc)

    # ---------------- main body ----------------
    body = doc.add_section(WD_SECTION.NEW_PAGE)
    body.page_width = Cm(21.0)
    body.page_height = Cm(29.7)
    body.top_margin = Cm(2.54)
    body.bottom_margin = Cm(2.54)
    body.left_margin = Cm(3.5)
    body.right_margin = Cm(2.0)
    add_footer_page_numbers(body, fmt="decimal", start=1, show=True)
    add_header_text(body, "%s  \u2014  Final Year Project Report" % CONFIG["project_title"])

    chapter_1(doc)
    chapter_2(doc)
    chapter_3(doc)
    chapter_4(doc)
    chapter_5(doc)
    chapter_6(doc)
    chapter_7(doc)
    chapter_8(doc)
    chapter_9(doc)

    build_references(doc)
    appendix_a(doc)
    appendix_b(doc)
    appendix_c(doc)
    if CONFIG["include_appendix_source_code"]:
        appendix_d(doc)

    out_path = BASE_DIR / CONFIG["output_docx"]
    doc.save(str(out_path))
    return out_path


def main():
    print("\n[1/3] drawing engineering diagrams ...")
    draw_architecture()
    draw_workflow()
    draw_use_case()
    draw_dfd_level0()
    draw_dfd_level1()
    draw_dfd_level2()
    draw_nlp_pipeline()
    draw_class_diagram()
    draw_ai_factory()
    draw_dedup()
    draw_sequence()
    draw_er_diagram()
    draw_chart_concept_scores()
    draw_chart_ngrams()
    draw_chart_pipeline_stats()
    draw_chart_jaccard()
    draw_chart_longitudinal()
    for number, caption, file_name in FIG_INDEX:
        print("      figure %-5s %s" % (number, caption))
    print("      diagrams written to : %s" % DIAGRAMS)

    print("\n[2/3] preparing screenshot assets ...")
    created, real = build_screenshot_assets()
    print("      placeholders generated : %d" % len(created))
    print("      real screenshots found : %d" % len(real))
    print("      screenshot guide       : %s" % (ASSETS / "SCREENSHOT_GUIDE.md"))

    print("\n[3/3] assembling the Word document ...")
    out = build_report()
    size_mb = out.stat().st_size / (1024.0 * 1024.0)
    print("      report written to : %s  (%.2f MB)" % (out, size_mb))
    print("\nDone. Open the file in MS Word and allow Word to update the fields so that the "
          "Table of Contents, the List of Figures, the List of Tables and the page numbers are "
          "populated.")


if __name__ == "__main__":
    main()
