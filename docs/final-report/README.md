# ExamEase — Final Project Report (Word) + Report Generator

This folder contains the complete, ready-to-submit project report for ExamEase together with the
Python script that generates it, so any change in the code base or in the student details can be
reflected in the document by re-running one command.

```
docs/final-report/
├── ExamEase_Report_Generator.py        # the generator (diagrams + DOCX assembly)
├── ExamEase_Final_Project_Report.docx  # the final report  (~22,000 words, 33 figures, 31 tables)
├── README.md                           # this file
└── assets/
    ├── SCREENSHOT_GUIDE.md             # which screenshot belongs to which figure
    ├── diagrams/*.png                  # 17 engineering diagrams + result charts
    └── screenshots/*.png               # 16 screenshot placeholders (Figures 8.1 – 8.16)
```

## 1. Generate (or regenerate) the report

```bash
pip install python-docx matplotlib     # one time

cd docs/final-report
python ExamEase_Report_Generator.py
```

The script is self-contained and prints a manifest of every figure it draws before assembling the
document. Running it again overwrites the DOCX and the diagrams; **real screenshots that already
exist in `assets/screenshots/` are never overwritten** — a placeholder is only created when the
file is missing.

## 2. Add your screenshots (16 figures)

Two ways, both keep the frame, size and alignment:

1. **Inside Word (easiest):** right click a placeholder → *Change Picture* → *This Device* →
   select your PNG.
2. **Before generating:** save your screenshots as
   `assets/screenshots/fig_8_1.png` … `fig_8_16.png` and re-run the script.

`assets/SCREENSHOT_GUIDE.md` lists exactly what to capture for every figure and from which route of
the running application (`/register`, `/dashboard`, `/generate`, `/quiz/{id}`, `/results/{id}`, …).

## 3. After opening the DOCX in MS Word

* Word asks to update fields → click **Yes** (builds the Table of Contents, List of Figures, List
  of Tables and page numbers). If nothing happens, press `Ctrl + A` then `F9` → *Update entire
  table*.
* The pagination is left-justified 1.5 line spacing, Times New Roman 12 pt, A4 with 3.5 cm binding
  margin — typical university (“black book”) formatting.

## 4. Personalise the report

Open `ExamEase_Report_Generator.py` and edit the `CONFIG` dictionary at the top:

```python
"student_name": "Shubham Dalvi",
"roll_number":  "______________",
"guide_name":   "Prof. ____________________",
"college_name": "________________ College of Engineering",
"university_name": "University of Mumbai",
"academic_year": "2026 - 2027",
"month_year": "October 2026",
```

Everything else — certificate, declaration, title page, header and footer — is filled from this
dictionary, so no manual editing of the document is required.

## 5. Contents of the report

| Part | Contents |
| :--- | :--- |
| Front matter | Title page, certificate, declaration, acknowledgement, abstract with keywords, Table of Contents, List of Figures, List of Tables, list of abbreviations, organisation of the report |
| Chapter 1 | Introduction — background, motivation, problem statement, objectives, scope, novelty, report organisation |
| Chapter 2 | Literature survey — classical NLP techniques, question generation approaches, document extraction, grounding techniques, comparison table (Table 2.1), research gap |
| Chapter 3 | System analysis and design — existing system, functional requirements, layered architecture, workflow, use case diagram + descriptions, DFD levels 0/1/2, module design, hardware/software requirements, constraints |
| Chapter 4 | Methodology and implementation — the seven pipeline stages with algorithms and code listings, ingestion service, service layer class diagram, grounded prompt engineering, provider factory, validation and Jaccard de-duplication, complexity analysis, sequence diagram |
| Chapter 5 | Database design — ER diagram, all four collection schemas, embedded document structures, index plan, integrity and deletion policy |
| Chapter 6 | Results and discussion — measured pipeline output, top ranked concepts, n-gram analysis, Jaccard evaluation, sample generated questions, API measurements, document size experiment, comparison with existing tools |
| Chapter 7 | Security, privacy and performance — bcrypt, JWT HTTP-only cookies, authorisation, input validation, in-memory document handling, pooling |
| Chapter 8 | Implementation and testing — 16 application screens, test strategy, 24 unit cases, 15 integration cases and 10 boundary cases with results |
| Chapter 9 | Conclusion and future scope — objective-to-outcome mapping, seven extensions, closing remark |
| Back matter | References (21 entries), Appendix A (REST API table), Appendix B (environment + commands), Appendix C (16 week project plan), Appendix D (source code of 14 core modules) |

All measurements printed in Chapter 6 come from actually executing the project's pipeline
(`src/services/nlp/`) on a 52-sentence lecture note, so the figures, tables and the narrative agree
with each other.
