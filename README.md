# 🎓 ExamEase — Academic NLP & AI-Powered Question Generator

[![Next.js](https://img.shields.io/badge/Next.js-16.3-black?style=flat&logo=next.js)](https://nextjs.org/)
[![React](https://img.shields.io/badge/React-19.2-blue?style=flat&logo=react)](https://react.dev/)
[![Tailwind CSS](https://img.shields.io/badge/TailwindCSS-v4-38bdf8?style=flat&logo=tailwindcss)](https://tailwindcss.com/)
[![MongoDB](https://img.shields.io/badge/MongoDB-7.6-green?style=flat&logo=mongodb)](https://www.mongodb.com/)
[![Gemini](https://img.shields.io/badge/Google_Gemini-2.5_Flash-8e75ff?style=flat&logo=google)](https://aistudio.google.com/)

> **ExamEase** transforms academic lecture notes, textbook chapters, and PDF documents into syllabus-grounded, verified Multiple Choice Questions (MCQs) using a hybrid **deterministic NLP pipeline + Google Gemini 2.5 Flash** architecture.

📖 **For the exhaustive, complete technical documentation from A to Z, please read [PROJECT_DOCUMENTATION.md](file:///c:/shubham%20dalvi%20project/ExamEase/PROJECT_DOCUMENTATION.md).**

---

## 🌟 Key Highlights

- **Classical NLP Pre-processing**: 7-stage deterministic NLP engine (Tokenization, POS Tagging, Sentence Segmentation, Lemmatization, N-Gram frequency analysis, and Concept/Definition scoring).
- **Universal PDF Extraction via Gemini Vision OCR**: Seamlessly handles text, scanned pages, handwritten notes, and documents with non-standard font encodings (`Identity-H` / CID).
- **Grounded AI Generation**: Generates MCQs strictly grounded on extracted source concepts (zero external hallucinations).
- **Automated Quality & Deduplication Filter**: Jaccard lexical similarity analysis eliminates duplicate questions; guarantees 4 balanced options and verified citations.
- **Interactive Teacher Dashboard**: Edit questions in real-time, modify distractors, or regenerate single questions in-place.
- **Student Examination Engine**: Built-in test taking interface with timers, question navigator, instant scoring, and concept citations.
- **Enterprise-Grade Security**: Salted `bcryptjs` password hashing and stateless JWT authentication stored in HTTP-only cookies.

---

## 🚀 Quick Start Guide

### 1. Prerequisites
- **Node.js**: `v20.x` or `v22.x`
- **MongoDB**: Running locally on port `27017` or a MongoDB Atlas URI

### 2. Installation
```bash
# Clone the repository and navigate inside
cd ExamEase

# Install dependencies
npm install
```

### 3. Environment Configuration
Create a `.env.local` file by copying the template:
```bash
copy .env.example .env.local
```

Inside `.env.local`:
```env
DATABASE_URL=mongodb://127.0.0.1:27017/examease
JWT_SECRET=your_super_secret_jwt_key
GEMINI_API_KEY=your_gemini_api_key_here
NEXT_PUBLIC_APP_URL=http://localhost:3000
```
*(Note: If `GEMINI_API_KEY` is omitted, ExamEase automatically falls back to an offline mock generator).*

### 4. Running the Application

```bash
# Start development server with Turbopack
npm run dev

# Or build for production
npm run build
npm run start
```

Open **[http://localhost:3000](http://localhost:3000)** in your browser.

---

## 📁 Project Structure

```text
ExamEase/
├── PROJECT_DOCUMENTATION.md   # Complete A-to-Z Technical Reference
├── .env.example               # Environment variables template
├── src/
│   ├── app/                   # Next.js App Router (Pages, Layouts & API routes)
│   ├── components/            # UI components (MCQ preview cards, NLP concept inspector, navbar)
│   ├── lib/                   # MongoDB client (db.ts) and JWT authentication (auth.ts)
│   ├── services/
│   │   ├── ai/                # Gemini 2.5 Flash service & Mock AI service
│   │   ├── nlp/               # Tokenizer, POS tagger, lemmatizer, concept scorer & deduplicator
│   │   ├── pdf.service.ts     # Gemini Vision OCR PDF extractor
│   │   └── mcq.service.ts     # Generation, validation & deduplication coordinator
│   └── types/                 # TypeScript interfaces and domain schemas
```


---

## 📄 Final Project Report (Word document + Python generator)

The complete, university-format project report is generated programmatically from a single
Python script, so every diagram, table, figure caption and page number stays consistent with
the code base.

```bash
# 1. installs (only needed once)
pip install python-docx matplotlib

# 2. move into the report folder
cd docs/final-report

# 3. generate the report
python ExamEase_Report_Generator.py
```

Generated artefacts:

| File | Description |
| :--- | :--- |
| `docs/final-report/ExamEase_Final_Project_Report.docx` | Final report — title page, certificate, declaration, acknowledgement, abstract, Table of Contents, List of Figures, List of Tables, abbreviations, 9 chapters, references and 4 appendices |
| `docs/final-report/assets/diagrams/*.png` | 17 engineering diagrams: layered architecture, activity workflow, use case, DFD levels 0/1/2, NLP pipeline, class diagram, ER diagram, sequence diagram, AI factory fallback, de-duplication flow and result charts |
| `docs/final-report/assets/screenshots/*.png` | 16 pre-formatted screenshot placeholders (Figures 8.1 – 8.16) |
| `docs/final-report/assets/SCREENSHOT_GUIDE.md` | Which screenshot belongs to which figure and how to capture it |

**Adding your own screenshots:** right click a placeholder inside MS Word → *Change Picture* →
*This Device*, or drop a real PNG with the same file name into
`docs/final-report/assets/screenshots/` and re-run the script.

**After opening the DOCX:** Word asks to update fields — click **Yes** (or press `Ctrl + A`
then `F9`) so that the Table of Contents, the List of Figures, the List of Tables and the page
numbers are populated.

Edit the `CONFIG` dictionary at the top of `ExamEase_Report_Generator.py` to change the student
name, roll number, guide name, college, university or academic year, then regenerate.

---

## 📄 License & Attribution

Developed for academic NLP research and automated examination generation.
