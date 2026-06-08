# Karya Ilmiah: Bachelor's Thesis

This workspace is dedicated to the writing, formatting, and management of the Final Report for **Ryan Al-Thariq** (NIM: **G64101234**) in the Department of Computer Science (_Ilmu Komputer_), Faculty of Mathematics and Natural Sciences (_FMIPA_).

The project is structured to strictly adhere to the official **Pedoman Penulisan Karya Ilmiah (PPKI) IPB Edisi ke-4**.

---

## 📝 Thesis Information

- **Title**: _Implementasi Sistem Notifikasi Terintegrasi Berbasis API Gateway untuk Efisiensi Koordinasi Tim Multi-Platform: Studi Kasus Tasku.site_
- **Focus**: Software engineering, API Gateways, event-driven webhooks, and asynchronous message broker queues (Redis) integrated into a multi-platform notification hub ([Tasku.site](https://tasku.site)).

---

## 📁 Directory Structure

Below is the repository structure containing the manuscript files, guidelines, scripts, and the AI agent's brain configuration:

```text
Karya Ilmiah/
├── .agent/                               # AI Thesis Assistant Configuration & Knowledge
│   ├── facts/
│   │   ├── glossary.md                   # Key terms (PPKI, CSE Harvard, etc.)
│   │   └── project_map.md                # Overview of workspace file tree
│   ├── memory/
│   │   └── 2026-06-05.md                 # Collaborative session log
│   ├── rules/
│   │   ├── constraints.md                # Strict restrictions (margins, font, sources)
│   │   └── style.md                      # Language tone, capitalization, italics rules
│   ├── skills/
│   │   └── ppki-ipb-skripsi/             # Core IPB Thesis writing skill
│   │       ├── SKILL.md                  # Main directives for thesis structure & grammar
│   │       ├── prompts/
│   │       │   └── example-system-prompt.md # Example system instructions
│   │       ├── references/               # Official reference markdown files
│   │       │   ├── angka-satuan-lambang.md  # Rules for numbers, symbols, and SI units
│   │       │   ├── contoh-daftar-pusaka.md  # Detailed CSE Harvard bibliography examples
│   │       │   └── tata-nama-ilmiah.md      # Rules for biological nomenclature
│   │       └── scripts/                  # Automated DOCX generation scripts
│   │           ├── ipb_logo.png          # Logo inserted into cover pages
│   │           ├── requirements.txt      # Python dependencies (python-docx)
│   │           ├── generate_template.py  # Script to generate empty thesis layout template
│   │           └── generate_tasku_paper.py # Script to generate the full drafted thesis draft
│   └── tasks/
│       ├── backlog.md                    # Future tasks and prioritizing queue
│       └── in_progress.md                # Log of active and next steps
├── AGENT.md                              # AI Agent Brain Map (entrypoint index)
├── GEMINI.md                             # Workspace conventions and rules
├── Karya Ilmiah_Tasku_PPKI IPB_V01.docx  # Thesis Draft (Version 1)
├── Karya Ilmiah_Tasku_PPKI IPB_V02.docx  # Thesis Draft (Version 2)
├── Karya Ilmiah_Tasku_PPKI IPB_V03.docx  # Current Thesis Manuscript (Version 3)
├── PPKI IPB.docx                         # Original PPKI IPB guide document
├── PPKI IPB.pdf                          # Original PPKI IPB guide PDF (Source of Truth)
├── LAMPIRAN PPKI IPB.pdf                 # Appendices of formatting examples PDF
├── PPKI_IPB_Template.docx                # Empty structured thesis template
└── README.md                             # Project Documentation (This file)
```

### Key File Links:

- **Brain Index**: [AGENT.md](file:///c:/a_ryans/Bachelor%20Degree/Karya%20Ilmiah/AGENT.md)
- **Workspace Config**: [GEMINI.md](file:///c:/a_ryans/Bachelor%20Degree/Karya%20Ilmiah/GEMINI.md)
- **Active Tasks**: [in_progress.md](file:///c:/a_ryans/Bachelor%20Degree/Karya%20Ilmiah/.agent/tasks/in_progress.md)
- **Future Work**: [backlog.md](file:///c:/a_ryans/Bachelor%20Degree/Karya%20Ilmiah/.agent/tasks/backlog.md)
- **Past Sessions**: [2026-06-05.md](file:///c:/a_ryans/Bachelor%20Degree/Karya%20Ilmiah/.agent/memory/2026-06-05.md)

---

## 🛠️ Automated Scripts

The repository contains Python scripts that automate the generation of documents formatted according to IPB standards.

### Dependencies

Before running the scripts, make sure you have installed the required libraries listed in [requirements.txt](file:///c:/a_ryans/Bachelor%20Degree/Karya%20Ilmiah/.agent/skills/ppki-ipb-skripsi/scripts/requirements.txt):

```bash
pip install -r .agent/skills/ppki-ipb-skripsi/scripts/requirements.txt
```

### 1. Template Generator

- **File**: [generate_template.py](file:///c:/a_ryans/Bachelor%20Degree/Karya%20Ilmiah/.agent/skills/ppki-ipb-skripsi/scripts/generate_template.py)
- **Purpose**: Generates [PPKI_IPB_Template.docx](file:///c:/a_ryans/Bachelor%20Degree/Karya%20Ilmiah/PPKI_IPB_Template.docx), which contains the correct mirror margins, page number styles (Roman vs Arabic), covers, title page, and empty section placeholders.
- **Usage**:
    ```bash
    python .agent/skills/ppki-ipb-skripsi/scripts/generate_template.py
    ```

### 2. Thesis Manuscript Generator

- **File**: [generate_tasku_paper.py](file:///c:/a_ryans/Bachelor%20Degree/Karya%20Ilmiah/.agent/skills/ppki-ipb-skripsi/scripts/generate_tasku_paper.py)
- **Purpose**: Generates the complete initial draft of Ryan Al-Thariq's thesis containing the actual written sections (from Bab I: Pendahuluan to Bab V: Simpulan dan Saran) as well as the CSE Harvard bibliography.
- **Usage**:
    ```bash
    python .agent/skills/ppki-ipb-skripsi/scripts/generate_tasku_paper.py
    ```

---

## 🤖 AI Assistant Capabilities (`ppki-ipb-skripsi`)

The workspace includes a dedicated AI agent skill configured in [SKILL.md](file:///c:/a_ryans/Bachelor%20Degree/Karya%20Ilmiah/.agent/skills/ppki-ipb-skripsi/SKILL.md) to help write, review, and format the thesis. The assistant follows strict formatting and linguistic rules:

### 1. Formatting Specifications (PPKI IPB)

- **Paper Size**: A4
- **Margins**: Left/Inside: 4 cm, Top/Bottom/Right/Outside: 3 cm (Mirror Margins enabled)
- **Font**: Times New Roman, 12 pt for main text (1.0 single spacing, 1 cm first-line indent, justified), 14 pt bold for Chapter Titles
- **Abstract**: Maximum 200 words, single paragraph, single page for both Indonesian (Abstrak) and English (Abstract)

### 2. Indonesian Scientific Language & Style (PUEBI/KBBI)

- Automatic grammar checks to correct common informalities and spelling mistakes.
- Verification of scientific taxonomic names (genus and below) to be correctly italicized (e.g., _Oryza sativa_).
- Refer to [tata-nama-ilmiah.md](file:///c:/a_ryans/Bachelor%20Degree/Karya%20Ilmiah/.agent/skills/ppki-ipb-skripsi/references/tata-nama-ilmiah.md) and [angka-satuan-lambang.md](file:///c:/a_ryans/Bachelor%20Degree/Karya%20Ilmiah/.agent/skills/ppki-ipb-skripsi/references/angka-satuan-lambang.md) for detailed guidelines.

### 3. Citations & Bibliography (CSE Harvard Style)

- Strict in-text citations: `(Author Year)` (e.g., `(Naim 2020)`) or `(Author1 and Author2 Year)` (e.g., `(Naim and Keraf 2020)`)
- For 3 or more authors: `(Author et al. Year)` (e.g., `(Naim et al. 2020)`)
- Alphabetically sorted bibliography in hanging indent format.
- Refer to [contoh-daftar-pusaka.md](file:///c:/a_ryans/Bachelor%20Degree/Karya%20Ilmiah/.agent/skills/ppki-ipb-skripsi/references/contoh-daftar-pusaka.md) for full reference examples.

---

## 🚀 Ongoing Roadmap & Backlog

The following tasks are managed under the `.agent/tasks/` directory:

1. **Bibliography Verification**: Ensure all cited references are correctly formatted in CSE Harvard style and cross-referenced.
2. **Abstract Verification**: Verify that the English and Indonesian abstracts fit on a single page and are under 200 words.
3. **Scientific Nomenclature Scanner**: Scan the document to verify that all species and genus names are correctly formatted.
4. **Halaman Pengesahan**: Generate the approval sheet structure with precise spacing.

## Important Notes

> **This project is intended solely as an academic formatting and educational aid for students.**

### Disclaimer

- All names, organizations, institutions, datasets, scenarios, and examples used in this project are **dummy data** and are provided for demonstration, educational, and formatting purposes only.
- Any resemblance to actual persons, organizations, products, services, or events is purely coincidental.
- The generated content should be treated as a **sample academic document** and not as a final submission-ready thesis, dissertation, journal article, or scientific publication.
- Students are responsible for validating, revising, and adapting all generated content according to their institution's academic guidelines, supervisor feedback, and research requirements.
- References, citations, figures, tables, and appendices should be independently verified before use in any official academic submission.
- This project does not replace academic supervision, peer review, or institutional review processes.

### Usage Guidelines

- Use this project as a reference for:
    - Academic document structure
    - Formatting standards
    - Writing organization
    - Chapter composition
    - Citation placement
    - Thesis or final project preparation

- Do not use generated dummy data, assumptions, analyses, or conclusions as real research findings without proper validation and supporting evidence.

### Academic Integrity

Users are expected to comply with their institution's policies regarding:

- Academic honesty
- Citation practices
- Research ethics
- Plagiarism prevention
- Responsible use of AI-assisted writing tools
