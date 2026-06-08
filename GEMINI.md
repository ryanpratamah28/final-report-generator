# Project Overview: Karya Ilmiah (Bachelor's Thesis) - IPB University

This workspace is dedicated to the writing, formatting, and management of a Bachelor's Thesis (Skripsi S-1) at **IPB University**. The project strictly adheres to the **Pedoman Penulisan Karya Ilmiah (PPKI) IPB Edisi ke-4**. It includes specialized guidelines, reference materials, and an AI agent skill designed to assist with scientific writing in Indonesian.

## Directory Structure

- **Root Directory**: Contains the primary manuscript (`PPKI IPB.docx`) and its PDF version. These are the central files for the thesis.
- **.agent/skills/ppki-ipb-skripsi/**: Contains the core instructional logic for the AI assistant.
    - `SKILL.md`: The primary directive for how the AI should assist with the thesis (formatting, language, structure).
    - `references/`: Detailed reference guides for:
        - `tata-nama-ilmiah.md`: Scientific nomenclature for organisms.
        - `angka-satuan-lambang.md`: Rules for numbers, SI units, and symbols.
        - `contoh-daftar-pusaka.md`: Comprehensive bibliography examples (CSE Harvard style).

## Key Standards & Conventions

### 1. Formatting Requirements (PPKI IPB)

- **Paper**: A4, HVS 80g.
- **Font**: Times New Roman, 12 pt (Main text), 14 pt (Chapter Titles, Bold, Caps).
- **Margins**: Left (4 cm), Right (3 cm), Top (3 cm), Bottom (3 cm).
- **Spacing**: Single spacing (1.0).
- **Paragraph Indent**: 1 cm.
- **Alignment**: Justified.

### 2. Language & Style

- **Language**: Formal Indonesian (Bahasa Indonesia Ilmiah) following PUEBI/KBBI.
- **Tense**: Passive voice is generally preferred for results and methodology.
- **Scientific Names**: Must be in _italics_ (genus level and below).

### 3. Citations & Bibliography

- **Style**: CSE Harvard.
- **Citation Example**: `(Naim 2020)` or `Naim (2020)`.
- **Minimum References**: 20 sources (primarily journals from the last 10 years).
- **Bibliography Format**: `Author AB. Year. Title. Journal. Volume(Issue):Pages.`

## Usage Guidelines

- **Drafting & Editing**: Use the `ppki-ipb-skripsi` skill for drafting chapters, checking formatting, and correcting Indonesian grammar.
- **Validation**: Always verify that any generated text complies with the strict margin and spacing rules detailed in `SKILL.md`.
- **References**: Consult the markdown files in `.agent/skills/ppki-ipb-skripsi/references/` for specific technical rules regarding nomenclature and units.
- **Word Document**: The primary work should be maintained in `PPKI IPB.docx`. Use Word-specific tools or manual verification to ensure final formatting matches the 4-3-3-3 margin requirement.

## TODO / Future Tasks

- [ ] Verify bibliography matches CSE Harvard style for all sources.
- [ ] Perform a final check of scientific names for italics.
- [ ] Ensure Abstract is under 200 words and fits on one page.
