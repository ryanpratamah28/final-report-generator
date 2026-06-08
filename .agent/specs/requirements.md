# Product Requirements Document (PRD): IPB Thesis Generator

## 1. Vision
Empower IPB students to produce high-quality, perfectly formatted scientific papers by automating the formatting and linguistic overhead of the PPKI IPB Edisi ke-4.

## 2. Core Features
- **Auto-Formatting**: Apply 4-3-3-3 margins, TNR font, and 1.0 spacing to `.docx` files.
- **Citation Engine**: Automatically format in-text citations and bibliographies in CSE Harvard style.
- **Scientific Editor**: Automatically italicize taxonomy and correct Indonesian grammar (PUEBI).
- **Structure Builder**: Generate the standard structure from Title Page to Bibliography.

## 3. User Stories
- *As a student*, I want to paste my raw results and have the agent format them into a proper "Bab IV: Hasil dan Pembahasan" section.
- *As a student*, I want the agent to check if my citations match the bibliography list.
- *As a student*, I want to ask "is this margin correct?" and have the agent fix it automatically.

## 4. Acceptance Criteria
- Output must be a valid `.docx` file.
- All formatting must pass a visual check against the `PPKI IPB.pdf` guidelines.
- Bibliography must be alphabetically sorted.
