# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## What Bar-Skills Is

Bar-Skills is a beverage consulting and education practice run by Oleks Iurchenko, based in Kingston, Ontario (bar-skills.com | oleks@bar-skills.com | 416 994 3223). Services span three lines: a private 1:1 Craft Bartender Course, a self-paced Home Bar System online course, and private event hosting (cocktail masterclasses, wine service). B2B consulting to Kingston-area restaurants is a fourth revenue stream in active outreach.

## The Craft Bartender Course

A 9-session private 1:1 course. Each session is structured as 30 min Theory + 60 min Practical. Theory is always 6 topics (T1–T6) with a Knowledge Test at the end. Practical builds run in sequence with station orientation at the start.

**Sessions and topics:**

| Session | Title | Key builds / focus |
|---------|-------|--------------------|
| S01 | Foundations | Mise en place, professional bar setup |
| S02 | Spirits Knowledge | Single malt Scotch, tequila categories (blanco/reposado) |
| S03 | Stirred Builds | Negroni, Martini, Whiskey Sour; vermouth handling |
| S04 | Balance and Technique | Live diagnosis (Hot/Balanced/Thin), flavour modification, float; builds: Daiquiri, Peach Daiquiri, French 75, Penicillin |
| S05 | Fermented Foundations & The Spritz Family | Wine structure, beer basics, sparkling service; builds: Kir Royale, Aperol Spritz, wine tasting, beer comparison |
| S06 | Original Builds | Smoked Old Fashioned, demerara syrup, recipe spec creation |
| S07 | Shrubs & Zero-Proof | Shrubs (1:1:1 formula), zero-proof builds, Mojito muddling, guest complaint handling (Acknowledge / Apologize / Act) |
| S08 | Speed, Flow, and the Capstone | Simultaneous builds (Negroni + Daiquiri sequencing), batching/pre-building, capstone from memory |
| S09 | Menu Engineering | Pour cost, the four-cell matrix (Star/Plowhorse/Puzzle/Dog), menu design principles, inventory; builds: Black Russian, White Russian, Classic Cosmopolitan, Elevated Cosmopolitan |

Each session produces two documents: a **Student Book (SB)** and a **Teacher's Manual (TM)**. The TM adds timing markers, facilitator scripts ("WHAT TO SAY / WHAT TO SHOW"), WATCH OUT notes, discussion prompts, and answer keys. The SB has REFLECT questions, PRO TIP, and NOTE call-out boxes.

A **Comprehensive Exam** (`BarSkills_AnswerKey_v1.docx`) exists as an instructor-only marking guide covering all 9 sessions (20 MC questions + short answer).

## The Home Bar System (Online Course)

A self-paced course hosted at **learn.bar-skills.com** (Tutor LMS). It serves as the entry product and upsell path into the Craft Bartender Course.

- **Price:** $89 CAD standard | $79 CAD founding launch | $99 CAD post-traction
- **Runtime:** ~2–3 hours of video + 5 downloadable PDFs
- **Audience:** Beginners and home hosts, primarily Ontario

**Four lessons and their deliverables:**

| Lesson | Topic | PDF deliverable |
|--------|-------|-----------------|
| 1 | The Plan — stocking system, space, LCBO shopping list | Home Bar Blueprint |
| 2 | The Kit — tools, glassware, ice | Equipment Checklist |
| 3 | The Four Templates — build dozens of drinks from four ratios | Recipe and Ratio Cards |
| 4 | The Host — mise en place, batching, guest math, simple menu design | Hosting Playbook |

The fifth PDF (LCBO Bottle Buying Guide) doubles as the lead magnet, SEO article, email opt-in, and Instagram carousel. All LCBO product selections must be non-US (US alcohol unavailable in Ontario since early 2025).

**The Four Templates (course spine):**

| Template | Formula | Example unlocks |
|----------|---------|----------------|
| Old Fashioned | Spirit + sugar + bitters | Old Fashioned, sweet-and-bitter family |
| The Sour | Spirit + citrus + sweet | Daiquiri, Margarita, Whiskey Sour, Gimlet |
| The Highball | Spirit + bubbles | Gin and Tonic, Tom Collins, Mojito-style |
| Stirred / Spirit-Forward | Spirit + vermouth (+ bitters) | Manhattan, Martini, Negroni-style |

## This Repository

The repo currently contains `generate_s09.py`, a Python script that programmatically generates the Session 09 Word documents.

**Running it:**
```bash
pip install python-docx
python generate_s09.py
```

Outputs `BarSkills_S09_SB.docx` and `BarSkills_S09_TM.docx` to the working directory.

**Script architecture** (`generate_s09.py`, ~1582 lines):

- **Helpers (L11–L173):** `python-docx` wrappers — `set_cell_bg`, `add_heading`, `add_body`, `add_labelled_block`, `add_box_block`, `add_two_col_table`, `add_divider`, `add_quote`, `add_reflect`, `add_transition`, `add_theory_header_sb`, `add_theory_header_tm`, `add_build_header`
- **`build_sb()` (L175–L903):** Student Book — Cover → Welcome → T1–T6 → four cocktail builds → Menu Analysis Exercise → Knowledge Test
- **`build_tm()` (L904–end):** Teacher's Manual — same flow with Session Control table, timing markers, facilitator notes, WATCH OUT blocks, answer keys

**Brand colors:** `1F3864` (dark blue headings), `7F7F7F` (grey labels/footers), `EBF0F7` / `F2F2F2` (table row alternates).

Sessions S01–S08 were built manually as Word documents stored in Google Drive. The Python generation approach (used for S09) is the pattern for future sessions.

## Document Naming Convention

| Format | Example |
|--------|---------|
| Older sessions (Drive) | `Bar-Skills S 04 SB.docx`, `Bar-Skills S 04 TM.docx` |
| Newer sessions (repo) | `BarSkills_S09_SB.docx`, `BarSkills_S09_TM.docx` |
| Proposals | `BarSkills_Proposal_ClientName_vN.pdf` |
| Course specs | `BarSkills_HomeBarSystem_CourseSpec_v1.docx` |

## Business Context

- **Geography:** Kingston, Ontario; also serves Prince Edward County and Eastern Ontario
- **Credentials:** WSET Level 3, 20+ years hospitality, National Bartending Competition Finalist (Top 3)
- **B2B outreach:** tracked in `BarSkills_OutreachCampaign_v1` (Google Sheets), targeting Kingston and Loyalist-area independent restaurants
- **Private events:** cocktail masterclasses + sommelier wine service; Option A (cocktails only, $950 CAD) / Option B (full afternoon with wine service); client provides glassware and alcohol, Bar-Skills brings everything else
- **Local distillery partnerships:** being cultivated with KL Craft Distillery, Bare Bones, Wild Lot, Kinsip, Stillus/Elements Distillery for Home Bar System course features
