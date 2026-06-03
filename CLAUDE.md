# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## What This Repo Does

This is a single-script Python project that generates `.docx` course materials for **Bar-Skills**, a craft bartender training program by Oleks Iurchenko (bar-skills.com). The script currently produces Session 09 ("Menu Engineering") as two Word documents.

## Running the Script

```bash
pip install python-docx
python generate_s09.py
```

Outputs:
- `BarSkills_S09_SB.docx` — Student Book
- `BarSkills_S09_TM.docx` — Teacher's Manual

Both files are saved to the working directory. The script ends with `build_sb()`, `build_tm()`, `print('Done.')`.

## Architecture

**`generate_s09.py`** (~1582 lines) is structured in three layers:

1. **Formatting helpers** (L11–L173): Low-level `python-docx` wrappers — `set_cell_bg`, `add_heading`, `add_body`, `add_labelled_block`, `add_box_block`, `add_two_col_table`, `add_divider`, `add_quote`, `add_reflect`, `add_transition`, and document-type-specific headers (`add_theory_header_sb`, `add_theory_header_tm`, `add_build_header`).

2. **`build_sb()`** (L175–L903): Builds the Student Book. Sections follow this order: Cover → Welcome → Theory blocks T1–T6 → Four cocktail builds (Black Russian, White Russian, Classic Cosmopolitan, Elevated Cosmopolitan) → Menu Analysis Exercise → Knowledge Test.

3. **`build_tm()`** (L904–end): Builds the Teacher's Manual with the same session flow but adds timing markers (minutes per section), facilitator notes, answer keys, and a Session Control block at the top.

**Brand colors** used throughout: `1F3864` (dark blue headings), `7F7F7F` (grey labels/footers), `EBF0F7` / `F2F2F2` (table alternating rows).

## Extending to New Sessions

To add a new session, copy `generate_s09.py`, rename it (e.g., `generate_s10.py`), update the cover block strings, replace the theory/build content, and adjust the `doc.save(...)` paths at the end of each build function. All formatting helpers are self-contained and can be reused as-is.
