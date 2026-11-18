# Changelog

## v1.3.0 - 2026-11-18

- A label and its number ("Stage 4", "Lesson 3", "Step 2", "Part 1") now stay together across a line break, in every language, so a number never drops to a line of its own.
- New warning: the narration says a quantity ("ten questions", "twenty minutes") that no text on the slide carries. Anything the learner must do or reach has to be on screen.
- New warning: the narration says an item is required or mandatory but the slide does not.
- New warnings for patterns that confuse viewers: a total or overall figure inside a `process` slide (use `cards`), a title that promises a count that differs from the list under it, a "this lesson" question in a multi-video lesson with no scope, a slide missing from the middle of an "X N of M" label chain.
- The example closing slide now writes "within seven days" on screen, as its narration says.

## v1.2.0 - 2026-11-17

- New optional deck key `footer_text`: replaces the footer text on every slide except the title slide, which keeps `instructor · contact`. Without it nothing changes.
- The course name on the title slide is now 28 pt (was the 24 pt floor) and steps down to 24 pt only if it does not fit on one line. Decks without a long course name differ from v1.1.0 only in this line on the cover.

## v1.1.0 - 2026-11-17

Fixes from a first run by a new user on Windows 11 (Python 3.12 in a venv, LibreOffice and Poppler from winget).

- README: PowerShell commands for Windows next to the bash ones, for both the personal and the per-project install. Every install now makes a `.venv` in the project and installs `requirements.txt` into it.
- README: how to install LibreOffice and Poppler (`pdftoppm` on PATH) on Windows (winget), macOS (Homebrew) and Linux (apt, dnf), how to check it, and what to do when `pdftoppm` is not found. A note on Arial and Liberation Sans for the text measurement.
- SKILL.md: Claude runs the scripts with the project's `.venv` Python when there is one (`.venv/Scripts/python` on Windows, `.venv/bin/python` elsewhere), otherwise `python`.
- SKILL.md: Claude takes or asks for the instructor name and contact for the footer, and when none is given, says in the report that every slide still reads "Your name · yoursite.com" and how to change it.
- `build_slides.py`: the part-divider check now counts an `agenda` repeated with `current` set to part 2 or later, as SKILL.md always allowed. Before, a deck that used repeated agendas instead of `section` slides got a false warning ("The agenda lists 3 parts but there are only 0 'section' slides"). Decks without enough dividers still get the warning, now naming both kinds.
- `references/slide-types.md`: a table of the labels and prefixes the script adds by itself ("By the end of this lesson you can", "Next lesson: ", "Source: " and the rest, in English and Vietnamese), how to change or hide each, the extra English text `split_videos.py` writes, and what the footer prints by default.

## v1.0.0 - 2026-11-12

First public release.
