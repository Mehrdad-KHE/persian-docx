---
name: persian-docx
description: Write Word (.docx) files in Persian that are correct on both halves — the mechanics (right-to-left throughout, Vazirmatn font, Persian digits, English terms that do not scramble the line) and the prose (fluent, human Farsi that reads like someone thought in Persian, not like a translation). Use whenever the owner asks for a Word document, report, contract, letter or any deliverable in Persian, and whenever a .docx has to be produced or edited for a Farsi reader. Do NOT use for English-only documents.
---

# persian-docx

Two halves, both required. A file can be perfectly right-to-left and still be
unreadable because the sentences were built in English and translated word by
word. **Part 1 is the machinery. Part 2 is the writing, and it matters more.**

# Part 1 — the machinery

python-docx has **no** RTL support. A document built with it looks Persian and
reads broken: paragraphs left-aligned, English words jumping to the wrong side,
digits in Latin, and the font ignored for Arabic-script glyphs because Word
takes those from the *complex-script* font, not the Latin one.

`persian_docx.py` in this folder fixes all of that. Use it — do not hand-roll.

## Use it

```python
import sys, os
sys.path.insert(0, os.path.dirname(__file__))   # the folder this SKILL.md is in
from persian_docx import PersianDoc

d = PersianDoc("قراردادِ شراکت")
d.h2("بخش ۱ — طرفین")
d.p("این قرارداد بین مهرداد و [نامِ کامل] بسته می‌شود.")
d.bullet("مالکیتِ نرم‌افزار منتقل نمی‌شود.")
d.page_break()
d.save(r"D:\path\out.docx")
```

`PersianDoc(digits=True)` · `title` `h1` `h2` `h3` `p` `bullet` `page_break`
`save`. Pass `digits=False` to keep Latin numerals, `justify=False` on a short
standalone line.

Before handing a file over, run the auditor. Its output must be empty:

```python
from persian_docx import audit
print(audit(r"path.docx"))
```

## What it does, and why each part is needed

- **`w:bidi` on every paragraph and on the section.** Without it the paragraph is
  LTR and the last word of a Persian sentence lands on the left.
- **`w:rtl` and `w:cs` on every run**, plus `w:rFonts w:cs="Vazirmatn"`. Word
  picks the font for Arabic script from the complex-script slot; setting only
  `font.name` leaves the text in Word's default and the file looks nothing like
  Vazirmatn.
- **Right alignment** — separate from bidi, and both are required.
- **`w:jc` set to `start`, never `right`.** This one is counter-intuitive and it
  is what breaks most Persian documents: inside a bidi paragraph Word reads
  `right` as the *logical* right, which lands on the physical **left**. Headings
  written with `right` hug the wrong margin while justified body text looks
  fine — so the bug hides. Use `start` (reader's side), `both` (justified),
  `center`. The auditor rejects anything else.
- **Real heading styles** — `Title`, `Heading 1/2/3`. Bold text in a `Normal`
  paragraph is not a heading: Word builds no navigation pane, no table of
  contents, and no structure from it. Each built-in style is LTR and carries its
  own Latin font, so the skill rewrites the style before first use.
- **A separate display face for headings** (`B Titr`) against `Vazirmatn` for
  body. One typeface throughout reads flat in Persian.
- **Justified body text.** A ragged left edge is what makes a Persian page look
  unfinished.
- **`«` `»` instead of `"`.** Straight quotes are wrong in Persian; the skill
  converts them and the auditor flags any that survive.
- **A written `•` marker instead of Word's list.** Word's own bullet renders as
  `/` in an RTL Persian document. An arrow glyph is also wrong — it reads as
  Word's own heading-collapse triangle.
- **Persian digits**, except inside `[...]` placeholders and after `$`. A blank
  that reads `[۱۲]` is a blank nobody can fill, and `$۸۰۰` is not a real amount.

**Not a bug:** content that vanishes under a heading is Word collapsing that
heading, not the file. Click the triangle beside it.

# Part 2 — the writing

**Think in Persian and write. Never build the sentence in English first.** The
tell is unmistakable and the owner has named it: «این در یک‌طرفه نیست», «روی شن
ساخته می‌شود» — each one an English idiom carried across word by word. A Persian
reader does not think with them.

**Verbs carry Persian sentences, not nouns.** «نتیجهٔ موردِ انتظار» and «درجهٔ
اطمینان» are office-speak; «چه انتظاری داریم» and «چقدر مطمئنیم» are speech.
When a sentence stacks three nouns in a row, rewrite it around a verb.

**Short sentences with a verb.** A two-line sentence becomes two sentences. If a
paragraph can be cut without loss, cut it.

**Never translate an English idiom.** No equivalent? Say plainly what you mean.

**English words only where English is the honest answer:** a company name, an
email address, a file path, a clause label, a technical term with no Persian
equal. Everything else in Persian.

**Months — the rule has two halves, and using the wrong half is a real error.**
- **Persian month names** (آگوست, سپتامبر, نوامبر) belong in prose written *for*
  the owner: a report, an explanation, a summary of what happened.
- **English month names** (August, September, November) belong in anything that
  lives in the Canadian world or comes from an English source: a contract, a
  date that will be read by a bank or a lawyer or a government office, a
  deadline copied from an English letter, a translated article. Writing
  «۲۱ آگوست» on a document a Canadian office will act on is worse than useless —
  it does not match the source and cannot be checked against it.
- **Never French month names** (août, septembre). Canada is bilingual; the
  owner's documents are not. English.
- A date that must stay verifiable keeps its original form: `August 21, 2026`.

**In a legal or contractual document, add one line under each clause saying what
that clause actually does.** The reader should never have to decode a clause to
learn its effect. That line is Persian, plain, and about consequence — «اگر این
بند نباشد، نصفِ نرم‌افزار مالِ او می‌شود» beats a restatement of the clause.

**Say the number.** «مبلغِ قابلِ توجهی» tells the reader nothing; «۷۸ هزار دلار»
tells them everything. Where the number is not known yet, leave the `[...]`
placeholder — never paper over a gap with an adjective.

**Read the finished paragraph aloud in your head.** If you stumble, the reader
will too, and the sentence goes back.

## Rules for the content itself

- Keep an English term in English when it has no honest Persian equal — a
  clause name, `escrow`, a file path, a company name. Everything else Persian.
- Put the English term inside the Persian sentence, not in a line of its own;
  the run-level RTL handles the ordering.
- Placeholders stay as `[...]` and get listed once at the top of the document,
  so the reader sees in one place what is still unfilled.
- Months in Persian: آگوست, سپتامبر — never August.

## Fonts

**Vazirmatn** for body and headings by default — SIL OFL, free to use and
redistribute: https://github.com/rastikerdar/vazirmatn/releases

A document names a font; it does not carry it. If the reader's machine lacks
Vazirmatn, Word substitutes and the page stops looking Persian. Check first and
say so rather than shipping silently:

```python
from persian_docx import fonts_present
fonts_present()          # {"Vazirmatn": True}
```

To set a display face for headings — only one you hold a licence for; **B Titr
and most Persian display faces are not free to redistribute**:

```python
import os; os.environ["PERSIAN_DOCX_TITLE_FONT"] = "B Titr"
```

or set `FONT` / `TITLE_FONT` in the module. Nothing else in the skill is
machine-specific.

## Check it

```bash
python persian_docx.py out.docx      # writes a demo covering every feature
```

Writes a demo file with a Persian paragraph, an English term, a number and a
placeholder — open it and every one of the four should look right.
