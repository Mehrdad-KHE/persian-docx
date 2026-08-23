# persian-docx

An agent skill for writing Word (`.docx`) documents in Persian that are actually
correct — right-to-left throughout, real heading styles, justified body text,
Persian digits and quotation marks, and English terms that stay where they
belong in the line.

Ask any coding agent to "write me a Word file in Persian" and it reaches for
`python-docx`. The result opens, looks Persian, and is wrong in four or five
ways at once. This skill is the accumulated fix.

---

## The problem, concretely

`python-docx` has no RTL support at all. Nothing in its API sets paragraph
direction, and its `alignment` property writes a value Word reinterprets in a
bidi context. So a file built the obvious way comes out like this:

- The last word of a Persian sentence sits on the **left**.
- Headings hug the **wrong margin** — while the body text looks fine, so nobody
  notices until the document is in front of a client.
- The font you asked for is **ignored** for every Persian glyph.
- Quotation marks are `"..."` instead of `«...»`.
- Bullets render as `/`.
- Numbers are `1450`, not `۱۴۵۰`.
- Word sees **no headings at all**, so there is no navigation pane and no table
  of contents — the "headings" are just bold paragraphs.

Each of these has a one-line fix in the XML. Finding all of them takes a day.

---

## Install

```bash
npx skills add Mehrdad-KHE/persian-docx@persian-docx -g
```

Or clone and point your agent at the folder. The only dependency is
`python-docx`:

```bash
pip install python-docx
```

Python 3.8+. Windows, macOS and Linux.

---

## Use

```python
from persian_docx import PersianDoc

d = PersianDoc()
d.title("قرارداد شراکت")
d.h1("بخش اول — طرفین")
d.h2("۱. تعریف‌ها")
d.p("این قرارداد میان دو طرف بسته می‌شود و مبلغ آن [مبلغ] است.")
d.h3("۱.۱ نرم‌افزار")
d.bullet("مالکیت نرم‌افزار منتقل نمی‌شود.")
d.bullet("شرکت فقط پروانهٔ استفاده می‌گیرد.")
d.page_break()
d.save("out.docx")
```

### API

| call | what it gives you |
|---|---|
| `PersianDoc(digits=True)` | a document with RTL sections, Vazirmatn body text and 2.5 cm margins. `digits=False` keeps Latin numerals |
| `.title(text)` | Word's `Title` style, centred, display face |
| `.h1 / .h2 / .h3 (text)` | real `Heading 1/2/3` styles, rewritten RTL |
| `.p(text, justify=True)` | body paragraph. `justify=False` for a short standalone line |
| `.bullet(text)` | a bulleted line with the marker and indent handled |
| `.page_break()` · `.save(path)` | |
| `audit(path)` | opens a finished file and lists what is still wrong |
| `fonts_present(*names)` | which font families are actually installed |

### Check before you ship

```python
from persian_docx import audit
print(audit("out.docx"))     # must be []
```

The auditor is the point of the skill as much as the builder is. It re-opens a
finished document and reports missing `w:bidi`, wrong alignment values, runs
without the complex-script font, straight quotes left behind, and a document
with no real headings. Run it on files other tools produced, too — it is a
useful diagnostic on its own.

---

## What it fixes, and why each one matters

**`w:bidi` on every paragraph, run and section.** Paragraph direction is not
alignment. Without the flag the paragraph is LTR no matter how you align it, and
punctuation drifts to the wrong end of the line.

**`w:jc` set to `start`, never `right`.** This is the one that catches everyone.
Inside a bidi paragraph Word reads `right` as the *logical* right, which lands
on the physical **left**. Justified body text looks correct either way — both
edges are flush — so the bug hides in the body and only shows in the headings.
Use `start` (the reader's side), `both`, or `center`. The auditor rejects
anything else.

**The complex-script font slot (`w:rFonts w:cs`).** Word picks the typeface for
Arabic-script glyphs from a different slot than the Latin one. Setting
`run.font.name` leaves every Persian character in Word's default and your font
choice does nothing.

**Real heading styles.** Bold text in a `Normal` paragraph is not a heading.
Word builds the navigation pane, the table of contents and the document map from
styles — and its built-in heading styles are LTR with their own Latin font, so
each one has to be rewritten before first use.

**Justified body text.** In Persian typesetting a ragged edge reads as
unfinished far more than it does in English.

**`«` `»` instead of `"`.** Straight quotes are simply not Persian punctuation.
The builder converts them; the auditor flags any that survive.

**A written `•` marker instead of Word's list.** Word's own bullet renders as
`/` in an RTL Persian document. An arrow glyph is worse — it reads as Word's
heading-collapse triangle and the reader tries to click it.

**Persian digits, except inside `[...]` and after `$`.** A blank that reads
`[۱۲]` is a blank nobody can fill, and `$۸۰۰` is not an amount anyone can bank.

---

## The other half: the writing

A file can be flawless right-to-left and still unreadable, because the sentences
were built in English and carried across word by word. The skill carries
guidance for that too, and it matters more than the mechanics:

- **Think in Persian, do not translate.** Rendered idioms are unmistakable to a
  native reader and invisible to the person who wrote them.
- **Verbs carry Persian sentences, not stacked nouns.** «چه انتظاری داریم» over
  «نتیجهٔ مورد انتظار».
- **Short sentences.** A two-line sentence becomes two.
- **English only where English is the honest answer** — a company name, a file
  path, a term with no equivalent.
- **Months have two rules.** Persian month names in prose written for a Persian
  reader; English month names in anything a bank, a lawyer or a government
  office will act on, and in anything translated from an English source — a date
  that cannot be checked against its original is worse than useless. Never
  French.
- **In a contract, one plain line under each clause saying what it does** — its
  consequence, not a restatement.
- **Say the number.** "A significant amount" tells the reader nothing.

---

## Fonts

**Vazirmatn** (SIL OFL, free to use and redistribute) is the default for both
body and headings: https://github.com/rastikerdar/vazirmatn/releases

A `.docx` names a font; it does not carry it. If the reader's machine lacks the
font, Word substitutes silently and the page stops looking Persian. Check:

```python
from persian_docx import fonts_present
fonts_present()                       # {"Vazirmatn": True}
fonts_present("Vazirmatn", "B Titr")  # check any family
```

To give headings a display face:

```python
import os
os.environ["PERSIAN_DOCX_TITLE_FONT"] = "B Titr"
```

Use only a font you hold a licence for — most Persian display faces are **not**
free to redistribute, which is why the default ships as Vazirmatn for
everything.

---

## Not a bug

Content that disappears under a heading is Word collapsing that heading, not the
file. Click the triangle beside it. (This one cost an hour, so it is written
down.)

---

## Contributing

Issues and pull requests welcome — particularly from Persian, Arabic, Hebrew and
Urdu users. The RTL mechanics are shared across all four scripts; the typography
conventions are not, and this skill currently encodes Persian ones. Arabic and
Hebrew variants would be a natural addition.

If you hit a rendering problem this skill does not catch, please open an issue
with the offending `.docx` — every rule here came from a file that looked wrong
on someone's screen.

---

## Author

Mehrdad Kheirollahi. Built while drafting a Persian contract, after fixing the
same four things by hand one time too many.

## Licence

MIT
