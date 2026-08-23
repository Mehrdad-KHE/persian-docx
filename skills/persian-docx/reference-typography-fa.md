# Persian writing conventions

The rules `persian-docx` follows, and the ones it can enforce automatically.
`normalize()` applies everything marked **auto**; the rest is on the writer.

## Letters — **auto**

- **ی and ک must be Persian**, never the Arabic ي and ك. They look almost
  identical and break search, sorting and every font that lacks the Arabic
  glyph. This is the single most common defect in Persian text produced by
  software.
- **Arabic-Indic digits ٠١٢٣ are not Persian digits ۰۱۲۳.** Four of them differ
  in shape.
- **ة at the end of a Persian word becomes ه** («سرمایة» → «سرمایه»), except in
  Arabic phrases kept intact.

## نیم‌فاصله — the half-space (ZWNJ) — **auto**

Persian words join without a space but must not fuse. The zero-width non-joiner
does that, and text without it reads as broken:

| written with a space | correct | wrong (fused) |
|---|---|---|
| می شود | می‌شود | میشود |
| کتاب ها | کتاب‌ها | کتابها |
| نمی توان | نمی‌توان | نمیتوان |
| خانه ی من | خانهٔ من | خانه ی من |

Applies to: می/نمی prefixes · ها/های/هایی plurals · تر/ترین · ‌ام/‌ات/‌اش
suffixes · بی/هم prefixes.

## Punctuation — **auto** where marked

- **Comma is ، not ,** and semicolon is ؛ not ;. **auto**
- **Question mark is ؟ not ?** **auto**
- **Quotes are « » not " "** — and nothing else. Not “ ”, not ' '. **auto**
- **No space before a punctuation mark, one space after.** «سلام ،» is wrong;
  «سلام،» is right. **auto**
- **Ellipsis is a single … or three dots, never four or more.** **auto**
- **Parentheses hug their content**: (مانند این) — no inner spaces. **auto**
- **Em dash —** for a break in thought, spaced on both sides. At most one per
  paragraph; a second means the sentence should have been two.
- **Colon** introduces a list or a quotation, nothing else.

## Numbers

- **Persian digits in prose**: ۱۴۵۰، ۲۳ درصد.
- **Thousands separator is ٬** (U+066C), not the Latin comma: ۷۸٬۰۰۰.
- **Latin digits stay Latin** inside a version number, a file name, a phone
  number, a postal code, an amount in a foreign currency (`$12,500`), and inside
  a `[...]` placeholder — a blank reading `[۱۲]` is one nobody can fill.
- **Percent**: «۲۳ درصد» in prose; the ٪ sign only in tables.

## Bullets and lists

- **The marker is •** — Word's own list renders as `/` in an RTL document, and
  an arrow glyph reads as Word's heading-collapse triangle.
- **A bullet is not a sentence fragment with a full stop.** Either every item
  is a complete sentence ending in a dot, or none is. Do not mix.
- **Nested lists in Persian rarely help.** Two levels is the limit; past that,
  use headings.
- **Numbered lists take Persian digits with a dot**: «۱. …» — not «1)» and not
  «۱-» which reads as a range.

## Spacing and layout

- **One space between words, never two.** **auto**
- **No blank line inside a paragraph** — spacing is a paragraph property.
- **Body text justified.** A ragged edge is what makes a Persian page look
  unfinished.
- **Headings are never justified** — they align to the reader's side.

## English inside Persian

- The word keeps its Latin script and sits inside the Persian sentence; the
  run-level RTL puts it in the right place. Do not park it on its own line.
- **No Persian plural on an English word** — «فایل‌ها» not «فایلs».
- A term with an honest Persian equivalent uses the Persian one. Keep English
  for a name, an address, a path, a clause label, or a term with no equal.

## Words that mark a translation, not Persian

| translated | Persian |
|---|---|
| بدینوسیله اعلام می‌گردد | اعلام می‌کنیم |
| مورد بررسی قرار گرفت | بررسی شد |
| نتیجهٔ مورد انتظار | چه انتظاری داریم |
| در راستای | برای |
| اتخاذ تصمیم نمود | تصمیم گرفت |
| می‌باشد | است |

Passive voice and noun stacks are where translated Persian hides. Persian runs
on verbs.

## Automatic check

```python
from persian_docx import normalize, check_text
normalize("می شود و کتاب ها ,")   # → «می‌شود و کتاب‌ها،»
check_text(text)                   # lists what normalize cannot decide for you
```
