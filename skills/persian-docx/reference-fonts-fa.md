# Using Persian fonts: the conventions

Which face, which size, which leading — the norms Iranian practice actually
follows. `KINDS` in the module encodes the size and leading rows.

## Size: Persian is set larger than Latin

Persian letterforms sit smaller inside the em and depend on dots and diacritics
to tell letters apart, so the readability threshold is higher. Every bilingual
Iranian spec encodes the gap: Persian ≈ Latin + 1 to 3 points in the same slot.

| document | Persian body | Latin in the same line | leading |
|---|---|---|---|
| contract | 12 | Times New Roman 11 | 1.5 |
| official letter (نامهٔ اداری) | 13–14 | 11–12 | 1.5 |
| thesis (پایان‌نامه) | 13 | TNR 12 / Arial 10 | 1.5 |
| book | 13.5 | 12 | 1.3 |
| web | 16 px | 15 px | 1.8–2.0 |

Sizes relative to a 13 pt body: footnotes 11 · captions 9 **bold** · text
inside tables ≤ 12 · page numbers and running heads 10 · abstract 12.
Captions go smallest and take bold to compensate.

**Where sources disagree:** administrative-letter body is quoted as 12, 13 and
14 depending on the source, and typing services default contracts to 16 because
they expect photocopying. For a contract that will be printed and signed, 12–13
is what a notary actually sees.

Source for the thesis column: Irandoc «راه», table 3-1 — the national
thesis-registration guide most university شیوه‌نامه‌ها are written against.

## Leading

Persian ascenders and descenders reach further than Latin ones, and the dots
under ب پ ی collide with the line below at Latin leading (W3C alreq).

- Word documents: **1.5**. Some university guides say 1.15; abstracts single.
- Print books: **1.3** — the measure is narrower.
- Web: **1.8–2.0**, never below 1.5. The browser default of about 1.2 is too
  tight for Persian.

## Headings against body

One display face for headings, one text face for body. The heading must be
visibly **heavier**, not slanted — Persian takes its hierarchy from weight and
family, never from slope.

A display face in body text is the most common Persian typographic mistake:
the weight that makes a title carry destroys a paragraph. B Titr, Lalezar and
every other display face belong in headings only.

Free pairings:
- **Vazirmatn Black over Vazirmatn Regular** — one family, no metric mismatch.
- **Lalezar over Vazirmatn or Estedad** — real contrast, still OFL.

## Bold and italic

**Bold (سیاه) is the emphasis device in Persian.** Do not bold book body text;
the administrative-letter tradition bolds nearly everything, which is a form
convention and not a typographic one.

Mechanically slanting a Persian face — what Word does when you press the italic
button — is wrong: an Arabic-script word has almost no vertical strokes, so the
slant barely registers and the letterforms distort. Persian editorial practice
does use **ایرانیک**, a purpose-designed oblique, for titles of books and
journals inside running text. If the face has no ایرانیک companion, use quotes
« » instead of faking a slant.

## Justification and kashida (کشیدگی)

Justify with **spacing first**. Kashida is one mechanism among six and the
weakest one to lean on: W3C alreq warns that relying on it alone produces
rivers, and that long kashidas give uneven colour. Tatweel is worse — it is a
fixed-width character, not an elastic one.

Use kashida only in a wide measure at a generous size in formal single-column
body. Never in tables, narrow columns, small sizes or UI. In Word, plain
Ctrl+J blows the word gaps open on Persian; **Justify Low** is the setting that
behaves.

## Mixing a Latin face with a Persian one

Do it in the style, not by hand. In Word: Ctrl+D → **Complex Script** field
takes the Persian face, **Latin Text** field takes the Latin one, and the sizes
are set independently so the Latin can run 1–3 pt smaller. This module already
writes both slots (`w:rFonts` `w:cs` and `w:ascii`/`w:hAnsi`).

Pick a Latin face whose x-height and stroke weight sit level with the Persian
face at the reduced size. Iranian specs pair Times New Roman with serif
documents and Arial with sans, and always set Arial about 2 pt smaller than
Times New Roman in the same slot.

## What Persian font articles get wrong

Most «best Persian fonts» round-ups recommend B Nazanin, B Lotus, B Mitra,
IRANSans and Dana without mentioning that none can be legally redistributed —
see the font table in SKILL.md. The typographic advice in those articles is
usually sound; the licence advice is absent.
