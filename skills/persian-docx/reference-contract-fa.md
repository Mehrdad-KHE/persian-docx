# How a Persian contract is written

Reference for `persian-docx` when the document is a contract, an agreement, a
memorandum of understanding or anything a lawyer will read. Distilled from
standard Iranian commercial templates. It is about **form**, not law.

## The skeleton

A Persian contract runs in this order, and a reader looks for it in this order:

1. **عنوان قرارداد** — a noun phrase, not a sentence: «قرارداد مشارکت مدنی»،
   «قرارداد واگذاری حق استفاده از نرم‌افزار».
2. **مقدمه و طرفین** — who is bound. Each party gets full legal name, national
   ID or registration number, address and contact, then the short label used
   throughout: «که از این پس "طرف اول" نامیده می‌شود».
3. **ماده ۱ — موضوع قرارداد.** One paragraph. What the contract is *for*.
   Everything after it elaborates.
4. **مواد** in a fixed sequence: مدت · مبلغ و نحوهٔ پرداخت · تعهدات هر طرف
   (each party its own ماده) · ضمانت اجرا · فسخ · فورس‌ماژور · حل اختلاف ·
   نشانی و ابلاغ.
5. **مادهٔ پایانی — نسخ.** The count-and-force line, always last:
   «این قرارداد در [تعداد] ماده و [تعداد] تبصره، در [تعداد] نسخهٔ واحدالمتن و
   دارای اعتبار یکسان تنظیم شد و به امضای طرفین رسید.»
6. **امضاها** — a labelled block per party, with date. Witnesses if the deal
   warrants them.

## ماده and تبصره

- **ماده** is a numbered article — one obligation or one subject each. When a
  ماده needs two unrelated things, it should have been two مواد.
- **تبصره** is an exception or a qualification *to the ماده directly above it*,
  numbered within it. It never introduces a new subject and never stands alone.
- Numbering is Persian digits: ماده ۱، ماده ۲، تبصرهٔ ۱.
- A ماده that fills half a page is a drafting failure, not thoroughness.

## Sentence habits that mark a real contract

- **Present tense, obligation form**: «متعهد می‌شود»، «موظف است»،
  «حق دارد»، «مجاز نیست». Not «خواهد بود» or conditional softening.
- **Define once, then use the label.** After «که از این پس "نرم‌افزار" نامیده
  می‌شود», the word never gets described again.
- **No adjectives of degree.** «مبلغ قابل توجهی» has no place; the number does.
- **No English legalese carried across.** «طرفین بدینوسیله توافق می‌نمایند» is
  a translation of *the parties hereby agree* and reads like one. «طرفین توافق
  کردند» is Persian.
- **Amounts written twice** where they matter: digits then words —
  «۷۸٬۰۰۰ (هفتاد و هشت هزار) دلار کانادا». Currency always named.
- **Dates carry their calendar.** In a Canadian contract keep the Gregorian
  date in English form (`August 21, 2026`); add the Persian equivalent only if
  both parties actually use it.

## Typography, on top of the general rules in SKILL.md

- ماده headings are `h2`; تبصره sits in the body under its ماده, not as a
  heading.
- Party labels and defined terms in **bold** on first appearance only.
- Blank placeholders as `[...]`, never a row of dots — dots invite someone to
  write on the wrong copy.
- Signature block at the end, one line per party, right-aligned like everything
  else, with room under each name for a date.

## Where a translated contract differs, and why it matters here

A contract governed by Ontario law that is being *read* in Persian is not an
Iranian contract. Keep the source document's own structure and clause
numbering — the Persian version has to be checkable line by line against the
English one that gets signed. Borrow the Persian conventions for **language**
(obligation verbs, defined labels, the نسخ line), not for **structure**.

State plainly at the top which text governs:
«متن انگلیسی مبنای امضا و تفسیر است؛ این ترجمه برای فهم طرفین تهیه شده است.»

Sources consulted for the conventions above:
edalatsara.com/letter-of-partnership · gharardadha.com · moshirin.com
