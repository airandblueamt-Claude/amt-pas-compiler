# AMT PAS — Document Specification

Reverse-engineered from the gold sample `2506038-TCS-009 v.00 PAS (Sample).pdf`
(a 56-page Public Address System submittal). Page size **A4** (595.32 × 841.92 pt).

## Page order
1. **Cover** (page 1)
2. **Table of Contents** (page 2)
3. For each of the 8 sections, in order: a **divider page**, then the section **content**.

## Chrome (on AMT-authored pages: cover, TOC, dividers, placeholders)
- **Header**: AMT primary logo, top-left (`assets/amt-logo.png`).
- **Footer**: `Ref.: <ref_no>-v.<version>` line (bottom-left) above the full-width
  AMT contact banner (`assets/amt-footer-banner.png`) — CR / Tel / Fax / P.O.Box /
  email / www.amt-arabia.com + the red "NETWORKING LOW CURRENT SYSTEM" diamond.
- **Company seal** (`assets/amt-seal.png`) placed centre/lower on cover, TOC and dividers.

> Rendered Excel tables (§1–4) and appended source PDFs (§5–8) keep their **own**
> footers (the Excel print footer / the vendor datasheet footer) — the AMT banner is
> NOT overlaid on them, matching the sample.

## Cover layout
- Optional client/project logo, centred near the top (`client_logo` in config).
- Title block (centred): Arabic title, English title, `Ref. #<mts_ref> v.<version>`,
  then the large Arabic company name.
- **Revision table** (blue header): Ref | Version | Date | Author | Remarks.
- **Sign-off table** (blue header): three columns — Ref BY | Prepared by | Approved by,
  with a roles row (Sales Coordinator / Document Controller / Presales Manager) and an
  initials row.

## Table of Contents layout
Blue header row: `S.No | Table of Content جدول المحتويات | Page.No`, then one row per
section: number, English title (bold serif), Arabic title (RTL), and the physical page
number where that section's divider falls.

## The 8 sections
| # | English | Arabic | Content source |
|---|---------|--------|----------------|
| 1 | Tender BOQ | جدول كميات العقد | `1-*/**.xlsx` |
| 2 | Material Sheet Provided by AMT | جدول المواد المقدمة من AMT | `2-*/**.xlsx` |
| 3 | Material Traceability Sheet & Tender BoQ Compliance | جدول تتبع المواد والمطابقة مع جدول كميات العقد | `3-*/**.xlsx` |
| 4 | Material Selection | اختيار المواد | `4-*/**.xlsx` |
| 5 | Product's Datasheet & Catalogue | النشرة الفنية والكتالوج للمنتجات | `5-*/**.pdf` (recursive) |
| 6 | Vendor Partnership Certificate | شهادة شراكة مع المورد | `6-*/**.pdf` (optional) |
| 7 | Warranty Certificate | شهادة الضمان | `7-*/**.docx`/`.pdf` |
| 8 | Layout and Single Line Diagram | مخطط أحادي لجميع الأنظمة | `8-*/**.pdf` + `.docx` (recursive) |

## Divider layout
AMT chrome + a single title row: `N-  <English>` on the left (bold serif) and the
Arabic title on the right (RTL), with the seal centred lower.

## Ordering rules for appended documents (§5–8)
- Shallower paths before deeper sub-folders, then case-insensitive filename.
- "master" plans before "part" plans.
- Ignore `*:Zone.Identifier` / `*:com.dropbox.attributes` streams and `~$` temp files.

## Pagination
Two-pass: render & measure all content first, compute each section's starting page,
then draw the TOC with the real page numbers. A QA check asserts the merged page count
equals the computed total.
