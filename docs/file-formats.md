# File formats and eligibility

[Getting started](../README.md)

## CALPADS report

Use the certified **1.23 Graduates and Completers Student List** for the reporting cycle and schools covered by your NSC agreement. A separate file for each school and a single combined student list both work. Every student must appear only once across the selected reports. The program reads the school in the student’s row, so a file name or a current enrollment in a separate system cannot change the graduating campus. [CALPADS Report 1.23 definition](https://documentation.calpads.org/Reports/EOY2/Report1.23_GraduatesandCompletersStudentList/)

Export a flat table with the column headings in the first 25 lines of a text report or the first 25 rows of the worksheet. A report title, filter details or blank lines may precede the headings. Supported formats include a single-worksheet `.xlsx` and delimited `.csv`, `.tsv` or `.txt`. Text files can use commas, tabs, pipes or carets. Save text as UTF-8, or UTF-16 with its byte-order mark.

| Field in the report | How the program uses it |
| --- | --- |
| School Name | The CALPADS side of the school match. |
| SSID | Ten digits used to detect duplicate reports and match the optional supplement. Preserve leading zeroes. |
| Student Name, or separate legal first and last name columns | A combined name must use `Last, First Middle`. Middle name is optional. |
| Birth Date | Used for matching by NSC and for checking that the graduation date is plausible. |
| Student Enrollment Exit Date | Places the record in the selected cycle and supplies the default graduation date. |
| Student Completion Status | The three-digit code determines eligibility and the diploma type. A description after the code is supported. |
| Student Exit Reason, if present | Must identify a completer exit (`E230` or `Completer Exit`). |
| Local ID, if present | Used in NSC column J in local ID mode (the default). Supplied local IDs in that mode must be unique among included graduates and contain 1–15 letters or digits. |

A year of **2026**, for example, covers enrollment exit dates from August 16, 2025 through August 15, 2026, inclusive. The tool infers this year from all enrollment exit dates in the selected reports, including campuses that will later be excluded. Mixed cycles stop preparation. A scripted run can supply `--year` to check the expected cycle. The date on the computer and the diploma dates in the supplement do not determine the reporting year.

### Completion status

| CALPADS code | Default action |
| --- | --- |
| 100 | Include as `Regular Diploma`. |
| 120, 320, 330 | Exclude. |
| Any other code | Pause until the district specifies an NSC diploma type or explicitly excludes the code. |

For an unfamiliar code, check the credential awarded, then choose its NSC diploma type or exclude the code. The confirmed mapping applies to all graduates with that code for the agreement and reporting year. CALPADS code 250 represents an adult education high school diploma and requires this choice. [Completion settings](command-line.md#district-completion-settings) · [NSC accepted diploma types](https://help.studentclearinghouse.org/sths/knowledge-base/frequently-asked-questions/)

## Account Reference Sheet

Use the current Account Reference Sheet for the same service agreement as the reports. In the StudentTracker 3.0 web application, export it from **Accounts and Settings**. [NSC account features](https://help.studentclearinghouse.org/sths/knowledge-base/studenttracker3-and-2-key-differences-for-secondary-institutions/)

- **Excel** (`.xls` or `.xlsx`). The reader looks for a sheet with Organization Name, Service Agreement/Account Number, Total High School Count, and the table with ACT Code, High School Name and HS Operating Status. The count is checked against the school rows. Open campuses are offered as choices.
- **PDF** (`.pdf`). The reader supports a text-based reference sheet with an account name and number, participating organization count and numbered school table. Poppler’s `pdftotext` must be installed. For a scanned sheet or a different layout, export the Excel sheet. [PDF setup](troubleshooting.md#the-pdf-account-sheet-will-not-open)

A confirmed school exclusion omits that campus from the test, full file and a comparison that uses the workbook. The exclusion is tied to its reporting cycle. For another cycle, create a fresh review. A flag leaves the decision pending.

A unique exact match after name normalization fills in the choice. A similar spelling is a suggestion in column C. The reviewer selects the registered campus in D.

The crosswalk matches a CALPADS campus to the **six-digit ACT code** on this sheet. The program writes the registered school name associated with that code. If that spelling exceeds NSC’s 50-character field limit, the program shows the school and length, keeps the spelling from the sheet and allows generation. NSC may return a validation error. Run the Graduate test before submitting the full file. [NSC graduates file specification](https://help.studentclearinghouse.org/sths/knowledge-base/formatting-and-submitting-a-graduates-file/)

## District supplement

A supplement can provide district diploma dates, known FERPA blocks and the seven optional academic and program fields. Export **one CSV** with `SSID` and at least one supported column. The three-column example below works for dates and blocks. [Prepare the supplement from your SIS](district-supplement.md)

```csv
SSID,GraduationDate,FERPABlock
```

| Column | What to enter |
| --- | --- |
| `SSID` | The ten-digit California Statewide Student Identifier, saved as text. |
| `GraduationDate` | The actual diploma date. Prefer `YYYY-MM-DD`. The parser also accepts `YYYYMMDD` and US month/day dates with a four-digit year. Leave unknown dates empty. |
| `FERPABlock` | `Y` for a known block, `N` for no block, or blank for unknown. This column is optional. For a dates-only file, use just `SSID,GraduationDate`. |

Put privacy flags in the supplement. If a locally modified CALPADS report contains a populated `FERPABlock` column, preparation pauses so you can transfer those flags to the supplement and use the original certified export.

For blocks only, use `SSID,FERPABlock`. If your source stores flags as Yes/No or numeric codes, verify their local meanings before converting them to Y/N. NSC’s output value is `N` for no known block or unknown. Include every known block for this cohort when preparing the full file. For the full file, the guide asks for confirmation of no known blocks if the selected inputs provide no block column. A three-column supplement can contain blank flags for unknown status. [NSC privacy guidance](https://help.studentclearinghouse.org/sths/knowledge-base/ferpa-compliance/) · [NSC file submission guide](https://help.studentclearinghouse.org/sths/knowledge-base/formatting-and-submitting-a-graduates-file/)

### How the supplement affects each graduate

| What is on file for an SSID in the CALPADS roster | Graduation date in NSC column L | FERPA block in NSC column M |
| --- | --- | --- |
| One distinct, usable diploma date within the reporting cycle | District diploma date | `Y` if any matching row is marked `Y`, otherwise `N`. |
| Duplicate rows with the same usable in-cycle date, or that date plus only blank, unreadable or implausible entries | District diploma date | Same rule. |
| One plausible diploma date outside the reporting cycle | Review the indicated source rows. Confirm the verified diploma date, choose the CALPADS exit date, or correct the source and retry. | Same rule. |
| Multiple distinct plausible diploma dates for one SSID | Correct the supplement to one verified date, or explicitly choose the CALPADS exit date. | Same rule. |
| Only blank or unreadable dates, or no date in the birth-to-transmission range | CALPADS exit date. | Same rule. |
| No matching row or no supplement | CALPADS exit date | `N` for a standard 1.23 export. |

Test and full files require a transmission date on or before today. A diploma date is plausible when it falls after birth and on or before the file transmission date. A plausible date outside the exit cycle can be correct. NSC requests the actual diploma date. When the district confirms one, the file header includes the graduation year or range present in the submitted detail rows.

Only eligible graduates from included CALPADS schools enter the NSC file. District-wide supplement rows with an empty or malformed SSID and a blank or negative block are skipped with a notice. If such a row carries a known block, preparation stops and identifies the source row. An invalid block value also stops preparation so the district can correct it. A FERPA block on any matching row remains in effect whichever graduation date is selected. In the guide, **B** uses the saved supplement for blocks only and keeps CALPADS dates for everyone. **S** skips every field in that supplement for this run.

NSC requests a high school graduation date. When exit dates and diploma dates differ in your district, supply the recorded diploma dates wherever possible and review the fallback cases. The optional comparison groups the cohort by source dates and decision status. You can preview pending decisions before choosing a date or mapping a completion code. [How to compare dates](command-line.md#compare-exit-dates-with-diploma-dates) · [NSC required fields](https://help.studentclearinghouse.org/sths/knowledge-base/frequently-asked-questions/)

## Graduates text files

The generated UTF-8 text file has tab-separated fields and Windows-style line endings. The same record layout is used for the sample and full file. The header starts with `PH3` and contains the agreement, organization, transmission date and the calendar year or range of years of the submitted graduation dates. Each graduate row starts with `PD3` and ends with `ED`. The trailer starts with `PT3` and counts **all lines**, including the header and trailer. [NSC trailer lesson](https://www.studentclearinghouse.org/academy/courses/studenttracker-for-high-schools/lessons/file-submission-create-the-trailer-row-1-minute/) · [NSC file submission guide](https://help.studentclearinghouse.org/sths/knowledge-base/formatting-and-submitting-a-graduates-file/)

| Excel column in a student row | What to check |
| --- | --- |
| B | The literal `NO SSN`. |
| C–E | First, middle and last legal names in that order. |
| I | Date of birth in `YYYYMMDD`. |
| J | Optional student identifier. The default writes the CALPADS local ID when available. The district can choose the ten-digit SSID, or leave J blank for everyone. Matching the supplement always uses the SSID. |
| K, L | Diploma type and selected graduation date in `YYYYMMDD`. |
| M | FERPA block indicator. A known block stays on the record as `Y`. |
| N, O | Registered school name and six-digit ACT code. Import O as text to preserve leading zeroes. |

Name formatting affects only the generated file. The formatter folds supported accents and letter combinations into Latin letters, normalizes spaces, apostrophes and hyphens, removes other punctuation, and limits each name field to 40 characters. A warning identifies shortening. If a required name becomes empty or includes letters the formatter cannot represent, the program identifies the source report and row for human review.

The Graduate test selects the first eligible graduate for each included campus and permits up to 1,000 student rows. The full file covers all eligible graduates and permits up to 250,000 student rows. [NSC test uploads](https://help.studentclearinghouse.org/sths/knowledge-base/submitting-a-test-file/) · [NSC field requirements and limits](https://help.studentclearinghouse.org/sths/knowledge-base/frequently-asked-questions/)

### Optional demographics and programs

In the terminal guide, press **Return** to include available demographics or choose **2** to exclude them for this file. When included, the guide fills the following fields from each graduate’s certified 1.23 row. The same choices apply to test and full files. A missing column, an empty cell or an unrecognized value leaves the corresponding NSC field blank. The terminal identifies fields with unsupported entries.

| NSC column | CALPADS source | Values written |
| --- | --- | --- |
| P · Gender | Gender | `M` or `F`. NSC's layout has no code for nonbinary. |
| Q · Race/ethnicity | Ethnicity/Race | `AM` American Indian/Alaska Native, `AS` Asian, `BL` Black/African American, `HI` Hispanic, `PI` Hawaiian/Pacific Islander, `WH` White, `MU` multiple races. |
| R · Economic disadvantage | Socio-Economically Disadvantaged | `Y` or `N` from the reported indicator. |
| W · English learner | ELAS Status | `Y` for current English Learner (EL). `N` for English Only (EO), Initially Fluent (IFEP) or Reclassified Fluent (RFEP). An undetermined status stays blank. |
| Z · Disability | Special Education Status | Press **Return** to use the reported special education participation indicator. Choose **2** to leave Z blank for this file. |

Report 1.23 also contains indicators such as foster youth, homelessness and Seal of Biliteracy. The NSC layout has no dedicated columns for those measures. The program code in AA is a district-supplied label. [CALPADS report definition](https://documentation.calpads.org/Reports/EOY2/Report1.23_GraduatesandCompletersStudentList/) · [NSC graduates file guide](https://theclearinghouse.download/sthssubmitgrads)

The supplement supplies grade 8 math and English language arts assessment results in **S and T**, high school math and English language arts assessment results in **U and V**, math semesters in **X**, dual enrollment in **Y**, and the district’s program code in **AA**. Each field is optional and can be omitted or left blank. Matching by SSID applies only to graduates on the selected CALPADS roster. [Supplement headings and accepted values](district-supplement.md#choose-the-district-fields)

## Sources checked for the 2026 cycle

| Source | Checked | Coverage |
| --- | --- | --- |
| [CALPADS Report 1.23, updated May 1, 2026](https://documentation.calpads.org/Reports/EOY2/Report1.23_GraduatesandCompletersStudentList/) | November 24, 2026 | Enrollment exit window, report headings and completion codes. |
| [NSC Graduates File Guide](https://theclearinghouse.download/sthssubmitgrads) and [NSC field FAQ](https://help.studentclearinghouse.org/sths/knowledge-base/frequently-asked-questions/) | November 24, 2026 | Diploma dates, detail layout, codes and graduation period. |
| [Clearinghouse Academy trailer lesson](https://www.studentclearinghouse.org/academy/courses/studenttracker-for-high-schools/lessons/file-submission-create-the-trailer-row-1-minute/) | November 24, 2026 | The trailer count includes the header, each detail row and the trailer itself. |
