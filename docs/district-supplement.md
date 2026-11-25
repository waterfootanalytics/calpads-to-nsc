# Prepare the optional district supplement

Use this CSV to add recorded diploma dates, known FERPA blocks, or optional academic and program information. The certified CALPADS 1.23 reports supply the graduate roster and the optional demographics. The guide matches each supplemental record to a graduate by California Statewide Student Identifier (SSID).

## Choose the district fields

Include `SSID` and at least one other heading from the table. You can keep using a dates-only or dates-and-blocks CSV and leave all seven academic and program fields empty. A blocks-only file can contain just `SSID,FERPABlock`.

| CSV heading | Select from district records |
| --- | --- |
| `SSID` | California’s ten-digit SSID. Keep it as text, including leading zeroes. |
| `GraduationDate` | Date the diploma was awarded. Use `YYYY-MM-DD`, or leave it blank when unknown. |
| `FERPABlock` | The district’s recorded block applicable to StudentTracker. Use `Y` for a known block, `N` for a documented no, or blank for unknown. |
| `Grade8MathAssessment` | Eighth-grade state math assessment on the NSC scale (column S). |
| `Grade8ELAAssessment` | Eighth-grade state English language arts/reading assessment on the NSC scale (T). |
| `HighSchoolMathAssessment` | Final high-school state math assessment on the NSC scale (U). |
| `HighSchoolELAAssessment` | Final high-school state English language arts/reading assessment on the NSC scale (V). |
| `MathSemesters` | Total semesters of math completed for credit, including qualifying nontraditional terms. Enter a whole number from 0 through 99 (X). |
| `DualEnrollment` | Whether the graduate was enrolled in both high-school and postsecondary coursework (column Y). Enter `Y` or `N`. The guide also accepts Yes and No. |
| `ProgramCode` | A district-defined label returned in the student-level detail report. Enter up to 50 printable characters, with no quotes, dashes or hyphens. Separate multiple codes with an underscore or period, for example `AVID.FAFSA` (AA). |

For all four assessments, map the statewide assessment’s reported performance levels to the NSC scale. Use `1` for below standard, `2` for meeting the standard, or `3` for exceeding the standard. Leave a cell blank when no result or verified mapping is available. [NSC graduates-file guide](https://help.studentclearinghouse.org/sths/knowledge-base/formatting-and-submitting-a-graduates-file/)

Leave optional fields blank when they do not apply. A program label cannot be the placeholder `N/A`. The tool also rejects a leading `=`, `+` or `@` in a program label because the review step opens the text file in a spreadsheet. [NSC submission guidance](https://help.studentclearinghouse.org/sths/knowledge-base/common-graduates-file-submission-errors/) · [NSC program code guidance](https://studentclearinghouse.org/academy/courses/studenttracker-for-high-schools/lessons/file-submission-program-code/)

To prepare the full set of columns, use this header on a single line.

```csv
SSID,GraduationDate,FERPABlock,Grade8MathAssessment,Grade8ELAAssessment,HighSchoolMathAssessment,HighSchoolELAAssessment,MathSemesters,DualEnrollment,ProgramCode
```

If a district records blocks separately, join the block list to the other fields by SSID. Include graduates with a known block even when their diploma date is blank. If more than one row matches a graduate, a known block on any matching row is retained. Academic and program values are combined field by field. Identical entries and blanks are fine. When two nonblank values for one field disagree, the guide identifies the field and both supplement row numbers so you can correct the CSV.

## Export from your student information system

The export steps and field names below were checked against the linked vendor documentation. Use the diploma and FERPA fields your district maintains.

### PowerSchool SIS

1. Select the graduating students for the relevant year, including those moved to the district’s graduated-students school where applicable. Open **Quick Export → Fields**. PowerSchool’s California reporting documentation identifies the state SSID as `[Students]State_StudentNumber`. [California SSID field](https://ps-compliance.powerschool-docs.com/pssis-ca/latest/csis-reporting)
2. Add the district’s verified diploma-award field and, if available, its applicable block field. These can live in student extensions. Keep column headings, choose a comma delimiter and export the selection. Rename the columns as listed above. [PowerSchool Quick Export](https://ps.powerschool-docs.com/pssis-admin/latest/quick-export)

### Aeries

1. Open **Query** for the district or schools and year that contain the graduate records. Include inactive students. Start a `LIST` query with the `STU` table and apply your district’s graduate selection. Select `STU.CID` for the state ID. For awarded high-school diplomas, select `STU.DG` (School Completion Date), which Aeries defines as the graduation date for high schools. The separate `STU.EGD` field holds the expected graduation date. [Aeries student field definitions](https://support.aeries.com/support/solutions/articles/14000113683-aeries-api-student-related-end-points) · [Including inactive students](https://support.aeries.com/support/solutions/articles/14000094525-student-data-management-queries)
2. Add a block field only after verifying its meaning for your district, or bring in a separate block list by SSID. Use **Excel** in Query to export the selected fields. Save the prepared columns as CSV. [Aeries Query instructions](https://support.aeries.com/support/solutions/articles/14000081220-beginning-query)

### Infinite Campus

1. In **Filter Designer**, create or select a saved student-data filter. California reporting uses State ID (`Person.stateID`) for SSID. The Graduation record’s Diploma Date (`graduation.diplomaDate`) records the credential award. Include the applicable district-verified block field or merge a separate block list by SSID. [Infinite Campus California SSID mapping](https://kb.infinitecampus.com/help/calpads-ssid-enrollment-senr) · [Graduation fields](https://kb.infinitecampus.com/help/graduation)
2. Check year and status criteria using **Search** to verify the filter includes graduates. Open **Data Export**, select the saved filter and select the calendars covering the intended schools and year. Choose **Delimited Values (CSV)**, then **Export**. Align the exported headings and flag values with the table above. [Infinite Campus Data Export](https://kb.infinitecampus.com/help/data-export)

### School Pathways

1. Open **Admin → Create A Report**. Start with **Students**, select the state ID and link any additional tables needed for your district’s verified diploma and block fields. Set restrictions for the graduating class, including students who have left. [School Pathways student report](https://help.schoolpathways.com/knowledge/student-data-report-sample?hsLang=en) · [State ID and SSID](https://help.schoolpathways.com/knowledge/ccgi-data-export)
2. Use **Field Exported As** for the required headings. Review the restrictions in the report summary, run the report and save the selected columns as CSV. [School Pathways Create A Report](https://help.schoolpathways.com/knowledge/create-a-report-basics-video)

### Q by Aequitas Solutions

1. Open **Reports** under the relevant Q menu and select the district report or saved **Report Set** that exposes the state ID and verified diploma and block fields. Check the student and school selections for the reporting year. A report set saves report settings. If the report lacks a needed field, use a report or a separate SSID-matched file that contains the field.
2. In a report that supports export, choose **Export (Excel)** under **Report Output**, then **Create Report**. Keep the required columns, rename the headings and save as CSV. [Aequitas Q reports](https://support.myaequitas.com/support/solutions/articles/154000229161)

## Check the export and save it

1. Before saving, confirm with district records that the selected ID is the state ID, the date is the award date, and the applicable block values are mapped correctly. Check a known blocked record if the district has one. In a spreadsheet, import the SSID column as **Text** to preserve leading zeroes.
2. Save a UTF-8 CSV with headings on row one, for example `District supplement.csv`, directly in the working folder’s `input` directory. Put intermediate exports elsewhere so the guide can recognize the final file. [Accepted columns and date rules](file-formats.md#district-supplement)
3. Select the supplement when preparing a Graduate test file. Compare columns L and M and any academic or program columns you supplied against the source. The test picks one graduate per school. To check the wider cohort before answering date questions, run the grouped date comparison. Check that the district’s known blocks are represented in the full file before submitting. [Date comparison and full-file options](command-line.md#compare-exit-dates-with-diploma-dates)
