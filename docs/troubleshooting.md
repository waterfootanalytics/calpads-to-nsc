# Troubleshooting

[Back to the getting-started guide](../README.md)

## I closed the terminal and the command is missing

Each newly opened terminal needs to know where Python installed the tool. Open a terminal in the **downloaded source folder**, the one containing `pyproject.toml` and `.venv`. Reactivate the existing environment.

| System | Command in the source folder |
| --- | --- |
| macOS or Linux | `source .venv/bin/activate` |
| Windows PowerShell | `.\.venv\Scripts\Activate.ps1` |

Then change to your working folder and enter `prepare-nsc-graduates`. You can also use `python -m calpads_to_nsc` in the active environment. If the `.venv` folder is missing or the module cannot be found, repeat the installation steps from the folder with `pyproject.toml`. [Installation instructions](../README.md#download-and-install)

## PowerShell blocks activation

You can use the virtual environment’s Python directly without activating it. Open PowerShell in the extracted source folder. After creating the environment, use its Python to install and check the tool.

```powershell
$NscPython = (Resolve-Path .\.venv\Scripts\python.exe).Path
& $NscPython -m pip install .
& $NscPython -m calpads_to_nsc --help
```

Keep this PowerShell window open. Change to your working folder using `cd` and launch the guide with `& $NscPython -m calpads_to_nsc`. The variable holds the full path to Python, so it works after the folder changes. In a new window, repeat the first line from the source folder to set the variable again.

## The guide cannot find a file

1. Return to the working folder. Check in Finder or File Explorer that the current account sheet, official CALPADS exports and any optional supplement are directly inside its `input` folder.
2. Keep old exports, comparison CSVs and unrelated spreadsheets outside `input`. A blank or partial download can be replaced with a fresh export.
3. Launch the guide from the working folder or its `input` folder. Select **4 — Change input files** to choose the correct folder or file with your mouse. When the folder contains multiple candidate sheets or supplements, choose the one for this agreement and year.
4. If the supplement is missing, check that it is a CSV with `SSID` and at least one supported supplement heading on row one. For dates and blocks, use `SSID,GraduationDate,FERPABlock`. If you already have known blocks for this cohort, include them in the file before preparing the full submission.

On Windows or Linux the native file picker needs Python’s Tk support and a desktop session. On a headless system you can run the [command-line workflow](command-line.md) with explicit filenames. To finish the guide, enter **Q** at a terminal prompt. In a file dialog choose **Cancel** to return to the terminal.

## The PDF account sheet will not open

An original Excel Account Reference Sheet (`.xls` or `.xlsx`) can be used directly. For a text-based reference PDF, check the reader installation with this terminal command.

```sh
pdftotext -v
```

If this command cannot be found, install Poppler and make sure `pdftotext` is on your PATH. On a Mac with Homebrew, run `brew install poppler`. On Windows the executable is `pdftotext.exe`. You can also use the Excel export directly. [Poppler](https://poppler.freedesktop.org/) · [Homebrew package](https://formulae.brew.sh/formula/poppler)

On Windows, if `pdftotext -v` shows `xpdfreader.com`, run `where.exe pdftotext` to see which copy Windows finds first. Put the folder containing Poppler’s `pdftotext.exe` ahead of the Xpdf folder on PATH, reopen Terminal and check that the version output names the Poppler Developers.

If you can select and copy individual letters in the PDF, it contains text. The parser also needs the account labels and numbered school table in a supported layout. A scan or a different layout can be replaced with an Excel export from NSC’s **Accounts and Settings**.

## A school is flagged or its name has changed

Open **Match schools** and find the CALPADS name in the message. Use the **Look up a school** tab to compare candidate names and the six-digit ACT codes in the NSC agreement. For a renamed campus, use a district directory or records of its earlier name to verify the identity. Once you identify the campus, select it in the workbook dropdown and save. Matching a school to another account’s similarly named campus would change where its graduates are reported.

If the guide reports an unsupported workbook layout or a change to the account sheet or school list, create a fresh review workbook from the selected district inputs. Check the campus choices and exclusions in the new workbook. The workbook tracks the contents of the reference file, so reusing a filename does not carry approval to a different sheet. If an expected campus is absent from the account, verify the agreement and get a current reference sheet before the full submission.

### Two CALPADS schools share one ACT code

Each included CALPADS school needs its own registered ACT code in this tool. Check both rows against the current agreement and the district’s records of campus names and reorganizations. If both schools genuinely report under one code, keep them flagged while the district verifies the submission arrangement with its Clearinghouse contact. A shared-code submission needs a workflow that supports that arrangement. Choose **Exclude from this submission** only when the district actually intends to leave that campus out.

### Restore an excluded campus

Open **1 — Review school matches** in Terminal. Choose **M**, select the campus, then select its registered school and ACT code. In Excel, choose the registered school in column D and clear the exclusion reason and confirmation in F and G. Save the workbook and check the matches again.

### A similar name is shown but the school still needs a match

Column C shows a suggestion. Verify the campus and ACT code against district records, then select the registered school in the dropdown in column D. Keep the campus flagged if its identity is still uncertain.

## A registered school name exceeds 50 characters

Look at the ACT code and the spelling printed in the warning. If it matches that code on the current sheet, continue with the Graduate test file. The tool uses the registered spelling, and NSC can still flag the field length. If the test rejects it, use the test result’s proposed solution or obtain a corrected reference from NSC, then regenerate. [NSC test result instructions](https://help.studentclearinghouse.org/sths/knowledge-base/submitting-a-test-file/)

## A validation message names a report, row or date

| Message or condition | What to check |
| --- | --- |
| Missing or ambiguous headers | Use the flat 1.23 student-list export. Put the column headings within the first 25 lines or worksheet rows. Keep one worksheet for an Excel report. For a supplement CSV, keep the headings on row one. |
| Missing SSID | Open the named file and row. Obtain the state ID from district records and correct the export. |
| Malformed SSID | Check the indicated cell. Store all ten digits as text, including leading zeroes. |
| Repeated SSID in CALPADS | Open both indicated rows. Check whether a combined report and a campus report contain the same graduate. Correct the input coverage and retry. |
| Exit dates in different cycles | Open the reports at the listed rows and check their certified years. Put each cycle in its own input folder. For a scripted run, also check any explicit `--year`. |
| Unknown completion status | Confirm which credential was awarded. Review the code, its NSC diploma type or exclusion, and the reporting cycle at the confirmation prompt. This choice applies across the cohort. [Completion settings](command-line.md#district-completion-settings) |
| One student’s name cannot be formatted | Check the reported row against the legal student record. Use a verified spelling or correct the export encoding, then replace the source export and retry. |
| Warning about shortened or formatted student names | Check the affected names locally before submission. The formatter removes unsupported punctuation and enforces a 40-character field width in output. |
| An expected diploma date did not appear | Verify that the SSID matches, the date has a four-digit year, and there is one distinct usable date. For an out-of-cycle date or conflicting dates, check the indicated rows locally and review the saved decision. If you chose blocks only, the run uses CALPADS dates for everyone. [Compare dates](command-line.md#compare-exit-dates-with-diploma-dates) |
| Some district supplement rows were skipped | Check that the supplement uses the ten-digit state ID. District-wide exports can include rows without usable IDs. A known block on an unmatchable row stops preparation and identifies the source row to correct. |
| The full file asks for confirmation about FERPA blocks | Select the supplement containing the reviewed block column. If the district has verified that no one in the cohort has a known block, follow the confirmation prompt. [Prepare a supplement](district-supplement.md) |
| Demographic fields are blank | Choose **Include available demographics** when creating the file. The terminal lists recognized headings and reports unsupported entries. Check the [source fields and NSC codes](file-formats.md#optional-demographics-and-programs). |
| An expected block did not appear | Check the ten-digit SSID, the `FERPABlock` column name and the `Y` value. Choose the supplement at the generation prompt. Generate a fresh file after any correction. |

## Where did the file go?

A successful run prints **Graduate test file saved** or **Graduates file saved** and displays the complete path. With the recommended `input` folder, both text files and the school review workbook are in its sibling `outputs` folder. For a shared folder with a custom name, look in `outputs` inside that folder. Selecting a separate reports-only folder places `outputs` beside that folder. The timestamp helps identify your run.

If the guide says **No graduates file was saved**, the validation error above that message explains what to fix. Correct the input and choose **R** to retry. If a workbook is locked, save and close it in Excel. For command-line runs with an explicit filename, choose a new output filename or an existing directory to allow timestamped names.

Open saved text files for inspection only. If Excel converts a code to a number, close without saving and import the original as tab-delimited text with text columns. Use the generated `.txt` for the Clearinghouse upload. You can reopen a saved full file’s review and submission instructions with menu option **5**.

## A save was interrupted or another session appears to be editing

- **An output is zero bytes.** A save may have been interrupted after reserving the filename. Run generation again with a new output name, or use the guide’s timestamped name. Review the newly generated file before uploading. Keep earlier completed submissions separate.
- **Editing is still blocked after everyone has closed Excel and the guide.** An interrupted save can leave an empty hidden folder beside the affected file, such as `.School review.xlsx.editing` or `.School review.decisions.json.editing`. If you have confirmed there is no active editor, remove only that empty marker folder and retry. Preserve the workbook and decisions file.

## How do I clear the saved input locations?

The guide remembers local input paths. It determines the reporting year from the selected CALPADS reports. You can choose another working folder for a new cycle. To forget saved preferences, close the guide and remove only the `locations.json` file for your user account.

| System | Preferences file |
| --- | --- |
| macOS | `~/Library/Application Support/calpads-to-nsc/locations.json` |
| Windows | `%APPDATA%\calpads-to-nsc\locations.json` |
| Linux | `$XDG_CONFIG_HOME/calpads-to-nsc/locations.json`, normally `~/.config/calpads-to-nsc/locations.json` |

## I need help with submission

Use the StudentTracker web application for the small **Graduate test**. Upload the full file using your district’s **secure FTP** account. An FTP receipt confirms the transfer. NSC then validates the file and contacts your institution about corrections. [NSC test upload instructions](https://help.studentclearinghouse.org/sths/knowledge-base/submitting-a-test-file/) · [NSC submission and validation FAQ](https://help.studentclearinghouse.org/sths/knowledge-base/frequently-asked-questions/)
