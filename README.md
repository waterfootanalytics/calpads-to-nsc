# CALPADS to NSC

Create a National Student Clearinghouse (NSC) StudentTracker for High Schools Graduates File from certified CALPADS Report 1.23 exports. The tool runs on your computer. It creates a workbook for matching school names, a Graduate test file for NSC to check, and a full file for submission.

## Download and install

1. Install Python 3.11 or newer if you need it. [Download Python](https://www.python.org/downloads/)
2. Download **Source code (zip)** for the version you want from [Releases](https://github.com/waterfootanalytics/calpads-to-nsc/releases). Extract it and open the folder containing this README and `pyproject.toml`.
3. Open a terminal in that folder and run the commands for your computer. On macOS, open Terminal, type `cd` followed by a space, drag the folder into the window, and press Return. On Windows, right-click inside the folder in File Explorer and choose **Open in Terminal**.

**macOS or Linux**

```sh
python3 -m venv .venv
source .venv/bin/activate
python -m pip install .
```

**Windows PowerShell**

```powershell
py -3 -m venv .venv
.\.venv\Scripts\Activate.ps1
python -m pip install .
```

## Gather your files

Create a folder for the graduating class. Inside it, create a folder named `input` and put these files directly inside it. You can keep their original names.

| File | What to use |
| --- | --- |
| **CALPADS reports** | The certified **1.23 Graduates and Completers Student List** for the reporting year. Use one flat file per school or a combined list. Accepted formats are `.xlsx`, `.csv`, `.tsv` and `.txt`. |
| **NSC Account Reference Sheet** | Download the sheet for your agreement from **Accounts and Settings** in StudentTracker. Use the original Excel file (`.xls` or `.xlsx`) or an NSC PDF with selectable text. PDF input uses Poppler’s `pdftotext` command. [PDF setup](docs/troubleshooting.md#the-pdf-account-sheet-will-not-open) |
| **District supplement** (optional) | A CSV keyed by `SSID`. Include any district diploma dates, FERPA blocks, assessment results, math semesters, dual enrollment or program codes you have. For diploma dates and FERPA blocks, use the headings `SSID,GraduationDate,FERPABlock`. [Prepare a supplement](docs/district-supplement.md) |

## Run the guide

Start the guide in the terminal window you used for installation.

```sh
prepare-nsc-graduates
```

The screens below are excerpts from a first run. **Return** and **Enter** mean the same key. When a prompt shows a default, press Enter to accept it. Type **Q** and press Enter to exit.

### 1. Select the inputs and create the workbook

If the guide opens a folder picker, click your class’s `input` folder. If you started in the class folder, the guide finds `input` automatically. The **Input check** shows which files it found. Check the reporting year it derives from the CALPADS reports, then create your workbook.

```text
School review workbook
  1  Create a new workbook
  2  Use an existing workbook
  Q  Exit guide
Press Return for 1, or enter your choice >
```

Press **Enter** here and again at **Where should the new workbook be saved?** to create `outputs/School review.xlsx`. The guide opens the school review in Terminal.

### 2. Match each school to its Clearinghouse name

When a school needs a match, the guide shows its CALPADS name and lists the available Clearinghouse names with their six-digit ACT codes. At **School number or six-digit ACT code >**, enter the number beside the correct campus or its ACT code. If you’re unsure, **F** flags the campus and pauses the review.

To work in Excel, enter **X**. On **Match schools**, choose the registered school in column D. Use **Look up a school** to resolve a flagged match. To leave a campus out, choose **Exclude from this submission**, enter a school-level reason in F and type `EXCLUDE` in G. Save and close the workbook, return to Terminal, and press **Enter** to check your changes.

```text
Every school has a match.
Return  Keep saved matches
M       View or change a match or exclusion
X       Open Excel
Q       Exit guide
>
```

Press **Enter** to keep the matches and continue. On a later run, choose **1** from the main menu to revisit them.

### 3. Make and review the Graduate test file

At the main menu, type **2** and press **Enter**, as shown here. Your reporting year and folder appear above the choices.

```text
CALPADS completion reporting year · 2025–26 (Aug 16, 2025–Aug 15, 2026)

  1  Review school matches
  2  Make a Graduate test file
  3  Make the full graduates file
  4  Change input files
  5  Review and submit a saved full file
  6  Review graduation decisions
  7  Check a saved NSC graduates file
  Q  Exit guide
Press Return for 1, or enter your choice > 2
```

At the supplement prompt, **Enter** uses the CSV shown. If no CSV was found, **Enter** continues with CALPADS exit dates, or **Y** opens a file picker. If a diploma date or completion code needs a decision, the guide identifies the source rows and explains the choices.

At the demographics prompt, press **Return** to include the available fields. Choose **2** to leave them blank for this file.

```text
Include demographics from CALPADS?
  1  Include available demographics
  2  Exclude demographics for this file
```

The guide lists the fields it found. If your reports include Special Education Status, press **Return** at the next prompt to use it for NSC’s disability indicator (column Z). Choose **2** to leave Z blank for this file. Empty or unrecognized entries stay blank. District academic and program fields come from the supplement. [See the column map](docs/file-formats.md#optional-demographics-and-programs)

Before saving, review the excluded campuses, completion codes, diploma-date sources and any field warnings. If the supplement has a FERPA column but no usable Y or N entries, the summary flags that all graduates will receive N (unknown). Press **Return** to save or **M** to go back.

Look for **Graduate test file saved** and the path to your timestamped `TEST_Graduates_…txt` file. It contains one eligible graduate per included school. On macOS and Windows, you’ll see these review choices.

```text
Return  Open this file in Excel
C       I reviewed the file. Show me how to submit it.
T       Show the number of graduates in this file
M       Main menu
Q       Exit guide
>
```

Press **Enter** to open the file (LibreOffice Calc on Linux). If an import window opens, choose **Tab** as the delimiter and **Text** for the columns to preserve leading zeros. Check the student ID in **J**, diploma date in **L**, FERPA block in **M**, school name in **N** and ACT code in **O**. If you included demographics, check **P** (gender), **Q** (race/ethnicity), **R** (economic disadvantage), **W** (English learner) and, if selected, **Z** (disability). Close the file without saving edits. Back in Terminal, enter **C** for upload instructions.

## Submit your files

| Step | File in `outputs` | Where to upload it |
| --- | --- | --- |
| **1. Check the test file** | `TEST_Graduates_…txt` | Sign in to StudentTracker and choose **Submit a Test File → Graduate**. Review the results. If NSC finds an error, correct the input and generate the test file again. [NSC test instructions](https://help.studentclearinghouse.org/sths/knowledge-base/submitting-a-test-file/) |
| **2. Submit the full class** | `Graduates_…txt` | Once the test passes, choose **3 — Make the full graduates file** in the terminal guide. Review the file, sign in to NSC’s secure FTP site, choose **Upload → Browse**, and select the full file. [NSC secure FTP login](https://ftps.nslc.org/) |

If your inputs omit the FERPA block column, the guide asks you to confirm there are no known blocks before saving the full file.

At the upload instructions, **O** opens the website, **M** takes you back to the main menu, and **Enter** finishes and exits. If you exit after the test, run `prepare-nsc-graduates` again when you’re ready for option **3**. Both files are in the class’s `outputs` folder, next to `input`. Choose **7** to check the layout of a saved test or full text file. The guide also checks the file when you request upload instructions. To review a previously saved full file and see its upload instructions again, choose **5**.

[File formats](docs/file-formats.md) · [Command-line reference](docs/command-line.md) · [Troubleshooting](docs/troubleshooting.md)

## License

Waterfoot Analytics maintains this project independently of the National Student Clearinghouse and the California Department of Education. The MIT License permits use, modification and redistribution and includes a warranty disclaimer and limitation of liability. [MIT License](LICENSE)
