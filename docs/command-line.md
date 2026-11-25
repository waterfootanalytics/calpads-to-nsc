# Command-line reference

[Back to the guided workflow](../README.md)

Running `prepare-nsc-graduates` with no arguments starts the interactive guide. After installing the project, you can run the same program as `python -m calpads_to_nsc`. The examples below work in an activated virtual environment on macOS, Windows PowerShell and Linux. Each command is a single line. Run `prepare-nsc-graduates --version` to see the installed version or `prepare-nsc-graduates --help` to list the available options. To diagnose a failed command, rerun it with `--debug`. The additional diagnostic gives the error type and the program function where it occurred.

## Set up a working folder

Start in the working folder containing `input`. These examples assume the reference is `input/Account Reference Sheet.xlsx`, the optional file is `input/District supplement.csv`, and your CALPADS exports are beside them. Replace the quoted names with your actual file names. The reports determine the reporting year. Add `--year 2026` to assert that every report belongs to the cycle ending August 15, 2026.

Create `outputs` once if it does not already exist.

```sh
mkdir outputs
```

With this layout, pass `input --input-folder` to identify the reports in the mixed folder. If your reports are all in a separate reports-only folder, pass its path in place of `input` and omit `--input-folder`. The positional argument also accepts one report file. Keep generated workbooks and files outside the input folder.

## Check inputs and create the school review workbook

```sh
python -m calpads_to_nsc input --input-folder --check-headers
python -m calpads_to_nsc input --input-folder --reference-sheet "input/Account Reference Sheet.xlsx" --create-crosswalk "outputs/School review.xlsx" --no-prompt
```

The header check looks for the required CALPADS columns in the first 25 lines or worksheet rows of each report. Open `outputs/School review.xlsx`, review the suggestions and complete the choices under **Match schools**. Save the workbook. You can also do the review in the terminal by replacing `--no-prompt` in the next command with `--terminal-review`.

```sh
python -m calpads_to_nsc input --input-folder --reference-sheet "input/Account Reference Sheet.xlsx" --crosswalk "outputs/School review.xlsx" --check-crosswalk --no-prompt
```

On **Match schools**, an intentional exclusion needs a school-level reason in F and the word `EXCLUDE` in G. Confirmed exclusions appear by campus name at validation and file generation. An unresolved or flagged school stops validation and the message names the row to review. After correcting it, rerun the check. Use a fresh workbook if the reference sheet changes.

## Compare exit dates with diploma dates

Save a grouped preview to a **new** CSV to see all date differences and pending completion decisions.

```sh
python -m calpads_to_nsc input --input-folder --reference-sheet "input/Account Reference Sheet.xlsx" --crosswalk "outputs/School review.xlsx" --supplement "input/District supplement.csv" --compare-dates "outputs/date-comparison.csv" --no-prompt
```

| Result column | How to read it |
| --- | --- |
| `Eligibility` | Whether the completion code qualifies or still needs a district decision. |
| `Status` | How the date was selected or which date question needs review. |
| `CALPADSExitDate`, `SupplementGraduationDate` | The reported exit date and a single plausible supplemental date, when available, including a date awaiting review. |
| `DaysDifference` | Exit date minus the supplemental date shown. A positive number means the exit was later. Blank means there is no single date to compare. |
| `DateSource`, `NSCGraduationDate` | The source and date selected for an eligible graduate. Both are blank while eligibility or the date is pending. |
| `Students` | The number of CALPADS students in that group. |

You can run a date comparison before reviewing campus names. Omit both `--reference-sheet` and `--crosswalk` to compare the full roster. To compare a subset using a rehearsal workbook and `--defer-school`, include both.

Filter the CSV on `Eligibility` and `Status` to find groups that need review. The preview uses saved decisions, skips confirmed exclusions and records pending groups without prompting. Use the same school workbook and supplement in the guide to resolve the decisions, then rerun the comparison with a new filename. Test and full-file generation require each decision to be resolved.

A supplement covering more students than the CALPADS roster contributes only fields matched to eligible CALPADS graduates. To retain all exit dates while still using the supplement for blocks, add `--calpads-exit-dates`. In this mode the supplement supplies FERPA blocks only. A date comparison requires a `GraduationDate` heading.

## Validate and generate

The dry run applies the school choices, diploma rules, supplement and output field checks. Address errors before generating the full cohort.

```sh
python -m calpads_to_nsc input --input-folder --reference-sheet "input/Account Reference Sheet.xlsx" --crosswalk "outputs/School review.xlsx" --supplement "input/District supplement.csv" --dry-run --no-prompt
python -m calpads_to_nsc input --input-folder --reference-sheet "input/Account Reference Sheet.xlsx" --crosswalk "outputs/School review.xlsx" --supplement "input/District supplement.csv" --test-output outputs --no-prompt
```

Review the test file and check it through your district’s StudentTracker web account. Once ready to prepare the full file, run the following command. Submission is a separate manual step using the district’s secure FTP account. [Review and submit](../README.md#submit-your-files)

```sh
python -m calpads_to_nsc input --input-folder --reference-sheet "input/Account Reference Sheet.xlsx" --crosswalk "outputs/School review.xlsx" --supplement "input/District supplement.csv" --output outputs --no-prompt
```

- **Without a supplement**, omit `--supplement` and its filename from the dry run or generation command. CALPADS exit dates supply column L. For full-file generation with no supplement or a dates-only supplement, supply `--confirm-no-ferpa-blocks` only after the district has confirmed that the cohort has no known blocks. The guide asks for this confirmation interactively. A known block belongs in a supplement with a `FERPABlock` column.
- **To choose the optional identifier in NSC column J**, use `--student-id local` (the default, local ID if present), `--student-id ssid` (state ID) or `--student-id blank` (empty for everyone). Use the same choice for the dry run, Graduate test and full file. The program always joins the supplement by SSID.
- **For timestamped names**, pass the path to an existing output directory. To choose a specific name, pass a new path such as `outputs/Graduates_2026.txt`. Use letters, digits, hyphens or underscores in the filename. An existing file is preserved.
- **For an interactive review after saving**, use `--review-test` with `--test-output` or `--review-full` with `--output`, and omit `--no-prompt`.
- **For an explicit transmission date**, add `--date YYYYMMDD`. The default is today’s local date. Test and full files can use today or an earlier transmission date. For planning, a dry run or grouped date comparison can use a later date. Birth dates must precede graduation, and a selected graduation date must be on or before the transmission date.

### Include demographics and program fields

Add `--include-demographics` to the dry run, test and full commands to include available CALPADS demographics. Add `--use-special-education-for-disability` alongside it if the district uses special education participation for column Z. To check all five fields in a Graduate test file, for example, run the following command.

```sh
python -m calpads_to_nsc input --input-folder --reference-sheet "input/Account Reference Sheet.xlsx" --crosswalk "outputs/School review.xlsx" --supplement "input/District supplement.csv" --include-demographics --use-special-education-for-disability --test-output outputs --no-prompt
```

The supplement can contribute assessments, math semesters, dual enrollment and a district program code with either demographic setting. Omit the unused columns or leave their cells empty. [Supplement headings and accepted values](district-supplement.md)

### Review decisions from a terminal

The interactive guide handles both completion codes and diploma-date conflicts. Menu option **6 — Review graduation decisions** also lets you revisit a saved choice. For the same review from a command prompt, run this from your working folder.

```sh
python -m calpads_to_nsc input --input-folder --reference-sheet "input/Account Reference Sheet.xlsx" --crosswalk "outputs/School review.xlsx" --supplement "input/District supplement.csv" --dry-run --review-decisions
```

Check the identified rows in district records. A question about one graduate shows source filenames and row numbers. The software keeps the actual dates and SSID off the terminal. For multiple conflicting dates, correct the supplement or confirm the CALPADS fallback. For one plausible out-of-cycle diploma date, you may also affirm that the supplied diploma date is correct. For any completion-code choice, review the exact code, NSC diploma type or exclusion, and the reporting cycle. Type `APPLY` to map the code or `EXCLUDE` to omit it. Press Return to go back and choose again. Subsequent runs list the completion codes and mappings applied. **Q** leaves saved decisions available for your next run.

The guide saves graduation decisions beside the workbook in a private file such as `School review.decisions.json`. A completion decision applies to one agreement and reporting cycle. A date decision applies to the reviewed report and supplement rows. Keep the workbook and JSON file together when moving the working folder. Changing a source filename or the evidence in a reviewed row can require a new decision.

Use the same workbook for the date comparison, dry run, Graduate test and full submission. The grouped comparison shows pending decisions. Unattended file generation requires a decision for each question.

## District completion settings

If a report contains a completion code outside the built-in mapping, a district decision is needed to include that credential under an accepted diploma type or exclude it. Document the decision locally. For scripted district-wide credential policies, an optional UTF-8 JSON file accepts the `completion_types` object and takes precedence over a saved choice for the same code. A three-digit code maps to an exact NSC diploma type or to JSON `null` to exclude it. [Built-in completion rules](file-formats.md#completion-status) · [NSC diploma types](https://help.studentclearinghouse.org/sths/knowledge-base/frequently-asked-questions/)

The following file illustrates a district decision to **exclude code 250**. Use it only when it reflects your district’s decision.

```json
{
  "completion_types": {
    "250": null
  }
}
```

Save the file as `completion-policy.json` in the private working folder. Add `--settings completion-policy.json` to every applicable date comparison, dry run, test and full generation command for that cycle. Other codes retain their built-in behavior. If no additional codes need a decision, you can skip the settings file.

## Rehearse with an incomplete school list

To test a partial set of school matches, add `--rehearsal` when creating a **new** review workbook. If a school cannot be matched, leave it blank or flagged in the workbook. Supply `--defer-school "Exact CALPADS campus name"` when checking that rehearsal, comparing dates, running a dry run, or generating a Graduate test file. Repeat the option for additional campuses. Keep at least one matched school. Full submission requires a submission workbook with a match or a confirmed exclusion for every CALPADS school.

Use the built-in help to see all options in your installed version.

```sh
python -m calpads_to_nsc --help
```

## Check a saved graduates file

Run this check with the original test or full text file. It requires no CALPADS exports, school workbook or supplement.

```sh
prepare-nsc-graduates --check-file "outputs/your-graduates-file.txt"
```

The check reads the header, all 28 fields in each detail row, and the trailer. It checks field positions, lengths, dates, allowed values, and the total record count. A structural or field error returns a nonzero exit status and identifies the row or field. Warnings preserve the spelling of school and student names.
