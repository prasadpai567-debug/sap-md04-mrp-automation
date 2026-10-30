# SAP MD04 MRP Date Automation

Excel VBA and SAP GUI Scripting portfolio project for batch extraction of selected MRP dates into an Excel tracking sheet.

> Portfolio project for SAP procurement, materials planning, and business-process automation roles.

## Repository status

The current public repository is documentation-only. It contains:

- README.md
- APPLICATION_ANALYSIS.md

The original macro-enabled workbook, exported VBA modules, screenshots, and operational data are intentionally not published here. Cloning the repository alone does not provide a runnable application.

The behavior described below is based on the external workbook review documented in APPLICATION_ANALYSIS.md. The implementation artifact is not currently included for independent source-code review.

## Business problem

A materials-planning user may need to open SAP transactions repeatedly for many materials, enter the same material and MRP-area values, locate a relevant MRP element, read a date, and transcribe the result into Excel.

This is repetitive, time-sensitive work that can be affected by manual entry and inconsistent navigation.

## Automation objective

The documented automation uses Excel as an input/output control sheet and SAP GUI Scripting to process a batch of material records. For each nonblank material row, it is designed to:

1. Connect to an existing authenticated SAP GUI session.
2. Run SAP transaction MD03 with the material and MRP-area values.
3. Open SAP transaction MD04 for the material.
4. Apply a recorded SAP list filter.
5. Search the visible MRP table for an element containing PLA.OR.
6. Capture the first associated nonblank date.
7. Write the result to the Excel Release Date column.
8. Write Not Found when no qualifying date is identified.

This project does not create purchase orders, send supplier emails, or communicate with external parties.

## Documented workflow

| Stage | Input | Processing | Output |
| --- | --- | --- | --- |
| Excel preparation | Material Number and MRP AREA | Read rows beginning at row 2 | Batch input |
| SAP setup | Existing SAP GUI session | Obtain the scripting engine and select a session | Automation context |
| MD03 | Material and MRP area | Enter fields and submit the transaction | SAP planning screen |
| MD04 | Material | Open the stock/requirements list and apply the recorded filter | Filtered MRP table |
| MRP search | Visible table rows | Find PLA.OR and the first nonblank associated date | Candidate result |
| Excel update | Candidate result | Write to column C | Release Date or Not Found |
| Cleanup | Completed batch | Return SAP to /N and release objects | Completion message |

A more detailed workflow is available in docs/application-workflow.md.

## Documented implementation scope

The external workbook analysis reports:

- One substantive VBA procedure named SAP_Final_Solution.
- One worksheet named Sheet1.
- Input columns for Material Number and MRP AREA.
- Release Date output in column C.
- A clickable MD04 worksheet shape.
- Late-bound SAP GUI objects.
- A maximum scan of 20 MRP table rows per material.
- No UserForm, report sheet, structured error log, or retry queue.

These points are documented observations, not claims that source code is currently present in this repository.

## Requirements for a future sanitized demo

A functional demonstration would require:

- Windows.
- Microsoft Excel Desktop with macro support.
- SAP GUI for Windows.
- SAP GUI Scripting enabled and permitted by the SAP system.
- An authenticated SAP session with appropriate authorization.
- A sanitized workbook and synthetic input values.

The automation is not designed for Excel for the web, Excel for Mac, or SAP environments with incompatible screen layouts and control IDs.

## Safe setup and usage

When a sanitized workbook is approved for publication:

1. Use synthetic material numbers and non-production MRP-area values.
2. Open SAP GUI and log in only to an approved non-production or test environment.
3. Open the sanitized workbook in Excel Desktop.
4. Keep the intended input sheet active.
5. Populate Material Number and MRP AREA beginning in row 2.
6. Review the rows before running.
7. Run SAP_Final_Solution only after confirming the SAP session and target environment.
8. Review Release Date results and investigate Not Found values manually.

Do not upload or use a production workbook as a portfolio artifact. Do not disable Office security globally. See docs/setup-and-troubleshooting.md.

## Error handling and operational safeguards

The documented implementation has important limitations:

- It may select the first available SAP connection and session.
- SAP control IDs are hard-coded and may vary by system, version, or screen variant.
- A fixed one-second wait may be too short for a slow SAP response.
- On Error Resume Next can hide some UI errors.
- A target row after the first 20 inspected rows may not be found.
- Not Found can represent several different conditions, including missing data, a layout mismatch, or a recoverable UI error.
- The completion count may reflect row positions rather than successful nonblank records.
- MD03 may affect planning results, so production use requires authorization, testing, and operational controls.

## Security and data handling

The public repository must contain documentation, synthetic values, and redacted screenshots only.

Never publish:

- Real material numbers, MRP areas, plants, dates, or comments.
- SAP system, client, user, or connection identifiers.
- Credentials, tokens, connection strings, or passwords.
- Supplier, customer, employee, or procurement records.
- Original workbook metadata containing organization-specific paths.
- Proprietary VBA project binaries without authorization.

The current published Markdown files do not contain a runnable workbook or known credentials. This does not replace an organizational data-classification review.

## Testing status

No automated tests or runnable implementation are currently included. Excel, VBA, SAP GUI, and external-system execution has not been validated from this repository.

Future tests should cover:

- Input validation.
- MRP-element/date matching.
- Result mapping to Excel rows.
- Missing controls and timeout states.
- Mocked SAP responses without connecting to production.

## Benefits supported by the documented workflow

The automation is intended to reduce repetitive SAP navigation and manual date transcription, process multiple material records consistently, and keep results in a familiar Excel format.

No measured time savings, throughput figures, or error-rate improvements are claimed because the repository contains no measurement evidence.

## Known limitations

This is a Windows desktop automation concept that depends on SAP GUI screen structure, user authorization, SAP GUI Scripting settings, and workbook state. It is not a standalone SAP API integration and is not production-ready based on the current published repository.

## Planned improvements

- Publish redacted VBA source only after confidentiality review.
- Add a sanitized workbook or input template.
- Select SAP sessions explicitly rather than by position.
- Replace broad error suppression with scoped handling and per-row status logging.
- Add configurable timeouts, scan limits, and SAP parameters.
- Add validation, retry, cancellation, and resumable processing.
- Add mocked tests and redacted screenshots.
- Add release notes for validated portfolio versions.

## Recruiter summary

Built an Excel VBA and SAP GUI Scripting automation concept for procurement and materials-planning operations. The workflow processes material/MRP-area inputs, navigates SAP MD03 and MD04, identifies relevant PLA.OR planning elements, extracts the first available date, and writes standardized results back to Excel. The project demonstrates SAP process knowledge, spreadsheet automation, batch processing, operational risk awareness, and responsible handling of business data.

## License

This repository uses a portfolio-use, all-rights-reserved notice. The materials are not an authorization to copy or deploy the automation in a production environment.
