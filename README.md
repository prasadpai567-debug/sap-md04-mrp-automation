# SAP MD04 MRP Run Automation

Excel VBA automation for retrieving selected MRP dates from SAP and writing the results back to an Excel input/output table.

> Portfolio project demonstrating Excel VBA, SAP GUI Scripting, batch processing, and procurement/planning process automation.

## Project overview

The application uses Excel as a lightweight control panel for a repetitive SAP planning task. A user supplies material numbers and MRP areas in Excel. The macro then:

1. Connects to an existing SAP GUI session.
2. Runs transaction MD03 for each material and MRP area.
3. Opens transaction MD04 for the material.
4. Applies the recorded SAP list filter.
5. Searches the visible MRP table for an element containing PLA.OR.
6. Captures the first nonblank date associated with that element.
7. Writes the result to the Release Date column in Excel.

If a qualifying date is not found, the macro writes Not Found.

The tool addresses the manual effort involved in opening SAP transactions, entering the same material data repeatedly, locating the relevant MRP line, and transcribing dates into a tracking sheet.

## Key features

- Excel-driven batch input: reads material numbers from column A and MRP areas from column B.
- SAP GUI integration: uses the running SAP GUI Scripting engine through late-bound COM objects.
- Automated MD03 execution: enters the material, MRP area, and plant/MRP-area value before submitting the transaction.
- Automated MD04 navigation: opens the material stock/requirements list and applies the configured list filter.
- Targeted MRP search: checks up to 20 rows for an MRP element containing PLA.OR.
- First valid date capture: ignores blank dates and stops at the first nonblank date associated with the target element.
- Excel result writing: writes the extracted date or Not Found to column C.
- Progress completion message: reports the number of processed rows after the loop completes.
- Workbook control: includes a rounded MD04 shape assigned to the SAP_Final_Solution macro.
- Filterable worksheet: the input/output range has standard Excel filtering enabled.
- Duplicate-value highlighting: the workbook contains duplicate-value conditional formatting rules, mainly around material-number ranges.

## Application workflow

### 1. Prepare the Excel sheet

On Sheet1, use the following columns:

| Column | Header | Purpose |
| --- | --- | --- |
| A | Material Number | SAP material to process |
| B | MRP AREA | MRP area and plant value passed to SAP |
| C | Release Date | Output written by the macro |
| D | Comments | Optional user-maintained status or notes; the macro does not populate it |

The macro starts at row 2 and determines the final loop row from the last nonblank cell in column A. Blank material cells are skipped.

### 2. Connect to SAP

The macro looks for SAPGUI, obtains the scripting engine, and selects the first available connection and first session. SAP Logon must already be open, and the user must already be logged in to a session that is authorized to run the required transactions.

### 3. Run MD03

For each material, the macro navigates to /NMD03, enters the material number, MRP area, and plant/MRP-area value, and submits the screen.

### 4. Run MD04 and apply the filter

The macro navigates to /NMD04, enters the material number, submits the screen, and applies the recorded filter through SAP GUI control IDs. The filter key used in the code is SAP00001.

### 5. Inspect the MRP table

The macro reads the SAP table tblSAPMM61RTC_EZ. It uses the visible-row count and vertical scrollbar position to inspect at most 20 table rows. It reads the MRP element from DELB0, checks whether the text contains PLA.OR, then reads the associated date from DAT02.

### 6. Write the result

The first nonblank matching date is written to column C. If no matching nonblank date is found, the initial value Not Found remains in the cell.

### 7. Clean up

After processing, the macro returns SAP to /N, releases its object variables, and displays an Automation Complete message.

## Technical details

### Languages and technologies

- Visual Basic for Applications (VBA)
- Excel Open XML workbook format with macros (.xlsm)
- SAP GUI Scripting API
- COM automation through late-bound Object variables
- Excel worksheet controls, shapes, AutoFilter, and conditional formatting

### VBA project components

| Component | Current implementation |
| --- | --- |
| Module1 | Contains the SAP_Final_Solution procedure and all SAP automation logic |
| Module2 | Present but empty |
| ThisWorkbook | Contains generated component attributes only; no workbook event procedure was found |
| Sheet1 | Contains generated component attributes only; no worksheet event procedure was found |
| UserForms | None found |

The workbook contains standard stdole and Office object-library references. The SAP objects are declared as Object, so a separate compile-time SAP reference is not required by the code.

### Workbook structure observed

- One visible worksheet: Sheet1
- Four working columns: Material Number, MRP AREA, Release Date, and Comments
- No worksheet formulas were found
- No Excel tables, charts, or embedded images were found
- One drawing shape is linked to SAP_Final_Solution
- The saved workbook has an active AutoFilter and hidden rows reflecting a prior filtered view

## Business benefits

- Reduces repetitive SAP navigation and data entry.
- Processes multiple material records from one Excel control sheet.
- Applies the same MRP search method consistently for each material.
- Reduces manual transcription of dates from SAP to Excel.
- Provides a standard Not Found result when the expected qualifying date is unavailable.
- Keeps the output in a familiar Excel format that can be filtered and reviewed.

These are workflow benefits inferred from the implemented logic. The workbook does not contain measured time-savings, error-rate, or throughput metrics.

## Requirements

- Windows desktop environment
- Microsoft Excel Desktop with VBA and macro-enabled workbook support
- SAP GUI for Windows
- SAP GUI Scripting enabled on the client and permitted by the SAP system
- An active, authenticated SAP connection and session
- User authorization for the SAP transactions and fields used by the script

This application is not designed for Excel for the web, Excel for Mac, or an SAP session that exposes a different screen layout or control IDs.

## Installation and setup

1. Clone or download the repository.
2. Use a trusted local copy of the macro-enabled workbook.
3. If the workbook is distributed as a portfolio sample, use sanitized demo data and remove organization-specific metadata before opening it.
4. Install or open SAP GUI for Windows and sign in to the intended SAP environment.
5. Confirm that SAP GUI Scripting is available and that the required transactions can be run manually.
6. Open the .xlsm file in Excel Desktop.
7. Keep Sheet1 active and populate columns A and B starting at row 2.
8. Review the material list before running. The macro processes rows through the last nonblank cell in column A, including rows that are hidden by an Excel filter.
9. Run the SAP_Final_Solution macro from the MD04 worksheet button or from Developer → Macros.
10. Review the Release Date column and investigate any Not Found results in SAP.

Do not disable Office security globally. Use a trusted, sanitized copy and enable macros only when the workbook and SAP environment are known to be safe.

## Usage example

Example input layout:

    Material Number | MRP AREA | Release Date | Comments
    DEMO-001        | 1800     |              |
    DEMO-002        | 1800     |              |

After the run, column C is populated with the first matching nonblank date returned from SAP or with Not Found.

## Current limitations and operational safeguards

- The code uses the first SAP connection and first SAP session (Children(0)). Close unrelated SAP sessions or confirm the session order before running.
- The code uses ThisWorkbook.ActiveSheet, so the intended input sheet must be active.
- The SAP screen paths and control IDs are hard-coded. SAP screen variants, upgrades, or customizations may require code updates.
- Only the first 20 rows of the SAP MRP table are inspected for each material.
- Not Found can represent a missing target element, a blank target date, an unexpected SAP layout, or a recoverable lookup error.
- Several SAP reads are wrapped in On Error Resume Next, so some UI errors may be silently bypassed.
- The macro contains a one-second wait before reading the MRP table, which may be insufficient for a slow SAP connection.
- The completion count is based on lastRow - 1, not the number of nonblank material rows actually processed.
- The macro does not create a detailed execution log, retry failed materials, validate input fields, or write an error summary.
- Running MD03 may change planning results in SAP. Test in a non-production environment before operational use.

## Suggested repository structure

    sap-md04-mrp-automation/
    ├── README.md
    ├── LICENSE
    ├── docs/
    │   ├── application-workflow.md
    │   ├── setup-and-troubleshooting.md
    │   └── screenshots/
    ├── src/
    │   └── vba/
    │       ├── Module1.bas
    │       ├── Module2.bas
    │       ├── Sheet1.cls
    │       └── ThisWorkbook.cls
    ├── sample-data/
    │   └── input-template.xlsx
    └── workbook/
        └── sanitized-demo.xlsm

### Recommended GitHub contents

- README.md and supporting documentation
- Exported VBA modules after removing sensitive comments, identifiers, and environment-specific values
- A blank or synthetic input template
- Redacted screenshots
- A sanitized demo workbook only if it cannot connect to a real company system
- License and contribution files if the project will be shared publicly

### Do not upload

- The original production workbook with real material numbers, MRP areas, dates, or comments
- A workbook containing an absolute company SharePoint/OneDrive path
- SAP logon details, connection strings, credentials, tokens, or screenshots showing system/client/user information
- The embedded VBA project binary if it contains proprietary code or project-protection metadata
- Real supplier, customer, employee, or procurement records

## Screenshot recommendations

Use synthetic values and redact all SAP and company identifiers.

1. Excel input screen: show the four-column layout and the MD04 button.
2. SAP MD03 screen: show the material, MRP area, and plant fields with demo values.
3. SAP MD04 list: show the target PLA.OR line and its date, with system and material identifiers redacted.
4. Completed Excel output: show populated dates and a Not Found example.
5. Completion message: show the macro completion dialog without real counts or business identifiers.

## Implemented features versus future improvements

### Implemented

- Excel-to-SAP batch processing
- MD03 execution per input row
- MD04 navigation and recorded list filtering
- Controlled scan of MRP table rows
- PLA.OR matching and first nonblank date extraction
- Excel output writing and completion notification
- Basic defensive handling for missing SAP objects and optional SAP controls

### Future improvements

- Add a configuration sheet for SAP transaction parameters, filter keys, scan limits, and wait times.
- Replace On Error Resume Next with explicit error handling and a per-material error log.
- Validate material numbers and MRP areas before starting SAP automation.
- Support reliable pagination or full-table scanning beyond 20 rows.
- Detect and select a named SAP connection/session instead of using Children(0).
- Add a dry-run mode and a test mode using a mock SAP response.
- Separate the SAP automation layer from the Excel input/output layer.
- Add a progress indicator, cancellation control, and resumable processing.
- Preserve true Excel date values and standardize date parsing.
- Add automated tests for parsing, output mapping, and failure states.

## Resume and portfolio description

Built an Excel VBA and SAP GUI Scripting automation tool that batch-processes material/MRP-area inputs, runs SAP MD03 and MD04, searches MRP rows for PLA.OR, extracts the first valid related date, and writes standardized results back to Excel with completion and exception handling.

## License and data-use note

Add a license appropriate to the code owner and organization. This repository should contain synthetic or redacted data only. The automation is intended for demonstration and controlled internal use unless it is validated against the target SAP environment.
