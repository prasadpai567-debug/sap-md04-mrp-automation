# Application analysis: MD04_MRP RUN.xlsm

## Scope and method

The original .xlsm file was reviewed without saving or modifying it. The review covered the OOXML package, worksheet XML, workbook metadata, drawing assignments, VBA project streams, VBA source code, references, and the visible Excel workbook state.

The original workbook remains unchanged in its source location.

## Executive findings

This is a focused Excel-to-SAP automation utility rather than a multi-sheet reporting application. Excel supplies a list of materials and MRP areas. VBA controls a running SAP GUI session, executes MD03 and MD04 for each material, searches the MD04 table for an MRP element containing PLA.OR, captures the first nonblank associated date, and writes the result into the same worksheet.

The workbook has one substantive macro, one empty standard module, no user forms, no workbook or worksheet event logic, and one worksheet shape assigned to the macro.

## Workbook structure

### Worksheet

- Sheet: Sheet1
- Used range reported by the workbook: A1:H159
- Primary data range: A1:D159
- Headers:
  - Material Number
  - MRP AREA
  - Release Date
  - Comments
- Input rows begin at row 2.
- The macro uses column A to determine the last row.
- MRP area values in the inspected file include more than one area code. These are operational data and should be replaced with synthetic values before publication.

### Saved workbook state

- An AutoFilter is applied to A1:D159.
- The saved view contains a date filter on the Release Date column and a nonblank filter on Comments.
- Many rows are hidden because of the saved filter state.
- The macro does not use the visible-row state of the Excel sheet. It loops through all worksheet rows from 2 through the last nonblank material row, including hidden rows.
- The worksheet contains no formulas, Excel tables, charts, data-validation rules, or embedded images.
- Duplicate-value conditional formatting rules are present across several material-number ranges and on one isolated cell. They are worksheet-level formatting and are not used by the VBA procedure.

### Drawing and macro assignment

The drawing XML contains one rounded rectangle with visible text MD04. Its macro attribute points to [0]!SAP_Final_Solution, so the workbook provides a clickable control for the main procedure.

## VBA project inventory

| Component | Finding |
| --- | --- |
| Module1 | Contains SAP_Final_Solution, the complete automation procedure |
| Module2 | Empty standard module |
| ThisWorkbook | Generated VBA component attributes only; no event procedure |
| Sheet1 | Generated VBA component attributes only; no worksheet event procedure |
| UserForms | None found in the VBA project streams |
| .frx form resources | None found |

The embedded project includes standard references corresponding to the OLE Automation (stdole) and Office object libraries. The SAP GUI objects are late-bound and declared as Object variables.

## Detailed macro logic

### Initialization and SAP connection

SAP_Final_Solution declares Excel and SAP GUI objects, then attempts to obtain SAPGUI with GetObject("SAPGUI"). It retrieves the scripting engine, the first connection, and the first session. If SAP GUI is not available, it shows a critical error and exits.

### Excel setup

The procedure sets wsInput to ThisWorkbook.ActiveSheet and finds the final row with:

    lastRow = wsInput.Cells(wsInput.Rows.Count, "A").End(xlUp).Row

It then maximizes the active SAP window.

### Per-material processing

For each row from 2 to lastRow, the procedure reads:

- Column A as matNR
- Column B as mrpArea

It initializes extractedDate to Not Found and skips the row if the material number is blank.

### MD03 step

The code navigates to /NMD03, presses the SAP command button, enters:

- RM61X-MATNR = material number
- RM61X-BERID = MRP area
- RM61X-WERKS = MRP area

It submits the screen twice with sendVKey 0.

### MD04 step

The code navigates to /NMD04, enters the material number in the MD04 material field, and presses the command button twice. It then attempts to apply the recorded filter:

- Toolbar button: wnd[0]/tbar[1]/btn[29]
- Filter control: RM61R-FILBZ
- Filter key: SAP00001

The filter block uses On Error Resume Next so the macro continues if the control is unavailable or the filter is already active.

### Table scanning

The code waits one second, obtains the SAP table tblSAPMM61RTC_EZ, and reads its visible row count. It sets a maximum scan limit of 20 rows or the table row count if smaller.

For each row in that limit, it:

1. Adjusts the vertical scrollbar when the next row is outside the current visible range.
2. Reads MDEZ-DELB0 as text, with a context-field fallback.
3. Checks for PLA.OR using a case-insensitive substring search.
4. Reads the date from MDEZ-DAT02, with a text-field fallback.
5. Accepts the first nonblank date and exits the scan.

### Excel output and cleanup

The extracted text is written to wsInput.Cells(i, 3), which is column C. After all rows are processed, SAP is sent to /N, object variables are released, and a completion message reports lastRow - 1 as the processed count.

## Implemented behavior that should be described accurately

- The macro runs SAP GUI transactions. It is not a formula-only Excel tool.
- It reads existing SAP screen text; it does not export a full MD04 table.
- It searches at most 20 SAP table rows per material.
- It captures the first nonblank date associated with an MRP element containing PLA.OR.
- It writes results to the same worksheet and does not create a report sheet.
- It does not create purchase orders, communicate with suppliers, or send email.
- It does not contain a UserForm.
- It does not contain a structured error log or a retry queue.

## Risks and limitations

1. Session selection: Children(0) may select the wrong connection or session when several SAP sessions are open.
2. Hard-coded screen IDs: SAP GUI control IDs can change with versions, screen variants, customizations, or different transaction states.
3. Partial scan: a target row after the first 20 table rows is not considered.
4. Silent recovery: repeated On Error Resume Next blocks can hide missing controls and return Not Found instead of a precise error.
5. Timing sensitivity: the fixed one-second wait may be too short or unnecessarily long depending on SAP response time.
6. Input assumptions: MRP area is written to both the SAP MRP-area and plant fields without a separate validation or configuration step.
7. Output typing: the SAP date is written as text returned by the GUI. The workbook does not guarantee a normalized Excel date value.
8. Completion count: the message counts row positions, not successfully processed, nonblank records.
9. Planning impact: MD03 is an MRP transaction and can change planning results. Production use requires authorization, testing, and operational controls.

## Publication and confidentiality review

The original workbook should not be uploaded to a public repository unchanged.

### Data to remove or replace

- Real material numbers
- MRP area and plant codes
- Release dates and operational comments
- Any additional rows hidden by the saved filter state
- Real SAP system/client/user identifiers shown in screenshots

### Metadata to remove

Workbook metadata includes an absolute SharePoint/OneDrive-style path tied to an organizational tenant and user directory. Remove or sanitize that metadata before publication.

### VBA project protection

The VBA project contains standard project-protection metadata fields. No plaintext password or credential was observed in the extracted project text. Do not publish the original embedded project binary if the code or protection metadata is proprietary, and do not publish or attempt to bypass any project password.

### What was not observed

No plaintext SAP username, password, API token, supplier list, email address, or database connection string was found in the inspected VBA source or worksheet values. This does not replace a full organizational data-classification review.

## Recommended GitHub packaging

Export the VBA components into readable source files for review, but use a sanitized workbook and synthetic input data:

    README.md
    docs/
      application-workflow.md
      setup-and-troubleshooting.md
      screenshots/
    src/vba/
      Module1.bas
      Module2.bas
      Sheet1.cls
      ThisWorkbook.cls
    sample-data/
      input-template.xlsx
    workbook/
      sanitized-demo.xlsm

The original production workbook should remain outside the public repository. If a functional .xlsm is included, document clearly that it requires SAP GUI Scripting and must not be pointed at a production system without validation.
