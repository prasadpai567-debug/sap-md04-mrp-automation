# Application workflow

## Business workflow

The documented process starts with a list of material numbers and MRP areas in Excel. The operator reviews the input, confirms the SAP environment, and starts the macro.

For each nonblank material row:

1. The automation connects to the existing SAP GUI scripting engine.
2. MD03 is opened and the material, MRP area, and plant/MRP-area value are entered.
3. MD04 is opened for the material.
4. A recorded list filter is applied.
5. The visible MRP table is scanned for an element containing PLA.OR.
6. The first nonblank associated date is selected.
7. The result is written to the Release Date column.
8. If no qualifying result is identified, Not Found remains in the output cell.

After the batch, the automation returns SAP to /N, releases object variables, and displays a completion message.

## Data mapping

| Excel field | Documented use |
| --- | --- |
| Column A: Material Number | Material passed to MD03 and MD04 |
| Column B: MRP AREA | MRP-area and plant/MRP-area input |
| Column C: Release Date | First qualifying date or Not Found |
| Column D: Comments | Optional user-maintained notes; not populated by the macro |

## Technical boundaries

The documented process reads visible SAP screen values. It does not export a complete MD04 table, create purchase orders, send emails, or generate a separate report sheet.

The documented scan is limited to 20 MRP table rows per material. SAP screen control IDs and the first-session selection are environment-dependent.

## Portfolio demonstration

A safe demonstration should use:

- Synthetic material numbers.
- Redacted SAP screenshots.
- A non-production or mock environment.
- No supplier, customer, employee, or company-identifying data.
- A clearly labeled result showing both a found date and Not Found.
