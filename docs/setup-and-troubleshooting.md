# Setup and troubleshooting

This guide applies to a future sanitized demonstration artifact. The current repository contains documentation only and cannot be run by itself.

## Safety checklist

Before any execution:

- Use a trusted, sanitized workbook.
- Use synthetic input data.
- Confirm the SAP system and client are approved for testing.
- Confirm the active SAP session is the intended session.
- Confirm SAP GUI Scripting is enabled and authorized.
- Review all material rows before starting.
- Do not run the automation against production without formal validation and approval.

## Common conditions

### SAP GUI is not available

Confirm that SAP Logon is open, the user is authenticated, and SAP GUI Scripting is enabled on the client and permitted by the SAP system.

### The wrong SAP session is selected

The documented implementation may select the first connection and first session. Close unrelated sessions or use an explicitly selected session in a future revision.

### A control ID is not found

SAP control IDs can vary by system, screen variant, customization, or transaction state. Record the failure, verify the screen manually, and do not bypass the issue by suppressing errors.

### Results show Not Found

Not Found is not a single failure reason. It may mean that:

- No matching PLA.OR element was visible.
- The associated date was blank.
- The target row was beyond the scan limit.
- SAP returned a different layout.
- A recoverable UI read failed.

Investigate the material manually in SAP and record the reason.

### The run is slow or incomplete

The fixed wait time may not match the SAP response time. A future revision should use explicit readiness checks, bounded retries, and a per-row status log.

### The completion count looks wrong

The documented count is based on row positions rather than successful nonblank records. Treat it as an informational message only until a validated counter is implemented.

## Reporting a defect

For a reproducible issue, record:

- Synthetic material and MRP-area values.
- SAP transaction and screen state.
- Expected result.
- Actual result.
- Whether the row was visible within the first 20 scanned rows.
- Any visible SAP message.
- Sanitized screenshots, if allowed.
