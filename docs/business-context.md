# Business context

Hotlist Workflow Automation addresses the preparation stage between receiving an invoice inventory and analysts beginning invoice processing.

The existing spreadsheet process already supported inventory tracking, formulas, reference lookups, SLA calculations, and operational reporting. Staff still coordinated scope, media and status review, historical comparisons, related invoices, available capacity, allocation, and notification.

The objective is to standardize these decisions, reduce repeated administrative effort, preserve useful history, and deliver an actionable work queue. The reported preparation improvement is approximately 30–45 minutes to approximately 30 seconds; it is not evidence of faster invoice resolution or a measured error-rate reduction.

Primary users are the coordinator who prepares each run and analysts who process assigned work. Reference maintainers keep fictional demo mappings current; a designated coordinator controls batch processing in a shared workbook.

Success criteria for the demo include reconciled input/output counts, group integrity, availability compliance, explainable assignment, preserved progress, and reproducible synthetic scenarios. See [acceptance scenarios](../validation/test-scenarios.md).

## Operational workflow and public demonstration

The project owner reports that the workflow was developed from an internal operational process and is currently used to support recurring Hotlist preparation, assignment, and analyst follow-up. The repository contains a fully synthetic demonstration rather than the operational workbook or production data. No deployment scale, user count, or adoption metric is asserted.

| Stage | Role |
| --- | --- |
| Previous inventory | Spreadsheet-supported tracking, formulas, references, operational reporting, and SLA management, with manual preparation and coordination. |
| Current operational solution | The privately held workbook currently used for preparation, assignment, and analyst follow-up; the primary reference for workflow and usability. |
| Public portfolio demo | Hotlist_Workflow_Automation_Demo.xlsm, built from clean files with fictional records to demonstrate workflow and engineering principles. |

The working queue was reorganized around analyst needs: ownership and invoice context are prioritized, key fields remain visible during navigation, and progress fields are directly editable in the same shared inventory. This is a user-experience and operational-standardization improvement alongside automation. Shared Microsoft 365 behavior in the public demo still requires environment-specific validation.
