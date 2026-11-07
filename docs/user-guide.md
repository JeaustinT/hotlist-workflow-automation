# User guide — synthetic demo

The clean workbook exists at demo/Hotlist_Workflow_Automation_Demo.xlsm. First run the desktop compile and acceptance checks in validation/desktop-checklist.md; source packaging and Python checks do not prove Excel execution.

## Requirements

Use desktop Excel on Windows with macros permitted under your organization's policy. CREATE OUTLOOK DRAFT additionally requires a compatible desktop Outlook profile. Excel Web cannot run the actions. Use only the supplied synthetic fixtures.

## Worksheet roles

Control configures availability, relative capacity, and scope. Raw Hotlist accepts input. Assigned Work is the analyst queue. Exceptions & Review explains excluded or unresolved rows. Assignment Summary displays workload and weighted targets. Inventory History stores run snapshots and progress. Team Directory contains fictional example.com contacts.

## Action order on Control

1. **SAVE PROGRESS** — reconcile edited Agent and Agent Comment values with history; resolve reported unmatched or ambiguous rows and confirm intentional clearing.
2. **PROCESS HOTLIST** — after reviewing availability, weights, and input, validate and assign a new run. Existing progress must reconcile before replacement. Successful publication clears input records while retaining headers.
3. **CREATE OUTLOOK DRAFT** — prepare the current run summary and available analysts' fictional recipients. Review the draft; sending is a separate user action. The demo will not send automatically or include fixed corporate CC recipients.
4. **EXPORT HISTORY** — reconcile progress and choose a destination for a macro-free synthetic snapshot.

These are interface positions, not a requirement to create a draft or export for every run. Analysts filter Assigned Work by owner and edit only designated progress fields. SAVE PROGRESS is distinct from saving the workbook or AutoSave.

If validation or reconciliation fails, inspect the reported rows and correct input without replacing prior work. If processing fails, follow the recovery result and reopen the last saved version when recovery cannot complete. In shared use, one coordinator runs actions and pauses analyst edits during batch publication.

## Opening and resetting

Save a working copy outside the repository before opening. Follow Excel security prompts only for this reviewed file under your normal macro policy. Use Alt+F11, Debug → Compile VBAProject before running the four Control actions. If Excel repairs the file or cannot compile it, stop and record the error; do not claim the demo is validated.

The initial Assigned Work and history contain the same 12-row seed, while Raw Hotlist contains 60 next-run rows. The first seed record has progress text for clearing tests. Restore a clean copy to reset; alternatively rebuild with tools/build_demo.py. Do not paste CSV dates as ambiguous text: copy typed cells from Sample_Raw_Hotlist.xlsx, or import ISO dates with explicit typing.

Use availability_profiles.csv to configure alternate available analysts and weights in Control. Seed history includes a previously available analyst who is unavailable in the default next run. Current notification recipients come from the completed run summary, not changed post-run availability flags.

## Assigned Work layout — retained in demo 0.3.2

| Columns | Fields | Purpose |
| --- | --- | --- |
| A:C | Assigned Analyst; Invoice Type; Run ID | Assignment and traceability at the start of each row. |
| D:H | Location; Publication ID; Publication Name; Client ID; Invoice Number | Core invoice context retained with the assignment block while scrolling. |
| I:M | Invoice Date; Amount; Vendor ID; Batch ID; Status | Primary invoice context. |
| N:O | Hotlist Comment 1; Hotlist Comment 2 | Existing workflow context, separate from analyst progress. |
| P:Q | Agent; Agent Comment | Adjacent, highlighted, unlocked progress fields. |
| R:W | Client Name; Vendor Name; Input Date; Activity Date; Currency; Line ID | Secondary context and stable source-record identity. |
| X:AG | Resolved Media; Workstream; Group Key; Assignment Basis; Volume Class; Duplicate in Run; Invoice Key; Exclusion Reason; Run At UTC; Work Saved At UTC | Classification and technical fields, hidden by default and available by unhiding the right-side columns. |

Freeze panes are set to **I2**: row 1 and A:H stay visible during horizontal/vertical navigation. The narrower demo schema does not require reproducing the operational workbook's exact viewport. The long Run ID is displayed in a compact column; select its cell to read the complete value in the formula bar. The initial zoom is 85% for the wider working queue.

Primary analyst-facing fields are A:Q. Source comments in N:O are read as workflow context; analysts record their progress in P:Q. Both progress columns are yellow and editable. They share the same inventory as the invoice context, rather than requiring a separate update file.

Raw Hotlist, Inventory History, and fixture CSVs retain their canonical order. Assigned Work is a display projection, not a change to invoice identity, grouping, allocation, or history relationships. The VBA read/write mapping and write verification translate between the two orders. Synthetic inputs, histories, and expected assignment CSVs are unchanged.

History exports freeze only row 1 (**A2**, zero horizontal split). This is distinct from the wider Assigned Work freeze. The export implementation uses the destination workbook window's SplitRow/FreezePanes settings and must still be confirmed in desktop testing.
