# Public demo specification

Status: approved Phase 3 design implemented in clean synthetic artifacts. The exact 0.3.2 binary passed desktop compilation and main-workflow acceptance; edge/recovery/shared-client coverage remains pending. See [the acceptance record](../validation/phase-3-results.md). Four genuine synthetic Excel Desktop screenshots are approved and included.

## Authoritative version

The reviewed source under src/vba and the clean workbook generated for the same demo revision form one release pair. Version 0.3.2 is recorded in Control and source. Documentation follows that pair. Historical operational files are reference evidence, not release artifacts. Do not reuse an embedded operational VBA project.

Control exposes these four actions in order: SAVE PROGRESS, PROCESS HOTLIST, CREATE OUTLOOK DRAFT, EXPORT HISTORY. Align labels, assigned procedures, English messages, export behavior, and documentation before capture or release.

## Synthetic model

Use these deliberately fictional labels only, with no operational mappings reused:

| Entity | Proposed fictional convention |
| --- | --- |
| Analysts | Alex Morgan, Jordan Lee, Taylor Smith, Casey Brown, Morgan Davis |
| Analyst IDs | DEMO-ANL-A through DEMO-ANL-E |
| Email | alex.morgan@example.com and equivalent fictional example.com addresses |
| Client | Demo Northstar Client; DEMO-CLI-A |
| Vendor | Demo BluePeak Vendor; DEMO-VEN-A |
| Publication | Demo Vertex Publication; DEMO-PUB-A |
| Other organization | Demo Summit Organization; DEMO-ORG-A |
| Locations | DEMO-LOC-A, DEMO-LOC-B; DEMO-LOC-X outside configured scope |
| Invoice | DEMO-INV-A001 and subsequent fictional sequence |
| Batch / line | DEMO-BAT-A and unique DEMO-LINE-A001 sequence |
| Amount | Explicitly fictional decimal values such as 125.00; one documented demo currency |
| Dates | Fixed ISO 8601 dates within a fictional two-run period, unrelated to operational dates |
| Comments | Neutral fictional text such as “Demo follow-up requested” |
| History | Two small constructed runs supporting the acceptance scenarios |
| Mappings | Constructed exact-match, fallback, Print, and unresolved examples |

This table defines conventions, not a production-looking historical dataset. Phase 3 must check candidate labels against confidential source entities before fixture generation.

## Data contract

Keep identifiers as text. Required fields cover Location, Publication ID, Client ID, Vendor ID, Invoice Number, Line ID, status, source activity comments, input date, amount, and analyst progress. Use structured Activity Date for reliable volume classification. Qualifying activity is strictly after the previous run calendar date, not on the same date. Dates are typed and displayed consistently; fixture interchange uses ISO 8601. Keep original source activity and analyst progress separate.

Use a minimal schema sufficient to demonstrate the workflow rather than copying every operational field. Validate required headers, duplicates, unknown analyst IDs, missing emails, nonpositive weights, invalid dates, and blank identities. Reconcile input rows as assigned + exceptions, excluding fully blank rows. No eligible row may disappear or receive two owners.

## Implemented technical decisions

| Area | Proposal | Benefit and tradeoff |
| --- | --- | --- |
| Invoice identity | Normalized Location + Vendor + Client + Invoice Number | Reduces unrelated-number collisions; changing identity fields creates a new identity. Publication corrections remain linked. Requires domain confirmation. |
| Line identity | Unique synthetic Line ID per row; progress key includes it | Prevents ambiguous updates among repeated invoice lines; requires stable IDs. |
| Run ID | UTC timestamp prefix plus full UUID, e.g. timestamp-UUID format | Human-sortable prefix with collision-resistant suffix; longer than a timestamp. Validate uniqueness against history and stop with a retry instruction on conflict. |
| Historical ownership | Latest valid run timestamp, with deterministic Run ID tie resolution | Independent of worksheet sort order; invalid historical dates require review. |
| Save Progress | Preflight all keys and report matched, updated, unchanged, unmatched, ambiguous, skipped, and errors | No silent success for unresolved records; blocking errors require correction. |
| Clearing | Propagate deliberate clearing after a specific confirmation | History reflects corrections; confirmation prevents accidental deletion. |
| Activity | Structured Activity Date rather than free-text date extraction | Predictable classification; adds one input field. |
| Recipients | Available analysts from configurable fictional directory; no fixed CC | Reusable configuration; missing addresses block draft creation. |
| Reprocessing | Full ordered input signature plus explicit rerun review; use a new Run ID for an approved rerun | Reduces accidental repeated runs without treating a hash as invoice identity. Do not promise global idempotency. |

Public class names and assignment rules are defined in [terminology](terminology.md) and [assignment logic](assignment-logic.md). Saved progress counts and clearing semantics are defined in [historical management](historical-management.md).

## Reasonable transaction control

1. Require a single coordinator; set an in-session busy guard and retain prior Excel application settings.
2. Reconcile prior progress as a separately completed operation. Abort on unresolved relevant rows; preserve input and existing outputs.
3. Validate all input/configuration and calculate assignments, exceptions, summaries, and history rows in memory.
4. Check reconciled counts, group ownership, valid recipients where relevant, and unique run identifiers before any publication.
5. Back up affected ranges and control metadata in memory. Maintain a saved recovery copy outside the repository before processing the demo.
6. Publish outputs and append history in a guarded section. Record successful run metadata and clear input last.
7. On a handled failure, restore affected ranges, remove newly appended rows, restore input and metadata, and report whether rollback completed. Always restore prior application settings and release the guard.
8. If recovery fails or Excel terminates, use the saved prior version. Do not claim database-level atomicity or distributed locking.

## Acceptance and exclusions

Screenshots may depict the accepted synthetic happy path after the [capture/privacy checklist](../assets/screenshots/README.md) passes. Keep untested [scenario branches](../validation/test-scenarios.md) disclosed; images must not imply all 32 scenarios passed. No automatic email sending, live external reference refresh, production records, production integration, immutable audit system, or concurrent macro execution is planned. Performance remains unmeasured until the benchmark protocol is executed.

## Approved implementation and next gate

The approved identity, structured date, continuity priority, clearing safeguard, and practical recovery controls are implemented in source. The workbook contains 60 input records and a 12-row seed. Five analysts use example.com addresses. Four standard modules are embedded without compiled caches; empty workbook/worksheet modules are generated by the clean builder.

A cell-based token/user/UTC lock covers all four actions. Recovery is blocked until 30 minutes have elapsed, then requires coordinator confirmation. This does not provide atomic cross-client locking; one coordinator and a pause in analyst edits remain required.

For this manageable demo, use a full ordered normalized input signature rather than a hash; reject signatures over Excel's cell limit. Duplicate Line IDs block the run. Save Progress reports proposed counts on blocked/canceled saves and committed counts only on success. Prior progress reconciliation is a separate completed operation and is not undone if a subsequent new-run commit fails.

The preloaded seed is generated, not VBA-executed. The main desktop workflow is accepted. Current work covers portfolio diagrams, four approved genuine Excel screenshots, and release-readiness review. Remaining edge/recovery/shared-client tests stay documented. Public publication remains unauthorized.

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
