# Public terminology standard

This is the single naming reference for future documentation, code, workbook labels, fixtures, and screenshots. All public-facing content uses professional English.

| Concept | Public name |
| --- | --- |
| Configuration worksheet | Control |
| Incoming inventory | Raw Hotlist |
| Digital invoice work queue | Assigned Work |
| Exclusions and unresolved items | Exceptions & Review |
| Cumulative run records | Inventory History |
| Allocation totals and targets | Assignment Summary |
| Analyst contact configuration | Team Directory |
| Specific media mapping | Media Reference |
| Publication-level mapping | Publication Fallback |

Actions appear in this order: **SAVE PROGRESS**, **PROCESS HOTLIST**, **CREATE OUTLOOK DRAFT**, **EXPORT HISTORY**.

| Classification | Meaning |
| --- | --- |
| NEW | Invoice identity absent from prior history |
| REPEATED | Invoice identity present in prior history |
| MISSING INVOICE NUMBER | Invalid identity requiring review |
| NEW VOLUME | New identity or repeated identity with qualifying new activity |
| CARRYOVER / REVIEW | Repeated work without qualifying new activity; does not prove no work occurred |
| UNTOUCHED REPEAT | Repeated item whose two source activity comments are blank; a documented proxy |
| DIGITAL | Eligible workstream when other validation passes |
| PRINT | Excluded workstream |
| REVIEW | Classification unresolved or input invalid |
| Released to Clear | Completed source status excluded from assignment |

Invoice identity, duplicate-within-run flags, and volume classification are separate concepts. A repeated invoice can count as NEW VOLUME. Group Key means Location + Vendor + Client; weight means relative analyst capacity, not complexity. See [assignment logic](assignment-logic.md).
