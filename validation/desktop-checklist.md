# Excel Desktop acceptance checklist

Status: **0.3.2 main end-to-end desktop workflow PASS; edge/recovery/shared-client coverage remains pending.**

Version **0.3.2**, workbook from commit **465b598937632220c40edfecb42b986db8fded33**.

Pre-open SHA-256 verified by the desktop tester against the repository manifest:

```text
33573d72b238baab6c96716912a113ea0ea36da8c71b8e09f44e40527981424f
```

Results were reported by the desktop tester on 2026-11-05 using Windows desktop Excel and desktop Outlook. Exact application builds/bitness were not supplied. These are real desktop observations for this binary, distinct from the assistant's Python/static checks. They do not establish coverage of every acceptance scenario or a performance benchmark.

## Completed runtime checks

| Area | Desktop result | Verified evidence |
| --- | --- | --- |
| Artifact identity | PASS | Exact downloaded 0.3.2 binary hash matches the repository manifest before opening. |
| Workbook open | PASS | Opened without repair. |
| Macro enablement | PASS | Macros enabled successfully after Windows unblock. |
| VBA references | PASS | No `MISSING:` references observed. |
| VBA compilation | PASS | Debug → Compile VBAProject succeeded for the identified binary. |
| Assigned Work UX | PASS | Assignment fields lead, invoice context follows, Agent / Agent Comment are adjacent and highlighted, technical fields are initially hidden. I2 freezes row 1 and A:H; Assigned Analyst and Invoice Number remain visible while accessing later columns. |
| SAVE PROGRESS | PASS | An Agent / Agent Comment update propagated from Assigned Work to Inventory History and appeared in the exported history. This verifies the successful update path, not clearing or unmatched/ambiguous handling. |
| PROCESS HOTLIST | PASS | 48 assigned, 12 exceptions, status COMPLETED; Raw Hotlist cleared and the normal processing lock released after success. |
| Assignment totals | PASS | Alex Morgan 20; Jordan Lee 8; Taylor Smith 20; total 48. These are fictional analysts in the demo. |
| Exception routing/counts | PASS | 4 Print media, 3 out-of-scope location, 2 Released to Clear, 3 unresolved media classification; total 12. |
| Inventory History | PASS | 60 historical records after processing; saved Agent / Agent Comment persisted. |
| EXPORT HISTORY | PASS | One Synthetic History worksheet, 60 records plus header, no VBA, external links, or workbook connections; filters, formatting, and saved progress retained. A2 freezes only row 1, with no horizontal split. |
| CREATE OUTLOOK DRAFT | PASS | Draft opened in desktop Outlook; not automatically sent. Correct synthetic subject, Run ID, assignment summary, disclaimer, and recipients; CC empty. |

The Outlook summary was Alex Morgan 15 New Volume / 5 Carryover / 20 Assigned; Jordan Lee 8 / 0 / 8; Taylor Smith 16 / 4 / 20. To contained only `alex.morgan@example.com`, `jordan.lee@example.com`, and `taylor.smith@example.com`. CC was empty. Correct synthetic subject, Run ID, and disclaimer were observed; the message was not automatically sent.

Outlook's real configured sender in From is expected and is not a defect. Future public captures must exclude/redact that account address. No screenshot is authorized or created by this update.

## Pending runtime checklist

Use a separate synthetic working copy and retain a recovery copy. Record revision, original hash, Excel/Outlook versions, setup, observed result, and exact errors. A successful normal run does not mark these pending checks PASS.

1. **Progress clearing:** clear saved Agent/Agent Comment values. Confirm the field count; test confirm and cancel separately. Check exact history updates and dependent processing/export behavior.
2. **Unmatched/ambiguous progress:** alter a relevant key or create ambiguous history matches. Verify blocking status, counts, no silent success, and no partial history updates.
3. **Progress counters:** verify unchanged rows, fully blank non-record rows, invalid populated rows, updated/cleared counts, and error reporting independently.
4. **Rollback:** inject FAIL_AFTER_HISTORY or force a destination-write failure. Verify restoration of prior outputs, table extents, input, history extent, metadata, and Excel settings. Verify explicit incomplete-recovery reporting when appropriate. Previously completed progress saves remain a separate operation.
5. **Negative input/configuration:** test duplicate/missing Line ID and invoice keys, missing headers, invalid typed dates/weights, no available analysts, unknown IDs, and invalid mappings. Check preservation of prior work/input.
6. **Locks:** verify active-lock refusal for all actions, recovery blocked below 30 minutes, confirmed stale recovery above the threshold, and cancel retaining the lock. The observed successful lock release does not cover these cases.
7. **Alternate profiles:** apply ALTERNATE/ALL fixtures and different weights in clean copies. Check availability compliance, weighted targets, and full group integrity at row level.
8. **Ownership/activity rules:** isolate new/active/untouched repeated cases, unavailable historical owners, conflicting votes/ties, and date boundaries. The baseline totals alone do not verify every decision branch.
9. **Identity/history edge cases:** compare same invoice number across entities and multiple Line IDs per invoice; reorder historical rows; check latest ownership and line-specific progress.
10. **Run identity/reruns:** request IDs within one second; test identical-input cancel and approved rerun paths.
11. **Office cancellation/configuration:** cancel export destination selection; test missing/invalid draft recipients and post-run availability changes. Successful export/draft tests do not cover these negative paths.
12. **Shared environment:** in the intended Teams/SharePoint setup, test AutoSave, lock-cell visibility, paused analyst edits, and multi-user coauthoring. Keep one coordinator for batch actions. Cell check/set locking is not atomic across clients; do not claim full concurrent transaction safety.

## Evidence boundaries

The [results](phase-3-results.md) and [scenario matrix](test-scenarios.md) distinguish tester-reported desktop results from Python/static evidence. The historical investigation preserves its original build-time state. The current manifest records compile PASS through hash-bound desktop evidence, only for the exact original 0.3.2 artifact above. No rebuild or modified working copy is substituted for that artifact.

No performance benchmark, screenshots, public release, or automatic merge is included. The current private operational solution remains separate from this synthetic portfolio demonstration.
