# Demo acceptance scenarios

Status: **0.3.2 main desktop workflow PASS**, reported by the tester for commit 465b598937632220c40edfecb42b986db8fded33 and pre-open SHA-256 33573d72b238baab6c96716912a113ea0ea36da8c71b8e09f44e40527981424f. The 32 scenarios below remain acceptance criteria; successful baseline outputs do not establish complete runtime coverage of every rule or failure path.

For additional isolated scenarios, record availability, positive weights, scope, mappings, and starting history explicitly; use equal weights only where the scenario calls for them. The accepted baseline used the distributed synthetic fixture. Each additional case starts from a documented clean state; multi-run cases explicitly preserve history.

| ID | Scenario and setup | Expected result |
| --- | --- | --- |
| 01 | New Digital invoice absent from history | NEW and NEW VOLUME; assigned once to an available analyst; history receives the same owner and Run ID. |
| 02 | New Print invoice | Appears in Exceptions & Review with Print reason; no assigned workload or Digital history row. |
| 03 | Out-of-scope location | Exception with scope reason; excluded from assignment and eligible counts. |
| 04 | Released to Clear invoice | Exception with completed-status reason; no assignment or Digital history row. |
| 05 | Repeated composite identity with prior activity but no qualifying new activity | REPEATED and CARRYOVER / REVIEW; remains eligible when otherwise valid. |
| 06 | Repeated identity with both source comments blank | UNTOUCHED REPEAT flag; eligible historical-owner vote generated; no claim that blank comments prove inactivity. |
| 07 | Repeated identity with qualifying structured activity date | REPEATED but NEW VOLUME; no untouched-repeat vote when comments show activity. |
| 08 | Untouched repeat with historical owner available | Entire group returns to the selected historical owner; assignment basis explains exception. |
| 09 | Untouched repeat with historical owner unavailable | No assignment to unavailable owner; group follows eligible votes or normal balancing. |
| 10 | Multiple available analysts | Every eligible group has one available owner; unavailable analysts receive zero work. |
| 11 | Weights 100%, 50%, 100%; controlled equal-size groups and no continuity exception | Targets follow 2:1:2 ratio; allocation follows projected weighted-load rule and recorded tie-breaks; exact equality is not required. |
| 12 | One large Group Key plus smaller groups | Largest group considered first; stays intact; summary exposes target deviations. |
| 13 | Several related invoices in one group | Same Location + Vendor + Client rows share one owner even across publications. |
| 14 | Exact media mapping conflicts with fallback | Valid exact mapping takes precedence; resolved source is explainable. |
| 15 | No exact mapping, valid publication fallback | Fallback classification applies and correct workstream is selected. |
| 16 | Neither mapping resolves media | REVIEW exception; no Digital assignment. |
| 17 | Edit Agent and Agent Comment, then SAVE PROGRESS | Unique run/identity/line match updated; result reports matched and updated counts, no added history row. |
| 18 | Save progress in run one, then process run two | Previous progress retained; new snapshot appended; historical lookup reflects prior run without overwriting its ownership. |
| 19 | EXPORT HISTORY after edits | Progress reconciles first; output is .xlsx with synthetic values/formatting and no macros, connections, or links. Canceling file selection creates no export. |
| 20 | CREATE OUTLOOK DRAFT | Only configured available analysts in To; no fixed CC, all addresses example.com; current summary and run date shown; no automatic send. Missing email blocks creation. |
| 21 | Same invoice number under different vendor or client | Composite identities remain distinct; no false repeat or owner carryover. |
| 22 | Two invoices with same identity and distinct Line IDs | Line-specific progress stays separate; invoice-level repeat and group logic remain consistent. |
| 23 | Duplicate identity + Line ID or missing required key | Validation blocks ambiguous assignment/progress; no partial write. |
| 24 | Progress row unmatched to history | Nonzero unmatched count and incomplete/blocked status; no silent success or replacement/export. |
| 25 | Intentional clearing of previously saved progress | Confirmation shows affected fields; confirm clears matching history, cancel preserves history and blocks dependent actions. |
| 26 | Unchanged progress, empty non-record row, invalid populated row | Distinct unchanged/skipped/error counts reconcile with evaluated count; populated invalid rows are not silently skipped. |
| 27 | Two Run IDs requested in the same second | Unique timestamp-plus-UUID values; uniqueness check rejects an injected duplicate. |
| 28 | Inject handled failure after output write but before final commit | Prior outputs, history extent, input, and metadata restored; application settings and busy guard restored; failure reported. |
| 29 | No available analysts, invalid weight, missing header, or invalid date | Clear blocking validation; prior work and input preserved. |
| 30 | Re-submit identical input | Explicit repeat-input review; cancel produces no new run; approved rerun gets a new Run ID and is labeled accordingly. |
| 31 | Reorder historical worksheet rows | Most recent valid timestamp determines ownership; no dependence on physical row order. |
| 32 | Reconciliation and version alignment | Assigned + exception counts equal nonblank input; group ownership invariant holds; four buttons match reviewed procedures and version. |

For export and Outlook cases inspect generated artifacts in the supported desktop environment. Shared hosting, AutoSave, and simultaneous editing need a separate environment check; do not infer them from local scenario success.

## Actual evidence and desktop coverage

The Python checker passes 78 structural/source/model/receipt assertions; 11 binary/build regression tests also pass (the original 10 plus a hash-bound desktop-evidence regression). These checks do not execute VBA. Separately, the tester verified the exact 0.3.2 binary's first-open, macro enablement after Windows unblock, references, compilation, analyst UX/I2 navigation, normal progress save, processing, exception counts, history, export/A2 freeze, and Outlook draft. See [runtime results](phase-3-results.md).

“PASS (reported portion)” below accepts only the stated observation. Remaining criteria in the same scenario are still pending; no row implies unreported counters, cancellation, failure handling, or per-record rule inspection.

| Scenario IDs | Actual desktop evidence | Remaining runtime coverage |
| --- | --- | --- |
| 01 | Baseline processing produced 48 assignments; no specific new-invoice identity/owner audit was reported | Targeted NEW / NEW VOLUME row and historical-identity check pending. |
| 02 | PASS (reported routing/count): 4 Print media exceptions | Separate negative/row-specific checks not inferred. |
| 03 | PASS (reported routing/count): 3 out-of-scope location exceptions | Separate configuration/row-specific edge checks not inferred. |
| 04 | PASS (reported routing/count): 2 Released to Clear exceptions | Separate row-specific edge checks not inferred. |
| 05–09 | Baseline summary includes 39 New Volume and 9 Carryover rows, but no isolated repeat/continuity branch verification was reported | Repeat/activity classification, untouched-owner rules, unavailable owner, and vote/tie checks pending. |
| 10 | PASS (default-profile totals): Alex Morgan 20, Jordan Lee 8, Taylor Smith 20; total 48 | Alternate availability combinations and row-level compliance tests pending. |
| 11–15 | Baseline output totals match the model | Isolated weight ratios, large-group ordering/integrity, related invoices, exact-match precedence, and fallback branch checks pending. |
| 16 | PASS (reported routing/count): 3 unresolved media classification exceptions | Additional unresolved-mapping edge/configuration tests pending. |
| 17 | PASS (successful update path): edited progress propagated to history and exported history | Detailed reconciliation counters, skipped/error branches, unmatched/ambiguous keys, and clearing not verified. |
| 18 | PASS (reported persistence path): saved progress persisted through processing; history reached 60 records | Historical-owner lookup and preservation of prior ownership need separate targeted inspection. |
| 19 | PASS (successful export): one Synthetic History sheet, 60 records plus header; no VBA, external links, or connections; filters, formatting, progress, A2/header-only freeze retained | Export cancellation and failure/recovery paths pending. |
| 20 | PASS (successful draft): correct three example.com recipients, empty CC, synthetic subject, Run ID, summary/disclaimer; no automatic send | Missing/invalid email and post-run availability-change cases pending. Real configured From is expected; no real sender may appear in future public screenshots. |
| 21–22 | Static/model identity and distinct-line fixtures pass | Targeted desktop composite/line identity verification pending. |
| 23–26 | Source/static checks only; successful normal progress saving does not establish these cases | Invalid/duplicate/missing keys, unmatched/ambiguous handling, clearing confirmation/cancel, and counters pending. |
| 27–31 | Static/model/source evidence only | Same-second Run IDs, rollback injection, negative input, reruns, and reordered-history tests pending. |
| 32 | PASS (reported portions): exact 0.3.2 binary compiled; four main actions ran; 48 + 12 = 60 input rows reconciled; status COMPLETED; input cleared and normal lock released; I2/P:Q/technical visibility verified | Per-group runtime invariant and recovery/concurrency checks remain pending. |

Outlook runtime summary: Alex Morgan 15 New Volume / 5 Carryover / 20 Assigned; Jordan Lee 8 / 0 / 8; Taylor Smith 16 / 4 / 20. Total 39 / 9 / 48. No unreported per-invoice decisions are inferred from this aggregate summary.

The runtime working copy had 60 historical records after processing. The distributed workbook remains the unchanged clean 0.3.2 fixture (12 seed history rows and 60 raw rows), not the tester's saved working copy. The machine-generated manifest now records compilation PASS through the hash-bound desktop-acceptance.json evidence. Different artifact bytes receive no inherited desktop PASS.

Active/stale locks, rollback/fault injection, ambiguous progress, negative inputs, alternate profiles, and shared Teams/SharePoint/coauthoring tests remain pending as detailed in the [desktop checklist](desktop-checklist.md). Shared-client behavior is environment-dependent; advisory locking is not a distributed transaction guarantee. No screenshots or public release have been created.
