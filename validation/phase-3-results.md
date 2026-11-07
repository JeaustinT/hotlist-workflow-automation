# Phase 3 validation results

## Current desktop acceptance — main workflow PASS

Version **0.3.2**, workbook from commit **465b598937632220c40edfecb42b986db8fded33**.

Pre-open SHA-256 verified by the desktop tester against the repository manifest:

```text
33573d72b238baab6c96716912a113ea0ea36da8c71b8e09f44e40527981424f
```

Results were reported by the desktop tester on 2026-11-05 using Windows desktop Excel and desktop Outlook. Exact application builds/bitness were not supplied. These are real desktop observations for this binary, distinct from the assistant's Python/static checks. They do not establish coverage of every acceptance scenario or a performance benchmark.

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

## Verified Outlook details

Only these synthetic recipients appeared in To: `alex.morgan@example.com`, `jordan.lee@example.com`, and `taylor.smith@example.com`. CC was empty. The draft had the correct synthetic subject, Run ID, and disclaimer; no automatic send occurred.

| Fictional analyst | New Volume | Carryover | Assigned |
| --- | --- | --- | --- |
| Alex Morgan | 15 | 5 | 20 |
| Jordan Lee | 8 | 0 | 8 |
| Taylor Smith | 16 | 4 | 20 |
| Total | 39 | 9 | 48 |

The Outlook client's configured real sender account appears in From. This is expected client behavior, not an automation defect. No real sender address is recorded here. Future public portfolio captures must exclude or redact the real sender; no screenshots are created in this phase.

## Static/model/binary validation — separate evidence

The assistant reran the existing checks in the cloud environment; these do not execute Excel:

- **78 structural/source/model/receipt assertions: PASS.** Workbook version remains 0.3.2; exact headers, I2 layout, progress mapping, source alignment, and receipt are verified. The model expects 48 assigned / 12 exceptions and 20 / 8 / 20 workloads, consistent with the reported desktop run.
- **11 binary/build regression tests: PASS.** The original 10 tests are retained; Phase 4 adds a check that desktop compile evidence cannot transfer to a different workbook hash. Deliberately mismatched embedded modules, stale declarations/calls, duplicate project parts, and missing modules are rejected. Failed staging validation cannot replace the prior artifact; a bogus legacy project cache is ignored.
- The final XLSM is reopened and its four actual standard-module sources are extracted with oletools and cross-checked with direct OLE decompression against src/vba. Module-by-module comparison is PASS.
- The accepted XLSM, VBA, fixtures, version, and generated manifest are unchanged by this documentation update. The distributed clean file retains 12 seed records and 60 raw rows; the tester's processed working copy reached 60 historical records.
- The manifest now records `excel_desktop_compile: PASS`, supported by validation/desktop-acceptance.json and bound to the exact accepted hash/version. The verifier never derives runtime success from source equality: an artifact without matching desktop evidence is reported as NOT VERIFIED FOR THIS ARTIFACT. The accepted workbook is unchanged; only metadata/tooling were updated in Phase 4.
- Reproduce with tools/verify_embedded_vba.py, validation/validate_demo.py, and validation/test_embedded_vba.py using the commands in tools/README.md. No new timing trials or benchmark claims were introduced.

## Remaining desktop tests

| Still-pending desktop coverage | Boundary |
| --- | --- |
| Rollback and fault injection | No runtime failure injection or recovery/restoration result was reported. |
| Unmatched and ambiguous progress | Successful saving does not verify blocking behavior, error counters, or duplicate/missing history matches. |
| Progress clearing and cancellation | Clear-value confirmation, confirm/cancel behavior, skipped/unchanged/error counters, and export cancellation remain unverified. |
| Negative input/configuration | Invalid identities, duplicate Line IDs, missing headers, invalid dates/weights, no available analysts, unknown IDs, and bad mappings remain pending. |
| Active/stale lock handling | Normal lock release passed; blocked concurrent actions, stale-age threshold, confirmation, and recovery remain pending. |
| Alternate profiles and targeted allocation cases | No alternate availability/weight profile, ownership-vote/tie case, isolated activity-date boundary, or per-group runtime integrity review was reported. |
| Other identity/repeat edge cases | Same-number/different-entity behavior, line-specific repeated identities, same-second Run IDs, rerun detection, and physically reordered history need targeted desktop checks. |
| Negative Outlook/configuration | Missing/invalid email and availability changes after a run remain pending. |
| Teams/SharePoint and multiple users | AutoSave, lock visibility, simultaneous editing, coauthoring, and cross-client behavior remain environment-dependent and unverified. The workbook lock is advisory, not an atomic distributed transaction lock. |

These limitations prevent claims of exhaustive runtime acceptance or validated shared-client concurrency. They do not negate the verified main end-to-end workflow. The [scenario record](test-scenarios.md) marks successful portions and pending branches separately.

## Confidentiality and merge scope

The privacy rescan passed: 32,096 source-derived candidates were checked; generic/standard metadata and synthetic-value matches were reviewed, with no employee-name, corporate-domain, or literal operational-path matches. Only example.com email addresses occur. Local documentation links passed.

Privacy validation covers repository text, extracted visible/hidden workbook content, metadata, embedded VBA, and source-derived candidate comparisons. Reviewed matches are generic interface/code terms, standard Office metadata, or independently generated synthetic values. No confidential source workbook, operational records, private comparison lists, or corporate email addresses are tracked. Only the clean synthetic demo and sample workbook are distributed.

The project owner's operational-use statement and the distinction between the current private workflow and public synthetic demo remain unchanged. No deployment scale, adoption count, invoice-processing speedup, or synthetic-demo benchmark is claimed.

This acceptance supports review for a private documentation/demo merge with the pending coverage disclosed. Public visibility, screenshots, and automatic merging remain unauthorized.

## Historical failures and artifact investigation

The initial build failed compilation at a reserved `input` identifier. A later reported 0.3.1 failing copy had different bytes from the exact committed artifact. The [investigation](embedded-vba-investigation.md) preserves those historical findings and the build/receipt safeguards. Its pending statuses describe the state before the successful 0.3.2 desktop test above; they are not the current runtime acceptance result.
