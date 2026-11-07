# Historical management

Status: the exact 0.3.2 binary passed normal progress-update persistence, cumulative history (60 records after processing), and history-export desktop checks. Clearing/cancellation, unmatched/ambiguous progress, counters, and recovery remain pending; see [the acceptance record](../validation/phase-3-results.md).

Inventory History retains a run snapshot of assigned Digital work, original synthetic fields, composite invoice identity, unique Line ID, group, owner, assignment basis, classifications, timestamp, and Run ID. Exceptions & Review is a current-run view; a complete exception audit trail is outside the initial scope.

Progress matching uses Run ID + composite invoice identity + unique Line ID. Duplicate or ambiguous keys are errors, not arbitrary first matches. History is cumulative, but progress fields are mutable; this is not an immutable change log.

## Save Progress result

Report total evaluated rows, matched rows, updated rows, unchanged rows, unmatched rows, ambiguous rows, skipped rows, values cleared, and validation/runtime errors. Matched = updated + unchanged. Evaluated = matched + unmatched + ambiguous + skipped + validation-error rows. Runtime failures receive a separate operation status.

Skip only empty non-record rows. A relevant populated record with an invalid or missing key is an error. Preflight all matches before updates; unmatched, ambiguous, or invalid relevant rows block the save and block replacement/export. Do not report complete success with unresolved records.

## Deliberate clearing

Clearing Agent or Agent Comment propagates to history only after a confirmation summarizing affected fields and rows. Blank-to-blank is unchanged. Canceling preserves history and blocks dependent replacement/export until the user resolves or restores the pending edits. Validate and back up changes before applying the save batch.

## Export

EXPORT HISTORY requires successful progress reconciliation, then a destination selection. Export synthetic history values and formatting only to `.xlsx`; exclude macros, connections, and links. Cancellation leaves the master unchanged except for an explicitly completed progress save. Export is a snapshot, not an incremental backup or anonymization operation.

The demo represents two deliberately small synthetic runs to prove repeat and progress behavior. Phase 3 includes a 12-row seed and 60-row next input. These are Python-generated synthetic fixtures, not VBA-executed history.
