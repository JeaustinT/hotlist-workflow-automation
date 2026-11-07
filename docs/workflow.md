# Workflow

Status: the exact 0.3.2 binary passed the main end-to-end desktop workflow. Edge, recovery, cancellation, and shared-client checks remain pending; see [the acceptance record](../validation/phase-3-results.md).

1. The coordinator opens the desktop workbook and reviews Control availability and positive relative weights.
2. Load synthetic records into Raw Hotlist using the agreed schema. Review date and identity validity.
3. SAVE PROGRESS reconciles existing Assigned Work edits with Inventory History. Resolve unmatched or ambiguous records before processing another run.
4. PROCESS HOTLIST validates input and configuration, classifies rows, looks up history, forms groups, and calculates assignment and reconciliation totals in memory.
5. Review any blocking errors. On successful commit, publish Assigned Work, Exceptions & Review, Assignment Summary, and history under one Run ID. Clear input only after successful writes.
6. Analysts filter Assigned Work by assigned analyst and update Agent and Agent Comment.
7. SAVE PROGRESS persists updates; intentional clearing requires the documented confirmation.
8. CREATE OUTLOOK DRAFT produces a reviewable notification for configured available analysts. Sending remains a human action.
9. EXPORT HISTORY reconciles pending progress and exports a macro-free values-and-formatting snapshot after selecting a destination.

Preparation is one processing action after setup. Progress, email, and export are distinct actions. In shared use, one coordinator runs macros; simultaneous batch processing is unsupported. See [transaction controls](demo-specification.md) and [user guide](user-guide.md).
