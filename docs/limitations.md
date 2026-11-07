# Engineering boundaries

- VBA requires desktop Excel; Excel Web does not execute VBA.
- The initial demo target is Windows desktop Excel. Other platforms require separate validation, particularly dictionary and Office automation behavior.
- Outlook draft generation depends on compatible desktop Outlook and automation availability. Classic desktop Outlook is the intended initial target; compatibility with other clients is not assumed.
- Whole-group allocation is a heuristic. Group size and historical continuity can prevent exact mathematical equality or target attainment.
- Workload uses eligible invoice row counts and group counts, not actual processing complexity, time, or analyst skill.
- Reference mappings require maintenance. Unresolved classifications are routed to review.
- Shared Microsoft 365 operation requires suitable Teams/SharePoint configuration, permissions, and desktop clients. AutoSave and coauthoring require validation; they do not provide macro transactions or concurrent-run locking.
- One coordinator runs batch actions. A workbook-cell lock records user, UTC time, and token. It blocks visible active locks but is advisory, not an atomic distributed lock.
- History snapshots with editable progress are not an immutable audit log. The initial demo does not archive every excluded input across runs.
- Implemented rollback controls reduce partial-write exposure but cannot guarantee recovery from Excel termination or storage failure. A saved prior version remains necessary.
- Composite identity assumes invoice uniqueness within Location + Vendor + Client. Ambiguous line items require a unique Line ID and review.
- The approximately 30–45 minute to approximately 30 second result is reported operational experience. No demo performance result has been measured.

These boundaries guide a focused, explainable demonstration. The exact 0.3.2 binary passed tester-reported Excel Desktop compilation and main-workflow acceptance. Edge/recovery and shared-client coverage remains pending; see [runtime evidence and limits](../validation/phase-3-results.md). The source-only VBA project intentionally contains no compiled caches. Input size is limited by a 32,767-character exact signature stored in one cell; oversized input is rejected. This is a small demonstration, not a scalable inventory platform.
