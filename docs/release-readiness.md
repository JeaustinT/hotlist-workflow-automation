# Public-release preparation — Phase 4

Status: **private review only**. This phase does not authorize a visibility change, release, license addition, branch deletion, or automatic merge.

## Ready for review

- The exact accepted demo remains version 0.3.2, SHA-256 `33573d72b238baab6c96716912a113ea0ea36da8c71b8e09f44e40527981424f`. It has not been rebuilt for presentation work.
- README leads with the business problem, process redesign, precisely scoped reported preparation improvement, analyst usability, and validated working solution.
- Before/after and architecture SVGs contain only generic workflow text and neutral styling. They are explanatory diagrams, not runtime screenshots.
- Four approved genuine synthetic Excel screenshots are distributed under assets/screenshots/ with their captions and capture/privacy methodology. Full-resolution review found no private identity, account UI, production data, or unrelated desktop content; no unnecessary PNG metadata was present. Do not use the validation Outlook capture with the real From account.
- Compile PASS is now accurately recorded through validation/desktop-acceptance.json and the manifest. Unknown/modified workbook hashes cannot inherit that PASS.
- Build, binary-source verification, expected fixtures, acceptance scenarios, and regression tooling are retained because they demonstrate engineering quality. Historical investigation material is kept under validation rather than dominating the README.
- Remaining rollback, progress-error/clearing, negative-input, stale-lock, alternate-profile, targeted allocation/identity, and shared-client tests remain disclosed. These limits do not negate the accepted portfolio happy path; no all-scenarios runtime claim is made.

## Privacy and repository review

Review tracked text, VBA, accepted/sample workbooks, hidden cells, metadata, links/connections/caches, SVG text/attributes, and relevant tracked-file history against the confidential sources. No original operational workbook, source package, private audit list, production records, real corporate recipients, or source reference mappings may be published. Preview renders and analysis remain outside the repository.

Git author/committer metadata contains the maintainer's identity and a personal email address. This is account/history metadata, not production data in the demo. **The owner should confirm whether that identity/email exposure is intended before public visibility.** No address is reproduced here, and no history rewrite was performed. A rewrite would change commit references and needs separate approval and an evidence-update plan. No corporate email domain was found in the inspected commit metadata.

Future captures also need review of Excel's signed-in profile, title/path, sensitivity labels, tenant URLs, notifications, and Control lock-user cells. The latter can contain the local user's name during a run even though the input data is synthetic. Outlook's configured From address is expected client behavior but must never appear in a public capture.

## Remaining release decisions

1. Review the final screenshot PR and approve its content. Keep the repository private until explicitly authorized otherwise.
2. Confirm the intended public Git author/email metadata policy.
3. Genuine synthetic captures are approved and integrated; retain the privacy review process for future replacements.
4. Decide licensing. **Recommendation: MIT for the generic portfolio code/tooling if ownership and permission to license the work are confirmed.** No license has been added; do not imply reuse rights have already been granted. A different license or an explicitly unlicensed portfolio is an owner decision.
5. Approve any later public visibility change and release separately. Retain the prior development branch during this phase; optional deletion can be considered only with explicit instruction after preparation/review.

## Recommendation

The synthetic implementation, precise impact narrative, accepted main desktop workflow, diagrams, and engineering evidence are suitable for private portfolio review. Hold public visibility until the owner reviews author/email exposure and licensing, and explicitly authorizes publication. A full concurrency/recovery certification is not claimed or required to describe the validated portfolio happy path; broader operational claims still require the pending tests.

## Phase 4 review evidence (before screenshot integration)

- Accepted workbook and all VBA/sample fixtures remain byte-identical to merged main; no workbook rebuild was performed.
- 78 structural/model/binary/receipt checks and 11 binary/build regression tests pass (the original 10 plus the new desktop-evidence hash guard).
- Both SVGs passed XML/accessibility/external-resource checks, were rendered locally, and were visually reviewed. Preview PNGs remain outside the repository.
- The current tree and 98 unique historical tracked blobs were compared against 32,096 confidential candidates, including extracted workbook XML/VBA. The 111 matches were reviewed as generic field/UI/code words, standard Office metadata/theme constants, or independently generated synthetic values. No employee-name, corporate-domain, or literal operational-path matches were found in the reviewed content. Email addresses in content use example.com only. Local links pass.
- Git author/committer metadata was reviewed separately and contains maintainer personal-email metadata, as disclosed above. It has not been removed or approved for public exposure by this phase.
- No actual screenshot, Outlook capture, release, license, visibility change, branch deletion, or automatic merge was created.

## Approved screenshot integration review

- Four owner-approved genuine Excel PNGs were inspected at full resolution, decoded successfully, and checked for metadata. Supplied image bytes were preserved; no Outlook capture is included.
- README image paths and local Markdown links pass. All 78 structural/model/binary checks and 11 regression tests pass again.
- Privacy rescan covered current text, workbook XML/VBA, screenshot OCR, and 116 unique historical tracked blobs against 32,096 confidential candidates. The same 111 generic/standard-metadata/synthetic-value matches remained; no employee/corporate-domain or literal operational-path matches were found. Content email domains remain example.com only. Visual review separately covered account UI and unrelated desktop content.
- The accepted 0.3.2 workbook hash remains unchanged. VBA source, sample data, and runtime-validation evidence were not modified.
- Repository visibility remains private. Public author/email policy, licensing disposition, and explicit publication approval remain owner decisions.
