# Assignment logic

Status: implemented in demonstration source; static/model validation completed. The 0.3.2 desktop baseline totals (48 assigned / 12 exceptions; workloads 20 / 8 / 20) passed. Isolated identity, allocation, continuity, and edge-case runtime checks remain pending; see [the acceptance record](../validation/phase-3-results.md).

## Eligibility and identity

Route Released to Clear, out-of-scope, Print, unresolved-media, and invalid-identity records to Exceptions & Review with explicit reasons. Resolve media by Location + Client + Publication first, then Publication Fallback. Only validated DIGITAL rows enter assignment.

Use Location + Vendor + Client + Invoice Number for historical identity. Group Key remains Location + Vendor + Client. Preserve identifiers as text and leading zeros; normalize case and surrounding whitespace. Escape or length-prefix components rather than joining ambiguous raw delimiters. Missing identity fields block the entire run. Publication is excluded from identity so a publication correction does not automatically create a new invoice; this is the approved domain assumption for the synthetic demo.

Multiple rows sharing an identity are flagged in Duplicate in Run. Preserve related line items when unique Line IDs exist; Line ID must be globally unique within each input run; repeated Line IDs are a blocking error. Ambiguous duplicate identities without Line IDs require review rather than silent deduplication.

## Allocation

1. Include only available analysts with validated positive capacity weights.
2. Sort groups by descending eligible row count, then Group Key for reproducibility.
3. For untouched repeats, vote for available historical owners. Most votes wins; resolve ties by weighted load and then stable analyst ID. This exception may outweigh balance.
4. Otherwise choose the analyst minimizing `(assigned rows + group rows) / weight`.
5. Resolve ties by assigned row count and then stable analyst ID.
6. Keep all eligible rows of a group with one analyst and record the assignment basis.

Weighted target = total eligible rows × analyst weight / total available weight. Report actual, target, and difference. A 50% weight represents half the relative capacity of 100%; it does not guarantee an exact half-count when groups are indivisible.

## Activity and history

NEW/REPEATED reflects prior composite identity. NEW VOLUME is a separate metric: a new identity or repeated identity with validated activity dated strictly after the preceding successful run calendar date (UTC). Use a structured ISO date field rather than parsing free text. Activity Date may be blank, meaning no qualifying new activity. Nonblank dates must be typed Excel dates; invalid dates block processing. Never infer completed work from comment wording alone.

Untouched-repeat eligibility means REPEATED, no qualifying new activity, and both Hotlist Comment 1 and Hotlist Comment 2 blank. It is a proxy, not proof of inactivity. Other repeated work is CARRYOVER / REVIEW unless qualifying new activity is recorded. Most recent historical owner is selected by recorded run time, not physical row position. The legacy preferred-owner table is not an assignment input.

## Distinct keys and ownership priority

Invoice Key identifies the same invoice across runs. Line ID identifies one source record and is unique within an input run. Group Key determines which records must stay with one analyst; it is not invoice identity. Progress uses Run ID + Invoice Key + Line ID.

Repeated + untouched + available previous owner uses **Repeated untouched - previous owner**. Previous owner unavailable, new invoice, or repeated invoice with qualifying new activity uses **Weighted balanced allocation**, except that other eligible untouched repeats in the same group may determine the whole group's owner. No Group Key is permanently bound to a person. Highest available-owner vote wins; tied votes use lowest current weighted workload, then stable analyst ID.
