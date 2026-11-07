# Impact and measurement

## Reported operational result

| Measure | Reported value |
| --- | --- |
| Previous preparation and assignment | Approximately 30–45 minutes |
| Automated preparation and assignment | Approximately 30 seconds |
| Time saved per Hotlist | Approximately 29.5–44.5 minutes |
| Reduction | Approximately 98.3%–98.9% |

Reduction = `(baseline seconds - automated seconds) / baseline seconds * 100`.

At 1,800 seconds the reduction is 98.33%; at 2,700 seconds it is 98.89%. These rounded calculations express the owner's reported operational experience. They are not independent benchmarks, and the approximate inputs do not support greater precision.

The claim concerns administrative preparation, validation, assignment, and coordination before invoice processing. It does not establish invoice throughput, resolution time, financial savings, error-rate improvements, or annual labor savings.

## Future benchmark protocol

Record two timing boundaries separately:

- Processing runtime: prepared input and configuration, immediately before PROCESS HOTLIST, through successful output/history publication.
- End-to-end preparation: begin receipt/review of the inventory through a ready assigned queue and the agreed communication milestone. Include manual configuration and draft review consistently on both sides.

Do not compare a narrow processing runtime with a broad manual baseline without labeling the scope difference. Source acquisition, email sending, export, and analyst processing must be explicitly included or excluded in each measurement.

Record Excel/OS version, hardware, input and history row counts, group sizes, analyst count, weights, workbook location, AutoSave state, successful trial count, warm/cold conditions, failures, and median/range. Use only synthetic data for reproducible demo trials. Do not assume that a small demo reproduces production timing. See [benchmark status](../validation/benchmark-results.md).
