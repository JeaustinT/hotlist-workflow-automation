# Embedded VBA investigation and 0.3.2 acceptance

> Historical build-time investigation. The exact 0.3.2 binary subsequently passed tester-reported desktop compilation and the main end-to-end workflow. See [current acceptance and pending coverage](phase-3-results.md). Pending statuses below preserve the earlier investigation state, not the current runtime result.

## Desktop failure and artifact identity

The tester reported that a downloaded 0.3.1 workbook failed Debug → Compile VBAProject at the old reserved `input` declaration, despite corrected repository source. That tested copy is invalid for acceptance. At investigation time, desktop compilation had not yet been accepted. The exact 0.3.2 artifact subsequently passed desktop compilation and the main workflow; see phase-3-results.md.

Investigation extracted the actual XLSM Git objects and independently fetched the workbook from GitHub's contents API at the full 17f6d41 commit. The fetched bytes match the Git blob. All four embedded standard modules in that artifact match their corresponding source files at that commit.

| Artifact | SHA-256 | Extracted HotlistProcessing |
| --- | --- | --- |
| Original Phase 3, eabdf94 | c397aea8e6494d5a3e20cd12b3ce88d8b79ef0c51a46e0ed7128da9f3b4ee4d9 | Old `input` declaration present. |
| Corrective build, aa72b8b | 5c4100a2092eed4449324ef157f9fa4074b8c327993a0039e3ac366370f5a612 | Corrected `inputRows`; source matches. |
| 0.3.1 at 17f6d41, Git and fresh GitHub API download | 0c53abad3bbd1e5ec8f58c63e0800ddacd06e628ba28ed1d0779ee9a4496b0de | Corrected `inputRows`; all four modules match. |
| Failing 0.3.1 copy, hash supplied by tester | 301165f2f959ed0fa75926c96eeb0554ea61c87bd7d34e2cb9fe8550a124baab | Tester reports stale `input` and compile failure. This copy was not available for extraction here. |

**Confirmed finding:** the failing copy's bytes differ from the committed/freshly fetched 0.3.1 artifact. The exact point at which they diverged is unresolved. A stale download, another local copy, or a later save cannot be selected as the cause without evidence. No stale-source persistence was reproduced in build_demo.py or vba_container.py. It would be incorrect to claim a proven builder-cache defect.

## Builder and previous validation review

- build_demo.py reads each `.bas` from src/vba on every invocation and passes that text to make_project. No operational workbook or template is an input.
- vba_container.py is a stateless serializer/compressor; it contains no saved HotlistProcessing source and imports no cached VBA project.
- The previous fixed tools/.vba-build.bin path was written anew, embedded on close, and removed. No such leftover file was found during inspection. This shared temporary filename was nevertheless an avoidable concurrency/stale-path risk.
- The previous standalone validator did reopen the XLSM and compare oletools-extracted module source. It was not merely comparing input hashes. The exact 17f6d41 artifact passes that comparison when rerun here.
- The confirmed release-control gaps were that the builder did not itself require post-build verification before replacing its output, and delivery lacked a byte-level receipt to identify the file being tested. Those gaps are addressed; they do not prove where the tester's different copy originated.

## 0.3.2 changes

1. Embed a new in-memory VBA project through XlsxWriter's stream API. The fixed temporary .bin path is no longer read or written.
2. Build the workbook in an isolated temporary directory, close it, then verify its actual ZIP/CFB/module bytes before atomic replacement of the output path. A failed pre-publication check exits with an error and leaves the previous artifact and receipt untouched.
3. Reopen the published final path and verify it again. Write demo/embedded-vba-manifest.json only after this succeeds.
4. Check ZIP-part uniqueness and the workbook's single VBA relationship, then use oletools project-directory extraction and direct OLE stream decompression. Both results must agree.
5. Require exactly the four standard modules and compare their full normalized source with src/vba. Normalization changes only BOM/line endings, not attributes, identifiers, spaces, or comments. Record independent extracted/source SHA-256 values.
6. Reject declarations using the earlier reserved identifiers, `input.Add`, and `Dim line As String`; require `Dim inputRows As New Collection` and the other corrected identifiers in the embedded text itself.
7. Bind the receipt to the complete workbook SHA-256, embedded project hash, version, and four module hashes. The standalone structural validator recomputes and compares the receipt.

The verifier targets the clean builder's source-only, zero-source-offset project format. It is not a general validator for a workbook that Excel has subsequently recompiled and saved. Hash the freshly downloaded file **before opening/saving it in Excel**.

## Actual validation

| Embedded standard module | Extracted/source comparison | Independent stream comparison | Stale-identifier checks |
| --- | --- | --- | --- |
| DemoUtilities | PASS | PASS | PASS: appendRows and sheetName present. |
| HistoryManagement | PASS | PASS | PASS: no banned declarations. |
| HotlistProcessing | PASS | PASS | PASS: inputRows and lineId present. |
| OfficeIntegration | PASS | PASS | PASS: exportWorkbook present. |

- **78 structural/source/model checks: PASS.** Existing 0.3.1 UX is preserved: assignment fields A:C, analyst-facing context through Q, Agent/Agent Comment P:Q, technical X:AG hidden, I2 freeze, and history-export A2 freeze source retained. Operational narrative remains qualified and synthetic.
- **10 final-binary/build regression tests: PASS.** They mutate each embedded module in an actual XLSM, inject stale declarations/calls even with matching bad expected source, omit a module, add duplicate project ZIP entries, force a post-build failure, and seed a bogus legacy cached project file. Invalid artifacts are rejected; the bogus cache is not used.
- Model expectations remain 48 assigned / 12 exceptions and loads 20 / 8 / 20. Synthetic fixture bytes are unchanged.
- Privacy checks are rerun before commit, including hidden content, metadata, embedded VBA, and confidential candidate comparisons. No operational workbook or records are build inputs or distributed assets.
- At investigation time, desktop actions had not been accepted. Subsequent hash-verified desktop testing passed compilation and the main workflow. Recovery/edge/shared-client coverage remains pending, and no timing benchmark is asserted.

## Reproduce and identify the download

```sh
python tools/build_demo.py
python tools/verify_embedded_vba.py demo/Hotlist_Workflow_Automation_Demo.xlsm --version 0.3.2
python validation/validate_demo.py
python validation/test_embedded_vba.py
```

Use the workbook and manifest from the same corrective commit, not main. Before opening the download in Excel, compare its hash to workbook_sha256 in the manifest:

```powershell
Get-FileHash .\Hotlist_Workflow_Automation_Demo.xlsm -Algorithm SHA256
```

Then open a fresh working copy and repeat Debug → Compile VBAProject. The current receipt records compile PASS only through matching hash-bound desktop evidence; new or different artifact bytes still need their own acceptance.

Final privacy rescan: 32,096 confidential candidate strings checked against repository text and extracted workbook content. The 109 matches were reviewed as generic labels/code terms (including “manifest”), standard Office metadata/theme constants, and independently generated synthetic date/amount coincidences. Employee-name, corporate-domain, and literal operational-path match counts were zero. Only example.com email addresses occur. Original confidential workbooks remain unchanged and outside the repository.
