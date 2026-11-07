# Clean build tools

These tools are included to reproduce and inspect the synthetic artifacts. They never read confidential operational files.

Use Python 3.11 or later in an isolated environment:

```sh
python -m pip install -r tools/requirements.txt
python tools/build_demo.py
python tools/verify_embedded_vba.py demo/Hotlist_Workflow_Automation_Demo.xlsm --version 0.3.2
python validation/validate_demo.py
python validation/test_embedded_vba.py
```

The builder creates the clean `.xlsm`, a typed-date sample `.xlsx`, CSV fixtures, expected outputs, and availability profiles. It embeds the four exact `.bas` files under src/vba plus generated, empty document modules. All workbook content is synthetic.

vba_container.py writes the CFB/MS-OVBA container directly. It generates registered Excel reference metadata and a source-only VBA project, without importing any operational or third-party VBA binary. There are no compiled performance caches. Office must compile the project on first open; parsing/extraction does not establish Excel compatibility. Complete the desktop checklist before calling it runtime validated.

The reference model in build_demo.py is an executable specification for synthetic expected results. It does not interpret or run VBA. The validation script checks domain assertions, structure, source alignment, and packaging. Neither tool measures runtime performance.

Do not edit the embedded workbook source independently. Edit the reviewed `.bas` files, rebuild, validate, and perform the desktop pass. Keep a separately saved working copy when testing in Excel. Python ZIP timestamps may differ across builds even when extracted content is unchanged.

The builder now embeds a fresh in-memory project and validates the closed staged XLSM before atomic replacement. It reopens the published XLSM again and writes demo/embedded-vba-manifest.json with workbook/project hashes and module-by-module extracted/source results. A legacy tools/.vba-build.bin is not an input. Temporary build directories are isolated and removed on failure.

verify_embedded_vba.py follows the actual workbook VBA relationship, extracts/decompresses its project sources with oletools, cross-checks direct OLE module streams, rejects duplicate parts/modules and stale identifiers, and compares normalized full text with src/vba. The normalization preserves all source except BOM/line-ending differences. It targets this clean builder's zero-offset source-only format, not a working copy recompiled/saved by Excel. None of these checks executes VBA.

The binary regression suite corrupts real XLSM artifacts and proves that mismatches and old identifiers are rejected. It also verifies that failed build checks cannot replace the prior artifact, and that a bogus legacy cache file is ignored. Review validation/embedded-vba-investigation.md for the 0.3.1 tester report and independently verified artifact hashes.

## Desktop acceptance metadata

validation/desktop-acceptance.json records the tester's real Windows desktop Excel compile/main-workflow PASS for the exact accepted 0.3.2 hash. The verifier reads that evidence only after extracting/comparing the actual XLSM. It reports compile PASS only when the evidence hash/version matches; a different binary receives NOT VERIFIED FOR THIS ARTIFACT. This is an external acceptance record, not a claim that Python compiled VBA. The current manifest was refreshed without rebuilding or changing the accepted workbook.

The original 10 binary/build tests remain, plus a regression proving desktop PASS cannot transfer to modified artifact bytes (11 total). The build/validation tools are retained as engineering evidence. No captured runtime working copy or preview PNG is a build input or release artifact.
