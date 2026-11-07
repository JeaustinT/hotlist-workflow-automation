"""Regression tests against real XLSM files with altered embedded project bytes."""
from pathlib import Path
import re
import io
import contextlib
from unittest.mock import patch
import shutil
import sys
import tempfile
import unittest
import warnings
import zipfile
from oletools.olevba import VBA_Parser
ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / 'tools'))
import build_demo
from vba_container import make_project
from verify_embedded_vba import verify, EmbeddedSourceError, EXPECTED


class FinalBinaryTests(unittest.TestCase):
    def setUp(self):
        self.temp = tempfile.TemporaryDirectory(prefix='hotlist-binary-test-')
        self.folder = Path(self.temp.name)
        self.base = ROOT / 'demo/Hotlist_Workflow_Automation_Demo.xlsm'
        parser = VBA_Parser(str(self.base))
        try:
            self.modules = {Path(name).stem: (text, name.endswith('.cls')) for _, _, name, text in parser.extract_macros()}
        finally:
            parser.close()
        self.sequence = 0

    def tearDown(self):
        self.temp.cleanup()

    def changed_artifact(self, replacements=None, mirror_source=False, omit=None):
        self.sequence += 1
        folder = self.folder / str(self.sequence)
        source = folder / 'src/vba'
        source.mkdir(parents=True)
        for path in (ROOT / 'src/vba').glob('*.bas'):
            shutil.copyfile(path, source / path.name)
        modules = dict(self.modules)
        for name, replacement in (replacements or {}).items():
            text, document = modules[name]
            text = replacement(text)
            modules[name] = (text, document)
            if mirror_source:
                (source / (name + '.bas')).write_text(text)
        if omit:
            del modules[omit]
        project = make_project(modules)
        artifact = folder / 'altered.xlsm'
        with zipfile.ZipFile(self.base) as original, zipfile.ZipFile(artifact, 'w', zipfile.ZIP_DEFLATED) as output:
            for entry in original.infolist():
                output.writestr(entry, project if entry.filename == 'xl/vbaProject.bin' else original.read(entry.filename))
        return artifact, source

    def test_final_artifact_matches_all_four_modules(self):
        result = verify(self.base, ROOT / 'src/vba', '0.3.2')
        self.assertEqual(set(result['modules']), set(EXPECTED))
        self.assertTrue(all(r['match'] and r['direct_decompression_match'] for r in result['modules'].values()))

    def test_each_module_mismatch_is_rejected(self):
        for name in EXPECTED:
            with self.subTest(module=name):
                artifact, source = self.changed_artifact({name: lambda s: s + "\r\n' Deliberately different embedded text\r\n"})
                with self.assertRaisesRegex(EmbeddedSourceError, 'source mismatch'):
                    verify(artifact, source, '0.3.2')

    def test_stale_input_declaration_rejected_even_if_expected_source_matches(self):
        artifact, source = self.changed_artifact({'HotlistProcessing': lambda s: re.sub(r'\binputRows\b', 'input', s)}, mirror_source=True)
        with self.assertRaisesRegex(EmbeddedSourceError, 'reserved identifier'):
            verify(artifact, source, '0.3.2')

    def test_stale_input_add_rejected_even_if_expected_source_matches(self):
        artifact, source = self.changed_artifact({'HotlistProcessing': lambda s: s.replace('inputRows.Add', 'input.Add')}, mirror_source=True)
        with self.assertRaisesRegex(EmbeddedSourceError, 'input.Add'):
            verify(artifact, source, '0.3.2')

    def test_stale_line_declaration_rejected_even_if_expected_source_matches(self):
        artifact, source = self.changed_artifact({'HotlistProcessing': lambda s: s + '\r\nPrivate Sub RegressionOnly()\r\nDim line As String\r\nEnd Sub\r\n'}, mirror_source=True)
        with self.assertRaisesRegex(EmbeddedSourceError, 'reserved identifier|Dim line'):
            verify(artifact, source, '0.3.2')

    def test_other_previously_fixed_identifiers_are_guarded(self):
        for module, corrected, stale in [('DemoUtilities', 'appendRows', 'append'), ('DemoUtilities', 'sheetName', 'name'), ('OfficeIntegration', 'exportWorkbook', 'output')]:
            with self.subTest(identifier=corrected):
                artifact, source = self.changed_artifact({module: lambda s: re.sub(r'\b' + corrected + r'\b', stale, s)}, mirror_source=True)
                with self.assertRaisesRegex(EmbeddedSourceError, 'reserved identifier'):
                    verify(artifact, source, '0.3.2')

    def test_duplicate_zip_project_is_rejected(self):
        artifact, source = self.changed_artifact()
        with warnings.catch_warnings():
            warnings.simplefilter('ignore', UserWarning)
            with zipfile.ZipFile(artifact, 'a') as package:
                package.writestr('xl/vbaProject.bin', package.read('xl/vbaProject.bin'))
        with self.assertRaisesRegex(EmbeddedSourceError, 'Duplicate ZIP'):
            verify(artifact, source, '0.3.2')

    def test_missing_module_is_rejected(self):
        artifact, source = self.changed_artifact(omit='HistoryManagement')
        with self.assertRaisesRegex(EmbeddedSourceError, 'Unexpected embedded standard modules'):
            verify(artifact, source, '0.3.2')


    def isolated_root(self):
        root = self.folder / 'isolated'
        (root / 'demo/sample-data').mkdir(parents=True)
        (root / 'tools').mkdir()
        shutil.copytree(ROOT / 'src/vba', root / 'src/vba')
        return root

    def test_failed_postbuild_check_does_not_replace_prior_artifact(self):
        root = self.isolated_root()
        artifact = root / 'demo/Hotlist_Workflow_Automation_Demo.xlsm'
        artifact.write_bytes(b'PRIOR ARTIFACT MUST NOT BE REPLACED')
        receipt = root / 'demo/embedded-vba-manifest.json'
        receipt.write_text('prior receipt')
        with patch.object(build_demo, 'ROOT', root), patch.object(build_demo, 'verify', side_effect=EmbeddedSourceError('injected post-build failure')):
            with self.assertRaisesRegex(EmbeddedSourceError, 'injected post-build failure'):
                build_demo.build()
        self.assertEqual(artifact.read_bytes(), b'PRIOR ARTIFACT MUST NOT BE REPLACED')
        self.assertEqual(receipt.read_text(), 'prior receipt')
        self.assertEqual(list((root / 'demo').glob('.hotlist-build-*')), [])

    def test_desktop_acceptance_does_not_transfer_to_different_binary(self):
        artifact, source = self.changed_artifact({'HistoryManagement': lambda s: s + "\r\n' Different validated source fixture\r\n"}, mirror_source=True)
        # Keep evidence beside the fixture repository's src directory.
        actual_evidence_dir = source.resolve().parents[1] / 'validation'
        actual_evidence_dir.mkdir(exist_ok=True)
        shutil.copyfile(ROOT / 'validation/desktop-acceptance.json', actual_evidence_dir / 'desktop-acceptance.json')
        result = verify(artifact, source, '0.3.2')
        self.assertEqual(result['excel_desktop_compile'], 'NOT VERIFIED FOR THIS ARTIFACT')
        accepted = verify(self.base, ROOT / 'src/vba', '0.3.2')
        self.assertEqual(accepted['excel_desktop_compile'], 'PASS')

    def test_clean_build_ignores_legacy_cached_project_file(self):
        root = self.isolated_root()
        cached = root / 'tools/.vba-build.bin'
        cached.write_bytes(b'INVALID STALE CACHED PROJECT')
        with patch.object(build_demo, 'ROOT', root), contextlib.redirect_stdout(io.StringIO()):
            build_demo.build()
        result = verify(root / 'demo/Hotlist_Workflow_Automation_Demo.xlsm', root / 'src/vba', '0.3.2')
        self.assertTrue(all(m['match'] for m in result['modules'].values()))
        self.assertEqual(cached.read_bytes(), b'INVALID STALE CACHED PROJECT')
        self.assertEqual(list((root / 'demo').glob('.hotlist-build-*')), [])


if __name__ == '__main__':
    unittest.main(verbosity=2)
