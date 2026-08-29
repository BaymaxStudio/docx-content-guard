import tempfile
import unittest
from pathlib import Path

from docx import Document

from docx_content_guard import build_html, main, normalize_text, run_comparison


def write_docx(path: Path, paragraphs: list[str]) -> None:
    document = Document()
    for paragraph in paragraphs:
        document.add_paragraph(paragraph)
    document.save(path)


class ContentGuardTests(unittest.TestCase):
    def test_normalize_text_ignores_common_format_differences(self) -> None:
        self.assertEqual(normalize_text('“研究”　结论'), normalize_text('"研究" 结论'))

    def test_reports_non_format_content_change(self) -> None:
        with tempfile.TemporaryDirectory() as folder:
            root = Path(folder)
            before = root / 'before.docx'
            after = root / 'after.docx'
            write_docx(before, ['研究结论保持不变。'])
            write_docx(after, ['研究结论发生变化。'])

            result = run_comparison(before, after)

            self.assertEqual(len(result['unmatched']), 1)
            self.assertEqual(result['unmatched'][0]['direction'], 'changed')

    def test_report_escapes_file_names(self) -> None:
        result = {
            'doi_removed': 0,
            'ref_renumbered': 0,
            'blank_diff': 0,
            'normalized_count': 0,
            'total_paragraphs': 1,
            'unmatched': [],
        }

        report = build_html('<before>.docx', '<after>.docx', result)

        self.assertIn('&lt;before&gt;.docx', report)
        self.assertNotIn('<before>.docx', report)

    def test_cli_writes_report_without_opening_browser(self) -> None:
        with tempfile.TemporaryDirectory() as folder:
            root = Path(folder)
            before = root / 'before.docx'
            after = root / 'after.docx'
            output = root / 'report.html'
            write_docx(before, ['正文内容。'])
            write_docx(after, ['正文内容。'])

            exit_code = main([
                str(before),
                str(after),
                '--output',
                str(output),
                '--no-open',
            ])

            self.assertEqual(exit_code, 0)
            self.assertTrue(output.is_file())
            self.assertIn('当前覆盖范围内未发现内容变化', output.read_text(encoding='utf-8'))


if __name__ == '__main__':
    unittest.main()
