import tempfile
import unittest
from pathlib import Path

from openpyxl import Workbook, load_workbook

from inspect_error_structure import inspect_folder, inspect_workbook, write_report
from sohoa.domain import HEADERS


class InspectErrorStructureTests(unittest.TestCase):
    def test_reports_near_match_with_expected_and_actual_column_header(self):
        with tempfile.TemporaryDirectory() as temporary:
            path = Path(temporary) / 'short_unit_header.xlsx'
            workbook = Workbook()
            headers = list(HEADERS)
            headers[22] = 'Mã đơn vị'
            workbook.active.append(headers)
            workbook.save(path)
            workbook.close()

            findings = inspect_workbook(path)

            self.assertEqual(findings[0]['status'], 'Đọc được')
            self.assertIn(
                'Cột 23 "Mã đơn vị" được nhận diện là "Mã đơn vị (Mã moet)"',
                findings[0]['details'],
            )

    def test_short_header_without_all_identity_fields_is_not_readable(self):
        with tempfile.TemporaryDirectory() as temporary:
            path = Path(temporary) / 'shifted.xlsx'
            workbook = Workbook()
            workbook.active.append([HEADERS[0], HEADERS[4]])
            workbook.save(path)
            workbook.close()

            findings = inspect_workbook(path)

            self.assertEqual(findings[0]['status'], 'Không đạt cấu trúc')
            self.assertIn('Không nhận diện đủ tiêu đề', findings[0]['details'])

    def test_recognizes_reordered_columns_and_known_name_aliases(self):
        with tempfile.TemporaryDirectory() as temporary:
            folder = Path(temporary)
            workbook = Workbook()
            worksheet = workbook.active
            worksheet.title = 'Data'
            headers = list(HEADERS)
            headers[4] = 'Họ và tên'
            headers[5], headers[6] = headers[6], headers[5]
            worksheet.append(headers)
            workbook.save(folder / 'invalid.xlsx')
            workbook.close()

            findings = inspect_folder(folder)
            self.assertEqual(len(findings), 1)
            self.assertEqual(findings[0]['status'], 'Đọc được')
            self.assertIn('Họ và tên', findings[0]['details'])
            self.assertIn('Ngày sinh', findings[0]['details'])

    def test_follows_phase_one_data_sheet_selection_rule(self):
        with tempfile.TemporaryDirectory() as temporary:
            folder = Path(temporary)
            workbook = Workbook()
            workbook.active.title = 'Data'
            workbook.active.append(['Wrong header'])
            other = workbook.create_sheet('DuLieu')
            other.append(HEADERS)
            workbook.save(folder / 'data_first.xlsx')
            workbook.close()

            findings = inspect_folder(folder)

            self.assertEqual(len(findings), 2)
            self.assertEqual(findings[0]['status'], 'Không đạt cấu trúc')
            self.assertEqual(findings[0]['selected_by_phase1'], 'Có')
            self.assertEqual(findings[1]['status'], 'Đọc được')
            self.assertEqual(
                findings[1]['selected_by_phase1'],
                'Không (Pha 1 bỏ qua sheet này)',
            )

    def test_reports_unreadable_workbook_and_writes_summary(self):
        with tempfile.TemporaryDirectory() as temporary:
            folder = Path(temporary)
            (folder / 'broken.xlsx').write_bytes(b'not an Excel workbook')

            findings = inspect_folder(folder)
            report_path = write_report(folder, findings)

            self.assertEqual(findings[0]['status'], 'Không đọc được file')
            workbook = load_workbook(report_path, read_only=True)
            self.assertEqual(workbook.sheetnames, ['Tong_quan', 'Chi_tiet_sheet'])
            summary = list(workbook['Tong_quan'].values)
            self.assertIn('Không đọc được', summary[1][1])
            workbook.close()

    def test_accepts_canonical_header_within_first_ten_rows(self):
        with tempfile.TemporaryDirectory() as temporary:
            folder = Path(temporary)
            workbook = Workbook()
            worksheet = workbook.active
            for _ in range(9):
                worksheet.append(['Ghi chú'])
            worksheet.append(HEADERS)
            workbook.save(folder / 'valid.xlsx')
            workbook.close()

            findings = inspect_folder(folder)

            self.assertEqual(findings[0]['status'], 'Đọc được')
            self.assertEqual(findings[0]['header_row'], 10)


if __name__ == '__main__':
    unittest.main()
