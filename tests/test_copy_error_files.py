import tempfile
import unittest
from pathlib import Path

from openpyxl import Workbook

from copy_error_files import copy_reported_files
from sohoa.application import merge


class CopyErrorFilesTests(unittest.TestCase):
    def test_copies_file_named_by_a_real_phase_one_report(self):
        with tempfile.TemporaryDirectory() as temporary:
            root = Path(temporary)
            source = root / 'source'
            destination = root / 'copied'
            source.mkdir()
            destination.mkdir()
            invalid_file = source / 'unreadable.xlsx'
            workbook = Workbook()
            workbook.active.append(['Not the required header'])
            workbook.save(invalid_file)
            workbook.close()

            output = merge(source, root / 'output')
            copied, missing, existing = copy_reported_files(
                output / 'bao_cao_loi.xlsx', source, destination
            )

            self.assertEqual(copied, ['unreadable.xlsx'])
            self.assertEqual(missing, [])
            self.assertEqual(existing, [])
            self.assertEqual(
                (destination / 'unreadable.xlsx').read_bytes(),
                invalid_file.read_bytes(),
            )

    def test_copies_only_reported_workbooks_without_overwriting(self):
        with tempfile.TemporaryDirectory() as temporary:
            root = Path(temporary)
            source = root / 'source'
            destination = root / 'copied'
            source.mkdir()
            destination.mkdir()
            (source / 'bad.xlsx').write_bytes(b'bad workbook')
            (source / 'good.xlsx').write_bytes(b'good workbook')
            (destination / 'existing.xlsm').write_bytes(b'keep this')

            report_path = root / 'bao_cao_loi.xlsx'
            workbook = Workbook()
            worksheet = workbook.active
            worksheet.append(['Lỗi', 'Hướng xử lý'])
            worksheet.append([f'{source / "bad.xlsx"}: không tìm thấy tiêu đề', 'Sửa file'])
            worksheet.append(['existing.xlsm: cấu trúc không hợp lệ', 'Sửa file'])
            worksheet.append(['missing.xlsx: không đọc được', 'Sửa file'])
            worksheet.append(['Không liên quan tới file Excel', 'Kiểm tra'])
            workbook.save(report_path)
            workbook.close()

            copied, missing, existing = copy_reported_files(
                report_path, source, destination
            )

            self.assertEqual(copied, ['bad.xlsx'])
            self.assertEqual(missing, ['existing.xlsm', 'missing.xlsx'])
            self.assertEqual(existing, [])
            self.assertEqual((destination / 'bad.xlsx').read_bytes(), b'bad workbook')

    def test_does_not_overwrite_destination_file(self):
        with tempfile.TemporaryDirectory() as temporary:
            root = Path(temporary)
            source = root / 'source'
            destination = root / 'copied'
            source.mkdir()
            destination.mkdir()
            (source / 'bad.xlsx').write_bytes(b'source')
            (destination / 'bad.xlsx').write_bytes(b'keep')
            report_path = root / 'bao_cao_loi.xlsx'
            workbook = Workbook()
            workbook.active.append(['bad.xlsx: cấu trúc không hợp lệ'])
            workbook.save(report_path)
            workbook.close()

            copied, missing, existing = copy_reported_files(
                report_path, source, destination
            )

            self.assertEqual(copied, [])
            self.assertEqual(missing, [])
            self.assertEqual(existing, ['bad.xlsx'])
            self.assertEqual((destination / 'bad.xlsx').read_bytes(), b'keep')


if __name__ == '__main__':
    unittest.main()
