import tempfile
import unittest
from pathlib import Path
from openpyxl import load_workbook
from sohoa.domain import StudentRecord, MatchResult, REVIEW
from sohoa.report import ReportWriter

class ReportNoteTests(unittest.TestCase):
    def test_candidate_note_identifies_exact_birth_difference(self):
        def record(name,dob):
            values=['']*24; values[4]=name; values[6]=dob; values[12]='1987'
            return StudentRecord(tuple(values),'sample.xlsx','Data',11)
        source=record('Cao Hồng Đức','10/03/1973')
        school=record('CAO HỒNG ĐỨC','10/03/1976')
        with tempfile.TemporaryDirectory() as tmp:
            writer=ReportWriter(tmp,'1987',1,1)
            writer.add(MatchResult(REVIEW,school=school,candidates=[source],reason='Chưa xác định được cặp; không tính khớp.'))
            writer.add(MatchResult(REVIEW,source=source,candidates=[school],reason='Chưa ghép chính xác; có 1 hồ sơ trường cần xác minh.'))
            path=writer.finish()[0]
            wb=load_workbook(path)
            try:
                self.assertEqual(wb['Can_xac_minh']['Y2'].value,'1')
                candidate=wb['Can_xac_minh']['Z2'].value
                self.assertIn('Cao Hồng Đức',candidate)
                self.assertIn('10/03/1973',candidate)
                self.assertEqual(wb['Can_xac_minh']['AA2'].value,'Chưa xác định được cặp; không tính khớp.')
                difference=wb['Lech_thong_tin']['Z2'].value
                self.assertIn('lệch ngày sinh',difference)
                self.assertNotIn('lệch họ tên',difference)
            finally: wb.close()
