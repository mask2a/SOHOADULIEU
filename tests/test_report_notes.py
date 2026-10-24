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
            path=writer.finish()[0]
            wb=load_workbook(path)
            try:
                for sheet in ('Lech_thong_tin','Thua_Thieu'):
                    note=wb[sheet]['K2'].value
                    self.assertIn('Khớp họ tên, năm tốt nghiệp',note)
                    self.assertIn('Lệch ngày sinh',note)
                    self.assertIn('10/03/1973',note)
                    self.assertIn('10/03/1976',note)
                    self.assertEqual(wb[sheet]['A2'].value,REVIEW)
            finally: wb.close()
