import tempfile
import unittest
from pathlib import Path
from sohoa.domain import StudentRecord, MATCH, NORMALIZED, REVIEW, ONLY_SCHOOL, ONLY_SOURCE, birth
from sohoa.storage import RecordStore
from sohoa.matching import reconcile

def student(name='Nguyễn Văn An',dob='01/02/1980',year='1995',cert=''):
    v=['']*24; v[4]=name; v[6]=dob; v[12]=year; v[1]=cert
    return StudentRecord(tuple(v),'sample.xlsx','Data',6)

class MatchingTests(unittest.TestCase):
    def run_case(self,sources,schools):
        with tempfile.TemporaryDirectory() as d:
            store=RecordStore(Path(d)/'test.sqlite')
            try:
                for r in sources: store.add(r,'source')
                for r in schools: store.add(r,'school')
                return list(reconcile(store))
            finally: store.close()

    def test_normalization(self):
        result=self.run_case([student()],[student(' NGUYỄN  VĂN AN ','1/2/1980')])
        self.assertEqual([r.status for r in result],[NORMALIZED])

    def test_auxiliary_differences_ignored(self):
        result=self.run_case([student(cert='001')],[student(cert='999')])
        self.assertEqual(result[0].status,MATCH)

    def test_wrong_birth_never_matches_even_with_certificate(self):
        result=self.run_case([student(cert='001')],[student(dob='02/02/1980',cert='001')])
        self.assertEqual([r.status for r in result],[REVIEW])

    def test_duplicates_not_paired_by_order(self):
        result=self.run_case([student(),student()],[student(),student()])
        self.assertTrue(all(r.status==REVIEW for r in result))

    def test_certificate_disambiguates(self):
        result=self.run_case([student(cert='1'),student(cert='2')],[student(cert='2'),student(cert='1')])
        self.assertEqual([r.status for r in result],[MATCH,MATCH])

    def test_unrelated_rows(self):
        result=self.run_case([student()],[student('Lê Thị B','03/03/1981','1996')])
        self.assertEqual({r.status for r in result},{ONLY_SOURCE,ONLY_SCHOOL})

    def test_year_only_not_exact(self):
        result=self.run_case([student(dob='1980')],[student()])
        self.assertEqual(result[0].status,REVIEW)
        self.assertFalse(birth('31/02/1980')[1])

    def test_missing_source_birth_proposes_supplement(self):
        result=self.run_case([student(dob='')],[student()])
        self.assertEqual(result[0].status,REVIEW)
        self.assertIn('bổ sung',result[0].reason)

if __name__=='__main__': unittest.main()
