import csv
import tempfile
import unittest
from pathlib import Path
from openpyxl import Workbook, load_workbook
from sohoa.domain import HEADERS, AppError
from sohoa.application import merge, compare, compare_batch, JobContext
from sohoa.report import ReportWriter
from sohoa.domain import MatchResult, REVIEW, StudentRecord

def write_input(path, rows, header_row=1, headers=HEADERS):
    wb=Workbook(); ws=wb.active; ws.title='Data'
    for _ in range(header_row-1): ws.append(['Ghi chú'])
    ws.append(headers)
    if header_row==4: ws.append([f'({i})' for i in range(1,25)])
    for row in rows: ws.append(row)
    wb.save(path); wb.close()

def row(name='Nguyễn Văn An', dob='01/02/1980', year='1995', council='THCS A'):
    v=['']*24; v[0]='1'; v[4]=name; v[6]=dob; v[12]=year; v[13]=council
    return v

class WorkflowTests(unittest.TestCase):
    def test_batch_compare_matches_years_and_names_reports_by_year(self):
        import json
        with tempfile.TemporaryDirectory() as tmp:
            root=Path(tmp); sources=root/'sources'; schools=root/'schools'
            sources.mkdir(); schools.mkdir()
            write_input(sources/'THCS.CQ.1987_output.xlsx',
                        [row(name='Học sinh 1987',year='1987')],4)
            write_input(sources/'THCS.CQ.1989_output.xlsx',
                        [row(name='Học sinh 1989',year='1989')],4)
            write_input(sources/'THCS.CQ.1990_output.xlsx',
                        [row(name='Không có trường gửi',year='1990')],4)
            write_input(schools/'1987.xlsx',
                        [row(name='Học sinh 1987',year='1987')])
            write_input(schools/'1988.xlsx',
                        [row(name='Không có nguồn',year='1988')])
            write_input(schools/'1989.xlsx',
                        [row(name='Học sinh 1989',year='1989')])

            output=compare_batch(sources,schools,root/'out')

            self.assertEqual({path.name for path in output.glob('*_Check.xlsx')},
                             {'1987_Check.xlsx','1989_Check.xlsx'})
            manifest=json.loads((output/'pham_vi.json').read_text(encoding='utf-8'))
            self.assertEqual(manifest['skipped_school_years'],['1988'])
            self.assertEqual(manifest['skipped_source_years'],['1990'])
            for year in ('1987','1989'):
                wb=load_workbook(output/f'{year}_Check.xlsx',read_only=True)
                self.assertEqual(wb['Tong_hop']['B2'].value,year)
                self.assertEqual(wb['Tong_hop']['B7'].value,1)
                wb.close()

    def test_merge_reads_only_first_data_sheet(self):
        from sohoa.excel import records
        with tempfile.TemporaryDirectory() as tmp:
            root=Path(tmp); data=root/'data'; data.mkdir()
            source=data/'source.xlsx'
            write_input(source,[row(name='Học sinh trong Data')])
            wb=load_workbook(source)
            wb.active.title='temp'
            wb.active.title='data'
            ws=wb.create_sheet('trượt')
            ws.append(HEADERS); ws.append(row(name='Học sinh trong trượt'))
            wb.save(source); wb.close()

            parsed=list(records(source,1,canonical=True))
            self.assertEqual(len(parsed),1)
            self.assertEqual(parsed[0].values[4],'Học sinh trong Data')

            output=merge(data,root/'out')
            wb=load_workbook(output/'1995.xlsx')
            self.assertEqual(wb['Data'].max_row,2)
            self.assertEqual(wb['Data']['E2'].value,'Học sinh trong Data')
            wb.close()

    def test_year_named_sheet_is_read_and_lineage_preserves_name(self):
        from sohoa.excel import records, inspect
        with tempfile.TemporaryDirectory() as tmp:
            source=Path(tmp)/'THCS.CQ.1987.xlsx'
            write_input(source,[row(year='1987')],4)
            wb=load_workbook(source); wb['Data'].title='1987'; wb.save(source); wb.close()
            self.assertEqual(inspect(source,4)[:24],list(HEADERS))
            parsed=list(records(source,4,canonical=True))
            self.assertEqual(len(parsed),1)
            self.assertEqual(parsed[0].sheet,'1987')
            self.assertEqual(parsed[0].row,6)

    def test_merge_accepts_arbitrary_excel_names_and_ignores_lock_files(self):
        import json
        with tempfile.TemporaryDirectory() as tmp:
            root=Path(tmp); data=root/'other_folder'; data.mkdir()
            write_input(data/'du lieu truong A.XLSX',[row(year='1987')],4)
            write_input(data/'ban bo sung.xlsx',[row(year='1988')],4)
            (data/'~$dang_mo.xlsx').write_text('Excel lock file',encoding='utf-8')
            (data/'notes.txt').write_text('ignore',encoding='utf-8')
            (data/'not_a_file.xlsx').mkdir()
            alias=root/'schools.csv'; alias.write_text('alias,school_id,school_name\n',encoding='utf-8')
            output=merge(data,root/'out',alias)
            wb=load_workbook(output/'1987.xlsx')
            self.assertEqual(wb['Data'].max_row,2)
            self.assertTrue((output/'1988.xlsx').exists())
            wb.close()
            self.assertEqual(json.loads((output/'pham_vi.json').read_text(encoding='utf-8'))['years_read'],[1987,1988])
            self.assertFalse((output/'bao_cao_loi.xlsx').exists())

    def test_split_folder_ids_increase_without_overwrite(self):
        from sohoa.application import run_folder
        with tempfile.TemporaryDirectory() as tmp:
            root=Path(tmp)/'output'
            first=run_folder(root,'Gop_tach')
            self.assertEqual(first.name,'TACH_NAM_1')
            marker=first/'keep.txt'; marker.write_text('keep',encoding='utf-8')
            (root/'TACH_NAM_7').mkdir()
            (root/'TACH_NAM_invalid').mkdir()
            (root/'TACH_TRUONG.99').mkdir()
            second=run_folder(root,'Gop_tach')
            self.assertEqual(second.name,'TACH_NAM_8')
            self.assertEqual(marker.read_text(encoding='utf-8'),'keep')

    def test_three_schools_without_aliases_and_extra_note_on_number_row(self):
        with tempfile.TemporaryDirectory() as tmp:
            root=Path(tmp); data=root/'data'; data.mkdir()
            source=data/'THCS.CQ.1987_output.xlsx'
            write_input(source,[row(council=name,year='1987') for name in ('Phú Xuyên','Bắc Sơn','Cây Thị')],4)
            wb=load_workbook(source); wb['Data']['Y5']='Ghi chú ngoài bảng'; wb['Data']['X5']='(24'; wb.save(source); wb.close()
            alias=root/'schools.csv'; alias.write_text('alias,school_id,school_name,from_year,to_year\n',encoding='utf-8')
            output=merge(data,root/'out',alias)
            self.assertEqual({p.name for p in output.glob('*.xlsx')},{'1987.xlsx'})
            for path in output.glob('*.xlsx'):
                wb=load_workbook(path); self.assertEqual(wb['Data'].max_row,4); wb.close()

    def test_merge_compare_lineage_and_highlight(self):
        with tempfile.TemporaryDirectory() as tmp:
            root=Path(tmp); data=root/'data'; data.mkdir()
            write_input(data/'THCS.CQ.1995_output.xlsx',[row()],4)
            alias=root/'schools.csv'; alias.write_text('alias,school_id,school_name,from_year,to_year\nTHCS A,A,Trường A,1987,2005\n',encoding='utf-8')
            merged=merge(data,root/'out',alias)
            source=merged/'1995.xlsx'
            wb=load_workbook(source); self.assertEqual(wb.sheetnames,['Data','THONGKE']); self.assertEqual(wb['Data'].max_row,2); wb.close()
            school=root/'school.xlsx'; write_input(school,[row(dob='02/02/1980')])
            report=compare([source],school,root/'out',years={1995})
            wb=load_workbook(report/'1995_Check.xlsx')
            self.assertEqual(wb.sheetnames,['Tong_hop','Khop','Chi_co_trong_nguon_chuan',
                                            'Lech_thong_tin','Truong_khong_co_trong_nguon',
                                            'Can_xac_minh'])
            self.assertEqual(wb['Tong_hop'].max_row,14)
            self.assertEqual(wb['Khop'].max_row,1)
            self.assertEqual(wb['Lech_thong_tin'].max_row,2)
            self.assertEqual(wb['Lech_thong_tin']['Y2'].value,REVIEW)
            self.assertEqual(wb['Lech_thong_tin']['E2'].value,'Nguyễn Văn An')
            self.assertIn('lệch ngày sinh',wb['Lech_thong_tin']['Z2'].value)
            self.assertIn('dòng 6',wb['Lech_thong_tin']['AD2'].value)
            self.assertEqual(wb['Truong_khong_co_trong_nguon'].max_row,2)
            self.assertIn('Dữ liệu trường gửi lên mà không có trong dữ liệu nguồn chuẩn',
                          wb['Truong_khong_co_trong_nguon']['A1'].value)
            self.assertEqual(wb['Can_xac_minh'].max_row,2)
            wb.close()

    def test_compare_single_source_with_header_on_row_four_to_split_school_file(self):
        with tempfile.TemporaryDirectory() as tmp:
            root=Path(tmp); source=root/'THCS.CQ.1987_output.xlsx'; school=root/'1987.xlsx'
            write_input(source,[row(year='1987')],4)
            wb=Workbook(); ws=wb.active; ws.title='Data'
            ws.append((*HEADERS,'Nguồn dữ liệu'))
            ws.append((*row(year='1987'),'Tệp: source.xlsx | Sheet: Data | Dòng: 6'))
            stats=wb.create_sheet('THONGKE'); stats.append(['STT','TÊN TRƯỜNG','SỐ LƯỢNG'])
            wb.save(school); wb.close()
            report=compare([source],school,root/'out',years={1987})
            wb=load_workbook(report/'THCS.CQ.1987_output_Check.xlsx')
            self.assertEqual(wb['Tong_hop'].max_row,14)
            self.assertEqual(wb['Khop'].max_row,2)
            self.assertEqual(wb['Chi_co_trong_nguon_chuan'].max_row,2)
            self.assertEqual(wb['Lech_thong_tin'].max_row,1)
            self.assertEqual(wb['Truong_khong_co_trong_nguon'].max_row,2)
            self.assertEqual(wb['Can_xac_minh'].max_row,1)
            self.assertEqual(wb['Khop']['A2'].value,'Khớp hoàn toàn')
            headers=[cell.value for cell in wb['Khop'][1]]
            school_location=wb['Khop'].cell(2,headers.index('Vị trí file trường')+1).value
            self.assertEqual(school_location,'1987.xlsx | Data | dòng 2')
            self.assertNotIn('C:\\Users\\Admin',school_location)
            wb.close()

    def test_compare_rejects_multiple_source_files(self):
        with tempfile.TemporaryDirectory() as tmp:
            root=Path(tmp); source=root/'source.xlsx'; school=root/'school.xlsx'
            write_input(source,[row()]); write_input(school,[row()])
            with self.assertRaisesRegex(AppError,'đúng một file nguồn chuẩn'):
                compare([source,source],school,root/'out',years={1995})

    def test_source_only_records_are_in_separate_sheet(self):
        with tempfile.TemporaryDirectory() as tmp:
            root=Path(tmp); source=root/'source.xlsx'; school=root/'school.xlsx'
            write_input(source,[row()])
            write_input(school,[row(name='Lê Thị B',dob='03/04/1982',year='1996')])
            report=compare([source],school,root/'out')
            wb=load_workbook(report/'source_Check.xlsx')
            self.assertEqual(wb['Khop'].max_row,1)
            self.assertEqual(wb['Chi_co_trong_nguon_chuan'].max_row,3)
            self.assertEqual(wb['Chi_co_trong_nguon_chuan']['A3'].value,'1')
            self.assertEqual(wb['Chi_co_trong_nguon_chuan']['E3'].value,'Nguyễn Văn An')
            self.assertEqual(wb['Truong_khong_co_trong_nguon'].max_row,3)
            self.assertEqual(wb['Truong_khong_co_trong_nguon']['E3'].value,'Lê Thị B')
            self.assertEqual(wb['Lech_thong_tin'].max_row,1)
            wb.close()

    def test_exact_identity_among_candidates_counts_as_match_and_lists_short_locations(self):
        with tempfile.TemporaryDirectory() as tmp:
            root=Path(tmp); source=root/'source.xlsx'; school=root/'school.xlsx'
            write_input(source,[row(name='Lê Thị Hồng An',dob='24/11/1973',year='1987')])
            write_input(school,[
                row(name='LÊ THỊ HỒNG AN',dob='24/11/1973',year='1987'),
                row(name='LÊ THỊ HỒNG AN',dob='24/11/1973',year='1987'),
                row(name='Vũ Đoàn Trung',dob='24/11/1973',year='1987'),
                row(name='Nguyễn Văn Khác',dob='24/11/1973',year='1987'),
            ])
            report=compare([source],school,root/'out',years={1987})
            wb=load_workbook(report/'source_Check.xlsx')
            try:
                self.assertEqual(wb['Khop'].max_row,2)
                headers=[cell.value for cell in wb['Khop'][1]]
                note_column=headers.index('Ghi chú đối soát')+1
                note=wb['Khop'].cell(2,note_column).value
                self.assertIn('Có 2 hồ sơ file trường khớp thông tin',note)
                self.assertIn('school.xlsx | Data | dòng 2',note)
                self.assertIn('school.xlsx | Data | dòng 3',note)
                self.assertNotIn(str(root),note)
                self.assertFalse(any(
                    str(root).casefold() in str(cell.value).casefold()
                    for sheet in wb.worksheets for row_cells in sheet.iter_rows()
                    for cell in row_cells if cell.value is not None
                ))
                self.assertEqual(wb['Lech_thong_tin'].max_row,1)
                self.assertEqual(wb['Can_xac_minh'].max_row,3)
                self.assertEqual(wb['Tong_hop']['B7'].value,1)
                self.assertEqual(wb['Tong_hop']['B8'].value,2)
            finally:
                wb.close()

    def test_report_lists_duplicate_candidate_sources_once_each(self):
        with tempfile.TemporaryDirectory() as tmp:
            root=Path(tmp); source=root/'1987.xlsx'; school=root/'school.xlsx'
            write_input(source,[row(),row()])
            write_input(school,[row()])
            report=compare([source],school,root/'out',years={1995})
            wb=load_workbook(report/'1987_Check.xlsx')
            self.assertEqual(wb['Tong_hop']['B3'].value,2)
            self.assertEqual(wb['Tong_hop']['B4'].value,1)
            self.assertEqual(wb['Tong_hop']['B10'].value,2)
            self.assertEqual(wb['Lech_thong_tin'].max_row,3)
            self.assertEqual(wb['Chi_co_trong_nguon_chuan'].max_row,2)
            self.assertEqual(wb['Can_xac_minh'].max_row,2)
            self.assertEqual(wb['Truong_khong_co_trong_nguon'].max_row,2)
            wb.close()

    def test_compare_has_no_year_limit_and_blank_scope_includes_all_years(self):
        import json
        with tempfile.TemporaryDirectory() as tmp:
            root=Path(tmp); source=root/'source.xlsx'; school=root/'school.xlsx'
            write_input(source,[row(year='1983'),row(name='Lê Thị B',dob='03/04/1982',year='2026')])
            write_input(school,[row(year='1983'),row(name='Lê Thị B',dob='03/04/1982',year='2026')])
            report=compare([source],school,root/'out')
            metadata=json.loads((report/'pham_vi.json').read_text(encoding='utf-8'))
            self.assertEqual(metadata['years'],'Tất cả năm trong file')
            self.assertEqual(metadata['counts'],{'source':2,'school':2})
            wb=load_workbook(report/'source_Check.xlsx')
            self.assertEqual(wb['Tong_hop']['B7'].value,2)
            self.assertEqual(wb['Khop'].max_row,3)
            self.assertEqual(wb['Chi_co_trong_nguon_chuan'].max_row,2)
            self.assertEqual(wb['Lech_thong_tin'].max_row,1)
            wb.close()

    def test_merge_sorts_by_school_attended_then_sequence(self):
        with tempfile.TemporaryDirectory() as tmp:
            root=Path(tmp); data=root/'data'; data.mkdir()
            rows=[]
            for school,seq,council in [('THCS B','1','A'),('thcs a','2','Z'),('THCS A','1','Z')]:
                values=row(year='1984',council=council)
                values[11]=school; values[0]=seq; values[21]='Ghi chú gốc'; rows.append(values)
            write_input(data/'a.xlsx',rows)
            output=merge(data,root/'out')
            self.assertEqual(output.name,'TACH_NAM_1')
            wb=load_workbook(output/'1984.xlsx')
            self.assertEqual(wb.sheetnames,['Data','THONGKE'])
            self.assertEqual(list(wb['THONGKE'].values),[
                ('STT','TÊN TRƯỜNG','SỐ LƯỢNG'),
                (1,'THCS A',2),(2,'THCS B',1),('TỔNG','Tổng số trường: 2',3)])
            content=list(wb['Data'].values); exported=content[1:]; wb.close()
            self.assertEqual(content[0],(*HEADERS,'Nguồn dữ liệu'))
            self.assertEqual([(v[11],v[0]) for v in exported],[('THCS A','1'),('thcs a','2'),('THCS B','1')])
            self.assertEqual([v[24] for v in exported],[f'Tệp: a.xlsx | Sheet: Data | Dòng: {r}' for r in (4,3,2)])
            self.assertTrue(all(v[21]=='Ghi chú gốc' for v in exported))

    def test_merge_template_sheets_and_invalid_years(self):
        import sqlite3
        with tempfile.TemporaryDirectory() as tmp:
            root=Path(tmp); data=root/'data'; data.mkdir()
            write_input(data/'a.xlsx',[row(year='1984'),row(year=''),row(year='bad')])
            wb=load_workbook(data/'a.xlsx'); wb.create_sheet('HuongDan').append(['H??ng d?n']); wb.create_sheet('DanhMuc'); wb.save(data/'a.xlsx'); wb.close()
            write_input(data/'b.xlsx',[row(year='1984'),row(year='2005')],4)
            output=merge(data,root/'out')
            wb=load_workbook(output/'1984.xlsx'); self.assertEqual(wb['Data'].max_row,3); wb.close()
            self.assertTrue((output/'2005.xlsx').exists())
            wb=load_workbook(output/'Chua_xac_dinh_nam.xlsx'); self.assertEqual(wb['Data'].max_row,3); wb.close()
            with sqlite3.connect(output/'truy_vet.sqlite') as db:
                self.assertEqual(db.execute('SELECT COUNT(*) FROM exports').fetchone()[0],5)
            db.close()
            self.assertTrue((output/'bao_cao_loi.xlsx').exists())

    def test_mapping_and_structure_rejection(self):
        with tempfile.TemporaryDirectory() as tmp:
            root=Path(tmp); source=root/'source.xlsx'; school=root/'school.xlsx'
            write_input(source,[row()]); write_input(school,[['Nguyễn Văn An','01/02/1980','1995']],headers=['Tên','Sinh','Năm'])
            report=compare([source],school,root/'out',mapping={4:0,6:1,12:2},years={1995})
            self.assertTrue((report/'source_Check.xlsx').exists())
            with self.assertRaises(AppError): compare([source],school,root/'out',mapping={4:0},years={1995})
            self.assertTrue(list((root/'out').glob('*/bao_cao_loi.xlsx')))

    def test_report_split_and_formula_text(self):
        with tempfile.TemporaryDirectory() as tmp:
            v=row(name='=SUM(1,2)'); record=StudentRecord(tuple(v),'file','Data',2)
            writer=ReportWriter(tmp,'1995',3,0,row_limit=4)
            for _ in range(3): writer.add(MatchResult(REVIEW,source=record))
            paths=writer.finish(); self.assertEqual(len(paths),2)
            wb=load_workbook(paths[0]); self.assertEqual(wb['Lech_thong_tin']['E2'].data_type,'s'); wb.close()

    def test_cancel_records_incomplete_state(self):
        from sohoa.domain import Cancelled
        with tempfile.TemporaryDirectory() as tmp:
            root=Path(tmp); source=root/'source.xlsx'; write_input(source,[row()])
            with self.assertRaises(Cancelled):
                compare([source],source,root/'out',context=JobContext(cancelled=lambda:True))
            state=next((root/'out').glob('*/TRANG_THAI.txt')).read_text(encoding='utf-8')
            self.assertIn('CHƯA HOÀN TẤT',state)

if __name__=='__main__': unittest.main()
