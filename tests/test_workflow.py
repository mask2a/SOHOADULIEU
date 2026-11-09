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
    def test_merge_maps_reordered_headers_and_common_aliases(self):
        from sohoa.excel import merge_records
        with tempfile.TemporaryDirectory() as tmp:
            source=Path(tmp)/'school.xlsx'
            wb=Workbook(); ws=wb.active; ws.title='Data'
            ws.append(['Năm TN','Ngày sinh','Họ và tên','Tên trường đã học','Xếp loại tốt nghiệp'])
            ws.append(['1995','01/02/1980','Nguyễn Văn An','THCS A','Khá'])
            wb.save(source); wb.close()

            parsed=list(merge_records(source))

            self.assertEqual(len(parsed),1)
            self.assertEqual(parsed[0].values[4],'Nguyễn Văn An')
            self.assertEqual(parsed[0].values[6],'01/02/1980')
            self.assertEqual(parsed[0].values[11],'THCS A')
            self.assertEqual(parsed[0].values[12],'1995')
            self.assertEqual(parsed[0].values[20],'Khá')

    def test_merge_preserves_unmapped_column_values_in_output(self):
        with tempfile.TemporaryDirectory() as tmp:
            root=Path(tmp); data=root/'data'; data.mkdir()
            source=data/'school.xlsx'
            wb=Workbook(); ws=wb.active; ws.title='Data'
            ws.append(['Năm TN','Ngày sinh','Họ và tên','Ghi chú lạ'])
            ws.append(['1995','01/02/1980','Nguyễn Văn An','Giữ nguyên nội dung'])
            wb.save(source); wb.close()

            output=merge(data,root/'out')

            wb=load_workbook(output/'1995.xlsx',read_only=True)
            rows=list(wb['Data'].values)
            wb.close()
            self.assertEqual(rows[0],(*HEADERS,'Nguồn dữ liệu','Dữ liệu cột nguồn chưa ánh xạ'))
            self.assertEqual(rows[1][25],'Ghi chú lạ: Giữ nguyên nội dung')

    def test_merge_reports_file_stages_and_separate_percentage_updates(self):
        from sohoa.application import JobContext
        with tempfile.TemporaryDirectory() as tmp:
            root=Path(tmp); data=root/'data'; data.mkdir()
            write_input(data/'school.xlsx',[row()])
            events=[]

            merge(data,root/'out',context=JobContext(events.append))

            messages=[event.message for event in events]
            self.assertIn('Đang đọc tệp: school.xlsx',messages)
            self.assertIn('Bắt đầu giai đoạn xuất dữ liệu theo năm.',messages)
            self.assertIn('Đang xuất năm tốt nghiệp: 1995',messages)
            self.assertTrue(any(
                not event.log_message and event.current==event.total==100000
                for event in events
            ))
            percentages=[
                event.current for event in events
                if not event.log_message and event.total==100000
            ]
            self.assertEqual(percentages,sorted(percentages))

    def test_merge_preserves_rows_when_optional_headers_are_missing(self):
        from sohoa.excel import merge_records
        with tempfile.TemporaryDirectory() as tmp:
            source=Path(tmp)/'school.xlsx'
            wb=Workbook(); ws=wb.active; ws.title='Data'
            ws.append(['Họ và tên','Ngày sinh','Năm TN'])
            ws.append(['Nguyễn Văn An','01/02/1980','1995'])
            wb.save(source); wb.close()

            parsed=list(merge_records(source))

            self.assertEqual(len(parsed),1)
            self.assertEqual(parsed[0].values[4],'Nguyễn Văn An')
            self.assertEqual(parsed[0].values[12],'1995')
            self.assertEqual(parsed[0].values[20],'')

    def test_merge_infers_blank_header_from_neighboring_canonical_columns(self):
        from sohoa.excel import merge_records
        with tempfile.TemporaryDirectory() as tmp:
            source=Path(tmp)/'school.xlsx'
            headers=list(HEADERS)
            headers[0]=''
            headers[7]=''
            values=row()
            values[7]='Nữ'
            wb=Workbook(); ws=wb.active; ws.title='Data'
            ws.append(headers)
            ws.append(values)
            wb.save(source); wb.close()

            parsed=list(merge_records(source))

            self.assertEqual(len(parsed),1)
            self.assertEqual(parsed[0].values[0],'1')
            self.assertEqual(parsed[0].values[7],'Nữ')

    def test_merge_uses_populated_column_when_optional_header_is_duplicated(self):
        from sohoa.excel import merge_records
        with tempfile.TemporaryDirectory() as tmp:
            source=Path(tmp)/'school.xlsx'
            wb=Workbook(); ws=wb.active; ws.title='Data'
            ws.append(['STT','Stt','Năm TN','Họ và tên','Ngày sinh'])
            ws.append([None,'7','1995','Nguyễn Văn An','01/02/1980'])
            wb.save(source); wb.close()
            logs=[]

            parsed=list(merge_records(source,logs.append))

            self.assertEqual(len(parsed),1)
            self.assertEqual(parsed[0].values[0],'7')
            self.assertTrue(any('dùng cột 2 có 1 ô có dữ liệu' in line for line in logs))

    def test_merge_aggregates_numeric_code_warnings_and_reports_read_progress(self):
        from sohoa.excel import merge_records
        with tempfile.TemporaryDirectory() as tmp:
            source=Path(tmp)/'school.xlsx'
            wb=Workbook(); ws=wb.active; ws.title='Data'
            ws.append(['Số hiệu bằng','Họ và tên','Ngày sinh','Năm TN'])
            for index in range(5001):
                ws.append([index+1,f'Học sinh {index}','01/02/1980','1995'])
            wb.save(source); wb.close()
            logs=[]; progress=[]

            parsed=list(merge_records(source,logs.append,lambda current,total: progress.append((current,total))))

            self.assertEqual(len(parsed),5001)
            warnings=[line for line in logs if 'Cần kiểm tra mã dạng số' in line]
            self.assertEqual(len(warnings),1)
            self.assertIn('5001 ô',warnings[0])
            self.assertIn((0,5001),progress)
            self.assertIn((5000,5001),progress)
            self.assertEqual(progress[-1],(5001,5001))

    def test_merge_summarizes_large_unmapped_column_diagnostics(self):
        from sohoa.excel import merge_records
        with tempfile.TemporaryDirectory() as tmp:
            source=Path(tmp)/'wide.xlsx'
            wb=Workbook(); ws=wb.active; ws.title='Data'
            ws.append([*HEADERS,*(['']*1000)])
            ws.append([*row(),*(['Giá trị bổ sung']*1000)])
            wb.save(source); wb.close()
            logs=[]

            parsed=list(merge_records(source,logs.append))

            warnings=[line for line in logs if 'có dữ liệu ở cột chưa nhận diện' in line]
            self.assertEqual(len(parsed),1)
            self.assertEqual(len(parsed[0].unmapped),1000)
            self.assertEqual(len(warnings),1)
            self.assertIn('cột 25-1024 (không có tiêu đề): 1000 ô trên 1000 cột',warnings[0])
            self.assertLess(len(warnings[0]),500)

    def test_merge_ignores_formatting_extent_but_preserves_far_right_data(self):
        from openpyxl.styles import PatternFill
        from sohoa.excel import inspect, merge_records
        with tempfile.TemporaryDirectory() as tmp:
            source=Path(tmp)/'wide.xlsx'
            wb=Workbook(); ws=wb.active; ws.title='Data'
            ws.append((*HEADERS,'Cột bổ sung'))
            values=row()
            values[4]='Học sinh thử'
            ws.append((*values,'Giữ dữ liệu cột bổ sung'))
            ws.cell(row=2,column=16384,value='Giữ dữ liệu rất xa bên phải')
            ws.cell(row=20,column=16384).fill=PatternFill('solid',fgColor='FFFFFF')
            wb.save(source); wb.close()
            wb=load_workbook(source,read_only=True)
            self.assertEqual(wb['Data'].max_column,16384)
            wb.close()

            headers=inspect(source,1,'Data')
            parsed=list(merge_records(source))

            self.assertEqual(len(headers),25)
            self.assertEqual(len(parsed),1)
            self.assertIn(('Cột bổ sung','Giữ dữ liệu cột bổ sung'),parsed[0].unmapped)
            self.assertIn(
                ('Cột nguồn 16384 (không có tiêu đề)','Giữ dữ liệu rất xa bên phải'),
                parsed[0].unmapped,
            )

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
                        [row(name='HỌC SINH 1989',year='1989')])

            output=compare_batch(sources,schools,root/'out')

            self.assertEqual({path.name for path in output.glob('*_Check.xlsx')},
                             {'1987_Check.xlsx','1989_Check.xlsx'})
            manifest=json.loads((output/'pham_vi.json').read_text(encoding='utf-8'))
            self.assertEqual(manifest['skipped_school_years'],['1988'])
            self.assertEqual(manifest['skipped_source_years'],['1990'])
            for year in ('1987','1989'):
                wb=load_workbook(output/f'{year}_Check.xlsx',read_only=True)
                try:
                    self.assertEqual(wb['TONGHOP']['B2'].value,year)
                    self.assertEqual(wb['TONGHOP']['B7'].value,1)
                    if year == '1989':
                        self.assertEqual(wb['TONGHOP']['B5'].value,0)
                        self.assertEqual(wb['TONGHOP']['B6'].value,1)
                finally:
                    wb.close()
                self.assertTrue((output/'import'/f'import_{year}.xlsx').is_file())
                self.assertTrue((
                    output/'import'/'kiemtralai'/
                    f'import_khong_co_trong_SGD{year}.xlsx'
                ).is_file())
            imported_1987=load_workbook(
                output/'import'/'import_1987.xlsx',read_only=True,
            )
            try:
                self.assertEqual(imported_1987['Data']['A5'].value,'(1)')
                self.assertEqual(imported_1987['Data']['E6'].value,'Học sinh 1987')
                self.assertEqual(imported_1987['Data'].max_row,6)
            finally:
                imported_1987.close()
            self.assertFalse((output/'import'/'import_1990.xlsx').exists())

    def test_batch_writes_requested_sheets_summary_and_import_workbooks(self):
        with tempfile.TemporaryDirectory() as tmp:
            root=Path(tmp); sources=root/'sources'; schools=root/'schools'
            sources.mkdir(); schools.mkdir()
            source=sources/'THCS.CQ.1987_output.xlsx'
            write_input(source,[
                row(name='Chỉ có Sở',dob='03/04/1982',year='1987',
                    council='THCS Không gửi'),
                row(name='Học sinh khớp',year='1987',council='THCS Có nguồn'),
                row(name='Cần kiểm tra',dob='10/11/1980',year='1987',
                    council='THCS Chờ xác minh'),
            ],headers=(*HEADERS,'Nguồn dữ liệu','Dữ liệu cột nguồn chưa ánh xạ'))
            wb=load_workbook(source)
            stats=wb.create_sheet('THONGKE')
            stats.append(['STT','TÊN TRƯỜNG','SỐ LƯỢNG'])
            stats.append([1,'THCS Có nguồn',1])
            stats.append([2,'THCS Không gửi',1])
            stats.append(['TỔNG','Tổng số trường: 2',2])
            wb.save(source); wb.close()
            wb=Workbook(); ws=wb.active; ws.title='Data'
            ws.append(['Năm TN','Họ và tên','Ngày sinh','Tên trường đã học'])
            ws.append(['1987','Học sinh khớp','01/02/1980','THCS Có nguồn'])
            ws.append(['1987','Chỉ có trường','05/06/1983','THCS Trường gửi'])
            ws.append(['1987','Cần kiểm tra','','THCS Chờ xác minh'])
            wb.save(schools/'1987.xlsx'); wb.close()

            output=compare_batch(sources,schools,root/'out')

            report=load_workbook(output/'1987_Check.xlsx',read_only=True)
            try:
                self.assertEqual(report.sheetnames,[
                    'TONGHOP','DULIEU_OK','DULIEU_CHI_CO_TRONG_SGD',
                    'DULIEU_KHONGCO_TRONG_SGD','DULIEU_LECH_THONG_TIN',
                    'DULIEU_CAN_XAC_MINH',
                ])
                summary=list(report['TONGHOP'].values)
                self.assertEqual(summary[1][0:2],('Năm','1987'))
                self.assertEqual(summary[2][1],3)
                self.assertEqual(summary[3][1],3)
                self.assertEqual(summary[4][1],1)
                self.assertEqual(summary[5][1],0)
                self.assertEqual(summary[6][1],1)
                self.assertEqual(summary[7][2],'33.33%')
                self.assertEqual(summary[8][2],'33.33%')
                self.assertEqual(summary[9][1],4)
                self.assertEqual(summary[10][1],1)
                self.assertEqual(summary[11][1],1)
                self.assertEqual(summary[12][1],2)
                self.assertEqual(
                    report['DULIEU_CHI_CO_TRONG_SGD']['E3'].value,'Chỉ có Sở',
                )
                self.assertIn(
                    'THCS.CQ.1987_output.xlsx | Data | dòng 2',
                    report['DULIEU_CHI_CO_TRONG_SGD']['Y3'].value,
                )
                self.assertEqual(
                    report['DULIEU_KHONGCO_TRONG_SGD']['E3'].value,'Chỉ có trường',
                )
                self.assertEqual(
                    report['DULIEU_KHONGCO_TRONG_SGD']['Y3'].value,
                    '1987.xlsx | Data | dòng 3',
                )
                self.assertEqual(
                    report['DULIEU_LECH_THONG_TIN']['E3'].value,'Cần kiểm tra',
                )
                self.assertEqual(
                    report['DULIEU_CAN_XAC_MINH']['E3'].value,'Cần kiểm tra',
                )
                self.assertIn('góc nhìn',report['DULIEU_LECH_THONG_TIN']['A1'].value.lower())
                self.assertIn('góc nhìn',report['DULIEU_CAN_XAC_MINH']['A1'].value.lower())
            finally:
                report.close()

            matched=load_workbook(output/'import'/'import_1987.xlsx',read_only=True)
            self.assertEqual(matched.sheetnames,['Data','THONGKE'])
            self.assertEqual(matched['Data'].max_row,2)
            self.assertEqual(matched['Data']['E2'].value,'Học sinh khớp')
            self.assertEqual(matched['THONGKE']['C2'].value,1)
            matched.close()

            school_only=load_workbook(
                output/'import'/'kiemtralai'/'import_khong_co_trong_SGD1987.xlsx',
                read_only=True,
            )
            self.assertEqual(school_only.sheetnames,['Data','THONGKE'])
            self.assertEqual(school_only['Data'].max_row,2)
            self.assertEqual(school_only['Data']['E2'].value,'Chỉ có trường')
            self.assertEqual(school_only['Data']['Y2'].value,
                             'Tệp: 1987.xlsx | Sheet: Data | Dòng: 3')
            self.assertEqual(school_only['THONGKE']['C2'].value,1)
            school_only.close()

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
            self.assertEqual(wb.sheetnames,[
                'TONGHOP','DULIEU_OK','DULIEU_CHI_CO_TRONG_SGD',
                'DULIEU_KHONGCO_TRONG_SGD','DULIEU_LECH_THONG_TIN',
                'DULIEU_CAN_XAC_MINH',
            ])
            self.assertEqual(wb['TONGHOP'].max_row,13)
            self.assertEqual(wb['DULIEU_OK'].max_row,1)
            self.assertEqual(wb['DULIEU_LECH_THONG_TIN'].max_row,3)
            self.assertEqual(wb['DULIEU_LECH_THONG_TIN']['Y3'].value,REVIEW)
            self.assertEqual(wb['DULIEU_LECH_THONG_TIN']['E3'].value,'Nguyễn Văn An')
            self.assertIn('lệch ngày sinh',wb['DULIEU_LECH_THONG_TIN']['Z3'].value)
            self.assertIn('dòng 6',wb['DULIEU_LECH_THONG_TIN']['AD3'].value)
            self.assertEqual(wb['DULIEU_KHONGCO_TRONG_SGD'].max_row,2)
            self.assertIn('Dữ liệu trường gửi lên mà không có trong dữ liệu nguồn chuẩn',
                          wb['DULIEU_KHONGCO_TRONG_SGD']['A1'].value)
            self.assertEqual(wb['DULIEU_CAN_XAC_MINH'].max_row,3)
            wb.close()

    def test_compare_uses_batch_format_and_exports_imports_per_year(self):
        with tempfile.TemporaryDirectory() as tmp:
            root=Path(tmp); source=root/'source.xlsx'; school=root/'school.xlsx'
            write_input(source,[
                row(name='Học sinh 1987',year='1987'),
                row(name='Học sinh 1988',year='1988',council='THCS 1988'),
            ],headers=(*HEADERS,'Nguồn dữ liệu','Dữ liệu cột nguồn chưa ánh xạ'))
            wb=Workbook(); ws=wb.active; ws.title='Data'
            ws.append(['Năm TN','Họ và tên','Ngày sinh','Tên trường đã học'])
            ws.append(['1987','Học sinh 1987','01/02/1980','THCS A'])
            ws.append(['1988','Chỉ có trường','05/06/1983','THCS B'])
            wb.save(school); wb.close()

            output=compare([source],school,root/'out',years={1987,1988})

            report=load_workbook(output/'source_Check.xlsx',read_only=True)
            try:
                self.assertEqual(report.sheetnames,[
                    'TONGHOP','DULIEU_OK','DULIEU_CHI_CO_TRONG_SGD',
                    'DULIEU_KHONGCO_TRONG_SGD','DULIEU_LECH_THONG_TIN',
                    'DULIEU_CAN_XAC_MINH',
                ])
                self.assertEqual(report['TONGHOP']['B2'].value,'1987, 1988')
                self.assertEqual(report['TONGHOP']['B7'].value,1)
            finally:
                report.close()

            matched=load_workbook(output/'import'/'import_1987.xlsx',read_only=True)
            try:
                self.assertEqual(matched['Data'].max_row,2)
                self.assertEqual(matched['Data']['E2'].value,'Học sinh 1987')
            finally:
                matched.close()
            matched_1988=load_workbook(
                output/'import'/'import_1988.xlsx',read_only=True,
            )
            try:
                self.assertEqual(matched_1988['Data'].max_row,1)
            finally:
                matched_1988.close()

            review=load_workbook(
                output/'import'/'kiemtralai'/
                'import_khong_co_trong_SGD1988.xlsx',read_only=True,
            )
            try:
                self.assertEqual(review['Data'].max_row,2)
                self.assertEqual(review['Data']['E2'].value,'Chỉ có trường')
                self.assertEqual(
                    review['Data']['Y2'].value,
                    'Tệp: school.xlsx | Sheet: Data | Dòng: 3',
                )
            finally:
                review.close()

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
            self.assertEqual(wb['TONGHOP'].max_row,13)
            self.assertEqual(wb['DULIEU_OK'].max_row,2)
            self.assertEqual(wb['DULIEU_CHI_CO_TRONG_SGD'].max_row,2)
            self.assertEqual(wb['DULIEU_LECH_THONG_TIN'].max_row,2)
            self.assertEqual(wb['DULIEU_KHONGCO_TRONG_SGD'].max_row,2)
            self.assertEqual(wb['DULIEU_CAN_XAC_MINH'].max_row,2)
            self.assertEqual(wb['DULIEU_OK']['A2'].value,'Khớp hoàn toàn')
            headers=[cell.value for cell in wb['DULIEU_OK'][1]]
            school_location=wb['DULIEU_OK'].cell(2,headers.index('Vị trí file trường')+1).value
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
            self.assertEqual(wb['DULIEU_OK'].max_row,1)
            self.assertEqual(wb['DULIEU_CHI_CO_TRONG_SGD'].max_row,3)
            self.assertEqual(wb['DULIEU_CHI_CO_TRONG_SGD']['A3'].value,'1')
            self.assertEqual(wb['DULIEU_CHI_CO_TRONG_SGD']['E3'].value,'Nguyễn Văn An')
            self.assertEqual(wb['DULIEU_KHONGCO_TRONG_SGD'].max_row,3)
            self.assertEqual(wb['DULIEU_KHONGCO_TRONG_SGD']['E3'].value,'Lê Thị B')
            self.assertEqual(wb['DULIEU_LECH_THONG_TIN'].max_row,2)
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
                self.assertEqual(wb['DULIEU_OK'].max_row,2)
                headers=[cell.value for cell in wb['DULIEU_OK'][1]]
                note_column=headers.index('Ghi chú đối soát')+1
                note=wb['DULIEU_OK'].cell(2,note_column).value
                self.assertIn('Có 2 hồ sơ file trường khớp thông tin',note)
                self.assertIn('school.xlsx | Data | dòng 2',note)
                self.assertIn('school.xlsx | Data | dòng 3',note)
                self.assertNotIn(str(root),note)
                self.assertFalse(any(
                    str(root).casefold() in str(cell.value).casefold()
                    for sheet in wb.worksheets for row_cells in sheet.iter_rows()
                    for cell in row_cells if cell.value is not None
                ))
                self.assertEqual(wb['DULIEU_LECH_THONG_TIN'].max_row,2)
                self.assertEqual(wb['DULIEU_CAN_XAC_MINH'].max_row,4)
                self.assertEqual(wb['TONGHOP']['B7'].value,1)
                self.assertEqual(wb['TONGHOP']['B8'].value,2)
            finally:
                wb.close()

    def test_report_lists_duplicate_candidate_sources_once_each(self):
        with tempfile.TemporaryDirectory() as tmp:
            root=Path(tmp); source=root/'1987.xlsx'; school=root/'school.xlsx'
            write_input(source,[row(),row()])
            write_input(school,[row()])
            report=compare([source],school,root/'out',years={1995})
            wb=load_workbook(report/'1987_Check.xlsx')
            self.assertEqual(wb['TONGHOP']['B3'].value,2)
            self.assertEqual(wb['TONGHOP']['B4'].value,1)
            self.assertEqual(wb['TONGHOP']['B13'].value,2)
            self.assertEqual(wb['DULIEU_LECH_THONG_TIN'].max_row,4)
            self.assertEqual(wb['DULIEU_CHI_CO_TRONG_SGD'].max_row,2)
            self.assertEqual(wb['DULIEU_CAN_XAC_MINH'].max_row,3)
            self.assertEqual(wb['DULIEU_KHONGCO_TRONG_SGD'].max_row,2)
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
            self.assertEqual(wb['TONGHOP']['B7'].value,2)
            self.assertEqual(wb['DULIEU_OK'].max_row,3)
            self.assertEqual(wb['DULIEU_CHI_CO_TRONG_SGD'].max_row,2)
            self.assertEqual(wb['DULIEU_LECH_THONG_TIN'].max_row,2)
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
            self.assertEqual(content[0],(*HEADERS,'Nguồn dữ liệu','Dữ liệu cột nguồn chưa ánh xạ'))
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
