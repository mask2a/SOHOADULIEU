"""Streaming, source-centric Excel reconciliation report."""
from collections import Counter
from pathlib import Path
import pickle
import tempfile
from openpyxl import Workbook
from openpyxl.cell import WriteOnlyCell
from openpyxl.styles import PatternFill
from openpyxl.utils import get_column_letter
from .domain import AppError, HEADERS, IDENTITY, MATCH, NORMALIZED, REVIEW, ONLY_SOURCE, ONLY_SCHOOL, text
from .excel import append_text, save_atomic, MAX_ROWS

MATCHED = (MATCH, NORMALIZED)
IDENTITY_LABELS = ('họ tên', 'ngày sinh', 'năm tốt nghiệp')


def _candidate_summary(records):
    entries = []
    for record in records[:3]:
        name, dob, year = (record.values[i] for i in IDENTITY)
        entries.append(f'{name} | {dob} | {year} | {record.location}')
    if len(records) > 3:
        entries.append(f'… và {len(records) - 3} hồ sơ khác')
    return '; '.join(entries)


def _difference_summary(source, candidates):
    if any(source.keys == candidate.keys for candidate in candidates):
        return 'Có ứng viên trùng bộ ba; cần xác minh'
    differences = []
    for candidate in candidates:
        for index, (field_index, label) in enumerate(zip(IDENTITY, IDENTITY_LABELS)):
            source_value = text(source.values[field_index]).strip()
            candidate_value = text(candidate.values[field_index]).strip()
            if not source_value and candidate_value:
                difference = f'thiếu {label} ở nguồn chuẩn'
            elif source_value and not candidate_value:
                difference = f'thiếu {label} ở file trường'
            elif source.keys[index] != candidate.keys[index]:
                difference = f'lệch {label}'
            else:
                continue
            if difference not in differences:
                differences.append(difference)
    return '; '.join(differences) or 'Không khớp duy nhất; cần xác minh'


class ReportWriter:
    def __init__(self, folder, scope, source_count, school_count, row_limit=MAX_ROWS,
                 output_name='Bao_cao_Check'):
        self.folder = Path(folder)
        self.scope = scope
        self.source_count, self.school_count = source_count, school_count
        self.limit = row_limit
        self.output_name = str(output_name)
        self.stats = Counter()
        self.part = 0
        self.paths = []
        self.spool = tempfile.TemporaryFile()

    def add(self, result):
        if result.status in MATCHED:
            self.stats['matched'] += 1
            self.stats['matched_school'] += len(result.matching_schools) or 1
            self.stats[result.status] += 1
        elif result.source:
            self.stats['source_unmatched'] += 1
            self.stats['source_review' if result.status == REVIEW else 'source_only'] += 1
        elif result.school:
            self.stats['school_review' if result.status == REVIEW else 'school_only'] += 1
        pickle.dump(result, self.spool, protocol=pickle.HIGHEST_PROTOCOL)

    def _summary_rows(self):
        matched = self.stats['matched']
        matched_school = self.stats['matched_school']
        return [
            ('Phạm vi năm', self.scope, ''),
            ('Tổng hồ sơ nguồn chuẩn', self.source_count, '100%'),
            ('Tổng hồ sơ trường gửi', self.school_count, '100%'),
            ('Khớp hoàn toàn', self.stats[MATCH], ''),
            ('Khớp sau chuẩn hóa', self.stats[NORMALIZED], ''),
            ('Tổng hồ sơ nguồn khớp', matched, ''),
            ('Tỷ lệ hồ sơ file trường khớp / hồ sơ trường gửi', matched_school,
             f'{matched_school / self.school_count:.2%}' if self.school_count else 'Không có mẫu số'),
            ('Độ bao phủ / hồ sơ nguồn chuẩn', matched,
             f'{matched / self.source_count:.2%}' if self.source_count else 'Không có mẫu số'),
            ('Nguồn chưa khớp', self.source_count - matched, ''),
            ('  Chỉ có trong nguồn chuẩn (sheet riêng)', self.stats['source_only'], ''),
            ('  Nguồn cần xác minh', self.stats['source_review'], ''),
            ('Trường chỉ có (không tìm thấy ứng viên nguồn)', self.stats['school_only'], ''),
            ('Trường cần xác minh (không kết luận là thừa)', self.stats['school_review'], ''),
        ]

    def _open(self):
        self.part += 1
        self.wb = Workbook(write_only=True)
        self.summary = self.wb.create_sheet('Tong_hop')
        self.matched = self.wb.create_sheet('Khop')
        self.source_only = self.wb.create_sheet('Chi_co_trong_nguon_chuan')
        self.diff = self.wb.create_sheet('Lech_thong_tin')
        self.extra = self.wb.create_sheet('Truong_khong_co_trong_nguon')
        self.review = self.wb.create_sheet('Can_xac_minh')
        self.sheet_rows = {ws.title: 1 for ws in (
            self.matched, self.source_only, self.diff, self.extra, self.review
        )}
        self.summary.freeze_panes = 'A2'
        self.summary.column_dimensions['A'].width = 55
        self.summary.column_dimensions['B'].width = 28
        self.summary.column_dimensions['C'].width = 24
        append_text(self.summary, ['Chỉ tiêu', 'Số lượng / giá trị', 'Tỷ lệ'], row_number=1)
        for row in self._summary_rows():
            cells = []
            for value in row:
                cell = WriteOnlyCell(self.summary, value=value)
                if isinstance(value, str):
                    cell.data_type = 's'
                cells.append(cell)
            self.summary.append(cells)

        matched_headers = (['Trạng thái'] +
                           [f'Nguồn chuẩn — {header}' for header in HEADERS] +
                           [f'File trường — {header}' for header in HEADERS] +
                           ['Vị trí nguồn chuẩn', 'Vị trí file trường', 'Ghi chú đối soát'])
        diff_headers = list(HEADERS) + [
            'Trạng thái', 'Thông tin lệch', 'Số ứng viên file trường',
            'Thông tin ứng viên (tối đa 3)', 'Ghi chú', 'Vị trí nguồn chuẩn'
        ]
        source_only_headers = list(HEADERS) + ['Vị trí nguồn chuẩn']
        extra_headers = list(HEADERS) + ['Vị trí file trường']
        review_headers = list(HEADERS) + [
            'Số ứng viên nguồn', 'Vị trí ứng viên nguồn (tối đa 3)', 'Ghi chú',
            'Vị trí file trường gửi'
        ]
        for ws, headers in ((self.matched, matched_headers), (self.source_only, source_only_headers),
                            (self.diff, diff_headers), (self.extra, extra_headers),
                            (self.review, review_headers)):
            ws.freeze_panes = 'A2'
            for index in range(1, len(headers) + 1):
                ws.column_dimensions[get_column_letter(index)].width = 28
            if ws is self.extra:
                append_text(ws, ['Dữ liệu trường gửi lên mà không có trong dữ liệu nguồn chuẩn'])
                append_text(ws, headers, row_number=2)
                self.sheet_rows[ws.title] = 2
            elif ws is self.source_only:
                append_text(ws, ['Dữ liệu chỉ có trong nguồn chuẩn'])
                append_text(ws, headers, row_number=2)
                self.sheet_rows[ws.title] = 2
            else:
                append_text(ws, headers, row_number=1)

    def _capacity(self, ws):
        if self.sheet_rows[ws.title] >= self.limit - 1:
            self._save()
            self._open()
            return getattr(self, {
                'Khop': 'matched',
                'Chi_co_trong_nguon_chuan': 'source_only',
                'Lech_thong_tin': 'diff',
                'Truong_khong_co_trong_nguon': 'extra',
                'Can_xac_minh': 'review',
            }[ws.title])
        return ws

    def _append_match(self, result):
        school_values = result.school.values if result.school else ('',) * len(HEADERS)
        values = [result.status, *result.source.values, *school_values,
                  result.source.location, result.school.location if result.school else '',
                  result.reason]
        cells = []
        for index, value in enumerate(values):
            cell = WriteOnlyCell(self.matched, value=text(value))
            cell.data_type = 's'
            source_index = index - 1 if 1 <= index <= len(HEADERS) else None
            school_index = index - 1 - len(HEADERS) if len(HEADERS) + 1 <= index <= 2 * len(HEADERS) else None
            if (source_index is not None and source_index in IDENTITY and
                    result.school is not None and
                    result.source.values[source_index] != result.school.values[source_index]) or (
                    school_index is not None and school_index in IDENTITY and
                    result.school is not None and
                    result.source.values[school_index] != result.school.values[school_index]):
                cell.fill = PatternFill('solid', fgColor='FFD966')
            cells.append(cell)
        self.matched.append(cells)

    def _render(self, result):
        if result.status in MATCHED:
            self.matched = self._capacity(self.matched)
            self._append_match(result)
            self.sheet_rows[self.matched.title] += 1
        elif result.status == ONLY_SOURCE:
            self.source_only = self._capacity(self.source_only)
            append_text(self.source_only, (*result.source.values, result.source.location))
            self.sheet_rows[self.source_only.title] += 1
        elif result.source:
            self.diff = self._capacity(self.diff)
            append_text(self.diff, (*result.source.values, result.status,
                                    _difference_summary(result.source, result.candidates),
                                    len(result.candidates),
                                    _candidate_summary(result.candidates), result.reason,
                                    result.source.location))
            self.sheet_rows[self.diff.title] += 1
        elif result.status == ONLY_SCHOOL:
            self.extra = self._capacity(self.extra)
            append_text(self.extra, (*result.school.values, result.school.location))
            self.sheet_rows[self.extra.title] += 1
        elif result.school:
            self.review = self._capacity(self.review)
            append_text(self.review, (*result.school.values, len(result.candidates),
                                      _candidate_summary(result.candidates), result.reason,
                                      result.school.location))
            self.sheet_rows[self.review.title] += 1

    def _save(self):
        path = self.folder / f'{self.output_name}_{self.part:03}.xlsx'
        save_atomic(self.wb, path)
        self.paths.append(path)

    def finish(self, check=lambda: None):
        try:
            if self.stats['matched'] + self.stats['source_unmatched'] != self.source_count:
                raise AppError('Số hồ sơ nguồn phân loại không khớp tổng số dòng nguồn.')
            if self.stats['matched_school'] + self.stats['school_only'] + self.stats['school_review'] != self.school_count:
                raise AppError('Số hồ sơ trường phân loại không khớp tổng số dòng trường gửi.')
            self.spool.seek(0)
            self._open()
            while True:
                check()
                try:
                    result = pickle.load(self.spool)
                except EOFError:
                    break
                self._render(result)
            self._save()
            if len(self.paths) == 1:
                final = self.folder / f'{self.output_name}.xlsx'
                self.paths[0].replace(final)
                self.paths[0] = final
            return self.paths
        finally:
            self.spool.close()
