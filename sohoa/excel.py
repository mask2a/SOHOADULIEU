"""Excel adapter: parsing, mapping validation, safe cell output."""
from collections import Counter
from copy import copy
from datetime import date, datetime
from pathlib import Path
import re
from openpyxl import load_workbook, Workbook
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter
import math
from .domain import HEADERS, IDENTITY, UNMAPPED_HEADER, StudentRecord, StructureError, AppError, Cancelled, normalize, text

MAX_ROWS = 1_048_576

_HEADER_ALIASES = {
    normalize('Họ và tên'): 4,
    normalize('Họ tên'): 4,
    normalize('Ngày sinh'): 6,
    normalize('Ngày tháng năm sinh'): 6,
    normalize('Năm TN'): 12,
    normalize('Năm TN THCS'): 12,
    normalize('Năm tốt nghiệp THCS'): 12,
    normalize('Xếp loại tốt nghiệp'): 20,
    normalize('Xếp loại TN,'): 20,
    normalize('Mã đơn vị'): 22,
}
_CANONICAL_HEADERS = {normalize(header): index for index, header in enumerate(HEADERS)}


def _header_mapping(headers):
    mapping = {}
    duplicates = {}
    for column, header in enumerate(headers):
        key = normalize(header)
        canonical = _CANONICAL_HEADERS.get(key, _HEADER_ALIASES.get(key))
        if canonical is None:
            continue
        if canonical in mapping:
            duplicates.setdefault(canonical, [mapping[canonical]]).append(column)
        else:
            mapping[canonical] = column
    return mapping, duplicates


def flexible_header_mapping(headers):
    """Map known aliases and safely infer blank cells from neighboring headers."""
    mapping, duplicates = _header_mapping(headers)
    if 0 not in mapping and len(headers) > 1 and mapping.get(1) == 1:
        mapping[0] = 0

    ordered = sorted((column, index) for index, column in mapping.items())
    if any(left[1] >= right[1] for left, right in zip(ordered, ordered[1:])):
        return mapping, duplicates
    for (left_column, left_index), (right_column, right_index) in zip(
            ordered, ordered[1:]):
        gap = right_column - left_column
        if gap <= 1 or gap != right_index - left_index:
            continue
        for offset in range(1, gap):
            column = left_column + offset
            index = left_index + offset
            if (index not in mapping and column < len(headers)
                    and not normalize(headers[column])):
                mapping[index] = column
    return mapping, duplicates


def data_sheet(wb, path):
    if wb.worksheets and wb.worksheets[0].title.casefold() == 'data':
        return wb.worksheets[0]
    if len(wb.worksheets) != 1:
        raise StructureError([f'{path}: cần duy nhất một sheet dữ liệu (tên bất kỳ); hiện có {wb.sheetnames}'])
    return wb.worksheets[0]


def export_import_workbook(template_path, output_path, keep_rows=None,
                           records_to_append=(), check=lambda: None):
    template_path = Path(template_path)
    output_path = Path(output_path)
    wb = None
    try:
        wb = load_workbook(template_path)
        ws = data_sheet(wb, template_path)
        header_row = canonical_header_row(template_path)
        data_start = header_row + (2 if header_row == 4 else 1)
        original_max_row = ws.max_row
        style_row = min(data_start, original_max_row)
        row_styles = [
            copy(ws.cell(style_row, column)._style)
            for column in range(1, ws.max_column + 1)
        ] if style_row else []

        if keep_rows is None:
            if original_max_row >= data_start:
                ws.delete_rows(data_start, original_max_row - data_start + 1)
        else:
            output_row = data_start
            for row_number in range(data_start, original_max_row + 1):
                if row_number % 1000 == 0:
                    check()
                if row_number not in keep_rows:
                    continue
                if output_row != row_number:
                    for column in range(1, ws.max_column + 1):
                        source_cell = ws.cell(row_number, column)
                        target_cell = ws.cell(output_row, column)
                        target_cell.value = source_cell.value
                        target_cell.data_type = source_cell.data_type
                        target_cell._style = copy(source_cell._style)
                        target_cell._hyperlink = (
                            copy(source_cell.hyperlink) if source_cell.hyperlink else None
                        )
                        target_cell.comment = (
                            copy(source_cell.comment) if source_cell.comment else None
                        )
                    source_dimension = ws.row_dimensions[row_number]
                    target_dimension = ws.row_dimensions[output_row]
                    target_dimension.height = source_dimension.height
                    target_dimension.hidden = source_dimension.hidden
                output_row += 1
            if output_row <= original_max_row:
                ws.delete_rows(output_row, original_max_row - output_row + 1)

        headers = [text(ws.cell(header_row, column).value)
                   for column in range(1, ws.max_column + 1)]
        for record in records_to_append:
            check()
            values = [''] * len(headers)
            for index, value in enumerate(record.values[:24]):
                values[index] = text(value)
            unmapped = '\n'.join(
                f'{header}: {value}' for header, value in record.unmapped
            )
            for column, header in enumerate(headers):
                key = normalize(header)
                if key == normalize('Nguồn dữ liệu'):
                    values[column] = (
                        f'Tệp: {Path(record.file).name} | Sheet: {record.sheet} | '
                        f'Dòng: {record.row}'
                    )
                elif key == normalize(UNMAPPED_HEADER):
                    values[column] = unmapped

            row_number = ws.max_row + 1
            for column, value in enumerate(values, 1):
                cell = ws.cell(row_number, column, value=value)
                if isinstance(value, str):
                    cell.data_type = 's'
                if column <= len(row_styles):
                    cell._style = copy(row_styles[column - 1])

        _refresh_school_statistics(wb, ws, data_start)
        tmp = output_path.with_name(output_path.stem + '.partial.xlsx')
        try:
            wb.save(tmp)
            tmp.replace(output_path)
        finally:
            if tmp.exists():
                tmp.unlink()
    except Cancelled:
        raise
    except Exception as exc:
        raise AppError(f'Không tạo được file import {output_path}: {exc}') from exc
    finally:
        if wb is not None:
            wb.close()


def _refresh_school_statistics(wb, data_ws, data_start):
    stats = next(
        (sheet for sheet in wb.worksheets if sheet.title.casefold() == 'thongke'),
        None,
    )
    if stats is None:
        return

    style_row = 2 if stats.max_row >= 2 else 1
    styles = [copy(stats.cell(style_row, column)._style) for column in range(1, 4)]
    if stats.max_row > 1:
        stats.delete_rows(2, stats.max_row - 1)

    schools = {}
    for row_number in range(data_start, data_ws.max_row + 1):
        if not any(data_ws.cell(row_number, column).value is not None
                   for column in range(1, len(HEADERS) + 1)):
            continue
        name = text(data_ws.cell(row_number, 12).value).strip()
        key = normalize(name)
        display, count = schools.get(key, (name, 0))
        schools[key] = (display, count + 1)

    def append(values):
        row_number = stats.max_row + 1
        for column, value in enumerate(values, 1):
            cell = stats.cell(row_number, column, value=value)
            if isinstance(value, str):
                cell.data_type = 's'
            cell._style = copy(styles[column - 1])

    for index, key in enumerate(sorted(schools), 1):
        name, count = schools[key]
        append([index, name or '(Chưa có tên trường)', count])
    named_count = sum(bool(key) for key in schools)
    append(['TỔNG', f'Tổng số trường: {named_count}',
            sum(count for _, count in schools.values())])


def _canonical_header_row(ws):
    expected = tuple(map(normalize, HEADERS))
    ws.reset_dimensions()
    for row_no, cells in enumerate(ws.iter_rows(min_row=1, max_row=10, values_only=True), 1):
        if tuple(normalize(value) for value in cells[:24]) == expected:
            return row_no
    return None


def _flexible_header_row(ws):
    ws.reset_dimensions()
    for row_no, cells in enumerate(ws.iter_rows(min_row=1, max_row=10, values_only=True), 1):
        mapping, _ = flexible_header_mapping(cells)
        if all(index in mapping for index in IDENTITY):
            return row_no
    return None


def canonical_header_row(path):
    try:
        wb = load_workbook(path, read_only=True, data_only=False)
    except Exception as exc:
        raise AppError(f'Không đọc được {path}: {exc}') from exc
    try:
        ws = data_sheet(wb, path)
        row_no = _canonical_header_row(ws)
        if row_no is None:
            raise StructureError([f'{path}: không tìm thấy bộ 24 cột chuẩn trong 10 dòng đầu.'])
        return row_no
    finally:
        wb.close()


def inspect(path, header_row, sheet_name=None):
    try:
        wb = load_workbook(path, read_only=True, data_only=False)
        try:
            ws = wb[sheet_name] if sheet_name else data_sheet(wb, path)
            ws.reset_dimensions()
            return [text(c.value) for c in next(ws.iter_rows(min_row=header_row, max_row=header_row))]
        finally:
            wb.close()
    except (StructureError, AppError):
        raise
    except Exception as exc:
        raise AppError(f'Không đọc được {path}: {exc}') from exc

def auto_mapping(headers):
    mapping, _ = _header_mapping(headers)
    return mapping

def records(path, header_row, mapping=None, canonical=False, log=lambda _: None,
            sheet_name=None, flexible_headers=False,
            progress=lambda current,total: None):
    headers = inspect(path, header_row, sheet_name)
    if flexible_headers:
        mapping, duplicates = flexible_header_mapping(headers)
    else:
        mapping = auto_mapping(headers) if mapping is None else mapping
        duplicates = {}
    issues = []
    required = IDENTITY if flexible_headers else range(24) if canonical else IDENTITY
    for i in required:
        if i not in mapping:
            issues.append(f'{path}: thiếu ánh xạ cột {HEADERS[i]}')
    used = list(mapping.values())
    if len(used) != len(set(used)):
        issues.append(f'{path}: một cột đầu vào đang được ánh xạ cho nhiều cột chuẩn')
    for i, j in mapping.items():
        if i not in range(24) or j < 0 or j >= len(headers):
            issues.append(f'{path}: ánh xạ ngoài phạm vi {i} -> {j}')
    if canonical and not flexible_headers:
        if tuple(normalize(h) for h in headers[:24]) != tuple(map(normalize, HEADERS)):
            issues.append(f'{path}: bộ 24 cột chuẩn không đúng thứ tự hoặc tên')
    if issues:
        raise StructureError(issues)
    if canonical and flexible_headers:
        for index, column in mapping.items():
            if normalize(headers[column]) != normalize(HEADERS[index]):
                if normalize(headers[column]):
                    log(
                        f'{Path(path).name}: tiêu đề "{headers[column]}" ở cột '
                        f'{column + 1} được nhận diện theo vị trí là "{HEADERS[index]}".'
                    )
                else:
                    log(
                        f'{Path(path).name}: cột {column + 1} không có tiêu đề; '
                        f'nhận diện theo vị trí là "{HEADERS[index]}".'
                    )
        missing_optional = [HEADERS[i] for i in range(len(HEADERS)) if i not in mapping]
        if missing_optional:
            log(
                f'{Path(path).name}: không có tiêu đề cho '
                f'{", ".join(missing_optional)}; các cột này sẽ để trống.'
            )
    wb = load_workbook(path, read_only=True, data_only=False)
    try:
        ws = wb[sheet_name] if sheet_name else data_sheet(wb, path)
        reported_max_row = ws.max_row
        if flexible_headers:
            ws.reset_dimensions()
        if flexible_headers and duplicates:
            populated = Counter()
            conflicting = set()
            for row in ws.iter_rows(min_row=header_row + 1, values_only=True):
                for index, columns in duplicates.items():
                    values = [
                        text(row[column]).strip()
                        for column in columns
                        if column < len(row) and text(row[column]).strip()
                    ]
                    populated.update(
                        column for column in columns
                        if column < len(row) and text(row[column]).strip()
                    )
                    if len(set(values)) > 1:
                        conflicting.add(index)
            ambiguous = [index for index in duplicates if index in IDENTITY and index in conflicting]
            if ambiguous:
                raise StructureError([
                    f'{path}: các cột trùng tiêu đề định danh '
                    f'{", ".join(HEADERS[index] for index in ambiguous)} có dữ liệu khác nhau.'
                ])
            for index, columns in duplicates.items():
                chosen = max(columns, key=lambda column: (populated[column], -column))
                mapping[index] = chosen
                log(
                    f'{Path(path).name}: tiêu đề {HEADERS[index]} xuất hiện nhiều lần '
                    f'ở cột {", ".join(str(column + 1) for column in columns)}; '
                    f'dùng cột {chosen + 1} có {populated[chosen]} ô có dữ liệu.'
                )

        used_columns = set(mapping.values())
        unmapped_columns = [
            column for column in range(len(headers)) if column not in used_columns
        ] if flexible_headers else []
        unmapped_values = Counter()
        numeric_code_warnings = Counter()
        total_data_rows = max(0, (reported_max_row or header_row) - header_row)
        progress(0, total_data_rows)
        for row_no, cells in enumerate(ws.iter_rows(min_row=header_row + 1), header_row + 1):
            rows_scanned = row_no - header_row
            if rows_scanned % 100 == 0:
                progress(rows_scanned, total_data_rows)
            raw = [c.value for c in cells]
            if flexible_headers:
                while raw and (raw[-1] is None or not text(raw[-1]).strip()):
                    raw.pop()
            row_unmapped_columns = list(unmapped_columns)
            if len(raw) > len(headers):
                row_unmapped_columns.extend(range(len(headers), len(raw)))
            for column in row_unmapped_columns:
                if column < len(raw) and raw[column] is not None and text(raw[column]).strip():
                    unmapped_values[column] += 1
            if not any(v is not None and text(v).strip() for v in raw):
                continue
            selected = [raw[j] for j in mapping.values() if j < len(raw)]
            nonempty = [text(v).strip() for v in selected if v is not None and text(v).strip()]
            if (canonical and header_row == 4 and row_no == 5) or (nonempty and all(re.fullmatch(r'\(\d+\)?', v) for v in nonempty)):
                log(f'Bỏ dòng số cột: {Path(path).name}:{row_no}')
                continue
            if sum(normalize(v) == normalize(HEADERS[i]) for i, j in mapping.items() if j < len(raw) for v in [raw[j]]) >= 3:
                log(f'Bỏ tiêu đề lặp: {Path(path).name}:{row_no}')
                continue
            values = [''] * 24
            for i, j in mapping.items():
                if j < len(raw):
                    v = raw[j]
                    if isinstance(v, (date, datetime)):
                        v = v.strftime('%d/%m/%Y')
                    # Recover displayed leading zeros only when Excel explicitly supplies a zero mask.
                    if i in (1, 2, 5, 22) and isinstance(v, (int, float)):
                        mask = cells[j].number_format
                        if re.fullmatch(r'0+', mask) and float(v).is_integer():
                            v = str(int(v)).zfill(len(mask))
                        else:
                            numeric_code_warnings[i] += 1
                    values[i] = text(v)
            unmapped = tuple(
                (
                    (headers[column] if column < len(headers) else '')
                    or f'Cột nguồn {column + 1} (không có tiêu đề)',
                    text(raw[column]),
                )
                for column in row_unmapped_columns
                if column < len(raw) and raw[column] is not None
                and text(raw[column]).strip()
            )
            yield StudentRecord(
                tuple(values), str(Path(path).resolve()), ws.title, row_no,
                unmapped=unmapped,
            )
        progress(total_data_rows, total_data_rows)
        for index, count in sorted(numeric_code_warnings.items()):
            log(
                f'Cần kiểm tra mã dạng số, có thể mất số 0 đầu: {Path(path).name}, '
                f'{count} ô ở cột {HEADERS[index]}.'
            )
        if unmapped_values:
            details = _unmapped_column_details(unmapped_values, headers)
            log(
                f'{Path(path).name}: có dữ liệu ở cột chưa nhận diện, không đưa vào '
                f'các cột chuẩn: {details}.'
            )
    except (OSError, ValueError) as exc:
        raise AppError(f'Lỗi đọc {path}: {exc}') from exc
    finally:
        wb.close()


def merge_records(path, log=lambda _: None, progress=lambda current,total: None):
    """Read sheets with recognizable identity headers, using only the first Data tab."""
    wb = load_workbook(path, read_only=True, data_only=False)
    sheets = []
    try:
        worksheets = wb.worksheets
        if worksheets and worksheets[0].title.casefold() == 'data':
            worksheets = worksheets[:1]
        for ws in worksheets:
            row_no = _flexible_header_row(ws)
            if row_no is not None:
                sheets.append((ws.title, row_no))
    finally:
        wb.close()
    if not sheets:
        raise StructureError([
            f'{path}: không nhận diện đủ Họ tên, Ngày sinh và Năm tốt nghiệp '
            'trong 10 dòng đầu.'
        ])
    for name, header in sheets:
        yield from records(path, header, canonical=True, log=log, sheet_name=name,
                           flexible_headers=True, progress=progress)

def fit_row(ws, values, row_number):
    lines = 1
    for index, value in enumerate(values, 1):
        width = ws.column_dimensions[get_column_letter(index)].width or 13
        capacity = max(1, int(width * 0.85) - 2)
        lines = max(lines, sum(max(1, math.ceil(len(part) / capacity))
                               for part in text(value).split('\n')))
    ws.row_dimensions[row_number].height = min(409, max(30, lines * 16 + 8))


def setup_data_layout(ws):
    widths = (8,20,24,30,30,22,20,12,16,18,38,42,18,38,38,22,42,30,30,24,18,42,22,38,65,85)
    for index, width in enumerate(widths, 1):
        ws.column_dimensions[get_column_letter(index)].width = width
    ws.freeze_panes = 'E2'


def append_text(ws, values, row_number=None):
    from openpyxl.cell import WriteOnlyCell
    cells = []
    for value in values:
        c = WriteOnlyCell(ws, value=text(value))
        c.data_type = 's'
        if row_number is not None:
            c.alignment = Alignment(wrap_text=True, vertical='top')
            if row_number == 1:
                c.font = Font(bold=True)
                c.fill = PatternFill('solid', fgColor='D9E2F3')
        cells.append(c)
    if row_number is not None:
        fit_row(ws, values, row_number)
    ws.append(cells)

def add_school_statistics(wb, schools):
    """Summarize the records in this output part, using normalized school names."""
    from openpyxl.cell import WriteOnlyCell
    ws = wb.create_sheet('THONGKE')
    ws.column_dimensions['A'].width = 12
    ws.column_dimensions['B'].width = 65
    ws.column_dimensions['C'].width = 20
    ws.freeze_panes = 'A2'
    row_number = 0

    def add(values, bold=False):
        nonlocal row_number
        row_number += 1
        fit_row(ws, values, row_number)
        cells = []
        for value in values:
            cell = WriteOnlyCell(ws, value=value)
            if isinstance(value, str):
                cell.data_type = 's'
            cell.font = Font(bold=bold)
            cell.alignment = Alignment(wrap_text=True, vertical='top')
            if bold:
                cell.fill = PatternFill('solid', fgColor='D9E2F3')
            cells.append(cell)
        ws.append(cells)

    add(['STT', 'TÊN TRƯỜNG', 'SỐ LƯỢNG'], bold=True)
    for index, key in enumerate(sorted(schools), 1):
        name, count = schools[key]
        add([index, name or '(Chưa có tên trường)', count])
    named_count = sum(bool(key) for key in schools)
    add(['TỔNG', f'Tổng số trường: {named_count}', sum(count for _, count in schools.values())], bold=True)


def save_atomic(wb, path):
    path = Path(path)
    tmp = path.with_name(path.stem + '.partial.xlsx')
    try:
        wb.save(tmp)
        tmp.replace(path)
    except Exception as exc:
        raise AppError(f'Không ghi được {path}: {exc}') from exc
    finally:
        wb.close()

def _unmapped_column_details(unmapped_values, headers, max_groups=20):
    groups = []
    for column, count in sorted(unmapped_values.items()):
        header = text(headers[column]).strip() if column < len(headers) else ''
        label = header[:80] or 'không có tiêu đề'
        if groups and groups[-1][1] + 1 == column and groups[-1][2] == label:
            groups[-1][1] = column
            groups[-1][3] += count
            groups[-1][4] += 1
        else:
            groups.append([column, column, label, count, 1])

    details = []
    for start, end, label, count, column_count in groups[:max_groups]:
        columns = f'cột {start + 1}'
        if end > start:
            columns += f'-{end + 1}'
        quantity = f'{count} ô'
        if column_count > 1:
            quantity += f' trên {column_count} cột'
        details.append(f'{columns} ({label}): {quantity}')

    if len(groups) > max_groups:
        remaining = groups[max_groups:]
        omitted_columns = sum(group[4] for group in remaining)
        omitted_values = sum(group[3] for group in remaining)
        details.append(f'... và {omitted_columns} cột khác ({omitted_values} ô)')
    return '; '.join(details)

def error_report(path, issues):
    wb = Workbook(write_only=True)
    ws = wb.create_sheet('Loi_du_lieu')
    append_text(ws, ['Lỗi', 'Hướng xử lý'])
    for issue in issues:
        append_text(ws, [issue, 'Kiểm tra cấu trúc/ánh xạ; sửa và gửi lại file.'])
    save_atomic(wb, path)
