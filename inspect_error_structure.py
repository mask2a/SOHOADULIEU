"""Diagnose workbook structure problems in a folder of rejected Pha 1 files."""
import argparse
from pathlib import Path
import sys
import tkinter as tk
from tkinter import filedialog, messagebox

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment

from sohoa.domain import HEADERS, IDENTITY, normalize
from sohoa.excel import flexible_header_mapping


REPORT_NAME = 'phan_tich_loi_cau_truc.xlsx'
SUPPORTED_EXTENSIONS = {'.xlsx', '.xlsm'}


def _header_candidate(worksheet):
    candidates = []
    for row_number, values in enumerate(
        worksheet.iter_rows(min_row=1, max_row=10, values_only=True), 1
    ):
        mapping, _ = flexible_header_mapping(values)
        if all(index in mapping for index in IDENTITY):
            candidates.append((len(mapping), row_number, values, mapping))
    if not candidates:
        return None
    _, row_number, values, mapping = max(
        candidates, key=lambda candidate: (candidate[0], -candidate[1])
    )
    aliases = []
    for index, column in mapping.items():
        if normalize(values[column]) == normalize(HEADERS[index]):
            continue
        if normalize(values[column]):
            aliases.append(
                f'Cột {column + 1} "{values[column]}" được nhận diện là "{HEADERS[index]}"'
            )
        else:
            aliases.append(
                f'Cột {column + 1} trống được suy ra là "{HEADERS[index]}" theo vị trí'
            )
    missing_optional = [
        HEADERS[index] for index in range(len(HEADERS)) if index not in mapping
    ]
    return row_number, mapping, aliases, missing_optional


def _sheet_diagnosis(worksheet):
    candidate = _header_candidate(worksheet)
    if candidate is not None:
        header_row, _, aliases, missing_optional = candidate
        details = 'Nhận diện được Họ tên, Ngày sinh và Năm tốt nghiệp; Pha 1 có thể đọc sheet này.'
        if aliases:
            details += ' Tiêu đề thay thế: ' + '; '.join(aliases) + '.'
        if missing_optional:
            details += ' Cột phụ không có tiêu đề và sẽ để trống: ' + '; '.join(missing_optional) + '.'
        return {
            'sheet': worksheet.title,
            'status': 'Đọc được',
            'header_row': header_row,
            'details': details,
        }

    return {
        'sheet': worksheet.title,
        'status': 'Không đạt cấu trúc',
        'header_row': '',
        'details': 'Không nhận diện đủ tiêu đề Họ tên, Ngày sinh và Năm tốt nghiệp trong 10 dòng đầu.',
    }


def inspect_workbook(path):
    """Return per-sheet structural findings using Pha 1's sheet-selection rules."""
    path = Path(path)
    try:
        workbook = load_workbook(path, read_only=True, data_only=False)
    except Exception as exc:
        return [{
            'file': path.name,
            'sheet': '',
            'status': 'Không đọc được file',
            'header_row': '',
            'details': f'{type(exc).__name__}: {exc}',
            'selected_by_phase1': 'Không',
        }]

    try:
        worksheets = workbook.worksheets
        if worksheets and worksheets[0].title.casefold() == 'data':
            selected = worksheets[:1]
        else:
            selected = worksheets
        selected_names = {worksheet.title for worksheet in selected}
        diagnoses = [_sheet_diagnosis(worksheet) for worksheet in worksheets]
        usable_names = {
            result['sheet'] for result in diagnoses
            if result['status'] == 'Đọc được' and result['sheet'] in selected_names
        }
        for result in diagnoses:
            result['file'] = path.name
            result['selected_by_phase1'] = (
                'Có' if result['sheet'] in selected_names else 'Không (Pha 1 bỏ qua sheet này)'
            )
            if (result['sheet'] in selected_names and not usable_names
                    and result['status'] != 'Đọc được'):
                result['details'] += ' Không có sheet hợp lệ nào được Pha 1 chọn.'
        return diagnoses
    finally:
        workbook.close()


def inspect_folder(folder):
    """Inspect Excel workbooks directly inside folder (not subfolders)."""
    folder = Path(folder)
    if not folder.is_dir():
        raise NotADirectoryError(f'Không tìm thấy thư mục: {folder}')
    files = sorted(
        (path for path in folder.iterdir()
         if path.is_file()
         and path.suffix.casefold() in SUPPORTED_EXTENSIONS
         and not path.name.startswith('~$')
         and not path.name.casefold().endswith('.partial.xlsx')
         and path.name.casefold() != REPORT_NAME.casefold()),
        key=lambda path: path.name.casefold(),
    )
    if not files:
        raise FileNotFoundError(
            f'Không tìm thấy file .xlsx hoặc .xlsm trong thư mục: {folder}'
        )
    return [result for path in files for result in inspect_workbook(path)]


def _file_summary(findings):
    by_file = {}
    for finding in findings:
        by_file.setdefault(finding['file'], []).append(finding)

    rows = []
    for filename, items in by_file.items():
        usable = any(
            item['status'] == 'Đọc được' and item['selected_by_phase1'] == 'Có'
            for item in items
        )
        if any(item['status'] == 'Không đọc được file' for item in items):
            status = 'Không đọc được'
        elif usable:
            status = 'Có sheet Pha 1 đọc được'
        else:
            status = 'Không có sheet Pha 1 đọc được'
        problem_sheets = [
            item['sheet'] for item in items
            if item['selected_by_phase1'] == 'Có'
            and item['status'] != 'Đọc được'
        ]
        rows.append([
            filename,
            status,
            ', '.join(problem_sheets),
            '; '.join(
                item['details'] for item in items
                if item['selected_by_phase1'] == 'Có'
                and item['status'] != 'Đọc được'
            ),
        ])
    return rows


def write_report(folder, findings):
    """Write a summary and sheet-level diagnostics workbook into folder."""
    folder = Path(folder)
    workbook = Workbook()
    summary = workbook.active
    summary.title = 'Tong_quan'
    summary.append(['Tệp', 'Kết luận', 'Sheet có vấn đề', 'Chi tiết'])
    for row in _file_summary(findings):
        summary.append(row)

    detail = workbook.create_sheet('Chi_tiet_sheet')
    detail.append([
        'Tệp', 'Sheet', 'Kết quả cấu trúc', 'Dòng tiêu đề / dòng gần nhất',
        'Pha 1 chọn sheet', 'Phân tích',
    ])
    for item in findings:
        detail.append([
            item['file'], item['sheet'], item['status'], item['header_row'],
            item['selected_by_phase1'], item['details'],
        ])

    for worksheet in workbook.worksheets:
        worksheet.freeze_panes = 'A2'
        worksheet.auto_filter.ref = worksheet.dimensions
        for column, width in {'A': 38, 'B': 32, 'C': 34, 'D': 110, 'E': 26, 'F': 110}.items():
            worksheet.column_dimensions[column].width = width
        for row in worksheet.iter_rows():
            for cell in row:
                cell.alignment = Alignment(wrap_text=True, vertical='top')

    destination = folder / REPORT_NAME
    temporary = folder / (Path(REPORT_NAME).stem + '.partial.xlsx')
    try:
        workbook.save(temporary)
        temporary.replace(destination)
    except Exception:
        workbook.close()
        raise
    workbook.close()
    return destination


def _choose_folder():
    root = tk.Tk()
    root.withdraw()
    try:
        folder = filedialog.askdirectory(
            parent=root, title='Chọn thư mục chứa các file Excel lỗi'
        )
        if not folder:
            return
        findings = inspect_folder(folder)
        report = write_report(folder, findings)
        bad_files = sum(
            row[1] != 'Có sheet Pha 1 đọc được' for row in _file_summary(findings)
        )
        messagebox.showinfo(
            'Đã kiểm tra cấu trúc',
            f'Đã kiểm tra {len(_file_summary(findings))} file; '
            f'{bad_files} file không có sheet Pha 1 đọc được.\n'
            f'Báo cáo chi tiết: {report}',
            parent=root,
        )
    except Exception as exc:
        messagebox.showerror('Không thể kiểm tra thư mục', str(exc), parent=root)
        raise
    finally:
        root.destroy()


def main(argv=None):
    parser = argparse.ArgumentParser(
        description='Phân tích lỗi cấu trúc file Excel bị từ chối ở Pha 1.'
    )
    parser.add_argument(
        'folder', nargs='?', help='Thư mục chứa các file .xlsx/.xlsm cần kiểm tra'
    )
    args = parser.parse_args(argv)
    if not args.folder:
        _choose_folder()
        return 0

    try:
        findings = inspect_folder(args.folder)
        report = write_report(args.folder, findings)
    except Exception as exc:
        print(f'Lỗi: {exc}', file=sys.stderr)
        return 1
    summaries = _file_summary(findings)
    bad_files = sum(row[1] != 'Có sheet Pha 1 đọc được' for row in summaries)
    print(f'Đã kiểm tra {len(summaries)} file; {bad_files} file không có sheet Pha 1 đọc được.')
    print(f'Báo cáo chi tiết: {report}')
    return 0


if __name__ == '__main__':
    raise SystemExit(main())
