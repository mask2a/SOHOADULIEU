"""Copy source workbooks named in a Pha 1 error report."""
from pathlib import Path
import re
import shutil
from tkinter import Tk, filedialog, messagebox

from openpyxl import load_workbook


_WORKBOOK_NAME = re.compile(
    r'([^\\/:*?"<>|\r\n]+\.xls(?:x|m)?)(?=\s*(?::|$))',
    re.IGNORECASE,
)


def _reported_filenames(report_path):
    workbook = load_workbook(report_path, read_only=True, data_only=True)
    try:
        worksheet = workbook[workbook.sheetnames[0]]
        filenames = set()
        for row in worksheet.iter_rows(min_col=1, max_col=1, values_only=True):
            issue = str(row[0] or '').strip()
            if not issue or issue.casefold() == 'lỗi':
                continue
            for match in _WORKBOOK_NAME.finditer(issue):
                filenames.add(match.group(1))
        return sorted(filenames, key=str.casefold)
    finally:
        workbook.close()


def copy_reported_files(report_path, source_folder, destination_folder):
    """Copy reported workbooks from source_folder without overwriting existing files."""
    report_path = Path(report_path)
    source_folder = Path(source_folder)
    destination_folder = Path(destination_folder)
    if not report_path.is_file():
        raise FileNotFoundError(f'Không tìm thấy báo cáo lỗi: {report_path}')
    if not source_folder.is_dir():
        raise NotADirectoryError(f'Không tìm thấy thư mục file nguồn: {source_folder}')
    if not destination_folder.is_dir():
        raise NotADirectoryError(f'Không tìm thấy thư mục đích: {destination_folder}')
    if source_folder.resolve() == destination_folder.resolve():
        raise ValueError('Thư mục nguồn và thư mục đích phải khác nhau.')

    filenames = _reported_filenames(report_path)
    if not filenames:
        raise ValueError('Không tìm thấy tên file Excel trong cột "Lỗi" của báo cáo.')

    copied, missing, existing = [], [], []
    for filename in filenames:
        source = source_folder / filename
        destination = destination_folder / filename
        if not source.is_file():
            missing.append(filename)
        elif destination.exists():
            existing.append(filename)
        else:
            shutil.copy2(source, destination)
            copied.append(filename)
    return copied, missing, existing


def _choose_directory(root, title):
    return filedialog.askdirectory(parent=root, title=title)


def main():
    root = Tk()
    root.withdraw()
    try:
        report = filedialog.askopenfilename(
            parent=root,
            title='Chọn báo cáo lỗi của Pha 1',
            filetypes=[('Excel workbook', '*.xlsx *.xlsm')],
        )
        if not report:
            return
        source = _choose_directory(root, 'Chọn thư mục chứa các file Excel gốc')
        if not source:
            return
        destination = _choose_directory(root, 'Chọn thư mục lưu bản sao file lỗi')
        if not destination:
            return

        copied, missing, existing = copy_reported_files(report, source, destination)
        lines = [f'Đã sao chép: {len(copied)} file.']
        if missing:
            lines.append(f'Không tìm thấy trong thư mục nguồn: {len(missing)} file.')
            lines.extend(f'  - {name}' for name in missing[:10])
        if existing:
            lines.append(f'Đã bỏ qua vì trùng tên ở thư mục đích: {len(existing)} file.')
            lines.extend(f'  - {name}' for name in existing[:10])
        messagebox.showinfo('Hoàn tất', '\n'.join(lines), parent=root)
    except Exception as exc:
        messagebox.showerror('Không thể sao chép file lỗi', str(exc), parent=root)
        raise
    finally:
        root.destroy()


if __name__ == '__main__':
    main()
