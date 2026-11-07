from dataclasses import dataclass, field
from datetime import date, datetime
from pathlib import PureWindowsPath
import re
import unicodedata

HEADERS = ('STT', 'Số hiệu bằng', 'Số vào sổ gốc cấp văn bằng', 'Tên văn bằng',
 'Họ, chữ đệm và tên', 'Số định danh cá nhân', 'Ngày, tháng, năm sinh',
 'Giới tính', 'Dân tộc', 'Quốc tịch', 'Nơi sinh', 'Tên trường đã học',
 'Năm tốt nghiệp', 'Hội đồng thi', 'Địa danh nơi cơ quan cấp bằng đặt trụ sở',
 'Ngày tháng năm cấp bằng', 'Tên cơ quan cấp bằng', 'Chức danh người ký bằng',
 'Họ, chữ đệm, tên người ký bằng', 'Tình trạng văn bằng', 'Xếp loại TN',
 'Ghi chú', 'Mã đơn vị (Mã moet)', 'Tên đơn vị')
IDENTITY = (4, 6, 12)
MATCH = 'Khớp hoàn toàn'
NORMALIZED = 'Khớp sau chuẩn hóa'
REVIEW = 'Cần xác minh'
ONLY_SOURCE = 'Chỉ có trong nguồn chuẩn'
ONLY_SCHOOL = 'Chỉ có trong file trường'

class AppError(Exception):
    """Expected failure safe to display to the operator."""

class Cancelled(AppError):
    pass

class StructureError(AppError):
    def __init__(self, issues):
        self.issues = issues
        super().__init__('File không đạt cấu trúc; xem bao_cao_loi.xlsx.')

def text(value):
    if value is None:
        return ''
    if isinstance(value, float) and value.is_integer():
        return str(int(value))
    return str(value)

def normalize(value):
    return ' '.join(unicodedata.normalize('NFC', text(value)).split()).upper()

def birth(value):
    if isinstance(value, (date, datetime)):
        return value.strftime('%d/%m/%Y'), True
    value = text(value).strip()
    if re.fullmatch(r'\d{4}', value):
        return value, False
    for fmt in ('%d/%m/%Y', '%d-%m-%Y', '%d.%m.%Y', '%Y-%m-%d', '%Y-%m-%d %H:%M:%S'):
        try:
            return datetime.strptime(value, fmt).strftime('%d/%m/%Y'), True
        except ValueError:
            pass
    return value, False

def short_location(file, sheet, row, origin=''):
    if origin:
        parts = [part.strip() for part in text(origin).split('|')]
        if len(parts) >= 3 and 'dòng' in parts[-1].casefold():
            filename = re.sub(r'^(?:tệp|file)\s*:\s*', '', parts[0], flags=re.IGNORECASE)
            sheet_name = re.sub(r'^sheet\s*:\s*', '', parts[1], flags=re.IGNORECASE)
            row_label = re.sub(r'^dòng\s*:\s*', 'dòng ', parts[-1], flags=re.IGNORECASE)
            return f'{PureWindowsPath(filename).name} | {sheet_name} | {row_label}'
    return f'{PureWindowsPath(file).name} | {sheet} | dòng {row}'

@dataclass(frozen=True)
class StudentRecord:
    values: tuple
    file: str
    sheet: str
    row: int
    origin: str = ''

    @property
    def keys(self):
        dob, _ = birth(self.values[6])
        return normalize(self.values[4]), dob, text(self.values[12]).strip()

    @property
    def valid(self):
        name, _, year = self.keys
        return bool(name) and birth(self.values[6])[1] and bool(re.fullmatch(r'\d{4}', year))

    @property
    def certificate(self):
        return normalize(self.values[1])

    @property
    def location(self):
        return short_location(self.file, self.sheet, self.row, self.origin)

@dataclass
class MatchResult:
    status: str
    source: StudentRecord | None = None
    school: StudentRecord | None = None
    reason: str = ''
    candidates: list[StudentRecord] = field(default_factory=list)
    matching_schools: list[StudentRecord] = field(default_factory=list)

@dataclass(frozen=True)
class Progress:
    message: str
    current: int = 0
    total: int = 0
