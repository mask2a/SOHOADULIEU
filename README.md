# Đối soát số hóa tốt nghiệp THCS — Offline

Ứng dụng desktop Python/Tkinter cho dữ liệu 1987–2005. Chỉ đối soát **Họ tên, Ngày sinh, Năm tốt nghiệp**. Các cột khác được giữ nguyên để tham khảo; số hiệu bằng chỉ hỗ trợ phân biệt người trùng khóa hoặc tìm ứng viên.

## Cấu trúc và chiều phụ thuộc

```text
main.py                    Điểm khởi động
sohoa/
  domain.py                DTO, lỗi nghiệp vụ, chuẩn hóa tên/ngày sinh
  matching.py              Quyết định ghép và phân loại, không phụ thuộc Excel/UI
  application.py           Hai interface nghiệp vụ: merge(), compare()
  storage.py               SQLite, chỉ mục tìm kiếm và truy vết
  excel.py                 Đọc/kiểm tra Excel, ánh xạ, ghi file an toàn
  report.py                Báo cáo ba sheet, màu ô, thống kê Pandas
  presentation.py          Hai tab ttk, worker, Queue, hủy và log
config/schools.csv          Bảng ánh xạ trường do người dùng duyệt
tests/
  test_matching.py         Kiểm thử chính sách ghép
  test_workflow.py         Kiểm thử Excel → SQLite → báo cáo
requirements.txt
requirements-build.txt
build.ps1
```

Presentation gọi application; application điều phối domain/matching và các adapter Excel/SQLite/report. Domain không import Tkinter, Pandas hoặc Openpyxl. Không tạo lớp interface rỗng chỉ để bọc một implementation. Seam với GUI là `JobContext(emit, cancelled)`, truyền `Progress` qua Queue. `StudentRecord` giữ 24 giá trị gốc và vị trí nguồn; `MatchResult` giữ trạng thái, hai phía và ứng viên.

## Tải mã nguồn bằng Git

Cài Git và Python có Tcl/Tk trên Windows. Mở PowerShell tại thư mục muốn lưu dự án, chạy:

```powershell
git clone https://github.com/mask2a/SOHOADULIEU.git
cd SOHOADULIEU
python -m venv .venv
.\.venv\Scripts\python.exe -m pip install -r requirements.txt
.\.venv\Scripts\python.exe main.py
```

Dữ liệu Excel và kết quả xuất không được đưa lên GitHub. Sau khi tải mã nguồn, chép các file Excel của bạn vào thư mục riêng (ví dụ `DATA1`), rồi chọn thư mục đó trong ứng dụng. Kết quả được tạo tại thư mục đầu ra đã chọn, trong `TACH_NAM_<ID>`.

Để lấy phiên bản mới, đóng ứng dụng, mở PowerShell trong thư mục `SOHOADULIEU` và chạy:

```powershell
git pull --ff-only origin main
.\.venv\Scripts\python.exe -m pip install -r requirements.txt
.\.venv\Scripts\python.exe main.py
```

Nếu Git báo có thay đổi cục bộ, chạy `git status` để kiểm tra và lưu các thay đổi trước khi cập nhật.

## Cài và chạy

```powershell
python -m venv .venv
.\.venv\Scripts\python.exe -m pip install -r requirements-build.txt
.\.venv\Scripts\python.exe main.py
```

Có thể chạy `python main.py` trong terminal: nếu Python hiện tại thiếu Pandas/Openpyxl, launcher tự chuyển sang `.venv` của dự án nếu có. VS Code cũng được cấu hình chọn interpreter này. Đường dẫn dữ liệu mặc định lấy theo vị trí mã nguồn/exe, không theo thư mục terminal; cấu hình cũ trỏ tới thư mục không tồn tại sẽ được dò lại. Exe trong `dist` hoặc `dist/fixed` có thể tìm `data` ở gốc dự án. Nếu mang exe sang máy khác, chọn thư mục dữ liệu thực tế bằng nút Chọn.

Cần bản Python có Tcl/Tk (chọn Tcl/Tk khi cài Python trên Windows). Cài thư viện là bước chuẩn bị; ứng dụng không sử dụng mạng khi chạy. Có thể chuẩn bị wheel trên máy có mạng rồi cài bằng `pip --no-index --find-links <thu_muc_wheel> -r requirements-build.txt` trên máy offline; wheel phải đúng phiên bản Python/Windows.

## Pha 1

Chọn thư mục dữ liệu (ví dụ `DATA1`) và bấm **Gộp và tách**. Chương trình đọc tất cả file `.xlsx`/`.xlsm` trực tiếp trong thư mục, gộp các bản ghi rồi tách theo **Năm tốt nghiệp ở cột M**. Không cần bảng ánh xạ trường. File khóa `~$` và file `.partial.xlsx` được bỏ qua.

Mỗi sheet dữ liệu cần đủ 24 cột chuẩn, tiêu đề nằm trong 10 dòng đầu (hỗ trợ dòng 1 của mẫu mới và dòng 4 của mẫu cũ). Sheet hướng dẫn/danh mục không có bộ tiêu đề này được bỏ qua. Bỏ dòng số cột, tiêu đề lặp và dòng trống. Giữ nguyên các bản ghi, kể cả bản ghi trùng.

Kết quả nằm trong `output/TACH_NAM_<ID>` mới, mỗi năm có một file như `1984.xlsx`, `1985.xlsx`… Chỉ xuất năm có dữ liệu; không giới hạn năm đọc vào. Dòng thiếu hoặc sai năm nằm trong `Chua_xac_dinh_nam.xlsx` và có báo cáo lỗi. Mỗi file có sheet `Data`, tiêu đề dòng 1, đủ 24 cột. Vượt giới hạn Excel thì chia phần `_001`, `_002`…

Trong mỗi năm, dữ liệu được sắp xếp tăng dần theo **Tên trường đã học** (cột L), chuẩn hóa hoa/thường và khoảng trắng khi so sánh. Trong cùng tên trường, sắp theo STT gốc; giữ nguyên nội dung các ô.

File tách có thêm cột **Nguồn dữ liệu** ở cuối (cột Y), ghi tên tệp, sheet và số dòng trong Excel nguồn, ví dụ `Tệp: truong.xlsx | Sheet: Data | Dòng: 6`. Cột `Ghi chú` gốc vẫn được giữ nguyên.

Các sheet xuất Data và THONGKE có độ rộng cột phù hợp, bật xuống dòng trong ô và tính chiều cao dòng theo độ dài nội dung khi xuất. Nếu sửa nội dung sau đó trong Excel, dùng AutoFit Row Height để cập nhật chiều cao.

Sheet **THONGKE** nằm sau **Data**, gồm **STT, TÊN TRƯỜNG, SỐ LƯỢNG**, thống kê theo cột `Tên trường đã học` trong chính file đó. Dòng cuối ghi tổng số trường và tổng số bản ghi. Tên trường được nhóm sau khi chuẩn hóa hoa/thường và khoảng trắng. Bản ghi trống tên trường vẫn được đếm ở dòng `(Chưa có tên trường)` nhưng không cộng vào số trường. Các tổng là số đã tính khi xuất file; nếu sửa Data thì cần chạy xuất lại để cập nhật.

Giữ `truy_vet.sqlite` cùng các file kết quả để đối soát truy về file/sheet/dòng gốc. `pham_vi.json` lưu danh sách đầu vào và các năm đã đọc. File nguồn sai cấu trúc được bỏ toàn bộ và ghi vào báo cáo lỗi, trạng thái đợt chạy là chưa đầy đủ. Phạm vi đối soát hỗ trợ 1984–2005.

## Pha 2

1. Chọn một hoặc nhiều phần nguồn chuẩn **của cùng một trường**. Không chọn file nhóm chưa xác định. Người vận hành phải xác nhận đúng trường vì ứng dụng không suy đoán trường từ tên file gửi.
2. Chọn file trường, dòng tiêu đề và bấm đọc tiêu đề. File có duy nhất một sheet, tên bất kỳ.
3. Xác nhận ánh xạ ba cột bắt buộc; số hiệu bằng tùy chọn. Tên cột không cần giống mẫu. Ánh xạ được nhớ theo bộ tiêu đề.
4. Chọn năm kiểm tra; xuất báo cáo.

File sai cấu trúc bị từ chối; `bao_cao_loi.xlsx` liệt kê các lỗi ánh xạ. Dòng thiếu/sai dữ liệu vẫn đi vào cần xác minh. Dòng có năm hợp lệ ngoài phạm vi bị loại và đếm trong `pham_vi.json`; dòng năm trống/sai được giữ để xác minh, không âm thầm bỏ.

Tên so sau NFC, chữ hoa, gọn khoảng trắng; không bỏ dấu, đảo từ hoặc sửa OCR. Ngày chuỗi hiểu ngày/tháng/năm, chấp nhận thêm ISO `YYYY-MM-DD`; không đoán ngày không hợp lệ. Chỉ có năm không đủ để tính khớp. Mã đọc dưới dạng văn bản; số 0 đầu chỉ phục hồi khi Excel có định dạng toàn số 0 rõ ràng, nếu không thì log để kiểm tra.

Ghép chính xác duy nhất theo bộ ba trước. Trùng bộ ba có thể phân biệt bằng số hiệu bằng duy nhất trong nhóm ở cả hai bên. Không được thì cần xác minh. Ứng viên có ít nhất hai thành phần định danh giống nhau (không tính hai ô trống) hoặc số hiệu bằng trùng. Không tự ghép gần đúng, không xóa trùng, không dùng số hiệu bằng để biến bộ ba sai thành khớp.

Nguồn chuẩn thiếu, trường có: ghi đề nghị bổ sung trong bằng chứng ứng viên, chưa xác nhận vì chưa đủ bộ ba. Nguồn có, trường thiếu: ghi trường thiếu thông tin. Ứng dụng không sửa nguồn, không có màn hình duyệt. Các ứng viên đã tham gia nhóm xác minh không đồng thời bị kết luận chỉ có một bên; một ứng viên có thể xuất hiện ở nhiều nhóm, nhưng không được coi là cặp đã ghép.

## Báo cáo

- `Tong_hop`: thống kê đầu sheet, tỷ lệ khớp/tổng dòng trường và độ bao phủ/tổng dòng chuẩn, kết quả chi tiết phía dưới.
- `Lech_thong_tin`: giá trị hai phía cạnh nhau, gồm bằng chứng ứng viên chưa xác nhận và khác biệt cách trình bày. Ô có giá trị gốc khác nhau tô vàng. **Ứng viên không phải cặp đã xác nhận.**
- `Thua_Thieu`: chỉ có trong chuẩn, chỉ có trong trường, cần xác minh; có lý do và nguồn dòng.

Khớp hoàn toàn và khớp sau chuẩn hóa đều đạt bộ ba. Các cột ngoài bộ ba không ảnh hưởng tỷ lệ. Không có mẫu số thì hiển thị rõ, không trả 100%. Thống kê đếm kết quả nghiệp vụ, không cộng dòng ứng viên. Báo cáo quá giới hạn được chia nhiều file, mỗi file vẫn ba sheet và có thống kê toàn đợt giống nhau.

Mỗi đợt có thư mục riêng và `TRANG_THAI.txt`. Chỉ dùng kết quả khi trạng thái hoàn tất (pha 1 có thể ghi chưa đầy đủ). Ghi `.partial.xlsx` rồi đổi tên khi thành công. Hủy có hiệu lực giữa các dòng/công đoạn; thao tác Openpyxl đang nén một file cần hoàn thành trước khi nhận hủy. Không ghi đè đầu vào. Log và cấu hình giao diện ở `%USERPROFILE%\THCS_Offline`; log chứa đường dẫn và thông tin xử lý, không gửi ra mạng.

## Kiểm thử và build

```powershell
.\.venv\Scripts\python.exe -m unittest discover -s tests -v
.\.venv\Scripts\python.exe tools/build.py
```

Hoặc chạy `build.ps1` nếu máy cho phép script PowerShell. Kết quả `dist/THCS_Offline.exe`. CSV ánh xạ là file ngoài do người dùng chọn, không nhúng danh mục học sinh vào exe.

Exe có kèm mẫu `schools.csv`. Ứng dụng ưu tiên file đã chọn còn tồn tại, sau đó `config/schools.csv` cạnh exe; nếu không có thì tạo bản có thể chỉnh sửa ở `%USERPROFILE%\THCS_Offline\schools.csv`. Cấu hình cũ trỏ tới file đã mất sẽ được thay bằng đường dẫn hợp lệ. Không ghi đè bảng ánh xạ đã có. Mẫu mặc định chỉ chứa tiêu đề; cần điền danh mục trường đã duyệt.

`sohoa/runtime.py` hỗ trợ cả bản Python đóng gói Tcl/Tk dạng ZIP: giải nén script thư viện vào `.runtime/` của dự án khi chạy nguồn/build. Trong exe, PyInstaller đóng gói Tcl/Tk và chọn thư mục tài nguyên tự động; không cần Python trên máy đích.

Trước đưa vào sử dụng, chạy bộ mẫu ẩn danh đã được duyệt và thử exe trên máy Windows mục tiêu không Python, ngắt mạng. Bộ test tự động không thay thế nghiệm thu này.

Kiểm tra nhanh exe trên máy đích: `dist\THCS_Offline.exe --self-test output\smoke`. Chế độ này khởi tạo Tkinter ẩn và chạy một ca Excel/SQLite/báo cáo với học sinh giả; thành công tạo `output\smoke\result.json`. Không dùng dữ liệu thật.

## Giới hạn cần biết

- Chưa cung cấp bộ chuyển mã TCVN3/VNI có màn hình duyệt. Cần chuyển mã và kiểm tra trên bản sao trước khi đưa vào; NFC không sửa được font cũ. Không tự sửa OCR.
- Với nhóm trùng cực lớn, danh sách ứng viên có thể lớn; Excel được stream nhưng một nhóm ứng viên vẫn cần RAM. Không có tuyên bố giới hạn RAM tuyệt đối.
- Không thể phát hiện hai bên cùng sai giống nhau. Không xác định được người trùng cả bộ ba nếu thiếu thông tin phân biệt.
