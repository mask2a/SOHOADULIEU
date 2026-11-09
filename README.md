# Đối soát số hóa tốt nghiệp THCS — Offline

Ứng dụng desktop Python/Tkinter xử lý dữ liệu số hóa tốt nghiệp THCS. Chỉ đối soát **Họ tên, Ngày sinh, Năm tốt nghiệp**; không giới hạn năm. Các cột khác được giữ nguyên để tham khảo; số hiệu bằng chỉ hỗ trợ phân biệt người trùng khóa hoặc tìm ứng viên.

## Cấu trúc và chiều phụ thuộc

```text
main.py                    Điểm khởi động
sohoa/
  domain.py                DTO, lỗi nghiệp vụ, chuẩn hóa tên/ngày sinh
  matching.py              Quyết định ghép và phân loại, không phụ thuộc Excel/UI
  application.py           Nghiệp vụ: merge(), compare(), compare_batch()
  storage.py               SQLite, chỉ mục tìm kiếm và truy vết
  excel.py                 Đọc/kiểm tra Excel, ánh xạ, ghi file an toàn
  report.py                Báo cáo đối soát theo nguồn, thống kê và hồ sơ phân loại
  presentation.py          Ba tab ttk, worker, Queue, hủy và log
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

Trên Windows, cài Python 3 có Tcl/Tk, sau đó mở PowerShell tại thư mục dự án. Tạo môi trường riêng và cài các thư viện:

```powershell
python -m venv .venv
.\.venv\Scripts\python.exe -m pip install -r requirements-build.txt
.\.venv\Scripts\python.exe main.py
```

Có thể dùng `requirements.txt` thay cho `requirements-build.txt` nếu chỉ chạy ứng dụng, không build exe. Nếu không muốn cài Python trên máy sử dụng, có thể build exe theo mục **Kiểm thử và build** rồi chuyển `dist\THCS_Offline.exe` sang máy Windows đích; cần thử nghiệm exe trên máy đích trước khi dùng dữ liệu thật.

Có thể chạy `python main.py` trong terminal: nếu Python hiện tại thiếu Pandas/Openpyxl, launcher tự chuyển sang `.venv` của dự án nếu có. VS Code cũng được cấu hình chọn interpreter này. Đường dẫn dữ liệu mặc định lấy theo vị trí mã nguồn/exe, không theo thư mục terminal; cấu hình cũ trỏ tới thư mục không tồn tại sẽ được dò lại. Exe trong `dist` hoặc `dist/fixed` có thể tìm `data` ở gốc dự án. Nếu mang exe sang máy khác, chọn thư mục dữ liệu thực tế bằng nút Chọn.

Cần bản Python có Tcl/Tk (chọn Tcl/Tk khi cài Python trên Windows). Cài thư viện là bước chuẩn bị; ứng dụng không sử dụng mạng khi chạy. Có thể chuẩn bị wheel trên máy có mạng rồi cài bằng `pip --no-index --find-links <thu_muc_wheel> -r requirements-build.txt` trên máy offline; wheel phải đúng phiên bản Python/Windows.

## Pha 1

Chọn thư mục dữ liệu (ví dụ `DATA1`) và bấm **Gộp và tách theo năm**. Chương trình đọc tất cả file `.xlsx`/`.xlsm` trực tiếp trong thư mục, gộp các bản ghi rồi tách theo **Năm tốt nghiệp ở cột M**. Không cần bảng ánh xạ trường. File khóa `~$` và file `.partial.xlsx` được bỏ qua.

Tiêu đề cần có trong 10 dòng đầu và phải nhận diện được **Họ tên, Ngày sinh, Năm tốt nghiệp**. Các cột được ánh xạ theo tên, không phụ thuộc thứ tự; chấp nhận một số tên thay thế thường gặp như `Họ và tên`, `Ngày sinh`, `Năm TN`, `Xếp loại tốt nghiệp`. Cột tiêu đề trống chỉ được suy ra theo vị trí khi nằm giữa các tiêu đề chuẩn liền kề; cột chuẩn thiếu tiêu đề sẽ để trống và được ghi log. Dữ liệu ở cột chưa nhận diện không bị gán đoán vào cột chuẩn; nội dung được giữ nguyên trong cột phụ riêng của kết quả. Sheet không có đủ ba tiêu đề định danh bị từ chối, sheet hướng dẫn/danh mục không phù hợp được bỏ qua. Bỏ dòng số cột, tiêu đề lặp và dòng trống. Giữ nguyên các bản ghi, kể cả bản ghi trùng. Vùng Excel chỉ có định dạng nhưng không có nội dung không làm tăng số cột phải quét; dữ liệu thật ở các cột xa bên phải vẫn được giữ lại. Log chỉ ghi file đang đọc và giai đoạn xử lý; phần trăm tiến độ hiển thị riêng bên phải thanh tiến trình, cập nhật đến một chữ số thập phân khi đọc file và xuất dữ liệu theo năm. Giao diện ưu tiên phần trăm mới nhất để không bị các mốc cũ trong hàng đợi làm đứng chỉ báo. Cảnh báo mã dạng số được tổng hợp theo cột thay vì ghi lặp từng dòng; cảnh báo dữ liệu ở các cột chưa nhận diện được gộp dải cột liên tiếp để log không bị kéo dài bất thường.

Kết quả nằm trong `output/TACH_NAM_<ID>` mới, mỗi năm có một file như `1984.xlsx`, `1985.xlsx`… Chỉ xuất năm có dữ liệu; không giới hạn năm đọc vào. Dòng thiếu hoặc sai năm nằm trong `Chua_xac_dinh_nam.xlsx` và có báo cáo lỗi. Mỗi file có sheet `Data`, tiêu đề dòng 1, 24 cột chuẩn và cột `Nguồn dữ liệu`, cùng cột `Dữ liệu cột nguồn chưa ánh xạ` để bảo toàn nội dung từ cột chưa nhận diện. Vượt giới hạn Excel thì chia phần `_001`, `_002`…

Báo cáo lỗi của Pha 1 có tên `bao_cao_loi.xlsx` (định dạng `.xlsx`), trong đó cột **Lỗi** ghi tên file nguồn không đọc được. Để gom các file nguồn này vào thư mục riêng, chạy `.\.venv\Scripts\python.exe copy_error_files.py`, rồi chọn báo cáo lỗi, thư mục chứa file Excel gốc và thư mục đích. Script chỉ sao chép, không sửa file nguồn; nếu tên file đã có ở thư mục đích thì bỏ qua, không ghi đè.

Để phân tích vì sao các file đã gom bị lỗi cấu trúc, chạy `.\.venv\Scripts\python.exe inspect_error_structure.py` rồi chọn thư mục chứa chúng. Có thể truyền đường dẫn thư mục trực tiếp: `.\.venv\Scripts\python.exe inspect_error_structure.py "C:\duong dan\thu muc"`. Báo cáo `phan_tich_loi_cau_truc.xlsx` được tạo trong thư mục đó, gồm kết luận theo file và chi tiết theo sheet (các tiêu đề định danh không nhận diện được, tên thay thế được ánh xạ, cột phụ thiếu tiêu đề, file không đọc được, hoặc sheet bị Pha 1 bỏ qua). Công cụ chỉ kiểm tra cấu trúc tiêu đề, không sửa file và không xác nhận nội dung từng bản ghi.

Trong mỗi năm, dữ liệu được sắp xếp tăng dần theo **Tên trường đã học** (cột L), chuẩn hóa hoa/thường và khoảng trắng khi so sánh. Trong cùng tên trường, sắp theo STT gốc; giữ nguyên nội dung các ô.

File tách có thêm cột **Nguồn dữ liệu** (cột Y), ghi tên tệp, sheet và số dòng trong Excel nguồn, ví dụ `Tệp: truong.xlsx | Sheet: Data | Dòng: 6`. Cột **Dữ liệu cột nguồn chưa ánh xạ** (cột Z) giữ nguyên giá trị từ các cột chưa nhận diện, kèm tiêu đề nguồn hoặc số cột nếu tiêu đề trống. Cột `Ghi chú` gốc vẫn được giữ nguyên.

Các sheet xuất Data và THONGKE có độ rộng cột phù hợp, bật xuống dòng trong ô và tính chiều cao dòng theo độ dài nội dung khi xuất. Nếu sửa nội dung sau đó trong Excel, dùng AutoFit Row Height để cập nhật chiều cao.

Sheet **THONGKE** nằm sau **Data**, gồm **STT, TÊN TRƯỜNG, SỐ LƯỢNG**, thống kê theo cột `Tên trường đã học` trong chính file đó. Dòng cuối ghi tổng số trường và tổng số bản ghi. Tên trường được nhóm sau khi chuẩn hóa hoa/thường và khoảng trắng. Bản ghi trống tên trường vẫn được đếm ở dòng `(Chưa có tên trường)` nhưng không cộng vào số trường. Các tổng là số đã tính khi xuất file; nếu sửa Data thì cần chạy xuất lại để cập nhật.

Giữ `truy_vet.sqlite` cùng các file kết quả để đối soát truy về file/sheet/dòng gốc. `pham_vi.json` lưu danh sách đầu vào và các năm đã đọc. File nguồn sai cấu trúc được bỏ toàn bộ và ghi vào báo cáo lỗi, trạng thái đợt chạy là chưa đầy đủ. Pha 1 và Pha 2 không giới hạn năm; báo cáo chỉ đối soát dữ liệu trong các file đã chọn.

## Pha 2

1. Ở tab **Đối soát và báo cáo**, chọn đúng một file nguồn chuẩn. Chương trình tự tìm bộ 24 cột chuẩn trong 10 dòng đầu, nên hỗ trợ mẫu có tiêu đề ở dòng 4 cũng như file đã tách có tiêu đề ở dòng 1.
2. Chọn file trường gửi cần đối soát. Với file Pha 1 có sheet `Data` và `THONGKE`, chọn file năm cần so; với file khác, chọn dòng tiêu đề thích hợp.
3. Bấm đọc tiêu đề, xác nhận ánh xạ **Họ tên**, **Ngày sinh**, **Năm tốt nghiệp**; ánh xạ **Số hiệu bằng** là tùy chọn. Tên cột file trường không cần giống mẫu. Ánh xạ được nhớ theo bộ tiêu đề.
4. Để trống **Năm kiểm tra** để lấy mọi năm có trong hai file. Hoặc nhập một hay nhiều năm, phân cách bằng dấu phẩy; có thể dùng khoảng năm, ví dụ `1983, 1987, 2026-2030`. Chọn thư mục kết quả rồi bấm **Xuất báo cáo đối soát**.
5. Khi ứng dụng báo hoàn tất, mở thư mục `Doi_soat_<thời gian>_<mã>` mới nhất bên trong thư mục kết quả. Chỉ dùng workbook khi `TRANG_THAI.txt` ghi **HOÀN TẤT**. Nếu báo chưa hoàn tất hoặc đã hủy, không dùng các file đang dang dở.

Pha 2 phù hợp khi cần so một cặp file, kể cả khi tên file không theo quy ước năm. Báo cáo là **một workbook cho toàn phạm vi năm đã chọn**, tên lấy theo file nguồn và thêm `_Check.xlsx`. Cùng thư mục đợt chạy có `import/` với các workbook được tách riêng theo từng năm hợp lệ.

File sai cấu trúc bị từ chối; `bao_cao_loi.xlsx` liệt kê các lỗi ánh xạ. Dòng thiếu/sai dữ liệu vẫn đi vào cần xác minh. Nếu chọn phạm vi năm, dòng có năm hợp lệ ngoài phạm vi được đếm trong `pham_vi.json`; dòng năm trống/sai được giữ để xác minh, không âm thầm bỏ.

Tên so sau NFC, chữ hoa, gọn khoảng trắng; không bỏ dấu, đảo từ hoặc sửa OCR. Ngày chuỗi hiểu ngày/tháng/năm, chấp nhận thêm ISO `YYYY-MM-DD`; không đoán ngày không hợp lệ. Chỉ có năm không đủ để tính khớp. Mã đọc dưới dạng văn bản; số 0 đầu chỉ phục hồi khi Excel có định dạng toàn số 0 rõ ràng, nếu không thì cảnh báo số ô cần kiểm tra được tổng hợp theo cột.

Ghép chính xác theo bộ ba trước. Nếu một hồ sơ nguồn có nhiều dòng file trường trùng bộ ba, các dòng đó được tính là khớp thông tin và ghi rõ số lượng cùng từng vị trí trong ghi chú; hồ sơ nguồn chỉ tính một lần. Nếu nhiều hồ sơ nguồn cùng bộ ba thì chỉ ghép khi số hiệu bằng phân biệt duy nhất ở cả hai bên, nếu không cần xác minh. Ứng viên có ít nhất hai thành phần định danh giống nhau (không tính hai ô trống) hoặc số hiệu bằng trùng. Không tự ghép gần đúng, không xóa trùng, không dùng số hiệu bằng để biến bộ ba sai thành khớp.

Nguồn chuẩn thiếu, trường có: ghi đề nghị bổ sung trong bằng chứng ứng viên, chưa xác nhận vì chưa đủ bộ ba. Nguồn có, trường thiếu: ghi trường thiếu thông tin. Ứng dụng không sửa nguồn, không có màn hình duyệt. Các ứng viên đã tham gia nhóm xác minh không đồng thời bị kết luận chỉ có một bên; một ứng viên có thể xuất hiện ở nhiều nhóm, nhưng không được coi là cặp đã ghép.

### Chuẩn bị file và dùng kết quả Pha 2

- Giữ nguyên file nguồn và file trường; nên đóng chúng trong Excel trước khi chạy. Đối soát chỉ đọc dữ liệu, không sửa workbook đầu vào.
- File nguồn chuẩn cần có bộ cột chuẩn. Pha 2 tự dò tiêu đề nguồn trong 10 dòng đầu; file trường có thể dùng tên cột khác nếu ánh xạ đúng. Chọn chính xác dòng tiêu đề file trường trước khi đọc tiêu đề.
- Trong ánh xạ, bắt buộc chọn đúng cột **Họ tên**, **Ngày sinh**, **Năm tốt nghiệp**. **Số hiệu bằng** có thể để `(Không dùng)` nếu không có. Ánh xạ sai có thể khiến kết quả phân loại sai; kiểm tra lại tên cột và một vài dòng mẫu trước khi chạy.
- Nếu chọn năm cụ thể, hồ sơ có năm hợp lệ ngoài phạm vi được loại khỏi lần đối soát và số lượng được ghi trong `pham_vi.json`. Hồ sơ không có năm hợp lệ vẫn được giữ để rà soát, không bị âm thầm bỏ.
- Báo cáo có `TONGHOP`, `DULIEU_OK`, `DULIEU_CHI_CO_TRONG_SGD`, `DULIEU_KHONGCO_TRONG_SGD`, `DULIEU_LECH_THONG_TIN`, `DULIEU_CAN_XAC_MINH`. Hai sheet cuối là hai góc nhìn; một hồ sơ có thể xuất hiện trong cả hai và không có nghĩa là hai lỗi độc lập.
- Trong `import/import_<năm>.xlsx`, dữ liệu lấy từ workbook nguồn Sở, giữ lại những dòng nguồn đã khớp để nhập sang hệ thống khác. Kiểm tra cấu trúc và các cột bắt buộc của phần mềm nhận trước khi nhập.
- Trong `import/kiemtralai/import_khong_co_trong_SGD<năm>.xlsx`, hồ sơ chỉ có ở file trường được ánh xạ vào cấu trúc nguồn Sở; cột không ánh xạ được để trống và `Nguồn dữ liệu` ghi tên file, sheet, dòng gốc. Đây là dữ liệu cần rà soát trước khi dùng, không phải kết luận đã được Sở xác nhận.
- Chỉ import sau khi đã kiểm tra đúng năm, số dòng, định dạng cột và danh tính file nguồn. Giữ nguyên thư mục kết quả cùng `truy_vet.sqlite` và `pham_vi.json` để truy lại nguồn gốc khi cần.

## Pha 3 — đối soát theo lô

1. Ở tab **Đối soát theo lô**, chọn thư mục nguồn chuẩn và thư mục trường gửi.
2. Nguồn cần có tên dạng `THCS.CQ.1987_output.xlsx`; file trường tương ứng cần tên `1987.xlsx`. Chương trình chỉ xử lý các năm có đủ cả hai file; năm chỉ có ở một thư mục được bỏ qua.
3. Bấm **Đối soát các năm trùng**. Mỗi năm được xử lý độc lập, tối đa 4 năm chạy đồng thời; kết quả nằm chung trong thư mục đợt chạy mới với tên như `1987_Check.xlsx`. Mỗi file sử dụng tiêu đề ở dòng 1 và tự ánh xạ các cột theo tên. Các sheet lần lượt là `TONGHOP`, `DULIEU_OK`, `DULIEU_CHI_CO_TRONG_SGD`, `DULIEU_KHONGCO_TRONG_SGD`, `DULIEU_LECH_THONG_TIN`, `DULIEU_CAN_XAC_MINH`.

`pham_vi.json` liệt kê các năm đã đối soát, các năm bị bỏ qua và lỗi nếu có. Nếu một năm lỗi cấu trúc, các năm còn lại vẫn được xử lý; đợt chạy ghi trạng thái chưa đầy đủ và tạo `bao_cao_loi_<năm>.xlsx` cho năm lỗi.

Pha 3 dành cho nhiều cặp file đã được tách theo năm. Tên file phải khớp mẫu: file nguồn `THCS.CQ.<năm>_output.xlsx`, file trường `<năm>.xlsx`; phần `<năm>` là bốn chữ số. Chương trình chỉ ghép đúng năm trùng, không đoán tên file gần giống và không đối soát các năm chỉ có ở một phía. Các cặp được xử lý độc lập nên lỗi cấu trúc của một năm không ngăn các năm còn lại chạy; tuy vậy trạng thái chung sẽ là chưa đầy đủ và cần xem `pham_vi.json` cùng `bao_cao_loi_<năm>.xlsx`.

Kết quả Pha 3 nằm trong thư mục `Doi_soat_<thời gian>_<mã>` mới. Mỗi năm có báo cáo `<năm>_Check.xlsx` và hai file import trong `import/` cùng `import/kiemtralai/`, theo cùng quy tắc nội dung như Pha 2. Dùng Pha 3 khi muốn báo cáo và import tách riêng theo từng năm; dùng Pha 2 khi cần một workbook tổng hợp cho một cặp file.

`TONGHOP` ghi phạm vi năm (Pha 2) hoặc năm đang xử lý (Pha 3), tổng hồ sơ hai phía, khớp hoàn toàn, khớp sau chuẩn hóa, tổng hồ sơ nguồn khớp và hai tỷ lệ: số hồ sơ trường khớp trên tổng hồ sơ trường gửi; số hồ sơ nguồn khớp trên tổng hồ sơ nguồn. `Dữ liệu chưa khớp` cộng số hồ sơ chưa khớp ở cả hai phía (mỗi hồ sơ mỗi phía tính riêng); `Dữ liệu cần kiểm tra` là tổng hồ sơ nguồn Sở trừ tổng hồ sơ khớp. Các dòng chỉ có ở Sở và chỉ có ở trường gửi được thống kê riêng. Nếu một tỷ lệ không có mẫu số, báo cáo ghi rõ thay vì hiển thị 100%.

`DULIEU_OK` giữ bảng đối chiếu hai phía và vị trí từng nguồn. `DULIEU_CHI_CO_TRONG_SGD` chỉ ghi hồ sơ nguồn không có hồ sơ trường tương ứng. `DULIEU_KHONGCO_TRONG_SGD` chỉ ghi hồ sơ trường gửi không có hồ sơ nguồn tương ứng. Hai sheet cuối là hai góc nhìn của trường hợp chưa phân định: `DULIEU_LECH_THONG_TIN` nhìn từ hồ sơ nguồn Sở, nêu thông tin lệch và ứng viên để rà soát; `DULIEU_CAN_XAC_MINH` nhìn từ hồ sơ trường gửi, nêu trường hợp chưa thể xác nhận hồ sơ nguồn cụ thể hoặc dữ liệu định danh thiếu/sai. Hai sheet có thể phản ánh cùng một nhóm hồ sơ; không tự kết luận dữ liệu sai hoặc hồ sơ thừa.

Ngoài báo cáo, mỗi năm tạo `import/import_<năm>.xlsx` bằng cách giữ nguyên cấu trúc workbook nguồn Sở và chỉ giữ lại các dòng nguồn đã khớp. Tạo thêm `import/kiemtralai/import_khong_co_trong_SGD<năm>.xlsx`; dữ liệu chỉ có trong file trường được ánh xạ vào cấu trúc workbook nguồn Sở, các cột không có dữ liệu tương ứng để trống và cột `Nguồn dữ liệu` ghi file/sheet/dòng gửi lên. Nếu workbook nguồn có sheet `THONGKE`, thống kê được tính lại theo các dòng trong file import.

## Báo cáo đối soát

Pha 2 tạo một workbook báo cáo cho phạm vi năm đã chọn, còn Pha 3 tạo một workbook cho mỗi năm. Cả hai dùng cùng sáu sheet:

- **TONGHOP**: tổng hồ sơ hai phía, khớp hoàn toàn, khớp sau chuẩn hóa, tổng hồ sơ nguồn khớp và hai tỷ lệ: số hồ sơ trường khớp trên tổng hồ sơ trường gửi; số hồ sơ nguồn khớp trên tổng hồ sơ nguồn. “Dữ liệu chưa khớp” cộng số hồ sơ chưa khớp ở cả hai phía (mỗi hồ sơ mỗi phía tính riêng); “Dữ liệu cần kiểm tra” là tổng hồ sơ nguồn Sở trừ tổng hồ sơ khớp. Các dòng chỉ có ở Sở và chỉ có ở trường gửi được thống kê riêng.
- **DULIEU_OK**: các hồ sơ được xác nhận trùng theo họ tên, ngày sinh, năm tốt nghiệp. Nếu có nhiều dòng file trường cùng khớp với một hồ sơ nguồn duy nhất, ghi chú cho biết số lượng và vị trí từng dòng.
- **DULIEU_CHI_CO_TRONG_SGD**: hồ sơ có trong nguồn chuẩn nhưng không có hồ sơ trường tương ứng.
- **DULIEU_KHONGCO_TRONG_SGD**: hồ sơ file trường không có hồ sơ tương ứng trong nguồn chuẩn.
- **DULIEU_LECH_THONG_TIN**: nhìn từ phía nguồn chuẩn; mỗi hồ sơ nguồn chưa được kết luận khớp chỉ liệt kê một lần. Cột “Thông tin lệch” nêu những khác biệt giữa nguồn và các ứng viên trường để hỗ trợ rà soát.
- **DULIEU_CAN_XAC_MINH**: nhìn từ phía file trường; hồ sơ trường có ứng viên nhưng chưa thể xác nhận ghép với hồ sơ nguồn cụ thể, hoặc thiếu/sai thông tin định danh. Không tự kết luận các hồ sơ này là khớp hay thừa.

`DULIEU_LECH_THONG_TIN` và `DULIEU_CAN_XAC_MINH` là hai góc nhìn của quy trình (nguồn chuẩn và file trường); chúng có thể phản ánh cùng một nhóm chưa phân định được, không nhất thiết là hai loại lỗi riêng biệt. “Lệch thông tin” là dấu hiệu để đối chiếu, không tự nó chứng minh dữ liệu trường sai. Vị trí trong báo cáo chỉ ghi tên tệp, sheet và dòng, không ghi đường dẫn thư mục. Khi một sheet vượt giới hạn Excel, báo cáo được chia phần với hậu tố số.

Khớp hoàn toàn và khớp sau chuẩn hóa đều đạt bộ ba. Các cột ngoài bộ ba không ảnh hưởng tỷ lệ. Không có mẫu số thì hiển thị rõ, không trả 100%. Báo cáo quá giới hạn được chia nhiều file, mỗi file có đủ các sheet và thống kê toàn đợt giống nhau.

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
