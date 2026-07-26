# Project Overview

## Vì sao dự án được tạo

Việc tính Qdd/Qdư ban đầu thực hiện **thủ công**: đối chiếu danh sách lệnh điều độ với dữ liệu công tơ (CSV 6001, 6303) bằng tay, tính diện tích công suất theo từng chu kỳ 30 phút, rồi so sánh với dải dung sai để xác định công suất dư. Cách làm này tốn thời gian và dễ sai sót khi:

- Có nhiều lệnh trong ngày, đặc biệt khi lệnh mới ngắt ngang một ramp đang chạy dở.
- Ramp chạy qua nửa đêm, cần chuyển tiếp trạng thái sang ngày kế tiếp.
- Phải phân biệt cách chọn công suất hiệu lực khác nhau giữa lệnh SO, MO bình thường, và MO bị dừng sớm.

## Những vấn đề của cách tính thủ công

1. Không có công cụ audit lại từng bước tính — khó phát hiện sai lệch nằm ở đâu (nhập liệu, chọn sai cột thời điểm, hay công thức).
2. Không nhất quán giữa người tính khác nhau hoặc giữa các ngày khác nhau.
3. CSV xuất ra từ các nền tảng khác nhau (Windows tách sẵn 50 cột, macOS dán CSV dồn vào một cột) khiến việc nhập liệu thủ công dễ lỗi định dạng.

## Mục tiêu

- Tự động hoá toàn bộ pipeline: Danh sách lệnh + CSV → Ramp Engine → 48 chu kỳ Qdd/Qdư → Báo cáo.
- **Không hộp đen**: dữ liệu vào, kết quả ra và các bước trung gian đều hiển thị trên Sheet để người dùng nghiệp vụ tự kiểm tra.
- Có bộ quy tắc nghiệp vụ (Business Rules) và test case (UAT) tường minh, để bất kỳ thay đổi thuật toán nào trong tương lai đều có cơ sở đối chiếu.

## Lợi ích

- Giảm thời gian tính toán từ thủ công xuống còn vài thao tác bấm nút (15 nút chức năng trên sheet điều khiển).
- Giảm sai sót do nhầm cột (ví dụ BĐTH và thời điểm hoàn thành — xem R04) hay do đọc sai công suất hiệu lực theo loại lệnh (R01-R03).
- Có lịch sử tháng (`LICH_SU_THANG`, sheet `LS_...`) để tra cứu lại kết quả các ngày trước mà không cần tính lại.

## Đối tượng sử dụng

- Người vận hành/kỹ thuật viên nhà máy nhiệt điện, trực tiếp nhập danh sách lệnh và CSV công tơ hàng ngày.
- Người phụ trách nghiệp vụ/quản lý, đối chiếu và xác nhận Qdd/Qdư dùng cho báo cáo tháng.

## Kiến trúc hiện tại

```
Danh sách lệnh (LENH_GOC)  ─┐
CSV Qdc (công tơ 6001)      ├─→  Sheet của nhà máy (Google Sheets)
CSV Qmp (công tơ 6303)      ─┘        │
                                       ▼
                    Lọc lệnh hiệu lực, chọn P hiệu lực (R01-R03)
                                       │
                                       ▼
                    Ramp Engine dựng đường cong P(t) (R05-R06)
                                       │
                                       ▼
                    Tích phân diện tích theo 48 chu kỳ (R08)
                                       │
                                       ▼
                    KET_QUA (Qdd, Qdd_V, Qdc, Qmp, Qdư mỗi chu kỳ)
                                       │
                                       ▼
                    File báo cáo Excel/PDF  +  BAO_CAO_THANG
```

Toàn bộ phép tính nằm trong thư viện dùng chung `QDD-Core-Library`, có bộ test khoá lại từng quy tắc. Dữ liệu đầu vào và kết quả đều nằm ở các sheet nhìn thấy được, nên người dùng nghiệp vụ tự đối chiếu được từng bước — xem [05_System_Architecture.md](05_System_Architecture.md).

Xem chi tiết từng bước thuật toán ở [04_Algorithm_Specification.md](04_Algorithm_Specification.md) và schema đầy đủ ở [06_Database_Design.md](06_Database_Design.md).
