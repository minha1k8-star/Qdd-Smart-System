# QDD Smart System (QSS)

Công cụ tính toán và xuất báo cáo **Qdd** (công suất điều độ) và **Qdu/Qdư** (công suất dư) cho tổ máy nhiệt điện, dựa trên danh sách lệnh điều độ và dữ liệu CSV công tơ đo đếm.

> Đọc [docs/AI_CONTEXT.md](docs/AI_CONTEXT.md) trước nếu bạn là AI agent (Claude Code, Codex, ChatGPT...) đang tiếp nhận dự án này.

## Trạng thái hiện tại

Bản **Google Sheets + Apps Script** đã chạy được và **đã kiểm chứng bằng dữ liệu vận hành thật**: ngày 17/07 và 18/07/2026 tổ S1 khớp **tuyệt đối (0,0000 MW / 48 chu kỳ)** với bảng tính tay độc lập. Xem [docs/15_Accuracy_Validation_2026-07.md](docs/15_Accuracy_Validation_2026-07.md).

Bản **Excel + VBA v1.3.1** (trong [`legacy/`](legacy/)) là hệ thống gốc đã dùng nhiều tháng, giữ lại làm chuẩn đối chiếu.

Tiến độ chi tiết: [docs/PROJECT_STATUS.md](docs/PROJECT_STATUS.md) · [ROADMAP.md](ROADMAP.md) · [CHANGELOG.md](CHANGELOG.md)

## Triển khai cho nhà máy của bạn

👉 **[docs/16_Huong_Dan_Trien_Khai.md](docs/16_Huong_Dan_Trien_Khai.md)** — hướng dẫn từng bước lấy code về tự tạo Google Sheets + Apps Script riêng (kèm cách xử lý các lỗi thường gặp khi cài đặt).

Hai lựa chọn: **copy Sheet mẫu** (~5 phút, không cần cài gì) hoặc **tự dựng từ GitHub** (~20 phút, tự chủ hoàn toàn mã nguồn).

Hồ sơ sáng kiến / phương án kỹ thuật: [docs/17_Thuyet_Minh_Sang_Kien.md](docs/17_Thuyet_Minh_Sang_Kien.md)

## Bài toán nghiệp vụ

1. Nhà máy nhận **lệnh điều độ** (SO — lệnh đóng máy, MO — lệnh mở/thay đổi công suất), mỗi lệnh có thời điểm bắt đầu thực hiện (BĐTH), công suất ra lệnh, công suất hoàn thành.
2. Công cụ dựng lại đường cong công suất theo thời gian trong ngày (48 chu kỳ × 30 phút) bằng mô hình **ramp** (tăng/giảm tải tuyến tính theo tốc độ cấu hình) + nội suy khi có lệnh mới ngắt ngang.
3. Từ đường cong đó tính **Qdd** (công suất điều độ trung bình mỗi chu kỳ) và quy đổi **Qdd_V**.
4. Đối chiếu với **Qdc** và **Qmp** (từ 2 CSV công tơ, mã công tơ cấu hình theo từng tổ máy/nhà máy) để tính **Qdu** — phần công suất dư ngoài dải dung sai ±3%.

Chi tiết đầy đủ: [docs/03_Business_Rules.md](docs/03_Business_Rules.md) (quy tắc nghiệp vụ) và [docs/04_Algorithm_Specification.md](docs/04_Algorithm_Specification.md) (thuật toán).

## Cấu trúc repo

```
Qdd-Smart-System/
├── README.md
├── AGENTS.md              # Quy tắc làm việc cho AI/con người
├── CLAUDE.md               # Điểm vào cho Claude Code
├── CHANGELOG.md
├── ROADMAP.md
├── docs/
│   ├── AI_CONTEXT.md              # Đọc đầu tiên
│   ├── 00_Project_Overview.md
│   ├── 03_Business_Rules.md
│   ├── 04_Algorithm_Specification.md
│   ├── 06_Database_Design.md
│   ├── 09_Test_Cases.md
│   ├── 14_Knowledge_Transfer.md
│   ├── 15_Accuracy_Validation_2026-07.md  # Đối chiếu với dữ liệu vận hành thật
│   ├── 16_Huong_Dan_Trien_Khai.md         # Triển khai cho nhà máy mới
│   ├── 17_Thuyet_Minh_Sang_Kien.md        # Bản thảo hồ sơ sáng kiến
│   └── PROJECT_STATUS.md
├── legacy/
│   └── ...                 # Bản Excel/VBA v1.3.1 đang chạy chính thức
├── src/
│   ├── QDD-Core-Library/     # Apps Script Library (Giai đoạn 2/3) - xem src/QDD-Core-Library/README.md
│   └── NhaMay-Mau-Template/  # Sheet mẫu nhà máy đầu tiên, gắn Library - xem src/NhaMay-Mau-Template/README.md
├── tools/
│   └── reference_engine/   # Bản tái hiện thuật toán bằng Python để kiểm tra/đối chiếu
└── test-data/
    └── ...                 # Dữ liệu thực tế dùng để kiểm tra độ chính xác (gitignore)
```

## Công nghệ

- **Đang dùng**: Google Sheets + Apps Script, với **thư viện dùng chung** `QDD-Core-Library` chứa toàn bộ thuật toán — sửa một lần, mọi nhà máy cập nhật được. Lý do chọn kiến trúc này (thay vì Web App tập trung): [docs/05_System_Architecture.md](docs/05_System_Architecture.md).
- **Bản gốc**: Microsoft Excel + VBA (1 module ~4000 dòng), công thức Excel kiểm toán được — giữ trong `legacy/` làm chuẩn đối chiếu.
- **Kiểm thử**: bộ test chạy bằng Node (`src/QDD-Core-Library/tests/run_tests.js`), không cần Google và không cần dữ liệu thật.

## Đóng góp

Đọc [AGENTS.md](AGENTS.md) trước khi thay đổi bất kỳ quy tắc nghiệp vụ hoặc thuật toán nào.
