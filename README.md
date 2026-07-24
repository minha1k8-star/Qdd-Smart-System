# QDD Smart System (QSS)

Công cụ tính toán và xuất báo cáo **Qdd** (công suất điều độ) và **Qdu/Qdư** (công suất dư) cho tổ máy nhiệt điện, dựa trên danh sách lệnh điều độ và dữ liệu CSV công tơ đo đếm.

> Đọc [docs/AI_CONTEXT.md](docs/AI_CONTEXT.md) trước nếu bạn là AI agent (Claude Code, Codex, ChatGPT...) đang tiếp nhận dự án này.

## Trạng thái hiện tại

Hệ thống chính thức đang chạy là công cụ **Excel + VBA** (bản v1.3.1), xem thư mục [`legacy/`](legacy/). Đây là công cụ đã dùng dữ liệu thực tế để kiểm chứng, **không phải bản demo**.

Repo này đang ở **Giai đoạn 1 — Documentation Foundation**: hệ thống hoá toàn bộ tri thức nghiệp vụ/kỹ thuật đã tích luỹ thành tài liệu chuẩn trước khi cân nhắc viết thêm mã nguồn mới. Xem tiến độ chi tiết ở [docs/PROJECT_STATUS.md](docs/PROJECT_STATUS.md) và định hướng dài hạn ở [ROADMAP.md](ROADMAP.md).

## Bài toán nghiệp vụ

1. Nhà máy nhận **lệnh điều độ** (SO — lệnh đóng máy, MO — lệnh mở/thay đổi công suất), mỗi lệnh có thời điểm bắt đầu thực hiện (BĐTH), công suất ra lệnh, công suất hoàn thành.
2. Công cụ dựng lại đường cong công suất theo thời gian trong ngày (48 chu kỳ × 30 phút) bằng mô hình **ramp** (tăng/giảm tải tuyến tính theo tốc độ cấu hình) + nội suy khi có lệnh mới ngắt ngang.
3. Từ đường cong đó tính **Qdd** (công suất điều độ trung bình mỗi chu kỳ) và quy đổi **Qdd_V**.
4. Đối chiếu với **Qdc** (từ CSV công tơ 6001) và **Qmp** (từ CSV công tơ 6303) để tính **Qdu** — phần công suất dư ngoài dải dung sai ±3%.

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
│   └── PROJECT_STATUS.md
├── legacy/
│   └── ...                 # Bản Excel/VBA v1.3.1 đang chạy chính thức
└── test-data/
    └── ...                 # Dữ liệu thực tế dùng để kiểm tra độ chính xác
```

## Công nghệ

- **Hiện tại**: Microsoft Excel Desktop (Windows/macOS) + VBA (1 module, ~4000 dòng), công thức Excel kiểm toán được.
- **Định hướng tương lai** (chưa triển khai): Google Apps Script + Web App + Google Sheets — xem [ROADMAP.md](ROADMAP.md).

## Đóng góp

Đọc [AGENTS.md](AGENTS.md) trước khi thay đổi bất kỳ quy tắc nghiệp vụ hoặc thuật toán nào.
