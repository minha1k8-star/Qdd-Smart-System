# Hướng dẫn triển khai cho nhà máy mới

Tài liệu này dành cho người muốn **lấy code từ GitHub về tự tạo Google Sheets + Apps Script riêng** cho nhà máy của mình.

Có **2 cách**. Đọc mục "Chọn cách nào" bên dưới trước khi bắt tay.

---

## Chọn cách nào

| | Cách A — Copy Sheet mẫu | Cách B — Tự dựng từ GitHub |
|---|---|---|
| Thời gian | ~5 phút | ~20 phút |
| Cần cài phần mềm | Không | Node.js + clasp |
| Tự chủ mã nguồn | Phụ thuộc bản gốc | Hoàn toàn tự chủ |
| Phù hợp khi | Dùng thử, hoặc nhà máy chỉ cần chạy | Muốn tự sửa code, tự quản lý phiên bản |

Cả 2 cách đều ra kết quả giống nhau về chức năng.

---

## Cách A — Copy Sheet mẫu (nhanh nhất)

1. Mở Sheet mẫu (xin link từ người quản trị hệ thống).
2. **Tệp → Tạo bản sao** (File → Make a copy). Bản sao mang theo toàn bộ mã nguồn Apps Script.
3. Mở bản sao vừa tạo, tải lại trang (F5).
4. Menu **QDD Smart System → Thiết lập sheet**.
5. Vào sheet `CAI_DAT`, sửa lại toàn bộ cho đúng nhà máy của bạn:
   - Tên nhà máy, tốc độ ramp, hệ số Qdd_V, dung sai
   - **Mã công tơ Qdc/Qmp cho từng tổ máy** — bắt buộc phải sửa, giá trị mặc định chỉ đúng cho Duyên Hải 1
   - Nhãn báo cáo (vd `S1DH1` → đổi thành ký hiệu nhà máy bạn)
6. Đọc sheet `HUONG_DAN` (ngoài cùng bên trái) để biết cách dùng.

> **Điều kiện**: thư viện `QDD-Core-Library` phải được chia sẻ ở mức "Bất kỳ ai có đường liên kết — Người xem". Nếu chưa, bản sao sẽ báo lỗi không truy cập được thư viện. Xem mục "Chia sẻ thư viện" bên dưới.

---

## Cách B — Tự dựng từ GitHub

### B0. Chuẩn bị máy

Cần **Node.js**. Kiểm tra bằng:

```bash
node --version
```

> ⚠️ **Quan trọng**: `clasp` (công cụ của Google) **không chạy được với Node.js phiên bản quá mới** — Node 22+ thường lỗi `Error retrieving access token: FetchError: ... Premature close` ngay ở bước đăng nhập. Dùng **Node 20** để chắc chắn.
>
> Trên macOS, cài thêm Node 20 mà không ảnh hưởng bản đang có:
> ```bash
> brew install node@20
> ```
> rồi thêm tiền tố `PATH="/opt/homebrew/opt/node@20/bin:$PATH"` trước mỗi lệnh `npx clasp ...`.

### B1. Lấy code về

```bash
git clone https://github.com/minha1k8-star/Qdd-Smart-System.git
cd Qdd-Smart-System
```

### B2. Bật Apps Script API cho tài khoản Google

Mở **https://script.google.com/home/usersettings** → bật **Google Apps Script API**. Không bật thì bước tạo project sẽ báo lỗi *"User has not enabled the Apps Script API"*.

### B3. Đăng nhập clasp

```bash
cd src/QDD-Core-Library
npm install
npx clasp login
```

Trình duyệt mở ra → chọn tài khoản Google → **Allow**.

> Nếu gặp lỗi `Premature close`: đang dùng Node quá mới. Chạy lại bằng Node 20 như mục B0.
> Đừng dùng `--no-localhost` — Google đã bỏ hỗ trợ cách đó, sẽ báo `Lỗi 400: invalid_request`.

### B4. Thư viện dùng chung hay tự tạo?

**Phương án 1 — Dùng chung thư viện có sẵn** (khuyến nghị khi nhân rộng nhiều nhà máy):
Bỏ qua bước tạo thư viện, dùng luôn Script ID sẵn có:
```
10_vjTSgVjZodA7xTkJ_qJaGom3JDx_tnYE0YgWA_cphh1Q7g_lTKMLUO
```
Lợi ích: sửa lỗi thuật toán một lần, mọi nhà máy cập nhật được. Đây là lý do chọn kiến trúc thư viện dùng chung (xem [05_System_Architecture.md](05_System_Architecture.md)).

**Phương án 2 — Tự tạo thư viện riêng** (khi muốn hoàn toàn tự chủ):

```bash
# vẫn đang ở src/QDD-Core-Library
npx clasp create --type standalone --title "QDD-Core-Library"
npx clasp push
npx clasp version "Ban dau"
```

Ghi lại **Script ID** hiện trong file `.clasp.json` vừa sinh ra, và **số version** (lần đầu là `1`).

### B5. Tạo Google Sheets + Apps Script cho nhà máy

```bash
cd ../NhaMay-Mau-Template
npm install
npx clasp create --type sheets --title "QDD - <Tên nhà máy của bạn>"
```

Lệnh này tạo **một Google Sheets mới** kèm Apps Script gắn liền, và in ra link Sheet.

### B6. Trỏ đúng thư viện

Mở file `appsscript.json` trong `src/NhaMay-Mau-Template`, sửa 2 giá trị:

```json
{
  "dependencies": {
    "libraries": [
      {
        "userSymbol": "QDDCoreLibrary",
        "libraryId": "<Script ID của thư viện>",
        "version": "<số version>"
      }
    ]
  }
}
```

- Dùng chung thư viện có sẵn → giữ nguyên `libraryId`, đặt `version` là số mới nhất (hỏi người quản trị, hoặc xem [CHANGELOG.md](../CHANGELOG.md)).
- Tự tạo thư viện riêng → điền Script ID và version của bạn ở bước B4.

### B7. Đẩy code lên

```bash
npx clasp push
```

### B8. Thiết lập trong Sheet

1. Mở Sheet vừa tạo (link in ra ở bước B5), tải lại trang (F5).
2. Menu **QDD Smart System → Thiết lập sheet**.
3. Vào `CAI_DAT` điền đúng thông số nhà máy (như mục 5 của Cách A).
4. Đọc sheet `HUONG_DAN` để biết quy trình sử dụng.

---

## Chia sẻ thư viện (bắt buộc nếu người khác cùng dùng)

Sheet gọi sang thư viện, nên **mọi người dùng Sheet đều phải có quyền xem thư viện**. Nếu thiếu bước này, người khác mở Sheet sẽ **không thấy menu** vì script chết ngay từ đầu.

1. Mở `https://drive.google.com/file/d/<Script ID thư viện>/view`
2. **Chia sẻ** → Quyền truy cập chung → **Bất kỳ ai có đường liên kết** → vai trò **Người xem**

Thư viện chỉ chứa mã thuật toán, không chứa dữ liệu vận hành, nên mở quyền xem là an toàn.

## Chia sẻ Sheet cho người dùng

- Vai trò phải là **Người chỉnh sửa (Editor)** — người chỉ có quyền Xem sẽ không chạy được Apps Script, không thấy menu.
- Khi dùng thật, nên **mời từng người theo email** thay vì để "bất kỳ ai có link đều chỉnh sửa".
- Lần đầu mỗi người bấm một chức năng, Google hỏi cấp quyền: chọn tài khoản → **Nâng cao (Advanced)** → **Chuyển đến … (unsafe)** → **Cho phép (Allow)**. Cảnh báo này là bình thường với mọi script chưa qua kiểm duyệt của Google.

---

## Cập nhật khi có phiên bản mới

**Cập nhật phần Sheet (giao diện, nhập liệu, báo cáo):**
```bash
cd src/NhaMay-Mau-Template
git pull
npx clasp push
```

**Cập nhật phần thuật toán (thư viện):**

- Nếu **dùng chung thư viện**: mở Apps Script Editor của Sheet → **Thư viện (Libraries)** → chọn version mới → **Lưu**. Việc này **không tự động** — đúng theo kiến trúc đã chọn, để mỗi nhà máy tự quyết thời điểm cập nhật.
- Nếu **tự quản lý thư viện**:
  ```bash
  cd src/QDD-Core-Library
  git pull
  npx clasp push
  npx clasp version "Mo ta thay doi"
  ```
  rồi sửa số `version` trong `appsscript.json` của Sheet và `npx clasp push` lại.

> ⚠️ Sửa code thư viện mà **quên tạo version mới** thì Sheet vẫn chạy code cũ — đây là lỗi rất dễ mắc.

---

## Kiểm tra trước khi dùng thật

Chạy bộ test của thư viện (không cần Google, không cần dữ liệu thật):

```bash
cd src/QDD-Core-Library
node tests/run_tests.js
```

Phải thấy **tất cả test pass**. Bộ test này khoá lại các lỗi thật đã từng gặp (xem [CHANGELOG.md](../CHANGELOG.md)), nên nếu có test đỏ thì đừng đưa vào dùng.

Sau đó, nên **đối chiếu ít nhất 2 ngày có số liệu tính tay** trước khi tin kết quả — cách làm xem [15_Accuracy_Validation_2026-07.md](15_Accuracy_Validation_2026-07.md).

---

## Sự cố thường gặp

| Hiện tượng | Nguyên nhân & cách xử lý |
|---|---|
| `Premature close` khi `clasp login` | Node.js quá mới. Dùng Node 20 (mục B0). |
| `Lỗi 400: invalid_request` khi đăng nhập | Đã dùng `--no-localhost`. Google bỏ hỗ trợ cách này — dùng `npx clasp login` bình thường. |
| `User has not enabled the Apps Script API` | Chưa bật ở https://script.google.com/home/usersettings (mục B2). |
| Người khác mở Sheet **không thấy menu** | Chưa chia sẻ thư viện, hoặc họ chỉ có quyền Xem. |
| Sửa code thư viện nhưng Sheet không đổi | Quên tạo version mới, hoặc Sheet chưa trỏ sang version mới. |
| Kết quả Qdd **phẳng bằng P0 cả ngày** | Không lệnh nào được nhận — kiểm tra sheet `LENH` có dữ liệu đúng ngày/tổ máy không. Sidebar cũng cảnh báo trường hợp này. |
| Thiếu P0 khi tính ngày đầu tiên | Điền tay 1 dòng vào `P0_NGAY`. Các ngày sau hệ thống tự ghi. |
