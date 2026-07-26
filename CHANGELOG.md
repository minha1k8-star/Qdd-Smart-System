# Changelog

Mỗi mục ghi thay đổi và **lý do** của thay đổi đó, không chỉ mô tả đã sửa gì. Mới nhất ở trên.

## v1.0 — Bản chính thức đầu tiên (26/07/2026)

Bản Google Sheets + Apps Script được đưa vào sử dụng thật tại nhà máy Duyên Hải 1. Các mục dưới đây là nhật ký thay đổi trong quá trình xây dựng, mới nhất ở trên.

### Nhật ký thao tác (sheet NHAT_KY)

Nhiều người cùng dùng một Sheet thì khi số liệu có vấn đề không truy được ai đã làm gì. Nay mọi thao tác chạy từ sidebar đều ghi 1 dòng vào sheet `NHAT_KY`: **thời gian · người dùng · thao tác · chi tiết · kết quả**.

- Ghi **cả thao tác thất bại** (`✗ <lỗi>`) — dòng lỗi thường là dòng hữu ích nhất khi truy vết.
- Dòng mới nhất chèn lên đầu; giữ tối đa 2000 dòng, cũ hơn tự xoá bớt.
- Việc ghi log **không bao giờ làm hỏng thao tác chính**: mọi lỗi khi ghi đều nuốt lặng. Mất một dòng nhật ký nhẹ hơn nhiều so với làm hỏng một lần tính đã chạy xong. Lỗi của thao tác chính vẫn được ném tiếp ra sidebar như cũ — `runLogged_` chỉ quan sát, không nuốt lỗi.

**Người dùng phải tự khai tên** một lần (lưu trong `UserProperties`, riêng theo từng tài khoản Google). Lý do: `Session.getActiveUser().getEmail()` chỉ trả về email khi người dùng **cùng miền Google Workspace** với chủ Sheet — nhà máy dùng Gmail cá nhân nên hàm đó luôn trả chuỗi rỗng.

> Ghi rõ giới hạn: tên là **tự khai**, không phải danh tính đã xác thực. Dùng để phối hợp công việc, không dùng làm bằng chứng quy trách nhiệm — việc đó dùng `Tệp → Lịch sử phiên bản` của Google Sheets.

### Tên file xuất đọc là biết ngay kỳ báo cáo

Tên file xuất trước đây là `BaoCao_QDD_<timestamp>.xlsx` — nhìn không biết của ngày/tháng nào, tải về vài file là lẫn. Nay đặt theo đúng kỳ báo cáo và tổ máy:

| Trường hợp | Tên file |
|---|---|
| Báo cáo tháng 7/2026, cả 2 tổ | `BaoCao_Qdd_Qdu_Thang-07-2026_S1-S2.xlsx` |
| Đúng 1 ngày | `BaoCao_Qdd_Qdu_23-07-2026_S1.xlsx` |
| Khoảng ngày | `BaoCao_Qdd_Qdu_23-07-2026_den_25-07-2026_S1-S2.xlsx` |

- Nhãn tháng lấy theo **tháng người dùng chọn**, không suy từ danh sách ngày có dữ liệu — tháng thiếu vài ngày vẫn phải mang tên tháng đó.
- Xuất lại cùng một kỳ thì thêm hậu tố ` (2)`, ` (3)`… thay vì để nhiều file trùng tên hệt nhau trên Drive.

### Kết quả hiển thị 2 số thập phân

Mọi con số kết quả giờ hiển thị **đúng 2 số thập phân** thay vì chuỗi dài kiểu `590.1833333333333`: sheet `KET_QUA` (Qdd → Qdư), `P0_NGAY` (P0 và Ramp tiếp đến), `BAO_CAO_THANG`, **file Excel/PDF xuất ra** (cả 48 dòng lẫn hàng Tổng ngày), và các thông báo/cảnh báo trên sidebar.

**Chỉ đổi ĐỊNH DẠNG HIỂN THỊ, không làm tròn giá trị thật trong ô.** Lý do: P0 của ngày kế tiếp và hàng Tổng ngày đều đọc lại từ các ô này — nếu làm tròn hẳn số gốc thì mỗi ngày lệch một ít ở P0 và **cộng dồn dần qua chuỗi ngày liên tiếp**, Hàng Tổng ngày vì vậy là tổng của số đầy đủ, có thể chênh vài phần trăm đơn vị so với việc cộng tay các số đã làm tròn hiển thị.

> Sheet `KET_QUA` được áp lại định dạng cho **toàn bộ** các dòng mỗi lần tính, nên chỉ cần tính lại một lần bất kỳ là dữ liệu cũ cũng hiển thị theo định dạng mới.

### Gộp sidebar còn 6 mục: bỏ "Lưu CSV 1 file" và "Tính 1 ngày"

Sidebar có 8 mục, trong đó 2 cặp làm cùng một việc: *Lưu CSV công tơ* (1 file, chọn tay tổ máy/loại dữ liệu) trùng với *Tải CSV hàng loạt*, và *Tính 1 ngày* trùng với *Tính hàng loạt*. Bản hàng loạt làm được cả trường hợp 1 file/1 ngày, nên bỏ hẳn hai mục kia:

- **Mục 1 "Tải CSV công tơ"** — chọn 1 hoặc nhiều file đều được, luôn tự nhận diện tổ máy/loại dữ liệu theo TÊN file và tự đọc ngày từ nội dung. Bỏ hẳn việc chọn tay: vừa mất công vừa từng gây lưu nhầm Qmp vào ô Qdc (rủi ro mà mục cũ phải có riêng một lớp kiểm tra để chặn).
- **Mục 3 "Tính"** — nhận Từ ngày/Đến ngày; tính 1 ngày là trường hợp Từ ngày = Đến ngày, và chọn Từ ngày xong hệ thống **tự điền Đến ngày** giống hệt nên vẫn chỉ 1 thao tác. Giờ chỉ còn **một luồng tính duy nhất** cần kiểm chứng thay vì hai.
- Đánh số lại còn 6 mục: Tải CSV · Nhập danh sách lệnh · Tính · Báo cáo tháng · Xuất báo cáo · Dọn dữ liệu cũ. Cập nhật theo: sheet `HUONG_DAN`, `README.md` của template.
- Xoá `sidebar_saveCsv`, `sidebar_calcOneDay` và phần JS client chỉ phục vụ hai mục đã bỏ (`onCsvFileSelected`, `parseCsvDateGuess`...).

**Đánh đổi đã cân nhắc**: mục "Lưu CSV 1 file" từng là đường thoát khi tên file bị đổi khác chuẩn (nhận diện theo tên thất bại). Nay trường hợp đó phải **đổi lại tên file cho đúng chuẩn** `<ngày><tháng><mã công tơ>.CSV` rồi tải lại — thông báo lỗi đã ghi rõ điều này.

### Sửa lệch +14 giờ khi nhập file Excel (nghiêm trọng, sai âm thầm)

**Triệu chứng**: nhập danh sách lệnh từ file Excel xong, mọi lệnh trong `LENH` lệch **đúng +14 giờ** so với file gốc → lệnh buổi tối nhảy sang ngày hôm sau, ngày 17/07 chỉ còn 2 lệnh thay vì 4, kết quả tính sai dây chuyền. Không có lỗi nào được báo.

**Nguyên nhân**: ngày-giờ trong `.xlsx` là giá trị "trần", không kèm múi giờ. Bản Google Sheets tạm do Drive tạo ra khi chuyển đổi lấy **múi giờ mặc định của tài khoản Google** (`America/Los_Angeles`, −07 giờ hè), trong khi Sheet đích dùng `Asia/Ho_Chi_Minh` (+07). Apps Script đọc ô `18:36` thành 18:36 giờ Los Angeles rồi ghi sang Sheet đích thành `08:36` hôm sau — chênh lệch đúng 14 giờ.

**Khắc phục**: `alignTimeZoneWithTargetSheet_` đặt múi giờ file tạm bằng múi giờ Sheet đích trước khi đọc (`setSpreadsheetTimeZone` chỉ đổi cách diễn giải, không sửa số liệu gốc).

**Chặn tái diễn**: sau mỗi lần nhập, sidebar hiện rõ **khoảng BĐTH thực sự đã ghi vào `LENH`** kèm nhắc đối chiếu với file gốc — lệch ngày/giờ lộ ra ngay thay vì phải phát hiện qua kết quả tính sai. Ghi thêm vào mục "cạm bẫy kỹ thuật" trong `AGENTS.md`.

> Cách dán tay vào `LENH` không dính lỗi này (dữ liệu vào thẳng Sheet đúng múi giờ) — chỉ đường nhập qua file Excel mới bị.

### Nhập danh sách lệnh thẳng từ file Excel, bỏ LENH_STAGING

Trước đây muốn nhập lệnh phải **mở file gốc, copy toàn bộ, dán vào sheet** (`LENH` hoặc `LENH_STAGING`) rồi mới bấm nhập. Nay sidebar mục 2 **chọn thẳng file Excel** (`.xlsx`/`.xlsm`/`.xls`) là xong.

Cách làm: file được tải lên Drive kèm yêu cầu chuyển sang Google Sheets (Drive API qua `UrlFetchApp`, không cần bật Advanced Service nào — nhà máy khác chỉ cần `clasp push` là chạy), đọc xong **xoá file tạm ngay**, kể cả khi nhập lỗi. Google chuyển đổi giữ nguyên ô ngày-giờ thành `Date` thật nên không dính lỗi BĐTH thành chữ như khi dán tay.

- **Tự dò bảng lệnh**: quét tất cả các sheet trong file, 25 dòng đầu mỗi sheet, tìm dòng tiêu đề có đủ 9 cột bắt buộc. File gốc có dòng logo/tiêu đề phía trên bảng, hay bảng nằm ở sheet thứ hai, đều nhận được. Không tìm thấy thì báo rõ từng sheet thiếu cột nào.
- **Bỏ hẳn sheet `LENH_STAGING`** — không còn bước trung gian nên không còn chỗ để dữ liệu cũ nằm lại gây nhầm. Sheet này bị xoá tự động khi chạy "Thiết lập sheet" (như đã làm với `CSV_STAGING` trước đây). `importCommandsFromStaging_()` đổi thành `importCommandTable_(table, firstDataRowNumber)` nhận thẳng bảng dữ liệu; toàn bộ luật dò cột theo TÊN tiêu đề, gộp theo ID Lệnh, sắp xếp theo BĐTH giữ nguyên.
- **BĐTH dạng chữ cũng đọc được**: thêm `coerceBdth_` nhận cả ô ngày-giờ thật lẫn chuỗi `dd/MM/yyyy HH:mm[:ss]` (và biến thể dấu `-`). Trước đây ô BĐTH lưu dạng text làm cả dòng bị loại, báo lỗi khó hiểu. Hàm này kiểm tra theo đặc điểm (`typeof v.getTime === 'function'`) chứ không dùng `instanceof Date` — theo đúng cạm bẫy đã ghi trong AGENTS.md.
- Báo cáo sau khi nhập ghi rõ đọc từ **sheet nào, tiêu đề ở dòng nào, bao nhiêu dòng dữ liệu**, và dòng lỗi được đánh số **theo đúng số dòng trong file gốc** để dễ tra.

Thư viện `QDD-Core-Library` **không đổi** (đây là thay đổi ở lớp nhập liệu của Sheet mẫu), vẫn version 4.

### Không tự tick "dọn lệnh + CSV" sau khi tính

Hai ô **"Dọn lệnh + CSV sau khi tính xong"** trước đây **mặc định được tick sẵn**, nên dữ liệu nguồn bị xoá ngay sau mỗi lần tính nếu người dùng không để ý bỏ tick. Nay **mặc định KHÔNG tick** — muốn dọn thì tự tick.

Lý do: xoá lệnh + CSV là thao tác **không hoàn tác được**, trong khi nhu cầu tính lại/đối chiếu một ngày là thường xuyên. Mặc định an toàn phải là giữ dữ liệu.

Cập nhật kèm: sheet `HUONG_DAN` (mục G1) và `src/NhaMay-Mau-Template/README.md`.

### Sheet HUONG_DAN ngay trong file

Thêm sheet **`HUONG_DAN`** (luôn nằm ngoài cùng bên trái) chứa hướng dẫn sử dụng đầy đủ ngay trong Google Sheets — để người dùng mới hoặc người tiếp nhận sau này đọc là làm được mà không cần mở tài liệu bên ngoài. Nội dung gồm: chuẩn bị lần đầu, quy trình hàng ngày, làm nhiều ngày cùng lúc, xuất báo cáo, **ý nghĩa từng cảnh báo tự động**, vai trò từng sheet, và các lưu ý quan trọng (tắt "dọn dữ liệu" khi đang đối chiếu, quyền Chỉnh sửa mới thấy menu...).

Sheet này do hệ thống ghi lại mỗi lần chạy "Thiết lập sheet".

### File xuất bám sát bản "Kiểm tra Qdu" gốc

- **Tổ máy chưa có dữ liệu vẫn giữ nguyên bảng** trong file xuất (để trống, bỏ hàng tổng) thay vì biến mất — đúng như bản gốc luôn có cả S1 và S2 cạnh nhau, bên nào chưa nhập thì bỏ trống. Ngày mà không tổ máy nào có dữ liệu thì vẫn bỏ qua, không tạo tab rỗng.
- Cột **Qdư âm/dương** trong file xuất: chỉ ghi `âm`/`dương`, khi Qdư = 0 thì để dấu **`-`** đúng như bản gốc (không ghi "trong ±3%" — chữ đó giữ lại cho sheet `KET_QUA` nội bộ vì dễ đọc hơn khi làm việc).

### R07: ramp vắt qua nửa đêm đã chạy tiếp sang ngày sau

Hoàn thiện nốt quy tắc **R07 (carry-over)** — trước đó mới chỉ đúng công suất *tại* 24:00, còn phần ramp **chưa hoàn tất** thì ngày hôm sau giữ nguyên công suất đó cho tới khi có lệnh mới, sai so với thực tế (tổ máy vẫn đang tăng/giảm tải nốt).

- `Segments.carryOverOf` phát hiện ramp còn dở dang lúc 24:00 (mục tiêu + thời gian còn lại).
- `calculateDay` nhận `carryTarget`: chèn một **lệnh nối ảo tại 00:00:00** hướng tới mục tiêu đó. Không cần truyền thời gian còn lại — Ramp Engine tự tính đúng phần còn lại từ (P0, mục tiêu, tốc độ). 
- `P0_NGAY` thêm cột **`Ramp tiếp đến (MW)`**; sau mỗi lần tính, hệ thống tự ghi cả P0 lẫn mục tiêu ramp cho ngày kế tiếp.
- Sidebar báo rõ khi một ngày được nối tiếp ramp từ ngày trước.

45/45 test pass (thêm 9 test cho R07). Thư viện lên **version 4**.

> Dữ liệu 07/2026 hiện không có ngày nào ramp vắt qua nửa đêm (gần nhất 18/07 kết thúc 23h52), nên tính năng này chưa kiểm chứng được bằng dữ liệu thật — mới xác nhận qua test.

### P0 phải là công suất tại 24:00, không phải Qdd chu kỳ 48

**Triệu chứng**: ngày 19/07 lệch đều +29,43 MW ở 36 chu kỳ đầu so với bảng tính tay, dù ngày 17 và 18 khớp tuyệt đối.

**Nguyên nhân**: P0 của ngày sau lấy **Qdd chu kỳ 48** của ngày trước — đó là công suất TRUNG BÌNH khoảng 23:30–24:00, không phải công suất TẠI 24:00. Ngày 18/07 lúc đó đang giảm tải 533,1 → 435,7 nên trung bình là 465,131 còn giá trị cuối mới là 435,7. Sai ở điểm khởi đầu nên sai lan ra cả ngày.

**Khắc phục**:
- Thêm `Segments.endPowerOfDay` / `AreaIntegration.endPowerOfDay`; `calculateDay` trả kèm `endPower`.
- Sau mỗi lần tính, Sheet **tự ghi P0 cho ngày kế tiếp** vào `P0_NGAY` (ghi chú "Tự động từ cuối ngày ..."), **không ghi đè** dòng người dùng nhập tay. Bỏ hẳn cách suy P0 từ Qdd chu kỳ 48.

**Lỗi thứ hai phát hiện cùng lúc**: trong `Segments.build`, khi ramp bị **cắt giữa chừng** (lệnh mới đến, hoặc hết ngày), công suất cuối đoạn vẫn ghi là mục tiêu `D` thay vì giá trị nội suy tại điểm cắt → sai độ dốc đoạn → sai diện tích. Đã sửa theo đúng công thức gốc `DOAN_CONG_SUAT!F`.

**Sửa thêm**: mục "Dọn dữ liệu cũ" xoá nhầm cả dòng P0 của ngày kế tiếp — chính là giá trị cần để tính ngày tiếp theo. Giờ giữ lại.

Thư viện lên **version 3**. 36/36 test pass (thêm 5 test khoá riêng 2 lỗi này).

### Sửa lỗi `instanceof Date` qua Library boundary (nghiêm trọng)

**Triệu chứng**: mọi ngày tính ra Qdd phẳng đúng bằng P0 suốt 48 chu kỳ, không lệnh nào được áp dụng — không báo lỗi gì.

**Nguyên nhân**: `CommandFilter.selectEffective` kiểm tra `c.bdth instanceof Date`. Khi Sheet gọi sang `QDD-Core-Library`, mỗi scope có **constructor `Date` riêng**, nên Date tạo ở script gọi KHÔNG thoả `instanceof Date` bên trong thư viện → toàn bộ lệnh bị loại âm thầm. Dữ liệu `LENH` hoàn toàn hợp lệ (đã kiểm chứng kiểu dữ liệu qua gviz: CS ra lệnh/hoàn thành là number, BĐTH là datetime) — lỗi thuần tuý ở code.

**Khắc phục**: thay bằng kiểm tra theo đặc điểm (`typeof v.getTime === 'function'`), không phụ thuộc constructor. Thêm test hồi quy mô phỏng Date đến từ scope khác (31/31 test pass). Thư viện lên **version 2**, Sheet mẫu trỏ sang version 2.

**Xác nhận bằng dữ liệu thật**: ngày 17/07/2026 tổ S1 — sai lệch **0,0000 MW trên cả 48 chu kỳ** so với bảng tính tay `Kiểm tra Qdu`.

> Bài học ghi lại cho người/AI tiếp nhận: **không dùng `instanceof` cho dữ liệu truyền qua ranh giới Apps Script Library** (Date, Array, Error...). Dùng duck typing.

### LENH giống file gốc, cảnh báo lệnh 0-0, dọn dữ liệu

**Sự cố đã xử lý (dữ liệu thật)**: kết quả tính ra sai toàn bộ 20 tổ hợp (Qdd phẳng = P0, không lệnh nào được nhận). Hai nguyên nhân độc lập:
1. Sheet `LENH` còn cấu trúc 9 cột cũ trong khi code đọc theo vị trí cột của cấu trúc mới → lệch cột → mọi lệnh bị loại. **Khắc phục**: `readAllCommands_` và `importCommandsFromStaging_` giờ dò cột theo **TÊN tiêu đề**, không theo vị trí; `LENH` được dựng lại đúng 25 cột giống hệt file gốc (Nhà máy ở cột B) và tự ánh xạ dữ liệu cũ sang đúng cột.
2. Ngày 01/07 lưu nhầm file `01076303.CSV` vào ô Qdc (chọn sai ở mục "Lưu CSV"). **Khắc phục**: chặn ngay khi tên file ứng với mã công tơ khác lựa chọn.

**Tính năng mới**
- **UAT-34 — cảnh báo lệnh "0-0"**: khi ngày tính có lệnh CS ra lệnh = CS hoàn thành = 0 (trip/ngừng sự cố), sidebar cảnh báo rõ rằng lệnh đó không được tính theo quy tắc nghiệp vụ và Qdd có thể đang cao hơn thực tế. Chỉ cảnh báo, không tự sửa số (đúng nguyên tắc đã thống nhất).
- **Dọn dữ liệu nguồn sau khi tính** (tuỳ chọn, mặc định bật): xoá lệnh + CSV của đúng (ngày, tổ máy) vừa tính xong.
- **Dọn dữ liệu cũ** (mục 8): giữ CAI_DAT + kết quả N ngày gần nhất (để P0 vẫn tự suy được), xoá phần còn lại — dùng khi bàn giao hoặc khi dữ liệu tích luỹ nhiều tháng.
- Nhập danh sách lệnh có 2 đường: dán thẳng vào `LENH` (đã giống hệt file gốc) hoặc qua `LENH_STAGING` để gộp thông minh theo ID Lệnh.

**Khác**: `KET_QUA` gộp chu kỳ + khung giờ vào 1 ô (`01 [00:00-00:30]`) giống file gốc; file xuất báo cáo bám layout `Kiểm tra Qdu` (đơn vị MWh, nhãn tổ máy S1DH1/S2DH1 cấu hình được, thứ tự cột Qdd→Qdd_V→Qdc→Qmp→Qdư→âm/dương, hàng Tổng ngày); kết quả sắp xếp ngày mới nhất lên đầu; thao tác xoá dòng chuyển sang đọc-lọc-ghi-lại (nhanh hơn nhiều với vài nghìn dòng); mã công tơ S2 mặc định 6002/6301.

### Sửa lỗi xuất Excel, tự sắp xếp theo ngày, dọn CSV_STAGING

- **Sửa lỗi xuất Excel lần 2**: `DriveApp.getFileById().getAs(MimeType.MICROSOFT_EXCEL)` không hỗ trợ chuyển Google Sheets sang .xlsx — đổi sang gọi link xuất trực tiếp của Google (`docs.google.com/.../export?format=...`) qua `UrlFetchApp` + `ScriptApp.getOAuthToken()`.
- **`CSV_DATA` và `KET_QUA` tự sắp xếp lại theo Ngày** sau mỗi lần lưu/tính (trước đó chỉ nối thêm vào cuối theo thứ tự thao tác, khó kiểm soát khi xem trực tiếp trong Sheet).
- **Dọn sheet `CSV_STAGING`** — không còn dùng từ khi chuyển sang đọc CSV trực tiếp trong sidebar; "Thiết lập sheet" giờ tự xoá sheet này nếu còn sót lại từ bản cũ.
- Đã đẩy lên Sheet thật.

### Sửa lỗi xuất file + Báo cáo tháng xuất file gộp

- **Sửa lỗi xuất Excel/PDF**: `setFrozenColumns(1)` xung đột với ô tiêu đề gộp toàn bộ cột ở hàng 1 (Google Sheets báo lỗi "không thể cố định cột chỉ chứa một phần ô hợp nhất") — bỏ cố định cột, chỉ giữ cố định hàng.
- **Đổi vị trí mục 2/3** trong sidebar: "Tính 1 ngày" lên mục 2 (đi cùng "Lưu CSV" mục 1), "Tải CSV hàng loạt" xuống mục 3 (đi cùng "Tính hàng loạt" mục 4) — nhóm đúng theo luồng dùng đơn lẻ vs hàng loạt.
- **"Báo cáo tháng" giờ xuất file thật**: ngoài cập nhật `BAO_CAO_THANG`, giờ xuất luôn 1 file gộp toàn bộ các ngày đã tính trong tháng (mỗi ngày 1 tab, tổ máy cạnh nhau), dùng chung cơ chế với mục "Xuất báo cáo" — thêm lựa chọn tổ máy + định dạng vào form.
- Đã đẩy lên Sheet thật.

### Xuất báo cáo ngày/khoảng ngày ra file riêng

- Thêm `ExportReport.js` + mục "6. Xuất báo cáo" trong sidebar: chọn khoảng ngày + tổ máy + định dạng (Excel/PDF), xuất từ dữ liệu đã có trong `KET_QUA` (không tính lại) thành 1 file lưu vào Google Drive, sidebar hiện link tải.
- Layout xuất: **mỗi ngày 1 tab**, các tổ máy đã chọn nằm **cạnh nhau trong cùng tab** — khớp layout file báo cáo gốc (S1/S2 cạnh nhau), không phải mỗi tổ máy 1 tab riêng.
- Đã đẩy lên Sheet thật.

### Tải CSV hàng loạt

- **Tải CSV hàng loạt** (mục mới trong sidebar): chọn nhiều file CSV cùng lúc, tự nhận diện tổ máy/loại dữ liệu bằng cách so tên file với mã công tơ đã cấu hình (tận dụng quy ước đặt tên thật `<ngày><tháng><mã công tơ>.CSV`), tự đọc ngày từng file — phục vụ tính hàng loạt nhiều ngày mà không phải upload từng file một. File không tự nhận diện được báo riêng theo tên.

### Sidebar UI thay cho hộp thoại prompt

Phản hồi người dùng: thao tác qua `prompt()` gõ tay ngày tháng + phải tự điền P0 + CSV qua 2 bước cảm giác rời rạc, rời rạc và khó dùng. Đã làm lại:

- **`Sidebar.html` + `Controller.js`** thay toàn bộ menu prompt-based: 1 bảng điều khiển duy nhất, có lịch chọn ngày (`<input type=date>`), upload file CSV trực tiếp (đọc bằng FileReader phía trình duyệt, gửi thẳng lên server - bỏ hẳn bước Import vào CSV_STAGING).
- **Tự động suy ra P0** từ chu kỳ cuối ngày liền trước đã tính (`readOrInferP0_`) — chỉ còn cần nhập tay P0_NGAY cho đúng 1 lần (ngày đầu tiên dùng hệ thống). Đây là xấp xỉ, chưa phải carry-over R07 đầy đủ.
- Bỏ sheet `CSV_STAGING`, menu rút gọn còn 2 mục (Bảng điều khiển, Thiết lập sheet).
- Đã đẩy lên Sheet thật (Duyên Hải 1).

### Giai đoạn 2: triển khai thật + Sheet mẫu

- **Triển khai thật QDD-Core-Library lên Google Apps Script** (đăng nhập, tạo project, push code). Khắc phục lỗi "Premature close" của `clasp login` (do Node.js v24 quá mới) bằng cách cài thêm Node 20 LTS qua Homebrew chỉ để chạy đăng nhập. Publish version 1 làm Library, Script ID `10_vjTSgVjZodA7xTkJ_qJaGom3JDx_tnYE0YgWA_cphh1Q7g_lTKMLUO`.
- Thêm `.claspignore` (loại file test Node ra khỏi bản đẩy Apps Script - nếu không sẽ lỗi toàn bộ thư viện vì `require`/`fs` không tồn tại trong môi trường Apps Script).
- **Tạo Sheet mẫu nhà máy đầu tiên** (`src/NhaMay-Mau-Template/`, Duyên Hải 1) — Google Sheets thật, gắn `QDD-Core-Library`, có menu 6 chức năng: thiết lập sheet, lưu CSV (tận dụng File > Import có sẵn của Sheets), tính 1 ngày, tính hàng loạt nhiều ngày/nhiều tổ máy, tổng hợp báo cáo tháng trực tiếp từ kết quả đã có. Giới hạn: chưa có carry-over R07 (P0 phải nhập tay vào sheet `P0_NGAY`).

### Giai đoạn 2: khởi tạo QDD-Core-Library

- Khởi tạo `src/QDD-Core-Library/` (Apps Script Library, quản lý bằng `clasp`). Port CommandFilter (R01-R03, giữ đúng quy tắc UAT-32), RampEngine (R06, nội suy ngắt ramp), Segments, AreaIntegration (R08, tích phân hình thang), QddCalculator (R09-R14), CsvParser (R10/R12) — trực tiếp từ `tools/reference_engine/qdd_engine.py` đã kiểm chứng bằng dữ liệu thật.
- Thêm bộ test cục bộ chạy bằng Node (`tests/run_tests.js`, 17/17 pass), không cần deploy Apps Script, không cần dữ liệu thật — mã hoá quy tắc UAT-32 thành test tự động để tránh sửa nhầm lại.
- Thêm `BatchCalculator.js` (tính nhiều ngày/nhiều tổ máy cùng lúc) và `MonthlyReport.js` (tổng hợp báo cáo tháng trực tiếp từ dữ liệu gốc). Cập nhật `docs/05_System_Architecture.md` với thiết kế luồng nhập dữ liệu (vừa từng ngày vừa upload hàng loạt). Tổng 29/29 test pass.

### Khởi tạo dự án

- Khởi tạo repo GitHub `Qdd-Smart-System`.
- `CLAUDE.md`, `AGENTS.md`, `docs/AI_CONTEXT.md` — hướng dẫn làm việc cho AI agent tiếp nhận dự án.
- Tài liệu nền: `README.md`, `docs/00_Project_Overview.md`, `docs/03_Business_Rules.md` (14 quy tắc R01-R14), `docs/04_Algorithm_Specification.md` (đặc tả thuật toán), `docs/06_Database_Design.md`, `docs/09_Test_Cases.md`, `docs/14_Knowledge_Transfer.md`, `docs/PROJECT_STATUS.md`, `ROADMAP.md`.
- Chốt kiến trúc: **Google Sheets + Apps Script Library dùng chung**, không làm Web App trung tâm — lý do và các phương án đã cân nhắc ghi ở `docs/05_System_Architecture.md`.
- Kiểm chứng thuật toán bằng dữ liệu vận hành thật (10 ngày, 07/2026) qua bản tái hiện Python (`tools/reference_engine/`), đối chiếu với bảng tính tay độc lập. Kết quả và phần điều tra quy tắc lệnh 0-0: `docs/15_Accuracy_Validation_2026-07.md`.
