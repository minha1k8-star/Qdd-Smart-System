# Knowledge Transfer

Tài liệu này ghi lại kinh nghiệm thực tế trong quá trình phát triển công cụ Qdd/Qdư, tổng hợp từ lịch sử trao đổi phát triển (nhiều tháng làm việc, hàng chục vòng lặp sửa lỗi) và các file ghi chú đi kèm bản v1.3.1. Mục đích: để người/AI tiếp nhận dự án sau này không lặp lại các lỗi đã từng gặp.

## Vì sao chọn Excel/VBA thay vì Web App ngay từ đầu

Hướng đi ban đầu là xây dựng website nhận CSV rồi tự tính Qdd, nhưng bị đổi hướng giữa chừng: người phụ trách nghiệp vụ muốn một **file Excel tổng hợp** các sheet CSV + danh sách lệnh để tự tính ra kết quả, thay vì phụ thuộc vào một hệ thống web riêng. Lý do thực tế (suy ra từ bối cảnh): người dùng nghiệp vụ cần **kiểm tra/audit được từng phép tính bằng mắt thường** ngay trên file họ quen dùng, không muốn phụ thuộc vào một "hộp đen". Đây là lý do kiến trúc hiện tại ưu tiên công thức Excel hiển thị được, VBA chỉ điều phối (xem [00_Project_Overview.md](00_Project_Overview.md)).

## Lỗi CSV trên macOS

Khi mở CSV bằng Excel trên **macOS**, đôi khi toàn bộ một dòng dữ liệu bị dán vào **một ô duy nhất** (cột A) thay vì tự tách thành 50 cột như trên Windows. Công cụ phải tự phát hiện tình huống này và tách chuỗi thành 50 trường trước khi xử lý — nếu không, việc nhập CSV báo lỗi hoặc đọc sai dữ liệu dù bước "kiểm tra đọc thử CSV" trước đó báo thành công (từng có trường hợp kiểm tra đọc thành công nhưng bước nhập CSV chính thức vẫn báo lỗi, cho thấy hai đường code sai lệch nhau — cần đối chiếu kỹ khi sửa `Qdu_KiemTraDocCSV` so với luồng nhập CSV chính thức).

Xem quy tắc liên quan: UAT-09 (CSV Mac một cột), UAT-10 (CSV Windows 50 cột) ở [09_Test_Cases.md](09_Test_Cases.md).

## Lệch số liệu lớn khung giờ 19h30–20h00 (kiểm thử ngày 17)

Khi kiểm thử, ngày 18 tính đúng nhưng **ngày 17 bị lệch lớn** ở khung 19h30–20h00. Nguyên nhân gốc liên quan đến việc dùng sai cột làm mốc bắt đầu ramp — phải dùng **cột G (Thời điểm BĐTH)**, không phải cột H (thời điểm hoàn thành). Đây là lý do R04 trong [03_Business_Rules.md](03_Business_Rules.md) nhấn mạnh rõ ràng "không dùng H làm điểm bắt đầu". Bài học: khi thấy lệch số liệu tập trung ở một khung giờ cụ thể (không phải lệch toàn bộ ngày), nhiều khả năng là lỗi chọn sai mốc thời gian của một lệnh cụ thể, không phải lỗi công thức tổng thể.

## Lỗi báo cáo tháng hiểu sai định dạng tháng

Nhập xong ngày 17 thì lưu tháng bình thường, nhưng nhập xong ngày 18 thì hệ thống báo nhầm là đã "sang tháng tiếp theo". Nguyên nhân: cách xác định tháng hiện tại (`M_MONTH_KEY_CELL`, ô `N2`) từng bị hiểu sai định dạng ngày/tháng. Bài học: logic xác định tháng cho `LICH_SU_THANG` phải test riêng với dữ liệu nhiều ngày liên tiếp trong cùng tháng, không chỉ test với 1 ngày đơn lẻ.

## Báo cáo tháng cần xuất cả S1 và S2 trong cùng một báo cáo

Yêu cầu nghiệp vụ: báo cáo tháng phải có cả hai tổ máy S1 và S2 trong cùng một báo cáo theo đúng file mẫu gốc; nếu một tổ không có dữ liệu nhập thì để trống thay vì báo lỗi hoặc bỏ qua toàn bộ báo cáo.

## Âm/dương của Qdư

Trong báo cáo, cột "Qdư âm/dương" chỉ nên ghi **âm, dương, hoặc để trống** (khi Qdư = 0, tức nằm trong dải dung sai) — không hiển thị số 0 hay ký hiệu gây nhầm lẫn.

## Lỗi xuất báo cáo tháng khác nhau giữa macOS và Windows

Đây là nhóm lỗi tốn nhiều công sức nhất trong lịch sử dự án, vì hành vi VBA/Excel khác nhau giữa hai hệ điều hành:

- **macOS**: từng không xuất được báo cáo tháng; nguyên nhân liên quan đến việc file mẫu có **ô gộp (merged cell)** khiến việc ghi dữ liệu vào vùng bị gộp gây lỗi — đã vá ở v1.2.6. Có giai đoạn còn phải tự lưu báo cáo cạnh workbook thay vì dùng hộp thoại lưu file chuẩn do khác biệt hành vi dialog trên Mac (xem UAT-31).
- **Windows**: từng gặp lỗi "Cannot run the macro" khi xuất báo cáo tháng ở bản v1.2.6 — nguyên nhân là gọi macro báo cáo tháng bằng `Application.Run` (macro ở module khác/tên không khớp runtime). Đã loại bỏ hoàn toàn cách gọi này ở v1.3.0, hợp nhất mọi logic vào **một module duy nhất** để tránh phụ thuộc giữa các module khi biên dịch/chạy trên các máy khác nhau.

**Bài học quan trọng nhất**: kiến trúc nhiều module VBA gọi lẫn nhau qua `Application.Run` rất dễ vỡ khi triển khai trên các máy/hệ điều hành khác nhau. Quyết định chuyển sang **một module all-in-one** (v1.3.0 trở đi) là để loại bỏ hẳn nhóm lỗi này, đánh đổi bằng việc module lớn hơn (~4000 dòng) khó đọc hơn.

## Lỗi compile v1.3.0 → v1.3.1

Khi hợp nhất mọi thứ vào một module, 13 hằng số `Private Const M_...` dùng cho phần báo cáo tháng bị đặt **sau** `End Sub` của một thủ tục nào đó. VBA bắt buộc mọi khai báo cấp module (`Const`, `Dim` module-level) phải nằm **trước thủ tục đầu tiên** trong module — nếu không sẽ lỗi compile toàn bộ. Đã sửa bằng cách gom tất cả khai báo `Const` lên đầu module. **Bài học**: khi copy/paste code giữa các module VBA để hợp nhất, luôn kiểm tra lại vị trí các khai báo cấp module trước khi compile, không chỉ kiểm tra logic bên trong từng Sub/Function.

## P đầu ngày (P0) tự động

Từng có thắc mắc: "tại sao các bản trước khi có P0 tự động lại nhập được, còn bản có P0 tự động thì không?" — cho thấy tính năng tự động lấy P0 từ ngày liền trước (R07, `Qdu_TuTinhP0`, nút 9) có thể xung đột với luồng nhập liệu thủ công nếu không kiểm tra kỹ thứ tự thực hiện (nhập lệnh trước hay tính P0 trước). Xem UAT-02, UAT-03 — cần đảm bảo P0 nhập tay không bị ghi đè ngoài ý muốn khi bật tính năng tự động.

## "Chạy chính thức" không cần Qdd gốc

Yêu cầu nghiệp vụ thay đổi qua thời gian: ban đầu quy trình "chạy chính thức" yêu cầu nhập Qdd gốc để đối chiếu, sau đó được điều chỉnh để **không bắt buộc** — Qdd gốc chỉ còn dùng cho mục đích đối chiếu thủ công (nút 4), nếu thiếu dữ liệu thì bỏ qua thay vì chặn tính toán (xem R lists, UAT-25).

## Vì sao báo cáo tháng đọc Snapshot, không đọc dữ liệu hiện tại

(Suy ra từ kiến trúc `LICH_SU_THANG`/`LS_...` và quy tắc "không được xoá" trong AGENTS.md) — nếu báo cáo tháng đọc trực tiếp từ `TINH_TOAN` (dữ liệu ngày đang mở), báo cáo sẽ thay đổi mỗi khi người dùng mở lại workbook và tính một ngày khác. Snapshot đảm bảo báo cáo tháng phản ánh đúng kết quả **đã chốt tại thời điểm tính**, không bị ghi đè bởi thao tác sau đó.

## Giới hạn vận hành: lệnh "0-0" không được tự động phát hiện (07/2026, UAT-32)

Khi tổ máy dừng do sự cố (trip) và được ghi nhận bằng một lệnh có cả CS ra lệnh và CS hoàn thành đều bằng 0 ("lệnh 0-0"), hệ thống **chủ động loại bỏ** lệnh này khỏi tính toán — đây là quy tắc nghiệp vụ có chủ đích (xác nhận với người phụ trách nghiệp vụ), không phải lỗi. Tương tự, khi khởi động lại sau sự cố, lệnh chỉ được coi là "đã hoàn thành" và bắt đầu tính lại khi **CS ra lệnh và CS hoàn thành cùng bằng một giá trị tải thật** (ví dụ 435,7-435,7).

Hệ quả thực tế: nếu dữ liệu vận hành chỉ có đúng lệnh 0-0 (không có gì khác báo hiệu dừng máy), **công cụ sẽ âm thầm giữ nguyên công suất trước đó cho hết ngày** thay vì đưa về 0 — không có cảnh báo nào hiện ra để người vận hành biết mà can thiệp tay. Đây chính là nguyên nhân bảng tính tay ngày 07/07/2026 (tổ S2) cho kết quả khác công cụ: người tính tay biết sự cố đã xảy ra trong thực tế và tự điền tay, không phải công cụ tự suy ra được.

**Định hướng xử lý (kế hoạch, chưa triển khai)**: thay vì lặng lẽ giữ nguyên công suất sai, hệ thống nên **cảnh báo rõ ràng** khi phát hiện một lệnh 0-0 trong ngày đang tính — để người vận hành biết mà xử lý thủ công, thay vì báo cáo âm thầm sai số. Xem hạng mục tương ứng ở [ROADMAP.md](../ROADMAP.md) (Giai đoạn 2/3) và [UAT-34](09_Test_Cases.md).

## Việc còn tồn đọng (theo lịch sử trao đổi, cần xác nhận lại trên bản v1.3.1)

- Chưa rõ toàn bộ các lỗi trên đã được xác nhận hết trên bản v1.3.1 hay chỉ một phần — vì bộ UAT hiện tại (31 case) toàn bộ đang ở trạng thái "Chưa chạy" (xem [09_Test_Cases.md](09_Test_Cases.md)). Cần thực thi UAT-04 (ramp qua 00:00), UAT-08 (ngắt ramp), UAT-24/UAT-30/UAT-31 (xuất báo cáo tháng Mac/Windows) trước tiên vì đây là các nhóm lỗi lịch sử nặng nhất.
