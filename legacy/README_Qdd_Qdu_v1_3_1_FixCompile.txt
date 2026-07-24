CÔNG CỤ Qdd/QDƯ v1.3.1 – SỬA LỖI COMPILE ALL-IN-ONE

Nguyên nhân: các Private Const M_... của phần báo cáo tháng bị đặt sau End Sub. VBA yêu cầu mọi khai báo cấp module nằm trước thủ tục đầu tiên.

Cài đặt:
1. Xóa toàn bộ module modQdu cũ, gồm modQdu_v1_3_0_AllInOne.
2. Import duy nhất modQdu_v1_3_1_AllInOne_FixCompile.bas.
3. Debug > Compile VBAProject.
4. Lưu .xlsm, đóng Excel, mở lại và chạy Qdu_TaoNut.

Không xóa LICH_SU_THANG hoặc các sheet LS_...
