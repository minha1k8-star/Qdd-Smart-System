Attribute VB_Name = "modQdu_v1_3_1_AllInOne"
Option Explicit

' OFFICIAL RELEASE v1.3.1 - fixed module-level declaration order; single-module architecture
' Formula sheets remain visible for audit, verification and acceptance testing.
' CSV import uses only KwhGiao and 48 numeric periods; column A date is ignored.

Private Const SH_CTRL As String = "DIEU_KHIEN_VBA"
Private Const SH_CFG As String = "CAI_DAT"
Private Const SH_CMD As String = "LENH_GOC"
Private Const SH_EFF As String = "LENH_DIEU_DO"
Private Const SH_6001 As String = "CSV_6001"
Private Const SH_6303 As String = "CSV_6303"
Private Const SH_PROCESS As String = "XU_LY_LENH"
Private Const SH_CALC As String = "TINH_TOAN"
Private Const SH_REPORT As String = "BAO_CAO_QDU"
Private Const SH_LOG As String = "NHAT_KY_VBA"
Private Const SH_STATE As String = "TRANG_THAI_CONG_SUAT"

Private Const CMD_FIRST_ROW As Long = 4
Private Const CMD_LAST_ROW As Long = 203
Private Const CSV_LAST_ROW As Long = 50
Private Const MAX_COMMANDS As Long = 199
Private Const MAX_EFFECTIVE_COMMANDS As Long = 60
Private Const PERIOD_COUNT As Long = 48
Private Const EPS As Double = 0.000001

Private Const APP_NAME As String = "Cong cu tinh Qdd - Qdu"
Private Const APP_VERSION As String = "1.3.1"
Private Const APP_RELEASE_DATE As String = "23/07/2026"
Private Const SH_INFO As String = "THONG_TIN_HE_THONG"
Private Const SH_RULES As String = "QUY_TAC_NGHIEP_VU"
Private Const SH_UAT As String = "KIEM_THU_UAT"
Private Const DEFAULT_QDD_TOL As Double = 0.1

' Monthly history/report constants - declarations must precede procedures
Private Const M_SH_CFG As String = "CAI_DAT"
Private Const M_SH_CALC As String = "TINH_TOAN"
Private Const M_SH_REPORT As String = "BAO_CAO_QDU"
Private Const M_SH_CTRL As String = "DIEU_KHIEN_VBA"
Private Const M_SH_HISTORY As String = "LICH_SU_THANG"
Private Const M_SH_LOG As String = "NHAT_KY_VBA"
Private Const M_PREFIX As String = "LS_"
Private Const M_FIRST_ROW As Long = 4
Private Const M_APP_NAME As String = "Cong cu tinh Qdd - Qdu"
Private Const M_MONTH_KEY_CELL As String = "N2"
Private Const M_REPORT_FIRST_ROW As Long = 4
Private Const M_REPORT_LAST_ROW As Long = 51
Private Const M_REPORT_TOTAL_ROW As Long = 52

Public Sub Auto_Open()
    Qdu_KhoiDong
End Sub

Public Sub Qdu_KhoiDong()
    Dim details As String
    Dim monthInitOk As Boolean

    On Error Resume Next
    GetOrCreateStateSheet
    monthInitOk = QduMonth_Initialize()
    UpdateVersionDisplay
    Qdu_TaoNut
    RecalculateWorkbook

    If VerifyWorkbookStructure(False, details) Then
        ThisWorkbook.Worksheets(SH_CTRL).Range("B10").Value = _
            "He thong san sang - phien ban " & APP_VERSION
    Else
        ThisWorkbook.Worksheets(SH_CTRL).Range("B10").Value = _
            "CANH BAO: cau truc workbook chua day du."
        WriteLog "KHOI_DONG", "CANH_BAO", details
    End If
    On Error GoTo 0
End Sub

Public Sub Qdu_TaoNut()
    Dim ws As Worksheet
    Dim stateWs As Worksheet

    On Error GoTo EH
    Set ws = GetOrCreateControlSheet()
    Set stateWs = GetOrCreateStateSheet()
    DeleteQduButtons ws

    AddQduButton ws, "qduBtn1", "1. Nhap danh sach lenh", "Qdu_NhapLenh", "D4:E5"
    AddQduButton ws, "qduBtn2", "2. Nhap CSV 6001", "Qdu_NhapCSV6001", "F4:G5"
    AddQduButton ws, "qduBtn3", "3. Nhap CSV 6303", "Qdu_NhapCSV6303", "D6:E7"
    AddQduButton ws, "qduBtn4", "4. Nhap Qdd goc", "Qdu_NhapQddGoc", "F6:G7"
    AddQduButton ws, "qduBtn5", "5. Cap nhat cai dat", "Qdu_CapNhatCaiDat", "D8:E9"
    AddQduButton ws, "qduBtn6", "6. Tinh va kiem tra", "Qdu_TinhVaKiemTra", "F8:G9"
    AddQduButton ws, "qduBtn7", "7. CHAY CHINH THUC", "Qdu_ChayChinhThuc", "D10:E11"
    AddQduButton ws, "qduBtn8", "8. Xoa du lieu dau vao", "Qdu_XoaDuLieu", "F10:G11"
    AddQduButton ws, "qduBtn9", "9. Tu tinh P dau ngay", "Qdu_TuTinhP0", "D12:E13"
    AddQduButton ws, "qduBtn10", "10. Mo lich su P", "Qdu_MoTrangThai", "F12:G13"
    AddQduButton ws, "qduBtn11", "11. Sao luu du phong", "Qdu_SaoLuuDuPhong", "D14:E15"
    AddQduButton ws, "qduBtn12", "12. Xuat bao cao thang", "Qdu_XuatBaoCaoThang", "F14:G15"
    AddQduButton ws, "qduBtn13", "13. Kiem tra he thong", "Qdu_KiemTraHeThong", "D16:E17"
    AddQduButton ws, "qduBtn14", "14. Kiem tra doc CSV", "Qdu_KiemTraDocCSV", "F16:G17"
    AddQduButton ws, "qduBtn15", "15. Mo lich su thang", "Qdu_MoLichSuThang", "D18:G19"

    UpdateVersionDisplay
    ws.Range("B10").Value = "Da tao 15 nut. San sang van hanh."
    ws.Activate
    WriteLog "TAO_NUT", "OK", "Da tao 15 nut dieu khien."
    Exit Sub
EH:
    MsgBox "Khong tao duoc nut: " & Err.Description, vbExclamation, APP_NAME
End Sub

Public Sub Qdu_NhapLenh()
    If ImportCommandsInteractive(True) Then
        RecalculateWorkbook
    End If
End Sub

Public Sub Qdu_KiemTraDocCSV()
    Dim filePath As Variant
    Dim savedErrNumber As Long
    Dim savedErrDescription As String
    Dim rawData() As Variant
    Dim rowCount As Long
    Dim firstDataRow As Long
    Dim kwhRow As Long
    Dim parsedData() As Variant
    Dim rowsToCopy As Long
    Dim r As Long, c As Long
    Dim sourceIndex As Long
    Dim numberValue As Double
    Dim token As String
    Dim messageText As String

    On Error GoTo EH

    filePath = PickFile("Chon CSV de kiem tra kha nang doc", _
                        "CSV files (*.csv),*.csv")
    If VarType(filePath) = vbBoolean Then Exit Sub

    If Not ReadCsvFile50(CStr(filePath), rawData, rowCount) Then
        Err.Raise vbObjectError + 1130, , "ReadCsvFile50 tra ve False."
    End If

    firstDataRow = 1
    If IsGenericColumnHeader(CStr(rawData(1, 1)), _
                             CStr(rawData(1, 2))) Then
        firstDataRow = 2
    End If

    rowsToCopy = rowCount - firstDataRow + 1
    ReDim parsedData(1 To rowsToCopy, 1 To 50)

    For r = 1 To rowsToCopy
        sourceIndex = firstDataRow + r - 1
        parsedData(r, 1) = CleanCsvToken(CStr(rawData(sourceIndex, 1)))
        parsedData(r, 2) = CleanCsvToken(CStr(rawData(sourceIndex, 2)))

        For c = 3 To 50
            If IsNumeric(rawData(sourceIndex, c)) And _
               Trim$(CStr(rawData(sourceIndex, c))) <> "" Then
                parsedData(r, c) = CDbl(rawData(sourceIndex, c))
            Else
                token = CleanCsvToken(CStr(rawData(sourceIndex, c)))
                If token = "" Then
                    parsedData(r, c) = Empty
                ElseIf TryParseCsvNumber(token, numberValue) Then
                    parsedData(r, c) = numberValue
                Else
                    Err.Raise vbObjectError + 1131, , _
                        "Gia tri khong phai so tai dong " & sourceIndex & _
                        ", cot " & c & ": " & token
                End If
            End If
        Next c
    Next r

    kwhRow = FindKwhGiaoRow(parsedData, rowsToCopy)
    If kwhRow = 0 Then
        Err.Raise vbObjectError + 1132, , "Khong tim thay KwhGiao."
    End If
    If Not Has48NumericPeriods(parsedData, kwhRow) Then
        Err.Raise vbObjectError + 1133, , "KwhGiao khong du 48 so."
    End If
    messageText = "DOC CSV THANH CONG" & vbCrLf & _
                  "File: " & CStr(filePath) & vbCrLf & _
                  "So dong: " & rowCount & vbCrLf & _
                  "So truong moi dong: 50" & vbCrLf & _
                  "KwhGiao tai dong: " & kwhRow & vbCrLf & _
                  "Cot A (chi tham khao, khong kiem tra): " & _
                      CStr(parsedData(kwhRow, 1)) & vbCrLf & _
                  "Gia tri dau: " & parsedData(kwhRow, 3) & vbCrLf & _
                  "Gia tri cuoi: " & parsedData(kwhRow, 50)

    WriteLog "KIEM_TRA_DOC_CSV", "OK", _
             Replace(messageText, vbCrLf, " | ")
    MsgBox messageText, vbInformation, APP_NAME
    Exit Sub

EH:
    savedErrNumber = Err.Number
    savedErrDescription = Err.Description
    If savedErrNumber = 0 Then
        savedErrNumber = vbObjectError + 1139
        savedErrDescription = "Loi kiem tra CSV khong xac dinh."
    End If

    WriteLog "KIEM_TRA_DOC_CSV", "LOI", _
             "Ma=" & savedErrNumber & "; " & savedErrDescription

    MsgBox "KHONG DOC DUOC CSV" & vbCrLf & _
           "Ma loi: " & savedErrNumber & vbCrLf & _
           "Chi tiet: " & savedErrDescription, _
           vbCritical, APP_NAME
End Sub

Public Sub Qdu_NhapCSV6001()
    If ImportCsvInteractive(SH_6001, "6001", "B18", True) Then
        RecalculateWorkbook
    End If
End Sub

Public Sub Qdu_NhapCSV6303()
    If ImportCsvInteractive(SH_6303, "6303", "B19", True) Then
        RecalculateWorkbook
    End If
End Sub

Public Sub Qdu_NhapQddGoc()
    If ImportOriginalQddInteractive(True) Then
        RecalculateWorkbook
    End If
End Sub

Public Sub Qdu_CapNhatCaiDat()
    If UpdateSettingsInteractive(True) Then
        RecalculateWorkbook
    End If
End Sub

Public Sub Qdu_TinhVaKiemTra()
    RunCalculationAndValidation True
End Sub

Public Sub Qdu_ThuNghiemMotNgay()
    Qdu_ChayChinhThuc
End Sub

Public Sub Qdu_ChayChinhThuc()
    Dim answer As VbMsgBoxResult
    Dim details As String
    Dim statusText As String
    Dim monthlySaved As Boolean

    On Error GoTo EH

    If Not VerifyWorkbookStructure(True, details) Then Exit Sub

    answer = MsgBox( _
        "Quy trinh chinh thuc:" & vbCrLf & _
        "1. Kiem tra/cap nhat ngay, to may va P dau ngay" & vbCrLf & _
        "2. Nhap danh sach lenh" & vbCrLf & _
        "3. Nhap CSV 6001 va CSV 6303" & vbCrLf & _
        "4. Calculate Full Rebuild va kiem tra" & vbCrLf & _
        "5. Luu P cuoi ngay cho ngay tiep theo" & vbCrLf & _
        "6. Luu snapshot bao cao thang" & vbCrLf & vbCrLf & _
        "Qdd goc khong bat buoc. Backup chi chay khi bam nut 11." & vbCrLf & _
        "Tiep tuc?", _
        vbYesNo + vbQuestion, APP_NAME & " " & APP_VERSION)
    If answer <> vbYes Then Exit Sub

    If Not SettingsReady() Then
        If Not UpdateSettingsInteractive(True) Then Exit Sub
    End If

    If Not ImportCommandsInteractive(False) Then Exit Sub
    If Not ImportCsvInteractive(SH_6001, "6001", "B18", False) Then Exit Sub
    If Not ImportCsvInteractive(SH_6303, "6303", "B19", False) Then Exit Sub

    RunCalculationAndValidation False
    statusText = CStr(ThisWorkbook.Worksheets(SH_CTRL).Range("B10").Value)

    If Left$(UCase$(statusText), 3) = "LOI" Then
        MsgBox statusText & vbCrLf & _
               "Khong xuat bao cao khi ket qua con loi.", _
               vbCritical, APP_NAME
        Exit Sub
    End If

    monthlySaved = QduMonth_SaveCurrentDay()

    If monthlySaved Then
        answer = MsgBox(statusText & vbCrLf & vbCrLf & _
                        "Da luu ngay vao lich su thang." & vbCrLf & _
                        "Co xuat bao cao thang thanh file .xlsx khong?", _
                        vbYesNo + vbQuestion, APP_NAME)
        If answer = vbYes Then
            monthlySaved = QduMonth_ExportMonthlyReport()
        End If
    Else
        MsgBox statusText & vbCrLf & vbCrLf & _
               "Ket qua tinh van hop le, nhung chua luu duoc lich su thang." & _
               vbCrLf & "Phan tinh toan khong bi huy.", _
               vbExclamation, APP_NAME
    End If
    Exit Sub
EH:
    WriteLog "CHAY_CHINH_THUC", "LOI", _
             "Ma=" & Err.Number & "; " & Err.Description
    MsgBox "Quy trinh chinh thuc that bai: " & Err.Description, _
           vbCritical, APP_NAME
End Sub

Public Sub Qdu_SaoLuuDuPhong()
    If SaveBackupCopy(False) Then
        MsgBox "Da tao ban sao du phong.", vbInformation, APP_NAME
    End If
End Sub

Public Sub Qdu_XuatBaoCaoNgay()
    ExportDailyReport True
End Sub

Public Sub Qdu_XuatBaoCaoThang()
    Dim exportOk As Boolean
    exportOk = QduMonth_ExportMonthlyReport()
End Sub

Public Sub Qdu_MoLichSuThang()
    Dim openOk As Boolean
    openOk = QduMonth_OpenHistory()
End Sub

Public Sub Qdu_KiemTraHeThong()
    Dim details As String
    Dim macroText As String
    Dim monthText As String

    If IsMacroEnabledWorkbook() Then
        macroText = "Dinh dang macro: DAT (.xlsm/.xlsb)"
    Else
        macroText = "CANH BAO: hay Save As thanh .xlsm de luu VBA"
    End If

    If QduMonth_SelfTest() Then
        monthText = "Module bao cao thang: DAT"
    Else
        monthText = "Module bao cao thang: CHUA SAN SANG" & vbCrLf & _
                    "(Khong anh huong den tinh Qdd/Qdu)"
    End If

    If VerifyWorkbookStructure(False, details) Then
        MsgBox "Cau truc tinh toan: DAT" & vbCrLf & _
               macroText & vbCrLf & _
               monthText & vbCrLf & _
               "Phien ban core: " & APP_VERSION, _
               vbInformation, APP_NAME
    Else
        MsgBox "Cau truc tinh toan: CHUA DAT" & vbCrLf & _
               details & vbCrLf & macroText & vbCrLf & monthText, _
               vbCritical, APP_NAME
    End If
End Sub

Public Sub Qdu_XoaDuLieu()
    Dim answer As VbMsgBoxResult
    Dim oldCalc As XlCalculation

    answer = MsgBox("Xoa danh sach lenh, hai CSV va Qdd file goc?" & vbCrLf & _
                    "Cong thuc, tieu de va lich su cong suat se duoc giu nguyen.", _
                    vbYesNo + vbQuestion, "Xoa du lieu thu nghiem")
    If answer <> vbYes Then Exit Sub

    On Error GoTo EH
    oldCalc = Application.Calculation
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual

    ThisWorkbook.Worksheets(SH_CMD).Range("A4:Y203").ClearContents
    ThisWorkbook.Worksheets(SH_6001).Range("A1:AX50").ClearContents
    ThisWorkbook.Worksheets(SH_6303).Range("A1:AX50").ClearContents
    ThisWorkbook.Worksheets(SH_CALC).Range("O2:O49").ClearContents

    With ThisWorkbook.Worksheets(SH_CTRL)
        .Range("B10:B15").ClearContents
        .Range("B10").Value = "Da xoa du lieu thu nghiem."
        .Range("B17:B20").ClearContents
    End With

    Application.Calculation = xlCalculationAutomatic
    RecalculateWorkbook
    WriteLog "XOA_DU_LIEU", "OK", "Da xoa cac vung input."
    MsgBox "Da xoa du lieu cu. Cong thuc, tieu de va lich su cong suat van duoc giu nguyen.", vbInformation, APP_NAME
SafeExit:
    Application.Calculation = oldCalc
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    Exit Sub
EH:
    MsgBox "Khong xoa duoc du lieu: " & Err.Description, vbExclamation, APP_NAME
    WriteLog "XOA_DU_LIEU", "LOI", Err.Description
    Resume SafeExit
End Sub

Public Sub Qdu_MoBaoCao()
    ThisWorkbook.Worksheets(SH_REPORT).Activate
End Sub

Public Sub Qdu_TuTinhP0()
    Dim ws As Worksheet
    Dim testDate As Date
    Dim unitText As String

    On Error GoTo EH
    Set ws = ThisWorkbook.Worksheets(SH_CFG)
    If Not IsDate(ws.Range("B4").Value) Then
        Err.Raise vbObjectError + 1501, , "Chua co ngay tinh hop le."
    End If

    testDate = CDate(ws.Range("B4").Value)
    unitText = NormalizeUnitName(CStr(ws.Range("B5").Value))
    If unitText <> "S1" And unitText <> "S2" Then
        Err.Raise vbObjectError + 1502, , "Chua co to may S1/S2 hop le."
    End If

    If ApplyAutomaticInitialPower(testDate, unitText, False, True) Then
        RecalculateWorkbook
    End If
    Exit Sub
EH:
    MsgBox "Khong tu tinh duoc P dau ngay: " & Err.Description, vbExclamation, APP_NAME
End Sub

Public Sub Qdu_MoTrangThai()
    Dim ws As Worksheet
    Set ws = GetOrCreateStateSheet()
    ws.Activate
End Sub

Private Function ImportCommandsInteractive(ByVal showMessage As Boolean) As Boolean
    Dim filePath As Variant
    Dim srcWb As Workbook
    Dim srcWs As Worksheet
    Dim dstWs As Worksheet
    Dim cfgWs As Worksheet
    Dim ctrlWs As Worksheet
    Dim headerRow As Long
    Dim lastRow As Long
    Dim dataRows As Long
    Dim data As Variant
    Dim targetDate As Date
    Dim targetUnit As String
    Dim existingP0 As Variant
    Dim existingSource As String
    Dim existingCarry As String
    Dim returnSheetName As String
    Dim matchingRows As Long
    Dim autoRestored As Boolean
    Dim statusText As String

    On Error GoTo EH

    returnSheetName = SH_CTRL
    If ActiveWorkbook Is ThisWorkbook Then returnSheetName = ActiveSheet.Name

    Set cfgWs = ThisWorkbook.Worksheets(SH_CFG)
    Set ctrlWs = ThisWorkbook.Worksheets(SH_CTRL)

    ' CAI_DAT la nguon quyet dinh duy nhat.
    ' Nhap file lenh KHONG duoc thay doi ngay, to may hay P dau ngay.
    If Not IsDate(cfgWs.Range("B4").Value) Then
        Err.Raise vbObjectError + 1000, , _
            "Chua cai dat ngay tinh. Hay bam Cap nhat cai dat truoc khi nhap lenh."
    End If

    targetDate = Int(CDbl(CDate(cfgWs.Range("B4").Value)))
    targetUnit = NormalizeUnitName(CStr(cfgWs.Range("B5").Value))
    If targetUnit <> "S1" And targetUnit <> "S2" Then
        Err.Raise vbObjectError + 1004, , _
            "Chua cai dat to may S1/S2 hop le."
    End If

    existingP0 = cfgWs.Range("B6").Value
    existingSource = CStr(cfgWs.Range("B13").Value)
    existingCarry = CStr(cfgWs.Range("B15").Value)

    filePath = PickFile("Chon file danh sach lenh ket thuc", _
                        "Excel files (*.xlsx;*.xlsm;*.xls),*.xlsx;*.xlsm;*.xls")
    If VarType(filePath) = vbBoolean Then Exit Function

    Application.ScreenUpdating = False
    Application.EnableEvents = False

    Set srcWb = Workbooks.Open(CStr(filePath), ReadOnly:=True, _
                               UpdateLinks:=False, AddToMru:=False)
    Set srcWs = Nothing
    On Error Resume Next
    Set srcWs = srcWb.Worksheets("All")
    On Error GoTo EH
    If srcWs Is Nothing Then Set srcWs = srcWb.Worksheets(1)

    headerRow = FindCommandHeaderRow(srcWs)
    If headerRow = 0 Then Err.Raise vbObjectError + 1001, , _
        "Khong tim thay dong tieu de co ID lenh va 25 cot A:Y."

    lastRow = srcWs.Cells(srcWs.Rows.Count, 1).End(xlUp).Row
    If lastRow <= headerRow Then
        dataRows = 0
    Else
        dataRows = lastRow - headerRow
    End If

    If dataRows > MAX_COMMANDS Then Err.Raise vbObjectError + 1003, , _
        "File co " & dataRows & " dong. Cong cu ho tro toi da " & MAX_COMMANDS & _
        " dong de danh 1 dong cho ramp tiep dien qua 00:00."

    Set dstWs = ThisWorkbook.Worksheets(SH_CMD)
    dstWs.Range("A4:Y203").ClearContents

    If dataRows > 0 Then
        data = srcWs.Range(srcWs.Cells(headerRow + 1, 1), _
                           srcWs.Cells(lastRow, 25)).Value2
        dstWs.Range("A4").Resize(dataRows, 25).Value2 = data
    End If

    dstWs.Range("E4:F203").NumberFormat = "0.000"
    dstWs.Range("G4:H203").NumberFormat = "dd/mm/yyyy hh:mm:ss"
    dstWs.Range("R4:R203").NumberFormat = "dd/mm/yyyy hh:mm:ss"

    ' Dem so lenh thuoc DUNG ngay va DUNG to may dang cai dat.
    matchingRows = CountCommandsForDateUnit(targetDate, targetUnit)

    ' Khoi phuc ramp qua 00:00 neu P0 dang co nguon tu dong.
    ' Neu khong co lich su, van giu nguyen B4/B5/B6.
    If Left$(UCase$(Trim$(existingSource)), 7) = "TU DONG" Then
        autoRestored = ApplyAutomaticInitialPower(targetDate, targetUnit, False, False)
        If Not autoRestored Then
            If IsNumeric(existingP0) Then
                cfgWs.Range("B6").Value = CDbl(existingP0)
                cfgWs.Range("B6").NumberFormat = "0.000"
            End If
            cfgWs.Range("B13").Value = existingSource & _
                " - giu P0, chua tai tao duoc ramp qua 00:00"
            cfgWs.Range("B15").Value = "Can kiem tra"
        End If
    Else
        ' P0 nhap tay: tuyet doi khong goi ham tu dong.
        If IsNumeric(existingP0) Then
            cfgWs.Range("B6").Value = CDbl(existingP0)
            cfgWs.Range("B6").NumberFormat = "0.000"
        End If
        cfgWs.Range("B13").Value = existingSource
        cfgWs.Range("B15").Value = existingCarry
    End If

    ' Khóa lại ngày và tổ máy để tránh mọi thay đổi ngoài ý muốn.
    cfgWs.Range("B4").Value = targetDate
    cfgWs.Range("B4").NumberFormat = "dd/mm/yyyy"
    cfgWs.Range("B5").Value = targetUnit

    If IsNumeric(cfgWs.Range("B6").Value) Then
        statusText = "Da nhap " & dataRows & " dong; giu P dau ngay " & _
                     Format$(CDbl(cfgWs.Range("B6").Value), "0.000") & " MW."
    Else
        statusText = "Da nhap lenh nhung P dau ngay dang trong; hay cap nhat cai dat."
    End If

    If matchingRows = 0 Then
        statusText = statusText & " Khong co lenh hop le dung ngay/to may dang chon."
    Else
        statusText = statusText & " Tim thay " & matchingRows & _
                     " dong dung ngay/to may."
    End If

    ctrlWs.Range("B17").Value = CStr(filePath)
    ctrlWs.Range("B8").Value = cfgWs.Range("B13").Value
    ctrlWs.Range("B10").Value = statusText

    srcWb.Close SaveChanges:=False
    Set srcWb = Nothing

    ImportCommandsInteractive = True
    WriteLog "NHAP_LENH", "OK", _
             "File=" & CStr(filePath) & _
             "; tong_dong=" & dataRows & _
             "; ngay_khoa=" & Format$(targetDate, "dd/mm/yyyy") & _
             "; to_khoa=" & targetUnit & _
             "; dong_phu_hop=" & matchingRows & _
             "; P0=" & IIf(IsNumeric(cfgWs.Range("B6").Value), _
                           Format$(CDbl(cfgWs.Range("B6").Value), "0.000"), "TRONG")

    If showMessage Then
        MsgBox statusText, vbInformation, APP_NAME
    End If

SafeExit:
    On Error Resume Next
    If Not srcWb Is Nothing Then srcWb.Close SaveChanges:=False
    ThisWorkbook.Activate
    If returnSheetName <> "" Then ThisWorkbook.Worksheets(returnSheetName).Activate
    On Error GoTo 0
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    Exit Function

EH:
    WriteLog "NHAP_LENH", "LOI", Err.Description
    MsgBox "Khong nhap duoc danh sach lenh: " & Err.Description, _
           vbExclamation, APP_NAME
    Resume SafeExit
End Function

Private Function ImportCsvInteractive(ByVal targetSheet As String, _
                                      ByVal meterCode As String, _
                                      ByVal pathCell As String, _
                                      ByVal showMessage As Boolean) As Boolean
    Dim filePath As Variant
    Dim dstWs As Worksheet
    Dim cfgWs As Worksheet
    Dim ctrlWs As Worksheet
    Dim rawData() As Variant
    Dim textData() As Variant
    Dim numericData() As Variant
    Dim rowData() As Variant
    Dim rowsCount As Long
    Dim firstDataRow As Long
    Dim rowsToCopy As Long
    Dim r As Long, c As Long
    Dim sourceIndex As Long
    Dim kwhRow As Long
    Dim configuredDate As Date
    Dim numberValue As Double
    Dim token As String
    Dim numericCount As Long
    Dim savedErrNumber As Long
    Dim savedErrDescription As String
    Dim stageName As String
    Dim oldEvents As Boolean
    Dim oldScreenUpdating As Boolean
    Dim returnSheetName As String
    Dim cellValue As Variant
    Dim cellType As VbVarType

    On Error GoTo EH

    Set cfgWs = ThisWorkbook.Worksheets(SH_CFG)
    Set ctrlWs = ThisWorkbook.Worksheets(SH_CTRL)

    returnSheetName = SH_CTRL
    If ActiveWorkbook Is ThisWorkbook Then returnSheetName = ActiveSheet.Name

    stageName = "Kiem tra cai dat"
    If Not IsDate(cfgWs.Range("B4").Value) Then
        Err.Raise vbObjectError + 1090, , _
            "Chua cai dat ngay tinh. Hay bam Cap nhat cai dat truoc khi nhap CSV."
    End If
    configuredDate = Int(CDbl(CDate(cfgWs.Range("B4").Value)))

    stageName = "Chon file CSV"
    filePath = PickFile("Chon CSV cong to " & meterCode, _
                        "CSV files (*.csv),*.csv")
    If VarType(filePath) = vbBoolean Then Exit Function

    oldEvents = Application.EnableEvents
    oldScreenUpdating = Application.ScreenUpdating
    Application.EnableEvents = False
    Application.ScreenUpdating = False

    stageName = "Excel mo va doc CSV"
    If Not ReadCsvFile50(CStr(filePath), rawData, rowsCount) Then
        Err.Raise vbObjectError + 1100, , "Bo doc CSV tra ve False."
    End If
    If rowsCount <= 0 Then
        Err.Raise vbObjectError + 1101, , "CSV khong co dong du lieu."
    End If

    stageName = "Chuan hoa du lieu"
    firstDataRow = 1
    If IsGenericColumnHeader(CStr(rawData(1, 1)), _
                             CStr(rawData(1, 2))) Then
        firstDataRow = 2
    End If

    rowsToCopy = rowsCount - firstDataRow + 1
    If rowsToCopy <= 0 Then
        Err.Raise vbObjectError + 1103, , "CSV khong co dong du lieu thuc."
    End If
    If rowsToCopy > CSV_LAST_ROW Then
        Err.Raise vbObjectError + 1104, , _
            "CSV co " & rowsToCopy & " dong; vuot gioi han."
    End If

    ReDim textData(1 To rowsToCopy, 1 To 2)
    ReDim numericData(1 To rowsToCopy, 1 To 48)

    For r = 1 To rowsToCopy
        sourceIndex = firstDataRow + r - 1
        textData(r, 1) = CleanCsvToken(CStr(rawData(sourceIndex, 1)))
        textData(r, 2) = Trim$(CleanCsvToken(CStr(rawData(sourceIndex, 2))))

        For c = 3 To 50
            If IsNumeric(rawData(sourceIndex, c)) And _
               Trim$(CStr(rawData(sourceIndex, c))) <> "" Then
                numericData(r, c - 2) = CDbl(rawData(sourceIndex, c))
            Else
                token = CleanCsvToken(CStr(rawData(sourceIndex, c)))
                If token = "" Then
                    numericData(r, c - 2) = Empty
                ElseIf TryParseCsvNumber(token, numberValue) Then
                    numericData(r, c - 2) = CDbl(numberValue)
                Else
                    Err.Raise vbObjectError + 1102, , _
                        "Gia tri khong phai so tai dong " & sourceIndex & _
                        ", cot " & c & ": " & token
                End If
            End If
        Next c
    Next r

    stageName = "Tim dong KwhGiao"
    kwhRow = FindKwhGiaoRowInTextBlock(textData, rowsToCopy)
    If kwhRow = 0 Then
        Err.Raise vbObjectError + 1105, , _
            "Khong tim thay dong KwhGiao trong cot B."
    End If

    For c = 1 To 48
        If Not IsNumeric(numericData(kwhRow, c)) Then
            Err.Raise vbObjectError + 1106, , _
                "Dong KwhGiao khong du 48 gia tri so truoc khi ghi."
        End If
    Next c

    ' Ngay/cot A cua CSV chi luu de tham khao. Khong doc, khong so sanh,
    ' khong phu thuoc Regional Settings cua Windows hoac macOS.
    stageName = "Bo qua ngay CSV; chi kiem tra KwhGiao va 48 chu ky"

    stageName = "Lam sach sheet " & targetSheet
    Set dstWs = ThisWorkbook.Worksheets(targetSheet)

    If dstWs.ProtectContents Then
        Err.Raise vbObjectError + 1115, , _
            "Sheet " & targetSheet & " dang duoc bao ve."
    End If

    dstWs.Range("A1:AZ50").UnMerge
    dstWs.Range("A1:AZ50").Clear
    dstWs.Range("A1:B50").NumberFormat = "@"
    dstWs.Range("C1:AX50").NumberFormat = "General"

    stageName = "Ghi cot ngay va ten chi tieu"
    dstWs.Range("A1").Resize(rowsToCopy, 2).Value2 = textData

    stageName = "Ghi 48 chu ky so"
    ReDim rowData(1 To 1, 1 To 48)

    For r = 1 To rowsToCopy
        For c = 1 To 48
            If IsNumeric(numericData(r, c)) Then
                rowData(1, c) = CDbl(numericData(r, c))
            Else
                rowData(1, c) = Empty
            End If
        Next c
        dstWs.Range(dstWs.Cells(r, 3), dstWs.Cells(r, 50)).Value2 = rowData
    Next r

    dstWs.Range("C1:AX50").NumberFormat = "0.00"

    stageName = "Xac minh du lieu da ghi"
    numericCount = 0
    For c = 3 To 50
        cellValue = dstWs.Cells(kwhRow, c).Value2
        cellType = VarType(cellValue)
        If cellType = vbByte Or _
           cellType = vbInteger Or _
           cellType = vbLong Or _
           cellType = vbSingle Or _
           cellType = vbDouble Or _
           cellType = vbCurrency Or _
           cellType = vbDecimal Then
            numericCount = numericCount + 1
        End If
    Next c

    If numericCount <> PERIOD_COUNT Then
        Err.Raise vbObjectError + 1107, , _
            "Sau khi ghi chi co " & numericCount & _
            "/48 chu ky la kieu so that."
    End If

    If UCase$(Trim$(CStr(dstWs.Cells(kwhRow, 2).Value2))) <> "KWHGIAO" Then
        Err.Raise vbObjectError + 1108, , _
            "Nhan KwhGiao bi thay doi sau khi ghi."
    End If

    ctrlWs.Range(pathCell).Value = CStr(filePath)
    ctrlWs.Range("B10").Value = _
        "Da nhap CSV " & meterCode & ": 48/48 chu ky so."

    RecalculateWorkbook
    ImportCsvInteractive = True

    WriteLog "NHAP_CSV_" & meterCode, "OK", _
             "File=" & CStr(filePath) & _
             "; dong=" & rowsToCopy & _
             "; KwhGiao=" & kwhRow & _
             "; chu_ky_so=" & numericCount & _
             "; ngay_tinh_CAI_DAT=" & Format$(configuredDate, "dd/mm/yyyy") & _
             "; ngay_CSV=BO_QUA"

    If showMessage Then
        MsgBox "Da nhap CSV " & meterCode & " thanh cong." & vbCrLf & _
               "Ngay tinh: " & Format$(configuredDate, "dd/mm/yyyy") & _
               " (lay tu CAI_DAT)" & vbCrLf & _
               "Ngay trong CSV: bo qua, khong kiem tra." & vbCrLf & _
               "KwhGiao: 48/48 chu ky so.", _
               vbInformation, APP_NAME
    End If

SafeExit:
    On Error Resume Next
    Application.CutCopyMode = False
    Application.EnableEvents = oldEvents
    Application.ScreenUpdating = oldScreenUpdating
    ThisWorkbook.Activate
    If returnSheetName <> "" Then ThisWorkbook.Worksheets(returnSheetName).Activate
    On Error GoTo 0
    Exit Function

EH:
    savedErrNumber = Err.Number
    savedErrDescription = Err.Description
    If savedErrNumber = 0 Then
        savedErrNumber = vbObjectError + 1199
        savedErrDescription = "Loi khong xac dinh tai buoc " & stageName & "."
    End If

    WriteLog "NHAP_CSV_" & meterCode, "LOI", _
             "Buoc=" & stageName & _
             "; Ma=" & savedErrNumber & _
             "; " & savedErrDescription

    MsgBox "Khong nhap duoc CSV " & meterCode & "." & vbCrLf & _
           "Buoc loi: " & stageName & vbCrLf & _
           "Ma loi: " & savedErrNumber & vbCrLf & _
           "Chi tiet: " & savedErrDescription, _
           vbExclamation, APP_NAME

    Resume SafeExit
End Function

Private Function ImportOriginalQddInteractive(ByVal showMessage As Boolean) As Boolean
    Dim filePath As Variant
    Dim srcWb As Workbook
    Dim selectedRange As Range
    Dim output(1 To PERIOD_COUNT, 1 To 1) As Variant
    Dim i As Long

    On Error GoTo EH
    filePath = PickFile("Chon file Excel goc co 48 gia tri Qdd", _
                        "Excel files (*.xlsx;*.xlsm;*.xls),*.xlsx;*.xlsm;*.xls")
    If VarType(filePath) = vbBoolean Then Exit Function

    Application.ScreenUpdating = True
    Application.EnableEvents = False
    Set srcWb = Workbooks.Open(CStr(filePath), ReadOnly:=True, UpdateLinks:=False, AddToMru:=False)
    srcWb.Activate

    MsgBox "Trong file vua mo, hay chuyen den sheet can kiem tra va chon DUNG 48 o Qdd lien tuc theo cot hoac theo hang." & vbCrLf & _
           "Khong chon tieu de.", vbInformation, "Chon Qdd file goc"

    On Error Resume Next
    Set selectedRange = Application.InputBox( _
        Prompt:="Chon 48 o Qdd (00:00-00:30 den 23:30-24:00)", _
        Title:="Qdd file goc", Type:=8)
    On Error GoTo EH

    If selectedRange Is Nothing Then GoTo SafeExit
    If selectedRange.Cells.CountLarge <> PERIOD_COUNT Then Err.Raise vbObjectError + 1201, , _
        "Da chon " & selectedRange.Cells.CountLarge & " o. Can chon dung 48 o."
    If selectedRange.Rows.Count <> 1 And selectedRange.Columns.Count <> 1 Then _
        Err.Raise vbObjectError + 1202, , "Vung Qdd phai la mot hang hoac mot cot lien tuc."

    For i = 1 To PERIOD_COUNT
        If selectedRange.Rows.Count = 1 Then
            output(i, 1) = selectedRange.Cells(1, i).Value2
        Else
            output(i, 1) = selectedRange.Cells(i, 1).Value2
        End If
        If Not IsNumeric(output(i, 1)) Then Err.Raise vbObjectError + 1203, , _
            "Qdd thu " & i & " khong phai la so."
    Next i

    ThisWorkbook.Worksheets(SH_CALC).Range("O2:O49").Value2 = output
    ThisWorkbook.Worksheets(SH_CALC).Range("O1").Value = "Qdd file goc"
    ThisWorkbook.Worksheets(SH_CTRL).Range("B20").Value = CStr(filePath)
    ThisWorkbook.Worksheets(SH_CTRL).Range("B10").Value = "Da nhap 48 Qdd file goc."

    srcWb.Close SaveChanges:=False
    Set srcWb = Nothing
    ThisWorkbook.Activate

    ImportOriginalQddInteractive = True
    WriteLog "NHAP_QDD_GOC", "OK", "File=" & CStr(filePath)
    If showMessage Then MsgBox "Da nhap 48 gia tri Qdd file goc vao TINH_TOAN!O2:O49.", _
                               vbInformation, APP_NAME
SafeExit:
    On Error Resume Next
    If Not srcWb Is Nothing Then srcWb.Close SaveChanges:=False
    ThisWorkbook.Activate
    On Error GoTo 0
    Application.EnableEvents = True
    Exit Function
EH:
    WriteLog "NHAP_QDD_GOC", "LOI", Err.Description
    MsgBox "Khong nhap duoc Qdd file goc: " & Err.Description, vbExclamation, APP_NAME
    Resume SafeExit
End Function

Private Function UpdateSettingsInteractive(ByVal showMessage As Boolean) As Boolean
    Dim ws As Worksheet
    Dim dateText As String
    Dim unitText As String
    Dim powerValue As Variant
    Dim testDate As Date
    Dim autoPower As Double
    Dim carryActive As Boolean
    Dim carryTarget As Double
    Dim carrySource As String
    Dim remainingSeconds As Double
    Dim sourceText As String
    Dim foundAuto As Boolean
    Dim answer As VbMsgBoxResult
    Dim useAuto As Boolean

    On Error GoTo EH
    Set ws = ThisWorkbook.Worksheets(SH_CFG)

    dateText = InputBox("Nhap ngay tinh theo dd/mm/yyyy:", "Ngay tinh", _
                        IIf(IsDate(ws.Range("B4").Value), Format$(ws.Range("B4").Value, "dd/mm/yyyy"), ""))
    If dateText = "" Then Exit Function
    If Not ParseUserDate(dateText, testDate) Then Err.Raise vbObjectError + 1301, , _
        "Ngay tinh khong hop le. Hay nhap theo dd/mm/yyyy."

    unitText = InputBox("Nhap to may (S1 hoac S2):", "To may", CStr(ws.Range("B5").Value))
    If unitText = "" Then Exit Function
    unitText = NormalizeUnitName(unitText)
    If unitText <> "S1" And unitText <> "S2" Then Err.Raise vbObjectError + 1302, , _
        "To may chi nhan S1 hoac S2."

    ws.Range("B4").Value = testDate
    ws.Range("B4").NumberFormat = "dd/mm/yyyy"
    ws.Range("B5").Value = unitText

    foundAuto = TryGetPreviousPowerState(testDate, unitText, autoPower, carryActive, _
                                         carryTarget, carrySource, remainingSeconds, sourceText)
    useAuto = False

    If foundAuto Then
        If showMessage Then
            answer = MsgBox("Tim thay cong suat cuoi ngay " & Format$(testDate - 1, "dd/mm/yyyy") & _
                            " cua " & unitText & " = " & Format$(autoPower, "0.000") & " MW." & vbCrLf & _
                            IIf(carryActive, "Ramp con tiep dien qua 00:00 den " & _
                                Format$(carryTarget, "0.000") & " MW." & vbCrLf, "") & _
                            "Chon YES de dung tu dong; NO de nhap tay; CANCEL de huy.", _
                            vbYesNoCancel + vbQuestion, "Cong suat dau ngay")
            If answer = vbCancel Then Exit Function
            useAuto = (answer = vbYes)
        Else
            useAuto = True
        End If
    End If

    If useAuto Then
        If Not ApplyAutomaticInitialPower(testDate, unitText, False, False) Then
            Err.Raise vbObjectError + 1304, , "Khong ap dung duoc cong suat dau ngay tu dong."
        End If
    Else
        RemoveSyntheticCarryCommand
        powerValue = Application.InputBox( _
            Prompt:="Nhap cong suat tai 00:00:00 (MW):", _
            Title:="Cong suat dau ngay", Default:=IIf(IsNumeric(ws.Range("B6").Value), ws.Range("B6").Value, 0), Type:=1)
        If VarType(powerValue) = vbBoolean Then
            If powerValue = False Then Exit Function
        End If
        If Not IsNumeric(powerValue) Or CDbl(powerValue) < 0 Then Err.Raise vbObjectError + 1303, , _
            "Cong suat dau ngay khong hop le."

        ws.Range("B6").Value = CDbl(powerValue)
        ws.Range("B6").NumberFormat = "0.000"
        ws.Range("B13").Value = "Nhap tay ngay " & Format$(Now, "dd/mm/yyyy hh:mm:ss")
        ws.Range("B15").Value = "Khong"
    End If

    ThisWorkbook.Worksheets(SH_CTRL).Range("B10").Value = "Da cap nhat cai dat."
    ThisWorkbook.Worksheets(SH_CTRL).Range("B8").Value = ws.Range("B13").Value
    UpdateSettingsInteractive = True
    WriteLog "CAP_NHAT_CAI_DAT", "OK", _
             "Ngay=" & Format$(testDate, "dd/mm/yyyy") & "; to=" & unitText & _
             "; P0=" & ws.Range("B6").Value & "; nguon=" & ws.Range("B13").Value
    If showMessage Then MsgBox "Da cap nhat ngay, to may va cong suat dau ngay.", vbInformation, APP_NAME
    Exit Function
EH:
    WriteLog "CAP_NHAT_CAI_DAT", "LOI", Err.Description
    MsgBox "Khong cap nhat duoc cai dat: " & Err.Description, vbExclamation, APP_NAME
End Function

Private Sub RunCalculationAndValidation(ByVal showMessage As Boolean)
    Dim oldCalc As XlCalculation
    Dim oldScreen As Boolean
    Dim oldEvents As Boolean
    Dim errorCount As Long
    Dim errorDetails As String
    Dim validCommands As Long
    Dim processedCommands As Long
    Dim qddNumeric As Long
    Dim qdcNumeric As Long
    Dim qmpNumeric As Long
    Dim qduNumeric As Long
    Dim changedPeriods As Long
    Dim originalCount As Long
    Dim maxDifference As Variant
    Dim totalQdu As Double
    Dim statusText As String
    Dim messageText As String
    Dim stateSaved As Boolean
    Dim endPower As Double
    Dim endDetail As String
    Dim structureDetails As String
    Dim inputDetails As String
    Dim qddTolerance As Double
    Dim logResult As String
    Dim stageName As String
    Dim savedErrNumber As Long
    Dim savedErrDescription As String

    On Error GoTo EH
    stageName = "Khoi tao che do tinh"
    oldCalc = Application.Calculation
    oldScreen = Application.ScreenUpdating
    oldEvents = Application.EnableEvents

    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationAutomatic
    Application.StatusBar = "Dang kiem tra cau truc va du lieu..."

    stageName = "Kiem tra cau truc workbook"
    If Not VerifyWorkbookStructure(False, structureDetails) Then
        Err.Raise vbObjectError + 1600, , _
            "Cau truc workbook khong hop le: " & structureDetails
    End If

    stageName = "Kiem tra du lieu dau vao"
    If Not ValidateInputConsistency(inputDetails) Then
        If Trim$(inputDetails) = "" Then
            inputDetails = "Ham kiem tra dau vao khong tra ve chi tiet."
        End If
        Err.Raise vbObjectError + 1602, , _
            "Du lieu dau vao khong hop le: " & inputDetails
    End If

    stageName = "Doc dung sai doi chieu Qdd"
    qddTolerance = GetQddCompareTolerance()
    Application.StatusBar = "Dang Calculate Full Rebuild..."

    stageName = "Kiem tra P dau ngay"
    If Not IsNumeric(ThisWorkbook.Worksheets(SH_CFG).Range("B6").Value) Then
        Err.Raise vbObjectError + 1601, , _
            "Chua xac dinh cong suat dau ngay. Hay bam 9. Tu tinh P dau ngay hoac 5. Cap nhat cai dat."
    End If

    stageName = "Calculate Full Rebuild"
    RecalculateWorkbook

    stageName = "Cap nhat cot Qdu am/duong"
    UpdateQduSignDisplay ThisWorkbook.Worksheets(SH_REPORT)

    stageName = "Dem lenh va chu ky ket qua"
    validCommands = SafeLong(ThisWorkbook.Worksheets(SH_CFG).Range("B10").Value)
    processedCommands = Application.WorksheetFunction.Count(ThisWorkbook.Worksheets(SH_PROCESS).Range("B2:B61"))
    qddNumeric = Application.WorksheetFunction.Count(ThisWorkbook.Worksheets(SH_CALC).Range("D2:D49"))
    qdcNumeric = Application.WorksheetFunction.Count(ThisWorkbook.Worksheets(SH_CALC).Range("F2:F49"))
    qmpNumeric = Application.WorksheetFunction.Count(ThisWorkbook.Worksheets(SH_CALC).Range("J2:J49"))
    qduNumeric = Application.WorksheetFunction.Count(ThisWorkbook.Worksheets(SH_CALC).Range("K2:K49"))
    changedPeriods = CountChangedQddPeriods(ThisWorkbook.Worksheets(SH_CFG).Range("B6").Value)
    originalCount = Application.WorksheetFunction.Count(ThisWorkbook.Worksheets(SH_CALC).Range("O2:O49"))

    stageName = "Quet loi cong thuc Excel"
    errorCount = CountWorkbookFormulaErrors(errorDetails)

    stageName = "Tinh tong Qdu an toan"
    totalQdu = SumNumericCells(ThisWorkbook.Worksheets(SH_CALC).Range("K2:K49"))

    If originalCount = PERIOD_COUNT Then
        maxDifference = MaxAbsDifference(ThisWorkbook.Worksheets(SH_CALC).Range("D2:D49"), _
                                         ThisWorkbook.Worksheets(SH_CALC).Range("O2:O49"))
    Else
        maxDifference = ""
    End If

    stageName = "Danh gia ket qua"
    statusText = "OK"
    messageText = "Ket qua kiem tra:" & vbCrLf & _
                  "- Lenh hop le: " & validCommands & vbCrLf & _
                  "- Dong XU_LY_LENH: " & processedCommands & vbCrLf & _
                  "- Chu ky Qdd co so: " & qddNumeric & "/48" & vbCrLf & _
                  "- Chu ky Qdc co so: " & qdcNumeric & "/48" & vbCrLf & _
                  "- Chu ky Qmp co so: " & qmpNumeric & "/48" & vbCrLf & _
                  "- Chu ky Qdu co so: " & qduNumeric & "/48" & vbCrLf & _
                  "- Chu ky Qdd khac P dau ngay: "  & changedPeriods & vbCrLf & _
                  "- So o loi Excel: " & errorCount

    If errorCount > 0 Then statusText = "LOI: workbook con o loi Excel"
    If qddNumeric <> PERIOD_COUNT Then statusText = "LOI: Qdd khong du 48 chu ky"
    If qdcNumeric <> PERIOD_COUNT Then statusText = "LOI: Qdc khong du 48 chu ky - kiem tra CSV 6001"
    If qmpNumeric <> PERIOD_COUNT Then statusText = "LOI: Qmp khong du 48 chu ky - kiem tra CSV 6303"
    If qduNumeric <> PERIOD_COUNT Then statusText = "LOI: Qdu khong du 48 chu ky"
    If processedCommands <> validCommands Then statusText = "LOI: so dong XU_LY_LENH khong khop"
    If validCommands > 0 And changedPeriods = 0 And Left$(statusText, 3) <> "LOI" Then _
        statusText = "CANH BAO: co lenh nhung Qdd chua thay doi"

    If originalCount = PERIOD_COUNT And IsNumeric(maxDifference) Then
        If CDbl(maxDifference) > qddTolerance And Left$(statusText, 3) <> "LOI" Then
            statusText = "CANH BAO: sai lech Qdd vuot dung sai doi chieu"
        End If
    End If

    stageName = "Luu trang thai P cuoi ngay"
    If errorCount = 0 And _
       qddNumeric = PERIOD_COUNT And _
       qdcNumeric = PERIOD_COUNT And _
       qmpNumeric = PERIOD_COUNT And _
       qduNumeric = PERIOD_COUNT And _
       processedCommands = validCommands Then
        stateSaved = SaveCurrentPowerState(statusText, endPower, endDetail)
    End If

    stageName = "Cap nhat bang dieu khien"
    With ThisWorkbook.Worksheets(SH_CTRL)
        .Range("B10").Value = statusText
        .Range("B11").Value = errorCount
        .Range("B12").Value = changedPeriods
        .Range("B13").Value = maxDifference
        .Range("B14").Value = totalQdu
        .Range("B15").Value = qdcNumeric & " / " & qmpNumeric & " / " & qduNumeric
        .Range("B21").Value = Now
        .Range("B21").NumberFormat = "dd/mm/yyyy hh:mm:ss"
        .Range("B13:B14").NumberFormat = "0.000000"
    End With

    On Error Resume Next
    ThisWorkbook.Worksheets(SH_INFO).Range("B14").Value = Now
    ThisWorkbook.Worksheets(SH_INFO).Range("B14").NumberFormat = _
        "dd/mm/yyyy hh:mm:ss"
    On Error GoTo EH

    If originalCount = PERIOD_COUNT Then
        messageText = messageText & vbCrLf & _
                      "- Sai lech Qdd lon nhat: " & _
                      Format$(CDbl(maxDifference), "0.000000") & " MW" & _
                      " (dung sai: " & Format$(qddTolerance, "0.000") & " MW)"
    Else
        messageText = messageText & vbCrLf & _
                      "- Qdd goc: khong doi chieu (khong bat buoc)."
    End If
    messageText = messageText & vbCrLf & "- Tong Qdu: " & Format$(totalQdu, "0.000000") & " MWh"
    If stateSaved Then
        messageText = messageText & vbCrLf & "- P cuoi ngay tai 24:00: " & _
                      Format$(endPower, "0.000") & " MW (da luu lich su)"
        If endDetail <> "" Then messageText = messageText & vbCrLf & "- " & endDetail
    Else
        messageText = messageText & vbCrLf & "- Chua luu trang thai cong suat cuoi ngay."
    End If

    If errorCount > 0 And errorDetails <> "" Then
        messageText = messageText & vbCrLf & vbCrLf & "Mot so o loi:" & vbCrLf & errorDetails
    End If

    If Left$(UCase$(statusText), 3) = "LOI" Then
        logResult = "LOI"
    ElseIf Left$(UCase$(statusText), 4) = "CANH" Then
        logResult = "CANH_BAO"
    Else
        logResult = "OK"
    End If

    stageName = "Ghi nhat ky ket qua"
    WriteLog "TINH_KIEM_TRA", logResult, _
             Replace(messageText, vbCrLf, " | ")

    stageName = "Hien thi ket qua"
    If showMessage Then
        If Left$(statusText, 3) = "LOI" Then
            MsgBox statusText & vbCrLf & vbCrLf & messageText, vbExclamation, APP_NAME
        Else
            MsgBox statusText & vbCrLf & vbCrLf & messageText, vbInformation, APP_NAME
        End If
    End If
SafeExit:
    On Error Resume Next
    Application.StatusBar = False
    Application.Calculation = oldCalc
    Application.EnableEvents = oldEvents
    Application.ScreenUpdating = oldScreen
    On Error GoTo 0
    Exit Sub
EH:
    savedErrNumber = Err.Number
    savedErrDescription = Err.Description

    If savedErrNumber = 0 Then
        savedErrNumber = vbObjectError + 1699
        savedErrDescription = "Loi runtime khong co mo ta tu Excel."
    End If
    If Trim$(stageName) = "" Then stageName = "Khong xac dinh"

    WriteLog "TINH_KIEM_TRA", "LOI", _
             "Buoc=" & stageName & _
             "; Ma=" & savedErrNumber & _
             "; " & savedErrDescription

    MsgBox "Tinh hoac kiem tra that bai." & vbCrLf & _
           "Buoc loi: " & stageName & vbCrLf & _
           "Ma loi: " & savedErrNumber & vbCrLf & _
           "Chi tiet: " & savedErrDescription, _
           vbCritical, APP_NAME
    Resume SafeExit
End Sub


Private Function VerifyWorkbookStructure(ByVal showMessage As Boolean, _
                                         ByRef details As String) As Boolean
    Dim requiredSheets As Variant
    Dim item As Variant
    Dim problems As String

    On Error GoTo EH

    requiredSheets = Array(SH_CTRL, SH_CFG, SH_CMD, SH_EFF, _
                           SH_6001, SH_6303, SH_PROCESS, _
                           "DOAN_CONG_SUAT", "DIEN_TICH", SH_CALC, _
                           SH_REPORT, SH_STATE, SH_INFO, SH_RULES, SH_UAT)

    For Each item In requiredSheets
        If Not SheetExists(CStr(item)) Then
            problems = problems & "- Thieu sheet " & CStr(item) & vbCrLf
        End If
    Next item

    If problems = "" Then
        If Left$(UCase$(Trim$(CStr( _
           ThisWorkbook.Worksheets(SH_CMD).Range("A3").Value))), 2) <> "ID" Then
            problems = problems & "- LENH_GOC!A3 khong phai cot ID Lenh" & vbCrLf
        End If

        If Not ThisWorkbook.Worksheets(SH_EFF).Range("T3").HasFormula Then
            problems = problems & "- Thieu cong thuc LENH_DIEU_DO!T3" & vbCrLf
        End If
        If Not ThisWorkbook.Worksheets(SH_EFF).Range("W3").HasFormula Then
            problems = problems & "- Thieu cong thuc LENH_DIEU_DO!W3" & vbCrLf
        End If
        If Not ThisWorkbook.Worksheets(SH_EFF).Range("H62").HasFormula Then
            problems = problems & "- LENH_DIEU_DO chua mo rong den 60 lenh" & vbCrLf
        End If
        If Not ThisWorkbook.Worksheets(SH_PROCESS).Range("B61").HasFormula Then
            problems = problems & "- XU_LY_LENH chua mo rong den 60 lenh" & vbCrLf
        End If
        If Not ThisWorkbook.Worksheets("DOAN_CONG_SUAT").Range("B122").HasFormula Then
            problems = problems & "- DOAN_CONG_SUAT chua du 121 doan" & vbCrLf
        End If
        If Not ThisWorkbook.Worksheets("DIEN_TICH").Range("DV2").HasFormula Then
            problems = problems & "- DIEN_TICH chua mo rong den cot DV" & vbCrLf
        End If
        If Not ThisWorkbook.Worksheets(SH_CALC).Range("D2").HasFormula Then
            problems = problems & "- Thieu cong thuc Qdd tai TINH_TOAN!D2" & vbCrLf
        End If
        If Not ThisWorkbook.Worksheets(SH_CALC).Range("F2").HasFormula Then
            problems = problems & "- Thieu cong thuc Qdc tai TINH_TOAN!F2" & vbCrLf
        End If
        If Not ThisWorkbook.Worksheets(SH_CALC).Range("J2").HasFormula Then
            problems = problems & "- Thieu cong thuc Qmp tai TINH_TOAN!J2" & vbCrLf
        End If
        If Not ThisWorkbook.Worksheets(SH_CALC).Range("K2").HasFormula Then
            problems = problems & "- Thieu cong thuc Qdu tai TINH_TOAN!K2" & vbCrLf
        End If
        If Not IsNumeric(ThisWorkbook.Worksheets(SH_CFG).Range("B7").Value) Or _
           CDbl(ThisWorkbook.Worksheets(SH_CFG).Range("B7").Value) <= 0 Then
            problems = problems & "- Toc do tang/giam tai khong hop le" & vbCrLf
        End If
    End If

    details = Trim$(problems)
    VerifyWorkbookStructure = (problems = "")

    If showMessage Then
        If VerifyWorkbookStructure Then
            MsgBox "Cau truc workbook hop le." & vbCrLf & _
                   "Phien ban: " & APP_VERSION, vbInformation, APP_NAME
        Else
            MsgBox "Cau truc workbook chua dat:" & vbCrLf & problems, _
                   vbCritical, APP_NAME
        End If
    End If
    Exit Function
EH:
    details = Err.Description
    VerifyWorkbookStructure = False
    If showMessage Then
        MsgBox "Khong kiem tra duoc cau truc: " & Err.Description, _
               vbCritical, APP_NAME
    End If
End Function

Private Function ValidateInputConsistency(ByRef details As String) As Boolean
    Dim cfg As Worksheet
    Dim unitName As String
    Dim meterDetail As String
    Dim problems As String
    Dim originalCount As Long
    Dim duplicateCount As Long
    Dim effectiveCommandCount As Long

    On Error GoTo EH
    Set cfg = ThisWorkbook.Worksheets(SH_CFG)

    If Not IsDate(cfg.Range("B4").Value) Then
        problems = problems & "- Ngay tinh trong CAI_DAT chua hop le" & vbCrLf
    End If

    unitName = NormalizeUnitName(CStr(cfg.Range("B5").Value))
    If unitName <> "S1" And unitName <> "S2" Then
        problems = problems & "- To may phai la S1 hoac S2" & vbCrLf
    End If

    If Not IsNumeric(cfg.Range("B6").Value) Then
        problems = problems & "- Chua co P dau ngay" & vbCrLf
    End If

    If Not CheckMeterSheet(SH_6001, meterDetail) Then
        problems = problems & "- CSV 6001: " & meterDetail & vbCrLf
    End If

    If Not CheckMeterSheet(SH_6303, meterDetail) Then
        problems = problems & "- CSV 6303: " & meterDetail & vbCrLf
    End If

    ' Ngay trong hai CSV khong duoc dung de kiem tra du lieu.

    ' Cot T la co hop le 0/1 tren 200 dong cong thuc.
    ' COUNT dem ca gia tri 0, nen co the bao nham 200 lenh.
    effectiveCommandCount = SafeLong(Application.Sum( _
        ThisWorkbook.Worksheets(SH_EFF).Range("T3:T202")))

    If effectiveCommandCount > MAX_EFFECTIVE_COMMANDS Then
        problems = problems & "- So lenh hieu luc: " & _
                   effectiveCommandCount & ", vuot " & _
                   MAX_EFFECTIVE_COMMANDS & " lenh/ngay" & vbCrLf
    End If

    duplicateCount = CountDuplicateCommandIds()
    If duplicateCount > 0 Then
        problems = problems & "- Co " & duplicateCount & _
                   " ID lenh bi trung" & vbCrLf
    End If

    ' Qdd goc chi la du lieu doi chieu tu chon.
    ' Chi so sanh khi co du 48 gia tri; thieu/mot phan se duoc bo qua.
    originalCount = Application.WorksheetFunction.Count( _
        ThisWorkbook.Worksheets(SH_CALC).Range("O2:O49"))

    details = Trim$(problems)
    ValidateInputConsistency = (problems = "")
    Exit Function
EH:
    If Err.Number = 0 Then
        details = "Loi kiem tra dau vao khong co mo ta."
    Else
        details = "Ma " & Err.Number & ": " & Err.Description
    End If
    ValidateInputConsistency = False
End Function

Private Function CheckMeterSheet(ByVal sheetName As String, _
                                 ByRef detailText As String) As Boolean
    Dim ws As Worksheet
    Dim r As Long
    Dim kwhRow As Long
    Dim numericCount As Long

    On Error GoTo EH
    Set ws = ThisWorkbook.Worksheets(sheetName)

    For r = 1 To CSV_LAST_ROW
        If UCase$(Trim$(CStr(ws.Cells(r, 2).Value2))) = "KWHGIAO" Then
            kwhRow = r
            Exit For
        End If
    Next r

    If kwhRow = 0 Then
        detailText = "Khong tim thay KwhGiao"
        Exit Function
    End If

    numericCount = Application.WorksheetFunction.Count( _
        ws.Range(ws.Cells(kwhRow, 3), ws.Cells(kwhRow, 50)))
    If numericCount <> PERIOD_COUNT Then
        detailText = "KwhGiao chi co " & numericCount & "/48 gia tri so"
        Exit Function
    End If

    ' Cot A co the trong, sai ngay, bi dao ngay-thang hoac la so serial.
    detailText = "OK - ngay CSV bo qua"
    CheckMeterSheet = True
    Exit Function
EH:
    detailText = Err.Description
End Function

Private Function CountDuplicateCommandIds() As Long
    Dim ws As Worksheet
    Dim r As Long
    Dim r2 As Long
    Dim id1 As String
    Dim id2 As String

    Set ws = ThisWorkbook.Worksheets(SH_CMD)

    For r = CMD_FIRST_ROW To CMD_LAST_ROW - 1
        id1 = UCase$(Trim$(CStr(ws.Cells(r, 1).Value2)))

        If id1 <> "" And Left$(id1, 11) <> "AUTO_CARRY_" Then
            For r2 = r + 1 To CMD_LAST_ROW
                id2 = UCase$(Trim$(CStr(ws.Cells(r2, 1).Value2)))

                If id1 = id2 Then
                    CountDuplicateCommandIds = CountDuplicateCommandIds + 1
                    Exit For
                End If
            Next r2
        End If
    Next r
End Function

Private Function GetQddCompareTolerance() As Double
    Dim valueIn As Variant

    valueIn = ThisWorkbook.Worksheets(SH_CFG).Range("B16").Value
    If IsNumeric(valueIn) And CDbl(valueIn) >= 0 Then
        GetQddCompareTolerance = CDbl(valueIn)
    Else
        GetQddCompareTolerance = DEFAULT_QDD_TOL
    End If
End Function

Private Function SaveBackupCopy(ByVal silentMode As Boolean) As Boolean
    Dim baseFolder As String
    Dim backupFolder As String
    Dim backupPath As String
    Dim sep As String
    Dim extensionText As String

    On Error GoTo EH

    If ThisWorkbook.Path = "" Then
        If Not silentMode Then
            MsgBox "Hay luu workbook truoc khi sao luu.", _
                   vbExclamation, APP_NAME
        End If
        Exit Function
    End If

    sep = Application.PathSeparator
    baseFolder = ThisWorkbook.Path
    backupFolder = baseFolder & sep & "Backup_Qdu"
    extensionText = Mid$(ThisWorkbook.Name, _
                        InStrRev(ThisWorkbook.Name, "."))

    If Dir(backupFolder, vbDirectory) = "" Then
        On Error Resume Next
        MkDir backupFolder
        If Err.Number <> 0 Then
            Err.Clear
            backupFolder = baseFolder
        End If
        On Error GoTo EH
    End If

    backupPath = backupFolder & sep & _
                 "Backup_Qdd_Qdu_" & _
                 Format$(Now, "yyyymmdd_hhnnss") & _
                 "_v" & Replace(APP_VERSION, ".", "_") & _
                 extensionText

    ThisWorkbook.SaveCopyAs backupPath
    SaveBackupCopy = True

    WriteLog "SAO_LUU", "OK", backupPath
    If Not silentMode Then
        MsgBox "Da sao luu:" & vbCrLf & backupPath, _
               vbInformation, APP_NAME
    End If
    Exit Function
EH:
    WriteLog "SAO_LUU", "LOI", Err.Description
    If Not silentMode Then
        MsgBox "Khong sao luu duoc: " & Err.Description, _
               vbExclamation, APP_NAME
    End If
End Function

Private Function ExportDailyReport(ByVal showMessage As Boolean) As Boolean
    Dim sourceWs As Worksheet
    Dim outWb As Workbook
    Dim outWs As Worksheet
    Dim infoWs As Worksheet
    Dim cfg As Worksheet
    Dim savePath As Variant
    Dim defaultName As String
    Dim statusText As String
    Dim formulaCell As Range
    Dim oldAlerts As Boolean
    Dim savedErrNumber As Long
    Dim savedErrDescription As String

    On Error GoTo EH

    Set cfg = ThisWorkbook.Worksheets(SH_CFG)
    statusText = CStr(ThisWorkbook.Worksheets(SH_CTRL).Range("B10").Value)

    If Left$(UCase$(statusText), 3) = "LOI" Then
        Err.Raise vbObjectError + 1801, , _
            "Khong xuat bao cao khi ket qua con loi."
    End If

    If Not IsDate(cfg.Range("B4").Value) Then
        Err.Raise vbObjectError + 1802, , "Ngay tinh khong hop le."
    End If

    defaultName = "Bao_cao_Qdd_Qdu_" & _
                  Format$(CDate(cfg.Range("B4").Value), "yyyymmdd") & _
                  "_" & NormalizeUnitName(CStr(cfg.Range("B5").Value)) & _
                  ".xlsx"

    savePath = PickSaveFile(defaultName)
    If VarType(savePath) = vbBoolean Then Exit Function

    Set sourceWs = ThisWorkbook.Worksheets(SH_REPORT)
    sourceWs.Copy
    Set outWb = ActiveWorkbook
    Set outWs = outWb.Worksheets(1)
    outWs.Name = "BAO_CAO_QDU"

    For Each formulaCell In outWs.UsedRange.Cells
        If formulaCell.HasFormula Then
            formulaCell.Value = formulaCell.Value
        End If
    Next formulaCell

    Set infoWs = outWb.Worksheets.Add(Before:=outWs)
    infoWs.Name = "THONG_TIN"
    infoWs.Cells(1, 1).Value = "Noi dung"
    infoWs.Cells(1, 2).Value = "Gia tri"
    infoWs.Cells(2, 1).Value = "Cong cu"
    infoWs.Cells(2, 2).Value = APP_NAME
    infoWs.Cells(3, 1).Value = "Phien ban"
    infoWs.Cells(3, 2).Value = APP_VERSION
    infoWs.Cells(4, 1).Value = "Ngay tinh"
    infoWs.Cells(4, 2).Value = CDate(cfg.Range("B4").Value)
    infoWs.Cells(5, 1).Value = "To may"
    infoWs.Cells(5, 2).Value = NormalizeUnitName(CStr(cfg.Range("B5").Value))
    infoWs.Cells(6, 1).Value = "P dau ngay (MW)"
    infoWs.Cells(6, 2).Value = cfg.Range("B6").Value
    infoWs.Cells(7, 1).Value = "P cuoi ngay (MW)"
    infoWs.Cells(7, 2).Value = cfg.Range("B14").Value
    infoWs.Cells(8, 1).Value = "Tong Qdu (MWh)"
    infoWs.Cells(8, 2).Value = _
        ThisWorkbook.Worksheets(SH_CTRL).Range("B14").Value
    infoWs.Cells(9, 1).Value = "Trang thai"
    infoWs.Cells(9, 2).Value = statusText
    infoWs.Cells(10, 1).Value = "Thoi diem xuat"
    infoWs.Cells(10, 2).Value = Now

    infoWs.Range("A1:B1").Font.Bold = True
    infoWs.Range("A1:B1").Interior.Color = RGB(31, 78, 121)
    infoWs.Range("A1:B1").Font.Color = RGB(255, 255, 255)
    infoWs.Columns("A:B").AutoFit
    infoWs.Range("B4").NumberFormat = "dd/mm/yyyy"
    infoWs.Range("B10").NumberFormat = "dd/mm/yyyy hh:mm:ss"

    oldAlerts = Application.DisplayAlerts
    Application.DisplayAlerts = False
    outWb.SaveAs Filename:=CStr(savePath), _
                  FileFormat:=xlOpenXMLWorkbook
    outWb.Close SaveChanges:=False
    Application.DisplayAlerts = oldAlerts

    ExportDailyReport = True
    WriteLog "XUAT_BAO_CAO", "OK", CStr(savePath)

    If showMessage Then
        MsgBox "Da xuat bao cao:" & vbCrLf & CStr(savePath), _
               vbInformation, APP_NAME
    End If
    Exit Function
EH:
    savedErrNumber = Err.Number
    savedErrDescription = Err.Description

    On Error Resume Next
    Application.DisplayAlerts = oldAlerts
    If Not outWb Is Nothing Then outWb.Close SaveChanges:=False
    On Error GoTo 0

    If savedErrNumber = 0 Then
        savedErrNumber = vbObjectError + 1809
        savedErrDescription = "Loi xuat bao cao khong xac dinh."
    End If

    WriteLog "XUAT_BAO_CAO", "LOI", _
             "Ma=" & savedErrNumber & "; " & savedErrDescription
    If showMessage Then
        MsgBox "Khong xuat duoc bao cao:" & vbCrLf & _
               "Ma loi: " & savedErrNumber & vbCrLf & _
               "Chi tiet: " & savedErrDescription, _
               vbExclamation, APP_NAME
    End If
End Function

Private Function PickSaveFile(ByVal defaultName As String) As Variant
#If Mac Then
    PickSaveFile = Application.GetSaveAsFilename( _
        InitialFileName:=defaultName)
#Else
    PickSaveFile = Application.GetSaveAsFilename( _
        InitialFileName:=defaultName, _
        FileFilter:="Excel Workbook (*.xlsx),*.xlsx", _
        Title:="Luu bao cao Qdd/Qdu")
#End If
End Function

Private Function SheetExists(ByVal sheetName As String) As Boolean
    Dim ws As Worksheet

    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(sheetName)
    SheetExists = Not ws Is Nothing
    On Error GoTo 0
End Function




Private Function IsMacroEnabledWorkbook() As Boolean
    Dim lowerName As String

    lowerName = LCase$(ThisWorkbook.Name)
    IsMacroEnabledWorkbook = _
        (Right$(lowerName, 5) = ".xlsm" Or _
         Right$(lowerName, 5) = ".xlsb")
End Function

Private Sub UpdateVersionDisplay()
    On Error Resume Next

    With ThisWorkbook.Worksheets(SH_CTRL)
        .Range("B2").Value = APP_VERSION
        .Range("E2").Value = "CHINH THUC"
        .Range("G2").Value = APP_RELEASE_DATE
    End With

    With ThisWorkbook.Worksheets(SH_CFG)
        .Range("B17").Value = APP_VERSION
        .Range("B18").Value = "CHINH THUC"
    End With

    With ThisWorkbook.Worksheets(SH_INFO)
        .Range("B5").Value = APP_VERSION
        .Range("B6").Value = APP_RELEASE_DATE
    End With

    On Error GoTo 0
End Sub

Private Function GetUserName() As String
    GetUserName = Trim$(Environ$("Username"))
    If GetUserName = "" Then GetUserName = Trim$(Application.UserName)
    If GetUserName = "" Then GetUserName = "Unknown"
End Function

Private Function GetOrCreateStateSheet() As Worksheet
    Dim ws As Worksheet
    Dim headers As Variant
    Dim i As Long

    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(SH_STATE)
    On Error GoTo 0

    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        ws.Name = SH_STATE
    End If

    If Trim$(CStr(ws.Range("A1").Value)) = "" Then
        headers = Array("Ngay", "To may", "P dau ngay (MW)", _
            "P cuoi ngay tai 24:00 (MW)", "Nguon P dau ngay", "So lenh hop le", _
            "Thoi diem cap nhat", "Trang thai", "Ramp qua 00:00", _
            "Muc tieu ramp tiep (MW)", "Nguon lenh tiep", _
            "Thoi gian con lai (giay)", "Chi tiet")
        For i = LBound(headers) To UBound(headers)
            ws.Cells(1, i + 1).Value = headers(i)
        Next i
        ws.Range("A1:M1").Font.Bold = True
        ws.Range("A1:M1").Interior.Color = RGB(31, 78, 121)
        ws.Range("A1:M1").Font.Color = RGB(255, 255, 255)
        ws.Columns("A").NumberFormat = "dd/mm/yyyy"
        ws.Columns("C:D").NumberFormat = "0.000"
        ws.Columns("G").NumberFormat = "dd/mm/yyyy hh:mm:ss"
        ws.Columns("J").NumberFormat = "0.000"
        ws.Columns("L").NumberFormat = "0.000"
        ws.Columns("A:M").ColumnWidth = 20
        ws.Columns("E").ColumnWidth = 38
        ws.Columns("H").ColumnWidth = 28
        ws.Columns("M").ColumnWidth = 55
        ws.Range("A1:M1").AutoFilter
    End If

    Set GetOrCreateStateSheet = ws
End Function

Private Function FindPowerStateRow(ByVal stateDate As Date, ByVal unitName As String) As Long
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim r As Long

    Set ws = GetOrCreateStateSheet()
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    If lastRow < 2 Then Exit Function

    For r = 2 To lastRow
        If IsDate(ws.Cells(r, 1).Value) Then
            If Int(CDbl(CDate(ws.Cells(r, 1).Value))) = Int(CDbl(stateDate)) And _
               NormalizeUnitName(CStr(ws.Cells(r, 2).Value)) = NormalizeUnitName(unitName) Then
                FindPowerStateRow = r
                Exit Function
            End If
        End If
    Next r
End Function

Private Function TryGetPreviousPowerState(ByVal targetDate As Date, _
                                          ByVal unitName As String, _
                                          ByRef startPower As Double, _
                                          ByRef carryActive As Boolean, _
                                          ByRef carryTarget As Double, _
                                          ByRef carrySource As String, _
                                          ByRef remainingSeconds As Double, _
                                          ByRef sourceText As String) As Boolean
    Dim ws As Worksheet
    Dim stateRow As Long
    Dim previousDate As Date

    On Error GoTo Fail
    previousDate = DateAdd("d", -1, Int(CDbl(targetDate)))
    stateRow = FindPowerStateRow(previousDate, unitName)
    If stateRow = 0 Then Exit Function

    Set ws = GetOrCreateStateSheet()
    If Not IsNumeric(ws.Cells(stateRow, 4).Value) Then Exit Function

    startPower = CDbl(ws.Cells(stateRow, 4).Value)
    carryActive = (SafeLong(ws.Cells(stateRow, 9).Value) = 1 Or _
                   UCase$(Trim$(CStr(ws.Cells(stateRow, 9).Value))) = "TRUE" Or _
                   UCase$(Trim$(CStr(ws.Cells(stateRow, 9).Value))) = "CO")
    If IsNumeric(ws.Cells(stateRow, 10).Value) Then carryTarget = CDbl(ws.Cells(stateRow, 10).Value)
    carrySource = UCase$(Trim$(CStr(ws.Cells(stateRow, 11).Value)))
    If IsNumeric(ws.Cells(stateRow, 12).Value) Then remainingSeconds = CDbl(ws.Cells(stateRow, 12).Value)

    sourceText = "Tu dong tu P cuoi ngay " & Format$(previousDate, "dd/mm/yyyy") & _
                 " (" & unitName & ")"
    TryGetPreviousPowerState = True
    Exit Function
Fail:
    TryGetPreviousPowerState = False
End Function

Private Function SettingsReady() As Boolean
    Dim ws As Worksheet
    Dim unitText As String

    On Error GoTo SafeExit
    Set ws = ThisWorkbook.Worksheets(SH_CFG)

    If Not IsDate(ws.Range("B4").Value) Then Exit Function
    unitText = NormalizeUnitName(CStr(ws.Range("B5").Value))
    If unitText <> "S1" And unitText <> "S2" Then Exit Function
    If Not IsNumeric(ws.Range("B6").Value) Then Exit Function
    If CDbl(ws.Range("B6").Value) < 0 Then Exit Function

    SettingsReady = True
SafeExit:
End Function

Private Function CountCommandsForDateUnit(ByVal targetDate As Date, _
                                          ByVal unitName As String) As Long
    Dim ws As Worksheet
    Dim r As Long
    Dim commandDate As Date
    Dim commandUnit As String

    Set ws = ThisWorkbook.Worksheets(SH_CMD)

    For r = CMD_FIRST_ROW To CMD_LAST_ROW
        If Trim$(CStr(ws.Cells(r, 1).Value2)) <> "" Then
            commandUnit = NormalizeUnitName(CStr(ws.Cells(r, 3).Value2))
            If commandUnit = NormalizeUnitName(unitName) Then
                If IsDate(ws.Cells(r, 7).Value) Or IsNumeric(ws.Cells(r, 7).Value) Then
                    commandDate = CDate(ws.Cells(r, 7).Value)
                    If Int(CDbl(commandDate)) = Int(CDbl(targetDate)) Then
                        CountCommandsForDateUnit = CountCommandsForDateUnit + 1
                    End If
                End If
            End If
        End If
    Next r
End Function

Private Function ApplyAutomaticInitialPower(ByVal targetDate As Date, _
                                            ByVal unitName As String, _
                                            ByVal clearWhenMissing As Boolean, _
                                            ByVal showMessage As Boolean) As Boolean
    Dim ws As Worksheet
    Dim startPower As Double
    Dim carryActive As Boolean
    Dim carryTarget As Double
    Dim carrySource As String
    Dim remainingSeconds As Double
    Dim sourceText As String

    On Error GoTo EH
    Set ws = ThisWorkbook.Worksheets(SH_CFG)

    If Not TryGetPreviousPowerState(targetDate, unitName, startPower, carryActive, _
                                    carryTarget, carrySource, remainingSeconds, sourceText) Then
        RemoveSyntheticCarryCommand
        If clearWhenMissing Then ws.Range("B6").ClearContents
        ws.Range("B13").Value = "Chua co lich su ngay " & Format$(targetDate - 1, "dd/mm/yyyy") & _
                                " - can nhap tay"
        ws.Range("B15").Value = ""
        ThisWorkbook.Worksheets(SH_CTRL).Range("B8").Value = ws.Range("B13").Value
        If showMessage Then MsgBox ws.Range("B13").Value, vbExclamation, APP_NAME
        Exit Function
    End If

    ws.Range("B4").Value = Int(CDbl(targetDate))
    ws.Range("B4").NumberFormat = "dd/mm/yyyy"
    ws.Range("B5").Value = NormalizeUnitName(unitName)
    ws.Range("B6").Value = startPower
    ws.Range("B6").NumberFormat = "0.000"
    ws.Range("B13").Value = sourceText

    If carryActive Then
        If carrySource <> "SO" And carrySource <> "MO" Then
            If carryTarget >= startPower Then
                carrySource = "SO"
            Else
                carrySource = "MO"
            End If
        End If
        EnsureCarryOverCommand targetDate, unitName, startPower, carryTarget, _
                               carrySource, remainingSeconds
        ws.Range("B15").Value = "Co - tiep tuc den " & Format$(carryTarget, "0.000") & " MW"
    Else
        RemoveSyntheticCarryCommand
        ws.Range("B15").Value = "Khong"
    End If

    ThisWorkbook.Worksheets(SH_CTRL).Range("B8").Value = sourceText
    ApplyAutomaticInitialPower = True
    WriteLog "TU_TINH_P0", "OK", "Ngay=" & Format$(targetDate, "dd/mm/yyyy") & _
             "; to=" & unitName & "; P0=" & Format$(startPower, "0.000") & _
             "; carry=" & IIf(carryActive, "YES", "NO")
    If showMessage Then
        MsgBox "Da tu lay P dau ngay = " & Format$(startPower, "0.000") & " MW." & vbCrLf & _
               IIf(carryActive, "Da tao lenh tiep tuc ramp qua 00:00 den " & _
                   Format$(carryTarget, "0.000") & " MW.", "Khong co ramp qua 00:00."), _
               vbInformation, APP_NAME
    End If
    Exit Function
EH:
    WriteLog "TU_TINH_P0", "LOI", Err.Description
    If showMessage Then MsgBox "Khong tu tinh duoc P dau ngay: " & Err.Description, _
                               vbExclamation, APP_NAME
End Function

Private Sub RemoveSyntheticCarryCommand()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim dataRows As Long
    Dim data As Variant

    Set ws = ThisWorkbook.Worksheets(SH_CMD)
    If Left$(UCase$(Trim$(CStr(ws.Cells(CMD_FIRST_ROW, 1).Value))), 11) <> "AUTO_CARRY_" Then Exit Sub

    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    If lastRow <= CMD_FIRST_ROW Then
        ws.Range("A4:Y4").ClearContents
        Exit Sub
    End If

    dataRows = lastRow - CMD_FIRST_ROW
    data = ws.Range(ws.Cells(CMD_FIRST_ROW + 1, 1), ws.Cells(lastRow, 25)).Value2
    ws.Range("A4:Y203").ClearContents
    ws.Range("A4").Resize(dataRows, 25).Value2 = data
    ws.Range("E4:F203").NumberFormat = "0.000"
    ws.Range("G4:H203").NumberFormat = "dd/mm/yyyy hh:mm:ss"
    ws.Range("R4:R203").NumberFormat = "dd/mm/yyyy hh:mm:ss"
End Sub

Private Sub EnsureCarryOverCommand(ByVal targetDate As Date, _
                                   ByVal unitName As String, _
                                   ByVal startPower As Double, _
                                   ByVal targetPower As Double, _
                                   ByVal sourceType As String, _
                                   ByVal remainingSeconds As Double)
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim dataRows As Long
    Dim data As Variant
    Dim rowData(1 To 1, 1 To 25) As Variant

    Set ws = ThisWorkbook.Worksheets(SH_CMD)

    If Left$(UCase$(Trim$(CStr(ws.Cells(CMD_FIRST_ROW, 1).Value))), 11) <> "AUTO_CARRY_" Then
        lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
        If lastRow >= CMD_LAST_ROW Then
            Err.Raise vbObjectError + 1510, , _
                "Khong con dong trong LENH_GOC de chen ramp qua 00:00."
        End If

        If lastRow >= CMD_FIRST_ROW Then
            dataRows = lastRow - CMD_FIRST_ROW + 1
            data = ws.Range(ws.Cells(CMD_FIRST_ROW, 1), ws.Cells(lastRow, 25)).Value2
            ws.Range("A4:Y203").ClearContents
            ws.Range("A5").Resize(dataRows, 25).Value2 = data
        End If
    End If

    If sourceType <> "SO" And sourceType <> "MO" Then
        If targetPower >= startPower Then sourceType = "SO" Else sourceType = "MO"
    End If

    rowData(1, 1) = "AUTO_CARRY_" & Format$(targetDate, "yyyymmdd") & "_" & unitName
    rowData(1, 2) = "AUTO"
    rowData(1, 3) = unitName
    rowData(1, 4) = "Tiep tuc ramp qua 00:00"
    rowData(1, 5) = targetPower
    rowData(1, 6) = targetPower
    rowData(1, 7) = Int(CDbl(targetDate))
    rowData(1, 8) = Int(CDbl(targetDate)) + remainingSeconds / 86400#
    rowData(1, 9) = "VBA"
    rowData(1, 10) = "VBA"
    rowData(1, 11) = False
    rowData(1, 13) = "Tiep tuc lenh tu ngay truoc"
    rowData(1, 14) = "Tu dong"
    rowData(1, 15) = "Ramp qua 00:00"
    rowData(1, 16) = 1
    rowData(1, 17) = False
    rowData(1, 23) = 0
    rowData(1, 24) = 0
    rowData(1, 25) = sourceType

    ws.Range("A4:Y4").Value2 = rowData
    ws.Range("E4:F203").NumberFormat = "0.000"
    ws.Range("G4:H203").NumberFormat = "dd/mm/yyyy hh:mm:ss"
    ws.Range("R4:R203").NumberFormat = "dd/mm/yyyy hh:mm:ss"
End Sub

Private Function CalculateEndOfDayState(ByRef endPower As Double, _
                                        ByRef carryActive As Boolean, _
                                        ByRef carryTarget As Double, _
                                        ByRef carrySource As String, _
                                        ByRef remainingSeconds As Double, _
                                        ByRef detailText As String) As Boolean
    Dim cfg As Worksheet
    Dim processWs As Worksheet
    Dim effWs As Worksheet
    Dim commandCount As Long
    Dim lastRow As Long
    Dim startSeconds As Double
    Dim endSeconds As Double
    Dim startPower As Double
    Dim targetPower As Double
    Dim rateValue As Double
    Dim elapsedMinutes As Double
    Dim directionValue As Double
    Dim noteText As String

    On Error GoTo Fail
    Set cfg = ThisWorkbook.Worksheets(SH_CFG)
    Set processWs = ThisWorkbook.Worksheets(SH_PROCESS)
    Set effWs = ThisWorkbook.Worksheets(SH_EFF)

    If Not IsNumeric(cfg.Range("B6").Value) Then Exit Function
    endPower = CDbl(cfg.Range("B6").Value)
    carryActive = False
    carryTarget = 0
    carrySource = ""
    remainingSeconds = 0
    detailText = "Khong co lenh; P cuoi ngay bang P dau ngay."

    commandCount = Application.WorksheetFunction.Count(processWs.Range("B2:B61"))
    If commandCount = 0 Then
        CalculateEndOfDayState = True
        Exit Function
    End If

    lastRow = commandCount + 1
    If Not IsNumeric(processWs.Cells(lastRow, 2).Value) Or _
       Not IsNumeric(processWs.Cells(lastRow, 4).Value) Or _
       Not IsNumeric(processWs.Cells(lastRow, 6).Value) Or _
       Not IsNumeric(processWs.Cells(lastRow, 9).Value) Then Exit Function

    startSeconds = CDbl(processWs.Cells(lastRow, 2).Value)
    targetPower = CDbl(processWs.Cells(lastRow, 4).Value)
    startPower = CDbl(processWs.Cells(lastRow, 6).Value)
    endSeconds = CDbl(processWs.Cells(lastRow, 9).Value)

    noteText = UCase$(Trim$(CStr(effWs.Cells(commandCount + 2, 7).Value)))
    carrySource = Left$(noteText, 2)
    If carrySource <> "SO" And carrySource <> "MO" Then
        If targetPower >= startPower Then carrySource = "SO" Else carrySource = "MO"
    End If

    If endSeconds <= 86400# + EPS Then
        endPower = targetPower
        detailText = "Ramp cuoi da hoan thanh truoc 24:00."
    Else
        If Not IsNumeric(cfg.Range("B7").Value) Or CDbl(cfg.Range("B7").Value) <= 0 Then Exit Function
        rateValue = CDbl(cfg.Range("B7").Value)
        elapsedMinutes = (86400# - startSeconds) / 60#
        If elapsedMinutes < 0 Then elapsedMinutes = 0

        directionValue = Sgn(targetPower - startPower)
        endPower = startPower + directionValue * rateValue * elapsedMinutes
        If directionValue > 0 And endPower > targetPower Then endPower = targetPower
        If directionValue < 0 And endPower < targetPower Then endPower = targetPower

        carryActive = (Abs(endPower - targetPower) > EPS)
        carryTarget = targetPower
        remainingSeconds = endSeconds - 86400#
        If remainingSeconds < 0 Then remainingSeconds = 0
        detailText = "Ramp con tiep dien qua 00:00; P tai 24:00=" & _
                     Format$(endPower, "0.000") & " MW; muc tieu=" & _
                     Format$(targetPower, "0.000") & " MW."
    End If

    CalculateEndOfDayState = True
    Exit Function
Fail:
    CalculateEndOfDayState = False
End Function

Private Function SaveCurrentPowerState(ByVal validationStatus As String, _
                                       ByRef savedEndPower As Double, _
                                       ByRef savedDetail As String) As Boolean
    Dim cfg As Worksheet
    Dim ws As Worksheet
    Dim stateDate As Date
    Dim unitName As String
    Dim stateRow As Long
    Dim nextRow As Long
    Dim carryActive As Boolean
    Dim carryTarget As Double
    Dim carrySource As String
    Dim remainingSeconds As Double

    On Error GoTo EH
    Set cfg = ThisWorkbook.Worksheets(SH_CFG)
    If Not IsDate(cfg.Range("B4").Value) Then Exit Function
    If Not IsNumeric(cfg.Range("B6").Value) Then Exit Function

    stateDate = Int(CDbl(CDate(cfg.Range("B4").Value)))
    unitName = NormalizeUnitName(CStr(cfg.Range("B5").Value))
    If unitName <> "S1" And unitName <> "S2" Then Exit Function

    If Not CalculateEndOfDayState(savedEndPower, carryActive, carryTarget, _
                                   carrySource, remainingSeconds, savedDetail) Then Exit Function

    Set ws = GetOrCreateStateSheet()
    stateRow = FindPowerStateRow(stateDate, unitName)
    If stateRow = 0 Then
        nextRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row + 1
        If nextRow < 2 Then nextRow = 2
        stateRow = nextRow
    End If

    ws.Cells(stateRow, 1).Value = stateDate
    ws.Cells(stateRow, 2).Value = unitName
    ws.Cells(stateRow, 3).Value = CDbl(cfg.Range("B6").Value)
    ws.Cells(stateRow, 4).Value = savedEndPower
    ws.Cells(stateRow, 5).Value = CStr(cfg.Range("B13").Value)
    ws.Cells(stateRow, 6).Value = SafeLong(cfg.Range("B10").Value)
    ws.Cells(stateRow, 7).Value = Now
    ws.Cells(stateRow, 8).Value = validationStatus
    ws.Cells(stateRow, 9).Value = IIf(carryActive, 1, 0)
    If carryActive Then
        ws.Cells(stateRow, 10).Value = carryTarget
        ws.Cells(stateRow, 11).Value = carrySource
        ws.Cells(stateRow, 12).Value = remainingSeconds
    Else
        ws.Range(ws.Cells(stateRow, 10), ws.Cells(stateRow, 12)).ClearContents
    End If
    ws.Cells(stateRow, 13).Value = savedDetail

    ws.Cells(stateRow, 1).NumberFormat = "dd/mm/yyyy"
    ws.Cells(stateRow, 3).NumberFormat = "0.000"
    ws.Cells(stateRow, 4).NumberFormat = "0.000"
    ws.Cells(stateRow, 7).NumberFormat = "dd/mm/yyyy hh:mm:ss"
    ws.Cells(stateRow, 10).NumberFormat = "0.000"
    ws.Cells(stateRow, 12).NumberFormat = "0.000"

    cfg.Range("B14").Value = savedEndPower
    cfg.Range("B14").NumberFormat = "0.000"
    cfg.Range("B15").Value = IIf(carryActive, "Co", "Khong")

    SaveCurrentPowerState = True
    WriteLog "LUU_TRANG_THAI_P", "OK", "Ngay=" & Format$(stateDate, "dd/mm/yyyy") & _
             "; to=" & unitName & "; P0=" & Format$(cfg.Range("B6").Value, "0.000") & _
             "; Pend=" & Format$(savedEndPower, "0.000") & _
             "; carry=" & IIf(carryActive, "YES", "NO")
    Exit Function
EH:
    WriteLog "LUU_TRANG_THAI_P", "LOI", Err.Description
End Function

Private Sub RecalculateWorkbook()
    On Error Resume Next
    Application.CalculateFullRebuild
    If Err.Number <> 0 Then
        Err.Clear
        Application.CalculateFull
    End If
    On Error GoTo 0
    DoEvents
End Sub




Private Function ReadCsvFile50(ByVal filePath As String, _
                               ByRef data() As Variant, _
                               ByRef rowCount As Long) As Boolean
    Dim detailText As String

    If ReadCsvWithExcel50(filePath, data, rowCount, detailText) Then
        ReadCsvFile50 = True
        Exit Function
    End If

    If Trim$(detailText) = "" Then
        detailText = "Excel mo duoc file nhung bo doc khong tra ve chi tiet."
    End If

    Err.Raise vbObjectError + 1140, "ReadCsvFile50", detailText
End Function

Private Function ReadCsvWithExcel50(ByVal filePath As String, _
                                        ByRef data() As Variant, _
                                        ByRef rowCount As Long, _
                                        ByRef detailText As String) As Boolean
    Dim csvWb As Workbook
    Dim csvWs As Worksheet
    Dim lastCell As Range
    Dim lastRow As Long
    Dim lastCol As Long
    Dim r As Long, c As Long
    Dim cellValue As Variant
    Dim lineText As String
    Dim fields As Variant
    Dim fieldCount As Long
    Dim oldAlerts As Boolean
    Dim returnSheetName As String
    Dim savedErrNumber As Long
    Dim savedErrDescription As String
    Dim openedHere As Boolean
    Dim openDetail As String

    On Error GoTo EH

    returnSheetName = SH_CTRL
    If ActiveWorkbook Is ThisWorkbook Then returnSheetName = ActiveSheet.Name

    oldAlerts = Application.DisplayAlerts
    Application.DisplayAlerts = False

    ' Khong dung OpenText. Mot so Excel Windows bao loi 1004 tai phuong thuc nay.
    Set csvWb = FindOpenWorkbookByPath(filePath)

    If csvWb Is Nothing Then
        Set csvWb = OpenCsvWorkbookCompatible(filePath, openDetail)
        openedHere = True
    Else
        openDetail = "Tai su dung workbook CSV dang mo."
    End If

    If csvWb Is Nothing Then
        Err.Raise vbObjectError + 1147, "ReadCsvWithExcel50", _
            "Khong mo duoc CSV bang Workbooks.Open. " & openDetail
    End If

    Set csvWs = csvWb.Worksheets(1)

    Set lastCell = csvWs.Cells.Find( _
        What:="*", _
        After:=csvWs.Cells(1, 1), _
        LookIn:=xlFormulas, _
        LookAt:=xlPart, _
        SearchOrder:=xlByRows, _
        SearchDirection:=xlPrevious, _
        MatchCase:=False)

    If lastCell Is Nothing Then
        Err.Raise vbObjectError + 1141, "ReadCsvWithExcel50", _
            "Excel mo duoc CSV nhung sheet khong co du lieu."
    End If
    lastRow = lastCell.Row

    Set lastCell = csvWs.Cells.Find( _
        What:="*", _
        After:=csvWs.Cells(1, 1), _
        LookIn:=xlFormulas, _
        LookAt:=xlPart, _
        SearchOrder:=xlByColumns, _
        SearchDirection:=xlPrevious, _
        MatchCase:=False)
    lastCol = lastCell.Column

    If lastRow > CSV_LAST_ROW + 1 Then
        Err.Raise vbObjectError + 1143, "ReadCsvWithExcel50", _
            "CSV co " & lastRow & " dong; vuot gioi han."
    End If

    rowCount = lastRow
    ReDim data(1 To rowCount, 1 To 50)

    If lastCol >= 50 Then
        For r = 1 To rowCount
            For c = 1 To 50
                cellValue = csvWs.Cells(r, c).Value2
                If IsError(cellValue) Then
                    Err.Raise vbObjectError + 1144, "ReadCsvWithExcel50", _
                        "Gia tri loi tai dong " & r & ", cot " & c & "."
                End If
                data(r, c) = cellValue
            Next c
        Next r

        detailText = openDetail & " Da doc " & rowCount & _
                     " dong, Excel tach " & lastCol & " cot."
    Else
        For r = 1 To rowCount
            lineText = BuildCsvLineFromOpenedRow(csvWs, r, lastCol)

            If Len(Trim$(lineText)) = 0 Then
                Err.Raise vbObjectError + 1145, "ReadCsvWithExcel50", _
                    "Dong " & r & " trong CSV bi trong."
            End If

            fields = ParseCsvLine(lineText)
            fieldCount = UBound(fields) - LBound(fields) + 1

            If fieldCount < 50 Then
                Err.Raise vbObjectError + 1146, "ReadCsvWithExcel50", _
                    "Dong " & r & " chi tach duoc " & fieldCount & _
                    " truong; can 50 truong."
            End If

            For c = 1 To 50
                data(r, c) = CleanCsvToken( _
                    CStr(fields(LBound(fields) + c - 1)))
            Next c
        Next r

        detailText = openDetail & " Excel mo thanh " & lastCol & _
                     " cot; bo doc da ghep va tach lai 50 truong."
    End If

    ReadCsvWithExcel50 = True

SafeExit:
    On Error Resume Next
    If openedHere Then
        If Not csvWb Is Nothing Then csvWb.Close SaveChanges:=False
    End If
    Application.DisplayAlerts = oldAlerts
    ThisWorkbook.Activate
    If returnSheetName <> "" Then
        ThisWorkbook.Worksheets(returnSheetName).Activate
    End If
    On Error GoTo 0
    Exit Function

EH:
    savedErrNumber = Err.Number
    savedErrDescription = Err.Description

    If savedErrNumber = 0 Then
        savedErrNumber = vbObjectError + 1149
        savedErrDescription = "Loi CSV khong xac dinh."
    End If

    detailText = "Ma " & savedErrNumber & ": " & savedErrDescription
    ReadCsvWithExcel50 = False
    Resume SafeExit
End Function

Private Function OpenCsvWorkbookCompatible(ByVal filePath As String, _
                                           ByRef detailText As String) As Workbook
    Dim wb As Workbook
    Dim err1Number As Long
    Dim err1Text As String
    Dim err2Number As Long
    Dim err2Text As String

    On Error Resume Next
    Err.Clear
    Set wb = Workbooks.Open( _
        Filename:=filePath, _
        ReadOnly:=True, _
        UpdateLinks:=False, _
        AddToMru:=False)
    err1Number = Err.Number
    err1Text = Err.Description
    Err.Clear
    On Error GoTo 0

    If Not wb Is Nothing Then
        detailText = "Mo bang Workbooks.Open mac dinh."
        Set OpenCsvWorkbookCompatible = wb
        Exit Function
    End If

    On Error Resume Next
    Err.Clear
    Set wb = Workbooks.Open( _
        Filename:=filePath, _
        ReadOnly:=True, _
        UpdateLinks:=False, _
        AddToMru:=False, _
        Local:=True)
    err2Number = Err.Number
    err2Text = Err.Description
    Err.Clear
    On Error GoTo 0

    If Not wb Is Nothing Then
        detailText = "Mo bang Workbooks.Open Local=True."
        Set OpenCsvWorkbookCompatible = wb
        Exit Function
    End If

    detailText = "Lan 1: " & err1Number & " - " & err1Text & _
                 "; Lan 2: " & err2Number & " - " & err2Text
End Function

Private Function FindOpenWorkbookByPath(ByVal filePath As String) As Workbook
    Dim wb As Workbook
    Dim wantedPath As String
    Dim currentPath As String

    wantedPath = NormalizePathForCompare(filePath)

    For Each wb In Application.Workbooks
        On Error Resume Next
        currentPath = NormalizePathForCompare(wb.FullName)
        On Error GoTo 0

        If currentPath <> "" And currentPath = wantedPath Then
            Set FindOpenWorkbookByPath = wb
            Exit Function
        End If
    Next wb
End Function

Private Function NormalizePathForCompare(ByVal pathText As String) As String
    Dim s As String

    s = Trim$(pathText)
    s = Replace(s, "\", "/")

#If Mac Then
    NormalizePathForCompare = s
#Else
    NormalizePathForCompare = LCase$(s)
#End If
End Function

Private Function BuildCsvLineFromOpenedRow(ByVal ws As Worksheet, _
                                           ByVal rowNumber As Long, _
                                           ByVal lastCol As Long) As String
    Dim c As Long
    Dim token As String
    Dim output As String

    If lastCol <= 1 Then
        BuildCsvLineFromOpenedRow = CStr(ws.Cells(rowNumber, 1).Value2)
        Exit Function
    End If

    For c = 1 To lastCol
        If c > 1 Then output = output & ","
        token = CStr(ws.Cells(rowNumber, c).Value2)
        output = output & EncodeCsvToken(token)
    Next c

    BuildCsvLineFromOpenedRow = output
End Function

Private Function EncodeCsvToken(ByVal token As String) As String
    Dim s As String

    s = token
    If InStr(1, s, Chr$(34), vbBinaryCompare) > 0 Then
        s = Replace(s, Chr$(34), Chr$(34) & Chr$(34))
    End If

    If InStr(1, s, ",", vbBinaryCompare) > 0 Or _
       InStr(1, s, Chr$(34), vbBinaryCompare) > 0 Or _
       InStr(1, s, vbCr, vbBinaryCompare) > 0 Or _
       InStr(1, s, vbLf, vbBinaryCompare) > 0 Then
        s = Chr$(34) & s & Chr$(34)
    End If

    EncodeCsvToken = s
End Function


Private Function ParseCsvLine(ByVal lineText As String) As Variant
    Dim result() As String
    Dim currentValue As String
    Dim inQuotes As Boolean
    Dim ch As String
    Dim i As Long
    Dim itemIndex As Long

    ReDim result(0 To 0)
    currentValue = ""
    itemIndex = 0

    For i = 1 To Len(lineText)
        ch = Mid$(lineText, i, 1)

        If ch = """" Then
            If inQuotes And i < Len(lineText) And Mid$(lineText, i + 1, 1) = """" Then
                currentValue = currentValue & """"
                i = i + 1
            Else
                inQuotes = Not inQuotes
            End If
        ElseIf ch = "," And Not inQuotes Then
            result(itemIndex) = currentValue
            itemIndex = itemIndex + 1
            ReDim Preserve result(0 To itemIndex)
            currentValue = ""
        Else
            currentValue = currentValue & ch
        End If
    Next i

    result(itemIndex) = currentValue
    ParseCsvLine = result
End Function

Private Function CleanCsvToken(ByVal valueText As String) As String
    Dim s As String
    s = Trim$(valueText)
    s = Replace(s, Chr$(34), "")
    s = Replace(s, Chr$(160), "")
    s = Replace(s, "ï»¿", "")
    On Error Resume Next
    s = Replace(s, ChrW$(&HFEFF), "")
    On Error GoTo 0
    CleanCsvToken = Trim$(s)
End Function

Private Function TryParseCsvNumber(ByVal valueText As String, _
                                   ByRef resultValue As Double) As Boolean
    Dim s As String
    Dim ch As String
    Dim i As Long
    Dim dotCount As Long
    Dim exponentCount As Long

    s = CleanCsvToken(valueText)
    s = Replace(s, " ", "")
    If s = "" Then Exit Function

    ' CSV cong to dung dau cham thap phan. Val khong phu thuoc dau thap phan he thong.
    For i = 1 To Len(s)
        ch = Mid$(s, i, 1)
        Select Case ch
            Case "0" To "9"
            Case "+", "-"
                If i > 1 Then
                    If UCase$(Mid$(s, i - 1, 1)) <> "E" Then Exit Function
                End If
            Case "."
                dotCount = dotCount + 1
                If dotCount > 1 Then Exit Function
            Case "e", "E"
                exponentCount = exponentCount + 1
                If exponentCount > 1 Then Exit Function
            Case Else
                Exit Function
        End Select
    Next i

    resultValue = Val(s)
    TryParseCsvNumber = True
End Function

Private Function FindCommandHeaderRow(ByVal ws As Worksheet) As Long
    Dim r As Long
    Dim maxRow As Long
    maxRow = Application.WorksheetFunction.Min(30, ws.UsedRange.Row + ws.UsedRange.Rows.Count - 1)

    For r = 1 To maxRow
        If InStr(1, UCase$(Trim$(CStr(ws.Cells(r, 1).Value2))), "ID", vbTextCompare) > 0 _
           And Len(Trim$(CStr(ws.Cells(r, 7).Value2))) > 0 _
           And Len(Trim$(CStr(ws.Cells(r, 16).Value2))) > 0 _
           And Len(Trim$(CStr(ws.Cells(r, 25).Value2))) > 0 Then
            FindCommandHeaderRow = r
            Exit Function
        End If
    Next r
End Function

Private Function FindKwhGiaoRowInTextBlock(ByRef textData As Variant, _
                                                ByVal rowCount As Long) As Long
    Dim r As Long

    For r = 1 To rowCount
        If UCase$(Trim$(CStr(textData(r, 2)))) = "KWHGIAO" Then
            FindKwhGiaoRowInTextBlock = r
            Exit Function
        End If
    Next r
End Function

Private Function FindKwhGiaoRow(ByRef data As Variant, ByVal rowCount As Long) As Long
    Dim r As Long
    For r = 1 To rowCount
        If UCase$(Trim$(CStr(data(r, 2)))) = "KWHGIAO" Then
            FindKwhGiaoRow = r
            Exit Function
        End If
    Next r
End Function

Private Function Has48NumericPeriods(ByRef data As Variant, ByVal dataRow As Long) As Boolean
    Dim c As Long
    For c = 3 To 50
        If Not IsNumeric(data(dataRow, c)) Then Exit Function
    Next c
    Has48NumericPeriods = True
End Function

Private Function IsGenericColumnHeader(ByVal firstValue As String, ByVal secondValue As String) As Boolean
    firstValue = UCase$(Trim$(firstValue))
    secondValue = UCase$(Trim$(secondValue))
    IsGenericColumnHeader = (Left$(firstValue, 6) = "COLUMN" Or Left$(secondValue, 6) = "COLUMN")
End Function

Private Function PickFile(ByVal titleText As String, ByVal windowsFilter As String) As Variant
#If Mac Then
    PickFile = Application.GetOpenFilename(, , titleText)
#Else
    PickFile = Application.GetOpenFilename(windowsFilter, , titleText)
#End If
End Function

' LEGACY UNUSED: tu v1.1.2, nhap CSV khong goi ham nay.
Private Function ParseCsvDate(ByVal valueIn As Variant, ByRef resultDate As Date) As Boolean
    Dim s As String
    Dim datePart As String
    Dim p() As String
    Dim a As Long, b As Long, c As Long
    Dim d As Long, m As Long, y As Long
    Dim serialValue As Double
    Dim parsedDate As Date

    On Error GoTo Fail

    ' Chi chap nhan so serial Excel that. Khong dung IsDate/CDate cho chuoi
    ' vi hai ham do phu thuoc Regional Settings cua tung may.
    If IsNumeric(valueIn) And VarType(valueIn) <> vbString Then
        serialValue = CDbl(valueIn)
        If serialValue >= 1 And serialValue < 100000 Then
            resultDate = Int(serialValue)
            ParseCsvDate = True
            Exit Function
        End If
    End If

    s = Trim$(CleanCsvToken(CStr(valueIn)))
    If s = "" Then GoTo Fail

    ' Bo phan gio neu nguon co dang dd-mm-yyyy hh:mm:ss hoac ISO co ky tu T.
    s = Replace(s, "T", " ")
    If InStr(1, s, " ", vbBinaryCompare) > 0 Then
        datePart = Left$(s, InStr(1, s, " ", vbBinaryCompare) - 1)
    Else
        datePart = s
    End If

    datePart = Replace(datePart, "/", "-")
    datePart = Replace(datePart, ".", "-")
    p = Split(datePart, "-")
    If UBound(p) <> 2 Then GoTo Fail

    a = CLng(Trim$(p(0)))
    b = CLng(Trim$(p(1)))
    c = CLng(Trim$(p(2)))

    ' Ho tro hai dang ro rang:
    ' - dd-mm-yy / dd-mm-yyyy (dinh dang cong to dang dung)
    ' - yyyy-mm-dd (ISO)
    If Len(Trim$(p(0))) = 4 Then
        y = a
        m = b
        d = c
    Else
        d = a
        m = b
        y = c
        If y < 100 Then y = 2000 + y
    End If

    If y < 1900 Or y > 2199 Then GoTo Fail
    If m < 1 Or m > 12 Then GoTo Fail
    If d < 1 Or d > 31 Then GoTo Fail

    parsedDate = DateSerial(y, m, d)

    ' DateSerial tu dieu chinh ngay khong hop le, nen phai doi chieu lai.
    If Year(parsedDate) <> y Or Month(parsedDate) <> m Or Day(parsedDate) <> d Then
        GoTo Fail
    End If

    resultDate = parsedDate
    ParseCsvDate = True
    Exit Function
Fail:
    ParseCsvDate = False
End Function

Private Function ParseUserDate(ByVal dateText As String, ByRef resultDate As Date) As Boolean
    Dim s As String
    Dim p() As String
    Dim d As Long, m As Long, y As Long

    On Error GoTo Fail
    s = Trim$(dateText)
    s = Replace(s, "-", "/")
    p = Split(s, "/")
    If UBound(p) <> 2 Then GoTo Fail

    d = CLng(p(0))
    m = CLng(p(1))
    y = CLng(p(2))
    If y < 100 Then y = 2000 + y
    resultDate = DateSerial(y, m, d)
    ParseUserDate = True
    Exit Function
Fail:
    ParseUserDate = False
End Function

Private Function ConfirmOrSetDate(ByVal importedDate As Date, _
                                  ByVal sourceName As String) As Boolean
    Dim ws As Worksheet
    Dim currentValue As Variant

    Set ws = ThisWorkbook.Worksheets(SH_CFG)
    currentValue = ws.Range("B4").Value

    If Not IsDate(currentValue) Then
        MsgBox "Chua cai dat ngay tinh. Hay cap nhat cai dat truoc khi nhap " & _
               sourceName & ".", vbExclamation, APP_NAME
        Exit Function
    End If

    If Int(CDbl(CDate(currentValue))) <> Int(CDbl(importedDate)) Then
        MsgBox sourceName & " co ngay " & _
               Format$(importedDate, "dd/mm/yyyy") & _
               ", khac ngay dang cai dat " & _
               Format$(CDate(currentValue), "dd/mm/yyyy") & "." & vbCrLf & _
               "Du lieu khong duoc nhap va CAI_DAT khong bi thay doi.", _
               vbExclamation, APP_NAME
        Exit Function
    End If

    ConfirmOrSetDate = True
End Function

Private Function NormalizeUnitName(ByVal unitName As String) As String
    Dim s As String
    s = UCase$(Trim$(unitName))
    If Left$(s, 2) = "S1" Then
        NormalizeUnitName = "S1"
    ElseIf Left$(s, 2) = "S2" Then
        NormalizeUnitName = "S2"
    Else
        NormalizeUnitName = s
    End If
End Function

Private Function CountWorkbookFormulaErrors(ByRef details As String) As Long
    Dim ws As Worksheet
    Dim errorCells As Range
    Dim cell As Range
    Dim total As Long
    Dim shown As Long

    details = ""
    For Each ws In ThisWorkbook.Worksheets
        Set errorCells = Nothing
        On Error Resume Next
        Set errorCells = ws.UsedRange.SpecialCells(xlCellTypeFormulas, xlErrors)
        On Error GoTo 0

        If Not errorCells Is Nothing Then
            total = total + errorCells.Cells.Count
            For Each cell In errorCells.Cells
                If shown < 15 Then
                    If details <> "" Then details = details & vbCrLf
                    details = details & ws.Name & "!" & cell.Address(False, False) & " = " & CStr(cell.Text)
                    shown = shown + 1
                End If
            Next cell
        End If
    Next ws

    CountWorkbookFormulaErrors = total
End Function

Private Sub UpdateQduSignDisplay(ByVal reportWs As Worksheet)
    Dim pQdcColumn As Long
    Dim signColumn As Long
    Dim qduColumn As Long
    Dim c As Long
    Dim r As Long
    Dim headerText As String
    Dim qduValue As Variant
    Dim numericQdu As Double
    Dim lastPeriodRow As Long

    On Error GoTo EH

    lastPeriodRow = 3 + PERIOD_COUNT

    ' Tim cot P_Qdc trong hang tieu de. Theo mau bao cao:
    ' cot ngay truoc P_Qdc la Qdu am/duong,
    ' cot truoc nua la gia tri Qdu.
    For c = 1 To 20
        headerText = UCase$(Trim$(CStr(reportWs.Cells(3, c).Value2)))
        headerText = Replace(headerText, " ", "")
        If InStr(1, headerText, "P_QDC", vbTextCompare) > 0 Or _
           InStr(1, headerText, "PQDC", vbTextCompare) > 0 Then
            pQdcColumn = c
            Exit For
        End If
    Next c

    If pQdcColumn < 3 Then Exit Sub

    signColumn = pQdcColumn - 1
    qduColumn = pQdcColumn - 2

    For r = 4 To lastPeriodRow
        qduValue = reportWs.Cells(r, qduColumn).Value2

        If IsError(qduValue) Or Len(Trim$(CStr(qduValue))) = 0 Then
            reportWs.Cells(r, signColumn).ClearContents
        ElseIf IsNumeric(qduValue) Then
            numericQdu = CDbl(qduValue)

            If Abs(numericQdu) < 0.0000000001 Then
                reportWs.Cells(r, signColumn).ClearContents
            ElseIf numericQdu > 0 Then
                reportWs.Cells(r, signColumn).Value = QduPositiveLabel()
            Else
                reportWs.Cells(r, signColumn).Value = QduNegativeLabel()
            End If
        Else
            reportWs.Cells(r, signColumn).ClearContents
        End If
    Next r
    Exit Sub

EH:
    Err.Raise vbObjectError + 1685, "UpdateQduSignDisplay", _
              "Khong cap nhat duoc cot Qdu am/duong: " & Err.Description
End Sub

Private Function QduPositiveLabel() As String
    ' "Dương" dung ChrW de khong phu thuoc encoding file BAS.
    QduPositiveLabel = "D" & ChrW(&H1B0) & ChrW(&H1A1) & "ng"
End Function

Private Function QduNegativeLabel() As String
    ' "Âm"
    QduNegativeLabel = ChrW(&HC2) & "m"
End Function

Private Function SumNumericCells(ByVal sourceRange As Range) As Double
    Dim cell As Range
    Dim valueIn As Variant

    On Error GoTo EH

    For Each cell In sourceRange.Cells
        valueIn = cell.Value2

        If Not IsError(valueIn) Then
            If IsNumeric(valueIn) Then
                SumNumericCells = SumNumericCells + CDbl(valueIn)
            End If
        End If
    Next cell
    Exit Function

EH:
    Err.Raise vbObjectError + 1681, "SumNumericCells", _
              "Khong cong duoc vung " & sourceRange.Worksheet.Name & _
              "!" & sourceRange.Address(False, False) & _
              ": " & Err.Description
End Function

Private Function CountChangedQddPeriods(ByVal initialPower As Variant) As Long
    Dim ws As Worksheet
    Dim i As Long
    Dim v As Variant
    Dim p0 As Double

    If Not IsNumeric(initialPower) Then Exit Function
    p0 = CDbl(initialPower)
    Set ws = ThisWorkbook.Worksheets(SH_CALC)

    For i = 2 To 49
        v = ws.Cells(i, 4).Value
        If IsNumeric(v) Then
            If Abs(CDbl(v) - p0) > EPS Then CountChangedQddPeriods = CountChangedQddPeriods + 1
        End If
    Next i
End Function

Private Function MaxAbsDifference(ByVal calculatedRange As Range, ByVal originalRange As Range) As Double
    Dim i As Long
    Dim d As Double
    For i = 1 To PERIOD_COUNT
        If IsNumeric(calculatedRange.Cells(i, 1).Value) And IsNumeric(originalRange.Cells(i, 1).Value) Then
            d = Abs(CDbl(calculatedRange.Cells(i, 1).Value) - CDbl(originalRange.Cells(i, 1).Value))
            If d > MaxAbsDifference Then MaxAbsDifference = d
        End If
    Next i
End Function

Private Function SafeLong(ByVal valueIn As Variant) As Long
    If IsNumeric(valueIn) Then SafeLong = CLng(valueIn)
End Function

Private Function GetOrCreateControlSheet() As Worksheet
    On Error Resume Next
    Set GetOrCreateControlSheet = ThisWorkbook.Worksheets(SH_CTRL)
    On Error GoTo 0

    If GetOrCreateControlSheet Is Nothing Then
        Set GetOrCreateControlSheet = ThisWorkbook.Worksheets.Add(Before:=ThisWorkbook.Worksheets(1))
        GetOrCreateControlSheet.Name = SH_CTRL
        GetOrCreateControlSheet.Range("A1").Value = "QDU VBA CONTROL"
    End If
End Function

Private Sub AddQduButton(ByVal ws As Worksheet, ByVal shapeName As String, _
                         ByVal caption As String, ByVal macroName As String, _
                         ByVal targetAddress As String)
    Dim target As Range
    Dim shp As Shape

    Set target = ws.Range(targetAddress)
    Set shp = ws.Shapes.AddShape(5, target.Left + 2, target.Top + 2, target.Width - 4, target.Height - 4)

    With shp
        .Name = shapeName
        .OnAction = "'" & ThisWorkbook.Name & "'!" & macroName
        .Fill.ForeColor.RGB = RGB(31, 78, 121)
        .Line.ForeColor.RGB = RGB(23, 54, 93)
        .TextFrame.Characters.Text = caption
        .TextFrame.Characters.Font.Bold = True
        .TextFrame.Characters.Font.Color = RGB(255, 255, 255)
        .TextFrame.Characters.Font.Size = 10
        .TextFrame.HorizontalAlignment = xlHAlignCenter
        .Placement = xlMoveAndSize
    End With
End Sub

Private Sub DeleteQduButtons(ByVal ws As Worksheet)
    Dim i As Long
    For i = ws.Shapes.Count To 1 Step -1
        If Left$(ws.Shapes(i).Name, 6) = "qduBtn" Then ws.Shapes(i).Delete
    Next i
End Sub

Private Sub WriteLog(ByVal actionName As String, ByVal resultText As String, ByVal detailText As String)
    Dim ws As Worksheet
    Dim nextRow As Long

    ' Logging must never interrupt the main workflow.
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(SH_LOG)
    If ws Is Nothing Then
        On Error Resume Next
        Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        ws.Name = SH_LOG
        ws.Cells(1, 1).Value = "Thoi gian"
        ws.Cells(1, 2).Value = "Thao tac"
        ws.Cells(1, 3).Value = "Ket qua"
        ws.Cells(1, 4).Value = "Chi tiet"
        ws.Cells(1, 5).Value = "Nguoi dung"
        ws.Cells(1, 6).Value = "Phien ban"
        ws.Cells(1, 7).Value = "Workbook"
        ws.Range("A1:G1").Font.Bold = True
        ws.Range("A1:G1").Interior.Color = RGB(31, 78, 121)
        ws.Range("A1:G1").Font.Color = RGB(255, 255, 255)
        ws.Columns("A").NumberFormat = "dd/mm/yyyy hh:mm:ss"
        ws.Columns("A:G").ColumnWidth = 24
        ws.Columns("D").ColumnWidth = 70
    End If

    If ws Is Nothing Then Exit Sub

    If Trim$(CStr(ws.Cells(1, 6).Value)) = "" Then
        ws.Cells(1, 6).Value = "Phien ban"
    End If
    If Trim$(CStr(ws.Cells(1, 7).Value)) = "" Then
        ws.Cells(1, 7).Value = "Workbook"
    End If
    ws.Range("A1:G1").Font.Bold = True
    ws.Columns("A").NumberFormat = "dd/mm/yyyy hh:mm:ss"

    nextRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row + 1
    If nextRow < 2 Then nextRow = 2

    ws.Cells(nextRow, 1).Value = Now
    ws.Cells(nextRow, 2).Value = actionName
    ws.Cells(nextRow, 3).Value = resultText
    ws.Cells(nextRow, 4).Value = detailText
    ws.Cells(nextRow, 5).Value = GetUserName()
    ws.Cells(nextRow, 6).Value = APP_VERSION
    ws.Cells(nextRow, 7).Value = ThisWorkbook.Name
End Sub

' ============================================================
' MONTHLY HISTORY AND REPORT SECTION - SAME STANDARD MODULE
' No Application.Run, no cross-module lookup, no macro-name string.
' ============================================================

' MONTHLY REPORT SECTION v1.3.1 - export creates fresh sheets without Merge/Copy/Paste
' Isolated from the calculation core. Errors here must not block Qdd/Qdu calculation.


Public Function QduMonth_Initialize() As Boolean
    On Error GoTo EH
    GetOrCreateHistorySheet
    QduMonth_Initialize = True
    Exit Function
EH:
    MonthWriteLog "KHOI_TAO", "LOI", Err.Number & " - " & Err.Description
End Function

Public Function QduMonth_SelfTest() As Boolean
    Dim ws As Worksheet

    On Error GoTo EH
    If Not MonthSheetExists(M_SH_CFG) Then Err.Raise vbObjectError + 2101, , "Thieu CAI_DAT."
    If Not MonthSheetExists(M_SH_CALC) Then Err.Raise vbObjectError + 2102, , "Thieu TINH_TOAN."
    If Not MonthSheetExists(M_SH_REPORT) Then Err.Raise vbObjectError + 2103, , "Thieu BAO_CAO_QDU."

    Set ws = GetOrCreateHistorySheet()
    If UCase$(Trim$(CStr(ws.Range("B3").Value))) <> "NGAY" Then
        Err.Raise vbObjectError + 2104, , "Tieu de LICH_SU_THANG khong hop le."
    End If

    QduMonth_SelfTest = True
    Exit Function
EH:
    MonthWriteLog "SELF_TEST", "LOI", Err.Number & " - " & Err.Description
End Function

Public Function QduMonth_OpenHistory() As Boolean
    Dim ws As Worksheet

    On Error GoTo EH
    Set ws = GetOrCreateHistorySheet()
    ws.Visible = xlSheetVisible
    ws.Activate
    QduMonth_OpenHistory = True
    Exit Function
EH:
    MsgBox "Khong mo duoc lich su thang:" & vbCrLf & _
           Err.Number & " - " & Err.Description, _
           vbExclamation, M_APP_NAME
    MonthWriteLog "MO_LICH_SU", "LOI", Err.Number & " - " & Err.Description
End Function

Public Function QduMonth_SaveCurrentDay() As Boolean
    Dim cfg As Worksheet
    Dim calcWs As Worksheet
    Dim reportWs As Worksheet
    Dim historyWs As Worksheet
    Dim snapshotWs As Worksheet
    Dim targetDate As Date
    Dim unitName As String
    Dim monthKeyText As String
    Dim snapshotName As String
    Dim historyRow As Long
    Dim oldAlerts As Boolean
    Dim statusText As String
    Dim savedErrNumber As Long
    Dim savedErrDescription As String
    Dim dateDetail As String

    On Error GoTo EH

    Set cfg = ThisWorkbook.Worksheets(M_SH_CFG)
    Set calcWs = ThisWorkbook.Worksheets(M_SH_CALC)
    Set reportWs = ThisWorkbook.Worksheets(M_SH_REPORT)
    Set historyWs = GetOrCreateHistorySheet()

    If Not TryReadExcelDate(cfg.Range("B4").Value2, targetDate, dateDetail) Then
        Err.Raise vbObjectError + 2111, , _
            "Ngay tinh CAI_DAT!B4 khong hop le. " & dateDetail
    End If

    unitName = NormalizeMonthUnit(CStr(cfg.Range("B5").Value2))
    If unitName <> "S1" And unitName <> "S2" Then
        Err.Raise vbObjectError + 2112, , "To may phai la S1 hoac S2."
    End If

    monthKeyText = MonthKeyFromDate(targetDate)
    If Not EnsureHistoryMonth(historyWs, monthKeyText) Then Exit Function

    snapshotName = SnapshotSheetName(targetDate, unitName)

    oldAlerts = Application.DisplayAlerts
    Application.DisplayAlerts = False
    On Error Resume Next
    ThisWorkbook.Worksheets(snapshotName).Delete
    Err.Clear
    On Error GoTo EH
    Application.DisplayAlerts = oldAlerts

    reportWs.Copy After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count)
    Set snapshotWs = ActiveSheet
    snapshotWs.Name = snapshotName

    ConvertSnapshotFormulasToValues snapshotWs

    If snapshotWs.Range("A1").MergeCells Then
        snapshotWs.Range("A1").MergeArea.Cells(1, 1).Value = _
            CStr(snapshotWs.Range("A1").MergeArea.Cells(1, 1).Value) & _
            " - " & DateTextDMY(targetDate) & " - " & unitName
    Else
        snapshotWs.Range("A1").Value = CStr(snapshotWs.Range("A1").Value) & _
            " - " & DateTextDMY(targetDate) & " - " & unitName
    End If
    snapshotWs.Visible = xlSheetVeryHidden

    historyRow = FindHistoryRow(historyWs, targetDate, unitName)
    If historyRow = 0 Then
        historyRow = historyWs.Cells(historyWs.Rows.Count, "B").End(xlUp).Row + 1
        If historyRow < M_FIRST_ROW Then historyRow = M_FIRST_ROW
    End If

    statusText = CStr(ThisWorkbook.Worksheets(M_SH_CTRL).Range("B10").Value2)

    historyWs.Cells(historyRow, 1).Value = historyRow - M_FIRST_ROW + 1
    historyWs.Cells(historyRow, 2).Value = targetDate
    historyWs.Cells(historyRow, 3).Value = unitName
    historyWs.Cells(historyRow, 4).Value = cfg.Range("B6").Value2
    historyWs.Cells(historyRow, 5).Value = cfg.Range("B14").Value2
    historyWs.Cells(historyRow, 6).Value = cfg.Range("B10").Value2
    historyWs.Cells(historyRow, 7).Value = Application.Sum(calcWs.Range("F2:F49"))
    historyWs.Cells(historyRow, 8).Value = Application.Sum(calcWs.Range("J2:J49"))
    historyWs.Cells(historyRow, 9).Value = Application.Sum(calcWs.Range("E2:E49"))
    historyWs.Cells(historyRow, 10).Value = Application.Sum(calcWs.Range("K2:K49"))
    historyWs.Cells(historyRow, 11).Value = statusText
    historyWs.Cells(historyRow, 12).Value = snapshotName
    historyWs.Cells(historyRow, 13).Value = Now

    historyWs.Range("B" & historyRow).NumberFormat = "dd/mm/yyyy"
    historyWs.Range("D" & historyRow & ":J" & historyRow).NumberFormat = "0.000"
    historyWs.Range("M" & historyRow).NumberFormat = "dd/mm/yyyy hh:mm:ss"

    SetStoredMonthKey historyWs, monthKeyText
    SortAndRenumberHistory historyWs

    MonthWriteLog "LUU_NGAY", "OK", _
                  DateTextDMY(targetDate) & " - " & unitName & _
                  "; thang=" & monthKeyText
    QduMonth_SaveCurrentDay = True
    Exit Function

EH:
    savedErrNumber = Err.Number
    savedErrDescription = Err.Description

    On Error Resume Next
    Application.DisplayAlerts = oldAlerts
    If Not snapshotWs Is Nothing Then snapshotWs.Visible = xlSheetVeryHidden
    On Error GoTo 0

    MonthWriteLog "LUU_NGAY", "LOI", _
                  savedErrNumber & " - " & savedErrDescription

    MsgBox "Khong luu duoc lich su thang:" & vbCrLf & _
           "Ma loi: " & savedErrNumber & vbCrLf & savedErrDescription, _
           vbExclamation, M_APP_NAME
End Function

Public Function QduMonth_ExportMonthlyReport() As Boolean
    Dim historyWs As Worksheet
    Dim outWb As Workbook
    Dim summaryWs As Worksheet
    Dim usedNames As Collection
    Dim exportedDates As Collection
    Dim savePath As Variant
    Dim monthKeyText As String
    Dim defaultName As String
    Dim lastRow As Long
    Dim historyRow As Long
    Dim outputRow As Long
    Dim targetDate As Date
    Dim unitName As String
    Dim dateKeyText As String
    Dim dateDetail As String
    Dim oldAlerts As Boolean
    Dim usedAutomaticPath As Boolean
    Dim stageName As String
    Dim savedErrNumber As Long
    Dim savedErrDescription As String

    On Error GoTo EH
    stageName = "Khoi tao xuat thang"

    Set historyWs = GetOrCreateHistorySheet()
    If Not HistoryHasData(historyWs) Then
        Err.Raise vbObjectError + 2121, , "Chua co ngay nao trong lich su thang."
    End If

    monthKeyText = GetStoredMonthKey(historyWs)
    If Not IsValidMonthKey(monthKeyText) Then
        Err.Raise vbObjectError + 2122, , "Ma thang lich su khong hop le."
    End If

    defaultName = "Bao_cao_Qdd_Qdu_Thang_" & _
                  Left$(monthKeyText, 4) & "_" & Right$(monthKeyText, 2) & ".xlsx"

    stageName = "Chon duong dan luu"
    savePath = MonthPickSavePath(defaultName, usedAutomaticPath)
    If VarType(savePath) = vbBoolean Then Exit Function

    stageName = "Tao workbook bao cao"
    Set outWb = Workbooks.Add(xlWBATWorksheet)
    Set summaryWs = outWb.Worksheets(1)
    summaryWs.Name = "TONG_HOP"
    BuildSummaryHeader summaryWs, monthKeyText

    Set usedNames = New Collection
    usedNames.Add "TONG_HOP"
    Set exportedDates = New Collection

    lastRow = historyWs.Cells(historyWs.Rows.Count, "B").End(xlUp).Row
    outputRow = 4

    For historyRow = M_FIRST_ROW To lastRow
        If TryReadExcelDate(historyWs.Cells(historyRow, 2).Value2, _
                            targetDate, dateDetail) Then
            unitName = NormalizeMonthUnit(CStr(historyWs.Cells(historyRow, 3).Value2))

            summaryWs.Cells(outputRow, 1).Value = targetDate
            summaryWs.Cells(outputRow, 2).Value = unitName
            summaryWs.Cells(outputRow, 3).Resize(1, 7).Value = _
                historyWs.Cells(historyRow, 4).Resize(1, 7).Value
            summaryWs.Cells(outputRow, 10).Value = historyWs.Cells(historyRow, 11).Value
            outputRow = outputRow + 1

            dateKeyText = DateKeyFromDate(targetDate)
            If Not CollectionHasText(exportedDates, dateKeyText) Then
                stageName = "Tao sheet ngay " & DateTextDMY(targetDate)
                CreateCombinedDailySheet outWb, targetDate, usedNames
                exportedDates.Add dateKeyText
            End If
        End If
    Next historyRow

    stageName = "Hoan thien tong hop"
    FinishSummarySheet summaryWs, outputRow

    oldAlerts = Application.DisplayAlerts
    Application.DisplayAlerts = False
    stageName = "Luu file xlsx"
    outWb.SaveAs Filename:=CStr(savePath), FileFormat:=xlOpenXMLWorkbook
    outWb.Close SaveChanges:=False
    Application.DisplayAlerts = oldAlerts

    MonthWriteLog "XUAT_THANG", "OK", _
                  CStr(savePath) & _
                  IIf(usedAutomaticPath, "; duong_dan_tu_dong=TRUE", "")

    If usedAutomaticPath Then
        MsgBox "Hop chon noi luu cua Excel Mac khong hoat dong." & vbCrLf & _
               "Bao cao da duoc tu dong luu tai:" & vbCrLf & _
               CStr(savePath), _
               vbInformation, M_APP_NAME
    Else
        MsgBox "Da xuat bao cao thang:" & vbCrLf & CStr(savePath), _
               vbInformation, M_APP_NAME
    End If
    QduMonth_ExportMonthlyReport = True
    Exit Function

EH:
    savedErrNumber = Err.Number
    savedErrDescription = Err.Description

    On Error Resume Next
    Application.DisplayAlerts = oldAlerts
    If Not outWb Is Nothing Then outWb.Close SaveChanges:=False
    On Error GoTo 0

    MonthWriteLog "XUAT_THANG", "LOI", _
                  savedErrNumber & " - " & savedErrDescription
    MsgBox "Khong xuat duoc bao cao thang:" & vbCrLf & _
           "Buoc loi: " & stageName & vbCrLf & _
           "Ma loi: " & savedErrNumber & vbCrLf & savedErrDescription, _
           vbExclamation, M_APP_NAME
End Function

Private Function GetOrCreateHistorySheet() As Worksheet
    Dim ws As Worksheet
    Dim existingKey As String

    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(M_SH_HISTORY)
    On Error GoTo 0

    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add( _
            After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        ws.Name = M_SH_HISTORY
    End If

    With ws
        .Visible = xlSheetVisible
        If Not .Range("A1").MergeCells Then .Range("A1:M1").Merge

        .Range("A1").Value = "LICH SU BAO CAO Qdd / Qdu THEO THANG"
        .Range("A1").Font.Bold = True
        .Range("A1").Font.Size = 14
        .Range("A1").HorizontalAlignment = xlCenter
        .Range("A1").Interior.Color = RGB(31, 78, 121)
        .Range("A1").Font.Color = RGB(255, 255, 255)

        .Range("A2").Value = "Thang hien tai"
        .Range("D2").Value = "Quy tac"
        .Range("E2").Value = "Cung ngay/to may: cap nhat; sang thang moi: xuat/reset."

        .Cells(3, 1).Value = "STT"
        .Cells(3, 2).Value = "Ngay"
        .Cells(3, 3).Value = "To may"
        .Cells(3, 4).Value = "P dau ngay"
        .Cells(3, 5).Value = "P cuoi ngay"
        .Cells(3, 6).Value = "So lenh"
        .Cells(3, 7).Value = "Tong Qdc"
        .Cells(3, 8).Value = "Tong Qmp"
        .Cells(3, 9).Value = "Tong Qdd_V"
        .Cells(3, 10).Value = "Tong Qdu"
        .Cells(3, 11).Value = "Trang thai"
        .Cells(3, 12).Value = "Snapshot"
        .Cells(3, 13).Value = "Thoi diem luu"

        .Range("A3:M3").Font.Bold = True
        .Range("A3:M3").Interior.Color = RGB(217, 234, 247)
        .Range("A3:M3").HorizontalAlignment = xlCenter
        .Columns("A").ColumnWidth = 7
        .Columns("B").ColumnWidth = 12
        .Columns("C").ColumnWidth = 10
        .Columns("D:J").ColumnWidth = 14
        .Columns("K").ColumnWidth = 32
        .Columns("L").ColumnWidth = 22
        .Columns("M").ColumnWidth = 20
        .Columns("B").NumberFormat = "dd/mm/yyyy"
        .Columns("D:J").NumberFormat = "0.000"
        .Columns("M").NumberFormat = "dd/mm/yyyy hh:mm:ss"
        .Range("B2").NumberFormat = "@"
        .Range(M_MONTH_KEY_CELL).NumberFormat = "@"
        .Columns("N").Hidden = True
    End With

    existingKey = GetStoredMonthKey(ws)
    If IsValidMonthKey(existingKey) Then SetStoredMonthKey ws, existingKey

    Set GetOrCreateHistorySheet = ws
End Function

Private Function EnsureHistoryMonth(ByVal historyWs As Worksheet, _
                                    ByVal newMonthKey As String) As Boolean
    Dim oldMonthKey As String
    Dim answer As VbMsgBoxResult

    oldMonthKey = GetStoredMonthKey(historyWs)

    If oldMonthKey = "" Then
        SetStoredMonthKey historyWs, newMonthKey
        EnsureHistoryMonth = True
        Exit Function
    End If

    If oldMonthKey = newMonthKey Then
        EnsureHistoryMonth = True
        Exit Function
    End If

    If HistoryHasData(historyWs) Then
        answer = MsgBox( _
            "Da sang thang moi." & vbCrLf & _
            "Thang dang luu: " & Right$(oldMonthKey, 2) & "/" & Left$(oldMonthKey, 4) & vbCrLf & _
            "Thang ngay moi: " & Right$(newMonthKey, 2) & "/" & Left$(newMonthKey, 4) & vbCrLf & vbCrLf & _
            "YES: Xuat thang cu roi reset." & vbCrLf & _
            "NO: Reset khong xuat." & vbCrLf & _
            "CANCEL: Dung lai.", _
            vbYesNoCancel + vbQuestion, M_APP_NAME)

        If answer = vbCancel Then Exit Function
        If answer = vbYes Then
            If Not QduMonth_ExportMonthlyReport() Then Exit Function
        Else
            answer = MsgBox("Reset lich su thang cu ma khong xuat?", _
                            vbYesNo + vbExclamation, M_APP_NAME)
            If answer <> vbYes Then Exit Function
        End If
    End If

    ResetHistory historyWs, newMonthKey
    EnsureHistoryMonth = True
End Function

Private Sub ResetHistory(ByVal historyWs As Worksheet, _
                         ByVal newMonthKey As String)
    Dim ws As Worksheet
    Dim namesToDelete As Collection
    Dim item As Variant
    Dim oldAlerts As Boolean
    Dim lastRow As Long

    Set namesToDelete = New Collection
    For Each ws In ThisWorkbook.Worksheets
        If Left$(ws.Name, Len(M_PREFIX)) = M_PREFIX Then namesToDelete.Add ws.Name
    Next ws

    oldAlerts = Application.DisplayAlerts
    Application.DisplayAlerts = False
    For Each item In namesToDelete
        ThisWorkbook.Worksheets(CStr(item)).Delete
    Next item
    Application.DisplayAlerts = oldAlerts

    lastRow = historyWs.Cells(historyWs.Rows.Count, "B").End(xlUp).Row
    If lastRow >= M_FIRST_ROW Then historyWs.Range("A4:M" & lastRow).ClearContents

    SetStoredMonthKey historyWs, newMonthKey
    MonthWriteLog "RESET_THANG", "OK", newMonthKey
End Sub

Private Function HistoryHasData(ByVal historyWs As Worksheet) As Boolean
    HistoryHasData = _
        (historyWs.Cells(historyWs.Rows.Count, "B").End(xlUp).Row >= M_FIRST_ROW)
End Function

Private Function FindHistoryRow(ByVal historyWs As Worksheet, _
                                ByVal targetDate As Date, _
                                ByVal unitName As String) As Long
    Dim lastRow As Long
    Dim r As Long
    Dim rowDate As Date
    Dim dateDetail As String

    lastRow = historyWs.Cells(historyWs.Rows.Count, "B").End(xlUp).Row
    If lastRow < M_FIRST_ROW Then Exit Function

    For r = M_FIRST_ROW To lastRow
        If TryReadExcelDate(historyWs.Cells(r, 2).Value2, rowDate, dateDetail) Then
            If CLng(rowDate) = CLng(targetDate) And _
               UCase$(Trim$(CStr(historyWs.Cells(r, 3).Value2))) = _
               UCase$(Trim$(unitName)) Then
                FindHistoryRow = r
                Exit Function
            End If
        End If
    Next r
End Function

Private Sub SortAndRenumberHistory(ByVal historyWs As Worksheet)
    Dim lastRow As Long
    Dim r As Long

    lastRow = historyWs.Cells(historyWs.Rows.Count, "B").End(xlUp).Row
    If lastRow < M_FIRST_ROW Then Exit Sub

    If lastRow > M_FIRST_ROW Then
        historyWs.Range("A3:M" & lastRow).Sort _
            Key1:=historyWs.Range("B4"), Order1:=xlAscending, _
            Key2:=historyWs.Range("C4"), Order2:=xlAscending, _
            Header:=xlYes
    End If

    For r = M_FIRST_ROW To lastRow
        historyWs.Cells(r, 1).Value = r - M_FIRST_ROW + 1
    Next r
End Sub

Private Sub ConvertSnapshotFormulasToValues(ByVal ws As Worksheet)
    Dim formulaRange As Range
    Dim cell As Range
    Dim topLeft As Range

    On Error Resume Next
    Set formulaRange = ws.UsedRange.SpecialCells(xlCellTypeFormulas)
    On Error GoTo 0
    If formulaRange Is Nothing Then Exit Sub

    For Each cell In formulaRange.Cells
        If cell.MergeCells Then
            Set topLeft = cell.MergeArea.Cells(1, 1)
            If cell.Address = topLeft.Address Then
                topLeft.Value2 = topLeft.Value2
            End If
        Else
            cell.Value2 = cell.Value2
        End If
    Next cell
End Sub

Private Sub BuildSummaryHeader(ByVal ws As Worksheet, _
                               ByVal monthKeyText As String)
    With ws
        .Range("A1:J1").ClearContents
        .Range("A1:J1").UnMerge
        .Range("A1").Value = "BAO CAO Qdd / Qdu THANG " & _
                             Right$(monthKeyText, 2) & "/" & Left$(monthKeyText, 4)
        .Range("A1:J1").HorizontalAlignment = xlCenterAcrossSelection
        .Range("A1:J1").VerticalAlignment = xlCenter
        .Range("A1:J1").Interior.Color = RGB(31, 78, 121)
        .Range("A1:J1").Font.Color = RGB(255, 255, 255)
        .Range("A1:J1").Font.Bold = True
        .Range("A1:J1").Font.Size = 16
        .Rows(1).RowHeight = 26

        .Cells(3, 1).Value = "Ngay"
        .Cells(3, 2).Value = "To may"
        .Cells(3, 3).Value = "P dau ngay"
        .Cells(3, 4).Value = "P cuoi ngay"
        .Cells(3, 5).Value = "So lenh"
        .Cells(3, 6).Value = "Tong Qdc"
        .Cells(3, 7).Value = "Tong Qmp"
        .Cells(3, 8).Value = "Tong Qdd_V"
        .Cells(3, 9).Value = "Tong Qdu"
        .Cells(3, 10).Value = "Trang thai"
        .Range("A3:J3").Font.Bold = True
        .Range("A3:J3").Interior.Color = RGB(217, 234, 247)
        .Range("A3:J3").HorizontalAlignment = xlCenter
    End With
End Sub

Private Sub FinishSummarySheet(ByVal ws As Worksheet, _
                               ByVal outputRow As Long)
    With ws
        .Range("A4:A" & outputRow - 1).NumberFormat = "dd/mm/yyyy"
        .Range("C4:I" & outputRow - 1).NumberFormat = "0.000"
        .Columns("A:J").AutoFit

        .Cells(outputRow + 1, 1).Value = "SO NGAY/TO DA LUU"
        .Cells(outputRow + 1, 2).Value = outputRow - 4
        .Cells(outputRow + 2, 1).Value = "TONG Qdc (MWh)"
        .Cells(outputRow + 2, 2).Formula = "=SUM(F4:F" & CStr(outputRow - 1) & ")"
        .Cells(outputRow + 3, 1).Value = "TONG Qmp (MWh)"
        .Cells(outputRow + 3, 2).Formula = "=SUM(G4:G" & CStr(outputRow - 1) & ")"
        .Cells(outputRow + 4, 1).Value = "TONG Qdd_V (MWh)"
        .Cells(outputRow + 4, 2).Formula = "=SUM(H4:H" & CStr(outputRow - 1) & ")"
        .Cells(outputRow + 5, 1).Value = "TONG Qdu (MWh)"
        .Cells(outputRow + 5, 2).Formula = "=SUM(I4:I" & CStr(outputRow - 1) & ")"
        .Range("A" & CStr(outputRow + 1) & ":B" & CStr(outputRow + 5)).Font.Bold = True
        .Range("B" & CStr(outputRow + 2) & ":B" & CStr(outputRow + 5)).NumberFormat = "0.000"
    End With
End Sub

Private Function SnapshotSheetName(ByVal targetDate As Date, _
                                   ByVal unitName As String) As String
    SnapshotSheetName = M_PREFIX & Format$(targetDate, "yyyymm_dd") & "_" & unitName
End Function

Private Sub CreateCombinedDailySheet(ByVal outWb As Workbook, _
                                     ByVal targetDate As Date, _
                                     ByVal usedNames As Collection)
    Dim s1Ws As Worksheet
    Dim s2Ws As Worksheet
    Dim templateWs As Worksheet
    Dim dailyWs As Worksheet
    Dim oldVisibility As Long

    If MonthSheetExists(SnapshotSheetName(targetDate, "S1")) Then
        Set s1Ws = ThisWorkbook.Worksheets(SnapshotSheetName(targetDate, "S1"))
    End If
    If MonthSheetExists(SnapshotSheetName(targetDate, "S2")) Then
        Set s2Ws = ThisWorkbook.Worksheets(SnapshotSheetName(targetDate, "S2"))
    End If
    If s1Ws Is Nothing And s2Ws Is Nothing Then Exit Sub

    If Not s1Ws Is Nothing Then
        Set templateWs = s1Ws
    Else
        Set templateWs = s2Ws
    End If

    ' Tao sheet moi hoan toan. Khong copy worksheet, khong PasteSpecial,
    ' khong merge bat ky o nao trong duong xuat bao cao.
    Set dailyWs = outWb.Worksheets.Add( _
        After:=outWb.Worksheets(outWb.Worksheets.Count))
    dailyWs.Name = MakeDailySheetName(targetDate, usedNames)

    BuildCombinedDailyLayout dailyWs, targetDate

    oldVisibility = templateWs.Visible
    templateWs.Visible = xlSheetVisible
    dailyWs.Range("A4:A" & M_REPORT_TOTAL_ROW).Value2 = _
        templateWs.Range("A4:A" & M_REPORT_TOTAL_ROW).Value2
    dailyWs.Range("B3:K3").Value2 = templateWs.Range("B3:K3").Value2
    dailyWs.Range("L3:U3").Value2 = templateWs.Range("B3:K3").Value2
    CopyTemplateColumnWidths templateWs, dailyWs
    templateWs.Visible = oldVisibility

    If Not s1Ws Is Nothing Then
        oldVisibility = s1Ws.Visible
        s1Ws.Visible = xlSheetVisible
        dailyWs.Range("B4:K" & M_REPORT_TOTAL_ROW).Value2 = _
            s1Ws.Range("B4:K" & M_REPORT_TOTAL_ROW).Value2
        s1Ws.Visible = oldVisibility
    Else
        dailyWs.Range("B4:K" & M_REPORT_TOTAL_ROW).ClearContents
    End If

    If Not s2Ws Is Nothing Then
        oldVisibility = s2Ws.Visible
        s2Ws.Visible = xlSheetVisible
        dailyWs.Range("L4:U" & M_REPORT_TOTAL_ROW).Value2 = _
            s2Ws.Range("B4:K" & M_REPORT_TOTAL_ROW).Value2
        s2Ws.Visible = oldVisibility
    Else
        dailyWs.Range("L4:U" & M_REPORT_TOTAL_ROW).ClearContents
    End If

    NormalizeCombinedQduSignBlock dailyWs, 2, 11
    NormalizeCombinedQduSignBlock dailyWs, 12, 21
    ApplyCombinedDailyFormatting dailyWs, targetDate
End Sub





Private Sub BuildCombinedDailyLayout(ByVal ws As Worksheet, _
                                     ByVal targetDate As Date)
    ws.Cells.Clear

    ws.Range("A1:U1").UnMerge
    ws.Range("A2:U3").UnMerge

    ws.Range("A1").Value = "BAO CAO Qdd / Qdu NGAY " & DateTextDMY(targetDate)
    ws.Range("A1:U1").HorizontalAlignment = xlCenterAcrossSelection
    ws.Range("A1:U1").VerticalAlignment = xlCenter

    ws.Range("A2").Value = "Chu ky"
    ws.Range("B2").Value = "S1DH1"
    ws.Range("L2").Value = "S2DH1"

    ws.Range("B2:K2").HorizontalAlignment = xlCenterAcrossSelection
    ws.Range("L2:U2").HorizontalAlignment = xlCenterAcrossSelection
    ws.Range("A2:A3").HorizontalAlignment = xlCenter
    ws.Range("A2:A3").VerticalAlignment = xlCenter
End Sub

Private Sub CopyTemplateColumnWidths(ByVal sourceWs As Worksheet, _
                                     ByVal targetWs As Worksheet)
    Dim offsetValue As Long

    targetWs.Columns(1).ColumnWidth = sourceWs.Columns(1).ColumnWidth
    For offsetValue = 0 To 9
        targetWs.Columns(2 + offsetValue).ColumnWidth = _
            sourceWs.Columns(2 + offsetValue).ColumnWidth
        targetWs.Columns(12 + offsetValue).ColumnWidth = _
            sourceWs.Columns(2 + offsetValue).ColumnWidth
    Next offsetValue
End Sub

Private Sub ApplyCombinedDailyFormatting(ByVal ws As Worksheet, _
                                         ByVal targetDate As Date)
    Dim r As Long

    With ws.Range("A1:U1")
        .Interior.Color = RGB(31, 78, 121)
        .Font.Color = RGB(255, 255, 255)
        .Font.Bold = True
        .Font.Size = 15
        .HorizontalAlignment = xlCenterAcrossSelection
        .VerticalAlignment = xlCenter
    End With
    ws.Rows(1).RowHeight = 28

    With ws.Range("A2:U3")
        .Font.Bold = True
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = True
    End With

    With ws.Range("A2:A3")
        .Interior.Color = RGB(37, 86, 129)
        .Font.Color = RGB(255, 255, 255)
    End With
    With ws.Range("B2:K2")
        .Interior.Color = RGB(37, 86, 129)
        .Font.Color = RGB(255, 255, 255)
        .HorizontalAlignment = xlCenterAcrossSelection
    End With
    With ws.Range("L2:U2")
        .Interior.Color = RGB(37, 86, 129)
        .Font.Color = RGB(255, 255, 255)
        .HorizontalAlignment = xlCenterAcrossSelection
    End With
    With ws.Range("B3:K3")
        .Interior.Color = RGB(37, 86, 129)
        .Font.Color = RGB(255, 255, 255)
    End With
    With ws.Range("L3:U3")
        .Interior.Color = RGB(37, 86, 129)
        .Font.Color = RGB(255, 255, 255)
    End With

    ws.Range("A4:A" & M_REPORT_TOTAL_ROW).HorizontalAlignment = xlCenter
    ws.Range("B4:U" & M_REPORT_TOTAL_ROW).NumberFormat = "0.000"
    ws.Range("A" & M_REPORT_TOTAL_ROW).Value = "Tong ngay"
    ws.Range("A" & M_REPORT_TOTAL_ROW & ":U" & M_REPORT_TOTAL_ROW).Font.Bold = True

    ws.Columns("A").ColumnWidth = 22
    ws.Columns("K").ColumnWidth = 18
    ws.Columns("U").ColumnWidth = 18

    With ws.Range("A2:U" & M_REPORT_TOTAL_ROW).Borders
        .LineStyle = xlContinuous
        .Weight = xlThin
        .Color = RGB(191, 191, 191)
    End With

    With ws.PageSetup
        .Orientation = xlLandscape
        .Zoom = False
        .FitToPagesWide = 1
        .FitToPagesTall = False
        .PrintArea = "$A$1:$U$" & CStr(M_REPORT_TOTAL_ROW)
        .CenterHorizontally = True
    End With
End Sub

Private Function MonthPickSavePath(ByVal defaultName As String, _
                                   ByRef usedAutomaticPath As Boolean) As Variant
    Dim chosenPath As Variant
    Dim baseFolder As String
    Dim automaticPath As String
    Dim pickerError As Long
    Dim separatorText As String

    usedAutomaticPath = False
    separatorText = Application.PathSeparator

    ' Excel Mac co the bao loi 1004 neu truyen FileFilter hoac Title.
#If Mac Then
    On Error Resume Next
    Err.Clear
    chosenPath = Application.GetSaveAsFilename( _
        InitialFileName:=defaultName)
    pickerError = Err.Number
    Err.Clear
    On Error GoTo 0
#Else
    On Error Resume Next
    Err.Clear
    chosenPath = Application.GetSaveAsFilename( _
        InitialFileName:=defaultName, _
        FileFilter:="Excel Workbook (*.xlsx), *.xlsx", _
        Title:="Xuat bao cao Qdd/Qdu thang")
    pickerError = Err.Number
    Err.Clear
    On Error GoTo 0
#End If

    If pickerError = 0 Then
        MonthPickSavePath = chosenPath
        Exit Function
    End If

    ' Lan hai: goi toi gian, khong FileFilter/Title.
    On Error Resume Next
    Err.Clear
    chosenPath = Application.GetSaveAsFilename(defaultName)
    pickerError = Err.Number
    Err.Clear
    On Error GoTo 0

    If pickerError = 0 Then
        MonthPickSavePath = chosenPath
        Exit Function
    End If

    ' Neu hop chon Excel van loi, tu dong luu canh workbook.
    baseFolder = Trim$(ThisWorkbook.Path)
    If baseFolder = "" Then baseFolder = Trim$(CurDir$)

    If baseFolder = "" Then
        Err.Raise vbObjectError + 2131, "MonthPickSavePath", _
            "Excel khong mo duoc hop chon noi luu va khong xac dinh duoc thu muc workbook."
    End If

    automaticPath = baseFolder & separatorText & defaultName

    If Len(Dir$(automaticPath)) > 0 Then
        automaticPath = baseFolder & separatorText & _
                        Left$(defaultName, Len(defaultName) - 5) & "_" & _
                        Format$(Now, "yyyymmdd_hhnnss") & ".xlsx"
    End If

    usedAutomaticPath = True
    MonthPickSavePath = automaticPath
End Function

Private Sub NormalizeCombinedQduSignBlock(ByVal ws As Worksheet, _
                                          ByVal firstColumn As Long, _
                                          ByVal lastColumn As Long)
    Dim pQdcColumn As Long
    Dim signColumn As Long
    Dim qduColumn As Long
    Dim c As Long
    Dim r As Long
    Dim headerText As String
    Dim qduValue As Variant
    Dim numericQdu As Double

    For c = firstColumn To lastColumn
        headerText = UCase$(Trim$(CStr(ws.Cells(3, c).Value2)))
        headerText = Replace(headerText, " ", "")

        If InStr(1, headerText, "P_QDC", vbTextCompare) > 0 Or _
           InStr(1, headerText, "PQDC", vbTextCompare) > 0 Then
            pQdcColumn = c
            Exit For
        End If
    Next c

    If pQdcColumn < firstColumn + 2 Then Exit Sub

    signColumn = pQdcColumn - 1
    qduColumn = pQdcColumn - 2

    For r = M_REPORT_FIRST_ROW To M_REPORT_LAST_ROW
        qduValue = ws.Cells(r, qduColumn).Value2

        If IsError(qduValue) Or Len(Trim$(CStr(qduValue))) = 0 Then
            ws.Cells(r, signColumn).ClearContents
        ElseIf IsNumeric(qduValue) Then
            numericQdu = CDbl(qduValue)

            If Abs(numericQdu) < 0.0000000001 Then
                ws.Cells(r, signColumn).ClearContents
            ElseIf numericQdu > 0 Then
                ws.Cells(r, signColumn).Value = MonthPositiveLabel()
            Else
                ws.Cells(r, signColumn).Value = MonthNegativeLabel()
            End If
        Else
            ws.Cells(r, signColumn).ClearContents
        End If
    Next r
End Sub

Private Function MonthPositiveLabel() As String
    MonthPositiveLabel = "D" & ChrW(&H1B0) & ChrW(&H1A1) & "ng"
End Function

Private Function MonthNegativeLabel() As String
    MonthNegativeLabel = ChrW(&HC2) & "m"
End Function



Private Function MakeDailySheetName(ByVal targetDate As Date, _
                                    ByVal usedNames As Collection) As String
    Dim baseName As String
    Dim candidate As String
    Dim counter As Long

    baseName = Right$("0" & CStr(Day(targetDate)), 2) & "." & _
               Right$("0" & CStr(Month(targetDate)), 2) & "." & _
               Right$(CStr(Year(targetDate)), 2)
    candidate = baseName
    counter = 2

    Do While CollectionHasText(usedNames, candidate)
        candidate = Left$(baseName, 27) & "_" & CStr(counter)
        counter = counter + 1
    Loop

    usedNames.Add candidate
    MakeDailySheetName = candidate
End Function

Private Function MonthKeyFromDate(ByVal valueDate As Date) As String
    MonthKeyFromDate = CStr(Year(valueDate)) & Right$("0" & CStr(Month(valueDate)), 2)
End Function

Private Function DateKeyFromDate(ByVal valueDate As Date) As String
    DateKeyFromDate = CStr(Year(valueDate)) & _
                      Right$("0" & CStr(Month(valueDate)), 2) & _
                      Right$("0" & CStr(Day(valueDate)), 2)
End Function

Private Function DateTextDMY(ByVal valueDate As Date) As String
    DateTextDMY = Right$("0" & CStr(Day(valueDate)), 2) & "/" & _
                  Right$("0" & CStr(Month(valueDate)), 2) & "/" & CStr(Year(valueDate))
End Function

Private Function IsValidMonthKey(ByVal keyText As String) As Boolean
    Dim yearPart As Long
    Dim monthPart As Long
    Dim i As Long
    Dim ch As String

    keyText = Trim$(keyText)
    If Len(keyText) <> 6 Then Exit Function

    For i = 1 To 6
        ch = Mid$(keyText, i, 1)
        If ch < "0" Or ch > "9" Then Exit Function
    Next i

    yearPart = CLng(Left$(keyText, 4))
    monthPart = CLng(Right$(keyText, 2))
    IsValidMonthKey = yearPart >= 2000 And yearPart <= 9999 And _
                      monthPart >= 1 And monthPart <= 12
End Function

Private Function GetStoredMonthKey(ByVal historyWs As Worksheet) As String
    Dim keyText As String
    Dim legacyText As String
    Dim firstDate As Date
    Dim detailText As String

    keyText = Trim$(CStr(historyWs.Range(M_MONTH_KEY_CELL).Value2))
    If IsValidMonthKey(keyText) Then
        GetStoredMonthKey = keyText
        Exit Function
    End If

    legacyText = Trim$(CStr(historyWs.Range("B2").Value2))
    If InStr(legacyText, ".") > 0 Then
        legacyText = Left$(legacyText, InStr(legacyText, ".") - 1)
    End If
    If IsValidMonthKey(legacyText) Then
        SetStoredMonthKey historyWs, legacyText
        GetStoredMonthKey = legacyText
        Exit Function
    End If

    If HistoryHasData(historyWs) Then
        If TryReadExcelDate(historyWs.Cells(M_FIRST_ROW, 2).Value2, firstDate, detailText) Then
            keyText = MonthKeyFromDate(firstDate)
            SetStoredMonthKey historyWs, keyText
            GetStoredMonthKey = keyText
        End If
    End If
End Function

Private Sub SetStoredMonthKey(ByVal historyWs As Worksheet, ByVal keyText As String)
    historyWs.Range(M_MONTH_KEY_CELL).NumberFormat = "@"
    historyWs.Range(M_MONTH_KEY_CELL).Value2 = keyText
    historyWs.Range("B2").NumberFormat = "@"
    historyWs.Range("B2").Value2 = Right$(keyText, 2) & "/" & Left$(keyText, 4)
    historyWs.Range("C2").ClearContents
    historyWs.Columns("N").Hidden = True
End Sub

Private Function TryReadExcelDate(ByVal valueIn As Variant, _
                                  ByRef resultDate As Date, _
                                  ByRef detailText As String) As Boolean
    Dim serialValue As Double
    Dim dateValue As Date
    Dim textValue As String
    Dim normalized As String
    Dim parts As Variant
    Dim dayPart As Long
    Dim monthPart As Long
    Dim yearPart As Long

    On Error GoTo EH
    detailText = ""

    If IsError(valueIn) Or IsEmpty(valueIn) Then
        detailText = "Gia tri ngay bi trong hoac loi."
        Exit Function
    End If

    If IsNumeric(valueIn) Then
        serialValue = CDbl(valueIn)
        If serialValue <= 0 Then
            detailText = "So serial ngay khong hop le."
            Exit Function
        End If
        dateValue = CDate(serialValue)
        resultDate = DateSerial(Year(dateValue), Month(dateValue), Day(dateValue))
        TryReadExcelDate = True
        Exit Function
    End If

    textValue = Trim$(CStr(valueIn))
    normalized = Replace(Replace(textValue, ".", "-"), "/", "-")
    parts = Split(normalized, "-")
    If UBound(parts) <> 2 Then
        detailText = "Can dinh dang dd/mm/yyyy hoac yyyy-mm-dd."
        Exit Function
    End If

    If Len(Trim$(CStr(parts(0)))) = 4 Then
        yearPart = CLng(parts(0))
        monthPart = CLng(parts(1))
        dayPart = CLng(parts(2))
    Else
        dayPart = CLng(parts(0))
        monthPart = CLng(parts(1))
        yearPart = CLng(parts(2))
        If yearPart < 100 Then yearPart = 2000 + yearPart
    End If

    dateValue = DateSerial(yearPart, monthPart, dayPart)
    If Year(dateValue) <> yearPart Or Month(dateValue) <> monthPart Or Day(dateValue) <> dayPart Then
        detailText = "Ngay khong ton tai."
        Exit Function
    End If

    resultDate = dateValue
    TryReadExcelDate = True
    Exit Function

EH:
    detailText = "Khong doc duoc ngay: " & Err.Description
End Function

Private Function CollectionHasText(ByVal items As Collection, _
                                   ByVal valueText As String) As Boolean
    Dim item As Variant

    For Each item In items
        If UCase$(CStr(item)) = UCase$(valueText) Then
            CollectionHasText = True
            Exit Function
        End If
    Next item
End Function

Private Function NormalizeMonthUnit(ByVal unitText As String) As String
    Dim s As String

    s = UCase$(Trim$(unitText))
    If Left$(s, 2) = "S1" Then
        NormalizeMonthUnit = "S1"
    ElseIf Left$(s, 2) = "S2" Then
        NormalizeMonthUnit = "S2"
    End If
End Function

Private Function MonthSheetExists(ByVal sheetName As String) As Boolean
    Dim ws As Worksheet

    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(sheetName)
    MonthSheetExists = Not ws Is Nothing
    On Error GoTo 0
End Function

Private Sub MonthWriteLog(ByVal actionText As String, _
                          ByVal resultText As String, _
                          ByVal detailText As String)
    Dim ws As Worksheet
    Dim nextRow As Long

    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(M_SH_LOG)
    If ws Is Nothing Then Exit Sub

    nextRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row + 1
    If nextRow < 2 Then nextRow = 2

    ws.Cells(nextRow, 1).Value = Now
    ws.Cells(nextRow, 2).Value = "BAO_CAO_THANG_" & actionText
    ws.Cells(nextRow, 3).Value = resultText
    ws.Cells(nextRow, 4).Value = detailText
    ws.Cells(nextRow, 1).NumberFormat = "dd/mm/yyyy hh:mm:ss"
    On Error GoTo 0
End Sub
