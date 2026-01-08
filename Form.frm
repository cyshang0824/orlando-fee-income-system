VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Form 
   Caption         =   "住戶繳費系統"
   ClientHeight    =   8628.001
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   8676.001
   OleObjectBlob   =   "Form.frx":0000
   StartUpPosition =   2  '螢幕中央
End
Attribute VB_Name = "Form"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Latest_ExternalPath As String
Public wbExternal As Workbook
Public ws住戶 As Worksheet
Public 住戶總表_X As Long
Public 住戶總表_繳到幾月Y As Long
Public Current_Time, 棟樓別, 收據繳交月份



'------------------------------
' UserForm 初始化
'------------------------------
Private Sub UserForm_Initialize()
Dim ctrl As MSForms.Control
Dim i As Long
Dim sColumn As Long
Dim wbExternal_Init As Workbook
Dim ws住戶_Init As Worksheet
Dim isNewOpen As Boolean
Dim folderPath As String
Dim searchPattern As String
Dim fileName As String
Dim latestFileName As String
Dim latestTimestamp As Double
Dim currentTimestamp As Double
Dim timestampString As String
Dim dateTimeString As String
folderPath = GetFolderPath()
searchPattern = "住戶總表_????????_??????*.xlsx"    ' ★修正pattern
latestTimestamp = 0
latestFileName = ""
With Me
.BackColor = RGB(220, 230, 240)
.Caption = "住戶繳費系統"
End With
For Each ctrl In Me.Controls
If TypeOf ctrl Is MSForms.Label Or _
TypeOf ctrl Is MSForms.TextBox Or _
TypeOf ctrl Is MSForms.CommandButton Then
With ctrl
If TypeOf ctrl Is MSForms.Label Then
.BackColor = Me.BackColor
.BackStyle = fmBackStyleTransparent
End If
If TypeOf ctrl Is MSForms.TextBox Then
If .Name <> "棟樓別_2" And .Name <> "想繳幾月" Then
.Locked = True
End If
End If
End With
End If
Next ctrl
On Error GoTo ErrorHandler
isNewOpen = False
fileName = Dir(folderPath & searchPattern)
Do While fileName <> ""
timestampString = Mid(fileName, 6, 15)
dateTimeString = Left(timestampString, 4) & "/" & _
Mid(timestampString, 5, 2) & "/" & _
Mid(timestampString, 7, 2) & " " & _
Mid(timestampString, 10, 2) & ":" & _
Mid(timestampString, 12, 2) & ":" & _
Right(timestampString, 2)
On Error Resume Next
currentTimestamp = CDbl(CDate(dateTimeString))
On Error GoTo ErrorHandler
If currentTimestamp > latestTimestamp Then
latestTimestamp = currentTimestamp
latestFileName = fileName
End If
fileName = Dir
Loop
If latestFileName <> "" Then
Latest_ExternalPath = folderPath & latestFileName
Me.住戶總表Name_2.Caption = latestFileName
On Error Resume Next
Set wbExternal_Init = Workbooks(latestFileName)
On Error GoTo ErrorHandler
If wbExternal_Init Is Nothing Then
Set wbExternal_Init = Workbooks.Open(Latest_ExternalPath)
isNewOpen = True
End If
Else
Me.住戶總表Name_2.Caption = "找不到檔案"
GoTo NoFileError
End If
Set ws住戶_Init = wbExternal_Init.Sheets("住戶總表")
sColumn = ws住戶_Init.Columns("S").Column
Me.收款人.Clear
i = 2
Do While ws住戶_Init.Cells(i, sColumn).Value <> "" And i <= Rows.count
Me.收款人.AddItem ws住戶_Init.Cells(i, sColumn).Value
i = i + 1
Loop
ExitHandler:
If Not wbExternal_Init Is Nothing Then
wbExternal_Init.Close SaveChanges:=False
End If
Set ws住戶_Init = Nothing
Set wbExternal_Init = Nothing
Exit Sub
ErrorHandler:
Me.住戶總表Name_2.Caption = "載入錯誤"
MsgBox "載入收款人資料時發生錯誤，請檢查檔案和程式碼。", vbCritical, "程式碼執行錯誤"
Resume ExitHandler
NoFileError:
MsgBox "無法找到任何最新的備份檔案，請確認路徑或備份檔案是否存在。", vbCritical, "找不到檔案錯誤"
Exit Sub
End Sub

'------------------------------
' Search_Click
'------------------------------
Sub Search_Click()
If Trim(Me.想繳幾月.Text) = "" Or Val(Me.想繳幾月.Text) = 0 Then
Me.想繳幾月.Text = "1"
End If
Dim wbMain As Workbook
Dim 棟樓別_Search As String
Dim Found_棟樓別 As Boolean
Dim X As Long: X = 2
Dim Y As Long, Max_Y As Long
Dim 本次繳交月份 As Variant
Dim 舊車費要幾個月 As Integer, 新車費要幾個月 As Integer
Dim folderPath As String
Dim searchPattern As String
Dim fileName As String
Dim latestFileName As String
Dim latestTimestamp As Double
Dim currentTimestamp As Double
Dim timestampString As String
Dim dateTimeString As String
folderPath = GetFolderPath()
searchPattern = "住戶總表_????????_??????*.xlsx"    ' ★修正pattern
latestTimestamp = 0
latestFileName = ""
Set wbExternal = Nothing
On Error GoTo ErrorHandler_Search
If Not wbExternal Is Nothing Then
wbExternal.Close SaveChanges:=False
Set wbExternal = Nothing
Set ws住戶 = Nothing
End If
fileName = Dir(folderPath & searchPattern)
Do While fileName <> ""
timestampString = Mid(fileName, 6, 15)
dateTimeString = Left(timestampString, 4) & "/" & _
Mid(timestampString, 5, 2) & "/" & _
Mid(timestampString, 7, 2) & " " & _
Mid(timestampString, 10, 2) & ":" & _
Mid(timestampString, 12, 2) & ":" & _
Right(timestampString, 2)
On Error Resume Next
currentTimestamp = CDbl(CDate(dateTimeString))
On Error GoTo ErrorHandler_Search
If currentTimestamp > latestTimestamp Then
latestTimestamp = currentTimestamp
latestFileName = fileName
End If
fileName = Dir
Loop
If latestFileName = "" Then
Me.住戶總表Name_2.Caption = "找不到檔案"
MsgBox "無法找到任何最新的備份檔案，無法執行查詢。", vbCritical, "錯誤"
Exit Sub
End If
Latest_ExternalPath = folderPath & latestFileName
Me.住戶總表Name_2.Caption = latestFileName
On Error Resume Next
Set wbExternal = Workbooks(latestFileName)
On Error GoTo ErrorHandler_Search
If wbExternal Is Nothing Then
Set wbExternal = Workbooks.Open(Latest_ExternalPath)
End If
Set ws住戶 = wbExternal.Sheets("住戶總表")
' ======= 新增段落：預產生收據編號並填入收據範本 =======
Dim ws紀錄_External As Worksheet
Dim Max As Long, PreviousNumber As Long
Dim NextReceiptNo As String
Set ws紀錄_External = wbExternal.Sheets("歐藍朵大廈管理費繳費紀錄")
Max = 2
Do While ws紀錄_External.Cells(Max, 10) <> "" Or ws紀錄_External.Cells(Max, 1) <> "" Or ws紀錄_External.Cells(Max, 2) <> ""
Max = Max + 1
Loop
If Max > 2 And Left(ws紀錄_External.Cells(Max - 1, 10), 2) = "PC" Then
PreviousNumber = Val(Mid(ws紀錄_External.Cells(Max - 1, 10), 3)) + 1
NextReceiptNo = "PC" & Format(PreviousNumber, "0000")
Else
NextReceiptNo = "PC0001"
End If
ThisWorkbook.Sheets("收據範本").Cells(2, 5) = NextReceiptNo
ThisWorkbook.Sheets("收據範本").Cells(18, 5) = NextReceiptNo
' ========== 新增段落結束 ==========
Set wbMain = ThisWorkbook
wbMain.Activate
wbMain.Sheets("收據範本").Select
wbMain.Sheets("收據範本").Cells(2, 2).Select
棟樓別_Search = Me.棟樓別_2.Text
棟樓別 = 棟樓別_Search
Select Case "-" & Mid(棟樓別, InStr(棟樓別, "-") + 1)
Case "-2", "-4": 棟樓別 = "C-" & 棟樓別
Case "-1", "-3": 棟樓別 = "D-" & 棟樓別
Case "-5", "-7", "-9": 棟樓別 = "A-" & 棟樓別
Case "-6", "-8", "-10": 棟樓別 = "B-" & 棟樓別
End Select
Found_棟樓別 = False
X = 2
Do
If ws住戶.Cells(X, 3) = 棟樓別 Then
住戶總表_X = X
Y = 20
Do
Y = Y + 1
Loop Until ws住戶.Cells(X, Y) = ""
Max_Y = Y - 1
繳到幾月_2.Caption = ws住戶.Cells(1, Max_Y)
住戶總表_繳到幾月Y = Max_Y
If Val(想繳幾月.Text) > 12 Then MsgBox "不能繳超過一年"
If Val(想繳幾月.Text) = 1 Then
本次繳交月份 = Val(Left(繳到幾月_2.Caption, Len(繳到幾月_2.Caption) - 1)) + Val(想繳幾月.Text)
If 本次繳交月份 > 12 Then
本次繳交月份 = Year(Now) - 1911 + 1 & "/" & 本次繳交月份 - 12
收據繳交月份 = 本次繳交月份 & "月"
Else
收據繳交月份 = Year(Now) - 1911 & "/" & 本次繳交月份 & "月"
End If
Else
本次繳交月份 = Val(Left(繳到幾月_2.Caption, Len(繳到幾月_2.Caption) - 1)) + Val(想繳幾月.Text)
If 本次繳交月份 > 12 Then
本次繳交月份 = Year(Now) - 1911 + 1 & "/" & 本次繳交月份 - 12
End If
If CInt(Replace(繳到幾月_2.Caption, "月", "")) + 1 <= 12 Then
收據繳交月份 = Year(Now) - 1911 & "/" & CInt(Replace(繳到幾月_2.Caption, "月", "")) + 1 & "-" & 本次繳交月份 & "月"
Else
收據繳交月份 = Year(Now) + 1 - 1911 & "/" & CInt(Replace(繳到幾月_2.Caption, "月", "")) + 1 - 12 & "-" & 本次繳交月份 & "月"
End If
End If
收據繳交月份_2.Text = 收據繳交月份
舊車費要幾個月 = 0
新車費要幾個月 = 0
If Month(Now) >= 7 And Val(Left(繳到幾月_2.Caption, Len(繳到幾月_2.Caption) - 1)) <= 6 Then
舊車費要幾個月 = 6 - Val(Left(繳到幾月_2.Caption, Len(繳到幾月_2.Caption) - 1))
End If
If 舊車費要幾個月 < Val(想繳幾月.Text) Then
新車費要幾個月 = Val(想繳幾月.Text) - 舊車費要幾個月
Else
舊車費要幾個月 = Val(想繳幾月.Text)
End If
所有權人_2.Caption = ws住戶.Cells(X, 4)
管理費_2.Text = ws住戶.Cells(X, 6) * Val(想繳幾月.Text)
汽車車位_2.Caption = ws住戶.Cells(X, 7)
汽車清潔費_2.Text = ws住戶.Cells(X, 8) * 新車費要幾個月 + ws住戶.Cells(X, 17) * 舊車費要幾個月
機車車位_2.Caption = ws住戶.Cells(X, 9)
機車清潔費_2.Text = ws住戶.Cells(X, 10) * 新車費要幾個月 + ws住戶.Cells(X, 18) * 舊車費要幾個月
小計_2.Text = Val(管理費_2.Text) + Val(汽車清潔費_2.Text) + Val(機車清潔費_2.Text)
應繳金額_2.Caption = 小計_2.Text
If ws住戶.Cells(X, 13) <> "" Then
區權會抵扣_2.Text = 0
Else
區權會抵扣_2.Text = ws住戶.Cells(X, 12)
End If
If ws住戶.Cells(X, 15) <> "" Then
住戶回饋_2.Text = 0
Else
住戶回饋_2.Text = ws住戶.Cells(X, 13)
End If
應繳金額_2.Caption = Val(小計_2.Text) - Val(區權會抵扣_2.Text) - Val(住戶回饋_2.Text)
If 汽車車位_2.Caption = "" Then 汽車車位_2.Caption = "無"
If 汽車清潔費_2.Text = "" Then 汽車清潔費_2.Text = "0"
If 機車車位_2.Caption = "" Then 機車車位_2.Caption = "無"
If 機車清潔費_2.Text = "" Then 機車清潔費_2.Text = "0"
If 區權會抵扣_2.Text = "" Then 區權會抵扣_2.Text = "0"
If 住戶回饋_2.Text = "" Then 住戶回饋_2.Text = "0"
Dim CurrentYear As String, currentMonth As String, currentDay As String
Dim currentHour As String, currentMinute As String
CurrentYear = Format(Year(Now), "000")
currentMonth = Format(Month(Now), "00")
currentDay = Format(Day(Now), "00")
currentHour = Format(Hour(Now), "00")
currentMinute = Format(Minute(Now), "00")
Current_Time = "'" & CurrentYear & currentMonth & currentDay & currentHour & currentMinute
Call 填入收據範本
Found_棟樓別 = True
Exit Do
End If
X = X + 1
Loop Until ws住戶.Cells(X, 3) = ""
If Found_棟樓別 = False Then
MsgBox "沒有此住戶"
所有權人_2.Caption = ""
管理費_2.Text = ""
汽車車位_2.Caption = ""
汽車清潔費_2.Text = ""
機車車位_2.Caption = ""
機車清潔費_2.Text = ""
小計_2.Text = ""
應繳金額_2.Caption = ""
繳到幾月_2.Caption = ""
區權會抵扣_2.Text = ""
住戶回饋_2.Text = ""
End If
ExitHandler_Search:
If Not wbExternal Is Nothing Then
wbExternal.Close SaveChanges:=False
Set wbExternal = Nothing
Set ws住戶 = Nothing
End If
Exit Sub
ErrorHandler_Search:
Me.住戶總表Name_2.Caption = "查詢錯誤"
MsgBox "查詢時發生錯誤：" & Err.Description, vbCritical, "查詢失敗"
If Not wbExternal Is Nothing Then
wbExternal.Close SaveChanges:=False
End If
Set ws住戶 = Nothing
Set wbExternal = Nothing
Resume ExitHandler_Search
End Sub





Sub 儲存與印收據_Click()
    Dim wbExternal As Workbook
    Dim ws紀錄_External As Worksheet
    Dim ws住戶 As Worksheet
    Dim 區權會抵扣_num As Double
    Dim 住戶回饋_num As Double
    Dim todayFormatted As String
    Dim i As Long
    Dim PreviousNumber As Long
    Dim timestamp As String
    Dim backupPath As String
    Dim Max As Long
    Dim X As Long
    Const RecordSheetName As String = "歐藍朵大廈管理費繳費紀錄"
    Dim folderPath As String
    Dim searchPattern As String
    Dim fileName As String
    Dim latestFileName As String
    Dim latestTimestamp As Double
    Dim currentTimestamp As Double
    Dim timestampString As String
    Dim dateTimeString As String

    On Error GoTo ErrorHandler_Save

    ' 1. 自動抓最新檔名
    folderPath = GetFolderPath()
    searchPattern = "住戶總表_????????_??????*.xlsx"
    latestTimestamp = 0
    latestFileName = ""
    fileName = Dir(folderPath & searchPattern)
    Do While fileName <> ""
        timestampString = Mid(fileName, 6, 15)
        dateTimeString = Left(timestampString, 4) & "/" & _
                         Mid(timestampString, 5, 2) & "/" & _
                         Mid(timestampString, 7, 2) & " " & _
                         Mid(timestampString, 10, 2) & ":" & _
                         Mid(timestampString, 12, 2) & ":" & _
                         Right(timestampString, 2)
        On Error Resume Next
        currentTimestamp = CDbl(CDate(dateTimeString))
        On Error GoTo ErrorHandler_Save
        If currentTimestamp > latestTimestamp Then
            latestTimestamp = currentTimestamp
            latestFileName = fileName
        End If
        fileName = Dir
    Loop
    If latestFileName = "" Then
        Me.住戶總表Name_2.Caption = "找不到檔案"
        MsgBox "無法找到任何最新的備份檔案，無法執行儲存。", vbCritical, "錯誤"
        Exit Sub
    End If
    Latest_ExternalPath = folderPath & latestFileName
    Me.住戶總表Name_2.Caption = latestFileName

    Set wbExternal = Workbooks.Open(Latest_ExternalPath)
    Set ws紀錄_External = wbExternal.Sheets(RecordSheetName)
    Set ws住戶 = wbExternal.Sheets("住戶總表")

    ' 2. 再度去更新 NextReceiptNo 並填入收據範本
    Max = 2
    Do While ws紀錄_External.Cells(Max, 10) <> "" Or ws紀錄_External.Cells(Max, 1) <> "" Or ws紀錄_External.Cells(Max, 2) <> ""
        Max = Max + 1
    Loop
    Dim NextReceiptNo As String
    Dim receiptID As String
    If Max > 2 And Left(ws紀錄_External.Cells(Max - 1, 10), 2) = "PC" Then
        PreviousNumber = Val(Mid(ws紀錄_External.Cells(Max - 1, 10), 3)) + 1
        NextReceiptNo = "PC" & Format(PreviousNumber, "0000")
    Else
        NextReceiptNo = "PC0001"
    End If
    ThisWorkbook.Sheets("收據範本").Cells(2, 5) = NextReceiptNo
    ThisWorkbook.Sheets("收據範本").Cells(18, 5) = NextReceiptNo

    ' 3. 儲存收據資料、編號
    X = 1
    Do
        X = X + 1
        DoEvents
    Loop Until ws紀錄_External.Cells(X, 10) = "" And ws紀錄_External.Cells(X, 1) = "" And ws紀錄_External.Cells(X, 2) = ""
    Max = X

    ws紀錄_External.Select
    ws紀錄_External.Cells(Max, 10).Select
    ws紀錄_External.Cells(Max, 1) = Month(Now) & "/" & Day(Now)
    ws紀錄_External.Cells(Max, 2) = 棟樓別
    ws紀錄_External.Cells(Max, 3) = 所有權人_2.Caption
    ws紀錄_External.Cells(Max, 4) = 管理費_2.Text
    ws紀錄_External.Cells(Max, 5) = 汽車清潔費_2.Text
    ws紀錄_External.Cells(Max, 6) = 機車清潔費_2.Text
    ws紀錄_External.Cells(Max, 7) = 小計_2.Text

    If 區權會抵扣_2.Text = "無" Or 區權會抵扣_2.Text = "" Then
        區權會抵扣_num = 0
    Else
        區權會抵扣_num = Val(區權會抵扣_2.Text)
    End If

    If 住戶回饋_2.Text = "無" Or 住戶回饋_2.Text = "" Then
        住戶回饋_num = 0
    Else
        住戶回饋_num = Val(住戶回饋_2.Text)
    End If

    ws紀錄_External.Cells(Max, 8) = 住戶回饋_num + 區權會抵扣_num
    ws紀錄_External.Cells(Max, 9) = 應繳金額_2.Caption
    ws紀錄_External.Cells(Max, 10) = NextReceiptNo
    ws紀錄_External.Cells(Max, 11) = 收據繳交月份_2.Text
    ws紀錄_External.Cells(Max, 12) = Me.收款人.List(Me.收款人.ListIndex)

    If Not ws住戶 Is Nothing Then
        todayFormatted = Format(Date, "yyyy/m/d")
        For i = 1 To Val(想繳幾月.Text)
            ws住戶.Cells(住戶總表_X, 住戶總表_繳到幾月Y + i).Value = todayFormatted
        Next i
    End If

    timestamp = Format(Now, "yyyymmdd_hhmmss")
    receiptID = NextReceiptNo
    backupPath = folderPath & "住戶總表_" & timestamp & "_" & receiptID & ".xlsx"
    
    Application.DisplayAlerts = False
    wbExternal.SaveAs fileName:=backupPath, FileFormat:=xlOpenXMLWorkbook
    Application.DisplayAlerts = True

    wbExternal.Close SaveChanges:=False
    Set ws住戶 = Nothing
    Set ws紀錄_External = Nothing
    Set wbExternal = Nothing
    Call 刪除舊備份檔案

    
    Sheets("收據範本").PrintOut From:=1, To:=1
    清空表單
    Unload Form
    
    Exit Sub

ErrorHandler_Save:
    Application.DisplayAlerts = True
    MsgBox "儲存或備份檔案時發生錯誤：" & Err.Description, vbCritical, "處理錯誤"
    If Not wbExternal Is Nothing Then
        wbExternal.Close SaveChanges:=False
    End If
    Set ws住戶 = Nothing
    Set wbExternal = Nothing
    Set ws紀錄_External = Nothing
    Resume Next
End Sub








Sub 填入收據範本()
    Sheets("收據範本").Cells(3, 5) = Current_Time
    Sheets("收據範本").Cells(19, 5) = Current_Time
    Sheets("收據範本").Cells(2, 3) = "'" & 棟樓別
    Sheets("收據範本").Cells(18, 3) = "'" & 棟樓別
    Sheets("收據範本").Cells(3, 3) = 收據繳交月份_2.Text
    Sheets("收據範本").Cells(19, 3) = 收據繳交月份_2.Text
    Sheets("收據範本").Cells(5, 3) = 所有權人_2.Caption
    Sheets("收據範本").Cells(21, 3) = 所有權人_2.Caption
    Sheets("收據範本").Cells(7, 2) = 管理費_2.Text
    Sheets("收據範本").Cells(23, 2) = 管理費_2.Text
    Sheets("收據範本").Cells(11, 3) = 汽車車位_2.Caption
    Sheets("收據範本").Cells(27, 3) = 汽車車位_2.Caption
    Sheets("收據範本").Cells(7, 3) = 汽車清潔費_2.Text
    Sheets("收據範本").Cells(23, 3) = 汽車清潔費_2.Text
    Sheets("收據範本").Cells(11, 5) = 機車車位_2.Caption
    Sheets("收據範本").Cells(27, 5) = 機車車位_2.Caption
    Sheets("收據範本").Cells(7, 4) = 機車清潔費_2.Text
    Sheets("收據範本").Cells(23, 4) = 機車清潔費_2.Text
    Sheets("收據範本").Cells(7, 5) = 小計_2.Text
    Sheets("收據範本").Cells(23, 5) = 小計_2.Text
    Sheets("收據範本").Cells(9, 4) = 區權會抵扣_2.Text
    Sheets("收據範本").Cells(25, 4) = 區權會抵扣_2.Text
    Sheets("收據範本").Cells(9, 3) = 住戶回饋_2.Text
    Sheets("收據範本").Cells(25, 3) = 住戶回饋_2.Text
    Sheets("收據範本").Cells(9, 5) = 應繳金額_2.Caption
    Sheets("收據範本").Cells(25, 5) = 應繳金額_2.Caption
    Call listbox_print
End Sub

Sub listbox_print()
    Dim wsTarget As Worksheet
    Set wsTarget = ThisWorkbook.Sheets("收據範本")
    If 收款人.ListIndex <> -1 Then
        wsTarget.Cells(13, 5).Value = Me.收款人.List(Me.收款人.ListIndex)
        wsTarget.Cells(29, 5).Value = Me.收款人.List(Me.收款人.ListIndex)
    Else
        MsgBox "請先從列表中選擇一個收款人。"
    End If
End Sub

Sub 手動修改_Click()
    Dim ctrl As MSForms.Control
    For Each ctrl In Me.Controls
        If TypeOf ctrl Is MSForms.Label Or TypeOf ctrl Is MSForms.TextBox Or TypeOf ctrl Is MSForms.CommandButton Then
            With ctrl
                If TypeOf ctrl Is MSForms.TextBox Then
                    .Locked = False
                End If
            End With
        End If
    Next ctrl
End Sub

Sub 重新計算_Click()
    小計_2.Text = Val(管理費_2.Text) + Val(汽車清潔費_2.Text) + Val(機車清潔費_2.Text)
    應繳金額_2.Caption = Val(小計_2.Text) - Val(區權會抵扣_2.Text) - Val(住戶回饋_2.Text)
    Call 填入收據範本
End Sub



Sub 刪除舊備份檔案()
    Const MaxBackupsToKeep As Long = 10
    Dim folderPath As String
    folderPath = GetFolderPath()
    Dim fileName As String
    Dim fileList As Object
    Dim j As Long
    Dim currentTimestamp As Double
    Dim timestamp As String
    Dim dateTimeString As String
    Dim allTimestamps As Variant
    Dim i_sort As Long, j_sort As Long
    Dim filesToDelete As Long
    On Error GoTo ErrorHandler_Cleanup
    Set fileList = CreateObject("Scripting.Dictionary")

    fileName = Dir(folderPath & "住戶總表_????????_??????*.xlsx")    ' ← 支援有/無PC編號檔名
    Do While fileName <> ""
        timestamp = Mid(fileName, 6, 15)
        dateTimeString = Left(timestamp, 4) & "/" & Mid(timestamp, 5, 2) & "/" & Mid(timestamp, 7, 2) & " " & _
                         Mid(timestamp, 10, 2) & ":" & Mid(timestamp, 12, 2) & ":" & Right(timestamp, 2)
        On Error Resume Next
        currentTimestamp = CDbl(CDate(dateTimeString))
        If Err.Number = 0 And Not fileList.Exists(currentTimestamp) Then
            fileList.Add Key:=currentTimestamp, Item:=fileName
        End If
        Err.Clear
        On Error GoTo ErrorHandler_Cleanup
        fileName = Dir
    Loop

    If fileList.count > MaxBackupsToKeep Then
        allTimestamps = fileList.Keys
        Dim count As Long: count = UBound(allTimestamps)
        For i_sort = LBound(allTimestamps) To count - 1
            For j_sort = i_sort + 1 To count
                If allTimestamps(i_sort) < allTimestamps(j_sort) Then
                    Dim tempValue As Variant
                    tempValue = allTimestamps(i_sort)
                    allTimestamps(i_sort) = allTimestamps(j_sort)
                    allTimestamps(j_sort) = tempValue
                End If
            Next j_sort
        Next i_sort
        filesToDelete = fileList.count - MaxBackupsToKeep
        For j = MaxBackupsToKeep To fileList.count - 1
            Dim filePathToDelete As String
            currentTimestamp = allTimestamps(j)
            fileName = fileList.Item(currentTimestamp)
            filePathToDelete = folderPath & fileName
            On Error Resume Next
            Application.Wait Now + TimeValue("00:00:00")
            Kill filePathToDelete
            If Err.Number <> 0 Then
                Err.Clear
            End If
            On Error GoTo ErrorHandler_Cleanup
        Next j
    End If

ExitHandler_Cleanup:
    Set fileList = Nothing
    Exit Sub

ErrorHandler_Cleanup:
    MsgBox "執行備份清理程序時發生錯誤：" & Err.Description & vbCrLf & _
           "（請檢查 Google Drive 正在同步或鎖定檔案）", vbCritical, "檔案清理錯誤"
    Resume ExitHandler_Cleanup
End Sub





Function GetHostName() As String
    Dim objWMIService As Object
    Dim colItems As Object
    Dim objItem As Object
    Set objWMIService = GetObject("winmgmts:\\.\root\cimv2")
    Set colItems = objWMIService.ExecQuery("Select * from Win32_ComputerSystem")
    For Each objItem In colItems
        GetHostName = objItem.Name
        Exit For
    Next
End Function

Function GetFolderPath() As String
    Dim hostName As String
    hostName = GetHostName()
    If hostName = "TWR-LT-A62827" Then
        GetFolderPath = "D:\temp\管理費收入電腦系統\"
    Else
        GetFolderPath = "G:\我的雲端硬碟\管理費收入電腦系統\"
    End If
End Function





'------------------------------
' 清空表單副程式
'------------------------------
Sub 清空表單()
    Dim ctrl As MSForms.Control
    For Each ctrl In Me.Controls
        If TypeOf ctrl Is MSForms.TextBox Then
            ctrl.Text = ""
        End If
    Next ctrl

    ' 清空指定 Label
    繳到幾月_2.Caption = ""
    應繳金額_2.Caption = ""
    所有權人_2.Caption = ""
    汽車車位_2.Caption = ""
    機車車位_2.Caption = ""
End Sub

'------------------------------
' 印收據按鈕
'------------------------------
Private Sub 印收據_Click()
    If MsgBox("請確定印表機紙張放好了，並且印表機已經切過來了。是否要列印收據？", vbYesNo + vbQuestion, "列印前確認") = vbYes Then
        Sheets("收據範本").PrintOut From:=1, To:=1
        清空表單
    End If
End Sub
