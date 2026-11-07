Attribute VB_Name = "DemoUtilities"
Option Explicit

Public Const DEMO_VERSION As String = "0.3.2"
Public Const FIELD_COUNT As Long = 33
Public Const RAW_COUNT As Long = 20
Public Const COL_OWNER As Long = 21
Public Const COL_TYPE As Long = 22
Public Const COL_RUN As Long = 23
Public Const COL_IDENTITY As Long = 24
Public Const COL_GROUP As Long = 25
Public Const COL_BASIS As Long = 26
Public Const COL_VOLUME As Long = 27
Public Const COL_MEDIA As Long = 28
Public Const COL_STREAM As Long = 29
Public Const COL_REASON As Long = 30
Public Const COL_TIME As Long = 31
Public Const COL_SAVED As Long = 32
Private Type GuidValue
    A As Long
    B As Integer
    C As Integer
    D(0 To 7) As Byte
End Type
Private Type SystemTimeValue
    Year As Integer
    Month As Integer
    Weekday As Integer
    Day As Integer
    Hour As Integer
    Minute As Integer
    Second As Integer
    Millisecond As Integer
End Type
#If VBA7 Then
Private Declare PtrSafe Function CoCreateGuid Lib "ole32" (ByRef value As GuidValue) As Long
Private Declare PtrSafe Function StringFromGUID2 Lib "ole32" (ByRef value As GuidValue, ByVal buffer As LongPtr, ByVal size As Long) As Long
Private Declare PtrSafe Sub GetSystemTime Lib "kernel32" (ByRef value As SystemTimeValue)
#Else
Private Declare Function CoCreateGuid Lib "ole32" (ByRef value As GuidValue) As Long
Private Declare Function StringFromGUID2 Lib "ole32" (ByRef value As GuidValue, ByVal buffer As Long, ByVal size As Long) As Long
Private Declare Sub GetSystemTime Lib "kernel32" (ByRef value As SystemTimeValue)
#End If

Public Function WS(ByVal title As String) As Worksheet
    Set WS = ThisWorkbook.Worksheets(title)
End Function

Public Function NewMap() As Object
    Set NewMap = CreateObject("Scripting.Dictionary")
    NewMap.CompareMode = vbBinaryCompare
End Function

Public Function Normal(ByVal value As Variant) As String
    If IsError(value) Or IsNull(value) Then Err.Raise vbObjectError + 100, , "Invalid cell value."
    Normal = UCase$(Trim$(CStr(value)))
End Function

Public Function KeyPart(ByVal value As Variant) As String
    Dim text As String
    text = Normal(value)
    KeyPart = CStr(Len(text)) & ":" & text
End Function

Public Function InvoiceKey(ByVal row As Variant) As String
    InvoiceKey = KeyPart(row(1)) & KeyPart(row(6)) & KeyPart(row(4)) & KeyPart(row(8))
End Function

Public Function GroupKey(ByVal row As Variant) As String
    GroupKey = KeyPart(row(1)) & KeyPart(row(6)) & KeyPart(row(4))
End Function

Public Function RecordKey(ByVal row As Variant) As String
    RecordKey = KeyPart(row(COL_RUN)) & KeyPart(InvoiceKey(row)) & KeyPart(row(9))
End Function

Public Function ExactMediaKey(ByVal row As Variant) As String
    ExactMediaKey = KeyPart(row(1)) & KeyPart(row(4)) & KeyPart(row(2))
End Function

Public Function LastRow(ByVal sheet As Worksheet) As Long
    Dim found As Range
    Set found = sheet.Cells.Find(What:="*", After:=sheet.Cells(1, 1), LookIn:=xlFormulas, SearchOrder:=xlByRows, SearchDirection:=xlPrevious)
    LastRow = 1
    If Not found Is Nothing Then LastRow = found.row
End Function

Public Function RawHeaders() As Variant
    RawHeaders = Array("Location", "Publication ID", "Publication Name", "Client ID", "Client Name", "Vendor ID", "Vendor Name", "Invoice Number", "Line ID", "Batch ID", "Invoice Date", "Amount", "Input Date", "Status", "Hotlist Comment 1", "Hotlist Comment 2", "Activity Date", "Agent", "Agent Comment", "Currency")
End Function

Public Function AllHeaders() As Variant
    AllHeaders = Split(Join(RawHeaders(), "|") & "|Assigned Analyst|Invoice Type|Run ID|Invoice Key|Group Key|Assignment Basis|Volume Class|Resolved Media|Workstream|Exclusion Reason|Run At UTC|Work Saved At UTC|Duplicate in Run", "|")
End Function

Public Function AssignedHeaders() As Variant
    AssignedHeaders = Split("Assigned Analyst|Invoice Type|Run ID|Location|Publication ID|Publication Name|Client ID|Invoice Number|Invoice Date|Amount|Vendor ID|Batch ID|Status|Hotlist Comment 1|Hotlist Comment 2|Agent|Agent Comment|Client Name|Vendor Name|Input Date|Activity Date|Currency|Line ID|Resolved Media|Workstream|Group Key|Assignment Basis|Volume Class|Duplicate in Run|Invoice Key|Exclusion Reason|Run At UTC|Work Saved At UTC", "|")
End Function

' Internal rows and history retain their canonical order; only Assigned Work is projected.
Public Function WorksheetColumn(ByVal sheet As Worksheet, ByVal canonicalColumn As Long) As Long
    Dim displayColumns As Variant
    If sheet.Name = "Assigned Work" Then
        displayColumns = Array(4, 5, 6, 7, 18, 11, 19, 8, 23, 12, 9, 10, 20, 13, 14, 15, 21, 16, 17, 22, 1, 2, 3, 30, 26, 27, 28, 24, 25, 31, 32, 33, 29)
        WorksheetColumn = CLng(displayColumns(canonicalColumn - 1))
    Else
        WorksheetColumn = canonicalColumn
    End If
End Function

Public Sub CheckHeaders(ByVal sheet As Worksheet, ByVal expected As Variant)
    Dim i As Long
    For i = LBound(expected) To UBound(expected)
        If Normal(sheet.Cells(1, i + 1).Value2) <> Normal(expected(i)) Then Err.Raise vbObjectError + 101, , sheet.Name & ": expected header " & expected(i)
    Next i
End Sub

Public Function ReadRow(ByVal sheet As Worksheet, ByVal index As Long, ByVal count As Long) As Variant
    Dim result() As Variant, j As Long
    ReDim result(1 To FIELD_COUNT)
    For j = 1 To count
        result(j) = sheet.Cells(index, WorksheetColumn(sheet, j)).Value2
    Next j
    ReadRow = result
End Function

Public Function HasRawValues(ByVal row As Variant) As Boolean
    Dim i As Long
    For i = 1 To RAW_COUNT
        If Len(Normal(row(i))) > 0 Then HasRawValues = True: Exit Function
    Next i
End Function

Public Sub WriteRows(ByVal sheet As Worksheet, ByVal rows As Collection, Optional ByVal appendRows As Boolean = False)
    Dim start As Long, i As Long, j As Long, data() As Variant, row As Variant
    start = 2
    If appendRows Then
        start = LastRow(sheet) + 1
    Else
        If LastRow(sheet) > 1 Then sheet.Range(sheet.Cells(2, 1), sheet.Cells(LastRow(sheet), FIELD_COUNT)).ClearContents
    End If
    If rows.count = 0 Then
        ResizeOutputTable sheet
        Exit Sub
    End If
    ReDim data(1 To rows.count, 1 To FIELD_COUNT)
    For i = 1 To rows.count
        row = rows(i)
        For j = 1 To FIELD_COUNT
            data(i, WorksheetColumn(sheet, j)) = row(j)
        Next j
    Next i
    sheet.Cells(start, 1).Resize(rows.count, FIELD_COUNT).Value2 = data
    For Each row In Array(11, 13, 17)
        sheet.Cells(start, WorksheetColumn(sheet, CLng(row))).Resize(rows.count, 1).NumberFormat = "yyyy-mm-dd"
    Next row
    sheet.Cells(start, WorksheetColumn(sheet, COL_TIME)).Resize(rows.count, 2).NumberFormat = "yyyy-mm-dd hh:mm:ss"
    ResizeOutputTable sheet
End Sub

Public Function UtcNow() As Date
    Dim value As SystemTimeValue
    GetSystemTime value
    UtcNow = DateSerial(value.Year, value.Month, value.Day) + TimeSerial(value.Hour, value.Minute, value.Second)
End Function

Public Function NewRunId() As String
    Dim value As GuidValue, buffer As String
    buffer = String$(39, vbNullChar)
    If CoCreateGuid(value) <> 0 Then Err.Raise vbObjectError + 102, , "Cannot generate Run ID."
    If StringFromGUID2(value, StrPtr(buffer), 39) = 0 Then Err.Raise vbObjectError + 103, , "Cannot format Run ID."
    NewRunId = Format$(UtcNow(), "yyyymmdd\Thhnnss\Z") & "-" & LCase$(Mid$(buffer, 2, 36))
End Function

Public Function AcquireLock(ByVal action As String) As String
    Dim control As Worksheet, token As String
    Set control = WS("Control")
    If Len(CStr(control.Range("K9").Value2)) > 0 Then
        Err.Raise vbObjectError + 104, , "Workbook is locked by " & control.Range("K10").Value2 & " since " & control.Range("K11").Text & " UTC. Wait or use RecoverStaleLock after checking with the coordinator."
    End If
    token = NewRunId()
    control.Range("K9").Value2 = token
    control.Range("K10").Value2 = Application.UserName & " / " & action
    control.Range("K11").Value2 = CDbl(UtcNow())
    control.Range("K11").NumberFormat = "yyyy-mm-dd hh:mm:ss"
    AssertLock token
    AcquireLock = token
End Function

Public Sub AssertLock(ByVal token As String)
    If CStr(WS("Control").Range("K9").Value2) <> token Then Err.Raise vbObjectError + 105, , "Processing lock changed. Operation stopped."
End Sub

Public Sub ReleaseLock(ByVal token As String)
    If CStr(WS("Control").Range("K9").Value2) = token Then WS("Control").Range("K9:K11").ClearContents
End Sub

Public Sub RecoverStaleLock()
    Dim control As Worksheet
    Set control = WS("Control")
    If Len(CStr(control.Range("K9").Value2)) = 0 Then MsgBox "No active lock.": Exit Sub
    If Not IsNumeric(control.Range("K11").Value2) Then MsgBox "Invalid lock time; ask the coordinator to inspect the workbook.": Exit Sub
    If (CDbl(UtcNow()) - CDbl(control.Range("K11").Value2)) * 1440 < 30 Then MsgBox "Lock is less than 30 minutes old. Recovery is blocked.": Exit Sub
    If MsgBox("Confirm with the named coordinator that no macro is running. Clear this stale lock?", vbYesNo + vbExclamation) = vbYes Then control.Range("K9:K11").ClearContents
End Sub

Public Function SnapshotSheets(ByVal names As Variant) As Object
    Dim snapshots As Object, sheetName As Variant, sheet As Worksheet
    Set snapshots = NewMap()
    For Each sheetName In names
        Set sheet = WS(CStr(sheetName))
        snapshots.Add CStr(sheetName), sheet.Range(sheet.Cells(1, 1), sheet.Cells(LastRow(sheet), FIELD_COUNT)).Value2
    Next sheetName
    Set SnapshotSheets = snapshots
End Function

Public Function RestoreSheets(ByVal snapshots As Object) As Boolean
    On Error GoTo Failed
    Dim sheetName As Variant, data As Variant, sheet As Worksheet
    For Each sheetName In snapshots.Keys
        Set sheet = WS(CStr(sheetName))
        data = snapshots(sheetName)
        sheet.Range(sheet.Cells(1, 1), sheet.Cells(LastRow(sheet), FIELD_COUNT)).ClearContents
        sheet.Cells(1, 1).Resize(UBound(data, 1), UBound(data, 2)).Value2 = data
        If sheet.Name <> "Assignment Summary" And sheet.Name <> "Raw Hotlist" Then ResizeOutputTable sheet
    Next sheetName
    RestoreSheets = True
Failed:
End Function

Public Sub ValidateNumericDate(ByVal value As Variant, ByVal label As String, Optional ByVal allowBlank As Boolean = False)
    If Len(Normal(value)) = 0 And allowBlank Then Exit Sub
    If Not IsNumeric(value) Then Err.Raise vbObjectError + 106, , label & " must be a typed Excel date, not free text."
    If CDbl(value) < 36526 Or CDbl(value) > 73050 Then Err.Raise vbObjectError + 107, , label & " is outside the demo date range."
End Sub

Public Function NewRows() As Collection
    Set NewRows = New Collection
End Function

Public Sub ResizeOutputTable(ByVal sheet As Worksheet)
    If sheet.ListObjects.count = 0 Then Exit Sub
    Dim bottom As Long
    bottom = LastRow(sheet)
    If bottom < 2 Then bottom = 2
    sheet.ListObjects(1).Resize sheet.Range(sheet.Cells(1, 1), sheet.Cells(bottom, FIELD_COUNT))
End Sub

Public Function HasRecordValues(ByVal row As Variant) As Boolean
    Dim i As Long
    For i = 1 To FIELD_COUNT
        If Len(Normal(row(i))) > 0 Then HasRecordValues = True: Exit Function
    Next i
End Function
