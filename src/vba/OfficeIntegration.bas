Attribute VB_Name = "OfficeIntegration"
Option Explicit

Public Sub CreateOutlookDraft()
    Dim token As String, directory As Object, i As Long, id As String, email As String, recipients As String
    Dim outlook As Object, draft As Object, body As String
    On Error GoTo Failed
    token = AcquireLock("CREATE OUTLOOK DRAFT")
    If Len(CStr(WS("Control").Range("K5").Value2)) = 0 Then Err.Raise vbObjectError + 400, , "Process a run first."
    If CStr(WS("Control").Range("K12").Value2) <> "COMPLETED" Then Err.Raise vbObjectError + 401, , "No completed current run."
    Set directory = NewMap()
    For i = 2 To LastRow(WS("Team Directory"))
        id = Normal(WS("Team Directory").Cells(i, 1).Value2)
        email = Trim$(CStr(WS("Team Directory").Cells(i, 3).Value2))
        If LCase$(Right$(email, 12)) <> "@example.com" Or InStr(email, ";") > 0 Or InStr(email, " ") > 0 Or Len(email) < 13 Then Err.Raise vbObjectError + 402, , "Demo recipients must use example.com only."
        If directory.Exists(id) Then Err.Raise vbObjectError + 403, , "Duplicate directory ID."
        directory.Add id, email
    Next i
    body = "<p>Synthetic demonstration only.</p><p>Hotlist assignment summary for run " & Html(CStr(WS("Control").Range("K5").Value2)) & "</p><table border='1'><tr><th>Analyst</th><th>New volume</th><th>Carryover / Review</th><th>Assigned</th></tr>"
    ' Use the run summary rather than changed post-run availability settings.
    For i = 2 To LastRow(WS("Assignment Summary"))
        id = Normal(WS("Assignment Summary").Cells(i, 1).Value2)
        If Len(id) > 0 And id <> "TOTAL" Then
            If Not directory.Exists(id) Then Err.Raise vbObjectError + 404, , "Missing analyst email."
            If Len(recipients) > 0 Then recipients = recipients & ";"
            recipients = recipients & directory(id)
            body = body & "<tr><td>" & Html(CStr(WS("Assignment Summary").Cells(i, 2).Value2)) & "</td><td>" & WS("Assignment Summary").Cells(i, 4).Value2 & "</td><td>" & WS("Assignment Summary").Cells(i, 5).Value2 & "</td><td>" & WS("Assignment Summary").Cells(i, 6).Value2 & "</td></tr>"
        End If
    Next i
    If Len(recipients) = 0 Then Err.Raise vbObjectError + 405, , "No run recipients."
    Set outlook = CreateObject("Outlook.Application")
    Set draft = outlook.CreateItem(0)
    draft.To = recipients
    draft.Subject = "Synthetic Hotlist assignment - " & Format$(CDate(WS("Control").Range("K6").Value2), "yyyy-mm-dd")
    draft.HTMLBody = body & "</table><p>Review before sending. All contacts are fictional.</p>"
    draft.Display
Finished:
    If Len(token) > 0 Then ReleaseLock token
    Exit Sub
Failed:
    MsgBox "Outlook draft could not be created: " & Err.Description & vbCrLf & "Use compatible desktop Outlook. No message was sent.", vbExclamation
    Resume Finished
End Sub

Private Function Html(ByVal text As String) As String
    text = Replace(text, "&", "&amp;")
    text = Replace(text, "<", "&lt;")
    text = Replace(text, ">", "&gt;")
    Html = Replace(text, """", "&quot;")
End Function

Public Sub ExportHistory()
    Dim token As String, path As Variant, exportWorkbook As Workbook, source As Worksheet, target As Worksheet, i As Long
    On Error GoTo Failed
    token = AcquireLock("EXPORT HISTORY")
    If Not SaveProgressCore(token) Then GoTo Finished
    Set source = WS("Inventory History")
    If LastRow(source) < 2 Then Err.Raise vbObjectError + 406, , "No history to export."
    path = Application.GetSaveAsFilename("Synthetic_History_" & Format$(UtcNow(), "yyyy-mm-dd") & ".xlsx", "Excel Workbook (*.xlsx), *.xlsx")
    If VarType(path) = vbBoolean Then GoTo Finished
    AssertLock token
    Set exportWorkbook = Workbooks.Add(xlWBATWorksheet)
    Set target = exportWorkbook.Worksheets(1): target.Name = "Synthetic History"
    source.Range("A1:AG" & LastRow(source)).Copy
    target.Range("A1").PasteSpecial xlPasteValuesAndNumberFormats
    target.Range("A1").PasteSpecial xlPasteFormats
    Application.CutCopyMode = False
    For i = 1 To FIELD_COUNT
        target.Columns(i).ColumnWidth = source.Columns(i).ColumnWidth
    Next i
    target.Range("A1:AG" & LastRow(source)).AutoFilter
    ' Freeze only the header row in the exported history; no horizontal split.
    exportWorkbook.Activate
    target.Activate
    With exportWorkbook.Windows(1)
        .FreezePanes = False
        .SplitColumn = 0
        .SplitRow = 1
        .FreezePanes = True
    End With
    exportWorkbook.SaveAs Filename:=CStr(path), FileFormat:=xlOpenXMLWorkbook
    exportWorkbook.Close SaveChanges:=False
    Set exportWorkbook = Nothing
    MsgBox "Synthetic history exported as a macro-free workbook.", vbInformation
Finished:
    If Len(token) > 0 Then ReleaseLock token
    Exit Sub
Failed:
    MsgBox "History export failed: " & Err.Description, vbCritical
    On Error Resume Next
    Application.CutCopyMode = False
    If Not exportWorkbook Is Nothing Then exportWorkbook.Close SaveChanges:=False
    On Error GoTo 0
    Resume Finished
End Sub
