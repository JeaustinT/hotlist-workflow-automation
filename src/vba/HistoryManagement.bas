Attribute VB_Name = "HistoryManagement"
Option Explicit

Public Sub LoadHistory(ByVal seen As Object, ByVal previous As Object)
    Dim latest As Object, latestRun As Object, i As Long, row As Variant, key As String, stamp As Double, runId As String
    Set latest = NewMap(): Set latestRun = NewMap()
    For i = 2 To LastRow(WS("Inventory History"))
        row = ReadRow(WS("Inventory History"), i, FIELD_COUNT)
        If HasRawValues(row) Then
            key = InvoiceKey(row)
            ValidateNumericDate row(COL_TIME), "Historical Run At UTC"
            stamp = CDbl(row(COL_TIME)): runId = CStr(row(COL_RUN))
            seen(key) = True
            If Not latest.Exists(key) Then
                latest.Add key, -1#: latestRun.Add key, ""
            End If
            If stamp > CDbl(latest(key)) Or (stamp = CDbl(latest(key)) And runId > CStr(latestRun(key))) Then
                latest(key) = stamp: latestRun(key) = runId
                previous(key) = Normal(row(COL_OWNER))
            End If
        End If
    Next i
End Sub

Public Sub SaveProgress()
    Dim token As String
    On Error GoTo Failed
    token = AcquireLock("SAVE PROGRESS")
    Call SaveProgressCore(token)
Finished:
    If Len(token) > 0 Then ReleaseLock token
    Exit Sub
Failed:
    MsgBox "Save Progress stopped: " & Err.Description, vbCritical
    Resume Finished
End Sub

Public Function SaveProgressCore(ByVal token As String) As Boolean
    Dim index As Object, used As Object, changes As New Collection
    Dim i As Long, j As Long, row As Variant, historical As Variant, key As String, item As Variant
    Dim evaluated As Long, matched As Long, updated As Long, unchanged As Long, cleared As Long
    Dim unmatched As Long, skipped As Long, errors As Long, ambiguous As Long, target As Long, different As Boolean
    Dim saved As Object, writing As Boolean, priorEvents As Boolean, detail As String
    On Error GoTo Failed
    AssertLock token
    CheckHeaders WS("Assigned Work"), AssignedHeaders()
    CheckHeaders WS("Inventory History"), AllHeaders()
    Set index = NewMap(): Set used = NewMap()
    For i = 2 To LastRow(WS("Inventory History"))
        row = ReadRow(WS("Inventory History"), i, FIELD_COUNT)
        If HasRawValues(row) Then
            key = RecordKey(row)
            If index.Exists(key) Then index(key) = -1 Else index.Add key, i
        End If
    Next i
    For i = 2 To LastRow(WS("Assigned Work"))
        row = ReadRow(WS("Assigned Work"), i, FIELD_COUNT)
        evaluated = evaluated + 1
        If Not HasRecordValues(row) Then
            skipped = skipped + 1
        ElseIf Len(Normal(row(COL_RUN))) = 0 Or Len(Normal(row(9))) = 0 Or Len(Normal(row(1))) = 0 Or Len(Normal(row(4))) = 0 Or Len(Normal(row(6))) = 0 Or Len(Normal(row(8))) = 0 Then
            errors = errors + 1
        Else
            key = RecordKey(row)
            If used.Exists(key) Then
                ambiguous = ambiguous + 1
            Else
                used.Add key, True
                If Not index.Exists(key) Then
                    unmatched = unmatched + 1
                ElseIf CLng(index(key)) < 0 Then
                    ambiguous = ambiguous + 1
                Else
                    target = CLng(index(key))
                    historical = ReadRow(WS("Inventory History"), target, FIELD_COUNT)
                    matched = matched + 1: different = False
                    For j = 18 To 19
                        If CStr(row(j)) <> CStr(historical(j)) Then
                            different = True
                            If Len(CStr(historical(j))) > 0 And Len(CStr(row(j))) = 0 Then cleared = cleared + 1
                        End If
                    Next j
                    If different Then
                        updated = updated + 1
                        changes.Add Array(target, row(18), row(19))
                    Else
                        unchanged = unchanged + 1
                    End If
                End If
            End If
        End If
    Next i
    detail = "Rows evaluated: " & evaluated & vbCrLf & "Rows matched: " & matched & vbCrLf & "Rows updated: " & updated & vbCrLf & "Rows unchanged: " & unchanged & vbCrLf & "Values cleared: " & cleared & vbCrLf & "Rows not matched: " & unmatched & vbCrLf & "Rows skipped: " & skipped & vbCrLf & "Errors: " & errors & vbCrLf & "Ambiguous rows: " & ambiguous
    If errors + unmatched + ambiguous > 0 Then
        MsgBox "Progress reconciliation blocked; no values written. Update/clear counts below are proposed only." & vbCrLf & detail, vbExclamation
        Exit Function
    End If
    If cleared > 0 Then
        If MsgBox(cleared & " historical progress values will be cleared. Continue?", vbYesNo + vbExclamation) <> vbYes Then
            MsgBox "Canceled. No progress values changed. Replacement and export are blocked until edits are resolved. Counts below are proposed only." & vbCrLf & detail, vbInformation
            Exit Function
        End If
    End If
    Set saved = SnapshotSheets(Array("Inventory History"))
    priorEvents = Application.EnableEvents
    Application.EnableEvents = False
    writing = True
    For Each item In changes
        AssertLock token
        target = CLng(item(0))
        WS("Inventory History").Cells(target, 18).Value2 = item(1)
        WS("Inventory History").Cells(target, 19).Value2 = item(2)
        WS("Inventory History").Cells(target, COL_SAVED).Value2 = CDbl(UtcNow())
        If CStr(WS("Inventory History").Cells(target, 18).Value2) <> CStr(item(1)) Or CStr(WS("Inventory History").Cells(target, 19).Value2) <> CStr(item(2)) Then Err.Raise vbObjectError + 300, , "Progress write verification failed."
    Next item
    writing = False
    Application.EnableEvents = priorEvents
    SaveProgressCore = True
    MsgBox "Progress reconciliation completed." & vbCrLf & detail, vbInformation
    Exit Function
Failed:
    detail = Err.Description
    If writing Then
        detail = detail & vbCrLf & "Rollback completed: " & CStr(RestoreSheets(saved))
        Application.EnableEvents = priorEvents
    End If
    MsgBox "Progress save failed: " & detail, vbCritical
End Function
