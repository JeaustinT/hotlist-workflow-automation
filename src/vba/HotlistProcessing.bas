Attribute VB_Name = "HotlistProcessing"
Option Explicit

' All outputs are staged before publication. This guard is local and advisory.
Public Sub ProcessHotlist()
    Dim token As String, snapshots As Object, metadata As Variant, priorStatus As Variant
    Dim priorCalc As XlCalculation, priorEvents As Boolean, priorScreen As Boolean
    Dim writing As Boolean, recovery As Boolean, message As String
    priorCalc = Application.Calculation
    priorEvents = Application.EnableEvents
    priorScreen = Application.ScreenUpdating
    On Error GoTo Failed
    token = AcquireLock("PROCESS HOTLIST")
    CheckHeaders WS("Raw Hotlist"), RawHeaders()
    CheckHeaders WS("Assigned Work"), AssignedHeaders()
    CheckHeaders WS("Exceptions & Review"), AllHeaders()
    CheckHeaders WS("Inventory History"), AllHeaders()
    If CStr(WS("Control").Range("K4").Value2) <> DEMO_VERSION Then Err.Raise vbObjectError + 200, , "Workbook/source version mismatch."

    Dim people As Object, weights As Object, loads As Object, locations As Object
    Dim exact As Object, fallback As Object, seen As Object, previous As Object
    Set people = NewMap(): Set weights = NewMap(): Set loads = NewMap(): Set locations = NewMap()
    Set exact = NewMap(): Set fallback = NewMap(): Set seen = NewMap(): Set previous = NewMap()
    LoadConfiguration people, weights, loads, locations, exact, fallback
    LoadHistory seen, previous

    Dim inputRows As New Collection, assigned As New Collection, exceptions As New Collection
    Dim rows As Object, groups As Object, votes As Object, lineKeys As Object, identityCounts As Object
    Set rows = NewMap(): Set groups = NewMap(): Set votes = NewMap()
    Set lineKeys = NewMap(): Set identityCounts = NewMap()
    Dim i As Long, j As Long, row As Variant, key As String, lineId As String, signature As String
    Dim lastDivision As Double, runId As String, runAt As Double, stream As String, reason As String
    lastDivision = CDbl(WS("Control").Range("K6").Value2)
    runAt = CDbl(UtcNow()): runId = NewRunId()
    For i = 2 To LastRow(WS("Inventory History"))
        If CStr(WS("Inventory History").Cells(i, COL_RUN).Value2) = runId Then Err.Raise vbObjectError + 201, , "Run ID collision. Retry."
    Next i
    For i = 2 To LastRow(WS("Raw Hotlist"))
        row = ReadRow(WS("Raw Hotlist"), i, RAW_COUNT)
        If HasRawValues(row) Then
            For j = 1 To RAW_COUNT
                signature = signature & KeyPart(row(j))
            Next j
            inputRows.Add row
            If Len(Normal(row(1))) = 0 Or Len(Normal(row(4))) = 0 Or Len(Normal(row(6))) = 0 Or Len(Normal(row(8))) = 0 Or Len(Normal(row(9))) = 0 Then Err.Raise vbObjectError + 202, , "Missing identity field at input row " & i
            lineId = Normal(row(9))
            If lineKeys.Exists(lineId) Then Err.Raise vbObjectError + 203, , "Line ID must be unique within a run: " & lineId
            lineKeys.Add lineId, True
            ValidateNumericDate row(11), "Invoice Date"
            ValidateNumericDate row(13), "Input Date"
            ValidateNumericDate row(17), "Activity Date", True
            If Not IsNumeric(row(12)) Then Err.Raise vbObjectError + 204, , "Amount must be numeric at row " & i
            key = InvoiceKey(row)
            If Not identityCounts.Exists(key) Then identityCounts.Add key, 0
            identityCounts(key) = CLng(identityCounts(key)) + 1
            row(COL_IDENTITY) = key: row(COL_GROUP) = GroupKey(row)
            row(COL_RUN) = runId: row(COL_TIME) = runAt
            row(COL_TYPE) = "NEW": row(COL_VOLUME) = "NEW VOLUME"
            If seen.Exists(key) Then
                row(COL_TYPE) = "REPEATED": row(COL_VOLUME) = "CARRYOVER / REVIEW"
                If Len(Normal(row(17))) > 0 Then
                    If Int(CDbl(row(17))) > Int(lastDivision) Then row(COL_VOLUME) = "NEW VOLUME"
                End If
            End If
            stream = "REVIEW": reason = ""
            If Normal(row(14)) = "RELEASED TO CLEAR" Then
                reason = "Released to Clear"
            ElseIf Not locations.Exists(Normal(row(1))) Then
                reason = "Out-of-scope location"
            Else
                key = ExactMediaKey(row)
                If exact.Exists(key) Then
                    stream = CStr(exact(key))
                ElseIf fallback.Exists(Normal(row(2))) Then
                    stream = CStr(fallback(Normal(row(2))))
                End If
                If stream = "PRINT" Then reason = "Print media"
                If stream <> "DIGITAL" And stream <> "PRINT" Then reason = "Unresolved media classification"
            End If
            row(COL_MEDIA) = stream: row(COL_STREAM) = stream
            If Len(reason) > 0 Then
                row(COL_REASON) = reason: exceptions.Add row
            Else
                key = CStr(row(COL_GROUP))
                If Not groups.Exists(key) Then groups.Add key, NewRows()
                groups(key).Add CStr(i)
                rows.Add CStr(i), row
                ' New activity always uses balancing, even when comment fields are blank.
                If CStr(row(COL_TYPE)) = "REPEATED" And CStr(row(COL_VOLUME)) <> "NEW VOLUME" Then
                    If Len(Normal(row(15))) = 0 And Len(Normal(row(16))) = 0 Then
                        If previous.Exists(CStr(row(COL_IDENTITY))) Then
                            Dim owner As String, groupVotes As Object
                            owner = CStr(previous(CStr(row(COL_IDENTITY))))
                            If people.Exists(owner) Then
                                If Not votes.Exists(key) Then votes.Add key, NewMap()
                                Set groupVotes = votes(key)
                                If Not groupVotes.Exists(owner) Then groupVotes.Add owner, 0
                                groupVotes(owner) = CLng(groupVotes(owner)) + 1
                            End If
                        End If
                    End If
                End If
            End If
        End If
    Next i
    If inputRows.count = 0 Then Err.Raise vbObjectError + 205, , "Raw Hotlist has no records."
    If signature = CStr(WS("Control").Range("K7").Value2) Then
        If MsgBox("This input matches the previous successful run. Create an explicit rerun with a new Run ID?", vbYesNo + vbExclamation) <> vbYes Then GoTo Finished
    End If
    If Len(signature) > 32767 Then Err.Raise vbObjectError + 206, , "Input exceeds this small demo's signature capacity. Use a smaller fixture."

    AllocateGroups groups, votes, rows, people, weights, loads
    Dim rowIndex As Variant
    For Each rowIndex In rows.Keys
        row = rows(rowIndex)
        row(33) = (CLng(identityCounts(CStr(row(COL_IDENTITY)))) > 1)
        assigned.Add row
    Next rowIndex
    If assigned.count + exceptions.count <> inputRows.count Then Err.Raise vbObjectError + 207, , "Input reconciliation failed."
    If Not SaveProgressCore(token) Then GoTo Finished
    AssertLock token
    Set snapshots = SnapshotSheets(Array("Assigned Work", "Exceptions & Review", "Assignment Summary", "Inventory History", "Raw Hotlist"))
    metadata = WS("Control").Range("K5:K7").Value2
    priorStatus = WS("Control").Range("K12").Value2
    Application.ScreenUpdating = False: Application.EnableEvents = False: Application.Calculation = xlCalculationManual
    writing = True
    WriteRows WS("Assigned Work"), assigned
    WriteRows WS("Exceptions & Review"), exceptions
    BuildSummary people, weights, loads, assigned, exceptions.count, runId
    WriteRows WS("Inventory History"), assigned, True
    ' Optional fault injection is disabled in the distributed workbook.
    If CStr(WS("Control").Range("K15").Value2) = "FAIL_AFTER_HISTORY" Then Err.Raise vbObjectError + 208, , "Injected validation failure."
    VerifyRows WS("Assigned Work"), assigned, False
    VerifyRows WS("Exceptions & Review"), exceptions, False
    VerifyRows WS("Inventory History"), assigned, True
    AssertLock token
    WS("Control").Range("K5").Value2 = runId
    WS("Control").Range("K6").Value2 = runAt
    WS("Control").Range("K7").Value2 = signature
    WS("Control").Range("K12").Value2 = "COMPLETED"
    WS("Raw Hotlist").Range("A2:T" & LastRow(WS("Raw Hotlist"))).ClearContents
    writing = False
    MsgBox "Hotlist prepared. Assigned: " & assigned.count & "; exceptions: " & exceptions.count & vbCrLf & "Run ID: " & runId, vbInformation
Finished:
    On Error Resume Next
    Application.Calculation = priorCalc: Application.EnableEvents = priorEvents: Application.ScreenUpdating = priorScreen
    If Len(token) > 0 Then ReleaseLock token
    On Error GoTo 0
    Exit Sub
Failed:
    message = Err.Description
    If writing Then
        recovery = RestoreSheets(snapshots)
        On Error Resume Next
        WS("Control").Range("K5:K7").Value2 = metadata
        WS("Control").Range("K12").Value2 = priorStatus
        If Err.Number <> 0 Then recovery = False
        On Error GoTo 0
        message = message & vbCrLf & "Rollback completed: " & CStr(recovery) & ". If False, reopen the saved recovery copy."
    End If
    MsgBox "Processing stopped: " & message, vbCritical
    GoTo Finished
End Sub

Private Sub VerifyRows(ByVal sheet As Worksheet, ByVal expected As Collection, ByVal appended As Boolean)
    Dim start As Long, i As Long, j As Long, row As Variant, actualRow As Variant
    start = 2
    If appended Then start = LastRow(sheet) - expected.count + 1
    For i = 1 To expected.count
        row = expected(i)
        actualRow = ReadRow(sheet, start + i - 1, FIELD_COUNT)
        For j = 1 To FIELD_COUNT
            If CStr(actualRow(j)) <> CStr(row(j)) Then Err.Raise vbObjectError + 209, , "Write verification failed: " & sheet.Name
        Next j
    Next i
End Sub

Private Sub LoadConfiguration(ByVal people As Object, ByVal weights As Object, ByVal loads As Object, ByVal locations As Object, ByVal exact As Object, ByVal fallback As Object)
    Dim i As Long, id As String, flag As String, key As String, row As Variant, directory As Object
    Set directory = NewMap()
    CheckHeaders WS("Team Directory"), Array("Analyst ID", "Analyst Name", "Email")
    For i = 2 To LastRow(WS("Team Directory"))
        id = Normal(WS("Team Directory").Cells(i, 1).Value2)
        If Len(id) = 0 Or directory.Exists(id) Then Err.Raise vbObjectError + 210, , "Invalid or duplicate directory ID."
        directory.Add id, CStr(WS("Team Directory").Cells(i, 2).Value2)
    Next i
    For i = 5 To 9
        id = Normal(WS("Control").Cells(i, 1).Value2)
        flag = Normal(WS("Control").Cells(i, 3).Value2)
        If flag <> "YES" And flag <> "NO" Then Err.Raise vbObjectError + 211, , "Availability must be YES or NO."
        If Not directory.Exists(id) Then Err.Raise vbObjectError + 212, , "Unknown analyst ID."
        If flag = "YES" Then
            If people.Exists(id) Then Err.Raise vbObjectError + 213, , "Duplicate analyst ID."
            If Not IsNumeric(WS("Control").Cells(i, 4).Value2) Then Err.Raise vbObjectError + 214, , "Capacity weight must be numeric."
            If CDbl(WS("Control").Cells(i, 4).Value2) <= 0 Then Err.Raise vbObjectError + 215, , "Capacity weight must be positive."
            people.Add id, directory(id): weights.Add id, CDbl(WS("Control").Cells(i, 4).Value2): loads.Add id, 0
        End If
    Next i
    If people.count = 0 Then Err.Raise vbObjectError + 216, , "No available analysts."
    For i = 5 To 6
        id = Normal(WS("Control").Cells(i, 7).Value2)
        If Len(id) > 0 Then locations(id) = True
    Next i
    If locations.count = 0 Then Err.Raise vbObjectError + 217, , "No allowed locations configured."
    CheckHeaders WS("Media Reference"), Array("Location", "Client ID", "Publication ID", "Workstream")
    For i = 2 To LastRow(WS("Media Reference"))
        key = KeyPart(WS("Media Reference").Cells(i, 1).Value2) & KeyPart(WS("Media Reference").Cells(i, 2).Value2) & KeyPart(WS("Media Reference").Cells(i, 3).Value2)
        If exact.Exists(key) Then Err.Raise vbObjectError + 218, , "Duplicate exact media mapping."
        flag = Normal(WS("Media Reference").Cells(i, 4).Value2)
        If flag <> "DIGITAL" And flag <> "PRINT" Then Err.Raise vbObjectError + 219, , "Invalid exact workstream."
        exact.Add key, flag
    Next i
    CheckHeaders WS("Publication Fallback"), Array("Publication ID", "Workstream")
    For i = 2 To LastRow(WS("Publication Fallback"))
        key = Normal(WS("Publication Fallback").Cells(i, 1).Value2)
        flag = Normal(WS("Publication Fallback").Cells(i, 2).Value2)
        If fallback.Exists(key) Or (flag <> "DIGITAL" And flag <> "PRINT") Then Err.Raise vbObjectError + 220, , "Invalid fallback mapping."
        fallback.Add key, flag
    Next i
End Sub

Private Sub AllocateGroups(ByVal groups As Object, ByVal votes As Object, ByVal rows As Object, ByVal people As Object, ByVal weights As Object, ByVal loads As Object)
    If groups.count = 0 Then Exit Sub
    Dim keys As Variant, i As Long, j As Long, temp As Variant, key As String
    keys = groups.Keys
    For i = LBound(keys) To UBound(keys) - 1
        For j = i + 1 To UBound(keys)
            If groups(keys(j)).count > groups(keys(i)).count Or (groups(keys(j)).count = groups(keys(i)).count And CStr(keys(j)) < CStr(keys(i))) Then
                temp = keys(i): keys(i) = keys(j): keys(j) = temp
            End If
        Next j
    Next i
    Dim id As Variant, owner As String, basis As String, score As Double, best As Double, count As Long, bestVotes As Long
    Dim index As Variant, row As Variant, candidateVotes As Long
    For i = LBound(keys) To UBound(keys)
        key = CStr(keys(i)): owner = "": best = 1E+30: bestVotes = -1
        count = groups(key).count: basis = "Weighted balanced allocation"
        If votes.Exists(key) Then
            For Each id In votes(key).Keys
                candidateVotes = CLng(votes(key)(id))
                score = CDbl(loads(id)) / CDbl(weights(id))
                If candidateVotes > bestVotes Or (candidateVotes = bestVotes And score < best) Or (candidateVotes = bestVotes And Abs(score - best) < 0.0000001 And CStr(id) < owner) Then
                    owner = CStr(id): best = score: bestVotes = candidateVotes
                End If
            Next id
            basis = "Repeated untouched - previous owner"
        End If
        If Len(owner) = 0 Then
            For Each id In people.Keys
                score = (CDbl(loads(id)) + count) / CDbl(weights(id))
                If score < best - 0.0000001 Then
                    owner = CStr(id): best = score
                ElseIf Abs(score - best) < 0.0000001 Then
                    If CLng(loads(id)) < CLng(loads(owner)) Or (CLng(loads(id)) = CLng(loads(owner)) And CStr(id) < owner) Then owner = CStr(id)
                End If
            Next id
        End If
        loads(owner) = CLng(loads(owner)) + count
        For Each index In groups(key)
            row = rows(CStr(index))
            row(COL_OWNER) = owner: row(COL_BASIS) = basis
            rows(CStr(index)) = row
        Next index
    Next i
End Sub

Private Sub BuildSummary(ByVal people As Object, ByVal weights As Object, ByVal loads As Object, ByVal assigned As Collection, ByVal excluded As Long, ByVal runId As String)
    Dim sheet As Worksheet, id As Variant, row As Variant, i As Long, totalWeight As Double, fresh As Long, carry As Long
    Set sheet = WS("Assignment Summary")
    sheet.Range("A2:K100").ClearContents
    For Each id In people.Keys
        totalWeight = totalWeight + CDbl(weights(id))
    Next id
    i = 2
    For Each id In people.Keys
        fresh = 0: carry = 0
        For Each row In assigned
            If CStr(row(COL_OWNER)) = CStr(id) Then
                If CStr(row(COL_VOLUME)) = "NEW VOLUME" Then fresh = fresh + 1 Else carry = carry + 1
            End If
        Next row
        sheet.Cells(i, 1).Resize(1, 8).Value2 = Array(id, people(id), weights(id), fresh, carry, loads(id), assigned.count * CDbl(weights(id)) / totalWeight, CDbl(loads(id)) - assigned.count * CDbl(weights(id)) / totalWeight)
        i = i + 1
    Next id
    sheet.Cells(i, 1).Value2 = "TOTAL"
    sheet.Cells(i, 6).Value2 = assigned.count
    sheet.Range("J2").Value2 = "Run ID": sheet.Range("K2").Value2 = runId
    sheet.Range("J3").Value2 = "Exceptions": sheet.Range("K3").Value2 = excluded
End Sub
