Sub DSKY_Inputs_Only()

    Dim dictWB As Workbook
    Dim verbWS As Worksheet
    Dim nounWS As Worksheet
    Dim progWS As Worksheet
    Dim wsR1DAP As Worksheet
    Dim wsR2DAP As Worksheet

    Dim ws As Worksheet
    Dim cell As Range
    Dim txt As String

    Dim re As Object
    Dim matches As Object

    Dim verbCode As String
    Dim nounCode As String
    Dim progCode As String

    Dim verbDesc As String
    Dim nounDesc As String
    Dim nounScale As String
    Dim progDesc As String

    Dim found As Range
    Dim lastRow As Long

    Dim codeText As String
    Dim descOnly As String
    Dim beforeText As String
    Dim afterText As String

    Dim logText As String
    Dim f As Integer

    '----------------------------------------------------
    ' Open dictionary workbook
    '----------------------------------------------------
    On Error Resume Next
    Set dictWB = Workbooks("CSM Verb and noun list.xlsx")
    On Error GoTo 0

    If dictWB Is Nothing Then
        MsgBox "Open 'CSM Verb and noun list.xlsx' first.", vbCritical
        Exit Sub
    End If

    Set verbWS = dictWB.Worksheets("Verb")
    Set nounWS = dictWB.Worksheets("Noun")
    Set progWS = dictWB.Worksheets("Program")
    Set wsR1DAP = dictWB.Worksheets("R1 DAP")
    Set wsR2DAP = dictWB.Worksheets("R2 DAP")

    '----------------------------------------------------
    ' Regex object
    '----------------------------------------------------
    Set re = CreateObject("VBScript.RegExp")
    re.IgnoreCase = False
    re.Global = False

    logText = ""

    '----------------------------------------------------
    ' Scan Column I on all sheets
    '----------------------------------------------------
    For Each ws In ThisWorkbook.Worksheets
        For Each cell In ws.Range("I:I")

            If VarType(cell.Value) = vbString Then
                txt = Trim(cell.Value)

                '===========================================================
                ' STRICT DAP BLOCK — ONLY AT FIRST PREAMBLE
                '===========================================================
                If txt = "DSKY V 4 8 E" And _
                   Trim(ws.Cells(cell.Row + 1, "I").Value) = "DSKY V 2 4 E" Then

                    Dim r1Line As String, r2Line As String
                    Dim parts() As String
                    Dim d(1 To 5) As String
                    Dim i As Long
                    Dim fDAP As Range
                    Dim annotationDAP As String
                    Dim originalA As String

                    '--------------------------------------------------------
                    ' Annotate V48 (verb-only)
                    '--------------------------------------------------------
                    lastRow = verbWS.Cells(verbWS.Rows.Count, 2).End(xlUp).Row
                    Set found = verbWS.Range("B2:B" & lastRow).Find("48", LookIn:=xlValues, LookAt:=xlWhole)
                    If Not found Is Nothing Then
                        verbDesc = found.Offset(0, 1).Value
                        originalA = ws.Cells(cell.Row, "A").Value
                        ws.Cells(cell.Row, "A").Value = originalA & " - " & verbDesc
                        logText = logText & ws.Name & " | " & originalA & " -> " & ws.Cells(cell.Row, "A").Value & vbCrLf
                    End If

                    '--------------------------------------------------------
                    ' Annotate V24 (verb-only)
                    '--------------------------------------------------------
                    Set found = verbWS.Range("B2:B" & lastRow).Find("24", LookIn:=xlValues, LookAt:=xlWhole)
                    If Not found Is Nothing Then
                        verbDesc = found.Offset(0, 1).Value
                        originalA = ws.Cells(cell.Row + 1, "A").Value
                        ws.Cells(cell.Row + 1, "A").Value = originalA & " - " & verbDesc
                        logText = logText & ws.Name & " | " & originalA & " -> " & ws.Cells(cell.Row + 1, "A").Value & vbCrLf
                    End If

                    '--------------------------------------------------------
                    ' R1 DAP (strict)
                    '--------------------------------------------------------
                    r1Line = Trim(ws.Cells(cell.Row + 2, "I").Value)
                    parts = Split(Application.Trim(r1Line), " ")

                    ' Must be exactly 7 tokens: DSKY + 5 digits + E
                    If UBound(parts) <> 6 Then GoTo NextCell
                    If parts(0) <> "DSKY" Then GoTo NextCell
                    If parts(6) <> "E" Then GoTo NextCell

                    ' Must be numeric digits
                    For i = 1 To 5
                        If Not IsNumeric(parts(i)) Then GoTo NextCell
                        d(i) = parts(i)
                    Next i

                    annotationDAP = ""

                    For i = 1 To 5
                        Set fDAP = wsR1DAP.Columns("A").Find(d(i), LookIn:=xlValues, LookAt:=xlWhole)
                        If fDAP Is Nothing Then GoTo NextCell
                        annotationDAP = annotationDAP & fDAP.Offset(0, i).Value
                        If i < 5 Then annotationDAP = annotationDAP & " | "
                    Next i

                    originalA = ws.Cells(cell.Row + 2, "A").Value
                    ws.Cells(cell.Row + 2, "A").Value = originalA & " - " & annotationDAP
                    logText = logText & ws.Name & " | " & originalA & " -> " & ws.Cells(cell.Row + 2, "A").Value & vbCrLf

                    '--------------------------------------------------------
                    ' R2 DAP (strict)
                    '--------------------------------------------------------
                    r2Line = Trim(ws.Cells(cell.Row + 3, "I").Value)
                    parts = Split(Application.Trim(r2Line), " ")

                    If UBound(parts) <> 6 Then GoTo NextCell
                    If parts(0) <> "DSKY" Then GoTo NextCell
                    If parts(6) <> "E" Then GoTo NextCell

                    For i = 1 To 5
                        If parts(i) <> "0" And parts(i) <> "1" Then GoTo NextCell
                        d(i) = parts(i)
                    Next i

                    annotationDAP = ""

                    For i = 1 To 5
                        Set fDAP = wsR2DAP.Columns("A").Find(d(i), LookIn:=xlValues, LookAt:=xlWhole)
                        If fDAP Is Nothing Then GoTo NextCell
                        annotationDAP = annotationDAP & fDAP.Offset(0, i).Value
                        If i < 5 Then annotationDAP = annotationDAP & " | "
                    Next i

                    originalA = ws.Cells(cell.Row + 3, "A").Value
                    ws.Cells(cell.Row + 3, "A").Value = originalA & " - " & annotationDAP
                    logText = logText & ws.Name & " | " & originalA & " -> " & ws.Cells(cell.Row + 3, "A").Value & vbCrLf

                    GoTo NextCell
                End If

                '===========================================================
                ' EVERYTHING BELOW HERE = ORIGINAL BEHAVIOR
                '===========================================================

                ' PROGRAM ENTRY
                re.Pattern = "^DSKY\s+(\d)\s+(\d)\s+E$"
                If re.Test(txt) Then
                    Set matches = re.Execute(txt)
                    progCode = matches(0).SubMatches(0) & matches(0).SubMatches(1)

                    lastRow = progWS.Cells(progWS.Rows.Count, 1).End(xlUp).Row
                    Set found = progWS.Range("A2:A" & lastRow).Find(progCode, LookIn:=xlValues, LookAt:=xlWhole)
                    If found Is Nothing Then GoTo NextCell

                    progDesc = found.Offset(0, 1).Value

                    beforeText = ws.Cells(cell.Row, "A").Value
                    afterText = beforeText & " - " & progDesc

                    ws.Cells(cell.Row, "A").Value = afterText
                    logText = logText & ws.Name & " | " & beforeText & " -> " & afterText & vbCrLf

                    GoTo NextCell
                End If

                ' PROGRAM CALL
                re.Pattern = "^DSKY\s+V\s+3\s+7\s+E\s+(\d)\s+(\d)\s+E$"
                If re.Test(txt) Then
                    Set matches = re.Execute(txt)
                    progCode = matches(0).SubMatches(0) & matches(0).SubMatches(1)

                    lastRow = verbWS.Cells(verbWS.Rows.Count, 2).End(xlUp).Row
                    Set found = verbWS.Range("B2:B" & lastRow).Find("37", LookIn:=xlValues, LookAt:=xlWhole)
                    If found Is Nothing Then GoTo NextCell
                    verbDesc = found.Offset(0, 1).Value

                    lastRow = progWS.Cells(progWS.Rows.Count, 1).End(xlUp).Row
                    Set found = progWS.Range("A2:A" & lastRow).Find(progCode, LookIn:=xlValues, LookAt:=xlWhole)
                    If found Is Nothing Then GoTo NextCell
                    progDesc = found.Offset(0, 1).Value

                    codeText = "V37E " & progCode & "E"
                    descOnly = verbDesc & " - " & progDesc

                    beforeText = ws.Cells(cell.Row, "A").Value
                    afterText = beforeText & " - " & descOnly

                    ws.Cells(cell.Row, "A").Value = afterText
                    logText = logText & ws.Name & " | " & beforeText & " -> " & afterText & vbCrLf

                    GoTo NextCell
                End If

                ' VERB + NOUN
                re.Pattern = "^DSKY\s+V\s+(\d)\s+(\d)\s+N\s+(\d)\s+(\d)\s+E$"
                If re.Test(txt) Then
                    Set matches = re.Execute(txt)
                    verbCode = matches(0).SubMatches(0) & matches(0).SubMatches(1)
                    nounCode = matches(0).SubMatches(2) & matches(0).SubMatches(3)

                    lastRow = verbWS.Cells(verbWS.Rows.Count, 2).End(xlUp).Row
                    Set found = verbWS.Range("B2:B" & lastRow).Find(verbCode, LookIn:=xlValues, LookAt:=xlWhole)
                    If found Is Nothing Then GoTo NextCell
                    verbDesc = found.Offset(0, 1).Value

                    lastRow = nounWS.Cells(nounWS.Rows.Count, 1).End(xlUp).Row
                    Set found = nounWS.Range("A2:A" & lastRow).Find(nounCode, LookIn:=xlValues, LookAt:=xlWhole)
                    If found Is Nothing Then GoTo NextCell
                    nounDesc = found.Offset(0, 1).Value
                    nounScale = found.Offset(0, 2).Value

                    codeText = "V" & verbCode & "E N" & nounCode & "E"

                    If Trim(nounScale) = "" Then
                        descOnly = verbDesc & " - " & nounDesc
                    Else
                        descOnly = verbDesc & " - " & nounDesc & " (" & nounScale & ")"
                    End If

                    beforeText = ws.Cells(cell.Row, "A").Value
                    afterText = beforeText & " - " & descOnly

                    ws.Cells(cell.Row, "A").Value = afterText
                    logText = logText & ws.Name & " | " & beforeText & " -> " & afterText & vbCrLf

                    GoTo NextCell
                End If

                ' VERB ONLY
                re.Pattern = "^DSKY\s+V\s+(\d)\s+(\d)\s+E$"
                If re.Test(txt) Then
                    Set matches = re.Execute(txt)
                    verbCode = matches(0).SubMatches(0) & matches(0).SubMatches(1)

                    lastRow = verbWS.Cells(verbWS.Rows.Count, 2).End(xlUp).Row
                    Set found = verbWS.Range("B2:B" & lastRow).Find(verbCode, LookIn:=xlValues, LookAt:=xlWhole)
                    If found Is Nothing Then GoTo NextCell
                    verbDesc = found.Offset(0, 1).Value

                    codeText = "V" & verbCode & "E"
                    descOnly = verbDesc

                    beforeText = ws.Cells(cell.Row, "A").Value
                    afterText = beforeText & " - " & descOnly

                    ws.Cells(cell.Row, "A").Value = afterText
                    logText = logText & ws.Name & " | " & beforeText & " -> " & afterText & vbCrLf

                    GoTo NextCell
                End If

                ' NOUN ONLY
                re.Pattern = "^DSKY\s+N\s+(\d)\s+(\d)\s+E$"
                If re.Test(txt) Then
                    Set matches = re.Execute(txt)
                    nounCode = matches(0).SubMatches(0) & matches(0).SubMatches(1)

                    lastRow = nounWS.Cells(nounWS.Rows.Count, 1).End(xlUp).Row
                    Set found = nounWS.Range("A2:A" & lastRow).Find(nounCode, LookIn:=xlValues, LookAt:=xlWhole)
                    If found Is Nothing Then GoTo NextCell
                    nounDesc = found.Offset(0, 1).Value
                    nounScale = found.Offset(0, 2).Value

                    codeText = "N" & nounCode & "E"

                    If Trim(nounScale) = "" Then
                        descOnly = nounDesc
                    Else
                        descOnly = nounDesc & " (" & nounScale & ")"
                    End If

                    beforeText = ws.Cells(cell.Row, "A").Value
                    afterText = beforeText & " - " & descOnly

                    ws.Cells(cell.Row, "A").Value = afterText
                    logText = logText & ws.Name & " | " & beforeText & " -> " & afterText & vbCrLf

                    GoTo NextCell
                End If

                ' VERB MISSING E
                re.Pattern = "^DSKY\s+V\s+(\d)\s+(\d)$"
                If re.Test(txt) Then
                    Set matches = re.Execute(txt)
                    verbCode = matches(0).SubMatches(0) & matches(0).SubMatches(1)

                    lastRow = verbWS.Cells(verbWS.Rows.Count, 2).End(xlUp).Row
                    Set found = verbWS.Range("B2:B" & lastRow).Find(verbCode, LookIn:=xlValues, LookAt:=xlWhole)
                    If found Is Nothing Then GoTo NextCell
                    verbDesc = found.Offset(0, 1).Value

                    codeText = "V" & verbCode & "E"
                    descOnly = verbDesc

                    beforeText = ws.Cells(cell.Row, "A").Value
                    afterText = beforeText & " - " & descOnly

                    ws.Cells(cell.Row, "A").Value = afterText
                    logText = logText & ws.Name & " | " & beforeText & " -> " & afterText & vbCrLf

                    GoTo NextCell
                End If

            End If

NextCell:
        Next cell
    Next ws

    '----------------------------------------------------
    ' WRITE LOG TO DESKTOP
    '----------------------------------------------------
    f = FreeFile
    Open Environ("USERPROFILE") & "\Desktop\dsky_inputs_log.txt" For Output As #f
    Print #f, logText
    Close #f

    MsgBox "DSKY input processing complete. Log saved to Desktop.", vbInformation

End Sub
