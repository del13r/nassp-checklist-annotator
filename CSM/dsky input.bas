Sub DSKY_Inputs_Only()

    Dim dictWB As Workbook
    Dim verbWS As Worksheet
    Dim nounWS As Worksheet
    Dim progWS As Worksheet

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
                ' CASE 1 — PROGRAM ENTRY: DSKY X X E   (FIXED)
                '===========================================================
                re.Pattern = "^DSKY\s+(\d)\s+(\d)\s+E$"

                If re.Test(txt) Then

                    Set matches = re.Execute(txt)
                    progCode = matches(0).SubMatches(0) & matches(0).SubMatches(1)

                    lastRow = progWS.Cells(progWS.Rows.Count, 1).End(xlUp).Row
                    Set found = progWS.Range("A2:A" & lastRow).Find(progCode, LookIn:=xlValues, LookAt:=xlWhole)
                    If Not found Is Nothing Then
                        progDesc = found.Offset(0, 1).Value
                    Else
                        progDesc = "(program not found)"
                    End If

                    descOnly = progDesc
                    codeText = progCode & "E"

                    beforeText = ws.Cells(cell.Row, "A").Value

                    If Trim(beforeText) = "" Then
                        afterText = codeText & " - " & descOnly
                    Else
                        afterText = beforeText & " - " & descOnly
                    End If

                    ws.Cells(cell.Row, "A").Value = afterText
                    logText = logText & ws.Name & " | " & beforeText & " -> " & afterText & vbCrLf

                    GoTo NextCell
                End If

                '===========================================================
                ' CASE 2 — PROGRAM CALL: DSKY V 3 7 E X X E
                '===========================================================
                re.Pattern = "^DSKY\s+V\s+3\s+7\s+E\s+(\d)\s+(\d)\s+E$"

                If re.Test(txt) Then

                    Set matches = re.Execute(txt)
                    progCode = matches(0).SubMatches(0) & matches(0).SubMatches(1)

                    lastRow = verbWS.Cells(verbWS.Rows.Count, 2).End(xlUp).Row
                    Set found = verbWS.Range("B2:B" & lastRow).Find("37", LookIn:=xlValues, LookAt:=xlWhole)
                    If Not found Is Nothing Then verbDesc = found.Offset(0, 1).Value Else verbDesc = "(verb 37 not found)"

                    lastRow = progWS.Cells(progWS.Rows.Count, 1).End(xlUp).Row
                    Set found = progWS.Range("A2:A" & lastRow).Find(progCode, LookIn:=xlValues, LookAt:=xlWhole)
                    If Not found Is Nothing Then progDesc = found.Offset(0, 1).Value Else progDesc = "(program not found)"

                    codeText = "V37E " & progCode & "E"
                    descOnly = verbDesc & " - " & progDesc

                    beforeText = ws.Cells(cell.Row, "A").Value

                    If Trim(beforeText) = "" Then
                        afterText = codeText & " - " & descOnly
                    Else
                        afterText = beforeText & " - " & descOnly
                    End If

                    ws.Cells(cell.Row, "A").Value = afterText
                    logText = logText & ws.Name & " | " & beforeText & " -> " & afterText & vbCrLf

                    GoTo NextCell
                End If

                '===========================================================
                ' CASE 3 — VERB + NOUN: DSKY V X X N X X E
                '===========================================================
                re.Pattern = "^DSKY\s+V\s+(\d)\s+(\d)\s+N\s+(\d)\s+(\d)\s+E$"

                If re.Test(txt) Then

                    Set matches = re.Execute(txt)
                    verbCode = matches(0).SubMatches(0) & matches(0).SubMatches(1)
                    nounCode = matches(0).SubMatches(2) & matches(0).SubMatches(3)

                    lastRow = verbWS.Cells(verbWS.Rows.Count, 2).End(xlUp).Row
                    Set found = verbWS.Range("B2:B" & lastRow).Find(verbCode, LookIn:=xlValues, LookAt:=xlWhole)
                    If Not found Is Nothing Then verbDesc = found.Offset(0, 1).Value Else verbDesc = "(verb not found)"

                    lastRow = nounWS.Cells(nounWS.Rows.Count, 1).End(xlUp).Row
                    Set found = nounWS.Range("A2:A" & lastRow).Find(nounCode, LookIn:=xlValues, LookAt:=xlWhole)
                    If Not found Is Nothing Then
                        nounDesc = found.Offset(0, 1).Value
                        nounScale = found.Offset(0, 2).Value
                    Else
                        nounDesc = "(noun not found)"
                        nounScale = ""
                    End If

                    codeText = "V" & verbCode & "E N" & nounCode & "E"

                    If Trim(nounScale) = "" Then
                        descOnly = verbDesc & " - " & nounDesc
                    Else
                        descOnly = verbDesc & " - " & nounDesc & " (" & nounScale & ")"
                    End If

                    beforeText = ws.Cells(cell.Row, "A").Value

                    If Trim(beforeText) = "" Then
                        afterText = codeText & " - " & descOnly
                    Else
                        afterText = beforeText & " - " & descOnly
                    End If

                    ws.Cells(cell.Row, "A").Value = afterText
                    logText = logText & ws.Name & " | " & beforeText & " -> " & afterText & vbCrLf

                    GoTo NextCell
                End If

                '===========================================================
                ' CASE 4 — VERB-ONLY: DSKY V X X E
                '===========================================================
                re.Pattern = "^DSKY\s+V\s+(\d)\s+(\d)\s+E$"

                If re.Test(txt) Then

                    Set matches = re.Execute(txt)
                    verbCode = matches(0).SubMatches(0) & matches(0).SubMatches(1)

                    lastRow = verbWS.Cells(verbWS.Rows.Count, 2).End(xlUp).Row
                    Set found = verbWS.Range("B2:B" & lastRow).Find(verbCode, LookIn:=xlValues, LookAt:=xlWhole)
                    If Not found Is Nothing Then verbDesc = found.Offset(0, 1).Value Else verbDesc = "(verb not found)"

                    codeText = "V" & verbCode & "E"
                    descOnly = verbDesc

                    beforeText = ws.Cells(cell.Row, "A").Value

                    If Trim(beforeText) = "" Then
                        afterText = codeText & " - " & descOnly
                    Else
                        afterText = beforeText & " - " & descOnly
                    End If

                    ws.Cells(cell.Row, "A").Value = afterText
                    logText = logText & ws.Name & " | " & beforeText & " -> " & afterText & vbCrLf

                    GoTo NextCell
                End If

                '===========================================================
                ' CASE 5 — NOUN-ONLY: DSKY N X X E
                '===========================================================
                re.Pattern = "^DSKY\s+N\s+(\d)\s+(\d)\s+E$"

                If re.Test(txt) Then

                    Set matches = re.Execute(txt)
                    nounCode = matches(0).SubMatches(0) & matches(0).SubMatches(1)

                    lastRow = nounWS.Cells(nounWS.Rows.Count, 1).End(xlUp).Row
                    Set found = nounWS.Range("A2:A" & lastRow).Find(nounCode, LookIn:=xlValues, LookAt:=xlWhole)
                    If Not found Is Nothing Then
                        nounDesc = found.Offset(0, 1).Value
                        nounScale = found.Offset(0, 2).Value
                    Else
                        nounDesc = "(noun not found)"
                        nounScale = ""
                    End If

                    codeText = "N" & nounCode & "E"

                    If Trim(nounScale) = "" Then
                        descOnly = nounDesc
                    Else
                        descOnly = nounDesc & " (" & nounScale & ")"
                    End If

                    beforeText = ws.Cells(cell.Row, "A").Value

                    If Trim(beforeText) = "" Then
                        afterText = codeText & " - " & descOnly
                    Else
                        afterText = beforeText & " - " & descOnly
                    End If

                    ws.Cells(cell.Row, "A").Value = afterText
                    logText = logText & ws.Name & " | " & beforeText & " -> " & afterText & vbCrLf

                    GoTo NextCell
                End If

                '===========================================================
                ' CASE 6 — VERB MISSING E: DSKY V X X
                '===========================================================
                re.Pattern = "^DSKY\s+V\s+(\d)\s+(\d)$"

                If re.Test(txt) Then

                    Set matches = re.Execute(txt)
                    verbCode = matches(0).SubMatches(0) & matches(0).SubMatches(1)

                    lastRow = verbWS.Cells(verbWS.Rows.Count, 2).End(xlUp).Row
                    Set found = verbWS.Range("B2:B" & lastRow).Find(verbCode, LookIn:=xlValues, LookAt:=xlWhole)
                    If Not found Is Nothing Then verbDesc = found.Offset(0, 1).Value Else verbDesc = "(verb not found)"

                    codeText = "V" & verbCode & "E"
                    descOnly = verbDesc

                    beforeText = ws.Cells(cell.Row, "A").Value

                    If Trim(beforeText) = "" Then
                        afterText = codeText & " - " & descOnly
                    Else
                        afterText = beforeText & " - " & descOnly
                    End If

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
