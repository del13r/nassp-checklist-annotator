Sub DSKY_Outputs_Only()

    Dim dictWB As Workbook
    Dim verbWS As Worksheet
    Dim nounWS As Worksheet

    Dim ws As Worksheet
    Dim cell As Range
    Dim txt As String

    Dim re As Object
    Dim matches As Object

    Dim verbCode As String
    Dim nounCode As String

    Dim verbDesc As String
    Dim nounDesc As String
    Dim nounScale As String

    Dim found As Range
    Dim lastRow As Long

    Dim descOnly As String
    Dim beforeText As String
    Dim afterText As String

    Dim logText As String
    Dim f As Integer

    On Error Resume Next
    Set dictWB = Workbooks("LM Verb and noun list.xlsx")
    On Error GoTo 0

    If dictWB Is Nothing Then
        MsgBox "Open 'LM Verb and noun list.xlsx' first.", vbCritical
        Exit Sub
    End If

    Set verbWS = dictWB.Worksheets("Verb")
    Set nounWS = dictWB.Worksheets("Noun")

    Set re = CreateObject("VBScript.RegExp")
    re.IgnoreCase = False
    re.Global = False

    logText = ""

    For Each ws In ThisWorkbook.Worksheets
        For Each cell In ws.Range("A:A")

            If VarType(cell.Value) <> vbString Then GoTo NextCell

            txt = cell.Value

            If Trim(ws.Cells(cell.Row, "I").Value) <> "" Then GoTo NextCell
            If Trim(txt) = "" Then GoTo NextCell

            If InStr(1, txt, "F 06 33 (LAT,LONG,ALT)", vbTextCompare) > 0 Then GoTo NextCell

            ' FLASHING verb+noun
            re.Pattern = "^[ ]{0,2}F[ ]+(\d{2})[ ]+(\d{2})\b"
            If re.Test(txt) Then
                Set matches = re.Execute(txt)
                verbCode = matches(0).SubMatches(0)
                nounCode = matches(0).SubMatches(1)
                GoTo HardFilter
            End If

            ' FLASHING verb only
            re.Pattern = "^[ ]{0,2}F[ ]+(\d{2})\b"
            If re.Test(txt) Then
                Set matches = re.Execute(txt)
                verbCode = matches(0).SubMatches(0)
                nounCode = ""
                GoTo HardFilter
            End If

            ' STEADY verb+noun
            re.Pattern = "^[ ]{0,2}(\d{2})[ ]+(\d{2})\b"
            If re.Test(txt) Then
                Set matches = re.Execute(txt)
                verbCode = matches(0).SubMatches(0)
                nounCode = matches(0).SubMatches(1)
                GoTo HardFilter
            End If

            ' STEADY verb only
            re.Pattern = "^[ ]{0,2}(\d{2})\b"
            If re.Test(txt) Then
                Set matches = re.Execute(txt)
                verbCode = matches(0).SubMatches(0)
                nounCode = ""
                GoTo HardFilter
            End If

            GoTo NextCell

HardFilter:

            Dim pos As Long
            pos = InStr(txt, verbCode)

            If pos > 0 Then
                Dim nextChar As String
                nextChar = Mid$(txt, pos + Len(verbCode), 1)

                If nextChar <> "" And nextChar <> " " Then
                    GoTo NextCell
                End If
            End If

            '===========================================================
            ' SECOND HARD FILTER: reject alphabetic tokens (deg, ft, etc.)
            '===========================================================
            Dim remainder As String
            remainder = Trim(Mid$(txt, pos + Len(verbCode) + 1))

            If remainder <> "" Then
                Dim firstToken As String
                firstToken = Split(remainder, " ")(0)

                If Not IsNumeric(firstToken) Then
                    GoTo NextCell
                End If
            End If

            '===========================================================
            ' SKIP LM SELF-TEST PATTERNS (88, F 88, 88 88, F 88 88)
            '===========================================================
            If verbCode = "88" Then GoTo NextCell
            If nounCode = "88" Then GoTo NextCell

ProcessCodes:

            lastRow = verbWS.Cells(verbWS.Rows.Count, 2).End(xlUp).Row
            Set found = verbWS.Range("B2:B" & lastRow).Find(verbCode, LookIn:=xlValues, LookAt:=xlWhole)
            If Not found Is Nothing Then verbDesc = found.Offset(0, 1).Value Else verbDesc = "(verb not found)"

            If nounCode <> "" Then
                lastRow = nounWS.Cells(nounWS.Rows.Count, 1).End(xlUp).Row
                Set found = nounWS.Range("A2:A" & lastRow).Find(nounCode, LookIn:=xlValues, LookAt:=xlWhole)
                If Not found Is Nothing Then
                    nounDesc = found.Offset(0, 1).Value
                    nounScale = found.Offset(0, 2).Value
                Else
                    nounDesc = "(noun not found)"
                    nounScale = ""
                End If
            Else
                nounDesc = ""
                nounScale = ""
            End If

            If nounDesc = "" Then
                descOnly = verbDesc
            Else
                If Trim(nounScale) = "" Then
                    descOnly = verbDesc & " - " & nounDesc
                Else
                    descOnly = verbDesc & " - " & nounDesc & " (" & nounScale & ")"
                End If
            End If

            beforeText = txt
            afterText = beforeText & " - " & descOnly

            cell.Value = afterText

            logText = logText & ws.Name & " | " & beforeText & " -> " & afterText & vbCrLf

NextCell:
        Next cell
    Next ws

    f = FreeFile
    Open Environ("USERPROFILE") & "\Desktop\LM_dsky_outputs_log.txt" For Output As #f
    Print #f, logText
    Close #f

    MsgBox "DSKY output processing complete. Log saved to Desktop.", vbInformation

End Sub
