Sub Annotate_AGS_ALL()

    Dim ws As Worksheet
    Dim wsSel As Worksheet
    Dim wsIO As Worksheet
    Dim wsOut As Worksheet
    Dim wsMisc As Worksheet   ' <--- NEW
    Dim lastRow As Long
    Dim r As Long
    Dim txt As String
    Dim parts() As String
    Dim addr As String
    Dim selectorDigits As String
    Dim selectorCode As String
    Dim annotation As String
    Dim originalText As String

    Dim logAll As String

    Set wsSel = Workbooks("LM Verb and noun list.xlsx").Sheets("AGS Selector")
    Set wsIO = Workbooks("LM Verb and noun list.xlsx").Sheets("DEDA IO")
    Set wsOut = Workbooks("LM Verb and noun list.xlsx").Sheets("DEDA Output")
    Set wsMisc = Workbooks("LM Verb and noun list.xlsx").Sheets("DEDA Misc")   ' <--- NEW

    logAll = ""

    For Each ws In ThisWorkbook.Worksheets

        If ws.Name = "AGS Selector" Or ws.Name = "DEDA IO" Or ws.Name = "DEDA Output" Or ws.Name = "DEDA Misc" Then GoTo NextSheet
        If ws.Columns.Count < 9 Then GoTo NextSheet

        lastRow = ws.Cells(ws.Rows.Count, "I").End(xlUp).Row

        For r = 1 To lastRow

            txt = Trim(ws.Cells(r, "I").Value)
            originalText = ws.Cells(r, "A").Value

            If txt = "" Then GoTo NextRow
            If InStr(1, txt, "DEDA C", vbTextCompare) = 0 Then GoTo NextRow

            parts = Split(txt, " ")
            If UBound(parts) < 4 Then GoTo NextRow

            addr = parts(2) & parts(3) & parts(4)
            annotation = ""

            ' ---------------------------------------------------------
            ' SELECTOR LOGIC
            ' ---------------------------------------------------------
            If InStr(1, txt, "+") > 0 Then

                Dim p As Long, j As Long, ch As String
                selectorDigits = ""
                p = InStr(1, txt, "+")

                For j = p + 1 To Len(txt)
                    ch = Mid(txt, j, 1)
                    If ch Like "[0-9]" Then
                        selectorDigits = selectorDigits & ch
                    ElseIf ch = " " Then
                    Else
                        Exit For
                    End If
                Next j

                selectorCode = "+" & selectorDigits

                Dim rowSel As Long, lastSelRow As Long
                lastSelRow = wsSel.Cells(wsSel.Rows.Count, "A").End(xlUp).Row

                For rowSel = 1 To lastSelRow
                    If wsSel.Cells(rowSel, "A").Value = CLng(addr) Then
                        If Trim(wsSel.Cells(rowSel, "B").Value) = selectorCode Then
                            annotation = wsSel.Cells(rowSel, "C").Value
                            Exit For
                        End If
                    End If
                Next rowSel

                If annotation <> "" Then
                    ws.Cells(r, "A").Value = originalText & " - " & annotation
                    logAll = logAll & ws.Name & " | " & originalText & _
                             " -> " & ws.Cells(r, "A").Value & vbCrLf
                    GoTo NextRow
                End If

            End If

            ' ---------------------------------------------------------
            ' IO LOOKUP
            ' ---------------------------------------------------------
            Dim f As Range

            Set f = wsIO.Columns("A").Find(addr, LookIn:=xlValues, LookAt:=xlWhole)
            If f Is Nothing Then
                Set f = wsIO.Columns("A").Find(CLng(addr), LookIn:=xlValues, LookAt:=xlWhole)
            End If

            If Not f Is Nothing Then
                annotation = f.Offset(0, 1).Value
                Dim units As String: units = f.Offset(0, 2).Value
                If units <> "" Then annotation = annotation & " (" & units & ")"

                ws.Cells(r, "A").Value = originalText & " - " & annotation
                logAll = logAll & ws.Name & " | " & originalText & _
                         " -> " & ws.Cells(r, "A").Value & vbCrLf
                GoTo NextRow
            End If

            ' ---------------------------------------------------------
            ' OUTPUT LOOKUP
            ' ---------------------------------------------------------
            Set f = wsOut.Columns("A").Find(addr, LookIn:=xlValues, LookAt:=xlWhole)
            If f Is Nothing Then
                Set f = wsOut.Columns("A").Find(CLng(addr), LookIn:=xlValues, LookAt:=xlWhole)
            End If

            If Not f Is Nothing Then
                annotation = f.Offset(0, 1).Value
                Dim units2 As String: units2 = f.Offset(0, 2).Value
                If units2 <> "" Then annotation = annotation & " (" & units2 & ")"

                ws.Cells(r, "A").Value = originalText & " - " & annotation
                logAll = logAll & ws.Name & " | " & originalText & _
                         " -> " & ws.Cells(r, "A").Value & vbCrLf
                GoTo NextRow
            End If

            ' ---------------------------------------------------------
            ' NEW: DEDA MISC LOOKUP
            ' ---------------------------------------------------------
            Set f = wsMisc.Columns("A").Find(addr, LookIn:=xlValues, LookAt:=xlWhole)
            If f Is Nothing Then
                Set f = wsMisc.Columns("A").Find(CLng(addr), LookIn:=xlValues, LookAt:=xlWhole)
            End If

            If Not f Is Nothing Then
                annotation = f.Offset(0, 1).Value

                ws.Cells(r, "A").Value = originalText & " - " & annotation
                logAll = logAll & ws.Name & " | " & originalText & _
                         " -> " & ws.Cells(r, "A").Value & vbCrLf
                GoTo NextRow
            End If

            ' ---------------------------------------------------------
            ' UNMATCHED
            ' ---------------------------------------------------------
            logAll = logAll & ws.Name & " | [UNMATCHED] " & originalText & _
                     " was not found in AGS Selector, IO, Output, or Misc" & vbCrLf

NextRow:
        Next r

NextSheet:
    Next ws

    ' ---------------------------------------------------------
    ' WRITE UNIFIED LOG FILE
    ' ---------------------------------------------------------
    Dim ff As Integer
    ff = FreeFile
    Open Environ("USERPROFILE") & "\Desktop\deda_log_ags_all.txt" For Output As #ff
    Print #ff, logAll
    Close #ff

    MsgBox "AGS ALL annotation complete.", vbInformation

End Sub
