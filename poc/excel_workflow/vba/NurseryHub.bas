Attribute VB_Name = "NurseryHub"
'==============================================================================
'  Pacific Seeds — Nursery Hub VBA
'  ----------------------------------------------------------------------------
'  Aggregates every registered nursery into a single dashboard by reading the
'  shared registry.csv (written by each nursery workbook on Push-to-Hub).
'==============================================================================
Option Explicit

Public Const HUB_SHEET_HOME      As String = "Home"
Public Const HUB_SHEET_DASHBOARD As String = "Dashboard"
Public Const HUB_SHEET_SETTINGS  As String = "Settings"

Private Function GetSheet(ByVal name As String) As Worksheet
    On Error Resume Next
    Set GetSheet = ThisWorkbook.Sheets(name)
    On Error GoTo 0
End Function

Private Function HubSetting(ByVal key As String) As String
    Dim ws As Worksheet, r As Long
    Set ws = GetSheet(HUB_SHEET_SETTINGS)
    If ws Is Nothing Then Exit Function
    For r = 6 To 30
        If Trim$(CStr(ws.Cells(r, 1).Value)) = key Then
            HubSetting = CStr(ws.Cells(r, 2).Value)
            Exit Function
        End If
    Next r
End Function

Private Function ExpandTilde(ByVal p As String) As String
    If Len(p) > 0 And Left$(p, 1) = "~" Then
        Dim home As String
        home = Environ$("HOME")
        If Len(home) = 0 Then home = Environ$("USERPROFILE")
        ExpandTilde = home & Mid$(p, 2)
    Else
        ExpandTilde = p
    End If
End Function

Private Function RegistryPath() As String
    Dim folder As String, fileName As String
    folder = ExpandTilde(HubSetting("Registry folder"))
    fileName = HubSetting("Registry file")
    If Len(fileName) = 0 Then fileName = "registry.csv"
    If Len(folder) = 0 Then Exit Function
    If Right$(folder, 1) <> Application.PathSeparator Then folder = folder & Application.PathSeparator
    RegistryPath = folder & fileName
End Function

Public Sub hubRefreshDashboard()
    Dim path As String
    path = RegistryPath()
    If Len(path) = 0 Or Dir(path) = "" Then
        MsgBox "Registry not found at:" & vbCrLf & path & vbCrLf & vbCrLf & _
               "Either no nurseries have synced yet, or the path is wrong. " & _
               "Check the Settings tab.", vbExclamation, "No registry"
        Exit Sub
    End If

    Dim dash As Worksheet: Set dash = GetSheet(HUB_SHEET_DASHBOARD)
    If dash Is Nothing Then Exit Sub
    dash.Range("A6:Z" & dash.Rows.Count).Clear

    Dim ff As Integer, line As String, firstLine As Boolean
    firstLine = True
    Dim outRow As Long: outRow = 6
    Dim nCount As Long, totalP As Long, totalR As Long, totalE As Long

    ff = FreeFile
    Open path For Input As #ff
    Do While Not EOF(ff)
        Line Input #ff, line
        If firstLine Then
            firstLine = False
        Else
            If Len(Trim$(line)) > 0 Then
                Dim parts() As String
                parts = ParseCsvLine(line)
                Dim i As Long
                For i = LBound(parts) To UBound(parts)
                    dash.Cells(outRow, i + 1).Value = parts(i)
                Next i
                ' Border row
                Dim used As Range
                Set used = dash.Range(dash.Cells(outRow, 1), dash.Cells(outRow, 10))
                used.Borders.LineStyle = xlContinuous
                used.Borders.Color = RGB(216, 227, 237)
                ' Hyperlink the file path (col 8)
                Dim fp As String: fp = ""
                If UBound(parts) >= 7 Then fp = parts(7)
                If Len(fp) > 0 And Dir(fp) <> "" Then
                    dash.Hyperlinks.Add Anchor:=dash.Cells(outRow, 8), Address:=fp, _
                                        TextToDisplay:=fp
                End If
                ' Tally
                nCount = nCount + 1
                If UBound(parts) >= 3 Then totalP = totalP + CLng(Val(parts(3)))
                If UBound(parts) >= 4 Then totalR = totalR + CLng(Val(parts(4)))
                If UBound(parts) >= 5 Then totalE = totalE + CLng(Val(parts(5)))
                outRow = outRow + 1
            End If
        End If
    Loop
    Close #ff

    On Error Resume Next
    ThisWorkbook.Names("HUB_NurseryCount").RefersToRange.Value = nCount
    ThisWorkbook.Names("HUB_TotalPackets").RefersToRange.Value = totalP
    ThisWorkbook.Names("HUB_TotalReplacements").RefersToRange.Value = totalR
    ThisWorkbook.Names("HUB_TotalErrors").RefersToRange.Value = totalE
    On Error GoTo 0

    dash.Activate
    MsgBox nCount & " nursery(ies) loaded into Dashboard.", vbInformation, "Refreshed"
End Sub

Public Sub hubRegisterFolder()
    Dim folder As String
    folder = InputBox("Shared folder (full path) where nursery summaries are written:", _
                      "Registry folder", HubSetting("Registry folder"))
    If Len(folder) = 0 Then Exit Sub
    Dim ws As Worksheet: Set ws = GetSheet(HUB_SHEET_SETTINGS)
    Dim r As Long
    For r = 6 To 30
        If Trim$(CStr(ws.Cells(r, 1).Value)) = "Registry folder" Then
            ws.Cells(r, 2).Value = folder
            Exit For
        End If
    Next r
    MsgBox "Registry folder set to:" & vbCrLf & folder, vbInformation, "Saved"
End Sub

Public Sub hubOpenFolder()
    Dim folder As String
    folder = ExpandTilde(HubSetting("Registry folder"))
    If Len(folder) = 0 Then
        MsgBox "Set the Registry folder on the Settings tab first.", vbExclamation
        Exit Sub
    End If
    On Error Resume Next
    #If Mac Then
        MacScript "tell application ""Finder"" to open POSIX file """ & folder & """"
    #Else
        Shell "explorer """ & folder & """", vbNormalFocus
    #End If
    On Error GoTo 0
End Sub

Public Sub hubCreateFromTemplate()
    Dim tpl As String
    tpl = ExpandTilde(HubSetting("Template path"))
    If Len(tpl) = 0 Or Dir(tpl) = "" Then
        MsgBox "Template not found at:" & vbCrLf & tpl & vbCrLf & vbCrLf & _
               "Update the 'Template path' on the Settings tab.", vbExclamation
        Exit Sub
    End If
    Dim code As String
    code = InputBox("New nursery code (e.g. AUGT1-26S-IMI):", "New Nursery")
    If Len(code) = 0 Then Exit Sub

    Dim folder As String
    folder = ExpandTilde(HubSetting("Registry folder"))
    If Right$(folder, 1) <> Application.PathSeparator Then folder = folder & Application.PathSeparator
    Dim dst As String
    dst = folder & code & ".xlsm"
    If Dir(dst) <> "" Then
        MsgBox "A workbook already exists for that code:" & vbCrLf & dst, vbExclamation
        Exit Sub
    End If
    FileCopy tpl, dst
    Workbooks.Open dst
    MsgBox "Created new nursery workbook:" & vbCrLf & dst & vbCrLf & vbCrLf & _
           "Run Step 1 in the new workbook to initialise.", vbInformation, "Created"
End Sub

' Naive CSV line parser — handles quoted fields with embedded commas.
Private Function ParseCsvLine(ByVal s As String) As String()
    Dim out() As String
    ReDim out(0 To 100)
    Dim n As Long: n = 0
    Dim i As Long, ch As String, cur As String, inq As Boolean
    For i = 1 To Len(s)
        ch = Mid$(s, i, 1)
        If inq Then
            If ch = """" Then
                If i < Len(s) And Mid$(s, i + 1, 1) = """" Then
                    cur = cur & """": i = i + 1
                Else
                    inq = False
                End If
            Else
                cur = cur & ch
            End If
        Else
            If ch = """" Then
                inq = True
            ElseIf ch = "," Then
                out(n) = cur: n = n + 1: cur = ""
            Else
                cur = cur & ch
            End If
        End If
    Next i
    out(n) = cur
    ReDim Preserve out(0 To n)
    ParseCsvLine = out
End Function

Public Sub hubRunMacroByCell(ByVal target As Range)
    ' Hub doesn't store macro names in column G; routes by cell text instead.
    ' Use Cells(1,1) so merged-range targets resolve to a single cell value.
    Dim t As String
    t = LCase$(Trim$(CStr(target.Cells(1, 1).Value)))
    Select Case True
        Case InStr(t, "refresh") > 0:               hubRefreshDashboard
        Case InStr(t, "register") > 0:              hubRegisterFolder
        Case InStr(t, "open") > 0 And InStr(t, "folder") > 0: hubOpenFolder
        Case InStr(t, "new nursery") > 0 Or InStr(t, "template") > 0: hubCreateFromTemplate
    End Select
End Sub
