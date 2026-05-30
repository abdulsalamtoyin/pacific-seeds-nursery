Attribute VB_Name = "PSBootstrap"
'==============================================================================
'  Pacific Seeds — VBA self-installer (bootstrap)
'  ----------------------------------------------------------------------------
'  Paste this whole module into Nursery_Template.xlsx OR Nursery_Hub.xlsx
'  via Tools > Macros > Visual Basic Editor > Insert > Module.
'  Then put your cursor anywhere in InstallAll and press F5.
'
'  It will:
'    1. ask you to pick the 'vba/' folder once,
'    2. inject all the right .bas/.cls files into THIS workbook,
'    3. delete itself so it isn't left lying around,
'    4. tell you to Save As .xlsm.
'==============================================================================
Option Explicit

' VBE component types
Private Const vbext_ct_StdModule  As Long = 1
Private Const vbext_ct_ClassModule As Long = 2
Private Const vbext_ct_Document    As Long = 100

Public Sub InstallAll()
    Dim folder As String
    folder = PickFolder("Pick the 'vba/' folder containing NurseryTemplate.bas etc.")
    If Len(folder) = 0 Then Exit Sub
    If Right$(folder, 1) <> Application.PathSeparator Then _
        folder = folder & Application.PathSeparator

    Dim wbName As String: wbName = LCase$(ThisWorkbook.Name)
    Dim okCount As Long, errCount As Long

    ' Clean up orphans from any prior half-failed install first.
    Dim removed As Long: removed = RemoveOrphanModules()
    If removed > 0 Then Debug.Print "Removed " & removed & " orphan std module(s)."

    If InStr(wbName, "template") > 0 Then
        ImportStdModule folder & "NurseryTemplate.bas", "NurseryTemplate", okCount, errCount
        ReplaceComponent "ThisWorkbook", folder & "ThisWorkbook_Template.cls", okCount, errCount
        ReplaceSheetByName "Home", folder & "Sheet_Home_Template.cls", okCount, errCount
    ElseIf InStr(wbName, "hub") > 0 Then
        ImportStdModule folder & "NurseryHub.bas", "NurseryHub", okCount, errCount
        ReplaceSheetByName "Home", folder & "Sheet_Home_Hub.cls", okCount, errCount
    Else
        MsgBox "This workbook's name doesn't match 'Template' or 'Hub'." & vbCrLf & _
               "Rename to include one of those, or edit InstallAll.", vbExclamation
        Exit Sub
    End If

    Dim msg As String
    msg = okCount & " module(s) installed."
    If errCount > 0 Then msg = msg & vbCrLf & errCount & " error(s) — see Immediate window (Ctrl+G)."
    msg = msg & vbCrLf & vbCrLf & "Next: " & vbCrLf & _
          "  1) File > Save As..." & vbCrLf & _
          "  2) Choose ""Excel Macro-Enabled Workbook (.xlsm)""" & vbCrLf & _
          "  3) Save next to the original .xlsx." & vbCrLf & vbCrLf & _
          "This bootstrap module will now delete itself."
    MsgBox msg, vbInformation, "Pacific Seeds — Install Complete"

    SelfDelete
End Sub

'------------------------------------------------------------------------------
Private Sub ImportStdModule(ByVal path As String, ByVal moduleName As String, _
                            ByRef okCount As Long, ByRef errCount As Long)
    ' Idempotent: remove any prior module of the same name, then Import the .bas.
    ' VBComponents.Import handles the file format including Attribute VB_Name.
    If Not FileExists(path) Then
        Debug.Print "MISSING: " & path: errCount = errCount + 1: Exit Sub
    End If
    Dim existing As Object
    On Error Resume Next
    Set existing = ThisWorkbook.VBProject.VBComponents(moduleName)
    On Error GoTo 0
    If Not existing Is Nothing Then
        ThisWorkbook.VBProject.VBComponents.Remove existing
        Debug.Print "Removed existing " & moduleName
    End If
    On Error GoTo importFail
    ThisWorkbook.VBProject.VBComponents.Import path
    On Error GoTo 0
    okCount = okCount + 1
    Debug.Print "Imported " & moduleName & " from " & path
    Exit Sub
importFail:
    Debug.Print "IMPORT FAILED: " & path & " — " & Err.Number & " " & Err.Description
    errCount = errCount + 1
End Sub

Private Sub ReplaceComponent(ByVal compName As String, ByVal clsPath As String, _
                             ByRef okCount As Long, ByRef errCount As Long)
    If Not FileExists(clsPath) Then
        Debug.Print "MISSING: " & clsPath: errCount = errCount + 1: Exit Sub
    End If
    Dim comp As Object
    On Error Resume Next
    Set comp = ThisWorkbook.VBProject.VBComponents(compName)
    On Error GoTo 0
    If comp Is Nothing Then
        Debug.Print "COMPONENT NOT FOUND: " & compName
        errCount = errCount + 1: Exit Sub
    End If
    Dim body As String: body = StripClsHeader(ReadAllText(clsPath))
    Dim cm As Object: Set cm = comp.CodeModule
    If cm.CountOfLines > 0 Then cm.DeleteLines 1, cm.CountOfLines
    cm.AddFromString body
    okCount = okCount + 1
    Debug.Print "Replaced " & compName & " from " & clsPath
End Sub

Private Sub ReplaceSheetByName(ByVal sheetName As String, ByVal clsPath As String, _
                               ByRef okCount As Long, ByRef errCount As Long)
    ' Find the VBComponent for the sheet whose tab name = sheetName
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(sheetName)
    On Error GoTo 0
    If ws Is Nothing Then
        Debug.Print "SHEET NOT FOUND: " & sheetName
        errCount = errCount + 1: Exit Sub
    End If
    ReplaceComponent ws.CodeName, clsPath, okCount, errCount
End Sub

'------------------------------------------------------------------------------
Private Function StripClsHeader(ByVal text As String) As String
    ' Strip VBA file-format preamble (VERSION, BEGIN/END, MultiUse, Attribute).
    ' Attribute lines can appear anywhere; remove them globally, not just at top.
    Dim lines() As String: lines = Split(text, vbLf)
    Dim out As String, started As Boolean, i As Long, s As String, t As String
    For i = LBound(lines) To UBound(lines)
        s = lines(i)
        If Right$(s, 1) = vbCr Then s = Left$(s, Len(s) - 1)
        t = LTrim$(s)
        ' Always drop Attribute lines, wherever they appear.
        If Left$(t, 10) = "Attribute " Then GoTo nxt
        If Not started Then
            If Left$(t, 8) = "VERSION " Then GoTo nxt
            If t = "BEGIN" Or t = "END" Then GoTo nxt
            If Left$(t, 9) = "MultiUse " Then GoTo nxt
            If Len(t) = 0 Then GoTo nxt
            started = True
        End If
        out = out & s & vbCrLf
nxt:
    Next i
    StripClsHeader = out
End Function

Private Function ReadAllText(ByVal path As String) As String
    Dim ff As Integer, txt As String, line As String, first As Boolean
    ff = FreeFile
    Open path For Input As #ff
    first = True
    Do While Not EOF(ff)
        Line Input #ff, line
        If first Then first = False Else txt = txt & vbCrLf
        txt = txt & line
    Loop
    Close #ff
    ReadAllText = txt
End Function

Private Function FileExists(ByVal p As String) As Boolean
    On Error Resume Next
    FileExists = (Dir(p) <> "")
End Function

Private Function PickFolder(ByVal prompt As String) As String
    On Error Resume Next
    Dim sel As Variant
    #If Mac Then
        ' macOS: AppleScript chooser
        Dim scpt As String
        scpt = "POSIX path of (choose folder with prompt """ & prompt & """)"
        sel = MacScript(scpt)
    #Else
        Dim fd As FileDialog
        Set fd = Application.FileDialog(msoFileDialogFolderPicker)
        fd.Title = prompt
        If fd.Show = -1 Then sel = fd.SelectedItems(1)
    #End If
    If Not IsEmpty(sel) Then PickFolder = CStr(sel)
End Function

Private Function RemoveOrphanModules() As Long
    ' Remove any std module that is NOT one of the canonical names. Anything
    ' else is either a half-failed prior install (Module1, Module2…) or a
    ' stale copy of the bootstrap.
    Const KEEP As String = "|NurseryTemplate|NurseryHub|PSBootstrap|"
    Dim vbp As Object, comp As Object
    Set vbp = ThisWorkbook.VBProject
    Dim toRemove As Collection: Set toRemove = New Collection
    Dim n As Long
    For Each comp In vbp.VBComponents
        If comp.Type = vbext_ct_StdModule Then
            If InStr(KEEP, "|" & comp.Name & "|") = 0 Then
                toRemove.Add comp
            Else
                ' Also remove canonical modules whose bodies contain a stray
                ' Attribute line (from a previous AddFromString that swallowed it).
                On Error Resume Next
                Dim body As String
                body = comp.CodeModule.Lines(1, comp.CodeModule.CountOfLines)
                On Error GoTo 0
                If InStr(body, "Attribute VB_") > 0 Then toRemove.Add comp
            End If
        End If
    Next comp
    Dim it As Variant
    For Each it In toRemove
        On Error Resume Next
        vbp.VBComponents.Remove it
        If Err.Number = 0 Then n = n + 1
        On Error GoTo 0
    Next it
    RemoveOrphanModules = n
End Function

Private Sub SelfDelete()
    On Error Resume Next
    Dim me_ As Object
    Set me_ = ThisWorkbook.VBProject.VBComponents("PSBootstrap")
    If Not me_ Is Nothing Then
        ThisWorkbook.VBProject.VBComponents.Remove me_
    End If
End Sub
