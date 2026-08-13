Attribute VB_Name = "NurseryTemplate"
'==============================================================================
'  Pacific Seeds — Nursery Workflow (consolidated VBA module)
'  Mirrors the FastAPI/PWA app's 17-tab workbook output exactly.
'
'  Workflow:
'    1. User pastes PRISM export into the "Nursery site" tab (data row 6+).
'    2. Click  ▶ 1 Initialise from PRISM export   → validates + stamps.
'    3. Click  ▶ 2 Generate ALL workbook tabs     → builds every output tab:
'         Map · Material Map · Packet Prep · Nursery list · Fieldbook ·
'         BC0 labels · Date recording · Pulling bags · TFMSA Spray plots ·
'         Hy Heights.
'    4. (Optional) Step 3 re-sorts Packet Prep into LSD-radix racking order.
'==============================================================================
Option Explicit

' Tab names live in SpecConstants.bas, generated from
' excel_workflow/spec/nursery_spec.json by gen_spec_constants.py.
' Do not re-declare them here — the generator is the single source of truth,
' and a local copy is exactly the drift the spec file exists to prevent.

' Colours for the four digit columns on Packet Prep
Public Const CLR_DIGIT_THOUSANDS     As Long = &H000000   ' Black
Public Const CLR_DIGIT_HUNDREDS      As Long = &H0000C0   ' Red
Public Const CLR_DIGIT_TENS          As Long = &H00B050   ' Green
Public Const CLR_DIGIT_ONES          As Long = &HCD7806   ' Blue

' Header style values
Public Const HDR_NAVY                As Long = &H402A09   ' #092A40 in BGR
Public Const HDR_PALE                As Long = &HFCF3E9   ' #E9F3FC in BGR

'==============================================================================
'  M_COMMON — utilities
'==============================================================================
Public Function GetSheet(ByVal name As String) As Worksheet
    On Error Resume Next
    Set GetSheet = ThisWorkbook.Sheets(name)
    On Error GoTo 0
End Function

Public Function FindHeaderCol(ByVal ws As Worksheet, ByVal name As String, _
                              Optional ByVal headerRow As Long = 5) As Long
    Dim lastCol As Long, c As Long
    lastCol = ws.Cells(headerRow, ws.Columns.Count).End(xlToLeft).Column
    For c = 1 To lastCol
        If LCase$(Trim$(CStr(ws.Cells(headerRow, c).Value))) = LCase$(name) Then
            FindHeaderCol = c
            Exit Function
        End If
    Next c
    FindHeaderCol = 0
End Function

Public Function GetSetting(ByVal key As String) As Variant
    Dim ws As Worksheet, r As Long
    Set ws = GetSheet(SHEET_SETTINGS)
    If ws Is Nothing Then Exit Function
    For r = 6 To 30
        If Trim$(CStr(ws.Cells(r, 1).Value)) = key Then
            GetSetting = ws.Cells(r, 2).Value
            Exit Function
        End If
    Next r
End Function

Public Sub SetSetting(ByVal key As String, ByVal value As Variant)
    Dim ws As Worksheet, r As Long
    Set ws = GetSheet(SHEET_SETTINGS)
    If ws Is Nothing Then Exit Sub
    For r = 6 To 30
        If Trim$(CStr(ws.Cells(r, 1).Value)) = key Then
            ws.Cells(r, 2).Value = value
            Exit Sub
        End If
    Next r
End Sub

Public Function ExpandTilde(ByVal p As String) As String
    If Len(p) > 0 And Left$(p, 1) = "~" Then
        Dim home As String
        home = Environ$("HOME")
        If Len(home) = 0 Then home = Environ$("USERPROFILE")
        ExpandTilde = home & Mid$(p, 2)
    Else
        ExpandTilde = p
    End If
End Function

Public Function NowISO() As String
    NowISO = Format$(Now, "yyyy-mm-dd hh:nn:ss")
End Function

Public Function CurrentTech() As String
    Dim t As String
    t = CStr(GetSetting("Default tech"))
    If Len(t) = 0 Then t = Environ$("USER")
    If Len(t) = 0 Then t = Environ$("USERNAME")
    CurrentTech = t
End Function

' Helper: turn nullable cell value into a string ("" if blank)
Private Function S(ByVal v As Variant) As String
    If IsError(v) Or IsNull(v) Or IsEmpty(v) Then
        S = ""
    Else
        S = Trim$(CStr(v))
    End If
End Function

' Helper: TRUE if `gen` starts with "BC" (case-insensitive)
Private Function IsBCGen(ByVal gen As String) As Boolean
    IsBCGen = (Len(gen) >= 2) And (UCase$(Left$(gen, 2)) = "BC")
End Function

' Helper: TRUE if `gen` is BC* or generic "Fn" (NOT F1/F2/F8 — those are selections)
Private Function IsRecurrent(ByVal gen As String) As Boolean
    IsRecurrent = IsBCGen(gen) Or (UCase$(gen) = "FN")
End Function

' Helper: TRUE if `gen` = "F1" (hybrid for height tracking)
Private Function IsHybridF1(ByVal gen As String) As Boolean
    IsHybridF1 = (UCase$(gen) = "F1")
End Function


'==============================================================================
'  M_BUTTONS — double-click router for the Home page
'==============================================================================
Public Sub RunMacroByCell(ByVal target As Range)
    Dim macroName As String
    macroName = Trim$(CStr(target.Worksheet.Cells(target.Row, 7).Value))
    If Len(macroName) = 0 Then Exit Sub
    On Error Resume Next
    Application.Run macroName
    If Err.Number <> 0 Then
        MsgBox "Could not run macro '" & macroName & "'." & vbCrLf & _
               Err.Description, vbExclamation, "Pacific Seeds Workflow"
    End If
    On Error GoTo 0
End Sub


'==============================================================================
'  Step 1 — Initialise nursery (validate + stamp)
'==============================================================================
Public Sub btnInitNursery()
    Dim code As String
    code = InputBox("Enter the nursery code (e.g. AUGT1-26S-IMI):", _
                    "Initialise Nursery", CStr(GetSetting("Nursery code")))
    If Len(code) = 0 Then Exit Sub
    SetSetting "Nursery code", code

    ' Nursery type drives which conditional tabs and extras get built, so it
    ' has to be captured before any tab generation runs.
    Dim types As String
    types = InputBox( _
        "Nursery type(s) — comma separated." & vbCrLf & _
        "Options: Selection, AB, Hybrid, Other", _
        "Nursery type", CStr(GetSetting("Nursery types")))
    If Len(Trim(types)) = 0 Then Exit Sub
    SetSetting "Nursery types", types

    Dim fName As String
    fName = InputBox("File name for this workbook:", "File name", code)
    If Len(Trim(fName)) > 0 Then SetSetting "File name", fName

    Dim ws As Worksheet: Set ws = GetSheet(SHEET_NURSERY_SITE)
    If ws Is Nothing Then
        MsgBox "Sheet '" & SHEET_NURSERY_SITE & "' not found.", vbCritical
        Exit Sub
    End If
    If S(ws.Cells(6, 1).Value) = "" Then
        MsgBox "Paste the PRISM export into the '" & SHEET_NURSERY_SITE & _
               "' tab starting at row 6 (headers stay at row 5), then re-run step 1.", _
               vbInformation, "Nursery site is empty"
        ws.Activate
        Exit Sub
    End If

    Dim nd As Worksheet: Set nd = GetSheet(SHEET_NURSERY_DATA)
    If Not nd Is Nothing Then
        nd.Cells(6, 1).Value = "Nursery code"
        nd.Cells(6, 2).Value = code
        nd.Cells(7, 1).Value = "Initialised at"
        nd.Cells(7, 2).Value = NowISO()
        nd.Cells(8, 1).Value = "Breeder"
        nd.Cells(8, 2).Value = CStr(GetSetting("Breeder"))
        nd.Cells(9, 1).Value = "Season"
        nd.Cells(9, 2).Value = CStr(GetSetting("Season"))
    End If

    MsgBox "Initialised nursery '" & code & "'." & vbCrLf & vbCrLf & _
           "Next: click Step 2 to generate every workbook tab.", _
           vbInformation, "Done"
End Sub


'==============================================================================
'  Step 2 — Generate ALL workbook tabs from PRISM data (the master sub)
'==============================================================================
Public Sub btnGenerateAllTabs()
    Dim src As Worksheet: Set src = GetSheet(SHEET_NURSERY_SITE)
    If src Is Nothing Or S(src.Cells(6, 1).Value) = "" Then
        MsgBox "Paste the PRISM export into '" & SHEET_NURSERY_SITE & _
               "' tab first (headers row 5, data row 6+).", vbExclamation
        Exit Sub
    End If

    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual

    ' Read PRISM headers once
    Dim col_range As Long, col_row As Long, col_mat As Long, col_inb As Long
    Dim col_src As Long, col_cms As Long, col_gen As Long, col_com As Long
    Dim col_ped As Long, col_hyb As Long, col_trait As Long, col_plant As Long
    Dim col_loc As Long, col_sub As Long, col_ebp As Long, col_ebn As Long, col_ent As Long
    col_range = FindHeaderCol(src, "Range", 5)
    col_row   = FindHeaderCol(src, "Row", 5)
    col_mat   = FindHeaderCol(src, "Material ID", 5)
    col_inb   = FindHeaderCol(src, "Inbred Code", 5)
    col_src   = FindHeaderCol(src, "Source ID", 5)
    col_cms   = FindHeaderCol(src, "CMS reaction", 5)
    col_gen   = FindHeaderCol(src, "Generation", 5)
    col_com   = FindHeaderCol(src, "Comments", 5)
    col_ped   = FindHeaderCol(src, "Pedigree", 5)
    col_hyb   = FindHeaderCol(src, "Hybrid Code", 5)
    col_trait = FindHeaderCol(src, "Trait Name", 5)
    col_plant = FindHeaderCol(src, "Plant #", 5)
    col_loc   = FindHeaderCol(src, "Loc Seq#", 5)
    col_sub   = FindHeaderCol(src, "SubSeq Flag", 5)
    col_ebp   = FindHeaderCol(src, "Entry Book Project", 5)
    col_ebn   = FindHeaderCol(src, "Entry Book Name", 5)
    col_ent   = FindHeaderCol(src, "Entry #", 5)
    If col_range = 0 Or col_row = 0 Then
        MsgBox "Could not find Range/Row columns in row 5 of Nursery site.", vbCritical
        Application.Calculation = xlCalculationAutomatic
        Exit Sub
    End If

    Dim lastRow As Long
    lastRow = src.Cells(src.Rows.Count, col_range).End(xlUp).Row
    Dim nRows As Long: nRows = lastRow - 5

    ' Pull all packets into a 2D array indexed by packet#
    Dim p() As Variant
    ReDim p(1 To nRows, 1 To 17)
    Dim r As Long, k As Long: k = 0
    For r = 6 To lastRow
        k = k + 1
        p(k, 1)  = CLng(Val(src.Cells(r, col_range).Value))    ' Range
        p(k, 2)  = CLng(Val(src.Cells(r, col_row).Value))      ' Row
        p(k, 3)  = IIf(col_mat > 0,   src.Cells(r, col_mat).Value, "")
        p(k, 4)  = IIf(col_inb > 0,   src.Cells(r, col_inb).Value, "")
        p(k, 5)  = IIf(col_src > 0,   src.Cells(r, col_src).Value, "")
        p(k, 6)  = IIf(col_cms > 0,   src.Cells(r, col_cms).Value, "")
        p(k, 7)  = IIf(col_gen > 0,   src.Cells(r, col_gen).Value, "")
        p(k, 8)  = IIf(col_com > 0,   src.Cells(r, col_com).Value, "")
        p(k, 9)  = IIf(col_ped > 0,   src.Cells(r, col_ped).Value, "")
        p(k, 10) = IIf(col_hyb > 0,   src.Cells(r, col_hyb).Value, "")
        p(k, 11) = IIf(col_trait > 0, src.Cells(r, col_trait).Value, "")
        p(k, 12) = IIf(col_plant > 0, src.Cells(r, col_plant).Value, "")
        p(k, 13) = IIf(col_loc > 0,   src.Cells(r, col_loc).Value, "")
        p(k, 14) = IIf(col_sub > 0,   src.Cells(r, col_sub).Value, "")
        p(k, 15) = IIf(col_ebp > 0,   src.Cells(r, col_ebp).Value, "")
        p(k, 16) = IIf(col_ebn > 0,   src.Cells(r, col_ebn).Value, "")
        p(k, 17) = IIf(col_ent > 0,   src.Cells(r, col_ent).Value, "")
    Next r

    ' --- Build each tab -------------------------------------------------
    ' Packet Prep fans out to one tab per planting date, set by the Field Map
    ' wizard. Without the wizard there is a single split.
    Dim nSplits As Long: nSplits = PlantingDateCount()
    Dim sp As Long
    For sp = 1 To nSplits
        BuildPacketPrep p, CStr(GetSetting("Nursery code")), sp
    Next sp

    ' Material Map carries the identifiers; Field Map is the bare grid the
    ' wizard writes planting information alongside.
    BuildMap p, SHEET_MATERIAL_MAP, Array(3, 4, 10)  ' Material / Inbred / Hybrid
    BuildMap p, SHEET_FIELD_MAP, Array()             ' geometry only, no values
    BuildNurseryList p

    ' Fieldbook is deliberately NOT built here — it comes from
    ' 'Updated nursery site' via btnBuildFieldbook, after PRISM is re-exported.

    If IsNurseryType("AB") Then
        BuildBC0Labels p, CStr(GetSetting("Nursery code"))
        BuildDateRecording p, SHEET_DATE_RECORDING
        BuildDateRecording p, SHEET_PULLING_BAGS
    End If
    If IsNurseryType("Other") Then BuildTFMSASprayPlots p
    If IsNurseryType("Hybrid") Then BuildHyHeights p

    btnBuildHomeNav

    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    Application.CalculateFull

    MsgBox nRows & " packets read from PRISM." & vbCrLf & vbCrLf & _
           "Generated:" & vbCrLf & _
           "  • Packet Prep 1.." & nSplits & " (QR text + colored digits, Split no.)" & vbCrLf & _
           "  • Material Map (Material ID grid)" & vbCrLf & _
           "  • Field Map (grid — run the wizard for planting dates)" & vbCrLf & _
           "  • Nursery list" & vbCrLf & _
           "  • Nursery-type extras for: " & CStr(GetSetting("Nursery types")) & vbCrLf & vbCrLf & _
           "Fieldbook is built later, from the 'Updated nursery site' tab." & vbCrLf & vbCrLf & _
           "The Packet Prep tabs carry the QR CODE text payload column — " & _
           "send this workbook to the barcode-printing machine.", _
           vbInformation, "All tabs generated"
End Sub


'==============================================================================
'  Packet Prep — 25 cols, QR text + colored digit columns
'==============================================================================
Private Sub BuildPacketPrep(ByRef p As Variant, ByVal nurseryCode As String, _
                            ByVal splitNo As Long)
    Dim ws As Worksheet: Set ws = GetSheet(PacketPrepName(splitNo))
    If ws Is Nothing Then Exit Sub
    ws.Range("A6:AZ" & ws.Rows.Count).Clear

    Dim n As Long: n = UBound(p, 1)
    ' Compute spike + rack per packet (default: spike = range, rack = row)
    Dim spike() As Long, rack() As Long
    ReDim spike(1 To n)
    ReDim rack(1 To n)
    Dim i As Long
    For i = 1 To n
        spike(i) = CLng(p(i, 1))
        rack(i) = CLng(p(i, 2))
    Next i

    ' Output rows — only the packets belonging to this split
    Dim outRow As Long: outRow = 6
    For i = 1 To n
        If SplitForRow(CLng(p(i, 2))) <> splitNo Then GoTo NextPacket
        Dim plot As String: plot = CLng(p(i, 1)) & "_" & CLng(p(i, 2))
        ws.Cells(outRow, 1).Value = QRText(plot, p(i, 3), p(i, 4), p(i, 5), p(i, 6), p(i, 7), p(i, 8))
        ws.Cells(outRow, 2).Value = p(i, 1)       ' Range
        ws.Cells(outRow, 3).Value = p(i, 2)       ' Row
        ws.Cells(outRow, 4).Value = plot           ' Plot
        ws.Cells(outRow, 5).Value = spike(i)       ' SPIKE#
        ws.Cells(outRow, 6).Value = rack(i)        ' RACK ORDER
        ' Column 7 is Split no.; digit and data columns shift right by one.
        ws.Cells(outRow, 7).Value = splitNo
        ' Colored digit columns (cols 8-11)
        Dim digits As String: digits = Format$(rack(i), "0000")
        ws.Cells(outRow, 8).Value = Mid$(digits, 1, 1)
        ws.Cells(outRow, 8).Font.Color = CLR_DIGIT_THOUSANDS
        ws.Cells(outRow, 8).Font.Bold = True
        ws.Cells(outRow, 9).Value = Mid$(digits, 2, 1)
        ws.Cells(outRow, 9).Font.Color = CLR_DIGIT_HUNDREDS
        ws.Cells(outRow, 9).Font.Bold = True
        ws.Cells(outRow, 10).Value = Mid$(digits, 3, 1)
        ws.Cells(outRow, 10).Font.Color = CLR_DIGIT_TENS
        ws.Cells(outRow, 10).Font.Bold = True
        ws.Cells(outRow, 11).Value = Mid$(digits, 4, 1)
        ws.Cells(outRow, 11).Font.Color = CLR_DIGIT_ONES
        ws.Cells(outRow, 11).Font.Bold = True
        ' Remaining PRISM cols (Material ID through Entry #) → cols 12-26
        ws.Cells(outRow, 12).Value = p(i, 3)      ' Material ID
        ws.Cells(outRow, 13).Value = p(i, 4)      ' Inbred Code
        ws.Cells(outRow, 14).Value = p(i, 5)      ' Source ID
        ws.Cells(outRow, 15).Value = p(i, 6)      ' CMS reaction
        ws.Cells(outRow, 16).Value = p(i, 7)      ' Generation
        ws.Cells(outRow, 17).Value = p(i, 8)      ' Comments
        ws.Cells(outRow, 18).Value = p(i, 9)      ' Pedigree
        ws.Cells(outRow, 19).Value = p(i, 10)     ' Hybrid Code
        ws.Cells(outRow, 20).Value = p(i, 11)     ' Trait Name
        ws.Cells(outRow, 21).Value = p(i, 12)     ' Plant #
        ws.Cells(outRow, 22).Value = p(i, 13)     ' Loc Seq#
        ws.Cells(outRow, 23).Value = p(i, 14)     ' SubSeq Flag
        ws.Cells(outRow, 24).Value = p(i, 15)     ' Entry Book Project
        ws.Cells(outRow, 25).Value = p(i, 16)     ' Entry Book Name
        ws.Cells(outRow, 26).Value = p(i, 17)     ' Entry #
        outRow = outRow + 1
NextPacket:
    Next i

    ' Sort by SPIKE# asc then RACK ORDER asc
    Dim lastDataRow As Long: lastDataRow = outRow - 1
    If lastDataRow >= 7 Then
        With ws.Sort
            .SortFields.Clear
            .SortFields.Add Key:=ws.Range(ws.Cells(6, 5), ws.Cells(lastDataRow, 5)), Order:=xlAscending
            .SortFields.Add Key:=ws.Range(ws.Cells(6, 6), ws.Cells(lastDataRow, 6)), Order:=xlAscending
            .SetRange ws.Range(ws.Cells(6, 1), ws.Cells(lastDataRow, 25))
            .Header = xlNo
            .Apply
        End With
    End If
End Sub

' Returns the comma-joined QR text:
'   Plot,Material ID,Inbred Code,Source ID,CMS,Generation,Comments
Private Function QRText(ByVal plot As String, _
                        ByVal material As Variant, ByVal inbred As Variant, _
                        ByVal source As Variant, ByVal cms As Variant, _
                        ByVal gen As Variant, ByVal comments As Variant) As String
    QRText = plot & "," & S(material) & "," & S(inbred) & "," & _
             S(source) & "," & S(cms) & "," & S(gen) & "," & S(comments)
End Function


'==============================================================================
'  Map / Material Map — 2D grid (range × row) with value at each cell
'==============================================================================
'  valueIdxs is an array of indices into the packet array p():
'    Material Map -> Array(3, 4, 10)  Material ID / Inbred Code / Hybrid Code
'    Field Map    -> Array()          grid geometry only, no material info
Private Sub BuildMap(ByRef p As Variant, ByVal sheetName As String, _
                     ByVal valueIdxs As Variant)
    Dim ws As Worksheet: Set ws = GetSheet(sheetName)
    If ws Is Nothing Then Exit Sub
    ws.Cells.Clear

    Dim n As Long: n = UBound(p, 1)
    Dim maxRange As Long, maxRow As Long
    Dim i As Long
    For i = 1 To n
        If CLng(p(i, 1)) > maxRange Then maxRange = CLng(p(i, 1))
        If CLng(p(i, 2)) > maxRow Then maxRow = CLng(p(i, 2))
    Next i
    If maxRange = 0 Or maxRow = 0 Then Exit Sub

    ' Title
    ws.Cells(1, 1).Value = sheetName & " — " & CStr(GetSetting("Nursery code"))
    ws.Cells(1, 1).Font.Bold = True
    ws.Cells(1, 1).Font.Size = 14
    ws.Cells(2, 1).Value = "Range numbers down the side · Field rows across the top"
    ws.Cells(2, 1).Font.Italic = True

    ' Header row (field rows 1..maxRow)
    Const HDR_ROW As Long = 3
    ws.Cells(HDR_ROW, 1).Value = "Rng \ Row"
    ws.Cells(HDR_ROW, 1).Font.Bold = True
    ws.Cells(HDR_ROW, 1).Interior.Color = HDR_PALE
    ws.Cells(HDR_ROW, 1).HorizontalAlignment = xlCenter
    Dim fr As Long
    For fr = 1 To maxRow
        With ws.Cells(HDR_ROW, fr + 1)
            .Value = fr
            .Font.Bold = True
            .Interior.Color = HDR_PALE
            .HorizontalAlignment = xlCenter
        End With
    Next fr
    ' Right-side label
    With ws.Cells(HDR_ROW, maxRow + 2)
        .Value = "Rng"
        .Font.Bold = True
        .Interior.Color = HDR_PALE
        .HorizontalAlignment = xlCenter
    End With

    ' Build lookup of (range,row) → packet index for fast cell fill
    Dim dict As Object
    Set dict = CreateObject("Scripting.Dictionary")
    On Error Resume Next
    If dict Is Nothing Or Err.Number <> 0 Then
        ' Mac fallback: skip dict, do linear lookup. With n~1300 + cells~maxR*maxRow,
        ' this is too slow. So pre-build a flat array key.
        Err.Clear
    End If
    On Error GoTo 0

    Dim rng As Long, rw As Long
    ' Fast: build a flat 2D array of values, then write in one shot
    Dim grid() As Variant
    ReDim grid(1 To maxRange, 1 To maxRow)
    For i = 1 To n
        rng = CLng(p(i, 1)): rw = CLng(p(i, 2))
        If rng >= 1 And rng <= maxRange And rw >= 1 And rw <= maxRow Then
            ' An empty valueIdxs draws grid geometry only — that is the Field
            ' Map, which carries no material information. Otherwise stack each
            ' requested identifier on its own line within the cell.
            Dim v As String, k As Long, part As String
            v = ""
            For k = LBound(valueIdxs) To UBound(valueIdxs)
                part = S(p(i, CLng(valueIdxs(k))))
                If Len(part) > 14 Then part = Left$(part, 14)
                If Len(part) > 0 Then
                    If Len(v) > 0 Then v = v & vbLf
                    v = v & part
                End If
            Next k
            grid(rng, rw) = v
        End If
    Next i

    ' Write data rows — ranges descending (top range at top)
    Dim outR As Long: outR = HDR_ROW + 1
    For rng = maxRange To 1 Step -1
        With ws.Cells(outR, 1)
            .Value = rng
            .Font.Bold = True
            .Interior.Color = HDR_PALE
            .HorizontalAlignment = xlCenter
        End With
        For fr = 1 To maxRow
            If Len(CStr(grid(rng, fr))) > 0 Then
                ws.Cells(outR, fr + 1).Value = grid(rng, fr)
                ws.Cells(outR, fr + 1).HorizontalAlignment = xlCenter
                ws.Cells(outR, fr + 1).Font.Size = 9
            End If
        Next fr
        With ws.Cells(outR, maxRow + 2)
            .Value = rng
            .Font.Bold = True
            .Interior.Color = HDR_PALE
            .HorizontalAlignment = xlCenter
        End With
        outR = outR + 1
    Next rng

    ' Bottom header — row numbers repeat along the foot so the map stays
    ' readable when it is printed and read from the far end of the block.
    Dim footRow As Long: footRow = outR
    ws.Cells(footRow, 1).Value = "Rng \ Row"
    ws.Cells(footRow, 1).Font.Bold = True
    ws.Cells(footRow, 1).Interior.Color = HDR_PALE
    ws.Cells(footRow, 1).HorizontalAlignment = xlCenter
    For fr = 1 To maxRow
        With ws.Cells(footRow, fr + 1)
            .Value = fr
            .Font.Bold = True
            .Interior.Color = HDR_PALE
            .HorizontalAlignment = xlCenter
        End With
    Next fr
    With ws.Cells(footRow, maxRow + 2)
        .Value = "Rng"
        .Font.Bold = True
        .Interior.Color = HDR_PALE
        .HorizontalAlignment = xlCenter
    End With

    ' Grid borders and wrapping for the stacked identifiers
    Dim nVals As Long: nVals = UBound(valueIdxs) - LBound(valueIdxs) + 1
    With ws.Range(ws.Cells(HDR_ROW, 1), ws.Cells(footRow, maxRow + 2))
        .Borders.LineStyle = xlContinuous
        .Borders.Weight = xlThin
        .WrapText = (nVals > 1)
        .VerticalAlignment = xlCenter
    End With

    If nVals > 1 Then
        For outR = HDR_ROW + 1 To footRow - 1
            ws.Rows(outR).RowHeight = 12 * nVals + 4
        Next outR
    End If

    ' Column widths
    ws.Columns(1).ColumnWidth = 9
    For fr = 1 To maxRow
        ws.Columns(fr + 1).ColumnWidth = 11
    Next fr
    ws.Columns(maxRow + 2).ColumnWidth = 7

    ws.Cells(2, 1).Value = "Range numbers down both sides · Field rows across " & _
                           "top and bottom"
End Sub


'==============================================================================
'  Nursery list — Source ID (Hybrid Code) grouped by count
'==============================================================================
Private Sub BuildNurseryList(ByRef p As Variant)
    Dim ws As Worksheet: Set ws = GetSheet(SHEET_NURSERY_LIST)
    If ws Is Nothing Then Exit Sub
    ws.Range("A6:Z" & ws.Rows.Count).Clear

    Dim n As Long: n = UBound(p, 1)
    Dim labels() As String, counts() As Long, inbreds() As String, hybrids() As String
    ReDim labels(1 To n): ReDim counts(1 To n)
    ReDim inbreds(1 To n): ReDim hybrids(1 To n)
    Dim m As Long: m = 0
    Dim i As Long, j As Long
    For i = 1 To n
        Dim sid As String, hc As String, lbl As String
        sid = S(p(i, 5))
        hc = S(p(i, 10))
        If Len(sid) = 0 Then sid = "(unknown)"
        If Len(hc) > 0 Then
            lbl = sid & " (" & hc & ")"
        Else
            lbl = sid
        End If
        ' Linear find
        Dim found As Boolean: found = False
        For j = 1 To m
            If labels(j) = lbl Then
                counts(j) = counts(j) + 1
                found = True
                Exit For
            End If
        Next j
        If Not found Then
            m = m + 1
            labels(m) = lbl
            counts(m) = 1
            inbreds(m) = S(p(i, 4))
            hybrids(m) = hc
        End If
    Next i

    ' Bubble sort by label (small dataset)
    Dim tmp As String, tmpN As Long
    For i = 1 To m - 1
        For j = i + 1 To m
            If labels(i) > labels(j) Then
                tmp = labels(i): labels(i) = labels(j): labels(j) = tmp
                tmpN = counts(i): counts(i) = counts(j): counts(j) = tmpN
                tmp = inbreds(i): inbreds(i) = inbreds(j): inbreds(j) = tmp
                tmp = hybrids(i): hybrids(i) = hybrids(j): hybrids(j) = tmp
            End If
        Next j
    Next i

    ' Write rows (start at row 6, leave cols A/B blank to match sample format)
    Dim qty As Double
    qty = CDbl(GetSetting("Qty per packet"))
    If qty <= 0 Then qty = 1.4

    Dim outR As Long: outR = 6
    For i = 1 To m
        ws.Cells(outR, 3).Value = labels(i)
        ws.Cells(outR, 4).Value = counts(i)
        ws.Cells(outR, 5).Value = Round(counts(i) * qty, 1)
        ws.Cells(outR, 6).Value = inbreds(i)
        ws.Cells(outR, 7).Value = hybrids(i)
        outR = outR + 1
    Next i
End Sub


'==============================================================================
'  Fieldbook — 10 cols, Range/Row/R_R/blanks/Material/Source/Gen/CMS/Comments
'==============================================================================
Private Sub BuildFieldbook(ByRef p As Variant)
    Dim ws As Worksheet: Set ws = GetSheet(SHEET_FIELDBOOK)
    If ws Is Nothing Then Exit Sub
    ws.Range("A6:Z" & ws.Rows.Count).Clear

    Dim n As Long: n = UBound(p, 1)
    Dim outR As Long: outR = 6
    Dim i As Long
    For i = 1 To n
        ws.Cells(outR, 1).Value = p(i, 1)     ' Range
        ws.Cells(outR, 2).Value = p(i, 2)     ' Row
        ws.Cells(outR, 3).Value = CLng(p(i, 1)) & "_" & CLng(p(i, 2))   ' R_R
        ' col 4 (Crossed bags) + col 5 (Bagging Info) left blank
        ws.Cells(outR, 6).Value = p(i, 3)     ' Material ID
        ws.Cells(outR, 7).Value = p(i, 5)     ' Source ID
        ws.Cells(outR, 8).Value = p(i, 7)     ' Gen
        ws.Cells(outR, 9).Value = p(i, 6)     ' CMS
        ws.Cells(outR, 10).Value = p(i, 8)    ' Comments
        outR = outR + 1
    Next i

    ' Sort by Range then Row
    Dim lastDataRow As Long: lastDataRow = outR - 1
    If lastDataRow >= 7 Then
        With ws.Sort
            .SortFields.Clear
            .SortFields.Add Key:=ws.Range(ws.Cells(6, 1), ws.Cells(lastDataRow, 1)), Order:=xlAscending
            .SortFields.Add Key:=ws.Range(ws.Cells(6, 2), ws.Cells(lastDataRow, 2)), Order:=xlAscending
            .SetRange ws.Range(ws.Cells(6, 1), ws.Cells(lastDataRow, 10))
            .Header = xlNo
            .Apply
        End With
    End If
End Sub


'==============================================================================
'  BC0 labels — BC* generations only
'==============================================================================
Private Sub BuildBC0Labels(ByRef p As Variant, ByVal nurseryCode As String)
    Dim ws As Worksheet: Set ws = GetSheet(SHEET_BC0_LABELS)
    If ws Is Nothing Then Exit Sub
    ws.Range("A6:Z" & ws.Rows.Count).Clear

    Dim n As Long: n = UBound(p, 1)
    Dim outR As Long: outR = 6
    Dim i As Long, gen As String
    For i = 1 To n
        gen = S(p(i, 7))
        If IsBCGen(gen) Then
            ws.Cells(outR, 1).Value = p(i, 1)         ' Range
            ws.Cells(outR, 2).Value = p(i, 2)         ' Row
            ' col 3 (Crossed bags), 4 (TFMSA), 5 (Pollen), 6 (TFMSA/Pollen) blank
            ws.Cells(outR, 7).Value = nurseryCode    ' Nursery Name
            ' col 8 (Bagging Info) blank
            ws.Cells(outR, 9).Value = p(i, 3)         ' Material ID
            ws.Cells(outR, 10).Value = p(i, 5)        ' Source ID
            ws.Cells(outR, 11).Value = gen            ' Gen
            ws.Cells(outR, 12).Value = p(i, 6)        ' CMS
            ws.Cells(outR, 13).Value = p(i, 8)        ' Comments
            outR = outR + 1
        End If
    Next i
End Sub


'==============================================================================
'  Date recording + Pulling bags — same data, same shape
'==============================================================================
Private Sub BuildDateRecording(ByRef p As Variant, ByVal sheetName As String)
    ' 11-column layout from the client's 'Date recording' sample.
    ' Group and O/E look like scaffolding but are load-bearing: Group is the
    ' two-row band index the '2-rows together' button orders by, and O/E
    ' carries the parity that drives the serpentine direction. Both are
    ' hidden rather than removed.
    Dim ws As Worksheet: Set ws = GetSheet(sheetName)
    If ws Is Nothing Then Exit Sub
    ws.Cells.Clear

    Dim headers As Variant
    headers = Array("Group", "Range", "Row", "O/E", "S 1", "S 2", _
                    "Material ID", "Source ID", "Gen", "CMS", "In. Code")
    Dim widths As Variant
    widths = Array(6.29, 6.86, 5.43, 9.14, 5.14, 5.57, 21.43, 28.57, 5.14, 4.71, 8.43)

    Dim c As Long
    For c = 0 To UBound(headers)
        ws.Cells(1, c + 1).Value = headers(c)
        ws.Cells(1, c + 1).Font.Bold = True
        ws.Columns(c + 1).ColumnWidth = widths(c)
    Next c

    Dim n As Long: n = UBound(p, 1)
    Dim outR As Long: outR = 2
    Dim i As Long, gen As String, rowNo As Long
    For i = 1 To n
        gen = S(p(i, 7))
        If IsRecurrent(gen) Then
            rowNo = CLng(p(i, 2))
            ws.Cells(outR, 1).Value = BandForRow(rowNo)     ' Group
            ws.Cells(outR, 2).Value = p(i, 1)               ' Range
            ws.Cells(outR, 3).Value = rowNo                 ' Row
            ws.Cells(outR, 4).Value = ParityForRow(rowNo)   ' O/E
            ' S 1 and S 2 stay empty — day-of-month numbers entered in field.
            ws.Cells(outR, 7).Value = p(i, 3)               ' Material ID
            ws.Cells(outR, 8).Value = p(i, 5)               ' Source ID
            ws.Cells(outR, 9).Value = gen                   ' Gen
            ws.Cells(outR, 10).Value = p(i, 6)              ' CMS
            ws.Cells(outR, 11).Value = p(i, 4)              ' In. Code
            outR = outR + 1
        End If
    Next i

    With ws.Range(ws.Cells(1, 1), ws.Cells(outR - 1, 11)).Borders
        .LineStyle = xlContinuous
        .Weight = xlThin
    End With

    ws.Columns(1).Hidden = True     ' Group
    ws.Columns(4).Hidden = True     ' O/E
End Sub


'==============================================================================
'  TFMSA Spray plots — BC* only
'==============================================================================
Private Sub BuildTFMSASprayPlots(ByRef p As Variant)
    Dim ws As Worksheet: Set ws = GetSheet(SHEET_TFMSA_SPRAY_PLOTS)
    If ws Is Nothing Then Exit Sub
    ws.Range("A6:Z" & ws.Rows.Count).Clear

    Dim n As Long: n = UBound(p, 1)
    Dim outR As Long: outR = 6
    Dim i As Long, gen As String
    For i = 1 To n
        gen = S(p(i, 7))
        If IsBCGen(gen) Then
            ws.Cells(outR, 1).Value = p(i, 1)     ' Range
            ws.Cells(outR, 2).Value = p(i, 2)     ' Row
            ws.Cells(outR, 3).Value = p(i, 5)     ' Source ID
            ws.Cells(outR, 4).Value = p(i, 6)     ' CMS
            ws.Cells(outR, 5).Value = gen          ' Gen
            outR = outR + 1
        End If
    Next i
End Sub


'==============================================================================
'  Hy Heights — F1 only
'==============================================================================
Private Sub BuildHyHeights(ByRef p As Variant)
    Dim ws As Worksheet: Set ws = GetSheet(SHEET_HY_HEIGHTS)
    If ws Is Nothing Then Exit Sub
    ws.Range("A6:Z" & ws.Rows.Count).Clear

    Dim n As Long: n = UBound(p, 1)
    Dim outR As Long: outR = 6
    Dim i As Long, gen As String
    For i = 1 To n
        gen = S(p(i, 7))
        If IsHybridF1(gen) Then
            ws.Cells(outR, 1).Value = p(i, 1)     ' Range
            ws.Cells(outR, 2).Value = p(i, 2)     ' Row
            ' col 3 (Height in CM) blank
            ws.Cells(outR, 4).Value = p(i, 3)     ' Material ID
            ws.Cells(outR, 5).Value = p(i, 5)     ' Source ID
            ws.Cells(outR, 6).Value = gen          ' Gen
            ws.Cells(outR, 7).Value = p(i, 6)     ' CMS
            outR = outR + 1
        End If
    Next i
End Sub


'==============================================================================
'  Step 3 — Sort for racking (LSD Radix)
'==============================================================================
Public Sub btnSortForRacking()
    ' Packet Prep is now one tab per split, so sort whichever split the user
    ' is looking at; fall back to split 1 if they ran this from elsewhere.
    Dim ws As Worksheet
    If InStr(ActiveSheet.Name, SHEET_PACKET_PREP) = 1 Then
        Set ws = ActiveSheet
    Else
        Set ws = GetSheet(PacketPrepName(1))
    End If
    If ws Is Nothing Then Exit Sub

    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, 2).End(xlUp).Row
    If lastRow < 7 Then
        MsgBox "Run Step 2 first — '" & ws.Name & "' has no packets yet.", _
               vbExclamation
        Exit Sub
    End If
    With ws.Sort
        .SortFields.Clear
        .SortFields.Add Key:=ws.Range(ws.Cells(6, 6), ws.Cells(lastRow, 6)), Order:=xlAscending
        .SortFields.Add Key:=ws.Range(ws.Cells(6, 5), ws.Cells(lastRow, 5)), Order:=xlAscending
        .SetRange ws.Range(ws.Cells(6, 1), ws.Cells(lastRow, 26))
        .Header = xlNo
        .Apply
    End With
    ws.Activate
    MsgBox ws.Name & " re-sorted in LSD-radix racking order (rack ↑, spike ↑).", _
           vbInformation, "Sorted"
End Sub


'==============================================================================
'  Field events (Steps 4–7)
'==============================================================================
Public Sub btnAddReplacement():    AddEventRow "replacement":          End Sub
Public Sub btnAddPlantingError():  AddEventRow "planting_error":       End Sub
Public Sub btnRecordSpray():       AddEventRow "spray":                End Sub
Public Sub btnRecordABPull():      AddEventRow "ab_pull":              End Sub

Private Sub AddEventRow(ByVal evType As String)
    Dim destSheet As String
    Select Case evType
        ' Replacements and planting errors share one tab now; Stage tells
        ' them apart on the row itself.
        Case "replacement":     destSheet = SHEET_REPLACEMENTS_AND_ERRORS
        Case "planting_error":  destSheet = SHEET_REPLACEMENTS_AND_ERRORS
        Case "spray":           destSheet = SHEET_TFMSA_SPRAY_PLOTS
        Case "ab_pull":         destSheet = SHEET_PULLING_BAGS
    End Select
    Dim ws As Worksheet: Set ws = GetSheet(destSheet)
    If ws Is Nothing Then Exit Sub

    Dim plot As String
    plot = InputBox("Plot (e.g. 2_8):", "New " & evType)
    If Len(plot) = 0 Then Exit Sub
    Dim r As Long: r = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row + 1
    If r < 6 Then r = 6

    Select Case evType
        Case "replacement"
            ws.Cells(r, 1).Value = plot  ' use Qrcode-ish slot
            ws.Cells(r, 2).Value = InputBox("Replacement Source ID:", "Replacement")
            ws.Cells(r, 3).Value = "Open"
            ws.Cells(r, 4).Value = InputBox("Note (optional):", "Replacement")
        Case "planting_error"
            ws.Cells(r, 1).Value = plot
            Dim parts() As String
            If InStr(plot, "_") > 0 Then
                parts = Split(plot, "_")
                ws.Cells(r, 2).Value = parts(0)  ' Range
                ws.Cells(r, 3).Value = parts(1)  ' Row
            End If
            ws.Cells(r, 4).Value = InputBox("Description of the error:", "Planting error")
            ws.Cells(r, 5).Value = InputBox("Severity (Low/Medium/High):", "Planting error", "Medium")
            ws.Cells(r, 6).Value = Format$(Date, "yyyy-mm-dd")
            ws.Cells(r, 7).Value = "Open"
        Case "spray"
            ws.Cells(r, 1).Value = plot
        Case "ab_pull"
            ws.Cells(r, 1).Value = plot
    End Select
    ws.Activate
    ws.Cells(r, 1).Select
End Sub


'==============================================================================
'  Updated PRISM + Refresh dashboard + Push to Hub
'==============================================================================
Public Sub btnImportUpdated()
    Dim ws As Worksheet: Set ws = GetSheet(SHEET_NURSERY_SITE)
    If Not ws Is Nothing Then ws.Activate
    MsgBox "Paste the refreshed PRISM 'Nursery site' export into this tab " & _
           "starting at row 6 (headers stay at row 5)." & vbCrLf & vbCrLf & _
           "Then click Step 2 to regenerate every tab.", _
           vbInformation, "Updated Nursery site"
End Sub

Public Sub btnRefreshDashboard()
    Application.CalculateFull
End Sub

Public Sub btnPushToHub()
    On Error GoTo cleanFail
    Dim folder As String
    folder = ExpandTilde(CStr(GetSetting("Hub registry folder")))
    If Len(folder) = 0 Then Exit Sub
    If Right$(folder, 1) <> Application.PathSeparator Then _
        folder = folder & Application.PathSeparator
    EnsureFolder folder

    Dim regPath As String: regPath = folder & "registry.csv"
    Dim code As String: code = CStr(GetSetting("Nursery code"))
    If Len(code) = 0 Then code = "(uncoded)"
    Dim season As String: season = CStr(GetSetting("Season"))
    Dim breeder As String: breeder = CStr(GetSetting("Breeder"))

    Dim packets As Long, reps As Long, errs As Long, sprays As Long
    packets = CountRows(SHEET_NURSERY_SITE)
    ' One tab holds both now, so split the count on the Stage column rather
    ' than counting the same rows twice.
    reps = CountByStage("Packeting")
    errs = CountByStage("Planting")
    sprays = CountRows(SHEET_TFMSA_SPRAY_PLOTS)

    Dim ff As Integer, line As String
    Dim outLines As Collection: Set outLines = New Collection
    outLines.Add "nursery_code,season,breeder,packets,replacements,errors,additionals,file_path,last_update"
    If FileExists(regPath) Then
        ff = FreeFile
        Open regPath For Input As #ff
        Dim header As Boolean: header = True
        Do While Not EOF(ff)
            Line Input #ff, line
            If header Then header = False: GoTo nxt
            If Len(Trim$(line)) = 0 Then GoTo nxt
            If InStr(1, line, code & ",") <> 1 Then outLines.Add line
nxt:
        Loop
        Close #ff
    End If
    Dim row As String
    row = code & "," & CsvEsc(season) & "," & CsvEsc(breeder) & "," & _
          packets & "," & reps & "," & errs & "," & sprays & "," & _
          CsvEsc(ThisWorkbook.FullName) & "," & CsvEsc(NowISO())
    outLines.Add row
    ff = FreeFile
    Open regPath For Output As #ff
    Dim it As Variant
    For Each it In outLines: Print #ff, CStr(it): Next it
    Close #ff
    On Error Resume Next
    ThisWorkbook.Names("DASH_LastSync").RefersToRange.Value = NowISO()
    On Error GoTo 0
    MsgBox "Pushed to Hub registry:" & vbCrLf & regPath, vbInformation, "Synced"
    Exit Sub
cleanFail:
    Debug.Print "btnPushToHub failed: " & Err.Number & " " & Err.Description
End Sub

'==============================================================================
'  v2 — ordering rules
'  Ports of excel_workflow/nursery_algos.py. Keep in step with the Python
'  originals; both are pinned by the same client-sample cases.
'==============================================================================
Public Function SpikeForRow(ByVal r As Long) As Long
    ' A two-cone planter's cones swap sides when it turns around, so spikes
    ' run 1,2 down the forward pass and 2,1 back up the reverse pass.
    Dim cycle As Variant
    cycle = Array(1, 2, 2, 1)
    SpikeForRow = cycle((r - 1) Mod 4)
End Function

Public Function RunDirection(ByVal r As Long) As String
    If (((r - 1) \ 2) Mod 2) = 0 Then
        RunDirection = "forward"
    Else
        RunDirection = "reverse"
    End If
End Function

Public Function BandForRow(ByVal r As Long) As Long
    ' Rows 2 and 3 are band 1; rows 26 and 27 are band 13.
    BandForRow = r \ 2
End Function

Public Function ParityForRow(ByVal r As Long) As String
    If (r Mod 2) = 1 Then ParityForRow = "O" Else ParityForRow = "E"
End Function

Public Function SerpentineKey(ByVal rangeNo As Long, ByVal rowNo As Long) As Long
    ' Rows descend on even ranges, ascend on odd. Negating on even ranges
    ' collapses the snake into one sortable number.
    If (rangeNo Mod 2) = 0 Then
        SerpentineKey = -rowNo
    Else
        SerpentineKey = rowNo
    End If
End Function


'==============================================================================
'  v2 — nursery type + split helpers
'==============================================================================
Public Function IsNurseryType(ByVal t As String) As Boolean
    Dim s As String
    s = "," & Replace(LCase(CStr(GetSetting("Nursery types"))), " ", "") & ","
    IsNurseryType = InStr(s, "," & LCase(t) & ",") > 0
End Function

Public Function PlantingDateCount() As Long
    Dim n As Long
    n = CLng(Val(GetSetting("Planting dates")))
    If n < 1 Then n = 1
    PlantingDateCount = n
End Function

Public Function PacketPrepName(ByVal splitNo As Long) As String
    PacketPrepName = SHEET_PACKET_PREP & " " & splitNo
End Function

Public Function SplitForRow(ByVal fieldRow As Long) As Long
    ' Reads the Field Map wizard's output block (cols T..X). Distinct planting
    ' dates appear in order, so the Nth distinct date is split N.
    ' Before the wizard has run there is one split and everything is in it.
    Const OUT_COL As Long = 20

    Dim ws As Worksheet: Set ws = GetSheet(SHEET_FIELD_MAP)
    If ws Is Nothing Then
        SplitForRow = 1
        Exit Function
    End If

    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, OUT_COL + 1).End(xlUp).Row
    If lastRow < 3 Then
        SplitForRow = 1
        Exit Function
    End If

    Dim dates As Object: Set dates = CreateObject("Scripting.Dictionary")
    Dim r As Long, dop As String, n As Long
    For r = 3 To lastRow
        dop = S(ws.Cells(r, OUT_COL).Value)
        If Len(dop) > 0 Then
            If Not dates.Exists(dop) Then
                n = n + 1
                dates.Add dop, n
            End If
            If CLng(Val(ws.Cells(r, OUT_COL + 1).Value)) = fieldRow Then
                SplitForRow = dates(dop)
                Exit Function
            End If
        End If
    Next r

    ' Row not covered by any planting date — keep it visible in split 1
    ' rather than dropping it silently from every Packet Prep tab.
    SplitForRow = 1
End Function

Private Function CountByStage(ByVal stageName As String) As Long
    Dim ws As Worksheet: Set ws = GetSheet(SHEET_REPLACEMENTS_AND_ERRORS)
    If ws Is Nothing Then Exit Function

    Dim lastRow As Long, r As Long, n As Long
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    For r = 6 To lastRow
        If LCase(S(ws.Cells(r, 1).Value)) = LCase(stageName) Then n = n + 1
    Next r
    CountByStage = n
End Function


'==============================================================================
'  v2 — Home navigation
'==============================================================================
Public Sub btnBuildHomeNav()
    Dim ws As Worksheet, home As Worksheet
    Dim r As Long
    Set home = GetSheet(SHEET_HOME)
    If home Is Nothing Then Exit Sub

    home.Range("I5:I200").Clear
    home.Cells(4, 9).Value = "GO TO TAB"
    home.Cells(4, 9).Font.Bold = True
    home.Columns(9).ColumnWidth = 26

    r = 5
    For Each ws In ThisWorkbook.Worksheets
        If ws.Name <> SHEET_HOME And ws.Visible = xlSheetVisible Then
            home.Hyperlinks.Add Anchor:=home.Cells(r, 9), Address:="", _
                SubAddress:="'" & ws.Name & "'!A1", TextToDisplay:=ws.Name
            r = r + 1
        End If
    Next ws
End Sub


'==============================================================================
'  v2 — Field Map wizard
'  Re-runnable: clears its own output block so repeated runs do not stack.
'==============================================================================
Public Sub btnFieldMapWizard()
    Const CLR_FORWARD As Long = 15792383   ' pale blue
    Const CLR_REVERSE As Long = 15794160   ' pale green
    Const OUT_COL As Long = 20             ' output block starts at column T

    Dim ws As Worksheet: Set ws = GetSheet(SHEET_FIELD_MAP)
    If ws Is Nothing Then Exit Sub

    Dim nDates As Long
    nDates = CLng(Val(InputBox("How many planting dates?", "Field Map", "2")))
    If nDates < 1 Then Exit Sub

    Dim qty As String
    qty = InputBox("Seed quantity per plot:", "Field Map", _
                   CStr(GetSetting("Seed qty per plot")))
    SetSetting "Seed qty per plot", qty

    Dim topRow As Long: topRow = 2
    ws.Range(ws.Cells(topRow, OUT_COL), ws.Cells(topRow + 400, OUT_COL + 4)).Clear

    ws.Cells(topRow, OUT_COL).Value = "Planting date"
    ws.Cells(topRow, OUT_COL + 1).Value = "Row"
    ws.Cells(topRow, OUT_COL + 2).Value = "Spike"
    ws.Cells(topRow, OUT_COL + 3).Value = "Run"
    ws.Cells(topRow, OUT_COL + 4).Value = "Seed qty/plot"
    ws.Range(ws.Cells(topRow, OUT_COL), ws.Cells(topRow, OUT_COL + 4)).Font.Bold = True

    Dim seen As Object: Set seen = CreateObject("Scripting.Dictionary")
    Dim outR As Long: outR = topRow + 1
    Dim i As Long, dop As String, rowsCsv As String
    Dim parts As Variant, j As Long, rowNo As Long, dir As String

    For i = 1 To nDates
        dop = InputBox("Date of planting " & i & " (e.g. 25-Feb-2026):", "Field Map")
        If Len(dop) = 0 Then Exit Sub
        rowsCsv = InputBox("Rows for planting date " & i & " (comma separated):", _
                           "Field Map")
        If Len(rowsCsv) = 0 Then Exit Sub

        parts = Split(rowsCsv, ",")
        For j = LBound(parts) To UBound(parts)
            rowNo = CLng(Val(Trim(parts(j))))
            If rowNo > 0 Then
                ' A row in two planting dates would print packets onto the
                ' wrong split, so refuse rather than silently overwrite.
                If seen.Exists(rowNo) Then
                    MsgBox "Row " & rowNo & " is already assigned to planting " & _
                           "date " & seen(rowNo) & ". Fix the row lists and " & _
                           "run the wizard again.", vbCritical, "Overlapping rows"
                    Exit Sub
                End If
                seen.Add rowNo, i

                dir = RunDirection(rowNo)
                ws.Cells(outR, OUT_COL).Value = dop
                ws.Cells(outR, OUT_COL + 1).Value = rowNo
                ws.Cells(outR, OUT_COL + 2).Value = SpikeForRow(rowNo)
                ws.Cells(outR, OUT_COL + 3).Value = dir
                ws.Cells(outR, OUT_COL + 4).Value = qty
                If dir = "forward" Then
                    ws.Range(ws.Cells(outR, OUT_COL), ws.Cells(outR, OUT_COL + 4)) _
                      .Interior.Color = CLR_FORWARD
                Else
                    ws.Range(ws.Cells(outR, OUT_COL), ws.Cells(outR, OUT_COL + 4)) _
                      .Interior.Color = CLR_REVERSE
                End If
                outR = outR + 1
            End If
        Next j
    Next i

    SetSetting "Planting dates", nDates
    MsgBox nDates & " planting date(s) recorded, " & (outR - topRow - 1) & _
           " rows assigned." & vbCrLf & vbCrLf & _
           "Re-run Step 2 to rebuild Packet Prep with these splits.", _
           vbInformation, "Field Map"
End Sub


'==============================================================================
'  v2 — Fieldbook, built from Updated nursery site
'==============================================================================
Public Sub btnBuildFieldbook()
    Dim src As Worksheet: Set src = GetSheet(SHEET_UPDATED_NURSERY_SITE)
    If src Is Nothing Then
        MsgBox "Tab '" & SHEET_UPDATED_NURSERY_SITE & "' not found.", vbCritical
        Exit Sub
    End If
    If S(src.Cells(6, 1).Value) = "" Then
        MsgBox "Paste the updated PRISM export into '" & _
               SHEET_UPDATED_NURSERY_SITE & "' first " & _
               "(headers row 5, data row 6+).", vbExclamation, "Nothing to build from"
        src.Activate
        Exit Sub
    End If
    BuildFieldbookV2 src
End Sub

Private Sub BuildFieldbookV2(ByVal src As Worksheet)
    Const CLR_BLINE As Long = 13561798    ' pale green

    Dim ws As Worksheet: Set ws = GetSheet(SHEET_FIELDBOOK)
    If ws Is Nothing Then Exit Sub
    ws.Cells.Clear

    Dim cRange As Long, cRow As Long, cMat As Long, cSrc As Long
    Dim cGen As Long, cCms As Long, cCom As Long
    cRange = FindHeaderCol(src, "Range", 5)
    cRow = FindHeaderCol(src, "Row", 5)
    cMat = FindHeaderCol(src, "Material ID", 5)
    cSrc = FindHeaderCol(src, "Source ID", 5)
    cGen = FindHeaderCol(src, "Generation", 5)
    cCms = FindHeaderCol(src, "CMS reaction", 5)
    cCom = FindHeaderCol(src, "Comments", 5)
    If cRange = 0 Or cRow = 0 Then
        MsgBox "Could not find Range/Row columns in row 5 of '" & _
               SHEET_UPDATED_NURSERY_SITE & "'.", vbCritical
        Exit Sub
    End If

    ws.Cells(1, FB_COL_RANGE).Value = "Range"
    ws.Cells(1, FB_COL_ROW).Value = "Row"
    ws.Cells(1, FB_COL_MATERIAL_ID).Value = "Material ID"
    ws.Cells(1, FB_COL_SOURCE_ID).Value = "Source ID"
    ws.Cells(1, FB_COL_GEN).Value = "Gen"
    ws.Cells(1, FB_COL_CMS).Value = "CMS"
    ws.Cells(1, FB_COL_PLOT).Value = "Plot"
    ws.Cells(1, FB_COL_COMMENTS).Value = "Comments"

    Dim lastSrc As Long, r As Long, outR As Long
    lastSrc = src.Cells(src.Rows.Count, cRange).End(xlUp).Row
    outR = 2
    For r = 6 To lastSrc
        Dim rangeNo As Long, rowNo As Long
        rangeNo = CLng(Val(src.Cells(r, cRange).Value))
        rowNo = CLng(Val(src.Cells(r, cRow).Value))
        ws.Cells(outR, FB_COL_RANGE).Value = rangeNo
        ws.Cells(outR, FB_COL_ROW).Value = rowNo
        ws.Cells(outR, FB_COL_MATERIAL_ID).Value = IIf(cMat > 0, src.Cells(r, cMat).Value, "")
        ws.Cells(outR, FB_COL_SOURCE_ID).Value = IIf(cSrc > 0, src.Cells(r, cSrc).Value, "")
        ws.Cells(outR, FB_COL_GEN).Value = IIf(cGen > 0, src.Cells(r, cGen).Value, "")
        ws.Cells(outR, FB_COL_CMS).Value = IIf(cCms > 0, src.Cells(r, cCms).Value, "")
        ws.Cells(outR, FB_COL_PLOT).Value = rangeNo & "_" & rowNo
        ws.Cells(outR, FB_COL_COMMENTS).Value = IIf(cCom > 0, src.Cells(r, cCom).Value, "")
        ' Scratch serpentine key, cleared once the sort is done.
        ws.Cells(outR, FB_COL_COMMENTS + 2).Value = SerpentineKey(rangeNo, rowNo)
        outR = outR + 1
    Next r

    Dim lastRow As Long: lastRow = outR - 1
    If lastRow >= 3 Then
        With ws.Sort
            .SortFields.Clear
            .SortFields.Add Key:=ws.Range(ws.Cells(2, FB_COL_RANGE), _
                                          ws.Cells(lastRow, FB_COL_RANGE)), _
                            Order:=xlAscending
            .SortFields.Add Key:=ws.Range(ws.Cells(2, FB_COL_COMMENTS + 2), _
                                          ws.Cells(lastRow, FB_COL_COMMENTS + 2)), _
                            Order:=xlAscending
            .SetRange ws.Range(ws.Cells(2, 1), ws.Cells(lastRow, FB_COL_COMMENTS + 2))
            .Header = xlNo
            .Apply
        End With
    End If
    ws.Columns(FB_COL_COMMENTS + 2).Clear

    With ws.Range(ws.Cells(1, 1), ws.Cells(lastRow, FB_COL_COMMENTS)).Borders
        .LineStyle = xlContinuous
        .Weight = xlThin
    End With
    ws.Range(ws.Cells(1, 1), ws.Cells(1, FB_COL_COMMENTS)).Font.Bold = True
    ws.Range(ws.Cells(1, 1), ws.Cells(1, FB_COL_COMMENTS)).HorizontalAlignment = xlCenter
    ws.Columns(FB_COL_MATERIAL_ID).ColumnWidth = 22
    ws.Columns(FB_COL_SOURCE_ID).ColumnWidth = 28
    ws.Columns(FB_COL_COMMENTS).ColumnWidth = 30

    If IsNurseryType("AB") Then
        For r = 2 To lastRow
            If UCase(S(ws.Cells(r, FB_COL_CMS).Value)) = "B" Then
                ws.Range(ws.Cells(r, 1), ws.Cells(r, FB_COL_COMMENTS)) _
                  .Interior.Color = CLR_BLINE
            End If
        Next r
    End If

    ws.Activate
    ActiveWindow.FreezePanes = False
    ws.Rows(2).Select
    ActiveWindow.FreezePanes = True

    With ws.PageSetup
        .Orientation = xlLandscape
        .PaperSize = xlPaperA4
        .PrintTitleRows = "$1:$1"
        .CenterFooter = "&P/&N"
        .RightHeader = ThisWorkbook.Name
        .Zoom = False
        .FitToPagesWide = 1
        .FitToPagesTall = False
    End With

    ws.Cells(1, 1).Select
    MsgBox (lastRow - 1) & " plots written to the Fieldbook." & vbCrLf & vbCrLf & _
           "Duplex short-edge is a printer setting — choose it in the print " & _
           "dialog; Excel cannot set it from here.", vbInformation, "Fieldbook"
End Sub


'==============================================================================
'  v2 — QR to split lookup
'  The payload is plot,material,inbred,source,cms,gen,comments — the plot is
'  the first field. The format is NOT changed here: already-printed labels
'  must keep working.
'==============================================================================
Public Sub btnResolveSplitFromQR()
    Const COL_QR As Long = 6
    Const COL_SPLIT As Long = 9

    Dim ws As Worksheet: Set ws = GetSheet(SHEET_REPLACEMENTS_AND_ERRORS)
    If ws Is Nothing Then Exit Sub

    Dim nDates As Long: nDates = PlantingDateCount()
    Dim lastRow As Long, r As Long, unresolved As Long, resolved As Long
    lastRow = ws.Cells(ws.Rows.Count, COL_QR).End(xlUp).Row

    For r = 6 To lastRow
        Dim payload As String, plot As String
        payload = S(ws.Cells(r, COL_QR).Value)
        If Len(payload) > 0 Then
            plot = Trim(Split(payload, ",")(0))
            Dim s As Long, found As Long: found = 0
            For s = 1 To nDates
                Dim pp As Worksheet: Set pp = GetSheet(PacketPrepName(s))
                If Not pp Is Nothing Then
                    Dim hit As Range
                    Set hit = pp.Columns(4).Find(What:=plot, LookAt:=xlWhole)
                    If Not hit Is Nothing Then
                        found = s
                        Exit For
                    End If
                End If
            Next s
            If found > 0 Then
                ws.Cells(r, COL_SPLIT).Value = found
                resolved = resolved + 1
            Else
                unresolved = unresolved + 1
            End If
        End If
    Next r

    If unresolved > 0 Then
        MsgBox resolved & " scan(s) resolved." & vbCrLf & unresolved & _
               " QR value(s) matched no plot in any Packet Prep tab — those " & _
               "rows were left blank.", vbExclamation, "Unresolved scans"
    Else
        MsgBox resolved & " scan(s) resolved to a split.", vbInformation
    End If
End Sub


'==============================================================================
'  v2 — Date recording ordering buttons (AB only)
'==============================================================================
Public Sub btnRecordByRangeAndPullBags()
    SortDateRecording False
End Sub

Public Sub btnRecord2RowsTogether()
    SortDateRecording True
End Sub

Private Sub SortDateRecording(ByVal byBand As Boolean)
    Const COL_GROUP As Long = 1
    Const COL_RANGE As Long = 2
    Const COL_ROW As Long = 3
    Const COL_LAST As Long = 11
    Const COL_KEY As Long = 13            ' scratch

    Dim ws As Worksheet: Set ws = GetSheet(SHEET_DATE_RECORDING)
    If ws Is Nothing Then Exit Sub

    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, COL_RANGE).End(xlUp).Row
    If lastRow < 2 Then
        MsgBox "No data on '" & SHEET_DATE_RECORDING & "' yet.", vbInformation
        Exit Sub
    End If

    Dim r As Long, rangeNo As Long, rowNo As Long
    For r = 2 To lastRow
        rangeNo = CLng(Val(ws.Cells(r, COL_RANGE).Value))
        rowNo = CLng(Val(ws.Cells(r, COL_ROW).Value))
        ws.Cells(r, COL_KEY).Value = SerpentineKey(rangeNo, rowNo)
    Next r

    With ws.Sort
        .SortFields.Clear
        If byBand Then
            ' Walk each two-row band end to end before moving on.
            .SortFields.Add Key:=ws.Range(ws.Cells(2, COL_GROUP), _
                                          ws.Cells(lastRow, COL_GROUP)), _
                            Order:=xlAscending
        End If
        .SortFields.Add Key:=ws.Range(ws.Cells(2, COL_RANGE), _
                                      ws.Cells(lastRow, COL_RANGE)), _
                        Order:=xlAscending
        .SortFields.Add Key:=ws.Range(ws.Cells(2, COL_KEY), _
                                      ws.Cells(lastRow, COL_KEY)), _
                        Order:=xlAscending
        .SetRange ws.Range(ws.Cells(2, 1), ws.Cells(lastRow, COL_KEY))
        .Header = xlNo
        .Apply
    End With

    ws.Columns(COL_KEY).Clear

    If byBand Then
        MsgBox "Ordered 2 rows together, snaking by range within each band.", _
               vbInformation, "Date recording"
    Else
        MsgBox "Ordered by range, snaking through rows.", _
               vbInformation, "Date recording"
    End If
End Sub


Private Function CountRows(ByVal sheetName As String) As Long
    Dim ws As Worksheet: Set ws = GetSheet(sheetName)
    If ws Is Nothing Then Exit Function
    Dim lr As Long: lr = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    If lr < 6 Then CountRows = 0 Else CountRows = lr - 5
End Function

Private Function CsvEsc(ByVal s As String) As String
    If InStr(s, ",") > 0 Or InStr(s, """") > 0 Or InStr(s, vbLf) > 0 Then
        CsvEsc = """" & Replace(s, """", """""") & """"
    Else
        CsvEsc = s
    End If
End Function

Private Function FileExists(ByVal p As String) As Boolean
    On Error Resume Next
    FileExists = (Dir(p) <> "")
End Function

Private Sub EnsureFolder(ByVal folderPath As String)
    Dim sep As String: sep = Application.PathSeparator
    Dim parts() As String: parts = Split(folderPath, sep)
    Dim acc As String, i As Long
    If Len(folderPath) > 0 And Left$(folderPath, 1) = sep Then acc = sep
    For i = LBound(parts) To UBound(parts)
        If Len(parts(i)) > 0 Then
            acc = acc & parts(i) & sep
            If Dir(acc, vbDirectory) = "" Then
                On Error Resume Next
                MkDir acc
                On Error GoTo 0
            End If
        End If
    Next i
End Sub
