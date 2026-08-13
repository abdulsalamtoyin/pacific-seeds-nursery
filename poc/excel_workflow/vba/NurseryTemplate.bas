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

' Tab names (must match build_workbooks.py)
Public Const SHEET_HOME              As String = "Home"
Public Const SHEET_SETTINGS          As String = "Settings"
Public Const SHEET_NURSERY_SITE      As String = "Nursery site"
Public Const SHEET_NURSERY_DATA      As String = "Nursery data"
Public Const SHEET_MAP               As String = "Map"
Public Const SHEET_MATERIAL_MAP      As String = "Material Map"
Public Const SHEET_PACKET_PREP       As String = "Packet Prep"
Public Const SHEET_NURSERY_LIST      As String = "Nursery list"
Public Const SHEET_FIELDBOOK         As String = "Fieldbook"
Public Const SHEET_REPLACEMENTS      As String = "Replacements done"
Public Const SHEET_PLANTING_ERRORS   As String = "Planting error noted"
Public Const SHEET_BC0_LABELS        As String = "BC0 labels"
Public Const SHEET_DATE_RECORDING    As String = "Date recording"
Public Const SHEET_PULLING_BAGS      As String = "Pulling bags"
Public Const SHEET_OPERATIONS        As String = "Operations"
Public Const SHEET_COMMENTS          As String = "Comments"
Public Const SHEET_BC0_TFMSA         As String = "BC0 TFMSA record"
Public Const SHEET_TFMSA_SPRAY       As String = "TFMSA Spray plots"
Public Const SHEET_HY_HEIGHTS        As String = "Hy Heights"

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

    ' Build each tab
    BuildPacketPrep p, CStr(GetSetting("Nursery code"))
    BuildMap p, SHEET_MAP, 10              ' value index 10 = Hybrid Code
    BuildMap p, SHEET_MATERIAL_MAP, 3      ' value index 3  = Material ID
    BuildNurseryList p
    BuildFieldbook p
    BuildBC0Labels p, CStr(GetSetting("Nursery code"))
    BuildDateRecording p, SHEET_DATE_RECORDING
    BuildDateRecording p, SHEET_PULLING_BAGS
    BuildTFMSASprayPlots p
    BuildHyHeights p

    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    Application.CalculateFull

    MsgBox nRows & " packets read from PRISM." & vbCrLf & vbCrLf & _
           "All 10 data tabs generated:" & vbCrLf & _
           "  • Packet Prep (25 cols, QR text + colored digits)" & vbCrLf & _
           "  • Map (Hybrid Code grid)" & vbCrLf & _
           "  • Material Map (Material ID grid)" & vbCrLf & _
           "  • Nursery list, Fieldbook, BC0 labels" & vbCrLf & _
           "  • Date recording, Pulling bags" & vbCrLf & _
           "  • TFMSA Spray plots, Hy Heights" & vbCrLf & vbCrLf & _
           "The Packet Prep tab has the QR CODE text payload column — " & _
           "send this workbook to the barcode-printing machine.", _
           vbInformation, "All tabs generated"
End Sub


'==============================================================================
'  Packet Prep — 25 cols, QR text + colored digit columns
'==============================================================================
Private Sub BuildPacketPrep(ByRef p As Variant, ByVal nurseryCode As String)
    Dim ws As Worksheet: Set ws = GetSheet(SHEET_PACKET_PREP)
    If ws Is Nothing Then Exit Sub
    ws.Range("A6:Z" & ws.Rows.Count).Clear

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

    ' Output rows
    Dim outRow As Long: outRow = 6
    For i = 1 To n
        Dim plot As String: plot = CLng(p(i, 1)) & "_" & CLng(p(i, 2))
        ws.Cells(outRow, 1).Value = QRText(plot, p(i, 3), p(i, 4), p(i, 5), p(i, 6), p(i, 7), p(i, 8))
        ws.Cells(outRow, 2).Value = p(i, 1)       ' Range
        ws.Cells(outRow, 3).Value = p(i, 2)       ' Row
        ws.Cells(outRow, 4).Value = plot           ' Plot
        ws.Cells(outRow, 5).Value = spike(i)       ' SPIKE#
        ws.Cells(outRow, 6).Value = rack(i)        ' RACK ORDER
        ' Colored digit columns (cols 7-10)
        Dim digits As String: digits = Format$(rack(i), "0000")
        ws.Cells(outRow, 7).Value = Mid$(digits, 1, 1)
        ws.Cells(outRow, 7).Font.Color = CLR_DIGIT_THOUSANDS
        ws.Cells(outRow, 7).Font.Bold = True
        ws.Cells(outRow, 8).Value = Mid$(digits, 2, 1)
        ws.Cells(outRow, 8).Font.Color = CLR_DIGIT_HUNDREDS
        ws.Cells(outRow, 8).Font.Bold = True
        ws.Cells(outRow, 9).Value = Mid$(digits, 3, 1)
        ws.Cells(outRow, 9).Font.Color = CLR_DIGIT_TENS
        ws.Cells(outRow, 9).Font.Bold = True
        ws.Cells(outRow, 10).Value = Mid$(digits, 4, 1)
        ws.Cells(outRow, 10).Font.Color = CLR_DIGIT_ONES
        ws.Cells(outRow, 10).Font.Bold = True
        ' Remaining PRISM cols (Material ID through Entry #) → cols 11-25
        ws.Cells(outRow, 11).Value = p(i, 3)      ' Material ID
        ws.Cells(outRow, 12).Value = p(i, 4)      ' Inbred Code
        ws.Cells(outRow, 13).Value = p(i, 5)      ' Source ID
        ws.Cells(outRow, 14).Value = p(i, 6)      ' CMS reaction
        ws.Cells(outRow, 15).Value = p(i, 7)      ' Generation
        ws.Cells(outRow, 16).Value = p(i, 8)      ' Comments
        ws.Cells(outRow, 17).Value = p(i, 9)      ' Pedigree
        ws.Cells(outRow, 18).Value = p(i, 10)     ' Hybrid Code
        ws.Cells(outRow, 19).Value = p(i, 11)     ' Trait Name
        ws.Cells(outRow, 20).Value = p(i, 12)     ' Plant #
        ws.Cells(outRow, 21).Value = p(i, 13)     ' Loc Seq#
        ws.Cells(outRow, 22).Value = p(i, 14)     ' SubSeq Flag
        ws.Cells(outRow, 23).Value = p(i, 15)     ' Entry Book Project
        ws.Cells(outRow, 24).Value = p(i, 16)     ' Entry Book Name
        ws.Cells(outRow, 25).Value = p(i, 17)     ' Entry #
        outRow = outRow + 1
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
Private Sub BuildMap(ByRef p As Variant, ByVal sheetName As String, _
                     ByVal valueIdx As Long)
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
            Dim v As String: v = S(p(i, valueIdx))
            If Len(v) > 14 Then v = Left$(v, 14)
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

    ' Column widths
    ws.Columns(1).ColumnWidth = 9
    For fr = 1 To maxRow
        ws.Columns(fr + 1).ColumnWidth = 11
    Next fr
    ws.Columns(maxRow + 2).ColumnWidth = 7
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
    Dim ws As Worksheet: Set ws = GetSheet(sheetName)
    If ws Is Nothing Then Exit Sub
    ws.Range("A6:Z" & ws.Rows.Count).Clear

    Dim n As Long: n = UBound(p, 1)
    Dim outR As Long: outR = 6
    Dim i As Long, gen As String
    For i = 1 To n
        gen = S(p(i, 7))
        If IsRecurrent(gen) Then
            ws.Cells(outR, 1).Value = p(i, 1)     ' Range
            ws.Cells(outR, 2).Value = p(i, 2)     ' Row
            ws.Cells(outR, 3).Value = CLng(p(i, 1)) & "_" & CLng(p(i, 2))   ' Plot
            ' cols 4 ("1") and 5 ("2") blank
            ws.Cells(outR, 6).Value = p(i, 3)     ' Material ID
            ws.Cells(outR, 7).Value = p(i, 5)     ' Source ID
            ws.Cells(outR, 8).Value = gen          ' Gen
            ws.Cells(outR, 9).Value = p(i, 6)     ' CMS
            ws.Cells(outR, 10).Value = p(i, 8)    ' Comments
            outR = outR + 1
        End If
    Next i
End Sub


'==============================================================================
'  TFMSA Spray plots — BC* only
'==============================================================================
Private Sub BuildTFMSASprayPlots(ByRef p As Variant)
    Dim ws As Worksheet: Set ws = GetSheet(SHEET_TFMSA_SPRAY)
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
    Dim ws As Worksheet: Set ws = GetSheet(SHEET_PACKET_PREP)
    If ws Is Nothing Then Exit Sub
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, 2).End(xlUp).Row
    If lastRow < 7 Then
        MsgBox "Run Step 2 first.", vbExclamation
        Exit Sub
    End If
    With ws.Sort
        .SortFields.Clear
        .SortFields.Add Key:=ws.Range(ws.Cells(6, 6), ws.Cells(lastRow, 6)), Order:=xlAscending
        .SortFields.Add Key:=ws.Range(ws.Cells(6, 5), ws.Cells(lastRow, 5)), Order:=xlAscending
        .SetRange ws.Range(ws.Cells(6, 1), ws.Cells(lastRow, 25))
        .Header = xlNo
        .Apply
    End With
    ws.Activate
    MsgBox "Packet Prep re-sorted in LSD-radix racking order (rack ↑, spike ↑).", _
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
        Case "replacement":     destSheet = SHEET_REPLACEMENTS
        Case "planting_error":  destSheet = SHEET_PLANTING_ERRORS
        Case "spray":           destSheet = SHEET_TFMSA_SPRAY
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
    reps = CountRows(SHEET_REPLACEMENTS)
    errs = CountRows(SHEET_PLANTING_ERRORS)
    sprays = CountRows(SHEET_TFMSA_SPRAY)

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
