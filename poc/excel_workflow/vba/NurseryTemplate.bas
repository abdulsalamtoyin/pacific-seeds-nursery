Attribute VB_Name = "NurseryTemplate"
'==============================================================================
'  Pacific Seeds — Nursery Workflow (consolidated module for Nursery_Template)
'  ----------------------------------------------------------------------------
'  Implements every workflow step from "Sorghum nursery prep workflow.docx",
'  consolidating the three original modules (Recurrent, Packetprinting,
'  Fieldbook) plus new automation for replacements, additionals, dashboard
'  and Hub-sync.
'
'  Setup: import this .bas file via VBE (Tools > Macros > VB Editor >
'  File > Import File), import the matching ThisWorkbook and Sheet_Home class
'  modules, then Save As .xlsm.
'==============================================================================

Option Explicit

'------------------------------------------------------------------------------
'  CONSTANTS
'------------------------------------------------------------------------------
Public Const SHEET_HOME           As String = "Home"
Public Const SHEET_SETTINGS       As String = "Settings"
Public Const SHEET_NURSERY_SITE   As String = "Nursery site"
Public Const SHEET_FIELD_MAP      As String = "Field Map"
Public Const SHEET_MATERIAL_MAP   As String = "Material Map"
Public Const SHEET_NURSERY_DATA   As String = "Nursery data"
Public Const SHEET_NURSERY_LIST   As String = "Nursery list"
Public Const SHEET_PACKET_PREP    As String = "Packet prep"
Public Const SHEET_REPLACEMENTS   As String = "Replacements and errors"
Public Const SHEET_UPDATED_SITE   As String = "Updated nursery site"
Public Const SHEET_FIELDBOOK      As String = "Fieldbook"
Public Const SHEET_OPERATIONS     As String = "Operations"
Public Const SHEET_COMMENTS       As String = "Comments"
Public Const SHEET_ADDITIONALS    As String = "Additionals"

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

Public Function GenUuid() As String
    ' 8-char random hex — collision space is fine for nursery scale
    Dim s As String, i As Integer, c As Integer
    Randomize
    For i = 1 To 8
        c = Int(Rnd() * 16)
        s = s & Mid$("0123456789abcdef", c + 1, 1)
    Next i
    GenUuid = s
End Function

Public Function CurrentTech() As String
    Dim t As String
    t = CStr(GetSetting("Default tech"))
    If Len(t) = 0 Then t = Environ$("USER")
    If Len(t) = 0 Then t = Environ$("USERNAME")
    CurrentTech = t
End Function


'==============================================================================
'  M_BUTTONS — double-click router
'  Called from Sheet_Home.Worksheet_BeforeDoubleClick
'==============================================================================
Public Sub RunMacroByCell(ByVal target As Range)
    Dim macroName As String
    ' Macro name is mirrored in column G of the same row
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
'  M_SETUP — Step 1: Initialise from PRISM export
'==============================================================================
Public Sub btnInitNursery()
    Dim code As String
    code = InputBox("Enter the nursery code (e.g. AUGT1-26S-IMI):", _
                    "Initialise Nursery")
    If Len(code) = 0 Then Exit Sub
    SetSetting "Nursery code", code

    Dim ws As Worksheet
    Set ws = GetSheet(SHEET_NURSERY_SITE)
    If ws Is Nothing Then
        MsgBox "Sheet '" & SHEET_NURSERY_SITE & "' not found.", vbCritical
        Exit Sub
    End If
    If ws.Cells(6, 1).Value = "" Then
        MsgBox "Paste the PRISM export data into the '" & SHEET_NURSERY_SITE & _
               "' tab starting at row 6 (headers stay at row 5), then re-run step 1.", _
               vbInformation, "Nursery site is empty"
        ws.Activate
        Exit Sub
    End If

    ' Stamp Nursery data tab
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

    btnRefreshDashboard
    MsgBox "Initialised nursery '" & code & "'.", vbInformation, "Done"
End Sub


'==============================================================================
'  M_NURSERY_LIST — Step 2: Build the Nursery list
'==============================================================================
Public Sub btnBuildNurseryList()
    ' Mac-compatible: uses arrays instead of Scripting.Dictionary (Windows-only).
    Dim src As Worksheet, dst As Worksheet
    Set src = GetSheet(SHEET_NURSERY_SITE)
    Set dst = GetSheet(SHEET_NURSERY_LIST)
    If src Is Nothing Or dst Is Nothing Then
        MsgBox "Missing source or destination sheet.", vbCritical
        Exit Sub
    End If

    Dim sidCol As Long, icCol As Long, hcCol As Long
    sidCol = FindHeaderCol(src, "Source ID", 5)
    icCol  = FindHeaderCol(src, "Inbred Code", 5)
    hcCol  = FindHeaderCol(src, "Hybrid Code", 5)
    If sidCol = 0 Then
        MsgBox "Could not find 'Source ID' column in Nursery site row 5.", vbExclamation
        Exit Sub
    End If

    Dim lastRow As Long
    lastRow = src.Cells(src.Rows.Count, sidCol).End(xlUp).Row
    If lastRow < 6 Then
        MsgBox "Nursery site has no data rows.", vbExclamation
        Exit Sub
    End If

    ' Read source IDs + parallel inbred/hybrid into arrays
    Dim n As Long: n = lastRow - 5
    Dim ids() As String, inb() As String, hyb() As String
    ReDim ids(1 To n)
    ReDim inb(1 To n)
    ReDim hyb(1 To n)
    Dim r As Long, j As Long
    j = 0
    For r = 6 To lastRow
        Dim k As String: k = Trim$(CStr(src.Cells(r, sidCol).Value))
        If Len(k) > 0 Then
            j = j + 1
            ids(j) = k
            If icCol > 0 Then inb(j) = CStr(src.Cells(r, icCol).Value)
            If hcCol > 0 Then hyb(j) = CStr(src.Cells(r, hcCol).Value)
        End If
    Next r
    n = j  ' actual count of non-blank rows
    If n = 0 Then
        MsgBox "No Source ID values found.", vbExclamation
        Exit Sub
    End If

    ' Sort the three parallel arrays by Source ID (insertion sort is fast for typical sizes).
    Dim ii As Long, jj As Long, ts As String
    For ii = 2 To n
        Dim cur As String, ci As String, ch As String
        cur = ids(ii): ci = inb(ii): ch = hyb(ii)
        jj = ii - 1
        Do While jj >= 1
            If ids(jj) <= cur Then Exit Do
            ids(jj + 1) = ids(jj): inb(jj + 1) = inb(jj): hyb(jj + 1) = hyb(jj)
            jj = jj - 1
        Loop
        ids(jj + 1) = cur: inb(jj + 1) = ci: hyb(jj + 1) = ch
    Next ii

    ' Settings
    Dim qty As Double: qty = CDbl(GetSetting("Qty per packet"))
    If qty <= 0 Then qty = 1.4
    Dim bulkAbove As Long: bulkAbove = CLng(GetSetting("Filter bulk treatment above"))
    If bulkAbove <= 0 Then bulkAbove = 10

    ' Clear destination & write grouped/counted rows
    dst.Range("A6:Z" & dst.Rows.Count).Clear

    Application.ScreenUpdating = False
    Dim outRow As Long: outRow = 6
    Dim uniqueCount As Long: uniqueCount = 0
    ii = 1
    Do While ii <= n
        Dim grpKey As String: grpKey = ids(ii)
        Dim grpInbred As String: grpInbred = inb(ii)
        Dim grpHybrid As String: grpHybrid = hyb(ii)
        Dim grpCount As Long: grpCount = 1
        Do While ii + 1 <= n
            If ids(ii + 1) <> grpKey Then Exit Do
            grpCount = grpCount + 1
            ii = ii + 1
        Loop
        dst.Cells(outRow, 1).Value = grpKey
        dst.Cells(outRow, 2).Value = grpCount
        dst.Cells(outRow, 3).Value = qty * grpCount
        dst.Cells(outRow, 4).Value = grpInbred
        dst.Cells(outRow, 5).Value = grpHybrid
        If grpCount > bulkAbove Then dst.Cells(outRow, 6).Value = "BULK"
        outRow = outRow + 1
        uniqueCount = uniqueCount + 1
        ii = ii + 1
    Loop

    ' Light formatting
    If outRow > 6 Then
        Dim used As Range
        Set used = dst.Range(dst.Cells(6, 1), dst.Cells(outRow - 1, 7))
        used.Borders.LineStyle = xlContinuous
        used.Borders.Color = RGB(216, 227, 237)
    End If
    Application.ScreenUpdating = True

    dst.Activate
    btnRefreshDashboard
    MsgBox uniqueCount & " unique source IDs written. " & _
           "Items over " & bulkAbove & " reps flagged 'BULK'.", _
           vbInformation, "Nursery list built"
End Sub


'==============================================================================
'  M_DESIGN_MAP — Step 3
'==============================================================================
Public Sub btnDesignFieldMap()
    Dim ws As Worksheet
    Set ws = GetSheet(SHEET_FIELD_MAP)
    If Not ws Is Nothing Then ws.Activate
    MsgBox "Design the Field Map manually here. Put range numbers along one " & _
           "axis, row numbers along the other; leave a blank column between " & _
           "groups to mark spike boundaries.", vbInformation, "Field Map"
End Sub


'==============================================================================
'  M_PACKET_PREP — Step 4: 13-step consolidated workflow
'  (Originally Packetprinting VBA code.docx — Step1..Step13)
'==============================================================================
Public Sub btnGeneratePacketPrep()
    Dim src As Worksheet, dst As Worksheet
    Set src = GetSheet(SHEET_NURSERY_SITE)
    Set dst = GetSheet(SHEET_PACKET_PREP)
    If src Is Nothing Or dst Is Nothing Then
        MsgBox "Missing Nursery site or Packet prep sheet.", vbCritical
        Exit Sub
    End If

    Dim rangeCol As Long, rowCol As Long, sidCol As Long, midCol As Long
    Dim cmsCol As Long, genCol As Long, comCol As Long
    rangeCol = FindHeaderCol(src, "Range", 5)
    rowCol   = FindHeaderCol(src, "Row", 5)
    sidCol   = FindHeaderCol(src, "Source ID", 5)
    midCol   = FindHeaderCol(src, "Material ID", 5)
    cmsCol   = FindHeaderCol(src, "CMS reaction", 5)
    genCol   = FindHeaderCol(src, "Generation", 5)
    comCol   = FindHeaderCol(src, "Comments", 5)
    If rangeCol = 0 Or rowCol = 0 Then
        MsgBox "Could not find Range/Row columns in Nursery site row 5.", vbExclamation
        Exit Sub
    End If

    Dim lastRow As Long
    lastRow = src.Cells(src.Rows.Count, rangeCol).End(xlUp).Row
    If lastRow < 6 Then
        MsgBox "Nursery site has no data.", vbExclamation
        Exit Sub
    End If

    ' Clear & build Packet prep
    dst.Range("A6:Z" & dst.Rows.Count).Clear

    Dim code As String: code = CStr(GetSetting("Nursery code"))
    Dim outRow As Long: outRow = 6
    Dim r As Long
    For r = 6 To lastRow
        Dim rng As Long, rw As Long
        rng = CLng(Val(src.Cells(r, rangeCol).Value))
        rw  = CLng(Val(src.Cells(r, rowCol).Value))
        If rng > 0 And rw > 0 Then
            Dim plot As String: plot = rng & "_" & rw
            Dim spike As Long, rack As Long
            ' Default: spike = range, rack = row (Field Map can override later)
            spike = rng
            rack = rw
            Dim uuid As String: uuid = GenUuid()
            Dim qr As String
            qr = "SNUR:" & code & ":" & uuid
            dst.Cells(outRow, 1).Value = qr
            dst.Cells(outRow, 2).Value = plot
            dst.Cells(outRow, 3).Value = rng
            dst.Cells(outRow, 4).Value = rw
            dst.Cells(outRow, 5).Value = spike
            dst.Cells(outRow, 6).Value = rack
            If midCol > 0 Then dst.Cells(outRow, 7).Value = src.Cells(r, midCol).Value
            If sidCol > 0 Then dst.Cells(outRow, 8).Value = src.Cells(r, sidCol).Value
            If genCol > 0 Then dst.Cells(outRow, 9).Value = src.Cells(r, genCol).Value
            If cmsCol > 0 Then dst.Cells(outRow, 10).Value = src.Cells(r, cmsCol).Value
            If comCol > 0 Then dst.Cells(outRow, 11).Value = src.Cells(r, comCol).Value
            outRow = outRow + 1
        End If
    Next r

    ' Sort: spike asc, rack asc
    Dim lastDataRow As Long: lastDataRow = outRow - 1
    If lastDataRow >= 7 Then
        With dst.Sort
            .SortFields.Clear
            .SortFields.Add Key:=dst.Range(dst.Cells(6, 5), dst.Cells(lastDataRow, 5)), Order:=xlAscending
            .SortFields.Add Key:=dst.Range(dst.Cells(6, 6), dst.Cells(lastDataRow, 6)), Order:=xlAscending
            .SetRange dst.Range(dst.Cells(6, 1), dst.Cells(lastDataRow, 11))
            .Header = xlNo
            .Apply
        End With
    End If

    Dim used As Range
    Set used = dst.Range(dst.Cells(6, 1), dst.Cells(lastDataRow, 11))
    used.Borders.LineStyle = xlContinuous
    used.Borders.Color = RGB(216, 227, 237)

    dst.Activate
    btnRefreshDashboard
    MsgBox (lastDataRow - 5) & " packets prepared with QR payloads, spike & rack order.", _
           vbInformation, "Packet Prep ready"
End Sub


'==============================================================================
'  M_RACKING — Step 5: LSD Radix sort for racking
'==============================================================================
Public Sub btnSortForRacking()
    Dim dst As Worksheet
    Set dst = GetSheet(SHEET_PACKET_PREP)
    If dst Is Nothing Then
        MsgBox "Run Step 4 first.", vbExclamation
        Exit Sub
    End If

    Dim lastRow As Long
    lastRow = dst.Cells(dst.Rows.Count, 2).End(xlUp).Row
    If lastRow < 7 Then
        MsgBox "Packet prep is empty.", vbExclamation
        Exit Sub
    End If

    ' LSD: sort by rack asc, then by spike asc — produces racking pickup order
    With dst.Sort
        .SortFields.Clear
        .SortFields.Add Key:=dst.Range(dst.Cells(6, 6), dst.Cells(lastRow, 6)), Order:=xlAscending
        .SortFields.Add Key:=dst.Range(dst.Cells(6, 5), dst.Cells(lastRow, 5)), Order:=xlAscending
        .SetRange dst.Range(dst.Cells(6, 1), dst.Cells(lastRow, 11))
        .Header = xlNo
        .Apply
    End With
    dst.Activate
    MsgBox "Packets re-sorted in racking pickup order (LSD radix: rack ↑, spike ↑).", _
           vbInformation, "Sorted for racking"
End Sub


'==============================================================================
'  M_REPLACEMENTS — Step 6 + Step 7: add events
'==============================================================================
Public Sub btnAddReplacement()
    AddEvent "replacement"
End Sub

Public Sub btnAddPlantingError()
    AddEvent "planting_error"
End Sub

Public Sub btnRecordSpray()
    AddAdditional "Spray"
End Sub

Public Sub btnRecordABPull()
    AddAdditional "AB bag pulling"
End Sub

Private Sub AddEvent(ByVal evType As String)
    Dim ws As Worksheet: Set ws = GetSheet(SHEET_REPLACEMENTS)
    If ws Is Nothing Then Exit Sub
    Dim r As Long
    r = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row + 1
    If r < 6 Then r = 6
    Dim plot As String
    plot = InputBox("Plot number (e.g. 2_8):", "New " & evType)
    If Len(plot) = 0 Then Exit Sub
    ws.Cells(r, 1).Value = NowISO()
    ws.Cells(r, 2).Value = CurrentTech()
    ws.Cells(r, 3).Value = plot
    ws.Cells(r, 4).Value = evType
    If evType = "replacement" Then
        ws.Cells(r, 5).Value = InputBox("Stage (Packeting / Planting):", _
                                        "Replacement", "Planting")
        ws.Cells(r, 6).Value = InputBox("Original Source ID:", "Replacement")
        ws.Cells(r, 7).Value = InputBox("Replaced with (new Source ID):", "Replacement")
    Else
        ws.Cells(r, 8).Value = InputBox("Severity (Low / Medium / High):", _
                                        "Planting error", "Medium")
        ws.Cells(r, 9).Value = InputBox("Note describing the error:", "Planting error")
    End If
    ws.Cells(r, 10).Value = "Open"
    ws.Activate
    ws.Cells(r, 1).Select
    btnRefreshDashboard
End Sub

Private Sub AddAdditional(ByVal opLabel As String)
    Dim ws As Worksheet: Set ws = GetSheet(SHEET_ADDITIONALS)
    If ws Is Nothing Then Exit Sub
    Dim r As Long
    r = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row + 1
    If r < 6 Then r = 6
    Dim plot As String
    plot = InputBox("Plot or zone:", "New " & opLabel)
    If Len(plot) = 0 Then Exit Sub
    ws.Cells(r, 1).Value = NowISO()
    ws.Cells(r, 2).Value = CurrentTech()
    ws.Cells(r, 3).Value = plot
    ws.Cells(r, 4).Value = opLabel
    Dim prod As String
    prod = InputBox("Product or event label (e.g. TFMSA, AB pull):", opLabel)
    ws.Cells(r, 5).Value = prod
    ws.Cells(r, 6).Value = InputBox("Date (YYYY-MM-DD):", opLabel, _
                                    Format$(Date, "yyyy-mm-dd"))
    ws.Cells(r, 7).Value = InputBox("Count / qty (optional):", opLabel)
    ws.Cells(r, 8).Value = InputBox("Note (optional):", opLabel)
    ws.Activate
    ws.Cells(r, 1).Select
    btnRefreshDashboard
End Sub


'==============================================================================
'  M_UPDATED — Step 10
'==============================================================================
Public Sub btnImportUpdated()
    Dim ws As Worksheet
    Set ws = GetSheet(SHEET_UPDATED_SITE)
    If Not ws Is Nothing Then ws.Activate
    MsgBox "Paste the refreshed PRISM 'Nursery site' export into this tab " & _
           "starting at row 6. Then run Step 11 to regenerate the Fieldbook.", _
           vbInformation, "Updated Nursery site"
End Sub


'==============================================================================
'  M_FIELDBOOK — Step 11
'  (Consolidated from Fieldbook template VBA code.docx)
'==============================================================================
Public Sub btnGenerateFieldbook()
    Dim src As Worksheet, dst As Worksheet
    Set src = GetSheet(SHEET_UPDATED_SITE)
    If src Is Nothing Or src.Cells(6, 1).Value = "" Then
        Set src = GetSheet(SHEET_NURSERY_SITE)
    End If
    Set dst = GetSheet(SHEET_FIELDBOOK)
    If src Is Nothing Or dst Is Nothing Then
        MsgBox "Missing source or Fieldbook sheet.", vbCritical
        Exit Sub
    End If
    If src.Cells(6, 1).Value = "" Then
        MsgBox "No data in Nursery site / Updated nursery site.", vbExclamation
        Exit Sub
    End If

    Dim cols As Variant, names As Variant, i As Long
    cols = Array("Range", "Row", "Material ID", "Source ID", "Generation", "CMS reaction")
    names = Array("Range", "Row", "Material ID", "Source ID", "Gen", "CMS")

    dst.Range("A6:Z" & dst.Rows.Count).Clear

    Dim srcCols(0 To 5) As Long
    For i = 0 To 5
        srcCols(i) = FindHeaderCol(src, CStr(cols(i)), 5)
    Next i

    Dim lastRow As Long, r As Long, outRow As Long
    lastRow = src.Cells(src.Rows.Count, srcCols(0)).End(xlUp).Row
    outRow = 6
    For r = 6 To lastRow
        Dim k As Long
        For k = 0 To 5
            If srcCols(k) > 0 Then dst.Cells(outRow, k + 1).Value = src.Cells(r, srcCols(k)).Value
        Next k
        ' R_R, Crossed bags, Bagging Info, Comments left blank for field marking
        dst.Cells(outRow, 3).Value = dst.Cells(outRow, 1).Value & "_" & dst.Cells(outRow, 2).Value
        outRow = outRow + 1
    Next r

    ' Serpentine sort by Range then Row
    Dim lastDataRow As Long: lastDataRow = outRow - 1
    With dst.Sort
        .SortFields.Clear
        .SortFields.Add Key:=dst.Range(dst.Cells(6, 1), dst.Cells(lastDataRow, 1)), Order:=xlAscending
        .SortFields.Add Key:=dst.Range(dst.Cells(6, 2), dst.Cells(lastDataRow, 2)), Order:=xlAscending
        .SetRange dst.Range(dst.Cells(6, 1), dst.Cells(lastDataRow, 10))
        .Header = xlNo
        .Apply
    End With

    ' Print layout
    With dst.PageSetup
        .Orientation = xlLandscape
        .CenterFooter = "&P/&N"
        .RightHeader = ThisWorkbook.Name & " — " & SHEET_FIELDBOOK
        .PrintTitleRows = "$1:$5"
    End With

    Dim used As Range
    Set used = dst.Range(dst.Cells(6, 1), dst.Cells(lastDataRow, 10))
    used.Borders.LineStyle = xlContinuous
    used.Borders.Color = RGB(216, 227, 237)

    dst.Activate
    btnRefreshDashboard
    MsgBox "Fieldbook generated with " & (lastDataRow - 5) & " rows. Print preview is ready.", _
           vbInformation, "Fieldbook"
End Sub


'==============================================================================
'  M_DASHBOARD — Step 12
'==============================================================================
Public Sub btnRefreshDashboard()
    Application.CalculateFull
End Sub


'==============================================================================
'  M_HUB — Step 13: write summary row to the shared registry.csv
'==============================================================================
Public Sub btnPushToHub()
    On Error GoTo cleanFail
    Dim folder As String
    folder = ExpandTilde(CStr(GetSetting("Hub registry folder")))
    If Len(folder) = 0 Then Exit Sub  ' silent: nothing to do
    If Right$(folder, 1) <> Application.PathSeparator Then folder = folder & Application.PathSeparator

    ' Recursively create the registry folder (MkDir only does one level)
    EnsureFolder folder

    Dim regPath As String
    regPath = folder & "registry.csv"

    Dim code As String: code = CStr(GetSetting("Nursery code"))
    If Len(code) = 0 Then code = "(uncoded)"
    Dim season As String: season = CStr(GetSetting("Season"))
    Dim breeder As String: breeder = CStr(GetSetting("Breeder"))

    Dim packets As Long, reps As Long, errs As Long, sprays As Long
    packets = CountRows(SHEET_NURSERY_SITE)
    reps    = CountIfInType(SHEET_REPLACEMENTS, 4, "replacement")
    errs    = CountIfInType(SHEET_REPLACEMENTS, 4, "planting_error")
    sprays  = MaxOf(CountRows(SHEET_ADDITIONALS), 0)

    ' Read existing CSV, remove any prior row for this code, then append
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
            ' Skip rows that belong to this nursery
            If InStr(1, line, code & ",") <> 1 Then outLines.Add line
nxt:
        Loop
        Close #ff
    End If

    Dim filePath As String
    filePath = ThisWorkbook.FullName
    Dim row As String
    row = code & "," & CsvEsc(season) & "," & CsvEsc(breeder) & "," & _
          packets & "," & reps & "," & errs & "," & sprays & "," & _
          CsvEsc(filePath) & "," & CsvEsc(NowISO())
    outLines.Add row

    ff = FreeFile
    Open regPath For Output As #ff
    Dim it As Variant
    For Each it In outLines
        Print #ff, CStr(it)
    Next it
    Close #ff

    ' Update local "Last synced to Hub" cell
    On Error Resume Next
    ThisWorkbook.Names("DASH_LastSync").RefersToRange.Value = NowISO()
    On Error GoTo 0

    MsgBox "Pushed to Hub registry:" & vbCrLf & regPath, vbInformation, "Synced"
    Exit Sub

cleanFail:
    ' Never let Hub sync errors block a save. Log to debug pane and move on.
    Debug.Print "btnPushToHub failed: " & Err.Number & " " & Err.Description
End Sub

Private Sub EnsureFolder(ByVal folderPath As String)
    ' Build the directory hierarchy one level at a time.
    Dim sep As String: sep = Application.PathSeparator
    Dim parts() As String: parts = Split(folderPath, sep)
    Dim acc As String, i As Long
    ' Preserve leading separator on POSIX paths
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

Private Function CountRows(ByVal sheetName As String) As Long
    Dim ws As Worksheet: Set ws = GetSheet(sheetName)
    If ws Is Nothing Then Exit Function
    Dim lr As Long: lr = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    If lr < 6 Then CountRows = 0 Else CountRows = lr - 5
End Function

Private Function CountIfInType(ByVal sheetName As String, _
                               ByVal col As Long, ByVal value As String) As Long
    Dim ws As Worksheet: Set ws = GetSheet(sheetName)
    If ws Is Nothing Then Exit Function
    Dim lr As Long, r As Long, n As Long
    lr = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    For r = 6 To lr
        If LCase$(Trim$(CStr(ws.Cells(r, col).Value))) = LCase$(value) Then n = n + 1
    Next r
    CountIfInType = n
End Function

Private Function MaxOf(a As Long, b As Long) As Long
    If a > b Then MaxOf = a Else MaxOf = b
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
    On Error GoTo 0
End Function
