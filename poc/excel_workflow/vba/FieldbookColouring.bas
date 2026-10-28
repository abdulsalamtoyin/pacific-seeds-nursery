Attribute VB_Name = "FieldbookColouring"
'==============================================================================
'  Fieldbook colouring for repeating values
'
'  Supplied by the client in 'Fieldbook colouring for repeating values VBA.docx'
'  and kept essentially verbatim so it stays recognisable to them.
'
'  Usage: filter the Fieldbook so only the rows of interest are visible, then
'  run StartColouring. Pick the column to group by, tick the columns to
'  colour, and apply. Only visible rows are touched.
'==============================================================================
Option Explicit

Public TargetColumn As Long
Public SelectedColumns As Collection

' Run this to start
Public Sub StartColouring()
    frmSelectColumns.Show
End Sub

' Colours ONLY VISIBLE (filtered) rows, and ONLY selected columns
Public Sub ApplyColouring_VisibleOnly()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim dict As Object
    Dim colourIndex As Long
    Dim cell As Range
    Dim key As Variant
    Dim col As Variant
    Dim rngKeyCol As Range
    Dim rngVisible As Range

    Set ws = ActiveSheet
    Set dict = CreateObject("Scripting.Dictionary")

    If TargetColumn = 0 Then
        MsgBox "Target column not set. Please choose a column in the form.", vbExclamation
        Exit Sub
    End If

    If SelectedColumns Is Nothing Or SelectedColumns.Count = 0 Then
        MsgBox "No columns selected to colour.", vbExclamation
        Exit Sub
    End If

    lastRow = ws.Cells(ws.Rows.Count, TargetColumn).End(xlUp).Row
    If lastRow < 2 Then
        MsgBox "No data found (need data starting from row 2).", vbExclamation
        Exit Sub
    End If

    Application.ScreenUpdating = False

    '-----------------------------
    ' 1) Clear existing fill ONLY for VISIBLE rows in the selected columns
    '-----------------------------
    On Error Resume Next
    For Each col In SelectedColumns
        Set rngVisible = ws.Range(ws.Cells(2, CLng(col)), ws.Cells(lastRow, CLng(col))).SpecialCells(xlCellTypeVisible)
        If Not rngVisible Is Nothing Then rngVisible.Interior.Pattern = xlNone
        Set rngVisible = Nothing
    Next col
    On Error GoTo 0

    '-----------------------------
    ' 2) Loop ONLY visible cells in the grouping column
    '-----------------------------
    Set rngKeyCol = ws.Range(ws.Cells(2, TargetColumn), ws.Cells(lastRow, TargetColumn))
    On Error Resume Next
    Set rngKeyCol = rngKeyCol.SpecialCells(xlCellTypeVisible)
    On Error GoTo 0

    If rngKeyCol Is Nothing Then
        Application.ScreenUpdating = True
        MsgBox "No visible rows found (filter may be hiding all rows).", vbInformation
        Exit Sub
    End If

    colourIndex = 0
    For Each cell In rngKeyCol.Cells
        key = cell.Value
        If Len(key) > 0 Then
            If Not dict.Exists(CStr(key)) Then
                dict.Add CStr(key), GetMutedColour(colourIndex)
                colourIndex = colourIndex + 1
            End If
            ' Colour ONLY selected columns on THIS visible row
            For Each col In SelectedColumns
                ws.Cells(cell.Row, CLng(col)).Interior.Color = dict(CStr(key))
            Next col
        End If
    Next cell

    Application.ScreenUpdating = True
    MsgBox "Done. Coloured visible (filtered) rows only.", vbInformation
End Sub

' Muted pastel palette (safe for black text)
Public Function GetMutedColour(i As Long) As Long
    Dim colours As Variant
    colours = Array( _
        RGB(242, 242, 242), _
        RGB(231, 240, 255), _
        RGB(237, 248, 241), _
        RGB(255, 245, 227), _
        RGB(245, 238, 255), _
        RGB(255, 236, 239), _
        RGB(240, 250, 255), _
        RGB(250, 250, 235))
    GetMutedColour = colours(i Mod (UBound(colours) + 1))
End Function
