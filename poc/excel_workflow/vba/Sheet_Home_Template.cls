VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet_Home_Template"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
'==============================================================================
'  Pacific Seeds — Home sheet event handlers (Nursery_Template)
'  Paste this code into the "Home" sheet's class module:
'     VBE > Project Explorer > Microsoft Excel Objects > Sheet "Home"
'  (Don't import as a regular module — it must live on the sheet.)
'==============================================================================
Option Explicit

Private Sub Worksheet_BeforeDoubleClick(ByVal Target As Range, Cancel As Boolean)
    ' If they double-clicked a styled button cell (rows that have a macro name in col G),
    ' run that macro and cancel the default edit behaviour.
    Dim macroName As String
    macroName = Trim$(CStr(Me.Cells(Target.Row, 7).Value))
    If Len(macroName) > 0 And Target.Column <= 5 Then
        Cancel = True
        RunMacroByCell Target
    End If
End Sub
