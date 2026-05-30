VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet_Home_Hub"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
'==============================================================================
'  Pacific Seeds — Home sheet event handlers (Nursery_Hub)
'  Paste this code into the Hub's "Home" sheet class module.
'==============================================================================
Option Explicit

Private Sub Worksheet_BeforeDoubleClick(ByVal Target As Range, Cancel As Boolean)
    If Len(Trim$(CStr(Target.Cells(1, 1).Value))) = 0 Then Exit Sub
    Cancel = True
    hubRunMacroByCell Target
End Sub
