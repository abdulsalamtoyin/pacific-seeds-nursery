VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook_Template"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
'==============================================================================
'  Paste into ThisWorkbook class module of Nursery_Template (.xlsm)
'==============================================================================
Option Explicit

Private Sub Workbook_Open()
    ' Auto-refresh the dashboard so opening the file shows live numbers.
    On Error Resume Next
    Application.Run "btnRefreshDashboard"
    On Error GoTo 0
End Sub

Private Sub Workbook_BeforeSave(ByVal SaveAsUI As Boolean, Cancel As Boolean)
    ' Auto-sync to Hub on save (best-effort; failures don't block).
    On Error Resume Next
    Application.Run "btnPushToHub"
    On Error GoTo 0
End Sub
