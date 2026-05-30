Attribute VB_Name = "FixTemplate"
'==============================================================================
'  One-shot Template repair
'  ---------------------------------------------------------------------
'  Paste into any blank module in Nursery_Template.xlsm, press F5, save.
'  Removes all orphan standard modules and re-imports NurseryTemplate.bas
'  from disk so the template is in a known-good state.
'==============================================================================
Option Explicit

Private Const NURSERY_BAS As String = _
    "/Users/toyinabdulsalam/Desktop/work/Sorghum/VBA codes - Sorghum nursery workflow/poc/excel_workflow/vba/NurseryTemplate.bas"

Public Sub FixTemplate()
    Dim vbp As Object: Set vbp = ThisWorkbook.VBProject
    Dim comp As Object
    Dim removed As Long, removedNames As String
    Dim toRemove As Collection: Set toRemove = New Collection

    ' Step 1: Collect every std module that isn't the canonical NurseryTemplate.
    For Each comp In vbp.VBComponents
        If comp.Type = 1 Then  ' vbext_ct_StdModule
            If LCase$(comp.Name) <> "nurserytemplate" And _
               LCase$(comp.Name) <> "fixtemplate" Then
                toRemove.Add comp
            End If
        End If
    Next comp

    ' Step 2: Remove orphans.
    Dim it As Variant
    For Each it In toRemove
        On Error Resume Next
        removedNames = removedNames & it.Name & ", "
        vbp.VBComponents.Remove it
        If Err.Number = 0 Then removed = removed + 1
        On Error GoTo 0
    Next it

    ' Step 3: Remove an existing NurseryTemplate (idempotent) then re-import.
    On Error Resume Next
    Dim existing As Object
    Set existing = vbp.VBComponents("NurseryTemplate")
    If Not existing Is Nothing Then vbp.VBComponents.Remove existing
    On Error GoTo 0

    If Dir(NURSERY_BAS) = "" Then
        MsgBox "Could not find:" & vbCrLf & NURSERY_BAS & vbCrLf & vbCrLf & _
               "Edit the NURSERY_BAS constant in this module if the path moved.", _
               vbCritical, "File not found"
        Exit Sub
    End If

    On Error GoTo importFail
    vbp.VBComponents.Import NURSERY_BAS
    On Error GoTo 0

    MsgBox "✅ Template repaired." & vbCrLf & vbCrLf & _
           "Removed orphan module(s): " & removed & vbCrLf & _
           IIf(Len(removedNames) > 0, "(" & Left$(removedNames, Len(removedNames) - 2) & ")", "") & vbCrLf & vbCrLf & _
           "Imported NurseryTemplate.bas (the workflow code)." & vbCrLf & vbCrLf & _
           "Now: ⌘S to save, then delete this FixTemplate module if you want.", _
           vbInformation, "Done"
    Exit Sub

importFail:
    MsgBox "Import of NurseryTemplate.bas failed:" & vbCrLf & _
           Err.Number & " — " & Err.Description, vbCritical, "Import failed"
End Sub
