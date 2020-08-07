Attribute VB_Name = "vbaSecurity"
Sub LockItUp()
' Locks up each sheet with give password

Dim strPW As String
Dim i As Integer

strPW = InputBox("Input password to lock workbook", "LockItUp")

With ActiveWorkbook
    For i = 1 To .Worksheets.Count
        .Worksheets(i).Protect strPW
    Next i
End With

End Sub

Sub DeleteAllVBAModules()
' !!!!!!!!!!!!!!!!!!!!!WARNING!!!!!!!!!!! Deletes all modules
' Please add reference to Microsoft Visual Basic for Applications Extensibility

Dim VBProj As VBIDE.VBProject
Dim VBComp As VBIDE.VBComponent
Dim CodeMod As VBIDE.CodeModule

Set VBProj = ActiveWorkbook.VBProject

For Each VBComp In VBProj.VBComponents
    If VBComp.Type = vbext_ct_Document Then
        Set CodeMod = VBComp.CodeModule
        With CodeMod
            .DeleteLines 1, .CountOfLines
        End With
    Else
        VBProj.VBComponents.Remove VBComp
    End If
Next VBComp

End Sub

