Attribute VB_Name = "Maintenance"


Sub test_import()

Call download_code("Data1", "C:\Users\tallar\Documents\Data1.bas")
  
End Sub

Sub download_code(ByVal Module As String, ByVal Path_new_module As String)
    
    'TODO tester si module existe
    For i = 1 To ThisWorkbook.VBProject.VBComponents.Count
    Destination = "C:\Users\tallar\Documents\Exports Modules\"
    If ThisWorkbook.VBProject.VBComponents(i).Type = 1 Then
    ThisWorkbook.VBProject.VBComponents(i).Export (Destination & ThisWorkbook.VBProject.VBComponents(i).Name & ".bas")
    End If
    
    
    
        If ThisWorkbook.VBProject.VBComponents(i).Name = Module Then
            Set vbcomp = ActiveWorkbook.VBProject.VBComponents(Module)
            ThisWorkbook.VBProject.VBComponents.Remove vbcomp
            Set vbcomp = Nothing
            ThisWorkbook.VBProject.VBComponents.Import Path_new_module
            GoTo fin
        End If
    
    Next
    
fin:
End Sub

