Attribute VB_Name = "Prints"
Sub ecran(ByVal Message As String)

    Call masque

    MsgBox Message
    
    Call affiche
    
End Sub

Sub detail(ByVal info As String)
    Sheets("Feuil1").Cells(10, 10) = Sheets("Feuil1").Cells(10, 10) & Chr(10) & info
    details.TextBox1.Value = Sheets("Feuil1").Cells(10, 10)
End Sub

