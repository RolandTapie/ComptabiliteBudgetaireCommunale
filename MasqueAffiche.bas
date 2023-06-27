Attribute VB_Name = "MasqueAffiche"
Function affiche()
    Dim sh As Worksheet
    For Each sh In ThisWorkbook.Worksheets
        With sh
            If Not .Name = "Feuil1" Then .Visible = True
        End With
    Next
    'Call logging(Now, Application.UserName, Application.Caption, "Afficher", "MasqueAffiche.affiche")
End Function

Function masque()
    Dim sh As Worksheet
    For Each sh In ThisWorkbook.Worksheets
        With sh
            'If Not .Name = "Feuil1" Then .Visible = False
        End With
    Next
    'Call logging(Now, Application.UserName, Application.Caption, "Masquer", "MasqueAffiche.masque")
End Function

Sub affich_app()
    Application.Visible = True
End Sub

