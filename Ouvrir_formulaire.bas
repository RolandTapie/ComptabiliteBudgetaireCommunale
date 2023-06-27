Attribute VB_Name = "Ouvrir_formulaire"
Function formulaire()
    UserForm1.Height = Application.Height
    UserForm1.Width = Application.Width
    UserForm1.Show
End Function

Function formulaire_mode_connexion()
    UserForm2.Height = 73
    UserForm2.Width = 286
    UserForm2.Show
End Function

