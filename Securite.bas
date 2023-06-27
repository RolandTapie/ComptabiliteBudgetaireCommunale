Attribute VB_Name = "Securite"
Function test_ID() As String
    While Count <= 3
        user = Application.InputBox("Votre identifiant :")
        If user = "" Then
            Count = Count + 1
        Else
            test_ID = user
            GoTo suite
        End If
    Wend
    Call exit_f
suite:
End Function

