Attribute VB_Name = "Cryptage"
Function crypting(ByVal Texte As String) As String

    resultat1 = ""
    resultat2 = ""

    longueur = Len(Texte)

    For i = 1 To longueur
        correction = Asc(Mid(Texte, i, 1))
        If i Mod 2 = 0 Then
            correction = correction + 1
        Else
            correction = correction - 1
        End If
        resultat1 = resultat1 & Chr(correction)

    Next



    For i = 1 To longueur
        correction = Asc(Mid(resultat1, i, 1))
        If i Mod 2 = 0 Then
            correction = correction - 1
        Else
            correction = correction + 1
        End If
        resultat2 = resultat2 & Chr(correction)

    Next

    If resultat2 = Texte Then
        crypting = resultat1
    Else
        crypting = "Cryptage incorrect"
    End If

End Function

Function decrypting(ByVal Texte As String) As String

    resultat1 = ""
    resultat2 = ""

    longueur = Len(Texte)





    For i = 1 To longueur
        correction = Asc(Mid(Texte, i, 1))
        If i Mod 2 = 0 Then
            correction = correction - 1
        Else
            correction = correction + 1
        End If
        'resultat = resultat & Asc(Mid(texte, i, 1))
        resultat1 = resultat1 & Chr(correction)

    Next


    For i = 1 To longueur
        correction = Asc(Mid(resultat1, i, 1))
        If i Mod 2 = 0 Then
            correction = correction + 1
        Else
            correction = correction - 1
        End If
        resultat2 = resultat2 & Chr(correction)
    Next

    If resultat2 = Texte Then
        decrypting = resultat1
    Else
        MsgBox resultat2 & Chr(10) & Texte
        Sheets("Feuil1").Cells(10, 2) = resultat2
        Sheets("Feuil1").Cells(11, 2) = Texte
    
        decrypting = "Cryptage incorrect" & Chr(10) & resultat2
    End If

End Function

Private Sub test()
    Texte = "CAT.1;CAT.2;CAT.Libellé;Détails;Code CG;Libellé;Rubrique du bilan;Préfixe du compte particulier;Référence code économique;Liaison"

    If Texte = decrypting(crypting(Texte)) Then
        MsgBox crypting(Texte)
        MsgBox "Cryptage OK"
    End If

End Sub

