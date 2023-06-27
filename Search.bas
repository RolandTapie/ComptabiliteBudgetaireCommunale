Attribute VB_Name = "Search"
Function Search( _
         ByVal table As String, _
         ByVal colonne As String, _
         ByVal valeur As String, _
         Optional ByVal cible As String, _
         Optional ByVal typeContrepartie As String) As String

    'On Error GoTo ErrorHandler
    Dim resultat As String
    Dim cibl As String
    Dim col, col_cible As Integer
    Dim dlignel, dlignec As Double
    Dim table_exist As Boolean


    table_exist = False

    cibl = ""

    If cible <> "" Then
        cibl = cible
    End If

    resultat = ""
    col = 0

    For i = 1 To Worksheets.Count

        If Sheets(i).Name = table Then
            table_exist = True
        End If

    Next

    If table_exist = False Then
        resultat = "table inexistente"
        GoTo fin
    Else
  
        dlignel = Sheets(table).Cells(Rows.Count, 1).End(xlUp).Row
        dlignec = 1
        While Sheets(table).Cells(1, dlignec + 1) <> ""
            dlignec = dlignec + 1
        Wend
    
        For i = 1 To dlignec
            If Sheets(table).Cells(1, i) = colonne Then
                col = i
                GoTo suite:
            End If
        Next
suite:
        If col = 0 Then
            resultat = "colonne inexistante"
            GoTo fin
        End If
    
    
    
        If cibl <> "" Then
            For i = 1 To dlignec
                If Sheets(table).Cells(1, i) = cibl Then
                    col_cible = i
                    GoTo suite1
                End If
            Next
suite1:
            If col = 0 Then
                resultat = "cible inexistente"
                GoTo fin
            End If
        End If
    
    
    
    
    
    
    
        If cibl <> "" Then
            For i = 1 To dlignel
                If valeur = Sheets(table).Cells(i, col) Then
                    
                    
                    
                    If resultat = "" Then
                        resultat = Sheets(table).Cells(i, col_cible)
                        If resultat = "" Then
                            resultat = "cible vide"
                        End If
                    End If
                    
                    
                    
                    
                End If
            Next
            GoTo fin
        End If
    
        'TODO Construire le chainage avec le CG et la contrepartie en tenant compte du tiers individuel/globalisé
    
        If typeContrepartie = "Individuel" Then
            For i = 1 To dlignel
                If CStr(valeur) = CStr(Sheets(table).Cells(i, col)) Then
                    For j = 2 To 7
                        If resultat = "" Then
                            resultat = Sheets(table).Cells(i, j)
                        Else
                            resultat = resultat & "|" & Sheets(table).Cells(i, j)
                        End If
                    Next
                    resultat = resultat & "|" & Sheets(table).Cells(i, 8)
                End If
            Next
        ElseIf typeContrepartie = "Globalisé" Then
            For i = 1 To dlignel
                If CStr(valeur) = CStr(Sheets(table).Cells(i, col)) Then
                    For j = 2 To 7
                        If resultat = "" Then
                            resultat = Sheets(table).Cells(i, j)
                        Else
                            resultat = resultat & "|" & Sheets(table).Cells(i, j)
                        End If
                    Next
                    If Sheets(table).Cells(i, 9) = "" Then
                        resultat = resultat & "|" & Sheets(table).Cells(i, 8)
                    Else
                        resultat = resultat & "|" & Sheets(table).Cells(i, 9)
                    End If
                End If
            Next
        Else
            For i = 1 To dlignel
                If CStr(valeur) = CStr(Sheets(table).Cells(i, col)) Then
                    For j = 2 To dlignec
                        If resultat = "" Then
                            resultat = Sheets(table).Cells(i, j)
                        Else
                            resultat = resultat & "|" & Sheets(table).Cells(i, j)
                        End If
                    Next
                End If
            Next
        End If
    
    
    
        If resultat = "" Then
            resultat = "Valeur indéfinie"
        End If
    
 

    End If

fin:
    Call logging(Now, Application.UserName, Application.Caption, "Recherche", "Search.search")
    Search = resultat

    Exit Function
ErrorHandler:
    Call logging(Now, Application.UserName, Application.Caption, "Recherche : Erreur", "Search.search")
    ecran ("Une erreur s'est produite lors du traitement" & Err.Description)
End Function

Private Sub test()
    Dim CG As String

    CG = Search("ECO", "Code ECO", "11X-01", "CG")



    If CG <> "cible vide" Or CG <> "Valeur indéfinie" Then
        ecran (Search("CG-> CE", "CG1", CG))
    Else
        ecran (CG)
    End If

End Sub

Function find(ByVal table As String, ByVal colonne As String, ByVal valeur As String, Optional ByVal cible As String, Optional ByVal typeContrepartie As String) As String

    Dim CG As String

    If cible <> "" Then
        CG = Search(table, colonne, valeur, cible)
    Else
        CG = Search(table, colonne, valeur)
    End If

    find = ""

    If CG <> "cible vide" And CG <> "Valeur indéfinie" And cible <> "" Then
        If CG = "" Then
            ecran ("Valeur indéfinie")
        Else
            find = Search("Correspondance", "CG2", CG, , typeContrepartie)
            ecran (Search("Correspondance", "CG2", CG, , typeContrepartie))
        End If
    Else
        find = CG
        ecran (CG)
    End If
    Call logging(Now, Application.UserName, Application.Caption, "Recherche", "Search.find")
End Function


