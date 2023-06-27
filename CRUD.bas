Attribute VB_Name = "CRUD"
Type Data
    data_1 As String
    data_2 As String
End Type

Function CRUD_data(ByVal table As String, ByVal cle As String, ByVal valeur As String, ByVal operation As String, Optional ByVal libelle As String, Optional ByVal data_1 As String, Optional ByVal data_2 As String, Optional ByVal data_3 As String, Optional ByVal data_4 As String, Optional ByVal data_5 As String) As String


    'table correspond à la feuille à manipuler
    'cle correspond au nom de la colonne contenant #valeur
    'valeur correspond à la donnée à retrouver dans #cle
    'operation correspond au type d'opération à effectuer ex: create, retrieve, update delete
    'libelle correspond à la description à reprendre dans l'enregistrement
    'types:data_1 correspond à un spécification à apporter à l'enregistrement ; ex: pour eco =dépense ou recette
    'service:data_2 correspond à un spécification à apporter à l'enregistrement ; ex: pour eco =dépense ou recette
    'CG:data_3 correspond à un spécification à apporter à l'enregistrement ; ex: pour eco =dépense ou recette
    'Regroupement:data_4 correspond à un spécification à apporter à l'enregistrement ; ex: pour eco =dépense ou recette

    On Error GoTo ErrorHandler

    Dim resultat, cibl As String
    Dim col, col_cible As Integer
    Dim dlignel, dlignec As Double
    Dim table_exist, trouve As Boolean

    dligne1 = Sheets(table).Cells(Rows.Count, 1).End(xlUp).Row + 1


    If operation = "create" Then
        If Search.Search(table, cle, valeur, cle) = valeur Then
            CRUD_data = valeur & " existe déjà dans la table"
            GoTo fin
        End If
        
        Sheets(table).Cells(dligne1, 1) = valeur
        Sheets(table).Cells(dligne1, 2) = valeur
        Sheets(table).Cells(dligne1, 3) = libelle
        Sheets(table).Cells(dligne1, 4) = ""
        Sheets(table).Cells(dligne1, 5) = valeur
        Sheets(table).Cells(dligne1, 6) = libelle
        Sheets(table).Cells(dligne1, 7) = data_1
        Sheets(table).Cells(dligne1, 8) = data_2
        Sheets(table).Cells(dligne1, 9) = data_3
        Sheets(table).Cells(dligne1, 10) = data_4
        Sheets(table).Cells(dligne1, 11) = data_5
        
        Call logging(Now, Application.UserName, Application.Caption, operation, table, valeur)
        CRUD_data = "create OK"

    ElseIf operation = "delete" Then

        dlignec = 1
        While Sheets(table).Cells(1, dlignec + 1) <> ""
            dlignec = dlignec + 1
        Wend
        For i = 1 To dlignec
            If Sheets(table).Cells(1, i) = cle Then
                col = i
                GoTo suite
            End If
        Next
suite:
        trouve = False
        For i = 2 To dligne1
            'Debug.Assert i <> 403
            'chercher la colonne
            'NOTE si valeur est = left(Sheets(table).Cells(i, col),3)&"-"&right(Sheets(table).Cells(i, col),2) permet supprimer le parent et tous les codes fils
        
            If Left(Sheets(table).Cells(i, col), 3) & "-" & Right(Sheets(table).Cells(i, col), 2) = valeur Or Sheets(table).Cells(i, col) = valeur Then
                Sheets(table).Rows(i).EntireRow.Delete
                i = i - 1
                Call logging(Now, Application.UserName, Application.Caption, operation, table, valeur)
                trouve = True
                CRUD_data = "delete OK"
                'GoTo go
            End If
        Next
        
        If trouve = False Then
            Call logging(Now, Application.UserName, Application.Caption, operation, table, valeur & "inexistante")
            CRUD_data = "Cet enregistrement n'existe pas/plus dans la table : delete OK"
        End If
        
        'go:
        
    Else
        Call logging(Now, Application.UserName, Application.Caption, operation & " : Operation Incorrecte", table, valeur)
        CRUD_data = "Opération incorrecte : 'create' ; 'retrive' ; 'delete' ; 'update'"
    End If

fin:

    Exit Function

ErrorHandler:
    Call logging(Now, Application.UserName, Application.Caption, operation & " : Erreur", table, valeur & "inexistante")
    CRUD_data = "Une erreur s'est produite lors du traitement"
 
End Function

Function add_delete_CG() As String

End Function

Function link_ECO_CG() As String

End Function

Function test_eco(ByVal code_eco As String) As Boolean
    test_eco = False

    If (Len(code_eco) = 6 Or Len(code_eco) = 8) And InStr(1, code_eco, "-") > 0 Then
        test_eco = True
    End If

End Function

Private Sub test()

    eco = "XXX-XX"
    CG = "2205xxx"

    If test_eco(eco) Then

        'tester si CG existe et recupérer compte general de liaison
        cgs = Search.Search("CG", "Code CG", CG, "Code CG")
    
        If cgs = "" Then
            ecran ("Compte général " & CG & " inexistant")
            Exit Sub
        End If
        Call logging(Now, Application.UserName, Application.Caption, "test", "CRUD.test")
        ecran (CRUD_data(eco, "create", "eco test", "D", "O", cgs, "80"))
        'ecran (CRUD_data(eco, "delete"))
    Else
        Call logging(Now, Application.UserName, Application.Caption, "test", "CRUD.test")
        ecran ("Structure de code " & eco & " incorrect")
    End If


End Sub

