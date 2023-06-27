Attribute VB_Name = "Vues"
 
'NOTE Vues des Listes View

Sub vue(ByVal liste As Integer, ByVal table As String, Optional ByVal condition As String, Optional ByVal colonne As String, Optional valeur As String)
 
    Dim dligne As Long
 
    'table = "ECO"
 
    If table = "ECO" Then
        UserForm1.Frame2.Caption = table
        dligne = Sheets(table).Cells(Rows.Count, 1).End(xlUp).Row
        With UserForm1.ListView1
    
            .FullRowSelect = True
            .ListItems.Clear
            .Gridlines = True
           
            'Définit le nombre de colonnes et Entêtes
            With .ColumnHeaders
                'Supprime les anciens entêtes
                .Clear
                'Ajoute 3 colonnes en spécifiant le nom de l'entête
                'et la largeur des colonnes
                .Add , , "Code ECO", 80
                .Add , , "Libellé", 50
                .Add , , "Type", 50
                .Add , , "Service", 50
                .Add , , "CG", 50
                .Add , , "CT", 50
                .Add , , "Globalisation", 50
            End With
           
            Call filtre(liste, table, 6, 11, , colonne, valeur)
           
            'Remplissage de la 1ere colonne (création de 3 lignes)
          
        End With
        UserForm1.ComboBox15.Clear
        UserForm1.ComboBox15.AddItem ("Code ECO")
        UserForm1.ComboBox15.AddItem ("Libellé")
        UserForm1.ComboBox15.AddItem ("Type")
        UserForm1.ComboBox15.AddItem ("Service")
        UserForm1.ComboBox15.AddItem ("CG")
        UserForm1.ComboBox15.AddItem ("CT")
        UserForm1.ComboBox15.AddItem ("Globalisation")
        '--------------------------------------------------
    End If
    
    If table = "CG" Then
        dligne = Sheets(table).Cells(Rows.Count, 1).End(xlUp).Row
        With UserForm1.ListView1
            .FullRowSelect = True
            .ListItems.Clear
            'Définit le nombre de colonnes et Entêtes
            With .ColumnHeaders
                'Supprime les anciens entêtes
                .Clear
                'Ajoute 3 colonnes en spécifiant le nom de l'entête
                'et la largeur des colonnes
                .Add , , "Code CG", 80
                .Add , , "Libellé", 80
                .Add , , "Rubrique", 50
                .Add , , "Préfixe du compte particulier", 50
                .Add , , "Ref. Eco?", 50
                .Add , , "Laision", 50
            End With
           

            Call filtre(liste, table, 6, 11, , colonne, valeur)

           
        
            'Remplissage de la 1ere colonne (création de 3 lignes)
          
        End With
        UserForm1.ComboBox15.Clear
        UserForm1.ComboBox15.AddItem ("Code CG")
        UserForm1.ComboBox15.AddItem ("Libellé")
        UserForm1.ComboBox15.AddItem ("Rubrique")
        UserForm1.ComboBox15.AddItem ("Préfixe du compte particulier")
        UserForm1.ComboBox15.AddItem ("Ref. Eco?")
        UserForm1.ComboBox15.AddItem ("Laision")
        '--------------------------------------------------
    End If
    
    If table = "Correspondance" Then
        dligne = Sheets(table).Cells(Rows.Count, 1).End(xlUp).Row
        With UserForm1.ListView1
            .FullRowSelect = True
            .ListItems.Clear
            'Définit le nombre de colonnes et Entêtes
            With .ColumnHeaders
                'Supprime les anciens entêtes
                .Clear
                'Ajoute 3 colonnes en spécifiant le nom de l'entête
                'et la largeur des colonnes
                .Add , , "Compte général", 80
                .Add , , "Libellé", 50
                .Add , , "Préfixe CP1", 50
                .Add , , "Débit/Credit?", 50
                .Add , , "Séquence", 50
                .Add , , "Libellé", 50
                .Add , , "Compte général de contrepartie individualisé", 80
                .Add , , "Compte général de contrepartie globalisé", 80
                .Add , , "Préfixe CP2", 50
                .Add , , "Débit/Credit?", 50
                .Add , , "Laision", 50
                .Add , , "Statut", 50
            End With
               
            Call filtre(liste, table, 3, 13, , colonne, valeur)
          
            'Remplissage de la 1ere colonne (création de 3 lignes)
          
        End With
        UserForm1.ComboBox15.Clear
        UserForm1.ComboBox15.AddItem ("Compte général")
        UserForm1.ComboBox15.AddItem ("Libellé")
        UserForm1.ComboBox15.AddItem ("Préfixe CP1")
        UserForm1.ComboBox15.AddItem ("Débit/Credit?")
        UserForm1.ComboBox15.AddItem ("Séquence")
        UserForm1.ComboBox15.AddItem ("Libellé")
        UserForm1.ComboBox15.AddItem ("Compte général de contrepartie individualisé")
        UserForm1.ComboBox15.AddItem ("Compte général de contrepartie globalisé")
        UserForm1.ComboBox15.AddItem ("Préfixe CP2")
        UserForm1.ComboBox15.AddItem ("Débit/Credit?")
        UserForm1.ComboBox15.AddItem ("Laision")
        UserForm1.ComboBox15.AddItem ("Statut")
        '--------------------------------------------------
    End If
    
    If table = "Tampon" Then
        dligne = Sheets(table).Cells(Rows.Count, 1).End(xlUp).Row
        With UserForm1.ListView2
    
            .FullRowSelect = True
            .ListItems.Clear
            .Gridlines = True
           
            'Définit le nombre de colonnes et Entêtes
            With .ColumnHeaders
                'Supprime les anciens entêtes
                .Clear
                'Ajoute 3 colonnes en spécifiant le nom de l'entête
                'et la largeur des colonnes
                .Add , , "Code ECO", 80
                .Add , , "Libellé", 50
                .Add , , "Type", 50
                .Add , , "Service", 50
                .Add , , "Type", 50
                .Add , , "CT", 50
                .Add , , "Globalisation", 50
                .Add , , "Liaison", 50
            End With
           
            Call filtre(liste, table, 2, 5, condition)
           
            'Remplissage de la 1ere colonne (création de 3 lignes)
          
        End With
        UserForm1.ComboBox15.Clear
        UserForm1.ComboBox15.AddItem ("Code ECO")
        UserForm1.ComboBox15.AddItem ("Libellé")
        UserForm1.ComboBox15.AddItem ("Type")
        UserForm1.ComboBox15.AddItem ("Service")
        UserForm1.ComboBox15.AddItem ("CG")
        UserForm1.ComboBox15.AddItem ("CT")
        UserForm1.ComboBox15.AddItem ("Globalisation")
        '--------------------------------------------------
    End If
    
    If table = "CP" Then
        dligne = Sheets(table).Cells(Rows.Count, 2).End(xlUp).Row
        With UserForm1.ListView3
    
            .FullRowSelect = True
            .ListItems.Clear
            .Gridlines = True
           
            'Définit le nombre de colonnes et Entêtes
            With .ColumnHeaders
                'Supprime les anciens entêtes
                .Clear
                'Ajoute 3 colonnes en spécifiant le nom de l'entête
                'et la largeur des colonnes
                .Add , , "CP", 80
                .Add , , "Details", 50
            End With
           
            Call filtre(liste, table, 2, 3, , colonne, valeur)
           
            'Remplissage de la 1ere colonne (création de 3 lignes)
          
        End With
        '--------------------------------------------------
    End If
    'Spécifie l'affichage en mode "Détails"
    Call logging(Now, Application.UserName, Application.Caption, "Initialisation ListView", "Vues.vue")
    UserForm1.ListView1.View = lvwReport
    UserForm1.ListView2.View = lvwReport
    UserForm1.ListView3.View = lvwReport
End Sub

'NOTE pour alimentation ou filtre des listeView
Sub filtre(ByVal liste As Integer, _
           ByVal table As String, _
           ByVal borneGauche As Integer, _
           ByVal borneDroite As Integer, _
           Optional ByVal condition As String, _
           Optional ByVal colonne As String, _
           Optional ByVal valeur As String)

    dligne = Sheets(table).Cells(Rows.Count, 1).End(xlUp).Row

    If liste = 1 Then
        With UserForm1.ListView1
        
            If colonne <> "" And valeur <> "" Then
                While Sheets(table).Cells(1, dlignec + 1) <> ""
                    dlignec = dlignec + 1
                Wend
                For i = 1 To dlignec
                    If Sheets(table).Cells(1, i) = colonne Then
                        col_cible = i
                        GoTo suite1
                    End If
                Next
suite1:
            End If
                    
            k = 0
                   
            If colonne <> "" And valeur <> "" Then
                'NOTE Chargement des ListeView en fonction des filtres
                For i = 2 To dligne
                    If InStr(1, LCase(Sheets(table).Cells(i, col_cible)), LCase(valeur)) > 0 Then
                        .ListItems.Add , , Sheets(table).Cells(i, borneGauche - 1)
                        k = k + 1
                        For j = borneGauche To borneDroite
                            .ListItems(k).ListSubItems.Add , , Sheets(table).Cells(i, j)
                        Next
                                
                    End If
                Next
            Else
                    
                If condition <> "" Then
                    'NOTE Chargement des ListeView pour les zones Tampon ; dépend de la condition pour filtrer sur la première Colonne
                    For i = 2 To dligne
                        If Sheets(table).Cells(i, 1) = condition Then
                            .ListItems.Add , , Sheets(table).Cells(i, borneGauche - 1)
                            k = k + 1
                            For j = borneGauche To borneDroite
                                .ListItems(k).ListSubItems.Add , , Sheets(table).Cells(i, j)
                            Next
                        End If
                    Next
                Else
                    'NOTE Chargement complète de la table dans la ListeView
                    For i = 2 To dligne
                        .ListItems.Add , , Sheets(table).Cells(i, borneGauche - 1)
                        For j = borneGauche To borneDroite
                            .ListItems(i - 1).ListSubItems.Add , , Sheets(table).Cells(i, j)
                        Next
                    Next
                End If
            End If
        End With
    ElseIf liste = 2 Then
        With UserForm1.ListView2

            If colonne <> "" And valeur <> "" Then
                While Sheets(table).Cells(1, dlignec + 1) <> ""
                    dlignec = dlignec + 1
                Wend
                For i = 1 To dlignec
                    If Sheets(table).Cells(1, i) = colonne Then
                        col_cible = i
                        GoTo suite2
                    End If
                Next
suite2:
            End If
                    
            k = 0
                   
            If colonne <> "" And valeur <> "" Then
                'NOTE Chargement des ListeView en fonction des filtres
                For i = 2 To dligne
                    If InStr(1, LCase(Sheets(table).Cells(i, col_cible)), LCase(valeur)) > 0 Then
                        .ListItems.Add , , Sheets(table).Cells(i, borneGauche - 1)
                        k = k + 1
                        For j = borneGauche To borneDroite
                            .ListItems(k).ListSubItems.Add , , Sheets(table).Cells(i, j)
                        Next
                                
                    End If
                Next
            Else
                    
                If condition <> "" Then
                    'NOTE Chargement des ListeView pour les zones Tampon ; dépend de la condition pour filtrer sur la première Colonne
                    For i = 2 To dligne
                        If Sheets(table).Cells(i, 1) = condition Then
                            .ListItems.Add , , Sheets(table).Cells(i, borneGauche - 1)
                            k = k + 1
                            For j = borneGauche To borneDroite
                                .ListItems(k).ListSubItems.Add , , Sheets(table).Cells(i, j)
                            Next
                        End If
                    Next
                Else
                    'NOTE Chargement complète de la table dans la ListeView
                    For i = 2 To dligne
                        .ListItems.Add , , Sheets(table).Cells(i, borneGauche - 1)
                        For j = borneGauche To borneDroite
                            .ListItems(i - 1).ListSubItems.Add , , Sheets(table).Cells(i, j)
                        Next
                    Next
                End If
            End If
        End With
    ElseIf liste = 3 Then
        With UserForm1.ListView3
            dligne = Sheets(table).Cells(Rows.Count, 2).End(xlUp).Row
            If colonne <> "" And valeur <> "" Then
                While Sheets(table).Cells(1, dlignec + 1) <> ""
                    dlignec = dlignec + 1
                Wend
                For i = 1 To dlignec
                    If Sheets(table).Cells(1, i) = colonne Then
                        col_cible = i
                        GoTo suite3
                    End If
                Next
suite3:
            End If
                    
            k = 0
                   
            If colonne <> "" And valeur <> "" Then
                'NOTE Chargement des ListeView en fonction des filtres
                For i = 2 To dligne
                    If InStr(1, LCase(Sheets(table).Cells(i, col_cible)), LCase(valeur)) > 0 Then
                        .ListItems.Add , , Sheets(table).Cells(i, borneGauche - 1)
                        k = k + 1
                        For j = borneGauche To borneDroite
                            .ListItems(k).ListSubItems.Add , , Sheets(table).Cells(i, j)
                        Next
                                
                    End If
                Next
            Else
                    
                If condition <> "" Then
                    'NOTE Chargement des ListeView pour les zones Tampon ; dépend de la condition pour filtrer sur la première Colonne
                    For i = 2 To dligne
                        If Sheets(table).Cells(i, 1) = condition Then
                            .ListItems.Add , , Sheets(table).Cells(i, borneGauche - 1)
                            k = k + 1
                            For j = borneGauche To borneDroite
                                .ListItems(k).ListSubItems.Add , , Sheets(table).Cells(i, j)
                            Next
                        End If
                    Next
                Else
                    'NOTE Chargement complète de la table dans la ListeView
                    For i = 2 To dligne
                        .ListItems.Add , , Sheets(table).Cells(i, borneGauche - 1)
                        For j = borneGauche To borneDroite
                            .ListItems(i - 1).ListSubItems.Add , , Sheets(table).Cells(i, j)
                        Next
                    Next
                End If
            End If
        End With
    End If
    Call logging(Now, Application.UserName, Application.Caption, "Filtre des ListViews", "Vues.filtre")
End Sub


