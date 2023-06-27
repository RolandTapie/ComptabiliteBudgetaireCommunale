Attribute VB_Name = "Clean_DATA"
Private Sub clean_ECO()
    Dim i, j, k, l, m, dligne, dlign, dligne_eco As Integer
    Dim cat_1, cat_2, cat_lib, code, CG, libelle, types, Service, Regroupement As String
    Dim trouve As Boolean

    Sheets("ECO").Cells.ClearContents

    Sheets("ECO").Cells(1, 1) = "Titre général"
    Sheets("ECO").Cells(1, 2) = "Sous-groupe"
    Sheets("ECO").Cells(1, 3) = "Titre des codes"
    Sheets("ECO").Cells(1, 4) = "Détails"
    Sheets("ECO").Cells(1, 5) = "Code ECO"
    Sheets("ECO").Cells(1, 6) = "Libellé"
    Sheets("ECO").Cells(1, 7) = "Type"
    Sheets("ECO").Cells(1, 8) = "Service"
    Sheets("ECO").Cells(1, 9) = "CG"
    Sheets("ECO").Cells(1, 10) = "CT"

    Sheets("Liste ECO").Select
    dligne_eco = Cells(Rows.Count, 1).End(xlUp).Row
    j = 2
    k = 2

    cat_1 = ""
    cat_2 = ""
    cat_lib = ""
    For i = 1 To dligne_eco
   
        If Sheets("Liste ECO").Cells(i, 4) = "X" And cat_1 = "" Then
            cat_1 = Sheets("Liste ECO").Cells(i, 1)
            cat_2 = Sheets("Liste ECO").Cells(i, 2)
            cat_lib = Sheets("Liste ECO").Cells(i, 3)
        End If
    
        If Sheets("Liste ECO").Cells(i, 1) = "Code" And Sheets("Liste ECO").Cells(i, 3) = "Libellé" Then
        
            m = i + 1
            While Sheets("Liste ECO").Cells(m, 1) <> "" And m <= dligne_eco And InStr(1, Sheets("Liste ECO").Cells(m, 1), "-") > 0
        
                Sheets("ECO").Cells(k, 1) = CStr(Left(Sheets("Liste ECO").Cells(m, 1), 1))
                Sheets("ECO").Cells(k, 2) = CStr(Left(Sheets("Liste ECO").Cells(m, 1), 2))
                Sheets("ECO").Cells(k, 3) = CStr(Left(Sheets("Liste ECO").Cells(m, 1), 3))
                Sheets("ECO").Cells(k, 5) = Sheets("Liste ECO").Cells(m, 1)
            
                If Len(Sheets("ECO").Cells(k, 5)) <> 6 Then
                    Sheets("ECO").Cells(k, 4) = "/!\"
                End If
            
                Sheets("ECO").Cells(k, 6) = Sheets("Liste ECO").Cells(m, 3)
                Sheets("ECO").Cells(k, 7) = Sheets("Liste ECO").Cells(m, 4)
                Sheets("ECO").Cells(k, 8) = Sheets("Liste ECO").Cells(m, 5)
                Sheets("ECO").Cells(k, 9) = CStr(Sheets("Liste ECO").Cells(m, 6))
                Sheets("ECO").Cells(k, 10) = CStr(Sheets("Liste ECO").Cells(m, 7))
                k = k + 1
            
            
                m = m + 1
            
            Wend
        
            cat_1 = ""
            cat_2 = ""

        End If

    Next

    Call logging(Now, Application.UserName, Application.Caption, "Chargement", "Chargement de la  Table ECO")
End Sub

Private Sub clean_CG()
    Dim i, j, k, l, m, dligne, dlign, dligne_eco As Integer
    Dim cat_1, cat_2, cat_lib, code, CG, libelle, types, Service, Regroupement, recherche As String
    Dim trouve As Boolean

    Sheets("CG").Cells.ClearContents

    Sheets("CG").Cells(1, 1) = "CAT.1"
    Sheets("CG").Cells(1, 2) = "Classe"
    Sheets("CG").Cells(1, 3) = "CAT.Libellé"
    Sheets("CG").Cells(1, 4) = "Détails"
    Sheets("CG").Cells(1, 5) = "Code CG"
    Sheets("CG").Cells(1, 6) = "Libellé"
    Sheets("CG").Cells(1, 7) = "Rubrique du bilan"
    Sheets("CG").Cells(1, 8) = "Préfixe du compte particulier"
    Sheets("CG").Cells(1, 9) = "Référence code économique"
    Sheets("CG").Cells(1, 10) = "Cardinalité"
    Sheets("CG").Cells(1, 10) = "Liaison"

    Sheets("Liste CG").Select
    dligne_eco = Cells(Rows.Count, 1).End(xlUp).Row
    j = 2
    k = 2

    cat_1 = ""
    cat_2 = ""
    cat_lib = ""

    For i = 1 To dligne_eco
   
    
        If Sheets("Liste CG").Cells(i, 1) = "Compte général" And Sheets("Liste CG").Cells(i, 3) = "Libellé" Then
        
            m = i + 1
            While Sheets("Liste CG").Cells(m, 1) <> "Compte général" And m <= dligne_eco
                If Sheets("Liste CG").Cells(m, 1) <> "" And Len(Sheets("Liste CG").Cells(m, 1)) >= 5 Then
                    Sheets("CG").Cells(k, 1) = Left(Sheets("Liste CG").Cells(m, 1), 1)
                    Sheets("CG").Cells(k, 2) = Left(Sheets("Liste CG").Cells(m, 1), 2)
                    Sheets("CG").Cells(k, 3) = Left(Sheets("Liste CG").Cells(m, 1), 3)
                    Sheets("CG").Cells(k, 5) = CStr(Sheets("Liste CG").Cells(m, 1))
                    code = Sheets("CG").Cells(k, 5)
                    Sheets("CG").Cells(k, 6) = Sheets("Liste CG").Cells(m, 3)
                    Sheets("CG").Cells(k, 7) = Sheets("Liste CG").Cells(m, 4)
                    Sheets("CG").Cells(k, 8) = Sheets("Liste CG").Cells(m, 5)
                    Sheets("CG").Cells(k, 9) = CStr(Sheets("Liste CG").Cells(m, 6))
                
                    If Sheets("CG").Cells(k, 9) <> "" Then
                        Sheets("CG").Cells(k, 9) = "N"
                        recherche = Search.Search("ECO", "CG", code)
                        If InStr(1, recherche, "|") Then
                            Sheets("CG").Cells(k, 9) = "O"
                            Sheets("CG").Cells(k, 11) = recherche
                        End If
                    End If
                
                    Sheets("CG").Cells(k, 10) = 0
                
                    Debug.Assert k <> 1000
                
                    k = k + 1
                End If
                m = m + 1
            Wend
        

        End If

    Next


    Sheets("CG").Select
    Call logging(Now, Application.UserName, Application.Caption, "Chargement", "Chargement de la  Table CG")
End Sub

Private Sub clean_Correspondance()
    Dim i, j, k, l, m, dligne, dlign, dligne_eco As Integer
    Dim cgs() As String
    Dim cat_1, cat_2, cat_lib, code, CG, libelle, types, Service, Regroupement, recherche As String
    Dim trouve As Boolean

    Sheets("Correspondance").Select
    Cells.ClearContents
    Sheets("Correspondance").Cells(1, 1) = "CG1"
    Sheets("Correspondance").Cells(1, 2) = "CG2"
    Sheets("Correspondance").Cells(1, 3) = "Libellé"
    Sheets("Correspondance").Cells(1, 4) = "Préf.CP1"
    Sheets("Correspondance").Cells(1, 5) = "D/C1"
    Sheets("Correspondance").Cells(1, 6) = "Seq"
    Sheets("Correspondance").Cells(1, 7) = "Mouvement"
    Sheets("Correspondance").Cells(1, 8) = "Compte général de contrepartie"
    Sheets("Correspondance").Cells(1, 9) = "Compte général de contrepartie globalisé"
    Sheets("Correspondance").Cells(1, 10) = "Préf.CP2"
    Sheets("Correspondance").Cells(1, 11) = "D/C2"
    Sheets("Correspondance").Cells(1, 12) = "code économique"
    Sheets("Correspondance").Cells(1, 13) = "statut"



    dligne_eco = Sheets("CG-> CE").Cells(Rows.Count, 1).End(xlUp).Row

    j = 2
    k = 2

    cat_1 = ""
    cat_2 = ""
    cat_lib = ""

    'For i = 2 To dligne_eco
    'If Sheets("CG-> CE").Cells(i, 1) = "" Then
    'Sheets("CG-> CE").Cells(i, 1) = Sheets("CG-> CE").Cells(i - 1, 1)
    'End If
   
    'Sheets("CG-> CE").Cells(i, 2) = Sheets("CG-> CE").Cells(i, 1)
   
    'GoTo suite
    
    'If Sheets("CG-> CE").Cells(i, 1) = "Compte général" And Sheets("CG-> CE").Cells(i, 3) = "Libellé" Then
        
    m = 2
    While m <= dligne_eco                        'And Sheets("CG-> CE").Cells(m, 1) <> "Compte général"
        If Sheets("CG-> CE").Cells(m, 1) <> "" Then
            Sheets("Correspondance").Cells(k, 1) = "1"
            Sheets("Correspondance").Cells(k, 2) = CStr(Sheets("CG-> CE").Cells(m, 2))
            code = Format(Sheets("Correspondance").Cells(k, 2), "00000")
                
            Sheets("Correspondance").Cells(k, 3) = Sheets("CG-> CE").Cells(m, 3)
            Sheets("Correspondance").Cells(k, 4) = Sheets("CG-> CE").Cells(m, 4)
            Sheets("Correspondance").Cells(k, 5) = CStr(Sheets("CG-> CE").Cells(m, 5))
                
            Sheets("Correspondance").Cells(k, 13) = "Défini"
                
            Sheets("Correspondance").Cells(k, 6) = Sheets("CG-> CE").Cells(m, 6)
            Sheets("Correspondance").Cells(k, 7) = Sheets("CG-> CE").Cells(m, 7)
            Sheets("Correspondance").Cells(k, 8) = Sheets("CG-> CE").Cells(m, 8)
                
            If InStr(1, Sheets("Correspondance").Cells(k, 8), "/") > 0 Then
                cgs = Split(Sheets("Correspondance").Cells(k, 8), "/")
                Sheets("Correspondance").Cells(k, 8) = cgs(0)
                Sheets("Correspondance").Cells(k, 9) = cgs(1)
            End If
                
            If InStr(1, Sheets("Correspondance").Cells(k, 8), "X") > 0 Or InStr(1, Sheets("Correspondance").Cells(k, 8), "x") > 0 Then
                Sheets("Correspondance").Cells(k, 13) = "Indéfini"
            End If
                
            Sheets("Correspondance").Cells(k, 10) = CStr(Sheets("CG-> CE").Cells(m, 10))
            Sheets("Correspondance").Cells(k, 11) = CStr(Sheets("CG-> CE").Cells(m, 11))
            Sheets("Correspondance").Cells(k, 12) = CStr(Sheets("CG-> CE").Cells(m, 12))
            'Debug.Assert k <> 1344
            If Sheets("Correspondance").Cells(k, 12) <> "" Then
                Sheets("Correspondance").Cells(k, 12) = "N"
                recherche = Search.Search("ECO", "CG", code)
                If InStr(1, recherche, "|") Then
                    Sheets("Correspondance").Cells(k, 12) = "O"
                    Sheets("Correspondance").Cells(k, 12) = recherche
                End If
            End If
                
            'Sheets("Correspondance").Cells(k, 12) = 0
                
            k = k + 1
        End If
        m = m + 1
    Wend
        

    'End If
    
    
suite:
    'Next

    Call logging(Now, Application.UserName, Application.Caption, "Chargement", "Chargement de la  Table de Correspondance ECO-CG")
End Sub

Private Sub all()
    Call clean_ECO
    Call clean_CG
    Call clean_Correspondance
    MsgBox "clean OK"
End Sub

