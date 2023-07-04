VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1 
   Caption         =   "Contrôle Code Economique et Comptabilité générale"
   ClientHeight    =   9300
   ClientLeft      =   -8030
   ClientTop       =   -34500
   ClientWidth     =   18350
   OleObjectBlob   =   "UserForm1.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim eco As String

Private Sub ComboBox1_Enter()

    If Search.Search("Regroupement", "CT", UserForm1.ComboBox1.Value, "CT") = "" Then
        MsgBox "Selectionner un code de regoupement valide"
        UserForm1.ComboBox1.Value = ""
    End If

End Sub

Private Sub ComboBox10_Enter()
    UserForm1.Label18.Caption = Search.Search("CG", "Code CG", UserForm1.ComboBox10.Value, "Préfixe du compte particulier")
End Sub

Private Sub ComboBox14_Enter()
    UserForm1.Label19.Caption = Search.Search("CG", "Code CG", UserForm1.ComboBox14.Value, "Préfixe du compte particulier")
End Sub

Private Sub CommandButton1_Click()

    On Error GoTo ErrorHandler

    Call affiche

    If UserForm1.TextBox4.Value <> "" Then

        Chainage = ""
        If UserForm1.OptionButton1.Value = True Then
            Chainage = "CG"
        End If
    
    
        eco = Split(UserForm1.ComboBox5.Value, ":")(0)
    
        'TODOOK tester le code éonomique ;si enfant, vérifier si parent existe , reprendre les caratéristiques du père
    
        'Récupération du Type de globalisation
        type_contrepartie = Search.Search("ECO", "Code ECO", eco, "Globalisation")
     
    
        detail (find(UserForm1.TextBox1.Value, UserForm1.TextBox2.Value, eco, Chainage, type_contrepartie))
    Else
        detail (find(UserForm1.TextBox1.Value, UserForm1.TextBox2.Value, eco))
    End If

    Call masque
    Exit Sub

ErrorHandler:
    Call masque
    ecran ("Une erreur s'est produite lors du traitement" & Err.Description)

End Sub

Private Sub CommandButton10_Click()
    Dim vues As clsVues
    Set vues = New clsVues
    Call vues.filtre(UserForm1.ListView1, UserForm1.Frame2.Caption, UserForm1.ComboBox15.Value, UserForm1.TextBox16.Value)
    Set vues = Nothing
End Sub

Private Sub CommandButton11_Click()
    'TODO implémenter l'association entre ECO et CG


End Sub

Private Sub CommandButton13_Click()
    Call coherence_eco
End Sub

Private Sub CommandButton2_Click()
    'On Error GoTo ErrorHandler


    type_contrepartie = "Individuel"

    Call affiche

    eco = UserForm1.TextBox5.Value
    
    If UserForm1.ComboBox4.Value = "" Then
        UserForm1.ComboBox4.Value = "CG: Non selectionné"
    End If

    If UserForm1.ToggleButton1.Caption = "CREATION" Or UserForm1.ToggleButton1.Caption = "LIAISON" Then
        operation = "create"
        libelle = UserForm1.TextBox7.Value
        types = UserForm1.ComboBox2.Value
        Service = UserForm1.ComboBox3.Value
        CG = Split(UserForm1.ComboBox4.Value, ":")(0)
        Regroupement = Split(UserForm1.ComboBox1.Value, ":")(0)
      
    Else
        operation = "delete"
    End If



    If UserForm1.CommandButton2.Caption = "Créer code economique" Or UserForm1.CommandButton2.Caption = "Réaliser la liaion" Then
        
        'TODOOK tester le code éonomique ;si enfant, vérifier si parent existe , reprendre les caratéristiques du père
        If Len(eco) = 8 And test_eco(eco) Then
            eco = Left(eco, 3) & "-" & Right(eco, 2)
            dligne = Sheets("ECO").Cells(Rows.Count, 1).End(xlUp).Row + 1
            For i = 2 To dligne
                If eco = Sheets("ECO").Cells(i, 5) Then
                    types = Sheets("ECO").Cells(i, 7)
                    Service = Sheets("ECO").Cells(i, 8)
                    CG = Sheets("ECO").Cells(i, 9)
                    Regroupement = Sheets("ECO").Cells(i, 10)
                    GoTo ecoOK
                End If
            Next
                    
            ecran ("Code Economique parent inexisant pour :" & UserForm1.TextBox5.Value)
            GoTo fin
                    
        End If
    
        'Récupération du Type de globalisation
        type_contrepartie = Search.Search("ECO", "Code ECO", eco, "Globalisation")
        If type_contrepartie = "" Then
            type_contrepartie = "Individuel"
        End If
                
ecoOK:
        'TODOOK implémenter la fonction de test
        If test_eco(eco) Then
                    
            cgs = Search.Search("CG", "Code CG", CG, "Code CG")
                    
            If cgs = "" Then
                reponse = MsgBox("Compte général " & CG & " inexistant" & vbCrLf & "Voulez-vous stocker l'opération en zone tampon?", vbYesNo)
                'TODOOK proposer la mise en tampon le temps de créer le compte général
                If reponse = vbYes Then
                         
                    Call mise_en_tampon("ECO", eco, libelle, types, Service, Regroupement, type_contrepartie)
                    Call vue(2, "Tampon")
                End If
                GoTo fin
            End If
                    
                    
            eco = UserForm1.TextBox5.Value
            ecran (CRUD_data("ECO", "Code ECO", eco, "create", libelle, types, Service, cgs, Regroupement, type_contrepartie))
            
            If UserForm1.CommandButton2.Caption = "Réaliser la liaion" Then
                'NOTE supprimer eco de la zone tampon
                dligne = Sheets("Tampon").Cells(Rows.Count, 1).End(xlUp).Row
                For i = 2 To dligne
                    If Sheets("Tampon").Cells(i, 2) = eco Then
                        Sheets("Tampon").Rows(i).EntireRow.Delete
                        i = i - 1
                    End If
                Next
                
                'NOTE Formulaire Initialisation Tampon ECo
                Call vue(2, "Tampon")
                
                
            End If
                   
        Else
            eco = UserForm1.TextBox5.Value
            ecran ("Structure de code " & eco & " incorrect")
        End If
        
        'ElseIf UserForm1.CommandButton2.Caption = "Réaliser la liaion" Then
    
    
    Else
        eco = UserForm1.TextBox5.Value
        
        'TODOOK si eco parent, vérifier si fils existent et lui demander s'il veut aussi supprimer les enfants
        dligne = Sheets("ECO").Cells(Rows.Count, 1).End(xlUp).Row
        
        fils = ""
        
        For i = 2 To dligne
            If eco = Left(Sheets("ECO").Cells(i, 5), 3) & "-" & Right(Sheets("ECO").Cells(i, 5), 2) And Len(Sheets("ECO").Cells(i, 5)) = 8 Then
                fils = fils & vbCrLf & Sheets("ECO").Cells(i, 5)
            End If
        Next
        
        If fils <> "" Then
            reponse = MsgBox("Ce code Economique a des codes fils" & vbCrLf & fils & vbCrLf & "Voulez-vous effectuer la suppression ?", vbYesNo)
            If reponse = vbYes Then
                ecran (CRUD_data("ECO", "Code ECO", eco, "delete"))
                GoTo fin
            Else
                ecran ("Suppression annulée")
                GoTo fin
            End If
        End If
        
        ecran (CRUD_data("ECO", "Code ECO", eco, "delete"))
        'BUG Tester les listes filtres (ListView1) après suppression d'un code ECO
        
    End If


fin:
    Call masque
    'Call UserForm_Initialize
    Exit Sub

ErrorHandler:
    Call masque
    Call UserForm_Initialize
    ecran ("Une erreur s'est produite lors du traitement" & Err.Description)
End Sub

Private Sub CommandButton3_Click()
    On Error GoTo ErrorHandler
    
    If UserForm1.OptionButton2 = True Then
        Call reporting.reporting("ECO")
    ElseIf UserForm1.OptionButton4 = True Then
        Call reporting.reporting("CG")
    ElseIf UserForm1.OptionButton3 = True Then
        Call reporting.reporting("Correspondance")
    End If
    
    Exit Sub

ErrorHandler:
    Call masque
    ecran ("Une erreur s'est produite lors du traitement")
End Sub

Private Sub CommandButton4_Click()
    details.Height = 306
    details.Width = 617
    details.Show
End Sub

Private Sub CommandButton5_Click()
    codeCG = UserForm1.TextBox12.Value
    libelle = UserForm1.TextBox13.Value
    rubrique = UserForm1.ComboBox7.Value
    debit = UserForm1.ComboBox16.Value
    Sequence = UserForm1.ComboBox17.Value
    prefixe = UserForm1.ComboBox8.Value
    contrepartieIndividuel = UserForm1.ComboBox10.Value
    contrepartieGlobalise = UserForm1.ComboBox14.Value
    prefixeContrepartie = UserForm1.Label18.Caption
    If debit = "Débit" Then
        debitContrepartie = "Crédit"
    Else
        debitContrepartie = "Débit"
    End If
    reference = "N"
    liaison = "0"
    
    If UserForm1.CommandButton5.Caption = "Créer CG" Then
        If Search.Search("CG", "Code CG", codeCG) <> "Valeur indéfinie" Then
            MsgBox "le CG exite déjà dans la table"
        Else
            Call CRUD_data("CG", "Code CG", codeCG, "create", libelle, rubrique, prefixe, reference, liaison)
            Call CRUD_data("Correspondance", "Compte général", codeCG, "create", libelle, prefixe, debit, Sequence, , contrepartieIndividuel, contrepartieGlobalise, prefixeContrepartie, debitContrepartie)
            MsgBox "le CG : " & codeCG & " créé" & vbCrLf & " et mise en correspondance effecttuée"
        End If
    ElseIf UserForm1.CommandButton5.Caption = "Supprimer CG" Then
        'TODO vérifier si c'est n'est pas lié à Eco ou en Contrepartie : si liée à CG, suppression impossible tant que eco existe ; si contrepartie, effacer la contrepartie avant suppression du CG
        Call CRUD_data("CG", "Code CG", codeCG, "delete")
    End If
    
    
End Sub

Private Sub CommandButton8_Click()
    Application.Visible = True
End Sub

Private Sub CommandButton9_Click()
    eco = Label8.Caption
    type_contrepartie = Search.Search("ECO", "Code ECO", eco, "Globalisation")
    detail (find(UserForm1.TextBox1.Value, UserForm1.TextBox2.Value, eco, "CG", type_contrepartie))
End Sub

Private Sub ListView1_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    ListView1.Sorted = False
    ListView1.SortKey = ColumnHeader.Index - 1
    
    If ListView1.SortOrder = lvwAscending Then
        ListView1.SortOrder = lvwDescending
    Else
        ListView1.SortOrder = lvwAscending
    End If
    
    ListView1.Sorted = True
End Sub

Private Sub ListView2_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    ListView2.Sorted = False
    ListView2.SortKey = ColumnHeader.Index - 1
    
    If ListView2.SortOrder = lvwAscending Then
        ListView2.SortOrder = lvwDescending
    Else
        ListView2.SortOrder = lvwAscending
    End If
    
    ListView2.Sorted = True
End Sub

Private Sub ListView3_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    ListView3.Sorted = False
    ListView3.SortKey = ColumnHeader.Index - 1
    
    If ListView3.SortOrder = lvwAscending Then
        ListView3.SortOrder = lvwDescending
    Else
        ListView3.SortOrder = lvwAscending
    End If
    
    ListView3.Sorted = True
End Sub

Private Sub ListView1_ItemClick(ByVal Item As MSComctlLib.ListItem)
  
    Label8.Caption = ListView1.SelectedItem.SubItems(4)
    
    If UserForm1.Frame2.Caption = "CG" Then
        
        UserForm1.Label10 = ListView1.SelectedItem.SubItems(4)
        UserForm1.ComboBox4.Value = ListView1.SelectedItem.text & ": Liason via zone tampon"
        UserForm1.TextBox12.Value = ListView1.SelectedItem.SubItems(4)
        UserForm1.TextBox13.Value = ListView1.SelectedItem.SubItems(5)
        UserForm1.ComboBox7.Value = ListView1.SelectedItem.SubItems(6)
        UserForm1.ComboBox8.Value = ListView1.SelectedItem.SubItems(7)
        UserForm1.ComboBox9.Value = ListView1.SelectedItem.SubItems(8) & ": > :" & ListView1.SelectedItem.SubItems(9)
    End If
    'MsgBox eco
End Sub

Private Sub ListView2_ItemClick(ByVal Item As MSComctlLib.ListItem)
    UserForm1.Label9.Caption = ListView2.SelectedItem.text
    UserForm1.TextBox5.Value = ListView2.SelectedItem.text
    UserForm1.TextBox7.Value = ListView2.SelectedItem.SubItems(1)
    UserForm1.ComboBox2.Value = ListView2.SelectedItem.SubItems(2)
    UserForm1.ComboBox3.Value = ListView2.SelectedItem.SubItems(3)
    UserForm1.ComboBox1.Value = ListView2.SelectedItem.SubItems(4)
    UserForm1.ComboBox12.Value = ListView2.SelectedItem.SubItems(5)

    UserForm1.ToggleButton1.Caption = "LIAISON"
    UserForm1.ToggleButton1.BackColor = RGB(0, 100, 0)
    UserForm1.TextBox7.Enabled = True
    UserForm1.TextBox8.Enabled = True
    UserForm1.TextBox9.Enabled = True
    UserForm1.TextBox10.Enabled = True
    UserForm1.TextBox11.Enabled = True
    UserForm1.ComboBox1.Enabled = True
    UserForm1.ComboBox2.Enabled = True
    UserForm1.ComboBox3.Enabled = True
    UserForm1.ComboBox4.Enabled = True
        
    UserForm1.CommandButton2.Caption = "Réaliser la liaion"
End Sub

Private Sub ListView3_BeforeLabelEdit(Cancel As Integer)
    'TODO alimenter listView3 avec les préfixes de CP

End Sub

Private Sub MultiPage1_Change()
    ' Call UserForm_Initialize
End Sub

Private Sub MultiPage2_Change()
    'Call UserForm_Initialize
End Sub

Private Sub OptionButton6_Click()
    Dim vues As clsVues
    Set vues = New clsVues
    
    Call vues.Init_vue(UserForm1.ListView1, "ECO", UserForm1.ComboBox15)
    UserForm1.Frame2.Caption = "ECO"
    Set vues = Nothing
    
End Sub

Private Sub OptionButton7_Click()
    Dim vues As clsVues
    Set vues = New clsVues
    
    Call vues.Init_vue(UserForm1.ListView1, "CG", UserForm1.ComboBox15)
    UserForm1.Frame2.Caption = "CG"
    Set vues = Nothing
End Sub

Private Sub OptionButton8_Click()
    Dim vues As clsVues
    Set vues = New clsVues
    
    Call vues.Init_vue(UserForm1.ListView1, "Correspondance", UserForm1.ComboBox15)
    UserForm1.Frame2.Caption = "Correspondance"
    
    Set vues = Nothing
End Sub

Private Sub ToggleButton1_Change()
    If UserForm1.ToggleButton1.Value = False Then
        UserForm1.ToggleButton1.Caption = "CREATION"
        UserForm1.ToggleButton1.BackColor = RGB(0, 255, 0)
        UserForm1.TextBox7.Enabled = True
        UserForm1.TextBox8.Enabled = True
        UserForm1.TextBox9.Enabled = True
        UserForm1.TextBox10.Enabled = True
        UserForm1.TextBox11.Enabled = True
        UserForm1.ComboBox1.Enabled = True
        UserForm1.ComboBox2.Enabled = True
        UserForm1.ComboBox3.Enabled = True
        UserForm1.ComboBox4.Enabled = True
        
        UserForm1.CommandButton2.BackColor = RGB(0, 255, 0)
        UserForm1.CommandButton2.Caption = "Créer code economique"
    Else
        UserForm1.ToggleButton1.Caption = "SUPPRESSION (à verrouiller???)"
        UserForm1.ToggleButton1.BackColor = RGB(255, 0, 0)
        UserForm1.TextBox7.Enabled = False
        UserForm1.TextBox8.Enabled = False
        UserForm1.TextBox9.Enabled = False
        UserForm1.TextBox10.Enabled = False
        UserForm1.TextBox11.Enabled = False
        UserForm1.ComboBox1.Enabled = False
        UserForm1.ComboBox2.Enabled = False
        UserForm1.ComboBox3.Enabled = False
        UserForm1.ComboBox4.Enabled = False
        
        UserForm1.CommandButton2.Caption = "Supprimer code economique"
        UserForm1.CommandButton2.BackColor = RGB(255, 0, 0)
        
    End If
End Sub

Private Sub ToggleButton2_Change()
    If UserForm1.ToggleButton2.Value = False Then
        UserForm1.ToggleButton2.Caption = "CREATION"
        UserForm1.ToggleButton2.BackColor = RGB(0, 255, 0)
     
        
        UserForm1.CommandButton5.BackColor = RGB(0, 255, 0)
        UserForm1.CommandButton5.Caption = "Créer CG"
    Else
        UserForm1.ToggleButton2.Caption = "SUPPRESSION (à verrouiller???)"
        UserForm1.ToggleButton2.BackColor = RGB(255, 0, 0)
       
        
        UserForm1.CommandButton5.Caption = "Supprimer CG"
        UserForm1.CommandButton5.BackColor = RGB(255, 0, 0)
        
    End If
End Sub

'ACTION Initialisation du Formulaire
Private Sub UserForm_Initialize()

    'Application.Visible = False

    Static etat As String


    UserForm1.Height = Application.Height
    UserForm1.Width = Application.Width
    UserForm1.Font.Name = "Courier new"
    

    Dim log As clsLogging
    Set log = New clsLogging

    If Sheets("Parametres").Cells(1, 1) <> "X" Then
        user = Application.InputBox("Votre identifiant ?", Identifiant)

        If Main.Main_identification(user) = False Then
            For Each ctrl In UserForm1.controls
                ctrl.Enabled = False
            Next
        Else
            'Connexion réussie
            etat = "OK"
            Sheets("Parametres").Cells(1, 1) = "X"
            Sheets("Parametres").Cells(2, 1) = user
        
        End If
    Else

    End If


    



    GoTo sans_controle

    '    UserForm1.Label5.Caption = Sheets("Feuil1").Cells(1, 1)
    '
    '    If UserForm1.Label5.Caption = "LOCAL" Then
    '        'tester si date de validité
    '        MsgBox "Control à implémenter"
    '
    '    Else
    '
    '        If UserForm1.Label3.Caption = "Not Connected" Then
    '            UserForm1.Label4.Caption = test_ID
    '
    '            If Internet_Check() Then
    '                UserForm1.Label3.BackColor = RGB(0, 255, 0)
    '                UserForm1.Label3.Caption = "Connected"
    '            Else
    '                UserForm1.Label3.BackColor = RGB(255, 0, 0)
    '                UserForm1.Label3.Caption = " Not Connected"
    '            End If
    '        End If
    '
    '    End If

sans_controle:


    Count = 0

    Call masque

    'UserForm1.Width = 950
    'UserForm1.Height = 419

    UserForm1.ToggleButton1.BackColor = RGB(0, 255, 0)
    UserForm1.CommandButton2.BackColor = RGB(0, 255, 0)
    
    UserForm1.Label5.Caption = user

    'NOTE Formulaire Initilisation code de regroupement
    dcombo = Sheets("Regroupement").Cells(Rows.Count, 1).End(xlUp).Row
    ComboBox1.Clear
    For i = 2 To dcombo
        ComboBox1.AddItem Sheets("Regroupement").Cells(i, 1) & ": " & Sheets("Regroupement").Cells(i, 2)
    Next


    'NOTE Formulaire Initilisation des Codes Eco
    dcombo = Sheets("ECO").Cells(Rows.Count, 5).End(xlUp).Row
    ComboBox5.Clear
    ComboBox5.Value = "Sélectionner un code économique"

    For i = 2 To dcombo
        ComboBox5.AddItem Sheets("ECO").Cells(i, 5) & ":  " & Sheets("ECO").Cells(i, 6) & ":  " & Sheets("ECO").Cells(i, 7) & ":  " & Sheets("ECO").Cells(i, 8) & ":  " & Sheets("ECO").Cells(i, 9) & ":  " & Sheets("ECO").Cells(i, 10)
    Next

    'NOTE Formulaire Initilisation des CG
    dcombo = Sheets("CG").Cells(Rows.Count, 5).End(xlUp).Row
    ComboBox4.Clear
    ComboBox10.Clear
    ComboBox14.Clear
    For i = 2 To dcombo
        If Len(Sheets("CG").Cells(i, 5)) >= 5 Then
            ComboBox4.AddItem Sheets("CG").Cells(i, 5) & ": " & Sheets("CG").Cells(i, 6)
            ComboBox10.AddItem Sheets("CG").Cells(i, 5)
            ComboBox14.AddItem Sheets("CG").Cells(i, 5)
        End If
    Next
    
    'NOTE Formulaire Initilisation Rubrique Compte Généraux
    dcombo = Sheets("Rubrique").Cells(Rows.Count, 1).End(xlUp).Row
    ComboBox7.Clear
    For i = 2 To dcombo

        ComboBox7.AddItem Sheets("Rubrique").Cells(i, 1)

    Next
    
    'NOTE Formulaire Initilisation des CP
    dcombo = Sheets("CP").Cells(Rows.Count, 2).End(xlUp).Row
    ComboBox8.Clear
    For i = 2 To dcombo

        ComboBox8.AddItem Sheets("CP").Cells(i, 2)

    Next

    'NOTE Formulaire Initilisation Rubrique Compte Généraux
    ComboBox12.Clear
    ComboBox12.AddItem "Individuel"
    ComboBox12.AddItem "Globalisé"
    
    'NOTE Formulaire Initilisation DEBIT/CREDIT
    ComboBox16.Clear
    ComboBox16.AddItem "Débit"
    ComboBox16.AddItem "Crédit"
    
    'NOTE Formulaire Initilisation SEQUENCE
    ComboBox17.Clear
    ComboBox17.AddItem "1"
    ComboBox17.AddItem "2"
    ComboBox17.AddItem "3"
    ComboBox17.AddItem "4"
    ComboBox17.AddItem "5"
    ComboBox17.AddItem "6"

    'NOTE Formulaire Initilisation Recette/Dépense
    ComboBox2.Clear
    ComboBox2.AddItem "Recette"
    ComboBox2.AddItem "Dépense"

    'NOTE Formulaire Initilisation Ordinaire/Extraordinaire
    ComboBox3.Clear
    ComboBox3.AddItem "Ordinaire"
    ComboBox3.AddItem "EXtraordinaire"

    

    'NOTE Formulaire Initialisation ListeView
    UserForm1.Frame2.Caption = "ECO"
    'Call vue(1, "ECO")

    'NOTE Formulaire Initialisation Tampon ECo
    'Call vue(2, "Tampon")                        ', "ECO")

    'NOTE Formulaire Initialisation CP
    'Call vue(3, "CP")

    Dim vues As clsVues
    Set vues = New clsVues
    
    Call vues.Init_vue(UserForm1.ListView1, "ECO", UserForm1.ComboBox15)
    Call vues.Init_vue(UserForm1.ListView2, "Tampon")
    Call vues.Init_vue(UserForm1.ListView3, "CP")
    
    
    
    Set vues = Nothing
    
    Call log.logging("Log", Now, Application.UserName, Application.Caption, "Initialisation du formulaire Prinicipale", "Userform1.UserForm_Initialize")


    Set log = Nothing
    
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    MsgBox "Fermeture de l'application"
End Sub

