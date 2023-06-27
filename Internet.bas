Attribute VB_Name = "Internet"
Function Internet_check() As Boolean
    'PURPOSE: Output TRUE if computer has an internet connection
    'SOURCE: www.thespreadsheetguru.com
    On Error GoTo ErrorHandling

    'Application.DisplayAlerts = False

    Dim config_source As String
    Dim root As String
    root = "https://raw.githubusercontent.com/RolandTapie/ComptabiliteBudgetaireCommunale/main/"
    Dim entite() As String
    Dim user() As String
    Dim details() As String


    config_source = root & "config.txt"
    download = root & "Entites.txt"
    datas = root & "data.txt"
    Users = root & "users.txt"
    'config_source = "https://www.yahoo.com"
    Dim objHTTP As Object
    'Test for Internet Connection
    On Error Resume Next
    Set objHTTP = CreateObject("MSXML2.XMLHTTP")
    objHTTP.Open "GET", config_source, False
    
    
    'objHTTP.Send
    
    
    Internet_check = (objHTTP.status = 200)
    Debug.Print objHTTP.status
    On Error GoTo 0
    'Report to User if Internet Connection not detected
    If Internet_check = False Then
        'MsgBox "pas d'internet"
    Else
        objHTTP.Open "GET", download, False
        
        'objHTTP.Send
        
        
        Internet_check = (objHTTP.status = 200)
        entite = Split(objHTTP.responseText, ":")
        
        For i = 0 To UBound(entite)
            'UserForm1.ComboBox6.AddItem entite(i)
        Next
        
        objHTTP.Open "GET", datas, False
        
        
        'objHTTP.Send
        
        
        Internet_check = (objHTTP.status = 200)
        datas = objHTTP.responseText
        'Call SaveToFile("C:\Users\tallar\Downloads\data2.txt", datas)
        
        'Call DeleteUrlCacheEntry(Users)
        
        objHTTP.Open "GET", Users, False
        
        
        'objHTTP.Send
        
        
        Internet_check = (objHTTP.status = 200)
        user = Split(objHTTP.responseText, "|")
        Call logging(Now, Application.UserName, Application.Caption, "Recuperation des data users", "Internet.Internet_Check")
        
        Count = 1
        For i = 0 To UBound(user)
            If InStr(user(i), UserForm1.Label4.Caption) > 0 Then
                details = Split(user(i), ":")
                login = details(0)
                pwd = details(1)
                
verification:
                pwd_user = Application.InputBox("Votre mot de passe :")
                Call logging(Now, Application.UserName, Application.Caption, "Verification des data users", "Internet.Internet_Check")
                
                If pwd = pwd_user Then
                    GoTo fin:
                ElseIf Count <= 3 Then
                    MsgBox "Données d'identification incorrectes : il vous reste " & 3 - Count & " tentatives"
                    Count = Count + 1
                    GoTo verification
                ElseIf Count > 3 Then
                    GoTo expulsion
                End If
                Call logging(Now, Application.UserName, Application.Caption, "Verification des data users", "Internet.Internet_Check")
            End If
        Next
        Debug.Assert UserForm1.Label4.Caption = ""
        MsgBox "Utilisateur : " & UserForm1.Label4.Caption & " inexistant"
expulsion:
        Call exit_f
        
        
    End If
    
    GoTo fin
    
ErrorHandling:
    Application.DisplayAlerts = True
    MsgBox ("Un erreur lors de la configuration")
    Call exit_f

fin:
    Application.DisplayAlerts = True
End Function

Private Sub test()
    Call logging(Now, Application.UserName, Application.Caption, "Test Connection", "Internet.test")
    MsgBox Internet_check()
End Sub

Function IsInternetConnected(Optional SupressMessage As Boolean) As Boolean
    'PURPOSE: Output TRUE if computer has an internet connection
    'SOURCE: www.thespreadsheetguru.com
    Dim objHTTP As Object
    'Test for Internet Connection
    On Error Resume Next
    Set objHTTP = CreateObject("MSXML2.XMLHTTP")
    objHTTP.Open "GET", "https://www.yahoo.com", False
    objHTTP.Send
    IsInternetConnected = (objHTTP.status = 200)
    On Error GoTo 0
    'Report to User if Internet Connection not detected
    If IsInternetConnected = False And SupressMessage = False Then
        MsgBox "No internet connection detected! Connection to the internet is required before proceeding", _
               vbCritical, "No Internet Detected"
    End If
End Function

Sub SaveToFile(sFileName, sContent)

    Dim lignes() As String
    Dim ligne() As String
    
    'Saves a string to a file and closes the file
    ' sample usage: SaveToFile "d:\test.txt", "test"
    
    Dim f As Integer
    Sheets("lecture").Cells.ClearContents
    
    f = FreeFile
    k = 1
    Open sFileName For Output As #f
    lignes = Split(sContent, Chr(10))
    For i = 0 To UBound(lignes)
    
        ligne = Split(lignes(i), ";")
        
        For j = 0 To UBound(ligne)
            Sheets("lecture").Cells(k, j + 1) = (ligne(j))
        Next
        
        k = k + 1
        Print #f, lignes(i)
    Next
    Close #f
    
    Call logging(Now, Application.UserName, Application.Caption, "Sauvegarde dans fichier txt", "Internet.SaveToFile")
End Sub


