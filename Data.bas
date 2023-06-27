Attribute VB_Name = "Data"
Function generate_eco_txt()
    'On Error GoTo ErrorHandler

    Dim resultat, cibl, data_file As String
    Dim col, col_cible As Integer
    Dim dlignel, dlignec As Double
    Dim table_exist As Boolean

    data_file = Replace("C:\Users\tallar\Downloads\data.txt", "\", "/")

    Call write_file("CG", data_file)
    Call read_file("CG", data_file)
fin:

    Exit Function

ErrorHandler:
    MsgBox "Une erreur s'est produite lors du traitement"
 
End Function

Sub write_file(ByVal table As String, ByVal path As String)

    Dim f As Integer
    f = FreeFile

    dligne1 = Sheets(table).Cells(Rows.Count, 1).End(xlUp).Row + 1

    'data_file = "C:\Users\tallar\Downloads\data.txt"
    data_file = path
    dlignec = 1
    While Sheets(table).Cells(1, dlignec + 1) <> ""
        dlignec = dlignec + 1
    Wend
       
    Open data_file For Output As #f

    For i = 1 To dligne1
        Texte = ""
        For j = 1 To dlignec
            If Texte = "" Then
                Texte = Sheets(table).Cells(i, j)
            Else
                Texte = Texte & ";" & Sheets(table).Cells(i, j)
            End If
        Next
        'Print #f, (crypting(texte))
        Print #f, ((Texte))
    Next
    Close #f
    Call logging(Now, Application.UserName, Application.Caption, "ecriture_fichier", "Data.write_file")
End Sub

Sub read_file(ByVal table As String, ByVal path As String)

    Dim f As Integer
    f = FreeFile
    Dim ligne As String
    dligne1 = Sheets(table).Cells(Rows.Count, 1).End(xlUp).Row + 1

    'data_file = "C:\Users\tallar\Downloads\data.txt"
    data_file = path
    dlignec = 1
    While Sheets("lecture").Cells(1, dlignec + 1) <> ""
        dlignec = dlignec + 1
    Wend
       
    Open data_file For Input As #f

    While Not EOF(f)
        Line Input #f, ligne
        Sheets("lecture").Cells(dlignec, 1) = decrypting(ligne)
        dlignec = dlignec + 1
    Wend

    Close #f

    Call logging(Now, Application.UserName, Application.Caption, "lecture_fichier", "Data.read_file")
End Sub

Sub append_file(ByVal fichier As String, ByVal Texte As String)

    Dim fso As New FileSystemObject
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set FTW = fso.OpenTextFile(fichier, ForAppending)

    FTW.Write Texte & vbCrLf
    FTW.Close

    Set fso = Nothing
    Call logging(Now, Application.UserName, Application.Caption, "append_fichier", "Data.append_file")
End Sub

Sub test()
    Dim base As clsTextEncode
    Set base = New clsTextEncode

    MsgBox base.DecodeBase64(base.EncodeBase64("test"))
  
    Set base = Nothing
    Call logging(Now, Application.UserName, Application.Caption, "test_EncodeB64", "Data.test")
End Sub

