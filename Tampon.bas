Attribute VB_Name = "Tampon"
Sub mise_en_tampon(ByVal statut As String, _
                   ByVal data_1 As String, _
                   ByVal data_2 As String, _
                   ByVal data_3 As String, _
                   ByVal data_4 As String, _
                   ByVal data_5 As String, _
                   ByVal data_6 As String, _
                   ByVal data_7 As String)

    dligne = Sheets("Tampon").Cells(Rows.Count, 1).End(xlUp).Row + 1
    Sheets("Tampon").Cells(dligne, 1) = statut
    Sheets("Tampon").Cells(dligne, 2) = data_1
    Sheets("Tampon").Cells(dligne, 3) = data_2
    Sheets("Tampon").Cells(dligne, 4) = data_3
    Sheets("Tampon").Cells(dligne, 5) = data_4
    Sheets("Tampon").Cells(dligne, 6) = data_5
    Sheets("Tampon").Cells(dligne, 7) = data_6
    Sheets("Tampon").Cells(dligne, 8) = data_7
       
    Call logging(Now, Application.UserName, Application.Caption, "Mise en tampon", "Tampon.mise_en_tampon")
    MsgBox "Mise en zone tampon effectué"
End Sub


