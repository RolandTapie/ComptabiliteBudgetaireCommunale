Attribute VB_Name = "log"
'NOTE implémentation de la procédure de logging dans l'application
Sub logging(ByVal TimeLog As Date, ByVal statut As String, _
            Optional ByVal data_1 As String, _
            Optional ByVal data_2 As String, _
            Optional ByVal data_3 As String, _
            Optional ByVal data_4 As String, _
            Optional ByVal data_5 As String, _
            Optional ByVal data_6 As String, _
            Optional ByVal data_7 As String)
 
    dligne = Sheets("Log").Cells(Rows.Count, 1).End(xlUp).Row + 1
    Sheets("Log").Cells(dligne, 1) = TimeLog
    Sheets("Log").Cells(dligne, 2) = statut
    Sheets("Log").Cells(dligne, 3) = data_1
    Sheets("Log").Cells(dligne, 4) = data_2
    Sheets("Log").Cells(dligne, 5) = data_3
    Sheets("Log").Cells(dligne, 6) = data_4
    Sheets("Log").Cells(dligne, 7) = data_5
    Sheets("Log").Cells(dligne, 8) = data_6
    Sheets("Log").Cells(dligne, 9) = data_7
        
End Sub

Sub test()
    Call logging(Now, Application.UserName, Application.Caption, "creation", "eco", "cg", "global", "data_6", "data_7")
End Sub

'NOTE implémentation de la procédure de sauvegarde des logging via log.txt
Sub logFile(ByVal path As String, ByVal table As String)
    dligne = Sheets(table).Cells(Rows.Count, 1).End(xlUp).Row
    lo = Sheets(table).Cells(Rows.Count, 10).End(xlUp).Row
    logs = ""
    
    For i = lo To dligne
        If Sheets(table).Cells(i, 10) <> "x" Then
            logi = ""
            For j = 1 To 9
                If logi = "" Then
                    logi = Sheets(table).Cells(i, j)
                Else
                    logi = logi & ";" & Replace(Sheets(table).Cells(i, j), ";", ".")
                End If
            Next
            Sheets(table).Cells(i, 10) = "x"
            
            
            If logs = "" Then
                logs = logi
                'MsgBox logs
            Else
                logs = logs & vbCrLf & logi
                'MsgBox logs
            End If
            
        End If
        
    Next
    If logs <> "" Then
        Call append_file(path, logs)
    End If
End Sub

Sub controls()

    On Error Resume Next

    i = 2
    For Each con In UserForm1.controls
        Sheets("Controls").Cells(i, 1) = con.Name
        Sheets("Controls").Cells(i, 2) = con.Caption
        Sheets("Controls").Cells(i, 3) = con.Value
        i = i + 1
        Application.Visible = True
    Next
End Sub


