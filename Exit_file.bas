Attribute VB_Name = "Exit_file"
Function exit_f()
    Dim ctl As MSForms.Control
    For Each ctl In UserForm1.controls
        ctl.Enabled = False                      ' toggle Enabled property, use True/False if you don't want to toggle
    Next ctl
    Call logging(Now, Application.UserName, Application.Caption, "exit_fichier", "Exit_file.exit_f")
    ThisWorkbook.Close SaveChanges:=False
End Function

