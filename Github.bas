Attribute VB_Name = "Github"
'https://api.github.com/repos/{username}/{repository_name}/contents/{file_path}
'https://api.github.com/repos/RolandTapie/ComptabiliteBudgetaireCommunale/contents/config.txt
Sub Github_log()
    Dim git As clsGithub
    Set git = New clsGithub
 
    Call git.Upload_GitHub("https://api.github.com/repos/RolandTapie/ComptabiliteBudgetaireCommunale/contents/log.txt", "Upload log" & CStr(Now()), ReadAllTextFile("C:\Users\tallar\Downloads\log.txt"))
    Call logging(Now, Application.UserName, Application.Caption, "test", "Github.test")

    Set git = Nothing
End Sub

Function ReadAllTextFile(ByVal path As String) As String
    Const ForReading = 1, ForWriting = 2
    Dim fso, f
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set f = fso.OpenTextFile(path, ForReading)
    ReadAllTextFile = f.ReadAll
End Function

