Attribute VB_Name = "Impression"
Sub SimpleImpressionEnPDF(ByVal table As String, ByVal nomFichier As String)
    Sheets(table).ExportAsFixedFormat Type:=xlTypePDF, Filename:="C:\Users\tallar\Downloads\" & nomFichier & ".pdf", Quality:=xlQualityStandard, _
                                      IncludeDocProperties:=False, IgnorePrintAreas:=False, OpenAfterPublish:=True
    Call logging(Now, Application.UserName, Application.Caption, "Impression_data", "Impression.SimpleImpressionEnPDF")
End Sub

Sub test()
    Call SimpleImpressionEnPDF("Reporting", "code économique")
    Call logging(Now, Application.UserName, Application.Caption, "test", "Impression.test")
End Sub


