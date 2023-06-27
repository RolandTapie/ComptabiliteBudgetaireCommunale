Attribute VB_Name = "Reporting"
Sub reporting(ByVal table As String)
    'TODOOK réaliser le reporting des codes économiques
    Dim i, j, k, l, m, dligne, dlign, dligne_eco As Integer
    Dim cat_1, cat_2, cat_lib, code, CG, libelle, types, Service, Regroupement As String
    Dim trouve As Boolean

    Sheets("Reporting").Cells.ClearContents

    Sheets("Reporting").Cells.Borders(xlDiagonalDown).LineStyle = xlNone
    Sheets("Reporting").Cells.Borders(xlDiagonalUp).LineStyle = xlNone
    Sheets("Reporting").Cells.Borders(xlEdgeLeft).LineStyle = xlNone
    Sheets("Reporting").Cells.Borders(xlEdgeTop).LineStyle = xlNone
    Sheets("Reporting").Cells.Borders(xlEdgeBottom).LineStyle = xlNone
    Sheets("Reporting").Cells.Borders(xlEdgeRight).LineStyle = xlNone
    Sheets("Reporting").Cells.Borders(xlInsideVertical).LineStyle = xlNone
    Sheets("Reporting").Cells.Borders(xlInsideHorizontal).LineStyle = xlNone
    
    
    Sheets("Reporting").Range("A:Z").UnMerge



    dligne_eco = Sheets(table).Cells(Rows.Count, 1).End(xlUp).Row

    ligne1 = 2
    ligne2 = 2
    ligne3 = 2

    k = 5

    If table <> "Correspondance" Then
        'TODOOK trier la table
        Call tri(table, "E")

        Sheets("Reporting").Select
        Sheets("Reporting").Cells(1, 1) = "Plan Comptable"
        Sheets("Reporting").Range("A1:g1").Merge
        Sheets("Reporting").Cells(1, 2) = table
        Range("A2:g2").Merge

        For i = 2 To dligne_eco
            ligne2 = i
            While Sheets(table).Cells(ligne2, 2) = Sheets(table).Cells(ligne2 + 1, 2) And ligne2 <= dligne_eco
                ligne2 = ligne2 + 1
            Wend
        
            X = k
        
            Sheets("Reporting").Cells(k, 1) = Sheets(table).Cells(i, 1)
            Sheets("Reporting").Range("A" & k & ":g" & k).Merge
            Sheets("Reporting").Range("A" & k).HorizontalAlignment = xlCenter
            Sheets("Reporting").Range("A" & k).VerticalAlignment = xlCenter
        
            k = k + 1
        
            Sheets("Reporting").Cells(k, 1) = Sheets(table).Cells(i, 2)
            Sheets("Reporting").Range("A" & k & ":g" & k).Merge
            Sheets("Reporting").Range("A" & k).HorizontalAlignment = xlCenter
            Sheets("Reporting").Range("A" & k).VerticalAlignment = xlCenter
        
            k = k + 1
        
            For m = 5 To 11
                Sheets("Reporting").Cells(k, m - 5 + 1) = Sheets(table).Cells(1, m)
            Next
            k = k + 1
        
            For j = i To ligne2
                Sheets("Reporting").Cells(k, 1) = Sheets(table).Cells(j, 5)
                Sheets("Reporting").Cells(k, 2) = Sheets(table).Cells(j, 6)
                Sheets("Reporting").Cells(k, 3) = Sheets(table).Cells(j, 7)
                Sheets("Reporting").Cells(k, 4) = Sheets(table).Cells(j, 8)
                Sheets("Reporting").Cells(k, 5) = Sheets(table).Cells(j, 9)
                Sheets("Reporting").Cells(k, 6) = Sheets(table).Cells(j, 10)
                Sheets("Reporting").Cells(k, 7) = Sheets(table).Cells(j, 11)
                k = k + 1
            Next
        
            Call quadrillage(Range("A" & X & ":g" & k))
        
            i = ligne2 + 1
            k = k + 3
        Next
        Worksheets("Reporting").PageSetup.PrintArea = "$A$1:$G$" & k
        'Worksheets("Reporting").PageSetup.Orientation = xlPortrait
        Worksheets("Reporting").PageSetup.Orientation = xlLandscape
    Else
        'TODOOK trier la table
        Call tri(table, "B")
    
        Sheets("Reporting").Select
        Sheets("Reporting").Cells(1, 1) = "Table de correspondance des Plans Comptables"
        Sheets("Reporting").Range("A1:g1").Merge
        Sheets("Reporting").Cells(1, 2) = "Liaison"
        Range("A2:g2").Merge
    
        For i = 2 To dligne_eco
            Sheets("Reporting").Cells(k, 1) = Sheets(table).Cells(i, 2)
            Sheets("Reporting").Cells(k, 2) = Sheets(table).Cells(i, 3)
            Sheets("Reporting").Cells(k, 3) = Sheets(table).Cells(i, 4)
            Sheets("Reporting").Cells(k, 4) = Sheets(table).Cells(i, 5)
            Sheets("Reporting").Cells(k, 5) = Sheets(table).Cells(i, 6)
            Sheets("Reporting").Cells(k, 6) = Sheets(table).Cells(i, 7)
            Sheets("Reporting").Cells(k, 7) = Sheets(table).Cells(i, 8)
            Sheets("Reporting").Cells(k, 8) = Sheets(table).Cells(i, 9)
            Sheets("Reporting").Cells(k, 9) = Sheets(table).Cells(i, 10)
            Sheets("Reporting").Cells(k, 10) = Sheets(table).Cells(i, 11)
            Sheets("Reporting").Cells(k, 11) = Sheets(table).Cells(i, 12)
            Sheets("Reporting").Cells(k, 12) = Sheets(table).Cells(i, 13)
            k = k + 1
        Next
        Worksheets("Reporting").PageSetup.PrintArea = "$A$1:$L$" & dligne_eco
    
        Call quadrillage(Range("$A$1:$L$" & dligne_eco))
    
        Sheets("Reporting").PageSetup.Orientation = xlLandscape
        Sheets("Reporting").Columns("A:Z").AutoFit
        Sheets("Reporting").Range("B1").EntireColumn.ColumnWidth = 30
        Sheets("Reporting").Range("F1").EntireColumn.ColumnWidth = 30
        Sheets("Reporting").Range("k1").EntireColumn.ColumnWidth = 30
    End If


    Call SimpleImpressionEnPDF("Reporting", table)
    Call logging(Now, Application.UserName, Application.Caption, "Reporting" & table, "Reporting.reporting")
End Sub

Sub test()
    Call reporting("ECO")
End Sub

Sub quadrillage(ByVal emplacement As Range)

    'emplacement.Select
    Application.CutCopyMode = False
    emplacement.Borders(xlDiagonalDown).LineStyle = xlNone
    emplacement.Borders(xlDiagonalUp).LineStyle = xlNone
    With emplacement.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With emplacement.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With emplacement.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With emplacement.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With emplacement.Borders(xlInsideVertical)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With emplacement.Borders(xlInsideHorizontal)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    'emplacement.Select
End Sub

Sub tri(ByVal table As String, ByVal colonne As String)

    dligne = Sheets(table).Cells(Rows.Count, 1).End(xlUp).Row
    
    Sheets(table).Select
    Range("A1:J" & dligne).Select
    ActiveWorkbook.Worksheets(table).Sort.SortFields.Clear
    ActiveWorkbook.Worksheets(table).Sort.SortFields.Add2 Key:=Range(colonne & "2:" & colonne & dligne), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:= _
        xlSortTextAsNumbers
    With ActiveWorkbook.Worksheets(table).Sort
        .SetRange Range("A1:J" & dligne)
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    Range("A1").Select
    Sheets("Reporting").Select
    Call logging(Now, Application.UserName, Application.Caption, "Tri", table)
End Sub

Sub testtri()
    Call tri("ECO", "E")
End Sub


