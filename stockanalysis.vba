Sub stockanalysis()
Dim WS_count As Integer
Dim lastrow As Long
'loop through the sheets
WS_count = ActiveWorkbook.Worksheets.count
For j = 1 To WS_count
   Worksheets(j).Activate
   lastrow = Cells(Rows.count, "A").End(xlUp).Row
   'writing headers
    Cells(1, 9) = "Ticker"
    Cells(1, 10) = "Total Stock Volume"
    Cells(1, 16) = "Ticker"
    Cells(1, 17) = "Value"
    Cells(1, 11) = "Yearly Change"
    Cells(1, 12) = "Percent Change"
    Cells(2, 15) = "Greatest % increase"
    Cells(3, 15) = "Greatest % dencrease"
    Cells(4, 15) = "Greatest Total Value"
    'variable for total stock volume
    Dim ticker As String
    Dim volume As Double
    Dim g As Integer
    volume = 0
    g = 2
    'variables for yearly and percent change
    Dim openprice As Double
    Dim closeprice As Double
    Dim count As Integer
    Dim change As Double
    Dim percentchange As Double
    Dim n As Integer
    count = 0
    n = 2
    'variables for greatest increase, decrease and total value
    Dim Gincreas As Double
    Dim Gdecrease As Double
    Dim tickeri As String
    Dim tickerd As String
    Dim Gtotal As Double
    Dim tickert As String
    
    'calculate total stock volume for each year
    For i = 2 To lastrow
        If Cells(i, 1).Value = Cells(i + 1, 1) Then
           volume = volume + Cells(i, 7)
        Else
           ticker = Cells(i, 1).Value
           volume = volume + Cells(i, 7)
           Cells(g, 9) = ticker
           Cells(g, 10) = volume
           g = g + 1
           volume = 0
       End If
       
    Next i
    'yearly change and percentage
    For y = 2 To lastrow
        If Cells(y, 1) = Cells(n, 9) Then
            count = count + 1
        'get the closing price and calculate yearly change
        Else
            count = 1
            closeprice = Cells(y - 1, 6)
            change = closeprice - openprice
        'calculate percent change
        If openprice <> 0 Then
            percentchange = change / openprice
        Else
            percentchange = 0
        End If
        Cells(n, 11) = change
        Cells(n, 12) = percentchange
        n = n + 1
        End If
        'get opening price
        If count = 1 Then
            openprice = Cells(y, 3)
        End If
    Next y
    Gincrease = 0
    Gdecrease = 0
    Gtotal = 0
    'get the values of greatest increase, decrease and total value
    For t = 2 To lastrow

        If Cells(t, 12) > Gincrease Then
            Gincrease = Cells(t, 12).Value
            tickeri = Cells(t, 9)
        ElseIf Cells(t, 12) < Gdecrease Then
            Gdecrease = Cells(t, 12)
            tickerd = Cells(t, 9)
        End If
        If Cells(t, 10) > Gtotal Then
           Gtotal = Cells(t, 10).Value
           tickert = Cells(t, 9)
        End If

    Next t
    Cells(2, 17) = Gincrease
    Cells(2, 16) = tickeri
    Cells(3, 17) = Gdecrease
    Cells(3, 16) = tickerd
    Cells(4, 17) = Gtotal
    Cells(4, 16) = tickert
    ' percent change format
    Columns("L:L").Select
    Selection.NumberFormat = "0.00%"
    'greatest values format
    Range("Q2:Q3").Select
    Selection.NumberFormat = "0.00%"
    Range("Q4").Select
    Selection.NumberFormat = "0"
    'yearly change color conditional formate
    Range("K2").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.FormatConditions.Add Type:=xlCellValue, Operator:=xlLess, _
        Formula1:="=0"
    Selection.FormatConditions(Selection.FormatConditions.count).SetFirstPriority
    With Selection.FormatConditions(1).Interior
        .PatternColorIndex = xlAutomatic
        .Color = 192
        .TintAndShade = 0
    End With
        Selection.FormatConditions.Add Type:=xlCellValue, Operator:=xlGreater, _
        Formula1:="=0"
    Selection.FormatConditions(Selection.FormatConditions.count).SetFirstPriority
    With Selection.FormatConditions(1).Interior
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorAccent6
        .TintAndShade = 0
    End With
    Selection.FormatConditions(1).StopIfTrue = False
Next j
End Sub
