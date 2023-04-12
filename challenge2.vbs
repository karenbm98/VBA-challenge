Attribute VB_Name = "Module1"
Sub allworksheets()

    Dim x As Worksheet
    
    Application.ScreenUpdating = False
   
    For Each x In Worksheets
        x.Select
        Call challenge2
    Next
    
    Application.ScreenUpdating = True

End Sub

Sub challenge2()

    Dim openingprice As Single
    Dim closingprice As Single

    totalstockvolume = 0
    Position = 2

    Cells(1, 9).Value = "Ticker"
    Cells(1, 10).Value = "Yearly Change"
    Cells(1, 11).Value = "Percent Change"
    Cells(1, 12).Value = "Total Stock Volume"
    Cells(1, 16).Value = "Ticker"
    Cells(1, 17).Value = "Value"
    Cells(2, 15).Value = "Greatest % increase"
    Cells(3, 15).Value = "Greatest % decrease"
    Cells(4, 15).Value = "Greatest total volume"

    For i = 2 To 759001

        totalstockvolume = Cells(i, 7).Value + totalstockvolume

        If Cells(i, 2).Value = 20180102 Or Cells(i, 2).Value = 20190102 Or Cells(i, 2).Value = 20200102 Then

            openingprice = Cells(i, 3).Value

        ElseIf Cells(i, 2).Value = 20181231 Or Cells(i, 2).Value = 20191231 Or Cells(i, 2).Value = 20201231 Then

            closingprice = Cells(i, 6).Value

            ticker = Cells(i, 1).Value
            yearlychange = closingprice - openingprice
            percentchange = closingprice / openingprice - 1

            If yearlychange < 0 Then
                
                Cells(Position, 10).Interior.Color = RGB(255, 0, 0)
           
            ElseIf yearlychange > 0 Then
                
                Cells(Position, 10).Interior.Color = RGB(0, 255, 0)

            End If

            Cells(Position, 9).Value = ticker
            Cells(Position, 10).Value = Round(yearlychange, 2)
            Cells(Position, 11).Value = FormatPercent(percentchange, 2)
            Cells(Position, 12).Value = totalstockvolume

            Position = Position + 1
            totalstockvolume = 0

        End If

    Next i

    greatestincrease = Application.WorksheetFunction.Max(Range("K:K"))
    greatestincreaserow = WorksheetFunction.Match(greatestincrease, Range("K:K"), 0)
    greatestincreaseticker = Cells(greatestincreaserow, 9)

    greatestdecrease = Application.WorksheetFunction.Min(Range("K:K"))
    greatestdecreaserow = WorksheetFunction.Match(greatestdecrease, Range("K:K"), 0)
    greatestdecreaseticker = Cells(greatestdecreaserow, 9)

    greatesttotalvolume = Application.WorksheetFunction.Max(Range("L:L"))
    greatesttotalvolumerow = WorksheetFunction.Match(greatesttotalvolume, Range("L:L"), 0)
    greatesttotalvolumeticker = Cells(greatesttotalvolumerow, 9)

    Cells(2, 17).Value = FormatPercent(greatestincrease, 2)
    Cells(2, 16).Value = greatestincreaseticker
    Cells(3, 17).Value = FormatPercent(greatestdecrease, 2)
    Cells(3, 16).Value = greatestdecreaseticker
    Cells(4, 17).Value = greatesttotalvolume
    Cells(4, 16).Value = greatesttotalvolumeticker

End Sub
