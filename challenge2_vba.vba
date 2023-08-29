Sub yearlySumary(year)

Dim arr As Variant

LastRowTotal = ThisWorkbook.Worksheets(year).Cells(Rows.Count, "A").End(xlUp).Row


Set data = ThisWorkbook.Worksheets(year)

total = 0
Change = 0
Start = 2
loopCount = 2

For i = 2 To LastRowTotal

    If data.Cells(i + 1, 1).Value <> data.Cells(i, 1).Value Then
    
    totalVolume = Cells(i, 7)
    
    opening = data.Cells(Start, 3)
    
    closing = data.Cells(i, 6)
    
    Change = closing - opening
    
    ticker = data.Cells(Start, "A")
    
    pt = ((closing / opening) - 1)
    
    data.Cells(loopCount, "I") = ticker
    
    
    data.Cells(loopCount, "J") = Change
    
    data.Cells(loopCount, "K") = pt
    
    data.Cells(loopCount, "L") = _
    Application.WorksheetFunction.sumIf(data.Range("A2:A" & LastRowTotal), "=" & ticker, data.Range("G2:G" & LastRowTotal))
    
    loopCount = loopCount + 1
    
    Start = i + 1
    
    Else
    
    End If
       
Next i


LastRowGrouped = data.Cells(Rows.Count, "I").End(xlUp).Row

arr = data.Range("I2:L" & LastRowGrouped)

ptRng = data.Range("K2:K" & LastRowGrouped)

volumeRng = data.Range("L2:L" & LastRowGrouped)


ptIncrease = Application.WorksheetFunction.Max(ptRng)

ptDecrease = Application.WorksheetFunction.Min(ptRng)

highestVolume = Application.WorksheetFunction.Max(volumeRng)


indexIncrease = Application.Match(ptIncrease, Application.Index(arr, 0, 3), 0)

indexDecrease = Application.Match(ptDecrease, Application.Index(arr, 0, 3), 0)

indexVolume = Application.Match(highestVolume, Application.Index(arr, 0, 4), 0)


data.Cells(2, "P") = arr(indexIncrease, 1)
data.Cells(3, "P") = arr(indexDecrease, 1)
data.Cells(4, "P") = arr(indexVolume, 1)

data.Cells(2, "Q") = ptIncrease
data.Cells(3, "Q") = ptDecrease
data.Cells(4, "Q") = highestVolume

End Sub


Sub runAllSheets()

WS_Count = ThisWorkbook.Worksheets.Count

For i = 1 To WS_Count
    wsName = ThisWorkbook.Worksheets(i).Name
    yearlySumary (wsName)
Next i

End Sub
