Sub Stock():

Dim Ticker As String
Dim YearlyChange As Double
Dim PercentageChange As Double
Dim TotalStockVolume As Double
Dim row As Integer
Dim lastrow As Long
Dim Summary_Table_Row As Integer
Dim OpeningPrice As Double
Dim ClosingPrice As Double
Dim GreatestIncrease As Double
Dim GreatestIncreasePercentage As String
Dim GreatestDecrease As Double
Dim GreatestVolume As Double
Dim GreatestIncreaseRow As Double
Dim GreatestDecreaseRow As Double
Dim GreatestVolumeRow As Double

For Each ws In Worksheets

    Dim Worksheet As String
    
    lastrow = ws.Cells(Rows.Count, 1).End(xlUp).row
    WorksheetName = ws.Name
    Ticker = ws.Cells(2, 1).Value
    TotalStockVolume = 0
    Summary_Table_Row = 2
    OpeningPrice = ws.Cells(2, 3).Value
    ClosingPrice = 0
    

    ws.Cells(1, 10).Value = "Ticker"
    ws.Cells(1, 11).Value = "YearlyChange"
    ws.Cells(1, 12).Value = "PercentageChange"
    ws.Cells(1, 13).Value = "TotalStockVolume"


    For row = 2 To lastrow
        'Setting/Printing Summary table
        If Cells(row + 1, 1).Value <> Cells(row, 1).Value Then
       
        ClosingPrice = ws.Cells(row, 6).Value
        ws.Cells(Summary_Table_Row, 10).Value = Ticker
        ws.Cells(Summary_Table_Row, 11).Value = ClosingPrice - OpeningPrice
        ws.Cells(Summary_Table_Row, 13).Value = TotalStockVolume + ws.Cells(row, 7).Value
        
                'Checking to see for percentage change divided by 0
                If OpeningPrice <> 0 Then
                ws.Cells(Summary_Table_Row, 12).Value = Format(ws.Cells(Summary_Table_Row, 11).Value / OpeningPrice, "Percent")
            
                Else
                ws.Cells(Summary_Table_Row, 12).Value = "Can't divide by 0 so null value as Opening Price=0"
                
                 End If
             
                'Conditional formatting of positive to green and negative to red
                If ClosingPrice - OpeningPrice > 0 Then
                ws.Cells(Summary_Table_Row, 11).Interior.ColorIndex = 4
                
                Else
                ws.Cells(Summary_Table_Row, 11).Interior.ColorIndex = 3
                
                End If
             
        Ticker = ws.Cells(row + 1, 1).Value
        OpeningPrice = ws.Cells(row + 1, 3).Value
        TotalStockVolume = ws.Cells(row + 1, 7).Value
        Summary_Table_Row = Summary_Table_Row + 1
        
        Else
        Ticker = ws.Cells(row, 1).Value
        TotalStockVolume = TotalStockVolume + ws.Cells(row, 7).Value
        
        End If
        
    Next row
    
     'Setting/Printing Greatest Percentage Increase and Decrease
     ws.Cells(1, 16).Value = "Ticker"
     ws.Cells(1, 17).Value = "Value"
     ws.Cells(2, 15).Value = "Greatest % Increase"
     ws.Cells(3, 15).Value = "Greatest % Decrease"
     ws.Cells(4, 15).Value = "Greatest Total Volume"
     
     GreatestIncrease = WorksheetFunction.Max(ws.Range("L:L"))
     'GreatestIncreasePercentage = FormatPercent(GreatestIncrease, 2)
     For i = 2 To 15
        If Cells(i, 12).Value = GreatestIncrease Then
        ws.Cells(2, 16).Value=Cells(i,
        
     Next i
    
     'GreatestIncreaseRow = WorksheetFunction.Match(GreatestIncrease, ws.Range("L:L"), 0)
     'MsgBox (GreatestIncrease)
     'MsgBox (GreatestIncreaseRow)
     'ws.Range("Q2") = "%" & WorksheetFunction.Max(ws.Range("K2:K" & RowCount)) * 100
    'GreatestIncreaseRow=WorksheetFunction.Match(WorksheetFunction.Max(ws.Range("K2:K" & RowCount)), ws.Range("K2:K" & RowCount), 0)
    'ws.Range("P2") = ws.Cells(increase_number + 1, 9)
     'ws.Cells(2, 16).Value = ws.Cells(GreatestIncreaseRow, 9).Value
     'ws.Cells(2, 17).Value = Format(ws.Cells(GreatestIncreaseRow, 11).Value, "Percent")
     
     'GreatestDecrease = WorksheetFunction.Min(ws.Range("L2:L15"))
     'GreatestDecreaseRow = WorksheetFunction.Match(GreatestDecrease, ws.Range("L2:L15"), 0)
     'ws.Cells(4, 16).Value = ws.Cells(GreatestDecreaseRow, 9).Value
     'ws.Cells(4, 17).Value = Format(ws.Cells(GreatestDecreaseRow, 11).Value, "Percent")
     
     'GreatestVolume = WorksheetFunction.Max(ws.Range("M:M"))
     'GreatestVolumeRow = WorksheetFunction.Match(GreatestVolume, ws.Range("M:M"), 0)
     'ws.Cells(5, 16).Value = ws.Cells(GreatestVolumeRow, 9).Value
     'ws.Cells(5, 17).Value = Format(ws.Cells(GreatestVolumeRow, 12).Value, "Percent")
     
Next ws


End Sub



