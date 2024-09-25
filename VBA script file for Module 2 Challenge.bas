Attribute VB_Name = "Module1"
Sub QuarterlyStockAnalysis()

    ' Initialize all VBA variables
    Dim ws As Worksheet
    Dim ticker As String
    Dim openPrice As Double
    Dim closePrice As Double
    Dim totalVolume As Double
    Dim lastRow As Long
    Dim i As Long
    Dim quarterlyChange As Double
    Dim percentChange As Double
    Dim outputRow As Long
    Dim greatestIncrease As Double
    Dim greatestDecrease As Double
    Dim greatestVolume As Double
    Dim greatestIncreaseTicker As String
    Dim greatestDecreaseTicker As String
    Dim greatestVolumeTicker As String

    ' Loop through each worksheet
    For Each ws In ThisWorkbook.Worksheets
    
        ' Find the last row of data
        lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
        
        ' Set the initial output row for results
        outputRow = 2
        
        ' Add headers for the output columns
        ws.Cells(1, 9).Value = "Ticker"
        ws.Cells(1, 10).Value = "Quarterly Change"
        ws.Cells(1, 11).Value = "Percent Change"
        ws.Cells(1, 12).Value = "Total Stock Volume"
        ws.Cells(1, 15).Value = "Ticker"
        ws.Cells(1, 16).Value = "Value"
        
        ' Initialize the values for greatest increase, decrease, and volume
        greatestIncrease = 0
        greatestDecrease = 0
        greatestVolume = 0
        
        i = 2
        Do While i <= lastRow
            ticker = ws.Cells(i, 1).Value
            openPrice = ws.Cells(i, 3).Value
            totalVolume = 0
            
            ' Loop through the ticker and sum volume until the next ticker or end of sheet
            Do While ws.Cells(i, 1).Value = ticker
                totalVolume = totalVolume + ws.Cells(i, 7).Value
                closePrice = ws.Cells(i, 6).Value
                i = i + 1
                If i > lastRow Then Exit Do
            Loop
            
            ' Calculate the quarterly change and percent change
            quarterlyChange = closePrice - openPrice
            percentChange = (quarterlyChange / openPrice)
            
            ' Output the data to columns I, J, K, and L
            ws.Cells(outputRow, 9).Value = ticker
            ws.Cells(outputRow, 10).Value = quarterlyChange
            ws.Cells(outputRow, 11).Value = percentChange
            ws.Cells(outputRow, 12).Value = totalVolume
            outputRow = outputRow + 1
            
            ' Check for the greatest increase, decrease, and total volume
            If percentChange > greatestIncrease Then
                greatestIncrease = percentChange
                greatestIncreaseTicker = ticker
            End If
            
            If percentChange < greatestDecrease Then
                greatestDecrease = percentChange
                greatestDecreaseTicker = ticker
            End If
            
            If totalVolume > greatestVolume Then
                greatestVolume = totalVolume
                greatestVolumeTicker = ticker
            End If
        Loop
        
        ' Output the greatest % increase, % decrease, and total volume
        ws.Cells(2, 14).Value = "Greatest % Increase"
        ws.Cells(2, 15).Value = greatestIncreaseTicker
        ws.Cells(2, 16).Value = greatestIncrease
        
        ws.Cells(3, 14).Value = "Greatest % Decrease"
        ws.Cells(3, 15).Value = greatestDecreaseTicker
        ws.Cells(3, 16).Value = greatestDecrease
        
        ws.Cells(4, 14).Value = "Greatest Total Volume"
        ws.Cells(4, 15).Value = greatestVolumeTicker
        ws.Cells(4, 16).Value = greatestVolume
        
        ' Apply conditional formatting for positive/negative changes in the quarterly change column (J)
        With ws.Range(ws.Cells(2, 10), ws.Cells(outputRow - 1, 10)).FormatConditions.Add(Type:=xlCellValue, Operator:=xlGreater, Formula1:="=0")
            .Interior.Color = RGB(0, 255, 0) ' Green
        End With
        With ws.Range(ws.Cells(2, 10), ws.Cells(outputRow - 1, 10)).FormatConditions.Add(Type:=xlCellValue, Operator:=xlLess, Formula1:="=0")
            .Interior.Color = RGB(255, 0, 0) ' Red
        End With
        

    Next ws

    MsgBox "Quarterly stock analysis complete for all sheets!"

End Sub

