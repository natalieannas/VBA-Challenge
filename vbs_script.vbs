VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Sub MultiyearStockData()

    Dim ws As Worksheet
    Dim lastrow As Long
    Dim I As Long
    Dim ticker As String
    Dim openingPrice As Double
    Dim closingPrice As Double
    Dim quarterlyChange As Double
    Dim percentChange As Double
    Dim totalVolume As Double
    Dim greatestIncrease As Double
    Dim greatestDecrease As Double
    Dim highestVolume As Double
    Dim percentIncreaseStock As String
    Dim percentDecreaseStock As String
    Dim greatestVolumeStock As String
    Dim outputRow As Long
    
    
    ' Loop through each worksheet in the workbook
    For Each ws In ThisWorkbook.Worksheets
        
        lastrow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
        totalVolume = 0
        openingPrice = ws.Cells(2, 3).Value
        
      ' Headers for the new columns
        ws.Cells(1, 9).Value = "Ticker"
        ws.Cells(1, 10).Value = "Quarterly Change"
        ws.Cells(1, 11).Value = "Percent Change"
        ws.Cells(1, 12).Value = "Total Volume"
        ws.Cells(1, 15).Value = "Ticker"
        ws.Cells(1, 16).Value = "Value"
        ws.Cells(2, 14).Value = "Greatest % Increase"
        ws.Cells(3, 14).Value = "Greatest % Decrease"
        ws.Cells(4, 14).Value = "Greatest Total Volume"
        
    ' Setting Values
    greatestIncrease = 0
    greatestDecrease = 0
    highestVolume = 0
    outputRow = 2
       
        
        
    ' Loop through each row of data
        For I = 2 To lastrow
            
            ' Check if the ticker changes
        If ws.Cells(I + 1, 1).Value <> ws.Cells(I, 1).Value Then
             ticker = ws.Cells(I, 1).Value
             closingPrice = ws.Cells(I, 6).Value
             totalVolume = totalVolume + ws.Cells(I, 7).Value
                
         ' Calculate Quarterly Changes
            quarterlyChange = closingPrice - openingPrice
                If openingPrice <> 0 Then
                    percentChange = ((closingPrice - openingPrice) / openingPrice)
                Else
                    percentChange = 0
                
                End If
                
                ws.Cells(outputRow, 11).NumberFormat = "0.00%"
                
        ' Write the results for the Quarter
                ws.Cells(outputRow, 9).Value = ticker
                ws.Cells(outputRow, 10).Value = quarterlyChange
                ws.Cells(outputRow, 11).Value = percentChange
                ws.Cells(outputRow, 12).Value = totalVolume
                
            
                
                ' Conditional Formatting
               If quarterlyChange > 0 Then
                    ws.Cells(outputRow, 10).Interior.ColorIndex = 4
                ElseIf quarterlyChange < 0 Then
                    ws.Cells(outputRow, 10).Interior.ColorIndex = 3
                Else
                    ws.Cells(outputRow, 10).Interior.Color = xlNone
                End If
                
                ' Check for greatest values
         If percentChange > greatestIncrease Then
             greatestIncrease = percentChange
                percentIncreaseStock = ticker
                End If
                
          If percentChange < greatestDecrease Then
               greatestDecrease = percentChange
                 percentDecreaseStock = ticker
                End If
                
          If totalVolume > highestVolume Then
            highestVolume = totalVolume
                greatestVolumeStock = ticker
                End If
            
            ws.Cells(2, 16).NumberFormat = "0.00%"
            ws.Cells(3, 16).NumberFormat = "0.00%"
                
        ' Reset values for the next ticker
             outputRow = outputRow + 1
           If I + 1 <= lastrow Then
           openingPrice = ws.Cells(I + 1, 3).Value
           End If
           totalVolume = 0
            Else
            'total volume
                totalVolume = totalVolume + ws.Cells(I, 7).Value
            End If
            
        Next I
        
        ' Write greatest values
        ws.Cells(2, 15).Value = percentIncreaseStock
        ws.Cells(2, 16).Value = greatestIncrease
        ws.Cells(3, 15).Value = percentDecreaseStock
        ws.Cells(3, 16).Value = greatestDecrease
        ws.Cells(4, 15).Value = greatestVolumeStock
        ws.Cells(4, 16).Value = highestVolume
        
    Next ws

End Sub


