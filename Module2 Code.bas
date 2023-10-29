Attribute VB_Name = "Module1"
Sub StockTicker()
   Dim ws As Worksheet
   For Each ws In ThisWorkbook.Worksheets
      Dim i As Long
      Dim lastrow As Long
      Dim Stock_Name As String
      Dim Volume_Total As Double
      Dim Open_Stock As Double
      Dim Close_Stock As Double
      Dim Year_Change As Double
      Dim Percent_Change As Double
      
      Volume_Total = 0
      
      ' Add headers for new columns
      ws.Range("I1").Value = "Ticker"
      ws.Range("J1").Value = "Year Change"
      ws.Range("K1").Value = "Percent Change"
      ws.Range("L1").Value = "Total Stock Volume"
      
      ' Keep track of the location for each stock in the summary table
      Dim Summary_Table_Row As Integer
      Summary_Table_Row = 2
      
      lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row
      
      For i = 2 To lastrow
          ' Check if we are still within the same stock brand
          If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
              Stock_Name = ws.Cells(i, 1).Value
              Open_Stock = ws.Cells(i, 3).Value
              Close_Stock = ws.Cells(i, 6).Value
              
              ' Add to the volume total
              Volume_Total = Volume_Total + ws.Cells(i, 7).Value
              
              ' Add year change formula
              Year_Change = Open_Stock - Close_Stock
              
              ' Add percent change formula
              If Open_Stock <> 0 Then
                  Percent_Change = Round((((Open_Stock - Close_Stock) / Open_Stock) * 100), 2)
              Else
                  Percent_Change = 0
              End If
              
              ' Print the stock name in the summary table
              ws.Cells(Summary_Table_Row, 9).Value = Stock_Name
              
              ' Print the year change in the summary table
              ws.Cells(Summary_Table_Row, 10).Value = Year_Change
              
              ' Print the percent change in the summary table
              ws.Cells(Summary_Table_Row, 11).Value = Percent_Change
              
              ' Print the volume total in the summary table
              ws.Cells(Summary_Table_Row, 12).Value = Volume_Total
              
              ' Change colors of Year Change based on the sign
              If Year_Change >= 0 Then
                  ws.Cells(Summary_Table_Row, 10).Interior.ColorIndex = 4 ' Green
              Else
                  ws.Cells(Summary_Table_Row, 10).Interior.ColorIndex = 3 ' Red
              End If
              
              ' Add 1 to the summary table row
              Summary_Table_Row = Summary_Table_Row + 1
              
              ' Reset the volume total
              Volume_Total = 0
          Else
              Volume_Total = Volume_Total + ws.Cells(i, 7).Value
          End If
      Next i
   Next ws
End Sub

