Sub SheetA():
  
'   Loop through all sheets.

        For Each ws In Worksheets
  
  Dim ticker_name As String

'   Set an initial variable for holding the total per ticker.
Dim ticker_total As Double
ticker_total = 0

'   Keep track of each ticker category in the summary table.
Dim Summary_Table_Row As Integer
Summary_Table_Row = 2

'   Count number of rows.

LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row

' Set year open

    Dim Open_Price As Double
    Open_Price = ws.Cells(2, 3).Value

' Loop through all tickers.
For i = 2 To LastRow

    '   Set ticker name.
    ticker_name = ws.Cells(i, 1).Value
    
    '   Check if we are still in the same ticker category, and if it is not...
    
    If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
    
          '   Add to the ticker total.
    ticker_total = ticker_total + ws.Cells(i, 7).Value
    
    '   Print the ticker category in the Summary Table.
    ws.Cells(Summary_Table_Row, 10).Value = ticker_name
    
    '   Print the ticker amount into the Summary Table.
    ws.Cells(Summary_Table_Row, 13).Value = ticker_total
    
    '   Reset the ticker total.
    ticker_total = 0
  
                 '  Insert total yearly difference and percentage total.
    
                            Year_close = ws.Cells(i, 6).Value
     
                            Yearly_Change = Year_close - Open_Price
     
                            Range("K" & Summary_Table_Row).Value = Yearly_Change
     
                            Range("L" & Summary_Table_Row).Value = Yearly_Change / Open_Price
     
                            Range("L:L").NumberFormat = "0.00%"
     
         '   Add 1 to the Summary Table.
    
    Summary_Table_Row = Summary_Table_Row + 1
     
     '  If the cell immediately following a row is the same ticker...
     
         Else
    
      '   Add to the ticker total.
    ticker_total = ticker_total + ws.Cells(i, 7).Value
     
        End If
     
        Next i
        
    Next ws
    
End Sub