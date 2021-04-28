Attribute VB_Name = "Module1"
Sub StockTicker()

  Dim Ticker As String

  Dim TotalVolume As Double
  TotalVolume = 0

' Keep track of the location for each ticker in the summary table
  Dim Summary_Table_Row As Integer
  Summary_Table_Row = 2

' Yearly change variables needed to be created and initiated for compute price change and percent change
  Dim OpenPrice As Double
  OpenPrice = Range("C2").Value  'Set initial price
  Dim ClosePrice As Double
  Dim PriceChange As Double
    
' Declare a variable for last row
  Dim LastRow As Long
  LastRow = Cells(Rows.Count, 1).End(xlUp).Row   'In lieu of hardcoding ending number, this finds and stores the value of the last row in variable LastRow for use in For Loop
  
' Loop through all stocks
  For i = 2 To LastRow
  
' Determine the change in ticker in column 1
    If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
  
' Upon a change in ticker, this will update and set the ticker name
        Ticker = Cells(i, 1).Value

' Aggregate stock volume
        TotalVolume = TotalVolume + Cells(i, 7).Value
  
' Output ticker name and volume to the summary table
        Range("I" & Summary_Table_Row).Value = Ticker
        Range("L" & Summary_Table_Row).Value = TotalVolume
        Range("L" & Summary_Table_Row).NumberFormat = "#,###"  'Format number with 000 separator
                
' Update close price at last day of the year
        ClosePrice = Cells(i, 6).Value
        
' Calculates the price change between open price and close price
        PriceChange = (ClosePrice - OpenPrice)
          
' Output price change to summary table
        Range("J" & Summary_Table_Row).Value = PriceChange
                
' Add one row to the summary table
        Summary_Table_Row = Summary_Table_Row + 1
        
' Reset TotalVolume to zero in order to aggregate the next stock ticker
        TotalVolume = 0
        
' Reset and update open price for the next stock ticker
        OpenPrice = Cells(i + 1, 3).Value
  
    Else
  
' Add to TotalVolume when the cell immediately following a row has the same ticker
        TotalVolume = TotalVolume + Cells(i, 7).Value
        
        
        
        
  
  
    End If

  Next i



End Sub
