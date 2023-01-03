VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Sub StockAnalysisFinal()

'Loop through all sheets
    For Each ws In Worksheets
    
    'Add Column Headers
        ws.Cells(1, 9).Value = "Ticker"
        ws.Cells(1, 10).Value = "Yearly Change"
        ws.Cells(1, 11).Value = "Percent Change"
        ws.Cells(1, 12).Value = "Total Stock Volume"
        ws.Cells(2, 14).Value = "Greatest % Increase"
        ws.Cells(3, 14).Value = "Greatest % Decrease"
        ws.Cells(4, 14).Value = "Greatest Total Volume"
        ws.Cells(1, 15).Value = "Ticker"
        ws.Cells(1, 16).Value = "Value"

    ''''SET VARIABLES''''
        ' Set an initial variable for holding the ticker
                Dim Ticker As String
          
        'Set an inital variable for holding the last row with reference to WS
                Dim LastRow As Long
                LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
    
        ' Set an initial variable for holding the volume per ticker
                Dim TickerVolume As Double
                TickerVolume = 0
    
        ' Keep track of the location for each ticker in the summary table
                Dim Summary_Table_Row As Integer
                Summary_Table_Row = 2
                
        'Set and define additional variables
                Dim StartPrice As Double
                StartPrice = 0
                
                Dim EndPrice As Double
                EndPrice = 0
                
                Dim YearlyPriceChange As Double
                YearlyPriceChange = 0
                
                Dim YearlyPriceChangePer As Double
                YearlyPriceChangePer = 0
                
                Dim MaxTickerName As String
                MaxTickerName = " "
                
                Dim MinTickerName As String
                MinTickerName = " "
                
                Dim MaxPercent As Double
                MaxPercent = 0
                
                Dim MinPercent As Double
                MinPercent = 0
                
                Dim MaxVolumeTickerName As String
                MaxVolumeTickerName = " "
                
                Dim MaxVolume As Variant
                MaxVolume = 0
                
                
    'Set initial value of beginning stock value for the first ticker
        StartPrice = ws.Cells(2, 3).Value

''''TICKER LOOP''''
    
    ' Loop through all tickers
    For i = 2 To LastRow
    
        '''''IF'''''
        
                ' Check if we are still within the same tickers, if it is not...
                If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
            
        '''''THEN'''''
        
                  ' Set the Ticker
                  Ticker = ws.Cells(i, 1).Value
                  
                  ' Add to the TickerVolume
                  TickerVolume = TickerVolume + ws.Cells(i, 7).Value
            
                  ' Print the Ticker in the Summary Table
                  ws.Range("I" & Summary_Table_Row).Value = Ticker
            
                  ' Print the TickerVolume to the Summary Table
                  ws.Range("L" & Summary_Table_Row).Value = TickerVolume
            
                  ' Add one to the summary table row
                  Summary_Table_Row = Summary_Table_Row + 1
                  
                  ' Reset the TickerVolume
                  TickerVolume = 0
            
        '''''ELSE'''''
                ' If the cell immediately following a row is the same ticker...
                Else
            
                ''''Ticker Volume''''
                  ' Add to the TickerVolume
                  TickerVolume = TickerVolume + ws.Cells(i, 7).Value
            
                End If

        '''''CLOSE LOOP''''''
        Next i

''''YEARLY CHANGE LOOP'''''

    'Loop through all tickers
    Summary_Table_Row = 2
    For i = 2 To LastRow
     
    '''''IF'''''
        
        ' Check if we are still within the same tickers, if it is not...
        If ws.Cells(i, 1).Value <> ws.Cells(i + 1, 1).Value Then
        
    '''''THEN'''''
    
            EndPrice = ws.Cells(i, 6).Value
    '''''ELSEIF'''''
    
        ElseIf ws.Cells(i, 1).Value <> ws.Cells(i - 1, 1).Value Then
            StartPrice = ws.Cells(i, 3).Value
    
    '''''END IF''''
        End If
        
    '''''IF'''''
        If StartPrice > 0 And EndPrice > 0 Then
    
        '''''THEN'''''
        YearlyPriceChange = EndPrice - StartPrice
        YearlyPriceChangePer = YearlyPriceChange / StartPrice
        
        'Print PriceChangeValues in cells
        ws.Cells(Summary_Table_Row, 10).Value = YearlyPriceChange
        ws.Cells(Summary_Table_Row, 11).Value = YearlyPriceChangePer
        ws.Cells(Summary_Table_Row, 11).NumberFormat = "0.00%"
        
        'Reset Price Values
        StartPrice = 0
        EndPrice = 0
        
        'Add row to summary table
        Summary_Table_Row = Summary_Table_Row + 1
        
    '''''END IF''''
        End If
  
    '''''CLOSE LOOP''''''
    Next i

'Find min and max values and assign to cells
    MaxPercent = WorksheetFunction.Max(ws.Range("K2:K" & LastRow))
    MinPercent = WorksheetFunction.Min(ws.Range("K2:K" & LastRow))
    MaxVolume = WorksheetFunction.Max(ws.Range("L2:L" & LastRow))
    ws.Cells(2, 16).Value = MaxPercent
    ws.Cells(2, 16).NumberFormat = "0.00%"
    ws.Cells(3, 16).Value = MinPercent
    ws.Cells(3, 16).NumberFormat = "0.00%"
    ws.Cells(4, 16).Value = MaxVolume

'''''FIND TICKER LOOP'''''
    

    'Loop through all tickers
    For i = 2 To LastRow
    
    'If statement WITH END IF
        If MaxPercent = ws.Cells(i, 11).Value Then
            ws.Cells(2, 15).Value = ws.Cells(i, 9).Value
        ElseIf MinPercent = ws.Cells(i, 11).Value Then
            ws.Cells(3, 15).Value = ws.Cells(i, 9).Value
        ElseIf MaxVolume = ws.Cells(i, 12).Value Then
            ws.Cells(4, 15).Value = ws.Cells(i, 9).Value
        End If

    '''''CLOSE LOOP''''''
    Next i
    
    
'''''RED OR GREEN INTERIOR LOOP'''''
    For i = 2 To LastRow
    
    'If statement WITH END IF
        If IsEmpty(ws.Cells(i, 10).Value) Then Exit For
        If ws.Cells(i, 10).Value > 0 Then
            ws.Cells(i, 10).Interior.ColorIndex = 4
        Else
            ws.Cells(i, 10).Interior.ColorIndex = 3
        End If
    
    '''''CLOSE LOOP''''''
    Next i
    

'''''Next WS''''''
Next ws

End Sub
 

