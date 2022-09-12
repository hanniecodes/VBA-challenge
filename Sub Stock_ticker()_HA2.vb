Sub Stock_ticker()
Dim ws As Worksheet
For Each ws In Worksheets

'define Lastrow, getting the last row
Lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row

'Name the colums
    ws.Range("I1").Value = "Ticker"
    ws.Range("J1").Value = "Yearly Change"
    ws.Range("K1").Value = "Percent Change"
    ws.Range("L1").Value = "Total Stock Volume"


' Set an initial variable for holding the ticker name
Dim Ticker As String
Dim TickerRow As Long
TickerRow = 1

Dim outputrow As Integer
outputrow = 2



Dim percent As Double


  ' Set an initial variable for holding the total volume per ticker
Dim volume As Double

    
  ' Keep track of the location for each ticker in the summary table
Dim Summary_Table_Row As Integer
Summary_Table_Row = 2

' Set an initial variable for holding the year open and closed
Dim year_open As Double
Dim year_close As Double
'Can't grab this in the loops so grab it here
    'year_open = ws.Cells(i, 3).Value



  ' Loop through all credit card purchases
For i = 2 To Lastrow
    
   If ws.Cells(i - 1, 1).Value <> ws.Cells(i, 1).Value Then
        year_open = ws.Cells(i, 3).Value
        volume = 0
        
        
    End If
    volume = volume + Cells(i, 7).Value

    
    
   ' Check if we are still within the same ticker, if it is not...
   If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
        TickerRow = TickerRow + 1
      ' Set the ticker
        Ticker = ws.Cells(i, 1).Value
        
        'grabb close value and do subtraction then write it to ticker row
        year_close = ws.Cells(i, 6).Value
         'Last row off AAB
        year_diff = year_close - year_open
     
              ws.Cells(TickerRow, 10).Value = year_diff
      
         ' Add to the Ticker
         ws.Cells(TickerRow, 9).Value = Ticker
         
         'Color should go here*******
        If ws.Range("J" & outputrow).Value > 0 Then
          ws.Range("J" & outputrow).Interior.ColorIndex = 4
            End If
              
        If ws.Range("J" & outputrow).Value < 0 Then
            ws.Range("J" & outputrow).Interior.ColorIndex = 3
            End If
                
         'Calculation
        percent = (year_diff / year_open)
        
         'Add the volume and percent
         ws.Range("L" & outputrow).Value = volume
         ws.Range("K" & outputrow).Value = percent
            
            outputrow = outputrow + 1
     
        
        ' Add to Percent
        ws.Range("K" & outputrow).NumberFormat = "0.00%"
    
            
    End If
            

    

  Next i
  Next ws

End Sub