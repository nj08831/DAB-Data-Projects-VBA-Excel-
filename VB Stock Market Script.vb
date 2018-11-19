Sub Tot_Volume()

'PART 1 - CREATE TABLE OF VOLUME SUMS BY TICKER AND YEAR


Dim r As Double
Dim Tot_Volume As Double
Dim counter As Long
Dim stock_price_start As Double
Dim stock_price_end As Double
Dim perc_change As Double
Dim stock_change As Double

'Set initial counters
Tot_Volume = 0
counter = 2
stock_price_start = 0


'Loop through all sheets
Dim ws As Variant

    For Each ws In Worksheets
        ' Find the last row of the combined sheet after each paste
        ' Add 1 to get first empty row
       
Dim lastRow As Double

    lastRow = Cells(Rows.Count, "D").End(xlUp).Row + 1
    'MsgBox (lastRow)
       
       
'Check the WorksheetName the process is working on
Dim WorksheetName As Variant

    WorksheetName = ws.Name
    'MsgBox WorksheetName
       
'Set initial counters
Tot_Volume = 0
counter = 2
stock_price_start = 0

ws.Cells(1, 9) = "Year"
ws.Cells(1, 10) = "Ticker"
ws.Cells(1, 11) = "Stock Change"
ws.Cells(1, 12) = "Percent Change %"
ws.Cells(1, 13) = "Total Stock Volume"
ws.Cells(1, 15) = "Year"
ws.Cells(1, 16) = "Ticker"

ws.Range("I1:P1").Font.Bold = True
ws.Range("I1:P1").HorizontalAlignment = xlRight

For r = 3 To lastRow

'fixing the start point

If r = 3 Then
stock_price_start = ws.Cells(r - 1, 3)
'MsgBox (stock_price_start)
End If


'If the ticker is equal to the prior dates ticker then total the volume

   If ws.Cells(r, 1) = ws.Cells(r - 1, 1) And Left(ws.Cells(r, 2), 4) = Left(ws.Cells(r - 1, 2), 4) Then
   
      Tot_Volume = Tot_Volume + ws.Cells(r, 7)
   
   ElseIf ws.Cells(r, 1) <> ws.Cells(r - 1, 1) Or Left(ws.Cells(r, 2), 4) <> Left(ws.Cells(r - 1, 2), 4) Then
          
'Assign ticker to a summary table and the total volume
   
   'ticker
   ws.Cells(counter, 10) = ws.Cells(r - 1, 1)
   ws.Cells(counter, 10).HorizontalAlignment = xlRight
   
   'total volume
   ws.Cells(counter, 13) = Tot_Volume
   ws.Cells(counter, 13).HorizontalAlignment = xlRight
   
   'year
   ws.Cells(counter, 9) = Left(ws.Cells(r - 1, 2), 4)
   ws.Cells(counter, 9).HorizontalAlignment = xlRight
   
   'assign stock price end and change
   stock_price_end = ws.Cells(r - 1, 6)
   
   'Check for initial stock price = 0
   If stock_price_start = 0 Then
   perc_change = 0
   Else
   perc_change = (stock_price_end - stock_price_start) / stock_price_start
   End If
   
   stock_change = stock_price_end - stock_price_start
   
   'MsgBox (stock_price_end)
   'MsgBox (perc_change)
   
   
   ws.Cells(counter, 12) = perc_change
   ws.Cells(counter, 12).NumberFormat = "0.00%"
   ws.Cells(counter, 12).HorizontalAlignment = xlRight
  
   ws.Cells(counter, 11) = stock_change
   ws.Cells(counter, 11).NumberFormat = "0.00"
   ws.Cells(counter, 11).HorizontalAlignment = xlRight
   
'reset counter
   counter = counter + 1
   Tot_Volume = ws.Cells(r, 7)
   stock_price_start = ws.Cells(r, 3)
   
   
End If

Next r

'PART II -- IDENTIFY AND INPUT VALUES - % greatest, % least, and maximum volume


   Dim TotalRows As Long

    TotalRows = ws.Cells(Rows.Count, 10).End(xlUp).Row
    'MsgBox ("total rows are" + Str(TotalRows))
   
   
Dim greatest As Double
Dim least As Double
Dim maximum As Double

Dim j As Double


Dim tickm As String
Dim tickl As String
Dim tickv As String

Dim yearm As Long
Dim yearl As Long
Dim yearv As Long

 
    greatest = 0
    least = 0
    maximum = 0
    
    For j = 2 To TotalRows
       
'percent change most
       
       If ws.Cells(j, 12) > greatest Then
       greatest = ws.Cells(j, 12).Value
       tickm = ws.Cells(j, 10).Value
       yearm = ws.Cells(j, 9).Value
        ' MsgBox (tickm)
       End If
       
'percent change least
       
       If ws.Cells(j, 12) < least Then
       least = ws.Cells(j, 12).Value
       tickl = ws.Cells(j, 10).Value
       yearl = ws.Cells(j, 9).Value
       End If
       
       If ws.Cells(j, 13) > maximum Then
       maximum = ws.Cells(j, 13).Value
       tickv = ws.Cells(j, 10).Value
       yearv = ws.Cells(j, 9).Value
       End If
       
    Next j
    
    
'greatest values
'MsgBox (tickm)
'MsgBox (greatest)

ws.Range("O2:Q4").HorizontalAlignment = xlRight
ws.Cells(2, 17).NumberFormat = "0.00%"
ws.Cells(3, 17).NumberFormat = "0.00%"

    ws.Cells(2, 17) = greatest
    ws.Cells(2, 16) = tickm
    ws.Cells(2, 15) = yearm
    
'least values
    ws.Cells(3, 17) = least
    ws.Cells(3, 16) = tickl
    ws.Cells(3, 15) = yearl
    
'most volume
    ws.Cells(4, 17) = maximum
    ws.Cells(4, 16) = tickv
    ws.Cells(4, 15) = yearv

Next ws
End Sub



