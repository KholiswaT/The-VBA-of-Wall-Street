Option Explicit

Sub ticker()


' Define variables


 Dim ws As Worksheet

' insert for loop for each worksheet after defining ws as worksheet
'for some reason, previous attempts didn't work if for loop was right before if statment

For Each ws In ThisWorkbook.Sheets


 Dim i As Long

 Dim ticker_type As String

 Dim summary_table_row As Integer

 Dim stockvol As Double

 Dim lastrow As Long

 Dim ticker_open As Double


 Dim ticker_closed As Double


 Dim yearchange As Double

 Dim percentchange As Double

 Dim maxticker As Long




     summary_table_row = 2

     stockvol = 0


' Define headers for each output

    ws.Cells(1, 9).Value = "Ticker"


    ws.Cells(1, 10).Value = "Yearly Change"


    ws.Cells(1, 11).Value = "Percent Change"


    ws.Cells(1, 12).Value = "Total Stock Volume"

    ws.Cells(2, 15).Value = "Greatest % Increase"

    ws.Cells(3, 15).Value = "Greatest % Decrease"

    ws.Cells(4, 15).Value = "Greatest Total Volume"

    ws.Cells(1, 16).Value = "Ticker"
    
    ws.Cells(1, 17).Value = "Value"

'Define last row within each dataset

 lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row


    For i = 2 To lastrow


 'Define opening price value for start date, which is the first value for each stock

    If ticker_open = 0 Then
        ticker_open = ws.Cells(i, 3).Value
    End If


 
    If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then

    ticker_type = ws.Cells(i, 1).Value


    stockvol = stockvol + ws.Cells(i, 7).Value

    
    ws.Range("I" & summary_table_row).Value = ticker_type


    ws.Range("L" & summary_table_row).Value = stockvol

  
   ticker_closed = ws.Cells(i, 6).Value


                   
        yearchange = ticker_closed - ticker_open

        percentchange = (yearchange / ticker_open)
                    
        ws.Range("K" & summary_table_row).Value = percentchange

        ws.Range("K" & summary_table_row).NumberFormat = "0.00%"

        ws.Range("J" & summary_table_row).Value = yearchange


           If ws.Range("J" & summary_table_row).Value > 0 Then

        
              ws.Range("J" & summary_table_row).Interior.ColorIndex = 4
        
               
                Else


              ws.Range("J" & summary_table_row).Interior.ColorIndex = 3


             End If



' Adds one row to summary table so next value is printed on subsequent row

       summary_table_row = summary_table_row + 1

' Resets volume total and start date opening price before moving on to next stock


       stockvol = 0

       ticker_open = 0

      

        Else


       stockvol = stockvol + Cells(i, 7).Value



 
        End If



      Next i




 Next ws





End Sub
