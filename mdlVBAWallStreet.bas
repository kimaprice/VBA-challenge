Attribute VB_Name = "mdlWallStreet"
Sub WallStreet()
    Dim ws As Worksheet
    Dim Ticker, GreatestIncTicker, GreatestDecTicker, GreatestVolTicker As String
    Dim TickerCount As Integer
    Dim TotalVolume, OpenAmt, CloseAmt, PercentChng, AmtChng, GreatestInc, GreatestDec, GreatestVol As Double
    Dim OpenDt, CloseDt As String

    
'Loop through sheets
    For Each ws In Worksheets
        
        'Put headers/labels in correct cells
        ws.Range("I1").Value = "Ticker"
        ws.Range("J1").Value = "Yearly Change"
        ws.Range("K1").Value = "Percent Change"
        ws.Range("K:K").NumberFormat = "0.00%"
        ws.Range("L1").Value = "Total Stock Volume"
        
        ws.Range("O2").Value = "Greatest % Increase"
        ws.Range("O3").Value = "Greatest % Decrease"
        ws.Range("O4").Value = "Greatest Total Volume"
        
        ws.Range("P1").Value = "Ticker"
        ws.Range("Q1").Value = "Value"
        
       'lastrow function to find the count of rows in the sheet
       lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row
       lastrow = lastrow

         Ticker = ws.Cells(1, 1).Value
         TickerCount = 0
         TotalVolume = 0
         PercentChng = 0
         AmtChng = 0
         CloseAmt = 0
         OpenAmt = 0
         CloseDate = ""
         OpenDate = ""
         GreatestInc = 0
         GreatestDec = 0
         GreatestVol = 0
         
         
        'Loop through data/rows
        For i = 2 To lastrow
        
            'check to see if it is a new ticker
            If Ticker <> ws.Cells(i, 1).Value Then
            
                'If new ticker, Output Previous Ticker Total Volume, Yearly Change, Percent Change
                If TickerCount > 0 Then
                    'Output Total Volume form previous Ticker
                    ws.Cells(TickerCount + 1, 12).Value = TotalVolume
                    'Calculate AmtChange for previous Ticker
                    AmtChng = CloseAmt - OpenAmt
                    'Output AmtChng for Previous Ticker
                    ws.Cells(TickerCount + 1, 10).Value = AmtChng
                    'Calculate PercentChng for Previous Ticker
                    PercentChng = AmtChng / OpenAmt
                    'Output PercentChng for Previous Ticker
                    ws.Cells(TickerCount + 1, 11).Value = PercentChng
                    
                    'Compare GreatestInc, if higher assign new amount
                    If GreatestInc < PercentChng Then
                        GreatestIncTicker = Ticker
                        GreatestInc = PercentChng
                    End If
                    
                    'Compare GreatestDec, if lower, assign new amount
                    If GreatestDec > PercentChng Then
                        GreatestDecTicker = Ticker
                        GreatestDec = PercentChng
                    End If
                    
                    'Compare GreatestVol, if higher, assign new amount
                    If GreatestVol < TotalVolume Then
                        GreatestVolTicker = Ticker
                        GreatestVol = TotalVolume
                    End If
                
                End If
                
                'Initiate the Total volume,Open and Close dates and amounts for the ticker by assigning the first value in the sheet
                TotalVolume = ws.Cells(i, 7).Value
                OpenDt = ws.Cells(i, 2).Value
                CloseDt = ws.Cells(i, 2).Value
                OpenAmt = ws.Cells(i, 3).Value
                CloseAmt = ws.Cells(i, 6).Value

                 'Enter the unique ticker values in column I
                TickerCount = TickerCount + 1
                Ticker = ws.Cells(i, 1).Value
                ws.Cells(TickerCount + 1, 9).Value = Ticker
            
            Else
            'Sum the volume amounts as looping through all rows for a ticker and output result to column L
                TotalVolume = TotalVolume + ws.Cells(i, 7).Value
                
            'Check if the Date is earlier than OpenDt
                If OpenDt > ws.Cells(i, 2).Value Then
                    'Set Opent dt and amount to the new earliest found date for the ticker
                    OpenDt = ws.Cells(i, 2).Value
                    OpenAmt = ws.Cells(i, 3).Value
                End If
                
            'Check if the Date is later than the CloseDt
                If CloseDt < ws.Cells(i, 2).Value Then
                    'Set Close date and amount to the new latest found date for the ticker
                    CloseDt = ws.Cells(i, 2).Value
                    CloseAmt = ws.Cells(i, 6).Value
                End If

            End If


        Next i
        
        'write the last ticker's total volume, Yearly Change, Percent Change
        
 
        'Output Total Volume for last Ticker
        ws.Cells(TickerCount + 1, 12).Value = TotalVolume
        'Calculate AmtChng for last ticker
        AmtChng = CloseAmt - OpenAmt
        'OutPut AmgChng for last ticker
        ws.Cells(TickerCount + 1, 10).Value = AmtChng
        'Calculate PercentChng for Last ticker
        PercentChng = AmtChng / OpenAmt
        'Output PercentChng for last ticker
        ws.Cells(TickerCount + 1, 11).Value = PercentChng
        
        'Output Greatest Values and Tickers
        ws.Range("P2").Value = GreatestIncTicker
        ws.Range("P3").Value = GreatestDecTicker
        ws.Range("P4").Value = GreatestVolTicker
        ws.Range("Q2").Value = GreatestInc
        ws.Range("Q2").NumberFormat = "0.00%"
        ws.Range("Q3").Value = GreatestDec
        ws.Range("Q3").NumberFormat = "0.00%"
        ws.Range("Q4").Value = GreatestVol
        
   'Apply Conditional Formatting
        
        'Delete previous conditional formats
        ws.Range("J:K").FormatConditions.Delete
        'Greater than 0 will be Green
        ws.Range("J:K").FormatConditions.Add Type:=xlCellValue, Operator:=xlLess, _
                Formula1:="0"
        ws.Range("J:K").FormatConditions(1).Interior.ColorIndex = 3
        'Less than 0 will be red
        ws.Range("J:K").FormatConditions.Add Type:=xlCellValue, Operator:=xlGreater, _
                Formula1:="0"
        ws.Range("J:K").FormatConditions(2).Interior.ColorIndex = 4
        
        ws.Range("J1:K1").FormatConditions.Delete
        
    Next ws

End Sub
