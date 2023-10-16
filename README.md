# VBA_Challenge

Credit:
All code provided is original.
All written code is independant.

Research & resources consistent of content within the course. Outside resources used were LearningAssistant & ChatGPT
Outside resources strictly used to assist me troubleshoot my code & not to plagerise.

About Files:
"vba_challenge SCRIPT": My assignment. NOTE: This document is saved as a word document which can be copied & pasted into the VBA page on either the alphabetical_testing or Multiple_year_stock_data workbooks. Please find a copy of the script underneath this READ.ME as an extra copy if needed.

"vba_challenge SCREENSHOTS": consists of screenshots of each worksheet within the "Multiple_year_stock_data" workbook provided before & after the code was run.
Before & After photos are identified via the header provided, the year can be identified in the bottom left of each screenshot (the worksheet [year] is highlighted).

VBA_Challenge Script Below!

Sub VBA_Challenge()

    'All variables among nested For "ws" loop & separate For "i" loops
    Dim ws As Worksheet
    Dim i As Long
    Dim LastRow As Long
    Dim TLastRow As Long
    Dim opening As Double
    Dim closing As Double
    Dim Ticker As String
    Dim YDif As Double
    Dim PrChng As Double
    Dim col As Range
    Dim v As Long
    Dim sum As Double
    Dim vTicker As String
    Dim GI As Double
    Dim GD As Double
    Dim GTV As Double
    
    
    'Setup headers for each column
        For Each ws In ThisWorkbook.Worksheets
            LastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
            ws.Cells(1, 9).Value = "Ticker"
            ws.Cells(1, 10).Value = "Yearly Change ($)"
            ws.Cells(1, 11).Value = "Percent Change"
            ws.Cells(1, 12).Value = "Total Stock Volume"
        
        'Variables used through separate For "i" loops, but consistent in For "ws" loop
            y = 2
            t = 2
            v = 2
            g = 1
            sum = 0
              
        'Create For Loop to fill TICKER column
            For i = 2 To LastRow
                
                'Utilise If function to compare adjacent i,1 values in order to draw out unique values
                    If ws.Cells(i, 1).Value <> ws.Cells(i - 1, 1).Value Then
                    ws.Cells(t, 9).Value = ws.Cells(i, 1).Value
                    t = t + 1
                    
                    End If
                    
                'Create variable reference to assist drawing correct values for YEARLY CHANGE & PERCENT CHANGE columns
                    Ticker = ws.Cells(y, 9).Value
            
                'Create fixed reference to co-assist drawing correct values for YEARLY CHANGE & PERCENT CHANGE columns
                OpenDateCode = "0102"
                CloseDateCode = "1231"
                
                'Utilise references to draw appropriate values for opening for ticker
                If Right(ws.Cells(i, 2).Value, 4) = OpenDateCode And ws.Cells(i, 1).Value = Ticker Then
                    opening = ws.Cells(i, 3).Value
                End If
            
        
              'Do the same for closing values. Once both variables have values, manipulate the data so it fits the YEARLY CHANGE & PERCENT CHANGE columns, then adjust the variable reference point so it looks at the next Ticker
                 If Right(ws.Cells(i, 2).Value, 4) = CloseDateCode And ws.Cells(i, 1).Value = Ticker Then
                closing = ws.Cells(i, 6).Value
                PrChng = closing / opening
                YDif = closing - opening
                ws.Cells(y, 10).Value = YDif
                ws.Cells(y, 11).Value = (PrChng - 1) * 100 & "%"
                
                y = y + 1
                 End If
            'End For Loop
            Next i
            
            'Create new For Loop to fill TOTAL STOCK VOLUME column
             For i = 2 To LastRow + 1
                 'Make reference point in TICKER column
                    vTicker = ws.Cells(v, 9).Value
                    
                'Use the reference point to sum all the TOTAL STOCK VOLUME data for the particular ticker, pull the total to the TOTAL STOCK VOLUME column, reset the "sum" variable, add the current i value & adjust the variable reference point so it can begin summing the values with the next ticker
                If ws.Cells(i, 1).Value = vTicker Then
                    sum = sum + ws.Cells(i, 7).Value
                ElseIf ws.Cells(i, 1).Value <> vTicker Then
                ws.Cells(v, 12).Value = sum
                        sum = 0 + ws.Cells(i, 7).Value
                        v = v + 1
                End If
            
           'End For Loop
           Next i
        
           'Create new For Loop for CONDITIONAL FORMATTING
                For i = 2 To LastRow
                
                'Utilise conditional formatting of cell colour in YEARLY CHANGE & PERCENT CHANGE columns. Have green for positive values & red for negative values
                    If ws.Cells(i, 9).Value = "" Then
                        ws.Cells(i, 10).Interior.ColorIndex = 0
                        ws.Cells(i, 11).Interior.ColorIndex = 0
                    ElseIf ws.Cells(i, 10).Value = 0 Then
                        ws.Cells(i, 10).Interior.Color = vbYellow
                        ws.Cells(i, 11).Interior.Color = vbYellow
                    ElseIf ws.Cells(i, 10).Value > 0 Then
                         ws.Cells(i, 10).Interior.Color = vbGreen
                         ws.Cells(i, 11).Interior.Color = vbGreen
                    Else: ws.Cells(i, 10).Interior.Color = vbRed
                          ws.Cells(i, 11).Interior.Color = vbRed
                    End If
        'End For Loop
        Next i
   
        
           'Create table for ADDITIONAL STATS
                ws.Cells(2, 15).Value = "Greatest % Increase"
                ws.Cells(3, 15).Value = "Greatest % Decrease"
                ws.Cells(4, 15).Value = "Greatest Total Volume"
                ws.Cells(1, 16).Value = "Ticker"
                ws.Cells(1, 17).Value = "Value"
                
                'Create For Loop to find corresponding TICKERS to values & paste them accordingly in ADDITIONAL STATS table
                TLastRow = ws.Cells(ws.Rows.Count, 11).End(xlUp).Row
                For i = 2 To TLastRow
                
            'Assign desired values to respectable variables
                GI = Application.WorksheetFunction.max(ws.Range("K2:K" & TLastRow))
                GD = Application.WorksheetFunction.min(ws.Range("K2:K" & TLastRow))
                GTV = Application.WorksheetFunction.max(ws.Range("L2:L" & TLastRow))
                
            
                If ws.Cells(i, 11).Value = GI Then
                    ws.Cells(2, 16).Value = ws.Cells(i, 9).Value
                    ws.Cells(2, 17).Value = ws.Cells(i, 11).Value * 100 & "%"
                End If
                
                If ws.Cells(i, 11).Value = GD Then
                    ws.Cells(3, 16).Value = ws.Cells(i, 9).Value
                    ws.Cells(3, 17).Value = ws.Cells(i, 11).Value * 100 & "%"
                End If
                
                If ws.Cells(i, 12).Value = GTV Then
                    ws.Cells(4, 16).Value = ws.Cells(i, 9).Value
                    ws.Cells(4, 17).Value = ws.Cells(i, 12).Value
                End If
                   
              Next i
        Next ws
End Sub

