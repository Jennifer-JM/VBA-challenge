Attribute VB_Name = "Module4"
Sub Multiple_yr_stock_data():

    ' Loop through all sheets in the workbook
    For Each ws In ThisWorkbook.Sheets
    
        ' Set variables
        Dim i As Long
        Dim Total As Double
        Dim Q_change As Double
        Dim Open_Price As Double, Close_Price As Double
        Dim Start As Long
        Dim lastRowA As Long
        Dim lastRowI As Long
        Dim Greatest_Vol As Double
        Dim Greatest_Percent_Change As Double
        Dim Greatest_Incr As Double
        Dim Greatest_Decr As Double
        Dim Greatest_Vol_Ticker As String
        Dim Greatest_Incr_Ticker As String
        Dim Greatest_Decr_Ticker As String
        Dim Current_Total_Vol As Double
        Dim Current_Percentage_Change As Double
        Dim Current_Ticker As String
        Dim New_Table_Row As Long
        Dim Ticker_Name As String
        Dim Percent_Change As Double
        
        ' Add new table column names
        ws.Cells(1, 9).Value = "Ticker"
        ws.Cells(1, 10).Value = "Quarterly Change"
        ws.Cells(1, 11).Value = "Percent Change"
        ws.Cells(1, 12).Value = "Total Stock Volume"
        ' Second new table column names
        ws.Cells(1, 16).Value = "Ticker"
        ws.Cells(1, 17).Value = "Value"
        ws.Cells(2, 15).Value = "Greatest % Increase"
        ws.Cells(3, 15).Value = "Greatest % Decrease"
        ws.Cells(4, 15).Value = "Greatest Total Volume"
        
        Start = 2
        
        New_Table_Row = 2
        
        Total = 0
        
        ' Find the last row with data in column A
        lastRowA = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    
            ' Loop through all rows (Tickers)
            For i = 2 To lastRowA
                ' Get ticker name
                 Ticker_Name = ws.Cells(i, 1).Value
                ' Check if we are still within the same ticker, if it is not then:
                If ws.Cells(i + 1, 1).Value <> Ticker_Name Then
                
                ' Add data to the new table (for each ticker)
                ws.Cells(New_Table_Row, 9).Value = Ticker_Name
                
                ' Calculate quarterly change
                    ' Get Close Price
                Close_Price = ws.Cells(i, 6).Value
                    ' Get Open Price
                Open_Price = ws.Cells(Start, 3).Value
                Q_change = Close_Price - Open_Price
                
                ' Add data to the new table (for each ticker)
                ws.Cells(New_Table_Row, 10).Value = Q_change
                
                    ' Format colors for Q_change
                    If Q_change > 0 Then
                    ws.Cells(New_Table_Row, 10).Interior.Color = RGB(0, 255, 0)  ' Green
                    ElseIf Q_change < 0 Then
                    ws.Cells(New_Table_Row, 10).Interior.Color = RGB(255, 0, 0)  ' Red
                    End If
                    
                    ' Calculate the percent change
                    If Open_Price <> 0 Then
                    Percent_Change = (Q_change / Open_Price)
                    ' Percentage format
                    ws.Cells(New_Table_Row, 11).Value = FormatPercent(Percent_Change, 2)
                    Else
                    ws.Cells(New_Table_Row, 11).Value = FormatPercent(0, 2)
                    End If
                    
                    ' Reset for next ticker
                    Total = 0
                    Start = i + 1
                    ' Increment the new table row
                    New_Table_Row = New_Table_Row + 1
                    
                Else
                    ' Total Stock Volume
                    Total = Total + ws.Cells(i, 7).Value
                End If
                
                ' Add data to the new table (for each ticker)
                ws.Cells(New_Table_Row, 12).Value = Total
                    
            Next i
            
            ' Find the last row with data in column I
            lastRowI = ws.Cells(ws.Rows.Count, "I").End(xlUp).Row
            
            ' Summary metrics
            Greatest_Vol = ws.Cells(Start, 12).Value
            Greatest_Incr = ws.Cells(Start, 11).Value
            Greatest_Decr = ws.Cells(Start, 11).Value
            Greatest_Vol_Ticker = ws.Cells(Start, 9).Value
            Greatest_Incr_Ticker = ws.Cells(Start, 9).Value
            Greatest_Decr_Ticker = ws.Cells(Start, 9).Value
            
            ' Loop for summary info
            For i = 2 To lastRowI
                Current_Total_Vol = ws.Cells(i, 12).Value
                Current_Percentage_Change = ws.Cells(i, 11).Value
                Current_Ticker = ws.Cells(i, 9).Value
                    
                ' Greatest volume calculation
                If Current_Total_Vol > Greatest_Vol Then
                    Greatest_Vol = Current_Total_Vol
                    Greatest_Vol_Ticker = Current_Ticker
                End If
                    
                ' Greatest percentage increase calculation
                If Current_Percentage_Change > Greatest_Incr Then
                    Greatest_Incr = Current_Percentage_Change
                    Greatest_Incr_Ticker = Current_Ticker
                End If
                    
                ' Greatest percentage decrease calculation
                If Current_Percentage_Change < Greatest_Decr Then
                    Greatest_Decr = Current_Percentage_Change
                    Greatest_Decr_Ticker = Current_Ticker
                End If
                    
            Next i
                
            ' Add data to summary metrics
            ws.Cells(2, 16).Value = Greatest_Incr_Ticker
            ws.Cells(2, 17).Value = Format(Greatest_Incr, "Percent")
            ws.Cells(3, 16).Value = Greatest_Decr_Ticker
            ws.Cells(3, 17).Value = Format(Greatest_Decr, "Percent")
            ws.Cells(4, 16).Value = Greatest_Vol_Ticker
            ws.Cells(4, 17).Value = Format(Greatest_Vol, "Scientific")
            
        ' Automatically adjust column width
        ws.Columns("A:Z").AutoFit
                
        Next ws
    End Sub
