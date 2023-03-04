Attribute VB_Name = "Module1"
Sub stockchange()
    Dim ticker1 As String
    Dim ticker2 As String
    Dim rownum As Double
    Dim volume As Double
    Dim totalvolume As Double
    Dim index As Double
    Dim LastRow As Double
    Dim open_val As Single
    Dim close_val As Single
    Dim change As Single
    Dim per_change As Single
    
    
    'Loop through all worksheets
    
    For Each ws In Worksheets
    
        index = 1
        tickervolume = 0
    
        'Get last row of current worksheet
        LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
    
        'Get opening value in first row of data
        open_val = ws.Cells(2, 3).Value
        
        'Insert column labels
        ws.Range("I1").Value = "Ticker"
        ws.Range("J1").Value = "Yearly Change"
        ws.Range("K1").Value = "Percentage Change"
        ws.Range("L1").Value = "Total Stock Volume"
        
        'adjust column width to display values
        ws.Columns("I").ColumnWidth = "6"
        ws.Columns("J").ColumnWidth = "14"
        ws.Columns("K").ColumnWidth = "17"
        ws.Columns("L").ColumnWidth = "25"
        
        'MsgBox (ws.Name)
    
    
        For rownum = 2 To LastRow
            ticker1 = ws.Cells(rownum, 1).Value
            ticker2 = ws.Cells(rownum + 1, 1).Value
        
            'get volume for current row
            volume = ws.Cells(rownum, 7).Value
        
            'update total volume
            totalvolume = totalvolume + volume
        
        
            If ticker2 <> ticker1 Then
        
            
                'Get closing value ticker1
                close_val = ws.Cells(rownum, 6)
            
                'Calculate change for the year
                change = close_val - open_val
            
                'calculate percentage change in stock price
                per_change = change / open_val
            
                'update index to determine row to print results
                index = index + 1
            
                'print ticker, change, volume
                ws.Cells(index, 9).Value = ticker1
                ws.Cells(index, 10).Value = FormatNumber(change)
                ws.Cells(index, 11).Value = FormatPercent(per_change, 2)
                ws.Cells(index, 12).Value = totalvolume
                
               
                
                'reset total volume for new ticker
                totalvolume = 0
                'Get opening value of next row
                open_val = ws.Cells(rownum + 1, 3).Value
            End If

        Next rownum
        
        '----------- Conditional format yearly change -------------------
        With ws.Range("J2", "J" & index).FormatConditions.Add(xlCellValue, xlGreater, "=0")
            .Interior.Color = vbGreen
        End With
        
        With ws.Range("J2", "J" & index).FormatConditions.Add(xlCellValue, xlLessEqual, "=0")
            .Interior.Color = vbRed
        End With
        
        '---------------  BONUS SECTION ---------------------------------
        'Get & Print greatest increase, decrease and volume
        '---------------------------------------------------------------

        
        'adjust column width to display values
        ws.Columns("O").ColumnWidth = "22"
        ws.Columns("P").ColumnWidth = "6"
        ws.Columns("Q").ColumnWidth = "25"
        
        'Print header/row labels for values
        ws.Range("O2").Value = "Greatest % Increase"
        ws.Range("O3").Value = "Greatest % Decrease"
        ws.Range("O4").Value = "Greatest Total Volume"
        ws.Range("P1").Value = "Ticker"
        ws.Range("Q1").Value = "Value"
        
        Dim maxValue As Double
        Dim minChangeValue As Double
        Dim maxChangeValue As Double
        Dim SearchRange As Range
        Dim ValueCell As Range
        Dim LRow As Long
        
        
        'get last row of new table
        LRow = ws.Cells(Rows.Count, 11).End(xlUp).Row
        
        
        'Find maximum & minimum change
        Set SearchRange = ws.Range("K2:K" & LRow)
        
        
        'find & print maximum change and associated ticker
        maxChangeValue = Application.WorksheetFunction.Max(SearchRange)
        Set ValueCell = ws.Range("K2:K" & LRow).Find(FormatPercent(maxChangeValue, 2))
        ws.Range("P2").Value = ValueCell.Offset(, -2)
        ws.Range("Q2").Value = FormatPercent(maxChangeValue, 2)
        
        'find & print minimum change and associated ticker
        minChangeValue = Application.WorksheetFunction.Min(SearchRange)
        Set ValueCell = ws.Range("K2:K" & LRow).Find(FormatPercent(minChangeValue, 2))
        ws.Range("P3").Value = ValueCell.Offset(, -2)
        ws.Range("Q3").Value = FormatPercent(minChangeValue, 2)
        
        'Find & print maximum volume and associated ticker symbol
        Set SearchRange = ws.Range("L2:L" & LRow)
        maxValue = Application.WorksheetFunction.Max(SearchRange)
        Set ValueCell = ws.Range("L2:L" & LRow).Find(maxValue)
        ws.Range("Q4").Value = maxValue
        ws.Range("P4").Value = ValueCell.Offset(, -3)

        
    Next ws
    
            
End Sub

