Attribute VB_Name = "Module1"
Sub vb_Homework()

    'Variables
    '----------------------------
    'Strings
    Dim ticker As String
    Dim tickerSummary As String
    Dim dateVal As String
    Dim mostIncreaseRowTicker As String
    Dim mostDecreaseRowTicker As String
    Dim mostVolumeTicker As String
    
    
    'Longs

    'Integers
    Dim dayOpen As Integer
    Dim monthOpen As Integer
    Dim dayClose As Integer
    Dim monthClose As Integer
    Dim day As Integer
    Dim month As Integer
    
    'Doubles
    Dim yearOpen As Double
    Dim yearClose As Double
    Dim yearChange As Double
    Dim yearPercent As Double
    Dim currentVolume As Double
    Dim totalVolume As Double
    Dim summaryRow As Double
    Dim lastRow As Double
    Dim lastRowSummary As Double
    Dim yearChangeSummary As Double
    Dim yearOpenSummary As Double
    Dim mostIncrease As Double
    Dim mostIncreaseRow As Double
    Dim mostDecrease As Double
    Dim mostVolume As Double
    Dim mostVolumeRow As Double
    
    Dim sht As Worksheet
    
    'Date
    Dim dateConvert As Date
    
    Dim starting_ws As Worksheet
    Dim ws_num As Integer
    Set starting_ws = ActiveSheet
    
    ws_num = ThisWorkbook.Worksheets.Count

    For ws = 1 To ws_num
        ThisWorkbook.Worksheets(ws).Activate
        
        Range("I1").Value = "Ticker"
        Range("J1").Value = "Yearly Change"
        Range("K1").Value = "Percent Change"
        Range("L1").Value = "Total Stock Volume"
        Range("N2").Value = "Greatest % Increase"
        Range("N3").Value = "Greatest % Decrease"
        Range("N4").Value = "Greatest Total Volume"
        Range("O1").Value = "Ticker"
        Range("P1").Value = "Value"
        
        Range("K:K").NumberFormat = "0.00%"
    
        'Set the start of the summary row.
        summaryRow = 2
        
        'Find the last row in the sheet.
        lastRow = Cells(Rows.Count, 1).End(xlUp).Row
        'MsgBox (lastRow)
        
        'Ticker Symbol
        '----------------------------
        For i = 2 To lastRow
            'Isolate the unique ticker symbols
            If Cells(i + 1, 1) <> Cells(i, 1) Then
                ticker = Cells(i, 1)
                Cells(summaryRow, 9) = ticker
                
                'Advance the summary row to the next row.
                summaryRow = summaryRow + 1
                
            End If
        Next i
        
        'find last row in the summary table
        lastRowSummary = Cells(Rows.Count, 9).End(xlUp).Row
        'MsgBox (lastRowSummary)
        
        summaryRow = 2
        
        'Yearly Change and Percent Change
        '----------------------------
        For i = lastRowSummary To lastRowSummary
            For j = 2 To lastRow
                dayOpen = 1
                monthOpen = 1
                dayClose = 30
                monthClose = 12
                
                dateVal = Cells(j, 2)
                dateConvert = DateSerial(Left(dateVal, 4), Mid(dateVal, 5, 2), Right(dateVal, 2))
                day = DatePart("d", dateConvert)
                month = DatePart("m", dateConvert)
                
                'MsgBox (dateConvert)
                'MsgBox (day)
                'MsgBox (month)
                
                If day = dayOpen And month = monthOpen Then
                    yearOpen = Cells(j, 3)
                    'MsgBox (yearOpen)
                
                ElseIf day = dayClose And month = monthClose Then
                    yearClose = Cells(j, 6)
                    'MsgBox (yearClose)
                    
                    yearChange = yearClose - yearOpen
                    Cells(summaryRow, 10) = yearChange
                    
                    If yearOpen = 0 Then
                        yearPercent = yearChange
                    Else
                        yearPercent = yearChange / yearOpen
                        Cells(summaryRow, 11) = yearPercent
                    End If
                    
                    summaryRow = summaryRow + 1
                    
                End If
                
            Next j
            
        Next i
    
        summaryRow = 2
        
        'Format Colors
        '---------------------------
        For i = 2 To lastRowSummary
            'Conditionally format the cell's color.
            If Cells(i, 10) > 0 Then
                'If the stock gained value.
                Cells(summaryRow, 10).Interior.ColorIndex = 4
            ElseIf Cells(i, 10) < 0 Then
                'If the stock lost value.
                Cells(summaryRow, 10).Interior.ColorIndex = 3
            End If
            summaryRow = summaryRow + 1
        Next i
        
        'Greatest % Increase
        '---------------------------
        For i = 2 To lastRowSummary
            Dim mostIncreaseTemp As Double
            mostIncreaseTemp = Cells(i, 11)
            If mostIncreaseTemp > mostIncrease Then
                mostIncrease = mostIncreaseTemp
            End If
        Next i
        
        For i = 2 To lastRowSummary
            If Cells(i, 11) = mostIncrease Then
                mostIncreaseRow = i
            End If
        Next i
            
        'MsgBox (Cells(mostIncreaseRow, 9))
        mostIncreaseRowTicker = Cells(mostIncreaseRow, 9)
        
        Range("P2").NumberFormat = "0.00%"
        Range("O2").Value = mostIncreaseRowTicker
        Range("P2").Value = mostIncrease
        
        'Greatest % Decrease
        '--------------------------
        For i = 2 To lastRowSummary
            Dim mostDecreaseTemp As Double
            mostDecreaseTemp = Cells(i, 11)
            If mostDecreaseTemp < mostDecrease Then
                mostDecrease = mostDecreaseTemp
            End If
        Next i
        
        For i = 2 To lastRowSummary
            If Cells(i, 11) = mostDecrease Then
                mostDecreaseRow = i
            End If
        Next i
            
        'MsgBox (Cells(mostIncreaseRow, 9))
        mostDecreaseRowTicker = Cells(mostDecreaseRow, 9)
        
        Range("P3").NumberFormat = "0.00%"
        Range("O3").Value = mostDecreaseRowTicker
        Range("P3").Value = mostDecrease
        
            'Total stock volume.
        '--------------------------
        'Loop to isolate the ticker in the summary table.
        For i = 2 To lastRowSummary
            'reset total volume
            totalVolume = 0
            
            'Loop to run down the main table and match the ticker symbols.
            For j = 2 To lastRow
                ticker = Cells(j, 1)
                tickerSummary = Cells(i, 9)
                currentVolume = Cells(j, 7)
                
                'Matching the ticker symbols.
                If ticker = tickerSummary Then
                    totalVolume = totalVolume + currentVolume
                End If
            Next j
            'Add the totalVolume for current ticker symbol to the summary table.
            Cells(i, 12) = totalVolume
        
        'move on to the next ticker symbol in the summary table.
        Next i
        
        
        'Greatest Total Volume
        '-------------------------
        For i = 2 To lastRowSummary
            Dim mostVolumeTemp As Double
            mostVolumeTemp = Cells(i, 12)
            If mostVolumeTemp > mostVolume Then
                mostVolume = mostVolumeTemp
            End If
        Next i
        
        For i = 2 To lastRowSummary
            If Cells(i, 12) = mostVolume Then
                mostVolumeRow = i
            End If
        Next i
            
        'MsgBox (Cells(mostIncreaseRow, 9))
        mostVolumeTicker = Cells(mostVolumeRow, 9)
        
        Range("O4").Value = mostVolumeTicker
        Range("P4").Value = mostVolume

    
    Next ws
    
    Dim sheet As Worksheet
    
    For Each sheet In ThisWorkbook.Worksheets
        sheet.Cells.EntireColumn.AutoFit
    Next sheet
    
End Sub

