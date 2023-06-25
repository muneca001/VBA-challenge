VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Sub StockData()
    Dim wsCount As Integer
    Dim ws As Integer
    Dim columnI As Double
    Dim lastRow As Double
    Dim tickerRow As Double
    Dim currentTicker As String
    Dim nextTicker As String
    Dim closeValue As Double
    Dim openValue As Double
    Dim yearlyChange As Double
    Dim precentChange As Double
    Dim roundedPercent As Double
    Dim volumeCounter As Double
    Dim inTicker As String
    Dim inValue As Double
    Dim deTicker As String
    Dim deValue As Double
    Dim volTicker As String
    Dim volValue As Double
    
    wsCount = ActiveWorkbook.Worksheets.Count
    
    For ws = 1 To wsCount
        Worksheets(ws).Activate
        
        lastRow = Cells(Rows.Count, 1).End(xlUp).Row
        tickerRow = 2
        openValue = Cells(2, 3).Value
    
        'Labels
        Range("I1").Value = "Ticker"
        Range("J1").Value = "Yearly Change"
        Range("K1").Value = "Percent Change"
        Range("L1").Value = "Total Stock Volume"
        Range("P1").Value = "Ticker"
        Range("Q1").Value = "Value"
        Range("O2").Value = "Greatest % increase"
        Range("O3").Value = "Greatest % decrease"
        Range("O4").Value = "Greatest total volume"
    
        'for loop which goes through every row
        For i = 2 To lastRow
            currentTicker = Cells(i, 1).Value
            nextTicker = Cells(i + 1, 1).Value
            'previousTicker = Cells(i - 1, 1).Value
        
            'comparing the values to that of next
            If currentTicker <> nextTicker Then
        
                'pt 1 populates ticker column
                Cells(tickerRow, 9).Value = currentTicker
            
                'pt 2 populates yearly change column w/ color formatting
                closeValue = Cells(i, 6).Value
                yearlyChange = closeValue - openValue
                Cells(tickerRow, 10) = yearlyChange
                    If Cells(tickerRow, 10).Value = Abs(Cells(tickerRow, 10).Value) Then
                        Cells(tickerRow, 10).Interior.ColorIndex = 4
                    Else
                    Cells(tickerRow, 10).Interior.ColorIndex = 3
                    End If
                
                'pt 3 populates percent change
                percentChange = (yearlyChange / openValue)
                Cells(tickerRow, 11) = percentChange
                'greatest volume, % increase/decrease, and percent sign formatting
                If Cells(tickerRow, 11).Value = Cells(2, 11).Value Then
                    inTicker = Cells(2, 9).Value
                    inValue = Cells(2, 11).Value
                    deTicker = Cells(2, 9).Value
                    deValue = Cells(2, 11).Value
                    volTicker = Cells(2, 9).Value
                    volValue = Cells(2, 12).Value
                ElseIf Cells(tickerRow, 11).Value > inValue Then
                    inTicker = Cells(tickerRow, 9).Value
                    inValue = Cells(tickerRow, 11).Value
                ElseIf Cells(tickerRow, 11).Value < deValue Then
                    deTicker = Cells(tickerRow, 9).Value
                    deValue = Cells(tickerRow, 11).Value
                ElseIf Cells(tickerRow, 12).Value > volValue Then
                    volTicker = Cells(tickerRow, 9).Value
                    volValue = Cells(tickerRow, 12).Value
                End If
                Cells(tickerRow, 11).NumberFormat = "0.00%"
                Cells(2, 17).NumberFormat = "0.00%"
                Cells(3, 17).NumberFormat = "0.00%"
                'pt 4 total stock volume for each ticker
                volumeCounter = volumeCounter + Cells(i, 7).Value
                Cells(tickerRow, 12) = volumeCounter
            
                tickerRow = tickerRow + 1
            Else
                volumeCounter = volumeCounter + Cells(i, 7).Value
                Cells(tickerRow, 12) = volumeCounter
                volumeCounter = 0
            End If
            
        Next i
    
        Cells(2, 16) = inTicker
        Cells(2, 17) = inValue
        Cells(3, 16) = deTicker
        Cells(3, 17) = deValue
        Cells(4, 16) = volTicker
        Cells(4, 17) = volValue
    
    Next ws
    
End Sub


