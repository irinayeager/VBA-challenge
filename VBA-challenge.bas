Attribute VB_Name = "Module1"
Sub Stocks()


' Variables
Dim ticker
Dim yearlyChange
Dim percentChange
Dim totalStockVolume As Double
Dim openPrice
Dim closePrice
Dim summaryTableRow
Dim yearStart
Dim wsCount As Integer

' hard challange variables
Dim greatestIncreaseTicker
Dim greatestDecreaseTicker
Dim greatestVolumeTicker
Dim greatestIncreaseValue
Dim greatestDecreaseValue
Dim greatestVolumeValue

greatestIncreaseTicker = "test"
greatestDecreaseTicker = "test"
greatestVolumeTicker = "test"
greatestIncreaseValue = 0
greatestDecreaseValue = 0
greatestVolumeValue = 0

'worksheet iterate
wsCount = ActiveWorkbook.Worksheets.Count

For ws = 1 To wsCount
    ' format worksheet by adding colomns
    Worksheets(ws).Range("I1") = "Ticker"
    Worksheets(ws).Range("J1") = "Yearly Change"
    Worksheets(ws).Range("K1") = "Percent Change"
    Worksheets(ws).Range("L1") = "Total Stock Volume"
    
    'Hard Solution formatting
     Worksheets(ws).Range("P1") = "Ticker"
     Worksheets(ws).Range("Q1") = "Value"
     Worksheets(ws).Range("O2") = "Greatest % Increase"
     Worksheets(ws).Range("O3") = "Greatest % Decrease"
     Worksheets(ws).Range("O4") = "Greatest Total Volume"
     Worksheets(ws).Range("Q2:Q3").NumberFormat = "0.00%"

    summaryTableRow = 2
    ' begin loop
    For I = 2 To Worksheets(ws).Cells(Rows.Count, 1).End(xlUp).Row
        ' set ticker and begin stock volume increment
        ticker = Worksheets(ws).Cells(I, 1)
        totalStockVolume = totalStockVolume + Cells(I, 7)
        
        'set opening price
        If openPrice = "" Then
            openPrice = Worksheets(ws).Cells(I, 3)
        End If
        
        'iterate
        If ticker <> Worksheets(ws).Cells((I + 1), 1) Then
        
        'set close price
        closePrice = Worksheets(ws).Cells(I, 6)
        'calc yearly change
        yearlyChange = openPrice - closePrice
        
        ' ticker output to worksheet
        Worksheets(ws).Range("I" & summaryTableRow).Value = ticker
        
        ' yearly change output to worksheet with formatting
        Worksheets(ws).Range("J" & summaryTableRow).Value = yearlyChange
        If yearlyChange > 0 Then
        Worksheets(ws).Range("J" & summaryTableRow).Interior.ColorIndex = 4 'Green
        Else
        Worksheets(ws).Range("J" & summaryTableRow).Interior.ColorIndex = 3 'Red
        End If
        
        ' percent changed output to worksheet with formatting
        If startPrice <> closePrice Then
            percentChange = yearlyChange / closePrice
        Else
            percentChange = 0
        End If
        Worksheets(ws).Range("K" & summaryTableRow).Value = percentChange
        Worksheets(ws).Range("K" & summaryTableRow).NumberFormat = "0.00%"
        
        ' stock volume output to worksheet
        Worksheets(ws).Range("L" & summaryTableRow).Value = totalStockVolume
        
        
        'hard check greatest volume
        If totalStockVolume > greatestVolumeValue Then
        greatestVolumeValue = totalStockVolume
        greatestVolumeTicker = ticker
        End If
        
        ' reset variables and increment
        summaryTableRow = summaryTableRow + 1
        totalStockVolume = 0
        
        End If
        
        'hard greatest increase/decrease %
        If percentChange > greatestIncreaseValue Then
        greatestIncreaseValue = percentChange
        greatestIncreaseTicker = ticker
        ElseIf percentChange < greatestDecreaseValue Then
        greatestDecreaseValue = percentChange
        greatestDecreaseTicker = ticker
        End If
    
    Next I
    
    'output hard challange
    Worksheets(ws).Range("P2") = greatestIncreaseTicker
    Worksheets(ws).Range("Q2") = greatestIncreaseValue
    Worksheets(ws).Range("P3") = greatestDecreaseTicker
    Worksheets(ws).Range("Q3") = greatestDecreaseValue
    Worksheets(ws).Range("P4") = greatestVolumeTicker
    Worksheets(ws).Range("Q4") = greatestVolumeValue

    'reset hard variables
    greatestIncreaseTicker = "test"
    greatestDecreaseTicker = "test"
    greatestVolumeTicker = "test"
    greatestIncreaseValue = 0
    greatestDecreaseValue = 0
    greatestVolumeValue = 0
    
Next ws

            
        
End Sub
