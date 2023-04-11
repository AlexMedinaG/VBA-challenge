# VBA-challenge

# loop for going through all the Worksheets through activating each worksheet and running the stocksAnalysis subprogram
Sub loopWorksheets()
    Dim ws As Worksheet
    
    For Each ws In ThisWorkbook.Worksheets
        ws.Activate
        stocksAnalysis
    Next ws
End Sub

# stocksAnalysis subprogram
Sub stocksAnalysis()
    
# declaring variables and start values    
    Dim tickerSymbol As String
    Dim summaryTable As Integer
    summaryTable = 2
    
    Dim openingPrice As Double
    Dim closingPrice As Double
    Dim totalStockVolume As Double

# giving names to table headers    
    Range("I1") = "Ticker"
    Range("J1") = "Yearly Change"
    Range("K1") = "Percent Change"
    Range("L1") = "Total Stock Volume"
    totalStockVolume = 0

# select cell to start    
    Range("A2").Select
    openingPrice = Range("C2").Value

# loop until the the last cell on the range is empty
    Do Until IsEmpty(ActiveCell)

# the loop looks for when the active cell is <> to the next active cell and if so, takes the active cell values for the ticker, and the corresponding values for prices using a horizontal offset
        If ActiveCell <> ActiveCell.Offset(1, 0) Then
            tickerSymbol = ActiveCell.Value
            closingPrice = ActiveCell.Offset(0, 5)
            Range("I" & summaryTable).Value = tickerSymbol
            Range("J" & summaryTable).Value = closingPrice - openingPrice
            
            totalStockVolume = totalStockVolume + ActiveCell.Offset(0, 6)
            Range("L" & summaryTable).Value = totalStockVolume

 # function to color the cells according to change          
            If Range("J" & summaryTable).Value > 0 Then
                Range("J" & summaryTable).Interior.Color = vbGreen
                Else
                Range("J" & summaryTable).Interior.Color = vbRed
            End If
            
            Range("K" & summaryTable).Value = FormatPercent((closingPrice - openingPrice) / openingPrice)
            summaryTable = summaryTable + 1
            openingPrice = ActiveCell.Offset(1, 2)
 # returns Stock Volume to 0 for next sum                
            totalStockVolume = 0
        Else

 # continues to sum the Stock Volume because the active cell and the next cell have the same value
        totalStockVolume = totalStockVolume + ActiveCell.Offset(0, 6)
        End If

# goes to next cell and closes the loop    
    ActiveCell.Offset(1, 0).Select
    Loop
    
# declaring variables and start values for next analysis        
    Dim greatestTable As Integer
    greatestTable = 2
    Dim greatestIncreaseValue As Double
    Dim greatestIncreaseTicker As String
    Dim greatestDecreaseValue As Double
    Dim greatestDecreaseTicker As String
    Dim greatestVolumeValue As Double
    Dim greatestVolumeTicker As String
    greatestIncreaseValue = 0
    greatestDecreaseValue = 0
    greatestVolumeValue = 0

# giving names to new table headers    
    Range("P1") = "Ticker"
    Range("Q1") = "Value"
    Range("O2") = "Greatest % Increase"
    Range("O3") = "Greatest % Decrease"
    Range("O4") = "Greatest Total Volume"

# select cell to start     
    Range("I2").Select

# the loop looks for all the cells of the new list and overrights each value to the "greatest" in each category accordingly 
    Do Until IsEmpty(ActiveCell)
        If ActiveCell.Offset(0, 2) > greatestIncreaseValue Then
            greatestIncreaseValue = ActiveCell.Offset(0, 2)
            greatestIncreaseTicker = ActiveCell.Value
        End If

        If ActiveCell.Offset(0, 2) < greatestDecreaseValue Then
            greatestDecreaseValue = ActiveCell.Offset(0, 2)
            greatestDecreaseTicker = ActiveCell.Value
        End If

        If ActiveCell.Offset(0, 3) > greatestVolumeValue Then
            greatestVolumeValue = ActiveCell.Offset(0, 3)
            greatestVolumeTicker = ActiveCell.Value
        End If

# after running through the loop, if prints the stored values from the loop before
    Range("P2").Value = greatestIncreaseTicker
    Range("Q2").Value = FormatPercent(greatestIncreaseValue)
    
    Range("P3").Value = greatestDecreaseTicker
    Range("Q3").Value = FormatPercent(greatestDecreaseValue)
    
    Range("P4").Value = greatestVolumeTicker
    Range("Q4").Value = greatestVolumeValue
    
    ActiveCell.Offset(1, 0).Select
    Loop

# returns the selection to the top of the table    
    Range("A1").Select
End Sub



