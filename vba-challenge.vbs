Sub checkStock():
    
    ' Set variables and default numbers.
    Dim tickerName As String
    Dim totalVolume As LongLong
    totalVolume = 0
    Dim rowCounter As Integer
    rowCounter = 2
    
    ' Create headers for Ticker, Year Change, Percentage Change, Volume
    Range("I1").Value = "Ticker"
    Range("J1").Value = "Year Change"
    Range("K1").Value = "Percentage Change"
    Range("L1").Value = "Total Stock Volume"
    
    ' Go down the rows.
    For i = 1 To Range("A1").End(xlDown).Row
    
        If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
            tickerName = Cells(i, 1).Value
            
        
            totalVolume = totalVolume + Cells(i + 1, 7).Value
            
        End If
        
    Next i
    
    ' Automatically change width, https://stackoverflow.com/questions/63862228/how-to-autofit-column-width-with-vba
    Range("I:L").EntireRow.AutoFit
End Sub
