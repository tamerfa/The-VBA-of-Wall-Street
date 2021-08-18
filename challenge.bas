Attribute VB_Name = "module2"
Sub challenge():
    Dim maxPercent As Double
    Dim maxTicker As String
    Dim maxVol As Double
    Dim maxVolTicker As String
    Dim minPercent As Double
    Dim minTicker As String
    
    For Each ws In Worksheets
        'Determine last row of summary table
        sLastrow = ws.Cells(Rows.Count, 9).End(xlUp).Row
        
        maxPercent = Application.WorksheetFunction.Max(ws.Range("K2:K" & sLastrow))
        maxVol = Application.WorksheetFunction.Max(ws.Range("L2:L" & sLastrow))
        minPercent = Application.WorksheetFunction.Min(ws.Range("K2:K" & sLastrow))
        
        For i = 2 To sLastrow
            If ws.Range("K" & i).Value = maxPercent Then
                maxTicker = ws.Range("I" & i).Value
            ElseIf ws.Range("K" & i).Value = minPercent Then
                minTicker = ws.Range("I" & i).Value
            End If
            
            If ws.Range("L" & i).Value = maxVol Then
                maxVolTicker = ws.Range("I" & i).Value
            End If
        Next i
        
        'Fill in the headers
        ws.Range("P1").Value = "Ticker"
        ws.Range("Q1").Value = "Value"
        ws.Range("O2").Value = "Greatest % Increase"
        ws.Range("O3").Value = "Greatest % Decrease"
        ws.Range("O4").Value = "Greatest Total Volume"
        
        'Fill in the data in the table
        ws.Range("P2").Value = maxTicker
        ws.Range("Q2").Value = maxPercent
        ws.Range("P3").Value = minTicker
        ws.Range("Q3").Value = minPercent
        ws.Range("P4").Value = maxVolTicker
        ws.Range("Q4").Value = maxVol
        
        'Adjust style of percent cells
        ws.Range("Q2").NumberFormat = "0.00%"
        ws.Range("Q3").NumberFormat = "0.00%"
            
        
        'Adjust table columns to autofit
        ws.Columns("O:Q").AutoFit
    Next ws
            
        
        
End Sub
