Attribute VB_Name = "stocks"
Sub stocks()
    
    'declare data shape vars
    Dim lastRow As Double
    Dim rng As Range
    Dim summaryRow As Integer
        
    'declare ticker vars
    Dim ticker As String
    Dim tickerIndex As Double
    Dim tickerCount As Double
    
    'declare stats vars
    Dim yrOpen As Double
    Dim yrClose As Double
    Dim perChange As Double
    Dim vol As Double
    
    'find the last row, assign it to lastRow & fill in I4
    ActiveSheet.UsedRange 'refresh usedrange'
    lastRow = ActiveSheet.UsedRange.Rows(ActiveSheet.UsedRange.Rows.Count).Row
    Debug.Print (lastRow)
    Range("i4").Value = lastRow
    
    For i = 2 To lastRow
        i = 2
        If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
            ticker = Cells(i, 1).Value
            tickerIndex = i
            
            'calc & set summary table values
            vol = vol + Cells(i, 7).Value
            Range ("j" & summaryRow.Value = ticker)
            Range ("m" & summaryRow.Value = vol)
            yrOpen = Cells(i, 3).Value
            
        End If
    Next i
    
End Sub
