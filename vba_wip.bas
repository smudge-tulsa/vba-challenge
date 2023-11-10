Attribute VB_Name = "Module1"
Attribute VB_Name = "Module1"
Sub stocks()

    Dim ticker As String
    Dim change As Double
    Dim vol As Double
    vol = 0
    Dim summary_table_row As Integer
    summary_table_row = 2
    Dim rowCount As Double
    Dim yrOpen As Double
    Dim yrClose As Double
    Dim yrChange As Double
    Dim perChange As Double
    Dim tickerIndex As Double
    tickerIndex = 0

' Find the number of rows and write to rowCount and basic stats in I
rowCount = Cells(Rows.Count, 1).End(xlUp).Row
Cells(4, 9).Value = rowCount

 For i = 2 To rowCount

    ' Check if we are still within the same ticker, if it is not...
    i = 2
    If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then

      ' Set the ticker
    ticker = Cells(i, 1).Value
    tickerIndex = i

      ' Add to Trade Vol
    vol = vol + Cells(i, 7).Value

      ' Print the ticker in the Summary Table
    Range("j" & summary_table_row).Value = ticker

     ' Print trade volume to the Summary Table
    Range("m" & summary_table_row).Value = vol

    'Print year open
    yrOpen = Cells(i, 3).Value
    Debug.Print tickerIndex, ticker, yrOpen
         
     ' Add one to the summary table row
    summary_table_row = summary_table_row + 1

    ' Reset the ticker volume
      vol = 0

    ' If the cell immediately following a row is the same ticker...
    Else

      ' Add to the ticker volume
      vol = vol + Cells(i, 7).Value

    End If

  Next i

'Next key
End Sub
