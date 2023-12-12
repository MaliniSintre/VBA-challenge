Attribute VB_Name = "Module1"
Sub Worksheet_loop():
    'Declare ws variable to loop through all worksheets
    Dim ws As Worksheet
    
    'Loop through all worksheets
    For Each ws In Worksheets
        ws.Select
        Call Stock
    Next ws
       
End Sub
Sub Stock()

Dim ticker As String
Dim openprice As Double
Dim closeprice As Double
Dim yearlyChange As Double
Dim percentChange As Double
Dim totalVolume As LongLong
Dim SummaryTable As Integer

Dim lastRow As Long
Dim rowcount As Double

'Find the last row of data in column A
lastRow = Cells(Rows.Count, 1).End(xlUp).Row


'Create headers for Summary table
    Cells(1, 9).Value = "Ticker"
    Cells(1, 10).Value = "Yearly Change"
    Cells(1, 11).Value = "Percent Change"
    Cells(1, 12).Value = "Total Stock Volume"

'Format 2 digits - referenced from https://stackoverflow.com/questions/42844778/vba-for-each-cell-in-range-format-as-percentage
Range("K2:K" & lastRow).NumberFormat = "0.00%"

'Copy first ticker to table
Cells(2, 9).Value = Cells(2, 1).Value

'Start summary table from row 2
SummaryRow = 2

'Reset the total stock volume
totalstockvalue = 0

   'Set initial value for opening price
    openprice = Cells(2, 3)

'Loop through each row to calc stock analysis
For i = 2 To lastRow

    'Add row's volume to totalstockvalue
    totalstockvalue = totalstockvalue + Cells(i, 7).Value
    
    'Check for new ticker to copy to table
    If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
    
    'Print totalstockvalue to table
    Range("L" & SummaryRow).Value = totalstockvalue
    
    'Reset totalstockvalue to 0
    totalstockvalue = 0
    
    'Setting the ticker name from column A
    ticker = Cells(i, 1).Value
    
    'Print new ticker to table
    Range("I" & SummaryRow).Value = ticker
      
    'Set closing price
    closeprice = Cells(i, 6).Value
    
    'Calculate yearly change
    yearlyChange = closeprice - openprice
    Range("J" & SummaryRow).Value = yearlyChange
    
    'Calculate percent change
    percentChange = (closeprice - openprice) / openprice
    Range("k" & SummaryRow).Value = percentChange
    
    'Reset open price value for next ticker
    openprice = Cells(i + 1, 3).Value
    
     'Add one to summaryRow
    SummaryRow = SummaryRow + 1
    
    End If

 Next i

'Conditional formatting
    For j = 2 To lastRow
    
    If Cells(j, 10).Value >= 0 Then
    Cells(j, 10).Interior.ColorIndex = 4 'Green for positive
    
    Else: Cells(j, 10).Interior.ColorIndex = 3 'Red for negative
    
    End If

Next j
'-------------------------------------------
'------------------------------------------

'BONUS PART

'Create headers
Range("O2").Value = "Greatest % increase"
Range("O3").Value = "Greatest % decrease"
Range("O4").Value = "Greatest Total volume"

Range("P1").Value = "Ticker"
Range("Q1").Value = "Value"

Dim maximumTicker As String
Dim maximumVol As LongLong
maximumVol = 1

For n = 2 To tablerowcount
    If (Cells(n, 12).Value > maximumVol) Then
        maximumVol = Cells(n, 12).Value
        maximumTicker = Cells(n, 9).Value
    End If
Next n

'Put values to table for greatest total volume
[P4] = maximumTicker
[Q4] = maximumVol

'Format 2 digits
Range("Q2:Q3").NumberFormat = "0.00%"

'Autofit - referenced from https://learn.microsoft.com/en-us/office/vba/api/excel.range.autofit
Columns("A:Q").AutoFit

End Sub
