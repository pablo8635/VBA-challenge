Sub tickandsum():

'What variables are we going to use for this problem?
Dim ticker As String
Dim total As Double
Dim i As Double
Dim j As Integer
Dim open_num As Double
Dim close_num As Double
Dim min As Double
Dim max As Double
Dim max_vol As Double
Dim n_ticker As String


'For many sheets
  ' LOOP THROUGH ALL SHEETS
    ' --------------------------------------------
    For Each ws In Worksheets


' Determine the Last Column Number
'LastColumn = ws.Cells(1, Columns.Count).End(xlToLeft).Column

'Determine last row number
LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row

'Declare headers
ws.Range("I1") = "Ticker"
ws.Range("J1") = "Yearly Change"
ws.Range("K1") = "Percentage change"
ws.Range("L1") = "Total Stock Volume"
ws.Range("O2") = "Greatest % Increase"
ws.Range("O3") = "Greatest % Decrease"
ws.Range("O4") = "Greatest Total Volume"
ws.Range("P1") = "Ticker"
ws.Range("Q1") = "Value"

'Lets declare the variables to start the loop
i = 2
j = 2
open_num = ws.Cells(i, 3).Value

For i = 2 To LastRow
    If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
        ticker = ws.Cells(i, 1).Value
        total = total + ws.Cells(i, 7).Value
        close_num = ws.Cells(i, 6).Value
        ws.Range("I" & j) = ticker
        ws.Range("J" & j) = close_num - open_num
        ws.Range("K" & j) = ((close_num / open_num) - 1)
        ws.Range("K" & j).NumberFormat = "0.00%"
        ws.Range("L" & j) = total
        
        
        total = 0
        open_num = ws.Cells(i + 1, 3).Value
        close_num = 0
        j = j + 1
        
        
    Else
        total = total + Cells(i, 7)
    
    End If

Next i
    
'.Interior.ColorIndex
For i = 2 To LastRow
    If ws.Cells(i, 10).Value > 0 Then
        ws.Cells(i, 10).Interior.ColorIndex = 4
    ElseIf ws.Cells(i, 10).Value < 0 Then
        ws.Cells(i, 10).Interior.ColorIndex = 3
    
    End If
    
Next i

'Max and Min functions
 
max = WorksheetFunction.max(ws.Range("K:K"))
    ws.Range("Q2") = max
    ws.Range("Q2").NumberFormat = "0.00%"
min = WorksheetFunction.min(ws.Range("K:K"))
    ws.Range("Q3") = min
    ws.Range("Q3").NumberFormat = "0.00%"
max_vol = WorksheetFunction.max(ws.Range("L:L"))
    ws.Range("Q4") = max_vol
    


'Change Worksheet

Next ws

End Sub





