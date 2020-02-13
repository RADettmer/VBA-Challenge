Attribute VB_Name = "Module1"
Sub VBAHomework()

'--------------------------------------------
'Loop for Worksheets
'--------------------------------------------

For Each ws In Worksheets 'loop for worksheets
Dim worksheetname As String
worksheetname = ws.Name

'--------------------------------------------
'Set up Variables
'--------------------------------------------

Dim vol_total As Double 'vol of ticker
Dim opened As Double 'first open price
Dim closed As Double 'closed price
Dim lastrow As Long 'rows of data in column A

Dim w As Double 'establish w loop counter as a variable
Dim x As Double 'establish x loop counter as a variable
Dim y As Double 'establish y loop counter as a variable

Dim greatest_inc As Double 'greatest percent increase
Dim greatest_dec As Double 'greatest percent decrease
Dim greatest_vol As Double 'greatest total volume

greatest_inc = 0
greatest_dec = 0
greatest_vol = 0

'set end of data
lastrow = ws.Cells(Rows.Count, "A").End(xlUp).Row

vol_total = 0 'vol of trading for ticker
opened = 0 'first open price
closed = 0 'closed price

w = 0 'establish w loop counter as a variable
x = 0 'establish x loop counter as a variable

'----------------------------------------------
'Collect ticker Data and Eliminate Duplicates
'----------------------------------------------

With ActiveSheet 'source code from Stack Overflow "Make a new column without duplicates"
    ws.Range("A1", ws.Range("A1").End(xlDown)).Copy Destination:=ws.Range("I1")
    ws.Range("I1", ws.Range("I1").End(xlDown)).RemoveDuplicates Columns:=1, Header:=xlNo
End With

'----------------------------------------------
'Insert Headers - This will help with formating
'----------------------------------------------

ws.Range("I1").Value = "Ticker"
ws.Range("J1").Value = "Yearly Change"
ws.Range("K1").Value = "Percent Change"
ws.Range("L1").Value = "Total Stock Volume"
ws.Range("P1").Value = "Ticker"
ws.Range("O2").Value = "Greatest % Increase"
ws.Range("O3").Value = "Greatest % Decrease"
ws.Range("O4").Value = "Greatest Total Volume"
ws.Range("Q1").Value = "Value"

'--------------------------------------------------------------------
'Loop to generate yearly change and percent change with conditional formating
'--------------------------------------------------------------------

For w = 2 To lastrow

    If ws.Cells(w - 1, "A").Value <> ws.Cells(w, "A").Value Then opened = ws.Cells(w, "C").Value 'look behind
    
    If ws.Cells(w + 1, "A").Value <> ws.Cells(w, "A").Value Then closed = ws.Cells(w, "F").Value 'look ahead
     
    If ws.Cells(w + 1, "A").Value <> ws.Cells(w, "A").Value Then ws.Cells(w, 10).Value = (closed - opened)
    If ws.Cells(w + 1, "A").Value <> ws.Cells(w, "A").Value Then ws.Cells(w, 11).Value = (closed - opened) / opened
    
    'Update Yearly Change column with color code
    If (closed - opened) < 0 Then
    ws.Cells(w, 10).Interior.ColorIndex = 3 'red for negative change
    Else: ws.Cells(w, 10).Interior.ColorIndex = 4 'green for positive change
    End If
     
    If ws.Cells(w - 1, "A").Value <> ws.Cells(w, "A").Value Then vol_total = 0 'look behind

    vol_total = ws.Cells(w, "G").Value + vol_total 'collects total vol
    
    If ws.Cells(w + 1, "A").Value <> ws.Cells(w, "A").Value Then ws.Cells(w, 12).Value = vol_total
           
Next w

'---------------------------------------------------------
'Formating columns and eliminate empty cells
'---------------------------------------------------------

ws.Columns("J").NumberFormat = "0.00" 'set yearly change column to show decimals
ws.Columns("K").NumberFormat = "0.00%" 'set percent change column to percent
ws.Range("Q2:Q3").NumberFormat = "0.00%" 'set output range as percent
ws.Range("I1:L1").Columns.AutoFit 'adjust column width to fit

'PURPOSE: Deletes single cells that are blank located inside a designated range
'source code from Stack Overflow "How to delete empty cells in Excel using VBA"
Dim RNG As Range
On Error Resume Next
Set RNG = Intersect(ws.UsedRange, ws.Range("J:L"))
RNG.SpecialCells(xlCellTypeBlanks).Delete shift:=xlUp

'-----------------------------------------------------------
'Generate greatest solutions
'-----------------------------------------------------------

greatest_inc = WorksheetFunction.Max(ws.Columns("K")) 'returns greatest % increase
greatest_dec = WorksheetFunction.Min(ws.Columns("K")) 'returns greatest % decrease
greatest_vol = WorksheetFunction.Max(ws.Columns("L")) 'returns greatest total volume

Dim tickergi As String
Dim tickergd As String
Dim tickergv As String
tickergi = ""
tickergd = ""
tickergv = ""

For x = 2 To lastrow 'run thru data set to determine ticker of greatest variables
    If greatest_inc = ws.Cells(x, "K").Value Then tickergi = ws.Cells(x, "I").Value
    If greatest_dec = ws.Cells(x, "K").Value Then tickergd = ws.Cells(x, "I").Value
    If greatest_vol = ws.Cells(x, "L").Value Then tickergv = ws.Cells(x, "I").Value
Next

'Output ticker and amounts of greatest variables
ws.Range("P2") = tickergi
ws.Range("Q2") = greatest_inc
ws.Range("P3") = tickergd
ws.Range("Q3") = greatest_dec
ws.Range("P4") = tickergv
ws.Range("Q4") = greatest_vol
 
'Formating remaining columns
ws.Range("O1:Q1").Columns.AutoFit 'adjust column width to fit

Next ws 'Go to next worksheet

End Sub


