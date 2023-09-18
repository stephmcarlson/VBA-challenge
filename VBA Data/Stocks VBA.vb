Sub stocks():

'pull open from first date and store
'if ticker value changes, pull close date
'calculate difference, calculate % and print
'running sum of volume, print in cell and reset


Dim volume As Currency
    'ChatGPT
Dim ticker As String
Dim first As Double
Dim last As Double
Dim rowcount As Long

Dim i As Integer
Dim j As Long



For i = 2018 To 2020
'ChatGPT
Dim lastRow As Long
With ThisWorkbook.Worksheets(CStr(i))
    lastRow = .Cells(.Rows.Count, "A").End(xlUp).Row
End With

'ChatGPT
ThisWorkbook.Worksheets(CStr(i)).Select

Range("K1").Value = "Ticker"
Range("L1").Value = "Yearly Change"
Range("M1").Value = "Percent Change"
Range("N1").Value = "Total Stock Volume"

ticker = 0
volume = 0
first = 0
last = 0
rowcount = 2

For j = 2 To lastRow

If Cells(j, 1).Value <> Cells((j - 1), 1).Value Then
ticker = Cells(j, 1).Value
first = Cells(j, 3).Value
volume = Cells(j, 7).Value
End If
    
    If Cells(j, 1).Value = Cells((j - 1), 1).Value Then
    volume = volume + Cells(j, 7).Value
    End If
    
        If Cells(j, 1).Value <> Cells((j + 1), 1).Value Then
        last = Cells(j, 6).Value
        
        Cells(rowcount, 11).Value = ticker
        Cells(rowcount, 15).Value = ticker
        Cells(rowcount, 12).Value = last - first
        Cells(rowcount, 13).Value = (Cells(rowcount, 12).Value) / first
        Cells(rowcount, 14).Value = volume

        volume = 0
        rowcount = rowcount + 1
        End If
        
Next j

'ChatGPT

    Dim rng As Range
    Dim cond1 As FormatCondition, cond2 As FormatCondition
    
    ' Define the range to apply conditional formatting to
    Set rng = ThisWorkbook.Worksheets(CStr(i)).Range("L:M")
    
    ' Clear any existing conditional formatting rules
    rng.FormatConditions.Delete
    
    ' Define and apply the first condition (green fill)
    Set cond1 = rng.FormatConditions.Add(Type:=xlCellValue, Operator:=xlGreater, Formula1:="0")
    With cond1
        .Interior.Color = RGB(0, 255, 0)  ' Green
    End With
    
    ' Define and apply the second condition (red fill)
    Set cond2 = rng.FormatConditions.Add(Type:=xlCellValue, Operator:=xlLess, Formula1:="0")
    With cond2
        .Interior.Color = RGB(255, 0, 0)  ' Red
    End With
    

Range("R1").Value = "Ticker"
Range("S1").Value = "Value"
Range("Q2").Value = "Greatest % Increase"
Range("Q3").Value = "Greatest % Decrease"
Range("Q4").Value = "Greatest Total Volume"

'ChatGPT
    Dim inc As Range
    Dim incValue As Double
    Dim incTicker As String

    ' Define the range to search
    Set inc = ThisWorkbook.Worksheets(CStr(i)).Range("M:M")

    incValue = Application.WorksheetFunction.Max(inc)

    Range("S2").Value = incValue
    
    Dim dec As Range
    Dim decValue As Double
    Dim decTicker As String

    ' Define the range to search
    Set dec = ThisWorkbook.Worksheets(CStr(i)).Range("M:M")

    decValue = Application.WorksheetFunction.Min(dec)

    Range("S3").Value = decValue
    
    Dim vol As Range
    Dim volValue As Currency
    Dim volTicker As String

    ' Define the range to search
    Set vol = ThisWorkbook.Worksheets(CStr(i)).Range("N:N")

    volValue = Application.WorksheetFunction.Max(vol)
    

    Range("S4").Value = volValue
    
'ChatGPT
ThisWorkbook.Sheets(CStr(i)).Range("R4").Formula = "=VLOOKUP(S4,N:O,2,0)"
ThisWorkbook.Sheets(CStr(i)).Range("R4").Copy
ThisWorkbook.Sheets(CStr(i)).Range("R4").PasteSpecial Paste:=xlPasteValues
Application.CutCopyMode = False
   
'ChatGPT
incTicker = Application.WorksheetFunction.VLookup(incValue, ThisWorkbook.Worksheets(CStr(i)).Range("M:O"), 3, False)
decTicker = Application.WorksheetFunction.VLookup(decValue, ThisWorkbook.Worksheets(CStr(i)).Range("M:O"), 3, False)


Range("R2").Value = incTicker
Range("R3").Value = decTicker

ThisWorkbook.Sheets(CStr(i)).Range("O:O").ClearContents


Next i


End Sub


