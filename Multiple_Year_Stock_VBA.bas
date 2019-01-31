Attribute VB_Name = "Module1"
Sub stockvolume()

'Apply Macro to every sheet in workbook
    
    'Set worksheet as variable
    Dim Active As Worksheet
    
    'Run For loop through worksheets
    For Each Active In Worksheets
        Active.Select
        Call RunCode

Next Active

End Sub
Sub RunCode()

'Set variables
Dim ticker As String
Dim volume As Double
Dim x As Double


'Create Headers for new columns

Range("I1").Value = "Ticker"
Range("J1").Value = "Total Stock Volume"

x = 2
Cells(x, 9).Value = Cells(x, 1).Value

'Find last row of each ticker value
    LastRow = Cells(Rows.Count, 1).End(xlUp).Row

'Create For Loop for Stock Volume
For j = 2 To LastRow

    If Cells(j, 1).Value = Cells(x, 9).Value Then
    volume = volume + Cells(j, 7).Value

Else
    
    Cells(x, 10).Value = volume
    volume = Cells(j, 7).Value
    x = x + 1
    Cells(x, 9).Value = Cells(j, 1).Value
    
End If

Next j

    Cells(x, 10).Value = volume
    
'Auto fit colums
    Columns("I:J").EntireColumn.AutoFit


End Sub
