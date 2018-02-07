Sub stocks()
Dim ws As Worksheet
Set ws = ActiveSheet
For Each ws In ThisWorkbook.Worksheets
ws.Activate
'MsgBox (current.Name)

Dim ticker As String
Dim totals As Double
Dim summary As Integer
Dim lastrow As Long

Cells(1, 9).Value = "Ticker"
Cells(1, 10).Value = "Total Stock Volume"

lastrow = Cells(Rows.Count, "A").End(xlUp).Row

summary = 2
totals = 0

For I = 2 To lastrow
    
    If Cells(I + 1, 1).Value <> Cells(I, 1).Value Then
        ticker = Cells(I, 1).Value
        totals = totals + Cells(I, 7).Value

        Range("i" & summary).Value = ticker
        Range("j" & summary).Value = totals

        summary = summary + 1
        
        'reset totals
        totals = 0
        
    ElseIf Cells(I + 1, 1).Value = Cells(I, 1).Value Then
        totals = totals + Cells(I, 7).Value

    End If
Next I
Next ws

End Sub

