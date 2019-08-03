VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Sub cycle()
Dim wk As Worksheet

For Each wk In ThisWorkbook.Worksheets
    wk.Activate
    Call volume
Next wk

End Sub

Sub volume():

Dim lastrow As Long
Dim total As Double
Dim count As Long
Dim opener As Double
Dim closer As Double
Dim per As Double
Dim yearly As Double

Cells(1, 9).Value = "Ticker"
Cells(1, 10).Value = "Total Stock Volume"
Cells(1, 11).Value = "Yearly Change"
Cells(1, 12).Value = "Percent Change"

count = 2
total = 0
lastrow = Cells(Rows.count, 1).End(xlUp).Row
opener = Cells(2, 3).Value
closer = 0


   For i = 2 To lastrow
        If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
            closer = Cells(i, 6).Value
            yearly = closer - opener
            If opener <> 0 Then
                per = yearly / opener
            Else
                per = 0
            End If
            
            total = total + Cells(i, 7).Value
            
            Cells(count, 9).Value = Cells(i, 1).Value
            Cells(count, 10).Value = total
            Cells(count, 11).Value = yearly
            Cells(count, 12).Value = per
            
            total = 0
            count = count + 1
            opener = Cells(i + 1, 3).Value
        Else
            total = total + Cells(i, 7).Value
      End If
    Next i

Cells(2, 15).Value = "Greatest % Increase"
Cells(3, 15).Value = "Greatest % Decrease"
Cells(4, 15).Value = "Greatest Total Volume"
Cells(1, 16).Value = "Ticker"
Cells(1, 17).Value = "Value"
Cells(2, 16).Value = Cells(2, 9).Value
Cells(2, 17).Value = Cells(2, 12).Value
Cells(2, 17).NumberFormat = "0.00%"
Cells(3, 16).Value = Cells(2, 9).Value
Cells(3, 17).Value = Cells(2, 12).Value
Cells(3, 17).NumberFormat = "0.00%"
Cells(4, 16).Value = Cells(2, 9).Value
Cells(4, 17).Value = Cells(2, 10).Value


    For j = 2 To count
        If Cells(j, 12).Value > Cells(2, 17).Value Then
            Cells(2, 16).Value = Cells(j, 9).Value
            Cells(2, 17).Value = Cells(j, 12).Value
        ElseIf Cells(j, 12).Value < Cells(3, 17).Value Then
            Cells(3, 16).Value = Cells(j, 9).Value
            Cells(3, 17).Value = Cells(j, 12).Value
        End If
        If Cells(j, 10).Value > Cells(4, 17).Value Then
            Cells(4, 16).Value = Cells(j, 9).Value
            Cells(4, 17).Value = Cells(j, 10).Value
        End If
    Next j

Range("A:Z").Columns.AutoFit

    For k = 2 To count
        If Cells(k, 11).Value > 0 Then
            Cells(k, 11).Interior.ColorIndex = 4
        ElseIf Cells(k, 11).Value < 0 Then
            Cells(k, 11).Interior.ColorIndex = 3
        End If
    Next k
    
End Sub

