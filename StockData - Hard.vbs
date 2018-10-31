Sub stock()

For Each ws In Worksheets
Dim i As Long
Dim ticker As String
Dim totalvolume As Double
Dim yearly As Double
Dim percent As Double
Dim monthopen As Double
Dim monthclose As Double
Dim Minchange As Double
Dim Maxchange As Double
Dim Maxvol As Double
totalvolume = 0

Dim Summary_Table_Row As Integer
Summary_Table_Row = 2

Dim LastRow As Long
LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row

ws.Cells(1, 9).Value = "Ticker"
ws.Cells(1, 10).Value = "Yearly Change"
ws.Cells(1, 11).Value = "Percent Change"
ws.Cells(1, 12).Value = "Total Stock Volume"
ws.Cells(1, 13).Value = "Month Close"
ws.Cells(1, 14).Value = "Month Open"
ws.Cells(2, 14).Value = ws.Cells(2, 3).Value
ws.Cells(1, 17).Value = "Ticker"
ws.Cells(1, 18).Value = "Value"
ws.Cells(2, 16).Value = "Greatest % Increase"
ws.Cells(3, 16).Value = "Greatest % Decrease"
ws.Cells(4, 16).Value = "Greatest Total Volume"

    For i = 2 To LastRow
    
    If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
         ticker = ws.Cells(i, 1).Value
         monthclose = ws.Cells(i, 6).Value
         monthopen = ws.Cells((i + 1), 3).Value
         totalvolume = totalvolume + ws.Cells(i, 7).Value
         ws.Range("I" & Summary_Table_Row).Value = ticker
         ws.Range("L" & Summary_Table_Row).Value = totalvolume
         ws.Range("M" & Summary_Table_Row).Value = monthclose
         ws.Range("N" & (Summary_Table_Row + 1)).Value = monthopen
         Summary_Table_Row = Summary_Table_Row + 1
         totalvolume = 0
    Else
          totalvolume = totalvolume + ws.Cells(i, 7).Value
    End If
    Next i

Dim LastRow2 As Long
LastRow2 = ws.Cells(Rows.Count, 9).End(xlUp).Row

    For j = 2 To LastRow2
    
              yearly = ws.Cells(j, 13).Value - ws.Cells(j, 14).Value
              ws.Cells(j, 10).Value = yearly
              If ws.Cells(j, 14).Value = 0 Then
              ws.Cells(j, 11).Value = "0%"
              Else
              percent = ws.Cells(j, 10).Value / ws.Cells(j, 14).Value
              ws.Cells(j, 11).Value = percent
              ws.Cells(j, 11).NumberFormat = "0.00%"
              If ws.Cells(j, 10).Value > 0 Then
              ws.Cells(j, 10).Interior.ColorIndex = 4
              Else
              ws.Cells(j, 10).Interior.ColorIndex = 3
              End If
        End If
    Next j

ws.Range("M:N").Delete


Minchange = WorksheetFunction.Min(ws.Range("K:K"))
Maxchange = WorksheetFunction.Max(ws.Range("K:K"))
Maxvol = Application.WorksheetFunction.Max(ws.Range("L:L"))
ws.Cells(2, 16).Value = Maxchange
ws.Cells(2, 16).NumberFormat = "0.00%"
ws.Cells(3, 16).Value = Minchange
ws.Cells(3, 16).NumberFormat = "0.00%"
ws.Cells(4, 16).Value = Maxvol


For K = 2 To LastRow2
    If ws.Cells(K, 11).Value = Maxchange Then
    ws.Cells(2, 15).Value = ws.Cells(K, 9).Value
    End If
Next K

For L = 2 To LastRow2
    If ws.Cells(L, 11).Value = Minchange Then
    ws.Cells(3, 15).Value = ws.Cells(L, 9).Value
    End If
Next L

For M = 2 To LastRow2
    If ws.Cells(M, 12).Value = Maxvol Then
    ws.Cells(4, 15).Value = ws.Cells(M, 9).Value
    End If
Next M

ws.Range("A:P").Columns.AutoFit


Next ws

End Sub
