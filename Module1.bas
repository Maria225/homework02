Attribute VB_Name = "Module1"
Sub Multiyear()

    Dim ws As Worksheet
    Dim Total As Double
    Dim J As Long
    Dim ticker As String
    Dim sumtable As Long
    Dim i As Long
    Dim firstcell As Double
    Dim yearchange As Double
    Dim percentchange As Double
    Dim GreatestIncrease As Double
    Dim Greatestdecrease As Double
    Dim Greatesttotal As Double
    Dim maxyearchange As Double
    Dim RowCount As Long
    Dim Ticker2 As String

    
    For Each ws In Worksheets
        Total = 0
        J = 0
        firstcell = ws.Cells(2, 3).Value
        
        ws.Range("i1").Value = "Ticker"
        ws.Range("l1").Value = "Total Stock Volume"
        ws.Range("j1").Value = "Yearly Change"
        ws.Range("k1").Value = "Percent Change"
        ws.Range("O2").Value = "Greatest % Increase"
        ws.Range("O3").Value = "Greatest % Decrease"
        Range("O4").Value = "Greatest Total Volume"
        Range("P1").Value = "Ticker"
        Range("Q1").Value = "Value"
        
        RowCount = ws.Cells(Rows.Count, "A").End(xlUp).Row
    
        For i = 2 To RowCount
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
                Total = Total + ws.Cells(i, 7).Value
                yearchange = (ws.Cells(i, 6) - firstcell)
                     If (ws.Cells(i, 6) - firstcell) >= 0 Then
                        ws.Range("j" & 2 + J).Interior.ColorIndex = 4
                        ElseIf (ws.Cells(i, 6) - firstcell) < 0 Then
                        ws.Range("j" & 2 + J).Interior.ColorIndex = 3
                     End If
                percentchange = yearchange / Cells(i, 3).Value
                     maxyearchange = Value1
                     ws.Range("Q2") = Value1
                                        
                ws.Cells(i, 3).NumberFormat = "0.00%"
                ws.Range("i" & 2 + J).Value = ws.Cells(i, 1).Value
                ws.Range("l" & 2 + J).Value = Total
                ws.Range("j" & 2 + J).Value = yearchange
                ws.Range("k" & 2 + J).Value = percentchange
                ws.Range("k" & 2 + J).NumberFormat = "0.00%"
               
               
                Total = 0
                J = J + 1
            Else
            Total = Total + ws.Cells(i, 7).Value
            End If
        Next i
       
    For i = 2 To RowCount

 If ws.Cells(i, 12) > GreatestIncrease Then
  GreatestIncrease = ws.Cells(i, 12)
  Ticker2 = ws.Cells(i, 9)
  ws.Range("Q4").Value = GreatestI
  ws.Range("P4").Value = Ticker2
  End If
  
   If ws.Cells(i, 11) > Greatesttotal Then
  Greatesttotal = ws.Cells(i, 11)
  Ticker2 = ws.Cells(i, 9)

  ws.Range("Q2").NumberFormat = "0.00%"
  ws.Range("Q2").Value = Greatesttotal
  
 ws.Range("P2").Value = Ticker2
  
  End If
  
 If ws.Cells(i, 11) < Greatestdecrease Then
 Greatestdecrease = ws.Cells(i, 11)
  Ticker2 = ws.Cells(i, 9)
  ws.Range("Q3").NumberFormat = "0.00%"
  ws.Range("Q3").Value = Greatestdecrease
  
 ws.Range("P3").Value = Ticker2
  
  End If
  
    If ws.Cells(i, 12) > GreatestIncrease Then
  GreatestIncrease = ws.Cells(i, 12)
  Ticker2 = ws.Cells(i, 9)
  Range("Q4").Value = GreatestIncrease
  Range("P4").Value = Ticker2
  End If
  
   If ws.Cells(i, 11) > Greatesttotal Then
  Greatesttotal = ws.Cells(i, 11)
  Ticker2 = Cells(i, 9)

  Range("Q2").NumberFormat = "0.00%"
  Range("Q2").Value = Greatesttotal
  Range("P2").Value = Ticker2
  
  End If
  
 If ws.Cells(i, 11) < Greatestdecrease Then
 Greatestdecrease = ws.Cells(i, 11)
  Ticker2 = ws.Cells(i, 9)
  ws.Range("Q3").NumberFormat = "0.00%"
  ws.Range("Q3").Value = Greatestdecrease
  
 ws.Range("P3").Value = Ticker2
  
  End If

  
  Next i
Next ws
    End Sub




