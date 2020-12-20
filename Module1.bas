Attribute VB_Name = "Module1"
Sub Stock()

Cells(1, 9).Value = "Ticker"
Cells(1, 10).Value = "Yearly Change"
Cells(1, 11).Value = "Percent Change"
Cells(1, 12).Value = "Total Stock Volume"
 
 
  ' Set an initial variable for holding the Ticker name
  Dim Open_Name As String
  
  Dim Stock_Total As Double
  Brand_Total = 0
  
  Dim Summary_Table_Row As Integer
  Summary_Table_Row = 2
   
   For i = 2 To 1000000
   
   If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
   
  Ticker_Name = Cells(i, 1).Value
  
  Stock_Total = Stock_Total + Cells(i, 7).Value
  
  Range("I" & Summary_Table_Row).Value = Ticker_Name
  Range("L" & Summary_Table_Row).Value = Stock_Total
  
  Summary_Table_Row = Summary_Table_Row + 1
  
  Stock_Total = 0
  
  Else
  
  Stock_Total = Stock_Total + Cells(i, 7).Value
  
  End If
  
  Next i
  
  
  End Sub
 Sub Interiorcolor()
  
  For i = 2 To 100000
    For j = 10 To 10
   If Cells(i, j) >= 0 Then
    Cells(i, j).Interior.ColorIndex = 4
    Else
    Cells(i, j).Interior.ColorIndex = 3
  End If
Next j
Next i

  End Sub
