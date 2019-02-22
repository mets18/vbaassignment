Attribute VB_Name = "Module1"
Sub Master()
   Dim wb As Worksheet
   Application.ScreenUpdating = False
   For Each wb In Worksheets
       wb.Select
       Call tickervolchg
   Next
   Application.ScreenUpdating = True
End Sub
 
 Sub tickervolchg()
 'Label and format the Columns
    Range("I1").Value = "Ticker"
    Range("J1").Value = "Yearly Change"
    Range("K1").Value = "Percent Change"
    Range("L1").Value = "Total Stock Volume"
    Columns("K").NumberFormat = "0.00%"
 
 
 ' Set an initial variable for holding the ticker
  Dim Ticker As String


  ' Set an initial variable for holding the total volume per ticker
  Dim Totalvol As Double
  Totalvol = 0
  
  ' Set an initial variable for opening value per ticker
  Dim oval As Double
  
  
  ' Set an initial variable for closing value per ticker
  Dim cval As Double
  
  ' Set integers for determining opening value
  Dim k As Integer
  k = 2
  

  ' Keep track of the location for each ticker in the summary table
  Dim Summary_Table_Row As Integer
  Summary_Table_Row = 2

  ' Determine last row
  Dim lastrow As Double
  lastrow = Cells(Rows.Count, 1).End(xlUp).Row
  
  ' Loop through all tickers
  Dim i As Double
  For i = 2 To lastrow

    ' find opening value
    If k = Summary_Table_Row Then
    oval = Cells(i, 3).Value
    k = k + 1
    End If

    ' Check if we are still within the same ticker, if it is not...
    If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then

      ' Set the Ticker
      Ticker = Cells(i, 1).Value

      ' Determine closing value
      cval = Cells(i, 6).Value
      
      ' Add to the Total Volume
      Totalvol = Totalvol + Cells(i, 7).Value

      ' Print the Ticker in the Summary Table
      Range("I" & Summary_Table_Row).Value = Ticker

      ' Print the yearly change
      Range("J" & Summary_Table_Row).Value = cval - oval
      
      ' Check Null value and print percent Change
            If oval <> 0 Then
                Range("K" & Summary_Table_Row).Value = (cval - oval) / oval
            Else
                Range("K" & Summary_Table_Row).Value = 0
            End If
      
      'Conditional formatting of yearly change
            If Range("J" & Summary_Table_Row).Value > 0 Then
                Range("J" & Summary_Table_Row).Interior.ColorIndex = 4
            Else
                Range("J" & Summary_Table_Row).Interior.ColorIndex = 3
            End If
      
      ' Print the Total Volume to the Summary Table
      Range("L" & Summary_Table_Row).Value = Totalvol
      
      ' Add one to the summary table row
      Summary_Table_Row = Summary_Table_Row + 1
      
      ' Reset the Total Volume
      Totalvol = 0

    ' If the cell immediately following a row is the same ticker
    Else
    

      ' Add to the Total Volume
      Totalvol = Totalvol + Cells(i, 7).Value

    End If

  Next i

End Sub
