Attribute VB_Name = "Module1"
Sub STOCK_TICKER()


For Each ws In Worksheets
    ws.Activate
    
' Last Row formula
LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row

' Assign Variables
Dim Ticker_Name As String
Dim Ticker_Total As String
    Ticker_Total = 0
Dim Summary_Table_Row As Integer
    Summary_Table_Row = 2
Dim Open_Value As Double
    Open_Value = Range("C2")
Dim Close_Value As Double

' Add Column Headers
Cells(1, 9).Value = "Ticker"
Cells(1, 10).Value = "Yearly Change"
Cells(1, 11).Value = "Percent Change"
Cells(1, 12).Value = "Total Stock Volume"

  ' Loop through all stock tickers
  For i = 2 To LastRow

    ' Check if we are still within the same stock ticker name, if it is not then
    If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then

      ' Set the Ticker Name
      Ticker_Name = Cells(i, 1).Value

      ' Add to the Ticker Total
      Ticker_Total = Ticker_Total + Cells(i, 7).Value

      ' Print the Ticker Name in the Summary Table
      Range("I" & Summary_Table_Row).Value = Ticker_Name
      
      ' Assign Closed_Value and Print the yearly change in the Summary Table
      ' and set the fill color by pos or neg value
      Closed_Value = Cells(i, 6)
      
      Range("J" & Summary_Table_Row).Value = Closed_Value - Open_Value
        
        If Range("J" & Summary_Table_Row).Value < 0 Then
            Range("J" & Summary_Table_Row).Interior.ColorIndex = 3
        Else
             Range("J" & Summary_Table_Row).Interior.ColorIndex = 4
        End If
        
      'Print Percent Change and format to percent.
      Range("K" & Summary_Table_Row).Value = Range("J" & Summary_Table_Row).Value / Open_Value
      Range("K" & Summary_Table_Row).Style = "Percent"
      

      ' Print the Ticker Amount to the Summary Table
      Range("L" & Summary_Table_Row).Value = Ticker_Total

      ' Add one to the summary table row and assign new Open_Value snd change to 1 if 0
      Summary_Table_Row = Summary_Table_Row + 1
      Open_Value = Cells(i + 1, 3)
       If Open_Value = 0 Then
        Open_Value = 1
        End If
      
      
      ' Reset the Ticker Total
      Ticker_Total = 0

    ' If the cell immediately following a row is the same ticker then
    Else

      ' Add to the Ticker Total
      Ticker_Total = Ticker_Total + Cells(i, 7).Value

    End If

  Next i
  
  Next ws
  
End Sub
