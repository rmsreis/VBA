'---------------------------------------
'PART I
'Easy
'Create a script that will loop through one year of stock data for each run and return the total volume each stock had over that year.
'You will also need to display the ticker symbol to coincide with the total stock volume.
'--------------------------------------------

Sub year_stock()

' Set an initial variable for stock ticker
  Dim Stock_Ticker As String
  
' Set an initial variable for holding the total volume per credit card brand
  Dim Volume_Total As Double
  Volume_Total = 0
  
 ' Grabbed the WorksheetName
        WorksheetName = ActiveSheet.Name
        
  'Name new cells for summary
  ActiveSheet.Cells(1, 9).Value = "Stock Ticker"
  ActiveSheet.Cells(1, 9).Interior.ColorIndex = 37
  
  ActiveSheet.Cells(1, 10).Value = "Total Volume " + WorksheetName
  ActiveSheet.Cells(1, 10).Interior.ColorIndex = 37
  
  ' Keep track of the location for each Stock Ticker in the summary table
  Dim Summary_Table_Row As Integer
  Summary_Table_Row = 2

'Determine the Last Row
        LastRow = ActiveSheet.Cells(Rows.Count, 1).End(xlUp).Row
  
'Loop through all tickers
  For i = 2 To LastRow
  
  ' Check if we are still within the same ticker, if it is not...
    If ActiveSheet.Cells(i + 1, 1).Value <> ActiveSheet.Cells(i, 1).Value Then
     
     ' Set the Ticker Name
      Stock_Ticker = ActiveSheet.Cells(i, 1).Value
    
    ' Add to the Total Volume
      Volume_Total = Volume_Total + ActiveSheet.Cells(i, 7).Value
      
      ' Print the Stock Ticker in the Summary Table
      ActiveSheet.Range("I" & Summary_Table_Row).Value = Stock_Ticker
      ActiveSheet.Range("I" & Summary_Table_Row).Interior.ColorIndex = 37


      ' Print the Total Volume to the Summary Table
      ActiveSheet.Range("J" & Summary_Table_Row).Value = Volume_Total
      ActiveSheet.Range("J" & Summary_Table_Row).Interior.ColorIndex = 37

      ' Add one to the summary table row
      Summary_Table_Row = Summary_Table_Row + 1
      
      ' Reset the Total Volume
      Volume_Total = 0
      
        ' If the cell immediately following a row is the same stock ticker...
    Else

      ' Add to the Total Volume
      Volume_Total = Volume_Total + ActiveSheet.Cells(i, 7).Value

    End If
    
    Next i


End Sub



