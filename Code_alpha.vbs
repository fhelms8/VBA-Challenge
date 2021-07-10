
beginning of year 41.81
ended year at     45.56

45.56 - 41.81 = 3.75

%change = 3.75/41.81 * 100 = 8.97% increase


Sub Alpha_Ticker()

 'Print out names for summary table'
    Cells(1, 9).Value = "Ticker"
    Cells(1, 10).Value = "Yearly Change"
    Cells(1, 11).Value = "Percent Change"
    Cells(1, 12).Value = "Total Stock Volume"
	
  ' Set an initial variables '
  Dim Ticker_Name As String

  Dim i As Long

  Dim yearly_change as Double
  Dim yearly_end as Double

  Dim Percentage_change as Double

  Dim yearly_start as Double

  Dim Vol_Total As Double

  Dim Summary_Table_Row As Integer

  ' initialize variable values
  
  Vol_Total = 0
  Yearly_start = Cells(2,3).Value
  Yearly_end = 0
  Percentage_change = 0
  Yearly_change = 0
  
   Summary_Table_Row = 2
  
   
  'Determine Last Row'
  
   LastRow = Cells(Rows.Count, 1).End(xlUp).Row

  ' Loop through all Ticker Names '
  For i = 2 To LastRow


    ' Check if we are still within the same Ticker Name, if it is not...'
    If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then

      ' Set the Ticker Name '
      Ticker_Name = Cells(i, 1).Value

      ' Add to the Volume Total '
      Vol_Total = Vol_Total + Cells(i, 7).Value
      yearly_end = Cells(i,6).Value
      yearly_change = yearly_end - yearly_start
      percent_change = yearly_change / yearly_start * 100


      ' Print the Ticker Name in the Summary Table '
      Range("I" & Summary_Table_Row).Value = Ticker_Name

      ' Print the Volume Total to the Summary Table '
      Range("L" & Summary_Table_Row).Value = Vol_Total

      ' Print the Yearly Change in Summary Table '
	Cells(Summary_table_row, 10).Value = Yearly_Change

      ' Print the Percent Change to Summary Table '
	Cells(Summary_table_row, 11).Value = Percent_Change




      ' Add one to the summary table row '
      Summary_Table_Row = Summary_Table_Row + 1

      ' do your percentage calc and write it
      
      ' Reset the Volume Total '
      Vol_Total = 0
      yearly_start = Cells(i + 1, 3).Value


    ' If the cell immediately following a row is the same brand... '
    Else

      ' Add to the Volume Total '
      Vol_Total = Vol_Total + Cells(i, 7).Value
      if (yearly_start) = 0 then
          yearly_start = cells(i,3).value  ' 29.61


	'Negative/Positive Color' 

	

	If Range("J" & Summary_Table_Row).Value < 0 Then Interior.ColorIndex = 3

	If Range("J" & Summary_Table_row).Value > 0 Then Interior.ColorIndex = 4

      End if

    End If

  Next i

End Sub




