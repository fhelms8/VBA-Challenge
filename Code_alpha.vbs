Sub Run_Alpha_Ticker()

  Dim ws As Worksheet
 For Each ws In ThisWorkbook.Worksheets

    Call Alpha_Ticker(ws)

    
    Next ws
    
End Sub

Sub Alpha_Ticker(ws As Worksheet)

 
 'Print out names for summary table'
    ws.Cells(1, 9).Value = "Ticker"
    ws.Cells(1, 10).Value = "Yearly Change"
    ws.Cells(1, 11).Value = "Percent Change"
    ws.Cells(1, 12).Value = "Total Stock Volume"
 
 ' Set an initial variables '
  Dim Ticker_Name As String

  Dim i As Long

  Dim yearly_change As Double
  Dim yearly_end As Double

  Dim Percentage_change As Double

  Dim yearly_start As Double

  Dim Vol_Total As Double

  Dim Summary_Table_Row As Integer

  ' initialize variable values
  
  Vol_Total = 0
  yearly_start = ws.Cells(2, 3).Value
  yearly_end = 0
  Percentage_change = 0
  yearly_change = 0
  
   Summary_Table_Row = 2
  
   
  'Determine Last Row'
  
   LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row

  ' Loop through all Ticker Names '
  For i = 2 To LastRow


    ' Check if we are still within the same Ticker Name, if it is not...'
    If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then

      ' Set the Ticker Name '
      Ticker_Name = ws.Cells(i, 1).Value

      ' Add to the Volume Total '
      Vol_Total = Vol_Total + ws.Cells(i, 7).Value
      yearly_end = ws.Cells(i, 6).Value
      yearly_change = yearly_end - yearly_start
      
      If yearly_start <> 0 Then
      Percent_Change = yearly_change / yearly_start * 100
      Else
      Percent_Change = CVErr(xlErrNA)
      End If


      ' Print the Ticker Name in the Summary Table '
      ws.Range("I" & Summary_Table_Row).Value = Ticker_Name

      ' Print the Volume Total to the Summary Table '
      ws.Range("L" & Summary_Table_Row).Value = Vol_Total

      ' Print the Yearly Change in Summary Table '
      ws.Cells(Summary_Table_Row, 10).Value = yearly_change
    
     ' Print Conditional Formatting for Colors '
     Set_Conditional_Formatting (ws.Cells(Summary_Table_Row, 10))

      ' Print the Percent Change to Summary Table '
        ws.Cells(Summary_Table_Row, 11).Value = Percent_Change



      ' Add one to the summary table row '
      Summary_Table_Row = Summary_Table_Row + 1

      
      ' Reset the Volume Total '
      Vol_Total = 0
      yearly_start = ws.Cells(i + 1, 3).Value


    ' If the cell immediately following a row is the same brand... '
    Else

      ' Add to the Volume Total '
      Vol_Total = Vol_Total + ws.Cells(i, 7).Value
      If (yearly_start) = 0 Then
          yearly_start = ws.Cells(i, 3).Value ' 29.61
       End If

    
    End If
  
  Next i
  

End Sub

Sub Set_Conditional_Formatting(r As Range)
'
' This sub was recorded by using Marco Recorder '


    r.ClearFormats
    
    r.FormatConditions.Add Type:=xlCellValue, Operator:=xlGreater, _
        Formula1:="=0"
    r.FormatConditions(r.FormatConditions.Count).SetFirstPriority
    With r.FormatConditions(1).Font
        .Color = -16752384
        .TintAndShade = 0
    End With
    With r.FormatConditions(1).Interior
        .PatternColorIndex = xlAutomatic
        .Color = 13561798
        .TintAndShade = 0
    End With
    r.FormatConditions(1).StopIfTrue = True
    r.FormatConditions.Add Type:=xlCellValue, Operator:=xlLess, _
        Formula1:="=0"
    r.FormatConditions(r.FormatConditions.Count).SetFirstPriority
    With r.FormatConditions(1).Font
        .Color = -16383844
        .TintAndShade = 0
    End With
    With r.FormatConditions(1).Interior
        .PatternColorIndex = xlAutomatic
        .Color = 13551615
        .TintAndShade = 0
    End With
    r.FormatConditions(1).StopIfTrue = True
    
End Sub





