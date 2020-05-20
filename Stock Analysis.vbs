Attribute VB_Name = "Module1"
' The ticker symbol.

  '* Yearly change from opening price at the beginning of a given year to the closing price at the end of that year.

  '* The percent change from opening price at the beginning of a given year to the closing price at the end of that year.

  '* The total stock volume of the stock.

'* You should also have conditional formatting that will highlight positive change in green and negative change in red.

Sub stocks()
  
  Dim ws As Worksheet
Dim starting_ws As Worksheet
Set starting_ws = ActiveSheet 'remember which worksheet is active in the beginning

For Each ws In ThisWorkbook.Worksheets
    ws.Activate
  
  
  
  
  
  ' Set an initial variable for holding the brand name
  Dim Ticker As String

  ' Set an initial variable for holding the total per credit card brand
  Dim Start As Double
  Start = 0
  
  Dim EndYear As Double
  EndYear = 0
  
  Dim Vol As Double
  Vol = 0
 
 Dim PerChange As Double
 PerChange = 0
 

  ' Keep track of the location for each credit card brand in the summary table
  Dim Summary_Table_Row As Integer
  Summary_Table_Row = 2

  ' Loop through all credit card purchases
  For i = 2 To Cells(Rows.Count, "A").End(xlUp).Row - 1
    
    ' Check if we are still within the same credit card brand, if it is not...
    If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then

      ' Set the Brand name
      Ticker = Cells(i, 1).Value
      Start = Cells(i, 3).Value
      EndYear = Cells(i, 6).Value
      
      Vol = Vol + Cells(i, 7).Value
      PerChange = ((Start / EndYear) * 100) - 100
' Add to the Brand Total

      ' Print the Credit Card Brand in the Summary Table
      Range("I" & Summary_Table_Row).Value = Ticker

      ' Print the Brand Amount to the Summary Table
      Range("J" & Summary_Table_Row).Value = Start
      ' Print the Brand Amount to the Summary Table
      Range("K" & Summary_Table_Row).Value = EndYear
       ' Print the Brand Amount to the Summary Table
      Range("L" & Summary_Table_Row).Value = Vol
    ' Print the Brand Amount to the Summary Table
      Range("M" & Summary_Table_Row).Value = PerChange
      ' Add one to the summary table row
      Summary_Table_Row = Summary_Table_Row + 1
      
      ' Reset the Brand Total
      Vol = 0
      PerChange = 0

    ' If the cell immediately following a row is the same brand...
    Else

       Vol = Vol + Cells(i, 7).Value

    End If

  Next i

 ' Loop through all credit card purchases
  For i = 2 To Cells(Rows.Count, "M").End(xlUp).Row - 1

' Check if the student's grade is greater than or equal to 90...
  If Cells(i, 13).Value > 0 Then

      ' Color the Passing grade green
      Cells(i, 13).Interior.ColorIndex = 4

  Else

      ' Color the Failing grade red
      Cells(i, 13).Interior.ColorIndex = 3

  End If

Next i


Dim Max As Double
Dim Min As Double
Dim MaxVol As Double


Max = WorksheetFunction.Max(Columns(13))
Min = WorksheetFunction.Min(Columns(13))
MaxVol = WorksheetFunction.Max(Columns(12))

'MsgBox Max
'MsgBox Min
'MsgBox MaxVol

For i = 2 To Cells(Rows.Count, "M").End(xlUp).Row - 1

  If Cells(i, 13) = Max Then
  
    Cells(i, 14) = "Greatest % increase"
    
  ElseIf Cells(i, 13) = Min Then
  
     Cells(i, 14) = "Greatest % decrease"
     
  ElseIf Cells(i, 12) = MaxVol Then
  
      Cells(i, 15) = "Greatest total volume"

End If

Next i


Next

starting_ws.Activate



End Sub

'1. Your solution will also be able to return the stock with the "Greatest % increase", "Greatest % decrease" and
'"Greatest total volume". The solution will look as follows:


