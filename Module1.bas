Attribute VB_Name = "Module1"
Sub year_stock_data()

' Create variables
Dim Summary_Table_Row As Integer
Dim opening_price As Double
Dim closing_price As Double
Dim yearly_change As Double
Dim percentage_change As Double

Total_Volume = 0
Summary_Table_Row = 2

 ' Bonus: counts the number of rows
   lastrow = Cells(Rows.Count, 1).End(xlUp).Row
 
 ' Set title row
        Range("J1").Value = "Ticker"
        Range("K1").Value = "Yearly Change"
        Range("L1").Value = "Percent Change"
        Range("M1").Value = "Total Stock Volume"

 
' Loop through each row
For i = 2 To lastrow
        
        
  ' Check if we are still within the same Ticker, if it is not...
    If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
 
  ' Set Ticker and Closing Price
     Ticker = Cells(i, 1).Value
     
     closing_price = Cells(i, 6).Value
 
    ' Add to the Total Volume
     Total_Volume = Total_Volume + Cells(i, 7).Value
 
  ' Print the Ticker to the summary Table
   Range("J" & Summary_Table_Row).Value = Ticker
 
  ' Print the Total Volume Table to the Summary Table
    Range("M" & Summary_Table_Row).Value = Total_Volume
 
     
   ' Print Yearly Change
     yearly_change = (closing_price - opening_price)
     
     Range("L" & Summary_Table_Row).NumberFormat = "0.00\%"
     
     ' Print Percent Change
     Percent_change = (yearly_change / opening_price) * 100
     
     ' Print the Yearly change to the summary table
      Range("K" & Summary_Table_Row).Value = yearly_change
     
    If Percent_change >= 0 Then
     Range("K" & Summary_Table_Row).Interior.ColorIndex = 4
     ElseIf Percent_change < 0 Then
      Range("K" & Summary_Table_Row).Interior.ColorIndex = 3
     
   End If
     
    ' Print the Percent Change to the Summary Table
    Range("L" & Summary_Table_Row).Value = Percent_change
     
  ' Add one to the summary table row
    Summary_Table_Row = Summary_Table_Row + 1
 
  ' Reset the Total
   Total_Volume = 0
 
  ElseIf Cells(i, 1).Value <> Cells(i - 1, 1).Value Then
  
    opening_price = Cells(i, 3).Value
     
     
 
 ' If the Cells immediately following a row is the same ticker...
  Else
 
  ' Add to the Total Volume
   Total_Volume = Total_Volume + Cells(i, 7).Value
   
     
   End If

 


   Next i
    
End Sub

Sub Reset_Button():

' Empty out the current data on the summary table
  Range("J" & Summary_Table_Row).Value = ""
  Range("M" & Summary_Table_Row).Value = ""
  Range("K" & Summary_Table_Row).Value = ""
  Range("L" & Summary_Table_Row).Value = ""

End Sub

