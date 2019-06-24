VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Sub WorksheetLoop()


' Loop through all of the worksheets in the active Workbook
For Each ws In Worksheets ' ws will be the unique identifier for each worksheet name

' Create a variable to capture total volume per ticker symbol and to also assign an initial value of zero.
Dim Total_Volume As Double
Total_Volume = 0

' Create a counter for each ticker so this can be used to determine the open amount at the beginning of the year
' and to also assign an initial value of zero.
Dim Counter As Double
Counter = 0

'Create all the Variables being used in this program
Dim Open_Amount As Double
Dim Close_Amount As Double
Dim Difference As Double
Dim Percentage As Double
Dim SumLastRow As Double
Dim Percent() As Variant
Dim GreatPerInc As Double
Dim GreatPerDec As Double
Dim GreatVolume As Double
Dim GreatPerTick As String
Dim GreatPDTick As String
Dim GreatVolTick As String

' Creating the Headers
ws.Cells(1, 10).Value = "Ticker"
ws.Cells(1, 11).Value = "Open"
ws.Cells(1, 12).Value = "Close"
ws.Cells(1, 13).Value = "Yearly Change"
ws.Cells(1, 14).Value = "Percent Change"
ws.Cells(1, 15).Value = "Total Stock Volume"
ws.Cells(1, 18).Value = "Ticker"
ws.Cells(1, 19).Value = "Value"

' Setting the start position for displaying the ticker symbol summary
Dim Summary_Table_Row As Integer
Summary_Table_Row = 2

' Determine the number of the last row and use this instead of hard coding Last row number in the code.
LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row 'Placing "ws" om front of Cells indicates the specific worksheet

' Loop through all rows of ticker data in each worksheet
For I = 2 To LastRow

' Check to see if we are still in the same ticker symbol.  Make sure "ws" is in front of each cell because the "ws"
' indicates the specific worksheet being used.
If ws.Cells(I + 1, 1).Value = ws.Cells(I, 1).Value Then

' Add to the total Volume
Total_Volume = Total_Volume + ws.Cells(I, 7).Value

' Add to counter
Counter = Counter + 1

' If the cell immediately following a row is not the same ticker
Else

' Set the ticker symbol and make sure "ws" is in front of each cell
Ticker_Symbol = ws.Cells(I, 1).Value

' Set the Close Value of the Stock Price and make sure "ws" is in front of cell
Close_Amount = ws.Cells(I, 6).Value

' Set the Open Value of the stock Price by subtracting counter
Open_Amount = ws.Cells(I - Counter, 3).Value

' Calculate Total volume and make sure "ws" is in front of each cell
Total_Volume = Total_Volume + ws.Cells(I, 7).Value

' Print the ticker symbol in the Summary Table, the same is true for Range, need to include "ws"
ws.Range("J" & Summary_Table_Row).Value = Ticker_Symbol

' Print the total volume in the summary table, the same is true for Range, need "ws"
ws.Range("O" & Summary_Table_Row).Value = Total_Volume

' Print the Open_Amount in the Summary Table
ws.Range("K" & Summary_Table_Row).Value = Open_Amount

' Print the Close_Amount in the Summary Table
ws.Range("L" & Summary_Table_Row).Value = Close_Amount

' Calculate and then Print the difference between open and close amounts
Difference = Close_Amount - Open_Amount

' This accounts for instances where the Open price is zero which will cause an error.
If Open_Amount = 0 Then
Percentage = 0
Else
Percentage = (Close_Amount / Open_Amount) - 1
End If

ws.Range("M" & Summary_Table_Row).Value = Difference
ws.Range("N" & Summary_Table_Row).Value = Percentage
ws.Range("N" & Summary_Table_Row).NumberFormat = "0.00%"

' Add an increment of 1 to Summary Table
Summary_Table_Row = Summary_Table_Row + 1

' Reset the total_Volume amd Counter
Total_Volume = 0
Counter = 0

End If

' The following Next command will start on the next ticker symbol
Next I

' This was code I originally wrote and ended up not using and didn't
' want to lose it just in case I did need it.
    ' Assign a range to an array
    ' Place the array content in range of cells
    ' ws.Range("Q2:Q290") = Percent
    
' Determine the last row number for the Summary_Table_Row in each Worksheet
' Percent = ws.Range("N2:N290").Value
SumLastRow = ws.Cells(Rows.Count, 10).End(xlUp).Row

' Need to set a starting value for Determining Greatest Increase and Decrease Percentage
' and for the Greatest Total Volume
GreatPerInc = 0
GreatPerDec = 0
GreatVolume = 0

' The following is a loop
For M = 1 To SumLastRow

' This section assigned the color coding.
If ws.Cells(M + 1, 13).Value > 0 Then
ws.Cells(M + 1, 13).Interior.ColorIndex = 4
ElseIf ws.Cells(M + 1, 13).Value < 0 Then
ws.Cells(M + 1, 13).Interior.ColorIndex = 3
End If

' The following three sections provide the code for finding the low/high percentage as well as the Hightest Volume
If ws.Cells(M + 1, 14).Value > GreatPerInc Then
GreatPerInc = ws.Cells(M + 1, 14).Value
GreatPerTick = ws.Cells(M + 1, 10).Value
End If

If ws.Cells(M + 1, 14).Value < GreatPerDec Then
GreatPerDec = ws.Cells(M + 1, 14).Value
GreatPDTick = ws.Cells(M + 1, 10).Value
End If

If ws.Cells(M + 1, 15).Value > GreatVolume Then
GreatVolume = ws.Cells(M + 1, 15).Value
GreatVolTick = ws.Cells(M + 1, 10).Value
End If

Next M

' Print the ticker symbol and the value for the Greatest Percent Increase.  Need to incoude "ws"
ws.Cells(2, 17).Value = "Greatest % Increase"
ws.Cells(2, 18).Value = GreatPerTick
ws.Cells(2, 19).Value = GreatPerInc
ws.Cells(2, 19).NumberFormat = "0.00%"

ws.Cells(3, 17).Value = "Greatest % Decrease"
ws.Cells(3, 18).Value = GreatPDTick
ws.Cells(3, 19).Value = GreatPerDec
ws.Cells(3, 19).NumberFormat = "0.00%"

ws.Cells(4, 17).Value = "Greatest Total Volume"
ws.Cells(4, 18).Value = GreatVolTick
ws.Cells(4, 19).Value = GreatVolume

' This was a double check and instead of deleting it, I made it a comment
' As a double check prints the value in an unused cell
' ws.Cells(1, 15).Value = SumLastRow

Next ws

End Sub
