Public Sub Main()

	Dim sh As Worksheet

	For Each sh In ThisWorkbook.Worksheets
		SortSomething sh
	Next sh

End Sub

Sub Multiple_Year_Stock()

	Dim Ticker As String
	Dim Volume_T As Double
		Volume_T = 0
	Dim Opening_Date As Double
	Dim Closing_Date As Double
	Dim Closing_Price As Double
	Dim Opening_Price As Long
			Opening_Price = 2
	Dim Percent_Closing As Double
	Dim Table_Row As Integer
		Table_Row = 2
	Dim Last_Row As Long
		Last_Row = Cells(Rows.Count, 1).End(xlUp).Row
		
	Dim WS_Count As Integer
	Dim I As Integer

	 WS_Count = ActiveWorkbook.Worksheets.Count
	 
	 For I = 1 To WS_Count

	   MsgBox ActiveWorkbook.Worksheets(I).Name
	 Next I

	Range("J1").Value = "Ticker"
	Range("K1").Value = "Yearly Change"
	Range("L1").Value = "Percent Change"
	Range("M1").Value = "Total Stock Volume"
	Range("Q1").Value = "Ticker Total"
	Range("R1").Value = "Value"
	Range("P2").Value = "Greatest % Increase"
	Range("P3").Value = "Greatest % Decrease"
	Range("P4").Value = "Greatest Total Volume"


		For x = 2 To Last_Row
		
		If Cells(x + 1, 1).Value <> Cells(x, 1).Value Then
		
			Ticker = Cells(x, 1).Value
			
			Volume_T = Volume_T + Cells(x, 7).Value
									
			Range("J" & Table_Row).Value = Ticker
			Range("M" & Table_Row).Value = Volume_T
			

			Opening_Date = Range("C" & Opening_Price)
			
			Closing_Date = Range("F" & x)
			
			Closing_Price = Closing_Date - Opening_Date
			
			Range("K" & Table_Row).Value = Closing_Price
			
			If Opening_Date = 0 Then
			Percent_Closing = 0
			
					Else
							Opening_Date = Range("C" & Opening_Price)
							Percent_Closing = Closing_Price / Opening_Date
			End If
			
					Range("L" & Table_Row).NumberFormat = "0.00%"
					Range("L" & Table_Row).Value = Percent_Closing


					Table_Row = Table_Row + 1
			
					Volume_Total = 0
			Else
			
				Volume_T = Volume_T + Cells(x, 7).Value
		
		End If
		
		If Range("K" & Table_Row).Value > 0 Then
		Range("K" & Table_Row).Interior.ColorIndex = 4
		Else
		Range("K" & Table_Row).Interior.ColorIndex = 3
		End If
		
		Next x
		
		Dim Last_Row_Sum As Long
		Last_Row_Sum = Cells(Rows.Count, 10).End(xlUp).Row
		
		For x = 2 To Last_Row_Sum
		
		If Range("L" & x).Value > Range("R2").Value Then
											Range("R2").Value = Range("L" & x).Value
											Range("Q2").Value = Range("J" & x).Value
									End If
									
	If Range("L" & x).Value < Range("R3").Value Then
											Range("R3").Value = Range("L" & x).Value
											Range("Q3").Value = Range("J" & x).Value
									End If
									
	If Range("M" & x).Value > Range("R4").Value Then
											Range("R4").Value = Range("M" & x).Value
											Range("Q4").Value = Range("J" & x).Value
									End If
		
		
		Next x
		
							Range("R2").NumberFormat = "0.00%"
							Range("R3").NumberFormat = "0.00%"
							
							For Each sht In ThisWorkbook.Worksheets
											sht.Cells.EntireColumn.AutoFit
  
  Next sht
							
			End Sub

