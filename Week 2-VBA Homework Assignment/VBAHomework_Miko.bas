Attribute VB_Name = "Module1"
Function loopAcrossSheets()

Dim tempSheet As Worksheet

For Each tempSheet In Worksheets

   tempSheet.Activate
   Call Stock
   
   
   Next tempSheet

End Function
Sub Stock()
Dim Ticker As String
Dim TotalStock_Volume As Double
Dim Summary_Table_Row As Integer
Summary_Table_Row = 2
Dim Last_Row As Double
Dim ws As Worksheets

Last_Row = Cells(Rows.Count, "A").End(xlUp).Row


For I = 2 To Last_Row

   If Cells(I + 1, 1).Value <> Cells(I, 1).Value Then
   Ticker = Cells(I, 1).Value
   TotalStock_Volume = TotalStock_Volume + Cells(I, 7).Value
   
   Range("K" & Summary_Table_Row).Value = Ticker
   Range("L" & Summary_Table_Row).Value = TotalStock_Volume
   
   Summary_Table_Row = Summary_Table_Row + 1
   Else
   
       TotalStock_Volume = TotalStock_Volume + Cells(I, 7).Value

   End If
   
   Next I

   
End Sub


