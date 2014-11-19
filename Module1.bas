Attribute VB_Name = "Module1"
Public Repeats As Integer

Sub Main()

Application.ScreenUpdating = False
Repeats = 0
'--------------------------------------------------------
Dim Last_Row As Integer
Last_Row = Find_Total(2, "Technical Image Id")

Dim Last_Column As Integer
Last_Column = Find_Total_Columns(2)

Call Remove_Empties(Last_Row, Last_Column, 2)
'--------------------------------------------------------
Last_Row = Find_Total(3, "Technical Image Id")

Last_Column = Find_Total_Columns(3)

Call Remove_Empties(Last_Row, Last_Column, 3)
'--------------------------------------------------------
Last_Row = Find_Total(4, "Technical Image Id")

Last_Column = Find_Total_Columns(4)

Call Remove_Empties(Last_Row, Last_Column, 4)
'--------------------------------------------------------
Last_Row = Find_Total(5, "Technical Image Id")

Last_Column = Find_Total_Columns(5)

Call Remove_Empties(Last_Row, Last_Column, 5)
'--------------------------------------------------------
MsgBox "All Blank Rows Deleted"

Application.ScreenUpdating = True

End Sub


