Attribute VB_Name = "Functions"
Function Find_Total_Columns(Sheet_Index) As Integer

Sheets(Sheet_Index).Select
Find_Total_Columns = Cells(1, Columns.Count).End(xlToLeft).Column

End Function
Function Find_Total(Sheet_Index, Column_To_Search) As Integer

'Finds column_to_search column in sheet_index and counts total of rows

Sheets(Sheet_Index).Select
Range("A1").Select

Dim found As Boolean
found = False

Do While found = False
    If ActiveCell.Value = Column_To_Search Then
        found = True
        Find_Total = ActiveCell.End(xlDown).Row
    Else
        ActiveCell.Offset(0, 1).Select
    End If
Loop

End Function
Sub Remove_Empties(Last_Row, Last_Column, Sheet_Index)

'Remove entirely blanks columns

Dim i As Integer
Dim c As Integer
Dim Blank As Boolean

Sheets(Sheet_Index).Select
Range("A1").Select
'----------------------------------
For c = 1 To Last_Column
    If c > Last_Column Then
        Blank = False
    Else
        Blank = True
    End If
    For i = 2 To Last_Row
        If IsEmpty(Cells(i, c)) Then
        Else
            Blank = False
        End If
    Next i
    If Blank = True Then
        Cells(i, c).Select
        'MsgBox "deleting"
        Repeats = Repeats + 1
        ActiveCell.EntireColumn.Delete
        Call Remove_Empties(Last_Row, Last_Column - 1, Sheet_Index)
        Last_Column = Last_Column - Repeats
    End If
Next c
'----------------------------------
'If Blank2 = True Then
    'MsgBox ("All blank columns deleted")
'Else
    'MsgBox ("No blank columns found")
'End If
End Sub
Sub About_Show()
    UserForm2.Show
End Sub
Sub Clean()
Application.ScreenUpdating = False
Sheets(2).Cells.Clear
Sheets(3).Cells.Clear
Sheets(4).Cells.Clear
Sheets(5).Cells.Clear
Application.ScreenUpdating = True
MsgBox ("All Sheets Wiped")
End Sub
Sub ReadMe()
    UserForm1.Show
End Sub
