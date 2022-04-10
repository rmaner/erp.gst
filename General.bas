Attribute VB_Name = "General"
Public Sub SaveGrid(ByRef fg As VSFlexGrid, Optional FName As String)
    On Error Resume Next
    fg.Row = 0
    If MsgBox("Do you wish to save the grid?", vbYesNo + vbQuestion) = vbYes Then
        If FName = "" Then
            mdiOne.CDlg.FileName = Format(Now, "DD-MMM-YY HHMM")
        Else
            mdiOne.CDlg.FileName = FName
        End If
            
        mdiOne.CDlg.Filter = CompanyName & " Excel Report |*.xls"
        mdiOne.CDlg.ShowSave
        If mdiOne.CDlg.CancelError = False Then fg.SaveGrid mdiOne.CDlg.FileName, flexFileExcel, flexXLSaveFixedCells
    End If
End Sub

Public Function FlxSum(ByRef fg As VSFlexGrid) As String
    Dim sum As Double
    sum = 0
    For i = 0 To fg.SelectedRows - 1
        If fg.SelectedRow(i) >= 1 Then
            sum = sum + Val(fg.TextMatrix(fg.SelectedRow(i), fg.Col))
        End If
    Next
    FlxSum = "Sum on col " & fg.Col & " is " & sum & " for " & fg.SelectedRows & " rows."
End Function

Public Function getLastDateOfMonth(pDate As Date) As Date
    Select Case Month(pDate)
        Case 1, 3, 5, 7, 8, 10, 12  'months having 31 days
            getLastDateOfMonth = CDate("31-" & Month(pDate) & "-" & Year(pDate))
        Case 2
            getLastDateOfMonth = CDate("28-" & Month(pDate) & "-" & Year(pDate))
        Case 4, 6, 9, 11
            getLastDateOfMonth = CDate("30-" & Month(pDate) & "-" & Year(pDate))
    End Select
End Function
