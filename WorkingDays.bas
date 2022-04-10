Attribute VB_Name = "WorkingDays"
Public DateChanged As Boolean
Public RealDate As Date
Public WeeklyOff As Integer

Public Function IsWorkingDay(dt As Date) As Boolean
    WeeklyOff = Val(mdiOne.sckGo.GReadINI("[WeeklyOff]"))
    NonWorkingDays = Split(mdiOne.sckGo.GReadINI("[NonWorkingDays]"), ",")
    For Each i In NonWorkingDays
        If CDate(dt) = i Or Weekday(dt, vbSunday) = WeeklyOff Then
            IsWorkingDay = False
            Exit Function
        End If
    Next
    IsWorkingDay = True
End Function

Public Function GetWorkingDay() As Date
    If DateChanged = False Then RealDate = Now
    While Not IsWorkingDay(Date)
        MsgBox "Switching to " & Format(Date - 1, "Long Date"), vbOKOnly + vbExclamation
        Date = Date - 1
        DateChanged = True
    Wend
    GetNextWorkingDay = Date
End Function

