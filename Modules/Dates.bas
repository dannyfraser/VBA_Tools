Attribute VB_Name = "Dates"
Option Explicit

Public Function DaysInMonth(TargetDate As Date) As Integer
    DaysInMonth = Day(DateAdd("d", -1, DateSerial(Year(TargetDate), Month(TargetDate) + 1, 1)))
End Function

Public Function IsLeapYear(Year As Integer) As Boolean
    IsLeapYear = (Month(DateSerial(Year, 2, 29)) = 2)
End Function
