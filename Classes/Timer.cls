VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Timer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Type ExactTime
    Year As Integer
    Month As Integer
    DayOfWeek As Integer
    Day As Integer
    Hour As Integer
    Minute As Integer
    Second As Integer
    Millisecond As Integer
End Type

Private Declare Function GetLocalTime Lib "kernel32" () As ExactTime
Private Declare Sub Sleep Lib "kernel32" (ByVal Milliseconds As Long)

Private this As TimerFields
Private Type TimerFields
    StartTime As ExactTime
    EndTime As ExactTime
    TimeSnap As ExactTime
End Type

Sub SetStartTime()
    this.StartTime = GetLocalTime
End Sub
Sub SetEndTime()
    this.EndTime = GetLocalTime
End Sub
Sub SnapTime()
    this.TimeSnap = GetLocalTime
End Sub


Function ExactTimeAsString() As String
    ToString GetLocalTime
End Function
Function StartTimeAsString() As String
    ToString this.StartTime
End Function
Function EndTimeAsString() As String
    ToString this.EndTime
End Function
Function SnapTimeAsString() As String
    ToString this.TimeSnap
End Function

Private Function ToString(t As ExactTime) As String

    ToString = Format(t.Day, "00") & "/" & _
        Format(t.Month, "00") & "/" & _
        Format(t.Year, "0000") & " " & _
        Format(t.Hour, "00") & ":" & _
        Format(t.Minute, "00") & ":" & _
        Format(t.Second, "00") & "." & _
        Format(t.Millisecond, "000")

End Function


Private Function GetTimeDifference(StartTime As ExactTime, EndTime As ExactTime) As ExactTime

    GetTimeDifference.Year = 0
    GetTimeDifference.Month = 0
    GetTimeDifference.Day = 0
    GetTimeDifference.Hour = 0
    GetTimeDifference.Minute = 0
    GetTimeDifference.Second = 0
    GetTimeDifference.Millisecond = 0
    
    If EndTime.Millisecond - StartTime.Millisecond < 0 Then
        GetTimeDifference.Millisecond = GetTimeDifference.Millisecond + 1000 + EndTime.Millisecond - StartTime.Millisecond
        GetTimeDifference.Second = GetTimeDifference.Second - 1
    Else
        GetTimeDifference.Millisecond = GetTimeDifference.Millisecond + EndTime.Millisecond - StartTime.Millisecond
    End If
    
    If GetTimeDifference.Second + EndTime.Second - StartTime.Second < 0 Then
        GetTimeDifference.Second = GetTimeDifference.Second + 60 + EndTime.Second - StartTime.Second
        GetTimeDifference.Minute = GetTimeDifference.Minute - 1
    Else
        GetTimeDifference.Second = GetTimeDifference.Second + EndTime.Second - StartTime.Second
    End If
    
    If GetTimeDifference.Minute + EndTime.Minute - StartTime.Minute < 0 Then
        GetTimeDifference.Minute = GetTimeDifference.Minute + 60 + EndTime.Minute - StartTime.Minute
        GetTimeDifference.Hour = GetTimeDifference.Hour - 1
    Else
        GetTimeDifference.Minute = GetTimeDifference.Minute + EndTime.Minute - StartTime.Minute
    End If
    
    If GetTimeDifference.Hour + EndTime.Hour - StartTime.Hour < 0 Then
        GetTimeDifference.Hour = GetTimeDifference.Hour + 60 + EndTime.Hour - StartTime.Hour
        GetTimeDifference.Day = GetTimeDifference.Day - 1
    Else
        GetTimeDifference.Hour = GetTimeDifference.Hour + EndTime.Hour - StartTime.Hour
    End If
    
End Function


Function GetTimeDifference_Milliseconds() As Long

        Dim TimeDifference As ExactTime
        TimeDifference = GetTimeDifference(this.StartTime, this.EndTime)
        
        Dim Milliseconds As Long
        Milliseconds = 0
        Milliseconds = Milliseconds + CLng(TimeDifference.Hour) * 60 * 60 * 1000
        Milliseconds = Milliseconds + CLng(TimeDifference.Minute) * 60 * 1000
        Milliseconds = Milliseconds + CLng(TimeDifference.Second) * 1000
        Milliseconds = Milliseconds + TimeDifference.Millisecond
        
        GetTimeDifference_Milliseconds = Milliseconds
        
End Function


Function GetTimeDifference_Seconds() As Long
        
        Dim TimeDifference As ExactTime
        TimeDifference = GetTimeDifference(this.StartTime, this.EndTime)
        
        Dim Seconds As Long
        Seconds = 0
        Seconds = Seconds + TimeDifference.Hour * 60 * 60
        Seconds = Seconds + TimeDifference.Minute * 60
        Seconds = Seconds + TimeDifference.Second
        Seconds = Seconds + TimeDifference.Millisecond / 1000
        
    GetTimeDifference_Seconds = Seconds
    
End Function


Function GetTimeDifference_Minutes() As Long

        Dim TimeDifference As ExactTime
        TimeDifference = GetTimeDifference(this.StartTime, this.EndTime)
        
        Dim Minutes As Long
        Minutes = 0
        Minutes = Minutes + TimeDifference.Hour * 60
        Minutes = Minutes + TimeDifference.Minute
        Minutes = Minutes + TimeDifference.Second / 60
        
    GetTimeDifference_Minutes = Minutes

End Function


Sub Wait(SecondsToWait As Integer)
        
    Sleep CLng(SecondsToWait * 1000)
    
End Sub


Function GetExactTime() As Date
    SnapTime
    With this.TimeSnap
        GetExactTime = DateSerial(.Year, .Month, .Day) + TimeSerial(.Hour, .Minute, .Second)
    End With
End Function

Function GetStartTime() As Date
    With this.StartTime
        GetStartTime = DateSerial(.Year, .Month, .Day) + TimeSerial(.Hour, .Minute, .Second)
    End With
End Function

Function GetEndTime() As Date
    With this.EndTime
        GetEndTime = DateSerial(.Year, .Month, .Day) + TimeSerial(.Hour, .Minute, .Second)
    End With
End Function
