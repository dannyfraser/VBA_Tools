VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Timer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'These API functions count extremely high-resolution processing rates (i.e. CPU cycles)
'QueryPerformanceCounter counts CPU cycles, and QueryPerformanceFrequency gives the number of cycles per second
'Thus we can work out extremely accurate run-times for processes.

'Currency is used as it is essentially a structured data type with integer & decimal parts
Private Declare Function QueryPerformanceCounter Lib "kernel32" (X As Currency) As Boolean
Private Declare Function QueryPerformanceFrequency Lib "kernel32" (X As Currency) As Boolean

Private StartCounter As Currency
Private EndCounter As Currency


Private Function Overhead() As Currency

    'This calculates the CPU overhead of actually running the performance counter

    Static OverheadCounter As Currency
    If OverheadCounter = 0 Then
        Dim c1 As Currency, c2 As Currency
        QueryPerformanceCounter c1
        QueryPerformanceCounter c2
        OverheadCounter = c2 - c1
    Else
        Overhead = OverheadCounter
    End If

End Function

Private Function Frequency() As Currency
    QueryPerformanceFrequency Frequency
End Function

Public Sub Start()
    QueryPerformanceCounter StartCounter
End Sub
Public Function GetRunTime() As Double
    QueryPerformanceCounter EndCounter
    GetRunTime = (EndCounter - StartCounter - Overhead) / Frequency
End Function

