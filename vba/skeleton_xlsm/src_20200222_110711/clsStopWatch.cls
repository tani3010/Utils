VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsStopWatch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Const clsName As String = "clsStopWatch"

Private Const DEGITS                 As Integer = 5
Private Const RESOLUTION_SECOND      As String = "s"
Private Const RESOLUTION_MILLISECOND As String = "ms"
Private Const RESOLUTION_MICROSECOND As String = "us"
Private Const RESOLUTION_NANOSECOND  As String = "ns"

Private TIME_BEGIN As Double

Private Sub Class_Initialize()
    TIME_BEGIN = Timer
End Sub

Public Sub show(Optional ByVal resolution As String = RESOLUTION_SECOND)
    Debug.Print getElapsedTime(resolution)
End Sub

Public Function getElapsedTime(Optional ByVal resolution As String = RESOLUTION_SECOND) As String
    Dim elapsed As Double
    Select Case resolution
        Case RESOLUTION_SECOND
            elapsed = Round(Timer - TIME_BEGIN, DEGITS)
        Case RESOLUTION_MILLISECOND
            elapsed = Round((Timer - TIME_BEGIN) * 1000#, DEGITS)
        Case RESOLUTION_MICROSECOND
            elapsed = Round((Timer - TIME_BEGIN) * 1000000#, DEGITS)
        Case RESOLUTION_NANOSECOND
            elapsed = Round((Timer - TIME_BEGIN) * 1000000000#, DEGITS)
        Case Else
            elapsed = Round(Timer - TIME_BEGIN, DEGITS)
            resolution = RESOLUTION_SECOND
    End Select

    getElapsedTime = "elapsed: " & elapsed & "[" & resolution & "]"
End Function

Public Sub reset()
    TIME_BEGIN = Timer
End Sub

