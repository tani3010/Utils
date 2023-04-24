VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsErrorHandler"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Const clsName As String = "clsErrorHandler"

Private Const SYMBOL_SUCCESS As String = "[SUCCESS]"
Private Const SYMBOL_WARNING As String = "[WARNING]"
Private Const SYMBOL_ERROR   As String = "[ERROR]"

Private CLASS_LOGGER As clsLogger

Private DIR_NAME  As String
Private FILE_NAME As String

Private Sub Class_Initialize()
    Set CLASS_LOGGER = New clsLogger
End Sub

Private Sub Class_Terminate()
    Set CLASS_LOGGER = Nothing
End Sub

Public Sub writeStandardLog(ByVal checkPoint As String, Optional ByVal msg As String = "successfully completed")
    Call CLASS_LOGGER.writeLog(SYMBOL_SUCCESS & checkPoint & ", " & msg)
End Sub

Public Sub writeWarningLog(ByVal checkPoint As String, Optional ByVal msg As String = "warning was occured")
    Call CLASS_LOGGER.writeLog(SYMBOL_WARNING & checkPoint & ", " & msg)
End Sub

Public Sub writeErrorLog(ByVal checkPoint As String, Optional ByVal msg As String = vbNullString)
    Call CLASS_LOGGER.writeLog(SYMBOL_ERROR & checkPoint & ", Error No." & Err.Number & ", Description: " & Err.Description)
    If msg <> vbNullString Then
        Call CLASS_LOGGER.writeLog(SYMBOL_ERROR & checkPoint & ", " & msg)
    End If
    Call enableControl(True)
    Call enableDisplayAlerts(True)
End Sub
