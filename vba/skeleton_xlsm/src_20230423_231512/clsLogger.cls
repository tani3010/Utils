VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsLogger"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Const clsName As String = "clsLogger"

Private DIR_NAME  As String
Private FILE_NAME As String
Private FULL_NAME As String

Public Property Let set_dirName(ByVal x As String)
    DIR_NAME = x
    FULL_NAME = getFullPath(DIR_NAME, FILE_NAME)
End Property

Public Property Let set_fileName(ByVal x As String)
    FILE_NAME = x
    FULL_NAME = getFullPath(DIR_NAME, FILE_NAME)
End Property

Private Sub Class_Initialize()
    With ThisWorkbook
        DIR_NAME = .Path
        FILE_NAME = paste(DELIM_EXT, .name, EXT_LOG)
        FULL_NAME = getFullPath(DIR_NAME, FILE_NAME)
    End With
End Sub

Public Sub writeLog(ByVal msg As String)
    Dim ptrFile As Long
On Error Resume Next
    ptrFile = FreeFile
    Open FULL_NAME For Append As #ptrFile
    Print #ptrFile, getTimeStamp, msg
    Close #ptrFile
End Sub

Private Function getTimeStamp() As String
    getTimeStamp = format(Now, FMT_DATE_YYYYMMDDHHMMSS_SEP)
End Function
