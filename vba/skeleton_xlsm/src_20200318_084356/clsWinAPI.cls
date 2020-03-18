VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsWinAPI"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Const clsName As String = "clsWinAPI"

#If VBA7 And Win64 Then
    Private Declare PtrSafe Sub sleep Lib "kernel32" Alias "Sleep" (ByVal ms As Long)
    Private Declare PtrSafe Function URLDownloadToFile Lib "urlmon" Alias "URLDownloadToFileA" (ByVal pCaller As Long, ByVal szURL As String, ByVal szFileName As String, ByVal dwReserved As Long, ByVal lpfnCB As Long) As Long
    Private Declare PtrSafe Function SHCreateDirectoryEx Lib "shell32" Alias "SHCreateDirectoryExA" (ByVal hWnd As Long, ByVal pszPath As String, ByVal psa As Long) As Long
    Private Declare PtrSafe Sub keybd_event Lib "user32.dll" (ByVal bVk As Byte, ByVal bScan As Byte, ByVal dwFlags As Long, ByVal dwExtraInfo As Long)
    Private Declare PtrSafe Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
    Private Declare PtrSafe Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
    Private Declare PtrSafe Function SetLayeredWindowAttributes Lib "user32" (ByVal hWnd As Long, ByVal crKey As Long, ByVal bAlpha As Byte, ByVal dwFlags As Long) As Long
    Private Declare PtrSafe Function GetLastActivePopup Lib "user32" (ByVal hwndOwnder As Long) As Long
    Private Declare PtrSafe Function SetEnvironmentVariable Lib "kernel32" Alias "SetEnvironmentVariableA" (ByVal lpName As String, ByVal lpValue As String) As Long
    Private Declare PtrSafe Function SetDefaultDllDirectories Lib "kernel32.dll" (ByVal DirectoryFlags As Long) As Long
    Private Declare PtrSafe Function addDLLDirectory Lib "kernel32.dll" Alias "AddDllDirectory" (ByVal dirName As String) As LongPtr
    Private Declare PtrSafe Function RemoveDllDirectory Lib "kernel32.dll" (ByVal cookie As LongPtr) As Long
#Else
    Private Declare Sub sleep Lib "kernel32" Alias "Sleep" (ByVal ms As Long)
    Private Declare Function URLDownloadToFile Lib "urlmon" Alias "URLDownloadToFileA" (ByVal pCaller As Long, ByVal szURL As String, ByVal szFileName As String, ByVal dwReserved As Long, ByVal lpfnCB As Long) As Long
    Private Declare Function SHCreateDirectoryEx Lib "shell32" Alias "SHCreateDirectoryExA" (ByVal hWnd As Long, ByVal pszPath As String, ByVal psa As Long) As Long
    Private Declare Sub keybd_event Lib "user32.dll" (ByVal bVk As Byte, ByVal bScan As Byte, ByVal dwFlags As Long, ByVal dwExtraInfo As Long)
    Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
    Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
    Private Declare Function SetLayeredWindowAttributes Lib "user32" (ByVal hWnd As Long, ByVal crKey As Long, ByVal bAlpha As Byte, ByVal dwFlags As Long) As Long
    Private Declare Function GetLastActivePopup Lib "user32" (ByVal hwndOwnder As Long) As Long
    Private Declare Function SetEnvironmentVariable Lib "kernel32" Alias "SetEnvironmentVariableA" (ByVal lpName As String, ByVal lpValue As String) As Long
    Private Declare Function SetDefaultDllDirectories Lib "kernel32.dll" (ByVal DirectoryFlags As Long) As Long
    Private Declare Function addDLLDirectory Lib "kernel32.dll" Alias "AddDllDirectory" (ByVal dirName As String) As Long
    Private Declare Function RemoveDllDirectory Lib "kernel32.dll" (ByVal cookie As Long) As Long
#End If

Private Const EXE_IE As String = "IEXPLORE.EXE"

' ===== wait time =====
Private Const TIME_WAIT_DEFAULT As Long = 5000
Private Const TIME_WAIT_DIALOG  As Long = 1000
Private Const TIME_WATI_BUTTON  As Long = 2000
Private Const TIME_WAIT_SEC_1   As Long = 1000

' ===== constant number for screen shot =====
Private Const SW_SHOWMAXIMIZED As Long = 3
Private Const KEYEVENTF_KEYUP  As Long = &H2
Private Const VK_LCONTROL      As Long = &HA2
Private Const VK_V             As Long = &H56
Private Const VK_SNAPSHOT      As Long = &H2C
Private Const VK_MENU          As Long = &H12

Private Const LWA_COLLORKEY As Long = &H1
Private Const LWA_ALPHA     As Long = &H2
Private Const GWL_EXSTYLE   As Long = (-20)
Private Const WS_EX_LAYERD  As Long = &H80000

Public Sub makeDir(ByVal dirName As String)
    Dim rc As Long
    If Dir(dirName, vbDirectory) = vbNullString Then
        rc = SHCreateDirectoryEx(0&, dirName, 0&)
    End If
End Sub

Public Sub closeAllIE(Optional waitTime As Long = TIME_WAIT_DEFAULT)
    Dim objSh As Object
    Dim objIE As Object
    Set objSh = CreateObject(OBJ_SHELL)
    Dim i As Integer
    For i = objSh.Windows.Count To 1 Step -1
        Set objIE = objSh.Windows.Item(i - 1)
        If Not objIE Is Nothing Then
            If InStr(UCase(objIE.FullName), EXE_IE) Then
                Call objIE.Quit
            End If
        End If
    Next i
    Set objIE = Nothing
    Set objSh = Nothing
    Call sleep(waitTime)
End Sub

Public Sub downloadFile(ByVal url As String, ByVal dirName As String, ByVal fileName As String)
    Dim rc As Long
    Call makeDir(dirName)
    If Not existFile(dirName, fileName) Then
        rc = URLDownloadToFile(0, url, getFullPath(dirName, fileName), 0, 0)
    End If
End Sub

Public Sub screenShot()
    Call keybd_event(VK_MENU, 0, 0, 0)
    Call keybd_event(VK_SNAPSHOT, 0, 0, 0)
    Call keybd_event(VK_SNAPSHOT, 0, KEYEVENTF_KEYUP, 0)
    Call keybd_event(VK_MENU, 0, KEYEVENTF_KEYUP, 0)
    Call sleep(TIME_WAIT_SEC_1)
    DoEvents
End Sub

Public Sub wait(Optional ByVal waitTime As Double = TIME_WAIT_DEFAULT)
    Call sleep(waitTime)
End Sub

Public Sub setOpacity(ByVal opacity As Double)
    Dim h As Long
    Dim attr As Long
    h = Application.hWnd
    attr = GetWindowLong(h, GWL_EXSTYLE)
    Call SetWindowLong(h, GWL_EXSTYLE, attr Or WS_EX_LAYERD)
    Call SetLayeredWindowAttributes(h, RGB(0, 0, 0), CByte(opacity * 255), LWA_ALPHA)
End Sub

Public Function getEnvironmentVariable(ByVal environmentVariableName As String) As String
    getEnvironmentVariable = Environ(environmentVariableName)
End Function

Public Sub setTemporaryEnvironmentVariable( _
    ByVal environmentVariableName As String, _
    ByVal environmentVariablePath As String, Optional ByVal enableOverride As Boolean = False)

    Dim currentEnvironmentPath As String
    currentEnvironmentPath = Environ(environmentVariableName)
    If enableOverride Then
        Call SetEnvironmentVariable(environmentVariableName, environmentVariablePath)
    Else
        Call SetEnvironmentVariable( _
            environmentVariableName, _
            paste(DELIM_SEMICOLON, environmentVariablePath, currentEnvironmentPath))
    End If
End Sub

Public Sub showAllEnvironmentVariables()
    Dim i As Integer
    Dim env As String
    i = 1
    Do
        env = Environ(i)
        If env <> vbNullString Then
            Debug.Print env
        End If
        i = i + 1
    Loop Until env = vbNullString
End Sub

Public Sub setDLLDirectory(ByVal targetDir As String)
    If Not existDir(targetDir) Then
        Exit Sub
    End If

    Dim b As Byte
    Dim p As Long
    b = SetDefaultDllDirectories(&H1000)
    p = addDLLDirectory(StrConv(targetDir, vbUnicode))
End Sub


