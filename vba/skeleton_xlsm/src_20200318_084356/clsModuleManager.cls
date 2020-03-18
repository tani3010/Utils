VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsModuleManager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Const clsName As String = "clsModuleManager"

' need to set reference [Microsoft Visual Basic for Applications Extensibility x.x]
' need to set macro security in devlopment tab

Private DIR_NAME As String
Private Enum MODULE_TYPE
    MODULE_BAS = 1
    MODULE_CLASS = 2
    MODULE_MSFORM = 3
    MODULE_ACTIVEXDESIGNER = 11
    MODULE_DOCUMENT = 100
End Enum

Public Property Let set_dirName(ByVal x As String)
    DIR_NAME = x
    Call Class_Initialize
End Property

Private Sub Class_Initialize()
    With ThisWorkbook
        DIR_NAME = paste(DELIM_PATH_WIN, .Path, Replace(.name, DELIM_EXT, DELIM_UNDERSCORE), _
                         "src_" & Format(Now, FMT_DATE_YYYYMMDD_HHMMSS))
    End With
End Sub

Public Sub exportAllModules()
    Dim iter As Variant
    Dim ext As String
    
    Call makeDir(DIR_NAME)
    For Each iter In ThisWorkbook.VBProject.VBComponents
        Select Case iter.Type
            Case MODULE_BAS
                ext = EXT_BAS
            Case MODULE_CLASS, MODULE_ACTIVEXDESIGNER, MODULE_DOCUMENT
                ext = EXT_CLS
            Case MODULE_MSFORM
                ext = EXT_FRM
            Case Else
                ext = EXT_TXT
        End Select
        
On Error Resume Next
        Call iter.Export(getFullPath(DIR_NAME, paste(DELIM_EXT, iter.name, ext)))
    Next iter
    Set iter = Nothing
End Sub

Private Function hasModule(ByVal moduleName As String) As Boolean
    hasModule = False
    Dim iter As Variant
    For Each iter In ThisWorkbook.VBProject.VBComponents
        If iter.name = moduleName Then
            hasModule = True
            Exit For
        End If
    Next iter
    Set iter = Nothing
End Function

Private Sub releaseModule(ByVal moduleName As String)
    If Not hasModule(moduleName) Then
        Exit Sub
    End If
    
    Call ThisWorkbook.VBProject.VBComponents(moduleName).Remove
End Sub

Private Sub releaseAllModules()
    Dim iter As Variant
    For Each iter In ThisWorkbook.VBProject.VBComponents
        Call releaseModule(iter.name)
    Next iter
    Set iter = Nothing
End Sub

Public Sub importModule(ByVal dirName As String, ByVal moduleName As String)
    If Not existFile(dirName, moduleName) Then
        Exit Sub
    End If
    Call releaseModule(moduleName)
    Call ThisWorkbook.VBProject.VBComponents.Import(getFullPath(dirName, moduleName))
End Sub

Public Sub importModuleFromFile(ByVal dirName As String, ByVal fileName As String)
    If Not existFile(dirName, fileName) Then
        Exit Sub
    End If
    Call releaseAllModules
    
    Dim ptrFile As Long
    Dim buff As String
    ptrFile = FreeFile
    Open getFullPath(dirName, fileName) For Input As #ptrFile
    Do Until EOF(ptrFile)
        Line Input #ptrFile, buff
        Call importModule(dirName, buff)
    Loop
    Close #ptrFile
End Sub
