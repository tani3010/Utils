Attribute VB_Name = "stdHackVBA"
Option Explicit
  
Private Const PAGE_EXECUTE_READWRITE = &H40

#If Mac Or Win16 Then
    ' not supported for this platform
#ElseIf VBA7 And Win64 Then
    Private Declare PtrSafe Sub MoveMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As LongPtr, Source As LongPtr, ByVal Length As LongPtr)
    Private Declare PtrSafe Function VirtualProtect Lib "kernel32" (lpAddress As LongPtr, ByVal dwSize As LongPtr, ByVal flNewProtect As LongPtr, lpflOldProtect As LongPtr) As LongPtr
    Private Declare PtrSafe Function GetModuleHandleA Lib "kernel32" (ByVal lpModuleName As String) As LongPtr
    Private Declare PtrSafe Function GetProcAddress Lib "kernel32" (ByVal hModule As LongPtr, ByVal lpProcName As String) As LongPtr
    Private Declare PtrSafe Function DialogBoxParam Lib "user32" Alias "DialogBoxParamA" (ByVal hInstance As LongPtr, ByVal pTemplateName As LongPtr, ByVal hWndParent As LongPtr, ByVal lpDialogFunc As LongPtr, ByVal dwInitParam As LongPtr) As Integer
    Dim pFunc As LongPtr
#Else
    Private Declare Sub MoveMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Long, Source As Long, ByVal Length As Long)
    Private Declare Function VirtualProtect Lib "kernel32" (lpAddress As Long, ByVal dwSize As Long, ByVal flNewProtect As Long, lpflOldProtect As Long) As Long
    Private Declare Function GetModuleHandleA Lib "kernel32" (ByVal lpModuleName As String) As Long
    Private Declare Function GetProcAddress Lib "kernel32" (ByVal hModule As Long, ByVal lpProcName As String) As Long
    Private Declare Function DialogBoxParam Lib "user32" Alias "DialogBoxParamA" (ByVal hInstance As Long, ByVal pTemplateName As Long, ByVal hWndParent As Long, ByVal lpDialogFunc As Long, ByVal dwInitParam As Long) As Integer
    Dim pFunc As Long
#End If

Dim HookBytes(0 To 5) As Byte
Dim OriginBytes(0 To 5) As Byte
Dim Flag As Boolean


#If VBA7 And Win64 Then
Private Function GetPtr(ByVal Value As LongPtr) As LongPtr
#Else
Private Function GetPtr(ByVal Value As Long) As Long
#End If
    GetPtr = Value
End Function

Private Sub RecoverBytes()
    If Flag Then
        Call MoveMemory(ByVal pFunc, ByVal VarPtr(OriginBytes(0)), 6)
    End If
End Sub

Private Function Hook() As Boolean
    Dim TmpBytes(0 To 5) As Byte
    
    #If VBA7 And Win64 Then
        Dim p As LongPtr
        Dim OriginProtect As LongPtr
    #Else
        Dim p As Long
        Dim OriginProtect As Long
    #End If
       
    Hook = False
    pFunc = GetProcAddress(GetModuleHandleA("user32.dll"), "DialogBoxParamA")
    If VirtualProtect(ByVal pFunc, 6, PAGE_EXECUTE_READWRITE, OriginProtect) <> 0 Then
        MoveMemory ByVal VarPtr(TmpBytes(0)), ByVal pFunc, 6
        If TmpBytes(0) <> &H68 Then
            MoveMemory ByVal VarPtr(OriginBytes(0)), ByVal pFunc, 6
            p = GetPtr(AddressOf MyDialogBoxParam)
            HookBytes(0) = &H68
            MoveMemory ByVal VarPtr(HookBytes(1)), ByVal VarPtr(p), 4
            HookBytes(5) = &HC3
            MoveMemory ByVal pFunc, ByVal VarPtr(HookBytes(0)), 6
            Flag = True
            Hook = True
        End If
    End If
End Function

#If VBA7 And Win64 Then
Private Function MyDialogBoxParam(ByVal hInstance As LongPtr, ByVal pTemplateName As LongPtr, ByVal hWndParent As LongPtr, ByVal lpDialogFunc As LongPtr, ByVal dwInitParam As LongPtr) As Integer
#Else
Private Function MyDialogBoxParam(ByVal hInstance As Long, ByVal pTemplateName As Long, ByVal hWndParent As Long, ByVal lpDialogFunc As Long, ByVal dwInitParam As Long) As Integer
#End If
    
    If pTemplateName = 4070 Then
        MyDialogBoxParam = 1
    Else
        Call RecoverBytes
        MyDialogBoxParam = DialogBoxParam(hInstance, pTemplateName, hWndParent, lpDialogFunc, dwInitParam)
        Call Hook
    End If
End Function
  
Public Sub UnprotectVBA()
    If Hook Then
        Call MsgBox("VBAProject has been unprotectedÅI", vbInformation, "Congratulations")
    End If
End Sub

