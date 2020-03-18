VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsInternetExplorer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Const clsName As String = "clsInternetExplorer"

' IE state type
Private Enum READY_STATE
    READYSTATE_UNINITIALIZED = 0
    READYSTATE_LOADING = 1
    READYSTATE_LOADED = 2
    READYSTATE_INTERACTIVE = 3
    READYSTATE_COMPLETE = 4
End Enum

' IE action type
Public Enum IE_ACTION_TYPE
    ACTION_CLICK = 0
    ACTION_CHECK
    ACTION_SELECT
    ACTION_SET_TXT
    ACTION_GET_TXT
    ACTION_GET_SRC
    ACTION_ONCHANGE
    ACTION_ONBLUR
    ACTION_VALUE
End Enum

' attribution name
Private Const ATTR_ALT      As String = "alt"
Private Const ATTR_CLASS    As String = "class"
Private Const ATTR_INNERTXT As String = "innertext"
Private Const ATTR_ONBLUR   As String = "onblur"
Private Const ATTR_ONCHANGE As String = "onchange"
Private Const ATTR_ONCLICK  As String = "onclick"
Private Const ATTR_NAME     As String = "name"
Private Const ATTR_SRC      As String = "src"
Private Const ATTR_TITLE    As String = "title"
Private Const ATTR_TYPE     As String = "type"
Private Const ATTR_VALUE    As String = "value"

Private WIN_API As clsWinAPI
Private objIE As Object

Private Sub Class_Initialize()
    Set WIN_API = New clsWinAPI
    Call WIN_API.closeAllIE
    Set objIE = CreateObject(OBJ_IE)
    Call visibleIE(True)
    Call initializeIEPosition
    Call initializeIESize
End Sub

Private Sub Class_Terminate()
    Call objIE.Quit
    Set WIN_API = Nothing
    Set objIE = Nothing
End Sub

Public Sub visibleIE(ByVal visible As Boolean)
    If Not objIE Is Nothing Then
        objIE.visible = visible
    End If
End Sub

Public Sub closeIE()
    Call objIE.Quit
    Call WIN_API.closeAllIE
End Sub

Public Sub initializeIESize( _
    Optional ByVal height As Integer = 800, _
    Optional ByVal width As Integer = 1200)
    
    If Not objIE Is Nothing Then
        With objIE
            .height = height
            .width = width
        End With
    End If
End Sub

Public Sub initializeIEPosition( _
    Optional ByVal top As Integer = 50, _
    Optional ByVal left As Integer = 100)
    
    If Not objIE Is Nothing Then
        With objIE
            .top = top
            .left = left
        End With
    End If
End Sub

Public Sub navigate(ByVal url As String, _
    Optional ByVal title As String = vbNullString, _
    Optional ByVal initializeZoomSize As Boolean = True)
    
    Call objIE.navigate(url)
    DoEvents
    Call WIN_API.wait
    If title <> vbNullString Then
        Set objIE = getIE(title)
    End If
    Call standByProcess
    
    If initializeZoomSize Then
        Call zoomIE(100)    ' set 100% zoom
    End If
End Sub

Public Sub zoomIE(ByVal res As Long)
    Call objIE.Execwb(63&, 0, CLng(res))
End Sub

Private Sub standByProcess()
    Do While objIE.Busy Or objIE.READYSTATE < READYSTATE_COMPLETE
        DoEvents
    Loop
End Sub

Public Sub handleEvent( _
    ByVal actionType As Integer, ByRef targetValue As Variant, ByVal frameName As String, _
    ByVal tagName As String, ByVal attrName As String, ByVal attrValue As String, _
    Optional loopToEnd As Boolean = False, Optional standby As Boolean = True)
    
    Dim iter As Object
    Dim contents As Object
    
    Call getContents(frameName, objIE, contents)

    With contents
        For Each iter In .document.getElementsByTagName(tagName)
            If checkAttributes(attrName, attrValue, iter) Then
                Call actEvent(actionType, targetValue, iter)
                If Not loopToEnd Then
                    Exit For
                End If
            End If
        Next iter
    End With
    
    If standby Then
        Call standByProcess
    End If
    
    Set iter = Nothing
    Set contents = Nothing
End Sub

Public Sub handleEventSameLevel( _
    ByVal actionType As Integer, ByRef targetValue As Variant, ByVal frameName As String, _
    ByVal tagName As String, ByVal attrName As String, ByVal attrValue As String, _
    ByVal tagNameSub As String, ByVal attrNameSub As String, ByVal attrValueSub As String, _
    Optional loopToEnd As Boolean = False, Optional standby As Boolean = True)
    
    Dim iter As Object
    Dim iterSub As Object
    Dim contents As Object
    Dim contentsSub As Object
    Call getContents(frameName, objIE, contents)

    For Each iter In contents.document.getElementsByTagName(tagName)
        If checkAttributes(attrName, attrValue, iter) Then
            Set contentsSub = iter.parentElement.getElementByTagName(tagNameSub)
            For Each iterSub In contentsSub
                If checkAttributes(attrNameSub, attrValueSub, iterSub) Then
                    Call actEvent(actionType, targetValue, iterSub)
                    If Not loopToEnd Then
                        Exit For
                    End If
                End If
            Next iterSub

            If Not loopToEnd Then
                Exit For
            End If
        End If
    Next iter
    
    If standby Then
        Call standByProcess
    End If
    
    Set iter = Nothing
    Set iterSub = Nothing
    Set contents = Nothing
    Set contentsSub = Nothing
End Sub

Private Function getContents(ByVal frameName As String, ByRef ie As Variant, ByRef ieOut As Variant) As Boolean
    getContents = False
    If ie.document.frames.Length = 0 Or frameName = vbNullString Then
        Set ieOut = ie
        Exit Function
    End If
    
    Dim i As Integer
    With ie.document
        For i = 0 To .frames.Length - 1
            If .frames(i).name = frameName Then
                getContents = True
                Set ieOut = .getElementById(.frames(i).name).contentWindow
                Exit For
            ElseIf getContents(frameName, .getElementById(.frames(i).name).contentWindow, ieOut) Then
                getContents = True
                Exit For
            End If
        Next i
    End With
End Function

Private Function getIE(ByVal target As String) As Object
    Dim objSh As Object
    Dim objIE As Object
    Set objSh = CreateObject(OBJ_SHELL)
    Dim i As Integer
    For i = 1 To objSh.Windows.Count
        Set objIE = objSh.Windows.Item(i - 1)
        If Not objIE Is Nothing Then
            Do While objIE.Busy Or objIE.READYSTATE < READYSTATE_COMPLETE
                DoEvents
            Loop
            
            If left(objIE.LocationName, Len(target)) = target Then
                Set getIE = objIE
                Exit For
            End If
        End If
        Set objIE = Nothing
    Next i
    Set objSh = Nothing
End Function

Public Sub setIE(ByVal target As String)
    Dim objSh As Object
    Dim ie As Object
    Set objSh = CreateObject(OBJ_SHELL)
    Dim i As Integer
    For i = 0 To objSh.Windows.Count
        Set ie = objSh.Windows.Item(i - 1)
        If Not ie Is Nothing And VarType(ie) <> vbDataObject Then
            Do While ie.Busy Or ie.READYSTATE < READYSTATE_COMPLETE
                DoEvents
            Loop
            
            If left(ie.LocationName, Len(target)) = target Then
                Set objIE = ie
                Exit For
            End If
        End If
    Next i
    Set objSh = Nothing
End Sub

Private Function checkAttributes( _
    ByVal attrName As String, _
    ByVal attrValue As String, _
    ByRef container As Variant) As Boolean
    
    checkAttributes = False
    With container
        Select Case LCase(attrName)
            Case ATTR_NAME
                checkAttributes = (.name Like attrValue)
            Case ATTR_VALUE
                checkAttributes = (.Value Like attrValue)
            Case ATTR_INNERTXT
                checkAttributes = (.innerText Like attrValue)
            Case ATTR_CLASS
                checkAttributes = (.className Like attrValue)
            Case ATTR_TYPE
                checkAttributes = (.Type Like attrValue)
            Case ATTR_TITLE
                checkAttributes = (.title Like attrValue)
            Case ATTR_ALT
                checkAttributes = (.alt Like attrValue)
            Case ATTR_SRC
                checkAttributes = (.src Like attrValue)
            Case Else
            
On Error Resume Next
                Dim iter As Variant
                For Each iter In .Attributes
                    If iter.name = attrName Then
                        If iter.Value Like attrValue Then
                            checkAttributes = True
                            Exit For
                        End If
                    End If
                Next
                Set iter = Nothing
        End Select
    End With
End Function

Private Sub actEvent( _
    ByVal actionType As Integer, ByRef targetValue As Variant, _
    ByRef container As Variant)
    
    With container
        Select Case actionType
            Case ACTION_CLICK
                Call .Click
            Case ACTION_CHECK
                .Checked = targetValue
            Case ACTION_VALUE, ACTION_SET_TXT
                .Value = targetValue
            Case ACTION_SELECT
                .Selected = targetValue
            Case ACTION_GET_SRC
                Call targetValue.Add(.src)
            Case ACTION_GET_TXT
                targetValue = .innerText
            Case ACTION_ONCHANGE
                Call .onChange
            Case ACTION_ONBLUR
                Call .onBlur
        End Select
    End With
End Sub

Public Sub wait(ByVal t As Long)
    Call WIN_API.wait(t)
End Sub

Public Sub screenShot()
    Call AppActivate(objIE)
    Call WIN_API.screenShot
End Sub
