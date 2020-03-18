VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsExcelBook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Const clsName As String = "clsExcelBook"

Private DIR_NAME  As String
Private FILE_NAME As String
Private IS_OPEN   As Boolean
Private IS_EXIST  As Boolean
Private IS_XLS    As Boolean    ' flag that whether is it able to open in excel application
Private WK_BOOK   As Workbook

Private Sub Class_Initialize()
    IS_OPEN = False
    IS_EXIST = False
    IS_XLS = False
    DIR_NAME = vbNullString
    FILE_NAME = vbNullString
End Sub

Private Sub Class_Terminate()
    If IS_OPEN And ThisWorkbook.FullName <> getFullPath(DIR_NAME, FILE_NAME) Then
        Call WK_BOOK.Close(saveChanges:=False)
    End If
    Set WK_BOOK = Nothing
End Sub

Public Property Let set_dirName(ByVal x As String)
    DIR_NAME = x
    IS_EXIST = IIf(FILE_NAME = vbNullString, False, existFile(DIR_NAME, FILE_NAME))
End Property

Public Property Let set_fileName(ByVal x As String)
    FILE_NAME = x
    IS_EXIST = existFile(DIR_NAME, FILE_NAME)
    Select Case getExtention(x)
        Case EXT_XLS, EXT_XLSX, EXT_XLSM, EXT_CSV, EXT_TXT, EXT_TSV
            IS_XLS = True
        Case Else
            IS_XLS = False
    End Select
End Property

Public Property Let set_wkBook(ByRef wkBook As Workbook)
    Set WK_BOOK = wkBook
    IS_OPEN = True
    IS_EXIST = True
    IS_XLS = True
End Property

Public Property Get get_dirName() As String: get_dirName = DIR_NAME: End Property
Public Property Get get_fileName() As String: get_fileName = FILE_NAME: End Property
Public Property Get get_fullPath() As String: get_fullPath = getFullPath(DIR_NAME, FILE_NAME): End Property
Public Property Get get_isExist() As String: get_isExist = IS_EXIST: End Property
Public Property Get get_isOpen() As String: get_isOpen = IS_OPEN: End Property

Private Function isAlreadyOpen() As Boolean
    isAlreadyOpen = False
    Dim ws As Variant
    For Each ws In Workbooks
        If ws.name = FILE_NAME Then
            isAlreadyOpen = True
            Exit For
        End If
    Next ws
    Set ws = Nothing
End Function

Public Sub openFile(Optional ByVal readOnly As Boolean = True, Optional ByVal updateLinks As Integer = 0)
On Error GoTo LABEL_ERROR
    If isAlreadyOpen Then
        Set WK_BOOK = Workbooks(FILE_NAME)
        IS_OPEN = True
    ElseIf IS_EXIST And IS_XLS Then
        Set WK_BOOK = Workbooks.Open(fileName:=get_fullPath, readOnly:=readOnly, updateLinks:=updateLinks)
        IS_OPEN = True
    End If
    
LABEL_ERROR:
    IS_OPEN = False
End Sub

Public Sub closeFile(Optional ByVal saveChanges As Boolean = False, _
                     Optional ByVal dirName As String = vbNullString, _
                     Optional ByVal fileName As String = vbNullString)
    If Not IS_OPEN Then
        Exit Sub
    End If
    
    Call enableDisplayAlerts(False)
    If saveChanges And (Not isNone(dirName)) And (Not isNone(fileName)) Then
        Call makeDir(dirName)
        Dim wkFormat As Integer
        wkFormat = getDefaultSaveFormat
        Call changeDefaultSaveFormat(getExtention(fileName))
        Call WK_BOOK.SaveAs(fileName:=getFullPath(dirName, fileName), ConflictResolution:=xlLocalSessionChanges)
        Call changeDefaultSaveFormat(wkFormat)
    End If
    
    Call WK_BOOK.Close(saveChanges:=False)
    IS_OPEN = False
    Call enableDisplayAlerts(True)
End Sub

Public Function hasSheet(ByVal shName As String) As Boolean
    hasSheet = False
    If IS_OPEN Then
        hasSheet = stdUtil.hasSheet(shName, WK_BOOK)
    End If
End Function

Public Function hasRangeName(ByVal rangeName As String) As Boolean
    hasRangeName = False
    If IS_OPEN Then
        hasRangeName = stdUtil.hasRangeName(rangeName, WK_BOOK)
    End If
End Function

Public Sub createSheet(ByVal shName As String)
    If Not hasSheet(shName) Then
        Call stdUtil.createSheet(shName, WK_BOOK)
    End If
End Sub

Public Sub deleteSheet(ByVal shName As String)
    If Not hasSheet(shName) Then
        Call stdUtil.deleteSheet(shName, WK_BOOK)
    End If
End Sub

Public Sub copy(ByVal fromSheetName As String, ByVal fromRangeName As String, _
                ByRef toRange As Range, Optional ByVal valued As Boolean = True)
    Call openFile
    
    If hasSheet(fromSheetName) Then
        With WK_BOOK.Worksheets(fromSheetName).Range(fromRangeName)
            Call .copy(Destination:=toRange)
            If valued Then
                toRange.Value = .Value
            End If
        End With
    End If
    
    Call closeFile
End Sub

Public Sub copy2(ByVal fromSheetName As String, ByVal fromRangeName As String, _
                 ByRef toWkBook As Workbook, Optional ByVal valued As Boolean = True)
    Call openFile
    
    If hasSheet(fromSheetName) Then
        With WK_BOOK.Worksheets(fromSheetName).copy(After:=toWkBook.Worksheets(toWkBook.Worksheets.Count))

On Error Resume Next
            If valued Then
                toWkBook.Worksheets(toWkBook.Worksheets.Count).Range(fromRangeName).Value = _
                    toWkBook.Worksheets(toWkBook.Worksheets.Count).Range(fromRangeName).Value
            End If
        End With
    End If
    
    Call closeFile
End Sub

Public Function getWorkBook() As Workbook
    Set getWorkBook = WK_BOOK
End Function

Public Sub deletePublishObjects()
    If Not IS_OPEN Then
        Exit Sub
    End If
    
    With WK_BOOK.PublishObjects
        Do While .Count > 1
            Call .delete
        Loop
    End With
End Sub

Public Sub setPublishObjects(ByVal shName As String, ByVal rangeName As String, ByVal fileName As String)
    If Not hasSheet(shName) Then
        Exit Sub
    End If
    
    With WK_BOOK.Worksheets(shName)
        Call deletePublishObjects
        Call .PublishObjects.Add( _
            SourceType:=xlSourceRange, _
            fileName:=fileName, _
            Sheet:=shName, _
            Source:=rangeName, _
            HtmlType:=xlHtmlStatic)
        Call .PublishObject(1).Publish(True)
    End With
    
    Call deletePublishObjects
End Sub

Public Sub publishPDF(ByVal dirName As String, ByVal fileName As String)
    Dim sh As Variant
    
    For Each sh In WK_BOOK.Worksheets
        If sh.HPageBreaks.Count = 0 Then
            With sh.PageSetup
                .Zoom = False
                .FitToPagesWide = 1
                .FitToPagesTall = 1
                .CenterHorizontally = False
                .TopMargin = Application.CentimetersToPoints(1)
                .BottomMargin = Application.CentimetersToPoints(1)
            End With
        End If
    Next sh
    
    Call WK_BOOK.ExportAsFixedFormat( _
        Type:=xlTypePDF, _
        Quality:=xlQualityStandard, _
        fileName:=getFullPath(dirName, fileName), _
        IncludeDocProperties:=True, _
        IgnorePrintAreas:=False)
    Set sh = Nothing
End Sub

Public Sub runMacro(ByVal macroName As String, ParamArray args() As Variant)
    If IS_OPEN Then
        Select Case UBound(args)
            Case -1
                Call Application.Run(paste(DELIM_MACRO, FILE_NAME, macroName))
            Case 0
                Call Application.Run(paste(DELIM_MACRO, FILE_NAME, macroName, args(0)))
            Case 1
                Call Application.Run(paste(DELIM_MACRO, FILE_NAME, macroName, args(0), args(1)))
            Case 2
                Call Application.Run(paste(DELIM_MACRO, FILE_NAME, macroName, args(0), args(1), args(2)))
            Case 3
                Call Application.Run(paste(DELIM_MACRO, FILE_NAME, macroName, args(0), args(1), args(2), args(3)))
            Case 4
                Call Application.Run(paste(DELIM_MACRO, FILE_NAME, macroName, args(0), args(1), args(2), args(3), args(4)))
        End Select
    End If
End Sub

Public Sub exportAsCSV(ByVal dirName As String, ByVal fileName As String, ByVal shName As String)
    If Not hasSheet(shName) Then
        Exit Sub
    End If
    
    If getExtention(fileName) <> EXT_CSV Then
        fileName = paste(DELIM_EXT, Replace(fileName, DELIM_EXT, DELIM_UNDERSCORE), EXT_CSV)
    End If
    
    Dim defaultSaveFormat As Long
    Dim endRow As Long
    Dim endCol As Long
    Dim newBook As Workbook
    
    defaultSaveFormat = getDefaultSaveFormat
    Call changeDefaultSaveFormat(xlCSV)
    Call enableDisplayAlerts(False)
    
    endRow = getMaxRow(WK_BOOK.Worksheets(shName), 1, 1)
    endCol = getMaxCol(WK_BOOK.Worksheets(shName), 1, 1)
    Call makeDir(dirName)
    
    Set newBook = Workbooks.Add
    With newBook
        .Worksheets(1).Range( _
            .Worksheets(1).Cells(1, 1), _
            .Worksheets(1).Cells(endRow, endCol)).Value _
        = WK_BOOK.Worksheets(shName).Range( _
                WK_BOOK.Worksheets(shName).Cells(1, 1), _
                WK_BOOK.Worksheets(shName).Cells(endRow, endCol)).Value
        Call .SaveAs(fileName:=getFullPath(dirName, fileName), FileFormat:=xlCSV, local:=True)
        Call .Close(saveChanges:=False)
    End With
    
    Set newBook = Nothing
    Call changeDefaultSaveFormat(defaultSaveFormat)
    Call enableDisplayAlerts(True)
End Sub
