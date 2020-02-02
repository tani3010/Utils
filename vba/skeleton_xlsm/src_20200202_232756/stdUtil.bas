Attribute VB_Name = "stdUtil"
Option Explicit

Private Const stdName As String = "stdUtil"

Public Sub initializeErrorHandler()
    If isNone(ERROR_HANDLER) Then
        Set ERROR_HANDLER = New clsErrorHandler
    End If
End Sub

Public Sub buildSetterGetter(ByVal memberName As String, ByVal memberTypeName As String, _
                             Optional ByVal isByRef As Boolean = False, Optional ByVal argName As String = "val")
    Dim ref As String
    Dim convertedMemberTypeName As String
    convertedMemberTypeName = StrConv(memberTypeName, vbProperCase)
    ref = IIf(isByRef, "ByRef", "ByVal")
    Debug.Print paste(vbNullString, "Public Property Let set_", memberName, "(", ref, " ", argName, " As ", convertedMemberTypeName, "): ", memberName, " = ", argName, ": End Property")
    Debug.Print paste(vbNullString, "Public Property Get get_", memberName, "() As ", memberTypeName, ": get_", memberName, " = ", memberName, ": End Property")
End Sub

Public Function getFullPath(ByVal dirName As String, ByVal FileName As String) As String
    getFullPath = IIf(Right(dirName, 1) = DELIM_PATH_WIN, dirName, dirName & DELIM_PATH_WIN) & FileName
End Function

Public Function existFile(ByVal dirName As String, ByVal FileName As String) As Boolean
    existFile = Dir(getFullPath(dirName, FileName)) <> vbNullString
End Function

Public Function existDir(ByVal dirName As String) As Boolean
    existDir = Dir(dirName, vbDirectory) <> vbNullString
End Function

Public Sub makeDir(ByVal dirName As String)
    Dim api As clsWinAPI
    Set api = New clsWinAPI
    Call api.makeDir(dirName)
    Set api = Nothing
End Sub

Public Sub copyDir(ByVal fromDirName As String, ByVal toDirName As String)
    Dim fso As Object
    
    If Not existDir(fromDirName) Or (fromDirName = toDirName) Then
        Exit Sub
    End If
    
    Set fso = CreateObject(OBJ_FSO)
    Call fso.CopyFolder(fromDirName, toDirName)
    
    Set fso = Nothing
End Sub

Public Sub renameDir(ByVal fromDirName As String, ByVal toDirName As String)
    Dim fso As Object
    Dim fl As Object
    
    If Not existDir(fromDirName) Then
        Exit Sub
    End If
    
    Set fso = CreateObject(OBJ_FSO)
    Set fl = fso.GetFolder(fromDirName)
    fl.name = getFileName(toDirName)
    
    Set fso = Nothing
    Set fl = Nothing
End Sub

Public Sub renameFile(ByVal dirName As String, ByVal fromFileName As String, ByVal toFileName As String)
    Dim fso As Object
    Dim fl As Object
    
    If Not existFile(dirName, fromFileName) Then
        Exit Sub
    End If
    
    Set fso = CreateObject(OBJ_FSO)
    Set fl = fso.GetFolder(getFullPath(dirName, fromFileName))
    fl.name = getFileName(toFileName)
    
    Set fso = Nothing
    Set fl = Nothing
End Sub

Public Function getDefaultSaveFormat() As Long
    getDefaultSaveFormat = Application.defaultSaveFormat
End Function

Public Sub changeDefaultSaveFormat(ByVal xlFormat As Variant)
    With Application
        Select Case VarType(xlFormat)
            Case vbInteger, vbLong, vbSingle, vbDouble
                .defaultSaveFormat = Int(xlFormat)
            Case vbString
                Select Case Trim(LCase(xlFormat))
                    Case EXT_XLS
                        ' saveFormatParcer = xlExcel8
                        .defaultSaveFormat = xlWorkbookNormal
                    
                    Case EXT_XLSX
                        ' saveFormatParcer = xlWorkbookDefault
                        .defaultSaveFormat = xlOpenXMLWorkbook
                    
                    Case EXT_XLSM
                        .defaultSaveFormat = xlOpenXMLWorkbookMacroEnabled
                    
                    Case EXT_XLAM
                        .defaultSaveFormat = xlOpenXMLAddIn
                
                    Case EXT_CSV
                        .defaultSaveFormat = xlCSV
                    
                    Case EXT_XML
                        .defaultSaveFormat = xlXMLSpreadsheet
                    
                End Select
        End Select
    End With
End Sub

Public Sub showAllAddin()
    Dim iter As Variant
On Error GoTo LABEL_FINALLY
    For Each iter In Application.AddIns
        Debug.Print iter.title
    Next iter
    
    For Each iter In Application.COMAddIns
        Debug.Print iter.progID
    Next iter
    
LABEL_FINALLY:
    Set iter = Nothing
End Sub

Public Function hasAddin(ByVal addinName As String) As Boolean
    Dim iter As Variant
    hasAddin = False
On Error GoTo LABEL_FINALLY
    For Each iter In Application.AddIns
        If iter.title = addinName Then
            hasAddin = True
            Exit For
        End If
    Next iter
    
    For Each iter In Application.COMAddIns
        If iter.progID = addinName Then
            hasAddin = True
            Exit For
        End If
    Next iter
    
LABEL_FINALLY:
    Set iter = Nothing
End Function

Public Function installedAddin(ByVal addinName As String, _
                               Optional ByVal isCOMAddin As Boolean = False) As Boolean
    installedAddin = False
On Error Resume Next
    If Not hasAddin(addinName) Then
        Exit Function
    End If
    
    If isCOMAddin Then
        installedAddin = Application.COMAddIns(addinName).Connect
    Else
        installedAddin = Application.AddIns(addinName).Installed
    End If
End Function

Public Sub enableAddin(ByVal addinName As String, ByVal enabled As Boolean, _
                       Optional ByVal isCOMAddin As Boolean = False)
    If Not hasAddin(addinName) Then
        Exit Sub
    End If
    
    If enabled = installedAddin(addinName) Then
        Exit Sub
    End If
    
    Call enableDisplayAlerts(False)
    If isCOMAddin Then
        Application.COMAddIns(addinName).Connect = enabled
    Else
        Application.AddIns(addinName).Installed = enabled
    End If
    Call enableDisplayAlerts(True)
End Sub

Public Sub enableControl(ByVal enabled As Boolean)
    With Application
        .Interactive = enabled
        .ScreenUpdating = enabled
        .EnableEvents = enabled
        .Calculation = IIf(enabled, xlCalculationAutomatic, xlCalculationManual)
        .Cursor = IIf(enabled, xlDefault, xlWait)
    End With
End Sub

Public Sub enableDisplayAlerts(ByVal enabled As Boolean)
    With Application
        .DisplayAlerts = enabled
        .AskToUpdateLinks = enabled
    End With
End Sub

Public Sub enableCutCopyMode(ByVal enabled As Boolean)
    Application.CutCopyMode = enabled
End Sub

Public Function getEOM(ByVal baseDate As Date, Optional ByVal n As Integer = 0) As Date
    getEOM = DateSerial(Year(baseDate), Month(baseDate) + n, 0)
End Function

Public Function hasRangeName(ByVal rangeName As String, ByRef wkBook As Workbook) As Boolean
    hasRangeName = False
    Dim iter As Variant
    For Each iter In wkBook.Names
        If iter.name = rangeName Then
            hasRangeName = True
            Exit For
        End If
    Next iter
    Set iter = Nothing
End Function

Public Function hasSheet(ByVal shName As String, ByRef wkBook As Workbook) As Boolean
    hasSheet = False
    Dim ws As Variant
    For Each ws In wkBook.Worksheets
        If ws.name = shName Then
            hasSheet = True
            Exit For
        End If
    Next ws
    Set ws = Nothing
End Function

Public Function hasBook(ByVal bookName As String) As Boolean
    hasBook = False
    Dim wb As Workbook
    For Each wb In Workbooks
        If wb.name = bookName Then
            hasBook = True
            Exit For
        End If
    Next wb
    Set wb = Nothing
End Function

Public Function hasBookAndSheet(ByVal shName As String, ByVal bookName As String) As Boolean
    hasBookAndSheet = False
    If Not hasBook(bookName) Then
        Exit Function
    End If
    
    If Not hasSheet(shName, Workbooks(bookName)) Then
        Exit Function
    End If
    hasBookAndSheet = True
End Function

Public Sub createSheet(ByVal shName As String, ByRef wkBook As Workbook)
    If hasSheet(shName, wkBook) Then
        Exit Sub
    End If
    
    With wkBook
        Call .Worksheets.Add(After:=.Worksheets(.Worksheets.Count))
        .Worksheets(.Worksheets.Count).name = shName
    End With
End Sub

Public Sub deleteSheet(ByVal shName As String, ByRef wkBook As Workbook)
    If Not hasSheet(shName, wkBook) Then
        Exit Sub
    End If
    
    With wkBook
        Call enableDisplayAlerts(False)
        Call wkBook.Worksheets(shName).delete
        Call enableDisplayAlerts(True)
    End With
End Sub

Public Sub formatSheets(Optional ByVal zoomRatio As Integer = 70)
    Dim ws As Variant
    For Each ws In ThisWorkbook.Worksheets
        With ws
            If ws.visible Then
                Call ws.Select
                Call ws.Range(CURSOR_POSI_DEFAULT).Select
                ActiveWindow.Zoom = zoomRatio
            End If
        End With
    Next ws
    Set ws = Nothing
    
    If hasSheet(ST_MAIN, ThisWorkbook) Then
        With ThisWorkbook.Worksheets(ST_MAIN)
            Call .Select
        End With
    End If
End Sub

Public Function getMaxRow(ByRef wkSheet As Worksheet, ByVal wsRow As Long, ByVal wsCol As Long) As Long
    With wkSheet
        If .Cells(wsRow, wsCol).Value = vbNullString Then
            getMaxRow = 0
        Else
            getMaxRow = .Cells(wsRow, wsCol).End(xlDown).row
        End If
    End With
End Function

Public Function getMaxCol(ByRef wkSheet As Worksheet, ByVal wsRow As Long, ByVal wsCol As Long) As Long
    With wkSheet
        If .Cells(wsRow, wsCol).Value = vbNullString Then
            getMaxCol = 0
        Else
            getMaxCol = .Cells(wsRow, wsCol).End(xlToRight).Column
        End If
    End With
End Function

Public Function getMaxLine(ByRef wkSheet As Worksheet, ByVal wsRow As Long, ByVal wsCol As Long, _
                           ByVal rowcol As String, ByVal targetIndex As String) As Long
    With wkSheet
        Select Case Trim(UCase(rowcol))
            Case "ROW"
                Do While .Cells(wsRow, wsCol).Value <> targetIndex
                    wsRow = wsRow + 1
                Loop
                getMaxLine = wsRow
            Case "COL"
                Do While .Cells(wsRow, wsCol).Value <> targetIndex
                    wsCol = wsCol + 1
                Loop
                getMaxLine = wsCol
            Case Else
                getMaxLine = 0
        End Select
    End With
End Function

Public Function containsObjInCol(ByVal key As String, ByRef coll As Collection) As Boolean
    Dim obj As Object
On Error Resume Next
    Set obj = coll(key)
    If Err Then
        containsObjInCol = False
    Else
        containsObjInCol = True
    End If
    Set obj = Nothing
End Function

Public Function containsVarInCol(ByVal key As String, ByRef coll As Collection) As Boolean
    Dim var As Variant
On Error Resume Next
    var = coll(key)
    If Err Then
        containsVarInCol = False
    Else
        containsVarInCol = True
    End If
    Set var = Nothing
End Function

Public Function containsVarInRecord(ByVal key As String, ByRef record As Object) As Boolean
    Dim var As Variant
On Error Resume Next
    var = record(key)
    If Err Then
        containsVarInRecord = False
    Else
        containsVarInRecord = True
    End If
    Set var = Nothing
End Function

Public Function compactSQL(ByVal sql As String) As String
    compactSQL = sql
    compactSQL = Replace(compactSQL, " + ", "+")
    compactSQL = Replace(compactSQL, " - ", "-")
    compactSQL = Replace(compactSQL, " * ", "*")
    compactSQL = Replace(compactSQL, " / ", "/")
    compactSQL = Replace(compactSQL, " = ", "=")
    compactSQL = Replace(compactSQL, " < ", "<")
    compactSQL = Replace(compactSQL, " > ", ">")
    compactSQL = Replace(compactSQL, " <> ", "<>")
    
    compactSQL = Replace(compactSQL, " ", "")
    compactSQL = Replace(compactSQL, ", ", ",")
    
'    compactSQL = Replace(compactSQL, "," & vbCrLf, ",")
'    compactSQL = Replace(compactSQL, "," & vbCr, ",")
'    compactSQL = Replace(compactSQL, "," & vbLf, ",")
End Function

Public Function getExtention(ByVal FileName As String) As String
    getExtention = IIf(InStr(FileName, DELIM_EXT), LCase(Mid(FileName, InStrRev(FileName, DELIM_EXT) + 1)), vbNullString)
End Function

Public Function getLastName(ByVal target As String, ByVal delim As String) As String
    getLastName = IIf(InStr(target, delim), LCase(Mid(target, InStrRev(target, delim) + 1)), vbNullString)
End Function

Public Function getFileColl(ByVal dirName As String, Optional ByVal recursive As Boolean = False) As Collection
    Dim buff As Variant
    Dim iter As Variant
    Dim stackColl As Collection
    Set getFileColl = New Collection
    If Not existDir(dirName) Then
        Exit Function
    End If
    
    buff = Dir(getFullPath(dirName, "*.*"))
    Do While buff <> vbNullString
        Call getFileColl.Add(getFullPath(dirName, buff))
        buff = Dir()
    Loop
    
    If Not recursive Then
        Exit Function
    End If
    
    With CreateObject(OBJ_FSO)
        Set stackColl = New Collection
        For Each buff In .GetFolder(dirName).subFolders
            Set stackColl = getFileColl(buff.Path, recursive)
            For Each iter In stackColl
                Call getFileColl.Add(iter)
            Next iter
        Next buff
    End With
    
    Set buff = Nothing
    Set iter = Nothing
End Function

Public Function getFileName(ByVal fullPath As String, Optional excludeExtention As Boolean = False) As String
    Dim i As Integer
    i = InStrRev(fullPath, DELIM_PATH_WIN, -1, vbTextCompare)
    getFileName = IIf(i = 0, fullPath, Mid(fullPath, i + 1))
    
    If excludeExtention Then
        i = InStrRev(getFileName, DELIM_EXT, -1, vbTextCompare)
        If i > 0 Then
            getFileName = left(getFileName, i - 1)
        End If
    End If
End Function

Public Function getDirName(ByVal fullPath As String) As String
    Dim i As Integer
    i = InStrRev(fullPath, DELIM_PATH_WIN)
    getDirName = left(fullPath, i)
End Function


Public Function parseURL(ByVal url As String) As String
    parseURL = IIf(InStr(url, DELIM_PATH_LIN), LCase(Mid(url, InStrRev(url, DELIM_PATH_LIN) + 1)), vbNullString)
End Function

Public Function date2str(ByVal dtDate As Date) As String
    date2str = Format(dtDate, FMT_DATE_YYYYMMDD)
End Function

Public Function str2date(ByVal strDate As String) As Date
    strDate = Replace(Trim(strDate), DELIM_DATE, vbNullString)
    If Len(strDate) = 8 Then
        str2date = CDate(Format(strDate, FMT_DATE_YYYYMMDD_SEP_NUM))
    Else
        str2date NULL_DATE
    End If
End Function

Public Sub deleteShapes(ByVal shName As String, ByRef wkBook As Workbook)
    If Not hasSheet(shName, wkBook) Then
        Exit Sub
    End If
    
    With wkBook.Worksheets(shName)
        Dim sp As Shape
        For Each sp In .Shapes
            Call sp.delete
        Next sp
        Set sp = Nothing
    End With
End Sub

Public Sub pasteFigure(ByVal shName As String, ByRef wkBook As Workbook, Optional deleteFlg As Boolean = True)
    Call createSheet(shName, wkBook)
    If deleteFlg Then
        Call deleteShapes(shName, wkBook)
    End If
    
    With wkBook.Worksheets(shName)
        Call .Select
        Call .Range(CURSOR_POSI_DEFAULT).Select
        Call .paste
    End With
    
    Call enableCutCopyMode(False)
End Sub

Public Sub deleteSeriesLinks(ByVal shName As String, ByRef wkBook As Workbook)
    If Not hasSheet(shName, wkBook) Then
        Exit Sub
    End If
    
    With wkBook.Worksheets(shName)
        Dim chartObj As ChartObject
        Dim chart As chart
        Dim srs As Series
        For Each chartObj In .ChartObjects
            Set chart = chartObj.chart
            For Each srs In chart.SeriesCollection
                srs.name = srs.naem
                On Error Resume Next
                srs.XValues = srs.XValues
                srs.Values = srs.Values
            Next srs
        Next chartObj
        Set chartObj = Nothing
        Set chart = Nothing
        Set srs = Nothing
    End With
End Sub

Public Sub showTasks()
    Dim iter As Variant
    Dim word As Object
    Set word = CreateObject(OBJ_WORD)
    For Each iter In word.Tasks
        Debug.Print iter.name
    Next iter
    Call word.Quit
    Set iter = Nothing
    Set word = Nothing
End Sub

Public Sub deleteTasks(ByVal taskName As String)
    Dim iter As Variant
    Dim word As Object
    Set word = CreateObject(OBJ_WORD)
    For Each iter In word.Tasks
        If iter.name Like taskName Then
            Call iter.Close
        End If
    Next iter
    Call word.Quit
    Set iter = Nothing
    Set word = Nothing
End Sub

Public Function getImageExtention(ByVal dirName As String, ByVal FileName As String) As String
    getImageExtention = vbNullString
    If Not existFile(dirName, FileName) Then
        Exit Function
    End If
    
    Dim buffer() As Byte
    Dim filePtr As Long
    filePtr = FreeFile
    Open getFullPath(dirName, FileName) For Binary As filePtr
    ReDim buffer(LOF(filePtr))
    Get filePtr, , buffer
    Close filePtr
    
    Select Case Hex(buffer(0))
        Case "FF"
            getImageExtention = EXT_JPG
        Case "89"
            getImageExtention = EXT_PNG
        Case "47"
            getImageExtention = EXT_GIF
    End Select
End Function

Public Function isNone(ByRef x As Variant) As Boolean
    isNone = False
    Select Case VarType(x)
        Case vbEmpty, vbNull
            isNone = True
        Case vbObject
            isNone = (x Is Nothing)
        Case vbString
            isNone = (Len(Trim(x)) = 0)
        Case Is >= vbArray
            On Error Resume Next
            isNone = (UBound(x) < LBound(x))
    End Select
End Function

Public Function paste(ByVal delim As String, ParamArray val() As Variant) As String
    Dim iter As Variant
    paste = vbNullString
    For Each iter In val
        paste = IIf(paste = vbNullString, iter, paste & delim & iter)
    Next iter
End Function

Public Function head(ByVal str As String, Optional ByVal n As Integer = 1) As String
    head = left(str, n)
End Function

Public Function tail(ByVal str As String, Optional ByVal n As Integer = 1) As String
    tail = Right(str, n)
End Function

Public Function headTrim(ByVal str As String, Optional n As Integer = 1) As String
    headTrim = left(str, Len(str) - n)
End Function

Public Function tailTrim(ByVal str As String, Optional n As Integer = 1) As String
    tailTrim = Right(str, Len(str) - n)
End Function

Public Sub showMessage(ByVal msg As String, _
                       Optional ByVal popup As Boolean = False, _
                       Optional ByVal msgStyle As Long)
    Application.StatusBar = msg
    If popup Then
        Call MsgBox(msg, msgStyle)
    End If
End Sub

Public Sub foregroundWindow(ByVal bookName As String)
    If hasBook(bookName) Then
        Call AppActivate(bookName)
    Else
        Call AppActivate(ThisWorkbook.name)
    End If
End Sub

Public Sub printOut(ByVal shName As String, ByRef wkBook As Workbook)
    If Not hasSheet(shName, wkBook) Then
        Exit Sub
    End If
    
    With wkBook.Worksheets(shName)
        Call .printOut
    End With
End Sub

Public Sub printOutALL(ByRef wkBook As Workbook)
    Dim ws As Variant
    For Each ws In wkBook.Worksheets
        Call printOut(ws.name, wkBook)
    Next ws
    Set ws = Nothing
End Sub

Public Sub deleteFile(ByVal dirName As String, ByVal FileName As String)
    If existFile(dirName, FileName) Then
        Call Kill(getFullPath(dirName, FileName))
    End If
End Sub

Public Sub foregroundWorkbook(ByVal bookName As String)
    If hasBook(bookName) Then
        Call AppActivate(bookName)
    Else
        Call AppActivate(ThisWorkbook.name)
    End If
End Sub

Public Sub recalcAll()
    Call Application.CalculateFullRebuild
End Sub

Public Sub unzip(ByVal dirName As String, ByVal FileName As String)
    If Not existFile(dirName, FileName) Then
        Exit Sub
    End If
    
    Dim files As Variant
    Select Case getExtention(FileName)
        Case EXT_ZIP, EXT_RAR
            With CreateObject(OBJ_SHELL)
                Set files = .Namespace(getFullPath(dirName, FileName)).Items
                Dim i As Long
                For i = 0 To files.Count - 1
                    If Not existFile(dirName, files.Item((i)).name) Then
                        Call .Namespace(getFullPath(dirName, FileName)).CopyHere(files.Item((i)), &H10)
                    End If
                Next i
                Set files = Nothing
            End With
    End Select
End Sub

Public Sub zip(ByVal dirName As String, ByVal FileName As String, Optional ByVal deleteOriginalFile As Boolean = True)
    If Not existFile(dirName, FileName) Then
        Exit Sub
    End If
    
    Dim tempStr As String
    Dim zipFileName As String
    Dim zipFullName As String
    Dim win As clsWinAPI
    Dim objFSO As Object
    Set win = New clsWinAPI
    Set objFSO = CreateObject(OBJ_FSO)
    
    tempStr = "PK" & Chr(5) & Chr(6) & String(18, 0)    ' meaningless string for initial writing
    zipFileName = Replace(FileName, getExtention(FileName), EXT_ZIP)
    zipFullName = getFullPath(dirName, zipFileName)
    If Not existFile(dirName, FileName) Then
        Call objFSO.CreateTextFile(zipFullName, True).Write(tempStr)
        Call objFSO.CreateTextFile(zipFullName, True).Close
        With CreateObject(OBJ_SHELL)
            Call .Namespace(getFullPath(zipFullName, vbNullString)).CopyHere(getFullPath(dirName, FileName), &H4 Or &H10)
            While .Namespace(getFullPath(zipFullName, vbNullString)).Items.Count < 1
                DoEvents
            Wend
            Call win.wait
        End With
    End If
    
    If deleteOriginalFile Then
        Call deleteFile(dirName, FileName)
    End If
    
    Set objFSO = Nothing
    Set win = Nothing
End Sub

Public Function getOrdinalFormat(ByVal targetDate As Date) As String
    Dim suffix As String
    Select Case Day(targetDate)
        Case 1, 21, 31
            suffix = "st"
        Case 2, 22
            suffix = "nd"
        Case 3, 23
            suffix = "rd"
        Case Else
            suffix = "th"
    End Select
    
    getOrdinalFormat = Day(targetDate) & suffix
End Function

Public Sub setOpacity(Optional ByVal opacity As Double = 1)
    Dim api As clsWinAPI
    Set api = New clsWinAPI
    If Math.Abs(opacity) <= 1 Then
        Call api.setOpacity(Math.Abs(opacity))
    End If
    Set api = Nothing
End Sub

Public Sub exportAllModules()
    Dim mgr As clsModuleManager
    Set mgr = New clsModuleManager
    Call mgr.exportAllModules
    Set mgr = Nothing
End Sub

Public Sub importModule(ByVal dirName As String, ByVal moduleName As String)
    Dim mgr As clsModuleManager
    Set mgr = New clsModuleManager
    Call mgr.importModule(dirName, moduleName)
    Set mgr = Nothing
End Sub

Public Function replaceDelimiter(ByVal line As String, _
                                 Optional ByVal fromDelim As String = DELIM_COMMA, _
                                 Optional ByVal toDelim As String = DELIM_SPACE) As String
                                 
    Dim iter As Variant
    Dim splitString As Variant
    Dim bufferedString As String
    Dim outputString As String
    
    bufferedString = vbNullString
    outputString = vbNullString
    splitString = Split(line, fromDelim)
    
    For Each iter In splitString
        If ((head(iter) = DELIM_DOUBLEQUOTE And tail(iter) = DELIM_DOUBLEQUOTE) Or _
            (head(iter) <> DELIM_DOUBLEQUOTE And tail(iter) <> DELIM_DOUBLEQUOTE)) And _
            bufferedString = vbNullString Then
            
            bufferedString = vbNullString
            outputString = IIf(outputString = vbNullString, iter, paste(fromDelim, outputString, iter))
        Else
            bufferedString = IIf(bufferedString = vbNullString, iter, paste(toDelim, bufferedString, iter))
            If (head(bufferedString) = DELIM_DOUBLEQUOTE And tail(bufferedString) = DELIM_DOUBLEQUOTE) Then
                outputString = IIf(outputString = vbNullString, bufferedString, paste(fromDelim, outputString, bufferedString))
                bufferedString = vbNullString
            End If
        End If
    Next iter
    
    Set iter = Nothing
    Set splitString = Nothing
    replaceDelimiter = outputString
End Function

Public Sub recordHelper(ByVal key As String, ByRef record As Object, ByRef var As Variant)
    If Not containsVarInRecord(key, record) Then
        Exit Sub
    End If
    
    If isNone(record(key)) Then
        Select Case record(key)
            Case vbInteger, vbLong
                var = 0
            Case vbSingle, vbDouble
                var = 0
            Case vbString
                var = vbNullString
            Case 202    ' VarWcChar
                var = vbNullString
            Case Else
                var = vbNullString
        End Select
    Else
        var = record(key)
    End If
End Sub

Public Sub doEventsMult(Optional times As Integer = 2)
    Dim i As Integer
    For i = 1 To times
        DoEvents
    Next i
End Sub

Public Function isStringMatched( _
    ByVal targetString As String, _
    ByVal pattern As String, _
    Optional ByVal globalMatch As Boolean = True, _
    Optional ByVal ignoreCase As Boolean = False) As Boolean
    
    Dim reg As Object
    Set reg = CreateObject(OBJ_REGEXP)
    With reg
        .Global = globalMatch
        .ignoreCase = ignoreCase
        .pattern = pattern
        isStringMatched = .test(targetString)
    End With
    Set reg = Nothing
End Function

Public Sub saveGraph(ByVal dirName As String, ByVal FileName As String, ByRef targetObject As ChartObject)
    Call makeDir(dirName)
    
    Dim ext As String
    Dim fullFileName As String
    ext = LCase(Trim(getExtention(FileName)))
    fullFileName = getFullPath(dirName, FileName)
    
    Call targetObject.Activate
    With targetObject.chart
        .PageSetup.Orientation = xlLandscape
        .PageSetup.LeftMargin = Application.InchesToPoints(0)
        .PageSetup.RightMargin = Application.InchesToPoints(0)
        .PageSetup.TopMargin = Application.InchesToPoints(0)
        .PageSetup.BottomMargin = Application.InchesToPoints(0)
        .PageSetup.HeaderMargin = Application.InchesToPoints(0)
        .PageSetup.FooterMargin = Application.InchesToPoints(0)
        .PageSetup.CenterHorizontally = True
        .PageSetup.CenterVertically = True
        
        Select Case ext
            Case EXT_PNG, EXT_JPG, EXT_JPEG, EXT_GIF, EXT_BMP
                Call .Export(FileName:=fullFileName)
                
            Case EXT_PDF
                Call .ExportAsFixedFormat( _
                    Type:=xlTypePDF, _
                    FileName:=fullFileName, _
                    IgnorePrintAreas:=True, _
                    OpenAfterPublish:=True, _
                    Quality:=xlQualityMinimum, _
                    IncludeDocProperties:=False)
                    
            Case EXT_XPS
                Call .ExportAsFixedFormat( _
                    Type:=xlTypeXPS, _
                    FileName:=fullFileName, _
                    IgnorePrintAreas:=True, _
                    OpenAfterPublish:=True, _
                    Quality:=xlQualityMinimum, _
                    IncludeDocProperties:=False)
            Case Else
        
        End Select
    End With
End Sub

Public Function readTextAsString(ByVal dirName As String, ByVal FileName As String) As String
    If Not existFile(dirName, FileName) Then
        readTextAsString = vbNullString
        Exit Function
    End If
    
    With CreateObject(OBJ_FSO)
        With .GetFile(getFullPath(dirName, FileName)).OpenAsTextStream
            readTextAsString = .ReadAll
            Call .Close
        End With
    End With
End Function
