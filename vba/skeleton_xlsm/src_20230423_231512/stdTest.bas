Attribute VB_Name = "stdTest"
Option Explicit

Private Const stdName As String = "stdTest"

Public Sub testIE()
    Dim ie As clsInternetExplorer
    Set ie = New clsInternetExplorer
    
    With ie
        Call .navigate("https://cc.minkabu.jp/pair/BTC_JPY")
        Call .setIE("ビットコイン")
        Dim tmp As String
        Call .handleEvent(ACTION_GET_TXT, tmp, vbNullString, "script", "data-component-name", "HeaderPrice")
        Debug.Print tmp
    End With
    
    Set ie = Nothing
End Sub

Public Sub testest()
    Dim a As Collection
    Set a = New Collection
    Debug.Print VarType(a)
    Debug.Print TypeName(a)
End Sub

Public Sub testSaveCharts()
    Dim dirName As String
    Dim fileName As String

    dirName = ThisWorkbook.Path
    With ThisWorkbook.Worksheets("TEST")
        Dim iter As ChartObject
        For Each iter In .ChartObjects
            fileName = "test.pdf"
            Call saveGraph(dirName, fileName, iter)
        Next iter
    End With
        
    Set iter = Nothing
End Sub

Public Sub test_term2date()
    Dim d As Date
    d = Date
    
    Debug.Print term2date(d, "0D", "JPY+USD")
    Debug.Print term2date(d, "1D", "TKY+NYK")
    Debug.Print term2date(d, "1Y", "TKY+NYK+LDN")
    Debug.Print term2date(d, "2Y", "JPY")
    Debug.Print term2date(d, "3Y", "JPY")
    Debug.Print term2date(d, "4Y", "JPY")
    Debug.Print term2date(d, "5Y", "JPY")
    Debug.Print term2date(d, "6Y", "JPY")
    Debug.Print term2date(d, "7Y", "JPY")
    Debug.Print term2date(d, "8Y", "JPY")
    Debug.Print term2date(d, "9Y", "JPY")
    Debug.Print term2date(d, "10Y", "JPY")
    Debug.Print term2date(d, "11Y", "JPY")
    Debug.Print term2date(d, "12Y", "JPY")
End Sub

Public Sub test_cdsRollDate()
    Dim n As Integer
    n = 10
    
    Dim i As Integer
    Dim d As Date
    For i = 2010 To 2010 + n
        d = CDate(i & "/09/20")
        d = term2date(d, "0D", "USD")
        Debug.Print d
    Next i
    
End Sub

Public Sub test_getDayCount()
    Dim d1 As Date
    Dim d2 As Date
    d1 = Date
    d2 = d1 + 1234
    
    Debug.Print getDayCount(d1, d2, "act/365", True)
    Debug.Print getDayCount(d1, d2, "a/365")
    Debug.Print getDayCount(d1, d2, "act365")
    Debug.Print getDayCount(d1, d2, "act/360")
    Debug.Print getDayCount(d1, d2, "act/365F")
    Debug.Print getDayCount(d1, d2, "act/ACT")
    Debug.Print getDayCount(d1, d2, "30/360")
    Debug.Print getDayCount(d1, d2, "30360")
    Debug.Print getDayCount(d1, d2, "30e360")
    Debug.Print getDayCount(d1, d2, "act/365")
End Sub

Public Sub test_getNextIMMdate()
    Dim d As Date
    d = Date
    
    Dim i As Integer
    For i = 1 To 15
        d = getNextIMMdate(d, True)
        Debug.Print "IMM-" & i, d
    Next i
End Sub

Public Sub test_getNextIMMcode()
    Dim d As Date
    Dim code As String
    d = Date
    
    Dim i As Integer
    For i = 1 To 15
        d = getNextIMMdate(d, False)
        code = getNextIMMcode(d, False)
        Debug.Print "IMM-" & i, d, code
    Next i
End Sub

Public Sub test_isHoliday()
    Dim baseDate As Date
    Dim city As String
    
    baseDate = Date
    city = "TKY+LDN+NYK"
    
    Dim i As Integer
    For i = 1 To 100
        baseDate = baseDate + 1
        Debug.Print baseDate, city, isHoliday(baseDate, city)
    Next i
End Sub

Public Sub test_DLLdirectorySearch()
    Dim coll As Collection
    Set coll = getFileColl(ThisWorkbook.Path, True)
    
    Const dllName As String = "QuantLibDLL.dll"
    
    Dim iter As Variant
    Dim api As clsWinAPI
    Set api = New clsWinAPI
    For Each iter In coll
        If getFileName(iter) = dllName Then
            Call api.setDLLDirectory(getDirName(iter))
        End If
    Next iter
    
    Set iter = Nothing
    Set coll = Nothing
    Set api = Nothing
End Sub

Public Sub test_getHolidayList()
    Dim beginDate As Date
    Dim endDate As Date
    
    beginDate = Date
    endDate = beginDate + 15000
    
    Dim tmp() As Date
    Call getHolidayList(tmp, beginDate, endDate, False, "TKY")
End Sub

Public Sub test_createSchedule()
    Dim effectiveDate As Date
    Dim terminationDate As Date
    
    effectiveDate = Date
    terminationDate = term2date(effectiveDate, "10Y", "JPY")
    
    Dim tmp() As Date
    Call createScheduleByTerminationDate(tmp, effectiveDate, terminationDate, "6M", "JPY+USD")
    
    Dim i As Integer
    For i = LBound(tmp) To UBound(tmp)
        Debug.Print "CF-" & i, tmp(i)
    Next i
End Sub

Public Sub test_createScheduleByPeriod()
    Dim effectiveDate As Date
    Dim terminationDate As Date
    
    effectiveDate = Date
    
    Dim tmp() As Date
    Call createScheduleByTerminationPeriod(tmp, effectiveDate, "10Y", "6M", "JPY+USD")
    
    Dim i As Integer
    For i = LBound(tmp) To UBound(tmp)
        Debug.Print "CF-" & i, tmp(i)
    Next i
End Sub

Public Sub test_createIRSScheduleByPeriod()
    Dim effectiveDate As Date
    Dim terminationDate As Date
    
    effectiveDate = Date
    
    ' fixed dates setting
    Dim fixedStartDates() As Date
    Dim fixedEndDates() As Date
    Dim fixedFixingDates() As Date
    Dim fixedPaymentDates() As Date
    
    ' floating dates setting
    Dim floatingStartDates() As Date
    Dim floatingEndDates() As Date
    Dim floatingFixingDates() As Date
    Dim floatingPaymentDates() As Date
    
    Call createIRSScheduleByTerminationPeriod( _
        fixedStartDates, fixedEndDates, fixedFixingDates, fixedPaymentDates, _
        floatingStartDates, floatingEndDates, floatingFixingDates, floatingPaymentDates, _
        effectiveDate, "10Y", "6M", "JPY+USD")
    
    Dim i As Integer
    Debug.Print "Fixed Leg"
    Debug.Print "CF-No.", "StartDate", "EndDate", "FixingDate", "PaymentDate"
    For i = LBound(fixedStartDates) To UBound(fixedStartDates)
        Debug.Print "CF-" & i, fixedStartDates(i), fixedEndDates(i), fixedFixingDates(i) ' , fixedPaymentDates(i)
    Next i

    Debug.Print "floating Leg"
    Debug.Print "CF-No.", "StartDate", "EndDate", "FixingDate", "PaymentDate"
    For i = LBound(floatingStartDates) To UBound(floatingStartDates)
        Debug.Print "CF-" & i, floatingStartDates(i), floatingEndDates(i), floatingFixingDates(i) ' , floatingPaymentDates(i)
    Next i
End Sub

Public Sub test_sortTerms()
    Dim terms(0 To 10) As String
    terms(0) = "1.5Y"
    terms(1) = "1Y"
    terms(2) = "1M"
    terms(3) = "2D"
    terms(4) = "5Y"
    terms(5) = "TOD"
    terms(6) = "ON"
    terms(7) = "2Y"
    terms(8) = "2Y3M"
    terms(9) = "4Y11M"
    terms(10) = "1W"
    
    Call sortTerms(terms)
    
    Dim i As Integer
    For i = LBound(terms) To UBound(terms)
        Debug.Print i, terms(i)
    Next i
End Sub

Public Sub test_buildSetterGetter()
    Dim valName As Variant
    Dim valType As Variant
    Dim valRef As Variant
    
    valName = Array("x1", "x2", "x3", "x4", "y1", "y2")
    valType = Array("string", "double", "integer", "variant", "long", "double")
    valRef = Array(True, True, False, True, False, False)
    
    Dim i As Integer
    For i = LBound(valName) To UBound(valName)
        Call buildSetterGetter(valName(i), valType(i), valRef(i))
    Next i
End Sub


Public Sub testDB()
    Dim db As clsDBAccess
    Set db = New clsDBAccess
    
    Dim sql As String
    With ThisWorkbook.Worksheets(ST_DBCONFIG)
        db.set_driver = .Range("C2").Value
        db.set_dbQuery = .Range("C3").Value
        sql = .Range("B6").Value
    End With
    Call db.executeSQL(sql, batchPaste:=False)
    Set db = Nothing
End Sub
