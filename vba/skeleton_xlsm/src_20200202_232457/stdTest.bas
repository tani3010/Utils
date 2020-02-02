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
        
    End With
    
    Set ie = Nothing
End Sub

Public Sub testSaveCharts()
    Dim dirName As String
    Dim FileName As String

    dirName = ThisWorkbook.Path
    With ThisWorkbook.Worksheets("TEST")
        Dim iter As ChartObject
        For Each iter In .ChartObjects
            FileName = "test.pdf"
            Call saveGraph(dirName, FileName, iter)
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

Public Sub test_getDayCount()
    Dim d1 As Date
    Dim d2 As Date
    d1 = Date
    d2 = d1 + 1234
    
    Debug.Print getDayCount(d1, d2, "act/365")
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
    For Each iter In coll
        If getFileName(iter) = dllName Then
            ' Call addDLLDirectory(getDirName(iter))
        End If
    Next iter
    
    Set iter = Nothing
    Set coll = Nothing
End Sub

Public Sub test_getHolidayList()
    Dim beginDate As Date
    Dim endDate As Date
    
    beginDate = Date
    endDate = beginDate + 15000
    
    Dim tmp() As Date
    Call getHolidayList(tmp, beginDate, endDate, False, "TKY")
End Sub

Public Sub outputHolidays()
    If Not hasSheet(ST_HOLIDAY, ThisWorkbook) Then
        Call createSheet(ST_HOLIDAY, ThisWorkbook)
    End If
    
    Call ThisWorkbook.Worksheets(ST_HOLIDAY).Cells.Clear
    
    Dim cities As Variant
    Dim beginDate As Date
    Dim endDate As Date
    beginDate = #1/1/2010#
    endDate = beginDate + 16000
    cities = Array("TKY", "LDN", "NYK", "EUR")
    
    Dim i As Integer
    Dim j As Integer
    Dim holidays() As Date
    For i = LBound(cities) To UBound(cities)
        Call getHolidayList(holidays, beginDate, endDate, False, cities(i))
        With ThisWorkbook.Worksheets(ST_HOLIDAY)
            .Cells(1, i + 1).Value = cities(i)
            For j = LBound(holidays) To UBound(holidays)
                .Cells(j + 2, i + 1).Value = holidays(j)
            Next j
        End With
    Next i
End Sub
