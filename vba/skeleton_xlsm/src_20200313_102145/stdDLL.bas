Attribute VB_Name = "stdDLL"
Option Explicit

Private Const stdName As String = "stdDLL"
Private Const DLL_NAME As String = "QuantLibDLL.dll"
Private Const DLL_DIRECTORY As String = "C:\Users\tani\Dropbox\workspace\cpp\QuantLibCode\QuantLibDLL\Release"
' Private Const DLL_DIRECTORY As String = "C:\Users\tani\Dropbox\workspace\cpp\QuantLibCode\QuantLibDLL\Debug"
' Private Const DLL_DIRECTORY As String = "C:\Users\tani3010\Dropbox\workspace\cpp\QuantLibCode\QuantLibDLL\Debug"
' Private Const DLL_DIRECTORY As String = "C:\Users\tani3010\Dropbox\workspace\cpp\QuantLibCode\QuantLibDLL\Release"

' https://qiita.com/mmYYmmdd/items/fc1d3cce6a39771c0f36

#If VBA7 And Win64 Then
    Private Declare PtrSafe Function QLDLL_interpolate Lib "QuantLibDLL.dll" (ByRef x As Double, ByRef y As Double, ByVal target As Double, ByVal interpolationType As String, ByVal size As Long) As Double
    Private Declare PtrSafe Sub QLDLL_interpolate1D Lib "QuantLibDLL.dll" (ByRef x As Double, ByRef y As Double, ByRef target As Double, ByRef output As Double, ByVal interpolationType As String, ByVal size As Long, ByVal targetSize As Long)
    Private Declare PtrSafe Function QLDLL_term2date Lib "QuantLibDLL.dll" (ByVal baseDate As Long, ByVal term As String, ByVal calendar As String, ByVal delim As String, ByVal slidingRule As String) As Long
    Private Declare PtrSafe Function QLDLL_isHoliday Lib "QuantLibDLL.dll" (ByVal baseDate As Long, ByVal calendar As String, ByVal delim As String) As Long
    Private Declare PtrSafe Function QLDLL_getDayCount Lib "QuantLibDLL.dll" (ByVal d1 As Long, ByVal d2 As Long, ByVal dc As String, ByVal isYearFraction As Boolean) As Double
    Private Declare PtrSafe Function QLDLL_getNextIMMdate Lib "QuantLibDLL.dll" (ByVal baseDate As Long, ByVal isMainCycle As Boolean) As Long
    Private Declare PtrSafe Function QLDLL_getNextIMMcode Lib "QuantLibDLL.dll" (ByVal baseDate As Long, ByVal isMainCycle As Boolean) As String
    Private Declare PtrSafe Sub QLDLL_getHolidayList Lib "QuantLibDLL.dll" (ByVal beginDate As Long, ByVal endDate As Long, ByVal isIncludeWeekend As Boolean, ByVal calendar As String, ByVal delim As String, ByRef holidays As Long, ByVal sizeHolidays As Long)
    Private Declare PtrSafe Sub QLDLL_createScheduleByTerminationDate Lib "QuantLibDLL.dll" (ByRef outputDates As Long, ByVal outputDatesSize As Long, ByVal effectiveDate As Long, ByVal terminationDate As Long, ByVal tenor As String, ByVal calendar As String, ByVal delim As String, ByVal convention As String, ByVal terminationDateConvention As String, ByVal dateGeneration As String, ByVal endOfMonth As Boolean, ByVal firstDate As Long, ByVal nextToLastDate As Long)
    Private Declare PtrSafe Sub QLDLL_createScheduleByTerminationPeriod Lib "QuantLibDLL.dll" (ByRef outputDates As Long, ByVal outputDatesSize As Long, ByVal effectiveDate As Long, ByVal terminationPeriod As String, ByVal tenor As String, ByVal calendar As String, ByVal delim As String, ByVal convention As String, ByVal terminationDateConvention As String, ByVal dateGeneration As String, ByVal endOfMonth As Boolean, ByVal firstDate As Long, ByVal nextToLastDate As Long)
    Private Declare PtrSafe Sub QLDLL_sortTerms Lib "QuantLibDLL.dll" (ByRef terms() As String)
#Else
    Private Declare Function QLDLL_interpolate Lib "QuantLibDLL.dll" (ByRef x As Double, ByRef y As Double, ByVal target As Double, ByVal interpolationType As String, ByVal size As Long) As Double
    Private Declare Sub QLDLL_interpolate1D Lib "QuantLibDLL.dll" (ByRef x As Double, ByRef y As Double, ByRef target As Double, ByRef output As Double, ByVal interpolationType As String, ByVal size As Long, ByVal targetSize As Long)
    Private Declare Function QLDLL_term2date Lib "QuantLibDLL.dll" (ByVal baseDate As Long, ByVal term As String, ByVal calendar As String, ByVal delim As String, ByVal slidingRule As String) As Long
    Private Declare Function QLDLL_isHoliday Lib "QuantLibDLL.dll" (ByVal baseDate As Long, ByVal calendar As String, ByVal delim As String) As Long
    Private Declare Function QLDLL_getDayCount Lib "QuantLibDLL.dll" (ByVal d1 As Long, ByVal d2 As Long, ByVal dc As String, ByVal isYearFraction As Boolean) As Double
    Private Declare Function QLDLL_getNextIMMdate Lib "QuantLibDLL.dll" (ByVal baseDate As Long, ByVal isMainCycle As Boolean) As Long
    Private Declare Function QLDLL_getNextIMMcode Lib "QuantLibDLL.dll" (ByVal baseDate As Long, ByVal isMainCycle As Boolean) As String
    Private Declare Sub QLDLL_getHolidayList Lib "QuantLibDLL.dll" (ByVal beginDate As Long, ByVal endDate As Long, ByVal isIncludeWeekend As Boolean, ByVal calendar As String, ByVal delim As String, ByRef holidays As Long, ByVal sizeHolidays As Long)
    Private Declare Sub QLDLL_createScheduleByTerminationDate Lib "QuantLibDLL.dll" (ByRef outputDates As Long, ByVal outputDatesSize As Long, ByVal effectiveDate As Long, ByVal terminationDate As Long, ByVal tenor As String, ByVal calendar As String, ByVal delim As String, ByVal convention As String, ByVal terminationDateConvention As String, ByVal dateGeneration As String, ByVal endOfMonth As Boolean, ByVal firstDate As Long, ByVal nextToLastDate As Long)
    Private Declare Sub QLDLL_createScheduleByTerminationPeriod Lib "QuantLibDLL.dll" (ByRef outputDates As Long, ByVal outputDatesSize As Long, ByVal effectiveDate As Long, ByVal terminationPeriod As String, ByVal tenor As String, ByVal calendar As String, ByVal delim As String, ByVal convention As String, ByVal terminationDateConvention As String, ByVal dateGeneration As String, ByVal endOfMonth As Boolean, ByVal firstDate As Long, ByVal nextToLastDate As Long)
    Private Declare Sub QLDLL_sortTerms Lib "QuantLibDLL.dll" (ByRef terms() As String)
#End If

Public Sub addDLLDirectory()
    Dim api As clsWinAPI
    Set api = New clsWinAPI
    Call api.setDLLDirectory(DLL_DIRECTORY)
    Set api = Nothing
End Sub

Public Sub sortTerms(ByRef terms() As String)
    Call QLDLL_sortTerms(terms)
End Sub

Public Function term2date(ByVal targetDate As Date, ByVal term As String, ByVal calendar As String, _
                          Optional ByVal delim As String = "+", _
                          Optional ByVal slidingRule As String = "MODIFIEDFOLLOWING") As Date
    term2date = CDate(QLDLL_term2date(CLng(targetDate), term, calendar, delim, slidingRule))
End Function

Public Function isHoliday(ByVal targetDate As Date, ByVal calendar As String, _
                          Optional ByVal delim As String = "+") As Boolean
    isHoliday = CBool(QLDLL_isHoliday(CLng(targetDate), calendar, delim))
End Function

Public Sub getHolidayList( _
    ByRef outputHolidays() As Date, _
    ByVal beginDate As Date, ByVal endDate As Date, ByVal isIncludeWeekend As Boolean, _
    ByVal calendar As String, Optional ByVal delim As String = "+", Optional ByVal buffSize As Long = 5000)
    
    Dim holidays() As Long
    ReDim holidays(0 To buffSize) As Long
    
    Call QLDLL_getHolidayList( _
        CLng(beginDate), CLng(endDate), isIncludeWeekend, calendar, delim, holidays(LBound(holidays)), buffSize)
    
    Dim i As Long
    ReDim outputHolidays(0 To buffSize) As Date
    For i = 0 To buffSize
        If holidays(i) = 0 Then
            Exit For
        Else
            outputHolidays(i) = CDate(holidays(i))
        End If
    Next i
    ReDim Preserve outputHolidays(0 To i - 1) As Date
End Sub

Public Function getDayCount(ByVal d1 As Long, ByVal d2 As Long, ByVal dc As String, _
                            Optional ByVal isYearFraction As Boolean = False) As Double
    getDayCount = QLDLL_getDayCount(d1, d2, dc, isYearFraction)
End Function

Public Function getNextIMMdate(ByVal baseDate As Long, Optional ByVal isMainCycle As Boolean = True) As Date
    getNextIMMdate = CDate(QLDLL_getNextIMMdate(baseDate, isMainCycle))
End Function

Public Function getNextIMMcode(ByVal baseDate As Long, Optional ByVal isMainCycle As Boolean = True) As String
    getNextIMMcode = QLDLL_getNextIMMcode(baseDate, isMainCycle)
End Function

Public Function interpolate(ByRef x As Range, ByRef y As Range, ByVal target As Double, ByVal interpolationType As String) As Double
    Dim xArray As Variant: xArray = WorksheetFunction.Transpose(x)
    Dim yArray As Variant: yArray = WorksheetFunction.Transpose(y)
    
    Dim x_() As Double
    Dim y_() As Double
    ReDim x_(LBound(xArray) To UBound(xArray))
    ReDim y_(LBound(yArray) To UBound(yArray))
    Dim i As Integer
    Dim j As Integer
    Dim arraySize As Long
    arraySize = 0
    j = LBound(x_)
    For i = LBound(x_) To UBound(x_)
        If IsNumeric(xArray(i)) And IsNumeric(yArray(i)) Then
            If Not isNone(xArray(i)) And Not isNone(yArray(i)) Then
                If Not IsError(xArray(i)) And Not IsError(yArray(i)) Then
                    x_(j) = xArray(i)
                    y_(j) = yArray(i)
                    arraySize = arraySize + 1
                    j = j + 1
                End If
            End If
        End If
    Next i
    
    ReDim Preserve x_(LBound(x_) To arraySize) As Double
    ReDim Preserve y_(LBound(x_) To arraySize) As Double
    interpolate = QLDLL_interpolate(x_(LBound(x_)), y_(LBound(x_)), target, interpolationType, arraySize)
End Function

Public Function interpolate1D(ByRef x As Range, ByRef y As Range, ByVal target As Range, ByVal interpolationType As String) As Variant
    Dim xArray As Variant: xArray = WorksheetFunction.Transpose(x)
    Dim yArray As Variant: yArray = WorksheetFunction.Transpose(y)
    Dim targetArray As Variant: targetArray = WorksheetFunction.Transpose(target)
    
    Dim x_() As Double
    Dim y_() As Double
    Dim target_() As Double
    Dim output_() As Double
    ReDim x_(LBound(xArray) To UBound(xArray))
    ReDim y_(LBound(yArray) To UBound(yArray))
    ReDim target_(LBound(targetArray) To UBound(targetArray))
    ReDim output_(LBound(targetArray) To UBound(targetArray))
    Dim i As Integer
    Dim j As Integer
    Dim arraySize As Long
    Dim targetSize As Long
    arraySize = 0
    j = LBound(x_)
    For i = LBound(x_) To UBound(x_)
        If IsNumeric(xArray(i)) And IsNumeric(yArray(i)) Then
            If Not isNone(xArray(i)) And Not isNone(yArray(i)) Then
                If Not IsError(xArray(i)) And Not IsError(yArray(i)) Then
                    x_(j) = xArray(i)
                    y_(j) = yArray(i)
                    arraySize = arraySize + 1
                    j = j + 1
                End If
            End If
        End If
    Next i
    
    ReDim Preserve x_(LBound(x_) To arraySize) As Double
    ReDim Preserve y_(LBound(x_) To arraySize) As Double
    
    targetSize = 0
    For i = LBound(target_) To UBound(target_)
        If Not isNone(targetArray(i)) Then
            If Not IsError(targetArray(i)) Then
                target_(i) = targetArray(i)
                targetSize = targetSize + 1
            End If
        End If
    Next i
    
    Call QLDLL_interpolate1D( _
        x_(LBound(x_)), y_(LBound(x_)), target_(LBound(target_)), output_(LBound(target_)), _
        interpolationType, arraySize, targetSize)
    interpolate1D = WorksheetFunction.Transpose(output_)
End Function

Public Sub outputHolidays()
    If Not hasSheet(ST_HOLIDAY, ThisWorkbook) Then
        Call createSheet(ST_HOLIDAY, ThisWorkbook)
    End If
    
    Call enableControl(False)
    Call ThisWorkbook.Worksheets(ST_HOLIDAY).Cells.Clear
    
    Dim cities As Variant
    Dim beginDate As Date
    Dim endDate As Date
    beginDate = Date
    endDate = term2date(beginDate, "60Y", "JPN")
    cities = Array("JPN", "USA", "GBR", "EUR", "AUS", "CAN", "CHN", "HK", "NZ", "SGP")
    
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
    
    Call enableControl(True)
End Sub

Public Sub createScheduleByTerminationDate( _
    ByRef outputDates() As Date, ByVal effectiveDate As Date, ByVal terminationDate As Date, _
    ByVal tenor As String, ByVal calendar As String, _
    Optional ByVal delim As String = "+", _
    Optional ByVal convention As String = "MODIFIEDFOLLOWING", _
    Optional ByVal terminationDateConvention As String = "MODIFIEDFOLLOWING", _
    Optional ByVal dateGeneration As String = "FORWARD", _
    Optional ByVal endOfMonth As Boolean = False, _
    Optional ByVal firstDate As Date = NULL_DATE, _
    Optional ByVal nextToLastDate As Date = NULL_DATE, _
    Optional ByVal buffSize As Long = 1000)
    
    Dim tmpDates() As Long
    ReDim tmpDates(0 To buffSize) As Long
    
    Call QLDLL_createScheduleByTerminationDate( _
        tmpDates(LBound(tmpDates)), buffSize, CLng(effectiveDate), CLng(terminationDate), _
        tenor, calendar, delim, convention, terminationDateConvention, _
        dateGeneration, endOfMonth, _
        IIf(firstDate = NULL_DATE, 0, CLng(firstDate)), _
        IIf(nextToLastDate = NULL_DATE, 0, CLng(nextToLastDate)))
    
    Dim i As Long
    ReDim outputDates(0 To buffSize) As Date
    For i = 0 To buffSize
        If tmpDates(i) = 0 Then
            Exit For
        Else
            outputDates(i) = CDate(tmpDates(i))
        End If
    Next i
    ReDim Preserve outputDates(0 To i - 1) As Date
End Sub

Public Sub createScheduleByTerminationPeriod( _
    ByRef outputDates() As Date, ByVal effectiveDate As Date, ByVal terminationPeriod As String, _
    ByVal tenor As String, ByVal calendar As String, _
    Optional ByVal delim As String = "+", _
    Optional ByVal convention As String = "MODIFIEDFOLLOWING", _
    Optional ByVal terminationDateConvention As String = "MODIFIEDFOLLOWING", _
    Optional ByVal dateGeneration As String = "FORWARD", _
    Optional ByVal endOfMonth As Boolean = False, _
    Optional ByVal firstDate As Date = NULL_DATE, _
    Optional ByVal nextToLastDate As Date = NULL_DATE, _
    Optional ByVal buffSize As Long = 1000)
    
    Dim tmpDates() As Long
    ReDim tmpDates(0 To buffSize) As Long
    
    Call QLDLL_createScheduleByTerminationPeriod( _
        tmpDates(LBound(tmpDates)), buffSize, CLng(effectiveDate), terminationPeriod, _
        tenor, calendar, delim, convention, terminationDateConvention, _
        dateGeneration, endOfMonth, _
        IIf(firstDate = NULL_DATE, 0, CLng(firstDate)), _
        IIf(nextToLastDate = NULL_DATE, 0, CLng(nextToLastDate)))
    
    Dim i As Long
    ReDim outputDates(0 To buffSize) As Date
    For i = 0 To buffSize
        If tmpDates(i) = 0 Then
            Exit For
        Else
            outputDates(i) = CDate(tmpDates(i))
        End If
    Next i
    ReDim Preserve outputDates(0 To i - 1) As Date
End Sub

