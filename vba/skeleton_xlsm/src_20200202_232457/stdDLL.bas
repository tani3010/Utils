Attribute VB_Name = "stdDLL"
Option Explicit

Private Const stdName As String = "stdDLL"
Private Const DLL_NAME As String = "QuantLibDLL.dll"
' Private Const DLL_DIRECTORY As String = "C:\Users\xxx\Dropbox\workspace\cpp\QuantLibCode\QuantLibDLL\Release"
Private Const DLL_DIRECTORY As String = "C:\Users\xxx\Dropbox\workspace\cpp\QuantLibCode\QuantLibDLL\Debug"

' https://qiita.com/mmYYmmdd/items/fc1d3cce6a39771c0f36

#If VBA7 And Win64 Then
    Private Declare PtrSafe Function QLDLL_interpolate Lib "QuantLibDLL.dll" (ByRef x As Double, ByRef y As Double, ByVal target As Double, ByVal interpolationType As String, ByVal size As Long) As Double
    Private Declare PtrSafe Function QLDLL_term2date Lib "QuantLibDLL.dll" (ByVal baseDate As Long, ByVal term As String, ByVal calendar As String, ByVal delim As String, ByVal slidingRule As String) As Long
    Private Declare PtrSafe Function QLDLL_isHoliday Lib "QuantLibDLL.dll" (ByVal baseDate As Long, ByVal calendar As String, ByVal delim As String) As Long
    Private Declare PtrSafe Function QLDLL_getDayCount Lib "QuantLibDLL.dll" (ByVal d1 As Long, ByVal d2 As Long, ByVal dc As String, ByVal isYearFraction As Boolean) As Double
    Private Declare PtrSafe Function QLDLL_getNextIMMdate Lib "QuantLibDLL.dll" (ByVal baseDate As Long, ByVal isMainCycle As Boolean) As Long
    Private Declare PtrSafe Function QLDLL_getNextIMMcode Lib "QuantLibDLL.dll" (ByVal baseDate As Long, ByVal isMainCycle As Boolean) As String
    Private Declare PtrSafe Sub QLDLL_getHolidayList Lib "QuantLibDLL.dll" (ByVal beginDate As Long, ByVal endDate As Long, ByVal isIncludeWeekend As Boolean, ByVal calendar As String, ByVal delim As String, ByRef holidays As Long, ByVal sizeHolidays As Long)
#Else
    Private Declare Function QLDLL_interpolate Lib "QuantLibDLL.dll" (ByRef x As Double, ByRef y As Double, ByVal target As Double, ByVal interpolationType As String, ByVal size As Long) As Double
    Private Declare Function QLDLL_term2date Lib "QuantLibDLL.dll" (ByVal baseDate As Long, ByVal term As String, ByVal calendar As String, ByVal delim As String, ByVal slidingRule As String) As Long
    Private Declare Function QLDLL_isHoliday Lib "QuantLibDLL.dll" (ByVal baseDate As Long, ByVal calendar As String, ByVal delim As String) As Long
    Private Declare Function QLDLL_getDayCount Lib "QuantLibDLL.dll" (ByVal d1 As Long, ByVal d2 As Long, ByVal dc As String, ByVal isYearFraction As Boolean) As Double
    Private Declare Function QLDLL_getNextIMMdate Lib "QuantLibDLL.dll" (ByVal baseDate As Long, ByVal isMainCycle As Boolean) As Long
    Private Declare Function QLDLL_getNextIMMcode Lib "QuantLibDLL.dll" (ByVal baseDate As Long, ByVal isMainCycle As Boolean) As String
    Private Declare Sub QLDLL_getHolidayList Lib "QuantLibDLL.dll" (ByVal beginDate As Long, ByVal endDate As Long, ByVal isIncludeWeekend As Boolean, ByVal calendar As String, ByVal delim As String, ByRef holidays As Long, ByVal sizeHolidays As Long)
#End If

Public Sub addDLLDirectory()
    Dim api As clsWinAPI
    Set api = New clsWinAPI
    Call api.setDLLDirectory(DLL_DIRECTORY)
    Set api = Nothing
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
    Dim arraySize As Long
    arraySize = 0
    For i = LBound(x_) To UBound(x_)
        If Not isNone(xArray(i)) And Not isNone(yArray(i)) Then
            If Not IsError(xArray(i)) And Not IsError(yArray(i)) Then
                x_(i) = xArray(i)
                y_(i) = yArray(i)
                arraySize = arraySize + 1
            End If
        End If
    Next i
    
    interpolate = QLDLL_interpolate(x_(LBound(x_)), y_(LBound(x_)), target, interpolationType, arraySize)
End Function

