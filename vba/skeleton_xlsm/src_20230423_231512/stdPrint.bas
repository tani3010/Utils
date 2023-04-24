Attribute VB_Name = "stdPrint"
Option Explicit

' https://qiita.com/o_pos/items/482455a06a6150516dd7

'以下の定数は全て変更可能 (他の部分の書き換えが必要ない)
Const ARRAY_SEGMENT = ", "       ' 配列等の要素の区切り文字
Const MAX_ARR_ELEMENTS = 100     ' 配列の要素の表示数の上限
Const MAX_ARR_TRACKING = 20      ' 配列のネスト(配列の次元とは違う)の表示の上限
Const MAX_COLLECTION_SIZE = 100  ' コレクションの要素の表示数の上限
Const MAX_DICTIONARY_SIZE = 100  ' dictionaryの要素の表示数の上限
Const WRITE_EMPTY = True         ' Emptyを表示するか (Falseなら、たとえば[Empty]が[]になる)

'配列の型を取得
Function getArrayType(arr) As String '(arr)を(arr())にするとValiant以外の型の配列を受け取れなくなる
  Dim s As String: s = TypeName(arr)
  Debug.Assert s Like "*(*)" '配列であることを確認
  getArrayType = left(s, Len(s) - 2)
End Function

Function p(src, Optional debugPrint As Boolean = True, Optional recursionCount_ As Long = 0) As String
  Dim result As String, srcType As String
  result = ""
  srcType = TypeName(src)
  Select Case srcType
  Case "String" '文字列
    result = """" & Replace(src, """", """""") & """"
  Case "Integer", "Double", "Single", "Long", "Currency", "Byte" '数値
    result = src
  Case "Decimal" '十進数
    result = src & " As Decimal"
  Case "Date" '日付
    result = """" & src & """ As Date"
  Case "Boolean" 'ブール値
    result = src
  Case "Collection" 'コレクション
    processCollection result, src, "Collection", recursionCount_
  Case "Dictionary" 'Scripting.Dictionary
    processDictionary result, src, recursionCount_
  Case Else
    If srcType Like "*(*)" Then '配列
      processArr result, src, srcType, recursionCount_
    ElseIf isShortCollection(src) Then 'コレクション的オブジェクト
      processCollection result, src, srcType, recursionCount_
    Else 'その他
      result = srcType
    End If
  End Select
  If recursionCount_ > 0 And result = "Empty" And Not WRITE_EMPTY Then result = ""
  If debugPrint Then Debug.Print result
  p = result
End Function

'配列の次元を取得 (Dim a() のaなら0)
Function getArrayDim(arr, Optional dimIter_ As Long = 1) As Long
  On Error GoTo onerror
  Dim upper As Long
  upper = UBound(arr, dimIter_)
  If upper = -1 Then
    'Function UBound2(a, b): UBound2 = UBound(a, b): End Function があるときに、
    'Dim arr() As String: UBound(arr, 1) はエラー、
    'Dim arr() As Range:  UBound(arr, 1) もエラー、
    'Dim arr() As Range:  Ubound2(arr, 1) は-1を返す
    'という現象への対処
    getArrayDim = 0
    Exit Function
  End If
  getArrayDim = upper * 0 + getArrayDim(arr, dimIter_ + 1)
  Exit Function
onerror:
  getArrayDim = dimIter_ - 1
End Function

Function isCollection(c) As Boolean
  On Error GoTo onerror
  c.Count
  isCollection = True
  Exit Function
onerror:
  isCollection = False
End Function

'Collectionかつ長さが一定以下ならtrue
Private Function isShortCollection(c) As Boolean
  On Error GoTo onerror
  isShortCollection = c.Count <= MAX_COLLECTION_SIZE
  Exit Function
onerror:
  isShortCollection = False
End Function

'配列の場合の処理
Private Sub processArr(result As String, src, srcType As String, recursionCount_ As Long)
  Dim arrType As String, arrdim As Long
  Dim i As Long, j As Long, k As Long
  arrType = left(srcType, Len(srcType) - 2)
  arrdim = getArrayDim(src)
  '例外処理
  If recursionCount_ > MAX_ARR_TRACKING Then
    result = arrType & "[...]": Exit Sub
  ElseIf arrdim = 0 Then
    result = "(0dim)[] As " & arrType: Exit Sub
  ElseIf arrdim > 3 Then
    'この3は多次元配列の展開をする上限の次元。
    '変更する場合は下のselect case部分の書き換えが必要。
    result = templateStr("($dim)[...] As $", arrdim, arrType): Exit Sub
  End If
  '配列を展開
  If LBound(src) <> 0 Then concat result, templateStr("(from$)", LBound(src))
  startArr result
  '3次元までしか表示しないので、適当に書いた
  Select Case arrdim
  Case 1
    For i = LBound(src) To min(UBound(src), LBound(src) + MAX_ARR_ELEMENTS - 1)
      pushArrMem result, src(i), recursionCount_
    Next
    If UBound(src) >= LBound(src) + MAX_ARR_ELEMENTS Then pushEllipsis result
  Case 2
    For i = LBound(src) To min(UBound(src), LBound(src) + MAX_ARR_ELEMENTS - 1)
      startArr result
      For j = LBound(src, 2) To min(UBound(src, 2), LBound(src, 2) + MAX_ARR_ELEMENTS - 1)
        pushArrMem result, src(i, j), recursionCount_
      Next
      If UBound(src, 2) >= LBound(src, 2) + MAX_ARR_ELEMENTS Then pushEllipsis result
      endArr result
    Next
    If UBound(src) >= LBound(src) + MAX_ARR_ELEMENTS Then pushEllipsis result
  Case 3
    For i = LBound(src) To min(UBound(src), LBound(src) + MAX_ARR_ELEMENTS - 1)
      startArr result
      For j = LBound(src, 2) To min(UBound(src, 2), LBound(src, 2) + MAX_ARR_ELEMENTS - 1)
        startArr result
        For k = LBound(src, 3) To min(UBound(src, 3), LBound(src, 3) + MAX_ARR_ELEMENTS - 1)
          pushArrMem result, src(i, j, k), recursionCount_
        Next
        If UBound(src, 3) >= LBound(src, 3) + MAX_ARR_ELEMENTS Then pushEllipsis result
        endArr result
      Next
      If UBound(src, 2) >= LBound(src, 2) + MAX_ARR_ELEMENTS Then pushEllipsis result
      endArr result
    Next
    If UBound(src) >= LBound(src) + MAX_ARR_ELEMENTS Then pushEllipsis result
  End Select
  endArr result, True
  concat result, " As " & arrType
End Sub

'WorksheetFunctionが無い環境用
Private Function min(a, b)
  Debug.Assert IsNumeric(a) And IsNumeric(b)
  If a < b Then
    min = a
  Else
    min = b
  End If
End Function

'WorksheetFunctionが無い環境用
Private Function max(a, b)
  Debug.Assert IsNumeric(a) And IsNumeric(b)
  If a < b Then
    max = b
  Else
    max = a
  End If
End Function

'コレクションの添え字の最小値(0か1以外から始まる場合には非対応, 空のCollectionなら常に1を返す)
Private Function collectionLBound(c)
  On Error GoTo onerror
  Dim tmp: tmp = c(0)
  collectionLBound = 0
  Exit Function
onerror:
  collectionLBound = 1
End Function

'コレクションの場合の処理
Private Sub processCollection(result As String, src, srcType As String, recursionCount_ As Long)
  result = "Collection"
  If srcType <> "Collection" Then concat result, templateStr("($)", srcType)
  Dim lower As Long: lower = collectionLBound(src)
  If lower <> 1 Then concat result, templateStr("(from$)", lower)
  startArr result
  If src.Count = 0 Then
    endArr result, True, True
  Else
    'regexpのexecuteメソッドの返り値など、コレクション風だけど
    '添え字が0から始まるオブジェクトもあるから、For i = 1 To min(src.Count, MAX_COLLECTION_SIZE) では不可
    Dim el, i As Long: i = 0
    For Each el In src
      If i >= MAX_COLLECTION_SIZE Then
        pushEllipsis result
        Exit For
      End If
      pushArrMem result, el, recursionCount_
      i = i + 1
    Next
    endArr result, True
  End If
End Sub

'Dictionaryの場合の処理
Private Sub processDictionary(result As String, src, recursionCount_ As Long)
  startDic result
  If src.Count = 0 Then
    endDic result, True
  Else
    Dim keys(): keys = src.keys

    Dim key, i As Long: i = 0
    For Each key In keys
      If i >= MAX_DICTIONARY_SIZE Then
        pushEllipsis result
        Exit For
      End If
      pushDicMem result, key, src.Item(key), recursionCount_
      i = i + 1
    Next
    endDic result
  End If
End Sub

'配列文字列の作成
Private Sub startArr(result As String) '配列の開始
  concat result, "["
End Sub

Private Sub startDic(result As String)
  concat result, "{"
End Sub

Private Sub pushArrMem(result As String, mem, recursionCount_ As Long) '配列要素の追加
  concat result, p(mem, False, recursionCount_ + 1), ARRAY_SEGMENT
End Sub

Private Sub pushDicMem(result As String, key, Value, recursionCount_ As Long) 'Dictionaryの要素の追加
  concat result, p(key, False, recursionCount_ + 1), " => ", p(Value, False, recursionCount_ + 1), ARRAY_SEGMENT
End Sub

Private Sub pushEllipsis(result As String)
  concat result, "..., "
End Sub

Private Sub endArr(result As String, Optional isLast As Boolean = False, Optional isEmptyArr As Boolean = False) '配列の終了
  If isEmptyArr Then
    result = result & "]"
  Else
    result = left(result, Len(result) - Len(ARRAY_SEGMENT)) & "]"
  End If
  If Not isLast Then concat result, ARRAY_SEGMENT
End Sub

Private Sub endDic(result As String, Optional isEmptyDic As Boolean = False) '配列の終了
  If isEmptyDic Then
    result = result & "}"
  Else
    result = left(result, Len(result) - Len(ARRAY_SEGMENT)) & "}"
  End If
End Sub

'文字列の結合
Private Sub concat(s As String, ParamArray strs())
  s = s & Join(strs, "")
End Sub

'templateStr("a$b$c",1,2,3) == "a1b2c" のように$を使って文字列を生成
Private Function templateStr(format As String, ParamArray val()) As String
  Dim arr, i As Long
  On Error GoTo onerror
  arr = Split(format, "$")
  templateStr = arr(0)
  For i = 0 To UBound(val)
    concat templateStr, val(i), arr(i + 1)
  Next
  Exit Function
onerror:
  Debug.Print "エラー: templateStrの引数が不正です"
End Function
