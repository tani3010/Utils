Attribute VB_Name = "stdPrint"
Option Explicit

' https://qiita.com/o_pos/items/482455a06a6150516dd7

'�ȉ��̒萔�͑S�ĕύX�\ (���̕����̏����������K�v�Ȃ�)
Const ARRAY_SEGMENT = ", "       ' �z�񓙂̗v�f�̋�؂蕶��
Const MAX_ARR_ELEMENTS = 100     ' �z��̗v�f�̕\�����̏��
Const MAX_ARR_TRACKING = 20      ' �z��̃l�X�g(�z��̎����Ƃ͈Ⴄ)�̕\���̏��
Const MAX_COLLECTION_SIZE = 100  ' �R���N�V�����̗v�f�̕\�����̏��
Const MAX_DICTIONARY_SIZE = 100  ' dictionary�̗v�f�̕\�����̏��
Const WRITE_EMPTY = True         ' Empty��\�����邩 (False�Ȃ�A���Ƃ���[Empty]��[]�ɂȂ�)

'�z��̌^���擾
Function getArrayType(arr) As String '(arr)��(arr())�ɂ����Valiant�ȊO�̌^�̔z����󂯎��Ȃ��Ȃ�
  Dim s As String: s = TypeName(arr)
  Debug.Assert s Like "*(*)" '�z��ł��邱�Ƃ��m�F
  getArrayType = left(s, Len(s) - 2)
End Function

Function p(src, Optional debugPrint As Boolean = True, Optional recursionCount_ As Long = 0) As String
  Dim result As String, srcType As String
  result = ""
  srcType = TypeName(src)
  Select Case srcType
  Case "String" '������
    result = """" & Replace(src, """", """""") & """"
  Case "Integer", "Double", "Single", "Long", "Currency", "Byte" '���l
    result = src
  Case "Decimal" '�\�i��
    result = src & " As Decimal"
  Case "Date" '���t
    result = """" & src & """ As Date"
  Case "Boolean" '�u�[���l
    result = src
  Case "Collection" '�R���N�V����
    processCollection result, src, "Collection", recursionCount_
  Case "Dictionary" 'Scripting.Dictionary
    processDictionary result, src, recursionCount_
  Case Else
    If srcType Like "*(*)" Then '�z��
      processArr result, src, srcType, recursionCount_
    ElseIf isShortCollection(src) Then '�R���N�V�����I�I�u�W�F�N�g
      processCollection result, src, srcType, recursionCount_
    Else '���̑�
      result = srcType
    End If
  End Select
  If recursionCount_ > 0 And result = "Empty" And Not WRITE_EMPTY Then result = ""
  If debugPrint Then Debug.Print result
  p = result
End Function

'�z��̎������擾 (Dim a() ��a�Ȃ�0)
Function getArrayDim(arr, Optional dimIter_ As Long = 1) As Long
  On Error GoTo onerror
  Dim upper As Long
  upper = UBound(arr, dimIter_)
  If upper = -1 Then
    'Function UBound2(a, b): UBound2 = UBound(a, b): End Function ������Ƃ��ɁA
    'Dim arr() As String: UBound(arr, 1) �̓G���[�A
    'Dim arr() As Range:  UBound(arr, 1) ���G���[�A
    'Dim arr() As Range:  Ubound2(arr, 1) ��-1��Ԃ�
    '�Ƃ������ۂւ̑Ώ�
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

'Collection�����������ȉ��Ȃ�true
Private Function isShortCollection(c) As Boolean
  On Error GoTo onerror
  isShortCollection = c.Count <= MAX_COLLECTION_SIZE
  Exit Function
onerror:
  isShortCollection = False
End Function

'�z��̏ꍇ�̏���
Private Sub processArr(result As String, src, srcType As String, recursionCount_ As Long)
  Dim arrType As String, arrdim As Long
  Dim i As Long, j As Long, k As Long
  arrType = left(srcType, Len(srcType) - 2)
  arrdim = getArrayDim(src)
  '��O����
  If recursionCount_ > MAX_ARR_TRACKING Then
    result = arrType & "[...]": Exit Sub
  ElseIf arrdim = 0 Then
    result = "(0dim)[] As " & arrType: Exit Sub
  ElseIf arrdim > 3 Then
    '����3�͑������z��̓W�J���������̎����B
    '�ύX����ꍇ�͉���select case�����̏����������K�v�B
    result = templateStr("($dim)[...] As $", arrdim, arrType): Exit Sub
  End If
  '�z���W�J
  If LBound(src) <> 0 Then concat result, templateStr("(from$)", LBound(src))
  startArr result
  '3�����܂ł����\�����Ȃ��̂ŁA�K���ɏ�����
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

'WorksheetFunction���������p
Private Function min(a, b)
  Debug.Assert IsNumeric(a) And IsNumeric(b)
  If a < b Then
    min = a
  Else
    min = b
  End If
End Function

'WorksheetFunction���������p
Private Function max(a, b)
  Debug.Assert IsNumeric(a) And IsNumeric(b)
  If a < b Then
    max = b
  Else
    max = a
  End If
End Function

'�R���N�V�����̓Y�����̍ŏ��l(0��1�ȊO����n�܂�ꍇ�ɂ͔�Ή�, ���Collection�Ȃ���1��Ԃ�)
Private Function collectionLBound(c)
  On Error GoTo onerror
  Dim tmp: tmp = c(0)
  collectionLBound = 0
  Exit Function
onerror:
  collectionLBound = 1
End Function

'�R���N�V�����̏ꍇ�̏���
Private Sub processCollection(result As String, src, srcType As String, recursionCount_ As Long)
  result = "Collection"
  If srcType <> "Collection" Then concat result, templateStr("($)", srcType)
  Dim lower As Long: lower = collectionLBound(src)
  If lower <> 1 Then concat result, templateStr("(from$)", lower)
  startArr result
  If src.Count = 0 Then
    endArr result, True, True
  Else
    'regexp��execute���\�b�h�̕Ԃ�l�ȂǁA�R���N�V������������
    '�Y������0����n�܂�I�u�W�F�N�g�����邩��AFor i = 1 To min(src.Count, MAX_COLLECTION_SIZE) �ł͕s��
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

'Dictionary�̏ꍇ�̏���
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

'�z�񕶎���̍쐬
Private Sub startArr(result As String) '�z��̊J�n
  concat result, "["
End Sub

Private Sub startDic(result As String)
  concat result, "{"
End Sub

Private Sub pushArrMem(result As String, mem, recursionCount_ As Long) '�z��v�f�̒ǉ�
  concat result, p(mem, False, recursionCount_ + 1), ARRAY_SEGMENT
End Sub

Private Sub pushDicMem(result As String, key, Value, recursionCount_ As Long) 'Dictionary�̗v�f�̒ǉ�
  concat result, p(key, False, recursionCount_ + 1), " => ", p(Value, False, recursionCount_ + 1), ARRAY_SEGMENT
End Sub

Private Sub pushEllipsis(result As String)
  concat result, "..., "
End Sub

Private Sub endArr(result As String, Optional isLast As Boolean = False, Optional isEmptyArr As Boolean = False) '�z��̏I��
  If isEmptyArr Then
    result = result & "]"
  Else
    result = left(result, Len(result) - Len(ARRAY_SEGMENT)) & "]"
  End If
  If Not isLast Then concat result, ARRAY_SEGMENT
End Sub

Private Sub endDic(result As String, Optional isEmptyDic As Boolean = False) '�z��̏I��
  If isEmptyDic Then
    result = result & "}"
  Else
    result = left(result, Len(result) - Len(ARRAY_SEGMENT)) & "}"
  End If
End Sub

'������̌���
Private Sub concat(s As String, ParamArray strs())
  s = s & Join(strs, "")
End Sub

'templateStr("a$b$c",1,2,3) == "a1b2c" �̂悤��$���g���ĕ�����𐶐�
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
  Debug.Print "�G���[: templateStr�̈������s���ł�"
End Function
