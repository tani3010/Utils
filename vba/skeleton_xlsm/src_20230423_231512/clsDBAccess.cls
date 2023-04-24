VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsDBAccess"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Const clsName As String = "clsDBAccess"

' variables for connection string
Private CON_STR            As String ' connection string
Private CON_DRIVER         As String
Private CON_SERVER         As String
Private CON_NETWORKADDRESS As String
Private CON_DBQUERY        As String
Private CON_USERID         As String
Private CON_PASSWORD       As String

' constant driver name
Private Const DRIVER_ADAPTIVE       As String = "{Adaptive Server Enterprise}"
Private Const DRIVER_MSACCESS       As String = "{Microsoft Access Driver (*.mdb, *.accdb)}"
Private Const DRIVER_MSEXCEL        As String = "{Microsoft Excel Driver (*.xls, *.xlsx, *.xlsm, *.xlsb)}"
' Private Const DRIVER_MSTEXT         As String = "{Microsoft Access Text Driver (*.txt, *.csv)}"
Private Const DRIVER_MSTEXT         As String = "{Microsoft Text Driver (*.txt; *.csv)}"
Private Const DRIVER_ORACLE         As String = "{Microsoft ODBC for Oracle}"
Private Const DRIVER_SQLSERVEROLEDB As String = "{SQLOLEDB}"
Private Const DRIVER_SQLSERVERODBC  As String = "{SQL Server}"
Private Const DRIVER_ODBC           As String = "{Adaptive Server Enterprise}"
Private Const DRIVER_SQLITE3        As String = "{SQLite3 ODBC Driver}"  ' need to install SQLite3 ODBC Driver

Private Const DB_FAIL_ONERROR As Long = 128

Public Property Let set_driver(ByVal x As String): CON_DRIVER = x: End Property
Public Property Let set_server(ByVal x As String): CON_SERVER = x: End Property
Public Property Let set_networkAddress(ByVal x As String): CON_NETWORKADDRESS = x: End Property
Public Property Let set_dbQuery(ByVal x As String): CON_DBQUERY = x: End Property
Public Property Let set_userId(ByVal x As String): CON_USERID = x: End Property
Public Property Let set_password(ByVal x As String): CON_PASSWORD = x: End Property

Private Sub Class_Initialize()
    CON_STR = vbNullString
    CON_DRIVER = vbNullString
    CON_SERVER = vbNullString
    CON_NETWORKADDRESS = vbNullString
    CON_DBQUERY = vbNullString
    CON_USERID = vbNullString
    CON_PASSWORD = vbNullString
End Sub

Private Sub Class_Terminate()
    Call Class_Initialize
End Sub

Private Function appendStringHelper( _
    ByVal baseStr As String, ByVal appendStr1 As String, ByVal appendStr2 As String) As String
    
    If appendStr1 = vbNullString Or appendStr2 = vbNullString Then
        appendStringHelper = baseStr
    Else
        appendStringHelper = baseStr & IIf(Right(baseStr, 1) = DELIM_EQUAL, appendStr1, appendStr1 & DELIM_EQUAL)
        appendStringHelper = appendStringHelper & IIf(Right(appendStr2, 1) = DELIM_SEMICOLON, appendStr2, appendStr2 & DELIM_SEMICOLON & vbCrLf)
    End If
End Function

Private Sub buildConnectionStringFor_AdaptiveServerEnterprise(Optional ByVal readOnly As Boolean = True)
    CON_STR = vbNullString
    CON_STR = appendStringHelper(CON_STR, "DRIVER", DRIVER_ADAPTIVE)
    CON_STR = appendStringHelper(CON_STR, "SRVR", CON_SERVER)
    CON_STR = appendStringHelper(CON_STR, "NA", CON_NETWORKADDRESS)
    CON_STR = appendStringHelper(CON_STR, "DB", CON_DBQUERY)
    CON_STR = appendStringHelper(CON_STR, "UID", CON_USERID)
    CON_STR = appendStringHelper(CON_STR, "PWD", CON_PASSWORD)
End Sub

Private Sub buildConnectionStringFor_SQLServerOLEDB(Optional ByVal readOnly As Boolean = True)
    CON_STR = vbNullString
    CON_STR = appendStringHelper(CON_STR, "Provider", DRIVER_SQLSERVEROLEDB)
    CON_STR = appendStringHelper(CON_STR, "Data Source", CON_NETWORKADDRESS)
    CON_STR = appendStringHelper(CON_STR, "Initial Catalog", CON_DBQUERY)
    CON_STR = appendStringHelper(CON_STR, "User ID", CON_USERID)
    CON_STR = appendStringHelper(CON_STR, "Password", CON_PASSWORD)
End Sub

Private Sub buildConnectionStringFor_SQLServerODBC(Optional ByVal readOnly As Boolean = True)
    CON_STR = vbNullString
    CON_STR = appendStringHelper(CON_STR, "Driver", DRIVER_SQLSERVERODBC)
    CON_STR = appendStringHelper(CON_STR, "Server", CON_NETWORKADDRESS)
    CON_STR = appendStringHelper(CON_STR, "Database", CON_DBQUERY)
    CON_STR = appendStringHelper(CON_STR, "Uid", CON_USERID)
    CON_STR = appendStringHelper(CON_STR, "Pwd", CON_PASSWORD)
End Sub

Private Sub buildConnectionStringFor_Oracle(Optional ByVal readOnly As Boolean = True)
    CON_STR = vbNullString
    CON_STR = appendStringHelper(CON_STR, "DRIVER", DRIVER_ORACLE)
    CON_STR = appendStringHelper(CON_STR, "SERVER", CON_SERVER)
    CON_STR = appendStringHelper(CON_STR, "UID", CON_USERID)
    CON_STR = appendStringHelper(CON_STR, "PWD", CON_PASSWORD)
End Sub

Private Sub buildConnectionStringFor_ODBC(Optional ByVal readOnly As Boolean = True)
    CON_STR = vbNullString
    CON_STR = appendStringHelper(CON_STR, "DSN", DRIVER_ODBC)
    CON_STR = appendStringHelper(CON_STR, "UID", CON_USERID)
    CON_STR = appendStringHelper(CON_STR, "PWD", CON_PASSWORD)
End Sub

Private Sub buildConnectionStringFor_MSACCESS(Optional ByVal readOnly As Boolean = True)
    CON_STR = vbNullString
    CON_STR = appendStringHelper(CON_STR, "Driver", DRIVER_MSACCESS)
    CON_STR = appendStringHelper(CON_STR, "DBQ", CON_DBQUERY)
    CON_STR = appendStringHelper(CON_STR, "Uid", CON_USERID)
    CON_STR = appendStringHelper(CON_STR, "Pwd", CON_PASSWORD)
End Sub

Private Sub buildConnectionStringFor_MSEXCEL(Optional ByVal readOnly As Boolean = True)
    CON_STR = vbNullString
    CON_STR = appendStringHelper(CON_STR, "Driver", DRIVER_MSEXCEL)
    CON_STR = appendStringHelper(CON_STR, "DBQ", CON_DBQUERY)
    If readOnly Then
        CON_STR = appendStringHelper(CON_STR, "ReadOnly", 1)
    End If
End Sub

Private Sub buildConnectionStringFor_MSTEXT(Optional ByVal readOnly As Boolean = True)
    CON_STR = vbNullString
    CON_STR = appendStringHelper(CON_STR, "Driver", DRIVER_MSTEXT)
    CON_STR = appendStringHelper(CON_STR, "DBQ", CON_DBQUERY)
    If readOnly Then
        CON_STR = appendStringHelper(CON_STR, "ReadOnly", 1)
    End If
End Sub

Private Sub buildConnectionStringFor_SQLITE3(Optional ByVal readOnly As Boolean = True)
    CON_STR = vbNullString
    CON_STR = appendStringHelper(CON_STR, "DRIVER", DRIVER_SQLITE3)
    CON_STR = appendStringHelper(CON_STR, "Database", CON_DBQUERY)
End Sub

Private Sub buildConnectionString(Optional ByVal readOnly As Boolean = True)
    Select Case CON_DRIVER
        Case DRIVER_ADAPTIVE
            Call buildConnectionStringFor_AdaptiveServerEnterprise(readOnly)
        
        Case DRIVER_MSACCESS
            Call buildConnectionStringFor_MSACCESS(readOnly)
        
        Case DRIVER_MSEXCEL
            Call buildConnectionStringFor_MSEXCEL(readOnly)
        
        Case DRIVER_MSTEXT
            Call buildConnectionStringFor_MSTEXT(readOnly)
        
        Case DRIVER_ODBC
            Call buildConnectionStringFor_ODBC(readOnly)
        
        Case DRIVER_ORACLE
            Call buildConnectionStringFor_Oracle(readOnly)
        
        Case DRIVER_SQLSERVERODBC
            Call buildConnectionStringFor_SQLServerODBC(readOnly)
        
        Case DRIVER_SQLSERVEROLEDB
            Call buildConnectionStringFor_SQLServerOLEDB(readOnly)
            
        Case DRIVER_SQLITE3
            Call buildConnectionStringFor_SQLITE3(readOnly)
        
        Case Else
    
    End Select
End Sub

Public Sub executeSQL(ByVal sql As String, Optional ByVal shName As String = ST_DATA, Optional ByVal batchPaste As Boolean = True)
    Dim con As Object
    Dim rs As Object
    
    Call buildConnectionString(True)
    Set con = CreateObject(OBJ_ADODBCON)
    
    Call enableControl(False)
    With con
        .ConnectionString = CON_STR
        Call .Open
        Set rs = .Execute(sql)
        
        Call createSheet(shName, ThisWorkbook)
        With ThisWorkbook.Worksheets(shName)
            Call .Cells.Clear
            Dim i As Integer
            For i = 1 To rs.fields.Count
                .Cells(1, i).Value = rs.fields(i - 1).name
            Next i
            
            If batchPaste Then
                Call .Range(CURSOR_POSI_DEFAULT2).CopyFromRecordset(rs)
            Else
                Dim row As Long
                Dim col As Integer
                row = 2
                
                While Not rs.EOF
                    For col = 0 To rs.fields.Count - 1
                        .Cells(row, col + 1).Value = rs.fields(col).Value
                    Next col
                    row = row + 1
                    Call rs.MoveNext
                Wend
            End If
        End With
        
        Call rs.Close
        Call .Close
    End With
    
    Set con = Nothing
    Set rs = Nothing
    Call enableControl(True)
End Sub

Public Sub executeSQL_withTransaction(ByVal sql As String, Optional ByVal shName As String = ST_DATA)

On Error GoTo LABEL_ERROR
    
    Dim con As Object
    Dim rs As Object
    
    Call initializeErrorHandler
    Set ERROR_HANDLER = New clsErrorHandler
    
    Call buildConnectionString(False)
    Set con = CreateObject(OBJ_ADODBCON)
    
    With con
        .ConnectionString = CON_STR
        Call .Open
        Call .beginTrans
        Set rs = .Execute(sql, DB_FAIL_ONERROR)
        
        Select Case rs.State
            Case 1
                Call rs.Close
            Case Else
            
        End Select
        
        Call .CommitTrans
        Call .Close
    End With
    
LABEL_SUCCESS:
    Call ERROR_HANDLER.writeStandardLog(paste("::", clsName, "executeSQL_withTransaction"))
    GoTo LABEL_FINALLY
    
LABEL_ERROR:
    Call ERROR_HANDLER.writeErrorLog(paste("::", clsName, "executeSQL_withTransaction"), vbNullString)
    Call con.RollbackTrans
    Call con.Close

LABEL_FINALLY:
    Set ERROR_HANDLER = Nothing
    Call enableControl(True)
    
    Set con = Nothing
    Set rs = Nothing
End Sub

