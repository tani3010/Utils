Attribute VB_Name = "stdDefine"
Option Explicit

Private Const stdName As String = "stdDefine"

' ===== Error Handler =====
Public Const CATCH_ERROR As Boolean = 0
Public ERROR_HANDLER As clsErrorHandler

' ===== LABEL =====
Public Enum LABEL
    LABEL_SUCCESS
    LABEL_ERROR
    LABEL_FINALLY
End Enum

' ===== worksheet name =====
Public Const ST_MAIN    As String = "MAIN"
Public Const ST_DATA    As String = "DATA"
Public Const ST_DEFINE  As String = "DEFINE"
Public Const ST_HOLIDAY As String = "HOLIDAY"
Public Const ST_TEST    As String = "TEST"

' ===== object name for late binding =====
Public Const OBJ_ADODBCON   As String = "ADODB.Connection"
Public Const OBJ_ADODBCMD   As String = "ADODB.Command"
Public Const OBJ_ADODBSTREM As String = "ADODB.Stream"
Public Const OBJ_DICTIONARY As String = "Scripting.Dictionary"
Public Const OBJ_EXCEL      As String = "Excel.Application"
Public Const OBJ_FSO        As String = "Scripting.FileSystemObject"
Public Const OBJ_IE         As String = "InternetExplorer.Application"
Public Const OBJ_OUTLOOK    As String = "Outlook.Application"
Public Const OBJ_NOTES_OLE  As String = "Notes.NotesSession"
Public Const OBJ_NOTES_COM  As String = "Lotus.NotesSession"
Public Const OBJ_POWERPOINT As String = "PowerPoint.Application"
Public Const OBJ_REGEXP     As String = "VBScript.RegExp"
Public Const OBJ_SHELL      As String = "Shell.Application"
Public Const OBJ_WORD       As String = "Word.Application"
Public Const OBJ_WSHELL     As String = "WScript.Shell"

' ===== extensions =====
Public Const EXT_XLS  As String = "xls"
Public Const EXT_XLSX As String = "xlsx"
Public Const EXT_XLSM As String = "xlsm"
Public Const EXT_XLAM As String = "xlam"
Public Const EXT_XLSB As String = "xlsb"
Public Const EXT_DOC  As String = "doc"
Public Const EXT_DOCX As String = "docx"
Public Const EXT_DOCM As String = "docm"
Public Const EXT_PPT  As String = "ppt"
Public Const EXT_PPTX As String = "pptx"
Public Const EXT_PPTM As String = "pptm"
Public Const EXT_TXT  As String = "txt"
Public Const EXT_TSV  As String = "tsv"
Public Const EXT_CSV  As String = "csv"
Public Const EXT_HTM  As String = "htm"
Public Const EXT_HTML As String = "html"
Public Const EXT_BMP  As String = "bmp"
Public Const EXT_JPG  As String = "jpg"
Public Const EXT_JPEG As String = "jpeg"
Public Const EXT_PDF  As String = "pdf"
Public Const EXT_PNG  As String = "png"
Public Const EXT_GIF  As String = "gif"
Public Const EXT_ZIP  As String = "zip"
Public Const EXT_RAR  As String = "rar"
Public Const EXT_LOG  As String = "log"
Public Const EXT_BAS  As String = "bas"
Public Const EXT_CLS  As String = "cls"
Public Const EXT_FRM  As String = "frm"
Public Const EXT_XML  As String = "xml"
Public Const EXT_XPS  As String = "xps"
Public Const EXT_SQL  As String = "sql"
Public Const EXT_DLL  As String = "dll"
Public Const EXT_JSON As String = "json"

' ===== format string =====
Public Const FMT_DATE_YYYYMMDDHHMMSS_SEP As String = "yyyy/mm/dd hh:mm:ss"
Public Const FMT_DATE_YYYYMMDD_HHMMSS    As String = "yyyymmdd_hhmmss"
Public Const FMT_DATE_YYYYMMDD_SEP_NUM   As String = "####/##/##"
Public Const FMT_DATE_YYYYMMDD_SEP       As String = "yyyy/mm/dd"
Public Const FMT_DATE_YYYYMMDD_SEP_SQL   As String = "'yyyy/mm/dd'"
Public Const FMT_DATE_YYYYMM_SEP         As String = "yyyy/mm"
Public Const FMT_DATE_YYYY               As String = "yyyy"
Public Const FMT_DATE_YYYYMMDD           As String = "yyyymmdd"
Public Const FMT_DATE_YYYYMM             As String = "yyyymm"
Public Const FMT_DATE_YYMMDD_SEP         As String = "yy/mm/dd"
Public Const FMT_DATE_YYMMDD             As String = "yymmdd"
Public Const FMT_DATE_MMDD               As String = "mmdd"
Public Const FMT_DATE_MMMDDYYYY          As String = "mmm.dd,yyyy"
Public Const FMT_DATE_YYYYMMDD_JPN       As String = "yyyy”NmmŒŽdd“ú"
Public Const FMT_DATE_GGGEMMDD_JPN       As String = "ggge”NmmŒŽdd“ú"
Public Const FMT_DATE_GGGEEMMDD_JPN      As String = "gggee”NmmŒŽdd“ú"
Public Const FMT_DATE_MMDD_JPN           As String = "mmŒŽdd“ú"
Public Const FMT_NUM_DIGITS_NUN_7        As String = "0000000"
Public Const FMT_NUM_DIGITS_NUM_3        As String = "000"

' ===== delimiter =====
Public Const DELIM_ATMARK      As String = "@"
Public Const DELIM_BACKQUOTE   As String = "`"
Public Const DELIM_COLON       As String = ":"
Public Const DELIM_COMMA       As String = ","
Public Const DELIM_CSV         As String = ","
Public Const DELIM_DATE        As String = "/"
Public Const DELIM_DOT         As String = "."
Public Const DELIM_DOUBLEQUOTE As String = """"
Public Const DELIM_EQUAL       As String = "="
Public Const DELIM_EXCLAIM     As String = "!"
Public Const DELIM_EXT         As String = "."
Public Const DELIM_HYPHEN      As String = "-"
Public Const DELIM_PATH_WIN    As String = "\"
Public Const DELIM_PATH_LIN    As String = "/"
Public Const DELIM_SEMICOLON   As String = ";"
Public Const DELIM_SINGLEQUOTE As String = "'"
Public Const DELIM_SPACE       As String = " "
Public Const DELIM_UNDERSCORE  As String = "_"
Public Const DELIM_MACRO       As String = "!"
Public Const DELIM_NAMESPACE   As String = "::"

' ===== initial date =====
Public Const NULL_DATE As Date = #1/1/1900#

' ===== cursor position =====
Public Const CURSOR_POSI_DEFAULT  As String = "A1"
Public Const CURSOR_POSI_DEFAULT2 As String = "A2"
