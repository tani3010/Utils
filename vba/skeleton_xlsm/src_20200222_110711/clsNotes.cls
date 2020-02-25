VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsNotes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Const clsName As String = "clsNotes"

Private objNotes As Object

Private Sub Class_Initialize()
    Set objNotes = CreateObject(OBJ_NOTES_OLE)
End Sub

Private Sub Class_Terminate()
    Set objNotes = Nothing
End Sub

