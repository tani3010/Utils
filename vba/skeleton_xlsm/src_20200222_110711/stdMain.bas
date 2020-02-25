Attribute VB_Name = "stdMain"
Option Explicit

Private Const stdName As String = "stdMain"

#Const CATCH_ERROR = True

Public Sub calcMain()
#If CATCH_ERROR Then
    On Error GoTo LABEL_ERROR
    Call enableControl(False)
#Else
    Call enableControl(True)
#End If

    Const name As String = "calcMain"
    Call initializeErrorHandler

    ' ========== PROCEDURE BEGIN ====================================
    
    
    
    ' ========== PROCEDURE END ======================================
    
LABEL_SUCCESS:
    Call ERROR_HANDLER.writeStandardLog(name)
    GoTo LABEL_FINALLY
    
LABEL_ERROR:
    Call ERROR_HANDLER.writeStandardLog(name, vbNullString)

LABEL_FINALLY:
    Set ERROR_HANDLER = Nothing
    Call enableControl(True)
End Sub
