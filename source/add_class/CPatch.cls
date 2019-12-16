VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CPatch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private m_Conn As ADODB.Connection

Public AddEditMode As SHOW_MODE_TYPE
Public QueryFlag As Long
Public OrderBy As Long
Public OrderType As Long
Public Flag As String
' End of generic data member

Private Sub Class_Initialize()
   Set m_Conn = glbDatabaseMngr.DBConnection
End Sub

Public Function IsPatch(VersionStr As String) As Boolean
On Error GoTo ErrorHandler
Dim SQL As String
Dim TempRs As ADODB.Recordset
Dim ItemCount As Long
Dim Result As Boolean

   Set TempRs = New ADODB.Recordset
   Result = True
   
   SQL = "SELECT * FROM PATCH_DELIVERY WHERE PATCH_VERS = '" & VersionStr & "' "
   Call glbDatabaseMngr.GetRs(SQL, "", False, ItemCount, TempRs, glbErrorLog)

   If TempRs.EOF Then
      Result = fasle
   End If
   
   If TempRs.State = adStateOpen Then
      TempRs.Close
   End If
   Set TempRs = Nothing
   
   IsPatch = Result
   Exit Function
   
ErrorHandler:
   IsPatch = False
End Function

Public Function Patch_1_0_12_9() As Boolean
On Error GoTo ErrorHandler
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String
Dim IsOK As Boolean

   VersionStr = "1_0_12_9"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
      
'   SQL = "CREATE TABLE TEST1 (FIELD1 STRING) "
'   m_Conn.Execute (SQL)
   
   m_Conn.BeginTrans
   HasBegin = True

   SQL = "INSERT INTO PATCH_DELIVERY "
   SQL = SQL & "( "
   SQL = SQL & "PATCH_DELI, "
   SQL = SQL & "PATCH_VERS, "
   SQL = SQL & "PATCH_DATE, "
   SQL = SQL & "PATCH_BY "
   SQL = SQL & ") VALUES "
   SQL = SQL & "( "
   SQL = SQL & "GEN_ID(PATCH_DELIVERY_SEQ, 1), "
   SQL = SQL & "'" & VersionStr & "', "
   SQL = SQL & "'" & ServerDateTime & "', "
   SQL = SQL & "" & glbUser.USER_ID & " "
   SQL = SQL & ")"
   m_Conn.Execute (SQL)

   m_Conn.CommitTrans
   HasBegin = False
   
   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError
   
   Patch_1_0_12_9 = True
   Exit Function
   
ErrorHandler:
   'debug.print Err.DESCRIPTION
   If HasBegin Then
      m_Conn.RollbackTrans
   End If
End Function

