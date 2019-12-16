VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CXlsCarkillFW"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_Conn As ADODB.Connection

Public AddEditMode As SHOW_MODE_TYPE
Public QueryFlag As Long
Public OrderBy As Long
Public OrderType As Long
Public Flag As String
' End of generic data member

Public XLS_FORWARD_ID As Long
Public FW_ROW As String
Public MAIN_FLAG As String
Public UPPER_FLAG As String
Private Sub Class_Initialize()
   Set m_Conn = glbDatabaseMngr.DBLegacyConnection
End Sub

'In this class we will not handle any error or any transaction
Public Function QueryData(Ind As Long, Rs As ADODB.Recordset, ItemCount As Long)
Dim SQL1 As String
Dim SQL2 As String
Dim SelectStr1 As String
Dim SelectStr2 As String
Dim FromStr As String
Dim WhereStr As String
Dim OrderByStr As String
Dim OrderTypeStr As String
Dim TempDate As String

If Ind = 1 Then
   SelectStr1 = "SELECT XS.*  "
   SelectStr2 = "SELECT COUNT(*)  "
   
   FromStr = "FROM XLS_FORWARD XS "
   
   WhereStr = ""
End If

    If XLS_FORWARD_ID > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (XLS_FORWARD_ID = " & ChangeQuote(Trim(XLS_FORWARD_ID)) & ") "
      Else
         WhereStr = WhereStr & "AND (XLS_FORWARD_ID = " & ChangeQuote(Trim(XLS_FORWARD_ID)) & ") "
      End If
   End If
   
   OrderByStr = ""
   OrderTypeStr = ""

   OrderByStr = "ORDER BY XS.MAIN_FLAG DESC , XS.UPPER_FLAG DESC ,  XS.FW_ROW "
   OrderTypeStr = "ASC "
   
   SQL1 = SelectStr1 & FromStr & WhereStr & OrderByStr & OrderTypeStr
   SQL2 = SelectStr2 & FromStr & WhereStr
   'debug.print SQL1
   Call glbDatabaseMngr.GetLegacyRs(SQL1, SQL2, True, ItemCount, Rs, glbErrorLog)
End Function

Public Function PopulateFromRS(Ind As Long, Rs As ADODB.Recordset)
  If Ind = 1 Then
    XLS_FORWARD_ID = NVLI(Rs("XLS_FORWARD_ID"), -1)
    FW_ROW = NVLS(Rs("FW_ROW"), "")
    MAIN_FLAG = NVLS(Rs("MAIN_FLAG"), "")
    UPPER_FLAG = NVLS(Rs("UPPER_FLAG"), "")
  End If
End Function
Public Function AddEditData()
Dim ServerDateTime As String
Dim SQL1 As String

   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
'   ENTERPRISE_CODE = glbCompanyCode
   
   If AddEditMode = SHOW_ADD Then
      SQL1 = "INSERT INTO XLS_FORWARD " & _
                     "( " & _
                     "   FW_ROW, " & _
                     "   MAIN_FLAG, " & _
                     "   UPPER_FLAG " & _
                     ") " & _
                     "VALUES " & _
                     "( " & _
                        "'" & ChangeQuote(Trim(FW_ROW)) & "', " & _
                        "'" & ChangeQuote(Trim(MAIN_FLAG)) & "', " & _
                        "'" & ChangeQuote(Trim(UPPER_FLAG)) & "' " & _
                      ") "

   Else 'Show edit
      SQL1 = "UPDATE XLS_FORWARD SET " & _
                     "   FW_ROW = '" & ChangeQuote(Trim(FW_ROW)) & "', " & _
                     "   MAIN_FLAG = '" & ChangeQuote(Trim(MAIN_FLAG)) & "', " & _
                     "   UPPER_FLAG = '" & ChangeQuote(Trim(UPPER_FLAG)) & "' " & _
                     "WHERE XLS_FORWARD_ID = " & XLS_FORWARD_ID
   End If
   
   If (AddEditMode = SHOW_ADD) Or (AddEditMode = SHOW_EDIT) Then
      m_Conn.Execute (SQL1)
   End If
End Function

Public Function DeleteData()
Dim SQL1 As String
      SQL1 = "DELETE FROM XLS_FORWARD WHERE XLS_FORWARD_ID = " & XLS_FORWARD_ID
      m_Conn.Execute (SQL1)
End Function
