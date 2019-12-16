VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CCombudget"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit


Private m_Conn As ADODB.Connection

Public ShowMode As SHOW_MODE_TYPE
Public Flag As String
Public QueryFlag As Long

Public COM_BUDGET_ID As Long
Public MANAGER_ID As String
Public MANAGER_NAME As String
Public FROM_DATE As Date
Public TO_DATE As Date

Public ORDER_BY As Long
Public ORDER_TYPE As Long
Public Temp_ID As Long
Public TempFlag As String
Public ImportExportItems As Collection
Private Sub Class_Initialize()
    Set m_Conn = glbDatabaseMngr.DBLegacyConnection
   Set ImportExportItems = New Collection
End Sub
Public Sub DeleteData()
Dim SQL As String
   
   SQL = "DELETE FROM COM_BUDGET WHERE (COM_BUDGET_ID = " & COM_BUDGET_ID & ")"
      m_Conn.Execute (SQL)
End Sub
Public Sub DeleteSubData()
Dim SQL As String
   
   SQL = "DELETE FROM COM_SUB_BUDGET WHERE (COM_BUDGET_ID = " & COM_BUDGET_ID & ")"
         m_Conn.Execute (SQL)
End Sub
Public Sub QueryData(Ind As Long, Rs As ADODB.Recordset, ItemCount As Long, Optional MustCount As Boolean = False)
Dim SQL1 As String
Dim SQL2 As String
Dim SelectStr1 As String
Dim SelectStr2 As String
Dim FromStr As String
Dim WhereStr As String
Dim GroupByStr As String
Dim OrderByStr As String
Dim TempStr As String
   
   If Ind = 1 Then
      SelectStr1 = "SELECT GE.* "
      SelectStr2 = "SELECT COUNT(*) "
      
      FromStr = "FROM COM_BUDGET GE "
   '   FromStr = FromStr & "LEFT OUTER JOIN OESLM SLM ON (GE.MANAGER_ID = SLM.SLMCOD) "
      
      WhereStr = ""
   End If
   
   If COM_BUDGET_ID > 0 Then
      If Len(WhereStr) > 0 Then
         TempStr = "AND "
      Else
         TempStr = "WHERE "
      End If
      WhereStr = TempStr & " (GE.COM_BUDGET_ID = " & COM_BUDGET_ID & ") "
   End If

   OrderByStr = "ORDER BY FROM_DATE "
   
   If ORDER_TYPE = 1 Then
      OrderByStr = OrderByStr & " ASC "
   ElseIf ORDER_TYPE = 2 Then
      OrderByStr = OrderByStr & " DESC "
   Else
      OrderByStr = OrderByStr & " ASC "
   End If
   
   SQL1 = SelectStr1 & FromStr & WhereStr & GroupByStr & OrderByStr
   SQL2 = SelectStr2 & FromStr & WhereStr & GroupByStr
   'debug.print SQL1
   Call glbDatabaseMngr.GetLegacyRs(SQL1, SQL2, True, ItemCount, Rs, glbErrorLog)
End Sub
Public Sub PopulateFromRS(Ind As Long, Rs As ADODB.Recordset)
   If Ind = 1 Then
      COM_BUDGET_ID = NVLI(Rs("COM_BUDGET_ID"), -1)
      FROM_DATE = InternalDateToDate(NVLS(Rs("FROM_DATE"), ""))
      TO_DATE = InternalDateToDate(NVLS(Rs("TO_DATE"), ""))
      MANAGER_ID = NVLS(Rs("MANAGER_ID"), "")
'      MANAGER_NAME = NVLS(Rs("SLMNAM"), "")
   End If
End Sub
Public Function AddEditData()
Dim ServerDateTime As String
Dim SQL1 As String
Dim ID As Long
Dim TempID1 As String

'   TempID1 = PARENT_ID
'   If PARENT_ID <= 0 Then
'      TempID1 = "NULL"
'   End If
      
'   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
'
   If ShowMode = SHOW_ADD Then
'      Call glbDatabaseMngr.GetSeqID("COM_BUDGET_SEQ", ID, glbErrorLog)
      
      SQL1 = "INSERT INTO COM_BUDGET " & _
                     "( " & _
                     "   FROM_DATE, " & _
                     "   TO_DATE, " & _
                    "    MANAGER_ID" & _
                     ") "
         SQL1 = SQL1 & _
                     "VALUES " & _
                     "( " & _
                        "'" & ChangeQuote(Trim(FROM_DATE)) & "'," & _
                        "'" & ChangeQuote(Trim(TO_DATE)) & "'," & _
                        "'" & ChangeQuote(Trim(MANAGER_ID)) & "' " & _
                     ") "
            Temp_ID = ID
            
'                         "'" & glbUser.USER_NAME & "'" & _
 '                       "'" & ServerDateTime & "'," & _

   Else 'Show edit
      SQL1 = "UPDATE COM_BUDGET SET " & _
                     "    FROM_DATE = '" & ChangeQuote(Trim(FROM_DATE)) & "', " & _
                     "    TO_DATE = '" & ChangeQuote(Trim(TO_DATE)) & "', " & _
                     "    MANAGER_ID = '" & ChangeQuote(Trim(MANAGER_ID)) & "' " & _
                     "WHERE COM_BUDGET_ID = " & COM_BUDGET_ID
 '                    "    TO_DATE = '" & ServerDateTime & "', " & _
 '                    "    MANAGER_ID = '" & glbUser.USER_NAME & "' " & _

   End If
   
   If (ShowMode = SHOW_ADD) Or (ShowMode = SHOW_EDIT) Then
          'debug.print SQL1
          m_Conn.Execute (SQL1)
   End If
End Function
Private Sub Class_Terminate()
   Set ImportExportItems = Nothing
End Sub


