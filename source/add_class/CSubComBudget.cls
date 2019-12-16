VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CSupCombudget"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_Conn As ADODB.Connection

Public ShowMode As SHOW_MODE_TYPE
Public Flag As String
Public QueryFlag As Long

Public COM_SUB_BUDGET_ID As Long
Public COM_BUDGET_ID As Long
Public SLM_ID As String
Public MASTER_AREA_ID As Long
Public BUDGET As String

Public FROM_DATE As Date
Public TO_DATE As Date
Public MANAGER_ID As String
Public MASTER_AREA_NAME As String

Public ORDER_BY As Long
Public ORDER_TYPE As Long

Public m_FieldList As Collection
Private FieldValue As Variant
Private Sub Class_Initialize()
   Set m_Conn = glbDatabaseMngr.DBLegacyConnection
   Set m_FieldList = New Collection
End Sub
Private Sub Class_Terminate()
   Set m_FieldList = Nothing
End Sub
Public Sub DeleteData()
Dim SQL As String
   SQL = "DELETE FROM COM_SUB_BUDGET WHERE (COM_SUB_BUDGET_ID = " & COM_SUB_BUDGET_ID & ")"
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
      SelectStr1 = "SELECT GSE.* "
      SelectStr1 = SelectStr1 & ",GE.FROM_DATE , GE.TO_DATE , GE.MANAGER_ID "
 '     SelectStr1 = SelectStr1 & ",CA.MASTER_AREA_NAME "
         
      SelectStr2 = "SELECT COUNT(*) "
      
      FromStr = "FROM COM_SUB_BUDGET GSE "
      FromStr = FromStr & "LEFT OUTER JOIN COM_BUDGET GE ON (GSE.COM_BUDGET_ID = GE.COM_BUDGET_ID) "
 '     FromStr = FromStr & "LEFT OUTER JOIN  COMMISSION_MASTER_AREA CA ON (GSE.MASTER_AREA_ID = CA.MASTER_AREA_ID) "
      
      WhereStr = ""
   End If
   
   If COM_BUDGET_ID > 0 Then
      If Len(WhereStr) > 0 Then
         TempStr = "AND "
      Else
         TempStr = "WHERE "
      End If
      WhereStr = TempStr & " (GE.COM_BUDGET_ID = " & COM_BUDGET_ID & ") " & vbCrLf
   End If
   
   If COM_SUB_BUDGET_ID > 0 Then
      If Len(WhereStr) > 0 Then
         TempStr = "AND "
      Else
         TempStr = "WHERE "
      End If
      WhereStr = TempStr & " (GE.COM_SUB_BUDGET_ID = " & COM_SUB_BUDGET_ID & ") " & vbCrLf
   End If
   
   OrderByStr = "ORDER BY GSE.COM_SUB_BUDGET_ID "
   
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
      COM_SUB_BUDGET_ID = NVLI(Rs("COM_SUB_BUDGET_ID"), -1)
      COM_BUDGET_ID = NVLI(Rs("COM_BUDGET_ID"), -1)
      SLM_ID = NVLS(Rs("SLM_ID"), "")
      MASTER_AREA_ID = NVLI(Rs("MASTER_AREA_ID"), -1)
'      MASTER_AREA_NAME = NVLS(Rs("MASTER_AREA_NAME"), "")
      FROM_DATE = InternalDateToDate(NVLS(Rs("FROM_DATE"), ""))
      TO_DATE = InternalDateToDate(NVLS(Rs("TO_DATE"), ""))
      MANAGER_ID = NVLS(Rs("MANAGER_ID"), "")
      BUDGET = NVLS(Rs("BUDGET"), "")
      Flag = NVLS(Rs("FLAG"), "")
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
      
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   
   If ShowMode = SHOW_ADD Then
      'Call glbDatabaseMngr.GetSeqID("ENTERPRISE_SEQ", ID, glbErrorLog)
      
      SQL1 = "INSERT INTO COM_SUB_BUDGET " & _
                     "( " & _
                     "   COM_BUDGET_ID, " & _
                     "   SLM_ID, " & _
                     "   BUDGET, " & _
                    "    FLAG" & _
                     ") "
         SQL1 = SQL1 & _
                     "VALUES " & _
                     "( " & _
                        "'" & ChangeQuote(Trim(COM_BUDGET_ID)) & "'," & _
                        "'" & ChangeQuote(Trim(SLM_ID)) & "'," & _
                        "'" & ChangeQuote(Trim(BUDGET)) & "'," & _
                        "'" & ChangeQuote(Trim(Flag)) & "' " & _
                     ") "
   Else 'Show edit
      SQL1 = "UPDATE COM_SUB_BUDGET SET " & _
                     "    COM_BUDGET_ID = '" & ChangeQuote(Trim(COM_BUDGET_ID)) & "', " & _
                     "    MANAGER_ID = '" & ChangeQuote(Trim(MANAGER_ID)) & "', " & _
                     "    BUDGET = '" & ChangeQuote(Trim(BUDGET)) & "', " & _
                     "    FLAG = '" & ChangeQuote(Trim(Flag)) & "', " & _
                     "WHERE COM_SUB_BUDGET_ID = " & COM_SUB_BUDGET_ID
   End If
   
   If (ShowMode = SHOW_ADD) Or (ShowMode = SHOW_EDIT) Then
           m_Conn.Execute (SQL1)
   End If
End Function
Public Sub SetValue(Value As Variant)
   FieldValue = Value
End Sub


