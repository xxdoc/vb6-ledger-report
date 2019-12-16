VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CAccountConfig"
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

Public ACCOUNT_CONFIG_ID As Long
Public SHEET As Long
Public ROW As Long
Public ROW2 As Long
Public COLLUMN_CODE As Long
Public COLLUMN_DEBIT As Long
Public COLLUMN_CREDIT As Long
Public COLLUMN_CHECK As Long
Public FROM_DATE As Date
Public TO_DATE As Date

Public KEY_ID As Long
Public KEY_LOOKUP As String

Private Sub Class_Initialize()
   Set m_Conn = glbDatabaseMngr.DBLegacyConnection
End Sub

'In this class we will not handle any error or any transaction
Public Function QueryData(Rs As ADODB.Recordset, ItemCount As Long)
Dim SQL1 As String
Dim SQL2 As String
Dim SelectStr1 As String
Dim SelectStr2 As String
Dim FromStr As String
Dim WhereStr As String
Dim OrderByStr As String
Dim OrderTypeStr As String

   SelectStr1 = "SELECT UG.*  "
   SelectStr2 = "SELECT COUNT(*)  "
   
   FromStr = "FROM ACCOUNT_CONFIG UG "
   
   WhereStr = ""
   
   OrderByStr = ""
   OrderTypeStr = ""

   If OrderBy = 1 Then
      OrderByStr = "ORDER BY ACCOUNT_CONFIG_ID "
   ElseIf OrderBy = 2 Then
      OrderByStr = "ORDER BY ACCOUNT_CONFIG_ID "
   Else
      OrderByStr = "ORDER BY ACCOUNT_CONFIG_ID "
   End If
   
   If OrderType = 1 Then
      OrderTypeStr = "ASC "
   ElseIf OrderType = 2 Then
      OrderTypeStr = "DESC "
   Else
      OrderTypeStr = "DESC "
   End If
   
   SQL1 = SelectStr1 & FromStr & WhereStr & OrderByStr & OrderTypeStr
   SQL2 = SelectStr2 & FromStr & WhereStr
   
   Call glbDatabaseMngr.GetLegacyRs(SQL1, SQL2, True, ItemCount, Rs, glbErrorLog)
End Function

Public Function PopulateFromRS(Ind As Long, Rs As ADODB.Recordset)
   ACCOUNT_CONFIG_ID = NVLI(Rs("ACCOUNT_CONFIG_ID"), -1)
   SHEET = NVLI(Rs("SHEET"), -1)
   ROW = NVLI(Rs("ROW"), -1)
   ROW2 = NVLI(Rs("ROW2"), -1)
   COLLUMN_CODE = NVLI(Rs("COLLUMN_CODE"), -1)
   COLLUMN_DEBIT = NVLI(Rs("COLLUMN_DEBIT"), -1)
   COLLUMN_CREDIT = NVLI(Rs("COLLUMN_CREDIT"), -1)
   COLLUMN_CHECK = NVLI(Rs("COLLUMN_CHECK"), -1)
   FROM_DATE = InternalDateToDate(NVLS(Rs("FROM_DATE"), ""))
   TO_DATE = InternalDateToDate(NVLS(Rs("TO_DATE"), ""))
End Function
Public Function AddEditData()
Dim ServerDateTime As String
Dim SQL1 As String
Dim ID As Long

   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   
   If AddEditMode = SHOW_ADD Then
      SQL1 = "INSERT INTO ACCOUNT_CONFIG " & _
                     "( " & _
                     "   SHEET, " & _
                     "   ROW, " & _
                     "   ROW2, " & _
                     "   COLLUMN_CODE, " & _
                     "   COLLUMN_DEBIT, " & _
                     "   COLLUMN_CREDIT, " & _
                     "   COLLUMN_CHECK, " & _
                     "   FROM_DATE, " & _
                     "   TO_DATE " & _
                     ") " & _
                     "VALUES " & _
                     "( " & _
                        SHEET & "," & _
                         ROW & "," & _
                         ROW2 & "," & _
                         COLLUMN_CODE & "," & _
                         COLLUMN_DEBIT & "," & _
                         COLLUMN_CREDIT & "," & _
                         COLLUMN_CHECK & "," & _
                         "'" & ChangeQuote(Trim(DateToStringInt(FROM_DATE))) & "'," & _
                         "'" & ChangeQuote(Trim(DateToStringInt(TO_DATE))) & "' " & _
                     ") "
         ACCOUNT_CONFIG_ID = ID
   Else 'Show edit
      SQL1 = "UPDATE ACCOUNT_CONFIG SET " & _
                     "   SHEET = " & SHEET & ", " & _
                     "   ROW = " & ROW & ", " & _
                     "   ROW2 = " & ROW2 & ", " & _
                     "   COLLUMN_CODE = " & COLLUMN_CODE & ", " & _
                     "   COLLUMN_DEBIT = " & COLLUMN_DEBIT & ", " & _
                     "   COLLUMN_CREDIT = " & COLLUMN_CREDIT & ", " & _
                     "   COLLUMN_CHECK = " & COLLUMN_CHECK & ", " & _
                     "   FROM_DATE = '" & Trim(ChangeQuote(DateToStringInt(FROM_DATE))) & "', " & _
                     "   TO_DATE = '" & Trim(ChangeQuote(DateToStringInt(TO_DATE))) & "' " & _
                     "WHERE ACCOUNT_CONFIG_ID = " & ACCOUNT_CONFIG_ID
   End If
   
   If (AddEditMode = SHOW_ADD) Or (AddEditMode = SHOW_EDIT) Then
      m_Conn.Execute (SQL1)
   End If
End Function

Public Function DeleteData()
Dim SQL1 As String
      SQL1 = "DELETE FROM ACCOUNT_CONFIG WHERE ACCOUNT_CONFIG_ID = " & ACCOUNT_CONFIG_ID
      m_Conn.Execute (SQL1)
End Function
