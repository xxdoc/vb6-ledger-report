VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CRealCredit"
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

Public ID As Long
Public ENTERPRISE_CODE As String
Public DOCUMENT_NO As String
Public REAL_CREDIT As Long
Public PAID_FLAG As Boolean

Public ORDER_BY As Long
Public ORDER_TYPE  As Long

Public CUSTOMER_CODE As String
Public CUSTOMER_NAME As String

Public KEY_ID As Long
Public KEY_LOOKUP As String
Private Sub Class_Initialize()
   Set m_Conn = glbDatabaseMngr.DBLegacyConnection '

End Sub

'In this class we will not handle any error or any transaction
Public Function QueryData(Rs As ADODB.Recordset, ItemCount As Long, Optional OwnerEnterprise As Boolean = True)
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
   
   FromStr = "FROM REAL_CREDIT UG "
   
   WhereStr = ""
   
   If OwnerEnterprise Then
      'ENTERPRISE_CODE = glbCompanyCode
   End If
   
   If ID > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (UG.ID= " & ID & ") "
      Else
         WhereStr = WhereStr & "AND (UG.ID = " & ID & ") "
      End If
   End If
   
   If Len(ENTERPRISE_CODE) > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (ENTERPRISE_CODE = '" & ChangeQuote(Trim(ENTERPRISE_CODE)) & "') "
      Else
         WhereStr = WhereStr & "AND (ENTERPRISE_CODE = '" & ChangeQuote(Trim(ENTERPRISE_CODE)) & "') "
      End If
   End If
   
   If Len(PAID_FLAG) > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (PAID_FLAG = " & ChangeQuote(Trim(PAID_FLAG)) & ") "
      Else
         WhereStr = WhereStr & "AND (PAID_FLAG = " & ChangeQuote(Trim(PAID_FLAG)) & ") "
      End If
   End If
   
   OrderByStr = ""
   OrderTypeStr = ""

   If OrderBy = 1 Then
      OrderByStr = "ORDER BY CUSTOMER_CODE,DOCUMENT_NO "
   ElseIf OrderBy = 2 Then
      
   Else
      OrderByStr = "ORDER BY CUSTOMER_CODE,DOCUMENT_NO "
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
   ID = NVLI(Rs("ID"), -1)
   DOCUMENT_NO = NVLS(Rs("DOCUMENT_NO"), "")
   REAL_CREDIT = NVLI(Rs("REAL_CREDIT"), 0)
   PAID_FLAG = NVLS(Rs("PAID_FLAG"), "")
   ENTERPRISE_CODE = NVLS(Rs("ENTERPRISE_CODE"), "")
   
   CUSTOMER_CODE = NVLS(Rs("CUSTOMER_CODE"), "")
   CUSTOMER_NAME = NVLS(Rs("CUSTOMER_NAME"), "")
End Function
Public Function AddEditData()
Dim ServerDateTime As String
Dim SQL1 As String

   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   
   ENTERPRISE_CODE = glbCompanyCode
   
   If AddEditMode = SHOW_ADD Then
      SQL1 = "INSERT INTO REAL_CREDIT " & _
                     "( " & _
                     "   ENTERPRISE_CODE, " & _
                     "   DOCUMENT_NO, " & _
                     "   REAL_CREDIT, " & _
                     "   PAID_FLAG ," & _
                     "   CUSTOMER_CODE ," & _
                     "   CUSTOMER_NAME " & _
                     ") " & _
                     "VALUES " & _
                     "( " & _
                        "'" & ChangeQuote(Trim(ENTERPRISE_CODE)) & "'," & _
                        "'" & ChangeQuote(Trim(DOCUMENT_NO)) & "'," & _
                        "" & REAL_CREDIT & "," & _
                         "" & ChangeQuote(Trim(PAID_FLAG)) & "," & _
                         "'" & ChangeQuote(Trim(CUSTOMER_CODE)) & "'," & _
                         "'" & ChangeQuote(Trim(CUSTOMER_NAME)) & "'" & _
                     ") "
         ID = ID
   Else 'Show edit
      SQL1 = "UPDATE REAL_CREDIT SET " & _
                     "   ENTERPRISE_CODE = '" & ChangeQuote(Trim(ENTERPRISE_CODE)) & "', " & _
                     "   DOCUMENT_NO = '" & ChangeQuote(Trim(DOCUMENT_NO)) & "', " & _
                     "   REAL_CREDIT = " & REAL_CREDIT & ", " & _
                     "   PAID_FLAG = " & ChangeQuote(Trim(PAID_FLAG)) & ", " & _
                     "   CUSTOMER_CODE = '" & ChangeQuote(Trim(CUSTOMER_CODE)) & "', " & _
                     "   CUSTOMER_NAME = '" & ChangeQuote(Trim(CUSTOMER_NAME)) & "' " & _
                     "WHERE ID = " & ID
   End If
   
   If (AddEditMode = SHOW_ADD) Or (AddEditMode = SHOW_EDIT) Then
      m_Conn.Execute (SQL1)
   End If
End Function

Public Function DeleteData()
Dim SQL1 As String
      SQL1 = "DELETE FROM REAL_CREDIT WHERE ID = " & ID
      m_Conn.Execute (SQL1)
End Function
