VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CDocumentCancel"
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

Public ORDER_BY As Long
Public ORDER_TYPE  As Long

Public KEY_ID As Long
Public KEY_LOOKUP As String
Private Sub Class_Initialize()
   Set m_Conn = glbDatabaseMngr.DBLegacyConnection
End Sub
'In this class we will not handle any error or any transaction
Public Function QueryData(Ind As Long, Rs As ADODB.Recordset, ItemCount As Long, Optional OwnerEnterprise As Boolean = True)
Dim SQL1 As String
Dim SQL2 As String
Dim SelectStr1 As String
Dim SelectStr2 As String
Dim FromStr As String
Dim WhereStr As String
Dim OrderByStr As String
Dim OrderTypeStr As String
   
   If Ind = 1 Then
      SelectStr1 = "SELECT UG.*  "
      SelectStr2 = "SELECT COUNT(*)  "
      
      FromStr = "FROM DOCUMENT_CANCEL UG "
   End If
   
   WhereStr = ""
   
   If OwnerEnterprise Then
      ENTERPRISE_CODE = glbCompanyCode
   End If
   
   If Len(ENTERPRISE_CODE) > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (ENTERPRISE_CODE = '" & ChangeQuote(Trim(ENTERPRISE_CODE)) & "') "
      Else
         WhereStr = WhereStr & "AND (ENTERPRISE_CODE = '" & ChangeQuote(Trim(ENTERPRISE_CODE)) & "') "
      End If
   End If
   
   OrderByStr = ""
   OrderTypeStr = ""

   If OrderBy = 1 Then
      OrderByStr = "ORDER BY DOCUMENT_NO "
   ElseIf OrderBy = 2 Then
      
   Else
      OrderByStr = "ORDER BY DOCUMENT_NO "
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
   ENTERPRISE_CODE = NVLS(Rs("ENTERPRISE_CODE"), "")
End Function
Public Function AddEditData()
Dim ServerDateTime As String
Dim SQL1 As String

   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   
   ENTERPRISE_CODE = glbCompanyCode
   
   If AddEditMode = SHOW_ADD Then
      SQL1 = "INSERT INTO DOCUMENT_CANCEL " & _
                     "( " & _
                     "   ENTERPRISE_CODE, " & _
                     "   DOCUMENT_NO " & _
                     ") " & _
                     "VALUES " & _
                     "( " & _
                        "'" & ChangeQuote(Trim(ENTERPRISE_CODE)) & "'," & _
                        "'" & ChangeQuote(Trim(DOCUMENT_NO)) & "'" & _
                     ") "
         ID = ID
   Else 'Show edit
      SQL1 = "UPDATE DOCUMENT_CANCEL SET " & _
                     "   ENTERPRISE_CODE = '" & ChangeQuote(Trim(ENTERPRISE_CODE)) & "', " & _
                     "   DOCUMENT_NO = '" & ChangeQuote(Trim(DOCUMENT_NO)) & "' " & _
                     "WHERE ID = " & ID
   End If
   
   If (AddEditMode = SHOW_ADD) Or (AddEditMode = SHOW_EDIT) Then
      m_Conn.Execute (SQL1)
   End If
End Function

Public Function DeleteData()
Dim SQL1 As String
      SQL1 = "DELETE FROM DOCUMENT_CANCEL WHERE ID = " & ID
      m_Conn.Execute (SQL1)
End Function
