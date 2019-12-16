VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CCustomer"
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

Public CUSTOMER_ID As Long
Public CUSTOMER_CODE  As String
Public CUSTOMER_NAME As String
Public CUSTOMER_TYPE_ID As Long
Public CUSTOMER_TYPE_NAME As String

Public ORDER_BY As Long
Public ORDER_TYPE  As Long

Public KEY_ID As Long
Public KEY_LOOKUP As String

Private Sub Class_Initialize()
   Set m_Conn = glbDatabaseMngr.DBLegacyConnection
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
Dim SubLen As Long
Dim NewStr As String

   SelectStr1 = "SELECT CU.*  "
   SelectStr1 = SelectStr1 & ",CT.CUSTOMER_TYPE_NAME "
   
   
   SelectStr2 = "SELECT COUNT(*)  "
   
   FromStr = "FROM (CUSTOMER CU "
   FromStr = FromStr & " LEFT OUTER JOIN CUSTOMER_TYPE  CT ON (CU.CUSTOMER_TYPE_ID = CT.CUSTOMER_TYPE_ID)) "
   
   WhereStr = ""
'   If OwnerEnterprise Then
'      ENTERPRISE_CODE = glbCompanyCode
'   End If
'
'   If Len(ENTERPRISE_CODE) > 0 Then
'      If Len(WhereStr) <= 0 Then
'         WhereStr = "WHERE (ENTERPRISE_CODE = '" & ChangeQuote(Trim(ENTERPRISE_CODE)) & "') "
'      Else
'         WhereStr = WhereStr & "AND (ENTERPRISE_CODE = '" & ChangeQuote(Trim(ENTERPRISE_CODE)) & "') "
'      End If
'   End If
   
   If CUSTOMER_ID > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (CUSTOMER_ID= " & CUSTOMER_ID & ") "
      Else
         WhereStr = WhereStr & "AND (CUSTOMER_ID = " & CUSTOMER_ID & ") "
      End If
   End If
   
   If Trim(CUSTOMER_CODE) <> "" Then
      If WildCard(CUSTOMER_CODE, SubLen, NewStr) Then
         If WhereStr = "" Then
            WhereStr = "WHERE (CUSTOMER_CODE LIKE '" & ChangeQuote(Trim(CUSTOMER_CODE)) & "')"
         Else
            WhereStr = WhereStr & "AND (CUSTOMER_CODE LIKE '" & ChangeQuote(Trim(CUSTOMER_CODE)) & "')"
         End If
      Else
         If WhereStr = "" Then
            WhereStr = "WHERE (CUSTOMER_CODE = '" & ChangeQuote(Trim(CUSTOMER_CODE)) & "')"
         Else
            WhereStr = WhereStr & "AND (CUSTOMER_CODE = '" & ChangeQuote(Trim(CUSTOMER_CODE)) & "')"
         End If
      End If
   End If
   
     If Trim(CUSTOMER_NAME) <> "" Then
      If WildCard(CUSTOMER_NAME, SubLen, NewStr) Then
         If WhereStr = "" Then
            WhereStr = "WHERE (CUSTOMER_NAME LIKE '" & ChangeQuote(Trim(CUSTOMER_NAME)) & "')"
         Else
            WhereStr = WhereStr & "AND (CUSTOMER_NAME LIKE '" & ChangeQuote(Trim(CUSTOMER_NAME)) & "')"
         End If
      Else
         If WhereStr = "" Then
            WhereStr = "WHERE (CUSTOMER_NAME = '" & ChangeQuote(Trim(CUSTOMER_NAME)) & "')"
         Else
            WhereStr = WhereStr & "AND (CUSTOMER_NAME = '" & ChangeQuote(Trim(CUSTOMER_NAME)) & "')"
         End If
      End If
   End If
   
   If CUSTOMER_TYPE_ID > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (CU.CUSTOMER_TYPE_ID= " & CUSTOMER_TYPE_ID & ") "
      Else
         WhereStr = WhereStr & "AND (CU.CUSTOMER_TYPE_ID = " & CUSTOMER_TYPE_ID & ") "
      End If
   End If
   
   OrderByStr = ""
   OrderTypeStr = ""
   
   If ORDER_BY = 1 Then
      OrderByStr = "ORDER BY CUSTOMER_CODE "
   ElseIf ORDER_BY = 2 Then
      OrderByStr = "ORDER BY CU.CUSTOMER_NAME ASC,CU.CUSTOMER_TYPE_ID ASC,CU.CUSTOMER_CODE "
   Else
      OrderByStr = "ORDER BY CUSTOMER_CODE "
   End If
   If ORDER_TYPE = 1 Then
      OrderTypeStr = "ASC "
   ElseIf OrderType = 2 Then
      OrderTypeStr = "DESC "
   Else
      OrderTypeStr = "ASC "
   End If
   
   SQL1 = SelectStr1 & FromStr & WhereStr & OrderByStr & OrderTypeStr
   SQL2 = SelectStr2 & FromStr & WhereStr
   
   Call glbDatabaseMngr.GetLegacyRs(SQL1, SQL2, True, ItemCount, Rs, glbErrorLog)
End Function

Public Function PopulateFromRS(Ind As Long, Rs As ADODB.Recordset)
    CUSTOMER_ID = NVLI(Rs("CUSTOMER_ID"), -1)
   CUSTOMER_CODE = NVLS(Rs("CUSTOMER_CODE"), "")
   CUSTOMER_NAME = NVLS(Rs("CUSTOMER_NAME"), "")
   CUSTOMER_TYPE_ID = NVLI(Rs("CUSTOMER_TYPE_ID"), -1)
   CUSTOMER_TYPE_NAME = NVLS(Rs("CUSTOMER_TYPE_NAME"), "")

End Function
Public Function AddEditData()
Dim ServerDateTime As String
Dim SQL1 As String
   
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   
   If AddEditMode = SHOW_ADD Then
      SQL1 = "INSERT INTO CUSTOMER " & _
                     "( " & _
                     "   CUSTOMER_CODE, " & _
                     "   CUSTOMER_NAME, " & _
                     "   CUSTOMER_TYPE_ID " & _
                     ") " & _
                     "VALUES " & _
                     "( " & _
                       "'" & ChangeQuote(Trim(CUSTOMER_CODE)) & "'" & "," & _
                        "'" & ChangeQuote(Trim(CUSTOMER_NAME)) & "'," & _
                         "" & CUSTOMER_TYPE_ID & "" & _
                     ") "
   Else 'Show edit
      SQL1 = "UPDATE CUSTOMER SET " & _
                     "   CUSTOMER_CODE = '" & ChangeQuote(Trim(CUSTOMER_CODE)) & "' " & ", " & _
                     "   CUSTOMER_NAME = '" & ChangeQuote(Trim(CUSTOMER_NAME)) & "' " & ", " & _
                     "   CUSTOMER_TYPE_ID = " & CUSTOMER_TYPE_ID & _
                     " WHERE CUSTOMER_ID = " & CUSTOMER_ID
   End If
   
   If (AddEditMode = SHOW_ADD) Or (AddEditMode = SHOW_EDIT) Then
      m_Conn.Execute (SQL1)
   End If
End Function
Public Function DeleteData()
Dim SQL1 As String
      SQL1 = "DELETE FROM CUSTOMER WHERE CUSTOMER_ID = " & CUSTOMER_ID
      m_Conn.Execute (SQL1)
End Function
