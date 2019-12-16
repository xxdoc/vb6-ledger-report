VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CBankCredit"
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

Public BANK_NUMBER As Long
Public BANK_ID As Long
Public BANK_NAME  As String
Public BANK_AMOUNT As String
Public BANK_INTEREST As String
Public BANK_FEE_AMOUNT As String
Public BANK_FEE_TYPE As Integer
Public BANK_DATE_BROUGHT As Date
Public BANK_AMOUNT_BROUGHT As String
Public BANK_GET_AMOUNT As String
Public CUSTOMER_ID As Long
Public CUSTOMER_NAME  As String
Public FROM_DATE As Date
Public TO_DATE As Date

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
Dim SubLen As Long
Dim NewStr As String
Dim TempDate As String
   
   If Ind = 1 Then
         SelectStr1 = "SELECT CU.*  "
         SelectStr1 = SelectStr1 & ",CT.BANK_NAME "
         SelectStr1 = SelectStr1 & ",BC.CUSTOMER_NAME "
         
         SelectStr2 = "SELECT COUNT(*)  "
         
         FromStr = "FROM (BANK_CREDIT CU "
         FromStr = FromStr & " LEFT OUTER JOIN BANK  CT ON (CU.BANK_ID = CT.BANK_ID)) "
         FromStr = FromStr & " LEFT OUTER JOIN BANK_CUSTOMER  BC ON (CU.CUSTOMER_ID = BC.CUSTOMER_ID) "
         
         WhereStr = ""
   ElseIf Ind = 2 Then
         SelectStr1 = "SELECT DISTINCT CU.BANK_ID,CU.CUSTOMER_ID,CU.BANK_GET_AMOUNT  "
         SelectStr1 = SelectStr1 & ",CT.BANK_NAME "
         SelectStr1 = SelectStr1 & ",BC.CUSTOMER_NAME "
         
         SelectStr2 = "SELECT COUNT(*)  "
         
         FromStr = "FROM (BANK_CREDIT CU "
         FromStr = FromStr & " LEFT OUTER JOIN BANK  CT ON (CU.BANK_ID = CT.BANK_ID)) "
         FromStr = FromStr & " LEFT OUTER JOIN BANK_CUSTOMER  BC ON (CU.CUSTOMER_ID = BC.CUSTOMER_ID) "
         
         WhereStr = ""
   End If
   
   If BANK_NUMBER > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (BANK_NUMBER= " & BANK_NUMBER & ") "
      Else
         WhereStr = WhereStr & "AND (BANK_NUMBER = " & BANK_NUMBER & ") "
      End If
   End If
   
   If BANK_ID > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (CU.BANK_ID= " & BANK_ID & ") "
      Else
         WhereStr = WhereStr & "AND (CU.BANK_ID = " & BANK_ID & ") "
      End If
   End If
   
   If CUSTOMER_ID > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (CU.CUSTOMER_ID= " & CUSTOMER_ID & ") "
      Else
         WhereStr = WhereStr & "AND (CU.CUSTOMER_ID = " & CUSTOMER_ID & ") "
      End If
   End If
   
   If FROM_DATE > 0 Then
      TempDate = DateToStringIntLow(Trim(FROM_DATE))
      If WhereStr = "" Then
         WhereStr = " WHERE (BANK_DATE_BROUGHT >=  '" & ChangeQuote(TempDate) & "')"
      Else
         WhereStr = WhereStr & " AND (BANK_DATE_BROUGHT>= '" & ChangeQuote(TempDate) & "')"
      End If
   End If
   
   If TO_DATE > 0 Then
      TempDate = DateToStringIntHi(Trim(TO_DATE))
      If WhereStr = "" Then
         WhereStr = " WHERE (BANK_DATE_BROUGHT<= '" & ChangeQuote(TempDate) & "')"
      Else
         WhereStr = WhereStr & " AND (BANK_DATE_BROUGHT<= '" & ChangeQuote(TempDate) & "')"
      End If
   End If
   
   OrderByStr = ""
   OrderTypeStr = ""
   If Ind = 1 Then
      If ORDER_BY = 1 Then
         OrderByStr = " ORDER BY BANK_DATE_BROUGHT "
      ElseIf ORDER_BY = 2 Then
         OrderByStr = " ORDER BY CU.BANK_ID "
      ElseIf ORDER_BY = 3 Then
         OrderByStr = " ORDER BY CU.CUSTOMER_ID "
      Else
         OrderByStr = " ORDER BY BANK_DATE_BROUGHT "
      End If
   Else
      OrderByStr = " ORDER BY CU.BANK_ID "
   End If
   If ORDER_TYPE = 1 Then
      OrderTypeStr = "ASC "
   ElseIf ORDER_TYPE = 2 Then
      OrderTypeStr = "DESC "
   Else
      OrderTypeStr = "ASC "
   End If
   
   SQL1 = SelectStr1 & FromStr & WhereStr & OrderByStr & OrderTypeStr
   SQL2 = SelectStr2 & FromStr & WhereStr
   
   Call glbDatabaseMngr.GetLegacyRs(SQL1, SQL2, True, ItemCount, Rs, glbErrorLog)
End Function

Public Function PopulateFromRS(Ind As Long, Rs As ADODB.Recordset)
   If Ind = 1 Then
      BANK_NUMBER = NVLI(Rs("BANK_Number"), -1)
      BANK_ID = NVLI(Rs("BANK_ID"), -1)
      BANK_NAME = NVLS(Rs("BANK_NAME"), "")
      BANK_AMOUNT = NVLS(Rs("BANK_AMOUNT"), "")
      BANK_INTEREST = NVLS(Rs("BANK_INTEREST"), "")
      BANK_FEE_AMOUNT = NVLS(Rs("BANK_FEE_AMOUNT"), "")
      BANK_FEE_TYPE = NVLI(Rs("BANK_FEE_TYPE"), -1)
      BANK_DATE_BROUGHT = InternalDateToDate(NVLS(Rs("BANK_DATE_BROUGHT"), ""))
      BANK_AMOUNT_BROUGHT = NVLS(Rs("BANK_AMOUNT_BROUGHT"), "")
      BANK_GET_AMOUNT = NVLS(Rs("BANK_GET_AMOUNT"), "")
      CUSTOMER_ID = NVLI(Rs("CUSTOMER_ID"), -1)
      CUSTOMER_NAME = NVLS(Rs("CUSTOMER_NAME"), "")
   ElseIf Ind = 2 Then
      BANK_ID = NVLI(Rs("BANK_ID"), -1)
      BANK_NAME = NVLS(Rs("BANK_NAME"), "")
      CUSTOMER_ID = NVLI(Rs("CUSTOMER_ID"), -1)
      CUSTOMER_NAME = NVLS(Rs("CUSTOMER_NAME"), "")
      BANK_GET_AMOUNT = NVLS(Rs("BANK_GET_AMOUNT"), "")
   End If

End Function
Public Function AddEditData()
Dim ServerDateTime As String
Dim SQL1 As String
   
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   
   If AddEditMode = SHOW_ADD Then
      SQL1 = "INSERT INTO BANK_CREDIT " & _
                     "( " & _
                     "   BANK_ID, " & _
                     "   BANK_AMOUNT, " & _
                     "   BANK_INTEREST, " & _
                     "   BANK_FEE_AMOUNT, " & _
                     "   BANK_FEE_TYPE, " & _
                     "   BANK_DATE_BROUGHT, " & _
                     "   BANK_AMOUNT_BROUGHT, " & _
                     "   BANK_GET_AMOUNT, " & _
                      "   CUSTOMER_ID " & _
                     ") " & _
                     "VALUES " & _
                     "( " & _
                       "'" & BANK_ID & "'," & _
                        "'" & ChangeQuote(Trim(BANK_AMOUNT)) & "'," & _
                        "'" & ChangeQuote(Trim(BANK_INTEREST)) & "'," & _
                        "'" & ChangeQuote(Trim(BANK_FEE_AMOUNT)) & "'," & _
                        "'" & BANK_FEE_TYPE & "'," & _
                        "'" & ChangeQuote(Trim(DateToStringInt(BANK_DATE_BROUGHT))) & "'," & _
                        "'" & ChangeQuote(Trim(BANK_AMOUNT_BROUGHT)) & "'," & _
                        "'" & ChangeQuote(Trim(BANK_GET_AMOUNT)) & "'," & _
                        "'" & CUSTOMER_ID & "'" & _
                     " ) "
   Else 'Show edit
      SQL1 = "UPDATE BANK_CREDIT SET " & _
                     "   BANK_ID = " & BANK_ID & ", " & _
                     "   BANK_AMOUNT = '" & ChangeQuote(Trim(BANK_AMOUNT)) & "' " & ", " & _
                     "   BANK_INTEREST = '" & ChangeQuote(Trim(BANK_INTEREST)) & "' " & ", " & _
                     "   BANK_FEE_AMOUNT = '" & ChangeQuote(Trim(BANK_FEE_AMOUNT)) & "' " & ", " & _
                       "   BANK_FEE_TYPE = " & BANK_FEE_TYPE & ", " & _
                      "   BANK_DATE_BROUGHT = '" & ChangeQuote(Trim(DateToStringInt(BANK_DATE_BROUGHT))) & "', " & _
                      "   BANK_AMOUNT_BROUGHT = '" & ChangeQuote(Trim(BANK_AMOUNT_BROUGHT)) & "' " & ", " & _
                      "   BANK_GET_AMOUNT = '" & ChangeQuote(Trim(BANK_GET_AMOUNT)) & "' " & ", " & _
                     "   CUSTOMER_ID = " & CUSTOMER_ID & _
                     " WHERE BANK_NUMBER = " & BANK_NUMBER
   End If

   If (AddEditMode = SHOW_ADD) Or (AddEditMode = SHOW_EDIT) Then
      m_Conn.Execute (SQL1)
   End If
End Function
Public Function DeleteData()
Dim SQL1 As String
      SQL1 = "DELETE FROM BANK_CREDIT WHERE BANK_NUMBER = " & BANK_NUMBER
      m_Conn.Execute (SQL1)
End Function
