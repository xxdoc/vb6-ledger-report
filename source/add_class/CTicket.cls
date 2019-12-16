VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CTicket"
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

Public TICKET_ID As Long
Public TICKET_DATE As Date
Public TICKET_NUMBER As String
Public BANK_DATE_BROUGHT As Date
Public TICKET_INVOICE As String
Public TICKET_DATE_DUE As Date
Public TICKET_DATE_CHECK As Date
Public TICKET_DATE_GET_CHECK As Date
Public TICKET_DATE_NEW As Date
Public CUSTOMER_ID As Long
Public CUSTOMER_NAME  As String
Public FROM_DATE As Date
Public TO_DATE As Date
Public FROM_DATE_RECEIVE As Date
Public TO_DATE_RECEIVE As Date
Public TEMP_TYPE As String
Public TEMP_CHECK As String
Public TEMP_COUNT As Long
Public TICKET_AMOUNT As Double
Public TICKET_INTEREST As String
Public CHECK_RECEIVED As Long
Public BANK_ID As Long
Public BANK_NAME  As String
Public MASTER_AREA As Long

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
Dim GroupByStr As String
Dim OrderByStr As String
Dim OrderTypeStr As String
Dim SubLen As Long
Dim NewStr As String
Dim TempDate As String
   If Ind = 1 Then
      SelectStr1 = "SELECT CU.*  "
      SelectStr1 = SelectStr1 & ",CT.CUSTOMER_NAME , BA.BANK_NAME  "
      
      SelectStr2 = "SELECT COUNT(*)  "
      
      FromStr = "FROM (TICKET CU "
      FromStr = FromStr & " LEFT OUTER JOIN BANK_CUSTOMER  CT ON (CU.CUSTOMER_ID = CT.CUSTOMER_ID)) "
      FromStr = FromStr & " LEFT OUTER JOIN BANK  BA ON (CU.BANK_ID = BA.BANK_ID) "
      
      WhereStr = ""
   ElseIf Ind = 2 Then
      SelectStr1 = "SELECT CU.TICKET_INVOICE, SUM(CU.TICKET_AMOUNT) AS TICKET_AMOUNT  "
      SelectStr2 = "SELECT COUNT(*)  "
   
      FromStr = "FROM TICKET CU "
      WhereStr = " WHERE CU.TICKET_AMOUNT > 0 "
   End If
   
   If TICKET_ID > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (TICKET_ID= " & TICKET_ID & ") "
      Else
         WhereStr = WhereStr & "AND (TICKET_ID = " & TICKET_ID & ") "
      End If
   End If
   
   If CUSTOMER_ID > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (CU.CUSTOMER_ID= " & CUSTOMER_ID & ") "
      Else
         WhereStr = WhereStr & "AND (CU.CUSTOMER_ID = " & CUSTOMER_ID & ") "
      End If
   End If
   
   If BANK_ID > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (CU.BANK_ID= " & BANK_ID & ") "
      Else
         WhereStr = WhereStr & "AND (CU.BANK_ID = " & BANK_ID & ") "
      End If
   End If
   
   If CHECK_RECEIVED > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (CHECK_RECEIVED= " & CHECK_RECEIVED & ") "
      Else
         WhereStr = WhereStr & "AND (CHECK_RECEIVED = " & CHECK_RECEIVED & ") "
      End If
   End If
   
     If Trim(TICKET_NUMBER) <> "" Then
      If WildCard(TICKET_NUMBER, SubLen, NewStr) Then
         If WhereStr = "" Then
            WhereStr = "WHERE (TICKET_NUMBER LIKE '" & ChangeQuote(Trim(TICKET_NUMBER)) & "')"
         Else
            WhereStr = WhereStr & "AND (TICKET_NUMBER LIKE '" & ChangeQuote(Trim(TICKET_NUMBER)) & "')"
         End If
      Else
         If WhereStr = "" Then
            WhereStr = "WHERE (TICKET_NUMBER = '" & ChangeQuote(Trim(TICKET_NUMBER)) & "')"
         Else
            WhereStr = WhereStr & "AND (TICKET_NUMBER = '" & ChangeQuote(Trim(TICKET_NUMBER)) & "')"
         End If
      End If
   End If
   
   If Trim(TICKET_INVOICE) <> "" Then
      If WildCard(TICKET_INVOICE, SubLen, NewStr) Then
         If WhereStr = "" Then
            WhereStr = "WHERE (TICKET_INVOICE LIKE '" & ChangeQuote(Trim(TICKET_INVOICE)) & "')"
         Else
            WhereStr = WhereStr & "AND (TICKET_INVOICE LIKE '" & ChangeQuote(Trim(TICKET_INVOICE)) & "')"
         End If
      Else
         If WhereStr = "" Then
            WhereStr = "WHERE (TICKET_INVOICE = '" & ChangeQuote(Trim(TICKET_INVOICE)) & "')"
         Else
            WhereStr = WhereStr & "AND (TICKET_INVOICE = '" & ChangeQuote(Trim(TICKET_INVOICE)) & "')"
         End If
      End If
   End If
   
   If FROM_DATE > 0 Then
      TempDate = DateToStringIntLow(Trim(FROM_DATE))
      If WhereStr = "" Then
         WhereStr = " WHERE (TICKET_DATE >=  '" & ChangeQuote(TempDate) & "')"
      Else
         WhereStr = WhereStr & " AND (TICKET_DATE>= '" & ChangeQuote(TempDate) & "')"
      End If
   End If

   If TO_DATE > 0 Then
      TempDate = DateToStringIntHi(Trim(TO_DATE))
      If WhereStr = "" Then
         WhereStr = " WHERE (TICKET_DATE<= '" & ChangeQuote(TempDate) & "')"
      Else
         WhereStr = WhereStr & " AND (TICKET_DATE<= '" & ChangeQuote(TempDate) & "')"
      End If
   End If
   
   If FROM_DATE_RECEIVE > 0 Then
      TempDate = DateToStringIntLow(Trim(FROM_DATE_RECEIVE))
      If WhereStr = "" Then
         WhereStr = " WHERE (TICKET_DATE_CHECK >=  '" & ChangeQuote(TempDate) & "')"
      Else
         WhereStr = WhereStr & " AND (TICKET_DATE_CHECK>= '" & ChangeQuote(TempDate) & "')"
      End If
   End If

   If TO_DATE_RECEIVE > 0 Then
      TempDate = DateToStringIntHi(Trim(TO_DATE_RECEIVE))
      If WhereStr = "" Then
         WhereStr = " WHERE (TICKET_DATE_CHECK<= '" & ChangeQuote(TempDate) & "')"
      Else
         WhereStr = WhereStr & " AND (TICKET_DATE_CHECK<= '" & ChangeQuote(TempDate) & "')"
      End If
   End If
   
   If MASTER_AREA > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (MASTER_AREA= " & MASTER_AREA & ") "
      Else
         WhereStr = WhereStr & "AND (MASTER_AREA = " & MASTER_AREA & ") "
      End If
   End If
   
   GroupByStr = ""
   If Ind = 2 Then
      GroupByStr = " GROUP BY CU.TICKET_INVOICE "
   End If
   
   OrderByStr = ""
   OrderTypeStr = ""
   If Ind = 1 Then
      If ORDER_BY = 1 Then
         OrderByStr = " ORDER BY TICKET_NUMBER, TICKET_INVOICE, TICKET_DATE, TICKET_DATE_DUE, TICKET_DATE_CHECK, TICKET_DATE_GET_CHECK "
      ElseIf ORDER_BY = 2 Then
         OrderByStr = " ORDER BY CU.TICKET_DATE "
      ElseIf ORDER_BY = 3 Then
         OrderByStr = " ORDER BY CU.CUSTOMER_ID "
      Else
         OrderByStr = " ORDER BY TICKET_NUMBER, TICKET_INVOICE, TICKET_DATE, TICKET_DATE_DUE, TICKET_DATE_CHECK, TICKET_DATE_GET_CHECK  "
      End If
      
      If ORDER_TYPE = 1 Then
         OrderTypeStr = "ASC "
      ElseIf ORDER_TYPE = 2 Then
         OrderTypeStr = "DESC "
      Else
         OrderTypeStr = "ASC "
      End If
    End If
   
   SQL1 = SelectStr1 & FromStr & WhereStr & GroupByStr & OrderByStr & OrderTypeStr
   SQL2 = SelectStr2 & FromStr & WhereStr & GroupByStr
   
   Call glbDatabaseMngr.GetLegacyRs(SQL1, SQL2, True, ItemCount, Rs, glbErrorLog)
End Function

Public Function PopulateFromRS(Ind As Long, Rs As ADODB.Recordset)
   If Ind = 1 Then
      TICKET_ID = NVLI(Rs("TICKET_ID"), -1)
      TICKET_DATE = InternalDateToDate(NVLS(Rs("TICKET_DATE"), ""))
      TICKET_NUMBER = NVLS(Rs("TICKET_NUMBER"), "")
      TICKET_INVOICE = NVLS(Rs("TICKET_INVOICE"), "")
      TICKET_DATE_DUE = InternalDateToDate(NVLS(Rs("TICKET_DATE_DUE"), ""))
      TICKET_DATE_CHECK = InternalDateToDate(NVLS(Rs("TICKET_DATE_CHECK"), ""))
      TICKET_DATE_GET_CHECK = InternalDateToDate(NVLS(Rs("TICKET_DATE_GET_CHECK"), ""))
      CUSTOMER_ID = NVLI(Rs("CUSTOMER_ID"), -1)
      CUSTOMER_NAME = NVLS(Rs("CUSTOMER_NAME"), "")
      TICKET_AMOUNT = NVLD(Rs("TICKET_AMOUNT"), 0)
      TICKET_INTEREST = NVLS(Rs("TICKET_INTEREST"), "")
      TICKET_DATE_NEW = InternalDateToDate(NVLS(Rs("TICKET_DATE_NEW"), ""))
      CHECK_RECEIVED = NVLI(Rs("CHECK_RECEIVED"), -1)
      BANK_ID = NVLI(Rs("BANK_ID"), -1)
      BANK_NAME = NVLS(Rs("BANK_NAME"), "")
      MASTER_AREA = NVLI(Rs("MASTER_AREA"), -1)
   ElseIf Ind = 2 Then
      TICKET_INVOICE = NVLS(Rs("TICKET_INVOICE"), "")
      TICKET_AMOUNT = NVLD(Rs("TICKET_AMOUNT"), 0)
   End If
End Function
Public Function AddEditData()
Dim ServerDateTime As String
Dim SQL1 As String
   
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   
   If AddEditMode = SHOW_ADD Then
      SQL1 = "INSERT INTO TICKET " & _
                     "( " & _
                     "   TICKET_DATE, " & _
                     "   TICKET_NUMBER, " & _
                     "   TICKET_INVOICE, " & _
                     "   TICKET_DATE_DUE, " & _
                     "   TICKET_DATE_CHECK, " & _
                     "   TICKET_DATE_GET_CHECK, " & _
                      "   CUSTOMER_ID, " & _
                      "   TICKET_AMOUNT, " & _
                      "   TICKET_INTEREST, " & "   TICKET_DATE_NEW, " & _
                      "   CHECK_RECEIVED, " & "   BANK_ID, " & "   MASTER_AREA " & _
                     ") " & _
                     "VALUES " & _
                     "( " & _
                       "'" & ChangeQuote(Trim(DateToStringInt(TICKET_DATE))) & "'" & "," & _
                        "'" & ChangeQuote(Trim(TICKET_NUMBER)) & "'," & _
                        "'" & ChangeQuote(Trim(TICKET_INVOICE)) & "'," & _
                        "'" & ChangeQuote(Trim(DateToStringInt(TICKET_DATE_DUE))) & "'," & _
                        "'" & ChangeQuote(Trim(DateToStringInt(TICKET_DATE_CHECK))) & "'," & _
                         "'" & ChangeQuote(Trim(DateToStringInt(TICKET_DATE_GET_CHECK))) & "'," & _
                         "'" & CUSTOMER_ID & "'," & "'" & ChangeQuote(Trim(TICKET_AMOUNT)) & "'," & _
                        "'" & ChangeQuote(Trim(TICKET_INTEREST)) & "'," & "'" & ChangeQuote(Trim(DateToStringInt(TICKET_DATE_NEW))) & "'," & _
                        "'" & ChangeQuote(Trim(CHECK_RECEIVED)) & "'," & "'" & ChangeQuote(Trim(BANK_ID)) & "'," & "'" & ChangeQuote(Trim(MASTER_AREA)) & "'" & _
                     " ) "
   Else 'Show edit
      SQL1 = "UPDATE TICKET SET " & _
                     "   TICKET_DATE = '" & ChangeQuote(Trim(DateToStringInt(TICKET_DATE))) & "' " & ", " & _
                     "   TICKET_NUMBER = '" & ChangeQuote(Trim(TICKET_NUMBER)) & "' " & ", " & _
                     "   TICKET_INVOICE = '" & ChangeQuote(Trim(TICKET_INVOICE)) & "' " & ", " & _
                     "   TICKET_DATE_DUE = '" & ChangeQuote(Trim(DateToStringInt(TICKET_DATE_DUE))) & "', " & _
                     "   TICKET_DATE_CHECK = '" & ChangeQuote(Trim(DateToStringInt(TICKET_DATE_CHECK))) & "', " & _
                     "    TICKET_DATE_GET_CHECK =  '" & ChangeQuote(Trim(DateToStringInt(TICKET_DATE_GET_CHECK))) & "' " & ", " & _
                      "   CUSTOMER_ID = " & CUSTOMER_ID & ", " & _
                     "   TICKET_AMOUNT = '" & ChangeQuote(Trim(TICKET_AMOUNT)) & "' " & ", " & _
                     "   TICKET_INTEREST = '" & ChangeQuote(Trim(TICKET_INTEREST)) & "' " & ", " & _
                     "   TICKET_DATE_NEW = '" & ChangeQuote(Trim(DateToStringInt(TICKET_DATE_NEW))) & "'," & _
                     "   CHECK_RECEIVED = '" & ChangeQuote(Trim(CHECK_RECEIVED)) & "'," & _
                     "   BANK_ID = '" & ChangeQuote(Trim(BANK_ID)) & "'," & _
                     "   MASTER_AREA = '" & ChangeQuote(Trim(MASTER_AREA)) & "'" & _
                     " WHERE TICKET_ID = " & TICKET_ID
   End If

   If (AddEditMode = SHOW_ADD) Or (AddEditMode = SHOW_EDIT) Then
      m_Conn.Execute (SQL1)
   End If
End Function
Public Function DeleteData()
Dim SQL1 As String
      SQL1 = "DELETE FROM TICKET WHERE TICKET_ID = " & TICKET_ID
      m_Conn.Execute (SQL1)
End Function
