VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CPromotionPayCustom"
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

Public PRO_ID As Long
Public PRO_DATE As String
Public SALECODE_PRO As String
Public SALENAME_PRO As String
Public CUSTOMERCODE_PRO As String
Public CUSTOMERNAME_PRO As String
Public STKCOD_PRO As String
Public STKNAME_PRO As String
Public AMOUNT_PRO As Double
Public YYYY_MM As String

Public FROM_PRO_DATE As Date
Public TO_PRO_DATE As Date

Public FROM_SALE_CODE As String
Public TO_SALE_CODE As String
Public FROM_CUSTOMER_CODE As String
Public TO_CUSTOMER_CODE As String
Public FROM_STOCK_CODE As String
Public TO_STOCK_CODE As String

Public KEY_ID As Long
Public KEY_LOOKUP As String

Private Sub Class_Initialize()
   Set m_Conn = glbDatabaseMngr.DBLegacyConnection
   
   FROM_PRO_DATE = -1
   TO_PRO_DATE = -1
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
Dim GroupByStr As String
Dim TempDate As String
Dim SubLen As Long
Dim NewStr As String

   If Ind = 1 Then
      SelectStr1 = "SELECT PYC.* "
      SelectStr2 = "SELECT COUNT(*)  "
      
      FromStr = "FROM PROMOTION_PAYCUSTOMER PYC "
   
   ElseIf Ind = 2 Then
      SelectStr1 = "SELECT DISTINCT PYC.SALECODE_PRO,PYC.SALENAME_PRO,PYC.CUSTOMERCODE_PRO,PYC.CUSTOMERNAME_PRO,PYC.STKCOD_PRO,PYC.STKNAME_PRO " 'รหัสและชื่อ Sale ลูกค้า สินค้า
      SelectStr2 = "SELECT COUNT(*)  "
      
      FromStr = "FROM PROMOTION_PAYCUSTOMER PYC "
      
   ElseIf Ind = 3 Then
      SelectStr1 = "SELECT SUM(PYC.AMOUNT_PRO) AS AMOUNT_PRO, PYC.SALECODE_PRO,PYC.CUSTOMERCODE_PRO,PYC.STKCOD_PRO,PYC.YYYY_MM "
      SelectStr2 = "SELECT COUNT(*)  "
      
      FromStr = "FROM PROMOTION_PAYCUSTOMER PYC "
   
   ElseIf Ind = 4 Then
      SelectStr1 = "SELECT DISTINCT PYC.STKCOD_PRO,PYC.STKNAME_PRO,PYC.CUSTOMERCODE_PRO,PYC.CUSTOMERNAME_PRO " 'รหัสและชื่อ Sale ลูกค้า สินค้า
      SelectStr2 = "SELECT COUNT(*)  "
      
      FromStr = "FROM PROMOTION_PAYCUSTOMER PYC "
      
   ElseIf Ind = 5 Then
      SelectStr1 = "SELECT SUM(PYC.AMOUNT_PRO) AS AMOUNT_PRO,PYC.STKCOD_PRO,PYC.CUSTOMERCODE_PRO,PYC.YYYY_MM "
      SelectStr2 = "SELECT COUNT(*)  "
      
      FromStr = "FROM PROMOTION_PAYCUSTOMER PYC "
   End If
   
   If PRO_ID > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (PRO_ID= " & PRO_ID & ") "
      Else
         WhereStr = WhereStr & "AND (PRO_ID = " & PRO_ID & ") "
      End If
   End If
   
    If FROM_PRO_DATE > -1 Then
      TempDate = DateToString(FROM_PRO_DATE)
      If WhereStr = "" Then
         WhereStr = "WHERE ((PYC.PRO_DATE) >= '" & ChangeQuote(TempDate) & "')"
      Else
         WhereStr = WhereStr & "AND ((PYC.PRO_DATE) >= '" & ChangeQuote(TempDate) & "')"
      End If
   End If

    If TO_PRO_DATE > -1 Then
      TempDate = DateToString(TO_PRO_DATE)
      If WhereStr = "" Then
         WhereStr = "WHERE ((PYC.PRO_DATE) <= '" & ChangeQuote(TempDate) & "')"
      Else
         WhereStr = WhereStr & "AND ((PYC.PRO_DATE) <= '" & ChangeQuote(TempDate) & "')"
      End If
   End If
   
   If Trim(SALECODE_PRO) <> "" Then
      If WildCard(SALECODE_PRO, SubLen, NewStr) Then
         If WhereStr = "" Then
            WhereStr = "WHERE (SALECODE_PRO LIKE '" & ChangeQuote(Trim(SALECODE_PRO)) & "')"
         Else
            WhereStr = WhereStr & "AND (SALECODE_PRO LIKE '" & ChangeQuote(Trim(SALECODE_PRO)) & "')"
         End If
      Else
         If WhereStr = "" Then
            WhereStr = "WHERE (SALECODE_PRO = '" & ChangeQuote(Trim(SALECODE_PRO)) & "')"
         Else
            WhereStr = WhereStr & "AND (SALECODE_PRO = '" & ChangeQuote(Trim(SALECODE_PRO)) & "')"
         End If
      End If
   End If
   
   If Trim(SALENAME_PRO) <> "" Then
      If WildCard(SALENAME_PRO, SubLen, NewStr) Then
         If WhereStr = "" Then
            WhereStr = "WHERE (SALENAME_PRO LIKE '" & ChangeQuote(Trim(SALENAME_PRO)) & "')"
         Else
            WhereStr = WhereStr & "AND (SALENAME_PRO LIKE '" & ChangeQuote(Trim(SALENAME_PRO)) & "')"
         End If
      Else
         If WhereStr = "" Then
            WhereStr = "WHERE (SALENAME_PRO = '" & ChangeQuote(Trim(SALENAME_PRO)) & "')"
         Else
            WhereStr = WhereStr & "AND (SALENAME_PRO = '" & ChangeQuote(Trim(SALENAME_PRO)) & "')"
         End If
      End If
   End If
   
   If Len(FROM_SALE_CODE) > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (PYC.SALECODE_PRO >= '" & Trim(ChangeQuote(FROM_SALE_CODE)) & "') "
      Else
         WhereStr = WhereStr & "AND (PYC.SALECODE_PRO >= '" & Trim(ChangeQuote(FROM_SALE_CODE)) & "') "
      End If
   End If
   
   If Len(TO_SALE_CODE) > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (PYC.SALECODE_PRO <= '" & Trim(ChangeQuote(TO_SALE_CODE)) & "') "
      Else
         WhereStr = WhereStr & "AND (PYC.SALECODE_PRO <= '" & Trim(ChangeQuote(TO_SALE_CODE)) & "') "
      End If
   End If
   
   If Trim(CUSTOMERCODE_PRO) <> "" Then
      If WildCard(CUSTOMERCODE_PRO, SubLen, NewStr) Then
         If WhereStr = "" Then
            WhereStr = "WHERE (CUSTOMERCODE_PRO LIKE '" & ChangeQuote(Trim(CUSTOMERCODE_PRO)) & "')"
         Else
            WhereStr = WhereStr & "AND (CUSTOMERCODE_PRO LIKE '" & ChangeQuote(Trim(CUSTOMERCODE_PRO)) & "')"
         End If
      Else
         If WhereStr = "" Then
            WhereStr = "WHERE (CUSTOMERCODE_PRO = '" & ChangeQuote(Trim(CUSTOMERCODE_PRO)) & "')"
         Else
            WhereStr = WhereStr & "AND (CUSTOMERCODE_PRO = '" & ChangeQuote(Trim(CUSTOMERCODE_PRO)) & "')"
         End If
      End If
   End If
   
   If Trim(CUSTOMERNAME_PRO) <> "" Then
      If WildCard(CUSTOMERNAME_PRO, SubLen, NewStr) Then
         If WhereStr = "" Then
            WhereStr = "WHERE (CUSTOMERNAME_PRO LIKE '" & ChangeQuote(Trim(CUSTOMERNAME_PRO)) & "')"
         Else
            WhereStr = WhereStr & "AND (CUSTOMERNAME_PRO LIKE '" & ChangeQuote(Trim(CUSTOMERNAME_PRO)) & "')"
         End If
      Else
         If WhereStr = "" Then
            WhereStr = "WHERE (CUSTOMERNAME_PRO = '" & ChangeQuote(Trim(CUSTOMERNAME_PRO)) & "')"
         Else
            WhereStr = WhereStr & "AND (CUSTOMERNAME_PRO = '" & ChangeQuote(Trim(CUSTOMERNAME_PRO)) & "')"
         End If
      End If
   End If
   
   If Len(FROM_CUSTOMER_CODE) > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (PYC.CUSTOMERCODE_PRO >= '" & Trim(ChangeQuote(FROM_CUSTOMER_CODE)) & "') "
      Else
         WhereStr = WhereStr & "AND (PYC.CUSTOMERCODE_PRO >= '" & Trim(ChangeQuote(FROM_CUSTOMER_CODE)) & "') "
      End If
   End If
   
   If Len(TO_CUSTOMER_CODE) > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (PYC.CUSTOMERCODE_PRO <= '" & Trim(ChangeQuote(TO_CUSTOMER_CODE)) & "') "
      Else
         WhereStr = WhereStr & "AND (PYC.CUSTOMERCODE_PRO <= '" & Trim(ChangeQuote(TO_CUSTOMER_CODE)) & "') "
      End If
   End If
   
   If Trim(STKCOD_PRO) <> "" Then
      If WildCard(STKCOD_PRO, SubLen, NewStr) Then
         If WhereStr = "" Then
            WhereStr = "WHERE (STKCOD_PRO LIKE '" & ChangeQuote(Trim(STKCOD_PRO)) & "')"
         Else
            WhereStr = WhereStr & "AND (STKCOD_PRO LIKE '" & ChangeQuote(Trim(STKCOD_PRO)) & "')"
         End If
      Else
         If WhereStr = "" Then
            WhereStr = "WHERE (STKCOD_PRO = '" & ChangeQuote(Trim(STKCOD_PRO)) & "')"
         Else
            WhereStr = WhereStr & "AND (STKCOD_PRO = '" & ChangeQuote(Trim(STKCOD_PRO)) & "')"
         End If
      End If
   End If
   
   If Trim(STKNAME_PRO) <> "" Then
      If WildCard(STKNAME_PRO, SubLen, NewStr) Then
         If WhereStr = "" Then
            WhereStr = "WHERE (STKNAME_PRO LIKE '" & ChangeQuote(Trim(STKNAME_PRO)) & "')"
         Else
            WhereStr = WhereStr & "AND (STKNAME_PRO LIKE '" & ChangeQuote(Trim(STKNAME_PRO)) & "')"
         End If
      Else
         If WhereStr = "" Then
            WhereStr = "WHERE (STKNAME_PRO = '" & ChangeQuote(Trim(STKNAME_PRO)) & "')"
         Else
            WhereStr = WhereStr & "AND (STKNAME_PRO = '" & ChangeQuote(Trim(STKNAME_PRO)) & "')"
         End If
      End If
   End If
   
   If Len(FROM_STOCK_CODE) > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (PYC.STKCOD_PRO >= '" & Trim(ChangeQuote(FROM_STOCK_CODE)) & "') "
      Else
         WhereStr = WhereStr & "AND (PYC.STKCOD_PRO >= '" & Trim(ChangeQuote(FROM_STOCK_CODE)) & "') "
      End If
   End If
   
   If Len(TO_STOCK_CODE) > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (PYC.STKCOD_PRO <= '" & Trim(ChangeQuote(TO_STOCK_CODE)) & "') "
      Else
         WhereStr = WhereStr & "AND (PYC.STKCOD_PRO <= '" & Trim(ChangeQuote(TO_STOCK_CODE)) & "') "
      End If
   End If
   
   If AMOUNT_PRO > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (PYC.AMOUNT_PRO = " & AMOUNT_PRO & ") "
      Else
         WhereStr = WhereStr & "AND (PYC.AMOUNT_PRO = " & AMOUNT_PRO & ") "
      End If
   End If
   
   GroupByStr = ""
   If Ind = 1 Then
'      GroupByStr = "GROUP BY  PF.MONTH_NUM, PF.YEAR_NUM, PF.CUSTOMER_CODE, PF.STKCOD "
   ElseIf Ind = 3 Then
      GroupByStr = "GROUP BY PYC.SALECODE_PRO,PYC.CUSTOMERCODE_PRO,PYC.STKCOD_PRO,PYC.YYYY_MM "
   ElseIf Ind = 5 Then
      GroupByStr = "GROUP BY PYC.STKCOD_PRO,PYC.CUSTOMERCODE_PRO,PYC.YYYY_MM "
   End If

   OrderByStr = ""
   OrderTypeStr = ""
   
   If Ind = 1 Then
      If OrderBy = 1 Then
         OrderByStr = "ORDER BY PYC.PRO_DATE "
      ElseIf OrderBy = 2 Then
         OrderByStr = "ORDER BY PYC.SALENAME_PRO "
      ElseIf OrderBy = 3 Then
         OrderByStr = "ORDER BY PYC.CUSTOMERNAME_PRO "
      ElseIf OrderBy = 4 Then
         OrderByStr = "ORDER BY PYC.STKNAME_PRO "
      Else
         OrderByStr = "ORDER BY PYC.SALECODE_PRO "
      End If
   ElseIf Ind = 2 Then
      OrderByStr = "ORDER BY PYC.SALECODE_PRO ASC,PYC.CUSTOMERCODE_PRO ASC,PYC.STKCOD_PRO "
   ElseIf Ind = 3 Then
      OrderByStr = "ORDER BY PYC.SALECODE_PRO ASC,PYC.CUSTOMERCODE_PRO ASC,PYC.STKCOD_PRO,PYC.YYYY_MM "
   ElseIf Ind = 4 Then
      OrderByStr = "ORDER BY PYC.STKCOD_PRO ASC,PYC.CUSTOMERCODE_PRO "
   ElseIf Ind = 5 Then
      OrderByStr = "ORDER BY PYC.STKCOD_PRO ASC,PYC.CUSTOMERCODE_PRO ASC,PYC.YYYY_MM "
   End If
   
   
   If OrderType = 1 Then
      OrderTypeStr = "ASC "
   ElseIf OrderType = 2 Then
      OrderTypeStr = "DESC "
   Else
      OrderTypeStr = "ASC "
   End If
   
   SQL1 = SelectStr1 & FromStr & WhereStr & GroupByStr & OrderByStr & OrderTypeStr
   SQL2 = SelectStr2 & FromStr & WhereStr
   
   Call glbDatabaseMngr.GetLegacyRs(SQL1, SQL2, True, ItemCount, Rs, glbErrorLog)
End Function

Public Function PopulateFromRS(Ind As Long, Rs As ADODB.Recordset)
   If Ind = 1 Then
      PRO_ID = NVLI(Rs("PRO_ID"), -1)
      PRO_DATE = NVLS(Rs("PRO_DATE"), "")
      SALECODE_PRO = NVLS(Rs("SALECODE_PRO"), "")
      SALENAME_PRO = NVLS(Rs("SALENAME_PRO"), "")
      CUSTOMERCODE_PRO = NVLS(Rs("CUSTOMERCODE_PRO"), "")
      CUSTOMERNAME_PRO = NVLS(Rs("CUSTOMERNAME_PRO"), "")
      STKCOD_PRO = NVLS(Rs("STKCOD_PRO"), "")
      STKNAME_PRO = NVLS(Rs("STKNAME_PRO"), "")
      AMOUNT_PRO = NVLD(Rs("AMOUNT_PRO"), 0)
      YYYY_MM = NVLS(Rs("YYYY_MM"), "")
   ElseIf Ind = 2 Then
      SALECODE_PRO = NVLS(Rs("SALECODE_PRO"), "")
      SALENAME_PRO = NVLS(Rs("SALENAME_PRO"), "")
      CUSTOMERCODE_PRO = NVLS(Rs("CUSTOMERCODE_PRO"), "")
      CUSTOMERNAME_PRO = NVLS(Rs("CUSTOMERNAME_PRO"), "")
      STKCOD_PRO = NVLS(Rs("STKCOD_PRO"), "")
      STKNAME_PRO = NVLS(Rs("STKNAME_PRO"), "")
   ElseIf Ind = 3 Then
      SALECODE_PRO = NVLS(Rs("SALECODE_PRO"), "")
      CUSTOMERCODE_PRO = NVLS(Rs("CUSTOMERCODE_PRO"), "")
      STKCOD_PRO = NVLS(Rs("STKCOD_PRO"), "")
      AMOUNT_PRO = NVLD(Rs("AMOUNT_PRO"), 0)
      YYYY_MM = NVLS(Rs("YYYY_MM"), "")
   ElseIf Ind = 4 Then
      STKCOD_PRO = NVLS(Rs("STKCOD_PRO"), "")
      STKNAME_PRO = NVLS(Rs("STKNAME_PRO"), "")
      CUSTOMERCODE_PRO = NVLS(Rs("CUSTOMERCODE_PRO"), "")
      CUSTOMERNAME_PRO = NVLS(Rs("CUSTOMERNAME_PRO"), "")
   ElseIf Ind = 5 Then
      STKCOD_PRO = NVLS(Rs("STKCOD_PRO"), "")
      CUSTOMERCODE_PRO = NVLS(Rs("CUSTOMERCODE_PRO"), "")
      AMOUNT_PRO = NVLD(Rs("AMOUNT_PRO"), 0)
      YYYY_MM = NVLS(Rs("YYYY_MM"), "")
   End If
End Function
Public Function AddEditData()
Dim ServerDateTime As String
Dim SQL1 As String
Dim ID As Long

   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   
   If AddEditMode = SHOW_ADD Then
      SQL1 = "INSERT INTO PROMOTION_PAYCUSTOMER " & _
                     "( " & _
                     "   PRO_DATE, " & _
                     "   SALECODE_PRO, " & _
                     "   SALENAME_PRO, " & _
                     "   CUSTOMERCODE_PRO, " & _
                     "   CUSTOMERNAME_PRO, " & _
                     "   STKCOD_PRO, " & _
                     "   STKNAME_PRO, " & _
                     "   AMOUNT_PRO, " & _
                     "   YYYY_MM " & _
                     ") " & _
                     "VALUES " & _
                     "( " & _
                        "'" & ChangeQuote(Trim(PRO_DATE)) & "'" & "," & _
                        "'" & ChangeQuote(Trim(SALECODE_PRO)) & "'" & "," & _
                        "'" & ChangeQuote(Trim(SALENAME_PRO)) & "'" & "," & _
                        "'" & ChangeQuote(Trim(CUSTOMERCODE_PRO)) & "'" & "," & _
                        "'" & ChangeQuote(Trim(CUSTOMERNAME_PRO)) & "'" & "," & _
                        "'" & ChangeQuote(Trim(STKCOD_PRO)) & "'" & "," & _
                        "'" & ChangeQuote(Trim(STKNAME_PRO)) & "'" & "," & _
                        "" & AMOUNT_PRO & "," & _
                        "'" & ChangeQuote(Trim(YYYY_MM)) & "'" & _
                     ") "
   Else
            SQL1 = "UPDATE PROMOTION_PAYCUSTOMER SET " & _
                        "   PRO_DATE = '" & ChangeQuote(Trim(PRO_DATE)) & "' ," & _
                        "   SALECODE_PRO = '" & ChangeQuote(Trim(SALECODE_PRO)) & "' ," & _
                        "   SALENAME_PRO = '" & ChangeQuote(Trim(SALENAME_PRO)) & "' ," & _
                        "   CUSTOMERCODE_PRO = '" & ChangeQuote(Trim(CUSTOMERCODE_PRO)) & "' ," & _
                        "   CUSTOMERNAME_PRO = '" & ChangeQuote(Trim(CUSTOMERNAME_PRO)) & "' ," & _
                        "   STKCOD_PRO = '" & ChangeQuote(Trim(STKCOD_PRO)) & "' ," & _
                        "   STKNAME_PRO = '" & ChangeQuote(Trim(STKNAME_PRO)) & "' ," & _
                        "   AMOUNT_PRO = " & AMOUNT_PRO & " ," & _
                        "   YYYY_MM = '" & ChangeQuote(Trim(YYYY_MM)) & "' " & _
                        " WHERE PRO_ID = " & PRO_ID
   End If
   
   If (AddEditMode = SHOW_ADD) Or (AddEditMode = SHOW_EDIT) Then
      m_Conn.Execute (SQL1)
   End If
End Function

Public Function DeleteData()
Dim SQL1 As String
      SQL1 = "DELETE FROM PROMOTION_PAYCUSTOMER WHERE PRO_ID  = " & PRO_ID
      m_Conn.Execute (SQL1)
End Function
