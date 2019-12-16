VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CPromotionYear"
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

Public PROYEAR_ID As Long
Public DATEYEAR_PRO As String
Public CTMCODYEAR_PRO As String
Public CTMNAMEYEAR_PRO As String
Public STKCODYEAR_PRO As String
Public STKNAMEYEAR_PRO As String
Public AMOUNTYEAR_PRO As Double
Public SUM_AMOUT As Double
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
      SelectStr1 = "SELECT PTY.* "
      SelectStr2 = "SELECT COUNT(*)  "
      
      FromStr = "FROM PROMOTION_YEAR PTY "
      
   ElseIf Ind = 2 Then
      SelectStr1 = "SELECT PTY.CTMCODYEAR_PRO ,PTY.STKCODYEAR_PRO ,PTY.YYYY_MM ,SUM(PTY.AMOUNTYEAR_PRO) AS SUM_AMOUT  "
      SelectStr2 = "SELECT COUNT(*)  "
      
      FromStr = "FROM PROMOTION_YEAR PTY "
   End If
   
   If PROYEAR_ID > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (PROYEAR_ID= " & PROYEAR_ID & ") "
      Else
         WhereStr = WhereStr & "AND (PROYEAR_ID = " & PROYEAR_ID & ") "
      End If
   End If
   
    If FROM_PRO_DATE > -1 Then
      TempDate = DateToString(FROM_PRO_DATE)
      If WhereStr = "" Then
         WhereStr = "WHERE ((PTY.DATEYEAR_PRO) >= '" & ChangeQuote(TempDate) & "')"
      Else
         WhereStr = WhereStr & "AND ((PTY.DATEYEAR_PRO) >= '" & ChangeQuote(TempDate) & "')"
      End If
   End If

    If TO_PRO_DATE > -1 Then
      TempDate = DateToString(TO_PRO_DATE)
      If WhereStr = "" Then
         WhereStr = "WHERE ((PTY.DATEYEAR_PRO) <= '" & ChangeQuote(TempDate) & "')"
      Else
         WhereStr = WhereStr & "AND ((PTY.DATEYEAR_PRO) <= '" & ChangeQuote(TempDate) & "')"
      End If
   End If
   
   If Trim(CTMCODYEAR_PRO) <> "" Then
      If WildCard(CTMCODYEAR_PRO, SubLen, NewStr) Then
         If WhereStr = "" Then
            WhereStr = "WHERE (CTMCODYEAR_PRO LIKE '" & ChangeQuote(Trim(CTMCODYEAR_PRO)) & "')"
         Else
            WhereStr = WhereStr & "AND (CTMCODYEAR_PRO LIKE '" & ChangeQuote(Trim(CTMCODYEAR_PRO)) & "')"
         End If
      Else
         If WhereStr = "" Then
            WhereStr = "WHERE (CTMCODYEAR_PRO = '" & ChangeQuote(Trim(CTMCODYEAR_PRO)) & "')"
         Else
            WhereStr = WhereStr & "AND (CTMCODYEAR_PRO = '" & ChangeQuote(Trim(CTMCODYEAR_PRO)) & "')"
         End If
      End If
   End If
   
   If Trim(CTMNAMEYEAR_PRO) <> "" Then
      If WildCard(CTMNAMEYEAR_PRO, SubLen, NewStr) Then
         If WhereStr = "" Then
            WhereStr = "WHERE (CTMNAMEYEAR_PRO LIKE '" & ChangeQuote(Trim(CTMNAMEYEAR_PRO)) & "')"
         Else
            WhereStr = WhereStr & "AND (CTMNAMEYEAR_PRO LIKE '" & ChangeQuote(Trim(CTMNAMEYEAR_PRO)) & "')"
         End If
      Else
         If WhereStr = "" Then
            WhereStr = "WHERE (CTMNAMEYEAR_PRO = '" & ChangeQuote(Trim(CTMNAMEYEAR_PRO)) & "')"
         Else
            WhereStr = WhereStr & "AND (CTMNAMEYEAR_PRO = '" & ChangeQuote(Trim(CTMNAMEYEAR_PRO)) & "')"
         End If
      End If
   End If
   
   If Len(FROM_CUSTOMER_CODE) > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (PTY.CTMCODYEAR_PRO >= '" & Trim(ChangeQuote(FROM_CUSTOMER_CODE)) & "') "
      Else
         WhereStr = WhereStr & "AND (PTY.CTMCODYEAR_PRO >= '" & Trim(ChangeQuote(FROM_CUSTOMER_CODE)) & "') "
      End If
   End If
   
   If Len(TO_CUSTOMER_CODE) > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (PTY.CTMCODYEAR_PRO <= '" & Trim(ChangeQuote(TO_CUSTOMER_CODE)) & "') "
      Else
         WhereStr = WhereStr & "AND (PTY.CTMCODYEAR_PRO <= '" & Trim(ChangeQuote(TO_CUSTOMER_CODE)) & "') "
      End If
   End If
   
   If Trim(STKCODYEAR_PRO) <> "" Then
      If WildCard(STKCODYEAR_PRO, SubLen, NewStr) Then
         If WhereStr = "" Then
            WhereStr = "WHERE (STKCODYEAR_PRO LIKE '" & ChangeQuote(Trim(STKCODYEAR_PRO)) & "')"
         Else
            WhereStr = WhereStr & "AND (STKCODYEAR_PRO LIKE '" & ChangeQuote(Trim(STKCODYEAR_PRO)) & "')"
         End If
      Else
         If WhereStr = "" Then
            WhereStr = "WHERE (STKCODYEAR_PRO = '" & ChangeQuote(Trim(STKCODYEAR_PRO)) & "')"
         Else
            WhereStr = WhereStr & "AND (STKCODYEAR_PRO = '" & ChangeQuote(Trim(STKCODYEAR_PRO)) & "')"
         End If
      End If
   End If
   
   If Trim(STKNAMEYEAR_PRO) <> "" Then
      If WildCard(STKNAMEYEAR_PRO, SubLen, NewStr) Then
         If WhereStr = "" Then
            WhereStr = "WHERE (STKNAMEYEAR_PRO LIKE '" & ChangeQuote(Trim(STKNAMEYEAR_PRO)) & "')"
         Else
            WhereStr = WhereStr & "AND (STKNAMEYEAR_PRO LIKE '" & ChangeQuote(Trim(STKNAMEYEAR_PRO)) & "')"
         End If
      Else
         If WhereStr = "" Then
            WhereStr = "WHERE (STKNAMEYEAR_PRO = '" & ChangeQuote(Trim(STKNAMEYEAR_PRO)) & "')"
         Else
            WhereStr = WhereStr & "AND (STKNAMEYEAR_PRO = '" & ChangeQuote(Trim(STKNAMEYEAR_PRO)) & "')"
         End If
      End If
   End If
   
   If Len(FROM_STOCK_CODE) > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (PTY.STKCODYEAR_PRO >= '" & Trim(ChangeQuote(FROM_STOCK_CODE)) & "') "
      Else
         WhereStr = WhereStr & "AND (PTY.STKCODYEAR_PRO >= '" & Trim(ChangeQuote(FROM_STOCK_CODE)) & "') "
      End If
   End If
   
   If Len(TO_STOCK_CODE) > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (PTY.STKCODYEAR_PRO <= '" & Trim(ChangeQuote(TO_STOCK_CODE)) & "') "
      Else
         WhereStr = WhereStr & "AND (PTY.STKCODYEAR_PRO <= '" & Trim(ChangeQuote(TO_STOCK_CODE)) & "') "
      End If
   End If
   
   If AMOUNTYEAR_PRO > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (PTY.AMOUNTYEAR_PRO = " & AMOUNTYEAR_PRO & ") "
      Else
         WhereStr = WhereStr & "AND (PTY.AMOUNTYEAR_PRO = " & AMOUNTYEAR_PRO & ") "
      End If
   End If
   
   GroupByStr = ""
   If Ind = 2 Then
      GroupByStr = "GROUP BY  PTY.CTMCODYEAR_PRO ,PTY.STKCODYEAR_PRO ,PTY.YYYY_MM  "
   End If

   OrderByStr = ""
   OrderTypeStr = ""
   
   If Ind = 1 Then
      If OrderBy = 1 Then
         OrderByStr = "ORDER BY PTY.DATEYEAR_PRO  "
      Else
         OrderByStr = "ORDER BY PTY.CTMCODYEAR_PRO  "
      End If
   ElseIf Ind = 2 Then
      If OrderBy = 1 Then
         OrderByStr = "ORDER BY PTY.DATEYEAR_PRO  "
      Else
         OrderByStr = "ORDER BY PTY.CTMCODYEAR_PRO  "
      End If
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
      PROYEAR_ID = NVLI(Rs("PROYEAR_ID"), -1)
      DATEYEAR_PRO = NVLS(Rs("DATEYEAR_PRO"), "")
      CTMCODYEAR_PRO = NVLS(Rs("CTMCODYEAR_PRO"), "")
      CTMNAMEYEAR_PRO = NVLS(Rs("CTMNAMEYEAR_PRO"), "")
      STKCODYEAR_PRO = NVLS(Rs("STKCODYEAR_PRO"), "")
      STKNAMEYEAR_PRO = NVLS(Rs("STKNAMEYEAR_PRO"), "")
      AMOUNTYEAR_PRO = NVLD(Rs("AMOUNTYEAR_PRO"), 0)
      YYYY_MM = NVLS(Rs("YYYY_MM"), "")
   ElseIf Ind = 2 Then
      CTMCODYEAR_PRO = NVLS(Rs("CTMCODYEAR_PRO"), "")
      STKCODYEAR_PRO = NVLS(Rs("STKCODYEAR_PRO"), "")
      YYYY_MM = NVLS(Rs("YYYY_MM"), "")
      SUM_AMOUT = NVLD(Rs("SUM_AMOUT"), 0)
   End If
End Function
Public Function AddEditData()
Dim ServerDateTime As String
Dim SQL1 As String
Dim ID As Long

   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   
   If AddEditMode = SHOW_ADD Then
      SQL1 = "INSERT INTO PROMOTION_YEAR " & _
                     "( " & _
                     "   DATEYEAR_PRO, " & _
                     "   CTMCODYEAR_PRO, " & _
                     "   CTMNAMEYEAR_PRO, " & _
                     "   STKCODYEAR_PRO, " & _
                     "   STKNAMEYEAR_PRO, " & _
                     "   AMOUNTYEAR_PRO, " & _
                     "   YYYY_MM " & _
                     ") " & _
                     "VALUES " & _
                     "( " & _
                        "'" & ChangeQuote(Trim(DATEYEAR_PRO)) & "'" & "," & _
                        "'" & ChangeQuote(Trim(CTMCODYEAR_PRO)) & "'" & "," & _
                        "'" & ChangeQuote(Trim(CTMNAMEYEAR_PRO)) & "'" & "," & _
                        "'" & ChangeQuote(Trim(STKCODYEAR_PRO)) & "'" & "," & _
                        "'" & ChangeQuote(Trim(STKNAMEYEAR_PRO)) & "'" & "," & _
                        "" & AMOUNTYEAR_PRO & "," & _
                        "'" & ChangeQuote(Trim(YYYY_MM)) & "'" & _
                     ") "
   Else
            SQL1 = "UPDATE PROMOTION_YEAR SET " & _
                        "   DATEYEAR_PRO = '" & ChangeQuote(Trim(DATEYEAR_PRO)) & "' ," & _
                        "   CTMCODYEAR_PRO = '" & ChangeQuote(Trim(CTMCODYEAR_PRO)) & "' ," & _
                        "   CTMNAMEYEAR_PRO = '" & ChangeQuote(Trim(CTMNAMEYEAR_PRO)) & "' ," & _
                        "   STKCODYEAR_PRO = '" & ChangeQuote(Trim(STKCODYEAR_PRO)) & "' ," & _
                        "   STKNAMEYEAR_PRO = '" & ChangeQuote(Trim(STKNAMEYEAR_PRO)) & "' ," & _
                        "   AMOUNTYEAR_PRO = " & AMOUNTYEAR_PRO & " ," & _
                        "   YYYY_MM = '" & ChangeQuote(Trim(YYYY_MM)) & "' " & _
                        " WHERE PROYEAR_ID = " & PROYEAR_ID
   End If
   
   If (AddEditMode = SHOW_ADD) Or (AddEditMode = SHOW_EDIT) Then
      m_Conn.Execute (SQL1)
   End If
End Function

Public Function DeleteData()
Dim SQL1 As String
      SQL1 = "DELETE FROM PROMOTION_YEAR WHERE PROYEAR_ID  = " & PROYEAR_ID
      m_Conn.Execute (SQL1)
End Function
