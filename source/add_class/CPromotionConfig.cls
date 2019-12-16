VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CPromotionConfig"
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

Public PROMOTION_ID As Long
Public MONTH_NUM As String
Public YEAR_NUM As String
Public CUSTOMER_NAME As String
Public CUSTOMER_CODE As String
Public STKCOD As String
Public AMOUNT As Double
Public SUM_AMOUNT As Double

Public FROM_PRO_DATE As Date
Public TO_PRO_DATE As Date


Public KEY_ID As Long
Public KEY_LOOKUP As String

Private Sub Class_Initialize()
   Set m_Conn = glbDatabaseMngr.DBLegacyConnection
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

   SelectStr1 = "SELECT PF.MONTH_NUM, PF.YEAR_NUM, PF.CUSTOMER_CODE, PF.STKCOD, SUM(AMOUNT) AS SUM_AMOUNT  "
   SelectStr2 = "SELECT COUNT(*)  "
   
   FromStr = "FROM PROMOTION_FREE PF "
   
   
   GroupByStr = ""
   If Ind = 1 Then
      GroupByStr = "GROUP BY  PF.MONTH_NUM, PF.YEAR_NUM, PF.CUSTOMER_CODE, PF.STKCOD  "
   Else
      GroupByStr = "GROUP BY  PF.MONTH_NUM, PF.YEAR_NUM, PF.CUSTOMER_CODE, PF.STKCOD  "
   End If

   OrderByStr = ""
   OrderTypeStr = ""

   If OrderBy = 1 Then
      OrderByStr = "ORDER BY PF.CUSTOMER_CODE "
   Else
      OrderByStr = "ORDER BY PF.CUSTOMER_CODE "
   End If
   
   If OrderType = 1 Then
      OrderTypeStr = "ASC "
   Else
      OrderTypeStr = "DESC "
   End If
   
   SQL1 = SelectStr1 & FromStr & WhereStr & GroupByStr & OrderByStr & OrderTypeStr
   SQL2 = SelectStr2 & FromStr & WhereStr
   
   Call glbDatabaseMngr.GetLegacyRs(SQL1, SQL2, True, ItemCount, Rs, glbErrorLog)
End Function

Public Function PopulateFromRS(Ind As Long, Rs As ADODB.Recordset)
   MONTH_NUM = NVLS(Rs("MONTH_NUM"), "")
   YEAR_NUM = NVLS(Rs("YEAR_NUM"), "")
   CUSTOMER_CODE = NVLS(Rs("CUSTOMER_CODE"), "")
   STKCOD = NVLS(Rs("STKCOD"), "")
   SUM_AMOUNT = NVLD(Rs("SUM_AMOUNT"), 0)
End Function
Public Function AddEditData()
Dim ServerDateTime As String
Dim SQL1 As String
Dim ID As Long

   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   
   If AddEditMode = SHOW_ADD Then
      SQL1 = "INSERT INTO PROMOTION_FREE " & _
                     "( " & _
                     "   MONTH_NUM, " & _
                     "   YEAR_NUM, " & _
                     "   CUSTOMER_NAME, " & _
                     "   CUSTOMER_CODE, " & _
                     "   STKCOD, " & _
                     "   AMOUNT " & _
                     ") " & _
                     "VALUES " & _
                     "( " & _
                        "'" & ChangeQuote(Trim(MONTH_NUM)) & "'" & "," & _
                        "'" & ChangeQuote(Trim(YEAR_NUM)) & "'" & "," & _
                        "'" & ChangeQuote(Trim(CUSTOMER_NAME)) & "'" & "," & _
                        "'" & ChangeQuote(Trim(CUSTOMER_CODE)) & "'" & "," & _
                        "'" & ChangeQuote(Trim(STKCOD)) & "'" & "," & _
                        "" & AMOUNT & "" & _
                     ") "
   End If
   
   If (AddEditMode = SHOW_ADD) Then
      m_Conn.Execute (SQL1)
   End If
End Function

Public Function DeleteData()
Dim SQL1 As String
      SQL1 = "DELETE FROM PROMOTION_FREE "
      m_Conn.Execute (SQL1)
End Function
