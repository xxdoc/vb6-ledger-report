VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CComMinusStk"
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

Public COM_MINUSSTK_ID As Long
Public IV_DOCDAT As Date
Public IV_COD As String
Public STK_COD As String
Public STK_NAME As String
Public MINUS_AMOUNT As String
Public FROM_DOC_DATE As Date
Public TO_DOC_DATE As Date
Public MINUS_COD As String

Public SLMCOD As String
Public SLMNAME As String

Private Sub Class_Initialize()
   Set m_Conn = glbDatabaseMngr.DBLegacyConnection
    FROM_DOC_DATE = -1
    TO_DOC_DATE = -1
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
Dim TempDate As String

If Ind = 1 Then
   SelectStr1 = "SELECT CMS.*  "
   SelectStr2 = "SELECT COUNT(*)  "
   
   FromStr = "FROM COM_MINUS_STK CMS "
   
   WhereStr = ""
'ElseIf ind = 2 Then

End If
   
'   If OwnerEnterprise Then
'      ENTERPRISE_CODE = glbCompanyCode
'   End If
   
   If IV_COD <> "" Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (CMS.IV_COD = '" & ChangeQuote(Trim(IV_COD)) & "') "
      Else
         WhereStr = WhereStr & "AND (CMS.IV_COD = '" & ChangeQuote(Trim(IV_COD)) & "') "
      End If
   End If
   
    If SLMCOD <> "" Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (CMS.SLMCOD = '" & ChangeQuote(Trim(SLMCOD)) & "') "
      Else
         WhereStr = WhereStr & "AND (CMS.SLMCOD = '" & ChangeQuote(Trim(SLMCOD)) & "') "
      End If
   End If
   
      If STK_COD <> "" Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (CMS.STK_COD = '" & ChangeQuote(Trim(STK_COD)) & "') "
      Else
         WhereStr = WhereStr & "AND (CMS.STK_COD = '" & ChangeQuote(Trim(STK_COD)) & "') "
      End If
   End If

   If COM_MINUSSTK_ID > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (CMS.COM_MINUSSTK_ID = " & COM_MINUSSTK_ID & ") "
      Else
         WhereStr = WhereStr & "AND (CMS.COM_MINUSSTK_ID = '" & COM_MINUSSTK_ID & ") "
      End If
   End If
   
      If MINUS_COD > "" Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (CMS.MINUS_COD = " & MINUS_COD & ") "
      Else
         WhereStr = WhereStr & "AND (CMS.MINUS_COD = '" & MINUS_COD & ") "
      End If
   End If


   If FROM_DOC_DATE > 0 Then                     ' ¨ÐàÍÒäÇé âËÅ´ÃÕ¾ÍÃìµ
      TempDate = DateToStringIntLow(Trim(FROM_DOC_DATE))
      If WhereStr = "" Then
         WhereStr = "WHERE (CMS.IV_DOCDAT >= '" & ChangeQuote(TempDate) & "') "
      Else
         WhereStr = WhereStr & "AND (CMS.IV_DOCDAT >= '" & ChangeQuote(TempDate) & "') "
      End If
   End If

   If TO_DOC_DATE > 0 Then
      TempDate = DateToStringIntHi(Trim(TO_DOC_DATE))
      If WhereStr = "" Then
         WhereStr = "WHERE (CMS.IV_DOCDAT <= '" & ChangeQuote(TempDate) & "') "
      Else
         WhereStr = WhereStr & "AND (CMS.IV_DOCDAT <= '" & ChangeQuote(TempDate) & "') "
      End If
   End If
   
   OrderByStr = ""
   OrderTypeStr = ""

   If OrderBy = 1 Then
      OrderByStr = "ORDER BY CMS.IV_DOCDAT "
   ElseIf OrderBy = 2 Then
      OrderByStr = "ORDER BY CMS.IV_COD "
   ElseIf OrderBy = 3 Then
      OrderByStr = "ORDER BY CMS.STK_COD "
   Else
      OrderByStr = "ORDER BY CMS.IV_DOCDAT "
   End If
   
   If OrderType = 1 Then
      OrderTypeStr = "ASC "
   ElseIf OrderType = 2 Then
      OrderTypeStr = "DESC "
   Else
      OrderTypeStr = "ASC "
   End If
   
   SQL1 = SelectStr1 & FromStr & WhereStr & OrderByStr & OrderTypeStr
   SQL2 = SelectStr2 & FromStr & WhereStr
   'debug.print SQL1
   Call glbDatabaseMngr.GetLegacyRs(SQL1, SQL2, True, ItemCount, Rs, glbErrorLog)
End Function

Public Function PopulateFromRS(Ind As Long, Rs As ADODB.Recordset)
If Ind = 1 Then
   COM_MINUSSTK_ID = NVLI(Rs("COM_MINUSSTK_ID"), -1)
   MINUS_COD = NVLS(Rs("MINUS_COD"), "")
   IV_DOCDAT = InternalDateToDate(NVLS(Rs("IV_DOCDAT"), ""))
   IV_COD = NVLS(Rs("IV_COD"), "")
   SLMCOD = NVLS(Rs("SLMCOD"), "")
   SLMNAME = NVLS(Rs("SLMNAME"), "")
   STK_COD = NVLS(Rs("STK_COD"), "")
   STK_NAME = NVLS(Rs("STK_NAME"), "")
   MINUS_AMOUNT = NVLS(Rs("MINUS_AMOUNT"), "")
End If
End Function
Public Function AddEditData()
Dim ServerDateTime As String
Dim SQL1 As String

   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   
'   ENTERPRISE_CODE = glbCompanyCode
   
   If AddEditMode = SHOW_ADD Then
      SQL1 = "INSERT INTO COM_MINUS_STK " & _
                     "( " & _
                     "   IV_DOCDAT, " & _
                     "   IV_COD, " & _
                     "   SLMCOD, " & _
                     "   SLMNAME, " & _
                     "   STK_COD, " & _
                     "   STK_NAME, " & _
                     "   MINUS_AMOUNT, " & _
                     "   MINUS_COD" & _
                     ") " & _
                     "VALUES " & _
                     "( " & _
                        "'" & ChangeQuote(Trim(DateToStringInt(IV_DOCDAT))) & "'," & _
                        "'" & ChangeQuote(Trim(IV_COD)) & "'," & _
                        "'" & ChangeQuote(Trim(SLMCOD)) & "'," & _
                        "'" & ChangeQuote(Trim(SLMNAME)) & "'," & _
                        "'" & ChangeQuote(Trim(STK_COD)) & "'," & _
                        "'" & ChangeQuote(Trim(STK_NAME)) & "'," & _
                         "" & ChangeQuote(Trim(MINUS_AMOUNT)) & "," & _
                         "'" & ChangeQuote(Trim(MINUS_COD)) & "'" & _
                     ") "

   Else 'Show edit
      SQL1 = "UPDATE COM_MINUS_STK SET " & _
                     "   IV_DOCDAT = '" & ChangeQuote(Trim(DateToStringInt(IV_DOCDAT))) & "', " & _
                     "   IV_COD = '" & ChangeQuote(Trim(IV_COD)) & "', " & _
                     "   SLMCOD = '" & ChangeQuote(Trim(SLMCOD)) & "', " & _
                     "   SLMNAME = '" & ChangeQuote(Trim(SLMNAME)) & "', " & _
                     "   STK_COD = '" & ChangeQuote(Trim(STK_COD)) & "', " & _
                     "   STK_NAME = '" & ChangeQuote(Trim(STK_NAME)) & "', " & _
                     "   MINUS_AMOUNT = " & ChangeQuote(Trim(MINUS_AMOUNT)) & ", " & _
                     "   MINUS_COD = '" & ChangeQuote(Trim(MINUS_COD)) & "' " & _
                     "WHERE COM_MINUSSTK_ID = " & COM_MINUSSTK_ID
   End If
   
   If (AddEditMode = SHOW_ADD) Or (AddEditMode = SHOW_EDIT) Then
      m_Conn.Execute (SQL1)
   End If
End Function

Public Function DeleteData()
Dim SQL1 As String
      SQL1 = "DELETE FROM COM_MINUS_STK WHERE COM_MINUSSTK_ID = " & COM_MINUSSTK_ID
      m_Conn.Execute (SQL1)
End Function
