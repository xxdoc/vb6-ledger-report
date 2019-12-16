VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CComIVcredit"
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

Public COM_IV_CREDIT_ID As Long
Public IV_DOCDAT As Date
Public IV_COD As String
Public SLMCOD As String
Public SLMNAME As String
Public CUS_ID As String
Public CUS_NAME As String
Public CR_TYPE As String
Public CR_DATA As Long
Public FROM_DOC_DATE As Date
Public TO_DOC_DATE As Date

Public ORDER_BY As Long
Public ORDER_TYPE  As Long

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
   SelectStr1 = "SELECT IVCR.*  "
   SelectStr2 = "SELECT COUNT(*)  "
   
   FromStr = "FROM COM_IV_CREDIT IVCR "
    
   WhereStr = ""
'ElseIf ind = 2 Then

End If
   
'   If OwnerEnterprise Then
'      ENTERPRISE_CODE = glbCompanyCode
'   End If
   
   If IV_COD <> "" Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (IV_COD = '" & ChangeQuote(Trim(IV_COD)) & "') "
      Else
         WhereStr = WhereStr & "AND (IV_COD = '" & ChangeQuote(Trim(IV_COD)) & "') "
      End If
   End If
   
      If SLMCOD <> "" Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (SLMCOD = '" & ChangeQuote(Trim(SLMCOD)) & "') "
      Else
         WhereStr = WhereStr & "AND (SLMCOD = '" & ChangeQuote(Trim(SLMCOD)) & "') "
      End If
   End If
   
     If CUS_ID <> "" Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (CUS_ID = '" & ChangeQuote(Trim(CUS_ID)) & "') "
      Else
         WhereStr = WhereStr & "AND (CUS_ID = '" & ChangeQuote(Trim(CUS_ID)) & "') "
      End If
   End If

   If COM_IV_CREDIT_ID > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (COM_IV_CREDIT_ID = " & COM_IV_CREDIT_ID & ") "
      Else
         WhereStr = WhereStr & "AND (COM_IV_CREDIT_ID = '" & COM_IV_CREDIT_ID & ") "
      End If
   End If


   If FROM_DOC_DATE > 0 Then                     ' ¨ÐàÍÒäÇé âËÅ´ÃÕ¾ÍÃìµ
      TempDate = DateToStringIntLow(Trim(FROM_DOC_DATE))
      If WhereStr = "" Then
         WhereStr = "WHERE (IV_DOCDAT >= '" & ChangeQuote(TempDate) & "') "
      Else
         WhereStr = WhereStr & "AND (IV_DOCDAT >= '" & ChangeQuote(TempDate) & "') "
      End If
   End If

   If TO_DOC_DATE > 0 Then
      TempDate = DateToStringIntHi(Trim(TO_DOC_DATE))
      If WhereStr = "" Then
         WhereStr = "WHERE (IV_DOCDAT <= '" & ChangeQuote(TempDate) & "') "
      Else
         WhereStr = WhereStr & "AND (IV_DOCDAT <= '" & ChangeQuote(TempDate) & "') "
      End If
   End If
   
   OrderByStr = ""
   OrderTypeStr = ""

   If OrderBy = 1 Then
      OrderByStr = "ORDER BY IV_DOCDAT "
   ElseIf OrderBy = 2 Then
      OrderByStr = "ORDER BY IV_COD "
   ElseIf OrderBy = 3 Then
      OrderByStr = "ORDER BY SLMCOD "
   ElseIf OrderBy = 4 Then
      OrderByStr = "ORDER BY CUS_ID "
   Else
      OrderByStr = "ORDER BY IV_DOCDAT "
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
  ' 'debug.print SQL1
   Call glbDatabaseMngr.GetLegacyRs(SQL1, SQL2, True, ItemCount, Rs, glbErrorLog)
End Function

Public Function PopulateFromRS(Ind As Long, Rs As ADODB.Recordset)
If Ind = 1 Then
   COM_IV_CREDIT_ID = NVLI(Rs("COM_IV_CREDIT_ID"), -1)
   IV_DOCDAT = InternalDateToDate(NVLS(Rs("IV_DOCDAT"), ""))
   IV_COD = NVLS(Rs("IV_COD"), "")
   SLMCOD = NVLS(Rs("SLMCOD"), "")
   SLMNAME = NVLS(Rs("SLMNAME"), "")
    CUS_ID = NVLS(Rs("CUS_ID"), "")
   CUS_NAME = NVLS(Rs("CUS_NAME"), "")
   CR_TYPE = NVLS(Rs("CR_TYPE"), "")
   CR_DATA = NVLI(Rs("CR_DATA"), -1)
End If
End Function
Public Function AddEditData()
Dim ServerDateTime As String
Dim SQL1 As String

   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   
'   ENTERPRISE_CODE = glbCompanyCode
   
   If AddEditMode = SHOW_ADD Then
      SQL1 = "INSERT INTO COM_IV_CREDIT " & _
                     "( " & _
                     "   IV_DOCDAT, " & _
                     "   IV_COD, " & _
                     "   SLMCOD, " & _
                     "   SLMNAME, " & _
                     "   CUS_ID, " & _
                     "   CUS_NAME, " & _
                     "   CR_TYPE, " & _
                     "   CR_DATA " & _
                     ") " & _
                     "VALUES " & _
                     "( " & _
                        "'" & ChangeQuote(Trim(DateToStringInt(IV_DOCDAT))) & "'," & _
                        "'" & ChangeQuote(Trim(IV_COD)) & "'," & _
                        "'" & ChangeQuote(Trim(SLMCOD)) & "'," & _
                        "'" & ChangeQuote(Trim(SLMNAME)) & "'," & _
                        "'" & ChangeQuote(Trim(CUS_ID)) & "'," & _
                        "'" & ChangeQuote(Trim(CUS_NAME)) & "'," & _
                         "'" & ChangeQuote(Trim(CR_TYPE)) & "'," & _
                         "" & ChangeQuote(Trim(CR_DATA)) & "" & _
                     ") "

   Else 'Show edit
      SQL1 = "UPDATE COM_IV_CREDIT SET " & _
                     "   IV_DOCDAT = '" & ChangeQuote(Trim(DateToStringInt(IV_DOCDAT))) & "', " & _
                     "   IV_COD = '" & ChangeQuote(Trim(IV_COD)) & "', " & _
                     "   SLMCOD = '" & ChangeQuote(Trim(SLMCOD)) & "', " & _
                     "   SLMNAME = '" & ChangeQuote(Trim(SLMNAME)) & "', " & _
                     "   CUS_ID = '" & ChangeQuote(Trim(CUS_ID)) & "', " & _
                     "   CUS_NAME = '" & ChangeQuote(Trim(CUS_NAME)) & "', " & _
                     "   CR_TYPE = '" & ChangeQuote(Trim(CR_TYPE)) & "', " & _
                     "   CR_DATA = " & ChangeQuote(Trim(CR_DATA)) & " " & _
                     "WHERE COM_IV_CREDIT_ID = " & COM_IV_CREDIT_ID
   End If
   
   If (AddEditMode = SHOW_ADD) Or (AddEditMode = SHOW_EDIT) Then
      m_Conn.Execute (SQL1)
   End If
End Function

Public Function DeleteData()
Dim SQL1 As String
      SQL1 = "DELETE FROM COM_IV_CREDIT WHERE COM_IV_CREDIT_ID = " & COM_IV_CREDIT_ID
      m_Conn.Execute (SQL1)
End Function
