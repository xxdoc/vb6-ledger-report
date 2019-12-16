VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CARMas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_Conn As ADODB.Connection

Public AddEditMode As SHOW_MODE_TYPE
Public QueryFlag As Long

Public Flag As String
' End of generic data member

Public CUSCOD As String
Public CUSTYP As String
Public PRENAM  As String
Public CUSNAM  As String
Public ADDR01  As String
Public ADDR02 As String
Public TELNUM As String
Public CONTRACT As String
Public PAYTRM  As Double
Public PAYCOND As String

Public TYPDES As String
Public SLMCOD As String
Public AREACOD As String
Public CRLINE As String
Public SLMNAM As String

Public SUM_DB1 As Double
Public SUM_DB2 As Double
Public SUM_DB3 As Double

Public FROM_CUSTOMER_CODE As String
Public TO_CUSTOMER_CODE As String

Public KEY_ID As Long
Public KEY_LOOKUP As String

Public FROM_SALE_CODE As String
Public TO_SALE_CODE As String

Public OrderBy As Long
Public OrderType As Long
Private Sub Class_Initialize()
   Set m_Conn = glbDatabaseMngr.DBConnection
End Sub

Public Function PackAddress()
Dim TempStr As String

   If Len(ADDR02) > 0 Then
      PackAddress = ADDR01 & vbCrLf & PackAddress
   Else
      PackAddress = ADDR01
   End If
End Function

'In this class we will not handle any error or any transaction
Public Function QueryData(Ind As Long, Rs As ADODB.Recordset, ItemCount As Long, Optional db As Long = 1)
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

   SelectStr1 = "SELECT APM.CUSCOD, APM.CUSTYP, APM.PRENAM, APM.CUSNAM, APM.ADDR01, APM.ADDR02, APM.TELNUM, APM.CONTRACT, APM.PAYTRM, APM.PAYCOND, IST.*  "
   SelectStr1 = SelectStr1 & ",APM.SLMCOD,APM.AREACOD,APM.CRLINE,OES.SLMNAM "
   SelectStr2 = "SELECT COUNT(*)  "
   
   FromStr = "FROM ARMAS APM "
   FromStr = FromStr & "LEFT OUTER JOIN OESLM OES ON (APM.SLMCOD = OES.SLMCOD) "
   FromStr = FromStr & "LEFT OUTER JOIN ISTAB IST ON (APM.CUSTYP = IST.TYPCOD) AND (IST.TABTYP = '45') "
      
   If Trim(CUSCOD) <> "" Then
      If WildCard(CUSCOD, SubLen, NewStr) Then
         If WhereStr = "" Then
            WhereStr = "WHERE (APM.CUSCOD LIKE '" & ChangeQuote(Trim(CUSCOD)) & "')"
         Else
            WhereStr = WhereStr & "AND (APM.CUSCOD LIKE '" & ChangeQuote(Trim(CUSCOD)) & "')"
         End If
      Else
         If WhereStr = "" Then
            WhereStr = "WHERE (APM.CUSCOD = '" & ChangeQuote(Trim(CUSCOD)) & "')"
         Else
            WhereStr = WhereStr & "AND (APM.CUSCOD = '" & ChangeQuote(Trim(CUSCOD)) & "')"
         End If
      End If
   End If
   
   If Trim(CUSNAM) <> "" Then
      If WildCard(CUSNAM, SubLen, NewStr) Then
         If WhereStr = "" Then
            WhereStr = "WHERE (APM.CUSNAM LIKE '" & ChangeQuote(Trim(CUSNAM)) & "')"
         Else
            WhereStr = WhereStr & "AND (APM.CUSNAM LIKE '" & ChangeQuote(Trim(CUSNAM)) & "')"
         End If
      Else
         If WhereStr = "" Then
            WhereStr = "WHERE (APM.CUSNAM = '" & ChangeQuote(Trim(CUSNAM)) & "')"
         Else
            WhereStr = WhereStr & "AND (APM.CUSNAM = '" & ChangeQuote(Trim(CUSNAM)) & "')"
         End If
      End If
   End If
   
   If Len(CUSTYP) > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (APM.CUSTYP = '" & Trim(ChangeQuote(CUSTYP)) & "') "
      Else
         WhereStr = WhereStr & "AND (APM.CUSTYP = '" & Trim(ChangeQuote(CUSTYP)) & "') "
      End If
   End If
   
   If Len(FROM_CUSTOMER_CODE) > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (APM.CUSCOD >= '" & Trim(ChangeQuote(FROM_CUSTOMER_CODE)) & "') "
      Else
         WhereStr = WhereStr & "AND (APM.CUSCOD >= '" & Trim(ChangeQuote(FROM_CUSTOMER_CODE)) & "') "
      End If
   End If
   
   If Len(TO_CUSTOMER_CODE) > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (APM.CUSCOD <= '" & Trim(ChangeQuote(TO_CUSTOMER_CODE)) & "') "
      Else
         WhereStr = WhereStr & "AND (APM.CUSCOD <= '" & Trim(ChangeQuote(TO_CUSTOMER_CODE)) & "') "
      End If
   End If
   
   If Len(SLMCOD) > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (APM.SLMCOD = '" & Trim(ChangeQuote(SLMCOD)) & "') "
      Else
         WhereStr = WhereStr & "AND (APM.SLMCOD = '" & Trim(ChangeQuote(SLMCOD)) & "') "
      End If
   End If
   
   If Len(FROM_SALE_CODE) > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (APM.SLMCOD  >= '" & Trim(ChangeQuote(FROM_SALE_CODE)) & "') "
      Else
         WhereStr = WhereStr & "AND (APM.SLMCOD  >= '" & Trim(ChangeQuote(FROM_SALE_CODE)) & "') "
      End If
   End If
   
   If Len(TO_SALE_CODE) > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (APM.SLMCOD  <= '" & Trim(ChangeQuote(TO_SALE_CODE)) & "') "
      Else
         WhereStr = WhereStr & "AND (APM.SLMCOD  <= '" & Trim(ChangeQuote(TO_SALE_CODE)) & "') "
      End If
   End If
   
   OrderByStr = ""
   OrderTypeStr = ""
   
   If OrderBy = 1 Then
      OrderByStr = "ORDER BY APM.CUSCOD "
   ElseIf OrderBy = 2 Then
      OrderByStr = "ORDER BY APM.CUSNAM "
   Else
      OrderByStr = "ORDER BY  APM.SLMCOD,APM.CUSCOD "
   End If
   
   If OrderType = 1 Then
      OrderTypeStr = " ASC "
   ElseIf OrderType = 2 Then
      OrderTypeStr = " DESC "
   Else
      OrderTypeStr = " ASC "
   End If
   
   SQL1 = SelectStr1 & FromStr & WhereStr & OrderByStr & OrderTypeStr
   SQL2 = SelectStr2 & FromStr & WhereStr
   ''debug.print SQL1
   If db = 2 Then
   
      Call glbDatabaseMngr.GetRs2(SQL1, SQL2, True, ItemCount, Rs, glbErrorLog)
   ElseIf db = 3 Then
      Call glbDatabaseMngr.GetRs3(SQL1, SQL2, True, ItemCount, Rs, glbErrorLog)
   Else
      Call glbDatabaseMngr.GetRs(SQL1, SQL2, True, ItemCount, Rs, glbErrorLog)
   End If
End Function

Public Function PopulateFromRS(Ind As Long, Rs As ADODB.Recordset, Optional Count As Long = 0)
   CUSCOD = Trim(NVLS(Rs("CUSCOD"), ""))
   CUSTYP = Trim(NVLS(Rs("CUSTYP"), ""))
   PRENAM = Trim(NVLS(Rs("PRENAM"), ""))
   CUSNAM = Trim(NVLS(Rs("CUSNAM"), ""))
   ADDR01 = Trim(NVLS(Rs("ADDR01"), ""))
   ADDR02 = Trim(NVLS(Rs("ADDR02"), ""))
   TELNUM = Trim(NVLS(Rs("TELNUM"), ""))
   CONTRACT = Trim(NVLS(Rs("CONTRACT"), ""))
   PAYTRM = Trim(NVLD(Rs("PAYTRM"), 0))
   PAYCOND = Trim(NVLS(Rs("PAYCOND"), ""))
   
   TYPDES = NVLS(Rs("TYPDES"), "")
   
   SLMCOD = Trim(NVLS(Rs("SLMCOD"), ""))
   AREACOD = Trim(NVLS(Rs("AREACOD"), ""))
   CRLINE = Trim(NVLS(Rs("CRLINE"), ""))
   SLMNAM = Trim(NVLS(Rs("SLMNAM"), ""))
   
   If Count = 0 Then
         KEY_LOOKUP = CUSCOD
         KEY_ID = Val(CUSCOD)
   Else
         KEY_ID = Count
         KEY_LOOKUP = NVLS(Rs("CUSCOD"), "")
   End If
End Function
