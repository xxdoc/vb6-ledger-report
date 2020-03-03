VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CARTrn"
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

Public RECTYP As String
Public DOCNUM As String
Public DOCDAT  As Date
Public CUSCOD  As String
Public SLMCOD As String
Public DUEDAT  As Date
Public AMOUNT  As Double
Public TOTAL  As Double
Public DISCAMT As Double
Public RCVAMT As Double
Public SONUM As String
Public DISC As Double

Public CUSTYP As String
Public FROM_CUSTOMER_CODE As String
Public TO_CUSTOMER_CODE As String
Public FROM_SLM_CODE As String
Public TO_SLM_CODE As String

Public DESCRPT As String
Public FROM_DOC_DATE As Date
Public TO_DOC_DATE As Date
Public FROM_DUE_DATE As Date
Public TO_DUE_DATE As Date
Public FROM_INV_DATE As Date
Public TO_INV_DATE As Date
Public FROM_DOCNUM As String
Public TO_DOCNUM As String
Public FROM_CMPL_DATE As Date
Public TO_CMPL_DATE As Date
Public CUSNAM As String
Public PRENAM As String
Public RecTypeSet As String
Public PAY_FOR As String
Public ADDR01 As String
Public ADDR02 As String

Public CRLINE  As String

Public KEY_ID As Long
Public KEY_LOOKUP As String

Public ReceiptChecks As Collection
Public ReceiptItems As Collection
Public JournalItems As Collection
Public BankTrans As Collection
Public ImportExportItems As Collection
Public db2 As Boolean

Private Sub Class_Initialize()
   Set m_Conn = glbDatabaseMngr.DBConnection
   Set ReceiptChecks = New Collection
   Set JournalItems = New Collection
   Set ReceiptItems = New Collection
   Set BankTrans = New Collection
   Set ImportExportItems = New Collection
   
   FROM_DOC_DATE = -1
   TO_DOC_DATE = -1
   FROM_DUE_DATE = -1
   TO_DUE_DATE = -1
   FROM_INV_DATE = -1
   TO_INV_DATE = -1
   FROM_CMPL_DATE = -1
   TO_CMPL_DATE = -1
End Sub

'In this class we will not handle any error or any transaction
Public Function QueryData(Ind As Long, Rs As ADODB.Recordset, ItemCount As Long, Optional CountRecord As Boolean = True, Optional db As Long = 1)
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
Dim GroupByStr As String

   If Ind = 1 Then
      SelectStr1 = "SELECT APT.*, DTOS(APT.DUEDAT) DUE_DATE, DTOS(APT.DOCDAT) DOC_DATE, APM.CUSNAM, APM.PRENAM,APM.CRLINE,APM.ADDR01, APM.ADDR02 "
      SelectStr2 = "SELECT COUNT(*)  "
      
      FromStr = "FROM ARTRN APT "
      FromStr = FromStr & "LEFT OUTER JOIN ARMAS APM ON (APT.CUSCOD = APM.CUSCOD) "
   ElseIf Ind = 2 Then
      SelectStr1 = "SELECT APT.CUSCOD, SUM(APT.NETAMT) AMOUNT "
      SelectStr2 = "SELECT COUNT(*)  "
      
      FromStr = "FROM ARTRN APT "
      
      WhereStr = " WHERE (RECTYP IN " & "('3', '4', '5')" & ")"
   ElseIf Ind = 3 Then
      SelectStr1 = "SELECT APT.SONUM, SUM(APT.NETAMT) TOTAL "
      SelectStr2 = "SELECT COUNT(*)  "
      
      FromStr = "FROM ARTRN APT "
      WhereStr = " WHERE (RECTYP IN " & "('3', '4', '5')" & ")"
   ElseIf Ind = 4 Then
      SelectStr1 = "SELECT APT.SLMCOD, SUM(APT.NETAMT) AMOUNT "
      SelectStr2 = "SELECT COUNT(*)  "
      
      FromStr = "FROM ARTRN APT "
      WhereStr = " WHERE (RECTYP IN " & "('3', '4', '5')" & ")"
   ElseIf Ind = 5 Then
      SelectStr1 = "SELECT APT.CUSCOD, SUM(APT.NETAMT) AMOUNT,SUM(APT.RCVAMT) RCVAMT "
      SelectStr2 = "SELECT COUNT(*)  "
      
      FromStr = "FROM ARTRN APT "
   ElseIf Ind = 6 Then
      SelectStr1 = "SELECT SUM(APT.NETAMT) AMOUNT,SUM(APT.RCVAMT) RCVAMT "
      SelectStr2 = "SELECT COUNT(*)  "
      
      FromStr = "FROM ARTRN APT "
   ElseIf Ind = 7 Then
      SelectStr1 = "SELECT APT.CUSCOD, SUM(APT.NETAMT) AMOUNT "
      SelectStr2 = "SELECT COUNT(*)  "
      
      FromStr = "FROM ARTRN APT "
      WhereStr = " WHERE (APT.RECTYP =  '5' ) "
      
      If FROM_INV_DATE > 0 Or TO_INV_DATE > 0 Then
         FromStr = FromStr & "INNER JOIN ARTRN APT2 ON (APT.SONUM = APT2.DOCNUM) "
      End If
   ElseIf Ind = 8 Then
      SelectStr1 = "SELECT APT2.DOCNUM, SUM(APT.NETAMT) AMOUNT "
      SelectStr2 = "SELECT COUNT(*)  "
      
      FromStr = "FROM ARTRN APT "
      WhereStr = " WHERE (APT.RECTYP =  '5' ) "
      
      FromStr = FromStr & "INNER JOIN ARTRN APT2 ON (APT.SONUM = APT2.DOCNUM) "
      
   ElseIf Ind = 9 Then
      SelectStr1 = "SELECT  SUM(APT.NETAMT) AMOUNT "
      SelectStr2 = "SELECT COUNT(*)  "
      
      FromStr = "FROM ARTRN APT "
      WhereStr = " WHERE (APT.RECTYP =  '5' ) "
      
      If FROM_INV_DATE > 0 Or TO_INV_DATE > 0 Then
         FromStr = FromStr & "INNER JOIN ARTRN APT2 ON (APT.SONUM = APT2.DOCNUM) "
      End If
   ElseIf Ind = 10 Then
      SelectStr1 = "SELECT APT.*, DTOS(APT.DUEDAT) DUE_DATE, DTOS(APT.DOCDAT) DOC_DATE, APM.CUSNAM, APM.PRENAM "
      SelectStr1 = SelectStr1 & ",DTOS(OSO.SODAT) SO_DATE,SLM.SLMNAM "
      SelectStr1 = SelectStr1 & ",IST.TYPDES AREA_NAME "
      SelectStr1 = SelectStr1 & ",APM.ADDR01,APM.ADDR02,APM.TELNUM "
      
      SelectStr1 = SelectStr1 & ",STC.SEQNUM,STC.STKCOD,STC.STKDES,STC.TRNQTY,STC.NETVAL,STC.UNITPR "
      SelectStr1 = SelectStr1 & ",IST1.TYPDES UNIT_NAME "
      
      SelectStr2 = "SELECT COUNT(*)  "
      
      FromStr = "FROM ARTRN APT "
      FromStr = FromStr & "LEFT OUTER JOIN ARMAS APM ON (APT.CUSCOD = APM.CUSCOD) "
      FromStr = FromStr & "LEFT OUTER JOIN OESLM SLM ON (APT.SLMCOD = SLM.SLMCOD) "
      FromStr = FromStr & "LEFT OUTER JOIN OESO OSO ON (APT.SONUM = OSO.SONUM) "
      FromStr = FromStr & "LEFT OUTER JOIN ISTAB IST ON (IST.TYPCOD = APT.AREACOD AND IST.TABTYP = '40' ) "
      
      FromStr = FromStr & "LEFT OUTER JOIN STCRD STC ON (STC.DOCNUM = APT.DOCNUM) "
      FromStr = FromStr & "LEFT OUTER JOIN ISTAB IST1 ON (IST1.TYPCOD = STC.TQUCOD AND IST1.TABTYP = '20' ) "
   ElseIf Ind = 11 Then
      SelectStr1 = "SELECT APT.DOCNUM, DTOS(APT.DOCDAT) DOC_DATE, APT.AMOUNT,DTOS(APT.DUEDAT) DUEDAT "
      SelectStr2 = "SELECT COUNT(*)  "
      
      FromStr = "FROM ARTRN APT "
  ElseIf Ind = 12 Then
      SelectStr1 = "SELECT APT.DOCNUM, DTOS(APT.CMPLDAT) DOC_DATE "
      'SelectStr1 = "SELECT APT.DOCNUM, APT.CMPLDAT DOC_DATE "
      SelectStr2 = "SELECT COUNT(*)  "
      
      FromStr = "FROM ARTRN APT "
      
      WhereStr = " WHERE (RECTYP IN " & "('3')" & ") "            ' IV
   ElseIf Ind = 13 Then
      SelectStr1 = "SELECT APT.DOCNUM, DTOS(APT.CMPLDAT) DOC_DATE "
      SelectStr2 = "SELECT COUNT(*)  "
      
      FromStr = "FROM ARTRN APT "
      
      WhereStr = " WHERE (RECTYP IN " & "('9')" & ") "                ' RE
   End If
   
   If Trim(DOCNUM) <> "" Then
      If WildCard(DOCNUM, SubLen, NewStr) Then
         If WhereStr = "" Then
            WhereStr = "WHERE (APT.DOCNUM = '" & ChangeQuote(Trim(DOCNUM)) & "')"
         Else
            WhereStr = WhereStr & "AND (APT.DOCNUM = '" & ChangeQuote(Trim(DOCNUM)) & "')"
         End If
      Else
         If WhereStr = "" Then
            WhereStr = "WHERE (APT.DOCNUM = '" & ChangeQuote(Trim(DOCNUM)) & "')"
         Else
            WhereStr = WhereStr & "AND (APT.DOCNUM = '" & ChangeQuote(Trim(DOCNUM)) & "')"
         End If
      End If
   End If
   
   If Trim(FROM_DOCNUM) <> "" Then
      If WildCard(FROM_DOCNUM, SubLen, NewStr) Then
         If WhereStr = "" Then
            WhereStr = "WHERE (APT.DOCNUM >= '" & ChangeQuote(Trim(FROM_DOCNUM)) & "')"
         Else
            WhereStr = WhereStr & "AND (APT.DOCNUM >= '" & ChangeQuote(Trim(FROM_DOCNUM)) & "')"
         End If
      Else
         If WhereStr = "" Then
            WhereStr = "WHERE (APT.DOCNUM >= '" & ChangeQuote(Trim(FROM_DOCNUM)) & "')"
         Else
            WhereStr = WhereStr & "AND (APT.DOCNUM >= '" & ChangeQuote(Trim(FROM_DOCNUM)) & "')"
         End If
      End If
   End If
   
   If Trim(TO_DOCNUM) <> "" Then
      If WildCard(TO_DOCNUM, SubLen, NewStr) Then
         If WhereStr = "" Then
            WhereStr = "WHERE (APT.DOCNUM <= '" & ChangeQuote(Trim(TO_DOCNUM)) & "')"
         Else
            WhereStr = WhereStr & "AND (APT.DOCNUM <= '" & ChangeQuote(Trim(TO_DOCNUM)) & "')"
         End If
      Else
         If WhereStr = "" Then
            WhereStr = "WHERE (APT.DOCNUM <= '" & ChangeQuote(Trim(TO_DOCNUM)) & "')"
         Else
            WhereStr = WhereStr & "AND (APT.DOCNUM <= '" & ChangeQuote(Trim(TO_DOCNUM)) & "')"
         End If
      End If
   End If
   
   If Trim(CUSCOD) <> "" Then
      If WildCard(CUSCOD, SubLen, NewStr) Then
         If WhereStr = "" Then
            WhereStr = "WHERE (APT.CUSCOD LIKE '" & ChangeQuote(Trim(CUSCOD)) & "')"
         Else
            WhereStr = WhereStr & "AND (APT.CUSCOD LIKE '" & ChangeQuote(Trim(CUSCOD)) & "')"
         End If
      Else
         If WhereStr = "" Then
            WhereStr = "WHERE (APT.CUSCOD = '" & ChangeQuote(Trim(CUSCOD)) & "')"
         Else
            WhereStr = WhereStr & "AND (APT.CUSCOD = '" & ChangeQuote(Trim(CUSCOD)) & "')"
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
            
   If Trim(SLMCOD) <> "" Then
      If WildCard(SLMCOD, SubLen, NewStr) Then
         If WhereStr = "" Then
            WhereStr = "WHERE (APT.SLMCOD LIKE '" & ChangeQuote(Trim(SLMCOD)) & "')"
         Else
            WhereStr = WhereStr & "AND (APT.SLMCOD LIKE '" & ChangeQuote(Trim(SLMCOD)) & "')"
         End If
      Else
         If WhereStr = "" Then
            WhereStr = "WHERE (APT.SLMCOD = '" & ChangeQuote(Trim(SLMCOD)) & "')"
         Else
            WhereStr = WhereStr & "AND (APT.SLMCOD = '" & ChangeQuote(Trim(SLMCOD)) & "')"
         End If
      End If
   End If
   
   If Trim(RECTYP) <> "" Then
      If WildCard(RECTYP, SubLen, NewStr) Then
         If WhereStr = "" Then
            WhereStr = "WHERE (RECTYP LIKE '" & ChangeQuote(Trim(RECTYP)) & "')"
         Else
            WhereStr = WhereStr & "AND (RECTYP LIKE '" & ChangeQuote(Trim(RECTYP)) & "')"
         End If
      Else
         If WhereStr = "" Then
            WhereStr = "WHERE (RECTYP = '" & RECTYP & "') "
         Else
            WhereStr = WhereStr & "AND (RECTYP = '" & RECTYP & "')"
         End If
      End If
   End If
                        
   If Trim(RecTypeSet) <> "" Then
      If WhereStr = "" Then
         WhereStr = "WHERE (RECTYP IN " & RecTypeSet & ") "
      Else
         WhereStr = WhereStr & "AND (RECTYP IN " & RecTypeSet & ")"
      End If
   End If

    If FROM_DUE_DATE > -1 Then
      TempDate = DateToStringIntLow(Trim(FROM_DUE_DATE))
      If WhereStr = "" Then
         WhereStr = "WHERE (DTOS(DUEDAT) >= '" & ChangeQuote(TempDate) & "')"
      Else
         WhereStr = WhereStr & "AND (DTOS(DUEDAT) >= '" & ChangeQuote(TempDate) & "')"
      End If
   End If

   If TO_DUE_DATE > -1 Then
      TempDate = DateToStringIntHi(Trim(TO_DUE_DATE))
      If WhereStr = "" Then
         WhereStr = "WHERE (DTOS(DUEDAT) <= '" & ChangeQuote(TempDate) & "')"
      Else
         WhereStr = WhereStr & "AND (DTOS(DUEDAT) <= '" & ChangeQuote(TempDate) & "')"
      End If
   End If
      
     If FROM_DOC_DATE > -1 Then
      TempDate = DateToStringIntLow(Trim(FROM_DOC_DATE))
      If WhereStr = "" Then
         WhereStr = " WHERE (DTOS(APT.DOCDAT) >= '" & ChangeQuote(TempDate) & "')"
      Else
         WhereStr = WhereStr & " AND (DTOS(APT.DOCDAT) >= '" & ChangeQuote(TempDate) & "')"
      End If
             End If

   If TO_DOC_DATE > -1 Then
      TempDate = DateToStringIntHi(Trim(TO_DOC_DATE))
      If WhereStr = "" Then
         WhereStr = " WHERE (DTOS(APT.DOCDAT) <= '" & ChangeQuote(TempDate) & "')"
      Else
         WhereStr = WhereStr & " AND (DTOS(APT.DOCDAT) <= '" & ChangeQuote(TempDate) & "')"
      End If
   End If
   
   If FROM_CMPL_DATE > -1 Then
      TempDate = DateToStringIntLow(Trim(FROM_CMPL_DATE))
      If WhereStr = "" Then
         WhereStr = " WHERE (DTOS(APT.CMPLDAT) >= '" & ChangeQuote(TempDate) & "')"
      Else
         WhereStr = WhereStr & " AND (DTOS(APT.CMPLDAT) >= '" & ChangeQuote(TempDate) & "')"
      End If
   End If

   If TO_CMPL_DATE > -1 Then
      TempDate = DateToStringIntHi(Trim(TO_CMPL_DATE))
      If WhereStr = "" Then
         WhereStr = " WHERE (DTOS(APT.CMPLDAT) <= '" & ChangeQuote(TempDate) & "')"
      Else
         WhereStr = WhereStr & " AND (DTOS(APT.CMPLDAT) <= '" & ChangeQuote(TempDate) & "')"
      End If
   End If
   
      
   If FROM_INV_DATE > -1 Then
      TempDate = DateToStringIntLow(Trim(FROM_INV_DATE))
      If WhereStr = "" Then
         WhereStr = "WHERE (DTOS(APT2.DOCDAT) >= '" & ChangeQuote(TempDate) & "')"
      Else
         WhereStr = WhereStr & "AND (DTOS(APT2.DOCDAT) >= '" & ChangeQuote(TempDate) & "')"
      End If
   End If

   If TO_INV_DATE > -1 Then
      TempDate = DateToStringIntHi(Trim(TO_INV_DATE))
      If WhereStr = "" Then
         WhereStr = "WHERE (DTOS(APT2.DOCDAT) <= '" & ChangeQuote(TempDate) & "')"
      Else
         WhereStr = WhereStr & "AND (DTOS(APT2.DOCDAT) <= '" & ChangeQuote(TempDate) & "')"
      End If
   End If
   
   If Len(FROM_CUSTOMER_CODE) > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (APT.CUSCOD >= '" & Trim(ChangeQuote(FROM_CUSTOMER_CODE)) & "') "
      Else
         WhereStr = WhereStr & "AND (APT.CUSCOD >= '" & Trim(ChangeQuote(FROM_CUSTOMER_CODE)) & "') "
      End If
   End If
   
   If Len(TO_CUSTOMER_CODE) > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (APT.CUSCOD <= '" & Trim(ChangeQuote(TO_CUSTOMER_CODE)) & "') "
      Else
         WhereStr = WhereStr & "AND (APT.CUSCOD <= '" & Trim(ChangeQuote(TO_CUSTOMER_CODE)) & "') "
      End If
   End If
   
   If Len(FROM_SLM_CODE) > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (APT.SLMCOD  >= '" & Trim(ChangeQuote(FROM_SLM_CODE)) & "') "
      Else
         WhereStr = WhereStr & "AND (APT.SLMCOD  >= '" & Trim(ChangeQuote(FROM_SLM_CODE)) & "') "
      End If
   End If
   
   If Len(TO_SLM_CODE) > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (APT.SLMCOD  <= '" & Trim(ChangeQuote(TO_SLM_CODE)) & "') "
      Else
         WhereStr = WhereStr & "AND (APT.SLMCOD  <= '" & Trim(ChangeQuote(TO_SLM_CODE)) & "') "
      End If
   End If
   
   GroupByStr = ""
   If Ind = 2 Then
      GroupByStr = "GROUP BY  APT.CUSCOD "
   ElseIf Ind = 3 Then
      GroupByStr = "GROUP BY  APT.SONUM "
   ElseIf Ind = 4 Then
      GroupByStr = "GROUP BY  APT.SLMCOD "
   ElseIf Ind = 5 Then
      GroupByStr = "GROUP BY  APT.CUSCOD "
   ElseIf Ind = 7 Then
      GroupByStr = "GROUP BY  APT.CUSCOD "
   ElseIf Ind = 8 Then
      GroupByStr = "GROUP BY  APT2.DOCNUM "
   End If
   
   OrderByStr = ""
   OrderTypeStr = ""
   
   If Ind = 1 Then
      If OrderBy = 1 Then
         OrderByStr = "ORDER BY DUEDAT ASC, DOCNUM "
      ElseIf OrderBy = 2 Then
         OrderByStr = "ORDER BY DUEDAT ASC, DOCNUM "
      ElseIf OrderBy = 3 Then
         OrderByStr = "ORDER BY APT.CUSCOD ASC,APT.DUEDAT ASC, DOCNUM "
      ElseIf OrderBy = 4 Then
        OrderByStr = "ORDER BY APT.CUSCOD ASC,DOCDAT ASC, DOCNUM "
      Else
         OrderByStr = "ORDER BY DUEDAT ASC, DOCNUM "
      End If
   ElseIf Ind = 2 Then
      OrderByStr = "ORDER BY APT.CUSCOD "
   ElseIf Ind = 3 Then
      OrderByStr = "ORDER BY APT.SONUM "
   ElseIf Ind = 4 Then
      OrderByStr = "ORDER BY APT.SLMCOD "
   ElseIf Ind = 5 Then
      OrderByStr = "ORDER BY APT.CUSCOD "
   ElseIf Ind = 10 Then
      OrderByStr = "ORDER BY APT.DOCNUM ASC,STC.SEQNUM "
   ElseIf Ind = 11 Then
      OrderByStr = "ORDER BY APT.DOCNUM "
     ElseIf Ind = 12 Then
      OrderByStr = "ORDER BY APT.DOCNUM "
     ElseIf Ind = 13 Then
      OrderByStr = "ORDER BY APT.DOCNUM "
   End If
   
   If Len(OrderByStr) > 0 Then
      If OrderType = 1 Then
         OrderTypeStr = "ASC "
      ElseIf OrderType = 2 Then
         OrderTypeStr = "DESC "
      Else
         OrderTypeStr = "ASC "
      End If
   End If
   SQL1 = SelectStr1 & FromStr & WhereStr & GroupByStr & OrderByStr & OrderTypeStr
   SQL2 = SelectStr2 & FromStr & WhereStr
 '  'debug.print SQL1
   If db2 = True Then
      Call glbDatabaseMngr.GetRs2(SQL1, SQL2, CountRecord, ItemCount, Rs, glbErrorLog)
   Else
     If db = 1 Then
        Call glbDatabaseMngr.GetRs(SQL1, SQL2, CountRecord, ItemCount, Rs, glbErrorLog)
     ElseIf db = 2 Then
         Call glbDatabaseMngr.GetRs2(SQL1, SQL2, CountRecord, ItemCount, Rs, glbErrorLog)
      ElseIf db = 3 Then
        Call glbDatabaseMngr.GetRs3(SQL1, SQL2, CountRecord, ItemCount, Rs, glbErrorLog)
      Else
         Call glbDatabaseMngr.GetRs(SQL1, SQL2, CountRecord, ItemCount, Rs, glbErrorLog)
      End If
   End If
End Function

Public Function PopulateFromRS(Ind As Long, Rs As ADODB.Recordset, Optional Count As Long)
   If Ind = 1 Then
      RECTYP = NVLS(Rs("RECTYP"), "")
      DOCNUM = NVLS(Rs("DOCNUM"), "")
      DOCDAT = InternalDateToDate(NVLS(Rs("DOC_DATE"), ""))
      CUSCOD = NVLS(Rs("CUSCOD"), "")
      DUEDAT = InternalDateToDate(NVLS(Rs("DUE_DATE"), ""))
      AMOUNT = NVLD(Rs("NETAMT"), 0)
      TOTAL = NVLD(Rs("TOTAL"), 0)
      CUSNAM = NVLS(Rs("CUSNAM"), "")
      PRENAM = NVLS(Rs("PRENAM"), "")
      RCVAMT = NVLD(Rs("RCVAMT"), 0)
      DISC = NVLD(Rs("DISC"), 0)
      DISCAMT = NVLD(Rs("DISCAMT"), 0)
      ADDR01 = NVLS(Rs("ADDR01"), "")
      ADDR02 = NVLS(Rs("ADDR02"), "")
      
      CRLINE = NVLD(Rs("CRLINE"), 0)
      KEY_LOOKUP = CUSCOD
      KEY_ID = Val(CUSCOD)
   ElseIf Ind = 2 Then
      CUSCOD = NVLS(Rs("CUSCOD"), "")
      AMOUNT = NVLD(Rs("AMOUNT"), 0)
   ElseIf Ind = 3 Then
      SONUM = NVLS(Rs("SONUM"), "")
      TOTAL = NVLD(Rs("TOTAL"), 0)
   ElseIf Ind = 4 Then
      SLMCOD = NVLS(Rs("SLMCOD"), "")
      AMOUNT = NVLD(Rs("AMOUNT"), 0)
   ElseIf Ind = 5 Then
      CUSCOD = NVLS(Rs("CUSCOD"), "")
      AMOUNT = NVLD(Rs("AMOUNT"), 0)
      RCVAMT = NVLD(Rs("RCVAMT"), 0)
   ElseIf Ind = 6 Then
      AMOUNT = NVLD(Rs("AMOUNT"), 0)
      RCVAMT = NVLD(Rs("RCVAMT"), 0)
   ElseIf Ind = 7 Then
      CUSCOD = NVLS(Rs("CUSCOD"), "")
      AMOUNT = NVLD(Rs("AMOUNT"), 0)
   ElseIf Ind = 8 Then
      DOCNUM = NVLS(Rs("DOCNUM"), "")
      AMOUNT = NVLD(Rs("AMOUNT"), 0)
   ElseIf Ind = 9 Then
      AMOUNT = NVLD(Rs("AMOUNT"), 0)
   ElseIf Ind = 10 Then
      DOCNUM = NVLS(Rs("DOCNUM"), "")
      DOCDAT = InternalDateToDate(NVLS(Rs("DOC_DATE"), ""))
      AMOUNT = NVLD(Rs("AMOUNT"), 0)
   ElseIf Ind = 11 Then
      DOCNUM = NVLS(Rs("DOCNUM"), "")
      DOCDAT = InternalDateToDate(NVLS(Rs("DOC_DATE"), ""))
      AMOUNT = NVLD(Rs("AMOUNT"), 0)
      DUEDAT = InternalDateToDate(NVLS(Rs("DUEDAT"), ""))

      KEY_ID = Count
      KEY_LOOKUP = NVLS(Rs("DOCNUM"), "")
   ElseIf Ind = 12 Then
      DOCNUM = NVLS(Rs("DOCNUM"), "")
      DOCDAT = InternalDateToDate(NVLS(Rs("DOC_DATE"), ""))
  ElseIf Ind = 13 Then
      DOCNUM = NVLS(Rs("DOCNUM"), "")
      DOCDAT = InternalDateToDate(NVLS(Rs("DOC_DATE"), ""))
   End If
End Function

Private Sub Class_Terminate()
   Set ReceiptChecks = Nothing
   Set JournalItems = Nothing
   Set ReceiptItems = Nothing
   Set BankTrans = Nothing
   Set ImportExportItems = Nothing
End Sub
Public Sub GenerateArHeader(FileID As Long, Rs As ADODB.Recordset)
Dim TempStr As String
   
   TempStr = "BD;"
   TempStr = TempStr & NVLS(Rs("DOCNUM"), "") & ";"
   TempStr = TempStr & ReFormatDate2(NVLS(Rs("DOC_DATE"), "")) & ";"
   
   TempStr = TempStr & NVLS(Rs("CUSCOD"), "") & ";"
   TempStr = TempStr & NVLS(Rs("PRENAM"), "") & ";"
   TempStr = TempStr & NVLS(Rs("CUSNAM"), "") & ";"
   
   TempStr = TempStr & NVLD(Rs("PAYTRM"), 0) & ";"
   TempStr = TempStr & ReFormatDate2(NVLS(Rs("DUE_DATE"), "")) & ";"
   
   TempStr = TempStr & NVLS(Rs("SONUM"), "") & ";"
   TempStr = TempStr & ReFormatDate2(NVLS(Rs("SO_DATE"), "")) & ";"
   
   TempStr = TempStr & NVLS(Rs("SLMCOD"), "") & ";"
   TempStr = TempStr & NVLS(Rs("SLMNAM"), "") & ";"
   TempStr = TempStr & NVLS(Rs("AREA_NAME"), "") & ";"
   TempStr = TempStr & NVLS(Rs("YOUREF"), "") & ";"
   
   TempStr = TempStr & NVLD(Rs("AMOUNT"), 0) & ";"
   TempStr = TempStr & NVLS(Rs("DISC"), "") & ";"
   TempStr = TempStr & NVLD(Rs("DISCAMT"), 0) & ";"
   TempStr = TempStr & NVLD(Rs("AFTDISC"), 0) & ";"
   
   TempStr = TempStr & NVLD(Rs("VATRAT"), 0) & ";"
   TempStr = TempStr & NVLD(Rs("VATAMT"), 0) & ";"
   TempStr = TempStr & NVLD(Rs("NETAMT"), 0) & ";"
   
   TempStr = TempStr & NVLS(Rs("ADDR01"), "") & ";"
   TempStr = TempStr & NVLS(Rs("ADDR02"), "") & ";"
   TempStr = TempStr & NVLS(Rs("TELNUM"), "") & ";"
   
   Print #FileID, TempStr
End Sub
Public Sub GenerateArDetail(FileID As Long, Rs As ADODB.Recordset)
Dim TempStr As String
   
   TempStr = "STC;"
   TempStr = TempStr & NVLS(Rs("SEQNUM"), "") & ";"
   
   TempStr = TempStr & NVLS(Rs("STKCOD"), "") & ";"
   TempStr = TempStr & NVLS(Rs("STKDES"), "") & ";"
   
   TempStr = TempStr & NVLD(Rs("TRNQTY"), 0) & ";"
   TempStr = TempStr & NVLD(Rs("UNITPR"), 0) & ";"
   TempStr = TempStr & NVLD(Rs("NETVAL"), 0) & ";"
      
   TempStr = TempStr & NVLS(Rs("UNIT_NAME"), "") & ";"
   
   Print #FileID, TempStr
End Sub
