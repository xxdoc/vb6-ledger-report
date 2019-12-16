VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CARRcIt"
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

Public RCPNUM   As String
Public DOCNUM  As String
Public RECTYP As String
Public RCVAMT  As Double
Public VATAMT As Double
'Public RCPNUM As String

Public FROM_PAY_DATE As Date
Public TO_PAY_DATE As Date
Public FROM_DOC_DATE As Date
Public TO_DOC_DATE As Date
Public CUSCOD As String
Public SLMCOD  As String

Public FROM_CUSTOMER_CODE As String
Public TO_CUSTOMER_CODE As String
Public FROM_SALE_CODE As String
Public TO_SALE_CODE As String

Public DOCDAT As Date
Public PAYDAT As Date

Public AMOUNT As Double     'ยอดบิล
Public CUSNAM As String
Public SLMNAM As String

Public RecTypeSet As String
Public KEY_ID As Long
Public KEY_LOOKUP As String
Public db2 As Boolean

Private Sub Class_Initialize()
   Set m_Conn = glbDatabaseMngr.DBConnection
   FROM_PAY_DATE = -1
   TO_PAY_DATE = -1
   FROM_DOC_DATE = -1
   TO_DOC_DATE = -1
End Sub

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
Dim TempDate As String
Dim GroupByStr As String

   If Ind = 1 Then
      SelectStr1 = "SELECT ART.* "
      SelectStr2 = "SELECT COUNT(*)  "
      
      FromStr = "FROM ARRCPIT ART "
   ElseIf Ind = 2 Then
      SelectStr1 = "SELECT ART.DOCNUM, SUM(ART.RCVAMT) RCVAMT, SUM(ART.VATAMT) VATAMT "
      SelectStr2 = "SELECT COUNT(*)  "
      
      FromStr = "FROM ARRCPIT ART "
      FromStr = FromStr & "LEFT OUTER JOIN ARTRN APT ON (ART.RCPNUM = APT.DOCNUM) "
   ElseIf Ind = 3 Then
      SelectStr1 = "SELECT APT.CUSCOD, SUM(ART.RCVAMT) RCVAMT, SUM(ART.VATAMT) VATAMT "
      SelectStr2 = "SELECT COUNT(*)  "
      
      FromStr = "FROM ARRCPIT ART "
      FromStr = FromStr & "LEFT OUTER JOIN ARTRN APT ON (ART.RCPNUM = APT.DOCNUM) "
      FromStr = FromStr & "LEFT OUTER JOIN ARTRN APT2 ON (ART.DOCNUM = APT2.DOCNUM) "
   ElseIf Ind = 4 Then
      SelectStr1 = "SELECT APT.SLMCOD, SUM(ART.RCVAMT) RCVAMT, SUM(ART.VATAMT) VATAMT "
      SelectStr2 = "SELECT COUNT(*)  "
      
      FromStr = "FROM ARRCPIT ART "
      FromStr = FromStr & "LEFT OUTER JOIN ARTRN APT ON (ART.RCPNUM = APT.DOCNUM) "
      FromStr = FromStr & "LEFT OUTER JOIN ARTRN APT2 ON (ART.DOCNUM = APT2.DOCNUM) "
   ElseIf Ind = 5 Then
      SelectStr1 = "SELECT APT.CUSCOD, SUM(ART.RCVAMT) RCVAMT "
      SelectStr2 = "SELECT COUNT(*)  "
      
      FromStr = "FROM ARRCPIT ART "
      FromStr = FromStr & "LEFT OUTER JOIN ARTRN APT ON (ART.RCPNUM = APT.DOCNUM) "
      If FROM_DOC_DATE > 0 Or TO_DOC_DATE > 0 Then
         FromStr = FromStr & "LEFT OUTER JOIN ARTRN APT2 ON (ART.DOCNUM = APT2.DOCNUM) "
      End If
   ElseIf Ind = 6 Then
      SelectStr1 = "SELECT DTOS(APT2.DOCDAT) DOC_DATE  "
      SelectStr2 = "SELECT COUNT(*)  "
      
      FromStr = "FROM ARRCPIT ART "
      FromStr = FromStr & "LEFT OUTER JOIN ARTRN APT ON (ART.RCPNUM = APT.DOCNUM) "
      FromStr = FromStr & "LEFT OUTER JOIN ARTRN APT2 ON (ART.DOCNUM = APT2.DOCNUM) "
   ElseIf Ind = 7 Then
      SelectStr1 = "SELECT DTOS(APT2.DOCDAT) DOC_DATE,DTOS(APT.DOCDAT) PAY_DATE ,SUM(ART.RCVAMT) RCVAMT "
      SelectStr2 = "SELECT COUNT(*)  "
      
      FromStr = "FROM ARRCPIT ART "
      FromStr = FromStr & "LEFT OUTER JOIN ARTRN APT ON (ART.RCPNUM = APT.DOCNUM) "
      FromStr = FromStr & "LEFT OUTER JOIN ARTRN APT2 ON (ART.DOCNUM = APT2.DOCNUM) "
   ElseIf Ind = 8 Then
      SelectStr1 = "SELECT SUM(ART.RCVAMT) RCVAMT "
      SelectStr2 = "SELECT COUNT(*)  "
      
      FromStr = "FROM ARRCPIT ART "
      FromStr = FromStr & "LEFT OUTER JOIN ARTRN APT ON (ART.RCPNUM = APT.DOCNUM) "
      FromStr = FromStr & "LEFT OUTER JOIN ARTRN APT2 ON (ART.DOCNUM = APT2.DOCNUM) "
     ElseIf Ind = 9 Then                        ' SUM IV
      SelectStr1 = "SELECT SUM(ART.RCVAMT) RCVAMT, ART.DOCNUM ,MAX(RCPNUM) RCPNUM "
      SelectStr2 = "SELECT COUNT(*)  "
      
      FromStr = "FROM ARRCPIT ART "
   ElseIf Ind = 10 Then
      SelectStr1 = "SELECT APT.DOCNUM RCPNUM,DTOS(APT.DOCDAT) PAYDAT "
      SelectStr1 = SelectStr1 & ",ART.RCVAMT,ART.RECTYP "
      SelectStr1 = SelectStr1 & ",APT2.DOCNUM,DTOS(APT2.DOCDAT) DOCDAT,APT2.AMOUNT "
      
      SelectStr1 = SelectStr1 & ",APT2.CUSCOD,APM.CUSNAM "
      SelectStr1 = SelectStr1 & ",APT2.SLMCOD,SLM.SLMNAM "
      
      SelectStr2 = "SELECT COUNT(*)  "
      
      FromStr = "FROM ARRCPIT ART "
      FromStr = FromStr & "LEFT OUTER JOIN ARTRN APT ON (ART.RCPNUM = APT.DOCNUM) "
      FromStr = FromStr & "LEFT OUTER JOIN ARTRN APT2 ON (ART.DOCNUM = APT2.DOCNUM) "
      
      FromStr = FromStr & "LEFT OUTER JOIN ARMAS APM ON (APT.CUSCOD = APM.CUSCOD) "
      FromStr = FromStr & "LEFT OUTER JOIN OESLM SLM ON (APT.SLMCOD = SLM.SLMCOD) "
      
   End If
   
   If Trim(DOCNUM) <> "" Then
      If WildCard(DOCNUM, SubLen, NewStr) Then
         If WhereStr = "" Then
            WhereStr = "WHERE (DOCNUM LIKE '" & ChangeQuote(Trim(DOCNUM)) & "')"
         Else
            WhereStr = WhereStr & "AND (DOCNUM LIKE '" & ChangeQuote(Trim(DOCNUM)) & "')"
         End If
      Else
         If WhereStr = "" Then
            WhereStr = "WHERE (DOCNUM = '" & ChangeQuote(Trim(DOCNUM)) & "')"
         Else
            WhereStr = WhereStr & "AND (DOCNUM = '" & ChangeQuote(Trim(DOCNUM)) & "')"
         End If
      End If
   End If
            
   If Trim(RCPNUM) <> "" Then
      If WildCard(RCPNUM, SubLen, NewStr) Then
         If WhereStr = "" Then
            WhereStr = "WHERE (RCPNUM LIKE '" & ChangeQuote(Trim(RCPNUM)) & "')"
         Else
            WhereStr = WhereStr & "AND (RCPNUM LIKE '" & ChangeQuote(Trim(RCPNUM)) & "')"
         End If
      Else
         If WhereStr = "" Then
            WhereStr = "WHERE (RCPNUM = '" & ChangeQuote(Trim(RCPNUM)) & "')"
         Else
            WhereStr = WhereStr & "AND (RCPNUM = '" & ChangeQuote(Trim(RCPNUM)) & "')"
         End If
      End If
   End If
            
   If Trim(RECTYP) <> "" Then
      If WildCard(RECTYP, SubLen, NewStr) Then
         If WhereStr = "" Then
            WhereStr = "WHERE (ART.RECTYP LIKE '" & ChangeQuote(Trim(RECTYP)) & "')"
         Else
            WhereStr = WhereStr & "AND (ART.RECTYP LIKE '" & ChangeQuote(Trim(RECTYP)) & "')"
         End If
      Else
         If WhereStr = "" Then
            WhereStr = "WHERE (ART.RECTYP = '" & ChangeQuote(Trim(RECTYP)) & "')"
         Else
            WhereStr = WhereStr & "AND (ART.RECTYP = '" & ChangeQuote(Trim(RECTYP)) & "')"
         End If
      End If
   End If

    If FROM_PAY_DATE > -1 Then
      TempDate = DateToStringIntLow(Trim(FROM_PAY_DATE))
      If WhereStr = "" Then
         WhereStr = "WHERE (DTOS(APT.DOCDAT) >= '" & ChangeQuote(TempDate) & "')"
      Else
         WhereStr = WhereStr & "AND (DTOS(APT.DOCDAT) >= '" & ChangeQuote(TempDate) & "')"
      End If
   End If

   If TO_PAY_DATE > -1 Then
      TempDate = DateToStringIntHi(Trim(TO_PAY_DATE))
      If WhereStr = "" Then
         WhereStr = "WHERE ((DTOS(APT.DOCDAT) <= '" & ChangeQuote(TempDate) & "') OR (APT.DOCDAT IS NULL)) "
      Else
         WhereStr = WhereStr & "AND ((DTOS(APT.DOCDAT) <= '" & ChangeQuote(TempDate) & "') OR (APT.DOCDAT IS NULL)) "
      End If
   End If

   If FROM_DOC_DATE > -1 Then
      TempDate = DateToStringIntLow(Trim(FROM_DOC_DATE))
      If WhereStr = "" Then
         WhereStr = "WHERE (DTOS(APT2.DOCDAT) >= '" & ChangeQuote(TempDate) & "')"
      Else
         WhereStr = WhereStr & "AND (DTOS(APT2.DOCDAT) >= '" & ChangeQuote(TempDate) & "')"
      End If
   End If

   If TO_DOC_DATE > -1 Then
      TempDate = DateToStringIntHi(Trim(TO_DOC_DATE))
      If WhereStr = "" Then
         WhereStr = "WHERE ((DTOS(APT2.DOCDAT) <= '" & ChangeQuote(TempDate) & "') OR (APT2.DOCDAT IS NULL)) "
      Else
         WhereStr = WhereStr & "AND ((DTOS(APT2.DOCDAT) <= '" & ChangeQuote(TempDate) & "') OR (APT2.DOCDAT IS NULL)) "
      End If
   End If
   
   If Trim(RecTypeSet) <> "" Then
      If WhereStr = "" Then
         WhereStr = "WHERE (ART.RECTYP IN " & RecTypeSet & ") "
      Else
         WhereStr = WhereStr & "AND (ART.RECTYP IN " & RecTypeSet & ")"
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
   
   If Len(FROM_SALE_CODE) > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (APT.SLMCOD >= '" & Trim(ChangeQuote(FROM_SALE_CODE)) & "') "
      Else
         WhereStr = WhereStr & "AND (APT.SLMCOD >= '" & Trim(ChangeQuote(FROM_SALE_CODE)) & "') "
      End If
   End If
   
   If Len(TO_SALE_CODE) > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (APT.SLMCOD <= '" & Trim(ChangeQuote(TO_SALE_CODE)) & "') "
      Else
         WhereStr = WhereStr & "AND (APT.SLMCOD <= '" & Trim(ChangeQuote(TO_SALE_CODE)) & "') "
      End If
   End If
   
   GroupByStr = ""
   If Ind = 2 Then
      GroupByStr = "GROUP BY  ART.DOCNUM "
   ElseIf Ind = 3 Then
      GroupByStr = "GROUP BY  APT.CUSCOD "
   ElseIf Ind = 4 Then
      GroupByStr = "GROUP BY  APT.SLMCOD "
   ElseIf Ind = 5 Then
      GroupByStr = "GROUP BY  APT.CUSCOD "
   ElseIf Ind = 7 Then
      GroupByStr = "GROUP BY  APT2.DOCDAT,APT.DOCDAT "
   ElseIf Ind = 9 Then
      GroupByStr = "GROUP BY  ART.DOCNUM "
   End If
   
   OrderByStr = ""
   OrderTypeStr = ""
   
   If Ind = 1 Then
      If OrderBy = 1 Then
         OrderByStr = " ORDER BY RCPNUM ASC, DOCNUM "
      ElseIf OrderBy = 2 Then
         OrderByStr = " ORDER BY RCPNUM ASC, DOCNUM "
      Else
         OrderByStr = " ORDER BY RCPNUM ASC, DOCNUM "
      End If
   ElseIf Ind = 2 Then
      OrderByStr = " ORDER BY ART.DOCNUM "
   ElseIf Ind = 3 Then
      OrderByStr = " ORDER BY APT.CUSCOD "
   ElseIf Ind = 4 Then
      OrderByStr = " ORDER BY APT.SLMCOD "
   ElseIf Ind = 5 Then
      OrderByStr = " ORDER BY  APT.CUSCOD "
   ElseIf Ind = 6 Then
      OrderByStr = " ORDER BY  APT2.DOCDAT "
      OrderType = 1  'บังคับ OrderType เป็น 1 นะครับเนื่องจากจะใช้ในโปรแกรม
   ElseIf Ind = 7 Then
      OrderByStr = " ORDER BY  APT.DOCDAT ASC,APT2.DOCDAT "
      OrderType = 1  'บังคับ OrderType เป็น 1 นะครับเนื่องจากจะใช้ในโปรแกรม
   ElseIf Ind = 9 Then
      OrderByStr = " ORDER BY ART.DOCNUM "
   ElseIf Ind = 10 Then
      OrderByStr = " ORDER BY APT.SLMCOD ASC,APT.DOCDAT ASC,APT.DOCNUM "
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
   If db2 = True Then
       Call glbDatabaseMngr.GetRs2(SQL1, SQL2, True, ItemCount, Rs, glbErrorLog)
   Else
      If db = 1 Then
         Call glbDatabaseMngr.GetRs(SQL1, SQL2, True, ItemCount, Rs, glbErrorLog)
      ElseIf db = 2 Then
         Call glbDatabaseMngr.GetRs2(SQL1, SQL2, True, ItemCount, Rs, glbErrorLog)
       ElseIf db = 3 Then
         Call glbDatabaseMngr.GetRs3(SQL1, SQL2, True, ItemCount, Rs, glbErrorLog)
      Else
         Call glbDatabaseMngr.GetRs(SQL1, SQL2, True, ItemCount, Rs, glbErrorLog)
      End If
   End If
End Function

Public Function PopulateFromRS(Ind As Long, Rs As ADODB.Recordset)
   If Ind = 1 Then
      RECTYP = NVLS(Rs("RECTYP"), "")
      RCPNUM = NVLS(Rs("RCPNUM"), "")
      DOCNUM = NVLS(Rs("DOCNUM"), "")
      RCVAMT = NVLD(Rs("RCVAMT"), 0)
      VATAMT = NVLD(Rs("VATAMT"), 0)
   ElseIf Ind = 2 Then
      DOCNUM = NVLS(Rs("DOCNUM"), "")
      RCVAMT = NVLD(Rs("RCVAMT"), 0)
      VATAMT = NVLD(Rs("VATAMT"), 0)
   ElseIf Ind = 3 Then
      CUSCOD = NVLS(Rs("CUSCOD"), "")
      RCVAMT = NVLD(Rs("RCVAMT"), 0)
      VATAMT = NVLD(Rs("VATAMT"), 0)
   ElseIf Ind = 4 Then
      SLMCOD = NVLS(Rs("SLMCOD"), "")
      RCVAMT = NVLD(Rs("RCVAMT"), 0)
      VATAMT = NVLD(Rs("VATAMT"), 0)
   ElseIf Ind = 5 Then
      CUSCOD = NVLS(Rs("CUSCOD"), "")
      RCVAMT = NVLD(Rs("RCVAMT"), 0)
   ElseIf Ind = 6 Then
      DOCDAT = InternalDateToDate(NVLS(Rs("DOC_DATE"), ""))
   ElseIf Ind = 7 Then
      DOCDAT = InternalDateToDate(NVLS(Rs("DOC_DATE"), ""))
      PAYDAT = InternalDateToDate(NVLS(Rs("PAY_DATE"), ""))
      RCVAMT = NVLD(Rs("RCVAMT"), 0)
   ElseIf Ind = 8 Then
      RCVAMT = NVLD(Rs("RCVAMT"), 0)
   ElseIf Ind = 9 Then
      DOCNUM = NVLS(Rs("DOCNUM"), "")
      RCVAMT = NVLD(Rs("RCVAMT"), 0)
      RCPNUM = NVLS(Rs("RCPNUM"), "")
   ElseIf Ind = 10 Then
      DOCDAT = InternalDateToDate(NVLS(Rs("DOCDAT"), ""))
      DOCNUM = NVLS(Rs("DOCNUM"), "")
      
      PAYDAT = InternalDateToDate(NVLS(Rs("PAYDAT"), ""))
      RCPNUM = NVLS(Rs("RCPNUM"), "")
      
      AMOUNT = NVLD(Rs("AMOUNT"), 0)
      RCVAMT = NVLD(Rs("RCVAMT"), 0)
      
      CUSCOD = NVLS(Rs("CUSCOD"), "")
      CUSNAM = NVLS(Rs("CUSNAM"), "")
      
      SLMCOD = NVLS(Rs("SLMCOD"), "")
      SLMNAM = NVLS(Rs("SLMNAM"), "")
      RECTYP = NVLS(Rs("RECTYP"), "")
   End If
End Function


