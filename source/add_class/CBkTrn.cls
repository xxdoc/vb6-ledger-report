VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CBkTrn"
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

Public TRNDAT  As Date
Public CHQNUM As String
Public CHQDAT As Date         'วันที่ DUE เช็ค
Public GETDAT As Date            'วันที่รับเช็ค
Public BNKCOD As String
Public CUSCOD As String
Public DEPCOD As String
Public PAYINDAT  As Date
Public REFNUM As String
Public AMOUNT As Double
Public REMARK  As String
Public VOUCHER As String

Public BNKNAM As String
Public BRANCH As String
Public SUPNAM As String
Public CUSNAM As String

Public KEY_ID As Long
Public KEY_LOOKUP As String

Public FROM_CHQDAT As Date
Public TO_CHQDAT As Date
Public FROM_GETDAT As Date
Public TO_GETDAT As Date
Public FROM_TRNDAT As Date

Public TO_PAY_DATE As Date

Public DOCDAT As Date
Public SUPCOD As String
Public TYPDES As String
Public SUPTYP As String
Public CUSTYP As String
Public Credit As Long

Public ACCNUM As String
Public DESCRP As String

Public SLMCOD As String
Public SLMNAM As String

Public FROM_CUSTOMER_CODE As String
Public TO_CUSTOMER_CODE As String
Public FROM_SALE_CODE As String
Public TO_SALE_CODE As String



Private Sub Class_Initialize()
   Set m_Conn = glbDatabaseMngr.DBConnection
   
   FROM_CHQDAT = -1
   TO_CHQDAT = -1
   FROM_GETDAT = -1
   TO_GETDAT = -1
   TO_PAY_DATE = -1
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
Dim SubLen As Long
Dim NewStr As String
Dim TempDate As String
Dim GroupByStr As String
   
   If Ind = 1 Then
      SelectStr1 = "SELECT BKT.*, DTOS(BKT.CHQDAT) CHQ_CHQDAT, DTOS(BKT.GETDAT) CHQ_GETDAT, BKM.BRANCH BNKBRANCH, IST.TYPDES BNKNAM, APM.SUPNAM "
      SelectStr2 = "SELECT COUNT(*)  "
      
      FromStr = "FROM BKTRN BKT "
      FromStr = FromStr & "LEFT OUTER JOIN BKMAS BKM ON (BKT.BNKACC = BKM.BNKACC) "
      FromStr = FromStr & "LEFT OUTER JOIN ISTAB IST ON ((BKT.BNKCOD = IST.TYPCOD) AND (IST.TABTYP = '01')) "
      FromStr = FromStr & "LEFT OUTER JOIN APMAS APM ON (BKT.CUSCOD = APM.SUPCOD) "
   ElseIf Ind = 2 Then
      SelectStr1 = "SELECT BKT.CUSCOD,SUM(BKT.NETAMT) AMOUNT "
      SelectStr2 = "SELECT COUNT(*)  "
      
      FromStr = "FROM BKTRN BKT "
      'FromStr = FromStr & "LEFT OUTER JOIN APRCPIT ART ON (BKT.VOUCHER = ART.RCPNUM) "
      FromStr = FromStr & "LEFT OUTER JOIN APMAS APM ON (BKT.CUSCOD = APM.SUPCOD) "
      
      WhereStr = " WHERE (APM.SUPCOD IS NOT NULL)  "
      'AND APM.SUPCOD = 'ว-0011'
   ElseIf Ind = 3 Then
      SelectStr1 = "SELECT BKT.CHQNUM, BKT.AMOUNT AMOUNT, DTOS(BKT.CHQDAT) CHQDAT, APM.SUPNAM, APM.SUPCOD, BKT.VOUCHER, IST.TYPDES,BKT.REMARK "
      SelectStr2 = "SELECT COUNT(*)  "
      
      FromStr = "FROM BKTRN BKT "
      FromStr = FromStr & "LEFT OUTER JOIN APMAS APM ON (BKT.CUSCOD = APM.SUPCOD) "
      FromStr = FromStr & "LEFT OUTER JOIN ISTAB IST ON ((APM.SUPTYP = IST.TYPCOD) AND (IST.TABTYP = '46')) "
      
      FromStr = FromStr & "LEFT OUTER JOIN APTRN APT2 ON (BKT.VOUCHER = APT2.DOCNUM) " 'เพิ่มเข้ามา 29/05/2555 เพื่อให้สามารถใช้ search ถึงวันที่รับชำระได้
      
      WhereStr = "WHERE (BKT.BKTRNTYP = 'QP') "
   ElseIf Ind = 4 Then
      SelectStr1 = "SELECT BKT.CHQNUM, DTOS(APT.DOCDAT) DOCDAT "
      'SelectStr1 = "SELECT DISTINCT BKT.CHQNUM "
      SelectStr2 = "SELECT COUNT(*)  "
      
      FromStr = "FROM BKTRN BKT "
      FromStr = FromStr & "LEFT OUTER JOIN APMAS APM ON (BKT.CUSCOD = APM.SUPCOD) "
      FromStr = FromStr & "LEFT OUTER JOIN ISTAB IST ON ((APM.SUPTYP = IST.TYPCOD) AND (IST.TABTYP = '46')) "
      
      FromStr = FromStr & "LEFT OUTER JOIN APRCPCQ ARQ ON (ARQ.CHQNUM = BKT.CHQNUM) "
      FromStr = FromStr & "LEFT OUTER JOIN APRCPIT APR ON (APR.RCPNUM = ARQ.RCPNUM) "
      FromStr = FromStr & "LEFT OUTER JOIN APTRN APT ON (APT.DOCNUM = APR.DOCNUM) "
      
      WhereStr = "WHERE (BKT.BKTRNTYP = 'QP') "
   ElseIf Ind = 5 Then
      SelectStr1 = "SELECT BKT.CHQNUM, BKT.AMOUNT AMOUNT, DTOS(BKT.CHQDAT) CHQDAT, APM.SUPNAM, APM.SUPCOD, BKT.VOUCHER, IST.TYPDES,BKT.REMARK,GL.ACCNUM,GL.DESCRP "
      SelectStr2 = "SELECT COUNT(*)  "
      
      FromStr = "FROM BKTRN BKT "
      FromStr = FromStr & "LEFT OUTER JOIN APMAS APM ON (BKT.CUSCOD = APM.SUPCOD) "
      FromStr = FromStr & "LEFT OUTER JOIN ISTAB IST ON ((APM.SUPTYP = IST.TYPCOD) AND (IST.TABTYP = '46')) "
      FromStr = FromStr & "LEFT  JOIN GLJNL GL ON (GL.VOUCHER = BKT.VOUCHER)  "
      
      FromStr = FromStr & "LEFT OUTER JOIN APTRN APT2 ON (BKT.VOUCHER = APT2.DOCNUM) " 'เพิ่มเข้ามา 29/05/2555 เพื่อให้สามารถใช้ search ถึงวันที่รับชำระได้
      
      WhereStr = "WHERE (BKT.BKTRNTYP = 'QP' AND GL.TRNTYP = '0') "
   ElseIf Ind = 6 Then
      SelectStr1 = "SELECT BKT.CHQNUM, BKT.AMOUNT AMOUNT,BKT.BRANCH, DTOS(BKT.PAYINDAT) PAYINDAT, DTOS(BKT.GETDAT) GETDAT, DTOS(BKT.CHQDAT) CHQDAT, ARM.CUSNAM, ARM.CUSCOD, BKT.VOUCHER, IST.TYPDES,BKT.REMARK "
      SelectStr1 = SelectStr1 & ",ART.SLMCOD,OES.SLMNAM "
      SelectStr1 = SelectStr1 & ",IST.SHORTNAM BNKCOD "
      SelectStr2 = "SELECT COUNT(*)  "
      
      FromStr = "FROM BKTRN BKT "
      FromStr = FromStr & "LEFT OUTER JOIN ARMAS ARM ON (BKT.CUSCOD = ARM.CUSCOD) "
      FromStr = FromStr & "LEFT OUTER JOIN ISTAB IST ON ((BKT.BNKCOD = IST.TYPCOD) AND (IST.TABTYP = '01')) "
      
      FromStr = FromStr & "LEFT OUTER JOIN ARTRN ART ON (BKT.VOUCHER = ART.DOCNUM) "
      FromStr = FromStr & "LEFT OUTER JOIN OESLM OES ON (OES.SLMCOD = ART.SLMCOD) "
      
      WhereStr = "WHERE (BKT.BKTRNTYP = 'QR') "
    ElseIf Ind = 7 Then
      SelectStr1 = "SELECT BKT.CHQNUM, DTOS(ART.DOCDAT) DOCDAT "
      SelectStr2 = "SELECT COUNT(*)  "
      
      FromStr = "FROM BKTRN BKT "
      FromStr = FromStr & "LEFT OUTER JOIN ARMAS ARM ON (BKT.CUSCOD = ARM.CUSCOD) "
      FromStr = FromStr & "LEFT OUTER JOIN ISTAB IST ON ((ARM.CUSTYP = IST.TYPCOD) AND (IST.TABTYP = '45')) "
      
      FromStr = FromStr & "LEFT OUTER JOIN ARRCPCQ ARQ ON (ARQ.CHQNUM = BKT.CHQNUM) "
      FromStr = FromStr & "LEFT OUTER JOIN ARRCPIT ARR ON (ARR.RCPNUM = ARQ.RCPNUM) "
      FromStr = FromStr & "LEFT OUTER JOIN ARTRN ART ON (ART.DOCNUM = ARR.DOCNUM) "
      
      WhereStr = "WHERE (BKT.BKTRNTYP = 'QR') "
    ElseIf Ind = 8 Then 'ใช้สำหรับรายงาน 2.3 และ 2.3.2 สรุปอายุเจ้าหนี้ แยกตามช่วงวันที่ ตามกลุ่มเจ้าหนี้
      SelectStr1 = "SELECT SUM(BKT.AMOUNT) AMOUNT, APM.SUPCOD "
      SelectStr2 = "SELECT COUNT(*)  "
      
      FromStr = "FROM BKTRN BKT "
      FromStr = FromStr & "LEFT OUTER JOIN APMAS APM ON (BKT.CUSCOD = APM.SUPCOD) "
      FromStr = FromStr & "LEFT OUTER JOIN ISTAB IST ON ((APM.SUPTYP = IST.TYPCOD) AND (IST.TABTYP = '46')) "
      
      FromStr = FromStr & "LEFT OUTER JOIN APTRN APT2 ON (BKT.VOUCHER = APT2.DOCNUM) " 'เพิ่มเข้ามา 29/05/2555 เพื่อให้สามารถใช้ search ถึงวันที่รับชำระได้
      
      WhereStr = "WHERE (BKT.BKTRNTYP = 'QP') "
   ElseIf Ind = 9 Then
      SelectStr1 = "SELECT BKT.CHQNUM , SUM(BKT.AMOUNT) AMOUNT ,BKT.CUSCOD , ARM.CUSNAM "  '  SUM(ARRCPIT.RCVAMT)
      SelectStr1 = SelectStr1 & ",ART.SLMCOD "
      
      SelectStr2 = "SELECT COUNT(*)  "
      
      FromStr = "FROM BKTRN BKT "
      FromStr = FromStr & "LEFT OUTER JOIN ARMAS ARM ON (BKT.CUSCOD = ARM.CUSCOD) "
     FromStr = FromStr & "LEFT OUTER JOIN ISTAB IST ON ((BKT.BNKCOD = IST.TYPCOD) AND (IST.TABTYP = '01')) "
 
      FromStr = FromStr & "LEFT OUTER JOIN ARTRN ART ON (BKT.VOUCHER = ART.DOCNUM) "
      FromStr = FromStr & "LEFT OUTER JOIN OESLM OES ON (OES.SLMCOD = ART.SLMCOD) "
      
            '  FromStr = FromStr & "LEFT OUTER JOIN ARRCPIT ARPIT ON (ARPIT.RCPNUM = BKT.VOUCHER)
            
       WhereStr = "WHERE (BKT.BKTRNTYP = 'QR') "
 ElseIf Ind = 10 Then
      SelectStr1 = "SELECT SUM(BKT.AMOUNT) AMOUNT, ARM.CUSCOD , ARM.CUSNAM "   ',BKT.CHQNUM
      SelectStr1 = SelectStr1 & ",ART.SLMCOD "  ',OES.SLMNAM
      SelectStr2 = "SELECT COUNT(*)  "
      
      FromStr = "FROM BKTRN BKT "
      FromStr = FromStr & "LEFT OUTER JOIN ARMAS ARM ON (BKT.CUSCOD = ARM.CUSCOD) "
     FromStr = FromStr & "LEFT OUTER JOIN ISTAB IST ON ((BKT.BNKCOD = IST.TYPCOD) AND (IST.TABTYP = '01')) "
      
      FromStr = FromStr & "LEFT OUTER JOIN ARTRN ART ON (BKT.VOUCHER = ART.DOCNUM) "
      FromStr = FromStr & "LEFT OUTER JOIN OESLM OES ON (OES.SLMCOD = ART.SLMCOD) "
      
      WhereStr = "WHERE (BKT.BKTRNTYP = 'QR') "
 ElseIf Ind = 11 Then
      SelectStr1 = "SELECT BKT.* , ARM.CUSCOD , ARM.CUSNAM "
      SelectStr1 = SelectStr1 & ",ART.SLMCOD "  ',OES.SLMNAM
      SelectStr2 = "SELECT COUNT(*)  "
      
      FromStr = "FROM BKTRN BKT "
      FromStr = FromStr & "LEFT OUTER JOIN ARMAS ARM ON (BKT.CUSCOD = ARM.CUSCOD) "
     FromStr = FromStr & "LEFT OUTER JOIN ISTAB IST ON ((BKT.BNKCOD = IST.TYPCOD) AND (IST.TABTYP = '01')) "
      
      FromStr = FromStr & "LEFT OUTER JOIN ARTRN ART ON (BKT.VOUCHER = ART.DOCNUM) "
      FromStr = FromStr & "LEFT OUTER JOIN OESLM OES ON (OES.SLMCOD = ART.SLMCOD) "
      
      WhereStr = "WHERE (BKT.BKTRNTYP = 'QR') "
   ElseIf Ind = 12 Then
      SelectStr1 = "SELECT BKT.CHQNUM, BKT.AMOUNT AMOUNT,BKT.BRANCH, DTOS(BKT.PAYINDAT) PAYINDAT, DTOS(BKT.GETDAT) GETDAT, DTOS(BKT.CHQDAT) CHQDAT, ARM.CUSNAM, ARM.CUSCOD, BKT.VOUCHER, IST.TYPDES,BKT.REMARK "
      SelectStr1 = SelectStr1 & ",ART.SLMCOD,OES.SLMNAM "
      SelectStr1 = SelectStr1 & ",IST.SHORTNAM BNKCOD "
      SelectStr2 = "SELECT COUNT(*)  "
      
      FromStr = "FROM BKTRN BKT "
      FromStr = FromStr & "LEFT OUTER JOIN ARMAS ARM ON (BKT.CUSCOD = ARM.CUSCOD) "
      FromStr = FromStr & "LEFT OUTER JOIN ISTAB IST ON ((BKT.BNKCOD = IST.TYPCOD) AND (IST.TABTYP = '01')) "
      
      FromStr = FromStr & "LEFT OUTER JOIN ARTRN ART ON (BKT.VOUCHER = ART.DOCNUM) "
      FromStr = FromStr & "LEFT OUTER JOIN OESLM OES ON (OES.SLMCOD = ART.SLMCOD) "
      
      WhereStr = "WHERE (BKT.BKTRNTYP = 'QR') "
      
  End If
   '
   If Trim(VOUCHER) <> "" Then
      If WildCard(VOUCHER, SubLen, NewStr) Then
         If WhereStr = "" Then
            WhereStr = "WHERE (VOUCHER LIKE '" & ChangeQuote(Trim(VOUCHER)) & "')"
         Else
            WhereStr = WhereStr & "AND (VOUCHER LIKE '" & ChangeQuote(Trim(VOUCHER)) & "')"
         End If
      Else
         If WhereStr = "" Then
            WhereStr = "WHERE (VOUCHER = '" & ChangeQuote(Trim(VOUCHER)) & "')"
         Else
            WhereStr = WhereStr & "AND (VOUCHER = '" & ChangeQuote(Trim(VOUCHER)) & "')"
         End If
      End If
   End If
                                       
   If Trim(CHQNUM) <> "" Then
      If WildCard(CHQNUM, SubLen, NewStr) Then
         If WhereStr = "" Then
            WhereStr = "WHERE (CHQNUM LIKE '" & ChangeQuote(Trim(CHQNUM)) & "')"
         Else
            WhereStr = WhereStr & "AND (CHQNUM LIKE '" & ChangeQuote(Trim(CHQNUM)) & "')"
         End If
      Else
         If WhereStr = "" Then
            WhereStr = "WHERE (CHQNUM = '" & ChangeQuote(Trim(CHQNUM)) & "')"
         Else
            WhereStr = WhereStr & "AND (CHQNUM = '" & ChangeQuote(Trim(CHQNUM)) & "')"
         End If
      End If
   End If
                                       
   If FROM_GETDAT > -1 Then
      TempDate = DateToStringIntLow(Trim(FROM_GETDAT))
      If WhereStr = "" Then
         WhereStr = "WHERE (DTOS(GETDAT) >= '" & ChangeQuote(TempDate) & "')"
      Else
         WhereStr = WhereStr & "AND (DTOS(GETDAT) >= '" & ChangeQuote(TempDate) & "')"
      End If
   End If
   
   If TO_GETDAT > -1 Then
      TempDate = DateToStringIntHi(Trim(TO_GETDAT))
      If WhereStr = "" Then
         WhereStr = "WHERE (DTOS(GETDAT) <= '" & ChangeQuote(TempDate) & "')"
      Else
         WhereStr = WhereStr & "AND (DTOS(GETDAT) <= '" & ChangeQuote(TempDate) & "')"
      End If
   End If
                                       
   If FROM_CHQDAT > -1 Then
      TempDate = DateToStringIntLow(Trim(FROM_CHQDAT))
      If WhereStr = "" Then
         WhereStr = "WHERE (DTOS(CHQDAT) >= '" & ChangeQuote(TempDate) & "')"
      Else
         WhereStr = WhereStr & "AND (DTOS(CHQDAT) >= '" & ChangeQuote(TempDate) & "')"
      End If
   End If
   
   If TO_CHQDAT > -1 Then
      TempDate = DateToStringIntHi(Trim(TO_CHQDAT))
      If WhereStr = "" Then
         WhereStr = "WHERE (DTOS(CHQDAT) <= '" & ChangeQuote(TempDate) & "')"
      Else
         WhereStr = WhereStr & "AND (DTOS(CHQDAT) <= '" & ChangeQuote(TempDate) & "')"
      End If
   End If
   
      If FROM_TRNDAT > -1 Then
      TempDate = DateToStringIntLow(Trim(FROM_TRNDAT))
      If WhereStr = "" Then
         WhereStr = "WHERE (DTOS(TRNDAT) >= '" & ChangeQuote(TempDate) & "')"
      Else
         WhereStr = WhereStr & "AND (DTOS(TRNDAT) >= '" & ChangeQuote(TempDate) & "')"
      End If
   End If
   
   If Trim(SUPCOD) <> "" Then
      If WildCard(SUPCOD, SubLen, NewStr) Then
         If WhereStr = "" Then
            WhereStr = "WHERE (APM.SUPCOD LIKE '" & ChangeQuote(Trim(SUPCOD)) & "')"
         Else
            WhereStr = WhereStr & "AND (APM.SUPCOD LIKE '" & ChangeQuote(Trim(SUPCOD)) & "')"
         End If
      Else
         If WhereStr = "" Then
            WhereStr = "WHERE (APM.SUPCOD = '" & ChangeQuote(Trim(SUPCOD)) & "')"
         Else
            WhereStr = WhereStr & "AND (APM.SUPCOD = '" & ChangeQuote(Trim(SUPCOD)) & "')"
         End If
      End If
   End If
                  
   
   If Trim(CUSCOD) <> "" Then
      If WildCard(CUSCOD, SubLen, NewStr) Then
         If WhereStr = "" Then
            WhereStr = "WHERE (ARM.CUSCOD LIKE '" & ChangeQuote(Trim(CUSCOD)) & "')"
         Else
            WhereStr = WhereStr & "AND (ARM.CUSCOD LIKE '" & ChangeQuote(Trim(CUSCOD)) & "')"
         End If
      Else
         If WhereStr = "" Then
            WhereStr = "WHERE (ARM.CUSCOD = '" & ChangeQuote(Trim(CUSCOD)) & "')"
         Else
            WhereStr = WhereStr & "AND (ARM.CUSCOD = '" & ChangeQuote(Trim(CUSCOD)) & "')"
         End If
      End If
   End If
   
   If Trim(FROM_CUSTOMER_CODE) <> "" Then
      If WhereStr = "" Then
         WhereStr = "WHERE (ARM.CUSCOD >= '" & ChangeQuote(Trim(FROM_CUSTOMER_CODE)) & "')"
      Else
         WhereStr = WhereStr & "AND (ARM.CUSCOD >= '" & ChangeQuote(Trim(FROM_CUSTOMER_CODE)) & "')"
      End If
   End If
   
   If Trim(TO_CUSTOMER_CODE) <> "" Then
      If WhereStr = "" Then
         WhereStr = "WHERE (ARM.CUSCOD <= '" & ChangeQuote(Trim(TO_CUSTOMER_CODE)) & "')"
      Else
         WhereStr = WhereStr & "AND (ARM.CUSCOD <= '" & ChangeQuote(Trim(TO_CUSTOMER_CODE)) & "')"
      End If
   End If
   
   If Len(SUPTYP) > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (APM.SUPTYP = '" & Trim(ChangeQuote(SUPTYP)) & "') "
      Else
         WhereStr = WhereStr & "AND (APM.SUPTYP = '" & Trim(ChangeQuote(SUPTYP)) & "') "
      End If
   End If
   
   If Len(CUSTYP) > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (ARM.CUSTYP = '" & Trim(ChangeQuote(CUSTYP)) & "') "
      Else
         WhereStr = WhereStr & "AND (ARM.CUSTYP = '" & Trim(ChangeQuote(CUSTYP)) & "') "
      End If
   End If
   
   If Len(ACCNUM) > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (GL.ACCNUM = '" & Trim(ChangeQuote(ACCNUM)) & "') "
      Else
         WhereStr = WhereStr & "AND (GL.ACCNUM = '" & Trim(ChangeQuote(ACCNUM)) & "') "
      End If
   End If
   
   If Trim(REMARK) <> "" Then
      If WildCard(REMARK, SubLen, NewStr) Then
         If WhereStr = "" Then
            WhereStr = "WHERE (BKT.REMARK LIKE '" & ChangeQuote(Trim(REMARK)) & "')"
         Else
            WhereStr = WhereStr & "AND (BKT.REMARK LIKE '" & ChangeQuote(Trim(REMARK)) & "')"
         End If
      Else
         If WhereStr = "" Then
            WhereStr = "WHERE (BKT.REMARK = '" & ChangeQuote(Trim(REMARK)) & "')"
         Else
            WhereStr = WhereStr & "AND (BKT.REMARK = '" & ChangeQuote(Trim(REMARK)) & "')"
         End If
      End If
   End If
      
   If Trim(FROM_SALE_CODE) <> "" Then
      If WhereStr = "" Then
         WhereStr = "WHERE (OES.SLMCOD >= '" & ChangeQuote(Trim(FROM_SALE_CODE)) & "')"
      Else
         WhereStr = WhereStr & "AND (OES.SLMCOD >= '" & ChangeQuote(Trim(FROM_SALE_CODE)) & "')"
      End If
   End If
   
   If Trim(TO_SALE_CODE) <> "" Then
      If WhereStr = "" Then
         WhereStr = "WHERE (OES.SLMCOD <= '" & ChangeQuote(Trim(TO_SALE_CODE)) & "')"
      Else
         WhereStr = WhereStr & "AND (OES.SLMCOD <= '" & ChangeQuote(Trim(TO_SALE_CODE)) & "')"
      End If
   End If
   
   If TO_PAY_DATE > -1 Then
      TempDate = DateToStringIntHi(Trim(TO_PAY_DATE))
      If WhereStr = "" Then
         WhereStr = "WHERE (DTOS(APT2.DOCDAT) <= '" & ChangeQuote(TempDate) & "')"
      Else
         WhereStr = WhereStr & "AND (DTOS(APT2.DOCDAT) <= '" & ChangeQuote(TempDate) & "')"
      End If
   End If
   
   GroupByStr = ""
   If Ind = 2 Then
      GroupByStr = " GROUP BY BKT.CUSCOD "
   ElseIf Ind = 8 Then
      GroupByStr = " GROUP BY APM.SUPCOD "
   ElseIf Ind = 9 Then
      GroupByStr = " GROUP BY BKT.CHQNUM , BKT.CUSCOD, ARM.CUSNAM, ART.SLMCOD  "                    ', TRNDAT
   ElseIf Ind = 10 Then
       GroupByStr = " GROUP BY ARM.CUSCOD , ARM.CUSNAM , ART.SLMCOD "   ', OES.SLMNAM
   End If
   
   OrderByStr = ""
   OrderTypeStr = ""
   
   If Ind = 1 Then
      If OrderBy = 1 Then
         OrderByStr = "ORDER BY CHQNUM "
      ElseIf OrderBy = 2 Then
         OrderByStr = "ORDER BY CHQNUM "
      Else
         OrderByStr = "ORDER BY CHQNUM "
      End If
   ElseIf Ind = 2 Then
      OrderByStr = "ORDER BY BKT.CUSCOD "
   ElseIf Ind = 3 Then
      If OrderBy = 1 Then
         OrderByStr = "ORDER BY BKT.CHQDAT ASC,APM.SUPNAM  "
      Else
         OrderByStr = "ORDER BY APM.SUPNAM ASC, BKT.CHQDAT "
      End If
   ElseIf Ind = 4 Then
      OrderByStr = "ORDER BY BKT.CHQNUM ASC,APT.DOCDAT "
   ElseIf Ind = 5 Then
      OrderByStr = "ORDER BY GL.ACCNUM ASC,BKT.CHQDAT "
   ElseIf Ind = 6 Then
      If OrderBy = 1 Then
         OrderByStr = "ORDER BY BKT.GETDAT ASC,ARM.CUSNAM  "
      ElseIf OrderBy = 2 Then
         OrderByStr = "ORDER BY ARM.CUSCOD ASC, BKT.GETDAT ASC,OES.SLMCOD "       ' เรียงตามลูกค้า
      ElseIf OrderBy = 3 Then
         OrderByStr = "ORDER BY OES.SLMCOD ASC, BKT.GETDAT ASC,ARM.CUSCOD "    ' เรียงตาม Sale
      ElseIf OrderBy = 4 Then
         OrderByStr = "ORDER BY OES.SLMCOD ASC, ARM.CUSCOD ASC , BKT.GETDAT "    ' เรียงตาม Sale , ลูกค้า , วันที่
      Else
         OrderByStr = "ORDER BY OES.SLMCOD ASC, BKT.GETDAT ASC,ARM.CUSCOD "
      End If
    ElseIf Ind = 7 Then
      OrderByStr = "ORDER BY BKT.CHQNUM ASC,ART.DOCDAT "
   ElseIf Ind = 8 Then
      OrderByStr = "ORDER BY APM.SUPCOD "
    ElseIf Ind = 9 Then
      OrderByStr = "ORDER BY BKT.CUSCOD "
   ElseIf Ind = 10 Then
         OrderByStr = "ORDER BY ARM.CUSCOD "    ' เรียงตาม Sale , ลูกค้า , วันที่
   ElseIf Ind = 11 Then
        OrderByStr = " ORDER BY BKT.CUSCOD "
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
   
   Call glbDatabaseMngr.GetRs(SQL1, SQL2, True, ItemCount, Rs, glbErrorLog)
End Function

Public Function PopulateFromRS(Ind As Long, Rs As ADODB.Recordset)
   If Ind = 1 Then
      TRNDAT = InternalDateToDate(NVLS(Rs("TRNDAT"), ""))
      CHQNUM = NVLS(Rs("CHQNUM"), "")
      CHQDAT = InternalDateToDate(NVLS(Rs("CHQ_CHQDAT"), ""))
      GETDAT = InternalDateToDate(NVLS(Rs("CHQ_GETDAT"), ""))
      BNKCOD = NVLS(Rs("BNKCOD"), "")
      CUSCOD = NVLS(Rs("CUSCOD"), "")
      DEPCOD = NVLS(Rs("DEPCOD"), "")
      PAYINDAT = InternalDateToDate(NVLS(Rs("PAYINDAT"), ""))
      REFNUM = NVLS(Rs("REFNUM"), "")
      AMOUNT = NVLD(Rs("AMOUNT"), 0)
      REMARK = NVLS(Rs("REMARK"), "")
      VOUCHER = NVLS(Rs("VOUCHER"), "")
      
      SUPNAM = NVLS(Rs("SUPNAM"), "")
      BNKNAM = NVLS(Rs("BNKNAM"), "")
      BRANCH = NVLS(Rs("BRANCH"), "")
   ElseIf Ind = 2 Then
      CUSCOD = NVLS(Rs("CUSCOD"), "")
      AMOUNT = NVLD(Rs("AMOUNT"), 0)
   ElseIf Ind = 3 Then
      CHQNUM = NVLS(Rs("CHQNUM"), "")
      AMOUNT = NVLD(Rs("AMOUNT"), 0)
      
      CHQDAT = InternalDateToDate(NVLS(Rs("CHQDAT"), ""))
      SUPCOD = NVLS(Rs("SUPCOD"), "")
      SUPNAM = NVLS(Rs("SUPNAM"), "")
      TYPDES = NVLS(Rs("TYPDES"), "")
      VOUCHER = NVLS(Rs("VOUCHER"), "")
      REMARK = NVLS(Rs("REMARK"), "")
   ElseIf Ind = 4 Then
      CHQNUM = NVLS(Rs("CHQNUM"), "")
      DOCDAT = InternalDateToDate(NVLS(Rs("DOCDAT"), ""))
   ElseIf Ind = 5 Then
      CHQNUM = NVLS(Rs("CHQNUM"), "")
      AMOUNT = NVLD(Rs("AMOUNT"), 0)
      
      CHQDAT = InternalDateToDate(NVLS(Rs("CHQDAT"), ""))
      SUPCOD = NVLS(Rs("SUPCOD"), "")
      SUPNAM = NVLS(Rs("SUPNAM"), "")
      TYPDES = NVLS(Rs("TYPDES"), "")
      VOUCHER = NVLS(Rs("VOUCHER"), "")
      REMARK = NVLS(Rs("REMARK"), "")
      DESCRP = NVLS(Rs("DESCRP"), "")
      
      ACCNUM = NVLS(Rs("ACCNUM"), "")
   ElseIf Ind = 6 Then
      CHQNUM = NVLS(Rs("CHQNUM"), "")
      AMOUNT = NVLD(Rs("AMOUNT"), 0)
      
      CHQDAT = InternalDateToDate(NVLS(Rs("CHQDAT"), ""))
      GETDAT = InternalDateToDate(NVLS(Rs("GETDAT"), ""))
      PAYINDAT = InternalDateToDate(NVLS(Rs("PAYINDAT"), ""))
      
      CUSCOD = NVLS(Rs("CUSCOD"), "")
      CUSNAM = NVLS(Rs("CUSNAM"), "")
      BNKCOD = NVLS(Rs("BNKCOD"), "")
      BNKNAM = NVLS(Rs("TYPDES"), "")
      BRANCH = NVLS(Rs("BRANCH"), "")
      VOUCHER = NVLS(Rs("VOUCHER"), "")
      REMARK = NVLS(Rs("REMARK"), "")
      
      SLMCOD = NVLS(Rs("SLMCOD"), "")
      SLMNAM = NVLS(Rs("SLMNAM"), "")
    ElseIf Ind = 7 Then
      CHQNUM = NVLS(Rs("CHQNUM"), "")
'       If (CHQNUM = "IV0040532") Then
'            'debug.print
'        End If
      DOCDAT = InternalDateToDate(NVLS(Rs("DOCDAT"), ""))
   ElseIf Ind = 8 Then
      AMOUNT = NVLD(Rs("AMOUNT"), 0)
      SUPCOD = NVLS(Rs("SUPCOD"), "")
   ElseIf Ind = 9 Then
        AMOUNT = NVLD(Rs("AMOUNT"), 0)
        CHQNUM = NVLS(Rs("CHQNUM"), "")
        CUSCOD = NVLS(Rs("CUSCOD"), "")
              SLMCOD = NVLS(Rs("SLMCOD"), "")
  CUSNAM = NVLS(Rs("CUSNAM"), "")
   ElseIf Ind = 10 Then
      AMOUNT = NVLD(Rs("AMOUNT"), 0)
      CUSCOD = NVLS(Rs("CUSCOD"), "")
      CUSNAM = NVLS(Rs("CUSNAM"), "")
      SLMCOD = NVLS(Rs("SLMCOD"), "")
     ElseIf Ind = 11 Then
      AMOUNT = NVLD(Rs("AMOUNT"), 0)
      CUSCOD = NVLS(Rs("CUSCOD"), "")
    CUSNAM = NVLS(Rs("CUSNAM"), "")
      SLMCOD = NVLS(Rs("SLMCOD"), "")
     CHQNUM = NVLS(Rs("CHQNUM"), "")
   End If
End Function

Public Sub DeleteAllData()
Dim SQL1 As String

   SQL1 = "DELETE FROM BKTRN "
   m_Conn.Execute (SQL1)
End Sub




