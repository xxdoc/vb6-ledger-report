VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CStcrd"
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
Public KEY As String
' End of generic data member
Public addData As String
Public RECTYP As String
Public DOCNUM As String
Public DOCDAT  As Date


Public SUPCOD  As String
Public CUSCOD  As String
Public PEOPLE  As String

Public STKCOD  As String
Public STKDES  As String
Public AMOUNT  As Double
Public TRNQTY  As Double      'นน
Public NETVAL  As Double      'ราคาสุทธิ
Public DISAMT As Double        'ส่วนลด
Public UNITPR As Double       'ราคาเฉลี่ย/หน่วย

Public LOCCOD As String       'คลัง NO

Public FROM_DOC_DATE As Date
Public TO_DOC_DATE As Date
Public FROM_DUE_DATE As Date
Public TO_DUE_DATE As Date
Public FROM_CMPL_DATE As Date
Public TO_CMPL_DATE As Date
Public FROM_SUPPLIER_CODE As String
Public TO_SUPPLIER_CODE As String
Public FROM_CUSTOMER_CODE As String
Public TO_CUSTOMER_CODE  As String
Public FROM_STOCK_CODE As String
Public TO_STOCK_CODE As String

Public POSOPR As String          'ประเภทเอกสารใน ตารางนี้   ------------->STCRD


Public RENUM As String
Public FROM_SLM_DATE As Date
Public TO_SLM_DATE As Date
Public CMPLDAT As Date         'วันที่รับชำระ
Public SLMNAM As String
Public AREACOD As String
Public SEQNUM As String
Public TYPDES As String
Public AREANAM As String

Public RO_TYPE As String

Public SUPNAM As String
Public SUPTYP As String
Public SUPTYPSET As String
Public SUPTYPNAM As String

Public CUSNAM As String

Public CHQNUM  As String
Public CHQDAT As Date            'วันที่ DUE เช็ค
Public GETDAT As Date            'วันที่ มารับเช็ค
Public PAYAMT As Double

Public FROM_SALE_CODE As String
Public TO_SALE_CODE As String
Public SLMCOD As String

Public KEY_ID As Long
Public KEY_LOOKUP As String
Public RecTypeSet As String

Public TEMP_ADDR1 As String
Public TEMP_ADDR2 As String
Public ADDR01 As String
Public ADDR02 As String
Public TELNUM As String

Public eachSLM As Collection
Public temp_eachSLM As CCondiIncenSum
Public db2 As Boolean

Private Sub Class_Initialize()
   Set m_Conn = glbDatabaseMngr.DBConnection
   
   FROM_DOC_DATE = -1
   TO_DOC_DATE = -1
   FROM_DUE_DATE = -1
   TO_DUE_DATE = -1
   FROM_CMPL_DATE = -1
   TO_CMPL_DATE = -1
   FROM_SLM_DATE = -1
   TO_SLM_DATE = -1
   
End Sub

'In this class we will not handle any error or any transaction
Public Function QueryData(Ind As Long, Rs As ADODB.Recordset, ItemCount As Long, Optional db_num As Long = 1)
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
Dim i As Long
   
   If Ind = 1 Then
      SelectStr1 = "SELECT STC.* "
      SelectStr1 = SelectStr1 & ",APT.AMOUNT,APT.SUPCOD,APT.DOCNUM,DTOS(STC.DOCDAT) DOCDAT "     'รหัสผู้ขาย หมายเลข และวันที่เอกสาร
      SelectStr1 = SelectStr1 & ",APM.SUPNAM,APM.SUPTYP "     'ชื่อ และ รหัสประเภท ผู้ขาย
      SelectStr1 = SelectStr1 & ",IST.TYPDES "       'ประเภทผู้ขาย
      SelectStr1 = SelectStr1 & ",DTOS(BKT.GETDAT) GETDAT,BKT.CHQNUM,DTOS(BKT.CHQDAT) CHQDAT "       'รายละเอียดเช็ค
      SelectStr1 = SelectStr1 & ",ART.PAYAMT,APM.ADDR01,APM.ADDR02,APM.TELNUM "
      
      SelectStr2 = "SELECT COUNT(*)  "
      
      FromStr = " FROM STCRD STC "
      FromStr = FromStr & "LEFT OUTER JOIN APTRN APT ON (APT.DOCNUM = STC.DOCNUM) "
      FromStr = FromStr & "LEFT OUTER JOIN APRCPIT ART ON (ART.DOCNUM = APT.DOCNUM) "
      FromStr = FromStr & "LEFT OUTER JOIN BKTRN BKT ON (BKT.VOUCHER = ART.RCPNUM) "
      FromStr = FromStr & "LEFT OUTER JOIN APMAS APM ON (APT.SUPCOD = APM.SUPCOD) "
      FromStr = FromStr & "LEFT OUTER JOIN ISTAB IST ON (IST.TYPCOD = APM.SUPTYP AND IST.TABTYP = '46' )  "  ' 46 เป็นประเภท MASTER ผู้ขาย
      
      WhereStr = "WHERE (APT.SUPCOD IS NOT NULL) "
   ElseIf Ind = 2 Then
      SelectStr1 = "SELECT STC.* "
       SelectStr2 = "SELECT COUNT(*)  "
      
      FromStr = " FROM STCRD STC "
   ElseIf Ind = 3 Then  'LoadCustomerStcrd
      SelectStr1 = "SELECT STC.* "
      SelectStr1 = SelectStr1 & ",ART.CUSCOD,DTOS(STC.DOCDAT) DOCDAT,ART.RECTYP,ART.DOCNUM "     'รหัสลูกค้า วันที่เอกสาร
      
      SelectStr2 = "SELECT COUNT(*)  "
      
      FromStr = " FROM STCRD STC "
      FromStr = FromStr & "LEFT OUTER JOIN ARTRN ART ON (ART.DOCNUM = STC.DOCNUM) "
      FromStr = FromStr & "LEFT OUTER JOIN ARMAS ARM ON (ART.CUSCOD = ARM.CUSCOD) "
      FromStr = FromStr & "LEFT OUTER JOIN ISTAB IST ON (IST.TYPCOD = ARM.CUSTYP AND IST.TABTYP = '45' )  "  ' 45 เป็นประเภท MASTER ลูกค้า
      
      WhereStr = "WHERE (ART.CUSCOD IS NOT NULL) "
   ElseIf Ind = 4 Then  'LoadDistinctCustomerStcrd
      SelectStr1 = "SELECT DISTINCT STC.STKCOD,STC.STKDES " 'รหัสสินค้า ชื่อ สินค้า
      SelectStr1 = SelectStr1 & ",ART.CUSCOD "     'รหัสลูกค้า
      SelectStr1 = SelectStr1 & ",ARM.CUSNAM "     'ชื่อ ลูกค้า
      
      SelectStr2 = "SELECT COUNT(*)  "
      
      FromStr = " FROM STCRD STC "
      FromStr = FromStr & "LEFT OUTER JOIN ARTRN ART ON (ART.DOCNUM = STC.DOCNUM) "
      
      FromStr = FromStr & "LEFT OUTER JOIN ARMAS ARM ON (ART.CUSCOD = ARM.CUSCOD) "
      FromStr = FromStr & "LEFT OUTER JOIN ISTAB IST ON (IST.TYPCOD = ARM.CUSTYP AND IST.TABTYP = '45' )  "  ' 45 เป็นประเภท MASTER ลูกค้า
      
      WhereStr = "WHERE (ART.CUSCOD IS NOT NULL) "
   ElseIf Ind = 5 Then  'LoadDistinctStcrdCustomer
      SelectStr1 = "SELECT DISTINCT STC.STKCOD,STC.STKDES " 'รหัสสินค้า ชื่อ สินค้า
      SelectStr1 = SelectStr1 & ",ART.CUSCOD "     'รหัสลูกค้า
      SelectStr1 = SelectStr1 & ",ARM.CUSNAM "     'ชื่อ ลูกค้า
      
      SelectStr2 = "SELECT COUNT(*)  "
      
      FromStr = " FROM STCRD STC "
      FromStr = FromStr & "LEFT OUTER JOIN ARTRN ART ON (ART.DOCNUM = STC.DOCNUM) "
      
      FromStr = FromStr & "LEFT OUTER JOIN ARMAS ARM ON (ART.CUSCOD = ARM.CUSCOD) "
      FromStr = FromStr & "LEFT OUTER JOIN ISTAB IST ON (IST.TYPCOD = ARM.CUSTYP AND IST.TABTYP = '45' )  "  ' 45 เป็นประเภท MASTER ลูกค้า
      
      WhereStr = "WHERE (ART.CUSCOD IS NOT NULL) "
   ElseIf Ind = 6 Then  'LoadSaleCustomerStcrd
      SelectStr1 = "SELECT STC.* "
      SelectStr1 = SelectStr1 & ",ART.CUSCOD,DTOS(STC.DOCDAT) DOCDAT,ART.RECTYP,ART.DOCNUM "     'รหัสลูกค้า วันที่เอกสาร
      
      SelectStr2 = "SELECT COUNT(*)  "
      
      FromStr = " FROM STCRD STC "
      FromStr = FromStr & "LEFT OUTER JOIN ARTRN ART ON (ART.DOCNUM = STC.DOCNUM) "
      FromStr = FromStr & "LEFT OUTER JOIN ARMAS ARM ON (ART.CUSCOD = ARM.CUSCOD) "
      FromStr = FromStr & "LEFT OUTER JOIN ISTAB IST ON (IST.TYPCOD = ARM.CUSTYP AND IST.TABTYP = '45' )  "  ' 45 เป็นประเภท MASTER ลูกค้า
      
      WhereStr = "WHERE (ART.CUSCOD IS NOT NULL) "
   
   ElseIf Ind = 7 Then  'LoadDistinctSaleCustomerStcrd
      SelectStr1 = "SELECT DISTINCT STC.STKCOD,STC.STKDES " 'รหัสสินค้า ชื่อ สินค้า
      SelectStr1 = SelectStr1 & ",ART.CUSCOD "     'รหัสลูกค้า
      SelectStr1 = SelectStr1 & ",ARM.CUSNAM "     'ชื่อ ลูกค้า
      SelectStr1 = SelectStr1 & ",STC.SLMCOD,OES.SLMNAM "
      
      SelectStr2 = "SELECT COUNT(*)  "
      
      FromStr = " FROM STCRD STC "
      FromStr = FromStr & "LEFT OUTER JOIN ARTRN ART ON (ART.DOCNUM = STC.DOCNUM) "
      
      FromStr = FromStr & "LEFT OUTER JOIN ARMAS ARM ON (ART.CUSCOD = ARM.CUSCOD) "
      FromStr = FromStr & "LEFT OUTER JOIN ISTAB IST ON (IST.TYPCOD = ARM.CUSTYP AND IST.TABTYP = '45' )  "  ' 45 เป็นประเภท MASTER ลูกค้า
      
      FromStr = FromStr & " LEFT OUTER JOIN OESLM OES ON (OES.SLMCOD = ART.SLMCOD) "
      
      WhereStr = " WHERE (STC.SLMCOD IS NOT NULL AND ART.CUSCOD IS NOT NULL) "
  ElseIf Ind = 8 Then
      SelectStr1 = "SELECT DISTINCT STC.STKCOD,STC.STKDES "
       SelectStr2 = "SELECT COUNT(*)  "
      
      FromStr = " FROM STCRD STC "
      
      WhereStr = " WHERE (STC.POSOPR ='9') "
 ElseIf Ind = 9 Then  'LoadDistinctSaleCustomerStcrd
      SelectStr1 = "SELECT STC.STKCOD ,STC.DOCNUM ,STC.SEQNUM ,DTOS(STC.DOCDAT) DOCDAT ,STC.PEOPLE ,STC.SLMCOD ,STC.NETVAL ,STC.STKDES ,STC.TRNQTY " 'รหัสสินค้า ชื่อ สินค้า
     ' SelectStr1 = SelectStr1 & ",DTOS(ART.CMPLDAT) CMPLDAT, ART.AREACOD "     'วันที่ขึ้นเงิน
      SelectStr1 = SelectStr1 & ",DTOS(ART.CMPLDAT) CMPLDAT, ART.DOCNUM RENUM "     'วันที่ขึ้นเงิน
      SelectStr1 = SelectStr1 & ",ART2.AREACOD "
      SelectStr1 = SelectStr1 & ",ARM.CUSNAM "     'ชื่อ ลูกค้า
     SelectStr1 = SelectStr1 & ",OES.SLMNAM "
      SelectStr1 = SelectStr1 & ",IST.TYPDES "
      
      SelectStr2 = "SELECT COUNT(*)  "
      
      FromStr = " FROM STCRD STC "
      FromStr = FromStr & "LEFT OUTER JOIN ARRCPIT ARP ON (ARP.DOCNUM = STC.DOCNUM) "
      
      FromStr = FromStr & "LEFT OUTER JOIN ARTRN ART ON (ART.DOCNUM = ARP.RCPNUM) "
      FromStr = FromStr & "LEFT OUTER JOIN ARTRN ART2 ON (ART2.DOCNUM = STC.DOCNUM) "   '
      
      FromStr = FromStr & "LEFT OUTER JOIN ARMAS ARM ON (STC.PEOPLE= ARM.CUSCOD ) "
      FromStr = FromStr & "LEFT OUTER JOIN ISTAB IST ON (IST.TYPCOD = ART2.AREACOD AND IST.TABTYP = '40' )  "  ' 40 เป็น เขตการขายของ SLM
      
      FromStr = FromStr & " LEFT OUTER JOIN OESLM OES ON (OES.SLMCOD = STC.SLMCOD) "
      
     WhereStr = " WHERE ( STC.POSOPR = '9' ) "
   ' WhereStr = " WHERE (ART.CMPLDAT IS NOT NULL) AND (( STC.POSOPR = '0' ) OR ( STC.POSOPR = '9' )) "

     ElseIf Ind = 10 Then
      SelectStr1 = "SELECT SUM(STC.NETVAL) NETVAL,STC.DOCNUM , DTOS(STC.DOCDAT) DOCDAT "
      SelectStr2 = "SELECT COUNT(*)  "
      
      FromStr = " FROM STCRD STC "
      
      WhereStr = " WHERE (STC.POSOPR ='9') "
      
     ElseIf Ind = 11 Then
     
      Set eachSLM = New Collection
      Call LoadIncenForSum(eachSLM, FROM_SLM_DATE, FROM_SLM_DATE)  ' ใส่ YEAR_ID
      
      SelectStr1 = "SELECT SUM(STC.TRNQTY) AS TRNQTY , STC.STKCOD , STC.STKDES "
      SelectStr2 = "SELECT COUNT(*)  "
      
      FromStr = " FROM STCRD STC "
      
      WhereStr = " WHERE (STC.POSOPR ='9') "
      WhereStr = WhereStr & " AND (STC.UNITPR > 0 ) "
      
      i = 0
      For Each temp_eachSLM In eachSLM
        If temp_eachSLM.Flag <> "N" Then
             i = i + 1
             If i = 1 Then
                WhereStr = WhereStr & "AND ( (STC.SLMCOD = '" & temp_eachSLM.SLMCOD & "' ) "
             Else
                WhereStr = WhereStr & " OR (STC.SLMCOD = '" & temp_eachSLM.SLMCOD & "' ) "
              End If
        End If
      Next temp_eachSLM
      
            WhereStr = WhereStr & ")"

    ElseIf Ind = 12 Then
    '  SelectStr1 = "SELECT DISTINCT STC.PEOPLE"
      SelectStr1 = "SELECT SUM(NETVAL) AS NETVAL, STC.PEOPLE AS CUSCOD "       ', DTOS(STC.DOCDAT) DOCDAT
       SelectStr2 = "SELECT COUNT(*)  "
      
      FromStr = " FROM STCRD STC "
      
      WhereStr = " WHERE (STC.POSOPR ='9') "
    ElseIf Ind = 13 Then  'LoadDistinctSaleCustomerStcrd
      SelectStr1 = "SELECT STC.STKCOD ,STC.DOCNUM ,STC.SEQNUM ,DTOS(STC.DOCDAT) DOCDAT ,STC.PEOPLE ,STC.SLMCOD ,STC.NETVAL ,STC.STKDES ,STC.TRNQTY ,STC.UNITPR " 'รหัสสินค้า ชื่อ สินค้า
      SelectStr1 = SelectStr1 & ",ART2.AREACOD "
      SelectStr1 = SelectStr1 & ",ARM.CUSNAM "     'ชื่อ ลูกค้า
      SelectStr1 = SelectStr1 & ",OES.SLMNAM "
      SelectStr1 = SelectStr1 & ",IST.TYPDES "
      SelectStr2 = "SELECT COUNT(*)  "
      
      FromStr = " FROM STCRD STC "
      FromStr = FromStr & "LEFT OUTER JOIN ARTRN ART2 ON (ART2.DOCNUM = STC.DOCNUM) "   '
      FromStr = FromStr & "LEFT OUTER JOIN ARMAS ARM ON (STC.PEOPLE= ARM.CUSCOD ) "
      FromStr = FromStr & "LEFT OUTER JOIN ISTAB IST ON (IST.TYPCOD = ART2.AREACOD AND IST.TABTYP = '40' )  "  ' 40 เป็น เขตการขายของ SLM
      FromStr = FromStr & " LEFT OUTER JOIN OESLM OES ON (OES.SLMCOD = STC.SLMCOD) "
      
      WhereStr = " WHERE ( STC.POSOPR = '9' ) AND (STC.TRNVAL > 0) "

 ElseIf Ind = 14 Then  'LoadDistinctSaleCustomerStcrd
      SelectStr1 = "SELECT STC.STKCOD ,STC.DOCNUM ,STC.SEQNUM ,DTOS(STC.DOCDAT) DOCDAT ,STC.PEOPLE ,STC.SLMCOD ,STC.NETVAL ,STC.STKDES ,STC.TRNQTY ,STC.UNITPR " 'รหัสสินค้า ชื่อ สินค้า
 '     SelectStr1 = SelectStr1 & ",DTOS(ART.CMPLDAT) CMPLDAT, ART.DOCNUM RENUM "     'วันที่ขึ้นเงิน
      SelectStr1 = SelectStr1 & ",ART2.AREACOD "
      SelectStr1 = SelectStr1 & ",ARM.CUSNAM "     'ชื่อ ลูกค้า
     SelectStr1 = SelectStr1 & ",OES.SLMNAM "
      SelectStr1 = SelectStr1 & ",IST.TYPDES "
      
      SelectStr2 = "SELECT COUNT(*)  "
      
      FromStr = " FROM STCRD STC "
  '    FromStr = FromStr & "LEFT OUTER JOIN ARRCPIT ARP ON (ARP.DOCNUM = STC.DOCNUM) "
      
 '     FromStr = FromStr & "LEFT OUTER JOIN ARTRN ART ON (ART.DOCNUM = ARP.RCPNUM) "
      FromStr = FromStr & "LEFT OUTER JOIN ARTRN ART2 ON (ART2.DOCNUM = STC.DOCNUM) "   '
      
      FromStr = FromStr & "LEFT OUTER JOIN ARMAS ARM ON (STC.PEOPLE= ARM.CUSCOD ) "
      FromStr = FromStr & "LEFT OUTER JOIN ISTAB IST ON (IST.TYPCOD = ART2.AREACOD AND IST.TABTYP = '40' )  "  ' 40 เป็น เขตการขายของ SLM
      
      FromStr = FromStr & " LEFT OUTER JOIN OESLM OES ON (OES.SLMCOD = STC.SLMCOD) "
      
     WhereStr = " WHERE ( STC.POSOPR = '9' ) AND (STC.TRNVAL > 0) "
   ' WhereStr = " WHERE (ART.CMPLDAT IS NOT NULL) AND (( STC.POSOPR = '0' ) OR ( STC.POSOPR = '9' )) "
 ElseIf Ind = 15 Then  'LoadDistinctSaleCustomerStcrd
      SelectStr1 = "SELECT DISTINCT STC.SLMCOD "
      SelectStr2 = "SELECT COUNT(*)  "
      
      FromStr = " FROM STCRD STC "
      WhereStr = " WHERE ( STC.POSOPR = '9' ) "
'
ElseIf Ind = 16 Then  'LoadDistinctSaleCustomerStcrd
       Set eachSLM = New Collection
      Call LoadIncenForSum(eachSLM, FROM_DOC_DATE, TO_DOC_DATE)  ' ใส่ YEAR_ID
      
      SelectStr1 = "SELECT STC.STKCOD ,STC.DOCNUM ,STC.SEQNUM ,DTOS(STC.DOCDAT) DOCDAT ,STC.PEOPLE ,STC.SLMCOD ,STC.NETVAL ,STC.STKDES ,STC.TRNQTY ,STC.UNITPR  " 'รหัสสินค้า ชื่อ สินค้า
      SelectStr1 = SelectStr1 & ", ARM.CUSNAM "     'ชื่อ ลูกค้า
      SelectStr1 = SelectStr1 & ", OES.SLMNAM "
      SelectStr2 = "SELECT COUNT(*)  "
      
      FromStr = " FROM STCRD STC "
      FromStr = FromStr & "LEFT OUTER JOIN ARMAS ARM ON (STC.PEOPLE = ARM.CUSCOD) "
      FromStr = FromStr & "LEFT OUTER JOIN ISTAB IST ON (IST.TYPCOD = ARM.CUSTYP AND IST.TABTYP = '45' )  "  ' 45 เป็นประเภท MASTER ลูกค้า
      FromStr = FromStr & " LEFT OUTER JOIN OESLM OES ON (OES.SLMCOD = STC.SLMCOD) "
      
      WhereStr = " WHERE (STC.SLMCOD IS NOT NULL AND STC.PEOPLE IS NOT NULL) AND (STC.POSOPR ='9') "

      WhereStr = WhereStr & " AND (STC.UNITPR > 0 ) "

      i = 0
      For Each temp_eachSLM In eachSLM
          i = i + 1
          If i = 1 Then
             WhereStr = WhereStr & "AND ( (STC.SLMCOD = '" & temp_eachSLM.SLMCOD & "' ) "
          Else
             WhereStr = WhereStr & " OR (STC.SLMCOD = '" & temp_eachSLM.SLMCOD & "' ) "
           End If
      Next temp_eachSLM
            WhereStr = WhereStr & ")"
            
    ElseIf Ind = 17 Then  'LoadDistinctSaleCustomerStcrd
      SelectStr1 = "SELECT STC.STKCOD ,STC.DOCNUM ,STC.SEQNUM ,DTOS(STC.DOCDAT) DOCDAT ,STC.PEOPLE ,STC.SLMCOD ,STC.NETVAL ,STC.STKDES ,STC.TRNQTY " 'รหัสสินค้า ชื่อ สินค้า
      SelectStr1 = SelectStr1 & ",ART2.AREACOD "
      SelectStr1 = SelectStr1 & ",ARM.CUSNAM "     'ชื่อ ลูกค้า
      SelectStr1 = SelectStr1 & ",OES.SLMNAM "
      SelectStr1 = SelectStr1 & ",IST.TYPDES "
      SelectStr2 = "SELECT COUNT(*)  "
      
      FromStr = " FROM STCRD STC "
      FromStr = FromStr & "LEFT OUTER JOIN ARTRN ART2 ON (ART2.DOCNUM = STC.DOCNUM) "   '
      FromStr = FromStr & "LEFT OUTER JOIN ARMAS ARM ON (STC.PEOPLE= ARM.CUSCOD ) "
      FromStr = FromStr & "LEFT OUTER JOIN ISTAB IST ON (IST.TYPCOD = ART2.AREACOD AND IST.TABTYP = '40' )  "  ' 40 เป็น เขตการขายของ SLM
      FromStr = FromStr & " LEFT OUTER JOIN OESLM OES ON (OES.SLMCOD = STC.SLMCOD) "
      
      WhereStr = " WHERE ( STC.POSOPR = '9' ) AND (STC.TRNVAL > 0) "
   ElseIf Ind = 18 Then
      SelectStr1 = "SELECT DISTINCT  STC.DOCNUM "
       SelectStr2 = "SELECT COUNT(*)  "
      
      FromStr = " FROM STCRD STC "
      
         WhereStr = " WHERE (STC.POSOPR ='9') AND (STC.TRNVAL > 0) "
   ElseIf Ind = 19 Then  'LoadDistinctSaleCustomerStcrd
      SelectStr1 = "SELECT STC.STKCOD ,STC.DOCNUM ,STC.SEQNUM ,DTOS(STC.DOCDAT) DOCDAT ,STC.PEOPLE ,STC.SLMCOD ,STC.NETVAL ,STC.STKDES ,STC.TRNQTY ,STC.UNITPR " 'รหัสสินค้า ชื่อ สินค้า
      SelectStr1 = SelectStr1 & ",ART2.AREACOD "
      SelectStr1 = SelectStr1 & ",ARM.CUSNAM "     'ชื่อ ลูกค้า
      SelectStr1 = SelectStr1 & ",OES.SLMNAM "
      SelectStr1 = SelectStr1 & ",IST.TYPDES "
      SelectStr2 = "SELECT COUNT(*)  "
      
      FromStr = " FROM STCRD STC "
      FromStr = FromStr & "LEFT OUTER JOIN ARTRN ART2 ON (ART2.DOCNUM = STC.DOCNUM) "   '
      FromStr = FromStr & "LEFT OUTER JOIN ARMAS ARM ON (STC.PEOPLE= ARM.CUSCOD ) "
      FromStr = FromStr & "LEFT OUTER JOIN ISTAB IST ON (IST.TYPCOD = ART2.AREACOD AND IST.TABTYP = '40' )  "  ' 40 เป็น เขตการขายของ SLM
      FromStr = FromStr & " LEFT OUTER JOIN OESLM OES ON (OES.SLMCOD = STC.SLMCOD) "
      
      WhereStr = " WHERE ( STC.POSOPR = '9' ) AND (STC.TRNVAL > 0) "

   ElseIf Ind = 20 Then  'LoadDistinctSaleCustomerStcrd
      SelectStr1 = "SELECT DISTINCT ART.CUSCOD,ARM.CUSNAM " 'รหัสลูกค้า ชื่อ ลูกค้า
      SelectStr2 = "SELECT COUNT(*)  "

      FromStr = " FROM STCRD STC "
      FromStr = FromStr & "LEFT OUTER JOIN ARTRN ART ON (ART.DOCNUM = STC.DOCNUM) "
      
      FromStr = FromStr & "LEFT OUTER JOIN ARMAS ARM ON (ART.CUSCOD = ARM.CUSCOD) "
      FromStr = FromStr & "LEFT OUTER JOIN ISTAB IST ON (IST.TYPCOD = ARM.CUSTYP AND IST.TABTYP = '45' )  "  ' 45 เป็นประเภท MASTER ลูกค้า
      
      FromStr = FromStr & " LEFT OUTER JOIN OESLM OES ON (OES.SLMCOD = ART.SLMCOD) "
      
      WhereStr = " WHERE (STC.SLMCOD IS NOT NULL AND ART.CUSCOD IS NOT NULL) "
      ElseIf Ind = 21 Then  'LoadSumSaleCustomerStcrd
      SelectStr1 = "SELECT  "
      SelectStr1 = SelectStr1 & " ART.CUSCOD ,STC.STKCOD ,ART.RECTYP,SUM(STC.NETVAL) NETVAL ,SUM(STC.TRNQTY) TRNQTY "
      
      SelectStr2 = "SELECT COUNT(*)  "
      
      FromStr = " FROM STCRD STC "
      FromStr = FromStr & "LEFT OUTER JOIN ARTRN ART ON (ART.DOCNUM = STC.DOCNUM) "
      FromStr = FromStr & "LEFT OUTER JOIN ARMAS ARM ON (ART.CUSCOD = ARM.CUSCOD) "
      FromStr = FromStr & "LEFT OUTER JOIN ISTAB IST ON (IST.TYPCOD = ARM.CUSTYP AND IST.TABTYP = '45' )  "  ' 45 เป็นประเภท MASTER ลูกค้า
      
      WhereStr = "WHERE (ART.CUSCOD IS NOT NULL) "
         
   ElseIf Ind = 22 Then  'LoadDistinctSaleCustomerStcrd
      SelectStr1 = "SELECT DISTINCT STC.STKCOD,STC.STKDES " 'รหัสสินค้า ชื่อสินค้า
      SelectStr2 = "SELECT COUNT(*)  "
      
      FromStr = " FROM STCRD STC "
      FromStr = FromStr & "LEFT OUTER JOIN ARTRN ART ON (ART.DOCNUM = STC.DOCNUM) "
      
      FromStr = FromStr & "LEFT OUTER JOIN ARMAS ARM ON (ART.CUSCOD = ARM.CUSCOD) "
      FromStr = FromStr & "LEFT OUTER JOIN ISTAB IST ON (IST.TYPCOD = ARM.CUSTYP AND IST.TABTYP = '45' )  "  ' 45 เป็นประเภท MASTER ผู้ขาย
      
      FromStr = FromStr & " LEFT OUTER JOIN OESLM OES ON (OES.SLMCOD = ART.SLMCOD) "
      
      WhereStr = " WHERE (STC.SLMCOD IS NOT NULL AND ART.CUSCOD IS NOT NULL) "
   

   ElseIf Ind = 23 Then  'LoadDistinctSaleCustomerStcrd
      SelectStr1 = "SELECT STC.STKCOD ,STC.DOCNUM ,STC.SEQNUM ,DTOS(STC.DOCDAT) DOCDAT ,STC.PEOPLE ,STC.SLMCOD ,STC.NETVAL ,STC.STKDES ,STC.TRNQTY ,STC.UNITPR " 'รหัสสินค้า ชื่อ สินค้า
  '    SelectStr1 = SelectStr1 & ",ART2.AREACOD "
      SelectStr1 = SelectStr1 & ",ARM.CUSNAM "     'ชื่อ ลูกค้า
      SelectStr1 = SelectStr1 & ",OES.SLMNAM "
   '   SelectStr1 = SelectStr1 & ",IST.TYPDES "
      SelectStr2 = "SELECT COUNT(*)  "
      
      FromStr = " FROM STCRD STC "
  '    FromStr = FromStr & "LEFT OUTER JOIN ARTRN ART2 ON (ART2.DOCNUM = STC.DOCNUM) "   '
      FromStr = FromStr & "LEFT OUTER JOIN ARMAS ARM ON (STC.PEOPLE= ARM.CUSCOD ) "
  '    FromStr = FromStr & "LEFT OUTER JOIN ISTAB IST ON (IST.TYPCOD = ART2.AREACOD AND IST.TABTYP = '40' )  "  ' 40 เป็น เขตการขายของ SLM
      FromStr = FromStr & " LEFT OUTER JOIN OESLM OES ON (OES.SLMCOD = STC.SLMCOD) "
      
      WhereStr = " WHERE ( STC.POSOPR = '9' ) "
      
   ElseIf Ind = 24 Then  ' ก๊อปมาจาก 23
      SelectStr1 = "SELECT STC.STKCOD ,STC.DOCNUM ,STC.SEQNUM ,DTOS(STC.DOCDAT) DOCDAT ,STC.PEOPLE ,STC.SLMCOD ,STC.NETVAL ,STC.STKDES ,STC.TRNQTY ,STC.UNITPR " 'รหัสสินค้า ชื่อ สินค้า
      SelectStr1 = SelectStr1 & ",ARM.CUSNAM "     'ชื่อ ลูกค้า
      SelectStr1 = SelectStr1 & ",OES.SLMNAM "
      SelectStr2 = "SELECT COUNT(*)  "
      
      FromStr = " FROM STCRD STC "
      FromStr = FromStr & "LEFT OUTER JOIN ARMAS ARM ON (STC.PEOPLE= ARM.CUSCOD ) "
      FromStr = FromStr & " LEFT OUTER JOIN OESLM OES ON (OES.SLMCOD = STC.SLMCOD) "
      
      WhereStr = " WHERE ( STC.POSOPR = '9' ) "
         
   ElseIf Ind = 25 Then
      SelectStr1 = "SELECT STC.* "
      SelectStr1 = SelectStr1 & ",APT.SUPCOD,APT.DOCNUM,DTOS(STC.DOCDAT) DOCDAT "     'รหัสผู้ขาย หมายเลข และวันที่เอกสาร
      SelectStr1 = SelectStr1 & ",APM.SUPNAM "     'ชื่อ
      
      SelectStr2 = "SELECT COUNT(*)  "
      
      FromStr = " FROM STCRD STC "
      FromStr = FromStr & "LEFT OUTER JOIN APTRN APT ON (APT.DOCNUM = STC.DOCNUM) "
      
      FromStr = FromStr & "LEFT OUTER JOIN APMAS APM ON (APT.SUPCOD = APM.SUPCOD) "
      FromStr = FromStr & "LEFT OUTER JOIN ISTAB IST ON (IST.TYPCOD = APM.SUPTYP AND IST.TABTYP = '46' )  "  ' 46 เป็นประเภท MASTER ผู้ขาย
      
      WhereStr = "WHERE (APT.SUPCOD IS NOT NULL) "
 ElseIf Ind = 26 Then  'LoadSumSaleCustomerStcrdVac,LoadSumSaleCustomerStcrdNonVac
      SelectStr1 = "SELECT  "
      SelectStr1 = SelectStr1 & " ART.CUSCOD ,ART.RECTYP,SUM(STC.NETVAL) NETVAL ,SUM(STC.TRNQTY) TRNQTY "
      
      SelectStr2 = "SELECT COUNT(*)  "
      
      FromStr = " FROM STCRD STC "
      FromStr = FromStr & "LEFT OUTER JOIN ARTRN ART ON (ART.DOCNUM = STC.DOCNUM) "
      FromStr = FromStr & "LEFT OUTER JOIN ARMAS ARM ON (ART.CUSCOD = ARM.CUSCOD) "
      FromStr = FromStr & "LEFT OUTER JOIN ISTAB IST ON (IST.TYPCOD = ARM.CUSTYP AND IST.TABTYP = '45' )  "  ' 45 เป็นประเภท MASTER ลูกค้า
      
      WhereStr = "WHERE (ART.CUSCOD IS NOT NULL) "
   ElseIf Ind = 27 Then  'LoadCostProductStcrd
      SelectStr1 = "SELECT STC.* "
      SelectStr1 = SelectStr1 & ",ART.CUSCOD,ARM.CUSNAM,ARM.ADDR01,ARM.ADDR02,DTOS(STC.DOCDAT) DOCDAT,ART.RECTYP,ART.DOCNUM "     'รหัสลูกค้า วันที่เอกสาร
      SelectStr1 = SelectStr1 & ",STC.SLMCOD,OES.SLMNAM "
      
      SelectStr2 = "SELECT COUNT(*)  "
      
      FromStr = " FROM STCRD STC "
      FromStr = FromStr & "LEFT OUTER JOIN ARTRN ART ON (ART.DOCNUM = STC.DOCNUM) "
      FromStr = FromStr & "LEFT OUTER JOIN ARMAS ARM ON (ART.CUSCOD = ARM.CUSCOD) "
      FromStr = FromStr & "LEFT OUTER JOIN ISTAB IST ON (IST.TYPCOD = ARM.CUSTYP AND IST.TABTYP = '45' )  "  ' 45 เป็นประเภท MASTER ลูกค้า
      FromStr = FromStr & "LEFT OUTER JOIN OESLM OES ON (OES.SLMCOD = ART.SLMCOD) "
      
      WhereStr = "WHERE (ART.CUSCOD IS NOT NULL) "
 End If
   
   If Trim(DOCNUM) <> "" Then
      If WildCard(DOCNUM, SubLen, NewStr) Then
         If WhereStr = "" Then
            WhereStr = "WHERE (STC.DOCNUM = '" & ChangeQuote(Trim(DOCNUM)) & "')"
         Else
            WhereStr = WhereStr & "AND (STC.DOCNUM = '" & ChangeQuote(Trim(DOCNUM)) & "')"
         End If
      Else
         If WhereStr = "" Then
            WhereStr = "WHERE (STC.DOCNUM = '" & ChangeQuote(Trim(DOCNUM)) & "')"
         Else
            WhereStr = WhereStr & "AND (STC.DOCNUM = '" & ChangeQuote(Trim(DOCNUM)) & "')"
         End If
      End If
   End If
   
   If Trim(SUPCOD) <> "" Then
      If WildCard(SUPCOD, SubLen, NewStr) Then
         If WhereStr = "" Then
            WhereStr = "WHERE (APT.SUPCOD LIKE '" & ChangeQuote(Trim(SUPCOD)) & "')"
         Else
            WhereStr = WhereStr & "AND (APT.SUPCOD LIKE '" & ChangeQuote(Trim(SUPCOD)) & "')"
         End If
      Else
         If WhereStr = "" Then
            WhereStr = "WHERE (APT.SUPCOD = '" & ChangeQuote(Trim(SUPCOD)) & "')"
         Else
            WhereStr = WhereStr & "AND (APT.SUPCOD = '" & ChangeQuote(Trim(SUPCOD)) & "')"
         End If
      End If
   End If
   
   If Trim(SUPTYP) <> "" Then
      If WhereStr = "" Then
         WhereStr = "WHERE (APM.SUPTYP = '" & ChangeQuote(Trim(SUPTYP)) & "')"
      Else
         WhereStr = WhereStr & "AND (APM.SUPTYP = '" & ChangeQuote(Trim(SUPTYP)) & "')"
      End If
   End If
   
   If Trim(SUPTYPSET) <> "" Then
      If WhereStr = "" Then
         WhereStr = "WHERE (APM.SUPTYP IN " & Trim(SUPTYPSET) & ")"
      Else
         WhereStr = WhereStr & "AND (APM.SUPTYP IN " & Trim(SUPTYPSET) & ")"
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
                        
   If Len(FROM_SUPPLIER_CODE) > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (APM.SUPCOD >= '" & Trim(ChangeQuote(FROM_SUPPLIER_CODE)) & "') "
      Else
         WhereStr = WhereStr & "AND (APM.SUPCOD >= '" & Trim(ChangeQuote(FROM_SUPPLIER_CODE)) & "') "
      End If
   End If
   
   If Len(TO_SUPPLIER_CODE) > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (APM.SUPCOD <= '" & Trim(ChangeQuote(TO_SUPPLIER_CODE)) & "') "
      Else
         WhereStr = WhereStr & "AND (APM.SUPCOD <= '" & Trim(ChangeQuote(TO_SUPPLIER_CODE)) & "') "
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
      
    If FROM_DOC_DATE > -1 Then                ' FROM_DOC_DATE as date =1/1/2012
      TempDate = DateToStringIntLow(Trim(FROM_DOC_DATE))
      If WhereStr = "" Then
         WhereStr = "WHERE (DTOS(STC.DOCDAT) >= '" & ChangeQuote(TempDate) & "') "
      Else
         WhereStr = WhereStr & "AND (DTOS(STC.DOCDAT) >= '" & ChangeQuote(TempDate) & "') "
      End If
   End If
  ' 'debug.print FROM_DOC_DATE
   
   If TO_DOC_DATE > -1 Then
      TempDate = DateToStringIntHi(Trim(TO_DOC_DATE))
      If WhereStr = "" Then
         WhereStr = "WHERE (DTOS(STC.DOCDAT) <= '" & ChangeQuote(TempDate) & "') "
      Else
         WhereStr = WhereStr & "AND (DTOS(STC.DOCDAT) <= '" & ChangeQuote(TempDate) & "') "
      End If
   End If
   
   If Len(FROM_CUSTOMER_CODE) > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (STC.PEOPLE >= '" & Trim(ChangeQuote(FROM_CUSTOMER_CODE)) & "') "
      Else
         WhereStr = WhereStr & "AND (STC.PEOPLE >= '" & Trim(ChangeQuote(FROM_CUSTOMER_CODE)) & "') "
      End If
   End If
   
   If Len(TO_CUSTOMER_CODE) > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (STC.PEOPLE <= '" & Trim(ChangeQuote(TO_CUSTOMER_CODE)) & "') "
      Else
         WhereStr = WhereStr & "AND (STC.PEOPLE <= '" & Trim(ChangeQuote(TO_CUSTOMER_CODE)) & "') "
      End If
   End If
   
   If Len(FROM_STOCK_CODE) > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (STC.STKCOD >= '" & Trim(ChangeQuote(FROM_STOCK_CODE)) & "') "
      Else
         WhereStr = WhereStr & "AND (STC.STKCOD >= '" & Trim(ChangeQuote(FROM_STOCK_CODE)) & "') "
      End If
   End If
   
   If Len(TO_STOCK_CODE) > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (STC.STKCOD <= '" & Trim(ChangeQuote(TO_STOCK_CODE)) & "') "
      Else
         WhereStr = WhereStr & "AND (STC.STKCOD <= '" & Trim(ChangeQuote(TO_STOCK_CODE)) & "') "
      End If
   End If
   
   If Trim(FROM_SALE_CODE) <> "" Then
      If WhereStr = "" Then
         WhereStr = "WHERE (STC.SLMCOD >= '" & ChangeQuote(Trim(FROM_SALE_CODE)) & "')"
      Else
         WhereStr = WhereStr & " AND (STC.SLMCOD >= '" & ChangeQuote(Trim(FROM_SALE_CODE)) & "')"
      End If
   End If
   
   If Trim(TO_SALE_CODE) <> "" Then
      If WhereStr = "" Then
         WhereStr = "WHERE (STC.SLMCOD <= '" & ChangeQuote(Trim(TO_SALE_CODE)) & "')"
      Else
         WhereStr = WhereStr & " AND (STC.SLMCOD <= '" & ChangeQuote(Trim(TO_SALE_CODE)) & "')"
      End If
   End If
   
      ' Commission
    If Ind = 9 And FROM_CMPL_DATE > -1 Then
      TempDate = DateToStringIntLow(Trim(FROM_CMPL_DATE))
      If WhereStr = "" Then
         WhereStr = "WHERE (DTOS(ART.CMPLDAT) >= '" & ChangeQuote(TempDate) & "') "
      Else
         WhereStr = WhereStr & "AND (DTOS(ART.CMPLDAT) >= '" & ChangeQuote(TempDate) & "') "
      End If
   End If

   If Ind = 9 And TO_CMPL_DATE > -1 Then
      TempDate = DateToStringIntHi(Trim(TO_CMPL_DATE))
      If WhereStr = "" Then
         WhereStr = "WHERE (DTOS(ART.CMPLDAT) <= '" & ChangeQuote(TempDate) & "') "
      Else
         WhereStr = WhereStr & "AND (DTOS(ART.CMPLDAT) <= '" & ChangeQuote(TempDate) & "') "
      End If
      'debug.print FROM_CMPL_DATE
   End If
   
   If Trim(AREACOD) <> "" Then
      If WhereStr = "" Then
         WhereStr = "WHERE (ART2.AREACOD = '" & Trim(AREACOD) & "') "
      Else
         WhereStr = WhereStr & "AND (ART2.AREACOD = '" & Trim(AREACOD) & "') "            '& Trim(AREACOD) &
      End If
   End If
   
   If Trim(SLMCOD) <> "" Then
      If WhereStr = "" Then
         WhereStr = "WHERE (STC.SLMCOD = '" & ChangeQuote(Trim(SLMCOD)) & "') "
      Else
         WhereStr = WhereStr & "AND (STC.SLMCOD = '" & ChangeQuote(Trim(SLMCOD)) & "') "
      End If
   End If
   
      If Trim(PEOPLE) <> "" Then
      If WhereStr = "" Then
         WhereStr = "WHERE (STC.PEOPLE = '" & ChangeQuote(Trim(PEOPLE)) & "') "
      Else
         WhereStr = WhereStr & "AND (STC.PEOPLE = '" & ChangeQuote(Trim(PEOPLE)) & "') "
      End If
   End If
   
      If Trim(STKCOD) <> "" Then
      If WhereStr = "" Then
         WhereStr = "WHERE (STC.STKCOD = '" & ChangeQuote(Trim(STKCOD)) & "') "
      Else
         WhereStr = WhereStr & "AND (STC.STKCOD = '" & ChangeQuote(Trim(STKCOD)) & "') "
      End If
   End If
   
   If Trim(STKDES) <> "" Then
      If WhereStr = "" Then
         WhereStr = "WHERE (STC.STKDES LIKE '" & ChangeQuote(Trim(STKDES)) & "') "
      Else
         WhereStr = WhereStr & "AND (STC.STKDES LIKE '" & ChangeQuote(Trim(STKDES)) & "') "
      End If
   End If
   
   If Trim(TEMP_ADDR1) <> "" Then
      If WhereStr = "" Then
         WhereStr = "WHERE (APM.ADDR01 LIKE '" & ChangeQuote(Trim(TEMP_ADDR1)) & "' OR APM.ADDR02 LIKE '" & ChangeQuote(Trim(TEMP_ADDR1)) & "') "
      Else
         WhereStr = WhereStr & "AND (APM.ADDR01 LIKE '" & ChangeQuote(Trim(TEMP_ADDR1)) & "' OR APM.ADDR02 LIKE '" & ChangeQuote(Trim(TEMP_ADDR1)) & "') "
      End If
   End If
   
   If Trim(TEMP_ADDR2) <> "" Then
      If WhereStr = "" Then
         WhereStr = "WHERE (APM.ADDR01 LIKE '" & ChangeQuote(Trim(TEMP_ADDR2)) & "' OR APM.ADDR02 LIKE '" & ChangeQuote(Trim(TEMP_ADDR2)) & "') "
      Else
         WhereStr = WhereStr & "AND (APM.ADDR01 LIKE '" & ChangeQuote(Trim(TEMP_ADDR2)) & "' OR APM.ADDR02 LIKE '" & ChangeQuote(Trim(TEMP_ADDR2)) & "') "
      End If
   End If
   
   'TEMP_ADDR1
   If Ind = 10 Then
     GroupByStr = "GROUP BY  STC.DOCNUM , STC.DOCDAT "
   ElseIf Ind = 11 Then
      GroupByStr = "GROUP BY  STC.STKCOD , STC.STKDES "
   ElseIf Ind = 12 Then
      GroupByStr = "GROUP BY  STC.PEOPLE "
   End If
   
   OrderByStr = ""
   OrderTypeStr = ""
   
   If Ind = 1 Then
      OrderByStr = " ORDER BY  APM.SUPTYP ASC,APT.SUPCOD ASC,STC.DOCDAT ASC,APT.DOCNUM ASC,STC.STKCOD   "
   ElseIf Ind = 2 Then
      OrderByStr = " ORDER BY  STC.STKCOD   "
   ElseIf Ind = 3 Then
      OrderByStr = " ORDER BY ART.CUSCOD ASC,STC.STKCOD   "
   ElseIf Ind = 4 Then
      OrderByStr = " ORDER BY ART.CUSCOD ASC,STC.STKCOD   "
   ElseIf Ind = 5 Then
      OrderByStr = " ORDER BY  STC.STKCOD ASC,ART.CUSCOD   "
   ElseIf Ind = 6 Then
      OrderByStr = " ORDER BY  STC.SLMCOD ASC,ART.CUSCOD ASC,STC.STKCOD   "
   ElseIf Ind = 7 Then
      If OrderBy = 2 Then
         OrderByStr = " ORDER BY  STC.SLMCOD ASC, STC.STKCOD ASC, ART.CUSCOD "
      Else
         OrderByStr = " ORDER BY  STC.SLMCOD ASC,ART.CUSCOD ASC,STC.STKCOD   "
      End If
   ElseIf Ind = 8 Then
      OrderByStr = " ORDER BY  STC.STKDES "
  ElseIf Ind = 9 Then
      OrderByStr = " ORDER BY  STC.DOCNUM ASC,  STC.DOCDAT ASC, ART.CMPLDAT "         'ART2.AREACOD ASC,
 ElseIf Ind = 10 Then
      OrderByStr = " ORDER BY  STC.DOCNUM   "         'ART2.AREACOD ASC,
  ElseIf Ind = 11 Then
      OrderByStr = " ORDER BY  STC.STKCOD   "         'ART2.AREACOD ASC,
 ElseIf Ind = 12 Then
      OrderByStr = " ORDER BY  STC.PEOPLE   "         'ART2.AREACOD ASC,
  ElseIf Ind = 13 Then
      If OrderBy = 1 Then
         OrderByStr = " ORDER BY  STC.STKCOD ASC,  STC.PEOPLE ASC , STC.DOCNUM "
      Else
         OrderByStr = " ORDER BY  STC.DOCNUM ASC,  STC.DOCDAT "         'ART2.AREACOD ASC,
      End If
  ElseIf Ind = 14 Then
      OrderByStr = " ORDER BY  STC.PEOPLE   "         'ART2.AREACOD ASC,
  ElseIf Ind = 15 Then
     OrderByStr = " ORDER BY  STC.SLMCOD   "         'ART2.AREACOD ASC,
   ElseIf Ind = 16 Then
      OrderByStr = " ORDER BY STC.STKCOD ASC, STC.SLMCOD ASC , STC.DOCNUM  "
   ElseIf Ind = 17 Then
      OrderByStr = " ORDER BY  STC.SLMCOD ASC,ART2.AREACOD ASC,STC.DOCNUM   "
    ElseIf Ind = 18 Then
      OrderByStr = " ORDER BY  STC.DOCNUM   "         'ART2.AREACOD ASC,
   ElseIf Ind = 19 Then
      OrderByStr = " ORDER BY  STC.PEOPLE ASC, STC.DOCNUM ASC,  STC.DOCDAT "         'ART2.AREACOD ASC,
   ElseIf Ind = 20 Then
      OrderByStr = " ORDER BY  ART.CUSCOD   "
    ElseIf Ind = 21 Then
      GroupByStr = "GROUP BY  ART.CUSCOD , STC.STKCOD ,ART.RECTYP "
   ElseIf Ind = 22 Then
      OrderByStr = " ORDER BY  STC.STKCOD   "
    ElseIf Ind = 23 Then
      OrderByStr = " ORDER BY STC.PEOPLE  ASC, STC.STKCOD "       'STC.PEOPLE ASC,
    ElseIf Ind = 24 Then
      OrderByStr = " ORDER BY STC.STKCOD "       'STC.PEOPLE ASC,
   ElseIf Ind = 25 Then
      OrderByStr = " ORDER BY  APT.SUPCOD ASC,STC.STKCOD ASC ,STC.DOCDAT ASC,APT.DOCNUM "
   ElseIf Ind = 26 Then
      GroupByStr = "GROUP BY  ART.CUSCOD ,ART.RECTYP "
   ElseIf Ind = 27 Then
      If OrderBy = 2 Then
         OrderByStr = " ORDER BY  STC.SLMCOD ASC,STC.DOCDAT ASC,ART.CUSCOD ASC,STC.STKCOD "
      Else
         OrderByStr = " ORDER BY  STC.DOCDAT ASC,ART.CUSCOD ASC,STC.STKCOD   "
      End If
  End If
   
   If OrderType = 1 Then
      OrderTypeStr = "ASC "
   ElseIf OrderType = 2 Then
      OrderTypeStr = "DESC "
   Else
      OrderTypeStr = "ASC "
   End If
   
  If Len(OrderByStr) <= 0 Then
   OrderTypeStr = ""
  End If
  
   SQL1 = SelectStr1 & FromStr & WhereStr & GroupByStr & OrderByStr & OrderTypeStr
   SQL2 = SelectStr2 & FromStr & WhereStr
 '  'debug.print SQL1
   
   If db_num = 2 Then
      Call glbDatabaseMngr.GetRs2(SQL1, SQL2, True, ItemCount, Rs, glbErrorLog)
   Else
      Call glbDatabaseMngr.GetRs(SQL1, SQL2, True, ItemCount, Rs, glbErrorLog)
   End If
   
End Function

Public Function PopulateFromRS(Ind As Long, Rs As ADODB.Recordset, Optional Count As Long)
   If Ind = 1 Then
      DOCNUM = NVLS(Rs("DOCNUM"), "")
      DOCDAT = InternalDateToDate(NVLS(Rs("DOCDAT"), ""))
      SUPCOD = NVLS(Rs("SUPCOD"), "")
      SUPNAM = NVLS(Rs("SUPNAM"), "")
      
      STKCOD = NVLS(Rs("STKCOD"), "")
      SLMCOD = NVLS(Rs("SLMCOD"), "")
      STKDES = NVLS(Rs("STKDES"), "")
      
      AMOUNT = NVLD(Rs("AMOUNT"), 0)
      TRNQTY = NVLD(Rs("TRNQTY"), 0)
      NETVAL = NVLD(Rs("NETVAL"), 0)
      UNITPR = NVLD(Rs("UNITPR"), 0)
      
      LOCCOD = NVLS(Rs("LOCCOD"), "")
      SUPTYP = NVLS(Rs("SUPTYP"), "")
      SUPTYPNAM = NVLS(Rs("TYPDES"), "")
      
      PAYAMT = NVLD(Rs("PAYAMT"), 0)
      
      CHQNUM = NVLS(Rs("CHQNUM"), "")
      GETDAT = InternalDateToDate(NVLS(Rs("GETDAT"), ""))
      CHQDAT = InternalDateToDate(NVLS(Rs("CHQDAT"), ""))
      
      ADDR01 = NVLS(Rs("ADDR01"), "")
      ADDR02 = NVLS(Rs("ADDR02"), "")
      TELNUM = NVLS(Rs("TELNUM"), "")
            
      SEQNUM = NVLS(Rs("SEQNUM"), "")
            
      POSOPR = NVLS(Rs("POSOPR"), "")
   ElseIf Ind = 2 Then
      STKCOD = NVLS(Rs("STKCOD"), "")
      STKDES = NVLS(Rs("STKDES"), "")
      
      TRNQTY = NVLD(Rs("TRNQTY"), 0)
      NETVAL = NVLD(Rs("NETVAL"), 0)
      UNITPR = NVLD(Rs("UNITPR"), 0)
   ElseIf Ind = 3 Then
      DOCDAT = InternalDateToDate(NVLS(Rs("DOCDAT"), ""))
      CUSCOD = NVLS(Rs("CUSCOD"), "")
      
      STKCOD = NVLS(Rs("STKCOD"), "")
      
      TRNQTY = NVLD(Rs("TRNQTY"), 0)
      NETVAL = NVLD(Rs("NETVAL"), 0)
      UNITPR = NVLD(Rs("UNITPR"), 0)
      
      RECTYP = NVLI(Rs("RECTYP"), -1)
      DOCNUM = NVLS(Rs("DOCNUM"), "")
   ElseIf Ind = 4 Then
      CUSCOD = NVLS(Rs("CUSCOD"), "")
      CUSNAM = NVLS(Rs("CUSNAM"), "")
      
      STKCOD = NVLS(Rs("STKCOD"), "")
      STKDES = NVLS(Rs("STKDES"), "")
   ElseIf Ind = 5 Then
      CUSCOD = NVLS(Rs("CUSCOD"), "")
      CUSNAM = NVLS(Rs("CUSNAM"), "")
      
      STKCOD = NVLS(Rs("STKCOD"), "")
      STKDES = NVLS(Rs("STKDES"), "")
   ElseIf Ind = 6 Then
      DOCDAT = InternalDateToDate(NVLS(Rs("DOCDAT"), ""))
      CUSCOD = NVLS(Rs("CUSCOD"), "")
      
      SLMCOD = NVLS(Rs("SLMCOD"), "")
      
      STKCOD = NVLS(Rs("STKCOD"), "")
      
      TRNQTY = NVLD(Rs("TRNQTY"), 0)
      NETVAL = NVLD(Rs("NETVAL"), 0)
      UNITPR = NVLD(Rs("UNITPR"), 0)
      
      RECTYP = NVLI(Rs("RECTYP"), -1)
      DOCNUM = NVLS(Rs("DOCNUM"), "")
   ElseIf Ind = 7 Then
      CUSCOD = NVLS(Rs("CUSCOD"), "")
      CUSNAM = NVLS(Rs("CUSNAM"), "")
      
      STKCOD = NVLS(Rs("STKCOD"), "")
      STKDES = NVLS(Rs("STKDES"), "")
      
      SLMCOD = NVLS(Rs("SLMCOD"), "")
      SLMNAM = NVLS(Rs("SLMNAM"), "")
   ElseIf Ind = 8 Then
      STKCOD = NVLS(Rs("STKCOD"), "")
      STKDES = NVLS(Rs("STKDES"), "")
      
      KEY_ID = Count                       ' สำหรับ lookup
      KEY_LOOKUP = NVLS(Rs("STKCOD"), "")
  ElseIf Ind = 9 Then
      DOCNUM = NVLS(Rs("DOCNUM"), "")
      DOCDAT = InternalDateToDate(NVLS(Rs("DOCDAT"), ""))
      CMPLDAT = InternalDateToDate(NVLS(Rs("CMPLDAT"), ""))
      SEQNUM = NVLS(Rs("SEQNUM"), "")
      
      SLMCOD = NVLS(Rs("SLMCOD"), "")
      SLMNAM = NVLS(Rs("SLMNAM"), "")
      
      STKCOD = NVLS(Rs("STKCOD"), "")
      STKDES = NVLS(Rs("STKDES"), "")
      
      NETVAL = NVLD(Rs("NETVAL"), 0)

      CUSCOD = NVLS(Rs("PEOPLE"), "")
      CUSNAM = NVLS(Rs("CUSNAM"), "")
      
      RENUM = NVLS(Rs("RENUM"), "")
      TYPDES = NVLS(Rs("TYPDES"), "")
      AREACOD = NVLS(Rs("AREACOD"), "")
      AREANAM = NVLS(Rs("TYPDES"), "")
      TRNQTY = NVLD(Rs("TRNQTY"), 0)
    ElseIf Ind = 10 Then
      DOCNUM = NVLS(Rs("DOCNUM"), "")
      DOCDAT = InternalDateToDate(NVLS(Rs("DOCDAT"), ""))
      NETVAL = NVLD(Rs("NETVAL"), 0)
    ElseIf Ind = 11 Then
           STKCOD = NVLS(Rs("STKCOD"), "")
           TRNQTY = NVLD(Rs("TRNQTY"), 0)
           STKDES = NVLS(Rs("STKDES"), "")
           
   ElseIf Ind = 12 Then
            CUSCOD = NVLS(Rs("CUSCOD"), "")
           NETVAL = NVLD(Rs("NETVAL"), 0)
         '  DOCDAT = InternalDateToDate(NVLS(Rs("DOCDAT"), ""))
  ElseIf Ind = 13 Then
      DOCNUM = NVLS(Rs("DOCNUM"), "")
      DOCDAT = InternalDateToDate(NVLS(Rs("DOCDAT"), ""))
      SEQNUM = NVLS(Rs("SEQNUM"), "")
      
      SLMCOD = NVLS(Rs("SLMCOD"), "")
      SLMNAM = NVLS(Rs("SLMNAM"), "")
      
      STKCOD = NVLS(Rs("STKCOD"), "")
      STKDES = NVLS(Rs("STKDES"), "")
      
      NETVAL = NVLD(Rs("NETVAL"), 0)

      CUSCOD = NVLS(Rs("PEOPLE"), "")
      CUSNAM = NVLS(Rs("CUSNAM"), "")
      
      UNITPR = NVLD(Rs("UNITPR"), 0)
      TYPDES = NVLS(Rs("TYPDES"), "")
      AREACOD = NVLS(Rs("AREACOD"), "")
      AREANAM = NVLS(Rs("TYPDES"), "")
      TRNQTY = NVLD(Rs("TRNQTY"), 0)
   ElseIf Ind = 14 Then
      DOCNUM = NVLS(Rs("DOCNUM"), "")
      DOCDAT = InternalDateToDate(NVLS(Rs("DOCDAT"), ""))
 '     CMPLDAT = InternalDateToDate(NVLS(Rs("CMPLDAT"), ""))
      SEQNUM = NVLS(Rs("SEQNUM"), "")
      
      SLMCOD = NVLS(Rs("SLMCOD"), "")
      SLMNAM = NVLS(Rs("SLMNAM"), "")
      
      STKCOD = NVLS(Rs("STKCOD"), "")
      STKDES = NVLS(Rs("STKDES"), "")
      
      NETVAL = NVLD(Rs("NETVAL"), 0)

      CUSCOD = NVLS(Rs("PEOPLE"), "")
      CUSNAM = NVLS(Rs("CUSNAM"), "")
      
      UNITPR = NVLD(Rs("UNITPR"), 0)
      TYPDES = NVLS(Rs("TYPDES"), "")
      AREACOD = NVLS(Rs("AREACOD"), "")
      AREANAM = NVLS(Rs("TYPDES"), "")
      TRNQTY = NVLD(Rs("TRNQTY"), 0)
   ElseIf Ind = 15 Then
      SLMCOD = NVLS(Rs("SLMCOD"), "")
  ElseIf Ind = 16 Then
      DOCNUM = NVLS(Rs("DOCNUM"), "")
      DOCDAT = InternalDateToDate(NVLS(Rs("DOCDAT"), ""))
      
      CUSCOD = NVLS(Rs("PEOPLE"), "")
      CUSNAM = NVLS(Rs("CUSNAM"), "")
      
      STKCOD = NVLS(Rs("STKCOD"), "")
      STKDES = NVLS(Rs("STKDES"), "")
      
      SLMCOD = NVLS(Rs("SLMCOD"), "")
      SLMNAM = NVLS(Rs("SLMNAM"), "")
      
      TRNQTY = NVLD(Rs("TRNQTY"), 0)
      NETVAL = NVLD(Rs("NETVAL"), 0)
      UNITPR = NVLD(Rs("UNITPR"), 0)
 ElseIf Ind = 17 Then
      DOCNUM = NVLS(Rs("DOCNUM"), "")
      DOCDAT = InternalDateToDate(NVLS(Rs("DOCDAT"), ""))
      
      SLMCOD = NVLS(Rs("SLMCOD"), "")
      SLMNAM = NVLS(Rs("SLMNAM"), "")
      
      STKCOD = NVLS(Rs("STKCOD"), "")
      STKDES = NVLS(Rs("STKDES"), "")
      
      NETVAL = NVLD(Rs("NETVAL"), 0)

      CUSCOD = NVLS(Rs("PEOPLE"), "")
      CUSNAM = NVLS(Rs("CUSNAM"), "")
      
      AREACOD = NVLS(Rs("AREACOD"), "")
      AREANAM = NVLS(Rs("TYPDES"), "")
  ElseIf Ind = 18 Then
      DOCNUM = NVLS(Rs("DOCNUM"), "")
  ElseIf Ind = 19 Then

      DOCNUM = NVLS(Rs("DOCNUM"), "")
      
      DOCDAT = InternalDateToDate(NVLS(Rs("DOCDAT"), ""))
      SEQNUM = NVLS(Rs("SEQNUM"), "")
      
      SLMCOD = NVLS(Rs("SLMCOD"), "")
      SLMNAM = NVLS(Rs("SLMNAM"), "")
      
      STKCOD = NVLS(Rs("STKCOD"), "")
      STKDES = NVLS(Rs("STKDES"), "")
      
      NETVAL = NVLD(Rs("NETVAL"), 0)

      CUSCOD = NVLS(Rs("PEOPLE"), "")
      CUSNAM = NVLS(Rs("CUSNAM"), "")
      
      UNITPR = NVLD(Rs("UNITPR"), 0)
      TYPDES = NVLS(Rs("TYPDES"), "")
      AREACOD = NVLS(Rs("AREACOD"), "")
      AREANAM = NVLS(Rs("TYPDES"), "")
      TRNQTY = NVLD(Rs("TRNQTY"), 0)

   ElseIf Ind = 22 Then
      STKCOD = NVLS(Rs("STKCOD"), "")
      STKDES = NVLS(Rs("STKDES"), "")
   ElseIf Ind = 20 Then
      CUSCOD = NVLS(Rs("CUSCOD"), "")
      CUSNAM = NVLS(Rs("CUSNAM"), "")
   ElseIf Ind = 21 Then
      CUSCOD = NVLS(Rs("CUSCOD"), "")
      
      STKCOD = NVLS(Rs("STKCOD"), "")
      TRNQTY = NVLD(Rs("TRNQTY"), 0)
      NETVAL = NVLD(Rs("NETVAL"), 0)
      
      RECTYP = NVLI(Rs("RECTYP"), -1)
    ElseIf Ind = 23 Then
      DOCNUM = NVLS(Rs("DOCNUM"), "")
      DOCDAT = InternalDateToDate(NVLS(Rs("DOCDAT"), ""))
      SEQNUM = NVLS(Rs("SEQNUM"), "")
      
      SLMCOD = NVLS(Rs("SLMCOD"), "")
      SLMNAM = NVLS(Rs("SLMNAM"), "")
      
      STKCOD = NVLS(Rs("STKCOD"), "")
      STKDES = NVLS(Rs("STKDES"), "")
      
      NETVAL = NVLD(Rs("NETVAL"), 0)

      CUSCOD = NVLS(Rs("PEOPLE"), "")
      CUSNAM = NVLS(Rs("CUSNAM"), "")
      
      UNITPR = NVLD(Rs("UNITPR"), 0)
   '   AREACOD = NVLS(Rs("AREACOD"), "")
      TRNQTY = NVLD(Rs("TRNQTY"), 0)
      
    ElseIf Ind = 24 Then
      DOCNUM = NVLS(Rs("DOCNUM"), "")
      DOCDAT = InternalDateToDate(NVLS(Rs("DOCDAT"), ""))
      SEQNUM = NVLS(Rs("SEQNUM"), "")
      
      SLMCOD = NVLS(Rs("SLMCOD"), "")
      SLMNAM = NVLS(Rs("SLMNAM"), "")
      
      STKCOD = NVLS(Rs("STKCOD"), "")
      STKDES = NVLS(Rs("STKDES"), "")
      
      NETVAL = NVLD(Rs("NETVAL"), 0)

      CUSCOD = NVLS(Rs("PEOPLE"), "")
      CUSNAM = NVLS(Rs("CUSNAM"), "")
      
      UNITPR = NVLD(Rs("UNITPR"), 0)
   '   AREACOD = NVLS(Rs("AREACOD"), "")
      TRNQTY = NVLD(Rs("TRNQTY"), 0)
   ElseIf Ind = 25 Then
      DOCNUM = NVLS(Rs("DOCNUM"), "")
      DOCDAT = InternalDateToDate(NVLS(Rs("DOCDAT"), ""))
      SUPCOD = NVLS(Rs("SUPCOD"), "")
      SUPNAM = NVLS(Rs("SUPNAM"), "")
      
      STKCOD = NVLS(Rs("STKCOD"), "")
      STKDES = NVLS(Rs("STKDES"), "")
      
      TRNQTY = NVLD(Rs("TRNQTY"), 0)
      NETVAL = NVLD(Rs("NETVAL"), 0)
      UNITPR = NVLD(Rs("UNITPR"), 0)
      
      POSOPR = NVLS(Rs("POSOPR"), "")
    ElseIf Ind = 26 Then
      CUSCOD = NVLS(Rs("CUSCOD"), "")
      TRNQTY = NVLD(Rs("TRNQTY"), 0)
      NETVAL = NVLD(Rs("NETVAL"), 0)
      
      RECTYP = NVLI(Rs("RECTYP"), -1)
   ElseIf Ind = 27 Then
      DOCDAT = InternalDateToDate(NVLS(Rs("DOCDAT"), ""))
      CUSCOD = NVLS(Rs("CUSCOD"), "")
      CUSNAM = NVLS(Rs("CUSNAM"), "")
      ADDR01 = Trim(NVLS(Rs("ADDR01"), ""))
      ADDR02 = Trim(NVLS(Rs("ADDR02"), ""))
      
      STKCOD = NVLS(Rs("STKCOD"), "")
      STKDES = NVLS(Rs("STKDES"), "")
      
      SLMCOD = NVLS(Rs("SLMCOD"), "")
      SLMNAM = NVLS(Rs("SLMNAM"), "")
      
      TRNQTY = NVLD(Rs("TRNQTY"), 0)
      NETVAL = NVLD(Rs("NETVAL"), 0)
      UNITPR = NVLD(Rs("UNITPR"), 0)
      
      RECTYP = NVLI(Rs("RECTYP"), -1)
      DOCNUM = NVLS(Rs("DOCNUM"), "")
   End If
End Function

