VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CReportAR005_3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'User Group Report
Implements CReportInterface

Private Const MODULE_NAME = "CReportAR005_3"
Public MODULE_DESC As String

Private WithEvents Vsp As VSPrinter
Attribute Vsp.VB_VarHelpID = -1
Private mstrErrMsg As String
Private mcolParam As Collection
Private mblnNewPage As Boolean
Private mblnHeader As Boolean
Private mblnEndPage As Boolean
Private mdblWid   As Double
Private mdteDate As Date
Private mdY As Double
Private Rs As ADODB.Recordset

Private Const TITLE_SIZE = 14
Private Const HEADER_SIZE = 12
Private Const DETAIL_SIZE = 12

Private Const TITLE_FORMAT = "<32|<3|<5;"

Private m_TempCol As Collection
Private m_HeaderFormat1 As String
Private m_HeaderFormat2 As String
Private m_HeaderText1 As String
Private m_HeaderText2 As String
Private SaleWidth As Double

Private m_Headers0 As CFieldList
Private m_Headers1 As CFieldList
Private m_Headers1_1 As CFieldList
Private m_Headers1_2 As CFieldList
Private m_Headers1_3 As CFieldList
Private m_Headers1_4 As CFieldList
Private m_Details1 As CFieldList
Private m_Details2 As CFieldList
Private m_Left As Double
Private m_YearSeqs As Collection

Private m_ReceiveAmounts As Collection
Private m_BkReceiveAmounts As Collection
Private m_BkChqAmounts As Collection
Private m_BktSame1_15_1 As Collection
Dim m_ReceiveCancelitem As Collection
Private m_ReceiveAllAmounts As Collection
Private m_MaxMins As Collection
Private m_DbnCdnByBills As Collection
Private m_RealCreditColl As Collection
Private m_CnAllAmounts As Collection
Dim collum As Long
Private numj As Long

Dim TempCCheck As CCheckCancel
Dim TempCheckCancel As CCheckCancel
'Dim ItemCount As Long
 Dim m_Rs As ADODB.Recordset
Dim m_BktAmountForCheckInCancel As Collection
 

Private Sub Class_Initialize()
   Set mcolParam = New Collection
   Set m_TempCol = New Collection
   
   Set m_Headers0 = New CFieldList
   Set m_Headers1 = New CFieldList
   Set m_Headers1_1 = New CFieldList
   Set m_Headers1_2 = New CFieldList
   Set m_Headers1_3 = New CFieldList
   Set m_Headers1_4 = New CFieldList
   Set m_Details1 = New CFieldList
   Set m_Details2 = New CFieldList
   Set m_YearSeqs = New Collection
   Set m_BkReceiveAmounts = New Collection
      Set m_BkChqAmounts = New Collection
   Set m_BktSame1_15_1 = New Collection
      Set m_ReceiveAmounts = New Collection
   Set m_MaxMins = New Collection
   Set m_ReceiveAllAmounts = New Collection
   Set m_DbnCdnByBills = New Collection
   Set m_RealCreditColl = New Collection
   Set m_CnAllAmounts = New Collection
End Sub

Private Sub Class_Terminate()
   Call ClearParam
   Set mcolParam = Nothing
   Set Vsp = Nothing
   Set m_TempCol = Nothing
   
   Set m_Headers0 = Nothing
   Set m_Headers1 = Nothing
   Set m_Details1 = Nothing
   Set m_YearSeqs = Nothing
   Set m_Headers1_1 = Nothing
   Set m_Headers1_2 = Nothing
   Set m_Headers1_3 = Nothing
   Set m_Headers1_4 = Nothing
   Set m_Details2 = Nothing
   Set m_ReceiveAmounts = Nothing
   Set m_BkReceiveAmounts = Nothing
   Set m_BkChqAmounts = Nothing
   Set m_BktSame1_15_1 = Nothing
   Set m_MaxMins = Nothing
   Set m_ReceiveAllAmounts = Nothing
   Set m_DbnCdnByBills = Nothing
   Set m_RealCreditColl = Nothing
   Set m_CnAllAmounts = Nothing
End Sub

Private Function CReportInterface_AddParam(varItem As Variant, strIndex As String) As Boolean
   Call mcolParam.Add(varItem, strIndex)
   CReportInterface_AddParam = True
End Function

Private Sub CReportInterface_ClearParam()
   Call ClearParam
End Sub

Private Property Get CReportInterface_ErrorMsg() As String
   CReportInterface_ErrorMsg = mstrErrMsg
End Property

Private Function CReportInterface_Preview() As Boolean
   CReportInterface_Preview = genDoc(True)
End Function

Private Function CReportInterface_PrintDoc() As Boolean
   CReportInterface_PrintDoc = genDoc(False)
End Function

Private Property Set CReportInterface_VsPrint(RHS As VSPrinter7LibCtl.IVSPrinter)
   Set Vsp = RHS
End Property

Private Sub ClearParam()
   Dim i As Long
   
   For i = 1 To mcolParam.Count
      mcolParam.Remove 1
   Next i

End Sub
'============================= Codes above used in every report ==========================

Private Sub printHeader()
Dim strFormat As String
Dim strPrint As String
Dim tbdT As TableBorderSettings
Dim blnBold As Boolean
Dim iSize As Integer
Dim alngX() As Long
Dim PrevID As Long
Dim TempStr1 As String
Dim TempStr1_1 As String
Dim Cf As CReportField
Dim ExportWidth As Long
Dim HeadCf As CReportField
Dim j As Long
Dim TempStr As String
Dim TempRs As ADODB.Recordset
Dim Amt As Double
Dim iCount As Long
Dim FromDate As Date
Dim ToDate As Date
Dim TempDate As Date
Dim TempWidth As Double
   
   tbdT = Vsp.TableBorder
   blnBold = Vsp.FontBold
   iSize = Vsp.FontSize
   
   Vsp.FontSize = HEADER_SIZE
   Vsp.FontBold = True
   Vsp.TableBorder = tbAll
   
   Call m_Headers0.ClearField
   Call m_Headers1.ClearField
   
   Set Cf = New CReportField
      
   '===
   numj = 0
   SaleWidth = 0
   If mcolParam("SUMMARY_MODE") = "Y" Or mcolParam("SALE_AND_CUSTOMER") = "Y" Then
      Call Cf.SetFieldValue(0, "^", "NO.", "^")
      Call m_Headers0.AddField(Cf)
   Else
      Call Cf.SetFieldValue(1, "^", "NO.", "^")
      Call m_Headers0.AddField(Cf)
   End If
      
   If mcolParam("SUMMARY_MODE") = "Y" Then
      Call Cf.SetFieldValue(2, "<", "", "<")
      Call m_Headers0.AddField(Cf)
      Call Cf.SetFieldValue(4, "<", "", "<")
      Call m_Headers0.AddField(Cf)
   ElseIf mcolParam("SALE_AND_CUSTOMER") = "Y" Then
      Call Cf.SetFieldValue(2, "<", "", "<")
      Call m_Headers0.AddField(Cf)
      Call Cf.SetFieldValue(6, "<", "", "<")
      Call m_Headers0.AddField(Cf)
      
      SaleWidth = SaleWidth + 6
   Else
      Call Cf.SetFieldValue(2, "^", "เลขที่", "^")
      Call m_Headers0.AddField(Cf)
      Call Cf.SetFieldValue(2, "^", "", "^")
      Call m_Headers0.AddField(Cf)
      SaleWidth = SaleWidth + 2
   End If
   
   If mcolParam("SUMMARY_MODE") = "Y" Or mcolParam("SALE_AND_CUSTOMER") = "Y" Then
      Call Cf.SetFieldValue(0, "^", "", "^")
      Call m_Headers0.AddField(Cf)
   Else
      Call Cf.SetFieldValue(2, "^", "", "^")
      Call m_Headers0.AddField(Cf)
      SaleWidth = SaleWidth + 2
   End If
   
   FromDate = mcolParam("FROM_MONTH")
   ToDate = mcolParam("TO_MONTH")
   Call GetFirstLastDate(mcolParam("FROM_MONTH"), FromDate, TempDate)
   Call GetFirstLastDate(mcolParam("TO_MONTH"), TempDate, ToDate)
      
   TempWidth = 0
   While (FromDate < ToDate)
      TempWidth = TempWidth + 2
      SaleWidth = SaleWidth + 2
      FromDate = DateAdd("M", 1, FromDate)
   Wend
   Call Cf.SetFieldValue(TempWidth + 4, "^", "ช่วง DUE", ">")
   Call m_Headers0.AddField(Cf)
   SaleWidth = SaleWidth + 4
   
   Call Cf.SetFieldValue(2, "^", "", "^")
   Call m_Headers0.AddField(Cf)
   SaleWidth = SaleWidth + 2
   
   If mcolParam("SUMMARY_MODE") = "Y" Or mcolParam("SALE_AND_CUSTOMER") = "Y" Then
      Call Cf.SetFieldValue(0, "^", "", "^")
      Call m_Headers0.AddField(Cf)
   Else
      Call Cf.SetFieldValue(2, "^", "เกิน", "^")
      Call m_Headers0.AddField(Cf)
      SaleWidth = SaleWidth + 2
   End If
   

   If mcolParam("SALE_AND_CUSTOMER") = "Y" Then
      Call Cf.SetFieldValue(2, "^", "เช็ครับ", "^")
      Call m_Headers0.AddField(Cf)
    Else
      Call Cf.SetFieldValue(2, "^", "หมายเหตุ", ">")
      Call m_Headers0.AddField(Cf)
    End If
    SaleWidth = SaleWidth + 2
   
   Call Cf.SetFieldValue(2, "^", "", ">")
   Call m_Headers0.AddField(Cf)
   SaleWidth = SaleWidth + 2
   
   '===
   collum = 0
   If mcolParam("SUMMARY_MODE") = "Y" Or mcolParam("SALE_AND_CUSTOMER") = "Y" Then
      Call Cf.SetFieldValue(0, "^", "NO.", "^")
      Call m_Headers1.AddField(Cf)
      collum = collum + 1
   Else
      Call Cf.SetFieldValue(1, "^", "NO.", "^")
      Call m_Headers1.AddField(Cf)
      collum = collum + 1
   End If
   numj = numj + 1
   
   If mcolParam("SUMMARY_MODE") = "Y" Then
      Call Cf.SetFieldValue(2, "<", "รหัส", "<")
      Call m_Headers1.AddField(Cf)
      Call Cf.SetFieldValue(4, "<", "พนักงานขาย", "<")
      Call m_Headers1.AddField(Cf)
      collum = collum + 2
   ElseIf mcolParam("SALE_AND_CUSTOMER") = "Y" Then
      Call Cf.SetFieldValue(2, "<", "รหัส", "<")
      Call m_Headers1.AddField(Cf)
      Call Cf.SetFieldValue(6, "<", "ลูกค้า", "<")
      Call m_Headers1.AddField(Cf)
      collum = collum + 2
   Else
      Call Cf.SetFieldValue(2, "^", "ใบส่งของ", "^")
      Call m_Headers1.AddField(Cf)
      Call Cf.SetFieldValue(2, "^", "วันที่ขาย", "^")
      Call m_Headers1.AddField(Cf)
      collum = collum + 2
   End If
   numj = numj + 2
            
   If mcolParam("SUMMARY_MODE") = "Y" Or mcolParam("SALE_AND_CUSTOMER") = "Y" Then
      Call Cf.SetFieldValue(0, "^", "", "^")
      Call m_Headers1.AddField(Cf)
      collum = collum + 1
   Else
      Call Cf.SetFieldValue(2, "^", "ครบกำหนด", "^")
      Call m_Headers1.AddField(Cf)
      collum = collum + 1
   End If
   numj = numj + 1
   
   FromDate = mcolParam("FROM_MONTH")
   ToDate = mcolParam("TO_MONTH")
   Call GetFirstLastDate(mcolParam("FROM_MONTH"), FromDate, TempDate)
   Call GetFirstLastDate(mcolParam("TO_MONTH"), TempDate, ToDate)
      
   Call Cf.SetFieldValue(2, "^", "< " & Format(Month(FromDate), "00") & "/" & Format(Year(FromDate) + 543, "00"), ">")
   Call m_Headers1.AddField(Cf)
    numj = numj + 1
    collum = collum + 1
   
   While (FromDate < ToDate)
      Call Cf.SetFieldValue(2, "^", Format(Month(FromDate), "00") & "/" & Format(Year(FromDate) + 543, "00"), ">")
      Call m_Headers1.AddField(Cf)
      FromDate = DateAdd("M", 1, FromDate)
      collum = collum + 1
      numj = numj + 1
   Wend
    
    FromDate = DateAdd("M", -1, FromDate)
    Call Cf.SetFieldValue(2, "^", Format(Month(FromDate), "00") & "/" & Format(Year(FromDate) + 543, "00") & ">", ">")
   Call m_Headers1.AddField(Cf)
   collum = collum + 1
   numj = numj + 1
    
' If DateToStringExtEx2(mcolParam("TO_PAY_DATE")) <> "" Then
   Call Cf.SetFieldValue(2, "^", "รวมหนี้ค้าง", ">")
   Call m_Headers1.AddField(Cf)
   collum = collum + 1
   numj = numj + 1
   
   If mcolParam("SUMMARY_MODE") = "Y" Or mcolParam("SALE_AND_CUSTOMER") = "Y" Then
      Call Cf.SetFieldValue(0, "^", "(วัน)", "^")
      Call m_Headers1.AddField(Cf)
   Else
      Call Cf.SetFieldValue(2, "^", "", "^")
      Call m_Headers1.AddField(Cf)
   End If
    numj = numj + 1
   
   If mcolParam("SALE_AND_CUSTOMER") = "Y" Then
       Call Cf.SetFieldValue(2, "^", "ล่วงหน้า", ">")
       Call m_Headers1.AddField(Cf)
   Else
       Call Cf.SetFieldValue(2, "^", "", ">")
       Call m_Headers1.AddField(Cf)
   End If
    numj = numj + 1
   
   Call Cf.SetFieldValue(2, "^", "วงเงิน", ">")
   Call m_Headers1.AddField(Cf)
   
   Set Cf = Nothing
            
  Call m_Headers0.GetString(1, TempStr1, TempStr1_1)
   strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
   strPrint = TempStr1_1
   Call Vsp.AddTable(strFormat, "", strPrint)
            
  Call m_Headers1.GetString(1, TempStr1, TempStr1_1)
   strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
   strPrint = TempStr1_1
   Call Vsp.AddTable(strFormat, "", strPrint)
            
   Set TempRs = Nothing
   
   Vsp.TableBorder = tbdT
   Vsp.FontBold = blnBold
   Vsp.FontSize = iSize
End Sub

Private Function ArrangeChq(AllCheck As Collection, CheckCancel As Collection, BkReceiveAmounts As Collection)
Dim ITEM As CBkTrn
Dim Ei As CBkTrn
Dim TempCCheck As CCheckCancel
    Set Ei = New CBkTrn
     'debug.print
     For Each ITEM In AllCheck
     Set TempCCheck = GetCheckCancelitem(CheckCancel, ITEM.CHQNUM, False)
                  If TempCCheck Is Nothing Then
                     ITEM.Flag = "N"
                      Set Ei = GetBkTrn(BkReceiveAmounts, Trim(ITEM.CUSCOD & "-" & ITEM.SLMCOD), False)
                      If Ei Is Nothing Then
                          Call BkReceiveAmounts.Add(ITEM, Trim(ITEM.CUSCOD & "-" & ITEM.SLMCOD))                       ' เมื่อไม่เจอว่าเคยมีอยู่ใน m_BkReceiveAmounts
                      Else
                          Ei.AMOUNT = Ei.AMOUNT + ITEM.AMOUNT                                                                           ' เจอว่าเคยมีแล้ว
                      End If
                Else
                    'debug.print
                End If
     Next ITEM
                     
End Function

Private Function initDoc() As Boolean
Dim strDate As String
Dim TempStr1 As String
Dim TempStr2 As String
Dim TempStr3 As String
Dim FromDate As String
Dim ToDate As String
Dim ToPaidDate As String
Dim iCount As Long

   mstrErrMsg = ""
   mblnHeader = True
   mblnNewPage = True
   mblnEndPage = True
   Vsp.PaperSize = pprA4 '
   Vsp.ORIENTATION = orLandscape
   Vsp.MarginBottom = 700
   Vsp.MarginFooter = 700
   Vsp.MarginHeader = 700
   Vsp.MarginLeft = 700
   Vsp.MarginRight = 700
   Vsp.MarginTop = 700
   Vsp.FontName = "AngsanaUPC"
   Vsp.FontSize = DETAIL_SIZE
   
   Call LoadReceiveAmountByBill(Nothing, m_ReceiveAmounts, -1, mcolParam("TO_PAY_DATE"))
     
   Call LoadBktChqAmountByCus(Nothing, m_BktAmountForCheckInCancel, mcolParam("FROM_CUSTOMER_CODE"), mcolParam("TO_CUSTOMER_CODE"), mcolParam("FROM_SALE_CODE"), mcolParam("TO_SALE_CODE"), DateAdd("D", 1, mcolParam("TO_DOC_DATE")))                 'สำหรับไปเทียบกับ เช็คยกเลิก
   Call TempCheckCancel.QueryData(m_Rs, iCount)                                      ' สำหรับเช็คยกเลิก
   Call LoadCheckCancel(m_ReceiveCancelitem, TempCheckCancel.CHECK_NO, m_Rs)
   
   Call ArrangeChq(m_BktAmountForCheckInCancel, m_ReceiveCancelitem, m_BkReceiveAmounts)
   
   Call LoadReceiveAmountByBill(Nothing, m_ReceiveAllAmounts, -1, -1)
   Call LoadARCNAmountByBill(Nothing, m_CnAllAmounts, DateAdd("D", 1, mcolParam("TO_DOC_DATE")), -1, -1, mcolParam("TO_DOC_DATE"))         'ต้องบวกยอดยกมาเพิ่ม ยอด CN ทั้งหมดของบิลหลัง CN แต่ เอกสาร LINK เป็นในช่วงยกมา
   
   FromDate = "จากวันที่ " & EmptyToString(DateToStringExtEx2(mcolParam("FROM_DOC_DATE")), "N/A")
   ToDate = " ถึง วันที่ " & EmptyToString(DateToStringExtEx2(mcolParam("TO_DOC_DATE")), "N/A")
   ToPaidDate = "    วันที่ชำระเงินถึงวันที่ " & EmptyToString(DateToStringExtEx2(mcolParam("TO_PAY_DATE")), "N/A") & "    วันที่พิมพ์ " & EmptyToString(DateToStringExtEx2(mcolParam("PRINT_DATE")), "N/A")
   
   MODULE_DESC = glbCompanyName & vbCrLf & mcolParam("REPORT_NAME") & _
                                       vbCrLf & FromDate & ToDate & "  " & ToPaidDate & " จากลูกค้า " & EmptyToString(mcolParam("FROM_CUSTOMER_CODE"), "N/A") & " ถึงลูกค้า " & EmptyToString(mcolParam("TO_CUSTOMER_CODE"), "N/A")
   
   Call SetReportConfig(Vsp, mcolParam("REPORT_KEY"))
   
   mdblWid = Vsp.PageWidth - Vsp.MarginLeft - Vsp.MarginRight
   If Not glbDatabaseMngr.GetServerDateTime(strDate, glbErrorLog) Then
      mstrErrMsg = "Error GetDateTime Error."
      Exit Function
   End If
   mdteDate = InternalDateToDate(strDate)
   initDoc = True
End Function

Private Function GetItemFromListIndex(TempCol As Collection, Ind As Long) As Object
   Set GetItemFromListIndex = TempCol(Ind)
End Function

Private Sub GenerateDetail(alngX() As Long, i As Long, ARt As CARTrn, Tot1() As Double, Tot2() As Double, Tot3() As Double, Optional HaveHeader As Boolean = False, Optional PrevKey1 As String, Optional PrevKey4 As String, Optional PrevCrline As Double)
Dim HeadCf As CReportField
Dim TempStr As String
Dim TempStr1 As String
Dim TempStr2 As String
Dim BodyCf As CReportField
Dim strFormat As String
Dim Ari As CARRcIt
Dim Ari2 As CARRcIt
Dim Apt1 As CARTrn
Dim ARt2 As CARTrn
Dim Ug As CBkTrn
Dim DueCount As Long
Dim MM As CMaxMin
Dim j As Long
Dim Sum As Double
Dim PrintDate As Date
Dim PaidAmount As Double
Dim PaidBalance As Double 'ชำระแล้วยกมา
Dim TempAmt As Double
Dim TempRealCredit As CRealCredit
Dim Width  As Double
Dim FromDate As Date
Dim ToDate As Date
Dim TempDate As Date
Dim TempFromdate As Date
Dim TempToDate As Date
Dim NewDueDate As Date
Dim SumBkReceiveAmounts As Double

Dim TempCCheck As CCheckCancel
Dim FoundInCheckCancel  As Boolean

Dim saleSum As Double
Dim totSum As Double
   
   Set BodyCf = New CReportField
   
   If mcolParam("PRINT_DATE") <= 0 Then
      PrintDate = Now
   Else
      PrintDate = mcolParam("PRINT_DATE")
   End If
   NewDueDate = DateAdd("D", mcolParam("CREDIT_DAY"), ARt.DOCDAT)
   DueCount = DateDiff("D", NewDueDate, PrintDate)
   
   Set Ari = GetARRcpItem(m_ReceiveAmounts, ARt.DOCNUM)
   Set Ari2 = GetARRcpItemEx(m_ReceiveAllAmounts, ARt.DOCNUM)
   
   If (ARt.CUSCOD = "45-005") Then
    'debug.print
   End If

   Set Ug = GetBkTrn(m_BkReceiveAmounts, ARt.CUSCOD & "-" & ARt.SLMCOD)              'ดึงล่วงหน้าที่ ลูกหนี้คนเดียวกัน
   Set ARt2 = GetARTrn(m_CnAllAmounts, ARt.DOCNUM)
   
   PaidBalance = ARt.RCVAMT - Ari2.RCVAMT - ARt2.AMOUNT
      
   '*********
      If ARt.CUSCOD <> PrevKey1 Then
         Tot1(numj) = Tot1(numj) + Ug.AMOUNT
         Tot2(numj) = Tot2(numj) + Ug.AMOUNT
     End If
      
      '-----------------------------      'แสดงหัวลูกค้า
      If (HaveHeader) And PrevKey1 <> Trim(ARt.CUSCOD) And (mcolParam("INCLUDE_CUSTOMER_MODE") = "Y") And mcolParam("SUMMARY_MODE") = "N" Then
         
         Call GenerateFooter("รวม ", PrevKey1, "", Tot3, m_Details1, 0, Val(PrevCrline))  'รวมตามลูกค้า
         Call m_Details1.GetString(2, TempStr1, TempStr2)
         strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
         Vsp.FontBold = True
        If Tot3(collum) <> 0 Or Tot3(numj) <> 0 Then
               Call Vsp.AddTable(strFormat, "", TempStr2)
       End If
         Vsp.FontBold = False
         
         HaveHeader = False
         For j = 1 To UBound(Tot3)
            Tot3(j) = 0
         Next j
      ElseIf (HaveHeader) And mcolParam("SALE_AND_CUSTOMER") = "Y" And PrevKey1 <> Trim(ARt.CUSCOD) Then
         Call GenerateFooter("", PrevKey1, PrevKey4, Tot3, m_Details1, 0, Val(PrevCrline))  'รวมตามลูกค้า   '1 รหัส 4 ชื่อ
         Call m_Details1.GetString(2, TempStr1, TempStr2)
         strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
         If Tot3(collum) <> 0 Or Tot3(numj) <> 0 Then
               Call Vsp.AddTable(strFormat, "", TempStr2)
         End If
         Vsp.FontBold = False
         
         HaveHeader = False
         For j = 1 To UBound(Tot3)
            Tot3(j) = 0
         Next j
      End If
         
      If Not (HaveHeader) And PrevKey1 <> Trim(ARt.CUSCOD) And (mcolParam("INCLUDE_CUSTOMER_MODE") = "Y") And mcolParam("SUMMARY_MODE") = "N" Then
         Call m_Details1.ClearField
         Width = m_Headers0.Fields(1).Width + m_Headers0.Fields(2).Width + m_Headers0.Fields(3).Width
         TempStr = ARt.CUSCOD
         Call BodyCf.SetFieldValue(Width, "^", TempStr)
         Call m_Details1.AddField(BodyCf)
         
         Width = SaleWidth - m_Headers0.Fields(3).Width
         TempStr = ARt.CUSNAM
         Call BodyCf.SetFieldValue(Width, "<", TempStr)
         Call m_Details1.AddField(BodyCf)
   
         Call m_Details1.GetString(1, TempStr1, TempStr2)
         strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
         Call Vsp.AddTable(strFormat, "", TempStr2)
         HaveHeader = True
      ElseIf Not (HaveHeader) And PrevKey1 <> Trim(ARt.CUSCOD) And (mcolParam("SALE_AND_CUSTOMER") = "Y") And mcolParam("SUMMARY_MODE") = "N" Then
         HaveHeader = True
      End If
      PrevKey1 = Trim(ARt.CUSCOD)
      PrevKey4 = Trim(ARt.CUSNAM)
      '-----------------------------      'แสดงหัวลูกค้า
      
      PrevCrline = ARt.CRLINE
      
      Call m_Details1.ClearField
      Set HeadCf = m_Headers1.Fields(1)
      TempStr = i
      Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
      Call m_Details1.AddField(BodyCf)
   
      Set HeadCf = m_Headers1.Fields(2)
      TempStr = ARt.DOCNUM
      Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
      Call m_Details1.AddField(BodyCf)

      Set HeadCf = m_Headers1.Fields(3)
      TempStr = DateToStringExtEx2(ARt.DOCDAT)
      Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
      Call m_Details1.AddField(BodyCf)
      
      Set HeadCf = m_Headers1.Fields(4)
      TempStr = DateToStringExtEx2(NewDueDate)
      Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
      Call m_Details1.AddField(BodyCf)
      
      j = 5
      Sum = 0
      
      FromDate = mcolParam("FROM_MONTH")
      ToDate = mcolParam("TO_MONTH")
      Call GetFirstLastDate(mcolParam("FROM_MONTH"), FromDate, TempDate)
      Call GetFirstLastDate(mcolParam("TO_MONTH"), TempDate, ToDate)
         
      If (NewDueDate < FromDate) Then
         TempAmt = (ARt.AMOUNT) - (Ari.RCVAMT + PaidBalance)
         If ARt.RECTYP = 5 Then 'ลดหนี้
            TempAmt = -1 * TempAmt
         End If
         TempStr = FormatNumber(TempAmt, "-")
         
         Tot1(j) = Tot1(j) + TempAmt                 '1
         Tot2(j) = Tot2(j) + TempAmt
         Tot3(j) = Tot3(j) + TempAmt
         Sum = Sum + TempAmt
                   Tot3(numj) = Ug.AMOUNT  'รวมล่วงหน้า    3=ลูกค้า 1=พนักงานขาย 2สุทธิ
      Else
         TempStr = FormatNumber(0, "-")
      End If
      
      Set HeadCf = m_Headers1.Fields(j)
      Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
      Call m_Details1.AddField(BodyCf)
      
      While (FromDate < ToDate)
      j = j + 1
      Call GetFirstLastDate(FromDate, TempFromdate, TempToDate)
      If (TempFromdate <= NewDueDate) And (NewDueDate <= TempToDate) Then
         TempAmt = (ARt.AMOUNT) - (Ari.RCVAMT + PaidBalance)
         If ARt.RECTYP = 5 Then 'ลดหนี้
            TempAmt = -1 * TempAmt
         End If
         TempStr = FormatNumber(TempAmt, "-")
         
         Tot1(j) = Tot1(j) + TempAmt                         '2
         Tot2(j) = Tot2(j) + TempAmt
         Tot3(j) = Tot3(j) + TempAmt
         Sum = Sum + TempAmt
         Tot3(numj) = Ug.AMOUNT  'รวมล่วงหน้า    3=ลูกค้า 1=พนักงานขาย 2สุทธิ
      Else
         TempStr = FormatNumber(0, "-")
      End If
         
         Set HeadCf = m_Headers1.Fields(j)
         Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
         Call m_Details1.AddField(BodyCf)
         
         FromDate = DateAdd("M", 1, FromDate)
      Wend
      
      j = j + 1
            If (NewDueDate > ToDate) Then
               TempAmt = (ARt.AMOUNT) - (Ari.RCVAMT + PaidBalance)
               If ARt.RECTYP = 5 Then 'ลดหนี้
                  TempAmt = -1 * TempAmt
               End If
               TempStr = FormatNumber(TempAmt, "-")
               
               Tot1(j) = Tot1(j) + TempAmt                  '3
               Tot2(j) = Tot2(j) + TempAmt
               Tot3(j) = Tot3(j) + TempAmt
               Sum = Sum + TempAmt
                         Tot3(numj) = Ug.AMOUNT   'รวมล่วงหน้า    3=ลูกค้า 1=พนักงานขาย 2สุทธิ
            Else
               TempStr = FormatNumber(0, "-")
            End If
      Set HeadCf = m_Headers1.Fields(j)
      Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
      Call m_Details1.AddField(BodyCf)
         
      j = j + 1
      Set HeadCf = m_Headers1.Fields(j)
      TempStr = FormatNumber(Sum, "-")
      Tot1(j) = Tot1(j) + TempAmt                             '4
      Tot2(j) = Tot2(j) + TempAmt
      Tot3(j) = Tot3(j) + TempAmt
                Tot3(numj) = Ug.AMOUNT 'รวมล่วงหน้า    3=ลูกค้า 1=พนักงานขาย 2สุทธิ

      Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
      Call m_Details1.AddField(BodyCf)

      j = j + 1
      Set HeadCf = m_Headers1.Fields(j)
      TempStr = FormatNumberInt(DueCount)
      Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
      Call m_Details1.AddField(BodyCf)
      
     j = j + 1                      ' ตรงนี้ที่ตอนแรกทำผิด
      Set HeadCf = m_Headers1.Fields(j)
      TempStr = ""
      Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
      Call m_Details1.AddField(BodyCf)

      j = j + 1                      ' ตรงนี้ที่ตอนแรกทำผิด
      Set HeadCf = m_Headers1.Fields(j)
      TempStr = ""
      Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
      Call m_Details1.AddField(BodyCf)
      
     If mcolParam("SALE_AND_CUSTOMER") = "N" Then
         Tot1(numj) = 0
          Tot2(numj) = 0
          Tot3(numj) = 0
     End If

                    If mcolParam("SUMMARY_MODE") = "N" Then
                          If mcolParam("SALE_AND_CUSTOMER") = "N" Then
                          Call m_Details1.GetString(1, TempStr1, TempStr2)
                          strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
                          Call Vsp.AddTable(strFormat, "", TempStr2)
                       End If
                    End If
'   Else
'      i = i - 1
'   End If
   Set BodyCf = Nothing
End Sub

Private Function genDoc(isPreview As Boolean) As Boolean
On Error GoTo ErrHandler
Dim RName As String
Dim i As Long
Dim j As Long
Dim strFormat As String
Dim alngX() As Long
Dim IsOK As Boolean
Dim Amt As Double
Dim HeadCf As CReportField
Dim BodyCf As CReportField
Dim TempStr1 As String
Dim TempStr2 As String
Dim Total1(100) As Double
Dim Total2(100) As Double
Dim Total3(100) As Double
Dim iCount As Long
Dim TempStr As String
Dim PrevKey1 As String
Dim PrevKey2 As String
Dim PrevKey3 As String
Dim PrevKey4 As String

Dim PrevTot1 As Double
Dim PrevTot2 As Double
Dim St As CIsTab
Dim Ar As COESLM
Dim Width As Double
Dim ARt As CARTrn
Dim TempRs As ADODB.Recordset
Dim HaveHeader As Boolean
Dim Ug1_15_1 As CBkTrn

Dim PrevKey5 As String
Dim PrevCrline As Double
Dim saleSum As Double
Dim totSum As Double
   RName = "genDoc"
'-----------------------------------------------------------------------------------------------------
'                                             Query Here
'-----------------------------------------------------------------------------------------------------
   
   Set Rs = New ADODB.Recordset
   Set TempRs = New ADODB.Recordset
   
  Set m_Rs = New ADODB.Recordset
  Set TempCheckCancel = New CCheckCancel
  Set m_ReceiveCancelitem = New Collection
  Set TempCCheck = New CCheckCancel
   Set m_BktAmountForCheckInCancel = New Collection
   
   If Not initDoc Then Exit Function
   
   Vsp.Preview = isPreview
'-----------------------------------------------------------------------------------------------------
'                                         Main Operation Here
'-----------------------------------------------------------------------------------------------------
   
   For j = 1 To UBound(Total1)
      Total1(j) = 0
      Total2(j) = 0
      Total3(j) = 0
   Next j
   
   Call StartExportFile(Vsp)
   Vsp.StartDoc
      i = 0
      Set BodyCf = New CReportField
      
      Set Ar = New COESLM
      Ar.FROM_SALE_CODE = mcolParam("FROM_SALE_CODE")
      Ar.TO_SALE_CODE = mcolParam("TO_SALE_CODE")
      Ar.OrderBy = mcolParam("ORDER_TYPE")
      Ar.OrderType = mcolParam("ORDER_BY")
      Call glbDaily.QuerySale(Ar, Rs, iCount, IsOK, glbErrorLog)
      
          totSum = 0

      While Not Rs.EOF
            
               saleSum = 0

'         For j = 1 To UBound(Total1)
'            Total1(j) = 0
'         Next j

         Call m_Details1.ClearField
         Call Ar.PopulateFromRS(1, Rs)
         
  
         If True Then  'มียอดค้างจ่าย
            '===Head
            If mcolParam("SUMMARY_MODE") = "N" Then
               Width = m_Headers0.Fields(1).Width + m_Headers0.Fields(2).Width
               TempStr = Ar.SLMCOD
               Call BodyCf.SetFieldValue(Width, "^", TempStr)
               Call m_Details1.AddField(BodyCf)
   
               Width = SaleWidth
               TempStr = Ar.SLMNAM
               Call BodyCf.SetFieldValue(Width, "<", TempStr)
               Call m_Details1.AddField(BodyCf)
   
               Call m_Details1.GetString(1, TempStr1, TempStr2)
               strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
               Call Vsp.AddTable(strFormat, "", TempStr2)
            End If
            PrevKey2 = Ar.SLMCOD
            PrevKey3 = Ar.SLMNAM
            
            '===

            '=== Detail
            Set ARt = New CARTrn
            ARt.SLMCOD = Ar.SLMCOD
            ARt.FROM_DOC_DATE = mcolParam("FROM_DOC_DATE")
            ARt.TO_DOC_DATE = mcolParam("TO_DOC_DATE")
            
            ARt.FROM_CUSTOMER_CODE = mcolParam("FROM_CUSTOMER_CODE")
            ARt.TO_CUSTOMER_CODE = mcolParam("TO_CUSTOMER_CODE")
            
            ARt.RECTYP = ""
            ARt.RecTypeSet = "('3', '4', '5' )"         '
            
         '  ARt.FROM_GETDAT = DateAdd("D", 1, mcolParam("TO_DOC_DATE"))
    
            If mcolParam("INCLUDE_CUSTOMER_MODE") = "Y" Or mcolParam("SALE_AND_CUSTOMER") = "Y" Or mcolParam("SUMMARY_MODE") = "Y" Then
               ARt.OrderBy = 3
            Else
               ARt.OrderBy = 1
            End If
            ARt.OrderType = 1
            Call ARt.QueryData(1, TempRs, iCount)
            
            HaveHeader = False
            
            While Not TempRs.EOF
               i = i + 1
               Call ARt.PopulateFromRS(1, TempRs)    ' อันใหม่เป็น 12
               'debug.print (ARt.CUSCOD)
                    
               Set Ug1_15_1 = GetBkTrn(m_BkReceiveAmounts, ARt.CUSCOD & "-" & Ar.SLMCOD, False)
               If Not (Ug1_15_1 Is Nothing) Then
                  Ug1_15_1.Flag = "Y"   'มีอยู่ใน 1_15_1
               End If
               
               Call GenerateDetail(alngX(), i, ARt, Total1, Total2, Total3, HaveHeader, PrevKey1, PrevKey4, PrevCrline)
               TempRs.MoveNext
            Wend
            
            If HaveHeader And mcolParam("INCLUDE_CUSTOMER_MODE") = "Y" And mcolParam("SUMMARY_MODE") = "N" Then
               Call GenerateFooter("รวม ", PrevKey1, "", Total3, m_Details1, 0, Val(PrevCrline))  'รวมตามลูกค้า
               Call m_Details1.GetString(2, TempStr1, TempStr2)
               strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
               Vsp.FontBold = True
               If Total3(collum) <> 0 Or Total3(numj) <> 0 Then
                     Call Vsp.AddTable(strFormat, "", TempStr2)
             End If
               Vsp.FontBold = False
               
               HaveHeader = False
               For j = 1 To UBound(Total3)
                  Total3(j) = 0
               Next j
            ElseIf HaveHeader And mcolParam("SALE_AND_CUSTOMER") = "Y" Then
               Call GenerateFooter("", PrevKey1, PrevKey4, Total3, m_Details1, 0, Val(PrevCrline))  'รวมตามลูกค้า
               Call m_Details1.GetString(2, TempStr1, TempStr2)
               strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
             If Total3(collum) <> 0 Or Total3(numj) <> 0 Then
                   Call Vsp.AddTable(strFormat, "", TempStr2)
              End If
               Vsp.FontBold = False
               
               HaveHeader = False
               For j = 1 To UBound(Total3)
                  Total3(j) = 0
               Next j
            End If
            
            Set ARt = Nothing
            
       If mcolParam("SALE_AND_CUSTOMER") = "Y" Then
                For Each Ug1_15_1 In m_BkReceiveAmounts
                     If Ug1_15_1.Flag = "N" And Ug1_15_1.SLMCOD = Ar.SLMCOD Then
                            Call Generate1_15_1(Ug1_15_1.CUSCOD, Ug1_15_1.CUSNAM, Ug1_15_1.AMOUNT, m_Details1)
                            Call m_Details1.GetString(2, TempStr1, TempStr2)
                            strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
                              Ug1_15_1.Flag = "Y"
                            Call Vsp.AddTable(strFormat, "", TempStr2)
                            Total1(numj) = Total1(numj) + Ug1_15_1.AMOUNT
                            Total2(numj) = Total2(numj) + Ug1_15_1.AMOUNT
                     End If
                Next Ug1_15_1
            End If
            
            If mcolParam("SUMMARY_MODE") = "Y" Then
               Call GenerateFooter("", PrevKey2, PrevKey3, Total1, m_Details1, 0, 0)  'รวมตามพนักงานขาย
               Call m_Details1.GetString(2, TempStr1, TempStr2)
               strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
               Call Vsp.AddTable(strFormat, "", TempStr2)
            Else
               Call GenerateFooter("รวม", PrevKey2, "", Total1, m_Details1, 0, 0)  'รวมตามพนักงานขาย
               Call m_Details1.GetString(2, TempStr1, TempStr2)
               strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
               Vsp.FontBold = True
               Call Vsp.AddTable(strFormat, "", TempStr2)
               Vsp.FontBold = False
            End If
'              If PrevKey2 <> Ar.SLMCOD Then
               For j = 1 To UBound(Total3)
                  Total1(j) = 0
               Next j
 '           End If
            '===
          End If

         Rs.MoveNext
      Wend
      
         If (mcolParam("SALE_AND_CUSTOMER") = "Y" And mcolParam("FROM_SALE_CODE") = "") Or mcolParam("TO_SALE_CODE") = "" Then
         
               Call m_Details1.ClearField
               Width = m_Headers0.Fields(1).Width + m_Headers0.Fields(2).Width
               TempStr = "( )"
               Call BodyCf.SetFieldValue(Width, "^", TempStr)
               Call m_Details1.AddField(BodyCf)
   
               Width = SaleWidth
               TempStr = "ไม่มีรหัส"
               Call BodyCf.SetFieldValue(Width, "<", TempStr)
               Call m_Details1.AddField(BodyCf)
   
               Call m_Details1.GetString(1, TempStr1, TempStr2)
               strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
               Call Vsp.AddTable(strFormat, "", TempStr2)
     
         
                For Each Ug1_15_1 In m_BkReceiveAmounts
                     If Ug1_15_1.Flag = "N" And Ug1_15_1.SLMCOD = "" Then
                            Call Generate1_15_1(Ug1_15_1.CUSCOD, Ug1_15_1.CUSNAM, Ug1_15_1.AMOUNT, m_Details1)
                            Call m_Details1.GetString(2, TempStr1, TempStr2)
                            strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
                              Ug1_15_1.Flag = "Y"
                            Call Vsp.AddTable(strFormat, "", TempStr2)
                            Total1(numj) = Total1(numj) + Ug1_15_1.AMOUNT
                            Total2(numj) = Total2(numj) + Ug1_15_1.AMOUNT
                     End If
                Next Ug1_15_1
                
                Call GenerateFooter("", "( )", "", Total1, m_Details1, 0, 0)  'รวมตามพนักงานขาย
               Call m_Details1.GetString(2, TempStr1, TempStr2)
               strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
               Vsp.FontBold = True
               Call Vsp.AddTable(strFormat, "", TempStr2)
               Vsp.FontBold = False
                
            End If
      
      If mcolParam("SUMMARY_MODE") = "Y" Then
         Call GenerateFooter("", "สุทธิ", "", Total2, m_Details1, 0, 0)
      Else
         Call GenerateFooter("สุทธิ", "", "", Total2, m_Details1, 0, 0)
      End If
      Call m_Details1.GetString(2, TempStr1, TempStr2)
      strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
      Vsp.FontBold = True
      Call Vsp.AddTable(strFormat, "", TempStr2)
      Vsp.FontBold = False
      
      Set Ar = Nothing
      Set BodyCf = Nothing
   Vsp.EndDoc
   Call CloseExportFile(Vsp)
   
   If TempRs.State = adStateOpen Then
      TempRs.Close
   End If
   Set TempRs = Nothing
   
   If Rs.State = adStateOpen Then
      Rs.Close
   End If
   Set Rs = Nothing
   
   genDoc = True
   Exit Function
   
ErrHandler:
   mstrErrMsg = "Error(" & RName & ")" & Err.Number & " : " & Err.DESCRIPTION
   Set Rs = Nothing
End Function


Private Sub Generate1_15_1(Txt1 As String, Txt2 As String, Tot As Double, Details As CFieldList)
Dim HeadCf As CReportField
Dim BodyCf As CReportField
Dim j As Long
Dim TempStr As String
        
   Call Details.ClearField
   Set BodyCf = New CReportField
   
   j = 0
   For Each HeadCf In m_Headers1.Fields
      j = j + 1
      If j = 2 Then
         TempStr = Txt1
      ElseIf j = 3 Then
         TempStr = Txt2
      ElseIf j = numj Then
         TempStr = FormatNumber(Tot, "")
      Else
         TempStr = "-"
      End If
      
      Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.Align, TempStr, HeadCf.BodyAlign)
      Call Details.AddField(BodyCf)
   Next HeadCf
      
   Set BodyCf = Nothing
End Sub

Private Sub GenerateFooter(Txt1 As String, Txt2 As String, Txt3 As String, Tot() As Double, Details As CFieldList, DateCount As Long, CreditLine As Double)
Dim HeadCf As CReportField
Dim BodyCf As CReportField
Dim j As Long
Dim TempStr As String

   Call Details.ClearField
   Set BodyCf = New CReportField
   
   j = 0
   For Each HeadCf In m_Headers1.Fields
      j = j + 1
      If j = 1 Then
         TempStr = Txt1
      ElseIf j = 2 Then
         TempStr = Txt2
      ElseIf j = 3 Then
         TempStr = Txt3
      ElseIf j = (m_Headers1.Fields.Count - 1) Or (j = (m_Headers1.Fields.Count - 2) And (mcolParam("SUMMARY_MODE") = "N")) Then
         TempStr = FormatNumber(Tot(j), "")
      ElseIf j = (m_Headers1.Fields.Count) Then
         TempStr = FormatNumber(CreditLine, "")
      Else
         TempStr = FormatNumber(Tot(j), "-")
      End If
      
      Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.Align, TempStr, HeadCf.BodyAlign)
      Call Details.AddField(BodyCf)
   Next HeadCf
      
   Set BodyCf = Nothing
End Sub

Private Sub VSP_EndDoc()
'This event occur when VSPrinter.EndDoc is used
End Sub

Private Sub VSP_EndPage()
   If Not mblnEndPage Then Exit Sub
'   Vsp.DrawLine Vsp.MarginLeft, mdY, Vsp.MarginLeft, Vsp.CurrentY
'   Vsp.DrawLine Vsp.PageWidth - Vsp.MarginRight, mdY, Vsp.PageWidth - Vsp.MarginRight, Vsp.CurrentY
'   Vsp.DrawLine Vsp.MarginLeft, Vsp.CurrentY, Vsp.PageWidth - Vsp.MarginRight, Vsp.CurrentY
End Sub

Private Sub VSP_Error()
'Error in runtime occur here press F1 in VSP.ErrorDescription to see more information
   mstrErrMsg = Vsp.ErrorDescription
End Sub

Private Sub VSP_NewPage()
Dim talnT As TextAlignSettings
Dim tbdT As TableBorderSettings
Dim blnBold As Boolean
Dim blnUnder As Boolean
Dim blnItalic As Boolean
Dim iSize As Integer
Dim sName As String
Dim strFormat As String
Dim dY(0 To 1) As Double
Dim alngX() As Long

   If Not mblnNewPage Then Exit Sub
   talnT = Vsp.TextAlign
   tbdT = Vsp.TableBorder
   blnBold = Vsp.FontBold
   blnUnder = Vsp.FontUnderline
   blnItalic = Vsp.FontItalic
   iSize = Vsp.FontSize
   sName = Vsp.FontName
   
   Vsp.FontSize = 8
   dY(0) = Vsp.CurrentY
   Vsp.TextAlign = taLeftBottom
   strFormat = VSP_CalTable(TITLE_FORMAT, mdblWid, alngX)
   Vsp.TableBorder = tbNone
   Call Vsp.AddTable(strFormat, "", "|หน้าที่|: " & Vsp.CurrentPage)
   Call Vsp.AddTable(strFormat, "", "|วัน|: " & DateToStringExt(mdteDate))
   Call Vsp.AddTable(strFormat, "", "|เวลา|: " & Right$(DateToStringExtEx(Now), 8))
   Vsp.CurrentY = dY(0)
   
   Vsp.FontSize = TITLE_SIZE
   Vsp.FontBold = True
   Vsp.TextAlign = taCenterBottom
   Vsp.Paragraph = MODULE_DESC
   Vsp.Paragraph = ""
   If mblnHeader Then
      Call printHeader
   End If
   mdY = Vsp.CurrentY

   Vsp.TextAlign = talnT
   Vsp.TableBorder = tbdT
   Vsp.FontBold = blnBold
   Vsp.FontUnderline = blnUnder
   Vsp.FontItalic = blnItalic
   Vsp.FontSize = iSize
   Vsp.FontName = sName
End Sub

Private Sub VSP_StartDoc()
'This event occur when VSPrinter.StartDoc is used and used to initialize some information before generating printed document
End Sub



