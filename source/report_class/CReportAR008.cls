VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CReportAR008"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'User Group Report
Implements CReportInterface

Private Const MODULE_NAME = "CReportAR008"
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

Private Const TITLE_SIZE = 12
Private Const HEADER_SIZE = 10
Private Const DETAIL_SIZE = 10

Private Const TITLE_FORMAT = "<32|<3|<5;"

Private m_TempCol As Collection
Private m_HeaderFormat1 As String
Private m_HeaderFormat2 As String
Private m_HeaderText1 As String
Private m_HeaderText2 As String

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

Private m_ArCusAmounts As Collection
Private m_ArCusAmountCns As Collection
Private m_ArCusAmountCashs As Collection

Private m_ArCusAmountTo As Collection
Private m_ArCusAmountToCn As Collection

Private m_ArCusPaids As Collection
Private m_ArCusPaidCns As Collection
Private m_ArCusPaidTo As Collection
Private m_ArCusPaidToCn As Collection

Private m_PaidAmounts As Collection
Private m_PaidAllAmounts As Collection
Private m_PaidAllAmountCns As Collection

Private m_MaxMins As Collection
Private m_DbnCdnByBills As Collection

Private m_5 As Collection
Private m_6 As Collection
Private m_7 As Collection
Private m_8 As Collection
Private m_9 As Collection
Private m_10 As Collection
Private m_11 As Collection

Private HaveValue As Boolean
Private HaveValueSum As Boolean
Private FirstDate As Date
Private LastDate As Date
Private m_DistinctMontnYear As Collection
Private m_SumMontnYear As Collection
Private m_SumMontnYearCN As Collection
Private m_ReceiveAmounts As Collection
Private m_ReceiveAmountCNs As Collection
Private m_ReceiveAllAmounts As Collection
Private m_ReceiveAllAmountCNs As Collection
Private m_CnAllAmounts As Collection
Private m_ArAmounts As Collection
Private m_ArAmountsCn As Collection
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
   Set m_PaidAmounts = New Collection
   Set m_MaxMins = New Collection
   Set m_ArCusAmounts = New Collection
   Set m_ArCusPaids = New Collection
   Set m_PaidAllAmounts = New Collection
   Set m_DbnCdnByBills = New Collection
   Set m_ArCusAmountTo = New Collection
   Set m_ArCusPaidTo = New Collection
   Set m_ArCusAmountCns = New Collection
   Set m_ArCusPaidCns = New Collection
   Set m_PaidAllAmountCns = New Collection
   Set m_ArCusAmountToCn = New Collection
   Set m_ArCusPaidToCn = New Collection
   Set m_ArCusAmountCashs = New Collection
   
   Set m_5 = New Collection
   Set m_6 = New Collection
   Set m_7 = New Collection
   Set m_8 = New Collection
   Set m_9 = New Collection
   Set m_10 = New Collection
   Set m_11 = New Collection
   
   Set m_DistinctMontnYear = New Collection
   Set m_SumMontnYear = New Collection
   Set m_SumMontnYearCN = New Collection
   Set m_ReceiveAmounts = New Collection
   Set m_ReceiveAmountCNs = New Collection
   
   Set m_ReceiveAllAmounts = New Collection
   Set m_ReceiveAllAmountCNs = New Collection
   Set m_CnAllAmounts = New Collection
   Set m_ArAmounts = New Collection
   Set m_ArAmountsCn = New Collection
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
   Set m_PaidAmounts = Nothing
   Set m_MaxMins = Nothing
   Set m_PaidAllAmounts = Nothing
   Set m_DbnCdnByBills = Nothing
   Set m_ArCusAmountTo = Nothing
   Set m_ArCusPaidTo = Nothing
   Set m_ArCusAmountCns = Nothing
   Set m_ArCusPaidCns = Nothing
   Set m_PaidAllAmountCns = Nothing
   Set m_ArCusAmountToCn = Nothing
   Set m_ArCusPaidToCn = Nothing
   Set m_ArCusAmountCashs = Nothing
   
   Set m_5 = Nothing
   Set m_6 = Nothing
   Set m_7 = Nothing
   Set m_8 = Nothing
   Set m_9 = Nothing
   Set m_10 = Nothing
   Set m_11 = Nothing
   
   Set m_DistinctMontnYear = Nothing
   Set m_SumMontnYear = Nothing
   Set m_SumMontnYearCN = Nothing
   Set m_ReceiveAmounts = Nothing
   Set m_ReceiveAmountCNs = Nothing
   
   Set m_ReceiveAllAmounts = Nothing
   Set m_ReceiveAllAmountCNs = Nothing
   Set m_CnAllAmounts = Nothing
   Set m_ArAmounts = Nothing
   Set m_ArAmountsCn = Nothing
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
Dim TempStr1 As String
Dim TempStr1_1  As String
Dim Cf As CReportField
Dim HeadCf As CReportField
Dim j As Long
Dim TempStr As String
Dim Width As Integer
Dim Acp As CARRcIt
Dim HaveYYYYMM As Boolean
   tbdT = Vsp.TableBorder
   blnBold = Vsp.FontBold
   iSize = Vsp.FontSize
   
   Vsp.FontSize = HEADER_SIZE
   Vsp.FontBold = True
   Vsp.TableBorder = tbAll
   
   Call m_Headers0.ClearField
   Call m_Headers1.ClearField
   
   Set Cf = New CReportField
   
   Call Cf.SetFieldValue(2, "^", "", "^")
   Call m_Headers1.AddField(Cf)
   
   HaveYYYYMM = False
   Width = 0
   For Each Acp In m_DistinctMontnYear
      If mcolParam("YYYYMM") <= (Year(Acp.DOCDAT) + 543 & Format(Month(Acp.DOCDAT), "00")) Then
         Width = Width + 2
         Call Cf.SetFieldValue(2, "^", Format(Month(Acp.DOCDAT), "00") & "/" & Year(Acp.DOCDAT) + 543, ">")
         Call m_Headers1.AddField(Cf)
      Else
         If (Not HaveYYYYMM) And (Len(mcolParam("YYYYMM")) > 0) Then
            Width = Width + 2
            Call Cf.SetFieldValue(2, "^", "ยกมา", ">")
            Call m_Headers1.AddField(Cf)
            HaveYYYYMM = True
         End If
      End If
   Next Acp
   Call Cf.SetFieldValue(3, "^", "รวม", ">")
   Call m_Headers1.AddField(Cf)
   
   Call Cf.SetFieldValue(2, "^", "เดือน", "^")
   Call m_Headers0.AddField(Cf)
   Call Cf.SetFieldValue(Width + 3, "^", "เดือน/ปี ของใบ INVOICE ที่รับชำระ", ">")
   Call m_Headers0.AddField(Cf)
   
   Set Cf = Nothing
   Call m_Headers0.GetString(1, TempStr1, TempStr1_1)
   strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
   strPrint = TempStr1_1
   Call Vsp.AddTable(strFormat, "", strPrint)
   
  Call m_Headers1.GetString(1, TempStr1, TempStr1_1)
   strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
   strPrint = TempStr1_1
   Call Vsp.AddTable(strFormat, "", strPrint)
            
   Vsp.TableBorder = tbdT
   Vsp.FontBold = blnBold
   Vsp.FontSize = iSize
End Sub

Private Function initDoc() As Boolean
Dim strDate As String
Dim TempStr1 As String
Dim TempStr2 As String
Dim TempStr3 As String
Dim FromDate As String
Dim ToDate As String
Dim TempFromdate As Date
Dim TempToDate As Date
   FirstDate = mcolParam("FROM_DOC_DATE")
   LastDate = mcolParam("TO_DOC_DATE")
   
   mstrErrMsg = ""
   mblnHeader = True
   mblnNewPage = True
   mblnEndPage = True
   Vsp.PaperSize = pprA4 '
   Vsp.ORIENTATION = orLandscape
   Vsp.MarginBottom = 300
   Vsp.MarginFooter = 300
   Vsp.MarginHeader = 300
   Vsp.MarginLeft = 300
   Vsp.MarginRight = 500
   Vsp.MarginTop = 300
   Vsp.FontName = "AngsanaUPC"
   Vsp.FontSize = DETAIL_SIZE
      
   TempFromdate = DateSerial(Left(mcolParam("YYYYMM"), 4) - 543, Right(mcolParam("YYYYMM"), 2), 1)
   
'   '-------------------------------- ยกมา --------------------------------
   
   Call LoadReceiveAmount(Nothing, m_ReceiveAmounts, -1, DateAdd("D", -1, TempFromdate), -1, -1, "('3','4')", mcolParam("FROM_CUSTOMER_CODE"), mcolParam("TO_CUSTOMER_CODE"))
   Call LoadReceiveAmount(Nothing, m_ReceiveAmountCNs, -1, DateAdd("D", -1, TempFromdate), -1, -1, "('5')", mcolParam("FROM_CUSTOMER_CODE"), mcolParam("TO_CUSTOMER_CODE"))
   
   Call LoadReceiveAmount(Nothing, m_ReceiveAllAmounts, -1, -1, -1, DateAdd("D", -1, TempFromdate), "('3','4')", mcolParam("FROM_CUSTOMER_CODE"), mcolParam("TO_CUSTOMER_CODE"))
   Call LoadReceiveAmount(Nothing, m_ReceiveAllAmountCNs, -1, -1, -1, DateAdd("D", -1, TempFromdate), "('5')", mcolParam("FROM_CUSTOMER_CODE"), mcolParam("TO_CUSTOMER_CODE"))
   
   Call LoadARCNAmount(Nothing, m_CnAllAmounts, TempFromdate, -1, -1, DateAdd("D", -1, TempFromdate), mcolParam("FROM_CUSTOMER_CODE"), mcolParam("TO_CUSTOMER_CODE"))         'ต้องบวกยอดยกมาเพิ่ม ยอด CN ทั้งหมดของบิลหลัง CN แต่ เอกสาร LINK เป็นในช่วงยกมา
   
   Call LoadARAmount(Nothing, m_ArAmounts, -1, DateAdd("D", -1, TempFromdate), mcolParam("FROM_CUSTOMER_CODE"), mcolParam("TO_CUSTOMER_CODE"), "('3','4')")
   Call LoadARAmount(Nothing, m_ArAmountsCn, -1, DateAdd("D", -1, TempFromdate), mcolParam("FROM_CUSTOMER_CODE"), mcolParam("TO_CUSTOMER_CODE"), "('5')")
   
   Call LoadDistinctDocMonthYearFromReceipt(Nothing, m_DistinctMontnYear, FirstDate, LastDate, "('3','4','5')", mcolParam("FROM_CUSTOMER_CODE"), mcolParam("TO_CUSTOMER_CODE"))
   Call LoadSumMonthYearFromReceipt(Nothing, m_SumMontnYear, FirstDate, LastDate, "('3','4')", mcolParam("FROM_CUSTOMER_CODE"), mcolParam("TO_CUSTOMER_CODE"))
   Call LoadSumMonthYearFromReceipt(Nothing, m_SumMontnYearCN, FirstDate, LastDate, "('5')", mcolParam("FROM_CUSTOMER_CODE"), mcolParam("TO_CUSTOMER_CODE"))
   
   FromDate = "จากวันที่รับชำระ " & EmptyToString(DateToStringExtEx2(mcolParam("FROM_DOC_DATE")), "N/A")
   ToDate = " ถึงวันที่รับชำระ " & EmptyToString(DateToStringExtEx2(mcolParam("TO_DOC_DATE")), "N/A")
   
   MODULE_DESC = glbCompanyName & vbCrLf & mcolParam("REPORT_NAME") & _
                                       vbCrLf & FromDate & "   " & ToDate & " จากลูกค้า " & EmptyToString(mcolParam("FROM_CUSTOMER_CODE"), "N/A") & " ถึงลูกค้า " & EmptyToString(mcolParam("TO_CUSTOMER_CODE"), "N/A")
   
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

Private Function genDoc(isPreview As Boolean) As Boolean
On Error GoTo ErrHandler
Dim RName As String
Dim i As Long
Dim j As Long
Dim K As Long
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
Dim ARt As CARTrn
Dim TempRs As ADODB.Recordset
Dim FromDate  As Date
Dim ToDate  As Date
Dim Acp As CARRcIt
Dim TempAcp  As CARRcIt
Dim Sum As Double
Dim HaveYYYYMM As Boolean
Dim TempAcpCN As CARRcIt

Dim TempArBalance As CARTrn
Dim TempArBalanceCn As CARTrn
Dim TempAcp1  As CARRcIt
Dim TempAcpCn1  As CARRcIt
Dim TempAcpAll  As CARRcIt
Dim TempAcpAllCn  As CARRcIt
Dim TempAcpSumCn As CARTrn
   RName = "genDoc"
'-----------------------------------------------------------------------------------------------------
'                                             Query Here
'-----------------------------------------------------------------------------------------------------
   
   Set TempRs = New ADODB.Recordset
   
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
      Set BodyCf = New CReportField
      i = 0
      
      If i = 0 Then
         GenerateItemNewPage
      End If
      
      FromDate = mcolParam("FROM_DOC_DATE")
      ToDate = mcolParam("TO_DOC_DATE")
      
      While (Year(FromDate) & Format(Month(FromDate), "00")) <= (Year(ToDate) & Format(Month(ToDate), "00"))
         Call m_Details1.ClearField
         '===Head
         K = 0
         Sum = 0
         K = K + 1
         Set HeadCf = m_Headers1.Fields(K)
         TempStr = Month(FromDate) & "/" & (Year(FromDate) + 543)
         Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
         Call m_Details1.AddField(BodyCf)
         '-----------------------------------------------------------------------------------------------------------------------------
         Amt = 0
         HaveYYYYMM = False
         For Each Acp In m_DistinctMontnYear
            Set TempAcp = GetObject("CARRcIt", m_SumMontnYear, Year(FromDate) & "-" & Format(Month(FromDate), "00") & "-" & Year(Acp.DOCDAT) & "-" & Format(Month(Acp.DOCDAT), "00"))
            Set TempAcpCN = GetObject("CARRcIt", m_SumMontnYearCN, Year(FromDate) & "-" & Format(Month(FromDate), "00") & "-" & Year(Acp.DOCDAT) & "-" & Format(Month(Acp.DOCDAT), "00"))
            
            If mcolParam("YYYYMM") <= (Year(Acp.DOCDAT) + 543 & Format(Month(Acp.DOCDAT), "00")) Then
               If (Not HaveYYYYMM) And (Len(mcolParam("YYYYMM")) > 0) Then
                  K = K + 1
                  Set HeadCf = m_Headers1.Fields(K)
                  TempStr = FormatNumberToNull(Amt, 0)
                  Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
                  Call m_Details1.AddField(BodyCf)
                  Sum = Sum + Amt
                  Total1(K) = Total1(K) + Amt
                  Total3(K) = Total3(K) + Amt
                  HaveYYYYMM = True
               End If
               
               Amt = TempAcp.RCVAMT - TempAcpCN.RCVAMT
               K = K + 1
               Set HeadCf = m_Headers1.Fields(K)
               TempStr = FormatNumberToNull(Amt, 0)
               Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
               Call m_Details1.AddField(BodyCf)
               Sum = Sum + Amt
               Total1(K) = Total1(K) + Amt
               Total3(K) = Total3(K) + Amt
            Else
               Amt = Amt + TempAcp.RCVAMT - TempAcpCN.RCVAMT
               'Amt = 0
            End If
         Next Acp
         '-----------------------------------------------------------------------------------------------------------------------------
         
         K = K + 1
         Set HeadCf = m_Headers1.Fields(K)
         TempStr = FormatNumberToNull(Sum, 0)
         'TempStr = Sum
         Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
         Call m_Details1.AddField(BodyCf)
         Total1(K) = Total1(K) + Sum
         Total3(K) = Total3(K) + Sum
         
         Call m_Details1.GetString(1, TempStr1, TempStr2)
         strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
         Call Vsp.AddTable(strFormat, "", TempStr2)
            
         FromDate = DateAdd("M", 1, FromDate)
      Wend
         
      Call GenerateFooter("รวม", Total1, m_Details1)
      Call m_Details1.GetString(2, TempStr1, TempStr2)
      strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
      Vsp.FontBold = True
      Call Vsp.AddTable(strFormat, "", TempStr2)
      Vsp.FontBold = False
   
      '----------ยอดขาย------------------------------------------
      Vsp = ""
      Set Rs = New ADODB.Recordset
      Call m_Details1.ClearField
         '===Head
         K = 0
         Sum = 0
         K = K + 1
         Set HeadCf = m_Headers1.Fields(K)
         TempStr = "ยอดขาย"
         Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
         Call m_Details1.AddField(BodyCf)
         '-----------------------------------------------------------------------------------------------------------------------------
         HaveYYYYMM = False
         Amt = 0
         For Each Acp In m_DistinctMontnYear
            Set ARt = New CARTrn
            Call GetFirstLastDate(Acp.DOCDAT, FromDate, ToDate)
            ARt.FROM_DOC_DATE = FromDate
            ARt.TO_DOC_DATE = ToDate
            ARt.FROM_CUSTOMER_CODE = mcolParam("FROM_CUSTOMER_CODE")
            ARt.TO_CUSTOMER_CODE = mcolParam("TO_CUSTOMER_CODE")
            ARt.RecTypeSet = "('3','4')"
            Call ARt.QueryData(6, Rs, iCount)
            
            Call ARt.PopulateFromRS(6, Rs)
            
            If mcolParam("YYYYMM") <= (Year(Acp.DOCDAT) + 543 & Format(Month(Acp.DOCDAT), "00")) Then
               If (Not HaveYYYYMM) And (Len(mcolParam("YYYYMM")) > 0) Then
                  K = K + 1
                  Set TempArBalance = GetObject("CARTrn", m_ArAmounts, "1")
                  Set TempArBalanceCn = GetObject("CARTrn", m_ArAmountsCn, "1")
                  
                  Set TempAcp1 = GetObject("CARRcIt", m_ReceiveAmounts, "1")
                  Set TempAcpCn1 = GetObject("CARRcIt", m_ReceiveAmountCNs, "1")
               
                  Set TempAcpAll = GetObject("CARRcIt", m_ReceiveAllAmounts, "1")
                  Set TempAcpAllCn = GetObject("CARRcIt", m_ReceiveAllAmountCNs, "1")
                  
                  Set TempAcpSumCn = GetObject("CARTrn", m_CnAllAmounts, "1")
                  
                  Amt = (TempArBalance.AMOUNT - TempArBalance.RCVAMT) - (TempArBalanceCn.AMOUNT - TempArBalanceCn.RCVAMT) - (TempAcp1.RCVAMT - TempAcpCn1.RCVAMT) + (TempAcpAll.RCVAMT - TempAcpAllCn.RCVAMT) + TempAcpSumCn.AMOUNT
                  
                  Set HeadCf = m_Headers1.Fields(K)
                  TempStr = FormatNumberToNull(Amt, 0)
                  Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
                  Call m_Details1.AddField(BodyCf)
                  Sum = Sum + Amt
                  Total2(K) = Total2(K) + Amt
                  Total3(K) = Total3(K) - Amt
                  HaveYYYYMM = True
               End If
               Amt = ARt.AMOUNT
               K = K + 1
               Set HeadCf = m_Headers1.Fields(K)
               TempStr = FormatNumberToNull(Amt, 0)
               Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
               Call m_Details1.AddField(BodyCf)
               Sum = Sum + Amt
               Total2(K) = Total2(K) + Amt
               Total3(K) = Total3(K) - Amt
            Else
               Amt = 0
            End If
         Next Acp
         '-----------------------------------------------------------------------------------------------------------------------------
         K = K + 1
         Set HeadCf = m_Headers1.Fields(K)
         TempStr = FormatNumberToNull(Sum, 0)
         Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
         Call m_Details1.AddField(BodyCf)
         Total2(K) = Total2(K) + Sum
         Total3(K) = Total3(K) - Sum
         
         Call m_Details1.GetString(1, TempStr1, TempStr2)
         strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
         Call Vsp.AddTable(strFormat, "", TempStr2)
         
         
      '-----------------ยอดขาย----------------------------------------
      '----------------รับคืน----------------------------------------
      Set Rs = New ADODB.Recordset
      Call m_Details1.ClearField
         '===Head
         K = 0
         Sum = 0
         K = K + 1
         Set HeadCf = m_Headers1.Fields(K)
         TempStr = "ลดหนี้"
         Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
         Call m_Details1.AddField(BodyCf)
         '-----------------------------------------------------------------------------------------------------------------------------
         HaveYYYYMM = False
         Amt = 0
         For Each Acp In m_DistinctMontnYear
            Set ARt = New CARTrn
            Call GetFirstLastDate(Acp.DOCDAT, FromDate, ToDate)
            ARt.FROM_DOC_DATE = FromDate
            ARt.TO_DOC_DATE = ToDate
            ARt.FROM_CUSTOMER_CODE = mcolParam("FROM_CUSTOMER_CODE")
            ARt.TO_CUSTOMER_CODE = mcolParam("TO_CUSTOMER_CODE")
            ARt.RecTypeSet = "('5')"
            Call ARt.QueryData(6, Rs, iCount)
            
            Call ARt.PopulateFromRS(6, Rs)
            
            If mcolParam("YYYYMM") <= (Year(Acp.DOCDAT) + 543 & Format(Month(Acp.DOCDAT), "00")) Then
               If (Not HaveYYYYMM) And (Len(mcolParam("YYYYMM")) > 0) Then
                  K = K + 1
                  Set HeadCf = m_Headers1.Fields(K)
                  TempStr = FormatNumberToNull(Amt, 0)
                  Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
                  Call m_Details1.AddField(BodyCf)
                  Sum = Sum + Amt
                  Total2(K) = Total2(K) - Amt
                  Total3(K) = Total3(K) + Amt
                  HaveYYYYMM = True
               End If
               Amt = ARt.AMOUNT
               K = K + 1
               Set HeadCf = m_Headers1.Fields(K)
               TempStr = FormatNumberToNull(Amt, 0)
               Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
               Call m_Details1.AddField(BodyCf)
               Sum = Sum + Amt
               Total2(K) = Total2(K) - Amt
               Total3(K) = Total3(K) + Amt
            Else
               Amt = Amt + ARt.AMOUNT
            End If
         Next Acp
         '-----------------------------------------------------------------------------------------------------------------------------
         
         K = K + 1
         Set HeadCf = m_Headers1.Fields(K)
         TempStr = FormatNumberToNull(Sum, 0)
         Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
         Call m_Details1.AddField(BodyCf)
         Total2(K) = Total2(K) - Sum
         Total3(K) = Total3(K) + Sum
         
         Call m_Details1.GetString(1, TempStr1, TempStr2)
         strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
         Call Vsp.AddTable(strFormat, "", TempStr2)
      
      '----------------รับคืน----------------------------------------
      
      Call GenerateFooter("รวม", Total2, m_Details1)
      Call m_Details1.GetString(2, TempStr1, TempStr2)
      strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
      Vsp.FontBold = True
      Call Vsp.AddTable(strFormat, "", TempStr2)
      Vsp.FontBold = False
      
      Call GenerateFooter("ผลต่าง", Total3, m_Details1)
      Call m_Details1.GetString(2, TempStr1, TempStr2)
      strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
      Vsp.FontBold = True
      Call Vsp.AddTable(strFormat, "", TempStr2)
      Vsp.FontBold = False
      
'      Set ARt = New CARTrn
'
'      ARt.FROM_DOC_DATE = mcolParam("FROM_DOC_DATE")
'      ARt.TO_DOC_DATE = mcolParam("TO_DOC_DATE")
'      ARt.FROM_CUSTOMER_CODE = mcolParam("FROM_CUSTOMER_CODE")
'      ARt.TO_CUSTOMER_CODE = mcolParam("TO_CUSTOMER_CODE")
'      ARt.OrderBy = 4
'      ARt.OrderType = 1
'      ARt.QueryFlag = -1
'      ARt.RecTypeSet = "('3','4')"
'      Call glbDaily.QueryARTran(ARt, Rs, iCount, IsOK, glbErrorLog)
'
'         While Not Rs.EOF
'            i = i + 1
'            Call m_Details1.ClearField
'            Call ARt.PopulateFromRS(1, Rs)
'
'            If PrevKey1 <> ARt.CUSCOD And i <> 1 Then
'                Call GenerateFooter("รวม", "ยอดขายงวด " & PrevKey2, "", Total1, m_Details1)
'               Call m_Details1.GetString(2, TempStr1, TempStr2)
'               strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
'               Vsp.FontBold = True
'               Call Vsp.AddTable(strFormat, "", TempStr2)
'               Vsp.FontBold = False
'
'                Call GenerateFooter("รวม", "ลูกค้า" & PrevKey1, "", Total2, m_Details1)
'               Call m_Details1.GetString(2, TempStr1, TempStr2)
'               strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
'               Vsp.FontBold = True
'               Call Vsp.AddTable(strFormat, "", TempStr2)
'               Vsp.FontBold = False
'
'               For j = 1 To UBound(Total1)
'                  Total1(j) = 0
'                  Total2(j) = 0
'               Next j
'
'               Call GenerateHeader(ARt.CUSCOD, ARt.CUSNAM, m_Details1)
'               Call m_Details1.GetString(2, TempStr1, TempStr2)
'               strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
'               Vsp.FontBold = True
'               Call Vsp.AddTable(strFormat, "", TempStr2)
'               Vsp.FontBold = False
'            ElseIf (PrevKey2 <> (Month(ARt.DOCDAT) & "/" & (Year(ARt.DOCDAT) + 543))) And i <> 1 Then
'                Call GenerateFooter("รวม", "ยอดขายงวด " & PrevKey2, "", Total1, m_Details1)
'               Call m_Details1.GetString(2, TempStr1, TempStr2)
'               strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
'               Vsp.FontBold = True
'               Call Vsp.AddTable(strFormat, "", TempStr2)
'               Vsp.FontBold = False
'
'               For j = 1 To UBound(Total1)
'                  Total1(j) = 0
'               Next j
'
'            ElseIf i = 1 Then
'                Call GenerateHeader(ARt.CUSCOD, ARt.CUSNAM, m_Details1)
'               Call m_Details1.GetString(2, TempStr1, TempStr2)
'               strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
'               Vsp.FontBold = True
'               Call Vsp.AddTable(strFormat, "", TempStr2)
'               Vsp.FontBold = False
'            End If
'
'            PrevKey1 = ARt.CUSCOD
'            PrevKey2 = (Month(ARt.DOCDAT) & "/" & (Year(ARt.DOCDAT) + 543))
'            Call m_Details1.ClearField
'            '===Head
'            Set HeadCf = m_Headers1.Fields(1)
'            TempStr = i
'            Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
'            Call m_Details1.AddField(BodyCf)
'
'            K = 1
'            K = K + 1
'            Set HeadCf = m_Headers1.Fields(K)
'            TempStr = DateToStringExtEx2(ARt.DOCDAT)
'            Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
'            Call m_Details1.AddField(BodyCf)
'
'            K = K + 1
'            Set HeadCf = m_Headers1.Fields(K)
'            TempStr = ARt.DOCNUM
'            Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
'            Call m_Details1.AddField(BodyCf)
'
'            K = K + 1
'            Set HeadCf = m_Headers1.Fields(K)
'            TempStr = FormatNumber(ARt.AMOUNT, "")
'            Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
'            Call m_Details1.AddField(BodyCf)
'
'            Total1(K) = Total1(K) + ARt.AMOUNT
'            Total2(K) = Total2(K) + ARt.AMOUNT
'            Total3(K) = Total3(K) + ARt.AMOUNT
'
'            Call m_Details1.GetString(1, TempStr1, TempStr2)
'            strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
'            Call Vsp.AddTable(strFormat, "", TempStr2)
'
'            Rs.MoveNext
'         Wend
'
'         Call GenerateFooter("รวม", "ยอดขายงวด " & PrevKey2, "", Total1, m_Details1)
'        Call m_Details1.GetString(2, TempStr1, TempStr2)
'        strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
'        Vsp.FontBold = True
'        Call Vsp.AddTable(strFormat, "", TempStr2)
'        Vsp.FontBold = False
'
'         Call GenerateFooter("รวม", "ลูกค้า " & PrevKey1, "", Total2, m_Details1)
'        Call m_Details1.GetString(2, TempStr1, TempStr2)
'        strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
'        Vsp.FontBold = True
'        Call Vsp.AddTable(strFormat, "", TempStr2)
'        Vsp.FontBold = False
'
'      Set BodyCf = Nothing
'
'      Call GenerateFooter("สุทธิ", "", "", Total3, m_Details1)
'      Call m_Details1.GetString(2, TempStr1, TempStr2)
'      strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
'      Vsp.FontBold = True
'      Call Vsp.AddTable(strFormat, "", TempStr2)
'      Vsp.FontBold = False
'
'      Vsp = ""
'      Vsp = ""
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

Private Sub GenerateFooter(Txt1 As String, Tot() As Double, Details As CFieldList)
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
      ElseIf j > 1 Then
         TempStr = FormatNumberToNull(Tot(j), 0)
      End If
      
      Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.Align, TempStr, HeadCf.BodyAlign)
      Call Details.AddField(BodyCf)
   Next HeadCf
      
   Set BodyCf = Nothing
End Sub
Private Sub GenerateHeader(Txt1 As String, Txt2 As String, Details As CFieldList)
Dim HeadCf As CReportField
Dim BodyCf As CReportField
Dim j As Long
Dim TempStr As String

   Call Details.ClearField
   Set BodyCf = New CReportField
   
   j = 0
   For Each HeadCf In m_Headers0.Fields
      j = j + 1
      If j = 1 Then
         TempStr = Txt1
      ElseIf j = 2 Then
         TempStr = Txt2
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
   If Vsp.PageCount > 1 Then
      Call GenerateItemNewPage
   End If
End Sub
Private Sub GenerateItemNewPage()
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
Private Sub SumValue(Coll As Collection, KEY As String, Value As Double)
Dim GroupTypeItem As CColumnValue
Dim TempColumnValue  As CColumnValue
   
   Set GroupTypeItem = GetObject("CColumnValue", Coll, KEY, False)
   If GroupTypeItem Is Nothing Then
      Set TempColumnValue = New CColumnValue
      TempColumnValue.KEY = KEY
      TempColumnValue.Value = Value
      Call Coll.Add(TempColumnValue, KEY)
   Else
      GroupTypeItem.Value = GroupTypeItem.Value + Value
   End If
End Sub
