VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CReportSaleApprove"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Implements CReportInterface

Private Const MODULE_NAME = "CReportSaleApprove"
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

Private m_Headers0 As CFieldList
Private m_Headers1 As CFieldList
Private m_Headers1_1 As CFieldList
Private m_Headers1_2 As CFieldList
Private m_Headers1_3 As CFieldList
Private m_Headers1_4 As CFieldList
Private m_Details1 As CFieldList
Private m_Details2 As CFieldList

Private m_CustomerTypes As Collection
Private m_ReceiveAmounts As Collection
Private m_ReceiveAllAmounts As Collection
Private m_MaxMins As Collection
Private m_ArCustAmounts As Collection
Private m_ArCustPaids As Collection
Private m_DbnCdnByBills As Collection
Private m_RealCreditColl As Collection
Private m_CnAllAmounts As Collection
Private ColProvinceMap As Collection
Private m_ReceiveFutureAmounts As Collection
Private m_ReceiveFutureAmountsFuture As Collection
Private m_RealCreditDocColl As Collection
Private m_RealCreditCusColl As Collection
Private m_ChqnumDocDate As Collection
Private RealCredit As Long
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
   Set m_CustomerTypes = New Collection
   Set m_ReceiveAmounts = New Collection
   Set m_MaxMins = New Collection
   Set m_ArCustAmounts = New Collection
   Set m_ArCustPaids = New Collection
   Set m_ReceiveAllAmounts = New Collection
   Set m_DbnCdnByBills = New Collection
   Set m_RealCreditColl = New Collection
   Set m_CnAllAmounts = New Collection
   Set ColProvinceMap = New Collection
   Set m_ReceiveFutureAmounts = New Collection
   Set m_ReceiveFutureAmountsFuture = New Collection
   Set m_RealCreditCusColl = New Collection
   Set m_RealCreditDocColl = New Collection
   Set m_ChqnumDocDate = New Collection
End Sub

Private Sub Class_Terminate()
   Call ClearParam
   Set mcolParam = Nothing
   Set Vsp = Nothing
   Set m_TempCol = Nothing
   
   Set m_Headers0 = Nothing
   Set m_Headers1 = Nothing
   Set m_Details1 = Nothing
   Set m_Headers1_1 = Nothing
   Set m_Headers1_2 = Nothing
   Set m_Headers1_3 = Nothing
   Set m_Headers1_4 = Nothing
   Set m_Details2 = Nothing
   Set m_CustomerTypes = Nothing
   Set m_ReceiveAmounts = Nothing
   Set m_MaxMins = Nothing
   Set m_ReceiveAllAmounts = Nothing
   Set m_DbnCdnByBills = Nothing
   Set m_RealCreditColl = Nothing
   Set m_CnAllAmounts = Nothing
   Set ColProvinceMap = Nothing
   Set m_ReceiveFutureAmounts = Nothing
   Set m_ReceiveFutureAmountsFuture = Nothing
   Set m_RealCreditDocColl = Nothing
   Set m_RealCreditCusColl = Nothing
   Set m_ChqnumDocDate = Nothing
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

   tbdT = Vsp.TableBorder
   blnBold = Vsp.FontBold
   iSize = Vsp.FontSize
   
   Vsp.FontSize = HEADER_SIZE
   Vsp.FontBold = True
   Vsp.TableBorder = tbAll
   
   Call m_Headers0.ClearField
   Call m_Headers1.ClearField
   Call m_Headers1_1.ClearField
   
   Set Cf = New CReportField
      
   '===
   Call Cf.SetFieldValue(1, "^", "NO.", "^")
   Call m_Headers0.AddField(Cf)
   Call Cf.SetFieldValue(3, "^", "", "^")
   Call m_Headers0.AddField(Cf)
   Call Cf.SetFieldValue(3, "^", "", "^")
   Call m_Headers0.AddField(Cf)
   Call Cf.SetFieldValue(3, "^", "", "^")
   Call m_Headers0.AddField(Cf)
   Call Cf.SetFieldValue(8, "^", "ยอดวัน", ">")
   Call m_Headers0.AddField(Cf)
   Call Cf.SetFieldValue(2, "^", "", "^")
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
            
   Set TempRs = Nothing
   
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
Dim ToPaidDate As String
   mblnNewPage = False

   Vsp.PaperSize = pprA4 '
   Vsp.ORIENTATION = orPortrait
   Vsp.NewPage
   Vsp.MarginBottom = 700
   Vsp.MarginFooter = 700
   Vsp.MarginHeader = 700
   Vsp.MarginLeft = 700
   Vsp.MarginRight = 700
   Vsp.MarginTop = 700
   Vsp.FontName = "AngsanaUPC"


   mdblWid = Vsp.PageWidth - Vsp.MarginLeft - Vsp.MarginRight
   
   If mcolParam("REAL_CREDIT_FLAG") = "Y" Then
      Call LoadRealCreditNo(Nothing, m_RealCreditDocColl)
      Call LoadRealCreditNo(Nothing, m_RealCreditCusColl, False)
   End If
   
   Call LoadCustomerType(Nothing, m_CustomerTypes)
   Call LoadReceiveAmountByBill(Nothing, m_ReceiveAmounts, -1, mcolParam("TO_PAY_DATE"))
   Call LoadReceiveAmountByBill(Nothing, m_ReceiveAllAmounts, -1, -1)
      
   Call LoadARCNAmountByBill(Nothing, m_CnAllAmounts, DateAdd("D", 1, mcolParam("TO_DOC_DATE")), -1, -1, mcolParam("TO_DOC_DATE"))         'ต้องบวกยอดยกมาเพิ่ม ยอด CN ทั้งหมดของบิลหลัง CN แต่ เอกสาร LINK เป็นในช่วงยกมา
   
   
   Call LoadChequeFuture(Nothing, m_ReceiveFutureAmounts, mcolParam("FROM_CHECK_DATE"), mcolParam("TO_CHECK_DATE"), mcolParam("CUSTOMER_CODE"), mcolParam("CUSTOMER_CODE"))
   Call LoadReceiveAmountByBill(Nothing, m_ReceiveFutureAmountsFuture, DateAdd("D", 1, mcolParam("TO_PAY_DATE")), -1)
'   Call LoadBktChqnumDocDate(Nothing, m_ChqnumDocDate, , , mcolParam("FROM_CHECK_DATE"), mcolParam("TO_CHECK_DATE"))
   Call LoadBktChqnumDocDateAR(Nothing, m_ChqnumDocDate, mcolParam("FROM_CHECK_DATE"), mcolParam("TO_CHECK_DATE")) '
   
    Call LoadProvinceMap(Nothing, ColProvinceMap)

'   If mcolParam("INTERVAL_TYPE") = 1 Then
'      Call LoadDueDateInterval1(Nothing, m_MaxMins)
'   ElseIf mcolParam("INTERVAL_TYPE") = 3 Then
'      Call LoadDueDateInterval3(Nothing, m_MaxMins)
'   ElseIf mcolParam("INTERVAL_TYPE") = 4 Then
'      Call LoadDueDateInterval5(Nothing, m_MaxMins)
'   Else
'      Call LoadDueDateInterval2(Nothing, m_MaxMins)
'   End If
   
'   Call LoadDueDateInterval6(Nothing, m_MaxMins)
   
   
   
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

Private Sub GenerateDetail(alngX() As Long, i As Long, ARt As CARTrn, Tot1() As Double, Tot2() As Double, HaveHeader As Boolean, PrevKey1 As String, PrevKey4 As String)
Dim HeadCf As CReportField
Dim TempStr As String
Dim TempStr1 As String
Dim TempStr2 As String
Dim BodyCf As CReportField
Dim strFormat As String
Dim Ari As CARRcIt
Dim Ari2 As CARRcIt
Dim Apt1 As CARTrn
Dim DueCount As Long
Dim MM As CMaxMin
Dim j As Long
Dim Sum As Double
Dim PrintDate As Date
Dim PaidAmount As Double
Dim PaidBalance As Double 'ชำระแล้วยกมา
Dim TempAmt As Double
Dim TempRealCredit As CRealCredit
Dim ARt2 As CARTrn
Dim Width As Double
Dim AdjustCredit As Double
Dim ARtFuture As CARRcIt
   
   If mcolParam("PRINT_DATE") <= 0 Then
      PrintDate = Now
   Else
      PrintDate = mcolParam("PRINT_DATE")
   End If
   
   Set TempRealCredit = GetObject("CRealCredit", m_RealCreditDocColl, Trim(ARt.DOCNUM), False)
   If TempRealCredit Is Nothing Then
      Set TempRealCredit = GetObject("CRealCredit", m_RealCreditCusColl, Trim(ARt.CUSCOD), False)
      If TempRealCredit Is Nothing Then
         If Val(mcolParam("CREDIT")) > 0 Then
           DueCount = DateDiff("D", DateAdd("D", Val(mcolParam("CREDIT")), ARt.DOCDAT), PrintDate)
           AdjustCredit = Val(mcolParam("CREDIT"))
         ElseIf mcolParam("NINETY_FLAG") = "Y" Then
            DueCount = DateDiff("D", DateAdd("D", 90, ARt.DOCDAT), PrintDate)
            AdjustCredit = 90
         Else
            DueCount = DateDiff("D", ARt.DUEDAT, PrintDate)
            AdjustCredit = DateDiff("D", ARt.DOCDAT, ARt.DUEDAT)
         End If
         RealCredit = AdjustCredit
      Else
         DueCount = DateDiff("D", DateAdd("D", TempRealCredit.REAL_CREDIT, ARt.DOCDAT), PrintDate)
         AdjustCredit = TempRealCredit.REAL_CREDIT
         RealCredit = TempRealCredit.REAL_CREDIT
      End If
   Else
      DueCount = DateDiff("D", DateAdd("D", TempRealCredit.REAL_CREDIT, ARt.DOCDAT), PrintDate)
      AdjustCredit = TempRealCredit.REAL_CREDIT
   End If
   Set Ari = GetARRcpItem(m_ReceiveAmounts, ARt.DOCNUM)
   Set Ari2 = GetARRcpItemEx(m_ReceiveAllAmounts, ARt.DOCNUM)
   Set ARt2 = GetARTrn(m_CnAllAmounts, ARt.DOCNUM)
   'รับชำระแล้วยกมาเท่ากับ APT.PAYAMT - ยอดรับชำระทั้งหมดของใบสำคัญจ่าย
   PaidBalance = ARt.RCVAMT - Ari2.RCVAMT - ARt2.AMOUNT
      
          j = 1
         Sum = 0

         If (-9999999 < DueCount) And (DueCount <= 0) Then 'ในดิว
            TempAmt = (ARt.AMOUNT) - (Ari.RCVAMT + PaidBalance)
            If ARt.RECTYP = 5 Then 'ลดหนี้
               TempAmt = -1 * TempAmt
            End If
            Tot2(j) = Tot2(j) + TempAmt
            Sum = Sum + TempAmt
            
            Set ARtFuture = GetARRcpItem(m_ReceiveFutureAmountsFuture, ARt.DOCNUM)
            TempAmt = ARtFuture.RCVAMT
            If ARt.RECTYP = 5 Then 'ลดหนี้
               TempAmt = -1 * TempAmt
            End If
             Tot1(j) = Tot1(j) + TempAmt
             
         End If

            j = j + 1
         If (0 < DueCount) And (DueCount <= 9999999) Then 'เกินดิว
            TempAmt = (ARt.AMOUNT) - (Ari.RCVAMT + PaidBalance)
            If ARt.RECTYP = 5 Then 'ลดหนี้
               TempAmt = -1 * TempAmt
            End If
            Tot2(j) = Tot2(j) + TempAmt
            Sum = Sum + TempAmt
            
            Set ARtFuture = GetARRcpItem(m_ReceiveFutureAmountsFuture, ARt.DOCNUM)
            TempAmt = ARtFuture.RCVAMT
            If ARt.RECTYP = 5 Then 'ลดหนี้
               TempAmt = -1 * TempAmt
            End If
             Tot1(j) = Tot1(j) + TempAmt
             
         End If
         
         j = j + 1
         Tot2(j) = Tot2(j) + Sum 'รวม
        
'         j = j + 1
'         Set ARtFuture = GetARRcpItem(m_ReceiveFutureAmountsFuture, ARt.DOCNUM)
'         TempAmt = ARtFuture.RCVAMT
'         If ARt.RECTYP = 5 Then 'ลดหนี้
'            TempAmt = -1 * TempAmt
'         End If
'         Tot2(j) = Tot2(j) + TempAmt


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
Dim PrevKey4 As String
Dim St As CIsTab
Dim Ar As CARMas

Dim Width As Double
Dim ARt As CARTrn
Dim TempRs As ADODB.Recordset
Dim TempARt As CARTrn
Dim TempARRcIt As CARRcIt
Dim HaveHeader As Boolean
   RName = "genDoc"
'-----------------------------------------------------------------------------------------------------
'                                             Query Here
'-----------------------------------------------------------------------------------------------------
   
   Set Rs = New ADODB.Recordset
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
      i = 0
      Set BodyCf = New CReportField

      Set Ar = New CARMas
      Ar.CUSCOD = mcolParam("CUSTOMER_CODE")
      Call glbDaily.QueryCustomer(Ar, Rs, iCount, IsOK, glbErrorLog)
      
      While Not Rs.EOF
         Call m_Details1.ClearField
         Call Ar.PopulateFromRS(1, Rs)

         HaveHeader = False
         If True Then  'มียอดค้างจ่าย
            '===Head
            '=== Detail
            Set ARt = New CARTrn
            ARt.CUSCOD = Ar.CUSCOD
            ARt.FROM_DOC_DATE = mcolParam("FROM_DOC_DATE")
            ARt.TO_DOC_DATE = mcolParam("TO_DOC_DATE")
            ARt.RECTYP = ""
            ARt.RecTypeSet = "('3', '4', '5')"
            ARt.OrderBy = 1
            ARt.OrderType = 1
            Call ARt.QueryData(1, TempRs, iCount, False)
            
            If Not TempRs.EOF Then
               For j = 1 To UBound(Total1)
                  Total1(j) = 0
               Next j
            End If
            
            While Not TempRs.EOF
               i = i + 1
               Call m_Details1.ClearField
               Call ARt.PopulateFromRS(1, TempRs)

               Call GenerateDetail(alngX(), i, ARt, Total1, Total2, HaveHeader, PrevKey1, PrevKey4)
               
               TempRs.MoveNext
            Wend
            
            Set ARt = Nothing
          End If
          Call GenerateSaleApprove(Ar, TempRs, Total1, Total2, m_Details1, alngX())
          Call GenerateCheque(Ar, Total2, m_Details1, alngX())
         Rs.MoveNext
      Wend
      
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

Private Sub GenerateFooter(Txt1 As String, Txt2 As String, Txt3 As String, Tot() As Double, Details As CFieldList, DateCount As Long)
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
'         If mcolParam("SUMMARY_FLAG") = "N" Then
'            TempStr = Txt1
'         End If
         TempStr = Txt1
      ElseIf j = 2 Then
         TempStr = Txt2
      ElseIf j = 3 Then
         TempStr = Txt3
      Else
         TempStr = FormatNumber(Tot(j), "-")
      End If
      
      If j = 3 Then
         Call BodyCf.SetFieldValue(m_Headers1.Fields(3).Width + m_Headers1.Fields(4).Width, HeadCf.Align, TempStr, "<")
         Call Details.AddField(BodyCf)
      ElseIf j = 4 Then
        'รวมกับ 3 แล้ว
      Else
         Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.Align, TempStr, HeadCf.BodyAlign)
         Call Details.AddField(BodyCf)
      End If
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
'      Call printHeader
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
Private Sub GenerateSaleApprove(Ar As CARMas, TempARt As ADODB.Recordset, Tot1() As Double, Tot2() As Double, Details As CFieldList, alngX() As Long)
Dim strFormat As String
Dim strPrint As String
Dim TempX As Double
Dim tempY As Double
Dim i As Integer
Dim PlusAmt As Double
Dim MinusAmt As Double
Dim PrintDate As Date
Dim SentDate As Date


   TempX = Vsp.CurrentX
   
     Vsp.TableBorder = tbNone
     Vsp.FontBold = True
      Vsp.FontSize = 16
      strFormat = VSP_CalTable("^3;", mdblWid, alngX)
      strPrint = glbCompanyName
      Call Vsp.AddTable(strFormat, "", strPrint)

      Vsp.FontSize = 14
      strFormat = VSP_CalTable("^20|>7|<13;", mdblWid, alngX)
      strPrint = "แบบฟอร์มขออนุมัติขาย|พนักงานขาย|" & Ar.SLMNAM & "(" & Ar.SLMCOD & ")" & ";"
      Call Vsp.AddTable(strFormat, "", strPrint)

    
         If mcolParam("PRINT_DATE") <= 0 Then
            PrintDate = Now
         Else
            PrintDate = mcolParam("PRINT_DATE")
         End If
   
      Vsp.FontSize = DETAIL_SIZE
      strFormat = VSP_CalTable(">4|<17|>4|<6|>4|<12|>4|10;", mdblWid, alngX)
      strPrint = "ชื่อลูกค้า|" & Ar.CUSNAM & "(" & Ar.CUSCOD & ")" & "|เครดิต|" & RealCredit & " วัน" & "|จังหวัด|" & GetProvince(Ar.ADDR02, ColProvinceMap) & "|วันที่  |" & DateToStringExtEx2(PrintDate)
      Call Vsp.AddTable(strFormat, "", strPrint)

      Vsp.CurrentX = TempX
     Call Vsp.DrawLine(Vsp.CurrentX, Vsp.CurrentY, Vsp.PageWidth - 500, Vsp.CurrentY)
      Vsp.FontBold = False
      strFormat = VSP_CalTable("<1|<3|>3|<7;", mdblWid, alngX)
      strPrint = "|วงเงิน|" & FormatNumber(Ar.CRLINE) & "|บาท;"
      Call Vsp.AddTable(strFormat, "", strPrint)
      strPrint = "|ยอดหนี้ปัจจุบัน|" & FormatNumber(Tot2(3)) & "|บาท;"
      PlusAmt = Tot2(3)
      Call Vsp.AddTable(strFormat, "", strPrint)
      
      If mcolParam("SENT_DATE") <= 0 Then
         SentDate = Now
      Else
         SentDate = mcolParam("SENT_DATE")
      End If
      strFormat = VSP_CalTable("<1|<3|>3|<2|>2|^3;", mdblWid, alngX)
      If Val(mcolParam("WANT_BY_MORE")) > 0 Then
         strPrint = "|ต้องการซื้อเพิ่มเติม|" & FormatNumber(Val(mcolParam("WANT_BY_MORE"))) & "|บาท|วันที่ส่งของ|" & DateToStringExtEx2(SentDate) & ";"
      Else
         strPrint = "|ต้องการซื้อเพิ่มเติม|" & FormatNumber(Val(mcolParam("WANT_BY_MORE"))) & "|บาท||;"
      End If
      PlusAmt = PlusAmt + Val(mcolParam("WANT_BY_MORE"))
      Call Vsp.AddTable(strFormat, "", strPrint)
      
      strPrint = "|หนี้ปัจจุบันรวมซื้อเพิ่ม|" & FormatNumber(PlusAmt) & "|บาท;"
      Call Vsp.AddTable(strFormat, "", strPrint)
      strPrint = "|คงเหลือวงเงิน|" & FormatNumber(Ar.CRLINE - PlusAmt) & "|บาท;"
      Call Vsp.AddTable(strFormat, "", strPrint)

     PlusAmt = 0
     Vsp.Paragraph = ""
      Call Vsp.DrawLine(Vsp.CurrentX, Vsp.CurrentY, Vsp.PageWidth - 500, Vsp.CurrentY)
      Vsp.FontBold = True
      Vsp.Paragraph = "ความเห็นฝ่ายข้อมูล"
      Vsp.FontBold = False
      strFormat = VSP_CalTable("<1|<3|>3|<7;", mdblWid, alngX)
      strPrint = "|ยอดหนี้ที่เกิน DUE|" & FormatNumber(Tot2(2) - Tot1(2)) & "|บาท;"
      PlusAmt = Tot2(2) - Tot1(2)
      Call Vsp.AddTable(strFormat, "", strPrint)
       
      Vsp.CurrentX = TempX
      strFormat = VSP_CalTable("<1|<3|>3|<7;", mdblWid, alngX)
      strPrint = "|รวมเช็ครับล่วงหน้า|" & FormatNumber(Tot1(1) + Tot1(2)) & "|บาท;"
      PlusAmt = PlusAmt + Tot1(1) + Tot1(2)
      Call Vsp.AddTable(strFormat, "", strPrint)
      
      strPrint = "|ยอดหนี้ที่ไม่เกิน DUE|" & FormatNumber(Tot2(1) - Tot1(1)) & "|บาท;"
      PlusAmt = PlusAmt + (Tot2(1) - Tot1(1))
      Call Vsp.AddTable(strFormat, "", strPrint)

      strPrint = "|รวม|" & FormatNumber(PlusAmt) & "|บาท;"
      Call Vsp.AddTable(strFormat, "", strPrint)
      
      strPrint = "|*ยอดเช็คคืน|" & FormatNumber(Val(mcolParam("CHEQUE_RE_AMOUNT"))) & "|บาท;"
      Call Vsp.AddTable(strFormat, "", strPrint)
'      PlusAmt = PlusAmt + (Val(mcolParam("CHEQUE_RE_AMOUNT")) * -1)
      
      Vsp.CurrentX = TempX
      Vsp.Paragraph = ""
      Call Vsp.DrawLine(Vsp.CurrentX, Vsp.CurrentY, Vsp.PageWidth - 500, Vsp.CurrentY)
      
      strFormat = VSP_CalTable("<1|<3|>3|<2|>2|^3;", mdblWid, alngX)
      If Val(mcolParam("WANT_BY_MORE")) > 0 Then
         strPrint = "|ต้องการซื้อเพิ่มเติม|" & FormatNumber(Val(mcolParam("WANT_BY_MORE"))) & "|บาท|วันที่ส่งของ|" & DateToStringExtEx2(SentDate) & ";"
      Else
         strPrint = "|ต้องการซื้อเพิ่มเติม|" & FormatNumber(Val(mcolParam("WANT_BY_MORE"))) & "|บาท||;"
      End If
      PlusAmt = PlusAmt + Val(mcolParam("WANT_BY_MORE"))
      Call Vsp.AddTable(strFormat, "", strPrint)
      
      strPrint = "|รวมหนี้ทั้งหมด|" & FormatNumber(PlusAmt) & "|บาท;"
      Call Vsp.AddTable(strFormat, "", strPrint)


   For i = 1 To 4
    '============
     Vsp.Paragraph = ""
    Call Vsp.DrawLine(Vsp.CurrentX, Vsp.CurrentY, Vsp.PageWidth - 500, Vsp.CurrentY)
      Vsp.FontBold = True
      If i = 1 Then
         Vsp.Paragraph = "ความเห็นฝ่ายการตลาด(ผจก.)"
      ElseIf i = 2 Then
         Vsp.Paragraph = "ความเห็นฝ่ายสินเชื่อ"
      ElseIf i = 3 Then
         Vsp.Paragraph = "ความเห็นผู้มีอำนาจอนุมัติ"
      ElseIf i = 4 Then
         Vsp.Paragraph = "ความเห็นผู้มีอำนาจอนุมัติ*ผู้บริหาร/กรณี CREDIT LIMIT > 500,000.00 บาท"
      End If
      Vsp.FontBold = False
      Vsp.Paragraph = ""

      Call Vsp.DrawLine(Vsp.CurrentX + 800, Vsp.CurrentY, Vsp.PageWidth - 1500, Vsp.CurrentY)
      Vsp.Paragraph = ""

      Call Vsp.DrawLine(Vsp.CurrentX + 800, Vsp.CurrentY, Vsp.PageWidth - 1500, Vsp.CurrentY)
      Vsp.Paragraph = ""

      Vsp.FontBold = True
      Call Vsp.DrawLine(Vsp.CurrentX + 500, Vsp.CurrentY, Vsp.CurrentX + 1000, Vsp.CurrentY)
      TempX = Vsp.CurrentX
      tempY = Vsp.CurrentY
      Vsp.CurrentX = Vsp.CurrentX + 1000
      Vsp.CurrentY = Vsp.CurrentY - 100
      Vsp = "    เห็นด้วย"
      Vsp.CurrentX = TempX
      Vsp.CurrentY = tempY
      Call Vsp.DrawLine(Vsp.PageWidth - 4500, Vsp.CurrentY, Vsp.PageWidth - 4000, Vsp.CurrentY)
      TempX = Vsp.CurrentX
      Vsp.CurrentX = Vsp.PageWidth - 4000
      Vsp.CurrentY = Vsp.CurrentY - 100
      Vsp = "    ไม่เห็นด้วย"
      Vsp.CurrentX = TempX
      Vsp.CurrentY = tempY

      Call Vsp.DrawLine(Vsp.CurrentX + 500, Vsp.CurrentY, Vsp.CurrentX + 500, Vsp.CurrentY + 250)
      Call Vsp.DrawLine(Vsp.CurrentX + 1000, Vsp.CurrentY, Vsp.CurrentX + 1000, Vsp.CurrentY + 250)
      Call Vsp.DrawLine(Vsp.PageWidth - 4500, Vsp.CurrentY, Vsp.PageWidth - 4500, Vsp.CurrentY + 250)
      Call Vsp.DrawLine(Vsp.PageWidth - 4000, Vsp.CurrentY, Vsp.PageWidth - 4000, Vsp.CurrentY + 250)

      Call Vsp.DrawLine(Vsp.CurrentX + 500, Vsp.CurrentY + 250, Vsp.CurrentX + 1000, Vsp.CurrentY + 250)
      Call Vsp.DrawLine(Vsp.PageWidth - 4500, Vsp.CurrentY + 250, Vsp.PageWidth - 4000, Vsp.CurrentY + 250)
      Vsp.FontBold = False

      Vsp.Paragraph = ""
      Vsp = "                                           ลงชื่อ    ........................................................................  วันที่   ............/............../................."
      '============
   Next i
   Vsp.Paragraph = ""
    Call Vsp.DrawLine(Vsp.CurrentX, Vsp.CurrentY, Vsp.PageWidth - 500, Vsp.CurrentY)
End Sub
Private Sub GenerateCheque(Ar As CARMas, Tot() As Double, Details As CFieldList, alngX() As Long)
Dim strFormat As String
Dim strPrint As String
Dim TempX As Double
Dim tempY As Double
Dim i As Integer
Dim PlusAmt As Double
Dim MinusAmt As Double
Dim PrintDate As Date
Dim SentDate As Date
Dim SumCheque As Double
Dim TempDocDate As Date
Dim TempStr As String



If mcolParam("SHOW_DETAIL_RCP_FUTURE") = "Y" Then
   Vsp.NewPage
'   tempY = Vsp.CurrentY
Else
  Exit Sub
End If

   TempX = Vsp.CurrentX
   
     Vsp.TableBorder = tbNone
     Vsp.FontBold = True
      Vsp.FontSize = 16
      strFormat = VSP_CalTable("^3;", mdblWid, alngX)
      strPrint = glbCompanyName
      Call Vsp.AddTable(strFormat, "", strPrint)

      Vsp.FontSize = 14
      strFormat = VSP_CalTable("^20|>7|<13;", mdblWid, alngX)
      strPrint = "แบบฟอร์มขออนุมัติขาย|พนักงานขาย|" & Ar.SLMNAM & "(" & Ar.SLMCOD & ")" & ";"
      Call Vsp.AddTable(strFormat, "", strPrint)
      
         If mcolParam("PRINT_DATE") <= 0 Then
            PrintDate = Now
         Else
            PrintDate = mcolParam("PRINT_DATE")
         End If
   
      Vsp.FontSize = DETAIL_SIZE
      strFormat = VSP_CalTable(">4|<17|>4|<6|>4|<12|>4|10;", mdblWid, alngX)
      strPrint = "ชื่อลูกค้า|" & Ar.CUSNAM & "(" & Ar.CUSCOD & ")" & "|เครดิต|" & RealCredit & " วัน" & "|จังหวัด|" & GetProvince(Ar.ADDR02, ColProvinceMap) & "|วันที่  |" & DateToStringExtEx2(PrintDate)
      Call Vsp.AddTable(strFormat, "", strPrint)

      Vsp.CurrentX = TempX
     Call Vsp.DrawLine(Vsp.CurrentX, Vsp.CurrentY, Vsp.PageWidth - 500, Vsp.CurrentY)
      Vsp.FontBold = False
      
      Dim ARtFuture As CBkTrn
      Dim SumFuture As Double

      For Each ARtFuture In m_ReceiveFutureAmounts
               i = i + 1
               If i = 1 Then
                  Vsp.CurrentX = TempX
                  Vsp.Paragraph = "รายละเอียดเช็ครับล่วงหน้า"
                  strFormat = VSP_CalTable("^2|^4|^4|^4|^4|^4|^4|^4;", mdblWid, alngX)
                  strPrint = "ลำดับ|เลขที่เช็ค|วันที่เช็ค|รับล่วงหน้า|ใบสำคัญรับ|ยอดเช็ค|วันที่ส่งของ|เครดิต(วัน);"
                  Call Vsp.AddTable(strFormat, "", strPrint)
               End If
   
                SumFuture = SumFuture + ARtFuture.AMOUNT
                
               If ARtFuture.AMOUNT > 0 Then
                  strFormat = VSP_CalTable(">2|^4|^4|^4|^4|>4|^4|^4;", mdblWid, alngX)
               
                 TempStr = CalculateCreditCompare(ARtFuture.CHQNUM, ARtFuture.CHQDAT, TempDocDate)
            
                  strPrint = i & "|" & ARtFuture.CHQNUM & "|" & DateToStringExtEx2(ARtFuture.CHQDAT) & "|" & DateToStringExtEx2(ARtFuture.GETDAT) & "|" & ARtFuture.VOUCHER & "|" & FormatNumber(ARtFuture.AMOUNT) & "|" & DateToStringExtEx2(TempDocDate) & "|" & CalculateCreditCompare(ARtFuture.CHQNUM, ARtFuture.CHQDAT, TempDocDate) & ";"
                  Call Vsp.AddTable(strFormat, "", strPrint)
                  
                  SumCheque = SumCheque + ARtFuture.AMOUNT
               End If

      Next ARtFuture
      Vsp.FontBold = True
      strPrint = "|รวม||||" & FormatNumber(SumCheque) & "||;"
      Call Vsp.AddTable(strFormat, "", strPrint)
   Vsp.Paragraph = ""
    Call Vsp.DrawLine(Vsp.CurrentX, Vsp.CurrentY, Vsp.PageWidth - 500, Vsp.CurrentY)
    
'    If tempY > Vsp.CurrentY Then
'      Call GenerateCheque(Ar, Tot(), m_Details1, alngX())
'    End If
End Sub
Private Function CalculateCreditCompare(CHQNUM As String, CHQDAT As Date, DOCDAT As Date) As Double
Dim TempBkTrn As CBkTrn
   Set TempBkTrn = GetObject("CBkTrn", m_ChqnumDocDate, CHQNUM, False)
   If Not (TempBkTrn Is Nothing) Then
      If TempBkTrn.DOCDAT > 0 Then
         CalculateCreditCompare = DateDiff("D", TempBkTrn.DOCDAT, CHQDAT)
         DOCDAT = TempBkTrn.DOCDAT
      End If
   Else
      CalculateCreditCompare = 0
   End If
End Function

