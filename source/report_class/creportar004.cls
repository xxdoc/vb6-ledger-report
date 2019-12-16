VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CReportAR004"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'User Group Report
Implements CReportInterface

Private Const MODULE_NAME = "CReportAR004"
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
Private m_Left As Double
Private m_YearSeqs As Collection

Private m_CustomerTypes As Collection
Private m_ReceiveAmounts As Collection
Private m_ReceiveAllAmounts As Collection
Private m_MaxMins As Collection
Private m_ArCustAmounts As Collection
Private m_ArCustPaids As Collection
Private m_RealCreditColl  As Collection
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
   Set m_CustomerTypes = New Collection
   Set m_ReceiveAmounts = New Collection
   Set m_MaxMins = New Collection
   Set m_ArCustAmounts = New Collection
   Set m_ArCustPaids = New Collection
   Set m_ReceiveAllAmounts = New Collection
   Set m_RealCreditColl = New Collection
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
   Set m_CustomerTypes = Nothing
   Set m_ReceiveAmounts = Nothing
   Set m_MaxMins = Nothing
   Set m_ReceiveAllAmounts = Nothing
   Set m_RealCreditColl = Nothing
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
   
   Set Cf = New CReportField
      
   '===
   Call Cf.SetFieldValue(1, "^", "NO.", "^")
   Call m_Headers0.AddField(Cf)
   Call Cf.SetFieldValue(2, "^", "", "^")
   Call m_Headers0.AddField(Cf)
   Call Cf.SetFieldValue(7, "^", "", "^")
   Call m_Headers0.AddField(Cf)
   Call Cf.SetFieldValue(6, "^", "จะครบกำหนด", ">")
   Call m_Headers0.AddField(Cf)
   Call Cf.SetFieldValue(8, "^", "เกินกำหนด", ">")
   Call m_Headers0.AddField(Cf)
   Call Cf.SetFieldValue(2, "^", "", "^")
   Call m_Headers0.AddField(Cf)
   Call Cf.SetFieldValue(2, "^", "เครดิต", "^")
   Call m_Headers0.AddField(Cf)
   '===
   
   Call Cf.SetFieldValue(1, "^", "NO.", "^")
   Call m_Headers1.AddField(Cf)
   Call Cf.SetFieldValue(2, "^", "รหัสลูกค้า", "^")
   Call m_Headers1.AddField(Cf)
   Call Cf.SetFieldValue(7, "^", "ชื่อลูกค้า", "<")
   Call m_Headers1.AddField(Cf)
   
   If mcolParam("INTERVAL_TYPE") = 1 Then
      Call Cf.SetFieldValue(2, "^", "> 60 วัน", ">")
      Call m_Headers1.AddField(Cf)
      Call Cf.SetFieldValue(2, "^", "30-60 วัน", ">")
      Call m_Headers1.AddField(Cf)
      Call Cf.SetFieldValue(2, "^", "< 30 วัน", ">")
      Call m_Headers1.AddField(Cf)
      
      Call Cf.SetFieldValue(2, "^", "1-15 วัน", ">")
      Call m_Headers1.AddField(Cf)
      Call Cf.SetFieldValue(2, "^", "16-30 วัน", ">")
      Call m_Headers1.AddField(Cf)
      Call Cf.SetFieldValue(2, "^", "31-60 วัน", ">")
      Call m_Headers1.AddField(Cf)
      Call Cf.SetFieldValue(2, "^", "> 60 วัน", ">")
      Call m_Headers1.AddField(Cf)
   Else
      Call Cf.SetFieldValue(2, "^", "> 60 วัน", ">")
      Call m_Headers1.AddField(Cf)
      Call Cf.SetFieldValue(2, "^", "30-60 วัน", ">")
      Call m_Headers1.AddField(Cf)
      Call Cf.SetFieldValue(2, "^", "< 30 วัน", ">")
      Call m_Headers1.AddField(Cf)
      
      Call Cf.SetFieldValue(2, "^", "1-90 วัน", ">")
      Call m_Headers1.AddField(Cf)
      Call Cf.SetFieldValue(2, "^", "91-180 วัน", ">")
      Call m_Headers1.AddField(Cf)
      Call Cf.SetFieldValue(2, "^", "181-365 วัน", ">")
      Call m_Headers1.AddField(Cf)
      Call Cf.SetFieldValue(2, "^", "> 365 วัน", ">")
      Call m_Headers1.AddField(Cf)
   End If
   
   Call Cf.SetFieldValue(2, "^", "รวมหนี้ค้าง", ">")
   Call m_Headers1.AddField(Cf)
   
   Call Cf.SetFieldValue(2, "^", "เครดิต", "^")
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

Private Function initDoc() As Boolean
Dim strDate As String
Dim TempStr1 As String
Dim TempStr2 As String
Dim TempStr3 As String
Dim FromDate As String
Dim ToDate As String
Dim ToPaidDate As String

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
   
   If mcolParam("REAL_CREDIT_FLAG") = "Y" Then
      Call LoadRealCreditNo(Nothing, m_RealCreditColl)
   End If
   
   Call LoadCustomerType(Nothing, m_CustomerTypes, mcolParam("CUSTOMER_TYPE"))
   Call LoadReceiveAmountByBill(Nothing, m_ReceiveAmounts, -1, mcolParam("TO_PAY_DATE"))
   Call LoadReceiveAmountByBill(Nothing, m_ReceiveAllAmounts, -1, -1)
   If mcolParam("INTERVAL_TYPE") = 1 Then
      Call LoadDueDateInterval1(Nothing, m_MaxMins)
   Else
      Call LoadDueDateInterval2(Nothing, m_MaxMins)
   End If
   'Call LoadARAmountBycust(Nothing, m_ArCustAmounts, mcolParam("FROM_DOC_DATE"), mcolParam("TO_DOC_DATE"))
   Call LoadARAmountByCust(Nothing, m_ArCustAmounts, -1, mcolParam("TO_DOC_DATE"))
   Call LoadPaidAmountByCust(Nothing, m_ArCustPaids, -1, mcolParam("TO_PAY_DATE"))
   
   FromDate = "จากวันที่ " & EmptyToString(DateToStringExtEx2(mcolParam("FROM_DOC_DATE")), "N/A")
   ToDate = " ถึง วันที่ " & EmptyToString(DateToStringExtEx2(mcolParam("TO_DOC_DATE")), "N/A")
   ToPaidDate = "    วันที่ชำระเงินถึงวันที่ " & EmptyToString(DateToStringExtEx2(mcolParam("TO_PAY_DATE")), "N/A") & "    วันที่พิมพ์ " & EmptyToString(DateToStringExtEx2(mcolParam("PRINT_DATE")), "N/A")
   
   'glbEnterPrise.ENTERPRISE_NAME & " , " & glbEnterPrise.BRANCH_NAME
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

Private Sub GenerateDetail(alngX() As Long, i As Long, ARt As CARTrn, Tot1() As Double, Tot2() As Double, Tot3() As Double)
Dim HeadCf As CReportField
Dim TempStr As String
Dim TempStr1 As String
Dim TempStr2 As String
Dim BodyCf As CReportField
Dim strFormat As String
Dim Ari As CARRcIt
Dim Ari2 As CARRcIt
Dim DueCount As Long
Dim MM As CMaxMin
Dim j As Long
Dim Sum As Double
Dim PrintDate As Date
Dim PaidAmount As Double
Dim PaidBalance As Double 'ชำระแล้วยกมา
Dim TempRealCredit  As CRealCredit
   
   Set BodyCf = New CReportField
   
   If mcolParam("PRINT_DATE") <= 0 Then
      PrintDate = Now
   Else
      PrintDate = mcolParam("PRINT_DATE")
   End If
   
   Set TempRealCredit = GetObject("CRealCredit", m_RealCreditColl, Trim(ARt.DOCNUM), False)
   If TempRealCredit Is Nothing Then
      If mcolParam("NINETY_FLAG") = "Y" Then
         DueCount = DateDiff("D", DateAdd("D", 90, ARt.DOCDAT), PrintDate)
      Else
         DueCount = DateDiff("D", ARt.DUEDAT, PrintDate)
      End If
   Else
      DueCount = DateDiff("D", DateAdd("D", TempRealCredit.REAL_CREDIT, ARt.DOCDAT), PrintDate)
   End If
   Set Ari = GetARRcpItem(m_ReceiveAmounts, ARt.DOCNUM)
   Set Ari2 = GetARRcpItemEx(m_ReceiveAllAmounts, ARt.DOCNUM)

   'ชำระแล้วยกมาเท่ากับ Art.PAYAMT - ยอดชำระทั้งหมดของใบสำคัญจ่าย
   PaidBalance = ARt.RCVAMT - Ari2.RCVAMT
   If ARt.AMOUNT > (Ari.RCVAMT + PaidBalance) Then  'จ่ายยังไม่ครบ
      If ARt.RECTYP = 5 Then 'ลดหนี้
         ARt.AMOUNT = -1 * ARt.AMOUNT
      End If

      Set HeadCf = m_Headers1.Fields(1)
      TempStr = i
      Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
      Call m_Details1.AddField(BodyCf)

      Set HeadCf = m_Headers1.Fields(2)
      TempStr = DateToStringExtEx2(ARt.DUEDAT)
      Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
      Call m_Details1.AddField(BodyCf)

      Set HeadCf = m_Headers1.Fields(3)
      TempStr = DateToStringExtEx2(ARt.DOCDAT)
      Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
      Call m_Details1.AddField(BodyCf)
      
      j = 4
      Sum = 0
      For Each MM In m_MaxMins
         Set HeadCf = m_Headers1.Fields(j)
         If (MM.MIN < DueCount) And (DueCount <= MM.MAX) Then
            TempStr = FormatNumber(ARt.AMOUNT - (Ari.RCVAMT + PaidBalance), "-")
            
            Tot1(j) = Tot1(j) + (ARt.AMOUNT - (Ari.RCVAMT + PaidBalance))
            Tot2(j) = Tot2(j) + (ARt.AMOUNT - (Ari.RCVAMT + PaidBalance))
            Tot3(j) = Tot3(j) + (ARt.AMOUNT - (Ari.RCVAMT + PaidBalance))
            Sum = Sum + (ARt.AMOUNT - (Ari.RCVAMT + PaidBalance))
         Else
            TempStr = FormatNumber(0, "-")
         End If
         Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
         Call m_Details1.AddField(BodyCf)
                  
         j = j + 1
      Next MM
            
      Set HeadCf = m_Headers1.Fields(j)
      TempStr = FormatNumber(Sum, "-")
      Tot1(j) = Tot1(j) + (ARt.AMOUNT - (Ari.RCVAMT + PaidBalance))
      Tot2(j) = Tot2(j) + (ARt.AMOUNT - (Ari.RCVAMT + PaidBalance))
      Tot3(j) = Tot3(j) + (ARt.AMOUNT - (Ari.RCVAMT + PaidBalance))
      Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
      Call m_Details1.AddField(BodyCf)
      
      j = j + 1
      Set HeadCf = m_Headers1.Fields(j)
      TempStr = FormatNumberInt(DueCount)
      Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
      Call m_Details1.AddField(BodyCf)
            
      Call m_Details1.GetString(1, TempStr1, TempStr2)
      strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
'      Call Vsp.AddTable(strFormat, "", TempStr2)
   End If
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
Dim St As CIsTab
Dim AR As CARMas
Dim Width As Double
Dim ARt As CARTrn
Dim TempRs As ADODB.Recordset
Dim TempARt As CARTrn
Dim TempARRcIt As CARRcIt

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
   
      Set BodyCf = New CReportField
      i = 0
      For Each St In m_CustomerTypes
         For j = 1 To UBound(Total2)
            Total2(j) = 0
         Next j

         '===Head
         Call m_Details1.ClearField
         Width = m_Headers0.Fields(1).Width + m_Headers0.Fields(2).Width
         TempStr = St.TYPCOD
         Call BodyCf.SetFieldValue(Width, "^", TempStr)
         Call m_Details1.AddField(BodyCf)

         Width = m_Headers0.Fields(3).Width + m_Headers0.Fields(4).Width + m_Headers0.Fields(5).Width + m_Headers0.Fields(6).Width + m_Headers0.Fields(7).Width
         TempStr = St.TYPDES
         Call BodyCf.SetFieldValue(Width, "<", TempStr)
         Call m_Details1.AddField(BodyCf)

         Call m_Details1.GetString(1, TempStr1, TempStr2)
         strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
         Call Vsp.AddTable(strFormat, "", TempStr2)
         '===

         Set BodyCf = New CReportField
   
         Set AR = New CARMas
         AR.FROM_CUSTOMER_CODE = mcolParam("FROM_CUSTOMER_CODE")
         AR.TO_CUSTOMER_CODE = mcolParam("TO_CUSTOMER_CODE")
         AR.CUSTYP = St.TYPCOD
         AR.OrderBy = mcolParam("ORDER_TYPE")
         AR.OrderType = mcolParam("ORDER_BY")
         Call glbDaily.QueryCustomer(AR, Rs, iCount, IsOK, glbErrorLog)
   
         While Not Rs.EOF
            For j = 1 To UBound(Total1)
               Total1(j) = 0
            Next j

            Call m_Details1.ClearField
            Call AR.PopulateFromRS(1, Rs)
            Set TempARt = GetARTrn(m_ArCustAmounts, AR.CUSCOD)
            Set TempARRcIt = GetARRcpItem(m_ArCustPaids, AR.CUSCOD)
            If TempARt.AMOUNT > TempARRcIt.RCVAMT Then  'มียอดค้างจ่าย
               '===Head
               Width = m_Headers0.Fields(1).Width + m_Headers0.Fields(2).Width
               TempStr = AR.CUSCOD
               Call BodyCf.SetFieldValue(Width, "^", TempStr)
               Call m_Details1.AddField(BodyCf)

               Width = m_Headers0.Fields(3).Width + m_Headers0.Fields(4).Width + m_Headers0.Fields(5).Width + m_Headers0.Fields(6).Width + m_Headers0.Fields(7).Width
               TempStr = AR.CUSNAM
               Call BodyCf.SetFieldValue(Width, "<", TempStr)
               Call m_Details1.AddField(BodyCf)
   
               Call m_Details1.GetString(1, TempStr1, TempStr2)
               strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
'               Call Vsp.AddTable(strFormat, "", TempStr2)
               '===
   
               '=== Detail
               Set ARt = New CARTrn
               ARt.CUSCOD = AR.CUSCOD
               ARt.FROM_DUE_DATE = mcolParam("FROM_DUE_DATE")
               ARt.TO_DUE_DATE = mcolParam("TO_DUE_DATE")
               ARt.FROM_DOC_DATE = mcolParam("FROM_DOC_DATE")
               ARt.TO_DOC_DATE = mcolParam("TO_DOC_DATE")
               ARt.RECTYP = ""
               ARt.RecTypeSet = "('3', '4', '5')"
               ARt.OrderBy = 1
               ARt.OrderType = 1
               Call ARt.QueryData(1, TempRs, iCount)
               
               i = i + 1
               
               While Not TempRs.EOF
                     
                     Call m_Details1.ClearField
                     Call ARt.PopulateFromRS(1, TempRs)
      
                     Call GenerateDetail(alngX(), i, ARt, Total1, Total2, Total3)
                     
                     TempRs.MoveNext
                  Wend
   
                  Call GenerateFooter(Trim(Str(i)), AR.CUSCOD, AR.CUSNAM, Total1, m_Details1, AR.PAYTRM)
                  Call m_Details1.GetString(2, TempStr1, TempStr2)
                  strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
                  Call Vsp.AddTable(strFormat, "", TempStr2)
                  
                  Set ARt = Nothing
                  '===
               End If
   
            Rs.MoveNext
         Wend
   
         Call GenerateFooter("รวม", "", "", Total2, m_Details1, 0)
         Call m_Details1.GetString(2, TempStr1, TempStr2)
         strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
         Vsp.FontBold = True
         Call Vsp.AddTable(strFormat, "", TempStr2)
         Vsp.FontBold = False
   
         Set AR = Nothing
      Next St
      Set BodyCf = Nothing

      Call GenerateFooter("สุทธิ", "", "", Total3, m_Details1, 0)
      Call m_Details1.GetString(2, TempStr1, TempStr2)
      strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
      Vsp.FontBold = True
      Call Vsp.AddTable(strFormat, "", TempStr2)
      Vsp.FontBold = False
      
      Call GenerateEnd
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
         TempStr = Txt1
      ElseIf j = 2 Then
         TempStr = Txt2
      ElseIf j = 3 Then
         TempStr = Txt3
      ElseIf j = m_Headers1.Fields.Count Then
         TempStr = DateCount
      Else
         TempStr = FormatNumber(Tot(j))
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
Private Sub GenerateEnd()
   If mcolParam("REAL_CREDIT_FLAG") = "Y" Then
      Vsp = "***ใช้เครดิตตามตารางเครดิตจริง"
   End If
   If mcolParam("NINETY_FLAG") = "Y" Then
      Vsp = "***เครดิต 90 วัน(ถ้าไม่มีในตารางจริง)"
   End If
End Sub

