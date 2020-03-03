VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CReportAR005_4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'User Group Report
Implements CReportInterface

Private Const MODULE_NAME = "CReportAR005_4"
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
Private m_ReceiveAmounts As Collection
Private m_ReceiveAllAmounts As Collection
Private m_CnAllAmounts As Collection
Private m_CustomerTypes As Collection
Private Temp_AnalyzeCustomer As Collection
Private FirstDate As Date
Private LastDate As Date
Private DateCountMonth As Long
Private TempPosition As Long
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
   Set m_ReceiveAmounts = New Collection
   Set m_ReceiveAllAmounts = New Collection
   Set m_CnAllAmounts = New Collection
   Set Temp_AnalyzeCustomer = New Collection
   Set m_CustomerTypes = New Collection
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
   Set m_ReceiveAllAmounts = Nothing
   Set m_CnAllAmounts = Nothing
   Set Temp_AnalyzeCustomer = Nothing
   Set m_CustomerTypes = Nothing
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
Dim TempDate As Long
Dim FirstDayOfMonth As Date

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
   Call Cf.SetFieldValue(5, "^", "ชื่อลูกค้า", "^")
   Call m_Headers0.AddField(Cf)

   Call Cf.SetFieldValue(5, "^", "", "<")
   Call m_Headers1.AddField(Cf)
   
   FirstDayOfMonth = FirstDate
   TempDate = 1
   For j = 1 To DateCountMonth
      If Weekday(FirstDayOfMonth) = 1 Or j = DateCountMonth Then
         If Weekday(FirstDayOfMonth) = 7 Then
            j = j - 1
         ElseIf Weekday(FirstDayOfMonth) = 1 Then
            j = j - 2
         End If
         If j > 0 Then
            If TempDate = j Then
               Call Cf.SetFieldValue(2, "^", j & "/" & Month(FirstDayOfMonth) & "/" & Year(FirstDayOfMonth), "^")
            Else
               Call Cf.SetFieldValue(2, "^", TempDate & "-" & j & "/" & Month(FirstDayOfMonth) & "/" & Year(FirstDayOfMonth), "^")
            End If
            Call m_Headers0.AddField(Cf)
         
            Call Cf.SetFieldValue(2, "^", "TARGET", ">")
            Call m_Headers1.AddField(Cf)
         End If
         j = j + 2
         TempDate = j + 1
      End If
      FirstDayOfMonth = DateAdd("d", 1, FirstDayOfMonth)
   Next j
   
   Call Cf.SetFieldValue(2, "^", "รวม", "^")
   Call m_Headers0.AddField(Cf)

   Call Cf.SetFieldValue(2, "^", "", ">")
   Call m_Headers1.AddField(Cf)
   
   Call Cf.SetFieldValue(2, "^", "วันที่ขาย", "^")
   Call m_Headers0.AddField(Cf)

   Call Cf.SetFieldValue(2, "^", "", "^")
   Call m_Headers1.AddField(Cf)
   
   Call Cf.SetFieldValue(2, "^", "วันที่ครบดิว", "^")
   Call m_Headers0.AddField(Cf)

   Call Cf.SetFieldValue(2, "^", "", "^")
   Call m_Headers1.AddField(Cf)
   
   Call Cf.SetFieldValue(2, "^", "วันที่ปัจจุบัน", "^")
   Call m_Headers0.AddField(Cf)

   Call Cf.SetFieldValue(2, "^", "", "^")
   Call m_Headers1.AddField(Cf)
   
   Call Cf.SetFieldValue(1, "^", "เกินดิว", "^")
   Call m_Headers0.AddField(Cf)

   Call Cf.SetFieldValue(1, "^", "(วัน)", ">")
   Call m_Headers1.AddField(Cf)
   
   Call Cf.SetFieldValue(2, "^", "เลขที่ INV.", "^")
   Call m_Headers0.AddField(Cf)

   Call Cf.SetFieldValue(2, "^", "", "^")
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
   
   Call LoadCustomerType(Nothing, m_CustomerTypes)
   Call LoadReceiveAmountByBill(Nothing, m_ReceiveAmounts, -1, -1)
   Call LoadReceiveAmountByBill(Nothing, m_ReceiveAllAmounts, -1, -1)
   Call LoadARCNAmountByBill(Nothing, m_CnAllAmounts, DateAdd("D", 1, mcolParam("TO_DOC_DATE")), -1, -1, mcolParam("TO_DOC_DATE"))         'ต้องบวกยอดยกมาเพิ่ม ยอด CN ทั้งหมดของบิลหลัง CN แต่ เอกสาร LINK เป็นในช่วงยกมา
   Call LoadAnalyzeCustomer(Nothing, Temp_AnalyzeCustomer)
   
   Call GetFirstLastDate(DateSerial(mcolParam("YEAR_NO") - 543, mcolParam("MONTH_ID"), 1), FirstDate, LastDate)
   DateCountMonth = DateDiff("D", FirstDate, LastDate) + 1
   
   FromDate = "จากวันที่ " & EmptyToString(DateToStringExtEx2(mcolParam("FROM_DOC_DATE")), "N/A")
   ToDate = " ถึง วันที่ " & EmptyToString(DateToStringExtEx2(mcolParam("TO_DOC_DATE")), "N/A")
   ToPaidDate = "   วันที่พิมพ์ " & EmptyToString(DateToStringExtEx2(mcolParam("PRINT_DATE")), "N/A")
  
   'glbEnterPrise.ENTERPRISE_NAME & " , " & glbEnterPrise.BRANCH_NAME
   MODULE_DESC = glbCompanyName & vbCrLf & mcolParam("REPORT_NAME") & "เดือน " & IntToThaiMonth(mcolParam("MONTH_ID")) & "/" & mcolParam("YEAR_NO") & _
                                       vbCrLf & FromDate & ToDate & ToPaidDate & " จากลูกค้า " & EmptyToString(mcolParam("FROM_CUSTOMER_CODE"), "N/A") & " ถึงลูกค้า " & EmptyToString(mcolParam("TO_CUSTOMER_CODE"), "N/A") & " " & GenerateCreditString
   
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
Dim Ari As CARRcIt
Dim Ari2 As CARRcIt
Dim Apt1 As CARTrn
Dim ARt2 As CARTrn
Dim St As CIsTab
Dim Ar As COESLM
Dim Width As Double
Dim ARt As CARTrn
Dim Ac As CAnalyzeCustomer
Dim PaidBalance As Double 'ชำระแล้วยกมา
Dim TempRs As ADODB.Recordset
Dim HaveHeader As Boolean
Dim TempDate As Long
Dim TempDocDate As Date
Dim FirstDayOfMonth As Date
Dim PrintDate As Date
Dim TempAmt As Double
Dim Sum As Double
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
      
      If mcolParam("SHOW_EXCEL") = "Y" Then
         Vsp.Paragraph = MODULE_DESC
         Call printHeader
      End If
      
      Set BodyCf = New CReportField
      
      Set Ar = New COESLM
      Ar.FROM_SALE_CODE = mcolParam("FROM_SALE_CODE")
      Ar.TO_SALE_CODE = mcolParam("TO_SALE_CODE")
      Ar.OrderBy = mcolParam("ORDER_TYPE")
      Ar.OrderType = mcolParam("ORDER_BY")
      Call glbDaily.QuerySale(Ar, Rs, iCount, IsOK, glbErrorLog)

      While Not Rs.EOF
         For j = 1 To UBound(Total1)
            Total1(j) = 0
         Next j

         Call m_Details2.ClearField
         Call Ar.PopulateFromRS(1, Rs)
         
         HaveHeader = True
         
         If True Then  'มียอดค้างจ่าย
            '===Head
'            Set HeadCf = m_Headers1.Fields(1)
'            TempStr = Ar.SLMNAM & "  -  " & Ar.SLMCOD
'            Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
'            Call m_Details1.AddField(BodyCf)
'
'            Call m_Details1.GetString(1, TempStr1, TempStr2)
'            strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
'            Call Vsp.AddTable(strFormat, "", TempStr2)

            PrevKey1 = Ar.SLMNAM & "  -  " & Ar.SLMCOD
            '=== Detail
            Set ARt = New CARTrn
            ARt.SLMCOD = Ar.SLMCOD
            ARt.FROM_DOC_DATE = mcolParam("FROM_DOC_DATE")
            ARt.TO_DOC_DATE = mcolParam("TO_DOC_DATE")
            If mcolParam("CUSTOMER_TYPE") > 0 Then
               Set St = GetItemFromListIndex(m_CustomerTypes, mcolParam("CUSTOMER_TYPE"))
               ARt.CUSTYP = St.TYPCOD
            Else
               ARt.CUSTYP = ""
            End If
            ARt.FROM_CUSTOMER_CODE = mcolParam("FROM_CUSTOMER_CODE")
            ARt.TO_CUSTOMER_CODE = mcolParam("TO_CUSTOMER_CODE")
            ARt.RECTYP = ""
            ARt.RecTypeSet = "('3', '4', '5')"
            ARt.OrderBy = 4
            ARt.OrderType = 1
            Call ARt.QueryData(1, TempRs, iCount)

            While Not TempRs.EOF
                  Call m_Details1.ClearField
                  
                  Call ARt.PopulateFromRS(1, TempRs)
                  Set Ari = GetARRcpItem(m_ReceiveAmounts, ARt.DOCNUM)
                  Set Ari2 = GetARRcpItemEx(m_ReceiveAllAmounts, ARt.DOCNUM)
                  Set ARt2 = GetARTrn(m_CnAllAmounts, ARt.DOCNUM)
                  'รับชำระแล้วยกมาเท่ากับ APT.PAYAMT - ยอดรับชำระทั้งหมดของใบสำคัญจ่าย
                  PaidBalance = ARt.RCVAMT - Ari2.RCVAMT - ARt2.AMOUNT
                  If (ARt.AMOUNT) > (Ari.RCVAMT + PaidBalance) Then  'รับยังไม่ครบ
                     Set HeadCf = m_Headers1.Fields(1)
                     TempStr = ARt.CUSNAM
                     Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
                     Call m_Details1.AddField(BodyCf)

                     Set Ac = GetAnalyzeCustomer(Temp_AnalyzeCustomer, Trim(ARt.DOCNUM))
                     Sum = 0
                     TempAmt = 0
                     FirstDayOfMonth = FirstDate
                     TempDate = 1
                      i = 2
                     For j = 1 To DateCountMonth
                        If Weekday(FirstDayOfMonth) = 1 Or j = DateCountMonth Then
                           If Weekday(FirstDayOfMonth) = 7 Then
                              j = j - 1
                           ElseIf Weekday(FirstDayOfMonth) = 1 Then
                              j = j - 2
                           End If
                           If j > 0 Then
                              Set HeadCf = m_Headers1.Fields(i)
                              If TempDate <= Day(Ac.DATE_OF_PAYMENT) And Day(Ac.DATE_OF_PAYMENT) <= j And Month(FirstDayOfMonth) = Month(Ac.DATE_OF_PAYMENT) And Year(FirstDayOfMonth) = Year(Ac.DATE_OF_PAYMENT) Then
                                 TempAmt = (ARt.AMOUNT) - (Ari.RCVAMT + PaidBalance)
                                 If ARt.RECTYP = 5 Then 'ลดหนี้
                                    TempAmt = -1 * TempAmt
                                 End If
                                 Total1(i) = Total1(i) + TempAmt
                                 Total2(i) = Total2(i) + TempAmt
                                 Sum = Sum + TempAmt
                                 TempStr = FormatNumberToNull(TempAmt)
                              Else
                                 TempStr = ""
                              End If
                              Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
                              Call m_Details1.AddField(BodyCf)
                              i = i + 1
                           End If
                           j = j + 2
                           TempDate = j + 1
                        End If
                        FirstDayOfMonth = DateAdd("d", 1, FirstDayOfMonth)
                     Next j
                     'รวม
                     TempPosition = i
                     Set HeadCf = m_Headers1.Fields(i)
                     TempStr = FormatNumberToNull(Sum)
                     Total1(i) = Total1(i) + Sum
                     Total2(i) = Total2(i) + Sum
                     Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
                     Call m_Details1.AddField(BodyCf)
                     'วันที่ขาย
                      i = i + 1
                     Set HeadCf = m_Headers1.Fields(i)
                     TempStr = DateToStringExtEx2(ARt.DOCDAT)
                     Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
                     Call m_Details1.AddField(BodyCf)
                     'วันที่ครบดิว
                     i = i + 1
                     Set HeadCf = m_Headers1.Fields(i)
                     If Val(mcolParam("CREDIT")) > 0 Then
                       TempDocDate = DateToStringExtEx2(DateAdd("D", Val(mcolParam("CREDIT")), ARt.DOCDAT))
                     Else
                        TempDocDate = DateToStringExtEx2(DateAdd("D", 90, ARt.DOCDAT))
                     End If
                     TempStr = DateToStringExtEx2(TempDocDate)
                     Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
                     Call m_Details1.AddField(BodyCf)
                     'วันที่ปัจจุบัน
                     If mcolParam("PRINT_DATE") <= 0 Then
                        PrintDate = Now
                     Else
                        PrintDate = mcolParam("PRINT_DATE")
                     End If
                     i = i + 1
                     Set HeadCf = m_Headers1.Fields(i)
                     TempStr = DateToStringExtEx2(PrintDate)
                     Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
                     Call m_Details1.AddField(BodyCf)
                     'วันที่เกินดิว(วัน)
                     i = i + 1
                     Set HeadCf = m_Headers1.Fields(i)
                     TempStr = FormatNumberInt(DateDiff("D", TempDocDate, PrintDate))
                     Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
                     Call m_Details1.AddField(BodyCf)
                     'เลขที่ INV.
                     i = i + 1
                     Set HeadCf = m_Headers1.Fields(i)
                     TempStr = ARt.DOCNUM
                     Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
                     Call m_Details1.AddField(BodyCf)
                     
                     If mcolParam("SHOW_ALL") = "Y" Or Sum <> 0 Then
                        '===Head
                        If HaveHeader Then
                           Call GenerateHeader(PrevKey1, m_Details2)
                           Call m_Details2.GetString(2, TempStr1, TempStr2)
                           strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
                           Vsp.FontBold = True
                           Call Vsp.AddTable(strFormat, "", TempStr2)
                           Vsp.FontBold = False
                        End If
                        
                        HaveHeader = False
                        Call m_Details1.GetString(1, TempStr1, TempStr2)
                        strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
                        Call Vsp.AddTable(strFormat, "", TempStr2)
                     End If
                  End If
               TempRs.MoveNext
            Wend
            '===
            If Total1(TempPosition) <> 0 Then
               Call GenerateFooter("รวม ", PrevKey1, "", Total1, m_Details1)
               Call m_Details1.GetString(2, TempStr1, TempStr2)
               strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
               Vsp.FontBold = True
               Call Vsp.AddTable(strFormat, "", TempStr2)
               Vsp.FontBold = False
            End If
            
            HaveHeader = True
          End If
         Rs.MoveNext
      Wend
      
      Call GenerateFooter("รวม ", "", "", Total2, m_Details1)
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
Private Sub GenerateFooter(Txt1 As String, Txt2 As String, Txt3 As String, Tot() As Double, Details As CFieldList)
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
         TempStr = Txt1 & "     " & Txt2
      ElseIf j > TempPosition Then
         TempStr = ""
      Else
         TempStr = FormatNumberToNull(Tot(j))
      End If
      
      Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.Align, TempStr, HeadCf.BodyAlign)
      Call Details.AddField(BodyCf)
   Next HeadCf
      
   Set BodyCf = Nothing
End Sub
Private Sub GenerateHeader(Txt1 As String, Details As CFieldList)
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
Private Function GenerateCreditString() As String
   If Val(mcolParam("CREDIT")) > 0 Then
      GenerateCreditString = GenerateCreditString & "***เครดิต " & Val(mcolParam("CREDIT")) & " วัน"
   Else
      GenerateCreditString = GenerateCreditString & "***เครดิต 90 วัน"
   End If
End Function


