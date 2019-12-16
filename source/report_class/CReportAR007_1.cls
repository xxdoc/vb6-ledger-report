VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CReportAR007_1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'User Group Report
Implements CReportInterface

Private Const MODULE_NAME = "CReportAR006"
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

Private m_CustomerTypes As Collection
Private m_CusCodeCall As Collection
Private m_CusCodeAmountCall As Collection

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
   Set m_CusCodeCall = New Collection
   Set m_CusCodeAmountCall = New Collection
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
   Set m_CusCodeCall = Nothing
   Set m_CusCodeAmountCall = Nothing
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

Dim TempData As CCustomer
Set TempData = New CCustomer
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

        Call Cf.SetFieldValue(1, "^", "เดือน", "^")
        Call m_Headers1.AddField(Cf)
        
        For Each TempData In m_CusCodeCall
                Call Cf.SetFieldValue(2, "^", Left(TempData.CUSTOMER_NAME, 10), ">")
                Call m_Headers1.AddField(Cf)
                TempWidth = TempWidth + 2
        Next TempData
         
        Call Cf.SetFieldValue(1, "^", "รวม", ">")
        Call m_Headers1.AddField(Cf)
        
        Call Cf.SetFieldValue(1, "^", "", "^")
        Call m_Headers0.AddField(Cf)
        
        Call Cf.SetFieldValue(TempWidth, "^", "ชื่อลูกค้า", "<")
        Call m_Headers0.AddField(Cf)
        
        Call Cf.SetFieldValue(1, "^", "", "^")
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
      
   FirstDate = mcolParam("FROM_DOC_DATE")
   LastDate = mcolParam("TO_DOC_DATE")
   
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
   Call LoadCusCode(Nothing, m_CusCodeCall, mcolParam("CUSTOMER_TYPE_ID"))
   Call LoadCusCodeAmount(Nothing, m_CusCodeAmountCall, mcolParam("FROM_DOC_DATE"), mcolParam("TO_DOC_DATE"))
   '-------------------------------- ยกมา --------------------------------
   Call LoadARAmountByCust2(Nothing, m_ArCusAmounts, -1, DateAdd("D", -1, FirstDate), "('3','4')")    'ยอดซื้อ ถึงวันที่
   Call LoadARAmountByCust2(Nothing, m_ArCusAmountCns, -1, DateAdd("D", -1, FirstDate), "('5')")     'ยอดซื้อ ถึงวันที่       ลดหนี้
   
   Call LoadARAmountByCust2(Nothing, m_ArCusAmountCashs, FirstDate, LastDate, "('9')")    'ยอดจ่ายเงินสด ถึงวันที่
   
   Call LoadPaidAmountByCust2(Nothing, m_ArCusPaids, -1, DateAdd("D", -1, FirstDate), -1, -1, "('3','4')")           'ยอดจ่ายถึงวันที่
   Call LoadPaidAmountByCust2(Nothing, m_ArCusPaidCns, -1, DateAdd("D", -1, FirstDate), -1, -1, "('5')")           'ยอดจ่ายถึงวันที่    ลดหนี้
   
   Call LoadPaidAmountByCust2(Nothing, m_PaidAllAmounts, -1, -1, -1, DateAdd("D", -1, FirstDate), "('3','4')")         'ยอดตัดบิลทั้งหมดของบิลภายในวันที่กำหนด
   Call LoadPaidAmountByCust2(Nothing, m_PaidAllAmountCns, -1, -1, -1, DateAdd("D", -1, FirstDate), "('5')")          'ยอดตัดบิลทั้งหมดของบิลภายในวันที่กำหนด ลดหนี้
   '-------------------------------- ยกมา --------------------------------
   
   '-------------------------------- เคลื่อนไหว --------------------------------
   Call LoadARAmountByCust2(Nothing, m_ArCusAmountTo, FirstDate, LastDate, "('3','4')")     'ยอดซื้อ ถึงวันที่
   Call LoadARAmountByCust2(Nothing, m_ArCusAmountToCn, FirstDate, LastDate, "('5')")     'ลดหนี้ ถึงวันที่
   
   Call LoadPaidAmountByCust2(Nothing, m_ArCusPaidTo, FirstDate, LastDate, -1, -1, "('3','4')")           'ยอดจ่ายถึงวันที่
   Call LoadPaidAmountByCust2(Nothing, m_ArCusPaidToCn, FirstDate, LastDate, -1, -1, "('5')")           'ยอดจ่ายถึงวันที่
   '-------------------------------- เคลื่อนไหว --------------------------------
   
   '-------------------------------- จ่ายเช็ค --------------------------------
'   Call LoadBktChqAmountBySup(Nothing, m_5, , , FirstDate, LastDate)                                                                                      'ยอดจ่ายจริง DUE เดือน4
'   Call LoadBktChqAmountBySup(Nothing, m_6, , LastDate, DateAdd("D", 1, LastDate))                                                                        'ยอดจ่าย POST DUE หลังเดือน4
'
'   Call LoadBktChqAmountBySup(Nothing, m_8, FirstDate, LastDate, FirstDate, LastDate)                                                       'จ่าย 4 DUE 4
'   Call LoadBktChqAmountBySup(Nothing, m_9, FirstDate, LastDate, DateAdd("D", 1, LastDate))                                        'จ่าย 4 DUEหลัง 4
'
'   Call LoadBktChqAmountBySup(Nothing, m_10, -1, DateAdd("D", -1, FirstDate), FirstDate, LastDate)                              'จ่ายก่อน 4 DUE 4
'   Call LoadBktChqAmountBySup(Nothing, m_11, -1, DateAdd("D", -1, FirstDate), DateAdd("D", 1, LastDate), -1)                             'จ่ายก่อน 4 DUEหลัง 4
   '-------------------------------- จ่ายเช็ค --------------------------------
   
   FromDate = "จากวันที่ " & EmptyToString(DateToStringExtEx2(mcolParam("FROM_DOC_DATE")), "N/A")
   ToDate = " ถึง วันที่ " & EmptyToString(DateToStringExtEx2(mcolParam("TO_DOC_DATE")), "N/A")
   
   MODULE_DESC = glbCompanyName & vbCrLf & mcolParam("REPORT_NAME") & _
                                       vbCrLf & FromDate & " " & ToDate & " "
   
  MODULE_DESC = MODULE_DESC & " สรุปยอดขาย กลุ่ม " & EmptyToString(mcolParam("CUSTOMER_TYPE_NAME"), "N/A")

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
Dim Count As Long
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

Dim ItemCount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CCustomer
Set TempData = New CCustomer

Dim TempFromdate As Date
Dim TempToDate As Date
Dim SumAmount As Long
Dim Artrn As CARTrn

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
   Count = 0
    TempFromdate = mcolParam("FROM_DOC_DATE")
    TempToDate = mcolParam("TO_DOC_DATE")
    While (TempFromdate <= TempToDate)
         For j = 1 To UBound(Total1)
             Total1(j) = 0
         Next j
                K = 2
                SumAmount = 0
            For Each TempData In m_CusCodeCall
                Set Artrn = GetARTrn(m_CusCodeAmountCall, TempData.CUSTOMER_CODE & "-" & Format("00", Month(TempFromdate)) & Year(TempFromdate))
                SumAmount = SumAmount + Artrn.AMOUNT
                Total1(K) = Artrn.AMOUNT
                Total2(K) = Total2(K) + Total1(K)
                K = K + 1
            Next TempData
                Total1(K) = SumAmount
                Total2(K) = Total2(K) + Total1(K)
                If SumAmount <> 0 Then
                    Call GenerateFooter2("" & (Month(TempFromdate) & "/" & (Year(TempFromdate) + 543)), Total1, m_Details1)                        'พิมพ์ยอดในแต่ละเดือน
                    Call m_Details1.GetString(2, TempStr1, TempStr2)
                    strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
                    Call Vsp.AddTable(strFormat, "", TempStr2)
                End If
                TempFromdate = DateAdd("M", 1, TempFromdate)
                Count = Count + 1
            Wend
                Call GenerateFooter2("", Total2, m_Details1)                                                 'พิมพ์รวมยอดของแต่ละบริษัท
                Call m_Details1.GetString(2, TempStr1, TempStr2)
                strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
                Vsp.FontBold = True
                Call Vsp.AddTable(strFormat, "", TempStr2)
                Vsp.FontBold = False
                
                For j = 1 To UBound(Total2)
                    Total2(j) = Total2(j) / Count
                Next j
                
                Call GenerateFooter2("เฉลี่ย/เดือน", Total2, m_Details1)                                                 'พิมพ์ยอดเฉลี่ย
                Call m_Details1.GetString(2, TempStr1, TempStr2)
                strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
                Call Vsp.AddTable(strFormat, "", TempStr2)

      Set BodyCf = Nothing
            
    Vsp = ""
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
         TempStr = Txt1
      ElseIf j = 2 Then
         TempStr = Txt2
      ElseIf j = 3 Then
         TempStr = Txt3
      ElseIf j > 3 Then
        j = 4
         TempStr = FormatNumber(Tot(j))
      End If
      Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.Align, TempStr, HeadCf.BodyAlign)
      Call Details.AddField(BodyCf)
   Next HeadCf
      
   Set BodyCf = Nothing
End Sub
Private Sub GenerateFooter2(Txt1 As String, Tot() As Double, Details As CFieldList)
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
      Else
         TempStr = FormatNumber(Tot(j))
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
