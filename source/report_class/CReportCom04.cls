VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CReportCom04"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'User Group Report
Implements CReportInterface

Private Const MODULE_NAME = "CReportCOM04"
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
Private TempRs As ADODB.Recordset

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
Private m_StcrdBySale As Collection
Private m_AreaCod As Collection
Private m_AreaFromIV As Collection
Private m_ConditionCommiss1 As Collection
Private m_ConditionCommiss2 As Collection
Private m_ConditionCommiss3 As Collection
Private m_ConditionCommiss4 As Collection

Private m_cusFromArea(20) As Collection
Private nonidentifyArea As Boolean
Dim Stb As CCommissMasterArea
Private coll_Minus As Collection
Private m_IVincomplete0 As Collection
Private m_IVincomplete12 As Collection
Private m_ReDocdat As Collection
Private temp_Area As CCommissMasterArea

Private collTotal As Collection
Private m_saleChartHead As Collection
Private m_saleChartChild As Collection
Private temp_Head As CCommissionChart
 Private temp_Child As CCommissionChart
 Private temp_Child2 As CCommissionChart
 Private m_SaleName As Collection
 
 Private temp_GoodsGroup As CGoodsGroup
 Private temp_ChildGoodsGroup As CGoodsGroup
 Private m_GoodsGroup As Collection
 
Dim L As Long
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
   Set temp_Area = New CCommissMasterArea
   Set temp_Head = New CCommissionChart
   Set temp_Child = New CCommissionChart
   Set temp_Child2 = New CCommissionChart
   Set collTotal = New Collection
   Set m_saleChartHead = New Collection
   Set m_saleChartChild = New Collection
   Set m_SaleName = New Collection
   Set m_ConditionCommiss1 = New Collection
   Set temp_GoodsGroup = New CGoodsGroup
   Set temp_ChildGoodsGroup = New CGoodsGroup
   Set m_GoodsGroup = New Collection
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
   Set m_StcrdBySale = Nothing
   Set m_AreaCod = Nothing
   Set m_AreaFromIV = Nothing
   Set m_ConditionCommiss1 = Nothing
      Set m_ConditionCommiss2 = Nothing
   Set m_ConditionCommiss3 = Nothing
   Set m_ConditionCommiss4 = Nothing
   For L = 1 To UBound(m_cusFromArea)
            Set m_cusFromArea(L) = Nothing
   Next L
  Set Stb = Nothing
  Set coll_Minus = Nothing
  Set m_IVincomplete0 = Nothing
  Set m_IVincomplete12 = Nothing
  Set m_ReDocdat = Nothing
  Set temp_Area = Nothing
  Set collTotal = Nothing
  Set m_saleChartHead = Nothing
  Set m_saleChartChild = Nothing
  Set temp_Head = Nothing
   Set temp_Child = Nothing
   Set temp_Child2 = Nothing
   Set m_SaleName = Nothing
   Set temp_GoodsGroup = Nothing
   Set temp_ChildGoodsGroup = Nothing
   Set m_GoodsGroup = Nothing
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
'Dim TempRs As ADODB.Recordset
Dim Amt As Double
Dim iCount As Long

   tbdT = Vsp.TableBorder
   blnBold = Vsp.FontBold
   iSize = Vsp.FontSize
   
   Call m_Headers0.ClearField
   Call m_Headers1.ClearField
   
   Vsp.FontSize = HEADER_SIZE
   Vsp.FontBold = True
   Vsp.TableBorder = tbAll
   
   Set Cf = New CReportField
      
      Call Cf.SetFieldValue(1, "^", "ercial1", "<")
      Call m_Headers0.AddField(Cf)
      Call Cf.SetFieldValue(4, "^", "", "^")
      Call m_Headers0.AddField(Cf)
      Call Cf.SetFieldValue(6, "^", "คอมมิชชั่น (เก็บเงิน)", "<")    '27
      Call m_Headers0.AddField(Cf)
      Call Cf.SetFieldValue(6, "^", "คอมมิชชั่น (ขาย)", "^")
      Call m_Headers0.AddField(Cf)
      Call Cf.SetFieldValue(2, "^", "(ค่าเช่ารถ)", "<")    '27
      Call m_Headers0.AddField(Cf)
      Call Cf.SetFieldValue(2, "^", "", "^")
      Call m_Headers0.AddField(Cf)
      Call Cf.SetFieldValue(2, "^", "", "<")    '27
      Call m_Headers0.AddField(Cf)
      Call Cf.SetFieldValue(2, "^", "", "<")    '27
      Call m_Headers0.AddField(Cf)
      
      Call Cf.SetFieldValue(1, "^", "Comm", "<")
      Call m_Headers1.AddField(Cf)
      Call Cf.SetFieldValue(4, "^", "พนักงายขาย", "<")
      Call m_Headers1.AddField(Cf)
      Call Cf.SetFieldValue(3, "^", "Sale", ">")
      Call m_Headers1.AddField(Cf)
      Call Cf.SetFieldValue(3, "^", "Supervise/Manager", ">")
      Call m_Headers1.AddField(Cf)
      Call Cf.SetFieldValue(3, "^", "Sale", ">")
      Call m_Headers1.AddField(Cf)
      Call Cf.SetFieldValue(3, "^", "Supervise/Manager", ">")
      Call m_Headers1.AddField(Cf)
      Call Cf.SetFieldValue(2, "^", "Acheivement", ">")
      Call m_Headers1.AddField(Cf)
      Call Cf.SetFieldValue(2, "^", "INCENTIVE", ">")
      Call m_Headers1.AddField(Cf)
      Call Cf.SetFieldValue(2, "^", "รวมทั้งสิ้น", ">")
      Call m_Headers1.AddField(Cf)
      Call Cf.SetFieldValue(2, "^", "หมายเหตุ", "<")
      Call m_Headers1.AddField(Cf)

   Set Cf = Nothing

   Call m_Headers1.GetString(1, TempStr1, TempStr1_1)
   strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
   strPrint = TempStr1_1
   Call Vsp.AddTable(strFormat, "", strPrint)
   
     Call m_Headers0.GetString(1, TempStr1, TempStr1_1)
   strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
   strPrint = TempStr1_1
   Call Vsp.AddTable(strFormat, "", strPrint)
    Set Cf = Nothing
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
   Vsp.ORIENTATION = orLandscape             ' orPortrait
   Vsp.MarginBottom = 300
   Vsp.MarginFooter = 300
   Vsp.MarginHeader = 300
   Vsp.MarginLeft = 300
   Vsp.MarginRight = 500
   Vsp.MarginTop = 300
   Vsp.FontName = "AngsanaUPC"
   Vsp.FontSize = DETAIL_SIZE
   
MODULE_DESC = glbCompanyName & vbCrLf & mcolParam("REPORT_NAME") & _
                             vbCrLf & "ประจำวันที่   " & EmptyToString(DateToStringExtEx2(mcolParam("FROM_CMPL_DATE")), "N/A") & "    ถึง    " & EmptyToString(DateToStringExtEx2(mcolParam("TO_CMPL_DATE")), "N/A")

Call Load3ComcollTotal(mcolParam("FROM_CMPL_DATE"), mcolParam("TO_CMPL_DATE"), mcolParam("FROM_CMPL_DATE"), mcolParam("TO_CMPL_DATE"), collTotal)
'Call LoadcollTotalV(mcolParam("FROM_DOC_DATE"), mcolParam("TO_DOC_DATE"), collTotalVAC, coll_MinusVAC, False)    ' เฉพาะวัคซีน
Call LoadColGoodsGroup(m_GoodsGroup)
Call LoadSaleHead(m_saleChartHead, mcolParam("FROM_CMPL_DATE"), mcolParam("TO_CMPL_DATE"))
Call LoadSaleChartChild(m_saleChartChild, mcolParam("FROM_CMPL_DATE"), mcolParam("TO_CMPL_DATE"))
Call LoadSale(Nothing, m_SaleName)
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
Dim PrevKey3 As Long   'รหัสเขต
Dim PrevKey4 As String    'เขตการขาย
Dim St As CCommissMasterArea
Dim Stcrd As CStcrd
Dim ArS As COESLM

Dim haveValueSale As Boolean
Dim haveData As Boolean
Dim haveData_ByArea As Boolean
Dim DueCount As Double
Dim temp_CollTot As CChartTotal
Dim temp_sale As COESLM

Dim TOTAL1_AMOUNT As Double
Dim TOTAL2_AMOUNT As Double
Dim GP As Double
Dim GROUP As Double
Dim MINUS As String
Dim REAL_COM1 As Double
Dim REAL_COM2 As Double
Dim REAL_INCENTIVE As Double

Dim text_AREA_NAME  As String

   
   RName = "genDoc"
'-----------------------------------------------------------------------------------------------------
'                                             Query Here
'-----------------------------------------------------------------------------------------------------
   
   Set Rs = New ADODB.Recordset
   Set TempRs = New ADODB.Recordset
   Set temp_CollTot = New CChartTotal
'   Set temp_CollMinus = New CChartTotal
   Set temp_sale = New COESLM
'
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
   i = 0  'นับจ.บรรทัดข้อมูลย่อย
   Set BodyCf = New CReportField

   Set ArS = New COESLM
   Call glbDaily.QuerySale(ArS, TempRs, iCount, IsOK, glbErrorLog)
   
   While Not TempRs.EOF          ' sale
   
         Call m_Details1.ClearField
         Call ArS.PopulateFromRS(1, TempRs)

 For Each temp_GoodsGroup In m_GoodsGroup
  
  Set temp_Head = GetSaleChart(m_saleChartHead, ArS.SLMCOD, False)       ' เซลล์คนนี้มีอยู่ใน Chart หรือเปล่า
  If Not (temp_Head Is Nothing) Then
                i = i + 1
                j = 0
               
                j = j + 1
               Set HeadCf = m_Headers1.Fields(j)
               TempStr = "*" & i
               Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
               Call m_Details1.AddField(BodyCf)
   
                j = j + 1
               Set HeadCf = m_Headers1.Fields(j)
               TempStr = ArS.SLMNAM & " + " & temp_Head.AREA_NAME
               Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
               Call m_Details1.AddField(BodyCf)
               
               Set temp_CollTot = GetTotMinChart(collTotal, Trim(ArS.SLMCOD & "-" & temp_Head.MASTER_AREA_ID), False)
               j = j + 1
               Set HeadCf = m_Headers1.Fields(j)
               TempStr = FormatNumber(temp_CollTot.REAL_COM2)       ' sale เก็บเงิน
               Total2(j) = Total2(j) + temp_CollTot.REAL_COM2
               Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
               Call m_Details1.AddField(BodyCf)
               
               j = j + 1                                                                            ' หัวหน้าเก็บเงิน
               Set HeadCf = m_Headers1.Fields(j)
               TempStr = ""
               Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
               Call m_Details1.AddField(BodyCf)
               
                 j = j + 1
               Set HeadCf = m_Headers1.Fields(j)                                    ' sale ขาย
               TempStr = FormatNumber(temp_CollTot.REAL_COM1)
                Total2(j) = Total2(j) + temp_CollTot.REAL_COM1
               Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
               Call m_Details1.AddField(BodyCf)
               
               j = j + 1                                                                                 ' หัวหน้าขาย
               Set HeadCf = m_Headers1.Fields(j)
               TempStr = ""
               Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
               Call m_Details1.AddField(BodyCf)
               
               j = j + 1                                                                                  'ค่าเช่ารถ
               Set HeadCf = m_Headers1.Fields(j)
               TempStr = ""
               Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
               Call m_Details1.AddField(BodyCf)
               
                j = j + 1                                                                            ' Incentive
               Set HeadCf = m_Headers1.Fields(j)
               TempStr = FormatNumber(temp_CollTot.REAL_INCENTIVE)
                Total2(j) = Total2(j) + temp_CollTot.REAL_INCENTIVE
               Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
               Call m_Details1.AddField(BodyCf)
               
                j = j + 1                                                ' รวมทั้งสิ้น
               Set HeadCf = m_Headers1.Fields(j)
               TempStr = FormatNumber(temp_CollTot.REAL_COM1 + temp_CollTot.REAL_COM2 + temp_CollTot.REAL_INCENTIVE)
                Total2(j) = Total2(j) + (temp_CollTot.REAL_COM1 + temp_CollTot.REAL_COM2 + temp_CollTot.REAL_INCENTIVE)
               Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
               Call m_Details1.AddField(BodyCf)
               
               j = j + 1                                                ' หมายเหตุ
               Set HeadCf = m_Headers1.Fields(j)
               TempStr = ""
               Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
               Call m_Details1.AddField(BodyCf)
               
              If mcolParam("SUMMARY_MODE") = "N" Then
                     Call m_Details1.GetString(1, TempStr1, TempStr2)
                    strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
                    Call Vsp.AddTable(strFormat, "", TempStr2)
               End If
' ----------------------------------------------------------
                  K = 1
            For Each temp_Child In m_saleChartChild                   ' วนหาลูก
                     Set temp_Child2 = GetSaleChart(m_saleChartChild, Trim(temp_Head.SALE_ID & "-" & temp_Child.SALE_ID & "-" & temp_Child.MASTER_AREA_ID), False)
                     If Not (temp_Child2 Is Nothing) Then
                               ' i = i + 1
                               j = 0
                               K = K + 1
                              Call m_Details1.ClearField
'
                j = j + 1
               Set HeadCf = m_Headers1.Fields(j)
               TempStr = i
               Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
               Call m_Details1.AddField(BodyCf)
   
               Set temp_sale = GetSlm(m_SaleName, temp_Child2.SALE_ID, False)
               If temp_sale Is Nothing Then
                     TempStr = ""
               Else
                     TempStr = temp_sale.SLMNAM
               End If
                j = j + 1
               Set HeadCf = m_Headers1.Fields(j)
               TempStr = TempStr & " + " & temp_Child.AREA_NAME
               Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
               Call m_Details1.AddField(BodyCf)
               
               Set temp_CollTot = GetTotMinChart(collTotal, Trim(temp_Child2.SALE_ID & "-" & temp_Child2.MASTER_AREA_ID), False)
               If temp_CollTot Is Nothing Then
                  REAL_COM1 = 0
                  REAL_COM2 = 0
                  REAL_INCENTIVE = 0
               Else
                  REAL_COM1 = temp_CollTot.REAL_COM1
                  REAL_COM2 = temp_CollTot.REAL_COM2
                  REAL_INCENTIVE = temp_CollTot.REAL_INCENTIVE
               End If
               
               j = j + 1
               Set HeadCf = m_Headers1.Fields(j)
               TempStr = FormatNumber(REAL_COM2)       ' sale เก็บเงิน
               Total1(j) = Total1(j) + REAL_COM2
               Total2(j) = Total2(j) + REAL_COM2
               Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
               Call m_Details1.AddField(BodyCf)
               
               j = j + 1                                                                            ' หัวหน้าเก็บเงิน
               Set HeadCf = m_Headers1.Fields(j)
               TempStr = ""
               Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
               Call m_Details1.AddField(BodyCf)
               
                 j = j + 1
               Set HeadCf = m_Headers1.Fields(j)                                    ' sale ขาย
               TempStr = FormatNumber(REAL_COM1)
               Total1(j) = Total1(j) + REAL_COM1
               Total2(j) = Total2(j) + REAL_COM1
               Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
               Call m_Details1.AddField(BodyCf)
               
               j = j + 1                                                                                 ' หัวหน้าขาย
               Set HeadCf = m_Headers1.Fields(j)
               TempStr = ""
               Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
               Call m_Details1.AddField(BodyCf)
               
               j = j + 1                                                                                  'ค่าเช่ารถ
               Set HeadCf = m_Headers1.Fields(j)
               TempStr = ""
               Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
               Call m_Details1.AddField(BodyCf)
               
                j = j + 1                                                                            ' Incentive
'                Total1(j) = Total1(j) + temp_Head.BUDGET
               Set HeadCf = m_Headers1.Fields(j)
               TempStr = FormatNumber(REAL_INCENTIVE)
               Total1(j) = Total1(j) + REAL_INCENTIVE
               Total2(j) = Total2(j) + REAL_INCENTIVE
               Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
               Call m_Details1.AddField(BodyCf)

                j = j + 1                                                ' รวมทั้งสิ้น
               Set HeadCf = m_Headers1.Fields(j)
               TempStr = FormatNumber(REAL_COM1 + REAL_COM2 + REAL_INCENTIVE)
               Total2(j) = Total2(j) + (REAL_COM1 + REAL_COM2 + REAL_INCENTIVE)
               Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
               Call m_Details1.AddField(BodyCf)
               
   
               j = j + 1                                                ' หมายเหตุ
               Set HeadCf = m_Headers1.Fields(j)
               TempStr = ""
               Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
               Call m_Details1.AddField(BodyCf)
               
                    If mcolParam("SUMMARY_MODE") = "N" Then
                           Call m_Details1.GetString(1, TempStr1, TempStr2)
                            strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
                            Call Vsp.AddTable(strFormat, "", TempStr2)
                     End If

                     End If
            Next temp_Child
            
'----------------------------------

        Call m_Details1.ClearField
        j = 0
        ' รวม manager
                j = j + 1
               Set HeadCf = m_Headers1.Fields(j)
               TempStr = "รวม"                         ' & Str(i)
                  Vsp.FontBold = True
               Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
               Call m_Details1.AddField(BodyCf)
   
                j = j + 1
               Set HeadCf = m_Headers1.Fields(j)
               TempStr = ArS.SLMNAM
               Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
               Call m_Details1.AddField(BodyCf)

               j = j + 1
               Set HeadCf = m_Headers1.Fields(j)
               TempStr = FormatNumber(Total1(j))                                                          ' sale เก็บเงิน
               Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
               Call m_Details1.AddField(BodyCf)
               
               j = j + 1                                                                            ' หัวหน้าเก็บเงิน
               Set HeadCf = m_Headers1.Fields(j)
               TempStr = FormatNumber(Total1(j - 1) * 0.15)
               Total2(j) = Total2(j) + (Total1(j - 1) * 0.15)
               Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
               Call m_Details1.AddField(BodyCf)
               
                 j = j + 1
               Set HeadCf = m_Headers1.Fields(j)                                    ' sale ขาย
               TempStr = FormatNumber(Total1(j))
               Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
               Call m_Details1.AddField(BodyCf)
               
               j = j + 1                                                                                 ' หัวหน้าขาย
               Set HeadCf = m_Headers1.Fields(j)
               TempStr = FormatNumber(Total1(j - 1) * 0.15)
               Total2(j) = Total2(j) + (Total1(j - 1) * 0.15)
               Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
               Call m_Details1.AddField(BodyCf)
               
               j = j + 1                                                                                  'ค่าเช่ารถ
               Set HeadCf = m_Headers1.Fields(j)
               TempStr = ""
               Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
               Call m_Details1.AddField(BodyCf)
               
                j = j + 1                                                                            ' Incentive
               Set HeadCf = m_Headers1.Fields(j)
               TempStr = ""
               Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
               Call m_Details1.AddField(BodyCf)

                j = j + 1                                                                            ' รวมทั้งสิ้น
               Set HeadCf = m_Headers1.Fields(j)
               TempStr = FormatNumber((Total1(3) + Total1(5)) * 0.15)
               Total2(j) = Total2(j) + ((Total1(3) + Total1(5)) * 0.15)
               Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
               Call m_Details1.AddField(BodyCf)
               
               j = j + 1                                                ' หมายเหตุ
               Set HeadCf = m_Headers1.Fields(j)
               TempStr = ""
               Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
               Call m_Details1.AddField(BodyCf)
               
                     Call m_Details1.GetString(1, TempStr1, TempStr2)
                    strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
                    Call Vsp.AddTable(strFormat, "", TempStr2)
                       Vsp.FontBold = False

                  For j = 1 To UBound(Total1)
                     Total1(j) = 0
                  Next j

       End If   ' มีค่า Temp_HEAD
       
       Next temp_GoodsGroup
     
        TempRs.MoveNext                                                            ' วนเซลล์
Wend
            
            ' รวมสุทธิ
         Call m_Details1.ClearField

           j = 0

                j = j + 1
               Set HeadCf = m_Headers1.Fields(j)
               TempStr = ""                                                                            ' & Str(i)
                  Vsp.FontBold = True
               Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
               Call m_Details1.AddField(BodyCf)
   
                j = j + 1
               Set HeadCf = m_Headers1.Fields(j)
               TempStr = "สุทธิ"
               Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
               Call m_Details1.AddField(BodyCf)

               j = j + 1
               Set HeadCf = m_Headers1.Fields(j)
               TempStr = FormatNumber(Total2(j))                                                                   ' sale เก็บเงิน
               Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
               Call m_Details1.AddField(BodyCf)
               
               j = j + 1                                                                            ' หัวหน้าเก็บเงิน
               Set HeadCf = m_Headers1.Fields(j)
               TempStr = FormatNumber(Total2(j))
               Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
               Call m_Details1.AddField(BodyCf)
               
                 j = j + 1
               Set HeadCf = m_Headers1.Fields(j)                                    ' sale ขาย
               TempStr = FormatNumber(Total2(j))
               Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
               Call m_Details1.AddField(BodyCf)
               
               j = j + 1                                                                                 ' หัวหน้าขาย
               Set HeadCf = m_Headers1.Fields(j)
               TempStr = FormatNumber(Total2(j))
               Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
               Call m_Details1.AddField(BodyCf)
               
               j = j + 1                                                                                  'ค่าเช่ารถ
               Set HeadCf = m_Headers1.Fields(j)
               TempStr = ""
               Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
               Call m_Details1.AddField(BodyCf)
               
                j = j + 1                                                                            ' Incentive
               Set HeadCf = m_Headers1.Fields(j)
               TempStr = FormatNumber(Total2(j))
               Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
               Call m_Details1.AddField(BodyCf)

                j = j + 1                                                                            ' รวมทั้งสิ้น
               Set HeadCf = m_Headers1.Fields(j)
               TempStr = FormatNumber(Total2(j))
               Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
               Call m_Details1.AddField(BodyCf)
               
   
               j = j + 1                                                ' หมายเหตุ
               Set HeadCf = m_Headers1.Fields(j)
               TempStr = ""
               Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
               Call m_Details1.AddField(BodyCf)
               
                     Call m_Details1.GetString(1, TempStr1, TempStr2)
                    strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
                    Call Vsp.AddTable(strFormat, "", TempStr2)
                       Vsp.FontBold = False

               Vsp.EndDoc
               Call CloseExportFile(Vsp)
               
               If Rs.State = adStateOpen Then
                     Rs.Close
               End If
               Set Rs = Nothing
               
               If TempRs.State = adStateOpen Then
                  TempRs.Close
               End If
               Set TempRs = Nothing
               
               genDoc = True
               Exit Function
   
ErrHandler:
   Err.DESCRIPTION = "กรุณาเลือก Database ที่ 2 เพื่อทำการติดต่อข้อมูลก่อน โดยไปที่ : โปรแกรม/ติดต่อ database ที่ 2"
   mstrErrMsg = "Error(" & RName & ")" & Err.Number & " : " & Err.DESCRIPTION
   
   Set Rs = Nothing
      Set TempRs = Nothing
End Function

Private Sub GenerateHeader(Txt1 As String, Txt2 As String, Details1 As CFieldList)
Dim HeadCf As CReportField
Dim BodyCf As CReportField
Dim j As Long
Dim TempStr As String

   Call Details1.ClearField
   Set BodyCf = New CReportField

   Set BodyCf = Nothing
End Sub

Private Sub GenerateFooter(Txt1 As String, Txt2 As String, Txt3 As String, Tot() As Double, Details As CFieldList, HaveValue As Boolean)
Dim HeadCf As CReportField
Dim BodyCf As CReportField
Dim j As Long
Dim TempStr As String

   Call Details.ClearField
   Set BodyCf = New CReportField
   
   HaveValue = False
   
   j = 0
   For Each HeadCf In m_Headers1.Fields
      j = j + 1
      If j = 1 Then
         TempStr = Txt1
      ElseIf j = 2 Then
      'debug.print Txt2
         TempStr = Txt2
      ElseIf j = 3 Then
         'debug.print Txt3
         TempStr = Txt3
     ElseIf j = 4 Then
        'debug.print FormatNumber(Tot(j))
         TempStr = FormatNumber(Tot(j))
      ElseIf j = 5 Then
        'debug.print FormatNumber(Tot(j))
         TempStr = FormatNumber(Tot(j))
      ElseIf j = 9 Then
         TempStr = ""
      ElseIf j = 10 Then
         TempStr = ""
      Else
         TempStr = FormatNumber(Tot(j))
      End If
      
      If HaveValue = False And Tot(j) <> 0 Then
         HaveValue = True
      End If
   Next HeadCf
      
   Set BodyCf = Nothing
End Sub

Private Sub VSP_EndDoc()
'This event occur when VSPrinter.EndDoc is used
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


