VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CReportCom01"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'User Group Report
Implements CReportInterface

Private Const MODULE_NAME = "CReportCOM01"
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

Private m_IVincomplete0 As Collection
Private m_IVincomplete12 As Collection
Private m_ReDocdat As Collection
Private temp_Area As CCommissMasterArea

Private collTotal As Collection
Private collTotalVAC As Collection
Private coll_Minus As Collection
Private coll_MinusVAC As Collection

Private m_saleChartHead As Collection
Private m_GPwithGroup As Collection
Private m_saleChartChild As Collection
Private temp_Head As CCommissionChart
 Private temp_Child As CCommissionChart
 Private temp_Child2 As CCommissionChart
 Private m_SaleName As Collection
 Private m_ComRecord As Collection
 
 Dim n_array As Long
 Private m_GoodsGroup As Collection
 Private m_ConditionCommiss1a(10) As Collection
 Private temp_GoodsGroup As CGoodsGroup
 Private temp_ChildGoodsGroup As CGoodsGroup
  
 Private m_ComPro1 As Collection
Private m_ComMasPro As Collection
Private temp_ComMasPro As CCommissMasterPromote
 
Dim L As Long

'Private havevalue As Boolean

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
   Set m_StcrdBySale = New Collection
   Set m_AreaCod = New Collection
   Set m_AreaFromIV = New Collection
   Set m_ConditionCommiss1 = New Collection
   Set m_ConditionCommiss2 = New Collection
   Set m_ConditionCommiss3 = New Collection
   Set m_ConditionCommiss4 = New Collection
   For L = 1 To UBound(m_cusFromArea)
         Set m_cusFromArea(L) = New Collection
   Next L
   Set Stb = New CCommissMasterArea
   Set m_IVincomplete0 = New Collection
   Set m_IVincomplete12 = New Collection
   Set m_ReDocdat = New Collection
   Set temp_Area = New CCommissMasterArea
   Set temp_Head = New CCommissionChart
   Set temp_Child = New CCommissionChart
   Set temp_Child2 = New CCommissionChart
   Set collTotal = New Collection
   Set collTotalVAC = New Collection
   Set coll_Minus = New Collection
   Set coll_MinusVAC = New Collection
   Set m_saleChartHead = New Collection
   Set m_saleChartChild = New Collection
   Set m_SaleName = New Collection
   Set m_ConditionCommiss1 = New Collection
   Set m_GPwithGroup = New Collection
   
      Set m_ComPro1 = New Collection
      Set m_ComMasPro = New Collection
      Set temp_ComMasPro = New CCommissMasterPromote
      
    Set m_GoodsGroup = New Collection
    For n_array = 0 To UBound(m_ConditionCommiss1a)
      Set m_ConditionCommiss1a(n_array) = New Collection
   Next n_array
   Set temp_GoodsGroup = New CGoodsGroup
   Set temp_ChildGoodsGroup = New CGoodsGroup
   
   Set m_ComRecord = New Collection
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
  Set coll_MinusVAC = Nothing
  Set m_IVincomplete0 = Nothing
  Set m_IVincomplete12 = Nothing
  Set m_ReDocdat = Nothing
  Set temp_Area = Nothing
  Set collTotal = Nothing
  Set collTotalVAC = Nothing
  Set m_saleChartHead = Nothing
  Set m_saleChartChild = Nothing
   Set temp_Head = Nothing
   Set temp_Child = Nothing
   Set temp_Child2 = Nothing
   Set m_SaleName = Nothing
   Set m_GPwithGroup = Nothing
   
   Set m_ComPro1 = Nothing
   Set m_ComMasPro = Nothing
   Set temp_ComMasPro = Nothing
   
   Set m_GoodsGroup = Nothing
    For n_array = 0 To UBound(m_ConditionCommiss1a)
      Set m_ConditionCommiss1a(n_array) = Nothing
   Next n_array
   Set temp_GoodsGroup = Nothing
   Set temp_ChildGoodsGroup = Nothing
   Set m_ComRecord = Nothing
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
   
   Call m_Headers1.ClearField
   
   Vsp.FontSize = HEADER_SIZE
   Vsp.FontBold = True
   Vsp.TableBorder = tbAll
   
   Set Cf = New CReportField
      
    Call Cf.SetFieldValue(3, "^", "", "^")
    Call m_Headers0.AddField(Cf)
    Call Cf.SetFieldValue(26, "^", "", "<")    '27
    Call m_Headers0.AddField(Cf)
      

   Call Cf.SetFieldValue(1, "^", "Commercial 1", "^")
   Call m_Headers1.AddField(Cf)
   
      Call Cf.SetFieldValue(4, "^", "พนักงานขาย", "<")
   Call m_Headers1.AddField(Cf)
   
      Call Cf.SetFieldValue(3, "^", "เขต", "<")
   Call m_Headers1.AddField(Cf)
   
   Call Cf.SetFieldValue(2, "^", "สินค้า", "<")
   Call m_Headers1.AddField(Cf)

      Call Cf.SetFieldValue(2, "^", "ยอดขาย", ">")
      Call m_Headers1.AddField(Cf)
      
      Call Cf.SetFieldValue(2, "^", "C/N", ">")
      Call m_Headers1.AddField(Cf)
      
    Call Cf.SetFieldValue(2, "^", "รวม", ">")
   Call m_Headers1.AddField(Cf)
   
      Call Cf.SetFieldValue(2, "^", "ยอดประเมิน", ">")
   Call m_Headers1.AddField(Cf)

      Call Cf.SetFieldValue(3, "^", "งบประมาณ", ">")
   Call m_Headers1.AddField(Cf)
   
         Call Cf.SetFieldValue(2, "^", "%", "^")
   Call m_Headers1.AddField(Cf)
   
            Call Cf.SetFieldValue(2, "^", "ค่าคอมขาย", ">")
   Call m_Headers1.AddField(Cf)
   
            Call Cf.SetFieldValue(2, "^", "GP.", "^")
   Call m_Headers1.AddField(Cf)
   
               Call Cf.SetFieldValue(2, "^", "สัดส่วน Index.", "^")
   Call m_Headers1.AddField(Cf)
   
            Call Cf.SetFieldValue(3, "^", "จ่ายจริง", ">")
   Call m_Headers1.AddField(Cf)
   
   Call m_Headers1.GetString(1, TempStr1, TempStr1_1)
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
                            vbCrLf & "วันที่ส่งสินค้า      " & EmptyToString(DateToStringExtEx2(mcolParam("FROM_DOC_DATE")), "N/A") & "      ถึง        " & EmptyToString(DateToStringExtEx2(mcolParam("TO_DOC_DATE")), "N/A")
  
Call LoadComNamePro(m_ComMasPro, mcolParam("FROM_DOC_DATE"), mcolParam("TO_DOC_DATE"))       'รายชื่อ เซลล์และลูกค้าคนไหนทีเข้าข่าย
Call LoadComPro("01", m_ComPro1, mcolParam("FROM_DOC_DATE"), mcolParam("TO_DOC_DATE"))      'เก็บเงินธรรมดา

Call LoadGPwithGroup(m_GPwithGroup, mcolParam("FROM_DOC_DATE"), mcolParam("TO_DOC_DATE"))
   
   n_array = 0
   Call LoadColGoodsGroup(m_GoodsGroup)                ' มีประเภทไรบ้าง    ' ดึง ชื่อสินค้าที่อยู่ใน master นั่นน และกรุ๊ปนั่นนนน
   For Each temp_GoodsGroup In m_GoodsGroup
      Call LoadCommission("01", m_ConditionCommiss1a(temp_GoodsGroup.GOODS_GROUP_ID), mcolParam("FROM_DOC_DATE"), mcolParam("TO_DOC_DATE"), temp_GoodsGroup.GOODS_GROUP_ID)
      n_array = n_array + 1
  Next temp_GoodsGroup
  
Call LoadSaleHead(m_saleChartHead, mcolParam("FROM_DOC_DATE"), mcolParam("TO_DOC_DATE"))
Call LoadSaleChartChild(m_saleChartChild, mcolParam("FROM_DOC_DATE"), mcolParam("TO_DOC_DATE"))
Call LoadSale(Nothing, m_SaleName)
Call LoadcollTotal(mcolParam("FROM_DOC_DATE"), mcolParam("TO_DOC_DATE"), collTotal, coll_Minus, True)
Call LoadcollTotal(mcolParam("FROM_DOC_DATE"), mcolParam("TO_DOC_DATE"), collTotalVAC, coll_MinusVAC, False)    ' เฉพาะวัคซีน
'Call LoadcollTotalV(mcolParam("FROM_DOC_DATE"), mcolParam("TO_DOC_DATE"), collTotal, coll_Minus, True)    ' ทั้งหมด
 
' Save_mode
If mcolParam("SAVE_MODE") = "Y" Then
    Call LoadComRecord(m_ComRecord, "01", mcolParam("TO_DOC_DATE"))
End If
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

Dim TempCConditionCommiss As CConditionCommission
Dim m_runConditionCommiss As CConditionCommission
Dim PercentNum1 As Double
Dim PercentSum As Double
Dim Percent100 As Double
Dim tempCusArea As CCommissionCustomerArea
Dim toCMPLdate As Date
Dim tempMinusStkcod As CComMinusStk
'Dim tempIVincomplete0 As CComIVincomplete
'Dim tempIVincomplete12 As CComIVincomplete
Dim tempREdoc As CARTrn
Dim NETVAL As Double
Dim PayCompleted  As Boolean
Dim CMPLDAT As Date
Dim RECOD As String
Dim D As CChartTotal
Dim temp_CollTot As CChartTotal
Dim temp_CollMinus As CChartTotal
Dim temp_GPwithGroup As CMasterFromToDetail
Dim temp_sale As COESLM

Dim TOTAL1_AMOUNT As Double
Dim TOTAL2_AMOUNT As Double
Dim GP As Double
Dim GROUP As Double
Dim MINUS As String

Dim text_AREA_NAME  As String
Dim TempComPro As CComMasSubPromote
Dim m_runComPro As CComMasSubPromote
Dim m_runPro As CCommissPromote

Dim CreditKey As Long
Dim asAmount As Double
Dim MoreAmount As Double
Dim AMOUNT As Double
Dim NUM_ONE_Y As Double
   
   RName = "genDoc"
'-----------------------------------------------------------------------------------------------------
'                                             Query Here
'-----------------------------------------------------------------------------------------------------
   
   Set Rs = New ADODB.Recordset
   Set TempRs = New ADODB.Recordset
   Set temp_CollTot = New CChartTotal
   Set temp_CollMinus = New CChartTotal
   Set temp_sale = New COESLM
   
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
    
      Set temp_Head = GetSaleChart(m_saleChartHead, Trim(ArS.SLMCOD & "-" & temp_GoodsGroup.GOODS_GROUP_ID), False) '
      If Not (temp_Head Is Nothing) Then
      If Not (m_ConditionCommiss1a(temp_GoodsGroup.GOODS_GROUP_ID).Count = 0) Then
      
                j = 0
                i = i + 1
                
                j = j + 1
               Set HeadCf = m_Headers1.Fields(j)
               TempStr = "*" & i
               Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
               Call m_Details1.AddField(BodyCf)
   
                j = j + 1
               Set HeadCf = m_Headers1.Fields(j)
               TempStr = ArS.SLMNAM
               Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
               Call m_Details1.AddField(BodyCf)
               
               j = j + 1
               Set HeadCf = m_Headers1.Fields(j)
               TempStr = temp_Head.AREA_NAME      ' ชื่อเขต
               Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
               Call m_Details1.AddField(BodyCf)
               
                j = j + 1
               Set HeadCf = m_Headers1.Fields(j)
               TempStr = temp_Head.GOODS_GROUP_NAME       ' สินค้า
               Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
               Call m_Details1.AddField(BodyCf)
                     
                                             
                     '   If temp_Head.SALE_ID = 15 Then
                      '     'debug.print
                    '    End If
            
                     Set temp_CollTot = GetTotMinChart(collTotalVAC, Trim(ArS.SLMCOD & "-" & temp_Head.MASTER_AREA_ID), False)        ' temp_GoodsGroup.GOODS_GROUP_ID
                     If Not (temp_CollTot Is Nothing) Then
                         TOTAL1_AMOUNT = temp_CollTot.TOTAL1_SUM
                         TOTAL2_AMOUNT = temp_CollTot.TOTAL2_SUM
                     Else
                         TOTAL1_AMOUNT = 0
                     End If
                     
                     Set temp_CollMinus = Nothing
                     Set temp_CollMinus = GetTotMinChart(coll_MinusVAC, Trim(ArS.SLMCOD & "-" & temp_Head.MASTER_AREA_ID), False)        'temp_GoodsGroup.GOODS_GROUP_ID
                     If Not (temp_CollMinus Is Nothing) Then
                        MINUS = temp_CollMinus.MINUS
                     Else
                        MINUS = 0
                     End If
            
            If temp_Head.GOODS_GROUP_ID = 2 Then
                     Set temp_CollTot = GetTotMinChart(collTotal, Trim(ArS.SLMCOD & "-" & temp_Head.MASTER_AREA_ID), False)
                     If Not (temp_CollTot Is Nothing) Then
                         TOTAL1_AMOUNT = temp_CollTot.TOTAL1_SUM - TOTAL1_AMOUNT
                         TOTAL2_AMOUNT = temp_CollTot.TOTAL2_SUM - TOTAL2_AMOUNT
                     Else
                         TOTAL1_AMOUNT = 0
                     End If
                     
                     Set temp_CollMinus = Nothing
                     Set temp_CollMinus = GetTotMinChart(coll_Minus, Trim(ArS.SLMCOD & "-" & temp_Head.MASTER_AREA_ID), False)
                     If Not (temp_CollMinus Is Nothing) Then
                        MINUS = temp_CollMinus.MINUS - MINUS
                     Else
                        MINUS = 0
                     End If
            End If
               
               j = j + 1                                                                            'ยอดขาย
               Set HeadCf = m_Headers1.Fields(j)
               TempStr = FormatNumber(((-1) * Val(MINUS)) + TOTAL1_AMOUNT)
               Total1(j) = Total1(j) + (((-1) * Val(MINUS)) + TOTAL1_AMOUNT)
               Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
               Call m_Details1.AddField(BodyCf)
              
               j = j + 1                                                                                'minus
                 Total1(j) = Total1(j) + ((-1) * Val(MINUS))
               Set HeadCf = m_Headers1.Fields(j)
               TempStr = FormatNumber(((-1) * Val(MINUS)))
               Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
               Call m_Details1.AddField(BodyCf)
               
               j = j + 1                                                                                 'รวม
               Set HeadCf = m_Headers1.Fields(j)
               Total1(j) = Total1(j) + TOTAL1_AMOUNT
               TempStr = FormatNumber(TOTAL1_AMOUNT)
               Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
               Call m_Details1.AddField(BodyCf)
               
               j = j + 1                                                                                  'ยอดประเมิน
               Total1(j) = Total1(j) + TOTAL2_AMOUNT
               Set HeadCf = m_Headers1.Fields(j)
               TempStr = FormatNumber(TOTAL2_AMOUNT)
               Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
               Call m_Details1.AddField(BodyCf)
               
                j = j + 1                                                                            'งบประมาณ   temp_Head.Budget
                Total1(j) = Total1(j) + temp_Head.BUDGET
               Set HeadCf = m_Headers1.Fields(j)
               TempStr = FormatNumber(temp_Head.BUDGET)
               Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
               Call m_Details1.AddField(BodyCf)
               
               If temp_Head.BUDGET = "0" Or temp_Head.BUDGET = "" Then
                             temp_Head.BUDGET = "1"
                End If
               
                j = j + 1                                                ' %
               Set HeadCf = m_Headers1.Fields(j)
               DueCount = Round((TOTAL2_AMOUNT / (Val(temp_Head.BUDGET) / 100)), 2)
               Total1(j) = Total1(j) + DueCount
               TempStr = DueCount
               Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
               Call m_Details1.AddField(BodyCf)
               
               'คำนวนคอมขาย
''''''''            Set temp_ComMasPro = GetComMasPro(m_ComMasPro, Trim(Stcrd.SLMCOD & "-" & Stcrd.CUSCOD), False)
''''''''            If Not (temp_ComMasPro Is Nothing) Then           ' หาโดยใช้  .SLMCOD .PEOPLE
''''''''                     Set m_runComPro = m_ComPro1.ITEM(1)
''''''''                     CreditKey = Val(m_runComPro.CREDIT_NAME)
''''''''                  '   'debug.print CreditKey & " ~ " & DueCount
''''''''                     For Each m_runComPro In m_ComPro1                   ' เก็บเงินธรรมดา
''''''''                           If (Val(m_runComPro.CREDIT_NAME)) >= DueCount Then
''''''''                                CreditKey = Val(m_runComPro.CREDIT_NAME)    'จะได้ เครดิตเป็น Key
''''''''                             '   'debug.print CreditKey & " ~ " & DueCount
''''''''                          End If
''''''''                     Next m_runComPro
''''''''
''''''''                     ' วน loop ของ YEAR_ID นี้ ,, 02(สินค้าธรรมดา) อยู่ในเครดิตไหน  แล้วมันจะได้ค่า MASTER_COMMISS_SUB_PROMOTE_ID เป็น Key
''''''''                     Set TempComPro = GetComMasSubPro(m_ComPro1, Trim(Stcrd.SLMCOD & "-" & Stcrd.CUSCOD & "-" & CreditKey), False)
''''''''                      PercentSum = (TempComPro.DetailsCom1.ITEM(1).SLM_PERCENT / 100)  'น้อยสุดป่ะ
''''''''
''''''''                     For Each m_runPro In TempComPro.DetailsCom1                    ' เก็บเงินธรรมดา
''''''''                           If (Val(m_runPro.NUM_ONE) + 7) >= DueCount Then
''''''''                                    PercentSum = (m_runPro.SLM_PERCENT / 100)
''''''''                          End If
''''''''                     Next m_runPro
''''''''
''''''''             Else
                     Set TempCConditionCommiss = m_ConditionCommiss1a(temp_GoodsGroup.GOODS_GROUP_ID).ITEM(1)
                     PercentSum = (TempCConditionCommiss.SLM_PERCENT / 100)
                     
                  'หาเปอร์เซ็นของ 100
                  For Each m_runConditionCommiss In m_ConditionCommiss1a(temp_GoodsGroup.GOODS_GROUP_ID)
                      If (m_runConditionCommiss.NUM_ONE) > 100 Then
                           Percent100 = (m_runConditionCommiss.SLM_PERCENT / 100)
                     End If
                  Next m_runConditionCommiss

                     For Each m_runConditionCommiss In m_ConditionCommiss1a(temp_GoodsGroup.GOODS_GROUP_ID)
                          If (m_runConditionCommiss.NUM_ONE) >= DueCount Then
                               PercentSum = (m_runConditionCommiss.SLM_PERCENT / 100)
                         End If
                      Next m_runConditionCommiss
                      
                    For Each m_runConditionCommiss In m_ConditionCommiss1a(temp_GoodsGroup.GOODS_GROUP_ID)   ' ทำเพื่อถ้าอยู่ที่ 100 แต่ไม่ถึง 110 ให้เอายอดประมาณการมาคิด ตัดยอดที่เขาขายได้จริงทิ้ง
                         If (m_runConditionCommiss.Flag) = "Y" Then
                              NUM_ONE_Y = m_runConditionCommiss.NUM_ONE
                        End If
                    Next m_runConditionCommiss
                      
       ''''''''     End If
       
              AMOUNT = (TOTAL2_AMOUNT * PercentSum)
                        
               If DueCount > 100 Then    ' กรณีเกิน 100% คิดเฉพาะยอดส่วนเกิน 100
                  If DueCount < NUM_ONE_Y Then
                    AMOUNT = temp_Head.BUDGET * Percent100
                  Else
                    asAmount = 0
                    MoreAmount = 0
                    asAmount = temp_Head.BUDGET * Percent100
                    
                    MoreAmount = TOTAL2_AMOUNT - temp_Head.BUDGET
                    MoreAmount = MoreAmount * PercentSum
            
                     AMOUNT = MoreAmount + asAmount
                  End If
               End If
                
               j = j + 1                                                ' ค่าคอมขาย
                Total1(j) = Total1(j) + AMOUNT
               Set HeadCf = m_Headers1.Fields(j)
               TempStr = FormatNumber(AMOUNT)               ' เอายอดขายมาคูณ ผลจากเงื่อนไข
               Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
               Call m_Details1.AddField(BodyCf)
               
       Set temp_GPwithGroup = GetObject("CMasterFromToDetail", m_GPwithGroup, Trim(temp_Head.SALE_ID), False)
         If Not (temp_GPwithGroup Is Nothing) Then
             GP = temp_GPwithGroup.GP
             GROUP = temp_GPwithGroup.MASTER_PARAMETER_VALUE
         Else
             GP = 0
             GROUP = 1
         End If
         
         j = j + 1                                                ' Gp
          Set HeadCf = m_Headers1.Fields(j)
          TempStr = FormatNumber(GP)
             Total1(j) = Total1(j) + (GP)
          Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
          Call m_Details1.AddField(BodyCf)
          
           j = j + 1                                                ' INDEX
          Set HeadCf = m_Headers1.Fields(j)
          TempStr = FormatNumber(GP / GROUP)
             Total1(j) = Total1(j) + (GP / GROUP)
          Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
          Call m_Details1.AddField(BodyCf)
          
          j = j + 1                                                ' จ่ายจริง
          Set HeadCf = m_Headers1.Fields(j)
          TempStr = FormatNumber(AMOUNT * (GP / GROUP))            ' เอายอดขายมาคูณ ผลจากเงื่อนไข
             Total1(j) = Total1(j) + (AMOUNT * (GP / GROUP))
          Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
          Call m_Details1.AddField(BodyCf)
               
              If mcolParam("SUMMARY_MODE") = "N" Then
                     Call m_Details1.GetString(1, TempStr1, TempStr2)
                    strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
                    Call Vsp.AddTable(strFormat, "", TempStr2)
                    
                     If mcolParam("SAVE_MODE") = "Y" Then
                          Call SaveToDB(temp_Head.SALE_ID, ArS.SLMNAM, "01", temp_Head.MASTER_AREA_ID, temp_Head.GOODS_GROUP_ID, TempStr, mcolParam("FROM_DOC_DATE"), mcolParam("TO_DOC_DATE"))                 'PreKey3=เลขเขต ,Total2
                     End If
               End If
                  K = 1
                    
            End If   'กรุ๊ปไม่ว่า หัวหน้า
                  
       For Each temp_Child In m_saleChartChild                   ' วนหาลูก
         '   For Each temp_ChildGoodsGroup In m_GoodsGroup

     '    If temp_Head.SALE_ID = "15" And temp_Child.SALE_ID = "04" And temp_Child.MASTER_AREA_ID = 5 Then
     '       'debug.print temp_ChildGoodsGroup.GOODS_GROUP_ID
     '    End If
         
                     Set temp_Child2 = GetSaleChart(m_saleChartChild, Trim(temp_Head.SALE_ID & "-" & temp_Child.SALE_ID & "-" & temp_Child.MASTER_AREA_ID & "-" & temp_Child.GOODS_GROUP_ID), False)
                     If Not (temp_Child2 Is Nothing) Then
                        If temp_Child2.BUDGET <> 0 And Not (temp_Child2.SALE_ID = temp_Head.SALE_ID And temp_Child2.MASTER_AREA_ID = temp_Head.MASTER_AREA_ID And temp_Child2.GOODS_GROUP_ID = temp_Head.GOODS_GROUP_ID) Then
                        
                        If Not (m_ConditionCommiss1a(temp_Child2.GOODS_GROUP_ID).Count = 0) Then ' กรุ๊ปไม่ว่าง
                               ' i = i + 1
                               j = 0
                               K = K + 1
                              Call m_Details1.ClearField
                               
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
                              Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
                              Call m_Details1.AddField(BodyCf)

                               
                              j = j + 1
                         '     text_AREA_NAME = text_AREA_NAME & ", " & temp_Child.AREA_NAME
                              Set HeadCf = m_Headers1.Fields(j)
                              TempStr = temp_Child2.AREA_NAME         ' ชื่อเขต
                              Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
                              Call m_Details1.AddField(BodyCf)
                              
                           j = j + 1
                           Set HeadCf = m_Headers1.Fields(j)
                           TempStr = temp_Child2.GOODS_GROUP_NAME       ' สินค้า
                           Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
                           Call m_Details1.AddField(BodyCf)
                           
                        If temp_Child2.SALE_ID = 15 Then
                           'debug.print
                        End If
                        
                           '   Set temp_CollTot = Nothing
                              Set temp_CollTot = GetTotMinChart(collTotalVAC, Trim(temp_Child2.SALE_ID & "-" & temp_Child2.MASTER_AREA_ID), False)        'temp_Child2.GOODS_GROUP_ID
                              If Not (temp_CollTot Is Nothing) Then
                                  TOTAL1_AMOUNT = temp_CollTot.TOTAL1_SUM
                                  TOTAL2_AMOUNT = temp_CollTot.TOTAL2_SUM
                              Else
                                  TOTAL1_AMOUNT = 0
                                  TOTAL2_AMOUNT = 0
                              End If
                                                            
                              Set temp_CollMinus = GetTotMinChart(coll_MinusVAC, Trim(temp_Child2.SALE_ID & "-" & temp_Child2.MASTER_AREA_ID), False)        'temp_Child2.GOODS_GROUP_ID
                              If Not (temp_CollMinus Is Nothing) Then
                                  MINUS = temp_CollMinus.MINUS
                              Else
                                  MINUS = 0
                              End If
                              
                 If temp_Child2.GOODS_GROUP_ID = 2 Then
                              Set temp_CollTot = GetTotMinChart(collTotal, Trim(temp_Child2.SALE_ID & "-" & temp_Child2.MASTER_AREA_ID), False)
                              If Not (temp_CollTot Is Nothing) Then
                                  TOTAL1_AMOUNT = temp_CollTot.TOTAL1_SUM - TOTAL1_AMOUNT
                                  TOTAL2_AMOUNT = temp_CollTot.TOTAL2_SUM - TOTAL2_AMOUNT
                              Else
                                  TOTAL1_AMOUNT = 0
                                  TOTAL2_AMOUNT = 0
                              End If
                                                            
                              Set temp_CollMinus = GetTotMinChart(coll_Minus, Trim(temp_Child2.SALE_ID & "-" & temp_Child2.MASTER_AREA_ID), False)
                              If Not (temp_CollMinus Is Nothing) Then
                                  MINUS = temp_CollMinus.MINUS - MINUS
                              Else
                                  MINUS = 0
                              End If
            End If
                              
                              
                              j = j + 1                                                                            'ยอดขาย
                              
                              Set HeadCf = m_Headers1.Fields(j)
                             Total1(j) = Total1(j) + (((-1) * Val(MINUS)) + TOTAL1_AMOUNT)
                              TempStr = FormatNumber(((-1) * Val(MINUS)) + TOTAL1_AMOUNT)
                              Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
                              Call m_Details1.AddField(BodyCf)

                              j = j + 1                                                                                'minus
                                Total1(j) = Total1(j) + ((-1) * Val(MINUS))
                              Set HeadCf = m_Headers1.Fields(j)
                              TempStr = FormatNumber(((-1) * Val(MINUS)))
                              Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
                              Call m_Details1.AddField(BodyCf)
                              
                              j = j + 1                                                                                 'รวม
                              Set HeadCf = m_Headers1.Fields(j)
                              TempStr = FormatNumber(TOTAL1_AMOUNT)
                             Total1(j) = Total1(j) + TOTAL1_AMOUNT
                              Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
                              Call m_Details1.AddField(BodyCf)
                              
                              j = j + 1                                                                                  'ยอดประเมิน
                              Total1(j) = Total1(j) + TOTAL2_AMOUNT
                              Set HeadCf = m_Headers1.Fields(j)
                              TempStr = FormatNumber(TOTAL2_AMOUNT)
                              Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
                              Call m_Details1.AddField(BodyCf)
                              
                               j = j + 1                                                                            'งบประมาณ   temp_Head.Budget
                               Total1(j) = Total1(j) + temp_Child2.BUDGET
                              Set HeadCf = m_Headers1.Fields(j)
                              TempStr = FormatNumber(temp_Child2.BUDGET)
                              Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
                              Call m_Details1.AddField(BodyCf)
                              
                             If temp_Child2.BUDGET = "0" Or temp_Child2.BUDGET = "" Then
                                      temp_Child2.BUDGET = "1"
                             End If
                             
'                         If temp_Child2.SALE_ID = "15" And temp_Child2.MASTER_AREA_ID = 13 Then
'                              'debug.print
'                         End If

                               j = j + 1                                                                            ' %
                              Set HeadCf = m_Headers1.Fields(j)
                              DueCount = Round((TOTAL2_AMOUNT / Val(temp_Child2.BUDGET)) * 100, 2)
                              Total1(j) = Total1(j) + DueCount
                               TempStr = DueCount
                              Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
                              Call m_Details1.AddField(BodyCf)
                              
                                 'คำนวนคอมขาย
'''''''''            Set temp_ComMasPro = GetComMasPro(m_ComMasPro, Trim(Stcrd.SLMCOD & "-" & Stcrd.CUSCOD), False)
'''''''''            If Not (temp_ComMasPro Is Nothing) Then           ' หาโดยใช้  .SLMCOD .PEOPLE
'''''''''                     Set m_runComPro = m_ComPro1.ITEM(1)
'''''''''                     CreditKey = Val(m_runComPro.CREDIT_NAME)
'''''''''                  '   'debug.print CreditKey & " ~ " & DueCount
'''''''''                     For Each m_runComPro In m_ComPro1                   ' เก็บเงินธรรมดา
'''''''''                           If (Val(m_runComPro.CREDIT_NAME)) >= DueCount Then
'''''''''                                CreditKey = Val(m_runComPro.CREDIT_NAME)    'จะได้ เครดิตเป็น Key
'''''''''                             '   'debug.print CreditKey & " ~ " & DueCount
'''''''''                          End If
'''''''''                     Next m_runComPro
'''''''''
'''''''''                     ' วน loop ของ YEAR_ID นี้ ,, 02(สินค้าธรรมดา) อยู่ในเครดิตไหน  แล้วมันจะได้ค่า MASTER_COMMISS_SUB_PROMOTE_ID เป็น Key
'''''''''                     Set TempComPro = GetComMasSubPro(m_ComPro1, Trim(Stcrd.SLMCOD & "-" & Stcrd.CUSCOD & "-" & CreditKey), False)
'''''''''                      PercentSum = (TempComPro.DetailsCom1.ITEM(1).SLM_PERCENT / 100)  'น้อยสุดป่ะ
'''''''''
'''''''''                     For Each m_runPro In TempComPro.DetailsCom1                    ' เก็บเงินธรรมดา
'''''''''                           If (Val(m_runPro.NUM_ONE) + 7) >= DueCount Then
'''''''''                                    PercentSum = (m_runPro.SLM_PERCENT / 100)
'''''''''                          End If
'''''''''                     Next m_runPro
'''''''''             Else
                                      

                                   Set TempCConditionCommiss = m_ConditionCommiss1a(temp_Child2.GOODS_GROUP_ID).ITEM(1)
                                   PercentSum = (TempCConditionCommiss.SLM_PERCENT / 100)
                                   
                                   'หาเปอร์เซ็นของ 100
                                    For Each m_runConditionCommiss In m_ConditionCommiss1a(temp_Child2.GOODS_GROUP_ID)
                                        If (m_runConditionCommiss.NUM_ONE) > 100 Then
                                             Percent100 = (m_runConditionCommiss.SLM_PERCENT / 100)
                                            ' NUM_ONE_select = m_runConditionCommiss.NUM_ONE
                                       End If
                                    Next m_runConditionCommiss
                                   
                                   For Each m_runConditionCommiss In m_ConditionCommiss1a(temp_Child2.GOODS_GROUP_ID)
                                        If (m_runConditionCommiss.NUM_ONE) >= DueCount Then
                                             PercentSum = (m_runConditionCommiss.SLM_PERCENT / 100)
                                       End If
                                    Next m_runConditionCommiss
                                    
                                    For Each m_runConditionCommiss In m_ConditionCommiss1a(temp_Child2.GOODS_GROUP_ID)   ' ทำเพื่อถ้าอยู่ที่ 100 แต่ไม่ถึง 110 ให้เอายอดประมาณการมาคิด ตัดยอดที่เขาขายได้จริงทิ้ง
                                        If (m_runConditionCommiss.Flag) = "Y" Then
                                             NUM_ONE_Y = m_runConditionCommiss.NUM_ONE
                                       End If
                                    Next m_runConditionCommiss
'''''''                 End If

                               AMOUNT = (TOTAL2_AMOUNT * PercentSum)
                                         
                                If DueCount > 100 Then    ' กรณีเกิน 100% คิดเฉพาะยอดส่วนเกิน 100
                                 If DueCount < NUM_ONE_Y Then
                                    AMOUNT = temp_Child2.BUDGET * Percent100
                                 Else
                                     asAmount = 0
                                     MoreAmount = 0
                                     'Set m_runConditionCommiss = GetObject("CConditionCommission", m_ConditionCommiss1, "110")
                                     asAmount = temp_Child2.BUDGET * Percent100
                             
                                     MoreAmount = TOTAL2_AMOUNT - temp_Child2.BUDGET
                                     MoreAmount = MoreAmount * PercentSum
                             
                                      AMOUNT = MoreAmount + asAmount
                                   End If
                                End If
                 
                                  j = j + 1                                                ' ค่าคอมขาย
                                 Set HeadCf = m_Headers1.Fields(j)
                                 TempStr = FormatNumber(AMOUNT)               ' เอายอดขายมาคูณ ผลจากเงื่อนไข
                                 Total1(j) = Total1(j) + (AMOUNT)
                                 Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
                                 Call m_Details1.AddField(BodyCf)

                               Set temp_GPwithGroup = GetObject("CMasterFromToDetail", m_GPwithGroup, Trim(temp_Child2.SALE_ID), False)
                                If Not (temp_GPwithGroup Is Nothing) Then
                                    GP = temp_GPwithGroup.GP
                                    GROUP = temp_GPwithGroup.MASTER_PARAMETER_VALUE
                                Else
                                    GP = 0
                                    GROUP = 1
                                End If
                                
                                j = j + 1                                                ' Gp
                                 Set HeadCf = m_Headers1.Fields(j)
                                 TempStr = FormatNumber(GP)               ' เอายอดขายมาคูณ ผลจากเงื่อนไข
                                    Total1(j) = Total1(j) + (GP)
                                 Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
                                 Call m_Details1.AddField(BodyCf)
                                 
                                  j = j + 1                                                ' INDEX
                                 Set HeadCf = m_Headers1.Fields(j)
                                 TempStr = FormatNumber(GP / GROUP)             ' เอายอดขายมาคูณ ผลจากเงื่อนไข
                                    Total1(j) = Total1(j) + (GP / GROUP)
                                 Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
                                 Call m_Details1.AddField(BodyCf)
                                 
                                  j = j + 1                                                ' จ่ายจริง
                                 Set HeadCf = m_Headers1.Fields(j)
                                 TempStr = FormatNumber((AMOUNT) * (GP / GROUP))             ' เอายอดขายมาคูณ ผลจากเงื่อนไข
                                    Total1(j) = Total1(j) + ((AMOUNT) * (GP / GROUP))
                                 Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
                                 Call m_Details1.AddField(BodyCf)
        
                    If mcolParam("SUMMARY_MODE") = "N" Then
                           Call m_Details1.GetString(1, TempStr1, TempStr2)
                            strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
                            Call Vsp.AddTable(strFormat, "", TempStr2)
                            
                        If mcolParam("SAVE_MODE") = "Y" Then
                          Call SaveToDB(temp_sale.SLMCOD, temp_sale.SLMNAM, "01", temp_Child2.MASTER_AREA_ID, temp_Child2.GOODS_GROUP_ID, TempStr, mcolParam("FROM_DOC_DATE"), mcolParam("TO_DOC_DATE"))                 'PreKey3=เลขเขต ,Total2
                        End If
                            
                     End If
                            
                     End If 'งบประมาณมีค่า
                 End If ' เจอใน temp_Child2
                 
               End If   ' กรุ๊ปไม่ว่าง
          
    '     Next temp_ChildGoodsGroup
         Next temp_Child
         
        Call m_Details1.ClearField
        
        j = 0
        
        ' รวม manager
'                j = j + 1
'               Set HeadCf = m_Headers1.Fields(j)
'               TempStr = "รวม" ' & Str(i)
'                  Vsp.FontBold = True
'               Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
'               Call m_Details1.AddField(BodyCf)
'
'                j = j + 1
'               Set HeadCf = m_Headers1.Fields(j)
'               TempStr = ArS.SLMNAM
'               Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
'               Call m_Details1.AddField(BodyCf)
'
'               j = j + 1
'               Set HeadCf = m_Headers1.Fields(j)
'               TempStr = "Commercial # " & Str(i)
'               Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
'               Call m_Details1.AddField(BodyCf)
'
'               text_AREA_NAME = ""
'
'               j = j + 1                                                                            'ยอดขาย
'               Set HeadCf = m_Headers1.Fields(j)
'               TempStr = FormatNumber(Total1(j))
'               Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
'               Call m_Details1.AddField(BodyCf)
'               Total2(j) = Total2(j) + Total1(j)
'
'               j = j + 1                                                                                'minus
'               Set HeadCf = m_Headers1.Fields(j)
'               TempStr = FormatNumber(Total1(j))
'               Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
'               Call m_Details1.AddField(BodyCf)
'               Total2(j) = Total2(j) + Total1(j)
'
'               j = j + 1                                                                                 'รวม
'               Set HeadCf = m_Headers1.Fields(j)
'               TempStr = FormatNumber(Total1(j))
'               Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
'               Call m_Details1.AddField(BodyCf)
'               Total2(j) = Total2(j) + Total1(j)
'
'               j = j + 1                                                                                  'ยอดประเมิน
'               Set HeadCf = m_Headers1.Fields(j)
'               TempStr = FormatNumber(Total1(j))
'               Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
'               Call m_Details1.AddField(BodyCf)
'               Total2(j) = Total2(j) + Total1(j)
'
'                j = j + 1                                                                            'งบประมาณ   temp_Head.Budget
'               Set HeadCf = m_Headers1.Fields(j)
'               TempStr = FormatNumber(Total1(j))
'               Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
'               Call m_Details1.AddField(BodyCf)
'               Total2(j) = Total2(j) + Total1(j)
'
'               If temp_Head.BUDGET = "0" Or temp_Head.BUDGET = "" Then
'                             temp_Head.BUDGET = "1"
'                End If
'
'                j = j + 1                                                ' %
'               Set HeadCf = m_Headers1.Fields(j)
'               DueCount = Round((Total1(j) / K), 2)
'               TempStr = ""                     'DueCount
'               Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
'               Call m_Details1.AddField(BodyCf)
               
               'คำนวนคอมขาย
'''''            Set temp_ComMasPro = GetComMasPro(m_ComMasPro, Trim(Stcrd.SLMCOD & "-" & Stcrd.CUSCOD), False)
'''''            If Not (temp_ComMasPro Is Nothing) Then           ' หาโดยใช้  .SLMCOD .PEOPLE
'''''                     Set m_runComPro = m_ComPro1.ITEM(1)
'''''                     CreditKey = Val(m_runComPro.CREDIT_NAME)
'''''                  '   'debug.print CreditKey & " ~ " & DueCount
'''''                     For Each m_runComPro In m_ComPro1                   ' เก็บเงินธรรมดา
'''''                           If (Val(m_runComPro.CREDIT_NAME)) >= DueCount Then
'''''                                CreditKey = Val(m_runComPro.CREDIT_NAME)    'จะได้ เครดิตเป็น Key
'''''                             '   'debug.print CreditKey & " ~ " & DueCount
'''''                          End If
'''''                     Next m_runComPro
'''''
'''''                     ' วน loop ของ YEAR_ID นี้ ,, 02(สินค้าธรรมดา) อยู่ในเครดิตไหน  แล้วมันจะได้ค่า MASTER_COMMISS_SUB_PROMOTE_ID เป็น Key
'''''                     Set TempComPro = GetComMasSubPro(m_ComPro1, Trim(Stcrd.SLMCOD & "-" & Stcrd.CUSCOD & "-" & CreditKey), False)
'''''                      PercentSum = (TempComPro.DetailsCom1.ITEM(1).SLM_PERCENT / 100)  'น้อยสุดป่ะ
'''''
'''''                     For Each m_runPro In TempComPro.DetailsCom1                    ' เก็บเงินธรรมดา
'''''                           If (Val(m_runPro.NUM_ONE) + 7) >= DueCount Then
'''''                                    PercentSum = (m_runPro.SLM_PERCENT / 100)
'''''                          End If
'''''                     Next m_runPro
'''''             Else
''                     Set TempCConditionCommiss = m_ConditionCommiss1.ITEM(1)
''                     PercentSum = (TempCConditionCommiss.SLM_PERCENT / 100)
''                     For Each m_runConditionCommiss In m_ConditionCommiss1
''                          If (m_runConditionCommiss.NUM_ONE) >= DueCount Then
''                               PercentSum = (m_runConditionCommiss.SLM_PERCENT / 100)
''                         End If
''                      Next m_runConditionCommiss
'''''             End If
             
'                               j = j + 1                                                ' ค่าคอมขาย
'               Set HeadCf = m_Headers1.Fields(j)
'               TempStr = FormatNumber(Total1(j))              ' เอายอดขายมาคูณ ผลจากเงื่อนไข
'               Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
'               Call m_Details1.AddField(BodyCf)
'               Total2(j) = Total2(j) + Total1(j)
'
'                            j = j + 1                                                ' Gp
'            Set HeadCf = m_Headers1.Fields(j)
'            TempStr = ""               ' เอายอดขายมาคูณ ผลจากเงื่อนไข
'            Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
'            Call m_Details1.AddField(BodyCf)
'
'                                                             j = j + 1                                                ' INDEX
'            Set HeadCf = m_Headers1.Fields(j)
'            Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
'            Call m_Details1.AddField(BodyCf)
'
'                                                             j = j + 1                                                ' จ่ายจริง
'            Set HeadCf = m_Headers1.Fields(j)
'               TempStr = FormatNumber(Total1(j))
'            Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
'            Call m_Details1.AddField(BodyCf)
'            Total2(j) = Total2(j) + Total1(j)
'
'                     Call m_Details1.GetString(1, TempStr1, TempStr2)
'                    strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
'                    Call Vsp.AddTable(strFormat, "", TempStr2)
'                       Vsp.FontBold = False
               
'                  For j = 1 To UBound(Total1)
'                     Total1(j) = 0
'                  Next j
'
       End If
     
     Next temp_GoodsGroup
     
        TempRs.MoveNext                                                            ' วนเซลล์
Wend
      
      '----------------------- รวมสุทธิ
        Call m_Details1.ClearField
        j = 0
        
        ' รวม manager
                j = j + 1
               Set HeadCf = m_Headers1.Fields(j)
               TempStr = ""
                  Vsp.FontBold = True
               Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
               Call m_Details1.AddField(BodyCf)
   
                j = j + 1
               Set HeadCf = m_Headers1.Fields(j)
               TempStr = ""
               Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
               Call m_Details1.AddField(BodyCf)
               
               j = j + 1
               Set HeadCf = m_Headers1.Fields(j)
               TempStr = "สุทธิ" ' & Str(i)
               Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
               Call m_Details1.AddField(BodyCf)
               
               j = j + 1
               Set HeadCf = m_Headers1.Fields(j)
               TempStr = ""
               Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
               Call m_Details1.AddField(BodyCf)
               
               text_AREA_NAME = ""
               
               j = j + 1                                                                            'ยอดขาย
               Set HeadCf = m_Headers1.Fields(j)
               TempStr = FormatNumber(Total1(j))                                     'Total2
               Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
               Call m_Details1.AddField(BodyCf)
               
               j = j + 1                                                                                'minus
               Set HeadCf = m_Headers1.Fields(j)
               TempStr = FormatNumber(Total1(j))
               Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
               Call m_Details1.AddField(BodyCf)
               
               j = j + 1                                                                                 'รวม
               Set HeadCf = m_Headers1.Fields(j)
               TempStr = FormatNumber(Total1(j))
               Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
               Call m_Details1.AddField(BodyCf)
               
               j = j + 1                                                                                  'ยอดประเมิน
               Set HeadCf = m_Headers1.Fields(j)
               TempStr = FormatNumber(Total1(j))
               Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
               Call m_Details1.AddField(BodyCf)
               
                j = j + 1                                                                            'งบประมาณ   temp_Head.Budget
               Set HeadCf = m_Headers1.Fields(j)
               TempStr = FormatNumber(Total1(j))
               Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
               Call m_Details1.AddField(BodyCf)
               
                j = j + 1                                                ' %
               Set HeadCf = m_Headers1.Fields(j)
               TempStr = ""                     'DueCount
               Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
               Call m_Details1.AddField(BodyCf)
               
                               j = j + 1                                                ' ค่าคอมขาย
               Set HeadCf = m_Headers1.Fields(j)
               TempStr = FormatNumber(Total1(j))              ' เอายอดขายมาคูณ ผลจากเงื่อนไข
               Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
               Call m_Details1.AddField(BodyCf)
               
                            j = j + 1                                                ' Gp
            Set HeadCf = m_Headers1.Fields(j)
            TempStr = ""               ' เอายอดขายมาคูณ ผลจากเงื่อนไข
            Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
            Call m_Details1.AddField(BodyCf)
            
                                                             j = j + 1                                                ' INDEX
            Set HeadCf = m_Headers1.Fields(j)
            Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
            Call m_Details1.AddField(BodyCf)
            
                                                             j = j + 1                                                ' จ่ายจริง
            Set HeadCf = m_Headers1.Fields(j)
               TempStr = FormatNumber(Total1(j))
            Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
            Call m_Details1.AddField(BodyCf)
               
                     Call m_Details1.GetString(1, TempStr1, TempStr2)
                    strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
                    Call Vsp.AddTable(strFormat, "", TempStr2)
                       Vsp.FontBold = False
               
      '-------------------

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
      
'      Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.Align, TempStr, HeadCf.BodyAlign)
'      Call Details.AddField(BodyCf)
      
      If HaveValue = False And Tot(j) <> 0 Then
         HaveValue = True
      End If
   Next HeadCf
      
   Set BodyCf = Nothing
End Sub

Private Sub SaveToDB(Txt1 As String, Txt2 As String, COMTYP As String, Txt3 As Long, Txt4 As Long, Tot As String, FromDate As Date, ToDate As Date)
' Call SavetoDB( PreKey1, PreKey2 , "02", PreKey3 ,Total2 , mcolParam("FROM_CMPL_DATE"), mcolParam("TO_CMPL_DATE"))
Dim IsOK As Boolean
Dim exit_Record As CComRecord
Dim comRecord As CComRecord
Set comRecord = New CComRecord

Set exit_Record = GetObject("CCom_Record", m_ComRecord, Txt1 & "-" & Txt3 & "-" & Txt4, False)
If exit_Record Is Nothing Then
        
  If ConfirmSaveRecord(Txt2 & " " & Tot & " บ.") Then
      comRecord.AddEditMode = SHOW_ADD
      comRecord.SLMCOD = Txt1
      comRecord.SLMNAME = Txt2
      comRecord.COMTYP = COMTYP
      comRecord.MASTER_AREA_ID = Txt3
      comRecord.GOODS_GROUP_ID = Txt4
      comRecord.COM_VALUE = Tot
      comRecord.FROMDAT = FromDate
      comRecord.TODAT = ToDate
   
      If Not glbDaily.AddEditComRecord(comRecord, IsOK, True, glbErrorLog) Then
         glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
      End If
   End If
Else
     If ConfirmSaveExist(Txt2 & " " & Tot & " บ.", exit_Record.SLMNAME & " " & exit_Record.COM_VALUE & " บ.") Then
         comRecord.AddEditMode = SHOW_EDIT
         comRecord.RECORD_ID = exit_Record.RECORD_ID
         comRecord.SLMCOD = Txt1
         comRecord.SLMNAME = Txt2
         comRecord.COMTYP = COMTYP
         comRecord.MASTER_AREA_ID = Txt3
         comRecord.GOODS_GROUP_ID = Txt4
         comRecord.COM_VALUE = Tot
         comRecord.FROMDAT = FromDate
         comRecord.TODAT = ToDate
         
          If Not glbDaily.AddEditComRecord(comRecord, IsOK, True, glbErrorLog) Then
            glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
         End If
      End If
End If
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


