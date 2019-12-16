VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CReportCom02"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'User Group Report
Implements CReportInterface

Private Const MODULE_NAME = "CReportCOM02"
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

Private m_ComPro2 As Collection
Private m_ComPro3 As Collection
Private m_ComPro4 As Collection
Private m_ComMasPro As Collection
Private temp_ComMasPro As CCommissMasterPromote

'Private m_cusFromArea(20) As Collection
Private nonidentifyArea As Boolean
Dim Stb As CCommissMasterArea
Private m_MinusStkcod As Collection
Private m_AllSaleInIV As Collection
Private m_IVincomplete0 As Collection
Private m_IVincomplete12 As Collection
Private m_ReDocdat As Collection
Private temp_Area As CCommissionCustomerArea
Private temp_Area2 As CCommissionCustomerArea  ' เอาไว้อ้าง ID นี้ ชื่อเขตอะไร
Public REsumIV As CARRcIt
Private m_REsumIV As Collection

Private tempIVcredit As CComIVcredit
Private m_IVcredit As Collection

Private tempComDonStkcod As CComDonStk
Private m_ComDonStkcod As Collection
Dim L As Long

Dim GOODS_GROUP_ID As Long
Dim GOODS_MASTER_ID As Long

Private temp_GoodsGroup As CGoodsGroup
Private m_GoodsGroup As Collection
Private temp_GoodsInGroup As CGoodsDetail
Private m_GoodsInGroup(10) As Collection
Private TempCConditionCommiss As CConditionCommission
Private m_ConditionCommiss2a(10) As Collection
Dim n_array As Long

Private FromSaleID As String
Private ToSaleID As String
Private m_ComRecord As Collection

Private stcrd_mixdb As Collection

'Private havevalue As Boolean

Private Sub Class_Initialize()
   Set stcrd_mixdb = New Collection

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
   
   For n_array = 0 To UBound(m_ConditionCommiss2a)
      Set m_ConditionCommiss2a(n_array) = New Collection
   Next n_array
   
   Set tempIVcredit = New CComIVcredit
   Set m_IVcredit = New Collection
   Set m_ComPro2 = New Collection
   Set m_ComPro3 = New Collection
   Set m_ComPro4 = New Collection
   Set m_ComMasPro = New Collection
   Set temp_ComMasPro = New CCommissMasterPromote
   
   Set Stb = New CCommissMasterArea
   Set m_AllSaleInIV = New Collection
   Set m_MinusStkcod = New Collection
   Set m_IVincomplete0 = New Collection
   Set m_IVincomplete12 = New Collection
   Set m_REsumIV = New Collection
   Set m_ReDocdat = New Collection
   Set temp_Area = New CCommissionCustomerArea
   Set temp_Area2 = New CCommissionCustomerArea
   Set REsumIV = New CARRcIt
   Set m_ComDonStkcod = New Collection
   Set tempComDonStkcod = New CComDonStk
   
   Set m_GoodsGroup = New Collection
   Set temp_GoodsGroup = New CGoodsGroup
   For n_array = 0 To UBound(m_GoodsInGroup)
      Set m_GoodsInGroup(n_array) = New Collection
   Next n_array
   
   Set m_ComRecord = New Collection
   Set temp_GoodsInGroup = New CGoodsDetail
End Sub

Private Sub Class_Terminate()
   Call ClearParam
   Set stcrd_mixdb = Nothing
   Set mcolParam = Nothing
   Set Vsp = Nothing
   Set m_TempCol = Nothing
   
   Set m_IVcredit = Nothing
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
 
   For n_array = 0 To UBound(m_ConditionCommiss2a)
      Set m_ConditionCommiss2a(n_array) = Nothing
   Next n_array
   
   Set m_ComPro2 = Nothing
   Set m_ComPro3 = Nothing
   Set m_ComPro4 = Nothing
   Set m_ComMasPro = Nothing
   Set temp_ComMasPro = Nothing
   Set Stb = Nothing
   Set m_MinusStkcod = Nothing
   Set m_IVincomplete0 = Nothing
   Set m_IVincomplete12 = Nothing
   Set m_REsumIV = Nothing
   Set m_ReDocdat = Nothing
   Set temp_Area = Nothing
   Set temp_Area2 = Nothing
   Set REsumIV = Nothing
   Set m_ComDonStkcod = Nothing
   Set tempComDonStkcod = Nothing
    Set m_AllSaleInIV = Nothing
    Set tempIVcredit = Nothing
    
   Set m_GoodsGroup = Nothing
   Set temp_GoodsGroup = Nothing
   For n_array = 0 To UBound(m_GoodsInGroup)
      Set m_GoodsInGroup(n_array) = Nothing
   Next n_array
   Set temp_GoodsInGroup = Nothing
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
      
 If mcolParam("SUMMARY_MODE") = "N" Then
   Call Cf.SetFieldValue(3, "^", "", "^")
   Call m_Headers0.AddField(Cf)
   Call Cf.SetFieldValue(26, "^", "", "<")    '27
   Call m_Headers0.AddField(Cf)
Else
   Call Cf.SetFieldValue(2, "^", "", "^")
   Call m_Headers0.AddField(Cf)
   Call Cf.SetFieldValue(14, "^", "", "<")    '27
   Call m_Headers0.AddField(Cf)
End If
      
      
If mcolParam("SUMMARY_MODE") = "N" Then
      Call Cf.SetFieldValue(1, "^", "NO.", "^")
      Call m_Headers1.AddField(Cf)
      
      Call Cf.SetFieldValue(2, "^", "INVOICE", "^")
      Call m_Headers1.AddField(Cf)
      
      Call Cf.SetFieldValue(6, "^", "ชื่อลูกค้า", "<")
      Call m_Headers1.AddField(Cf)
      
      Call Cf.SetFieldValue(4, "^", "ชื่อสินค้า", "<")
      Call m_Headers1.AddField(Cf)
Else
      Call Cf.SetFieldValue(0, "^", "NO.", "^")
      Call m_Headers1.AddField(Cf)
      
      Call Cf.SetFieldValue(2, "^", "รหัส", ">")
      Call m_Headers1.AddField(Cf)
      
      Call Cf.SetFieldValue(6, "^", "พนักงานขาย", "<")
      Call m_Headers1.AddField(Cf)
      
      Call Cf.SetFieldValue(0, "^", "ชื่อสินค้า", "<")
      Call m_Headers1.AddField(Cf)
End If


   Call Cf.SetFieldValue(2, "^", "ยอดขาย", ">")
   Call m_Headers1.AddField(Cf)
      
   Call Cf.SetFieldValue(2, "^", "ยอดประเมิน", ">")
   Call m_Headers1.AddField(Cf)
      
   Call Cf.SetFieldValue(2, "^", "ยอดคิดค่า com", ">")
   Call m_Headers1.AddField(Cf)
        
 If mcolParam("SUMMARY_MODE") = "N" Then
   Call Cf.SetFieldValue(2, "^", "วันส่งสินค้า", "^")
   Call m_Headers1.AddField(Cf)
      
   Call Cf.SetFieldValue(2, "^", "วันรับชำระ", "^")
    Call m_Headers1.AddField(Cf)
      
   Call Cf.SetFieldValue(2, "^", "เลขที่ใบเสร็จ", "^")
    Call m_Headers1.AddField(Cf)

   Call Cf.SetFieldValue(1, "^", "Cr.", "^")
   Call m_Headers1.AddField(Cf)
   
   Call Cf.SetFieldValue(1, "^", "%", "^")
   Call m_Headers1.AddField(Cf)
 Else
    Call Cf.SetFieldValue(0, "^", "วันส่งสินค้า", "^")
   Call m_Headers1.AddField(Cf)
      
   Call Cf.SetFieldValue(0, "^", "วันรับชำระ", "^")
    Call m_Headers1.AddField(Cf)
      
   Call Cf.SetFieldValue(0, "^", "เลขที่ใบเสร็จ", "^")
    Call m_Headers1.AddField(Cf)

   Call Cf.SetFieldValue(0, "^", "Cr", "^")
   Call m_Headers1.AddField(Cf)
   
   Call Cf.SetFieldValue(0, "^", "%", "^")
   Call m_Headers1.AddField(Cf)
 End If
 
      Call Cf.SetFieldValue(2, "^", "ค่าคอมมิชชั่น", ">")
      Call m_Headers1.AddField(Cf)

'      Call Cf.SetFieldValue(2, "^", "วันที่จ่ายค่าคอม", ">")
'      Call m_Headers1.AddField(Cf)
   
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
Dim YEAR_ID As Long

   mstrErrMsg = ""
   mblnHeader = True
   mblnNewPage = True
   mblnEndPage = True
   Vsp.PaperSize = pprA4 '
If mcolParam("SUMMARY_MODE") = "N" Then
   Vsp.ORIENTATION = orLandscape             ' orPortrait
Else
   Vsp.ORIENTATION = orPortrait
End If
   Vsp.MarginBottom = 300
   Vsp.MarginFooter = 300
   Vsp.MarginHeader = 300
   Vsp.MarginLeft = 300
   Vsp.MarginRight = 500
   Vsp.MarginTop = 300
   Vsp.FontName = "AngsanaUPC"
   Vsp.FontSize = DETAIL_SIZE
   
  Call LoadComNamePro(m_ComMasPro, mcolParam("FROM_CMPL_DATE"), mcolParam("TO_CMPL_DATE"))       'รายชื่อ เซลล์และลูกค้าคนไหนทีเข้าข่าย
   Call LoadComPro("02", m_ComPro2, mcolParam("FROM_CMPL_DATE"), mcolParam("TO_CMPL_DATE"))      'เก็บเงินธรรมดา
   Call LoadComPro("03", m_ComPro3, mcolParam("FROM_CMPL_DATE"), mcolParam("TO_CMPL_DATE"))      'เก็บเงินพิเศษ
   
   n_array = 0
   Call LoadColGoodsGroup(m_GoodsGroup)                ' มีประเภทไรบ้าง
      ' ดึง ชื่อสินค้าที่อยู่ใน master นั่นน และกรุ๊ปนั่นนนน
   For Each temp_GoodsGroup In m_GoodsGroup
      Call LoadCommission("02", m_ConditionCommiss2a(n_array), mcolParam("FROM_CMPL_DATE"), mcolParam("TO_CMPL_DATE"), temp_GoodsGroup.GOODS_GROUP_ID)
      Set TempCConditionCommiss = m_ConditionCommiss2a(n_array).ITEM(1)
      Call LoadGoodsDetailFromGM_GG(m_GoodsInGroup(n_array), TempCConditionCommiss.GOODS_MASTER_ID, temp_GoodsGroup.GOODS_GROUP_ID)                 ' มีประเภทไรบ้าง
      n_array = n_array + 1
   Next temp_GoodsGroup
   Call LoadCommission("02", m_ConditionCommiss2a(UBound(m_ConditionCommiss2a)), mcolParam("FROM_CMPL_DATE"), mcolParam("TO_CMPL_DATE"), 2)   '<< ต้องแก้ 2 ให้เป็น defail

   
   ' ต้องก๊อปอันบน
'  Call LoadCommission("02", m_ConditionCommiss2, mcolParam("FROM_CMPL_DATE"), mcolParam("TO_CMPL_DATE"), temp_GoodsGroup.GOODS_GROUP_ID)

   Call LoadCommission("03", m_ConditionCommiss3, mcolParam("FROM_CMPL_DATE"), mcolParam("TO_CMPL_DATE"), TempCConditionCommiss.GOODS_GROUP_ID)
   Call LoadCommission04(m_ConditionCommiss4, mcolParam("FROM_CMPL_DATE"), mcolParam("TO_CMPL_DATE"), TempCConditionCommiss.GOODS_GROUP_ID)

  Dim bak_ID As Long
'  Dim FromSaleID As String
'  Dim ToSaleID As String
  Dim FromSaleName As String
  Dim ToSaleName As String
  
   bak_ID = InStr(1, mcolParam("FROM_SALE_NAME"), " ")
   If bak_ID > 0 Then
     FromSaleID = Left(mcolParam("FROM_SALE_NAME"), bak_ID - 1)
     FromSaleName = Mid(mcolParam("FROM_SALE_NAME"), bak_ID + 3)
   End If
   bak_ID = InStr(1, mcolParam("TO_SALE_NAME"), " ")
   If bak_ID > 0 Then
      ToSaleID = Left(mcolParam("TO_SALE_NAME"), bak_ID - 1)
     ToSaleName = Mid(mcolParam("TO_SALE_NAME"), bak_ID + 3)
   End If
      
   MODULE_DESC = glbCompanyName & vbCrLf & mcolParam("REPORT_NAME") & _
                             vbCrLf & "วันที่ส่งสินค้า   " & EmptyToString(DateToStringExtEx2(mcolParam("FROM_DOC_DATE")), "N/A") & "    ถึง    " & EmptyToString(DateToStringExtEx2(mcolParam("TO_DOC_DATE")), "N/A") & "      วันที่รับชำระ    " & EmptyToString(DateToStringExtEx2(mcolParam("FROM_CMPL_DATE")), "N/A") & "    ถึง    " & EmptyToString(DateToStringExtEx2(mcolParam("TO_CMPL_DATE")), "N/A")
      If Len(mcolParam("FROM_SALE_NAME")) > 0 Then
              MODULE_DESC = MODULE_DESC & vbCrLf & "    พนักงานขาย  " & EmptyToString(FromSaleName, "N/A")
      End If
       If Len(mcolParam("TO_SALE_NAME")) > 0 Then
              MODULE_DESC = MODULE_DESC & "   ถึง    " & EmptyToString(ToSaleName, "N/A")
      End If
       If Len(mcolParam("AREA_TYPE_NAME")) > 0 Then
              MODULE_DESC = MODULE_DESC & "      เขต :  " & EmptyToString(mcolParam("AREA_TYPE_NAME"), "N/A")
      End If
      
  Call LoadIVcredit(m_IVcredit, mcolParam("FROM_DOC_DATE"), mcolParam("TO_DOC_DATE"))
  Call LoadMinusIV(m_MinusStkcod, mcolParam("FROM_DOC_DATE"), mcolParam("TO_DOC_DATE"))
  Call LoadREDocDat(m_ReDocdat, mcolParam("FROM_CMPL_DATE"), mcolParam("TO_CMPL_DATE")) ' ต้องเอามาหมด
 ' Call LoadREDocDat(m_ReDocdat, mcolParam("FROM_CMPL_DATE"), mcolParam("TO_CMPL_DATE"), True) ' ต้องเอามาหมด ของปี db2
  Call LoadComDonStk(m_ComDonStkcod, mcolParam("FROM_DOC_DATE"), mcolParam("TO_DOC_DATE"))
  Call LoadREsumIV(Nothing, m_REsumIV)
'  Call LoadAllSaleInIV(m_AllSaleInIV, mcolParam("FROM_DOC_DATE"), mcolParam("TO_DOC_DATE"))

   Call LoadYearId(YEAR_ID, mcolParam("FROM_CMPL_DATE"), mcolParam("TO_CMPL_DATE"))   ' เวลานี้ใช้ Area_Year ไหน
   If mcolParam("AREA_TYPE_ID") > 0 Then
      Call LoadAreaComReport(Nothing, m_AreaCod, YEAR_ID, Val(Left(mcolParam("AREA_TYPE_NAME"), 2)))
   Else
      Call LoadAreaComReport(Nothing, m_AreaCod, YEAR_ID)       ' โหลดไว้สำหรับ combo
   End If
    
   ' Save_mode
   If mcolParam("SAVE_MODE") = "Y" Then
       Call LoadComRecord(m_ComRecord, "02", mcolParam("TO_CMPL_DATE"))
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
Dim num As Long
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
Dim Total4(100) As Double
Dim iCount As Long
Dim TempStr As String
Dim PrevKey1 As String
Dim PrevKey2 As String
Dim PrevKey3 As Long   'รหัสเขต
Dim PrevKey4 As String    'เขตการขาย
Dim PrevKey5 As String    'วันที่
Dim St As CCommissMasterArea
Dim ArS As COESLM
Dim Stcrd As CStcrd
Dim strDueCount As String

Dim temp_AllSaleInIV As CStcrd

Dim haveValueSale As Boolean
Dim haveData As Boolean
Dim haveData_ByArea As Boolean
Dim DueCount As Long

'Dim TempCConditionCommiss As CConditionCommission
Dim m_runConditionCommiss As CConditionCommission
Dim PercentNum1 As Double
Dim PercentSum As Double

Dim CreditKey As Long

Dim tempCusArea As CCommissionCustomerArea
Dim toCMPLdate As Date
Dim tempMinusStkcod As CComMinusStk
Dim tempREdoc As CARTrn
Dim NETVAL As Double
Dim strNETVAL As String
Dim PayCompleted  As Boolean
Dim CMPLDAT As Date
Dim RECOD As String
Dim CorrectStkcod As Boolean

Dim TempComPro As CComMasSubPromote
Dim m_runComPro As CComMasSubPromote
Dim m_runPro As CCommissPromote
Dim stcrd_temp As CStcrd

Dim CheckMinusFlag As Boolean
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
      Total4(j) = 0
   Next j
   
   Call StartExportFile(Vsp)
   Vsp.StartDoc
   i = 1  'นับจ.บรรทัดข้อมูลย่อย
   K = 1
   Set BodyCf = New CReportField
      
   If mcolParam("TO_CMPL_DATE") < 0 Then
         toCMPLdate = DateSerial(9999, 12, 31)
   Else:
          toCMPLdate = mcolParam("TO_CMPL_DATE")
   End If
      
   Set ArS = New COESLM
   ArS.FROM_SALE_CODE = FromSaleID
   ArS.TO_SALE_CODE = ToSaleID
   Call glbDaily.QuerySale(ArS, TempRs, iCount, IsOK, glbErrorLog)

   While Not TempRs.EOF          ' sale
   
         Call m_Details1.ClearField
         Call ArS.PopulateFromRS(1, TempRs)
         
         For j = 1 To UBound(Total3)
            Total3(j) = 0
         Next j

         Call GenerateHeader("( " & ArS.SLMCOD & " )", ArS.SLMNAM, m_Details1)
         ' If mcolParam("SUMMARY_MODE") = "N" Then
             Call m_Details1.GetString(1, TempStr1, TempStr2)
             strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
             Vsp.FontBold = True
             Call Vsp.AddTable(strFormat, "", TempStr2)
             Vsp.FontBold = False
         'End If
         PrevKey1 = ArS.SLMCOD
         PrevKey2 = ArS.SLMNAM
         
           Set Stcrd = New CStcrd
      Stcrd.FROM_DOC_DATE = mcolParam("FROM_DOC_DATE")
      Stcrd.TO_DOC_DATE = mcolParam("TO_DOC_DATE")
      Stcrd.FROM_CUSTOMER_CODE = mcolParam("FROM_CUSTOMER_CODE")
      Stcrd.TO_CUSTOMER_CODE = mcolParam("TO_CUSTOMER_CODE")
      Stcrd.SLMCOD = ArS.SLMCOD
      Call Stcrd.QueryData(13, Rs, iCount, 2) ' ติดต่อกับดาต้าเบสที่ 2
      While Not Rs.EOF
         Call Stcrd.PopulateFromRS(13, Rs)
            Set stcrd_temp = GetObject("CStcrd", stcrd_mixdb, Stcrd.DOCNUM & "-" & Stcrd.STKCOD, False)
            If stcrd_temp Is Nothing Then  ' ถ้าไม่มีในคอเล็กชั่น คือ เพิ่มได้
               Set stcrd_temp = New CStcrd
               stcrd_temp.TRNQTY = Stcrd.TRNQTY
            '   stcrd_temp.CMPLDAT = Stcrd.CMPLDAT
               stcrd_temp.NETVAL = Stcrd.NETVAL
               stcrd_temp.DOCNUM = Stcrd.DOCNUM
               stcrd_temp.CUSNAM = Stcrd.CUSNAM
               stcrd_temp.STKDES = Stcrd.STKDES
               stcrd_temp.DOCDAT = Stcrd.DOCDAT
               stcrd_temp.STKCOD = Stcrd.STKCOD
               stcrd_temp.SLMCOD = Stcrd.SLMCOD
               stcrd_temp.CUSCOD = Stcrd.CUSCOD
               stcrd_temp.UNITPR = Stcrd.UNITPR
               Call stcrd_mixdb.Add(stcrd_temp, stcrd_temp.DOCNUM & "-" & stcrd_temp.STKCOD)
               Set stcrd_temp = Nothing
            End If
         Rs.MoveNext                                                                                            ' วนคน    '  **** วนทุกบรรทัดของ Stcrd
      Wend
      
      Set Stcrd = New CStcrd
      Stcrd.FROM_DOC_DATE = mcolParam("FROM_DOC_DATE")
      Stcrd.TO_DOC_DATE = mcolParam("TO_DOC_DATE")
      Stcrd.FROM_CUSTOMER_CODE = mcolParam("FROM_CUSTOMER_CODE")
      Stcrd.TO_CUSTOMER_CODE = mcolParam("TO_CUSTOMER_CODE")
      Stcrd.SLMCOD = ArS.SLMCOD
      Call Stcrd.QueryData(13, Rs, iCount)   ' ติดต่อกับดาต้าเบสที่ 1
      While Not Rs.EOF     '  **** วนทุกบรรทัดของ Stcrd
         Call Stcrd.PopulateFromRS(13, Rs)
            Set stcrd_temp = GetObject("CStcrd", stcrd_mixdb, Stcrd.DOCNUM & "-" & Stcrd.STKCOD, False)
            If stcrd_temp Is Nothing Then  ' ถ้าไม่มีในคอเล็กชั่น คือ เพิ่มได้
             Set stcrd_temp = New CStcrd
             stcrd_temp.TRNQTY = Stcrd.TRNQTY
      '       stcrd_temp.CMPLDAT = Stcrd.CMPLDAT
             stcrd_temp.NETVAL = Stcrd.NETVAL
             stcrd_temp.DOCNUM = Stcrd.DOCNUM
             stcrd_temp.CUSNAM = Stcrd.CUSNAM
             stcrd_temp.STKDES = Stcrd.STKDES
             stcrd_temp.DOCDAT = Stcrd.DOCDAT
             stcrd_temp.STKCOD = Stcrd.STKCOD
             stcrd_temp.SLMCOD = Stcrd.SLMCOD
             stcrd_temp.CUSCOD = Stcrd.CUSCOD
             stcrd_temp.UNITPR = Stcrd.UNITPR
             Call stcrd_mixdb.Add(stcrd_temp, stcrd_temp.DOCNUM & "-" & stcrd_temp.STKCOD)
             Set stcrd_temp = Nothing
            End If
         Rs.MoveNext                                                                                            ' วนคน    '  **** วนทุกบรรทัดของ Stcrd
      Wend
      

  L = 0              ' เปลี่ยนเซลล์ใหม่ ก็รอบเขตใหม่
  For Each temp_Area In m_AreaCod
  L = L + 1        ' ถ้าไม่ได้เลือก = ทุกเขต = วนเรื่อยๆ
        
    Set Stcrd = Nothing
    For Each Stcrd In stcrd_mixdb
    
      CorrectStkcod = False
      Set tempComDonStkcod = GetComDonStk(m_ComDonStkcod, Trim(Stcrd.STKCOD), False)
      If (tempComDonStkcod Is Nothing) Then                   ' มีในสินค้าที่ห้ามคิด com
         CorrectStkcod = True
      End If
  
  Set tempCusArea = GetCusAreaCom(temp_Area.ImportExportItems, Stcrd.CUSCOD, False)     ' !!!!!         ดึงคอลเล็คชั่นลูกที่อยู่ใน คอลเล็คชั่นแม่มาเก็บไว้
  If (Not (tempCusArea Is Nothing)) And CorrectStkcod Then                 ' อยู่ในเขตที่ผู้ใช้เลือก หรือ ผู้ใช้ไม่ได้เลือกเลย Or nonidentifyArea
 
      ' คำนวณส่วนลดก่อน
      Set tempMinusStkcod = GetMinusCommiss(m_MinusStkcod, Stcrd.DOCDAT & "-" & Stcrd.DOCNUM & "-" & Stcrd.STKCOD, False)
      If (tempMinusStkcod Is Nothing) Then   'ไม่เจอในคอเล็คชั่น = ไม่ใช่สินค้าพิเศษ
               NETVAL = Stcrd.NETVAL
               strNETVAL = FormatNumber(NETVAL)
      Else:
               NETVAL = (Stcrd.NETVAL + Val(tempMinusStkcod.MINUS_AMOUNT))
               strNETVAL = "*" & FormatNumber(NETVAL)
      End If
 
  ' เลือกมาเฉพาะที่จ่ายครบแล้ว RE.Key(IV)-NETVAL <=0 then
'   If Stcrd.DOCNUM = "IV0048055" Then      ' IV0043347
'    'debug.print
'   End If
  
  Set REsumIV = GetARRcpItem(m_REsumIV, Stcrd.DOCNUM, False)
  If Not (REsumIV Is Nothing) Then
         RECOD = REsumIV.RCPNUM
         Set tempREdoc = GetREDocDat(m_ReDocdat, Trim(Stcrd.DOCNUM), False)              ' เด๋วต้องแก้ วันที่รับชำระที่นี่
         If Not (tempREdoc Is Nothing) Then
             CMPLDAT = tempREdoc.DOCDAT
         Else
            CMPLDAT = -1
         End If
'
'      If i = 35 Then
'         'debug.print
'      End If
         If ((REsumIV.RCVAMT - NETVAL) >= 0) And CMPLDAT >= mcolParam("FROM_CMPL_DATE") And CMPLDAT <= mcolParam("TO_CMPL_DATE") Then
                 If (PrevKey5 <> Str(Month(Stcrd.DOCDAT)) & "/" & Str(Year(Stcrd.DOCDAT) + 543)) Or tempCusArea.MASTER_AREA_ID <> PrevKey3 Then        'And PrevKey5 <> ""
                        If mcolParam("SUMMARY_MODE") = "N" Then
                           Call GenerateFooter("", "      รวม      ", PrevKey5, Total4, m_Details1, haveData_ByArea)       ' เช็ค haveValue ใน Call GenerateFooter **** เซลล์ต้องมีค่าถึงจะแสดงบรรทัดรวม
                           Vsp.FontBold = True
                        Else
                           Call GenerateFooter("", "", PrevKey5, Total4, m_Details1, haveData_ByArea)        ' เช็ค haveValue ใน Call GenerateFooter **** เซลล์ต้องมีค่าถึงจะแสดงบรรทัดรวม
                           Vsp.FontBold = False
                        End If
                      
                         If mcolParam("SHOWZERO_MODE") = "Y" And haveData_ByArea = True Then             ' ติ๊ก = โชว์ 0 ด้วย
                               Call m_Details1.GetString(2, TempStr1, TempStr2)
                               strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
                               Call Vsp.AddTable(strFormat, "", TempStr2)
                                For j = 1 To UBound(Total4)
                                    Total4(j) = 0
                                 Next j
                         Else
                            If Total4(13) <> 0 Then
                               Call m_Details1.GetString(2, TempStr1, TempStr2)
                               strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
                               Call Vsp.AddTable(strFormat, "", TempStr2)
                               For j = 1 To UBound(Total4)
                                    Total4(j) = 0
                                Next j
                             End If
                         End If
                          Vsp.FontBold = False
                  End If
               
                  If tempCusArea.MASTER_AREA_ID <> PrevKey3 Then           '!
'                       If mcolParam("SUMMARY_MODE") = "N" Then
                           Call GenerateFooter("", "      รวม      ", PrevKey4, Total2, m_Details1, haveData_ByArea)        ' เช็ค haveValue ใน Call GenerateFooter **** เซลล์ต้องมีค่าถึงจะแสดงบรรทัดรวม
                           Vsp.FontBold = True
'                      Else
'                           Call GenerateFooter("", "รวม", PrevKey4, Total2, m_Details1, haveData_ByArea)        ' เช็ค haveValue ใน Call GenerateFooter **** เซลล์ต้องมีค่าถึงจะแสดงบรรทัดรวม
'                           Vsp.FontBold = False
'                      End If
                      
                      If haveData_ByArea = True Then
                        K = K + 1
                        Call m_Details1.GetString(2, TempStr1, TempStr2)
                        strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
                        Call Vsp.AddTable(strFormat, "", TempStr2)
                        Vsp.FontBold = False
                        
                        If mcolParam("SAVE_MODE") = "Y" Then
                          Call SaveToDB(PrevKey1, PrevKey2, "02", PrevKey3, 0, Total2, mcolParam("FROM_CMPL_DATE"), mcolParam("TO_CMPL_DATE"))             'PreKey3=เลขเขต ,Total2
                        End If
                        
                     End If
                     
                     If mcolParam("SUMMARY_MODE") = "N" Then
                        Set temp_Area2 = GetAreaCom(m_AreaCod, tempCusArea.MASTER_AREA_ID)
                        Call GenerateHeader("เขต : ", temp_Area2.MASTER_AREA_NAME, m_Details1)   '!
                        Call m_Details1.GetString(1, TempStr1, TempStr2)
                        strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
                        Vsp.FontBold = True
                        Call Vsp.AddTable(strFormat, "", TempStr2)
                        Vsp.FontBold = False
                     End If
                  
                     haveData_ByArea = False
                     For j = 1 To UBound(Total2)
                         Total2(j) = 0
                     Next j
             End If '      รวม      ของเดือน
         
            Call m_Details1.ClearField
            j = 0
            haveData = False
         
             j = j + 1
            Set HeadCf = m_Headers1.Fields(j)
            TempStr = i
            Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
            Call m_Details1.AddField(BodyCf)

             j = j + 1
            Set HeadCf = m_Headers1.Fields(j)
            TempStr = Stcrd.DOCNUM
            Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
            Call m_Details1.AddField(BodyCf)

              j = j + 1
            Set HeadCf = m_Headers1.Fields(j)
            TempStr = Left(Stcrd.CUSNAM, 44)
            Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
            Call m_Details1.AddField(BodyCf)
   
            j = j + 1
            Set HeadCf = m_Headers1.Fields(j)
            TempStr = Stcrd.STKDES
            Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
            Call m_Details1.AddField(BodyCf)

           PayCompleted = True
           RECOD = REsumIV.RCPNUM
           Set tempREdoc = GetREDocDat(m_ReDocdat, REsumIV.RCPNUM, False)
            If Not (tempREdoc Is Nothing) Then
                CMPLDAT = tempREdoc.DOCDAT
            End If
           
            j = j + 1
            Set HeadCf = m_Headers1.Fields(j) ' เป็น Collumn ที่ 5
            TempStr = strNETVAL
            Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
            Call m_Details1.AddField(BodyCf)
            
            '   Set m_runConditionCommiss = New CConditionCommission
            ' คำนวณค่าคอม
            CreditKey = 0
            
               DueCount = DateDiff("D", Stcrd.DOCDAT, CMPLDAT)  'ต้องคิดตรงนี้เพื่อไปคำนวณค่า
               strDueCount = DueCount
                Set tempIVcredit = GetObject("CComIVcredit", m_IVcredit, Stcrd.DOCNUM, False)
                 If Not (tempIVcredit Is Nothing) Then
                     If tempIVcredit.CR_TYPE = "I" Then
                        DueCount = DueCount + tempIVcredit.CR_DATA
                     Else
                         DueCount = DueCount - tempIVcredit.CR_DATA
                     End If
                     strDueCount = "*" & Str(DueCount)
                 End If

               Set TempCConditionCommiss = GetCheckCommiss(m_ConditionCommiss4, Stcrd.STKCOD, False)
               If (TempCConditionCommiss Is Nothing) Then   'ไม่เจอในคอเล็คชั่น = ไม่ใช่สินค้าพิเศษ
                    PercentNum1 = (100 / 100)
                    TempStr = ""
                  
                     ' ถ้าเป็น เซล์และลูกค้าโปรโมต ใช้ comPro5
                     Set temp_ComMasPro = GetComMasPro(m_ComMasPro, Trim(Stcrd.SLMCOD & "-" & Stcrd.CUSCOD), False)
                     If Not (temp_ComMasPro Is Nothing) Then           ' หาโดยใช้  .SLMCOD .PEOPLE
                         Set m_runComPro = m_ComPro2.ITEM(1)
                         CreditKey = Val(m_runComPro.CREDIT_NAME)
                         For Each m_runComPro In m_ComPro2                   ' เก็บเงินธรรมดา
                               If (Val(m_runComPro.CREDIT_NAME)) >= DueCount Then
                                    CreditKey = Val(m_runComPro.CREDIT_NAME)    'จะได้ เครดิตเป็น Key
                               End If
                          Next m_runComPro
                           
                          ' วน loop ของ YEAR_ID นี้ ,, 02(สินค้าธรรมดา) อยู่ในเครดิตไหน  แล้วมันจะได้ค่า MASTER_COMMISS_SUB_PROMOTE_ID เป็น Key
                          Set TempComPro = GetComMasSubPro(m_ComPro2, Trim(Stcrd.SLMCOD & "-" & Stcrd.CUSCOD & "-" & CreditKey), False)
                          PercentSum = (TempComPro.DetailsCom2.ITEM(1).SLM_PERCENT / 100)  'น้อยสุดป่ะ
                          For Each m_runPro In TempComPro.DetailsCom2                    ' เก็บเงินธรรมดา
                                If (Val(m_runPro.NUM_ONE) + 7) >= DueCount Then
                                        PercentSum = (m_runPro.SLM_PERCENT / 100)
                               End If
                           Next m_runPro
                     Else           ' ถ้าไม่ใช่สินค้าโปรโมต ให้ใช้ธรรมดา
                           
                        'เปลี่ยนไปใช้ m_ConditionCommiss2a แทน  ต้องเทียบว่า Groupไร , master ไหน เพื่อจะได้เลือก condition2a ถูก
                           PercentSum = 999
                           For num = 0 To n_array
                        Set temp_GoodsInGroup = GetObject("CGoodsGroup", m_GoodsInGroup(num), Stcrd.STKCOD)
                           If Not (temp_GoodsInGroup Is Nothing) Then                                              ' มีสินค้านี้อยู่ในกรุ๊ป
                                Set TempCConditionCommiss = m_ConditionCommiss2a(num).ITEM(1)
                                PercentSum = (TempCConditionCommiss.SLM_PERCENT / 100)
                                For Each m_runConditionCommiss In m_ConditionCommiss2a(num)                  ' เก็บเงินธรรมดา
                                     If (m_runConditionCommiss.NUM_ONE + 7) >= DueCount Then
                                          PercentSum = (m_runConditionCommiss.SLM_PERCENT / 100)
                                    End If
                                 Next m_runConditionCommiss
                            End If
                         Next num
                        If PercentSum = 999 Then                                                                                         ' ไม่มีสินค้านี้อยู่ในกรุ๊ป ให้เอา m_ConditionCommiss2a(num) ดีฟอล์
                                Set TempCConditionCommiss = m_ConditionCommiss2a(UBound(m_ConditionCommiss2a)).ITEM(1)
                                PercentSum = (TempCConditionCommiss.SLM_PERCENT / 100)
                                For Each m_runConditionCommiss In m_ConditionCommiss2a(UBound(m_ConditionCommiss2a))               ' เก็บเงินธรรมดา
                                     If (m_runConditionCommiss.NUM_ONE + 7) >= DueCount Then
                                          PercentSum = (m_runConditionCommiss.SLM_PERCENT / 100)
                                    End If
                                 Next m_runConditionCommiss
                          End If
                          
                        End If          'ใช้ธรรมดาไปแล้ว
               Else      ' งั้นก็เป็น สินค้าพิเศษ
               
                     PercentNum1 = (TempCConditionCommiss.SLM_PERCENT / 100)   'เป็นสินค้าพิเศษ
                     TempStr = "** "
                             
                     ' ถ้าเป็น เซล์และลูกค้าโปรโมต ใช้ comPro6
                     Set temp_ComMasPro = GetComMasPro(m_ComMasPro, Trim(Stcrd.SLMCOD & "-" & Stcrd.CUSCOD), False)
                      If Not (temp_ComMasPro Is Nothing) Then           ' หาโดยใช้  .SLMCOD .PEOPLE
                              For Each m_runComPro In m_ComPro3                   ' เก็บเงินธรรมดา
                                    If (Val(m_runComPro.CREDIT_NAME)) >= DueCount Then
                                         CreditKey = Val(m_runComPro.CREDIT_NAME)    'จะได้ เครดิตเป็น Key
                                   End If
                              Next m_runComPro
                                
                              ' วน loop ของ YEAR_ID นี้ ,, 02(สินค้าธรรมดา) อยู่ในเครดิตไหน  แล้วมันจะได้ค่า MASTER_COMMISS_SUB_PROMOTE_ID เป็น Key
                               Set TempComPro = GetComMasSubPro(m_ComPro3, Trim(Stcrd.SLMCOD & "-" & Stcrd.CUSCOD & "-" & CreditKey), False)
                               PercentSum = (TempComPro.DetailsCom2.ITEM(1).SLM_PERCENT / 100)  'น้อยสุดป่ะ
                               
                              For Each m_runPro In TempComPro.DetailsCom3                    ' เก็บเงินธรรมดา
                                    If (Val(m_runPro.NUM_ONE) + 7) >= DueCount Then
                                             PercentSum = (m_runPro.SLM_PERCENT / 100)
                                   End If
                              Next m_runPro
                         Else            ' ไม่ต้องคิดแบบโปรโมต
                              Set TempCConditionCommiss = m_ConditionCommiss3.ITEM(1)              ' เก็บเงินสินค้า Comodity
                              PercentSum = (TempCConditionCommiss.SLM_PERCENT / 100)
                              For Each m_runConditionCommiss In m_ConditionCommiss3
                                   If (m_runConditionCommiss.NUM_ONE + 7) >= DueCount Then
                                        PercentSum = (m_runConditionCommiss.SLM_PERCENT / 100)
                                  End If
                               Next m_runConditionCommiss
                   End If  ' สินค้าพิเศษ sale&ลูกค้าโปรโมตหรือเปล่า
               End If  ' สินค้าพิเศษหรือเปล่า

               j = j + 1
               Set HeadCf = m_Headers1.Fields(j)
               TempStr = TempStr & FormatNumber(NETVAL * PercentNum1)
               Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
               Call m_Details1.AddField(BodyCf)
               
            ' เหมือนคอลัมป์ก่อนหน้านี้ เด๊ะ
               j = j + 1
               Set HeadCf = m_Headers1.Fields(j)
               TempStr = FormatNumber(NETVAL * PercentNum1)
               Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
               Call m_Details1.AddField(BodyCf)
               
                j = j + 1
               Set HeadCf = m_Headers1.Fields(j)
               TempStr = DateToStringExtEx2(Stcrd.DOCDAT)                                                                                 ' วันที่ส่งสินค้า
               Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
               Call m_Details1.AddField(BodyCf)
      
                j = j + 1
               Set HeadCf = m_Headers1.Fields(j)
               If (mcolParam("FROM_CMPL_DATE") <= CMPLDAT) And (Stcrd.CMPLDAT <= toCMPLdate) And PayCompleted And CMPLDAT <> "29/12/2442" Then
                     TempStr = DateToStringExtEx2(CMPLDAT)                                                                                 ' วันที่รับชำระ
               Else:
                     TempStr = ""
               End If
               Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
               Call m_Details1.AddField(BodyCf)
               
                j = j + 1
               Set HeadCf = m_Headers1.Fields(j)
               If (mcolParam("FROM_CMPL_DATE") <= CMPLDAT) And (Stcrd.CMPLDAT <= toCMPLdate) And PayCompleted And CMPLDAT <> "29/12/2442" Then
                     TempStr = RECOD                                                                                         ' เลข RE
               Else:
                     TempStr = ""
               End If
               Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
               Call m_Details1.AddField(BodyCf)
      
                j = j + 1
               Set HeadCf = m_Headers1.Fields(j)
               If (mcolParam("FROM_CMPL_DATE") <= CMPLDAT) And (Stcrd.CMPLDAT <= toCMPLdate) And PayCompleted And CMPLDAT <> "29/12/2442" Then
                     TempStr = strDueCount                                                                                           ' เครดิต
               Else:
                     TempStr = ""
               End If
               Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
               Call m_Details1.AddField(BodyCf)
               
                  j = j + 1
               Set HeadCf = m_Headers1.Fields(j)
               TempStr = PercentSum                                                                                        ' %
               Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
               Call m_Details1.AddField(BodyCf)
   
   
               Amt = 0
                j = j + 1
               Set HeadCf = m_Headers1.Fields(j)
               If (mcolParam("FROM_CMPL_DATE") <= CMPLDAT) And (Stcrd.CMPLDAT <= toCMPLdate) And PayCompleted And CMPLDAT <> "29/12/2442" Then
                     Amt = (NETVAL * PercentNum1) * PercentSum
                     TempStr = FormatNumber(Amt)                                ' ค่าคอม
               Else:
                     TempStr = ""
               End If
               Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
               Call m_Details1.AddField(BodyCf)
               If haveData = False And ((NETVAL * PercentNum1) * PercentSum) <> 0 Then
                  haveData = True
               End If
               
               If Amt < 0 Then
                  CheckMinusFlag = True
               Else
                  CheckMinusFlag = False
               End If
               
               If mcolParam("SUMMARY_MODE") = "N" Then
                  If (mcolParam("SHOWZERO_MODE") = "Y") Or (mcolParam("SHOWZERO_MODE") = "N" And (((NETVAL * PercentNum1) * PercentSum) <> 0)) Then              ' ติ๊ก = โชว์ 0 ด้วย
                     If (mcolParam("ONLY_MINUS") = "Y" And CheckMinusFlag) Or (mcolParam("ONLY_MINUS") = "N") Then                           ' ติ๊ก = โชว์ เฉพาะ ลบ
                        Total4(5) = Total4(5) + NETVAL
                        Total3(5) = Total3(5) + NETVAL
                        Total2(5) = Total2(5) + NETVAL
                        Total1(5) = Total1(5) + NETVAL
                                   
                        Total4(6) = Total4(6) + (NETVAL * PercentNum1)
                        Total3(6) = Total3(6) + (NETVAL * PercentNum1)
                        Total2(6) = Total2(6) + (NETVAL * PercentNum1)
                        Total1(6) = Total1(6) + (NETVAL * PercentNum1)
                        
                        Total4(7) = Total4(7) + (NETVAL * PercentNum1)
                        Total3(7) = Total3(7) + (NETVAL * PercentNum1)
                        Total2(7) = Total2(7) + (NETVAL * PercentNum1)
                        Total1(7) = Total1(7) + (NETVAL * PercentNum1)
                        
                        Total4(13) = Total4(13) + Amt
                        Total3(13) = Total3(13) + Amt
                        Total2(13) = Total2(13) + Amt
                        Total1(13) = Total1(13) + Amt
                     
                        Call m_Details1.GetString(1, TempStr1, TempStr2)
                        strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
                        Call Vsp.AddTable(strFormat, "", TempStr2)
                     End If
                  End If
                  i = i + 1   'นับบรรทัดข้อมูลย่อย
               End If
      
                 PrevKey3 = tempCusArea.MASTER_AREA_ID  '!
                 Set temp_Area2 = GetAreaCom(m_AreaCod, tempCusArea.MASTER_AREA_ID)
                 PrevKey4 = temp_Area2.MASTER_AREA_NAME '!
                 PrevKey5 = Str(Month(Stcrd.DOCDAT)) & "/" & Str(Year(Stcrd.DOCDAT) + 543)
      End If   ' IVนี้ จ่ายREครบ
    End If     ' RE จาก IV ไม่เป็น Nothing เรียกวันที่ได้
 End If   ' ลูกค้าอยู่ในเขต
          
         If mcolParam("AREA_TYPE_ID") > 0 And Rs.AbsolutePosition = iCount Then    ' ถ้ามีการเลือกเขต ให้มีข้อมูลอยู่ในอเรย์นั้นๆ
              Exit For
         End If
                     
   Next Stcrd
Next temp_Area
   
      
            If mcolParam("SUMMARY_MODE") = "N" Then
               Call GenerateFooter("", "      รวม      ", PrevKey5, Total4, m_Details1, haveData_ByArea)       ' เช็ค haveValue ใน Call GenerateFooter **** เซลล์ต้องมีค่าถึงจะแสดงบรรทัดรวม
               Vsp.FontBold = True
            Else
               Call GenerateFooter("", "", PrevKey5, Total4, m_Details1, haveData_ByArea)        ' เช็ค haveValue ใน Call GenerateFooter **** เซลล์ต้องมีค่าถึงจะแสดงบรรทัดรวม
               Vsp.FontBold = False
            End If
            If haveData_ByArea = True Then
               K = K + 1
                  If mcolParam("SHOWZERO_MODE") = "Y" Then              ' ติ๊ก = โชว์ 0 ด้วย
                    Call m_Details1.GetString(2, TempStr1, TempStr2)
                    strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
                    Call Vsp.AddTable(strFormat, "", TempStr2)
                      For j = 1 To UBound(Total4)
                                  Total4(j) = 0
                    Next j
              Else
                     If Total4(13) <> 0 Then
                        Call m_Details1.GetString(2, TempStr1, TempStr2)
                        strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
                        Call Vsp.AddTable(strFormat, "", TempStr2)
                        For j = 1 To UBound(Total4)
                                  Total4(j) = 0
                        Next j
                     End If
                  End If
                  Vsp.FontBold = False
           End If
           

      
'       If mcolParam("SUMMARY_MODE") = "N" Then                                                         ' เขตพื้นที่
       
        Call GenerateFooter("", "      รวม      ", PrevKey4, Total2, m_Details1, haveData_ByArea)
         Vsp.FontBold = True
'      Else
'         Call GenerateFooter("", "รวม", PrevKey4, Total2, m_Details1, haveData_ByArea)
'         Vsp.FontBold = False
'      End If
      ' เช็ค haveValue ใน Call GenerateFooter **** บรรทัดสุดท้ายของเซลล์
       If haveData_ByArea = True Then
            K = K + 1
                 If mcolParam("SHOWZERO_MODE") = "Y" Then              ' ติ๊ก = โชว์ 0 ด้วย
                    Call m_Details1.GetString(2, TempStr1, TempStr2)
                    strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
                    Call Vsp.AddTable(strFormat, "", TempStr2)
                  Else
                     If Total2(13) <> 0 Then
                        Call m_Details1.GetString(2, TempStr1, TempStr2)
                        strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
                        Call Vsp.AddTable(strFormat, "", TempStr2)
                     End If
                  End If
            Vsp.FontBold = False
            
            If mcolParam("SAVE_MODE") = "Y" Then
              Call SaveToDB(PrevKey1, PrevKey2, "02", PrevKey3, 0, Total2, mcolParam("FROM_CMPL_DATE"), mcolParam("TO_CMPL_DATE"))             'PreKey3=เลขเขต ,Total2
            End If
       End If

      For j = 1 To UBound(Total2)
         Total2(j) = 0
      Next j

            TempStr = "      รวม      "
            Vsp.FontBold = True

      Call GenerateFooter("", TempStr, "  ** " & PrevKey2 & " **", Total3, m_Details1, haveValueSale)    ' เช็ค haveValue ใน Call GenerateFooter ****
      If haveValueSale = True Then                                                                     ' เซลล์ต้องมีค่า ถึงจะแสดงบรรทัดรวม   ****
            K = K + 1
                  If mcolParam("SHOWZERO_MODE") = "Y" Then              ' ติ๊ก = โชว์ 0 ด้วย
                    Call m_Details1.GetString(2, TempStr1, TempStr2)
                    strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
                    Call Vsp.AddTable(strFormat, "", TempStr2)
                  Else
                     If Total3(13) <> 0 Then
                        Call m_Details1.GetString(2, TempStr1, TempStr2)
                        strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
                        Call Vsp.AddTable(strFormat, "", TempStr2)
                     End If
                  End If
            Vsp.FontBold = False
      End If

           PrevKey3 = 0
           PrevKey4 = ""
           PrevKey5 = ""
           
            Set stcrd_mixdb = Nothing
            Set stcrd_mixdb = New Collection
           
   TempRs.MoveNext                                                            ' วนเซลล์
Wend
     
           Call GenerateFooter("สุทธิ", "", "", Total1, m_Details1, haveValueSale)
            Call m_Details1.GetString(2, TempStr1, TempStr2)
            strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
            Vsp.FontBold = True
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
   
                Set HeadCf = m_Headers0.Fields(1)
                TempStr = Txt1
                Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
                Call m_Details1.AddField(BodyCf)
    
                Set HeadCf = m_Headers0.Fields(2)
                TempStr = Txt2
                Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
                Call m_Details1.AddField(BodyCf)
    
    
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
         TempStr = Txt2
      ElseIf j = 3 Then
         TempStr = Txt3
     ElseIf j = 4 Then
         TempStr = ""
      ElseIf j = 8 Then
         TempStr = ""
      ElseIf j = 9 Then
         TempStr = ""
      ElseIf j = 10 Then
         TempStr = ""
      ElseIf j = 11 Then
         TempStr = ""
      ElseIf j = 12 Then
         TempStr = ""
      Else
         TempStr = FormatNumber(Tot(j))
      End If
      
      Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.Align, TempStr, HeadCf.BodyAlign)
      Call Details.AddField(BodyCf)
      
      If HaveValue = False And Tot(j) <> 0 Then
         HaveValue = True
      End If
   Next HeadCf
      
   Set BodyCf = Nothing
End Sub

Private Sub SaveToDB(Txt1 As String, Txt2 As String, COMTYP As String, Txt3 As Long, Txt4 As Long, Tot() As Double, FromDate As Date, ToDate As Date)
' Call SavetoDB( PreKey1, PreKey2 , "02", PreKey3 ,Total2 , mcolParam("FROM_CMPL_DATE"), mcolParam("TO_CMPL_DATE"))
Dim IsOK As Boolean
Dim exit_Record As CComRecord
Dim comRecord As CComRecord
Set comRecord = New CComRecord

Set exit_Record = GetObject("CCom_Record", m_ComRecord, Txt1 & "-" & Txt3 & "-" & Txt4, False)
If exit_Record Is Nothing Then
        
  If ConfirmSaveRecord(Txt2 & " " & Tot(13) & " บ.") Then
      comRecord.AddEditMode = SHOW_ADD
      comRecord.SLMCOD = Txt1
      comRecord.SLMNAME = Txt2
      comRecord.COMTYP = COMTYP
      comRecord.MASTER_AREA_ID = Txt3
      comRecord.GOODS_GROUP_ID = 1
      comRecord.COM_VALUE = Tot(13)
      comRecord.FROMDAT = FromDate
      comRecord.TODAT = ToDate
   
      If Not glbDaily.AddEditComRecord(comRecord, IsOK, True, glbErrorLog) Then
         glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
      End If
   End If
Else
     If ConfirmSaveExist(Txt2 & " " & Tot(13) & " บ.", exit_Record.SLMNAME & " " & exit_Record.COM_VALUE & " บ.") Then
         comRecord.AddEditMode = SHOW_EDIT
         comRecord.RECORD_ID = exit_Record.RECORD_ID
         comRecord.SLMCOD = Txt1
         comRecord.SLMNAME = Txt2
         comRecord.COMTYP = COMTYP
         comRecord.MASTER_AREA_ID = Txt3
         comRecord.GOODS_GROUP_ID = Txt4
         comRecord.COM_VALUE = Tot(13)
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


