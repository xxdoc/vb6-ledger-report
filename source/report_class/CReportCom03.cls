VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CReportCom03"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'User Group Report
Implements CReportInterface

Private Const MODULE_NAME = "CReportCOM03"
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
Private m_SumStcrd As Collection

Private m_ConditionCommiss5 As Collection
Private m_ConditionCommiss5_1 As Collection
Private m_ConditionCommiss5_2 As Collection
Public m_StkcodGroup As Collection

Private m_IncenPro5 As Collection
Private m_IncenPro5_1 As Collection
Private m_IncenPro5_2 As Collection
Public m_IncenProGroup As Collection

Private nonidentifyArea As Boolean
Dim Stb As CCommissMasterArea
Private m_MinusStkcod As Collection
Private m_ComRecord As Collection
Private m_NewCus As Collection
Private m_AllNewCus As Collection
Private m_IVincomplete0 As Collection
Private m_IVincomplete12 As Collection
Private m_ReDocdat As Collection
Private temp_Area As CCommissionCustomerArea
Private temp_Area2 As CCommissionCustomerArea
Private cr_condiCom As CConditionCommission
Dim L As Long
Dim NumCR As Long
Dim FirstDate As Date
Dim LastDate As Date
Dim CMPLFirstDate As Date
Dim CMPLLastDate As Date
Dim SixMonthFirst As Date
Dim SixMonthLast As Date
Dim dayFirst As Date
Dim dayLast As Date
Private m_REsumIV As Collection

Private m_IVcredit As Collection
Private tempIVcredit As CComIVcredit
Private stcrd_mixdb As Collection
Private FromSaleID As String
Private ToSaleID As String
Private Sub Class_Initialize()

   Set tempIVcredit = New CComIVcredit
   Set m_IVcredit = New Collection

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
   Set m_ConditionCommiss5 = New Collection
      Set m_SumStcrd = New Collection
      Set m_ConditionCommiss5_1 = New Collection
      Set m_ConditionCommiss5_2 = New Collection
      Set m_StkcodGroup = New Collection
      
      Set m_IncenPro5 = New Collection
      Set m_IncenPro5_1 = New Collection
      Set m_IncenPro5_2 = New Collection
      Set m_IncenProGroup = New Collection
      
     Set Stb = New CCommissMasterArea
     Set m_MinusStkcod = New Collection
      Set m_AllNewCus = New Collection
     Set m_NewCus = New Collection
     Set m_IVincomplete0 = New Collection
     Set m_IVincomplete12 = New Collection
     Set m_ReDocdat = New Collection
    Set temp_Area = New CCommissionCustomerArea
    Set temp_Area2 = New CCommissionCustomerArea
    Set cr_condiCom = New CConditionCommission
    Set m_REsumIV = New Collection
    Set stcrd_mixdb = New Collection
    
    Set m_ComRecord = New Collection
End Sub

Private Sub Class_Terminate()
   Call ClearParam
   
    Set m_IVcredit = Nothing
    Set tempIVcredit = Nothing
   
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
   Set m_ConditionCommiss5 = Nothing
      Set m_SumStcrd = Nothing
   Set m_ConditionCommiss5_1 = Nothing
   Set m_ConditionCommiss5_2 = Nothing
   Set m_StkcodGroup = Nothing
   
      Set m_IncenPro5 = Nothing
      Set m_IncenPro5_1 = Nothing
      Set m_IncenPro5_2 = Nothing
      Set m_IncenProGroup = Nothing
      
    Set m_AllNewCus = Nothing
  Set m_NewCus = Nothing
  Set Stb = Nothing
  Set m_MinusStkcod = Nothing
  Set m_IVincomplete0 = Nothing
  Set m_IVincomplete12 = Nothing
  Set m_ReDocdat = Nothing
  Set temp_Area = Nothing
  Set temp_Area2 = Nothing
  Set m_REsumIV = Nothing
  Set cr_condiCom = Nothing
  Set stcrd_mixdb = Nothing
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
      Call Cf.SetFieldValue(32, "^", "", "<")    '27
      Call m_Headers0.AddField(Cf)
      
If mcolParam("SUMMARY_MODE") = "N" Then
      Call Cf.SetFieldValue(1, "^", "NO.", "^")
      Call m_Headers1.AddField(Cf)

   Call Cf.SetFieldValue(2, "^", "INVOICE", "<")
   Call m_Headers1.AddField(Cf)
   
      Call Cf.SetFieldValue(5, "^", "ชื่อลูกค้า", "<")
   Call m_Headers1.AddField(Cf)
   
      Call Cf.SetFieldValue(5, "^", "ชื่อสินค้า", "<")
   Call m_Headers1.AddField(Cf)
Else
      Call Cf.SetFieldValue(0, "^", "NO.", "^")
      Call m_Headers1.AddField(Cf)
      
      Call Cf.SetFieldValue(2, "^", "รหัส", "<")
      Call m_Headers1.AddField(Cf)
      
    Call Cf.SetFieldValue(6, "^", "พนักงานขาย", "<")
   Call m_Headers1.AddField(Cf)
   
      Call Cf.SetFieldValue(0, "^", "ชื่อสินค้า", "<")
   Call m_Headers1.AddField(Cf)
End If

   Call Cf.SetFieldValue(2, "^", "จำนวน  (ขวด,ถุง)", "^")
   Call m_Headers1.AddField(Cf)
   
   Call Cf.SetFieldValue(2, "^", "ราคา   ต่อขวด", ">")
   Call m_Headers1.AddField(Cf)
      
   Call Cf.SetFieldValue(2, "^", "ยอดขาย", ">")
   Call m_Headers1.AddField(Cf)
        
 If mcolParam("SUMMARY_MODE") = "N" Then
   Call Cf.SetFieldValue(2, "^", "วันส่งสินค้า", "^")
   Call m_Headers1.AddField(Cf)
      
   Call Cf.SetFieldValue(2, "^", "วันรับชำระ", "^")
    Call m_Headers1.AddField(Cf)
      
   Call Cf.SetFieldValue(2, "^", "เลขที่   ใบเสร็จ", "^")
    Call m_Headers1.AddField(Cf)

   Call Cf.SetFieldValue(1, "^", "Cr.", "^")
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
 End If
      Call Cf.SetFieldValue(2, "^", "ลูกค้าใหม่", ">")
      Call m_Headers1.AddField(Cf)
 
      Call Cf.SetFieldValue(3, "^", "ค่า Incentive ต่อขวดและถุง", ">")
      Call m_Headers1.AddField(Cf)
 
      Call Cf.SetFieldValue(2, "^", "รวม", ">")
      Call m_Headers1.AddField(Cf)
      
      Call Cf.SetFieldValue(2, "^", "Total", ">")
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
Dim YEAR_ID As Long

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
   
   'เงื่อนไขพิเศษ
   Call LoadIncenPro05(m_IncenPro5, mcolParam("FROM_CMPL_DATE"), mcolParam("TO_CMPL_DATE"))  ' เงื่อนไข incentive
   Call LoadIncenPro05_2(m_IncenPro5_2, mcolParam("FROM_CMPL_DATE"), mcolParam("TO_CMPL_DATE"))  ' คัดมาเฉพาะสินค้า Incentive
   Call LoadGroupIncenPro(m_IncenProGroup, mcolParam("FROM_CMPL_DATE"), mcolParam("TO_CMPL_DATE"))
   
   'เงื่อนไขธรรมดา
   Call LoadIVcredit(m_IVcredit, mcolParam("FROM_DOC_DATE"), mcolParam("TO_DOC_DATE"))
   Call LoadCommission05(m_ConditionCommiss5, mcolParam("FROM_CMPL_DATE"), mcolParam("TO_CMPL_DATE"), 1) ' เงื่อนไข incentive
   Set cr_condiCom = m_ConditionCommiss5.ITEM(1)
   NumCR = cr_condiCom.INCEN_CR
   Call LoadCommission05_2(m_ConditionCommiss5_2, mcolParam("FROM_CMPL_DATE"), mcolParam("TO_CMPL_DATE"), 1) ' คัดมาเฉพาะสินค้า Incentive
   Call LoadGroupIncentive(m_StkcodGroup, mcolParam("FROM_CMPL_DATE"), mcolParam("TO_CMPL_DATE"), 1)
   
   Dim bak_ID As Long
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
 
    Call LoadMinusIV(m_MinusStkcod, mcolParam("FROM_DOC_DATE"), mcolParam("TO_DOC_DATE"))
    Call LoadREDocDat(m_ReDocdat, mcolParam("FROM_CMPL_DATE"), mcolParam("TO_CMPL_DATE")) ' ต้องเอามาหมด ของปีปัจจุบัน
'    Call LoadREDocDat(m_ReDocdat, mcolParam("FROM_CMPL_DATE"), mcolParam("TO_CMPL_DATE"), True) ' ต้องเอามาหมด ของปี db2
    Call LoadREsumIV(Nothing, m_REsumIV)
    'Call LoadREsumIV(Nothing, m_REsumIV, , True)

   Call LoadYearId(YEAR_ID, mcolParam("TO_CMPL_DATE"), mcolParam("TO_CMPL_DATE"))   ' เวลานี้ใช้ Area_Year ไหน
   If mcolParam("AREA_TYPE_ID") > 0 Then
      Call LoadAreaComReport(Nothing, m_AreaCod, YEAR_ID, Val(Left(mcolParam("AREA_TYPE_NAME"), 2)))
   Else
      Call LoadAreaComReport(Nothing, m_AreaCod, YEAR_ID)       ' โหลดไว้สำหรับ combo
   End If
   
   ' Save_mode
   If mcolParam("SAVE_MODE") = "Y" Then
       Call LoadComRecord(m_ComRecord, "03", mcolParam("TO_CMPL_DATE"))
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
Dim M As Long
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
Dim PrevKey5 As String    ' เดือน
'Dim PrevKey6 As String    ' ลูกค้า
Dim St As CCommissMasterArea
Dim ArS As COESLM
Dim Stcrd As CStcrd
Dim temp_Stcrd As CStcrd
Dim SumStkcod As CStcrd

Dim haveValueSale As Boolean
Dim haveData As Boolean
Dim haveData_ByArea As Boolean
Dim DueCount As Long
Dim EnableIncentive As Boolean
Dim EnableNumTwo As Boolean
Dim NUM_ONE As Double 'String
Dim NUM_TWO As Double
Dim FlagNewCus As Boolean
Dim SumTRNQTY As Double

Dim TempCConditionCommiss As CConditionCommission
Dim m_runConditionCommiss As CConditionCommission
'Dim CR_CondiCom As CConditionCommission
Dim m_runConditionCommiss2 As CConditionCommission
Dim PercentNum1 As Double
Dim PercentSum As Double
Dim NEWCUSVAL As Double
Dim tempCusArea As CCommissionCustomerArea
Dim toCMPLdate As Date
Dim tempMinusStkcod As CComMinusStk
Dim tempREdoc As CARTrn
Dim NETVAL As Double
Dim PayCompleted  As Boolean
Dim CMPLDAT As Date
Dim RECOD As String
Dim REsumIV As CARRcIt
Dim strDueCount As String

Dim m_runIncenPro2 As CIncentivePromote
Dim TempIncenPro As CIncentivePromote
Dim m_runIncenPro As CIncentivePromote
Dim stcrd_temp As CStcrd
   
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
      
   If mcolParam("TO_CMPL_DATE") = "27/12/2442" Then
         toCMPLdate = DateSerial(9999, 12, 31)
   Else:
          toCMPLdate = mcolParam("TO_CMPL_DATE")
   End If
      
   Set ArS = New COESLM
   ArS.FROM_SALE_CODE = FromSaleID
   ArS.TO_SALE_CODE = ToSaleID
   Call glbDaily.QuerySale(ArS, TempRs, iCount, IsOK, glbErrorLog)
   
    M = 0  'เดือนใหม่ ไม่อิงกะใคร
   While Not TempRs.EOF          ' sale
   
         Call m_Details1.ClearField
         Call ArS.PopulateFromRS(1, TempRs)
         
         For j = 1 To UBound(Total3)
            Total3(j) = 0
         Next j

         Call GenerateHeader("( " & ArS.SLMCOD & " )", ArS.SLMNAM, m_Details1)
          If mcolParam("SUMMARY_MODE") = "N" Then
             Call m_Details1.GetString(1, TempStr1, TempStr2)
             strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
             Vsp.FontBold = True
             Call Vsp.AddTable(strFormat, "", TempStr2)
             Vsp.FontBold = False
         End If
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
  

  '  **** วนทุกบรรทัดของ Stcrd
  Set tempCusArea = GetCusAreaCom(temp_Area.ImportExportItems, Stcrd.CUSCOD, False)     ' !!!!!         ดึงคอลเล็คชั่นลูกที่อยู่ใน คอลเล็คชั่นแม่มาเก็บไว้
  If (Not (tempCusArea Is Nothing)) Then                 ' อยู่ในเขตที่ผู้ใช้เลือก หรือ ผู้ใช้ไม่ได้เลือกเลย Or nonidentifyArea

   ' คำนวณส่วนลดก่อน
   Set tempMinusStkcod = GetMinusCommiss(m_MinusStkcod, Stcrd.DOCDAT & "-" & Stcrd.DOCNUM & "-" & Stcrd.STKCOD, False)
   If (tempMinusStkcod Is Nothing) Then   'ไม่เจอในคอเล็คชั่น = ราคาเท่าเดิม
      NETVAL = Stcrd.NETVAL
   Else:
      NETVAL = (Stcrd.NETVAL + Val(tempMinusStkcod.MINUS_AMOUNT))
   End If
 
' If Stcrd.DOCNUM = "IV0044485" Then
'      'debug.print
' End If

         
'         If Stcrd.STKCOD = "V200-1" Then
'            'debug.print
'         End If
 
  Set REsumIV = GetARRcpItem(m_REsumIV, Stcrd.DOCNUM, False)
  If Not (REsumIV Is Nothing) Then     ' เลือกมาเฉพาะ RE ที่จ่าย IV
     RECOD = REsumIV.RCPNUM
     Set tempREdoc = GetREDocDat(m_ReDocdat, Trim(Stcrd.DOCNUM), False)
     If Not (tempREdoc Is Nothing) Then
        CMPLDAT = tempREdoc.DOCDAT                                                                         ' เด๋วต้องแก้ตรงนี้ วันที่รับชำระ
     Else
         CMPLDAT = -1
     End If
          
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
                 
     ' เลือกมาเฉพาะที่จ่ายครบแล้ว RE.Key(IV)-NETVAL <=0 then
     If ((REsumIV.RCVAMT - NETVAL) >= 0) And CMPLDAT >= mcolParam("FROM_CMPL_DATE") And CMPLDAT <= mcolParam("TO_CMPL_DATE") Then
         SumTRNQTY = 0
         EnableIncentive = False
          '------------------------------------------
         Set m_runIncenPro2 = GetCheckIncenPro(m_IncenPro5_2, Stcrd.STKCOD, False)                 ' เงื่อนไขพิเศษ
         If (Not (m_runIncenPro2 Is Nothing)) Then        ' ต้องเป็นสินค้า Incentive ที่มีในเงื่อนไข
              Call LoadIncenPro05_1(m_IncenPro5_1, mcolParam("FROM_CMPL_DATE"), mcolParam("TO_CMPL_DATE"), Stcrd.STKCOD)   ' เงื่อนไข incentive
              '    โค๊ด if sumstkcod.TRNQTY  >= TempCConditionCommiss.NUM_ONE then ให้วนไป NUM_ONE ถัดไปของสินค้าตัวนั้น  ---- ต้องมี collection เก็บสินค้าใครมันหรือเปล่า
              '   จนกระทั้ง sumstkcod.TRNQTY <TempCConditionCommiss.NUM_ONE ก็ไม่ต้องเอาค่าใหม่ใส่ทับ และใช้ค่านั้นเป็น NUM_ONE
              NUM_ONE = 0                               ' น้อยสุด
              FlagNewCus = True
              PercentSum = 0
              NUM_TWO = 0
                          
           If DueCount <= NumCR Then    'EnableIncentive = True And
              For Each m_runIncenPro In m_IncenPro5_1
                   If Val(NETVAL / Stcrd.TRNQTY) >= Val(m_runIncenPro.NUM_ONE) Then
                       '   'debug.print Stcrd.DOCNUM & "--" & Stcrd.STKDES & "-" & SumTRNQTY & "-" & m_runIncenPro.NUM_ONE
                       NUM_ONE = m_runIncenPro.NUM_ONE    ' ใช้ m_runConditionCommiss.NUM_ONE เป็นคีย์ใน step3
                       EnableIncentive = True
                    End If
               Next m_runIncenPro
                                       
               Set TempIncenPro = GetCheckIncenPro(m_IncenPro5, Trim(Stcrd.STKCOD & "-" & NUM_ONE), False)     '         Call Cl.Add(TempData, Trim(TempData.STKCOD & "-" & TempData.NUM_ONE & "-" & TempData.NUM_TWO))    ' KEy
               If Not (TempIncenPro Is Nothing) Then
                   ' 'debug.print Stcrd.DOCNUM
                   PercentSum = Val(TempIncenPro.SLM_PERCENT)
                   If TempIncenPro.OPERATOR = "Y" Then
                        FlagNewCus = False
                    End If
                End If
             End If  ' ถ้าระยะเวลาเก็บเงิน <= 97
    '------------------------------------------
    ElseIf m_runIncenPro2 Is Nothing Then
       Set m_runConditionCommiss2 = GetCheckCommiss(m_ConditionCommiss5_2, Stcrd.STKCOD, False)                 ' เงื่อนไขปกติ
       If (Not (m_runConditionCommiss2 Is Nothing)) Then        ' ต้องเป็นสินค้า Incentive ที่มีในเงื่อนไข
            ' ถ้ามันมีคู่  Key เป็นสินค้าตัวนี้ สมาชิกใน Group >1 ให้ GetSum สินค้าคู่หูมา
           Set TempCConditionCommiss = GetCheckCommiss(m_StkcodGroup, Trim(Stcrd.STKCOD), False)     '         Call Cl.Add(TempData, Trim(TempData.STKCOD & "-" & TempData.NUM_ONE & "-" & TempData.NUM_TWO))    ' KEy
           If Not (TempCConditionCommiss Is Nothing) Then
                  For Each m_runConditionCommiss In m_StkcodGroup
                      If m_runConditionCommiss.GROUP1 = TempCConditionCommiss.GROUP1 Then
                           Call GetFirstLastDate(DateSerial(Year(Stcrd.DOCDAT), Month(Stcrd.DOCDAT), 1), FirstDate, LastDate)
'                           If Year(Stcrd.DOCDAT) < 2012 Then
'                              'debug.print
'                           End If
                           Call GetFirstLastDate(DateSerial(Year(CMPLDAT), Month(CMPLDAT), 1), CMPLFirstDate, CMPLLastDate)
                           Call LoadSumStcrdMonth(Nothing, m_SumStcrd, FirstDate, LastDate, CMPLFirstDate, CMPLLastDate)           ' โหลด sum TRNQTY groupby เดือนDOCDAT, สินค้า STKCOD     โดยดูผลรวมในเดือนนั้นๆ
                           Set SumStkcod = GetStkcodNoNew(m_SumStcrd, m_runConditionCommiss.STKCOD)
                           If Not (SumStkcod Is Nothing) Then
                               SumTRNQTY = SumTRNQTY + SumStkcod.TRNQTY
                           End If
                       End If
                    Next m_runConditionCommiss
            End If                ' หาคู่
                      
'                      If Stcrd.DOCNUM = "IV0043362" Then
'                        'debug.print
'                      End If
'
           Call LoadCommission05_1(m_ConditionCommiss5_1, mcolParam("FROM_CMPL_DATE"), mcolParam("TO_CMPL_DATE"), Stcrd.STKCOD)   ' เงื่อนไข incentive
         
'         If Stcrd.STKCOD = "V200-1" Then
'            'debug.print
'         End If
'
'         โค๊ด if sumstkcod.TRNQTY  >= TempCConditionCommiss.NUM_ONE then ให้วนไป NUM_ONE ถัดไปของสินค้าตัวนั้น  ---- ต้องมี collection เก็บสินค้าใครมันหรือเปล่า
         '   จนกระทั้ง sumstkcod.TRNQTY <TempCConditionCommiss.NUM_ONE ก็ไม่ต้องเอาค่าใหม่ใส่ทับ และใช้ค่านั้นเป็น NUM_ONE
            NUM_ONE = 0                               ' น้อยสุด
            For Each m_runConditionCommiss In m_ConditionCommiss5_1
                If Val(SumTRNQTY) >= Val(m_runConditionCommiss.NUM_ONE) Then
                   '   'debug.print Stcrd.DOCNUM & "--" & Stcrd.STKDES & "-" & SumTRNQTY & "-" & m_runConditionCommiss.NUM_ONE
                     NUM_ONE = m_runConditionCommiss.NUM_ONE    ' ใช้ m_runConditionCommiss.NUM_ONE เป็นคีย์ใน step3
                     EnableIncentive = True
               End If
             Next m_runConditionCommiss

            '   ดึงว่าสินค้าตัวที่รัน มีสิทธิ์มีสิทธิคิด ลูกค้าใหม่ ป่าว
            FlagNewCus = True
            PercentSum = 0
            NUM_TWO = 0
            If DueCount <= NumCR Then     'EnableIncentive = True And
                  For Each m_runConditionCommiss In m_ConditionCommiss5_1
                        If m_runConditionCommiss.NUM_ONE = NUM_ONE Then    ' ต้องดัก ไม่งั้นจะมาหมด ทั้งๆที่ NUM_ONE ไม่ใช่คู่ของมัน
                               If (NETVAL / Stcrd.TRNQTY) >= (m_runConditionCommiss.NUM_TWO) Then
                                  NUM_TWO = m_runConditionCommiss.NUM_TWO      ' ใช้ m_runConditionCommiss.NUM_TWO เป็นคีย์ใน step3
                                ' EnableNumTwo = True
                              End If
                          End If
                    Next m_runConditionCommiss

                  ' อ้างค่าเพื่อดูว่าใช้ค่า incentive เกณฑ์ไหน
                  Set TempCConditionCommiss = GetCheckCommiss(m_ConditionCommiss5, Trim(Stcrd.STKCOD & "-" & NUM_ONE & "-" & NUM_TWO), False)     '         Call Cl.Add(TempData, Trim(TempData.STKCOD & "-" & TempData.NUM_ONE & "-" & TempData.NUM_TWO))    ' KEy
                   If Not (TempCConditionCommiss Is Nothing) Then
                             PercentSum = Val(TempCConditionCommiss.SLM_PERCENT)
                              If TempCConditionCommiss.OPERATOR = "Y" Then
                                       FlagNewCus = False
                              End If
                   End If
             End If      ' อยู่ภายใน 97 วัน
                           
      End If ' อยู่ในเงื่อนไขธรรมดา
End If  ' อยู่ในเงื่อนไขพิเศษรึป่าว
'======== เลือกประเภทการคำนวน Pro หรือ ธรรมดาแล้ว ======'
      If EnableIncentive = True Then     ' ยอดเกิน เขียนบรรทัดนี้ได้
            If PrevKey5 <> Str(Month(Stcrd.DOCDAT)) & "/" & Str(Year(Stcrd.DOCDAT) + 543) And PrevKey5 <> "" Then
                  M = M + 1
                  If mcolParam("SUMMARY_MODE") = "N" Then
                     Call GenerateFooter("", "รวม", PrevKey5, Total4, m_Details1, haveData_ByArea)       ' เช็ค haveValue ใน Call GenerateFooter **** เซลล์ต้องมีค่าถึงจะแสดงบรรทัดรวม
                     Vsp.FontBold = True
                  Else
                     Call GenerateFooter("", Str(M), PrevKey5, Total4, m_Details1, haveData_ByArea)        ' เช็ค haveValue ใน Call GenerateFooter **** เซลล์ต้องมีค่าถึงจะแสดงบรรทัดรวม
                     Vsp.FontBold = False
                  End If

                 If haveData_ByArea = True Then
                      M = M + 1
                    If mcolParam("SHOWZERO_MODE") = "Y" Then   ' ติ๊ก = โชว์ 0 ด้วย
                        Call m_Details1.GetString(2, TempStr1, TempStr2)
                        strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
                        Call Vsp.AddTable(strFormat, "", TempStr2)
                     Else
                        If Total4(15) <> 0 Then
                          Call m_Details1.GetString(2, TempStr1, TempStr2)
                          strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
                          Call Vsp.AddTable(strFormat, "", TempStr2)
                        End If
                     End If
                End If
                     Vsp.FontBold = False
                  
                     For j = 1 To UBound(Total4)
                         Total4(j) = 0
                     Next j
            End If
         
               If tempCusArea.MASTER_AREA_ID <> PrevKey3 Then
                   If mcolParam("SUMMARY_MODE") = "N" Then
                     Call GenerateFooter("", "รวม", PrevKey4, Total2, m_Details1, haveData_ByArea)        ' เช็ค haveValue ใน Call GenerateFooter **** เซลล์ต้องมีค่าถึงจะแสดงบรรทัดรวม
                     Vsp.FontBold = True
                  Else
                     Call GenerateFooter("", "", PrevKey4, Total2, m_Details1, haveData_ByArea)        ' เช็ค haveValue ใน Call GenerateFooter **** เซลล์ต้องมีค่าถึงจะแสดงบรรทัดรวม
                     Vsp.FontBold = False
                  End If
                  
                   If haveData_ByArea = True Then
                     K = K + 1
                    If mcolParam("SHOWZERO_MODE") = "Y" Then   ' ติ๊ก = โชว์ 0 ด้วย
                        Call m_Details1.GetString(2, TempStr1, TempStr2)
                        strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
                        Call Vsp.AddTable(strFormat, "", TempStr2)
                     Else
                          If Total2(15) <> 0 Then
                          Call m_Details1.GetString(2, TempStr1, TempStr2)
                          strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
                          Call Vsp.AddTable(strFormat, "", TempStr2)
                          End If
                     End If
                     
                     If mcolParam("SAVE_MODE") = "Y" Then
                          Call SaveToDB(PrevKey1, PrevKey2, "03", PrevKey3, 0, Total2, mcolParam("FROM_CMPL_DATE"), mcolParam("TO_CMPL_DATE"))             'PreKey3=เลขเขต ,Total2
                     End If
                     
                  End If
                  
                  
                        Vsp.FontBold = False
                     
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
             End If
'
'             If Stcrd.DOCNUM = "IV0042361" Then
'               'debug.print
'             End If

               Call m_Details1.ClearField
               j = 0
               haveData = False
                
'                If i = 10 Then
'                  'debug.print
'                End If
                
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
               TempStr = Left(Stcrd.CUSNAM, 27)
               Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
               Call m_Details1.AddField(BodyCf)
      
               j = j + 1
               Set HeadCf = m_Headers1.Fields(j)
               TempStr = Stcrd.STKDES
               Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
               Call m_Details1.AddField(BodyCf)
               
              j = j + 1
               Set HeadCf = m_Headers1.Fields(j)
               TempStr = Stcrd.TRNQTY                                           ' จำนวน
               Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
               Call m_Details1.AddField(BodyCf)
                     Total4(j) = Total4(j) + (Val(Stcrd.TRNQTY))
                     Total3(j) = Total3(j) + (Val(Stcrd.TRNQTY))
                     Total2(j) = Total2(j) + (Val(Stcrd.TRNQTY))
                     Total1(j) = Total1(j) + (Val(Stcrd.TRNQTY))
               If haveData = False And Val(Stcrd.TRNQTY) <> 0 Then
                  haveData = True
               End If

              j = j + 1
               Set HeadCf = m_Headers1.Fields(j)
               TempStr = FormatNumber(NETVAL / Stcrd.TRNQTY)                                           ' ราคาต่อขวด
               Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
               Call m_Details1.AddField(BodyCf)
                     Total4(j) = Total4(j) + (Val(Stcrd.UNITPR))
                     Total3(j) = Total3(j) + (Val(Stcrd.UNITPR))
                     Total2(j) = Total2(j) + (Val(Stcrd.UNITPR))
                     Total1(j) = Total1(j) + (Val(Stcrd.UNITPR))
               If haveData = False And Val(Stcrd.UNITPR) <> 0 Then
                  haveData = True
               End If
               
              PayCompleted = True

               j = j + 1
               Set HeadCf = m_Headers1.Fields(j)
'               'debug.print Stcrd.DOCNUM
               TempStr = FormatNumber(NETVAL)                                   ' เงินที่จ่ายจริง
               Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
               Call m_Details1.AddField(BodyCf)
                     Total4(j) = Total4(j) + (NETVAL)
                     Total3(j) = Total3(j) + (NETVAL)
                     Total2(j) = Total2(j) + (NETVAL)
                     Total1(j) = Total1(j) + (NETVAL)
               If haveData = False And NETVAL <> 0 Then
                  haveData = True
               End If

                j = j + 1
               Set HeadCf = m_Headers1.Fields(j)
               TempStr = DateToStringExtEx2(Stcrd.DOCDAT)                                                                                 ' วันที่ส่งสินค้า
               Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
               Call m_Details1.AddField(BodyCf)
      
                j = j + 1
               Set HeadCf = m_Headers1.Fields(j)
               If (mcolParam("FROM_CMPL_DATE") <= CMPLDAT) And (CMPLDAT <= toCMPLdate) And PayCompleted And CMPLDAT > 0 Then
                      TempStr = DateToStringExtEx2(CMPLDAT)                                                                                 ' วันที่รับชำระ
               Else:
                     TempStr = ""
               End If
               Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
               Call m_Details1.AddField(BodyCf)
               
                j = j + 1
               Set HeadCf = m_Headers1.Fields(j)
               If (mcolParam("FROM_CMPL_DATE") <= CMPLDAT) And (CMPLDAT <= toCMPLdate) And PayCompleted And CMPLDAT > 0 Then
                     TempStr = RECOD                                                                                         ' เลข RE
               Else:
                     TempStr = ""
               End If
               Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
               Call m_Details1.AddField(BodyCf)
      
                j = j + 1
               Set HeadCf = m_Headers1.Fields(j)
               If (mcolParam("FROM_CMPL_DATE") <= CMPLDAT) And (CMPLDAT <= toCMPLdate) And PayCompleted And CMPLDAT > 0 Then
                     TempStr = strDueCount                                                                                         ' เครดิต
               Else:
                     TempStr = ""
               End If
               Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
               Call m_Details1.AddField(BodyCf)
                            
            ' ย้อนหลังประวัติลูกค้าคนนี้ 6เดือน
            NEWCUSVAL = 0                 ' กำหนดไว้ก่อนเลย ถ้าไม่ใช่ค่อยเปลี่ยน
            TempStr = ""
            
            If FlagNewCus = True And NETVAL >= 5000 And DueCount <= NumCR Then
               If Year(SixMonthFirst) = Year(SixMonthLast) Then
                   Call LoadNewCus(m_NewCus, SixMonthFirst, SixMonthLast, Stcrd.CUSCOD)    ' ต้องดูย้อนหลังว่าลูกค้าคนนี้ย้อนไป 6 เดือนที่แล้ว มีใบสั่งซื้อไหม
               Else
                  Call GetFirstLastDate(DateSerial(Year(SixMonthFirst), Month(1), 31), , dayLast)
                  Call LoadNewCus(m_NewCus, SixMonthFirst, dayLast, Stcrd.CUSCOD)          ' ปีเก่า เดือนนั้นๆ ถึงสิ้นปี
                  If m_NewCus.Count = 0 Then
                      Call GetFirstLastDate(DateSerial(Year(SixMonthLast), Month(12), 1), , dayFirst)
                      Call LoadNewCus(m_NewCus, dayFirst, SixMonthLast, Stcrd.CUSCOD)              ' ปีใหม่ ต้นปี ถึง เดือนนั้นๆ
                  End If
               End If
               Set temp_Stcrd = GetStkcodNoNew(m_AllNewCus, Stcrd.CUSCOD)
               If m_NewCus.Count = 0 And temp_Stcrd Is Nothing Then     '    And Stcrd.CUSCOD <> PrevKey6
                     NEWCUSVAL = 300
                     TempStr = FormatNumber(NEWCUSVAL)       '  แล้วเก็บไว้ใน Colection ว่าของคนนี้ให้แล้ว
                     Set temp_Stcrd = New CStcrd
                     temp_Stcrd.CUSCOD = Stcrd.CUSCOD
                     Call m_AllNewCus.Add(temp_Stcrd, Trim(Stcrd.CUSCOD))
                     Set temp_Stcrd = Nothing
               End If
         End If

               j = j + 1
              Set HeadCf = m_Headers1.Fields(j)
               If (mcolParam("FROM_CMPL_DATE") <= CMPLDAT) And (CMPLDAT <= toCMPLdate) And PayCompleted And CMPLDAT > 0 Then
                     Total4(j) = Total4(j) + NEWCUSVAL                              ' ลูกค้าใหม่
                     Total3(j) = Total3(j) + NEWCUSVAL
                     Total2(j) = Total2(j) + NEWCUSVAL
                     Total1(j) = Total1(j) + NEWCUSVAL
               Else:
                     TempStr = ""
               End If
               Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
               Call m_Details1.AddField(BodyCf)
               If haveData = False And (NEWCUSVAL) <> 0 Then
                  haveData = True
               End If
               
                j = j + 1
               Set HeadCf = m_Headers1.Fields(j)
               If (mcolParam("FROM_CMPL_DATE") <= CMPLDAT) And (CMPLDAT <= toCMPLdate) And PayCompleted And CMPLDAT > 0 Then
                     TempStr = FormatNumber(PercentSum)                                 ' ราคาต่อหน่วย
                     Total4(j) = Total4(j) + (PercentSum)
                     Total3(j) = Total3(j) + (PercentSum)
                     Total2(j) = Total2(j) + (PercentSum)
                     Total1(j) = Total1(j) + (PercentSum)
               Else:
                     TempStr = ""
               End If
               Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
               Call m_Details1.AddField(BodyCf)
               If haveData = False And (PercentSum) <> 0 Then
                  haveData = True
               End If

               j = j + 1
               Set HeadCf = m_Headers1.Fields(j)
               If (mcolParam("FROM_CMPL_DATE") <= CMPLDAT) And (CMPLDAT <= toCMPLdate) And PayCompleted And CMPLDAT > 0 Then
                     TempStr = FormatNumber(Val(Stcrd.TRNQTY) * PercentSum)                                 ' รวม ค่าคอม
                     Total4(j) = Total4(j) + (Val(Stcrd.TRNQTY) * PercentSum)
                     Total3(j) = Total3(j) + (Val(Stcrd.TRNQTY) * PercentSum)
                     Total2(j) = Total2(j) + (Val(Stcrd.TRNQTY) * PercentSum)
                     Total1(j) = Total1(j) + (Val(Stcrd.TRNQTY) * PercentSum)
               Else:
                     TempStr = ""
               End If
               Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
               Call m_Details1.AddField(BodyCf)
               If haveData = False And (Val(Stcrd.TRNQTY) * PercentSum) <> 0 Then
                  haveData = True
               End If
               
                j = j + 1
               Set HeadCf = m_Headers1.Fields(j)
               If (mcolParam("FROM_CMPL_DATE") <= CMPLDAT) And (CMPLDAT <= toCMPLdate) And PayCompleted And CMPLDAT > 0 Then
                     TempStr = FormatNumber(NEWCUSVAL + (Val(Stcrd.TRNQTY) * PercentSum))                                   ' total
                     Total4(j) = Total4(j) + NEWCUSVAL + (Val(Stcrd.TRNQTY) * PercentSum)
                     Total3(j) = Total3(j) + NEWCUSVAL + (Val(Stcrd.TRNQTY) * PercentSum)
                     Total2(j) = Total2(j) + NEWCUSVAL + (Val(Stcrd.TRNQTY) * PercentSum)
                     Total1(j) = Total1(j) + NEWCUSVAL + (Val(Stcrd.TRNQTY) * PercentSum)
               Else:
                     TempStr = ""
               End If
               Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
               Call m_Details1.AddField(BodyCf)
               If haveData = False And (Val(Stcrd.TRNQTY) * PercentSum) <> 0 Then
                  haveData = True
               End If
              
               If mcolParam("SUMMARY_MODE") = "N" And haveData = True Then
                 If mcolParam("SHOWZERO_MODE") = "Y" Then              ' ติ๊ก = โชว์ 0 ด้วย
                    Call m_Details1.GetString(1, TempStr1, TempStr2)
                    strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
                    Call Vsp.AddTable(strFormat, "", TempStr2)
                    i = i + 1
                  Else
                     If Val(NEWCUSVAL) + (Val(Stcrd.TRNQTY) * PercentSum) <> 0 Then
                        Call m_Details1.GetString(1, TempStr1, TempStr2)
                        strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
                        Call Vsp.AddTable(strFormat, "", TempStr2)
                        i = i + 1
                     End If
                  End If
               End If
               
                PrevKey3 = tempCusArea.MASTER_AREA_ID  '!
                Set temp_Area2 = GetAreaCom(m_AreaCod, tempCusArea.MASTER_AREA_ID)
                 PrevKey4 = temp_Area2.MASTER_AREA_NAME '!
                 PrevKey5 = Str(Month(Stcrd.DOCDAT)) & "/" & Str(Year(Stcrd.DOCDAT) + 543)
          '       PrevKey6 = Stcrd.CUSCOD

End If ' ยอดเกิน มีสิทธิ์ได้
      
' ******* ' ไม่ใช่สินค้า Incentive คิดแต่เป็นลูกค้าใหม่หรือเปล่าอย่างเดียว  ' ไม่อยู่ในเงื่อนไขธรรมดา
If m_runIncenPro2 Is Nothing And m_runConditionCommiss2 Is Nothing Then

         Call GetFirstLastDate(DateSerial(Year(Stcrd.DOCDAT), Month(Stcrd.DOCDAT) - 6, 1), SixMonthFirst)
         Call GetFirstLastDate(DateSerial(Year(Stcrd.DOCDAT), Month(Stcrd.DOCDAT) - 1, 1), , SixMonthLast)
         If Year(SixMonthFirst) = Year(SixMonthLast) Then
             Call LoadNewCus(m_NewCus, SixMonthFirst, SixMonthLast, Stcrd.CUSCOD)    ' ต้องดูย้อนหลังว่าลูกค้าคนนี้ย้อนไป 6 เดือนที่แล้ว มีใบสั่งซื้อไหม
         Else
            Call GetFirstLastDate(DateSerial(Year(SixMonthFirst), Month(1), 31), , dayLast)
            Call LoadNewCus(m_NewCus, SixMonthFirst, dayLast, Stcrd.CUSCOD)          ' ปีเก่า เดือนนั้นๆ ถึงสิ้นปี
            If m_NewCus.Count = 0 Then
                Call GetFirstLastDate(DateSerial(Year(SixMonthLast), Month(12), 1), , dayFirst)
                Call LoadNewCus(m_NewCus, dayFirst, SixMonthLast, Stcrd.CUSCOD)              ' ปีใหม่ ต้นปี ถึง เดือนนั้นๆ
            End If
         End If
         Set temp_Stcrd = GetStkcodNoNew(m_AllNewCus, Stcrd.CUSCOD)
         If m_NewCus.Count = 0 And temp_Stcrd Is Nothing And NETVAL > 5000 Then     '    And Stcrd.CUSCOD <> PrevKey6

            If PrevKey5 <> Str(Month(Stcrd.DOCDAT)) & "/" & Str(Year(Stcrd.DOCDAT) + 543) And PrevKey5 <> "" Then
                  M = M + 1
                  If mcolParam("SUMMARY_MODE") = "N" Then
                     Call GenerateFooter("", "รวม", PrevKey5, Total4, m_Details1, haveData_ByArea)       ' เช็ค haveValue ใน Call GenerateFooter **** เซลล์ต้องมีค่าถึงจะแสดงบรรทัดรวม
                     Vsp.FontBold = True
                  Else
                     Call GenerateFooter("", Str(M), PrevKey5, Total4, m_Details1, haveData_ByArea)        ' เช็ค haveValue ใน Call GenerateFooter **** เซลล์ต้องมีค่าถึงจะแสดงบรรทัดรวม
                     Vsp.FontBold = False
                  End If

                 If haveData_ByArea = True Then
                      M = M + 1
                     If mcolParam("SHOWZERO_MODE") = "Y" Then   ' ติ๊ก = โชว์ 0 ด้วย
                        Call m_Details1.GetString(2, TempStr1, TempStr2)
                        strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
                        Call Vsp.AddTable(strFormat, "", TempStr2)
                     Else
                          If Total4(15) <> 0 Then
                          Call m_Details1.GetString(2, TempStr1, TempStr2)
                          strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
                          Call Vsp.AddTable(strFormat, "", TempStr2)
                          End If
                     End If
                End If

                     Vsp.FontBold = False
                     For j = 1 To UBound(Total4)
                         Total4(j) = 0
                     Next j
            End If

               If tempCusArea.MASTER_AREA_ID <> PrevKey3 Then
                   If mcolParam("SUMMARY_MODE") = "N" Then
                     Call GenerateFooter("", "รวม", PrevKey4, Total2, m_Details1, haveData_ByArea)        ' เช็ค haveValue ใน Call GenerateFooter **** เซลล์ต้องมีค่าถึงจะแสดงบรรทัดรวม
                     Vsp.FontBold = True
                  Else
                     Call GenerateFooter("", "", PrevKey4, Total2, m_Details1, haveData_ByArea)        ' เช็ค haveValue ใน Call GenerateFooter **** เซลล์ต้องมีค่าถึงจะแสดงบรรทัดรวม
                     Vsp.FontBold = False
                  End If

                   If haveData_ByArea = True Then
                     K = K + 1
                     If mcolParam("SHOWZERO_MODE") = "Y" Then   ' ติ๊ก = โชว์ 0 ด้วย
                        Call m_Details1.GetString(2, TempStr1, TempStr2)
                        strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
                        Call Vsp.AddTable(strFormat, "", TempStr2)
                     Else
                          If Total2(15) <> 0 Then
                          Call m_Details1.GetString(2, TempStr1, TempStr2)
                          strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
                          Call Vsp.AddTable(strFormat, "", TempStr2)
                          End If
                     End If
                     
                     If mcolParam("SAVE_MODE") = "Y" Then
                        Call SaveToDB(PrevKey1, PrevKey2, "03", PrevKey3, 0, Total2, mcolParam("FROM_CMPL_DATE"), mcolParam("TO_CMPL_DATE"))             'PreKey3=เลขเขต ,Total2
                     End If
                        
                  End If
                        Vsp.FontBold = False

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
             End If

               Call m_Details1.ClearField
               j = 0
               haveData = False
               
                If i = 10 Then
                  'debug.print
                End If
                
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
               TempStr = "(*) " & Left(Stcrd.CUSNAM, 25)
               Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
               Call m_Details1.AddField(BodyCf)

               j = j + 1
               Set HeadCf = m_Headers1.Fields(j)
               TempStr = Stcrd.STKDES
               Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
               Call m_Details1.AddField(BodyCf)

              j = j + 1
               Set HeadCf = m_Headers1.Fields(j)
               TempStr = Stcrd.TRNQTY
               Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
               Call m_Details1.AddField(BodyCf)
                     Total4(j) = Total4(j) + (Val(Stcrd.TRNQTY))
                     Total3(j) = Total3(j) + (Val(Stcrd.TRNQTY))
                     Total2(j) = Total2(j) + (Val(Stcrd.TRNQTY))
                     Total1(j) = Total1(j) + (Val(Stcrd.TRNQTY))
               If haveData = False And Val(Stcrd.TRNQTY) <> 0 Then
                  haveData = True
               End If
               
               j = j + 1
               Set HeadCf = m_Headers1.Fields(j)
               TempStr = FormatNumber(NETVAL / Stcrd.TRNQTY)
               Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
               Call m_Details1.AddField(BodyCf)
               If haveData = False And Val(Stcrd.UNITPR) <> 0 Then
                  haveData = True
               End If

              PayCompleted = True
               j = j + 1
               Set HeadCf = m_Headers1.Fields(j)
               TempStr = FormatNumber(NETVAL)                                   ' เงินที่จ่ายจริง
               Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
               Call m_Details1.AddField(BodyCf)
                     Total4(j) = Total4(j) + (NETVAL)
                     Total3(j) = Total3(j) + (NETVAL)
                     Total2(j) = Total2(j) + (NETVAL)
                     Total1(j) = Total1(j) + (NETVAL)
               If haveData = False And NETVAL <> 0 Then
                  haveData = True
               End If

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

               DueCount = DateDiff("D", Stcrd.DOCDAT, CMPLDAT)  'ต้องคิดตรงนี้เพื่อไปคำนวณค่า
                j = j + 1
               Set HeadCf = m_Headers1.Fields(j)
               If (mcolParam("FROM_CMPL_DATE") <= CMPLDAT) And (Stcrd.CMPLDAT <= toCMPLdate) And PayCompleted And CMPLDAT <> "29/12/2442" Then
                     TempStr = DueCount                                                                                           ' เครดิต
               Else:
                     TempStr = ""
               End If
               Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
               Call m_Details1.AddField(BodyCf)

              If DueCount <= NumCR Then
                     NEWCUSVAL = 300
                     TempStr = FormatNumber(NEWCUSVAL)       '  แล้วเก็บไว้ใน Colection ว่าของคนนี้ให้แล้ว
                     Set temp_Stcrd = New CStcrd
                     temp_Stcrd.CUSCOD = Stcrd.CUSCOD
                     Call m_AllNewCus.Add(temp_Stcrd, Trim(Stcrd.CUSCOD))
                     Set temp_Stcrd = Nothing
               Else
                    TempStr = ""       '  แล้วเก็บไว้ใน Colection ว่าของคนนี้ให้แล้ว
               End If

               j = j + 1
              Set HeadCf = m_Headers1.Fields(j)
               If (mcolParam("FROM_CMPL_DATE") <= CMPLDAT) And (Stcrd.CMPLDAT <= toCMPLdate) And PayCompleted And CMPLDAT <> "29/12/2442" Then
                     Total4(j) = Total4(j) + NEWCUSVAL                              ' ลูกค้าใหม่
                     Total3(j) = Total3(j) + NEWCUSVAL
                     Total2(j) = Total2(j) + NEWCUSVAL
                     Total1(j) = Total1(j) + NEWCUSVAL
               Else:
                     TempStr = ""
               End If
               Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
               Call m_Details1.AddField(BodyCf)
               If haveData = False And (NEWCUSVAL) <> 0 Then
                  haveData = True
               End If


                j = j + 1
               Set HeadCf = m_Headers1.Fields(j)
               TempStr = ""
               Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
               Call m_Details1.AddField(BodyCf)
               If haveData = False And (PercentSum) <> 0 Then
                  haveData = True
               End If

               j = j + 1
               Set HeadCf = m_Headers1.Fields(j)
               TempStr = ""
               Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
               Call m_Details1.AddField(BodyCf)
               If haveData = False And (Val(Stcrd.TRNQTY) * PercentSum) <> 0 Then
                  haveData = True
               End If

                j = j + 1
               Set HeadCf = m_Headers1.Fields(j)
               If (mcolParam("FROM_CMPL_DATE") <= CMPLDAT) And (Stcrd.CMPLDAT <= toCMPLdate) And PayCompleted And CMPLDAT <> "29/12/2442" Then
                     TempStr = FormatNumber(NEWCUSVAL)                                   ' total
                     Total4(j) = Total4(j) + NEWCUSVAL
                     Total3(j) = Total3(j) + NEWCUSVAL
                     Total2(j) = Total2(j) + NEWCUSVAL
                     Total1(j) = Total1(j) + NEWCUSVAL
               Else:
                     TempStr = ""
               End If
               Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
               Call m_Details1.AddField(BodyCf)
               If haveData = False And (Val(Stcrd.TRNQTY) * PercentSum) <> 0 Then
                  haveData = True
               End If


               If mcolParam("SUMMARY_MODE") = "N" And haveData = True Then
                    Call m_Details1.GetString(1, TempStr1, TempStr2)
                    strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
                    Call Vsp.AddTable(strFormat, "", TempStr2)
                    i = i + 1   'นับบรรทัดข้อมูลย่อย
               End If

                PrevKey3 = tempCusArea.MASTER_AREA_ID  '!
                Set temp_Area2 = GetAreaCom(m_AreaCod, tempCusArea.MASTER_AREA_ID)
                 PrevKey4 = temp_Area2.MASTER_AREA_NAME '!
                 PrevKey5 = Str(Month(Stcrd.DOCDAT)) & "/" & Str(Year(Stcrd.DOCDAT) + 543)
          '       PrevKey6 = Stcrd.CUSCOD

End If   'เข้าข่ายสินค้านอกเหนือ แต่ต้องให้ Incentive
       End If   ' จบสินค้าตัวนั้นๆ

      End If   ' IVนี้ จ่ายREครบ
   End If     ' RE จาก IV ไม่เป็น Nothing
End If  'จบ เจอว่าอยู่ในเขตนั้นๆ
          
            If mcolParam("AREA_TYPE_ID") > 0 And Rs.AbsolutePosition = iCount Then    ' ถ้ามีการเลือกเขต ให้มีข้อมูลอยู่ในอเรย์นั้นๆ
                 Exit For
            End If
                   
'                Rs.MoveNext                                                                                            ' วนคน    '  **** วนทุกบรรทัดของ Stcrd
'         Wend
   Next Stcrd
Next temp_Area


            If mcolParam("SUMMARY_MODE") = "N" Then
               Call GenerateFooter("", "รวม", PrevKey5, Total4, m_Details1, haveData_ByArea)        ' เช็ค haveValue ใน Call GenerateFooter **** เซลล์ต้องมีค่าถึงจะแสดงบรรทัดรวม
               Vsp.FontBold = True
            Else
               Call GenerateFooter("", Str(M), PrevKey5, Total4, m_Details1, haveData_ByArea)        ' เช็ค haveValue ใน Call GenerateFooter **** เซลล์ต้องมีค่าถึงจะแสดงบรรทัดรวม
               Vsp.FontBold = False
            End If
            
              If haveData_ByArea = True Then
                  If mcolParam("SHOWZERO_MODE") = "Y" Then
                        Call m_Details1.GetString(2, TempStr1, TempStr2)
                        strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
                        Call Vsp.AddTable(strFormat, "", TempStr2)
                 Else
                          If Total4(15) <> 0 Then
                          Call m_Details1.GetString(2, TempStr1, TempStr2)
                          strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
                          Call Vsp.AddTable(strFormat, "", TempStr2)
                          End If
                End If
            End If
                 Vsp.FontBold = False
                 
         For j = 1 To UBound(Total4)
             Total4(j) = 0
         Next j

      
       If mcolParam("SUMMARY_MODE") = "N" Then                                                         ' เขตพื้นที่
        Call GenerateFooter("", "รวม", PrevKey4, Total2, m_Details1, haveData_ByArea)
         Vsp.FontBold = True
      Else
         Call GenerateFooter("", "", PrevKey4, Total2, m_Details1, haveData_ByArea)
         Vsp.FontBold = False
      End If
           ' เช็ค haveValue ใน Call GenerateFooter **** บรรทัดสุดท้ายของเซลล์
       If haveData_ByArea = True Then
            K = K + 1
                  If mcolParam("SHOWZERO_MODE") = "Y" Then
                        Call m_Details1.GetString(2, TempStr1, TempStr2)
                        strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
                        Call Vsp.AddTable(strFormat, "", TempStr2)
                 Else
                          If Total2(15) <> 0 Then
                          Call m_Details1.GetString(2, TempStr1, TempStr2)
                          strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
                          Call Vsp.AddTable(strFormat, "", TempStr2)
                          End If
                  End If
          If mcolParam("SAVE_MODE") = "Y" Then
                   Call SaveToDB(PrevKey1, PrevKey2, "03", PrevKey3, 0, Total2, mcolParam("FROM_CMPL_DATE"), mcolParam("TO_CMPL_DATE"))             'PreKey3=เลขเขต ,Total2
          End If
            
       End If
            Vsp.FontBold = False

      For j = 1 To UBound(Total2)
         Total2(j) = 0
      Next j

            TempStr = "รวม"
            Vsp.FontBold = True

      Call GenerateFooter("", TempStr, "  ** " & PrevKey2 & " **", Total3, m_Details1, haveValueSale)      ' เช็ค haveValue ใน Call GenerateFooter ****
      If haveValueSale = True Then                                                                       ' เซลล์ต้องมีค่า ถึงจะแสดงบรรทัดรวม   ****
                  If mcolParam("SHOWZERO_MODE") = "Y" Then
                        Call m_Details1.GetString(2, TempStr1, TempStr2)
                        strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
                        Call Vsp.AddTable(strFormat, "", TempStr2)
                 Else
                          If Total3(15) <> 0 Then
                          Call m_Details1.GetString(2, TempStr1, TempStr2)
                          strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
                          Call Vsp.AddTable(strFormat, "", TempStr2)
                          End If
                  End If
      End If
             Vsp.FontBold = False

           PrevKey3 = 0
           PrevKey4 = ""
           
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
     ElseIf j = 5 Then
         TempStr = ""
      ElseIf j = 6 Then
         TempStr = ""
      ElseIf j = 7 Then
         TempStr = ""
      ElseIf j = 8 Then
         TempStr = ""
      ElseIf j = 9 Then
         TempStr = ""
      ElseIf j = 10 Then
         TempStr = ""
       ElseIf j = 11 Then
         TempStr = ""
      ElseIf j = 13 Then
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
If exit_Record Is Nothing Then  ' ถ้าไม่มีในคอเล็กชั่น คือ เพิ่มได้
        
  If ConfirmSaveRecord(Txt2 & "-" & Tot(15) & " บ.") Then
      comRecord.AddEditMode = SHOW_ADD
      comRecord.SLMCOD = Txt1
      comRecord.SLMNAME = Txt2
      comRecord.COMTYP = COMTYP
      comRecord.MASTER_AREA_ID = Txt3
      comRecord.GOODS_GROUP_ID = 1
      comRecord.COM_VALUE = Tot(15)
      comRecord.FROMDAT = FromDate
      comRecord.TODAT = ToDate
   
      If Not glbDaily.AddEditComRecord(comRecord, IsOK, True, glbErrorLog) Then
         glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
      End If
   End If
Else
     If ConfirmSaveExist(Txt2 & "-" & Tot(15) & " บ.", exit_Record.SLMNAME & "-" & exit_Record.COM_VALUE & " บ.") Then
         comRecord.AddEditMode = SHOW_EDIT
         comRecord.RECORD_ID = exit_Record.RECORD_ID
         comRecord.SLMCOD = Txt1
         comRecord.SLMNAME = Txt2
         comRecord.COMTYP = COMTYP
         comRecord.MASTER_AREA_ID = Txt3
         comRecord.GOODS_GROUP_ID = Txt4
         comRecord.COM_VALUE = Tot(15)
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
