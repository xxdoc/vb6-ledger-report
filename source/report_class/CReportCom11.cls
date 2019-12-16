VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CReportCom11"
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
            Call Cf.SetFieldValue(18, "^", "", "<")    '27
      Call m_Headers0.AddField(Cf)
   Else
           Call Cf.SetFieldValue(2, "^", "", "^")
      Call m_Headers0.AddField(Cf)
            Call Cf.SetFieldValue(11, "^", "", "<")    '27
      Call m_Headers0.AddField(Cf)
   End If
      
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
      Call Cf.SetFieldValue(2, "^", "NO.", "<")
   Call m_Headers1.AddField(Cf)

   Call Cf.SetFieldValue(0, "^", "INVOICE", "<")
   Call m_Headers1.AddField(Cf)
   
         Call Cf.SetFieldValue(5, "^", "รหัส - ชื่อสินค้า", "<")
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
Else
   Call Cf.SetFieldValue(0, "^", "วันส่งสินค้า", "^")
   Call m_Headers1.AddField(Cf)
 End If
      
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
   
      MODULE_DESC = glbCompanyName & vbCrLf & mcolParam("REPORT_NAME") & _
                             vbCrLf & "วันที่ส่งสินค้า   " & EmptyToString(DateToStringExtEx2(mcolParam("FROM_DOC_DATE")), "N/A") & "    ถึง    " & EmptyToString(DateToStringExtEx2(mcolParam("TO_DOC_DATE")), "N/A") & _
                             vbCrLf & "จากรหัสสินค้า   " & EmptyToString(mcolParam("FROM_GOODS_CODE"), "N/A") & "    ถึง    " & EmptyToString(mcolParam("TO_GOODS_CODE"), "N/A")
   
   If Len(mcolParam("FROM_CUSTOMER_CODE")) > 0 Then
             MODULE_DESC = MODULE_DESC & "    ลูกค้า  " & EmptyToString(mcolParam("FROM_CUSTOMER_CODE"), "N/A")
     End If
      If Len(mcolParam("TO_CUSTOMER_CODE")) > 0 Then
             MODULE_DESC = MODULE_DESC & "   ถึง    " & EmptyToString(mcolParam("TO_CUSTOMER_CODE"), "N/A")
     End If
       
       If Len(mcolParam("AREA_TYPE_NAME")) > 0 Then
              MODULE_DESC = MODULE_DESC & "      เขต :  " & EmptyToString(mcolParam("AREA_TYPE_NAME"), "N/A")
      End If
 
    Call LoadMinusIV(m_MinusStkcod, mcolParam("FROM_DOC_DATE"), mcolParam("TO_DOC_DATE"))

   Call LoadYearId(YEAR_ID, mcolParam("TO_DOC_DATE"), mcolParam("TO_DOC_DATE"))   ' เวลานี้ใช้ Area_Year ไหน
   If mcolParam("AREA_TYPE_ID") > 0 Then
      Call LoadAreaComReport(Nothing, m_AreaCod, YEAR_ID, Val(Left(mcolParam("AREA_TYPE_NAME"), 2)))
   Else
      Call LoadAreaComReport(Nothing, m_AreaCod, YEAR_ID)       ' โหลดไว้สำหรับ combo
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
Dim PrevKey6 As String    ' สินค้า
Dim PrevKey7 As String    ' สินค้า As String    ' สินค้า
Dim PrevKey8 As String    ' รหัสลูกค้า
Dim PrevKey9 As String    ' ชื่อลูกค้า
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
      
      Set Stcrd = New CStcrd
      Stcrd.FROM_DOC_DATE = mcolParam("FROM_DOC_DATE")
      Stcrd.TO_DOC_DATE = mcolParam("TO_DOC_DATE")
      Stcrd.FROM_STOCK_CODE = mcolParam("FROM_GOODS_CODE")
      Stcrd.TO_STOCK_CODE = mcolParam("TO_GOODS_CODE")
      Stcrd.FROM_CUSTOMER_CODE = mcolParam("FROM_CUSTOMER_CODE")
      Stcrd.TO_CUSTOMER_CODE = mcolParam("TO_CUSTOMER_CODE")
      'Stcrd.SLMCOD = ArS.SLMCOD
      Call Stcrd.QueryData(23, Rs, iCount)   ' ติดต่อกับดาต้าเบสที่ 1
            While Not Rs.EOF     '  **** วนทุกบรรทัดของ Stcrd
         Call Stcrd.PopulateFromRS(23, Rs)
'            Set stcrd_temp = GetObject("CStcrd", stcrd_mixdb, Stcrd.DOCNUM & "-" & Stcrd.STKCOD, False)
'            If stcrd_temp Is Nothing Then  ' ถ้าไม่มีในคอเล็กชั่น คือ เพิ่มได้
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
             Call stcrd_mixdb.Add(stcrd_temp)   ' , stcrd_temp.DOCNUM & "-" & stcrd_temp.STKCOD
             Set stcrd_temp = Nothing
         '   End If
         Rs.MoveNext                                                                                            ' วนคน    '  **** วนทุกบรรทัดของ Stcrd
      Wend
      
For Each temp_Area In m_AreaCod
          
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
          
     ''     If mcolParam("SHOWCUS_MODE") = "Y" Then
           If (PrevKey8 <> Stcrd.STKCOD Or PrevKey6 <> Stcrd.CUSCOD) And PrevKey8 <> "" Then
                  If mcolParam("SUMMARY_MODE") = "N" Then
                     Call GenerateFooter("", "   รวมสินค้า", "   " & PrevKey8 & " - " & PrevKey9, Total3, m_Details1)        ' เช็ค haveValue ใน Call GenerateFooter **** เซลล์ต้องมีค่าถึงจะแสดงบรรทัดรวม
                     Vsp.FontBold = True
                  Else
                     Call GenerateFooter("                    " & PrevKey8, "", "    " & PrevKey9, Total3, m_Details1)            ' เช็ค haveValue ใน Call GenerateFooter **** เซลล์ต้องมีค่าถึงจะแสดงบรรทัดรวม
                     Vsp.FontBold = False
                  End If
                  Call m_Details1.GetString(2, TempStr1, TempStr2)
                  strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
                  Call Vsp.AddTable(strFormat, "", TempStr2)
                  Vsp.FontBold = False
                  
                  For j = 1 To UBound(Total3)
                      Total3(j) = 0
                  Next j
           End If
      ''   End If

           If mcolParam("SHOWCUS_MODE") = "Y" Then
              If PrevKey6 <> Stcrd.CUSCOD And PrevKey6 <> "" Then
                  If mcolParam("SUMMARY_MODE") = "N" Then
                     Call GenerateFooter("", "รวม", PrevKey6 & "  -  " & PrevKey7, Total1, m_Details1)         ' เช็ค haveValue ใน Call GenerateFooter **** เซลล์ต้องมีค่าถึงจะแสดงบรรทัดรวม
                     Vsp.FontBold = True
                  ElseIf mcolParam("SUMMARY_MODE") = "Y" And mcolParam("SHOWCUS_MODE") = "Y" Then
                     Call GenerateFooter("                  * รวมลูกค้า", "", PrevKey6 & "  -  " & PrevKey7, Total1, m_Details1)         ' เช็ค haveValue ใน Call GenerateFooter **** เซลล์ต้องมีค่าถึงจะแสดงบรรทัดรวม
                     Vsp.FontBold = True
                  Else
                     Call GenerateFooter("   " & Str(K), "", PrevKey6 & "  -  " & PrevKey7, Total1, m_Details1)         ' เช็ค haveValue ใน Call GenerateFooter **** เซลล์ต้องมีค่าถึงจะแสดงบรรทัดรวม
                     Vsp.FontBold = False
                  End If
                  Call m_Details1.GetString(2, TempStr1, TempStr2)
                  strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
                  Call Vsp.AddTable(strFormat, "", TempStr2)
                  Vsp.FontBold = False
                  K = K + 1
                  
                  For j = 1 To UBound(Total1)
                      Total1(j) = 0
                  Next j
               End If
            End If

               
         
               If tempCusArea.MASTER_AREA_ID <> PrevKey3 Then
                 If PrevKey3 <> 0 Then
'                   If mcolParam("SUMMARY_MODE") = "N" Then
                     Call GenerateFooter("รวม", "", PrevKey4, Total2, m_Details1)        ' เช็ค haveValue ใน Call GenerateFooter **** เซลล์ต้องมีค่าถึงจะแสดงบรรทัดรวม
                     Vsp.FontBold = True
'                  Else
'                     Call GenerateFooter("", "", PrevKey4, Total2, m_Details1)        ' เช็ค haveValue ใน Call GenerateFooter **** เซลล์ต้องมีค่าถึงจะแสดงบรรทัดรวม
'                     Vsp.FontBold = False
'                  End If
                        Call m_Details1.GetString(2, TempStr1, TempStr2)
                        strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
                        Call Vsp.AddTable(strFormat, "", TempStr2)
                        Vsp.FontBold = False
                        
                  For j = 1 To UBound(Total2)
                            Total2(j) = 0
                  Next j
                  End If
                  
                   
                        Set temp_Area2 = GetAreaCom(m_AreaCod, tempCusArea.MASTER_AREA_ID)    'เฮดเดอร์ใหม่
                        Call GenerateHeader("เขต : ", temp_Area2.MASTER_AREA_NAME, m_Details1)   '!
                        Call m_Details1.GetString(1, TempStr1, TempStr2)
                        strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
                        Vsp.FontBold = True
                        Call Vsp.AddTable(strFormat, "", TempStr2)
                        Vsp.FontBold = False
      
             End If

               If mcolParam("SHOWCUS_MODE") = "Y" Then  'mcolParam("SUMMARY_MODE") = "Y" And
                  If Stcrd.CUSCOD <> PrevKey6 Then
                        Call GenerateHeader("ลูกค้า                     ", Stcrd.CUSCOD & " - " & Stcrd.CUSNAM, m_Details1)    '!
                        Call m_Details1.GetString(1, TempStr1, TempStr2)
                        strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
                        Vsp.FontBold = True
                        Call Vsp.AddTable(strFormat, "", TempStr2)
                        Vsp.FontBold = False
                    End If
                  End If


'             If Stcrd.DOCNUM = "IV0042361" Then
'               'debug.print
'             End If

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
               TempStr = "       " & Stcrd.DOCNUM
               Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
               Call m_Details1.AddField(BodyCf)

                 j = j + 1
               Set HeadCf = m_Headers1.Fields(j)
               TempStr = "       " & Left(Stcrd.CUSNAM, 27)
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
               TempStr = FormatNumber(NETVAL)                                   ' เงินที่จ่ายจริง 7
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
               
               If mcolParam("SUMMARY_MODE") = "N" Then
                ' If mcolParam("SHOWZERO_MODE") = "Y" Then              ' ติ๊ก = โชว์ 0 ด้วย    And haveData = True
                    Call m_Details1.GetString(1, TempStr1, TempStr2)
                    strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
                    Call Vsp.AddTable(strFormat, "", TempStr2)
                    i = i + 1
                '  Else
'                     If Val(NEWCUSVAL) + (Val(Stcrd.TRNQTY) * PercentSum) <> 0 Then
'                        Call m_Details1.GetString(1, TempStr1, TempStr2)
'                        strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
'                        Call Vsp.AddTable(strFormat, "", TempStr2)
'                        i = i + 1
'                     End If
'                  End If
               End If
               
                PrevKey3 = tempCusArea.MASTER_AREA_ID  '!
                Set temp_Area2 = GetAreaCom(m_AreaCod, tempCusArea.MASTER_AREA_ID)
                 PrevKey4 = temp_Area2.MASTER_AREA_NAME '!
                 PrevKey5 = Str(Month(Stcrd.DOCDAT)) & "/" & Str(Year(Stcrd.DOCDAT) + 543)
                 PrevKey8 = Stcrd.STKCOD
                 PrevKey9 = Stcrd.STKDES
                 PrevKey6 = Stcrd.CUSCOD
                 PrevKey7 = Stcrd.CUSNAM
          
            If mcolParam("AREA_TYPE_ID") > 0 And Rs.AbsolutePosition = iCount Then    ' ถ้ามีการเลือกเขต ให้มีข้อมูลอยู่ในอเรย์นั้นๆ
                 Exit For
            End If
                   
 
 End If
   Next Stcrd
      
Next temp_Area
      
               
                  If mcolParam("SUMMARY_MODE") = "N" Then
                     Call GenerateFooter("", "   รวมสินค้า", "   " & PrevKey8 & " - " & PrevKey9, Total3, m_Details1)        ' เช็ค haveValue ใน Call GenerateFooter **** เซลล์ต้องมีค่าถึงจะแสดงบรรทัดรวม
                     Vsp.FontBold = True
                  Else
                     Call GenerateFooter("                    " & PrevKey8, "", "    " & PrevKey9, Total3, m_Details1)            ' เช็ค haveValue ใน Call GenerateFooter **** เซลล์ต้องมีค่าถึงจะแสดงบรรทัดรวม
                     Vsp.FontBold = False
                  End If
                  If mcolParam("SHOWCUS_MODE") = "N" Or mcolParam("SUMMARY_MODE") = "N" Then
                        Call m_Details1.GetString(2, TempStr1, TempStr2)
                        strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
                        Call Vsp.AddTable(strFormat, "", TempStr2)
                        Vsp.FontBold = False
                  End If
      
            If mcolParam("SHOWCUS_MODE") = "Y" Then
              '   If PrevKey8 <> Stcrd.CUSCOD And PrevKey8 <> "" Then
                  If mcolParam("SUMMARY_MODE") = "N" Then
                     Call GenerateFooter("", "   * รวมลูกค้า", "    " & PrevKey6 & "  -  " & PrevKey7, Total1, m_Details1)        ' เช็ค haveValue ใน Call GenerateFooter **** เซลล์ต้องมีค่าถึงจะแสดงบรรทัดรวม
                     Vsp.FontBold = True
                  Else
                     Call GenerateFooter("                    " & PrevKey8, "", "    " & PrevKey9, Total3, m_Details1)            ' เช็ค haveValue ใน Call GenerateFooter **** เซลล์ต้องมีค่าถึงจะแสดงบรรทัดรวม
                     Vsp.FontBold = False
                  End If
                  Call m_Details1.GetString(2, TempStr1, TempStr2)
                  strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
                  Call Vsp.AddTable(strFormat, "", TempStr2)
                  Vsp.FontBold = False
                  
                  For j = 1 To UBound(Total3)
                      Total3(j) = 0
                  Next j
              '   End If
            End If
      
                 If mcolParam("SUMMARY_MODE") = "N" Then
                   Call GenerateFooter("", "รวม", PrevKey6 & "  -  " & PrevKey7, Total1, m_Details1)       ' เช็ค haveValue ใน Call GenerateFooter **** เซลล์ต้องมีค่าถึงจะแสดงบรรทัดรวม
                     Vsp.FontBold = True
                  ElseIf mcolParam("SUMMARY_MODE") = "Y" And mcolParam("SHOWCUS_MODE") = "Y" Then
                     Call GenerateFooter("                  * รวมลูกค้า", "", PrevKey6 & "  -  " & PrevKey7, Total1, m_Details1)         ' เช็ค haveValue ใน Call GenerateFooter **** เซลล์ต้องมีค่าถึงจะแสดงบรรทัดรวม
                     Vsp.FontBold = True
                  Else
                     Call GenerateFooter("   " & Str(K), "", PrevKey6 & "  -  " & PrevKey7, Total1, m_Details1)         ' เช็ค haveValue ใน Call GenerateFooter **** เซลล์ต้องมีค่าถึงจะแสดงบรรทัดรวม
                     Vsp.FontBold = False
                  End If
                   
                  If mcolParam("SHOWCUS_MODE") = "Y" And mcolParam("SUMMARY_MODE") = "Y" Then
                        Call m_Details1.GetString(2, TempStr1, TempStr2)
                        strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
                        Call Vsp.AddTable(strFormat, "", TempStr2)
                        Vsp.FontBold = False
                  End If
             '     For j = 1 To UBound(Total1)
              '        Total2(j) = 0
            '      Next j
         
 '                 If mcolParam("SUMMARY_MODE") = "N" Then
                     Call GenerateFooter("รวม", "", PrevKey4, Total2, m_Details1)        ' เช็ค haveValue ใน Call GenerateFooter **** เซลล์ต้องมีค่าถึงจะแสดงบรรทัดรวม
                     Vsp.FontBold = True
'                  Else
'                     Call GenerateFooter(Str(K), "", PrevKey4, Total2, m_Details1)        ' เช็ค haveValue ใน Call GenerateFooter **** เซลล์ต้องมีค่าถึงจะแสดงบรรทัดรวม
'                     Vsp.FontBold = False
'                  End If
                        Call m_Details1.GetString(2, TempStr1, TempStr2)
                        strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
                        Call Vsp.AddTable(strFormat, "", TempStr2)
                        Vsp.FontBold = False
                        K = K + 1

     
           Call GenerateFooter("สุทธิ", "", "", Total4, m_Details1)
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
      ElseIf j = 4 Then
         TempStr = ""
     ElseIf j = 8 Then
         TempStr = ""
      ElseIf j = 9 Then
         TempStr = ""
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
