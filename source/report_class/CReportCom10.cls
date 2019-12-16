VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CReportCom10"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'User Group Report
Implements CReportInterface

Private Const MODULE_NAME = "CReportCOM10"
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

'Private m_cusFromArea(20) As Collection
Private nonidentifyArea As Boolean
Dim Stb As CCommissMasterArea
Private m_IVcenter As Collection
Private m_IVinArea As Collection
Private m_MinusStkcod As Collection
Private m_IVincomplete0 As Collection
Private m_IVincomplete12 As Collection
Private m_IVfromEx As Collection
Private m_REsumIV As Collection
Private m_ReDocdat As Collection
Private temp_Area As CCommissionCustomerArea
Private temp_Area2 As CCommissionCustomerArea
Public REsumIV As CARRcIt
Private tempComDonStkcod As CComDonStk
Private m_ComDonStkcod As Collection
Dim L As Long

Dim Total1(100) As Double
Dim Total2(100) As Double
Dim Total3(100) As Double
Dim strNETVAL As String
Dim strMinusCod As String
Dim haveData As Boolean
Dim i As Long
Dim NETVAL As Double
Dim PayCompleted  As Boolean
Dim CMPLDAT As Date
Dim TempCConditionCommiss As CConditionCommission
Dim m_runConditionCommiss As CConditionCommission
Dim RECOD As String
Dim PercentNum1 As Double
Dim tempREdoc As CARTrn
Dim DueCount As Long

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
   Set Stb = New CCommissMasterArea
   Set m_IVcenter = New Collection
   Set m_IVinArea = New Collection
   Set m_MinusStkcod = New Collection
   Set m_IVincomplete0 = New Collection
   Set m_IVincomplete12 = New Collection
   Set m_IVfromEx = New Collection
   Set m_REsumIV = New Collection
   Set m_ReDocdat = New Collection
   Set temp_Area = New CCommissionCustomerArea
   Set temp_Area2 = New CCommissionCustomerArea
   Set REsumIV = New CARRcIt
      Set m_ComDonStkcod = New Collection
   Set tempComDonStkcod = New CComDonStk
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
   Set Stb = Nothing
   Set m_IVcenter = Nothing
   Set m_IVinArea = Nothing
   Set m_MinusStkcod = Nothing
   Set m_IVincomplete0 = Nothing
   Set m_IVincomplete12 = Nothing
   Set m_IVfromEx = Nothing
   Set m_REsumIV = Nothing
   Set m_ReDocdat = Nothing
   Set temp_Area = Nothing
   Set temp_Area2 = Nothing
   Set REsumIV = Nothing
      Set m_ComDonStkcod = Nothing
   Set tempComDonStkcod = Nothing
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
      
   Call Cf.SetFieldValue(1, "^", "", "^")
   Call m_Headers0.AddField(Cf)
   Call Cf.SetFieldValue(2, "^", "", "<")    '27
   Call m_Headers0.AddField(Cf)
   Call Cf.SetFieldValue(2, "^", "", "^")
   Call m_Headers0.AddField(Cf)
   Call Cf.SetFieldValue(7, "^", "", "^")    '27
   Call m_Headers0.AddField(Cf)
   Call Cf.SetFieldValue(2, "^", "", ">")
   Call m_Headers0.AddField(Cf)
   Call Cf.SetFieldValue(2, "^", "", ">")    '27
   Call m_Headers0.AddField(Cf)
   Call Cf.SetFieldValue(2, "^", "", ">")
   Call m_Headers0.AddField(Cf)
   Call Cf.SetFieldValue(2, "^", "", ">")    '27
   Call m_Headers0.AddField(Cf)

      
'If mcolParam("SUMMARY_MODE") = "N" Then
      Call Cf.SetFieldValue(1, "^", "NO.", "^")
      Call m_Headers1.AddField(Cf)
      
      Call Cf.SetFieldValue(2, "^", "วันส่งสินค้า", "^")
      Call m_Headers1.AddField(Cf)
      
      Call Cf.SetFieldValue(2, "^", "INVOICE", "^")
      Call m_Headers1.AddField(Cf)
      
      Call Cf.SetFieldValue(2, "^", "รหัสสินค้า", "^")
      Call m_Headers1.AddField(Cf)
      
      Call Cf.SetFieldValue(5, "^", "ชื่อสินค้า", "<")
      Call m_Headers1.AddField(Cf)
'Else
'      Call Cf.SetFieldValue(0, "^", "NO.", "^")
'      Call m_Headers1.AddField(Cf)
'
'      Call Cf.SetFieldValue(0, "^", "รหัส", "^")
'      Call m_Headers1.AddField(Cf)
'
''      Call Cf.SetFieldValue(6, "^", "ชื่อลูกค้า", "<")
''      Call m_Headers1.AddField(Cf)
'
'      Call Cf.SetFieldValue(0, "^", "ชื่อสินค้า", "<")
'      Call m_Headers1.AddField(Cf)
'End If

   Call Cf.SetFieldValue(2, "^", "ยอดจริง", ">")
   Call m_Headers1.AddField(Cf)
      
   Call Cf.SetFieldValue(2, "^", "ส่วนลด", ">")
   Call m_Headers1.AddField(Cf)

   Call Cf.SetFieldValue(2, "^", "ยอดขาย", ">")
   Call m_Headers1.AddField(Cf)
   
   Call Cf.SetFieldValue(2, "^", "หมายเลข SR", "^")
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
   
   MODULE_DESC = glbCompanyName & vbCrLf & mcolParam("REPORT_NAME") & _
                             vbCrLf & "วันที่ส่งสินค้า   " & EmptyToString(DateToStringExtEx2(mcolParam("FROM_DOC_DATE")), "N/A") & "    ถึง    " & EmptyToString(DateToStringExtEx2(mcolParam("TO_DOC_DATE")), "N/A")
      If Len(mcolParam("FROM_CUSTOMER_CODE")) > 0 Then
              MODULE_DESC = MODULE_DESC & "    ลูกค้า  " & EmptyToString(mcolParam("FROM_CUSTOMER_CODE"), "N/A")
      End If
       If Len(mcolParam("TO_CUSTOMER_CODE")) > 0 Then
              MODULE_DESC = MODULE_DESC & "   ถึง    " & EmptyToString(mcolParam("TO_CUSTOMER_CODE"), "N/A")
      End If
'       If Len(mcolParam("AREA_TYPE_NAME")) > 0 Then
'              MODULE_DESC = MODULE_DESC & "      เขต :  " & EmptyToString(mcolParam("AREA_TYPE_NAME"), "N/A")
'      End If
  
'  Call LoadCommission04(m_ConditionCommiss4, mcolParam("FROM_DOC_DATE"), mcolParam("TO_DOC_DATE"))
'  Call LoadIVcenter(m_IVcenter, mcolParam("FROM_DOC_DATE"), mcolParam("TO_DOC_DATE"))
'  Call LoadREDocDat(m_ReDocdat, mcolParam("FROM_CMPL_DATE"), mcolParam("TO_CMPL_DATE")) ' ต้องเอามาหมด
'  Call LoadComDonStk(m_ComDonStkcod, mcolParam("FROM_DOC_DATE"), mcolParam("TO_DOC_DATE"))
'  Call LoadREsumIV(Nothing, m_REsumIV)
'  Call LoadIVExStcrd(m_IVfromEx, mcolParam("FROM_DOC_DATE"), mcolParam("TO_DOC_DATE"), Left(mcolParam("FROM_SALE_NAME"), 3), Left(mcolParam("TO_SALE_NAME"), 3))
'  Call LoadYearId(YEAR_ID, mcolParam("TO_DOC_DATE"), mcolParam("TO_DOC_DATE"))   ' เวลานี้ใช้ Area_Year ไหน
'     If mcolParam("AREA_TYPE_ID") > 0 Then
'      Call LoadAreaComReport(Nothing, m_AreaCod, YEAR_ID, Val(Left(mcolParam("AREA_TYPE_NAME"), 2)))
'   Else
'      Call LoadAreaComReport(Nothing, m_AreaCod, YEAR_ID)       ' โหลดไว้สำหรับ combo
'   End If
   
  Call LoadMinusIV(m_MinusStkcod, mcolParam("FROM_DOC_DATE"), mcolParam("TO_DOC_DATE"))

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

Dim iCount As Long
Dim TempStr As String
Dim PrevKey1 As String
Dim PrevKey2 As String
Dim PrevKey3 As Long   'รหัสเขต
Dim PrevKey4 As String    'เขตการขาย
Dim St As CCommissMasterArea
Dim ArS As COESLM
Dim Stcrd As CStcrd

Dim haveValueSale As Boolean
Dim enableIV As Boolean

Dim haveData_ByArea As Boolean

Dim PercentSum As Double
Dim tempCusArea As CCommissionCustomerArea
Dim toCMPLdate As Date
Dim tempMinusStkcod As CComMinusStk
Dim tempIVcenter As CComIVcenter

Dim temnetval As Long
Dim CorrectStkcod As Boolean
Dim IVoldinArea As Boolean
Dim stcrd_IVinArea As CStcrd
Dim ExStcrd As CStcrd
Dim tempMinus As Double
Dim tempExDocNum As String
   
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
   i = 1  'นับจ.บรรทัดข้อมูลย่อย
   K = 1
   Set BodyCf = New CReportField
      
      Set Stcrd = New CStcrd
      Stcrd.FROM_DOC_DATE = mcolParam("FROM_DOC_DATE")
      Stcrd.TO_DOC_DATE = mcolParam("TO_DOC_DATE")
      Stcrd.FROM_CUSTOMER_CODE = mcolParam("FROM_CUSTOMER_CODE")
      Stcrd.TO_CUSTOMER_CODE = mcolParam("TO_CUSTOMER_CODE")
'      Stcrd.FROM_SALE_CODE = mcolParam("FROM_SALE_CODE")
'      Stcrd.TO_SALE_CODE = mcolParam("TO_SALE_CODE")
      Call Stcrd.QueryData(19, Rs, iCount)

      While Not Rs.EOF

            Call Stcrd.PopulateFromRS(19, Rs)
               
            If Stcrd.CUSCOD <> PrevKey1 Then
                If PrevKey1 <> "" And Total1(7) <> 0 Then
                     Call GenerateFooter("", "      รวม      ", PrevKey1, PrevKey2, Total1, m_Details1, True)      ' เช็ค haveValue ใน Call GenerateFooter **** เซลล์ต้องมีค่าถึงจะแสดงบรรทัดรวม
                     Vsp.FontBold = True
                     Call m_Details1.GetString(2, TempStr1, TempStr2)
                     strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
                     Call Vsp.AddTable(strFormat, "", TempStr2)
                     
                     For j = 1 To UBound(Total1)
                         Total1(j) = 0
                      Next j
                  End If
            
              Vsp.FontBold = True
              Vsp = "    " & Stcrd.CUSCOD & "  :    " & Stcrd.CUSNAM
              Vsp.FontBold = False
              
               PrevKey1 = Stcrd.CUSCOD
               PrevKey2 = Stcrd.CUSNAM
            End If
      
            ' คำนวณส่วนลดก่อน
            Set tempMinusStkcod = GetMinusCommiss(m_MinusStkcod, Stcrd.DOCDAT & "-" & Stcrd.DOCNUM & "-" & Stcrd.STKCOD, False)
            If (tempMinusStkcod Is Nothing) Then   'ไม่เจอในคอเล็คชั่น = ไม่มีส่วนลด
                     tempMinus = 0
                     NETVAL = Stcrd.NETVAL
                     strNETVAL = FormatNumber(NETVAL)
                     strMinusCod = ""
                     enableIV = False
            Else:
                     tempMinus = tempMinusStkcod.MINUS_AMOUNT
                     NETVAL = (Stcrd.NETVAL + Val(tempMinusStkcod.MINUS_AMOUNT))
                     strNETVAL = "*" & FormatNumber(NETVAL)
                     strMinusCod = tempMinusStkcod.MINUS_COD
                     If mcolParam("SHOWNONAME_MODE") = "N" Then
                        enableIV = True
                     Else
                        If tempMinusStkcod.MINUS_COD = "" Then
                           enableIV = True
                        End If
                     End If
            End If
            
         If enableIV = True Then
         
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
               TempStr = DateToStringExtEx2(Stcrd.DOCDAT)                                                                                 ' วันที่ส่งสินค้า
               Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
               Call m_Details1.AddField(BodyCf)
   
                j = j + 1
               Set HeadCf = m_Headers1.Fields(j)
               TempStr = Stcrd.DOCNUM
               Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
               Call m_Details1.AddField(BodyCf)

               j = j + 1
               Set HeadCf = m_Headers1.Fields(j)
               TempStr = Stcrd.STKCOD
               Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
               Call m_Details1.AddField(BodyCf)
      
               j = j + 1
               Set HeadCf = m_Headers1.Fields(j)
               TempStr = Stcrd.STKDES
               Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
               Call m_Details1.AddField(BodyCf)


              j = j + 1
               Set HeadCf = m_Headers1.Fields(j)
                 TempStr = FormatNumber(Stcrd.NETVAL)
               Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
               Call m_Details1.AddField(BodyCf)
               Total1(j) = Total1(j) + Stcrd.NETVAL
               If haveData = False And Stcrd.NETVAL <> 0 Then
                  haveData = True
               End If
               
               j = j + 1
               Set HeadCf = m_Headers1.Fields(j)
                 TempStr = FormatNumber(tempMinus)
               Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
               Call m_Details1.AddField(BodyCf)
               Total1(j) = Total1(j) + tempMinus
               If haveData = False And tempMinus <> 0 Then
                  haveData = True
               End If
              
               j = j + 1
               Set HeadCf = m_Headers1.Fields(j)
                 TempStr = strNETVAL
               Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
               Call m_Details1.AddField(BodyCf)
               Total1(j) = Total1(j) + NETVAL
               If haveData = False And NETVAL <> 0 Then
                  haveData = True
               End If
            
               j = j + 1
               Set HeadCf = m_Headers1.Fields(j)
               TempStr = strMinusCod
               Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
               Call m_Details1.AddField(BodyCf)
               If haveData = False And NETVAL <> 0 Then
                  haveData = True
               End If

          '     If mcolParam("SUMMARY_MODE") = "N" And haveData = True Then
                    Call m_Details1.GetString(1, TempStr1, TempStr2)
                    strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
                    Call Vsp.AddTable(strFormat, "", TempStr2)
                    i = i + 1   'นับบรรทัดข้อมูลย่อย
        '       End If
          
          End If   'enableIV
            
                Rs.MoveNext                                                                                            ' วนลูกค้า
         Wend
         
         If Total1(7) <> 0 Then
               Call GenerateFooter("", "      รวม      ", PrevKey1, PrevKey2, Total1, m_Details1, True)      ' เช็ค haveValue ใน Call GenerateFooter **** เซลล์ต้องมีค่าถึงจะแสดงบรรทัดรวม
               Vsp.FontBold = True
               Call m_Details1.GetString(2, TempStr1, TempStr2)
               strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
               Call Vsp.AddTable(strFormat, "", TempStr2)
         End If
         
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
               
               'debug.print temnetval
               genDoc = True
               Exit Function
   
ErrHandler:
   
   mstrErrMsg = "Error(" & RName & ")" & Err.Number & " : " & Err.DESCRIPTION
   Set Rs = Nothing
      Set TempRs = Nothing
End Function


Private Sub GenerateData(tempMinus As Double, tempStcrd As CStcrd, Details1 As CFieldList)
Dim HeadCf As CReportField
Dim BodyCf As CReportField
Dim j As Long
Dim TempStr1 As String
Dim TempStr2 As String
Dim TempStr As String
Dim alngX() As Long
Dim strFormat As String

Call Details1.ClearField
Set BodyCf = New CReportField
   
               Call Details1.ClearField
               j = 0
               haveData = False
            
                j = j + 1
               Set HeadCf = m_Headers1.Fields(j)
               TempStr = i
               Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
               Call Details1.AddField(BodyCf)
   
                j = j + 1
               Set HeadCf = m_Headers1.Fields(j)
               TempStr = tempStcrd.DOCNUM
               Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
               Call Details1.AddField(BodyCf)

                 j = j + 1
               Set HeadCf = m_Headers1.Fields(j)
               TempStr = Left(tempStcrd.CUSNAM, 50)
               Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
               Call Details1.AddField(BodyCf)
      
               j = j + 1
               Set HeadCf = m_Headers1.Fields(j)
               TempStr = tempStcrd.STKDES
               Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
               Call Details1.AddField(BodyCf)

              PayCompleted = True
              RECOD = REsumIV.RCPNUM
               Set tempREdoc = GetREDocDat(m_ReDocdat, REsumIV.RCPNUM, False)
               If Not (tempREdoc Is Nothing) Then
                   CMPLDAT = tempREdoc.DOCDAT
               End If
               
                j = j + 1
               Set HeadCf = m_Headers1.Fields(j)
                 TempStr = FormatNumber(tempStcrd.NETVAL)
               Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
               Call m_Details1.AddField(BodyCf)
               Total3(j) = Total3(j) + tempStcrd.NETVAL
               Total2(j) = Total2(j) + tempStcrd.NETVAL
               Total1(j) = Total1(j) + tempStcrd.NETVAL
               If haveData = False And tempStcrd.NETVAL <> 0 Then
                  haveData = True
               End If
               
               j = j + 1
               Set HeadCf = m_Headers1.Fields(j)
                 TempStr = FormatNumber(tempMinus)
               Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
               Call m_Details1.AddField(BodyCf)
               Total3(j) = Total3(j) + tempMinus
               Total2(j) = Total2(j) + tempMinus
               Total1(j) = Total1(j) + tempMinus
               If haveData = False And tempMinus <> 0 Then
                  haveData = True
               End If
              
               j = j + 1
               Set HeadCf = m_Headers1.Fields(j)
                 TempStr = strNETVAL
               Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
               Call m_Details1.AddField(BodyCf)
               Total3(j) = Total3(j) + NETVAL
               Total2(j) = Total2(j) + NETVAL
               Total1(j) = Total1(j) + NETVAL
               If haveData = False And NETVAL <> 0 Then
                  haveData = True
               End If
            
               Set m_runConditionCommiss = New CConditionCommission
               
               ' คำนวณค่าคอม
               DueCount = DateDiff("D", tempStcrd.DOCDAT, CMPLDAT)  'ต้องคิดตรงนี้เพื่อไปคำนวณค่า
               Set TempCConditionCommiss = GetCheckCommiss(m_ConditionCommiss4, tempStcrd.STKCOD, False)
               If (TempCConditionCommiss Is Nothing) Then   'ไม่เจอในคอเล็คชั่น = ไม่ใช่สินค้าพิเศษ
                             PercentNum1 = (100 / 100)
                             TempStr = ""
               Else
                             PercentNum1 = (TempCConditionCommiss.SLM_PERCENT / 100)   'เป็นสินค้าพิเศษ
                             TempStr = "** "
               End If

               j = j + 1
               Set HeadCf = m_Headers1.Fields(j)
               TempStr = TempStr & FormatNumber(NETVAL * PercentNum1)
               Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
               Call Details1.AddField(BodyCf)
               Total3(j) = Total3(j) + (NETVAL * PercentNum1)
               Total2(j) = Total2(j) + (NETVAL * PercentNum1)
               Total1(j) = Total1(j) + (NETVAL * PercentNum1)
               If haveData = False And (NETVAL * PercentNum1) <> 0 Then
                  haveData = True
              End If

            ' เหมือนคอลัมป์ก่อนหน้านี้ เด๊ะ
               j = j + 1
               Set HeadCf = m_Headers1.Fields(j)
               TempStr = FormatNumber(NETVAL * PercentNum1)
               Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
               Call Details1.AddField(BodyCf)
               Total3(j) = Total3(j) + (NETVAL * PercentNum1)
               Total2(j) = Total2(j) + (NETVAL * PercentNum1)
               Total1(j) = Total1(j) + (NETVAL * PercentNum1)
               If haveData = False And (NETVAL * PercentNum1) <> 0 Then
                           haveData = True
              End If
            
                j = j + 1
               Set HeadCf = m_Headers1.Fields(j)
               TempStr = tempStcrd.DOCDAT                                                                                 ' วันที่ส่งสินค้า
               Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
               Call Details1.AddField(BodyCf)
               
       '        If mcolParam("SUMMARY_MODE") = "N" And haveData = True Then
                    Call Details1.GetString(1, TempStr1, TempStr2)
                    strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
                    Call Vsp.AddTable(strFormat, "", TempStr2)
                    i = i + 1   'นับบรรทัดข้อมูลย่อย
            '   End If

   Set BodyCf = Nothing
End Sub

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

Private Sub GenerateFooter(Txt1 As String, Txt2 As String, Txt3 As String, Txt4 As String, Tot() As Double, Details As CFieldList, HaveValue As Boolean)
Dim HeadCf As CReportField
Dim BodyCf As CReportField
Dim j As Long
Dim TempStr As String

   Call Details.ClearField
   Set BodyCf = New CReportField
   
   HaveValue = False
   
   j = 0
   For Each HeadCf In m_Headers0.Fields
      j = j + 1
      If j = 1 Then
         TempStr = Txt1
      ElseIf j = 2 Then
         TempStr = Txt2
      ElseIf j = 3 Then
         TempStr = Txt3
     ElseIf j = 4 Then
         TempStr = Txt4
     ElseIf j = 8 Then
         TempStr = ""
      ElseIf j = 9 Then
         TempStr = ""
      ElseIf j = 10 Then
         TempStr = ""
      Else
         TempStr = FormatNumber(Tot(j + 1))
      End If
      
      Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.Align, TempStr, HeadCf.BodyAlign)
      Call Details.AddField(BodyCf)
      
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


