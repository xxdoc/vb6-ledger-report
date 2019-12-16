VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CReportCom06"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'User Group Report
Implements CReportInterface

Private Const MODULE_NAME = "CReportCOM06"
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
Private eachSLM As Collection
Private m_Headers0 As CFieldList
Private m_Headers1 As CFieldList
Private m_Headers2 As CFieldList
'Private m_Headers1_1 As CFieldList
'Private m_Headers1_2 As CFieldList
'Private m_Headers1_3 As CFieldList
'Private m_Headers1_4 As CFieldList
Private m_Details1 As CFieldList
'Private m_Details2 As CFieldList
'Private m_StcrdBySale As Collection
'Private m_AreaCod As Collection
'Private m_AreaFromIV As Collection
'Private m_ConditionCommiss1 As Collection
'Private m_ConditionCommiss2 As Collection
'Private m_ConditionCommiss3 As Collection
'Private m_ConditionCommiss4 As Collection

'Private m_cusFromArea(20) As Collection
'Private nonidentifyArea As Boolean
'Dim Stb As CCommissMasterArea
'Private m_MinusStkcod As Collection
'Private m_IVincomplete0 As Collection
'Private m_IVincomplete12 As Collection
'Private m_REsumIV As Collection
'Private m_ReDocdat As Collection
'Private temp_Area As CCommissionCustomerArea
'Private temp_Area2 As CCommissionCustomerArea
'Public REsumIV As CARRcIt
'Private tempComDonStkcod As CComDonStk
'Private m_ComDonStkcod As Collection
'Dim L As Long

Private FirstDate As Date
Private LastDate As Date
Private dayFirst As Date
Private dayLast As Date
Private TempDate As Date

'Private m_SumStcrd As Collection
 Private TempEachSLM As CCondiIncenSum
 Private HaveHeader As Boolean
 Private stcrd_mixdb As Collection
 Private new_mixdb As Collection


'Private havevalue As Boolean

Private Sub Class_Initialize()
   Set new_mixdb = New Collection
   Set mcolParam = New Collection
   Set m_TempCol = New Collection
   Set stcrd_mixdb = New Collection
   Set m_Headers0 = New CFieldList
   Set m_Headers1 = New CFieldList
   Set m_Headers2 = New CFieldList
'   Set m_Headers1_2 = New CFieldList
'   Set m_Headers1_3 = New CFieldList
'   Set m_Headers1_4 = New CFieldList
   Set m_Details1 = New CFieldList
   Set eachSLM = New Collection
   Set TempEachSLM = New CCondiIncenSum
'   Set m_Details2 = New CFieldList
'   Set m_StcrdBySale = New Collection
'   Set m_AreaCod = New Collection
'   Set m_AreaFromIV = New Collection
'   Set m_ConditionCommiss1 = New Collection
'   Set m_ConditionCommiss2 = New Collection
'   Set m_ConditionCommiss3 = New Collection
'   Set m_ConditionCommiss4 = New Collection
'   Set Stb = New CCommissMasterArea
'   Set m_MinusStkcod = New Collection
'   Set m_IVincomplete0 = New Collection
'   Set m_IVincomplete12 = New Collection
'   Set m_REsumIV = New Collection
'   Set m_ReDocdat = New Collection
'   Set temp_Area = New CCommissionCustomerArea
'   Set temp_Area2 = New CCommissionCustomerArea
'   Set REsumIV = New CARRcIt
'   Set m_ComDonStkcod = New Collection
'   Set tempComDonStkcod = New CComDonStk
'   Set m_SumStcrd = New Collection
End Sub

Private Sub Class_Terminate()
   Call ClearParam
   Set mcolParam = Nothing
   Set Vsp = Nothing
   Set m_TempCol = Nothing
   Set stcrd_mixdb = Nothing
   Set m_Headers0 = Nothing
   Set m_Headers1 = Nothing
   Set m_Headers2 = Nothing
   Set m_Details1 = Nothing
   Set eachSLM = Nothing
   Set TempEachSLM = Nothing
   Set new_mixdb = Nothing
'   Set m_Headers1_1 = Nothing
'   Set m_Headers1_2 = Nothing
'   Set m_Headers1_3 = Nothing
'   Set m_Headers1_4 = Nothing
'   Set m_Details2 = Nothing
'   Set m_StcrdBySale = Nothing
'   Set m_AreaCod = Nothing
'   Set m_AreaFromIV = Nothing
'   Set m_ConditionCommiss1 = Nothing
'   Set m_ConditionCommiss2 = Nothing
'   Set m_ConditionCommiss3 = Nothing
'   Set m_ConditionCommiss4 = Nothing
'   Set Stb = Nothing
'   Set m_MinusStkcod = Nothing
'   Set m_IVincomplete0 = Nothing
'   Set m_IVincomplete12 = Nothing
'   Set m_REsumIV = Nothing
'   Set m_ReDocdat = Nothing
'   Set temp_Area = Nothing
'   Set temp_Area2 = Nothing
'   Set REsumIV = Nothing
'   Set m_ComDonStkcod = Nothing
'   Set tempComDonStkcod = Nothing
'   Set m_SumStcrd = Nothing
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
      Call m_Headers2.ClearField
   
   Vsp.FontSize = HEADER_SIZE
   Vsp.FontBold = True
   Vsp.TableBorder = tbAll
   
   Set Cf = New CReportField
      
   Call Cf.SetFieldValue(1, "^", "No.", "^")
   Call m_Headers0.AddField(Cf)
   If mcolParam("NEWCUS_MODE") = "Y" Then
      Call Cf.SetFieldValue(12, "^", "ลูกค้าใหม่", "<")    '27
      Call m_Headers0.AddField(Cf)
   Else
      Call Cf.SetFieldValue(12, "^", "เดือน", "<")    '27
      Call m_Headers0.AddField(Cf)
   End If
   Call Cf.SetFieldValue(2, "^", "จำนวน", "^")
   Call m_Headers0.AddField(Cf)
   Call Cf.SetFieldValue(2, "^", "ราคาต่อ  หน่วย", ">")    '27
   Call m_Headers0.AddField(Cf)
   Call Cf.SetFieldValue(3, "^", "ยอดขาย ไม่ตัดส่วนลด", ">")
   Call m_Headers0.AddField(Cf)
   Call Cf.SetFieldValue(3, "^", "พนักงานขาย", "^")
   Call m_Headers0.AddField(Cf)
   
     Call Cf.SetFieldValue(1, "^", "No.", "^")
     Call m_Headers1.AddField(Cf)
     
          Call Cf.SetFieldValue(1, "^", "รหัสลูกค้า", "^")
     Call m_Headers1.AddField(Cf)
      
     Call Cf.SetFieldValue(3, "^", "ชื่อลูกค้า", "<")
      Call m_Headers1.AddField(Cf)
      
      Call Cf.SetFieldValue(2, "^", "วันที่ส่งสินค้า", "^")
      Call m_Headers1.AddField(Cf)
      
      Call Cf.SetFieldValue(2, "^", "Invoice", "^")
      Call m_Headers1.AddField(Cf)
      
     Call Cf.SetFieldValue(1, "^", "รหัสสินค้า", "^")
      Call m_Headers1.AddField(Cf)
      
      Call Cf.SetFieldValue(3, "^", "ชื่อสินค้า", "<")
      Call m_Headers1.AddField(Cf)
      
      Call Cf.SetFieldValue(2, "^", "จำนวน", "^")
      Call m_Headers1.AddField(Cf)
      
     Call Cf.SetFieldValue(2, "^", "ราคาต่อหน่วย", ">")
      Call m_Headers1.AddField(Cf)

      Call Cf.SetFieldValue(3, "^", "ยอดขาย (ไม่ตัดส่วนลด)", ">")
      Call m_Headers1.AddField(Cf)
      
      Call Cf.SetFieldValue(3, "^", "พนักงานขาย", "^")
      Call m_Headers1.AddField(Cf)
      
   Call Cf.SetFieldValue(1, "^", "", "^")
   Call m_Headers2.AddField(Cf)
   Call Cf.SetFieldValue(22, "<", "", "<")    '27
   Call m_Headers2.AddField(Cf)

    If mcolParam("SUMMARY_MODE") = "Y" Or mcolParam("NEWCUS_MODE") = "Y" Then
      Call m_Headers0.GetString(1, TempStr1, TempStr1_1)
   Else
      Call m_Headers1.GetString(1, TempStr1, TempStr1_1)
   End If
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
Dim RunMonth As Long

   mstrErrMsg = ""
   mblnHeader = True
   mblnNewPage = True
   mblnEndPage = True
   Vsp.PaperSize = pprA4 '
   If mcolParam("SUMMARY_MODE") = "Y" Or mcolParam("NEWCUS_MODE") = "Y" Then
      Vsp.ORIENTATION = orPortrait
   Else
      Vsp.ORIENTATION = orLandscape
   End If
   Vsp.MarginBottom = 300
   Vsp.MarginFooter = 300
   Vsp.MarginHeader = 300
   Vsp.MarginLeft = 300
   Vsp.MarginRight = 500
   Vsp.MarginTop = 300
   Vsp.FontName = "AngsanaUPC"
   Vsp.FontSize = DETAIL_SIZE
         
   RunMonth = Val(mcolParam("RUN_MONTH")) - 1
   Call GetFirstLastDate(DateSerial(mcolParam("YEAR_NO") - 543, mcolParam("MONTH_ID") - RunMonth, 1), FirstDate, TempDate)
 Call GetFirstLastDate(DateSerial(mcolParam("YEAR_NO") - 543, mcolParam("MONTH_ID"), 1), TempDate, LastDate)

   'call LoadSaleCustomerStcrd(m_month1,FirstDate, LastDate, Optional FromCustomerCode As String, Optional ToCustomerCode As String, Optional FromStockCode As String, Optional ToStockCode As String, Optional FromSaleCode As String, Optional ToSaleCode As String)
'   If (RunMonth >= 2) Then          'เดือน  2
'      FirstDate = DateAdd("M", 1, FirstDate)
'      Call GetFirstLastDate(FirstDate, FirstDate, LastDate)
'
'   End If
'
   MODULE_DESC = glbCompanyName & vbCrLf & mcolParam("REPORT_NAME") & _
                             vbCrLf & "เดือน " & EmptyToString(mcolParam("MONTH_NAME"), "N/A") & "   ปี " & EmptyToString(mcolParam("YEAR_NO"), "N/A")
   MODULE_DESC = MODULE_DESC & "     รหัสสินค้า " & EmptyToString(mcolParam("FROM_STK_CODE"), "N/A") & "   ถึง    " & EmptyToString(mcolParam("TO_STK_CODE"), "N/A")

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
Dim PrevKey3 As String
Dim PrevKey4 As String
Dim Stcrd As CStcrd
Dim NewCus As Boolean
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
   Next j
   
   Call StartExportFile(Vsp)
   Vsp.StartDoc
   i = 1  'นับจ.บรรทัดข้อมูลย่อย
   K = 1
    HaveHeader = False
   
   Set BodyCf = New CReportField


        If Year(FirstDate) <> Year(LastDate) Then
            Call GetFirstLastDate(DateSerial(Year(FirstDate), Month(1), 31), , dayLast)  '311254
            Call GetFirstLastDate(DateSerial(Year(LastDate), Month(12), 1), dayFirst)
            
          Set Stcrd = New CStcrd
          Stcrd.FROM_DOC_DATE = FirstDate
          Stcrd.TO_DOC_DATE = dayLast
          Stcrd.FROM_CUSTOMER_CODE = mcolParam("FROM_CUSTOMER_CODE")
          Stcrd.TO_CUSTOMER_CODE = mcolParam("TO_CUSTOMER_CODE")
          Stcrd.FROM_STOCK_CODE = mcolParam("FROM_STK_CODE")
          Stcrd.TO_STOCK_CODE = mcolParam("TO_STK_CODE")
          Call Stcrd.QueryData(14, Rs, iCount, 2) ' ติดต่อกับดาต้าเบสที่ 2
          While Not Rs.EOF
             Call Stcrd.PopulateFromRS(14, Rs)
                Set stcrd_temp = GetObject("CStcrd", stcrd_mixdb, Stcrd.CUSCOD & "-" & Stcrd.DOCNUM & "-" & Stcrd.STKCOD, False)
                If stcrd_temp Is Nothing Then  ' ถ้าไม่มีในคอเล็กชั่น คือ เพิ่มได้
                   Set stcrd_temp = New CStcrd
                   stcrd_temp.CMPLDAT = Stcrd.CMPLDAT
                   stcrd_temp.NETVAL = Stcrd.NETVAL
                   stcrd_temp.DOCNUM = Stcrd.DOCNUM
                   stcrd_temp.CUSNAM = Stcrd.CUSNAM
                   stcrd_temp.STKDES = Stcrd.STKDES
                   stcrd_temp.DOCDAT = Stcrd.DOCDAT
                   stcrd_temp.STKCOD = Stcrd.STKCOD
                   stcrd_temp.SLMCOD = Stcrd.SLMCOD
                   stcrd_temp.CUSCOD = Stcrd.CUSCOD
                   stcrd_temp.UNITPR = Stcrd.UNITPR
                   stcrd_temp.SEQNUM = Stcrd.SEQNUM
                   stcrd_temp.SLMNAM = Stcrd.SLMNAM
                   stcrd_temp.TYPDES = Stcrd.TYPDES
                   stcrd_temp.AREACOD = Stcrd.AREACOD
                   stcrd_temp.AREANAM = Stcrd.AREANAM
                   stcrd_temp.TRNQTY = Stcrd.TRNQTY
                   Call stcrd_mixdb.Add(stcrd_temp, stcrd_temp.CUSCOD & "-" & stcrd_temp.DOCNUM & "-" & stcrd_temp.STKCOD)
                   Set stcrd_temp = Nothing
                End If
             Rs.MoveNext                                                                                            ' วนคน    '  **** วนทุกบรรทัดของ Stcrd
          Wend
                     FirstDate = dayFirst
      End If
      
      Set Stcrd = New CStcrd
      Stcrd.FROM_DOC_DATE = FirstDate
      Stcrd.TO_DOC_DATE = LastDate
      Stcrd.FROM_CUSTOMER_CODE = mcolParam("FROM_CUSTOMER_CODE")
      Stcrd.TO_CUSTOMER_CODE = mcolParam("TO_CUSTOMER_CODE")
      Stcrd.FROM_STOCK_CODE = mcolParam("FROM_STK_CODE")
      Stcrd.TO_STOCK_CODE = mcolParam("TO_STK_CODE")
      Call Stcrd.QueryData(14, Rs, iCount)   ' ติดต่อกับดาต้าเบสที่ 1
      While Not Rs.EOF     '  **** วนทุกบรรทัดของ Stcrd
         Call Stcrd.PopulateFromRS(14, Rs)
            Set stcrd_temp = GetObject("CStcrd", stcrd_mixdb, Stcrd.CUSCOD & "-" & Stcrd.DOCNUM & "-" & Stcrd.STKCOD, False)
            If stcrd_temp Is Nothing Then  ' ถ้าไม่มีในคอเล็กชั่น คือ เพิ่มได้
             Set stcrd_temp = New CStcrd
               stcrd_temp.CMPLDAT = Stcrd.CMPLDAT
               stcrd_temp.NETVAL = Stcrd.NETVAL
               stcrd_temp.DOCNUM = Stcrd.DOCNUM
               stcrd_temp.CUSNAM = Stcrd.CUSNAM
               stcrd_temp.STKDES = Stcrd.STKDES
               stcrd_temp.DOCDAT = Stcrd.DOCDAT
               stcrd_temp.STKCOD = Stcrd.STKCOD
               stcrd_temp.SLMCOD = Stcrd.SLMCOD
               stcrd_temp.CUSCOD = Stcrd.CUSCOD
               stcrd_temp.UNITPR = Stcrd.UNITPR
               stcrd_temp.SEQNUM = Stcrd.SEQNUM
               stcrd_temp.SLMNAM = Stcrd.SLMNAM
               stcrd_temp.TYPDES = Stcrd.TYPDES
               stcrd_temp.AREACOD = Stcrd.AREACOD
               stcrd_temp.AREANAM = Stcrd.AREANAM
               stcrd_temp.TRNQTY = Stcrd.TRNQTY
             Call stcrd_mixdb.Add(stcrd_temp, stcrd_temp.CUSCOD & "-" & stcrd_temp.DOCNUM & "-" & stcrd_temp.STKCOD)
             Set stcrd_temp = Nothing
            End If
         Rs.MoveNext                                                                                            ' วนคน    '  **** วนทุกบรรทัดของ Stcrd
      Wend
      
         Call ArrangeByLoopMin(stcrd_mixdb, new_mixdb)
         Set stcrd_mixdb = Nothing


    Set Stcrd = Nothing
    For Each Stcrd In new_mixdb
        
                 If (PrevKey3 <> Str(Month(Stcrd.DOCDAT)) & "/" & Str(Year(Stcrd.DOCDAT) + 543) Or Stcrd.CUSCOD <> PrevKey1) And PrevKey3 <> "" Then
                   If mcolParam("NEWCUS_MODE") = "N" Then
                      If mcolParam("SUMMARY_MODE") = "N" Then
                           Call GenerateFooter("", "      " & PrevKey3, "", Total2, m_Details1)      ' เช็ค haveValue ใน Call GenerateFooter **** เซลล์ต้องมีค่าถึงจะแสดงบรรทัดรวม
                           Vsp.FontBold = True
                     Else
                           Call GenerateFooter("", "      " & PrevKey3, "", Total2, m_Details1)        ' เช็ค haveValue ใน Call GenerateFooter **** เซลล์ต้องมีค่าถึงจะแสดงบรรทัดรวม
                           Vsp.FontBold = False
                     End If
                     Call m_Details1.GetString(2, TempStr1, TempStr2)
                     strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
                     Call Vsp.AddTable(strFormat, "", TempStr2)
                     Vsp.FontBold = False
                  End If ' new_cus
                     
                     If PrevKey3 <> Str(mcolParam("MONTH_ID")) & "/" & Str(mcolParam("YEAR_NO")) Then
                         NewCus = False           ' ลูกค้าใหม่
                     End If

                        For j = 1 To UBound(Total2)
                           Total2(j) = 0                  ' รวมรายลูกค้า
                        Next j
               End If

'               If PrevKey1 = "15-002" Then
'                  'debug.print
'               End If
               
               If Stcrd.CUSCOD <> PrevKey1 And PrevKey1 <> "" Then
                  If mcolParam("NEWCUS_MODE") = "N" Or (mcolParam("NEWCUS_MODE") = "Y" And NewCus = True) Then  ' ถ้าไม่ติ๊กจะโชว์หมด ถ้าติ๊กจะโชว์เฉพาะลูกค้าใหม่
                       If mcolParam("NEWCUS_MODE") = "Y" Then
                            Call GenerateFooter(Str(K), PrevKey2, PrevKey4, Total1, m_Details1)       ' เช็ค haveValue ใน Call GenerateFooter **** เซลล์ต้องมีค่าถึงจะแสดงบรรทัดรวม
                            Vsp.FontBold = False
                      ElseIf mcolParam("SUMMARY_MODE") = "N" Then
                            Call GenerateFooter("รวม", PrevKey2, PrevKey4, Total1, m_Details1)      ' เช็ค haveValue ใน Call GenerateFooter **** เซลล์ต้องมีค่าถึงจะแสดงบรรทัดรวม
                            Vsp.FontBold = True
                     Else
                            Call GenerateFooter("", PrevKey2, PrevKey4, Total1, m_Details1)       ' เช็ค haveValue ใน Call GenerateFooter **** เซลล์ต้องมีค่าถึงจะแสดงบรรทัดรวม
                            Vsp.FontBold = True
                      End If
                     Call m_Details1.GetString(2, TempStr1, TempStr2)
                     strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
                     Call Vsp.AddTable(strFormat, "", TempStr2)
                     Vsp.FontBold = False
                     K = K + 1
                     HaveHeader = False
                 End If
                     
                        For j = 1 To UBound(Total1)
                           Total1(j) = 0                  ' รวมรายลูกค้า
                        Next j
               End If
               
               'เฮดเดอร์
               If Not (HaveHeader) And mcolParam("NEWCUS_MODE") = "N" Then  '
                     If mcolParam("SUMMARY_MODE") = "N" And mcolParam("NEWCUS_MODE") = "N" Then
                          Call GenerateHeader("", "       ( " & Stcrd.CUSCOD & " )            " & Stcrd.CUSNAM, m_Details1)         ' เช็ค haveValue ใน Call GenerateFooter **** เซลล์ต้องมีค่าถึงจะแสดงบรรทัดรวม
                     Else
                           Call GenerateHeader(Str(K), Stcrd.CUSNAM & "    ( " & Stcrd.CUSCOD & " )", m_Details1)         ' เช็ค haveValue ใน Call GenerateFooter **** เซลล์ต้องมีค่าถึงจะแสดงบรรทัดรวม
                     End If
                  Vsp.FontBold = True
                  Call m_Details1.GetString(1, TempStr1, TempStr2)
                  strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
                  Call Vsp.AddTable(strFormat, "", TempStr2)
                  Vsp.FontBold = False
                  
                  HaveHeader = True
               End If
               
                If Stcrd.CUSCOD <> PrevKey1 Then
                     NewCus = True           ' ลูกค้าใหม่
                End If

               
               Call m_Details1.ClearField
               j = 0
            
                j = j + 1
               Set HeadCf = m_Headers1.Fields(j)
               TempStr = i
               Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
               Call m_Details1.AddField(BodyCf)
                              
                j = j + 1
               Set HeadCf = m_Headers1.Fields(j)
               TempStr = Stcrd.CUSCOD
               Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
               Call m_Details1.AddField(BodyCf)

               j = j + 1
               Set HeadCf = m_Headers1.Fields(j)
               TempStr = Left(Stcrd.CUSNAM, 25)
               Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
               Call m_Details1.AddField(BodyCf)

                j = j + 1
               Set HeadCf = m_Headers1.Fields(j)
               TempStr = DateToStringExtEx2(Stcrd.DOCDAT)
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
               TempStr = Left(Stcrd.STKDES, 25)
               Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
               Call m_Details1.AddField(BodyCf)
      
               j = j + 1
               Set HeadCf = m_Headers1.Fields(j)
               TempStr = Stcrd.TRNQTY
               Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
               Call m_Details1.AddField(BodyCf)
               Total1(j) = Total1(j) + Stcrd.TRNQTY
               Total2(j) = Total2(j) + Stcrd.TRNQTY
            
                j = j + 1
               Set HeadCf = m_Headers1.Fields(j)
               TempStr = FormatNumber(Stcrd.UNITPR)
               Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
               Call m_Details1.AddField(BodyCf)
               Total1(j) = Total1(j) + Stcrd.UNITPR
               Total2(j) = Total2(j) + Stcrd.UNITPR
             
               j = j + 1
               Set HeadCf = m_Headers1.Fields(j)
               TempStr = FormatNumber(Stcrd.NETVAL)
               Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
               Call m_Details1.AddField(BodyCf)
               Total1(j) = Total1(j) + Stcrd.NETVAL
               Total2(j) = Total2(j) + Stcrd.NETVAL
               
               j = j + 1
               Set HeadCf = m_Headers1.Fields(j)
               TempStr = Stcrd.SLMNAM
               Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
               Call m_Details1.AddField(BodyCf)


               If (mcolParam("SUMMARY_MODE") = "N" And mcolParam("NEWCUS_MODE") = "N") Then ' Or (mcolParam("NEWCUS_MODE") = "Y" And NewCus = True)
                  Call m_Details1.GetString(1, TempStr1, TempStr2)
                  strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
                  Call Vsp.AddTable(strFormat, "", TempStr2)
                End If
             
               i = i + 1
               PrevKey1 = Stcrd.CUSCOD
               PrevKey2 = Stcrd.CUSNAM & "   ( " & Stcrd.CUSCOD & " )"
               PrevKey3 = Str(Month(Stcrd.DOCDAT)) & "/" & Str(Year(Stcrd.DOCDAT) + 543)
               PrevKey4 = Left(Stcrd.SLMNAM, 15)

 Next Stcrd
 
            If mcolParam("NEWCUS_MODE") = "N" Then
                If mcolParam("SUMMARY_MODE") = "N" Then
                     Call GenerateFooter("", "      " & PrevKey3, "", Total2, m_Details1)       ' เช็ค haveValue ใน Call GenerateFooter **** เซลล์ต้องมีค่าถึงจะแสดงบรรทัดรวม
                     Vsp.FontBold = True
               Else
                     Call GenerateFooter("", "      " & PrevKey3, "", Total2, m_Details1)        ' เช็ค haveValue ใน Call GenerateFooter **** เซลล์ต้องมีค่าถึงจะแสดงบรรทัดรวม
                     Vsp.FontBold = False
               End If
               Call m_Details1.GetString(2, TempStr1, TempStr2)
               strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
               Call Vsp.AddTable(strFormat, "", TempStr2)
               Vsp.FontBold = False
             End If
             
           If mcolParam("NEWCUS_MODE") = "N" Or (mcolParam("NEWCUS_MODE") = "Y" And NewCus = True) Then
              If mcolParam("NEWCUS_MODE") = "Y" Then
                    Call GenerateFooter(Str(K), PrevKey2, PrevKey4, Total1, m_Details1)       ' เช็ค haveValue ใน Call GenerateFooter **** เซลล์ต้องมีค่าถึงจะแสดงบรรทัดรวม
                    Vsp.FontBold = False
              ElseIf mcolParam("SUMMARY_MODE") = "N" Then
                     Call GenerateFooter("รวม", PrevKey2, PrevKey4, Total1, m_Details1)      ' เช็ค haveValue ใน Call GenerateFooter **** เซลล์ต้องมีค่าถึงจะแสดงบรรทัดรวม
                     Vsp.FontBold = True
              Else
                     Call GenerateFooter("", PrevKey2, PrevKey4, Total1, m_Details1)        ' เช็ค haveValue ใน Call GenerateFooter **** เซลล์ต้องมีค่าถึงจะแสดงบรรทัดรวม
                     Vsp.FontBold = True
              End If
               Call m_Details1.GetString(2, TempStr1, TempStr2)
               strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
               Call Vsp.AddTable(strFormat, "", TempStr2)
               Vsp.FontBold = False
               K = K + 1
               HaveHeader = False
            End If
            
               Vsp.EndDoc
               Call CloseExportFile(Vsp)
               
               If Rs.State = adStateOpen Then
                     Rs.Close
               End If
               Set Rs = Nothing
               
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
   
                Set HeadCf = m_Headers2.Fields(1)
                TempStr = Txt1
                Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
                Call m_Details1.AddField(BodyCf)
    
                Set HeadCf = m_Headers2.Fields(2)
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
   For Each HeadCf In m_Headers0.Fields
      j = j + 1
      If j = 1 Then
         TempStr = Txt1
       ElseIf j = 2 Then
         TempStr = Txt2
      ElseIf j = 3 Then
         TempStr = Tot(8)
      ElseIf j = 4 Then
         TempStr = FormatNumber(Tot(9))
      ElseIf j = 5 Then
         TempStr = FormatNumber(Tot(10))
      Else
         TempStr = Txt3
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

Private Sub ArrangeByLoopMin(TempColl As Collection, OutPutColl As Collection)
'On Error GoTo ErrorHandler
Dim Temp As CStcrd
Dim TempOut As CStcrd
Dim MinKey As String
Dim MaxKey As String
Dim TempCode As String
Dim HasBegin As Boolean
Dim i As Long

   If TempColl.Count <= 0 Then
      Exit Sub
   End If
   
   i = 0
   For Each Temp In TempColl
      
      i = i + 1
      MinKey = ""
      MaxKey = ""
      Temp.KEY = Temp.CUSCOD & "-" & Temp.DOCNUM & "-" & Temp.STKCOD
'      If Temp.Key = "10-007-IV0041865-9103-1" Then
'         'debug.print
'      End If
      If OutPutColl.Count = 0 Then
         Call OutPutColl.Add(Temp, Temp.KEY)
      Else
         For Each TempOut In OutPutColl
            If TempOut.KEY < Temp.KEY Then
               MinKey = TempOut.KEY    ' after
            Else
               MaxKey = TempOut.KEY    ' before
            End If
            If MinKey <> "" And MaxKey <> "" Then
               Exit For
            End If
         Next TempOut
         
         If MinKey <> "" And MaxKey <> "" Then
           Call OutPutColl.Add(Temp, Temp.KEY, MaxKey)
         ElseIf MinKey = "" And MaxKey <> "" Then
            Call OutPutColl.Add(Temp, Temp.KEY, MaxKey)
         ElseIf MinKey <> "" And MaxKey = "" Then
            Call OutPutColl.Add(Temp, Temp.KEY, , MinKey)
         End If
         'Call TempColl.Remove(Temp.Key)
       End If
   Next Temp
   Set Temp = Nothing
'
'   Call OutPutColl.Add(TempColl.ITEM(TempCode))
'   Call TempColl.Remove(TempCode)
'
'   Call ArrangeByAmountMin(TempColl, OutPutColl)
'
'ErrorHandler:
'   'debug.print Err.DESCRIPTION
'   If HasBegin Then
'      m_Conn.RollbackTrans
'   End If
End Sub
