VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CReportAR007"
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

   tbdT = Vsp.TableBorder
   blnBold = Vsp.FontBold
   iSize = Vsp.FontSize
   
   Vsp.FontSize = HEADER_SIZE
   Vsp.FontBold = True
   Vsp.TableBorder = tbAll
   
   Call m_Headers0.ClearField
   Call m_Headers1.ClearField
   
   Set Cf = New CReportField
      
   Call Cf.SetFieldValue(4, "^", "รหัสลูกค้า", "^")
   Call m_Headers0.AddField(Cf)
   Call Cf.SetFieldValue(6, "^", "ลูกค้า", "<")
   Call m_Headers0.AddField(Cf)
   
   Call Cf.SetFieldValue(1, "^", "NO.", "^")
   Call m_Headers1.AddField(Cf)
   Call Cf.SetFieldValue(3, "^", "วันที่", "^")
   Call m_Headers1.AddField(Cf)
   Call Cf.SetFieldValue(3, "^", "เลขที่", "<")
   Call m_Headers1.AddField(Cf)
   Call Cf.SetFieldValue(3, "^", "จำนวนเงิน", ">")
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
      
   FirstDate = mcolParam("FROM_DOC_DATE")
   LastDate = mcolParam("TO_DOC_DATE")
   
   mstrErrMsg = ""
   mblnHeader = True
   mblnNewPage = True
   mblnEndPage = True
   Vsp.PaperSize = pprA4 '
   Vsp.ORIENTATION = orPortrait
   Vsp.MarginBottom = 700
   Vsp.MarginFooter = 700
   Vsp.MarginHeader = 700
   Vsp.MarginLeft = 700
   Vsp.MarginRight = 700
   Vsp.MarginTop = 700
   Vsp.FontName = "AngsanaUPC"
   Vsp.FontSize = DETAIL_SIZE
   
   Call LoadCustomerType(Nothing, m_CustomerTypes)
   
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
                                       vbCrLf & FromDate & "   " & ToDate & " " & " จากลูกค้า " & EmptyToString(mcolParam("FROM_CUSTOMER_CODE"), "N/A") & " ถึงลูกค้า " & EmptyToString(mcolParam("TO_CUSTOMER_CODE"), "N/A")
   
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
Dim ARt As CARTrn
Dim Width As Double

Dim TempRs As ADODB.Recordset
Dim Ar As COESLM
Dim Total4(100) As Double
Dim PrevKey3 As String
Dim PrevKey4 As String
Dim TypeFont As Boolean


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
   
      Set BodyCf = New CReportField
      i = 0
      
      Set Ar = New COESLM
      Ar.FROM_SALE_CODE = mcolParam("FROM_SALE_CODE")
      Ar.TO_SALE_CODE = mcolParam("TO_SALE_CODE")
      Call glbDaily.QuerySale(Ar, TempRs, iCount, IsOK, glbErrorLog)

              If mcolParam("SUMMARY_MODE") = "N" Then
          TypeFont = True
          Else:
                TypeFont = False
        End If
      
       While Not TempRs.EOF
      
                 For j = 1 To UBound(Total4)
                    Total4(j) = 0
                 Next j
        
                 Call m_Details1.ClearField
                 Call Ar.PopulateFromRS(1, TempRs)
                        
                Set BodyCf = New CReportField
                Width = m_Headers0.Fields(1).Width
                TempStr = Ar.SLMCOD                                          ' แสดงหัวบรรทัด sale
                Call BodyCf.SetFieldValue(Width, "^", TempStr)
                Call m_Details1.AddField(BodyCf)
    
                Width = m_Headers0.Fields(2).Width
                TempStr = Ar.SLMNAM
                Call BodyCf.SetFieldValue(Width, "<", TempStr)
                Call m_Details1.AddField(BodyCf)
    
                Call m_Details1.GetString(1, TempStr1, TempStr2)
                 strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
                    Vsp.FontBold = True
                    Call Vsp.AddTable(strFormat, "", TempStr2)
                    Vsp.FontBold = False
                   
                PrevKey3 = Ar.SLMCOD
                PrevKey4 = Ar.SLMNAM
                 
      
                    Set ARt = New CARTrn
                    ARt.SLMCOD = Ar.SLMCOD
                    ARt.FROM_DOC_DATE = mcolParam("FROM_DOC_DATE")
                    ARt.TO_DOC_DATE = mcolParam("TO_DOC_DATE")
                    ARt.FROM_CUSTOMER_CODE = mcolParam("FROM_CUSTOMER_CODE")
                    ARt.TO_CUSTOMER_CODE = mcolParam("TO_CUSTOMER_CODE")
                    ARt.OrderBy = 4
                    ARt.OrderType = 1
                    ARt.QueryFlag = -1
                    ARt.RecTypeSet = "('3','4')"
                    Call glbDaily.QueryARTran(ARt, Rs, iCount, IsOK, glbErrorLog)
                                                                                            
                     While Not Rs.EOF
                                i = i + 1
                                Call m_Details1.ClearField
                                Call ARt.PopulateFromRS(1, Rs)
                                
                                If PrevKey1 <> ARt.CUSCOD And i <> 1 Then
                                     If Total1(K) <> 0 Then
                                             Call GenerateFooter("รวม", "ยอดขายงวด " & PrevKey2, "", Total1, m_Details1)                    'รวม ยอดขาย ,,,,, total1
                                            Call m_Details1.GetString(2, TempStr1, TempStr2)
                                            strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
                                            Vsp.FontBold = TypeFont
                                            Call Vsp.AddTable(strFormat, "", TempStr2)
                                            Vsp.FontBold = False
                                      End If
                                    
                                        If Total4(K) <> 0 Then          'รวมของ sale คนนี้ ไม่เท่ากับ 0 ให้แสดงรวมลูกค้า
                                                 Call GenerateFooter("รวม", "ลูกค้า" & PrevKey1, "", Total2, m_Details1)                             'รวม ลูกค้า ,,,,, total2
                                                 Call m_Details1.GetString(2, TempStr1, TempStr2)
                                                 strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
                                                 Vsp.FontBold = True
                                                 Call Vsp.AddTable(strFormat, "", TempStr2)
                                                 Vsp.FontBold = False
                                              End If
                                   
                                        For j = 1 To UBound(Total1)
                                           Total1(j) = 0
                                           Total2(j) = 0
                                        Next j
                                        
                                        Call GenerateHeader(ARt.CUSCOD, ARt.CUSNAM, m_Details1)           ' เฮดเดอร์ แสดง ชื่อลูกค้า
                                        Call m_Details1.GetString(2, TempStr1, TempStr2)
                                        strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
                                        Vsp.FontBold = True
                                        Call Vsp.AddTable(strFormat, "", TempStr2)
                                        Vsp.FontBold = False
                                   
                                ElseIf (PrevKey2 <> (Month(ARt.DOCDAT) & "/" & (Year(ARt.DOCDAT) + 543))) And i <> 1 Then
                                     If Total1(K) <> 0 Then
                                            Call GenerateFooter("รวม", "ยอดขายงวด " & PrevKey2, "", Total1, m_Details1)                 'รวม ยอดขาย ,,,,, total1
                                            Call m_Details1.GetString(2, TempStr1, TempStr2)
                                            strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
                                            Vsp.FontBold = TypeFont
                                            Call Vsp.AddTable(strFormat, "", TempStr2)
                                            Vsp.FontBold = False
                                        End If
                                   
                                        For j = 1 To UBound(Total1)
                                           Total1(j) = 0
                                        Next j
                                   
                                ElseIf i = 1 Then
                                        Call GenerateHeader(ARt.CUSCOD, ARt.CUSNAM, m_Details1)            ' เฮดเดอร์ แสดง ชื่อลูกค้า
                                        Call m_Details1.GetString(2, TempStr1, TempStr2)
                                        strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
                                        Vsp.FontBold = True
                                        Call Vsp.AddTable(strFormat, "", TempStr2)
                                        Vsp.FontBold = False
                                End If
                                
                                PrevKey1 = ARt.CUSCOD
                                PrevKey2 = (Month(ARt.DOCDAT) & "/" & (Year(ARt.DOCDAT) + 543))
                                Call m_Details1.ClearField
                                '===Head
                                Set HeadCf = m_Headers1.Fields(1)
                                TempStr = i
                                Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
                                Call m_Details1.AddField(BodyCf)
                                
                                K = 1
                                K = K + 1
                                Set HeadCf = m_Headers1.Fields(K)
                                TempStr = DateToStringExtEx2(ARt.DOCDAT)
                                Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
                                Call m_Details1.AddField(BodyCf)
                                
                                K = K + 1
                                Set HeadCf = m_Headers1.Fields(K)
                                TempStr = ARt.DOCNUM
                                Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
                                Call m_Details1.AddField(BodyCf)
                                
                                K = K + 1
                                Set HeadCf = m_Headers1.Fields(K)
                                TempStr = FormatNumber(ARt.AMOUNT, "")
                                Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
                                Call m_Details1.AddField(BodyCf)
                                
                                Total1(K) = Total1(K) + ARt.AMOUNT
                                Total2(K) = Total2(K) + ARt.AMOUNT
                                Total3(K) = Total3(K) + ARt.AMOUNT
                                Total4(K) = Total4(K) + ARt.AMOUNT
                                
                                If mcolParam("SUMMARY_MODE") = "N" Then
                                        Call m_Details1.GetString(1, TempStr1, TempStr2)
                                        strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
                                        Call Vsp.AddTable(strFormat, "", TempStr2)
                                 End If
                                           
                                Rs.MoveNext
                             Wend
        
        If Total1(K) <> 0 Then
            Call GenerateFooter("รวม", "ยอดขายงวด " & PrevKey2, "", Total1, m_Details1)                'รวม ยอดขาย ,,,,, total1
            Call m_Details1.GetString(2, TempStr1, TempStr2)
            strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
            Vsp.FontBold = TypeFont
            Call Vsp.AddTable(strFormat, "", TempStr2)
            Vsp.FontBold = False
        End If
        
       If (Total4(K) <> 0) Then
            Call GenerateFooter("รวม", "ลูกค้า " & PrevKey1, "", Total2, m_Details1)                  ' รวม ลูกค้า ,,,,,,,,, totral2
            Call m_Details1.GetString(2, TempStr1, TempStr2)
            strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
            Vsp.FontBold = True
            Call Vsp.AddTable(strFormat, "", TempStr2)
            Vsp.FontBold = False
       End If
        
      Set BodyCf = Nothing
      Set ARt = Nothing
                                
             Call GenerateFooter("รวม", PrevKey3 & "  " & PrevKey4, "", Total4, m_Details1)                '
             Call m_Details1.GetString(2, TempStr1, TempStr2)
             strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
             Vsp.FontBold = True
             Call Vsp.AddTable(strFormat, "", TempStr2)
             Vsp.FontBold = False
             
        For j = 1 To UBound(Total3)
           Total1(j) = 0         ' tt3 ของลูกค้า 005_3
           Total2(j) = 0         ' tt3 ของลูกค้า 005_3
           Total4(j) = 0         ' tt3 ของลูกค้า 005_3
        Next j

         TempRs.MoveNext
      Wend


      Call GenerateFooter("สุทธิ", "", "", Total3, m_Details1)                             ' รวมสุทธิ ,,,,,,,,,, 3
      Call m_Details1.GetString(2, TempStr1, TempStr2)
      strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
      Vsp.FontBold = True
      Call Vsp.AddTable(strFormat, "", TempStr2)
      Vsp.FontBold = False
            
        Vsp = ""
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


