VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CReportAR016"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'User Group Report
Implements CReportInterface

Private Const MODULE_NAME = "CReportAR002"
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
Private m_Left As Double
Private Check_Summary_Flag As String
Private m_YearSeqs As Collection

Private Ar_mixDB As Collection
Private m_CustomerTypes As Collection

Private m_ReceiveAmounts_db1 As Collection
Private m_ReceiveAllAmounts_db1 As Collection
Private m_CnAllAmounts_db1 As Collection

Private m_ReceiveAmounts_db2 As Collection
Private m_ReceiveAllAmounts_db2 As Collection
Private m_CnAllAmounts_db2 As Collection

Private m_ReceiveAmounts_db3 As Collection
Private m_ReceiveAllAmounts_db3 As Collection
Private m_CnAllAmounts_db3 As Collection

Private m_MaxMins As Collection
Private m_ArCustAmounts As Collection
Private m_ArCustPaids As Collection
Private m_DbnCdnByBills As Collection
Private m_RealCreditColl As Collection

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
   Set m_MaxMins = New Collection
   Set m_ArCustAmounts = New Collection
   Set m_ArCustPaids = New Collection
   Set m_DbnCdnByBills = New Collection
   Set m_RealCreditColl = New Collection

   Set m_CnAllAmounts_db1 = New Collection
   Set m_ReceiveAmounts_db1 = New Collection
   Set m_ReceiveAllAmounts_db1 = New Collection
   Set m_CnAllAmounts_db2 = New Collection
   Set m_ReceiveAmounts_db2 = New Collection
   Set m_ReceiveAllAmounts_db2 = New Collection
   Set m_CnAllAmounts_db3 = New Collection
   Set m_ReceiveAmounts_db3 = New Collection
   Set m_ReceiveAllAmounts_db3 = New Collection
      
   Set Ar_mixDB = New Collection
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
   Set m_MaxMins = Nothing
   Set m_DbnCdnByBills = Nothing
   Set m_RealCreditColl = Nothing
   
   Set m_CnAllAmounts_db1 = Nothing
   Set m_ReceiveAmounts_db1 = Nothing
   Set m_ReceiveAllAmounts_db1 = Nothing
   Set m_CnAllAmounts_db2 = Nothing
   Set m_ReceiveAmounts_db2 = Nothing
   Set m_ReceiveAllAmounts_db2 = Nothing
   Set m_CnAllAmounts_db3 = Nothing
   Set m_ReceiveAmounts_db3 = Nothing
   Set m_ReceiveAllAmounts_db3 = Nothing
   Set Ar_mixDB = Nothing
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
   Call m_Headers1_1.ClearField
   
   Set Cf = New CReportField
   
         Call Cf.SetFieldValue(1, "^", "รหัสลูกหนี้", "^")
         Call m_Headers0.AddField(Cf)
         Call Cf.SetFieldValue(2, "^", "ชื่อลูกหนี้", "<")
         Call m_Headers0.AddField(Cf)
         Call Cf.SetFieldValue(5, "^", "หนี้คงค้าง", ">")
         Call m_Headers0.AddField(Cf)

         Call Cf.SetFieldValue(1, "^", "", "^")
         Call m_Headers1.AddField(Cf)
         Call Cf.SetFieldValue(2, "^", "", "<")
         Call m_Headers1.AddField(Cf)
         Call Cf.SetFieldValue(1, "^", EmptyToString(DateToStringExtEx2(mcolParam("TO_PAY_DATE1")), "N/A"), ">")
         Call m_Headers1.AddField(Cf)
         Call Cf.SetFieldValue(1, "^", EmptyToString(DateToStringExtEx2(mcolParam("TO_PAY_DATE2")), "N/A"), ">")
         Call m_Headers1.AddField(Cf)
         Call Cf.SetFieldValue(1, "^", "หนี้ (เพิ่ม)ลด" & vbCrLf & Right(EmptyToString(DateToStringExtEx2(mcolParam("TO_PAY_DATE1")), "N/A"), 4) & "  -  " & Right(EmptyToString(DateToStringExtEx2(mcolParam("TO_PAY_DATE2")), "N/A"), 4), ">")
         Call m_Headers1.AddField(Cf)
         Call Cf.SetFieldValue(1, "^", EmptyToString(DateToStringExtEx2(mcolParam("TO_PAY_DATE3")), "N/A"), ">")
         Call m_Headers1.AddField(Cf)
         Call Cf.SetFieldValue(1, "^", "หนี้ (เพิ่ม)ลด" & vbCrLf & Right(EmptyToString(DateToStringExtEx2(mcolParam("TO_PAY_DATE2")), "N/A"), 4) & "  -  " & Right(EmptyToString(DateToStringExtEx2(mcolParam("TO_PAY_DATE3")), "N/A"), 4), ">")
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
   
'---- db1
   Call LoadReceiveAmountByBill(Nothing, m_ReceiveAmounts_db1, -1, mcolParam("TO_PAY_DATE1"), 1)
   Call LoadReceiveAmountByBill(Nothing, m_ReceiveAllAmounts_db1, -1, -1, 1)
'   Call LoadARCNAmountByBill(Nothing, m_CnAllAmounts_db1, DateAdd("D", 1, -1), -1, -1, -1, 1)        'ต้องบวกยอดยกมาเพิ่ม ยอด CN ทั้งหมดของบิลหลัง CN แต่ เอกสาร LINK เป็นในช่วงยกมา
  Call LoadARCNAmountByBill(Nothing, m_CnAllAmounts_db1, DateAdd("D", 1, mcolParam("TO_PAY_DATE1")), -1, -1, mcolParam("TO_PAY_DATE1"), 1)        'ต้องบวกยอดยกมาเพิ่ม ยอด CN ทั้งหมดของบิลหลัง CN แต่ เอกสาร LINK เป็นในช่วงยกมา
'---- db2
   Call LoadReceiveAmountByBill(Nothing, m_ReceiveAmounts_db2, -1, mcolParam("TO_PAY_DATE2"), 2)
   Call LoadReceiveAmountByBill(Nothing, m_ReceiveAllAmounts_db2, -1, -1, 2)
'   Call LoadARCNAmountByBill(Nothing, m_CnAllAmounts_db2, DateAdd("D", 1, -1), -1, -1, -1, 2)        'ต้องบวกยอดยกมาเพิ่ม ยอด CN ทั้งหมดของบิลหลัง CN แต่ เอกสาร LINK เป็นในช่วงยกมา
  Call LoadARCNAmountByBill(Nothing, m_CnAllAmounts_db2, DateAdd("D", 1, mcolParam("TO_PAY_DATE2")), -1, -1, mcolParam("TO_PAY_DATE2"), 2)        'ต้องบวกยอดยกมาเพิ่ม ยอด CN ทั้งหมดของบิลหลัง CN แต่ เอกสาร LINK เป็นในช่วงยกมา
'---- db3
   Call LoadReceiveAmountByBill(Nothing, m_ReceiveAmounts_db3, -1, mcolParam("TO_PAY_DATE3"), 3)
   Call LoadReceiveAmountByBill(Nothing, m_ReceiveAllAmounts_db3, -1, -1, 3)
 '  Call LoadARCNAmountByBill(Nothing, m_CnAllAmounts_db3, DateAdd("D", 1, -1), -1, -1, -1, 3)        'ต้องบวกยอดยกมาเพิ่ม ยอด CN ทั้งหมดของบิลหลัง CN แต่ เอกสาร LINK เป็นในช่วงยกมา
   Call LoadARCNAmountByBill(Nothing, m_CnAllAmounts_db3, DateAdd("D", 1, mcolParam("TO_PAY_DATE3")), -1, -1, mcolParam("TO_PAY_DATE3"), 3)        'ต้องบวกยอดยกมาเพิ่ม ยอด CN ทั้งหมดของบิลหลัง CN แต่ เอกสาร LINK เป็นในช่วงยกมา
   
'   'debug.print "DB1**  " & m_ReceiveAmounts_db1.Count & "-" & m_ReceiveAllAmounts_db1.Count & "-" & m_CnAllAmounts_db1.Count
'   'debug.print "DB2**  " & m_ReceiveAmounts_db2.Count & "-" & m_ReceiveAllAmounts_db2.Count & "-" & m_CnAllAmounts_db2.Count
'   'debug.print "DB3**  " & m_ReceiveAmounts_db3.Count & "-" & m_ReceiveAllAmounts_db3.Count & "-" & m_CnAllAmounts_db3.Count

   ToPaidDate = "    วันที่ชำระเงิน " & EmptyToString(DateToStringExtEx2(mcolParam("TO_PAY_DATE1")), "N/A")
   ToPaidDate = ToPaidDate & " , " & EmptyToString(DateToStringExtEx2(mcolParam("TO_PAY_DATE2")), "N/A")
    ToPaidDate = ToPaidDate & " , " & EmptyToString(DateToStringExtEx2(mcolParam("TO_PAY_DATE3")), "N/A")

   MODULE_DESC = mcolParam("REPORT_NAME") & _
                                       vbCrLf & ToPaidDate & vbCrLf & " จากลูกค้า " & EmptyToString(mcolParam("FROM_CUSTOMER_CODE"), "N/A") & " ถึงลูกค้า " & EmptyToString(mcolParam("TO_CUSTOMER_CODE"), "N/A") & _
                                       vbCrLf & glbCompanyName
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


Private Sub calSum(ARt As CARTrn, m_Receive As Collection, m_ReceiveAll As Collection, m_CnAll As Collection, Sum As Double, Optional PrevKey1 As String = "", Optional PrevKey4 As String = "")
Dim HeadCf As CReportField
Dim TempStr As String
Dim TempStr1 As String
Dim TempStr2 As String
Dim BodyCf As CReportField
Dim strFormat As String
Dim Ari As CARRcIt
Dim Ari2 As CARRcIt
Dim Apt1 As CARTrn
Dim DueCount As Long
Dim MM As CMaxMin
Dim j As Long
Dim PrintDate As Date
Dim PaidAmount As Double
Dim PaidBalance As Double 'ชำระแล้วยกมา
Dim TempAmt As Double
Dim TempRealCredit As CRealCredit
Dim ARt2 As CARTrn
Dim Width As Double

'   'debug.print "---calsum " & m_Receive.Count & "-" & m_ReceiveAll.Count & "-" & m_CnAll.Count
'   'debug.print "DB1 " & m_ReceiveAmounts_db1.Count & "-" & m_ReceiveAllAmounts_db1.Count & "-" & m_CnAllAmounts_db1.Count
'   'debug.print "DB2 " & m_ReceiveAmounts_db2.Count & "-" & m_ReceiveAllAmounts_db2.Count & "-" & m_CnAllAmounts_db2.Count
'   'debug.print "DB3 " & m_ReceiveAmounts_db3.Count & "-" & m_ReceiveAllAmounts_db3.Count & "-" & m_CnAllAmounts_db3.Count
'   If ARt.CUSCOD = "10-007" Then
'      'debug.print ARt.CUSCOD & ARt.RCVAMT & "-IV-" & ARt.DOCNUM
'      'debug.print
'   End If

   Set Ari = GetARRcpItem(m_Receive, ARt.DOCNUM)
 '  If Not (Ari Is Nothing) Then
      PaidBalance = ARt.RCVAMT
'      If ARt.CUSCOD = "10-007" Then
'         'debug.print ARt.CUSCOD & "--" & ARt.RCVAMT & "--" & m_Receive.Count
'         'debug.print
'       End If
 '  End If
   
   Set Ari2 = GetARRcpItemEx(m_ReceiveAll, ARt.DOCNUM)
   If Not (Ari2 Is Nothing) Then
      PaidBalance = PaidBalance - Ari2.RCVAMT
   End If
   
   Set ARt2 = GetARTrn(m_CnAll, ARt.DOCNUM)
   If Not (ARt2 Is Nothing) Then
      PaidBalance = PaidBalance - ARt2.AMOUNT
   End If
   
   Sum = 0


            TempAmt = (ARt.AMOUNT) - (Ari.RCVAMT + PaidBalance)
            If ARt.RECTYP = 5 Then 'ลดหนี้
               TempAmt = -1 * TempAmt
            End If
            Sum = Sum + TempAmt

End Sub

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
Dim PrevKey4 As String
Dim St As CIsTab
Dim Ar As CARMas
Dim Ar_temp As CARMas
Dim Ar_temp2 As CARMas
Dim Sum1 As Double
Dim Sum2 As Double
Dim Sum3 As Double
Dim SumAll_DB1 As Double
Dim SumAll_DB2 As Double
Dim SumAll_DB3 As Double

Dim Width As Double
Dim ARt As CARTrn
Dim TempRs As ADODB.Recordset
Dim TempARt As CARTrn
Dim TempARRcIt As CARRcIt
Dim HaveHeader As Boolean
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
      Set BodyCf = New CReportField

      Set Ar = New CARMas
      Ar.FROM_CUSTOMER_CODE = mcolParam("FROM_CUSTOMER_CODE")
      Ar.TO_CUSTOMER_CODE = mcolParam("TO_CUSTOMER_CODE")
'      Ar.OrderBy = mcolParam("ORDER_TYPE")
'      Ar.OrderType = mcolParam("ORDER_BY")
      Call glbDaily.QueryCustomer(Ar, Rs, iCount, IsOK, glbErrorLog, 1)  ' DB1
       While Not Rs.EOF
         Call Ar.PopulateFromRS(1, Rs)
            Set Ar_temp = GetObject("CARMas", Ar_mixDB, Trim(Ar.CUSCOD), False)
            If Ar_temp Is Nothing Then  ' ถ้าไม่มีในคอเล็กชั่น คือ เพิ่มได้
               Set Ar_temp = New CARMas
               Ar_temp.CUSCOD = Ar.CUSCOD
               Ar_temp.CUSNAM = Ar.CUSNAM
               Call Ar_mixDB.Add(Ar_temp, Trim(Ar_temp.CUSCOD))
               Set Ar_temp = Nothing
            End If
         Rs.MoveNext
      Wend
      Set Ar = Nothing
      
      Set Ar = New CARMas
      Ar.FROM_CUSTOMER_CODE = mcolParam("FROM_CUSTOMER_CODE")
      Ar.TO_CUSTOMER_CODE = mcolParam("TO_CUSTOMER_CODE")
     Call glbDaily.QueryCustomer(Ar, Rs, iCount, IsOK, glbErrorLog, 2)  ' DB2
     While Not Rs.EOF
     Call Ar.PopulateFromRS(1, Rs)
     Set Ar_temp = GetObject("CARMas", Ar_mixDB, Trim(Ar.CUSCOD), False)
           If Ar_temp Is Nothing Then  ' ถ้าไม่มีในคอเล็กชั่น คือ เพิ่มได้
               Set Ar_temp = New CARMas
               Ar_temp.CUSCOD = Ar.CUSCOD
               Ar_temp.CUSNAM = Ar.CUSNAM
               Call Ar_mixDB.Add(Ar_temp, Trim(Ar_temp.CUSCOD))
               Set Ar_temp = Nothing
            End If
         Rs.MoveNext
      Wend
      Set Ar = Nothing
      
      Set Ar = New CARMas
      Ar.FROM_CUSTOMER_CODE = mcolParam("FROM_CUSTOMER_CODE")
      Ar.TO_CUSTOMER_CODE = mcolParam("TO_CUSTOMER_CODE")
      Call glbDaily.QueryCustomer(Ar, Rs, iCount, IsOK, glbErrorLog, 3)  ' DB3
      While Not Rs.EOF
      Call Ar.PopulateFromRS(1, Rs)
            Set Ar_temp = GetObject("CARMas", Ar_mixDB, Trim(Ar.CUSCOD), False)
            If Ar_temp Is Nothing Then  ' ถ้าไม่มีในคอเล็กชั่น คือ เพิ่มได้
               Set Ar_temp = New CARMas
               Ar_temp.CUSCOD = Ar.CUSCOD
               Ar_temp.CUSNAM = Ar.CUSNAM
               Call Ar_mixDB.Add(Ar_temp, Trim(Ar_temp.CUSCOD))
               Set Ar_temp = Nothing
            End If
         Rs.MoveNext
      Wend
      
   For Each Ar_temp In Ar_mixDB
   '   While Not Rs.EOF
'         For j = 1 To UBound(Total1)
'            Total1(j) = 0
'         Next j



'         HaveHeader = False
         If True Then  'มียอดค้างจ่าย
            '===Head
            '=== Detail
            Set ARt = New CARTrn
            ARt.TO_DOC_DATE = mcolParam("TO_PAY_DATE1")
            ARt.CUSCOD = Ar_temp.CUSCOD
            ARt.RECTYP = ""
            ARt.RecTypeSet = "('3', '4', '5')"
            ARt.OrderBy = 1
            ARt.OrderType = 1
            Call ARt.QueryData(1, TempRs, iCount, True, 1)
            While Not TempRs.EOF
               i = i + 1
        '       Call m_Details1.ClearField
               Call ARt.PopulateFromRS(1, TempRs)
               
               ' คำนวณ  ผลจาก DB1
               'Call GenerateDetail(alngX(), i, ARt, Total1, Total2, HaveHeader, PrevKey1, PrevKey4)
               Call calSum(ARt, m_ReceiveAmounts_db1, m_ReceiveAllAmounts_db1, m_CnAllAmounts_db1, Sum1)     ' ได้ผลหักแต่ละบิลของ db1 มา
                SumAll_DB1 = SumAll_DB1 + Sum1
'               เก็บใส่ตัว
               TempRs.MoveNext
            Wend

               Set Ar_temp2 = GetObject("CARMas", Ar_mixDB, Trim(ARt.CUSCOD), False)
               If Not (Ar_temp2 Is Nothing) Then
                        Ar_temp2.SUM_DB1 = SumAll_DB1
                        Set Ar_temp2 = Nothing
                        SumAll_DB1 = 0
               End If
               
               Set ARt = Nothing
                        Set ARt = New CARTrn
            ARt.TO_DOC_DATE = mcolParam("TO_PAY_DATE2")
            ARt.CUSCOD = Ar_temp.CUSCOD
            ARt.RECTYP = ""
            ARt.RecTypeSet = "('3', '4', '5')"
            ARt.OrderBy = 1
            ARt.OrderType = 1
            Call ARt.QueryData(1, TempRs, iCount, True, 2)
            While Not TempRs.EOF
               i = i + 1
         '      Call m_Details1.ClearField
               Call ARt.PopulateFromRS(1, TempRs)
               ' คำนวณ  ผลจาก DB1
               'Call GenerateDetail(alngX(), i, ARt, Total1, Total2, HaveHeader, PrevKey1, PrevKey4)
               Call calSum(ARt, m_ReceiveAmounts_db2, m_ReceiveAllAmounts_db2, m_CnAllAmounts_db2, Sum2)     ' ได้ผลหักแต่ละบิลของ db1 มา
               SumAll_DB2 = SumAll_DB2 + Sum2
'               เก็บใส่ตัว
               TempRs.MoveNext
            Wend

               Set Ar_temp2 = GetObject("CARMas", Ar_mixDB, Trim(ARt.CUSCOD), False)
               If Not (Ar_temp2 Is Nothing) Then
                        Ar_temp2.SUM_DB2 = SumAll_DB2
                        Set Ar_temp2 = Nothing
                        SumAll_DB2 = 0
               End If
            
               Set ARt = Nothing
            Set ARt = New CARTrn
           ARt.TO_DOC_DATE = mcolParam("TO_PAY_DATE3")
            ARt.CUSCOD = Ar_temp.CUSCOD
            ARt.RECTYP = ""
            ARt.RecTypeSet = "('3', '4', '5')"
            ARt.OrderBy = 1
            ARt.OrderType = 1
            Call ARt.QueryData(1, TempRs, iCount, True, 3)
            While Not TempRs.EOF
      '         Call m_Details1.ClearField
               Call ARt.PopulateFromRS(1, TempRs)
               ' คำนวณ  ผลจาก DB1
               'Call GenerateDetail(alngX(), i, ARt, Total1, Total2, HaveHeader, PrevKey1, PrevKey4)
               Call calSum(ARt, m_ReceiveAmounts_db3, m_ReceiveAllAmounts_db3, m_CnAllAmounts_db3, Sum3)     ' ได้ผลหักแต่ละบิลของ db1 มา
               SumAll_DB3 = SumAll_DB3 + Sum3
'               เก็บใส่ตัว
               TempRs.MoveNext
            Wend

               Set Ar_temp2 = GetObject("CARMas", Ar_mixDB, Trim(ARt.CUSCOD), False)
               If Not (Ar_temp2 Is Nothing) Then
                        Ar_temp2.SUM_DB3 = SumAll_DB3
                        Set Ar_temp2 = Nothing
                        SumAll_DB3 = 0
               End If
               
               
              Set ARt = Nothing
          End If
         Next Ar_temp

          For Each Ar_temp In Ar_mixDB
               Call m_Details1.ClearField
               
               j = 1
               Set HeadCf = m_Headers1.Fields(j)
               TempStr = Ar_temp.CUSCOD
               Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
               Call m_Details1.AddField(BodyCf)
               
                j = j + 1
               Set HeadCf = m_Headers1.Fields(j)
               TempStr = Ar_temp.CUSNAM
               Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
               Call m_Details1.AddField(BodyCf)
               
               j = j + 1
               Set HeadCf = m_Headers1.Fields(j)
               TempStr = FormatNumber(Ar_temp.SUM_DB1)
               Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
               Call m_Details1.AddField(BodyCf)
               Total1(j) = Total1(j) + Ar_temp.SUM_DB1
               
                j = j + 1
               Set HeadCf = m_Headers1.Fields(j)
               TempStr = FormatNumber(Ar_temp.SUM_DB2)
               Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
               Call m_Details1.AddField(BodyCf)
               Total1(j) = Total1(j) + Ar_temp.SUM_DB2
               
                j = j + 1
               Set HeadCf = m_Headers1.Fields(j)
               TempStr = FormatNumber(Ar_temp.SUM_DB1 - Ar_temp.SUM_DB2)
               Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
               Call m_Details1.AddField(BodyCf)
               Total1(j) = Total1(j) + (Ar_temp.SUM_DB1 - Ar_temp.SUM_DB2)
               
                j = j + 1
               Set HeadCf = m_Headers1.Fields(j)
               TempStr = FormatNumber(Ar_temp.SUM_DB3)
               Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
               Call m_Details1.AddField(BodyCf)
               Total1(j) = Total1(j) + (Ar_temp.SUM_DB3)
               
                j = j + 1
               Set HeadCf = m_Headers1.Fields(j)
               TempStr = FormatNumber(Ar_temp.SUM_DB2 - Ar_temp.SUM_DB3)
               Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
               Call m_Details1.AddField(BodyCf)
               Total1(j) = Total1(j) + (Ar_temp.SUM_DB2 - Ar_temp.SUM_DB3)
               
            If Ar_temp.SUM_DB1 <> 0 Or Ar_temp.SUM_DB2 <> 0 Or Ar_temp.SUM_DB3 <> 0 Then
               Call m_Details1.GetString(1, TempStr1, TempStr2)
               strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
               Call Vsp.AddTable(strFormat, "", TempStr2)
               Vsp.FontBold = False
            End If
               
         Next Ar_temp

              Call m_Details1.ClearField

               j = 1
               Set HeadCf = m_Headers1.Fields(j)
               TempStr = ""
               Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
               Call m_Details1.AddField(BodyCf)
               
                j = j + 1
               Set HeadCf = m_Headers1.Fields(j)
               TempStr = "รวม"
               Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
               Call m_Details1.AddField(BodyCf)
               
               j = j + 1
               Set HeadCf = m_Headers1.Fields(j)
               TempStr = FormatNumber(Total1(j))
               Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
               Call m_Details1.AddField(BodyCf)
               
                j = j + 1
               Set HeadCf = m_Headers1.Fields(j)
               TempStr = FormatNumber(Total1(j))
               Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
               Call m_Details1.AddField(BodyCf)
               
                j = j + 1
               Set HeadCf = m_Headers1.Fields(j)
               TempStr = FormatNumber(Total1(j))
               Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
               Call m_Details1.AddField(BodyCf)
               
                j = j + 1
               Set HeadCf = m_Headers1.Fields(j)
               TempStr = FormatNumber(Total1(j))
               Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
               Call m_Details1.AddField(BodyCf)
               
                j = j + 1
               Set HeadCf = m_Headers1.Fields(j)
               TempStr = FormatNumber(Total1(j))
               Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
               Call m_Details1.AddField(BodyCf)
               
               Call m_Details1.GetString(1, TempStr1, TempStr2)
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

Private Sub GenerateFooter(Txt1 As String, Txt2 As String, Txt3 As String, Tot() As Double, Details As CFieldList, DateCount As Long)
Dim HeadCf As CReportField
Dim BodyCf As CReportField
Dim j As Long
Dim TempStr As String

   Call Details.ClearField
   Set BodyCf = New CReportField
   
   j = 0
   For Each HeadCf In m_Headers1_1.Fields
      j = j + 1
      If j = 1 Then
         TempStr = Txt1
      ElseIf j = 2 Then
         TempStr = Txt2
      ElseIf j = 3 Then
         TempStr = Txt3
      ElseIf j = (m_Headers1_1.Fields.Count) Or (j = (m_Headers1_1.Fields.Count - 1) And (mcolParam("REAL_CREDIT_FLAG") = "Y" Or mcolParam("NINETY_FLAG") = "Y" Or Val(mcolParam("CREDIT")) > 0)) Then
         TempStr = ""
      Else
         TempStr = FormatNumber(Tot(j + 1), "-")
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
Private Sub GenerateEnd()
   If mcolParam("REAL_CREDIT_FLAG") = "Y" Then
      Vsp = "***ใช้เครดิตตามตารางเครดิตจริง"
   End If
   If Val(mcolParam("CREDIT")) > 0 Then
      Vsp = "***เครดิต " & Val(mcolParam("CREDIT")) & " วัน"
   ElseIf mcolParam("NINETY_FLAG") = "Y" Then
      Vsp = "***เครดิต 90 วัน(ถ้าไม่มีในตารางจริง)"
   End If
End Sub


