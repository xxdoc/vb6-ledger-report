VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CReportAP006"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'User Group Report
Implements CReportInterface

Private Const MODULE_NAME = "CReportAP002"
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

Private m_SupplierTypes As Collection
Private m_StockGroups As Collection
Private m_PaidAmounts As Collection
Private m_PaidAllAmounts As Collection
Private m_MaxMins As Collection
Private m_ApSupAmounts As Collection
Private m_ApSupPaids As Collection
Private m_DbnCdnByBills As Collection

Private HaveHeader As Boolean

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
   Set m_SupplierTypes = New Collection
   Set m_PaidAmounts = New Collection
   Set m_MaxMins = New Collection
   Set m_ApSupAmounts = New Collection
   Set m_ApSupPaids = New Collection
   Set m_PaidAllAmounts = New Collection
   Set m_DbnCdnByBills = New Collection
   Set m_StockGroups = New Collection
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
   Set m_SupplierTypes = Nothing
   Set m_PaidAmounts = Nothing
   Set m_MaxMins = Nothing
   Set m_PaidAllAmounts = Nothing
   Set m_DbnCdnByBills = Nothing
   Set m_StockGroups = Nothing
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
Dim SumHeader As Double

   tbdT = Vsp.TableBorder
   blnBold = Vsp.FontBold
   iSize = Vsp.FontSize
   
   Vsp.FontSize = HEADER_SIZE
   Vsp.FontBold = True
   Vsp.TableBorder = tbAll
   
   Call m_Headers0.ClearField
   Call m_Headers1.ClearField
   
   Set Cf = New CReportField
   
   Call Cf.SetFieldValue(1, "^", "N", "^")
   Call m_Headers0.AddField(Cf)
   Call Cf.SetFieldValue(6, "^", "รหัสประเภทผู้ขาย", "^")
   Call m_Headers0.AddField(Cf)
   
   
   SumHeader = 0
   '===
   Call Cf.SetFieldValue(1, "^", "NO.", "^")
   Call m_Headers1.AddField(Cf)
   Call Cf.SetFieldValue(2, "^", "รหัสผู้ขาย", "^")
   Call m_Headers1.AddField(Cf)
   Call Cf.SetFieldValue(4, "^", "ชื่อผู้ขาย", "<")
   Call m_Headers1.AddField(Cf)
   Call Cf.SetFieldValue(MakeHeader(2), "^", "เลขที่เอกสาร", "^")
   Call m_Headers1.AddField(Cf)
   SumHeader = SumHeader + MakeHeader(2)
   
   Call Cf.SetFieldValue(MakeHeader(2), "^", "วันที่", "^")
   Call m_Headers1.AddField(Cf)
   SumHeader = SumHeader + MakeHeader(2)
   
   Call Cf.SetFieldValue(MakeHeader(2), "^", "รหัส R/M", "^")
   Call m_Headers1.AddField(Cf)
   SumHeader = SumHeader + MakeHeader(2)
   
   Call Cf.SetFieldValue(MakeHeader(3), "^", "R/M", "<")
   Call m_Headers1.AddField(Cf)
   SumHeader = SumHeader + MakeHeader(3)
   
    
   Call Cf.SetFieldValue(2, "^", "นน", ">")
   Call m_Headers1.AddField(Cf)
   SumHeader = SumHeader + 2
   Call Cf.SetFieldValue(2, "^", "จำนวนเงิน", ">")
   Call m_Headers1.AddField(Cf)
   SumHeader = SumHeader + 2
   Call Cf.SetFieldValue(2, "^", "ราคา", ">")
   Call m_Headers1.AddField(Cf)
   SumHeader = SumHeader + 2
   
   Call Cf.SetFieldValue(MakeHeader(2), "^", "คลัง", "^")
   Call m_Headers1.AddField(Cf)
   SumHeader = SumHeader + MakeHeader(2)
   
   Call Cf.SetFieldValue(SumHeader, "^", "ประเภทผู้ขาย", "<")
   Call m_Headers0.AddField(Cf)
   
   SumHeader = 0
   
   Call Cf.SetFieldValue(MakeHeader(2), "^", "วันที่ชำระ", "^")
   Call m_Headers1.AddField(Cf)
   SumHeader = SumHeader + MakeHeader(2)
   
   Call Cf.SetFieldValue(MakeHeader(2), "^", "เลขที่เช็ค", "^")
   Call m_Headers1.AddField(Cf)
   SumHeader = SumHeader + MakeHeader(2)
   
   Call Cf.SetFieldValue(MakeHeader(2), "^", "เช็คลงวันที่", "^")
   Call m_Headers1.AddField(Cf)
   SumHeader = SumHeader + MakeHeader(2)
   
   Call Cf.SetFieldValue(2, "^", "จำนวนเงิน", ">")
   Call m_Headers1.AddField(Cf)
   SumHeader = SumHeader + 2
   
   If mcolParam("SHOW_SUM_PAID") = "Y" Then
        Call Cf.SetFieldValue(2, "^", "ยอดหนี้", ">")
        Call m_Headers1.AddField(Cf)
        SumHeader = SumHeader + 2
   Else
        Call Cf.SetFieldValue(0, "^", "ยอดหนี้", ">")
        Call m_Headers1.AddField(Cf)
        SumHeader = SumHeader + 0
   End If
   
   Call Cf.SetFieldValue(SumHeader, "^", "รายละเอียดชำระ", "<")
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
   
   Call LoadSupplierType(Nothing, m_SupplierTypes)
   
   'Call LoadPaidAmountByBill(Nothing, m_PaidAmounts, -1, mcolParam("TO_PAY_DATE"))
   'Call LoadPaidAmountByBill(Nothing, m_PaidAllAmounts, -1, -1)
   'Call LoadDueDateInterval1(Nothing, m_MaxMins)
   'Call LoadAPAmountBySup(Nothing, m_ApSupAmounts, -1, mcolParam("TO_DOC_DATE"))       'ยอดซื้อ ถึงวันที่
   'Call LoadPaidAmountBySup(Nothing, m_ApSupPaids, -1, mcolParam("TO_PAY_DATE"))             'ยอดจ่ายถึงวันที่
   'Call LoadDbnCdnByBill(Nothing, m_DbnCdnByBills, -1, mcolParam("TO_PAY_DATE"))
   FromDate = "จากวันที่ " & EmptyToString(DateToStringExtEx2(mcolParam("FROM_DOC_DATE")), "N/A")
   ToDate = " ถึง วันที่ " & EmptyToString(DateToStringExtEx2(mcolParam("TO_DOC_DATE")), "N/A")
   ToPaidDate = "    วันที่ชำระเงินถึงวันที่ " & EmptyToString(DateToStringExtEx2(mcolParam("TO_PAY_DATE")), "N/A")
   
   'glbEnterPrise.ENTERPRISE_NAME & " , " & glbEnterPrise.BRANCH_NAME
   MODULE_DESC = glbCompanyName & vbCrLf & mcolParam("REPORT_NAME") & _
                                       vbCrLf & FromDate & ToDate & "  " & ToPaidDate & _
                                       " จากรหัส RM " & EmptyToString(mcolParam("FROM_STOCK_CODE"), "N/A") & " ถึงรหัส RM " & EmptyToString(mcolParam("TO_STOCK_CODE"), "N/A")
   
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

Private Sub GenerateDetail(alngX() As Long, i As Long, Apt As CApTrn, Tot1() As Double, Tot2() As Double, Ap As CAPMas)
Dim HeadCf As CReportField
Dim TempStr As String
Dim TempStr1 As String
Dim TempStr2 As String
Dim BodyCf As CReportField
Dim strFormat As String
Dim Ari As CAPRcIt
Dim Ari2 As CAPRcIt
Dim Apt1 As CApTrn
Dim DueCount As Long
Dim MM As CMaxMin
Dim j As Long
Dim Sum As Double
Dim PrintDate As Date
Dim PaidAmount As Double
Dim PaidBalance As Double 'ชำระแล้วยกมา
Dim Width As Double
   
   Set BodyCf = New CReportField
   
   If mcolParam("PRINT_DATE") <= 0 Then
      PrintDate = Now
   Else
      PrintDate = mcolParam("PRINT_DATE")
   End If
   
   DueCount = DateDiff("D", Apt.DUEDAT, PrintDate)
   Set Ari = GetAPRcpItem(m_PaidAmounts, Apt.DOCNUM)
   Set Ari2 = GetAPRcpItemEx(m_PaidAllAmounts, Apt.DOCNUM)
   
'   Set Apt1 = GetAPTrn(m_DbnCdnByBills, Apt.DocNum)
   
   'ชำระแล้วยกมาเท่ากับ APT.PAYAMT - ยอดชำระทั้งหมดของใบสำคัญจ่าย
   PaidBalance = Apt.PAYAMT - Ari2.PAYAMT
   If Apt.AMOUNT > (Ari.PAYAMT + PaidBalance) Then  'จ่ายยังไม่ครบ
      If Apt.RECTYP = 5 Then 'ลดหนี้
         'จากการทดลองดูจาก flag CMPLAPP = 'Y' จะคูณด้วย -1 ถ้าเป็น 'N' จะไม่ต้องคูณ
         If Apt.CMPLAPP = "Y" Then
            'Apt.AMOUNT = -1 * Apt.AMOUNT
         ElseIf Apt.CMPLAPP = "N" Then
            'ไม่ต้องทำอะไร
         End If
      End If
      
      Set HeadCf = m_Headers1.Fields(1)
      TempStr = i
      Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
      Call m_Details1.AddField(BodyCf)
      
      Set HeadCf = m_Headers1.Fields(2)
      TempStr = Apt.SUPCOD
      Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
      Call m_Details1.AddField(BodyCf)

      Set HeadCf = m_Headers1.Fields(3)
      'TempStr = Format(Year(Apt.DUEDAT), "00") & "/" & Format(Month(Apt.DUEDAT), "00") & "/" & Format(Day(Apt.DUEDAT), "00")
      TempStr = DateToStringExtEx2(Apt.DUEDAT)
      Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
      Call m_Details1.AddField(BodyCf)
      
      Set HeadCf = m_Headers1.Fields(4)
      TempStr = Apt.DOCNUM
      Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
      Call m_Details1.AddField(BodyCf)

      Set HeadCf = m_Headers1.Fields(5)
      'TempStr = Format(Year(Apt.DOCDAT), "00") & "/" & Format(Month(Apt.DOCDAT), "00") & "/" & Format(Day(Apt.DOCDAT), "00")
      TempStr = DateToStringExtEx2(Apt.DOCDAT)
      Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
      Call m_Details1.AddField(BodyCf)
           
      If i = 163 Then
         'debug.print
      End If
      
      j = 6
      Sum = 0
      For Each MM In m_MaxMins
         'Set HeadCf = m_Headers1.Fields(j)
         If (MM.MIN < DueCount) And (DueCount <= MM.MAX) Then
            TempStr = FormatNumber(Apt.AMOUNT - (Ari.PAYAMT + PaidBalance), "-")
            
            'Tot1(j) = Tot1(j) + (Apt.AMOUNT - (Ari.PAYAMT + PaidBalance))
            'Tot2(j) = Tot2(j) + (Apt.AMOUNT - (Ari.PAYAMT + PaidBalance))
            If Apt.RECTYP = 5 Then 'ลดหนี้
               Sum = Sum + ((Ari.PAYAMT + PaidBalance) - Apt.AMOUNT)
            Else
               Sum = Sum + (Apt.AMOUNT - (Ari.PAYAMT + PaidBalance))
            End If
            
         Else
            TempStr = FormatNumber(0, "-")
         End If
'         Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
'         Call m_Details1.AddField(BodyCf)
                  
         'j = j + 1
      Next MM
            
      Set HeadCf = m_Headers1.Fields(j)
      TempStr = FormatNumber(Sum, "-")
      If Apt.RECTYP = 5 Then 'ลดหนี้
         Tot1(j) = Tot1(j) + ((Ari.PAYAMT + PaidBalance) - Apt.AMOUNT)
         Tot2(j) = Tot2(j) + ((Ari.PAYAMT + PaidBalance) - Apt.AMOUNT)
      Else
         Tot1(j) = Tot1(j) + (Apt.AMOUNT - (Ari.PAYAMT + PaidBalance))
         Tot2(j) = Tot2(j) + (Apt.AMOUNT - (Ari.PAYAMT + PaidBalance))
      End If
      Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
      Call m_Details1.AddField(BodyCf)
      
'      j = j + 1
'      Set HeadCf = m_Headers1.Fields(j)
'      TempStr = FormatNumberInt(DueCount)
'      Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
'      Call m_Details1.AddField(BodyCf)
      
      
      '----------------------------------------------------
      If Not (HaveHeader) Then
         Call m_Details2.ClearField
         
         Width = m_Headers1.Fields(1).Width + m_Headers1.Fields(2).Width
         TempStr = Ap.SUPCOD
         Call BodyCf.SetFieldValue(Width, "^", TempStr)
         Call m_Details2.AddField(BodyCf)

         Width = m_Headers1.Fields(3).Width + m_Headers1.Fields(4).Width + m_Headers1.Fields(5).Width + m_Headers1.Fields(6).Width
         TempStr = Ap.SUPNAM
         Call BodyCf.SetFieldValue(Width, "<", TempStr)
         Call m_Details2.AddField(BodyCf)
         
         Call m_Details2.GetString(1, TempStr1, TempStr2)
         strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
         Call Vsp.AddTable(strFormat, "", TempStr2)
         
         HaveHeader = True
      End If
      '----------------------------------------------------
      
      Call m_Details1.GetString(1, TempStr1, TempStr2)
      strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
      Call Vsp.AddTable(strFormat, "", TempStr2)
   Else
      i = i - 1
   End If
   Set BodyCf = Nothing
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
Dim PrevKey2 As String
Dim PrevKey3 As String
Dim PrevKey4 As String
Dim PrevKey5 As String
Dim PrevKey6 As String
Dim PrevKey7 As String

Dim Stc As CStcrd

Dim St As CIsTab
Dim Ap As CAPMas
Dim Width As Double
Dim Apt As CApTrn
Dim TempRs As ADODB.Recordset
Dim TempApt As CApTrn
Dim TempApRcIt As CAPRcIt
Dim HaveValue As Boolean
Dim Balance As Double
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
   
   mblnHeader = False
   Call StartExportFile(Vsp)
   Vsp.StartDoc
   Call printHeader
   mblnHeader = True
   
      i = 0
      Set BodyCf = New CReportField
      
      Set Stc = New CStcrd
      Stc.FROM_DOC_DATE = mcolParam("FROM_DOC_DATE")
      Stc.TO_DOC_DATE = mcolParam("TO_DOC_DATE")
      Stc.FROM_DUE_DATE = mcolParam("FROM_DUE_DATE")
      Stc.TO_DUE_DATE = mcolParam("TO_DUE_DATE")
      Stc.FROM_SUPPLIER_CODE = mcolParam("FROM_SUPPLIER_CODE")
      Stc.TO_SUPPLIER_CODE = mcolParam("TO_SUPPLIER_CODE")
      Stc.FROM_STOCK_CODE = mcolParam("FROM_STOCK_CODE")
      Stc.TO_STOCK_CODE = mcolParam("TO_STOCK_CODE")
      If mcolParam("SUPPLIER_TYPE") > 0 Then
         Set St = GetItemFromListIndex(m_SupplierTypes, mcolParam("SUPPLIER_TYPE"))
         Stc.SUPTYP = St.TYPCOD
      ElseIf Len(mcolParam("SUPPLIER_TYPE_SET")) > 0 Then
         Stc.SUPTYPSET = mcolParam("SUPPLIER_TYPE_SET")
      Else
         Stc.SUPTYP = ""
      End If
      
      Call Stc.QueryData(1, Rs, iCount)
      
      i = 0
      While Not Rs.EOF
         
         Call Stc.PopulateFromRS(1, Rs)
      
      
         If (PrevKey2 <> Stc.SUPCOD) Or (PrevKey4 <> Trim(Format(Month(Stc.DOCDAT), "00") & "/" & Format(Year(Stc.DOCDAT) + 543, "0000"))) Then
            If i > 0 Then
               If mcolParam("NO_SHOW_BILL") = "Y" Then
                    Call GenerateFooter("", PrevKey4, PrevKey5, Total3, m_Details1, 0)
               Else
                    Call GenerateFooter("รวม", PrevKey4, "", Total3, m_Details1, 0)
               End If
               Call m_Details1.GetString(2, TempStr1, TempStr2)
               strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
               If mcolParam("NO_SHOW_BILL") = "N" Then
                    Vsp.FontBold = True
               End If
               Call Vsp.AddTable(strFormat, "", TempStr2)
               Vsp.FontBold = False
               
               For j = 1 To UBound(Total1)
                  Total3(j) = 0
               Next j
            End If
            
            If i > 0 And (PrevKey2 <> Stc.SUPCOD) Then
               Call GenerateFooter("รวม", PrevKey2, "", Total1, m_Details1, 0)
               Call m_Details1.GetString(2, TempStr1, TempStr2)
               strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
               Vsp.FontBold = True
               Call Vsp.AddTable(strFormat, "", TempStr2)
               Vsp.FontBold = False
               
               For j = 1 To UBound(Total1)
                  Total1(j) = 0
               Next j
               
               Balance = 0
            End If
         End If
         
         PrevKey1 = Stc.SUPTYP
         PrevKey2 = Stc.SUPCOD
         PrevKey4 = Trim(Format(Month(Stc.DOCDAT), "00") & "/" & Format(Year(Stc.DOCDAT) + 543, "0000"))
         PrevKey5 = Stc.SUPNAM
         
         Call m_Details1.ClearField
         i = i + 1
         
         If i = 343 Then
            Debug.Print
         End If
         
         Set HeadCf = m_Headers1.Fields(1)
         If PrevKey3 <> Stc.DOCNUM Then
            TempStr = i
         Else
            TempStr = ""
         End If
         Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
         Call m_Details1.AddField(BodyCf)
         
         Set HeadCf = m_Headers1.Fields(2)
         If PrevKey3 <> Stc.DOCNUM Then
            TempStr = Stc.SUPCOD
         Else
            TempStr = ""
         End If
         Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
        Call m_Details1.AddField(BodyCf)
        
         Set HeadCf = m_Headers1.Fields(3)
         If PrevKey3 <> Stc.DOCNUM Then
            TempStr = Left(Stc.SUPNAM, 25)
         Else
            TempStr = ""
         End If
        Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
        Call m_Details1.AddField(BodyCf)
         
         Set HeadCf = m_Headers1.Fields(4)
         If PrevKey3 <> Stc.DOCNUM Then
            TempStr = Stc.DOCNUM
         Else
            TempStr = ""
         End If
        Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
        Call m_Details1.AddField(BodyCf)
        
         Set HeadCf = m_Headers1.Fields(5)
         If PrevKey3 <> Stc.DOCNUM Then
            TempStr = DateToStringExtEx2(Stc.DOCDAT)
        Else
            TempStr = ""
        End If
        Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
        Call m_Details1.AddField(BodyCf)
         
         Set HeadCf = m_Headers1.Fields(6)
         'If PrevKey3 <> Stc.DOCNUM Then
            TempStr = Stc.STKCOD
         'Else
         '   TempStr = ""
         'End If
        Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
        Call m_Details1.AddField(BodyCf)
         
         Set HeadCf = m_Headers1.Fields(7)
         'If PrevKey3 <> Stc.DOCNUM Then
            TempStr = Left(Stc.STKDES, 15)
         'Else
         '   TempStr = ""
         'End If
        Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
        Call m_Details1.AddField(BodyCf)
         
         Set HeadCf = m_Headers1.Fields(8)
         If PrevKey7 <> Trim(Stc.DOCNUM & "-" & Stc.STKCOD & "-" & Stc.TRNQTY & "-" & Stc.NETVAL & "-" & Stc.SEQNUM) Then
            If Val(Stc.POSOPR) = 7 Then  'ลดหนี้
               TempStr = FormatNumber(Stc.TRNQTY * -1)
               Total3(8) = Total3(8) - Stc.TRNQTY
               Total1(8) = Total1(8) - Stc.TRNQTY
               Total2(8) = Total2(8) - Stc.TRNQTY
            Else
               TempStr = FormatNumber(Stc.TRNQTY)
               Total3(8) = Total3(8) + Stc.TRNQTY
               Total1(8) = Total1(8) + Stc.TRNQTY
               Total2(8) = Total2(8) + Stc.TRNQTY
            End If
         Else
            TempStr = ""
         End If
        Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
        Call m_Details1.AddField(BodyCf)
        
         Set HeadCf = m_Headers1.Fields(9)
         If PrevKey7 <> Trim(Stc.DOCNUM & "-" & Stc.STKCOD & "-" & Stc.TRNQTY & "-" & Stc.NETVAL & "-" & Stc.SEQNUM) Then
         If Val(Stc.POSOPR) = 7 Then  'ลดหนี้
            TempStr = FormatNumber(Stc.NETVAL * -1)
            Total3(9) = Total3(9) - Stc.NETVAL
            Total1(9) = Total1(9) - Stc.NETVAL
            Total2(9) = Total2(9) - Stc.NETVAL
            Balance = Balance - Stc.NETVAL
         Else
            TempStr = FormatNumber(Stc.NETVAL)
            Total3(9) = Total3(9) + Stc.NETVAL
            Total1(9) = Total1(9) + Stc.NETVAL
            Total2(9) = Total2(9) + Stc.NETVAL
            Balance = Balance + Stc.NETVAL
         End If
         Else
            TempStr = ""
         End If
        Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
        Call m_Details1.AddField(BodyCf)
         
         Set HeadCf = m_Headers1.Fields(10)
         If PrevKey7 <> Trim(Stc.DOCNUM & "-" & Stc.STKCOD & "-" & Stc.TRNQTY & "-" & Stc.NETVAL & "-" & Stc.SEQNUM) Then
            TempStr = FormatNumber(Stc.UNITPR)
         Else
            TempStr = ""
         End If
        Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
        Call m_Details1.AddField(BodyCf)
         
         Set HeadCf = m_Headers1.Fields(11)
         If PrevKey7 <> Trim(Stc.DOCNUM & "-" & Stc.STKCOD & "-" & Stc.TRNQTY & "-" & Stc.NETVAL & "-" & Stc.SEQNUM) Then
            TempStr = Stc.LOCCOD
         Else
            TempStr = ""
         End If
        Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
        Call m_Details1.AddField(BodyCf)
         
         Set HeadCf = m_Headers1.Fields(12)
         If (PrevKey3 = Stc.DOCNUM And PrevKey6 = Stc.CHQNUM) Then
            TempStr = ""
         Else
            TempStr = DateToStringExtEx2(Stc.GETDAT)
         End If
         Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
         Call m_Details1.AddField(BodyCf)
        
         Set HeadCf = m_Headers1.Fields(13)
         If (PrevKey3 = Stc.DOCNUM And PrevKey6 = Stc.CHQNUM) Then
            TempStr = ""
         Else
            TempStr = Stc.CHQNUM
         End If
         Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
         Call m_Details1.AddField(BodyCf)
        
         Set HeadCf = m_Headers1.Fields(14)
         If (PrevKey3 = Stc.DOCNUM And PrevKey6 = Stc.CHQNUM) Then
            TempStr = ""
         Else
            TempStr = DateToStringExtEx2(Stc.CHQDAT)
         End If
         Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
         Call m_Details1.AddField(BodyCf)
        
         Set HeadCf = m_Headers1.Fields(15)
         If (PrevKey3 = Stc.DOCNUM And PrevKey6 = Stc.CHQNUM) Then
            TempStr = ""
         Else
            TempStr = FormatNumber(Stc.PAYAMT)
            Total3(15) = Total3(15) + Stc.PAYAMT
            Total1(15) = Total1(15) + Stc.PAYAMT
            Total2(15) = Total2(15) + Stc.PAYAMT
            Balance = Balance - Stc.PAYAMT
         End If
         Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
         Call m_Details1.AddField(BodyCf)
                        
        Set HeadCf = m_Headers1.Fields(16)
         TempStr = FormatNumber(Balance)
         Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
         Call m_Details1.AddField(BodyCf)
         
         If mcolParam("NO_SHOW_BILL") = "N" Then
            Call m_Details1.GetString(1, TempStr1, TempStr2)
            strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
            Call Vsp.AddTable(strFormat, "", TempStr2)
         End If
         
         PrevKey3 = Stc.DOCNUM
         PrevKey6 = Stc.CHQNUM
         PrevKey7 = Trim(Stc.DOCNUM & "-" & Stc.STKCOD & "-" & Stc.TRNQTY & "-" & Stc.NETVAL & "-" & Stc.SEQNUM)
         
         Rs.MoveNext
      Wend
      
      If mcolParam("NO_SHOW_BILL") = "Y" Then
        Call GenerateFooter("", PrevKey4, PrevKey5, Total3, m_Details1, 0)
      Else
        Call GenerateFooter("รวม", PrevKey4, "", Total3, m_Details1, 0)
      End If
      Call m_Details1.GetString(2, TempStr1, TempStr2)
      strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
      If mcolParam("NO_SHOW_BILL") = "N" Then
             Vsp.FontBold = True
        End If
      Call Vsp.AddTable(strFormat, "", TempStr2)
      Vsp.FontBold = False
      
      Call GenerateFooter("รวม", PrevKey2, "", Total1, m_Details1, 0)
      Call m_Details1.GetString(2, TempStr1, TempStr2)
      strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
      Vsp.FontBold = True
      Call Vsp.AddTable(strFormat, "", TempStr2)
      Vsp.FontBold = False
      
      Call GenerateFooter("สุทธิ", "", "", Total2, m_Details1, 0)
      Call m_Details1.GetString(2, TempStr1, TempStr2)
      strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
      Vsp.FontBold = True
      Call Vsp.AddTable(strFormat, "", TempStr2)
      Vsp.FontBold = False

      Set Ap = Nothing
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
   For Each HeadCf In m_Headers1.Fields
      j = j + 1
      If j = 1 Then
         TempStr = Txt1
      ElseIf j = 2 Then
         TempStr = Txt2
      ElseIf j = 3 Then
         TempStr = Txt3
      ElseIf j = 8 Or j = 9 Or j = 15 Then
         TempStr = FormatNumber(Tot(j))
      ElseIf j = 10 Then
         TempStr = FormatNumber(MyDiff(Tot(j - 1), Tot(j - 2)))
      ElseIf j = 16 Then
        TempStr = FormatNumber(Tot(9) - Tot(15))
      Else
         TempStr = ""
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
Public Function MakeHeader(DataWidth As Double) As Double
    If mcolParam("NO_SHOW_BILL") = "N" Then
         MakeHeader = DataWidth
    Else
         MakeHeader = 0
    End If
End Function



