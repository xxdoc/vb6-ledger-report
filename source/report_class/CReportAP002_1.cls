VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CReportAP002_1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'User Group Report
Implements CReportInterface

Private Const MODULE_NAME = "CReportAP002_1"
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
Private m_YearSeqs As Collection

Private m_SupplierTypes As Collection
Private m_PaidAmounts As Collection
Private m_PaidAllAmounts As Collection
Private m_MaxMins As Collection
Private m_ApSupAmounts As Collection
Private m_ApSupPaids As Collection
Private m_DbnCdnByBills As Collection
Private m_ApAllAmounts  As Collection
Private SupplierGroupColl As Collection
Private Supplier As Collection
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
   Set m_ApAllAmounts = New Collection
   Set SupplierGroupColl = New Collection
   Set Supplier = New Collection
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
   Set m_ApAllAmounts = Nothing
   Set SupplierGroupColl = Nothing
   Set Supplier = Nothing
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
      
   '===
   Call Cf.SetFieldValue(1, "^", "NO.", "^")
   Call m_Headers0.AddField(Cf)
   Call Cf.SetFieldValue(3, "^", "", "^")
   Call m_Headers0.AddField(Cf)
   Call Cf.SetFieldValue(3, "^", "", "^")
   Call m_Headers0.AddField(Cf)
   Call Cf.SetFieldValue(3, "^", "", "^")
   Call m_Headers0.AddField(Cf)
   Call Cf.SetFieldValue(6, "^", "จะครบกำหนด", ">")
   Call m_Headers0.AddField(Cf)
   Call Cf.SetFieldValue(8, "^", "เกินกำหนด", ">")
   Call m_Headers0.AddField(Cf)
   Call Cf.SetFieldValue(2, "^", "", "^")
   Call m_Headers0.AddField(Cf)
   Call Cf.SetFieldValue(2, "^", "เกินกำหนด", "^")
   Call m_Headers0.AddField(Cf)
   Call Cf.SetFieldValue(2, "^", "DIFF วัน", "^")
   Call m_Headers0.AddField(Cf)
   '===
   
   Call Cf.SetFieldValue(1, "^", "NO.", "^")
   Call m_Headers1.AddField(Cf)
   Call Cf.SetFieldValue(3, "^", "เลขที่ใบรับของ", "^")
   Call m_Headers1.AddField(Cf)
   Call Cf.SetFieldValue(3, "^", "วันที่รับของ", "^")
   Call m_Headers1.AddField(Cf)
   Call Cf.SetFieldValue(3, "^", "วันครบกำหนด", "^")
   Call m_Headers1.AddField(Cf)
   
   Call Cf.SetFieldValue(2, "^", "> 60 วัน", ">")
   Call m_Headers1.AddField(Cf)
   Call Cf.SetFieldValue(2, "^", "30-60 วัน", ">")
   Call m_Headers1.AddField(Cf)
   Call Cf.SetFieldValue(2, "^", "< 30 วัน", ">")
   Call m_Headers1.AddField(Cf)
   
   Call Cf.SetFieldValue(2, "^", "1-15 วัน", ">")
   Call m_Headers1.AddField(Cf)
   Call Cf.SetFieldValue(2, "^", "16-30 วัน", ">")
   Call m_Headers1.AddField(Cf)
   Call Cf.SetFieldValue(2, "^", "31-60 วัน", ">")
   Call m_Headers1.AddField(Cf)
   Call Cf.SetFieldValue(2, "^", "> 60 วัน", ">")
   Call m_Headers1.AddField(Cf)
   
   Call Cf.SetFieldValue(2, "^", "รวมหนี้ค้าง", ">")
   Call m_Headers1.AddField(Cf)
   
   Call Cf.SetFieldValue(2, "^", "เกินกำหนด", "^")
   Call m_Headers1.AddField(Cf)
   
   Call Cf.SetFieldValue(2, "^", "รับของ", "^")
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

   Call LoadSupplierType(Nothing, m_SupplierTypes)
   Call LoadPaidAmountByBill(Nothing, m_PaidAmounts, -1, mcolParam("TO_PAY_DATE"))
   Call LoadPaidAmountByBill(Nothing, m_PaidAllAmounts, -1, -1)
   Call LoadDueDateInterval1(Nothing, m_MaxMins)
   Call LoadAPAmountBySup(Nothing, m_ApSupAmounts, -1, mcolParam("TO_DOC_DATE"))
   Call LoadPaidAmountBySup(Nothing, m_ApSupPaids, -1, mcolParam("TO_PAY_DATE"))
   Call LoadPaidAmountBySup(Nothing, m_ApAllAmounts, -1, -1)         'ยอดตัดบิลทั้งหมดของบิลภายในวันที่กำหนด
   
   If mcolParam("DATA_TYPE_ID") > 0 Then
      Call LoadSupplierGroup(Nothing, SupplierGroupColl, , , mcolParam("DATA_TYPE_ID"), mcolParam("GROUP_TYPE_CODE"))
      If Not mcolParam("GROUP_TYPE_CODE") > 0 Then
         Call LoadSupplier(Nothing, Supplier)
         Call AddCollection(SupplierGroupColl, Supplier)
      End If
   End If
   
'   Call LoadDbnCdnByBill(Nothing, m_DbnCdnByBills, -1, mcolParam("TO_PAY_DATE"))
   FromDate = "จากวันที่ " & EmptyToString(DateToStringExtEx2(mcolParam("FROM_DOC_DATE")), "N/A")
   ToDate = " ถึง วันที่ " & EmptyToString(DateToStringExtEx2(mcolParam("TO_DOC_DATE")), "N/A")
   ToPaidDate = "    วันที่ชำระเงินถึงวันที่ " & EmptyToString(DateToStringExtEx2(mcolParam("TO_PAY_DATE")), "N/A") & "    วันที่พิมพ์ " & EmptyToString(DateToStringExtEx2(mcolParam("PRINT_DATE")), "N/A")

   'glbEnterPrise.ENTERPRISE_NAME & " , " & glbEnterPrise.BRANCH_NAME
   MODULE_DESC = glbCompanyName & vbCrLf & mcolParam("REPORT_NAME") & _
                                       vbCrLf & FromDate & ToDate & "  " & ToPaidDate
   
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

Private Sub GenerateDetail(alngX() As Long, i As Long, Apt As CApTrn, Tot1() As Double, Tot2() As Double, Tot3() As Double)
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
Dim RevCount As Long
Dim MM As CMaxMin
Dim j As Long
Dim Sum As Double
Dim PrintDate As Date
Dim PaidAmount As Double
Dim PaidBalance As Double 'ชำระแล้วยกมา
   
   Set BodyCf = New CReportField
   
   If mcolParam("PRINT_DATE") <= 0 Then
      PrintDate = Now
   Else
      PrintDate = mcolParam("PRINT_DATE")
   End If
   
   DueCount = DateDiff("D", Apt.DUEDAT, PrintDate)
   
   RevCount = DateDiff("D", Apt.DOCDAT, PrintDate)
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
                           'Apt.AMOUNT = -1 * Apt.AMOUNT
            'ไม่ต้องทำอะไร
         End If
      End If

      Set HeadCf = m_Headers1.Fields(1)
      TempStr = i
      Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
      Call m_Details1.AddField(BodyCf)
      
      Set HeadCf = m_Headers1.Fields(2)
      TempStr = Apt.DOCNUM
      Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
      Call m_Details1.AddField(BodyCf)

      Set HeadCf = m_Headers1.Fields(3)
      TempStr = DateToStringExtEx2(Apt.DOCDAT)
      Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
      Call m_Details1.AddField(BodyCf)
      
      Set HeadCf = m_Headers1.Fields(4)
      TempStr = DateToStringExtEx2(Apt.DUEDAT)
      Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
      Call m_Details1.AddField(BodyCf)
      
      j = 5
      Sum = 0
      For Each MM In m_MaxMins
         Set HeadCf = m_Headers1.Fields(j)
         If (MM.MIN < DueCount) And (DueCount <= MM.MAX) Then
            'TempStr = FormatNumber(Apt.AMOUNT - (Ari.PAYAMT + PaidBalance), "-")
            If Apt.RECTYP = 5 Then 'ลดหนี้
               TempStr = FormatNumber((Ari.PAYAMT + PaidBalance) - Apt.AMOUNT, "-")
            Else
               TempStr = FormatNumber(Apt.AMOUNT - (Ari.PAYAMT + PaidBalance), "-")
            End If
            
            If Apt.RECTYP = 5 Then 'ลดหนี้
               Tot1(j) = Tot1(j) + ((Ari.PAYAMT + PaidBalance) - Apt.AMOUNT)
               Tot2(j) = Tot2(j) + ((Ari.PAYAMT + PaidBalance) - Apt.AMOUNT)
               Tot3(j) = Tot3(j) + ((Ari.PAYAMT + PaidBalance) - Apt.AMOUNT)
               Sum = Sum + ((Ari.PAYAMT + PaidBalance) - Apt.AMOUNT)
            Else
               Tot1(j) = Tot1(j) + (Apt.AMOUNT - (Ari.PAYAMT + PaidBalance))
               Tot2(j) = Tot2(j) + (Apt.AMOUNT - (Ari.PAYAMT + PaidBalance))
               Tot3(j) = Tot3(j) + (Apt.AMOUNT - (Ari.PAYAMT + PaidBalance))
               Sum = Sum + (Apt.AMOUNT - (Ari.PAYAMT + PaidBalance))
            End If
         Else
            TempStr = FormatNumber(0, "-")
         End If
         Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
         Call m_Details1.AddField(BodyCf)
                  
         j = j + 1
      Next MM
            
      Set HeadCf = m_Headers1.Fields(j)
      TempStr = FormatNumber(Sum, "-")
      If Apt.RECTYP = 5 Then 'ลดหนี้
         Tot1(j) = Tot1(j) + ((Ari.PAYAMT + PaidBalance) - Apt.AMOUNT)
         Tot2(j) = Tot2(j) + ((Ari.PAYAMT + PaidBalance) - Apt.AMOUNT)
         Tot3(j) = Tot3(j) + ((Ari.PAYAMT + PaidBalance) - Apt.AMOUNT)
      Else
         Tot1(j) = Tot1(j) + (Apt.AMOUNT - (Ari.PAYAMT + PaidBalance))
         Tot2(j) = Tot2(j) + (Apt.AMOUNT - (Ari.PAYAMT + PaidBalance))
         Tot3(j) = Tot3(j) + (Apt.AMOUNT - (Ari.PAYAMT + PaidBalance))
      End If
      
      Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
      Call m_Details1.AddField(BodyCf)

      j = j + 1
      Set HeadCf = m_Headers1.Fields(j)
      TempStr = FormatNumberInt(DueCount)
      Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
      Call m_Details1.AddField(BodyCf)

      j = j + 1
      Set HeadCf = m_Headers1.Fields(j)
      TempStr = FormatNumberInt(RevCount)
      Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
      Call m_Details1.AddField(BodyCf)


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
Dim PrevKey2 As Double
Dim St As CIsTab
Dim Sg As CSupplierGroup
Dim Ap As CAPMas
Dim Width As Double
Dim Apt As CApTrn
Dim TempRs As ADODB.Recordset
Dim TempApt As CApTrn
Dim TempApRcIt As CAPRcIt
Dim TempApEx As CAPRcIt

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

      Set Ap = New CAPMas
      
      If mcolParam("SHOW_EXCEL") = "Y" Then
         Call printHeader
      End If
      
      For Each Sg In SupplierGroupColl
         
         If (PrevKey1 <> Sg.GROUP_TYPE_NAME) Then
            If i > 0 Then
               Call GenerateFooter("รวม", "" & PrevKey1, "", Total3, m_Details1, 0)
               Call m_Details1.GetString(2, TempStr1, TempStr2)
               strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
               Vsp.FontBold = True
               Call Vsp.AddTable(strFormat, "", TempStr2)
               Vsp.FontBold = False
            End If
            For j = 1 To UBound(Total3)
                    Total3(j) = 0
            Next j
            If Sg.GROUP_TYPE_CODE > 0 Then
'               Vsp = Sg.GROUP_TYPE_CODE & "  /  " & Sg.GROUP_TYPE_NAME
'               Vsp.Paragraph = Sg.GROUP_TYPE_CODE & "  /  " & Sg.GROUP_TYPE_NAME
               Vsp.TableBorder = tbNone
               Call GenerateHeader("" & Sg.GROUP_TYPE_CODE & "  /  " & Sg.GROUP_TYPE_NAME, m_Details1, 0)
               Call m_Details1.GetString(2, TempStr1, TempStr2)
               strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
               Call Vsp.AddTable(strFormat, "", TempStr2)
               Vsp.TableBorder = tbAll
            Else
               Vsp.TextColor = RGB(255, 0, 0)
'               Vsp = "ยังไม่ได้สร้างกลุ่ม"
               Vsp.TableBorder = tbNone
               Call GenerateHeader("ยังไม่ได้สร้างกลุ่ม", m_Details1, 0)
               Call m_Details1.GetString(2, TempStr1, TempStr2)
               strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
               Call Vsp.AddTable(strFormat, "", TempStr2)
               Vsp.TableBorder = tbAll
            End If

            Vsp.TextColor = RGB(0, 0, 0)
         End If
         
         PrevKey1 = Sg.GROUP_TYPE_NAME
     
         Set Ap = New CAPMas
         Set BodyCf = New CReportField
   
         Ap.FROM_SUPPLIER_CODE = mcolParam("FROM_SUPPLIER_CODE")
         Ap.TO_SUPPLIER_CODE = mcolParam("TO_SUPPLIER_CODE")
         If mcolParam("SUPPLIER_TYPE") > 0 Then
            Set St = GetItemFromListIndex(m_SupplierTypes, mcolParam("SUPPLIER_TYPE"))
            Ap.SUPTYP = St.TYPCOD
         Else
            Ap.SUPTYP = ""
         End If
         Ap.SUPCOD = Sg.SUPPLIER_CODE
         Ap.OrderBy = mcolParam("ORDER_TYPE")
         Ap.OrderType = mcolParam("ORDER_BY")
         Call glbDaily.QuerySupplier(Ap, Rs, iCount, IsOK, glbErrorLog)
   
         While Not Rs.EOF
            For j = 1 To UBound(Total1)
               Total1(j) = 0
            Next j
            
            Call m_Details1.ClearField
            Call Ap.PopulateFromRS(1, Rs)
            
            Set TempApt = GetAPTrn(m_ApSupAmounts, Ap.SUPCOD)
            Set TempApRcIt = GetAPRcpItem(m_ApSupPaids, Ap.SUPCOD)
            Set TempApEx = GetAPRcpItem(m_ApAllAmounts, Ap.SUPCOD)
               
            If TempApt.AMOUNT - TempApt.PAYAMT - TempApRcIt.PAYAMT + TempApEx.PAYAMT Then 'มียอดค้างจ่าย
               '===Head
'               Vsp = "   " & Ap.SUPCOD & "  /  " & Ap.SUPNAM
'               Vsp.Paragraph = Ap.SUPCOD & "  /  " & Ap.SUPNAM
               Vsp.TableBorder = tbNone
               Call GenerateHeader("" & Ap.SUPCOD & "  /  " & Ap.SUPNAM, m_Details1, 0)
               Call m_Details1.GetString(2, TempStr1, TempStr2)
               strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
               Call Vsp.AddTable(strFormat, "", TempStr2)
               Vsp.TableBorder = tbAll
               '===
               '=== Detail
               Set Apt = New CApTrn
               Apt.SUPCOD = Ap.SUPCOD
               Apt.FROM_DUE_DATE = mcolParam("FROM_DUE_DATE")
               Apt.TO_DUE_DATE = mcolParam("TO_DUE_DATE")
               Apt.FROM_DOC_DATE = mcolParam("FROM_DOC_DATE")
               Apt.TO_DOC_DATE = mcolParam("TO_DOC_DATE")
               Apt.RECTYP = ""
               Apt.RecTypeSet = "('3', '4', '5')"
               Apt.OrderBy = 1
               Apt.OrderType = 1
               Call Apt.QueryData(1, TempRs, iCount)
   
               While Not TempRs.EOF
                  i = i + 1
                  Call m_Details1.ClearField
                  Call Apt.PopulateFromRS(1, TempRs)
                  
                  Call GenerateDetail(alngX(), i, Apt, Total1, Total2, Total3)
                  
                  TempRs.MoveNext
               Wend
               Set Apt = Nothing
               
               Call GenerateFooter("รวม", "" & Ap.SUPCOD, "", Total1, m_Details1, 0)
               Call m_Details1.GetString(2, TempStr1, TempStr2)
               strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
               Vsp.FontBold = True
               Call Vsp.AddTable(strFormat, "", TempStr2)
               Vsp.FontBold = False
               i = 0
               '===
             End If
             
            Rs.MoveNext
         Wend
      Next Sg

      Call GenerateFooter("รวม", "" & PrevKey1, "", Total3, m_Details1, 0)
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
      ElseIf (j = 4) Then
         TempStr = ""
      ElseIf j = m_Headers1.Fields.Count Then
         TempStr = ""
       ElseIf j = (m_Headers1.Fields.Count - 1) Then
         TempStr = ""
      Else
         TempStr = FormatNumber(Tot(j))
      End If
      
      Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.Align, TempStr, HeadCf.BodyAlign)
      Call Details.AddField(BodyCf)
   Next HeadCf
      
   Set BodyCf = Nothing
End Sub
Private Sub GenerateHeader(Txt1 As String, Details As CFieldList, DateCount As Long)
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
         Call BodyCf.SetFieldValue(30, HeadCf.Align, TempStr, "<")
         Call Details.AddField(BodyCf)
      End If
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
Private Sub AddCollection(Src As Collection, Des As Collection)
Dim Sg As CSupplierGroup
Dim Apm As CAPMas
   For Each Apm In Des
      Set Sg = GetObject("CSupplierGroup", Src, Trim(Apm.SUPCOD), False)
      If Sg Is Nothing Then
          Set Sg = New CSupplierGroup
          Sg.SUPPLIER_CODE = Apm.SUPCOD
          Sg.SUPPLIER_NAME = Apm.SUPNAM
          Call Src.Add(Sg, Trim(Sg.SUPPLIER_CODE))
      End If
   Next Apm
End Sub
