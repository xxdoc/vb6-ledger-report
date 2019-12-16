VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CReportAP002_2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'User Group Report
Implements CReportInterface

Private Const MODULE_NAME = "CReportAP002_2"
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

Private Supplier As Collection

Private SupplierGroupColl As Collection
Private ApReceiveChequeColl As Collection
Private HaveHeader As Boolean
Private AllSumWidth As Double

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
   Set ApReceiveChequeColl = New Collection
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
   Set ApReceiveChequeColl = Nothing
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

Dim FromDate As Date
Dim ToDate As Date
Dim TempDate As Date

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
   Call m_Headers1.AddField(Cf)
   Call Cf.SetFieldValue(3, "^", "รหัสเจ้าหนี้", "^")
   Call m_Headers1.AddField(Cf)
   If mcolParam("NOT_SHOW_BILL") = "N" And mcolParam("SHOW_ONLY_GROUP") = "N" Then
      Call Cf.SetFieldValue(3, "^", "วันครบกำหนด", "^")
      Call m_Headers1.AddField(Cf)
      Call Cf.SetFieldValue(3, "^", "เลขที่ใบรับของ", "^")
      Call m_Headers1.AddField(Cf)
      Call Cf.SetFieldValue(3, "^", "วันที่รับของ", "^")
      Call m_Headers1.AddField(Cf)
   Else
      Call Cf.SetFieldValue(0, "^", "", "^")
      Call m_Headers1.AddField(Cf)
      Call Cf.SetFieldValue(0, "^", "", "^")
      Call m_Headers1.AddField(Cf)
      Call Cf.SetFieldValue(7, "^", "รายละเอียดเจ้าหนี้/กลุ่ม", "^")
      Call m_Headers1.AddField(Cf)
   End If
   
   Call Cf.SetFieldValue(2, "^", "รวมหนี้ค้าง", ">")
   Call m_Headers1.AddField(Cf)
   
   Call Cf.SetFieldValue(2, "^", "ยอดเช็ค", ">")
   Call m_Headers1.AddField(Cf)
   
   Call Cf.SetFieldValue(2, "^", "หนี้สุทธิ", ">")
   Call m_Headers1.AddField(Cf)
   AllSumWidth = 19
   
   FromDate = mcolParam("FROM_OVER_DUE")
   ToDate = mcolParam("TO_OVER_DUE")
   Call GetFirstLastDate(mcolParam("FROM_OVER_DUE"), FromDate, TempDate)
   Call GetFirstLastDate(mcolParam("TO_OVER_DUE"), TempDate, ToDate)
   
   Call Cf.SetFieldValue(2, "^", "< " & Format(Month(FromDate), "00") & "/" & Format(Year(FromDate) + 543, "00"), ">")
   Call m_Headers1.AddField(Cf)
   
   If Year(ToDate) = Year(ToDate) Then
      Call Cf.SetFieldValue(2, "^", Format(Month(FromDate), "00") & "-" & Format(Month(ToDate), "00") & "/" & Right(Format(Year(ToDate) + 543, "00"), 2), ">")
   Else
      Call Cf.SetFieldValue(2, "^", Format(Month(FromDate), "00") & "/" & Format(Year(FromDate) + 543, "00") & "-" & Format(Month(ToDate), "00") & "/" & Format(Year(ToDate) + 543, "00"), ">")
   End If
   Call m_Headers1.AddField(Cf)
    AllSumWidth = AllSumWidth + 4
    
   FromDate = mcolParam("FROM_DUE_DATE2")
   ToDate = mcolParam("TO_DUE_DATE2")
   Call GetFirstLastDate(mcolParam("FROM_DUE_DATE2"), FromDate, TempDate)
   Call GetFirstLastDate(mcolParam("TO_DUE_DATE2"), TempDate, ToDate)
      
   While (FromDate < ToDate)
      AllSumWidth = AllSumWidth + 2
      Call Cf.SetFieldValue(2, "^", Format(Month(FromDate), "00") & "/" & Format(Year(FromDate) + 543, "00"), ">")
      Call m_Headers1.AddField(Cf)
      FromDate = DateAdd("M", 1, FromDate)
   Wend
   
   AllSumWidth = AllSumWidth + 2
   FromDate = DateAdd("M", -1, FromDate)
   Call Cf.SetFieldValue(2, "^", Format(Month(FromDate), "00") & "/" & Format(Year(FromDate) + 543, "00") & ">", ">")
   Call m_Headers1.AddField(Cf)
   
   Set Cf = Nothing
           
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
   Vsp.MarginBottom = 300
   Vsp.MarginFooter = 300
   Vsp.MarginHeader = 300
   Vsp.MarginLeft = 300
   Vsp.MarginRight = 300
   Vsp.MarginTop = 300
   Vsp.FontName = "AngsanaUPC"
   Vsp.FontSize = DETAIL_SIZE

   Call LoadSupplierType(Nothing, m_SupplierTypes)
   Call LoadPaidAmountByBill(Nothing, m_PaidAmounts, -1, mcolParam("TO_PAY_DATE"))
   Call LoadPaidAmountByBill(Nothing, m_PaidAllAmounts, -1, -1)
   Call LoadDueDateInterval1(Nothing, m_MaxMins)
   Call LoadAPAmountBySup(Nothing, m_ApSupAmounts, -1, mcolParam("TO_DOC_DATE"))       'ยอดซื้อ ถึงวันที่
   Call LoadPaidAmountBySup(Nothing, m_ApSupPaids, -1, mcolParam("TO_PAY_DATE"))             'ยอดจ่ายถึงวันที่
   Call LoadPaidAmountBySup(Nothing, m_ApAllAmounts, -1, -1)         'ยอดตัดบิลทั้งหมดของบิลภายในวันที่กำหนด
   
'   If ((mcolParam("DATA_TYPE_ID") > 0) And (mcolParam("GROUP_TYPE_CODE") > 0)) Then
'        Call LoadSupplierGroup(Nothing, SupplierGroupColl, , , mcolParam("DATA_TYPE_ID"), mcolParam("GROUP_TYPE_CODE"))
'   End If
   
   Call LoadAPAmountByReceiveCheque(Nothing, ApReceiveChequeColl, -1, mcolParam("TO_PAY_DATE"), "('3')", DateAdd("D", 1, mcolParam("TO_PAY_DATE")), -1)
   
   Call LoadSupplier(Nothing, Supplier)
   Call LoadSupplierGroup(Nothing, SupplierGroupColl, , , mcolParam("DATA_TYPE_ID"))
   Call AddCollection(SupplierGroupColl, Supplier)
   
'   Call LoadDbnCdnByBill(Nothing, m_DbnCdnByBills, -1, mcolParam("TO_PAY_DATE"))
   FromDate = "จากวันที่ " & EmptyToString(DateToStringExtEx2(mcolParam("FROM_DOC_DATE")), "N/A")
   ToDate = " ถึง วันที่ " & EmptyToString(DateToStringExtEx2(mcolParam("TO_DOC_DATE")), "N/A")
   ToPaidDate = "    วันที่ชำระเงินถึงวันที่ " & EmptyToString(DateToStringExtEx2(mcolParam("TO_PAY_DATE")), "N/A") & "    วันที่พิมพ์ " & EmptyToString(DateToStringExtEx2(mcolParam("PRINT_DATE")), "N/A") & "    ประเภทข้อมูล " & EmptyToString(mcolParam("DATA_TYPE_NAME"), "N/A")

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

Private Sub GenerateDetail(alngX() As Long, i As Long, Apt As CApTrn, Tot1() As Double, Tot2() As Double, SupGroup As CSupplierGroup, Tot3() As Double)
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
Dim TempApTrn As CApTrn
Dim TempApt As CApTrn
Dim FromDate As Date
Dim ToDate As Date
Dim TempAmt As Double
Dim TempDate  As Date
   
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
      
'      If (Apt.DOCNUM = "RR55010250") Then
'        'debug.print
'      End If
      
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
            TempStr = FormatNumberToNull(Apt.AMOUNT - (Ari.PAYAMT + PaidBalance), Val(mcolParam("DECIMAL_AMOUNT")))
            
            'Tot1(j) = Tot1(j) + (Apt.AMOUNT - (Ari.PAYAMT + PaidBalance))
            'Tot2(j) = Tot2(j) + (Apt.AMOUNT - (Ari.PAYAMT + PaidBalance))
            If Apt.RECTYP = 5 Then 'ลดหนี้
               Sum = Sum + ((Ari.PAYAMT + PaidBalance) - Apt.AMOUNT)
            Else
               Sum = Sum + (Apt.AMOUNT - (Ari.PAYAMT + PaidBalance))
            End If
            
         Else
            TempStr = FormatNumberToNull(0, Val(mcolParam("DECIMAL_AMOUNT")))
         End If
'         Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
'         Call m_Details1.AddField(BodyCf)
                  
         'j = j + 1
      Next MM
            
      Set HeadCf = m_Headers1.Fields(j)
      TempStr = FormatNumberToNull(Sum, Val(mcolParam("DECIMAL_AMOUNT")))
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
      
'      Set TempApTrn = GetAPTrn(ApReceiveChequeColl, Apt.RCPNUM)
      
       j = j + 1
      Set HeadCf = m_Headers1.Fields(j)
      TempStr = "" 'FormatNumberToNull(TempApTrn.AMOUNT,val(mcolParam("DECIMAL_AMOUNT")))
      Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
      Call m_Details1.AddField(BodyCf)
      
       j = j + 1
      Set HeadCf = m_Headers1.Fields(j)
      TempStr = ""
      Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
      Call m_Details1.AddField(BodyCf)
      
      
      FromDate = mcolParam("FROM_OVER_DUE")
      ToDate = mcolParam("TO_OVER_DUE")
      Call GetFirstLastDate(mcolParam("FROM_OVER_DUE"), FromDate, TempDate)
      Call GetFirstLastDate(mcolParam("TO_OVER_DUE"), TempDate, ToDate)
      
      j = j + 1
      If Apt.DUEDAT < FromDate Then
         Set HeadCf = m_Headers1.Fields(j)
         TempStr = FormatNumberToNull(Sum, Val(mcolParam("DECIMAL_AMOUNT")))
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
      Else
          Set HeadCf = m_Headers1.Fields(j)
         TempStr = ""
         Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
         Call m_Details1.AddField(BodyCf)
      End If
      
      j = j + 1
      If Apt.DUEDAT >= FromDate And Apt.DUEDAT <= ToDate Then
         Set HeadCf = m_Headers1.Fields(j)
         TempStr = FormatNumberToNull(Sum, Val(mcolParam("DECIMAL_AMOUNT")))
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
      Else
         Set HeadCf = m_Headers1.Fields(j)
         TempStr = ""
         Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
         Call m_Details1.AddField(BodyCf)
      End If
          
      FromDate = mcolParam("FROM_DUE_DATE2")
      ToDate = mcolParam("TO_DUE_DATE2")
      Call GetFirstLastDate(mcolParam("FROM_DUE_DATE2"), FromDate, TempDate)
      Call GetFirstLastDate(mcolParam("TO_DUE_DATE2"), TempDate, ToDate)
   
      While (FromDate < ToDate)
         j = j + 1
         Call GetFirstLastDate(FromDate, FromDate, TempDate)
         If Apt.DUEDAT >= FromDate And Apt.DUEDAT <= TempDate Then
            Set HeadCf = m_Headers1.Fields(j)
            TempStr = FormatNumberToNull(Sum, Val(mcolParam("DECIMAL_AMOUNT")))
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
         Else
            Set HeadCf = m_Headers1.Fields(j)
            TempStr = ""
            Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
            Call m_Details1.AddField(BodyCf)
         End If
         FromDate = DateAdd("M", 1, FromDate)
      Wend
      
      
      j = j + 1
      If Apt.DUEDAT > ToDate Then
         Set HeadCf = m_Headers1.Fields(j)
         TempStr = FormatNumberToNull(Sum, Val(mcolParam("DECIMAL_AMOUNT")))
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
      Else
          Set HeadCf = m_Headers1.Fields(j)
         TempStr = ""
         Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
         Call m_Details1.AddField(BodyCf)
      End If
      
      '----------------------------------------------------
      If Not (HaveHeader) Then
         Call m_Details2.ClearField
         
         Width = m_Headers1.Fields(1).Width + m_Headers1.Fields(2).Width
         TempStr = SupGroup.SUPPLIER_CODE
         Call BodyCf.SetFieldValue(Width, "^", TempStr)
         Call m_Details2.AddField(BodyCf)

         Width = AllSumWidth - (m_Headers1.Fields(1).Width + m_Headers1.Fields(2).Width)
         TempStr = SupGroup.SUPPLIER_NAME
         Call BodyCf.SetFieldValue(Width, "<", TempStr)
         Call m_Details2.AddField(BodyCf)
         
         If mcolParam("NOT_SHOW_BILL") = "N" And mcolParam("SHOW_ONLY_GROUP") = "N" Then
            Call m_Details2.GetString(1, TempStr1, TempStr2)
            strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
            Call Vsp.AddTable(strFormat, "", TempStr2)
         End If
         
         HaveHeader = True
         
         For Each TempApt In ApReceiveChequeColl
            If TempApt.SUPCOD = SupGroup.SUPPLIER_CODE Then
                Call m_Details2.ClearField
                                      
               Width = m_Headers1.Fields(1).Width + m_Headers1.Fields(2).Width
                TempStr = TempApt.RCPNUM
                Call BodyCf.SetFieldValue(Width, "<", TempStr, "<")
                Call m_Details2.AddField(BodyCf)
            
               Width = m_Headers1.Fields(3).Width + m_Headers1.Fields(4).Width + m_Headers1.Fields(5).Width + m_Headers1.Fields(6).Width
                TempStr = TempApt.DOCNUM & " " & TempApt.CHQNUM
                Call BodyCf.SetFieldValue(Width, "<", TempStr, "<")
                Call m_Details2.AddField(BodyCf)
                TempApt.DOCNUM = ""                               'ตรงนี้เป็นตัวที่ใช้ SET ว่าเอาไปใช้แล้ว
                
'                Width = m_Headers1.Fields(5).Width + m_Headers1.Fields(6).Width
'                TempStr = TempApt.CHQNUM
'                Call BodyCf.SetFieldValue(Width, "<", TempStr, "<")
'                Call m_Details2.AddField(BodyCf)
                
                Set HeadCf = m_Headers1.Fields(7)
                TempStr = FormatNumberToNull(TempApt.AMOUNT, Val(mcolParam("DECIMAL_AMOUNT")))
                Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr, HeadCf.BodyAlign)
                Call m_Details2.AddField(BodyCf)
                
                Tot1(7) = Tot1(7) + TempApt.AMOUNT
                Tot2(7) = Tot2(7) + TempApt.AMOUNT
                Tot3(7) = Tot3(7) + TempApt.AMOUNT
         
                Set HeadCf = m_Headers1.Fields(8)
                TempStr = ""
                Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
                Call m_Details2.AddField(BodyCf)
                  
               j = 8
               FromDate = mcolParam("FROM_OVER_DUE")
               ToDate = mcolParam("TO_OVER_DUE")
               Call GetFirstLastDate(mcolParam("FROM_OVER_DUE"), FromDate, TempDate)
               Call GetFirstLastDate(mcolParam("TO_OVER_DUE"), TempDate, ToDate)
               
               j = j + 1
               If TempApt.CHQDAT < FromDate Then
                  Set HeadCf = m_Headers1.Fields(j)
                  TempStr = FormatNumberToNull(TempApt.AMOUNT, Val(mcolParam("DECIMAL_AMOUNT")))
                  Tot1(j) = Tot1(j) + TempApt.AMOUNT
                  Tot2(j) = Tot2(j) + TempApt.AMOUNT
                  Tot3(j) = Tot3(j) + TempApt.AMOUNT
                  Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
                  Call m_Details2.AddField(BodyCf)
               Else
                   Set HeadCf = m_Headers1.Fields(j)
                  TempStr = ""
                  Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
                  Call m_Details2.AddField(BodyCf)
               End If
               
               j = j + 1
               If TempApt.CHQDAT >= FromDate And TempApt.CHQDAT <= ToDate Then
                  Set HeadCf = m_Headers1.Fields(j)
                  TempStr = FormatNumberToNull(TempApt.AMOUNT, Val(mcolParam("DECIMAL_AMOUNT")))
                  Tot1(j) = Tot1(j) + TempApt.AMOUNT
                  Tot2(j) = Tot2(j) + TempApt.AMOUNT
                  Tot3(j) = Tot3(j) + TempApt.AMOUNT
                  Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
                  Call m_Details2.AddField(BodyCf)
               Else
                   Set HeadCf = m_Headers1.Fields(j)
                  TempStr = ""
                  Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
                  Call m_Details2.AddField(BodyCf)
               End If
               
               FromDate = mcolParam("FROM_DUE_DATE2")
               ToDate = mcolParam("TO_DUE_DATE2")
               Call GetFirstLastDate(mcolParam("FROM_DUE_DATE2"), FromDate, TempDate)
               Call GetFirstLastDate(mcolParam("TO_DUE_DATE2"), TempDate, ToDate)
   
               While (FromDate < ToDate)
                  j = j + 1
                  Call GetFirstLastDate(FromDate, FromDate, TempDate)
                  If TempApt.CHQDAT >= FromDate And TempApt.CHQDAT <= TempDate Then
                     Set HeadCf = m_Headers1.Fields(j)
                           TempStr = FormatNumberToNull(TempApt.AMOUNT, Val(mcolParam("DECIMAL_AMOUNT")))
                           Tot1(j) = Tot1(j) + TempApt.AMOUNT
                           Tot2(j) = Tot2(j) + TempApt.AMOUNT
                           Tot3(j) = Tot3(j) + TempApt.AMOUNT
                           Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
                           Call m_Details2.AddField(BodyCf)
                     Else
                        Set HeadCf = m_Headers1.Fields(j)
                        TempStr = ""
                        Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
                        Call m_Details2.AddField(BodyCf)
                     End If
                     FromDate = DateAdd("M", 1, FromDate)
                  Wend
      
      
               j = j + 1
               If TempApt.CHQDAT > ToDate Then
                  Set HeadCf = m_Headers1.Fields(j)
                  TempStr = FormatNumberToNull(TempApt.AMOUNT, Val(mcolParam("DECIMAL_AMOUNT")))
                  Tot1(j) = Tot1(j) + TempApt.AMOUNT
                  Tot2(j) = Tot2(j) + TempApt.AMOUNT
                  Tot3(j) = Tot3(j) + TempApt.AMOUNT
                  Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
                  Call m_Details2.AddField(BodyCf)
               Else
                   Set HeadCf = m_Headers1.Fields(j)
                  TempStr = ""
                  Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
                  Call m_Details2.AddField(BodyCf)
               End If
               
                If mcolParam("NOT_SHOW_BILL") = "N" And mcolParam("SHOW_ONLY_GROUP") = "N" Then
                  Call m_Details2.GetString(1, TempStr1, TempStr2)
                  strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
                  Vsp.FontBold = True
                  Call Vsp.AddTable(strFormat, "", TempStr2)
                  Vsp.FontBold = False
                End If
            End If
        Next TempApt
      End If
      '----------------------------------------------------
      
      If mcolParam("NOT_SHOW_BILL") = "N" And mcolParam("SHOW_ONLY_GROUP") = "N" Then
         Call m_Details1.GetString(1, TempStr1, TempStr2)
         strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
         Call Vsp.AddTable(strFormat, "", TempStr2)
      End If
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
Dim St As CIsTab
'Dim Ap As CAPMas
Dim Width As Double
Dim Apt As CApTrn
Dim TempRs As ADODB.Recordset
Dim TempApt As CApTrn
Dim TempApRcIt As CAPRcIt
Dim HaveValue As Boolean
Dim TempApEx As CAPRcIt
Dim FromDate  As Date
Dim ToDate As Date
Dim TempDate As Date
Dim SupGroup As CSupplierGroup

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
      
      For Each SupGroup In SupplierGroupColl
         If (i >= 0 And PrevKey1 <> Trim(Str(SupGroup.GROUP_TYPE_CODE))) Then
            
            If i > 0 Then
               Call GenerateFooter("รวม", "", PrevKey2, Total2, m_Details1, 0)
               Vsp.FontBold = True
               Call m_Details1.GetString(2, TempStr1, TempStr2)
               strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
               Call Vsp.AddTable(strFormat, "", TempStr2)
               Vsp.FontBold = False
               
               For j = 1 To UBound(Total1)
                  Total2(j) = 0
               Next j
            End If
         
            Call m_Details1.ClearField
            Width = AllSumWidth
            If SupGroup.GROUP_TYPE_CODE > 0 Then
               TempStr = SupGroup.GROUP_TYPE_NAME & " / " & SupGroup.GROUP_TYPE_CODE
            Else
               TempStr = "ยังไม่ได้สร้างกลุ่ม"
            End If
            Call BodyCf.SetFieldValue(Width, "<", TempStr)
            Call m_Details1.AddField(BodyCf)
            
            If mcolParam("SHOW_ONLY_GROUP") = "N" Then
               Call m_Details1.GetString(1, TempStr1, TempStr2)
               strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
               Vsp.FontBold = True
               Call Vsp.AddTable(strFormat, "", TempStr2)
               Vsp.FontBold = False
            End If
         End If
      
         i = i + 1
         PrevKey1 = SupGroup.GROUP_TYPE_CODE
         PrevKey2 = SupGroup.GROUP_TYPE_NAME
      
         
         For j = 1 To UBound(Total1)
            Total1(j) = 0
         Next j
         
            Call m_Details1.ClearField
         
             Set TempApt = GetAPTrn(m_ApSupAmounts, SupGroup.SUPPLIER_CODE)
             Set TempApRcIt = GetAPRcpItem(m_ApSupPaids, SupGroup.SUPPLIER_CODE)
             Set TempApEx = GetAPRcpItem(m_ApAllAmounts, SupGroup.SUPPLIER_CODE)
                
             ''debug.print (TempApt.AMOUNT - TempApt.PAYAMT - TempApRcIt.PAYAMT + TempApEx.PAYAMT)
             If TempApt.AMOUNT - TempApt.PAYAMT - TempApRcIt.PAYAMT + TempApEx.PAYAMT Then 'มียอดค้างจ่าย
                '===Head
                HaveHeader = False
                '===
                
                '=== Detail
                Set Apt = New CApTrn
                Apt.SUPCOD = SupGroup.SUPPLIER_CODE
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
                   
                   Call GenerateDetail(alngX(), i, Apt, Total1, Total2, SupGroup, Total3)
                   
                   TempRs.MoveNext
                Wend
                Set Apt = Nothing
                
                If HaveHeader Then
                  If mcolParam("SHOW_ONLY_GROUP") = "N" Then
                     Call GenerateFooter("รวม", SupGroup.SUPPLIER_CODE, SupGroup.SUPPLIER_NAME, Total1, m_Details1, 0)
                     Call m_Details1.GetString(2, TempStr1, TempStr2)
                     strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
                     Vsp.FontBold = True
                     Call Vsp.AddTable(strFormat, "", TempStr2)
                     Vsp.FontBold = False
                   End If
                End If
                '===
            End If
                  
            'debug.print
            HaveValue = False
            
            For Each TempApt In ApReceiveChequeColl
               If (Len(TempApt.DOCNUM) > 0) And TempApt.SUPCOD = SupGroup.SUPPLIER_CODE Then
                  
                  If Not HaveValue Then
                     Call m_Details2.ClearField
                                  
                      Width = m_Headers1.Fields(1).Width + m_Headers1.Fields(2).Width
                     TempStr = TempApt.SUPCOD
                     Call BodyCf.SetFieldValue(Width, "^", TempStr)
                     Call m_Details2.AddField(BodyCf)
            
                     Width = AllSumWidth - (m_Headers1.Fields(1).Width + m_Headers1.Fields(2).Width)
                     TempStr = TempApt.SUPNAM & " มีเฉพาะเช็คลงวันที่ล่วงหน้าเท่านั้น"
                     Call BodyCf.SetFieldValue(Width, "<", TempStr)
                     Call m_Details2.AddField(BodyCf)
                     
                     If mcolParam("NOT_SHOW_BILL") = "N" And mcolParam("SHOW_ONLY_GROUP") = "N" Then
                        Call m_Details2.GetString(1, TempStr1, TempStr2)
                        strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
                        Call Vsp.AddTable(strFormat, "", TempStr2)
                     End If
                   End If
                   HaveValue = True
                   
                   PrevKey3 = TempApt.SUPCOD
                                          
                  Call m_Details2.ClearField
         
                   Width = m_Headers1.Fields(1).Width + m_Headers1.Fields(2).Width
                   TempStr = TempApt.RCPNUM
                   Call BodyCf.SetFieldValue(Width, "<", TempStr, "<")
                   Call m_Details2.AddField(BodyCf)
               
                    Width = m_Headers1.Fields(3).Width + m_Headers1.Fields(4).Width + m_Headers1.Fields(5).Width + m_Headers1.Fields(6).Width
                   TempStr = TempApt.DOCNUM & " " & TempApt.CHQNUM
                   Call BodyCf.SetFieldValue(Width, "<", TempStr, "<")
                   Call m_Details2.AddField(BodyCf)
                   TempApt.DOCNUM = ""
                   
                   Set HeadCf = m_Headers1.Fields(7)
                   TempStr = FormatNumberToNull(TempApt.AMOUNT, Val(mcolParam("DECIMAL_AMOUNT")))
                   Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr, HeadCf.BodyAlign)
                   Call m_Details2.AddField(BodyCf)
                   
                   Total1(7) = Total1(7) + TempApt.AMOUNT
                   Total2(7) = Total2(7) + TempApt.AMOUNT
                   Total3(7) = Total3(7) + TempApt.AMOUNT
                  
                   Set HeadCf = m_Headers1.Fields(8)
                   TempStr = ""
                   Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
                   Call m_Details2.AddField(BodyCf)
                        
                    j = 8
                  FromDate = mcolParam("FROM_OVER_DUE")
                  ToDate = mcolParam("TO_OVER_DUE")
                  Call GetFirstLastDate(mcolParam("FROM_OVER_DUE"), FromDate, TempDate)
                  Call GetFirstLastDate(mcolParam("TO_OVER_DUE"), TempDate, ToDate)
                  
                  j = j + 1
                  If TempApt.CHQDAT < FromDate Then
                     Set HeadCf = m_Headers1.Fields(j)
                     TempStr = FormatNumberToNull(TempApt.AMOUNT, Val(mcolParam("DECIMAL_AMOUNT")))
                     Total1(j) = Total1(j) + TempApt.AMOUNT
                     Total2(j) = Total2(j) + TempApt.AMOUNT
                     Total3(j) = Total3(j) + TempApt.AMOUNT
                     Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
                     Call m_Details2.AddField(BodyCf)
                  Else
                      Set HeadCf = m_Headers1.Fields(j)
                     TempStr = ""
                     Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
                     Call m_Details2.AddField(BodyCf)
                  End If
                  
                  j = j + 1
                  If TempApt.CHQDAT >= FromDate And TempApt.CHQDAT <= ToDate Then
                     Set HeadCf = m_Headers1.Fields(j)
                     TempStr = FormatNumberToNull(TempApt.AMOUNT, Val(mcolParam("DECIMAL_AMOUNT")))
                     Total1(j) = Total1(j) + TempApt.AMOUNT
                     Total2(j) = Total2(j) + TempApt.AMOUNT
                     Total3(j) = Total3(j) + TempApt.AMOUNT
                     Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
                     Call m_Details2.AddField(BodyCf)
                  Else
                      Set HeadCf = m_Headers1.Fields(j)
                     TempStr = ""
                     Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
                     Call m_Details2.AddField(BodyCf)
                  End If
                  
                  FromDate = mcolParam("FROM_DUE_DATE2")
                  ToDate = mcolParam("TO_DUE_DATE2")
                  Call GetFirstLastDate(mcolParam("FROM_DUE_DATE2"), FromDate, TempDate)
                  Call GetFirstLastDate(mcolParam("TO_DUE_DATE2"), TempDate, ToDate)
      
                  While (FromDate < ToDate)
                     j = j + 1
                     Call GetFirstLastDate(FromDate, FromDate, TempDate)
                     If TempApt.CHQDAT >= FromDate And TempApt.CHQDAT <= TempDate Then
                        Set HeadCf = m_Headers1.Fields(j)
                              TempStr = FormatNumberToNull(TempApt.AMOUNT, Val(mcolParam("DECIMAL_AMOUNT")))
                              Total1(j) = Total1(j) + TempApt.AMOUNT
                              Total2(j) = Total2(j) + TempApt.AMOUNT
                              Total3(j) = Total3(j) + TempApt.AMOUNT
                              Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
                              Call m_Details2.AddField(BodyCf)
                        Else
                           Set HeadCf = m_Headers1.Fields(j)
                           TempStr = ""
                           Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
                           Call m_Details2.AddField(BodyCf)
                        End If
                        FromDate = DateAdd("M", 1, FromDate)
                     Wend
         
         
                  j = j + 1
                  If TempApt.CHQDAT > ToDate Then
                     Set HeadCf = m_Headers1.Fields(j)
                     TempStr = FormatNumberToNull(TempApt.AMOUNT, Val(mcolParam("DECIMAL_AMOUNT")))
                     Total1(j) = Total1(j) + TempApt.AMOUNT
                     Total2(j) = Total2(j) + TempApt.AMOUNT
                     Total3(j) = Total3(j) + TempApt.AMOUNT
                     Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
                     Call m_Details2.AddField(BodyCf)
                  Else
                      Set HeadCf = m_Headers1.Fields(j)
                     TempStr = ""
                     Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
                     Call m_Details2.AddField(BodyCf)
                  End If
                  
                  If mcolParam("NOT_SHOW_BILL") = "N" And mcolParam("SHOW_ONLY_GROUP") = "N" Then
                     Call m_Details2.GetString(1, TempStr1, TempStr2)
                     strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
                     Vsp.FontBold = True
                     Call Vsp.AddTable(strFormat, "", TempStr2)
                     Vsp.FontBold = False
                  End If
               End If
            Next TempApt
            
            If HaveValue Then
               If mcolParam("SHOW_ONLY_GROUP") = "N" Then
                  Call GenerateFooter("รวม", SupGroup.SUPPLIER_CODE, SupGroup.SUPPLIER_NAME, Total1, m_Details1, 0)
                  Call m_Details1.GetString(2, TempStr1, TempStr2)
                  strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
                  Vsp.FontBold = True
                  Call Vsp.AddTable(strFormat, "", TempStr2)
                  Vsp.FontBold = False
               End If
            End If
      Next SupGroup
              
'      Vsp = "เช็คจ่ายลงวันที่ล่วงหน้าส่วนที่เหลือ"
'
'      PrevKey1 = ""
'      For Each TempApt In ApReceiveChequeColl
'            If (Len(TempApt.DOCNUM) > 0) And (Len(mcolParam("FROM_SUPPLIER_CODE")) <= 0 Or TempApt.SUPCOD >= mcolParam("FROM_SUPPLIER_CODE")) And (Len(mcolParam("TO_SUPPLIER_CODE")) <= 0 Or TempApt.SUPCOD <= mcolParam("TO_SUPPLIER_CODE")) Then
'               '
'                If PrevKey1 <> TempApt.SUPCOD Then
'                  If (PrevKey1 <> "") Then
'                     Call GenerateFooter("รวม", PrevKey1, "", Total1, m_Details1, 0)
'                     Call m_Details1.GetString(2, TempStr1, TempStr2)
'                     strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
'                     Vsp.FontBold = True
'                     Call Vsp.AddTable(strFormat, "", TempStr2)
'                     Vsp.FontBold = False
'                  End If
'
'                   For j = 1 To UBound(Total1)
'                     Total1(j) = 0
'                  Next j
'
'                   Call m_Details2.ClearField
'
'                   Width = m_Headers1.Fields(1).Width + m_Headers1.Fields(2).Width
'                  TempStr = TempApt.SUPCOD
'                  Call BodyCf.SetFieldValue(Width, "^", TempStr)
'                  Call m_Details2.AddField(BodyCf)
'
'                  Width = AllSumWidth - (m_Headers1.Fields(1).Width + m_Headers1.Fields(2).Width)
'                  TempStr = TempApt.SUPNAM & " กลุ่ม : "
'                  Call BodyCf.SetFieldValue(Width, "<", TempStr)
'                  Call m_Details2.AddField(BodyCf)
'
'                  Call m_Details2.GetString(1, TempStr1, TempStr2)
'                  strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
'                  Call Vsp.AddTable(strFormat, "", TempStr2)
'                End If
'
'                PrevKey1 = TempApt.SUPCOD
'
'                Call m_Details2.ClearField
'
'                Width = m_Headers1.Fields(1).Width + m_Headers1.Fields(2).Width
'                TempStr = TempApt.RCPNUM
'                Call BodyCf.SetFieldValue(Width, "<", TempStr, "<")
'                Call m_Details2.AddField(BodyCf)
'
'                 Width = m_Headers1.Fields(3).Width + m_Headers1.Fields(4).Width + m_Headers1.Fields(5).Width + m_Headers1.Fields(6).Width
'                TempStr = TempApt.DOCNUM & " " & TempApt.CHQNUM
'                Call BodyCf.SetFieldValue(Width, "<", TempStr, "<")
'                Call m_Details2.AddField(BodyCf)
'                TempApt.DOCNUM = ""
'
'                Set HeadCf = m_Headers1.Fields(7)
'                TempStr = FormatNumberToNull(TempApt.AMOUNT, Val(mcolParam("DECIMAL_AMOUNT")))
'                Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr, HeadCf.BodyAlign)
'                Call m_Details2.AddField(BodyCf)
'
'                Total1(7) = Total1(7) + TempApt.AMOUNT
'                Total2(7) = Total2(7) + TempApt.AMOUNT
'                Total3(7) = Total3(7) + TempApt.AMOUNT
'
'                Set HeadCf = m_Headers1.Fields(8)
'                TempStr = ""
'                Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
'                Call m_Details2.AddField(BodyCf)
'
'                 j = 8
'               FromDate = mcolParam("FROM_OVER_DUE")
'               ToDate = mcolParam("TO_OVER_DUE")
'               Call GetFirstLastDate(mcolParam("FROM_OVER_DUE"), FromDate, TempDate)
'               Call GetFirstLastDate(mcolParam("TO_OVER_DUE"), TempDate, ToDate)
'
'               j = j + 1
'               If TempApt.CHQDAT < FromDate Then
'                  Set HeadCf = m_Headers1.Fields(j)
'                  TempStr = FormatNumberToNull(TempApt.AMOUNT, Val(mcolParam("DECIMAL_AMOUNT")))
'                  Total1(j) = Total1(j) + TempApt.AMOUNT
'                  Total2(j) = Total2(j) + TempApt.AMOUNT
'                  Total3(j) = Total3(j) + TempApt.AMOUNT
'                  Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
'                  Call m_Details2.AddField(BodyCf)
'               Else
'                   Set HeadCf = m_Headers1.Fields(j)
'                  TempStr = ""
'                  Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
'                  Call m_Details2.AddField(BodyCf)
'               End If
'
'               j = j + 1
'               If TempApt.CHQDAT >= FromDate And TempApt.CHQDAT <= ToDate Then
'                  Set HeadCf = m_Headers1.Fields(j)
'                  TempStr = FormatNumberToNull(TempApt.AMOUNT, Val(mcolParam("DECIMAL_AMOUNT")))
'                  Total1(j) = Total1(j) + TempApt.AMOUNT
'                  Total2(j) = Total2(j) + TempApt.AMOUNT
'                  Total3(j) = Total3(j) + TempApt.AMOUNT
'                  Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
'                  Call m_Details2.AddField(BodyCf)
'               Else
'                   Set HeadCf = m_Headers1.Fields(j)
'                  TempStr = ""
'                  Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
'                  Call m_Details2.AddField(BodyCf)
'               End If
'
'               FromDate = mcolParam("FROM_DUE_DATE2")
'               ToDate = mcolParam("TO_DUE_DATE2")
'               Call GetFirstLastDate(mcolParam("FROM_DUE_DATE2"), FromDate, TempDate)
'               Call GetFirstLastDate(mcolParam("TO_DUE_DATE2"), TempDate, ToDate)
'
'               While (FromDate < ToDate)
'                  j = j + 1
'                  Call GetFirstLastDate(FromDate, FromDate, TempDate)
'                  If TempApt.CHQDAT >= FromDate And TempApt.CHQDAT <= TempDate Then
'                     Set HeadCf = m_Headers1.Fields(j)
'                           TempStr = FormatNumberToNull(TempApt.AMOUNT, Val(mcolParam("DECIMAL_AMOUNT")))
'                           Total1(j) = Total1(j) + TempApt.AMOUNT
'                           Total2(j) = Total2(j) + TempApt.AMOUNT
'                           Total3(j) = Total3(j) + TempApt.AMOUNT
'                           Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
'                           Call m_Details2.AddField(BodyCf)
'                     Else
'                        Set HeadCf = m_Headers1.Fields(j)
'                        TempStr = ""
'                        Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
'                        Call m_Details2.AddField(BodyCf)
'                     End If
'                     FromDate = DateAdd("M", 1, FromDate)
'                  Wend
'
'
'               j = j + 1
'               If TempApt.CHQDAT > ToDate Then
'                  Set HeadCf = m_Headers1.Fields(j)
'                  TempStr = FormatNumberToNull(TempApt.AMOUNT, Val(mcolParam("DECIMAL_AMOUNT")))
'                  Total1(j) = Total1(j) + TempApt.AMOUNT
'                  Total2(j) = Total2(j) + TempApt.AMOUNT
'                  Total3(j) = Total3(j) + TempApt.AMOUNT
'                  Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
'                  Call m_Details2.AddField(BodyCf)
'               Else
'                   Set HeadCf = m_Headers1.Fields(j)
'                  TempStr = ""
'                  Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
'                  Call m_Details2.AddField(BodyCf)
'               End If
'
'                Call m_Details2.GetString(2, TempStr1, TempStr2)
'                strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
'                Vsp.FontBold = True
'                Call Vsp.AddTable(strFormat, "", TempStr2)
'                Vsp.FontBold = False
'            End If
'        Next TempApt
'
'      Call GenerateFooter("รวม", PrevKey1, "", Total2, m_Details1, 0)
'      Call m_Details1.GetString(2, TempStr1, TempStr2)
'      strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
'      Vsp.FontBold = True
'      Call Vsp.AddTable(strFormat, "", TempStr2)
'      Vsp.FontBold = False
                     
      Call GenerateFooter("สุทธิ", "", "", Total3, m_Details1, 0)
      Call m_Details1.GetString(2, TempStr1, TempStr2)
      strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
      Vsp.FontBold = True
      Call Vsp.AddTable(strFormat, "", TempStr2)
      Vsp.FontBold = False
   
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
Dim SumWidth As Double

   Call Details.ClearField
   Set BodyCf = New CReportField
   
   j = 0
   SumWidth = 0
   For Each HeadCf In m_Headers1.Fields
      j = j + 1
      If j = 1 Then
         TempStr = Txt1
      ElseIf j = 2 Then
         TempStr = Txt2
      ElseIf j = 5 Then
         TempStr = Txt3
      ElseIf j = 6 Then
         TempStr = FormatNumberToNull(Tot(j), Val(mcolParam("DECIMAL_AMOUNT")))
      ElseIf j = 7 Then
         TempStr = FormatNumberToNull(Tot(j), Val(mcolParam("DECIMAL_AMOUNT")))
      ElseIf j = 8 Then
         TempStr = FormatNumberToNull(Tot(j - 1) + Tot(j - 2), Val(mcolParam("DECIMAL_AMOUNT")))
      ElseIf j > 8 Then
         TempStr = FormatNumberToNull(Tot(j), Val(mcolParam("DECIMAL_AMOUNT")))
      Else
         TempStr = ""
      End If
      
      If j = 3 Or j = 4 Or j = 5 Then
         SumWidth = SumWidth + HeadCf.Width
         If j = 5 Then
            Call BodyCf.SetFieldValue(SumWidth, "<", TempStr, "<")
            Call Details.AddField(BodyCf)
         End If
      Else
         Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.Align, TempStr, HeadCf.BodyAlign)
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
Private Function CheckShow(SupCode As String) As Boolean
    If (SupplierGroupColl.Count <= 0) Then
        CheckShow = True
        Exit Function
    Else
        Dim TempSupplierGroup As CSupplierGroup
        Set TempSupplierGroup = GetSupplierGroup(SupplierGroupColl, SupCode)
        If (TempSupplierGroup Is Nothing) Then
            CheckShow = False
        Else
            CheckShow = True
        End If
    End If
End Function
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
      Else
         Sg.SUPPLIER_NAME = Apm.SUPNAM
      End If
   Next Apm
End Sub



