VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CReportAR011_1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'User Group Report
Implements CReportInterface

Private Const MODULE_NAME = "CReportAR011_1"
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
Private Maxdate As Date
Private Mindate As Date
Private TempWidth As Double
Private TempWidth2 As Double
Private TempCount As Long

Private m_Headers0 As CFieldList
Private m_Headers1 As CFieldList
Private m_Headers1_1 As CFieldList
Private m_Headers1_2 As CFieldList
Private m_Headers1_3 As CFieldList
Private m_Headers1_4 As CFieldList
Private m_Details1 As CFieldList
Private m_Details2 As CFieldList
Private m_Ticket As Collection
Private m_TicketCheckDate As Collection
Private m_TicketAmount As Collection
Private m_TicketReceive As Collection
Private m_TicketClear As Collection
Private m_BudgetTicket As Collection
Private m_BudgetTicketReceive As Collection
Private m_BudgetTicketClear As Collection
Private m_TicketData As Collection
Private m_BankCredit As Collection
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
   Set m_Ticket = New Collection
   Set m_TicketCheckDate = New Collection
   Set m_TicketAmount = New Collection
   Set m_TicketReceive = New Collection
   Set m_TicketClear = New Collection
   Set m_TicketData = New Collection
   Set m_BudgetTicket = New Collection
   Set m_BudgetTicketReceive = New Collection
   Set m_BudgetTicketClear = New Collection
   Set m_BankCredit = New Collection
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
   Set m_Ticket = Nothing
   Set m_TicketCheckDate = Nothing
   Set m_TicketAmount = Nothing
   Set m_TicketReceive = Nothing
   Set m_TicketClear = Nothing
   Set m_TicketData = Nothing
   Set m_BudgetTicket = Nothing
   Set m_BudgetTicketReceive = Nothing
   Set m_BudgetTicketClear = Nothing
   Set m_BankCredit = Nothing
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
Dim Check As Double
Dim TempStr1 As String
Dim TempStr1_1 As String
Dim Cf As CReportField
Dim ExportWidth As Long
Dim HeadCf As CReportField
Dim j As Long
Dim TempStr As String
Dim TempRs As ADODB.Recordset
Dim Amt As Double
Dim TempFromdate As Date
Dim TempToDate As Date
Dim TempWidthHeader As Double
Dim FirstDate As Date
Dim LastDate As Date

   tbdT = Vsp.TableBorder
   blnBold = Vsp.FontBold
   iSize = Vsp.FontSize
   
   Vsp.FontSize = HEADER_SIZE
   Vsp.FontBold = True
   Vsp.TableBorder = tbAll
   
   Call m_Headers0.ClearField
   Call m_Headers1.ClearField
   
   Set Cf = New CReportField
   
   TempWidth = 0
   TempWidthHeader = 0
   
   Call Cf.SetFieldValue(2, "^", "วันที่", "^")
   Call m_Headers1.AddField(Cf)
   TempWidthHeader = TempWidthHeader + 2
   
   Call Cf.SetFieldValue(2, "^", "ตั๋วเลขที่", "^")
   Call m_Headers1.AddField(Cf)
    TempWidthHeader = TempWidthHeader + 2
   
   Call Cf.SetFieldValue(2, "^", "INVOICE", "^")
   Call m_Headers1.AddField(Cf)
    TempWidthHeader = TempWidthHeader + 2
   
   Call Cf.SetFieldValue(2, "^", "วันที่", "^")
   Call m_Headers1.AddField(Cf)
    TempWidthHeader = TempWidthHeader + 2
   
   Call Cf.SetFieldValue(TempWidthHeader, "^", "ข้อมูลตั๋ว", "^")
   Call m_Headers0.AddField(Cf)
   TempWidthHeader = 0
   
   Call Cf.SetFieldValue(2, "^", "100%", ">")
   Call m_Headers1.AddField(Cf)
   TempWidthHeader = TempWidthHeader + 2
   
   Call Cf.SetFieldValue(2, "^", "" & m_BankCredit.ITEM(1).BANK_GET_AMOUNT & "%", ">")
   Call m_Headers1.AddField(Cf)
   TempWidthHeader = TempWidthHeader + 2
   
   Call Cf.SetFieldValue(2, "^", "" & 100 - (m_BankCredit.ITEM(1).BANK_GET_AMOUNT) & "%", ">")
   Call m_Headers1.AddField(Cf)
   TempWidthHeader = TempWidthHeader + 2
   
   Call Cf.SetFieldValue(2, "^", "รับคืน " & m_BankCredit.ITEM(1).BANK_GET_AMOUNT & "%", ">")
   Call m_Headers1.AddField(Cf)
   TempWidthHeader = TempWidthHeader + 2
   
   Call Cf.SetFieldValue(TempWidthHeader, "^", "จำนวนเงินตามตั๋ว", ">")
   Call m_Headers0.AddField(Cf)
   TempWidthHeader = 0
   
   Call Cf.SetFieldValue(2, "^", "วงเงินคงเหลือ", ">")
   Call m_Headers1.AddField(Cf)
   TempWidthHeader = TempWidthHeader + 2

   Call Cf.SetFieldValue(2, "^", "ภาระหนี้คงเหลือ", ">")
   Call m_Headers1.AddField(Cf)
   TempWidthHeader = TempWidthHeader + 2

   Check = FormatNumber(m_BankCredit.ITEM(1).BANK_INTEREST)
   If Check > 0 Then
      Call Cf.SetFieldValue(2, "^", "IN TEREST " & m_BankCredit.ITEM(1).BANK_INTEREST & "%", ">")
   Else
      Call Cf.SetFieldValue(2, "^", "IN TEREST ", ">")
   End If
   Call m_Headers1.AddField(Cf)
   TempWidthHeader = TempWidthHeader + 2

   Call Cf.SetFieldValue(2, "^", "ค่าธรรมเนียม", ">")
   Call m_Headers1.AddField(Cf)
   TempWidthHeader = TempWidthHeader + 2

   Call Cf.SetFieldValue(2, "^", "ปมก.รับ " & 100 - (m_BankCredit.ITEM(1).BANK_GET_AMOUNT) & "%", ">")
   Call m_Headers1.AddField(Cf)
   TempWidthHeader = TempWidthHeader + 2

   Call Cf.SetFieldValue(2, "^", "ปมก.รับ " & m_BankCredit.ITEM(1).BANK_GET_AMOUNT & "%", ">")
   Call m_Headers1.AddField(Cf)
   TempWidthHeader = TempWidthHeader + 2

   Call Cf.SetFieldValue(1, "^", "ระยะเวลาแพ็ค", "^")
   Call m_Headers1.AddField(Cf)
   TempWidthHeader = TempWidthHeader + 1

   Call Cf.SetFieldValue(2, "^", "วันที่ตั๋วครบกำหนด", "^")
   Call m_Headers1.AddField(Cf)
   TempWidthHeader = TempWidthHeader + 2

   Call Cf.SetFieldValue(2, "^", "วันที่หน้าเช็ค", "^")
   Call m_Headers1.AddField(Cf)
   TempWidthHeader = TempWidthHeader + 2

   Call Cf.SetFieldValue(2, "^", "วันที่เก็บเช็ค", "^")
   Call m_Headers1.AddField(Cf)
   TempWidthHeader = TempWidthHeader + 2
   
   Call Cf.SetFieldValue(TempWidthHeader, "^", "", ">")
   Call m_Headers0.AddField(Cf)
   
   Call GetFirstLastDate(mcolParam("FROM_DOC_DATE"), FirstDate, LastDate)
   Mindate = FirstDate
   TempFromdate = Mindate
   TempToDate = Maxdate
   TempCount = 18
   While (TempFromdate <= TempToDate)
   Call Cf.SetFieldValue(2, "^", IntToShortThaiMonth(Month(TempFromdate)) & " " & Right((Year(TempFromdate) + 543), 2), ">")
   Call m_Headers1.AddField(Cf)
   TempWidth = TempWidth + 2
   TempCount = TempCount + 1
   TempFromdate = DateAdd("M", 1, TempFromdate)
   Wend

   Call Cf.SetFieldValue(2, "^", "รวม", ">")
   Call m_Headers1.AddField(Cf)

   TempWidth = TempWidth + 2
   Call Cf.SetFieldValue(TempWidth, "^", "ดอกเบี้ยจ่าย", "^")
   Call m_Headers0.AddField(Cf)
   
   Call Cf.SetFieldValue(2, "^", "ด/บ ล่วงหน้า", ">")
   Call m_Headers0.AddField(Cf)
   
   Call Cf.SetFieldValue(2, "^", "คงเหลือ", ">")
   Call m_Headers1.AddField(Cf)
   
   Call Cf.SetFieldValue(2, "^", "", ">")
   Call m_Headers0.AddField(Cf)
   
   Call Cf.SetFieldValue(2, "^", "หมายเหตุ", "^")
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
Dim TempFromdate As Date
Dim TempToDate As Date
Dim TempDate As Date

   mstrErrMsg = ""
   mblnHeader = True
   mblnNewPage = True
   mblnEndPage = True
   Vsp.PaperSize = pprA3
   Vsp.ORIENTATION = orLandscape
   Vsp.MarginBottom = 700
   Vsp.MarginFooter = 700
   Vsp.MarginHeader = 700
   Vsp.MarginLeft = 700
   Vsp.MarginRight = 700
   Vsp.MarginTop = 700
   Vsp.FontName = "AngsanaUPC"
   Vsp.FontSize = DETAIL_SIZE
   
   TempFromdate = mcolParam("FROM_DOC_DATE")
   TempToDate = mcolParam("TO_DOC_DATE")
   
   Call LoadBankCredit(Nothing, m_BankCredit, mcolParam("BANK_ID"), mcolParam("CUSTOMER_ID"), TempFromdate, TempDate, 1)
   If m_BankCredit.Count = 0 Then
         mstrErrMsg = "ไม่มีข้อมูลวงเงินธนาคารในวันที่ที่กำหนด   กรุณาเลือก   --จากวันที่---    ใหม่อีกครั้ง"
         Exit Function
   End If
   
   Call LoadTicket(Nothing, m_Ticket, m_TicketCheckDate, Maxdate, Mindate, TempFromdate, TempToDate, mcolParam("CUSTOMER_ID"), mcolParam("BANK_ID"))
   Call LoadTicketReceive(Nothing, m_TicketReceive, m_TicketCheckDate, TempFromdate, TempToDate, mcolParam("CUSTOMER_ID"), mcolParam("BANK_ID"))
   Call LoadTicketClear(Nothing, m_TicketClear, TempFromdate, TempToDate, mcolParam("CUSTOMER_ID"), mcolParam("BANK_ID"))

   If mcolParam("SHOW_BUDGET") = "Y" Then   'ประมาณการ
      Call LoadBudgetTicket(Nothing, m_BudgetTicket, TempFromdate, TempToDate, mcolParam("CUSTOMER_ID"), mcolParam("BANK_ID"))
      Call LoadBudgetTicketReceive(Nothing, m_BudgetTicketReceive, TempFromdate, TempToDate, mcolParam("CUSTOMER_ID"), mcolParam("BANK_ID"))
      Call LoadBudgetTicketClear(Nothing, m_BudgetTicketClear, TempFromdate, TempToDate, mcolParam("CUSTOMER_ID"), mcolParam("BANK_ID"))
   End If
   
   Call LoadTicketData(Nothing, m_TicketData, m_Ticket, m_TicketReceive, m_TicketClear, m_BudgetTicket, m_BudgetTicketReceive, m_BudgetTicketClear, TempDate, TempToDate)
   Call LoadTicketAmount(Nothing, m_TicketAmount)

   FromDate = "จากวันที่ " & EmptyToString(DateToStringExtEx2(mcolParam("FROM_DOC_DATE")), "N/A")
   ToDate = " ถึง วันที่ " & EmptyToString(DateToStringExtEx2(mcolParam("TO_DOC_DATE")), "N/A")
   
   MODULE_DESC = glbCompanyName & vbCrLf & mcolParam("REPORT_NAME") & _
                                       vbCrLf & FromDate & " " & ToDate & " " & " วงเงินแพ็คตั๋ว กับ " & EmptyToString(mcolParam("BANK_NAME"), "N/A") & _
                                       "    บริษัท  " & EmptyToString(mcolParam("CUSTOMER_NAME"), "N/A")
   
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
Dim L As Long
Dim M As Long
Dim strFormat As String
Dim alngX() As Long
Dim IsOK As Boolean
Dim Amt As Double
Dim HeadCf As CReportField
Dim BodyCf As CReportField
Dim TempRs As ADODB.Recordset
Dim TempStr1 As String
Dim TempStr2 As String
Dim TempFromdate As Date
Dim TempToDate As Date
Dim TempDate As Date
Dim FromDate As Date
Dim ToDate As Date
Dim FirstDate As Date
Dim LastDate As Date
Dim Total1(100) As Double
Dim Total2(100) As Double
Dim Total3(100) As Double
Dim iCount As Double
Dim TempStr As String
Dim PrevKey1 As Double
Dim PrevKey2 As Double
Dim PrevKey3 As Double
Dim PrevKey4 As String
Dim PrevKey5 As Double
Dim PrevKey6 As Double
Dim PrevKey7 As Double
Dim PrevKey8 As Double
Dim PrevKey9 As String
Dim PrevKey10 As Date
Dim PrevKey11 As Double
Dim PrevKey12 As String
Dim TempData As CTicket
Dim Artrn As CARTrn
   
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
   
   Set BodyCf = New CReportField
   Set TempData = New CTicket
   
   FromDate = mcolParam("FROM_DOC_DATE")
   ToDate = mcolParam("TO_DOC_DATE")
   PrevKey1 = m_BankCredit.ITEM(1).BANK_AMOUNT
   PrevKey2 = m_BankCredit.ITEM(1).BANK_AMOUNT_BROUGHT
   PrevKey7 = PrevKey1 - PrevKey2
   PrevKey10 = m_BankCredit.ITEM(1).BANK_DATE_BROUGHT
   TempWidth2 = TempWidth
   K = 0
   L = 0
   M = 0

   For Each TempData In m_TicketData
         Call m_Details1.ClearField

         Set Artrn = GetARTrn(m_TicketAmount, TempData.TICKET_INVOICE)
         K = K + 1
         M = M + 1
         iCount = 0
         j = 0
          If PrevKey4 <> TempData.TICKET_NUMBER And K > 1 And Total1(5) <> 0 Then                        'รวมยอด
               Call GenerateFooter("", Total1, m_Details1)
               Call m_Details1.GetString(2, TempStr1, TempStr2)
               strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
               Vsp.FontBold = True
               Call Vsp.AddTable(strFormat, "", TempStr2)
               Vsp.FontBold = False
               For j = 1 To UBound(Total1)
                  Total1(j) = 0
               Next j
               Call m_Details1.ClearField
                j = 0
         End If
         
         If FromDate <= TempData.TICKET_DATE And ToDate >= TempData.TICKET_DATE Then
               If L = 0 Or PrevKey9 <> (Month(TempData.TICKET_DATE) & "/" & (Year(TempData.TICKET_DATE) + 543)) Then
                     If K > 1 And L > 0 Then
                           If Total2(9) < 0 Then
                              Vsp.TextColor = RGB(255, 0, 0)
                           End If

                           Call GenerateFooter("รวมส่งแพ็ค " & IntToShortThaiMonth(Month(PrevKey9)), Total2, m_Details2)                         '  รวมส่งแพ็ค
                           Call m_Details2.GetString(2, TempStr1, TempStr2)
                           strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
                           Vsp.FontBold = True
                           Call Vsp.AddTable(strFormat, "", TempStr2)
                           Vsp.FontBold = False
                           Vsp.TextColor = RGB(0, 0, 0)

                           For j = 1 To UBound(Total2)
                              Total2(j) = 0
                           Next j
                            j = 0
                     End If
               End If
         End If

         j = j + 1
         If PrevKey4 <> TempData.TICKET_NUMBER Then                                                                                                                    'วันที่
            TempStr = (Day(TempData.TICKET_DATE) & " " & IntToShortThaiMonth(Month(TempData.TICKET_DATE)) & " " & Right((Year(TempData.TICKET_DATE) + 543), 2))
         ElseIf TempData.TEMP_TYPE = "B" Or TempData.TEMP_TYPE = "C" Then
            TempStr = (Day(TempData.TICKET_DATE) & " " & IntToShortThaiMonth(Month(TempData.TICKET_DATE)) & " " & Right((Year(TempData.TICKET_DATE) + 543), 2))
         Else
            TempStr = ""
         End If
            Set HeadCf = m_Headers1.Fields(j)
            Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
            Call m_Details1.AddField(BodyCf)
         j = j + 1
         If PrevKey4 <> TempData.TICKET_NUMBER And (TempData.TEMP_TYPE <> "B" And TempData.TEMP_TYPE <> "C") Then                            ' ตั๋วเลขที่
            TempStr = TempData.TICKET_NUMBER
         ElseIf TempData.MASTER_AREA = 2 Then
            TempStr = TempData.TICKET_NUMBER
         Else
            TempStr = ""
         End If
         Set HeadCf = m_Headers1.Fields(j)
         Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
         Call m_Details1.AddField(BodyCf)
         j = j + 1
         Set HeadCf = m_Headers1.Fields(j)                                                                                                      ' INVOICE
         If TempData.TICKET_AMOUNT > 0 And (TempData.TEMP_TYPE <> "B" And TempData.TEMP_TYPE <> "C") Then
            If PrevKey12 <> TempData.TICKET_INVOICE Then
               TempStr = Artrn.DOCNUM
               M = 1
            Else
               TempStr = ""
            End If
         Else
            TempStr = Artrn.DOCNUM
            M = 1
         End If
         Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
         Call m_Details1.AddField(BodyCf)
         j = j + 1
         If TempData.TEMP_TYPE <> "B" And TempData.TEMP_TYPE <> "C" And M = 1 And TempData.MASTER_AREA <> 2 Then                        ' วันที่
            TempStr = (Day(Artrn.DOCDAT) & " " & IntToShortThaiMonth(Month(Artrn.DOCDAT)) & " " & Right((Year(Artrn.DOCDAT) + 543), 2))
         Else
            TempStr = ""
         End If
         Set HeadCf = m_Headers1.Fields(j)
         Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
         Call m_Details1.AddField(BodyCf)
         j = j + 1
         If TempData.TEMP_TYPE <> "B" And TempData.TEMP_TYPE <> "C" Then                          'จำนวนเงินตามตั๋ว 100
               If FromDate <= TempData.TICKET_DATE And ToDate >= TempData.TICKET_DATE Then
                  If TempData.TICKET_AMOUNT > 0 Then
                     Total1(j) = Total1(j) + TempData.TICKET_AMOUNT
                     Total2(j) = Total2(j) + TempData.TICKET_AMOUNT
                  Else
                     Total1(j) = Total1(j) + Artrn.AMOUNT
                     Total2(j) = Total2(j) + Artrn.AMOUNT
                  End If
               End If
               If TempData.TICKET_AMOUNT > 0 Then
                  TempStr = FormatNumber(TempData.TICKET_AMOUNT)
               Else
                  TempStr = FormatNumber(Artrn.AMOUNT)
               End If
         Else
               TempStr = ""
         End If
         Set HeadCf = m_Headers1.Fields(j)
         Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
         Call m_Details1.AddField(BodyCf)
         j = j + 1
         If TempData.TEMP_TYPE <> "B" And TempData.TEMP_TYPE <> "C" Then                                ' จำนวนเงินตามตั๋ว  80%
            If TempData.TICKET_AMOUNT > 0 Then
               PrevKey3 = TempData.TICKET_AMOUNT * (m_BankCredit.ITEM(1).BANK_GET_AMOUNT / 100)
            Else
               PrevKey3 = Artrn.AMOUNT * (m_BankCredit.ITEM(1).BANK_GET_AMOUNT / 100)
            End If
             If FromDate <= TempData.TICKET_DATE And ToDate >= TempData.TICKET_DATE Then
               Total1(j) = Total1(j) + PrevKey3
               Total2(j) = Total2(j) + PrevKey3
             End If
            TempStr = FormatNumber(PrevKey3)
         Else
            TempStr = ""
         End If
         Set HeadCf = m_Headers1.Fields(j)
         Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
         Call m_Details1.AddField(BodyCf)
         j = j + 1
         If TempData.TEMP_TYPE <> "B" And TempData.TEMP_TYPE <> "C" Then                                                 ' จำนวนเงินตามตั๋ว  20%
                If FromDate <= TempData.TICKET_DATE And ToDate >= TempData.TICKET_DATE Then
                  If TempData.TICKET_AMOUNT > 0 Then
                     Total1(j) = Total1(j) + (TempData.TICKET_AMOUNT - PrevKey3)
                     Total2(j) = Total2(j) + (TempData.TICKET_AMOUNT - PrevKey3)
                  Else
                     Total1(j) = Total1(j) + (Artrn.AMOUNT - PrevKey3)
                     Total2(j) = Total2(j) + (Artrn.AMOUNT - PrevKey3)
                  End If
               End If
               If TempData.TICKET_AMOUNT > 0 Then
                  TempStr = FormatNumber(TempData.TICKET_AMOUNT - PrevKey3)
               Else
                  TempStr = FormatNumber(Artrn.AMOUNT - PrevKey3)
               End If
         Else
               TempStr = ""
         End If
         Set HeadCf = m_Headers1.Fields(j)
         Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
         Call m_Details1.AddField(BodyCf)
          j = j + 1
         If TempData.TEMP_TYPE = "B" Then                                                                                                                   ' จำนวนเงินตามตั๋ว  รับคืน 80%
               If TempData.TICKET_AMOUNT > 0 Then
                  If TempData.TICKET_DATE_NEW > 0 Then
                     If TempData.TEMP_CHECK = "Y" Then
                        PrevKey3 = ((Artrn.AMOUNT * (m_BankCredit.ITEM(1).BANK_GET_AMOUNT / 100)) - (TempData.TICKET_AMOUNT * (TempData.TEMP_COUNT - 1)))
                     Else
                        PrevKey3 = CDbl(TempData.TICKET_AMOUNT)
                     End If
                  Else
                     PrevKey3 = CDbl(TempData.TICKET_AMOUNT) * (m_BankCredit.ITEM(1).BANK_GET_AMOUNT / 100)
                  End If
               Else
                  PrevKey3 = Artrn.AMOUNT * (m_BankCredit.ITEM(1).BANK_GET_AMOUNT / 100)
               End If
               If FromDate <= TempData.TICKET_DATE And ToDate >= TempData.TICKET_DATE Then
                  Total2(j) = Total2(j) + PrevKey3
               End If
               TempStr = FormatNumber(PrevKey3)
         Else
               TempStr = ""
         End If
         Set HeadCf = m_Headers1.Fields(j)
         Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
         Call m_Details1.AddField(BodyCf)
         j = j + 1
         If PrevKey9 <> (Month(TempData.TICKET_DATE) & "/" & (Year(TempData.TICKET_DATE) + 543)) And K <> 1 Then
               PrevKey2 = PrevKey1 - PrevKey7
               PrevKey10 = DateAdd("M", -1, TempData.TICKET_DATE)
               Call GetFirstLastDate(PrevKey10, FirstDate, LastDate)
               PrevKey10 = LastDate
         End If
            If TempData.TEMP_TYPE = "B" Then                                                                                                              ' วงเงินคงเหลือ
               PrevKey7 = PrevKey7 + PrevKey3
            ElseIf TempData.TEMP_TYPE = "C" Then
            Else
               PrevKey7 = PrevKey7 - PrevKey3
            End If
            If FromDate <= TempData.TICKET_DATE And ToDate >= TempData.TICKET_DATE Then
                     Total2(j) = PrevKey7
            End If
'            If PrevKey7 < 0 Then
'               TempX = Vsp.CurrentX
'               TempY = Vsp.CurrentY
'            End If
         Set HeadCf = m_Headers1.Fields(j)
         TempStr = FormatNumber(PrevKey7)
         Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
         Call m_Details1.AddField(BodyCf)
         j = j + 1
          If FromDate <= TempData.TICKET_DATE And ToDate >= TempData.TICKET_DATE Then
              Total2(j) = PrevKey1 - PrevKey7                                                                                                                         ' ภาระหนี้คงเหลือ
         End If
         Set HeadCf = m_Headers1.Fields(j)
         TempStr = FormatNumber(PrevKey1 - PrevKey7)
         Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
         Call m_Details1.AddField(BodyCf)
         j = j + 1
         If TempData.TEMP_TYPE <> "B" And TempData.TEMP_TYPE <> "C" And TempData.MASTER_AREA <> 2 Then                                              ' INTEREST
               If Len(TempData.TICKET_INTEREST) > 0 Then
                  PrevKey5 = ((PrevKey3 * TempData.TICKET_INTEREST * DateDiff("D", TempData.TICKET_DATE, TempData.TICKET_DATE_DUE)) / 36500)
               ElseIf Len(m_BankCredit.ITEM(1).BANK_INTEREST) > 0 Then
                  PrevKey5 = ((PrevKey3 * m_BankCredit.ITEM(1).BANK_INTEREST * DateDiff("D", TempData.TICKET_DATE, TempData.TICKET_DATE_DUE)) / 36500)
               Else
                  m_BankCredit.ITEM(1).BANK_INTEREST = 0
                  PrevKey5 = 0
               End If
                If FromDate <= TempData.TICKET_DATE And ToDate >= TempData.TICKET_DATE Then
                     Total1(j) = Total1(j) + PrevKey5
                     Total2(j) = Total2(j) + PrevKey5
               End If
               If PrevKey5 > 0 Then
                  TempStr = FormatNumber(PrevKey5)
               Else
                  TempStr = ""
               End If
         Else
               TempStr = ""
         End If
         Set HeadCf = m_Headers1.Fields(j)
         Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
         Call m_Details1.AddField(BodyCf)
          j = j + 1
         If TempData.TEMP_TYPE <> "B" And TempData.TEMP_TYPE <> "C" And FromDate <= TempData.TICKET_DATE And ToDate >= TempData.TICKET_DATE And TempData.MASTER_AREA <> 2 Then
               If m_BankCredit.ITEM(1).BANK_FEE_TYPE = 1 And PrevKey4 <> TempData.TICKET_NUMBER Then      ' ค่าธรรมเนียม
                  PrevKey8 = m_BankCredit.ITEM(1).BANK_FEE_AMOUNT
                  TempStr = FormatNumber(PrevKey8)
                  Total1(j) = Total1(j) + PrevKey8
                  Total2(j) = Total2(j) + PrevKey8
               ElseIf m_BankCredit.ITEM(1).BANK_FEE_TYPE <> 1 Then
                  If TempData.TICKET_AMOUNT > 0 Then
                     PrevKey8 = CDbl(TempData.TICKET_AMOUNT) * (m_BankCredit.ITEM(1).BANK_FEE_AMOUNT / 100)
                  Else
                     PrevKey8 = Artrn.AMOUNT * (m_BankCredit.ITEM(1).BANK_FEE_AMOUNT / 100)
                  End If
                  TempStr = FormatNumber(PrevKey8)
                  Total1(j) = Total1(j) + PrevKey8
                  Total2(j) = Total2(j) + PrevKey8
               Else
                  TempStr = ""
               End If
         Else
               TempStr = ""
         End If
         Set HeadCf = m_Headers1.Fields(j)
         Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
         Call m_Details1.AddField(BodyCf)
         j = j + 1
          If TempData.TEMP_TYPE = "C" Then                                                                                                               ' ปมก. รับ 20%
                If FromDate <= TempData.TICKET_DATE And ToDate >= TempData.TICKET_DATE Then
                  If TempData.TICKET_AMOUNT > 0 Then
                     If TempData.TICKET_DATE_NEW > 0 Then
                        Total2(j) = Total2(j) + (Artrn.AMOUNT - (Artrn.AMOUNT * (m_BankCredit.ITEM(1).BANK_GET_AMOUNT / 100)))
                     Else
                        Total2(j) = Total2(j) + (CDbl(TempData.TICKET_AMOUNT) - (CDbl(TempData.TICKET_AMOUNT) * (m_BankCredit.ITEM(1).BANK_GET_AMOUNT / 100)))
                     End If
                  Else
                     Total2(j) = Total2(j) + (Artrn.AMOUNT - (Artrn.AMOUNT * (m_BankCredit.ITEM(1).BANK_GET_AMOUNT / 100)))
                  End If
                End If
               If TempData.TICKET_AMOUNT > 0 Then
                  If TempData.TICKET_DATE_NEW > 0 Then
                     TempStr = FormatNumber(Artrn.AMOUNT - (Artrn.AMOUNT * (m_BankCredit.ITEM(1).BANK_GET_AMOUNT / 100)))
                  Else
                     TempStr = FormatNumber(TempData.TICKET_AMOUNT - (TempData.TICKET_AMOUNT * (m_BankCredit.ITEM(1).BANK_GET_AMOUNT / 100)))
                  End If
               Else
                  TempStr = FormatNumber(Artrn.AMOUNT - (Artrn.AMOUNT * (m_BankCredit.ITEM(1).BANK_GET_AMOUNT / 100)))
               End If
          Else
               TempStr = ""
          End If
         Set HeadCf = m_Headers1.Fields(j)
         Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
         Call m_Details1.AddField(BodyCf)
         j = j + 1
         If TempData.TEMP_TYPE <> "B" And TempData.TEMP_TYPE <> "C" And TempData.MASTER_AREA <> 2 Then                                                   ' ปมก. รับ 80%
                If PrevKey4 <> TempData.TICKET_NUMBER Then
                     PrevKey11 = PrevKey3 - PrevKey8 - PrevKey5
                     TempStr = FormatNumber(PrevKey11)
                Else
                     If m_BankCredit.ITEM(1).BANK_FEE_TYPE = 1 Then
                        PrevKey11 = PrevKey3 - PrevKey5
                        TempStr = FormatNumber(PrevKey11)
                     Else
                        PrevKey11 = PrevKey3 - PrevKey8 - PrevKey5
                        TempStr = FormatNumber(PrevKey11)
                     End If
               End If
               If FromDate <= TempData.TICKET_DATE And ToDate >= TempData.TICKET_DATE Then
                     Total1(j) = Total1(j) + PrevKey11
                     Total2(j) = Total2(j) + PrevKey11
               End If
         Else
               TempStr = ""
         End If
         Set HeadCf = m_Headers1.Fields(j)
         Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
         Call m_Details1.AddField(BodyCf)
         j = j + 1
         If TempData.TEMP_TYPE <> "B" And TempData.TEMP_TYPE <> "C" And TempData.MASTER_AREA <> 2 Then                              ' ระยะเวลาแพ็ค
               TempStr = DateDiff("D", TempData.TICKET_DATE, TempData.TICKET_DATE_DUE)
         Else
               TempStr = ""
         End If
         Set HeadCf = m_Headers1.Fields(j)
         Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
         Call m_Details1.AddField(BodyCf)
         j = j + 1
         If TempData.TEMP_TYPE <> "B" And TempData.TEMP_TYPE <> "C" And TempData.MASTER_AREA <> 2 Then                         'วันที่ตั๋วครบกำหนด
               TempStr = (Day(TempData.TICKET_DATE_DUE) & " " & IntToShortThaiMonth(Month(TempData.TICKET_DATE_DUE)) & " " & Right((Year(TempData.TICKET_DATE_DUE) + 543), 2))
         Else
               TempStr = ""
         End If
         Set HeadCf = m_Headers1.Fields(j)
         Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
         Call m_Details1.AddField(BodyCf)
         j = j + 1
         If TempData.TEMP_TYPE <> "B" And TempData.TEMP_TYPE <> "C" Then                                                               'วันที่หน้า เช็ค
               TempStr = (Day(TempData.TICKET_DATE_CHECK) & " " & IntToShortThaiMonth(Month(TempData.TICKET_DATE_CHECK)) & " " & Right((Year(TempData.TICKET_DATE_CHECK) + 543), 2))
         Else
               TempStr = ""
         End If
         Set HeadCf = m_Headers1.Fields(j)
         Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
         Call m_Details1.AddField(BodyCf)
         j = j + 1
         If TempData.TEMP_TYPE <> "B" And TempData.TEMP_TYPE <> "C" And TempData.MASTER_AREA <> 2 Then                                     'วันที่เก็บ เช็ค
               TempStr = (Day(TempData.TICKET_DATE_GET_CHECK) & " " & IntToShortThaiMonth(Month(TempData.TICKET_DATE_GET_CHECK)) & " " & Right((Year(TempData.TICKET_DATE_GET_CHECK) + 543), 2))
         Else
               TempStr = ""
         End If
         Set HeadCf = m_Headers1.Fields(j)
         Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
         Call m_Details1.AddField(BodyCf)
         
           'ดอกเบี้ยจ่าย
         If TempData.TEMP_TYPE <> "B" And TempData.TEMP_TYPE <> "C" And FromDate <= TempData.TICKET_DATE And ToDate >= TempData.TICKET_DATE And TempData.MASTER_AREA <> 2 Then
               TempFromdate = TempData.TICKET_DATE
               TempToDate = TempData.TICKET_DATE_CHECK
               i = 1
               While (TempFromdate <= TempToDate)
                        If i = 1 Then
                           Call GetFirstLastDate(TempData.TICKET_DATE, FirstDate, LastDate)
                           PrevKey6 = DateDiff("D", TempData.TICKET_DATE, LastDate)
                        Else
                           Call GetFirstLastDate(TempFromdate, FirstDate, LastDate)
                           If LastDate >= TempData.TICKET_DATE_CHECK Then
                              LastDate = TempData.TICKET_DATE_CHECK
                              PrevKey6 = DateDiff("D", FirstDate, LastDate) + 1
                           Else
                              PrevKey6 = DateDiff("D", FirstDate, LastDate) + 1
                           End If
                        End If
                        PrevKey6 = (PrevKey5 / DateDiff("D", TempData.TICKET_DATE, TempData.TICKET_DATE_DUE)) * PrevKey6
                        iCount = iCount + PrevKey6
                        j = j + 1
                         If FromDate <= TempData.TICKET_DATE And ToDate >= TempData.TICKET_DATE Then
                              Total1(j) = Total1(j) + PrevKey6
                              Total2(j) = Total2(j) + PrevKey6
                        End If
                        TempStr = FormatNumber(PrevKey6)
                        Set HeadCf = m_Headers1.Fields(j)
                        Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
                        Call m_Details1.AddField(BodyCf)
                        TempFromdate = DateAdd("M", 1, FirstDate)
                        i = i + 1
               Wend
         End If
         If j < TempCount Then
            For i = j + 1 To TempCount
               j = j + 1
               Set HeadCf = m_Headers1.Fields(j)                                                                                                                       'เพิ่มตารางส่วนที่ขาดจากดอกเบี้ยจ่าย
               TempStr = ""
               Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
               Call m_Details1.AddField(BodyCf)
            Next i
         End If

         j = j + 1
         If TempData.TEMP_TYPE <> "B" And TempData.TEMP_TYPE <> "C" And TempData.MASTER_AREA <> 2 Then                       'รวมดอกเบี้ยจ่าย
            If FromDate <= TempData.TICKET_DATE And ToDate >= TempData.TICKET_DATE Then
               Total1(j) = Total1(j) + iCount
               Total2(j) = Total2(j) + iCount
            End If
            TempStr = FormatNumber(iCount)
         Else
            TempStr = ""
         End If
         Set HeadCf = m_Headers1.Fields(j)
         Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
         Call m_Details1.AddField(BodyCf)
         
          j = j + 1
         If TempData.TEMP_TYPE <> "B" And TempData.TEMP_TYPE <> "C" And TempData.MASTER_AREA <> 2 Then                   'ด/บ ล่วงหน้า คงเหลือ
            If FromDate <= TempData.TICKET_DATE And ToDate >= TempData.TICKET_DATE Then
            Total1(j) = Total1(j) + (PrevKey5 - iCount)
            Total2(j) = Total2(j) + (PrevKey5 - iCount)
            End If
            TempStr = FormatNumber(PrevKey5 - iCount)
         Else
            TempStr = ""
         End If
         Set HeadCf = m_Headers1.Fields(j)
         Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
         Call m_Details1.AddField(BodyCf)
         
         j = j + 1
         If TempData.TEMP_TYPE <> "B" And TempData.TEMP_TYPE <> "C" And Len(TempData.TICKET_INTEREST) > 0 Then              'หมายเหตุ
            TempStr = "ดอกเบี้ย " & TempData.TICKET_INTEREST & " %"
         ElseIf TempData.TEMP_TYPE <> "B" And TempData.TEMP_TYPE <> "C" And TempData.CHECK_RECEIVED > 0 Then
            TempStr = "ระบุจน.หน้าเช็ค"
         Else
            TempStr = ""
         End If
         Set HeadCf = m_Headers1.Fields(j)
         Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
         Call m_Details1.AddField(BodyCf)
         
         If FromDate <= TempData.TICKET_DATE And ToDate >= TempData.TICKET_DATE Then
               If L = 0 Or PrevKey9 <> (Month(TempData.TICKET_DATE) & "/" & (Year(TempData.TICKET_DATE) + 543)) Then
                     If mcolParam("SHOW_EXCEL") = "Y" Then
                              If Vsp.CurrentPage = 1 Then
                                Call printHeader
                             End If
                       End If

                  If (PrevKey1 - PrevKey2) < 0 Then
                     Vsp.TextColor = RGB(255, 0, 0)
                  End If

                  Call GenerateHeader("ยอดยกมา ณ " & PrevKey10, "วงเงินรับ " & FormatNumber(m_BankCredit.ITEM(1).BANK_AMOUNT), FormatNumber(PrevKey1 - PrevKey2), FormatNumber(PrevKey2), TempWidth, m_Details2)
                  Call m_Details2.GetString(2, TempStr1, TempStr2)
                  strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
                  Vsp.FontBold = True
                  Call Vsp.AddTable(strFormat, "", TempStr2)
                  Vsp.FontBold = False
                  Vsp.TextColor = RGB(0, 0, 0)
                  
                     TempDate = ToDate
                     L = L + 1
               End If
               If PrevKey7 < 0 Then
                     Vsp.TextColor = RGB(255, 0, 0)
               End If
               Call m_Details1.GetString(1, TempStr1, TempStr2)
               strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
               Call Vsp.AddTable(strFormat, "", TempStr2)
               Vsp.TextColor = RGB(0, 0, 0)
         End If
         
         PrevKey9 = (Month(TempData.TICKET_DATE) & "/" & (Year(TempData.TICKET_DATE) + 543))
         PrevKey4 = TempData.TICKET_NUMBER
         PrevKey12 = TempData.TICKET_INVOICE
         
   Next TempData
    
         If Total1(5) <> 0 Then
               Call GenerateFooter("", Total1, m_Details1)                                                                                      '  รวมยอดบรรทัดสุดท้าย
               Call m_Details1.GetString(2, TempStr1, TempStr2)
               strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
               Vsp.FontBold = True
               Call Vsp.AddTable(strFormat, "", TempStr2)
               Vsp.FontBold = False
         End If
         
         If Total2(9) < 0 Then
                  Vsp.TextColor = RGB(255, 0, 0)
         End If
         
         Call GenerateFooter("รวมส่งแพ็ค " & IntToShortThaiMonth(Month(TempDate)), Total2, m_Details1)                         '  รวมส่งแพ็ค
         Call m_Details1.GetString(2, TempStr1, TempStr2)
         strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
         Vsp.FontBold = True
         Call Vsp.AddTable(strFormat, "", TempStr2)
         Vsp.FontBold = False
         Vsp.TextColor = RGB(0, 0, 0)
         
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
Private Sub GenerateFooter(Txt1 As String, Tot() As Double, Details As CFieldList)
Dim HeadCf As CReportField
Dim BodyCf As CReportField
Dim j As Long
Dim i As Long
Dim TempStr As String

   Call Details.ClearField
   Set BodyCf = New CReportField
   j = 0
   For Each HeadCf In m_Headers1.Fields
      j = j + 1
      If j = 1 Then
         Set HeadCf = m_Headers0.Fields(1)
         TempStr = Txt1
         Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.Align, TempStr, HeadCf.BodyAlign)
         Call Details.AddField(BodyCf)
      End If
      If j >= 5 Then
         TempStr = FormatNumber(Tot(j))
         If TempStr = "0.00" Then
            TempStr = ""
         End If
         Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.Align, TempStr, HeadCf.BodyAlign)
         Call Details.AddField(BodyCf)
      End If
   Next HeadCf
      
   Set BodyCf = Nothing
End Sub
Private Sub GenerateHeader(Txt1 As String, Txt2 As String, Txt3 As String, Txt4 As String, TempWidth As Double, Details As CFieldList)
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
         TempStr = ""
         Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.Align, TempStr, HeadCf.BodyAlign)
         Call Details.AddField(BodyCf)
      ElseIf j = 2 Then
         TempStr = Txt1 & "   " & Txt2
         Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.Align, TempStr, HeadCf.BodyAlign)
         Call Details.AddField(BodyCf)
      ElseIf j = 3 Then
         Set HeadCf = m_Headers1.Fields(9)
         TempStr = Txt3
         Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.Align, TempStr, HeadCf.BodyAlign)
         Call Details.AddField(BodyCf)
         
         Set HeadCf = m_Headers1.Fields(10)
         TempStr = Txt4
         Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.Align, TempStr, HeadCf.BodyAlign)
         Call Details.AddField(BodyCf)
      ElseIf j = 4 Then
         TempStr = ""
         Call BodyCf.SetFieldValue(15 + TempWidth, HeadCf.Align, TempStr, HeadCf.BodyAlign)
         Call Details.AddField(BodyCf)
      Else
         TempStr = ""
         Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.Align, TempStr, HeadCf.BodyAlign)
         Call Details.AddField(BodyCf)
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
   
   Vsp.TextColor = RGB(0, 0, 0)
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
