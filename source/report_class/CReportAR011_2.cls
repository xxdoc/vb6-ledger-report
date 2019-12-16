VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CReportAR011_2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'User Group Report
Implements CReportInterface

Private Const MODULE_NAME = "CReportAR011_2"
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
Private TempFromdate As Date
Private TempToDate As Date

Private m_Headers0 As CFieldList
Private m_Headers1 As CFieldList
Private m_Headers2 As CFieldList
Private m_Headers1_1 As CFieldList
Private m_Headers1_2 As CFieldList
Private m_Headers1_3 As CFieldList
Private m_Headers1_4 As CFieldList
Private m_Details1 As CFieldList
Private m_Details2 As CFieldList
Private m_Left As Double
Private m_BankCredit As Collection
Private m_Ticket As Collection
Private m_TicketClear As Collection
Private m_TicketAmount As Collection
Private m_BudgetTicket As Collection
Private m_BudgetTicketClear As Collection
Private Sub Class_Initialize()
   Set mcolParam = New Collection
   Set m_TempCol = New Collection
   
   Set m_Headers0 = New CFieldList
   Set m_Headers1 = New CFieldList
   Set m_Headers2 = New CFieldList
   Set m_Headers1_1 = New CFieldList
   Set m_Headers1_2 = New CFieldList
   Set m_Headers1_3 = New CFieldList
   Set m_Headers1_4 = New CFieldList
   Set m_Details1 = New CFieldList
   Set m_Details2 = New CFieldList
   Set m_BankCredit = New Collection
   Set m_Ticket = New Collection
   Set m_TicketClear = New Collection
   Set m_TicketAmount = New Collection
   Set m_BudgetTicket = New Collection
   Set m_BudgetTicketClear = New Collection
End Sub
Private Sub Class_Terminate()
   Call ClearParam
   Set mcolParam = Nothing
   Set Vsp = Nothing
   Set m_TempCol = Nothing
   
   Set m_Headers0 = Nothing
   Set m_Headers1 = Nothing
   Set m_Headers2 = Nothing
   Set m_Details1 = Nothing
   Set m_Headers1_1 = Nothing
   Set m_Headers1_2 = Nothing
   Set m_Headers1_3 = Nothing
   Set m_Headers1_4 = Nothing
   Set m_Details2 = Nothing
   Set m_BankCredit = Nothing
   Set m_Ticket = Nothing
   Set m_TicketClear = Nothing
   Set m_TicketAmount = Nothing
   Set m_BudgetTicket = Nothing
   Set m_BudgetTicketClear = Nothing
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
Dim TempDate As Date

   tbdT = Vsp.TableBorder
   blnBold = Vsp.FontBold
   iSize = Vsp.FontSize
   
   Vsp.FontSize = HEADER_SIZE
   Vsp.FontBold = True
   Vsp.TableBorder = tbAll
   
   Call m_Headers0.ClearField
   Call m_Headers1.ClearField
   Call m_Headers2.ClearField
   
   Set Cf = New CReportField
      
   Call Cf.SetFieldValue(4, "^", " ", "^")
   Call m_Headers2.AddField(Cf)
      
   Call Cf.SetFieldValue(2, "^", "ธนาคาร", "^")
   Call m_Headers0.AddField(Cf)
   
   Call Cf.SetFieldValue(2, "^", "บริษัท", "^")
   Call m_Headers0.AddField(Cf)
   
   Call Cf.SetFieldValue(2, "^", " ", "<")
   Call m_Headers1.AddField(Cf)
   
   Call Cf.SetFieldValue(2, "^", " ", "<")
   Call m_Headers1.AddField(Cf)
   
   TempDate = TempFromdate
   While (TempDate <= TempToDate)
         Call Cf.SetFieldValue(8, "^", IntToShortThaiMonth(Month(TempDate)) & " " & Right((Year(TempDate) + 543), 2), "^")
         Call m_Headers2.AddField(Cf)
   
         Call Cf.SetFieldValue(2, "^", "20%", "^")
         Call m_Headers0.AddField(Cf)
      
         Call Cf.SetFieldValue(2, "^", " ", ">")
         Call m_Headers1.AddField(Cf)
      
         Call Cf.SetFieldValue(4, "^", "80%", "^")
         Call m_Headers0.AddField(Cf)
      
         Call Cf.SetFieldValue(2, "^", "มีบิล", ">")
         Call m_Headers1.AddField(Cf)
      
         Call Cf.SetFieldValue(2, "^", "ปมก.", ">")
         Call m_Headers1.AddField(Cf)
         
         Call Cf.SetFieldValue(2, "^", "รวม", "^")
         Call m_Headers0.AddField(Cf)
      
         Call Cf.SetFieldValue(2, "^", " ", ">")
         Call m_Headers1.AddField(Cf)
         
         TempDate = DateAdd("M", 1, TempDate)
   Wend

   Set Cf = Nothing
            
   Call m_Headers2.GetString(1, TempStr1, TempStr1_1)
   strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
   strPrint = TempStr1_1
   Call Vsp.AddTable(strFormat, "", strPrint)
            
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
Dim FirstDate As Date
Dim LastDate As Date
Dim FirstDate2 As Date
Dim LastDate2 As Date
   
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

   Call GetFirstLastDate(mcolParam("FROM_DOC_DATE"), FirstDate, LastDate)
   Call GetFirstLastDate(mcolParam("TO_DOC_DATE"), FirstDate2, LastDate2)
   TempFromdate = FirstDate
   TempToDate = LastDate2
   
   Call LoadBankCredit(Nothing, m_BankCredit, , , , , 2)
   Call LoadTicket(Nothing, m_Ticket, , , , TempFromdate, TempToDate)
   Call LoadTicketClear(Nothing, m_TicketClear, TempFromdate, TempToDate)
   If mcolParam("SHOW_BUDGET") = "Y" Then   'ประมาณการ
         Call LoadBudgetTicket(Nothing, m_BudgetTicket, TempFromdate, TempToDate)
         Call LoadBudgetTicketClear(Nothing, m_BudgetTicketClear, TempFromdate, TempToDate)
   End If
   Call LoadTicketAmount(Nothing, m_TicketAmount)

   FromDate = "จากเดือน " & IntToThaiMonth(Month(mcolParam("FROM_DOC_DATE"))) & " " & Right((Year(mcolParam("FROM_DOC_DATE")) + 543), 4)
   ToDate = "ถึง เดือน " & IntToThaiMonth(Month(mcolParam("TO_DOC_DATE"))) & " " & Right((Year(mcolParam("FROM_DOC_DATE")) + 543), 4)
   
   MODULE_DESC = glbCompanyName & vbCrLf & mcolParam("REPORT_NAME") & "  " & FromDate & "  " & ToDate
   
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
Dim PrevKey1 As Long
Dim PrevKey2 As String
Dim Diff As Long
Dim TempBankCredit As CBankCredit
Dim TempTicket As CTicket
Dim TempTicketClear As CTicket
Dim TempBudgetTicket As CTicket
Dim TempBudgetTicketClear As CTicket
Dim Artrn As CARTrn

   RName = "genDoc"
'-----------------------------------------------------------------------------------------------------
'                                             Query Here
'-----------------------------------------------------------------------------------------------------
   
   Set Rs = New ADODB.Recordset
   
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
      
      If mcolParam("SHOW_EXCEL") = "Y" Then
            If Vsp.CurrentPage = 1 Then
                  Call printHeader
            End If
      End If
      
      Set TempBankCredit = New CBankCredit

      For Each TempBankCredit In m_BankCredit
         Call m_Details1.ClearField
         '20%
         Set TempTicketClear = New CTicket
         PrevKey1 = 0
         Diff = 0
         i = 3
         For Each TempTicketClear In m_TicketClear
               If TempFromdate <= TempTicketClear.TICKET_DATE And TempTicketClear.TICKET_DATE <= TempToDate And TempTicketClear.BANK_ID = TempBankCredit.BANK_ID And TempTicketClear.CUSTOMER_ID = TempBankCredit.CUSTOMER_ID Then
                     If PrevKey1 = 0 Then
                        PrevKey1 = Month(TempFromdate)
                     End If
                     If Month(TempTicketClear.TICKET_DATE) <> PrevKey1 Then
                        Diff = Month(TempTicketClear.TICKET_DATE) - PrevKey1
                        i = i + (4 * Diff)
                     End If
                     Set Artrn = GetARTrn(m_TicketAmount, TempTicketClear.TICKET_INVOICE)
                     Total1(i) = Total1(i) + (Artrn.AMOUNT * (100 - TempBankCredit.BANK_GET_AMOUNT) * 0.01)
                     Total1(i + 3) = Total1(i + 3) + (Artrn.AMOUNT * (100 - TempBankCredit.BANK_GET_AMOUNT) * 0.01)
                     Total2(i) = Total2(i) + (Artrn.AMOUNT * (100 - TempBankCredit.BANK_GET_AMOUNT) * 0.01)
                     Total2(i + 3) = Total2(i + 3) + (Artrn.AMOUNT * (100 - TempBankCredit.BANK_GET_AMOUNT) * 0.01)
                     PrevKey1 = Month(TempTicketClear.TICKET_DATE)
               End If
         Next TempTicketClear
         'ประมาณการ 20%
         Set TempBudgetTicketClear = New CTicket
         PrevKey1 = 0
         Diff = 0
         i = 3
         For Each TempBudgetTicketClear In m_BudgetTicketClear
               If TempFromdate <= TempBudgetTicketClear.TICKET_DATE And TempBudgetTicketClear.TICKET_DATE <= TempToDate And TempBudgetTicketClear.BANK_ID = TempBankCredit.BANK_ID And TempBudgetTicketClear.CUSTOMER_ID = TempBankCredit.CUSTOMER_ID Then
                     If PrevKey1 = 0 Then
                        PrevKey1 = Month(TempFromdate)
                     End If
                     If Month(TempBudgetTicketClear.TICKET_DATE) <> PrevKey1 Then
                        Diff = Month(TempBudgetTicketClear.TICKET_DATE) - PrevKey1
                        i = i + (4 * Diff)
                     End If
                     Total1(i) = Total1(i) + (TempBudgetTicketClear.TICKET_AMOUNT * (100 - TempBankCredit.BANK_GET_AMOUNT) * 0.01)
                     Total1(i + 3) = Total1(i + 3) + (TempBudgetTicketClear.TICKET_AMOUNT * (100 - TempBankCredit.BANK_GET_AMOUNT) * 0.01)
                     Total2(i) = Total2(i) + (TempBudgetTicketClear.TICKET_AMOUNT * (100 - TempBankCredit.BANK_GET_AMOUNT) * 0.01)
                     Total2(i + 3) = Total2(i + 3) + (TempBudgetTicketClear.TICKET_AMOUNT * (100 - TempBankCredit.BANK_GET_AMOUNT) * 0.01)
                     PrevKey1 = Month(TempBudgetTicketClear.TICKET_DATE)
               End If
         Next TempBudgetTicketClear
         'มีบิล 80%
         Set TempTicket = New CTicket
         PrevKey1 = 0
         Diff = 0
         i = 4
         For Each TempTicket In m_Ticket
               If TempFromdate <= TempTicket.TICKET_DATE And TempTicket.TICKET_DATE <= TempToDate And TempTicket.BANK_ID = TempBankCredit.BANK_ID And TempTicket.CUSTOMER_ID = TempBankCredit.CUSTOMER_ID Then
                     If PrevKey1 = 0 Then
                        PrevKey1 = Month(TempFromdate)
                     End If
                     If Month(TempTicket.TICKET_DATE) <> PrevKey1 Then
                        Diff = Month(TempTicket.TICKET_DATE) - PrevKey1
                        i = i + (4 * Diff)
                     End If
                     Set Artrn = GetARTrn(m_TicketAmount, TempTicket.TICKET_INVOICE)
                     Total1(i) = Total1(i) + (Artrn.AMOUNT * TempBankCredit.BANK_GET_AMOUNT * 0.01)
                     Total1(i + 2) = Total1(i + 2) + (Artrn.AMOUNT * TempBankCredit.BANK_GET_AMOUNT * 0.01)
                     Total2(i) = Total2(i) + (Artrn.AMOUNT * TempBankCredit.BANK_GET_AMOUNT * 0.01)
                     Total2(i + 2) = Total2(i + 2) + (Artrn.AMOUNT * TempBankCredit.BANK_GET_AMOUNT * 0.01)
                     PrevKey1 = Month(TempTicket.TICKET_DATE)
               End If
         Next TempTicket
         'ประมาณการ 80%
         Set TempBudgetTicket = New CTicket
         PrevKey1 = 0
         Diff = 0
         i = 5
         For Each TempBudgetTicket In m_BudgetTicket
               If TempFromdate <= TempBudgetTicket.TICKET_DATE And TempBudgetTicket.TICKET_DATE <= TempToDate And TempBudgetTicket.BANK_ID = TempBankCredit.BANK_ID And TempBudgetTicket.CUSTOMER_ID = TempBankCredit.CUSTOMER_ID Then
                     If PrevKey1 = 0 Then
                        PrevKey1 = Month(TempFromdate)
                     End If
                     If Month(TempBudgetTicket.TICKET_DATE) <> PrevKey1 Then
                        Diff = Month(TempBudgetTicket.TICKET_DATE) - PrevKey1
                        i = i + (4 * Diff)
                     End If
                     Total1(i) = Total1(i) + (TempBudgetTicket.TICKET_AMOUNT * TempBankCredit.BANK_GET_AMOUNT * 0.01)
                     Total1(i + 1) = Total1(i + 1) + (TempBudgetTicket.TICKET_AMOUNT * TempBankCredit.BANK_GET_AMOUNT * 0.01)
                     Total2(i) = Total2(i) + (TempBudgetTicket.TICKET_AMOUNT * TempBankCredit.BANK_GET_AMOUNT * 0.01)
                     Total2(i + 1) = Total2(i + 1) + (TempBudgetTicket.TICKET_AMOUNT * TempBankCredit.BANK_GET_AMOUNT * 0.01)
                     Total3(i) = Total3(i) + (TempBudgetTicket.TICKET_AMOUNT * TempBankCredit.BANK_GET_AMOUNT * 0.01)
                     PrevKey1 = Month(TempBudgetTicket.TICKET_DATE)
               End If
         Next TempBudgetTicket
         
         Call GenerateFooter("" & TempBankCredit.BANK_NAME, "" & Mid(TempBankCredit.CUSTOMER_NAME, 8, 30), Total1, m_Details1)
         Call m_Details1.GetString(2, TempStr1, TempStr2)
         strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
         Call Vsp.AddTable(strFormat, "", TempStr2)
         
            For j = 1 To UBound(Total1)
               Total1(j) = 0
            Next j
      Next TempBankCredit

      Call GenerateFooter("", "", Total2, m_Details1)
      Call m_Details1.GetString(2, TempStr1, TempStr2)
      strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
      Vsp.FontBold = True
      Call Vsp.AddTable(strFormat, "", TempStr2)
      Vsp.FontBold = False
      
      i = 3
      For j = 1 To UBound(Total2)
            If i = j Then
            Total3(j + 1) = Total2(j) + Total2(j + 1)
            i = i + 4
            End If
      Next j

      Call GenerateFooter("", "", Total3, m_Details1)
      Call m_Details1.GetString(2, TempStr1, TempStr2)
      strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
      Vsp.FontBold = True
      Call Vsp.AddTable(strFormat, "", TempStr2)
      Vsp.FontBold = False

   Set BodyCf = Nothing
   Vsp.EndDoc
   Call CloseExportFile(Vsp)
   
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
Private Sub GenerateFooter(Txt1 As String, Txt2 As String, Tot() As Double, Details As CFieldList)
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
      Else
         TempStr = FormatNumberToNull(Tot(j))
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
