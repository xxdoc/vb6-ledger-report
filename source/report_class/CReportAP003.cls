VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CReportAP003"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'User Group Report
Implements CReportInterface

Private Const MODULE_NAME = "CReportAP003"
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
Private m_MaxMins As Collection
Private m_ApSupAmounts As Collection
Private m_ApSupPaids As Collection
Private m_Apt2 As CApTrn

Private m_ImageWidth As Double
Private m_ImageHeight As Double
Private m_LedgerReport As Collection
Private m_TempCollection As Collection
Private m_TempAmount As Double

Public m_Cr As Collection
'Public OKClick As Boolean

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

Private Function GenerateReceiptText() As String
Dim Ri As CAPRcIt
Dim TempText As String

   TempText = ""
   For Each Ri In m_Apt2.ReceiptItems
      TempText = TempText & " " & Ri.DOCNUM
   Next Ri
   GenerateReceiptText = TempText
End Function

Private Sub GeneratePaidVocherHeader(Border As TableBorderSettings, alngX() As Long, HeaderOffset)
Dim Rc As CAPRcpCq
Dim Ri As CAPRcIt
Dim Bkt As CBkTrn
Dim Jnl As CGLJnl
Dim OldFont As Double

   Vsp.FontBold = True
   Vsp.FontSize = 14
   
   Vsp.StartTable
   Vsp.TableCell(tcCols) = 2
   Vsp.TableCell(tcRows) = 1
    Vsp.TableCell(tcRowHeight, 1) = 1.75 * 567 + HeaderOffset
    Vsp.TableCell(tcColWidth, 1, 1) = "17cm"
    Vsp.TableCell(tcColWidth, 1, 2) = "2.5cm"
    Vsp.TableCell(tcAlign, 1, 2) = taCenterBottom
    Vsp.TableCell(tcFontSize, 1, 2) = 16
    Vsp.TableCell(tcText, 1, 2) = " " & m_Apt2.DOCNUM
    
    Vsp.TableBorder = Border
    Vsp.EndTable
    
   Vsp.StartTable
   Vsp.TableCell(tcCols) = 2
   Vsp.TableCell(tcRows) = 1
    Vsp.TableCell(tcRowHeight) = 0.7 * 567
    Vsp.TableCell(tcColWidth, 1, 1) = "14.5cm"
    Vsp.TableCell(tcColWidth, 1, 2) = "5cm"
    
    Vsp.TableCell(tcAlign, 1, 2) = taLeftBottom
    Vsp.TableCell(tcText, 1, 2) = DateToStringExtEx2(m_Apt2.DOCDAT)
    Vsp.TableBorder = Border
    Vsp.EndTable
    
   Vsp.StartTable
   Vsp.TableCell(tcCols) = 3
   Vsp.TableCell(tcRows) = 1
    Vsp.TableCell(tcRowHeight) = 0.7 * 567
    Vsp.TableCell(tcColWidth, 1, 1) = "2.5cm"
    Vsp.TableCell(tcColWidth, 1, 2) = "12cm"
    Vsp.TableCell(tcColWidth, 1, 3) = "5cm"
    
    Vsp.TableCell(tcAlign, 1, 2) = taLeftBottom
    Vsp.TableCell(tcText, 1, 2) = m_Apt2.PRENAM & " " & m_Apt2.SUPNAM
    Vsp.TableCell(tcFontBold, 1, 3) = True
    Vsp.TableCell(tcFontSize, 1, 3) = 18
    Vsp.TableCell(tcAlign, 1, 3) = taCenterMiddle
    Vsp.TableCell(tcText, 1, 3) = "บันทึกบัญชีแล้ว"

    Vsp.TableBorder = Border
    Vsp.EndTable

   Vsp.StartTable
   Vsp.TableCell(tcCols) = 2
   Vsp.TableCell(tcRows) = 1
    Vsp.TableCell(tcRowHeight) = 0.8 * 567
    Vsp.TableCell(tcColWidth, 1, 1) = "5cm"
    Vsp.TableCell(tcColWidth, 1, 2) = "14.5cm"
    Vsp.TableCell(tcAlign, 1, 2) = taLeftBottom
    Vsp.TableCell(tcText, 1, 2) = GenerateReceiptText
    
    Vsp.TableBorder = Border
    Vsp.EndTable
    
   Vsp.StartTable
   Vsp.TableCell(tcCols) = 2
   Vsp.TableCell(tcRows) = 1
    Vsp.TableCell(tcRowHeight) = 0.8 * 567
    Vsp.TableCell(tcColWidth, 1, 1) = "3cm"
    Vsp.TableCell(tcColWidth, 1, 2) = "16.5cm"
    Vsp.TableCell(tcAlign, 1, 2) = taLeftBottom
    Vsp.TableCell(tcText, 1, 2) = m_Apt2.DESCRPT
    Vsp.TableBorder = Border
    Vsp.EndTable

   Vsp.StartTable
   Vsp.TableCell(tcCols) = 2
   Vsp.TableCell(tcRows) = 1
    Vsp.TableCell(tcRowHeight) = 0.8 * 567
    Vsp.TableCell(tcColWidth, 1, 1) = "3cm"
    Vsp.TableCell(tcColWidth, 1, 2) = "16.5cm"
    Vsp.TableBorder = Border
    Vsp.EndTable

   Vsp.StartTable
   Vsp.TableCell(tcCols) = 2
   Vsp.TableCell(tcRows) = 1
    Vsp.TableCell(tcRowHeight) = "0.7cm"
    Vsp.TableCell(tcColWidth, 1, 1) = "2.5cm"
    Vsp.TableCell(tcColWidth, 1, 2) = "17cm"
    Vsp.TableCell(tcAlign, 1, 2) = taLeftBottom
    Vsp.TableCell(tcText, 1, 2) = "      " & DateToStringExtEx2(m_Apt2.DUEDAT)
    Vsp.TableBorder = Border
    Vsp.EndTable

   Vsp.StartTable
   Vsp.TableCell(tcCols) = 6
   Vsp.TableCell(tcRows) = 1
    Vsp.TableCell(tcRowHeight) = 0.45 * 567   '0.85
    Vsp.TableCell(tcColWidth, 1, 1) = "4.5cm"
    Vsp.TableCell(tcColWidth, 1, 2) = "4cm"
    Vsp.TableCell(tcColWidth, 1, 3) = "1cm"
    Vsp.TableCell(tcColWidth, 1, 4) = "4cm"
    Vsp.TableCell(tcColWidth, 1, 5) = "1cm"
    Vsp.TableCell(tcColWidth, 1, 6) = "5cm"
    Vsp.TableBorder = Border
    Vsp.EndTable

   Vsp.StartTable
   Vsp.TableCell(tcCols) = 5
   Vsp.TableCell(tcRows) = 1
    Vsp.TableCell(tcRowHeight) = 1.3 * 567 '.84 * 567 เดิม
    Vsp.TableCell(tcColWidth, 1, 1) = "2.5cm"
    Vsp.TableCell(tcColWidth, 1, 2) = "3.5cm"
    
'    If mcolParam("REPORT_TYPE") = 1 Then
      Vsp.TableCell(tcAlign, 1, 2) = taRightBottom
      Vsp.TableCell(tcText, 1, 2) = FormatNumber(m_TempAmount) 'FormatNumber(m_Apt2.TOTAL)
      Vsp.TableCell(tcAlign, 1, 4) = taLeftBottom
      Vsp.TableCell(tcText, 1, 4) = ThaiBaht(m_TempAmount)
'    Else
'      Vsp.TableCell(tcAlign, 1, 2) = taRightBottom
'      Vsp.TableCell(tcText, 1, 2) = FormatNumber(Val(mcolParam("TOTAL_AMOUNT"))) 'FormatNumber(m_Apt2.TOTAL)
'      Vsp.TableCell(tcAlign, 1, 4) = taLeftBottom
'      Vsp.TableCell(tcText, 1, 4) = ThaiBaht(Val(mcolParam("TOTAL_AMOUNT")))
'    End If
    
    Vsp.TableCell(tcColWidth, 1, 3) = "1.5cm"
    Vsp.TableCell(tcColWidth, 1, 4) = "11cm"
    Vsp.TableCell(tcColWidth, 1, 5) = "1cm"
    Vsp.TableBorder = Border
    Vsp.EndTable
    
   Vsp.StartTable
   Vsp.TableCell(tcCols) = 3
   Vsp.TableCell(tcRows) = 1
    Vsp.TableCell(tcRowHeight) = 0.8 * 567
    Vsp.TableCell(tcColWidth, 1, 1) = "2cm"
    Vsp.TableCell(tcColWidth, 1, 2) = "1cm"
    If m_Apt2.BankTrans.Count <= 0 Then
      Vsp.TableCell(tcAlign, 1, 2) = taLeftBottom
'       Vsp.TableCell(tcText, 1, 2) = " X"
    End If
    Vsp.TableCell(tcColWidth, 1, 3) = "16.5cm"
    Vsp.TableBorder = Border
    Vsp.EndTable

   Vsp.StartTable
   Vsp.TableCell(tcCols) = 8
   Vsp.TableCell(tcRows) = 1
    Vsp.TableCell(tcRowHeight) = 0.8 * 567
    Vsp.TableCell(tcColWidth, 1, 1) = "2cm"
    Vsp.TableCell(tcColWidth, 1, 2) = "1cm"
    Vsp.TableCell(tcColWidth, 1, 3) = "1.5cm"
    Vsp.TableCell(tcColWidth, 1, 4) = "3.5cm"
    Vsp.TableCell(tcColWidth, 1, 5) = "1cm"
    Vsp.TableCell(tcColWidth, 1, 6) = "3.5cm"
    Vsp.TableCell(tcColWidth, 1, 7) = "3cm"
    Vsp.TableCell(tcColWidth, 1, 8) = "4cm"
    
    If m_Apt2.BankTrans.Count > 0 Then
      Vsp.TableCell(tcAlign, 1, 2) = taLeftBottom
'       Vsp.TableCell(tcText, 1, 2) = "  X"
       
       Set Bkt = m_Apt2.BankTrans(1)
      Vsp.TableCell(tcAlign, 1, 4) = taLeftBottom
       Vsp.TableCell(tcText, 1, 4) = Bkt.CHQNUM
      Vsp.TableCell(tcAlign, 1, 6) = taLeftBottom
       Vsp.TableCell(tcText, 1, 6) = DateToStringExtEx2(Bkt.CHQDAT)
      Vsp.TableCell(tcAlign, 1, 8) = taLeftBottom
       Vsp.TableCell(tcText, 1, 8) = DateToStringExtEx2(Bkt.GETDAT)
    End If
    Vsp.TableBorder = Border
    Vsp.EndTable

   Vsp.StartTable
   Vsp.TableCell(tcCols) = 7
   Vsp.TableCell(tcRows) = 1
    Vsp.TableCell(tcRowHeight) = 1.25 * 567
    Vsp.TableCell(tcColWidth, 1, 1) = "2.5cm"
    Vsp.TableCell(tcColWidth, 1, 2) = "1.3cm"
    Vsp.TableCell(tcColWidth, 1, 3) = "5.1cm"
    Vsp.TableCell(tcColWidth, 1, 4) = "0.1cm"
    Vsp.TableCell(tcColWidth, 1, 5) = "4cm"
    Vsp.TableCell(tcColWidth, 1, 6) = "1cm"
    Vsp.TableCell(tcColWidth, 1, 7) = "5.5cm"
    Vsp.TableBorder = Border
    If m_Apt2.BankTrans.Count > 0 Then
       Set Bkt = m_Apt2.BankTrans(1)
      Vsp.TableCell(tcAlign, 1, 3) = taLeftTop
       Vsp.TableCell(tcText, 1, 3) = Bkt.BNKNAM
      Vsp.TableCell(tcAlign, 1, 5) = taCenterTop
       Vsp.TableCell(tcText, 1, 5) = Bkt.BRANCH
   End If
    Vsp.EndTable
    
   Vsp.StartTable
   Vsp.TableCell(tcCols) = 1
   Vsp.TableCell(tcRows) = 1
    Vsp.TableCell(tcRowHeight) = 1.5 * 567
    Vsp.TableCell(tcColWidth, 1, 1) = "19.5cm"
    Vsp.EndTable
    
    Vsp.FontBold = False
End Sub

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

   Call GeneratePaidVocherHeader(tbNone, alngX, 0)
   
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
   Vsp.ORIENTATION = orPortrait
   Vsp.MarginBottom = 700
   Vsp.MarginFooter = 700
   Vsp.MarginHeader = 700
   Vsp.MarginLeft = 700
   Vsp.MarginRight = 700
   Vsp.MarginTop = 700
   Vsp.FontName = "AngsanaUPC"
   Vsp.FontSize = DETAIL_SIZE

   Call SetReportConfig(Vsp, mcolParam("REPORT_KEY"))
   Set m_TempCollection = mcolParam("JOURNAL")

   m_ImageWidth = 19.5 * 567
   m_ImageHeight = 26.8 * 567

'   Vsp.MarginLeft = (Vsp.PageWidth - m_ImageWidth) \ 2 - 300
'   Vsp.MarginRight = (Vsp.PageHeight - m_ImageWidth) \ 2
   
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

Private Sub GenerateBody(Border As TableBorderSettings, alngX() As Long)
Dim Jnl As CGLJnl
Dim TempStr As String
Dim IntPart As String
Dim DecimalPart As String
Dim PatchStr As String
Dim Left As Long
Dim i, j As Long
Dim SumDebit As Double
Dim SumCredit As Double
Dim TempCol As Collection
Dim Count As Long

   SumDebit = 0
   SumCredit = 0

   Vsp.FontBold = True
   Vsp.FontSize = 14

   If mcolParam("JOURNAL_TYPE") = 2 Then
      Left = 11 - m_Apt2.JournalItems.Count
      Set TempCol = m_Apt2.JournalItems
   ElseIf mcolParam("JOURNAL_TYPE") = 1 Then
      Left = 11 - m_TempCollection.Count
      Set TempCol = m_TempCollection
   End If
   Count = 0
   For Each Jnl In TempCol
      Count = Count + 1
      If Count > 11 Then
         Vsp.NewPage
         Count = 1
      End If
      Vsp.StartTable
      Vsp.TableCell(tcCols) = 6
      Vsp.TableCell(tcRows) = 1
       Vsp.TableCell(tcRowHeight, 1) = 0.84 * 567
       Vsp.TableCell(tcColWidth, 1, 1) = "8.3cm"
       Vsp.TableCell(tcColWidth, 1, 2) = "2.8cm"
       Vsp.TableCell(tcColWidth, 1, 3) = "3.1cm"
       Vsp.TableCell(tcColWidth, 1, 4) = "1cm"
       Vsp.TableCell(tcColWidth, 1, 5) = "3.1cm"
       Vsp.TableCell(tcColWidth, 1, 6) = "1.2cm"
       
       TempStr = FormatNumber(Jnl.AMOUNT)
       IntPart = Mid(TempStr, 1, InStr(1, TempStr, ".") - 1)
       DecimalPart = Mid(TempStr, InStr(1, TempStr, ".") + 1)
      If Jnl.TRNTYP = 0 Then 'Dr
         SumDebit = SumDebit + Jnl.AMOUNT
         Vsp.TableCell(tcAlign, 1, 3) = taRightMiddle
         Vsp.TableCell(tcAlign, 1, 4) = taLeftMiddle
         Vsp.TableCell(tcText, 1, 3) = IntPart
         Vsp.TableCell(tcText, 1, 4) = DecimalPart
         PatchStr = "  "
       Else 'Cr
         SumCredit = SumCredit + Jnl.AMOUNT
         Vsp.TableCell(tcAlign, 1, 5) = taRightMiddle
         Vsp.TableCell(tcAlign, 1, 6) = taLeftMiddle
         Vsp.TableCell(tcText, 1, 5) = IntPart
         Vsp.TableCell(tcText, 1, 6) = DecimalPart
         PatchStr = "         "
       End If
       Vsp.TableCell(tcText, 1, 1) = PatchStr & Jnl.ACCNAM
        Vsp.TableCell(tcText, 1, 2) = PatchStr & Jnl.ACCNUM
    
       Vsp.TableBorder = Border
       Vsp.EndTable
    Next Jnl
    
   For i = (Count + 1) To 11
      Vsp.StartTable
      Vsp.TableCell(tcCols) = 1
      Vsp.TableCell(tcRows) = 1
       Vsp.TableCell(tcRowHeight, 1) = 0.75 * 567
       Vsp.TableCell(tcColWidth, 1, 1) = "19.5cm"

       Vsp.TableBorder = Border
       Vsp.EndTable
   Next i

   Vsp.StartTable
   Vsp.TableCell(tcCols) = 5
   Vsp.TableCell(tcRows) = 1
   Vsp.TableCell(tcRowHeight, 1) = 1.6 * 567
   Vsp.TableCell(tcColWidth, 1, 1) = "11.1cm"
   Vsp.TableCell(tcColWidth, 1, 2) = "3.1cm"
   Vsp.TableCell(tcColWidth, 1, 3) = "1cm"
   Vsp.TableCell(tcColWidth, 1, 4) = "3.1cm"
   Vsp.TableCell(tcColWidth, 1, 5) = "1.2cm"
   
   Vsp.TableCell(tcAlign, 1, 1) = taLeftMiddle
   Vsp.TableCell(tcText, 1, 1) = "    " & ThaiBaht(SumDebit)
   
   TempStr = FormatNumber(SumDebit)                                                     ' sum Dr
   IntPart = Mid(TempStr, 1, InStr(1, TempStr, ".") - 1)
   DecimalPart = Mid(TempStr, InStr(1, TempStr, ".") + 1)
   Vsp.TableCell(tcAlign, 1, 2) = taRightMiddle
   Vsp.TableCell(tcText, 1, 2) = IntPart
   Vsp.TableCell(tcAlign, 1, 3) = taLeftMiddle
   Vsp.TableCell(tcText, 1, 3) = DecimalPart
   
   TempStr = FormatNumber(SumCredit)                                                   'Sum Cr
   IntPart = Mid(TempStr, 1, InStr(1, TempStr, ".") - 1)
   DecimalPart = Mid(TempStr, InStr(1, TempStr, ".") + 1)
   Vsp.TableCell(tcAlign, 1, 4) = taRightMiddle
   Vsp.TableCell(tcText, 1, 4) = IntPart
   Vsp.TableCell(tcAlign, 1, 5) = taLeftMiddle
   Vsp.TableCell(tcText, 1, 5) = DecimalPart

   Vsp.TableBorder = Border
   Vsp.EndTable
    
    Vsp.FontBold = False
End Sub

Private Sub GenerateApTranData(GlJnl As CGLJnl, Apt As CApTrn)
Dim Jnl As CGLJnl
Dim SumDebit As Double
Dim TempStr As String
Dim DescStr As String
Dim i As Long
Dim Flag As Boolean

   SumDebit = 0
   i = 0
   For Each Jnl In m_Apt2.JournalItems
'      If i = 0 Then
'         DescStr = Jnl.DESCRP
'      End If
      If Jnl.TRNTYP = 0 And Flag = False Then
         DescStr = Jnl.DESCRP
         Flag = True
      End If
      
      i = i + 1
      
      If Jnl.ACCNUM = mcolParam("ACCOUNT_NO") Then
         TempStr = Jnl.DESCRP
      End If
      
      If Jnl.TRNTYP = 0 Then
         SumDebit = SumDebit + Jnl.AMOUNT
      End If
   Next Jnl
   
   Apt.SUPNAM = TempStr
   Apt.TOTAL = SumDebit
   Apt.DOCNUM = GlJnl.VOUCHER
   Apt.DOCDAT = GlJnl.VOUDAT
   Apt.DUEDAT = GlJnl.VOUDAT
   Apt.DESCRPT = DescStr
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
Dim St As CIsTab
Dim Ap As CAPMas
Dim Width As Double
Dim Apt As CApTrn
Dim TempRs As ADODB.Recordset
Dim TempApt As CApTrn
Dim TempApRcIt As CAPRcIt
Dim DOCNUM As String
Dim ReportType As Long
Dim GlJnl As CGLJnl
Static AlreadyStart As Long
Static AlreadyNewPage As Long
Dim Jnl As CGLJnl

   RName = "genDoc"
'-----------------------------------------------------------------------------------------------------
'                                             Query Here
'-----------------------------------------------------------------------------------------------------
'   set  tempcolection =
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

   ReportType = mcolParam("REPORT_TYPE")
   If ReportType = 2 Then
      Set Apt = New CApTrn
'      Apt.SUPCOD = "" ' mcolParam("SUPPLIER_CODE")
      Apt.DOCNUM = mcolParam("DOCUMENT_NO")
     
      Apt.FROM_DOC_DATE = -1 'mcolParam("FROM_DOC_DATE")
      Apt.TO_DOC_DATE = -1 'mcolParam("TO_DOC_DATE")
      Apt.OrderBy = -1 'mcolParam("ORDER_TYPE")
      Apt.OrderType = -1 'mcolParam("ORDER_BY")
      Apt.RECTYP = 9
      Apt.QueryFlag = -1
      Call glbDaily.QueryAPTran(Apt, Rs, iCount, IsOK, glbErrorLog)
   ElseIf ReportType = 1 Then
      Set GlJnl = New CGLJnl
      GlJnl.VOUCHER = mcolParam("DOCUMENT_NO")
      GlJnl.FROM_VOUCHER_DATE = -1 ' mcolParam("FROM_DOC_DATE")
      GlJnl.TO_VOUCHER_DATE = -1 'mcolParam("TO_DOC_DATE")
'      GlJnl.JNLTYP = "01"
      GlJnl.OrderBy = -1 'mcolParam("ORDER_TYPE")
      GlJnl.OrderType = -1 'mcolParam("ORDER_BY")
      GlJnl.QueryFlag = -1
      Call glbDaily.QueryDistinctGlJnl(GlJnl, Rs, iCount, IsOK, glbErrorLog)
   End If
   
   While Not Rs.EOF
      If ReportType = 2 Then
         Call Apt.PopulateFromRS(1, Rs)
         DOCNUM = Apt.DOCNUM
      ElseIf ReportType = 1 Then
         Call GlJnl.PopulateFromRS(2, Rs)
         DOCNUM = GlJnl.VOUCHER
      End If
      
      Set m_Apt2 = New CApTrn
      m_Apt2.DOCNUM = DOCNUM
      m_Apt2.QueryFlag = 1
      Call glbDaily.QueryAPTran(m_Apt2, TempRs, iCount, IsOK, glbErrorLog)
'      If ReportType = 1 Then
      'จ่ายอื่น ๆ ให้ดึงมาจากที่ผู้ใช้ทำการกรอกเลขที่บัญชี
      
   If mcolParam("JOURNAL_TYPE") = 2 Then
      For Each Jnl In m_Apt2.JournalItems
         If Jnl.ACCNUM = mcolParam("ACCOUNT_NO") Then
            m_TempAmount = Jnl.AMOUNT
         End If
      Next Jnl
   ElseIf mcolParam("JOURNAL_TYPE") = 1 Then
      For Each Jnl In m_TempCollection
         If Jnl.ACCNUM = mcolParam("ACCOUNT_NO") Then
            m_TempAmount = Jnl.AMOUNT
         End If
      Next Jnl
   End If
   
'      End If
      
      If Not TempRs.EOF Then
         Call m_Apt2.PopulateFromRS(1, TempRs)
      Else
         Call GenerateApTranData(GlJnl, m_Apt2)
      End If
      
      If AlreadyNewPage = 0 Then
         AlreadyNewPage = 1
      Else
         Vsp.NewPage
      End If
      If AlreadyStart = 0 Then
         Vsp.StartDoc
         AlreadyStart = 1
      End If
      
      Call GenerateBody(tbNone, alngX)
      
      Vsp.FontBold = True
          
         Vsp.StartTable
         Vsp.TableCell(tcCols) = 2
         Vsp.TableCell(tcRows) = 1
          Vsp.TableCell(tcRowHeight) = 2 * 567
          Vsp.TableCell(tcColWidth, 1, 1) = "2.5cm"
          Vsp.TableCell(tcColWidth, 1, 2) = "3cm"
          Vsp.TableCell(tcAlign, 1, 2) = taLeftBottom
          If DateToStringExtEx2(mcolParam("PRINT_DATE")) <> "" Then
              Vsp.TableCell(tcText, 1, 2) = DateToStringExtEx2(mcolParam("PRINT_DATE"))
            Else
                Vsp.TableCell(tcText, 1, 2) = DateToStringExtEx2(Now)
            End If
          Vsp.EndTable
      
      Set m_Apt2 = Nothing
      Rs.MoveNext
   Wend
   Vsp.EndDoc
   Set Apt = Nothing
   Set GlJnl = Nothing

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
   
   If mcolParam("PICTURE_FLAG") = "Y" Then
      Vsp.DrawPicture mcolParam("BACK_GROUND"), Vsp.MarginLeft, Vsp.MarginTop, m_ImageWidth, m_ImageHeight, 10
   End If
   
'   Vsp.FontSize = TITLE_SIZE
'   Vsp.FontBold = True
'   Vsp.TextAlign = taCenterBottom
'   Vsp.Paragraph = MODULE_DESC
'   Vsp.Paragraph = ""
   If mblnHeader Then
      Call printHeader
   End If
'   mdY = Vsp.CurrentY
   
   
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
'Private Sub GenerateBodyAdd(Border As TableBorderSettings, alngX() As Long)
'Dim Jnl As CGLJnl
'Dim TempStr As String
'Dim IntPart As String
'Dim DecimalPart As String
'Dim PatchStr As String
'Dim Left As Long
'Dim i As Long
'Dim SumDebit As Double
'Dim SumCredit As Double
'
'   Left = 11 - m_Apt2.JournalItems.Count
'   SumDebit = 0
'   SumCredit = 0
'
'   Vsp.FontBold = True
'   Vsp.FontSize = 14
'                                                                                                               ' =============Body  2============
'   For i = 0 To 6
'      Vsp.StartTable
'      Vsp.TableCell(tcCols) = 6
'      Vsp.TableCell(tcRows) = 1
'       Vsp.TableCell(tcRowHeight, 1) = 0.75 * 567
'       Vsp.TableCell(tcColWidth, 1, 1) = "8.3cm"
'       Vsp.TableCell(tcColWidth, 1, 2) = "2.8cm"
'       Vsp.TableCell(tcColWidth, 1, 3) = "3.1cm"
'       Vsp.TableCell(tcColWidth, 1, 4) = "1cm"
'       Vsp.TableCell(tcColWidth, 1, 5) = "3.1cm"
'       Vsp.TableCell(tcColWidth, 1, 6) = "1.2cm"
'
'       TempStr = FormatNumber(frmDrCr.txtMoney(i).Text)
'       IntPart = Mid(TempStr, 1, InStr(1, TempStr, ".") - 1)
'       DecimalPart = Mid(TempStr, InStr(1, TempStr, ".") + 1)
'      If Jnl.TRNTYP = 0 Then                                              '===================Dr==========
'         SumDebit = SumDebit + Jnl.AMOUNT
'         Vsp.TableCell(tcAlign, 1, 3) = taRightMiddle
'         Vsp.TableCell(tcAlign, 1, 4) = taLeftMiddle
'         Vsp.TableCell(tcText, 1, 3) = IntPart
'         Vsp.TableCell(tcText, 1, 4) = DecimalPart
'         PatchStr = "  "
'       Else                                                                               '===================Cr===========
'         SumCredit = SumCredit + Jnl.AMOUNT
'         Vsp.TableCell(tcAlign, 1, 5) = taRightMiddle
'         Vsp.TableCell(tcAlign, 1, 6) = taLeftMiddle
'         Vsp.TableCell(tcText, 1, 5) = IntPart
'         Vsp.TableCell(tcText, 1, 6) = DecimalPart
'         PatchStr = "         "
'       End If                                                                          '==================ชื่อบัญชี=========
'       Vsp.TableCell(tcText, 1, 1) = PatchStr & Jnl.ACCNAM
'       Vsp.TableCell(tcText, 1, 2) = PatchStr & Jnl.ACCNUM
'
'      ' Vsp.TableBorder = Border
'       Vsp.TableBorder = tbAll
'
'       Vsp.EndTable
'    Next i
'
'   For i = 1 To Left
'      Vsp.StartTable
'      Vsp.TableCell(tcCols) = 1
'      Vsp.TableCell(tcRows) = 1
'       Vsp.TableCell(tcRowHeight, 1) = 0.75 * 567
'       Vsp.TableCell(tcColWidth, 1, 1) = "19.5cm"
'       Vsp.TableBorder = Border
'       Vsp.EndTable
'   Next i
'
'   Vsp.StartTable
'   Vsp.TableCell(tcCols) = 5
'   Vsp.TableCell(tcRows) = 1
'   Vsp.TableCell(tcRowHeight, 1) = 1.6 * 567
'   Vsp.TableCell(tcColWidth, 1, 1) = "11.1cm"
'   Vsp.TableCell(tcColWidth, 1, 2) = "3.1cm"
'   Vsp.TableCell(tcColWidth, 1, 3) = "1cm"
'   Vsp.TableCell(tcColWidth, 1, 4) = "3.1cm"
'   Vsp.TableCell(tcColWidth, 1, 5) = "1.2cm"
'
'   Vsp.TableCell(tcAlign, 1, 1) = taLeftMiddle
'   Vsp.TableCell(tcText, 1, 1) = "    " & ThaiBaht(SumDebit)
'
'   TempStr = FormatNumber(SumDebit)                                                     ' sum Dr
'   IntPart = Mid(TempStr, 1, InStr(1, TempStr, ".") - 1)
'   DecimalPart = Mid(TempStr, InStr(1, TempStr, ".") + 1)
'   Vsp.TableCell(tcAlign, 1, 2) = taRightMiddle
'   Vsp.TableCell(tcText, 1, 2) = IntPart
'   Vsp.TableCell(tcAlign, 1, 3) = taLeftMiddle
'   Vsp.TableCell(tcText, 1, 3) = DecimalPart
'
'   TempStr = FormatNumber(SumCredit)                                                   'Sum Cr
'   IntPart = Mid(TempStr, 1, InStr(1, TempStr, ".") - 1)
'   DecimalPart = Mid(TempStr, InStr(1, TempStr, ".") + 1)
'   Vsp.TableCell(tcAlign, 1, 4) = taRightMiddle
'   Vsp.TableCell(tcText, 1, 4) = IntPart
'   Vsp.TableCell(tcAlign, 1, 5) = taLeftMiddle
'   Vsp.TableCell(tcText, 1, 5) = DecimalPart
'   Vsp.TableBorder = Border
'   Vsp.EndTable
'
'    Vsp.FontBold = False
'End Sub




