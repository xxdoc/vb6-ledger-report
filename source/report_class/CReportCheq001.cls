VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CReportCheq001"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'User Group Report
Implements CReportInterface

Private Const MODULE_NAME = "CReportCheq001"
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

Private m_ImageWidth As Double
Private m_ImageHeight As Double
Private m_LedgerReport As Collection
Private m_TempCollection As Collection
Private m_TempGLJnl As CGLJnl
'Private m_Di As CDoItem
Private m_Total1(100)  As Double
Private m_Total2(100)  As Double
Private m_EndItem As Boolean

Public m_Cr As Collection
'Public OKClick As Boolean

Private Sub Class_Initialize()
   m_EndItem = False
   
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
   Set m_TempCollection = New Collection
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
   Set m_TempCollection = Nothing
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

   m_ImageWidth = 17.8 * 567
   m_ImageHeight = 8.9 * 567

   mdblWid = Vsp.PageWidth - Vsp.MarginLeft - Vsp.MarginRight
   If Not glbDatabaseMngr.GetServerDateTime(strDate, glbErrorLog) Then
      mstrErrMsg = "Error GetDateTime Error."
      Exit Function
   End If
   mdteDate = InternalDateToDate(strDate)
   initDoc = True
End Function

Private Sub GenerateChqItem(Bd As CBkTrn, Border As TableBorderSettings, alngX() As Long, CheqType As Long)
Dim CheqAmount As Double
Dim CheqDate As Date
Dim SupplierName As String
Dim Pos As Long

   If mcolParam("LINE_FLAG") = "Y" Then
      CheqDate = -1
   Else
      CheqDate = Bd.CHQDAT
   End If
   
   CheqAmount = Bd.AMOUNT
   
   If Len(mcolParam("PAYEE")) > 0 Then
      SupplierName = mcolParam("PAYEE")
   Else
      SupplierName = Bd.SUPNAM
      If Len(SupplierName) <= 0 Then
         'ให้ตัดเอา ข้อความที่อยู่ข้างหน้า * มาเท่านั้น
         Pos = InStr(1, Bd.REMARK, "*")
         If Pos > 0 Then
            SupplierName = Mid(Bd.REMARK, 1, Pos - 1)
         Else
            SupplierName = Bd.REMARK
         End If
      End If
   End If
         
   Vsp.FontSize = 16
   If CheqType = 4 Then
      Vsp.StartTable
      Vsp.TableCell(tcCols) = 3
      Vsp.TableCell(tcRows) = 1
       Vsp.TableCell(tcRowHeight, 1) = 1 * 567
      Vsp.TableCell(tcColWidth, 1, 1) = "12.1cm"
      Vsp.TableCell(tcColWidth, 1, 2) = "1cm"
      Vsp.TableCell(tcColWidth, 1, 3) = "4.7cm"
       Vsp.TableBorder = Border
       
       Vsp.TableCell(tcAlign, 1, 3) = taLeftMiddle
       Vsp.TableCell(tcText, 1, 3) = DateToStringExtEx2(CheqDate)
       Vsp.TableCell(tcFontBold) = True
       Vsp.EndTable
   
      Vsp.StartTable
      Vsp.TableCell(tcCols) = 1
      Vsp.TableCell(tcRows) = 1
       Vsp.TableCell(tcRowHeight, 1) = 0.5 * 567
      Vsp.TableCell(tcColWidth, 1, 1) = "17.8cm"
       Vsp.TableBorder = Border
       Vsp.TableCell(tcFontBold) = True
       Vsp.EndTable
   
      Vsp.StartTable
      Vsp.TableCell(tcCols) = 3
      Vsp.TableCell(tcRows) = 1
       Vsp.TableCell(tcRowHeight, 1) = 1 * 567
      Vsp.TableCell(tcColWidth, 1, 1) = "2.3cm"
      Vsp.TableCell(tcColWidth, 1, 2) = "13.5cm"
      Vsp.TableCell(tcColWidth, 1, 3) = "2cm"
      
       Vsp.TableCell(tcAlign, 1, 2) = taLeftMiddle
       Vsp.TableCell(tcText, 1, 2) = SupplierName

       Vsp.TableBorder = Border
       Vsp.TableCell(tcFontBold) = True
       Vsp.EndTable

      Vsp.StartTable
      Vsp.TableCell(tcCols) = 2
      Vsp.TableCell(tcRows) = 1
       Vsp.TableCell(tcRowHeight, 1) = 0.9 * 567
      Vsp.TableCell(tcColWidth, 1, 1) = "4cm"
      Vsp.TableCell(tcColWidth, 1, 2) = "13.8cm"
      
       Vsp.TableCell(tcAlign, 1, 2) = taLeftMiddle
       Vsp.TableCell(tcText, 1, 2) = "-" & ThaiBaht(CheqAmount) & "-"
       
       Vsp.TableBorder = Border
       Vsp.TableCell(tcFontBold) = True
       Vsp.EndTable
   
      Vsp.StartTable
      Vsp.TableCell(tcCols) = 2
      Vsp.TableCell(tcRows) = 1
       Vsp.TableCell(tcRowHeight, 1) = 1.1 * 567
      Vsp.TableCell(tcColWidth, 1, 1) = "11.2cm"
      Vsp.TableCell(tcColWidth, 1, 2) = "6.6cm"
      
       Vsp.TableCell(tcAlign, 1, 2) = taLeftMiddle
       Vsp.TableCell(tcText, 1, 2) = "-" & FormatNumber(CheqAmount) & "-"
       
       Vsp.TableBorder = Border
       Vsp.TableCell(tcFontBold) = True
       Vsp.EndTable
   ElseIf CheqType = 1 Then
      Vsp.StartTable
      Vsp.TableCell(tcCols) = 2
      Vsp.TableCell(tcRows) = 1
       Vsp.TableCell(tcRowHeight, 1) = 1 * 567
      Vsp.TableCell(tcColWidth, 1, 1) = "13cm"
      Vsp.TableCell(tcColWidth, 1, 2) = "4.8cm"
       Vsp.TableBorder = Border
       
       Vsp.TableCell(tcAlign, 1, 2) = taLeftMiddle
       Vsp.TableCell(tcText, 1, 2) = DateToStringExtEx2(CheqDate)
       Vsp.TableCell(tcFontBold) = True
       Vsp.EndTable
   
      Vsp.StartTable
      Vsp.TableCell(tcCols) = 1
      Vsp.TableCell(tcRows) = 1
       Vsp.TableCell(tcRowHeight, 1) = 0.6 * 567
      Vsp.TableCell(tcColWidth, 1, 1) = "17.8cm"
       Vsp.TableBorder = Border
       Vsp.TableCell(tcFontBold) = True
       Vsp.EndTable
   
      Vsp.StartTable
      Vsp.TableCell(tcCols) = 3
      Vsp.TableCell(tcRows) = 1
       Vsp.TableCell(tcRowHeight, 1) = 0.6 * 567
      Vsp.TableCell(tcColWidth, 1, 1) = "2.5cm"
      Vsp.TableCell(tcColWidth, 1, 2) = "11.5cm"
      Vsp.TableCell(tcColWidth, 1, 3) = "3.8cm"
      
       Vsp.TableCell(tcAlign, 1, 2) = taLeftMiddle
       Vsp.TableCell(tcText, 1, 2) = SupplierName

       Vsp.TableBorder = Border
       Vsp.TableCell(tcFontBold) = True
       Vsp.EndTable

      Vsp.StartTable
      Vsp.TableCell(tcCols) = 2
      Vsp.TableCell(tcRows) = 1
       Vsp.TableCell(tcRowHeight, 1) = 0.8 * 567
      Vsp.TableCell(tcColWidth, 1, 1) = "5cm"
      Vsp.TableCell(tcColWidth, 1, 2) = "12.8cm"
      
       Vsp.TableCell(tcAlign, 1, 2) = taLeftMiddle
       Vsp.TableCell(tcText, 1, 2) = "-" & ThaiBaht(CheqAmount) & "-"
       
       Vsp.TableBorder = Border
       Vsp.TableCell(tcFontBold) = True
       Vsp.EndTable
   
      Vsp.StartTable
      Vsp.TableCell(tcCols) = 1
      Vsp.TableCell(tcRows) = 1
       Vsp.TableCell(tcRowHeight, 1) = 0.4 * 567
      Vsp.TableCell(tcColWidth, 1, 1) = "17.8cm"
       Vsp.TableBorder = Border
       Vsp.TableCell(tcFontBold) = True
       Vsp.EndTable
      Vsp.StartTable
      
      Vsp.TableCell(tcCols) = 2
      Vsp.TableCell(tcRows) = 1
       Vsp.TableCell(tcRowHeight, 1) = 0.8 * 567
      Vsp.TableCell(tcColWidth, 1, 1) = "11.5cm"
      Vsp.TableCell(tcColWidth, 1, 2) = "6.3cm"
      
       Vsp.TableCell(tcAlign, 1, 2) = taLeftMiddle
       Vsp.TableCell(tcText, 1, 2) = "-" & FormatNumber(CheqAmount) & "-"
       
       Vsp.TableBorder = Border
       Vsp.TableCell(tcFontBold) = True
       Vsp.EndTable
       
   ElseIf CheqType = 2 Then
      
      Vsp.StartTable
      Vsp.TableCell(tcCols) = 1
      Vsp.TableCell(tcRows) = 1
       Vsp.TableCell(tcRowHeight, 1) = 0.4 * 567
      Vsp.TableCell(tcColWidth, 1, 1) = "17.8cm"
       Vsp.TableBorder = Border
       Vsp.TableCell(tcFontBold) = True
       Vsp.EndTable
       
      Vsp.StartTable
      Vsp.TableCell(tcCols) = 2
      Vsp.TableCell(tcRows) = 1
       Vsp.TableCell(tcRowHeight, 1) = 1 * 567
      Vsp.TableCell(tcColWidth, 1, 1) = "13cm"
      Vsp.TableCell(tcColWidth, 1, 2) = "4.8cm"
       Vsp.TableBorder = Border
       
       Vsp.TableCell(tcAlign, 1, 2) = taLeftMiddle
       Vsp.TableCell(tcText, 1, 2) = DateToStringExtEx2(CheqDate)
       Vsp.TableCell(tcFontBold) = True
       Vsp.EndTable
   
      Vsp.StartTable
      Vsp.TableCell(tcCols) = 1
      Vsp.TableCell(tcRows) = 1
       Vsp.TableCell(tcRowHeight, 1) = 0.4 * 567
      Vsp.TableCell(tcColWidth, 1, 1) = "17.8cm"
       Vsp.TableBorder = Border
       Vsp.TableCell(tcFontBold) = True
       Vsp.EndTable
   
      Vsp.StartTable
      Vsp.TableCell(tcCols) = 3
      Vsp.TableCell(tcRows) = 1
       Vsp.TableCell(tcRowHeight, 1) = 1 * 567
      Vsp.TableCell(tcColWidth, 1, 1) = "2cm"
      Vsp.TableCell(tcColWidth, 1, 2) = "13cm"
      Vsp.TableCell(tcColWidth, 1, 3) = "2.8cm"
      
       Vsp.TableCell(tcAlign, 1, 2) = taLeftMiddle
       Vsp.TableCell(tcText, 1, 2) = SupplierName

       Vsp.TableBorder = Border
       Vsp.TableCell(tcFontBold) = True
       Vsp.EndTable

      Vsp.StartTable
      Vsp.TableCell(tcCols) = 2
      Vsp.TableCell(tcRows) = 1
       Vsp.TableCell(tcRowHeight, 1) = 1 * 567
      Vsp.TableCell(tcColWidth, 1, 1) = "3.5cm"
      Vsp.TableCell(tcColWidth, 1, 2) = "14.3cm"
      
       Vsp.TableCell(tcAlign, 1, 2) = taLeftMiddle
       Vsp.TableCell(tcText, 1, 2) = "-" & ThaiBaht(CheqAmount) & "-"
       
       Vsp.TableBorder = Border
       Vsp.TableCell(tcFontBold) = True
       Vsp.EndTable
   
      
      Vsp.TableCell(tcCols) = 2
      Vsp.TableCell(tcRows) = 1
       Vsp.TableCell(tcRowHeight, 1) = 0.8 * 567
      Vsp.TableCell(tcColWidth, 1, 1) = "11.5cm"
      Vsp.TableCell(tcColWidth, 1, 2) = "6.3cm"
      
       Vsp.TableCell(tcAlign, 1, 2) = taLeftMiddle
       Vsp.TableCell(tcText, 1, 2) = "-" & FormatNumber(CheqAmount) & "-"
       
       Vsp.TableBorder = Border
       Vsp.TableCell(tcFontBold) = True
       Vsp.EndTable
   ElseIf CheqType = 6 Then
       
      Vsp.StartTable
      Vsp.TableCell(tcCols) = 2
      Vsp.TableCell(tcRows) = 1
       Vsp.TableCell(tcRowHeight, 1) = 1 * 567
      Vsp.TableCell(tcColWidth, 1, 1) = "13cm"
      Vsp.TableCell(tcColWidth, 1, 2) = "4.8cm"
       Vsp.TableBorder = Border
       
       Vsp.TableCell(tcAlign, 1, 2) = taLeftMiddle
       Vsp.TableCell(tcText, 1, 2) = DateToStringExtEx2(CheqDate)
       Vsp.TableCell(tcFontBold) = True
       Vsp.EndTable
   
      Vsp.StartTable
      Vsp.TableCell(tcCols) = 1
      Vsp.TableCell(tcRows) = 1
       Vsp.TableCell(tcRowHeight, 1) = 0.7 * 567
      Vsp.TableCell(tcColWidth, 1, 1) = "17.8cm"
       Vsp.TableBorder = Border
       Vsp.TableCell(tcFontBold) = True
       Vsp.EndTable
   
      Vsp.StartTable
      Vsp.TableCell(tcCols) = 3
      Vsp.TableCell(tcRows) = 1
       Vsp.TableCell(tcRowHeight, 1) = 0.75 * 567
      Vsp.TableCell(tcColWidth, 1, 1) = "3cm"
      Vsp.TableCell(tcColWidth, 1, 2) = "11.5cm"
      Vsp.TableCell(tcColWidth, 1, 3) = "3.3cm"
      
       Vsp.TableCell(tcAlign, 1, 2) = taLeftMiddle
       Vsp.TableCell(tcText, 1, 2) = SupplierName

       Vsp.TableBorder = Border
       Vsp.TableCell(tcFontBold) = True
       Vsp.EndTable

      Vsp.StartTable
      Vsp.TableCell(tcCols) = 2
      Vsp.TableCell(tcRows) = 1
       Vsp.TableCell(tcRowHeight, 1) = 0.75 * 567
      Vsp.TableCell(tcColWidth, 1, 1) = "3cm"
      Vsp.TableCell(tcColWidth, 1, 2) = "14.8cm"
      
       Vsp.TableCell(tcAlign, 1, 2) = taLeftMiddle
       Vsp.TableCell(tcText, 1, 2) = "-" & ThaiBaht(CheqAmount) & "-"
       
       Vsp.TableBorder = Border
       Vsp.TableCell(tcFontBold) = True
       Vsp.EndTable
   
      
      Vsp.TableCell(tcCols) = 2
      Vsp.TableCell(tcRows) = 1
       Vsp.TableCell(tcRowHeight, 1) = 1.1 * 567
      Vsp.TableCell(tcColWidth, 1, 1) = "11.5cm"
      Vsp.TableCell(tcColWidth, 1, 2) = "6.3cm"
      
       Vsp.TableCell(tcAlign, 1, 2) = taLeftMiddle
       Vsp.TableCell(tcText, 1, 2) = "-" & FormatNumber(CheqAmount) & "-"
       
       Vsp.TableBorder = Border
       Vsp.TableCell(tcFontBold) = True
       Vsp.EndTable
   ElseIf CheqType = 5 Then
      Vsp.StartTable
      Vsp.TableCell(tcCols) = 2
      Vsp.TableCell(tcRows) = 1
       Vsp.TableCell(tcRowHeight, 1) = 1.1 * 567
      Vsp.TableCell(tcColWidth, 1, 1) = "13cm"
      Vsp.TableCell(tcColWidth, 1, 2) = "4.8cm"
       Vsp.TableBorder = Border
       
       Vsp.TableCell(tcAlign, 1, 2) = taLeftMiddle
       Vsp.TableCell(tcText, 1, 2) = DateToStringExtEx2(CheqDate)
       Vsp.TableCell(tcFontBold) = True
       Vsp.EndTable
   
      Vsp.StartTable
      Vsp.TableCell(tcCols) = 1
      Vsp.TableCell(tcRows) = 1
       Vsp.TableCell(tcRowHeight, 1) = 0.6 * 567
      Vsp.TableCell(tcColWidth, 1, 1) = "17.8cm"
       Vsp.TableBorder = Border
       Vsp.TableCell(tcFontBold) = True
       Vsp.EndTable
   
      Vsp.StartTable
      Vsp.TableCell(tcCols) = 3
      Vsp.TableCell(tcRows) = 1
       Vsp.TableCell(tcRowHeight, 1) = 0.8 * 567
      Vsp.TableCell(tcColWidth, 1, 1) = "3.5cm"
      Vsp.TableCell(tcColWidth, 1, 2) = "11cm"
      Vsp.TableCell(tcColWidth, 1, 3) = "3.3cm"
      
       Vsp.TableCell(tcAlign, 1, 2) = taLeftMiddle
       Vsp.TableCell(tcText, 1, 2) = SupplierName

       Vsp.TableBorder = Border
       Vsp.TableCell(tcFontBold) = True
       Vsp.EndTable

      Vsp.StartTable
      Vsp.TableCell(tcCols) = 2
      Vsp.TableCell(tcRows) = 1
       Vsp.TableCell(tcRowHeight, 1) = 0.8 * 567
      Vsp.TableCell(tcColWidth, 1, 1) = "3.5cm"
      Vsp.TableCell(tcColWidth, 1, 2) = "14.3cm"
      
       Vsp.TableCell(tcAlign, 1, 2) = taLeftMiddle
       Vsp.TableCell(tcText, 1, 2) = "-" & ThaiBaht(CheqAmount) & "-"
       
       Vsp.TableBorder = Border
       Vsp.TableCell(tcFontBold) = True
       Vsp.EndTable
   
      
      Vsp.TableCell(tcCols) = 2
      Vsp.TableCell(tcRows) = 1
       Vsp.TableCell(tcRowHeight, 1) = 0.9 * 567
      Vsp.TableCell(tcColWidth, 1, 1) = "11.5cm"
      Vsp.TableCell(tcColWidth, 1, 2) = "6.3cm"
      
       Vsp.TableCell(tcAlign, 1, 2) = taLeftMiddle
       Vsp.TableCell(tcText, 1, 2) = "-" & FormatNumber(CheqAmount) & "-"
       
       Vsp.TableBorder = Border
       Vsp.TableCell(tcFontBold) = True
       Vsp.EndTable
          
   ElseIf CheqType = 3 Then
      Vsp.StartTable
      Vsp.TableCell(tcCols) = 1
      Vsp.TableCell(tcRows) = 1
       Vsp.TableCell(tcRowHeight, 1) = 0.9 * 567
      Vsp.TableCell(tcColWidth, 1, 1) = "17.8cm"
       Vsp.TableBorder = Border
       Vsp.TableCell(tcFontBold) = True
       Vsp.EndTable
       
      Vsp.StartTable
      Vsp.TableCell(tcCols) = 2
      Vsp.TableCell(tcRows) = 1
       Vsp.TableCell(tcRowHeight, 1) = 0.9 * 567
      Vsp.TableCell(tcColWidth, 1, 1) = "13cm"
      Vsp.TableCell(tcColWidth, 1, 2) = "4.8cm"
       Vsp.TableBorder = Border
       
       Vsp.TableCell(tcAlign, 1, 2) = taLeftMiddle
       Vsp.TableCell(tcText, 1, 2) = DateToStringExtEx2(CheqDate)
       Vsp.TableCell(tcFontBold) = True
       Vsp.EndTable
   
      Vsp.StartTable
      Vsp.TableCell(tcCols) = 3
      Vsp.TableCell(tcRows) = 1
       Vsp.TableCell(tcRowHeight, 1) = 0.9 * 567
      Vsp.TableCell(tcColWidth, 1, 1) = "2.5cm"
      Vsp.TableCell(tcColWidth, 1, 2) = "13cm"
      Vsp.TableCell(tcColWidth, 1, 3) = "2.3cm"
      
       Vsp.TableCell(tcAlign, 1, 2) = taLeftMiddle
       Vsp.TableCell(tcText, 1, 2) = SupplierName

       Vsp.TableBorder = Border
       Vsp.TableCell(tcFontBold) = True
       Vsp.EndTable

      Vsp.StartTable
      Vsp.TableCell(tcCols) = 2
      Vsp.TableCell(tcRows) = 1
       Vsp.TableCell(tcRowHeight, 1) = 0.75 * 567
      Vsp.TableCell(tcColWidth, 1, 1) = "2.5cm"
      Vsp.TableCell(tcColWidth, 1, 2) = "15.3cm"
      
       Vsp.TableCell(tcAlign, 1, 2) = taLeftMiddle
       Vsp.TableCell(tcText, 1, 2) = "-" & ThaiBaht(CheqAmount) & "-"
       
       Vsp.TableBorder = Border
       Vsp.TableCell(tcFontBold) = True
       Vsp.EndTable
   
      
      Vsp.TableCell(tcCols) = 2
      Vsp.TableCell(tcRows) = 1
       Vsp.TableCell(tcRowHeight, 1) = 0.8 * 567
      Vsp.TableCell(tcColWidth, 1, 1) = "11.5cm"
      Vsp.TableCell(tcColWidth, 1, 2) = "6.3cm"
      
       Vsp.TableCell(tcAlign, 1, 2) = taLeftMiddle
       Vsp.TableCell(tcText, 1, 2) = "-" & FormatNumber(CheqAmount) & "-"
       
       Vsp.TableBorder = Border
       Vsp.TableCell(tcFontBold) = True
       Vsp.EndTable
   End If
   
   Vsp.TableBorder = Border
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
Dim iCount As Long
Dim TempStr As String
Dim PrevKey1 As String
Dim PrevKey2 As String
Dim Width As Double
Dim TempRs As ADODB.Recordset
Dim Bkt As CBkTrn

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
   
   For j = 1 To UBound(m_Total1)
      m_Total1(j) = 0
      m_Total2(j) = 0
   Next j

   Set Bkt = New CBkTrn
   Bkt.CHQNUM = mcolParam("CHECK_NO")
   Call Bkt.QueryData(1, Rs, iCount)
   If Rs.EOF Then
      mstrErrMsg = "ไม่พบเลขที่เช็คที่ระบุ"
      genDoc = False
      Exit Function
   End If
   Call Bkt.PopulateFromRS(1, Rs)
   
   Vsp.StartDoc

   Call GenerateChqItem(Bkt, tbNone, alngX, mcolParam("CHECK_TYPE"))
   
   Vsp.EndDoc
   
   Set m_TempGLJnl = Nothing
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

Private Sub VSP_EndDoc()
'This event occur when VSPrinter.EndDoc is used
End Sub

Private Sub VSP_EndPage()
Dim alngX() As Long

   If Not mblnEndPage Then Exit Sub
   
   If m_EndItem Then
'      Call GenerateChqFooter(m_TempGLJnl, tbNone, alngX)
   End If
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
   
'   If mcolParam("PICTURE_FLAG") = "Y" Then
'      Vsp.DrawPicture mcolParam("BACK_GROUND"), Vsp.MarginLeft, Vsp.MarginTop, m_ImageWidth, m_ImageHeight, 10
'   End If
   
'   If mblnHeader Then
'      Call printHeader
'   End If
   
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

