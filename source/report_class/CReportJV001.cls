VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CReportJV001"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'User Group Report
Implements CReportInterface

Private Const MODULE_NAME = "CReportJV001"
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

Private Sub GenerateJvHeader(Border As TableBorderSettings, alngX() As Long, HeaderOffset)
Dim OldFont As Double
Dim Jnl As CGLJnl

   If m_TempCollection.Count <= 0 Then
      Set Jnl = New CGLJnl
   Else
      Set Jnl = m_TempCollection(1)
   End If
   
   Vsp.FontBold = True
   Vsp.FontSize = 14

   Vsp.StartTable
   Vsp.TableCell(tcCols) = 1
   Vsp.TableCell(tcRows) = 1
    Vsp.TableCell(tcRowHeight, 1) = 1.2 * 567 + HeaderOffset
   Vsp.TableCell(tcColWidth, 1, 1) = "18.4cm"
    Vsp.TableBorder = Border
    Vsp.EndTable

   Vsp.StartTable
   Vsp.TableCell(tcCols) = 3
   Vsp.TableCell(tcRows) = 1
    Vsp.TableCell(tcRowHeight) = 2 * 567
    Vsp.TableCell(tcColWidth, 1, 1) = "14cm"
    Vsp.TableCell(tcColWidth, 1, 2) = "1.4cm"
    Vsp.TableCell(tcColWidth, 1, 3) = "3cm"

   Vsp.TableCell(tcAlign, 1, 3) = taCenterMiddle
   Vsp.TableCell(tcText, 1, 3) = Jnl.VOUCHER
   
    Vsp.TableBorder = Border
    Vsp.EndTable

   Vsp.StartTable
   Vsp.TableCell(tcCols) = 3
   Vsp.TableCell(tcRows) = 1
    Vsp.TableCell(tcRowHeight) = 1.2 * 567
    Vsp.TableCell(tcColWidth, 1, 1) = "12cm"
    Vsp.TableCell(tcColWidth, 1, 2) = "1cm"
    Vsp.TableCell(tcColWidth, 1, 3) = "5.4cm"

   Vsp.TableCell(tcAlign, 1, 3) = taCenterTop
   Vsp.TableCell(tcText, 1, 3) = DateToStringExtEx2(Jnl.VOUDAT)

    Vsp.TableBorder = Border
    Vsp.EndTable

   Vsp.StartTable
   Vsp.TableCell(tcCols) = 1
   Vsp.TableCell(tcRows) = 1
    Vsp.TableCell(tcRowHeight) = 1.45 * 567
    Vsp.TableCell(tcColWidth, 1, 1) = "18.4cm"
    Vsp.TableBorder = Border
    Vsp.EndTable

    Vsp.FontBold = False
    Set Jnl = Nothing
End Sub
'
Private Sub GenerateJVFooter(Bd As CGLJnl, Border As TableBorderSettings, alngX() As Long)
Dim OldFont As Double

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

   Call GenerateJvHeader(tbNone, alngX, 0)
   
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

   m_ImageWidth = 18.4 * 567
   m_ImageHeight = 26.26 * 567

   mdblWid = Vsp.PageWidth - Vsp.MarginLeft - Vsp.MarginRight
   If Not glbDatabaseMngr.GetServerDateTime(strDate, glbErrorLog) Then
      mstrErrMsg = "Error GetDateTime Error."
      Exit Function
   End If
   mdteDate = InternalDateToDate(strDate)
   initDoc = True
End Function

Private Sub GenerateJvItem(Bd As CGLJnl, Border As TableBorderSettings, alngX() As Long)
Dim Jnl As CGLJnl
Dim i As Long
Dim j As Long
Dim RowHeight As Double
Dim MaxRow As Long
Dim LeftRow As Long
Dim TempStr As String
Dim IntPart As String
Dim DecimalPart As String

   MaxRow = 20
   RowHeight = 0.727 * 20 * 567 / MaxRow

   i = 0
   For Each Jnl In m_TempCollection
      i = i + 1

      '===============
      Vsp.StartTable
      Vsp.TableCell(tcCols) = 6
      Vsp.TableCell(tcRows) = 1
      Vsp.TableCell(tcFontSize) = 14

       Vsp.TableCell(tcRowHeight) = RowHeight
       Vsp.TableCell(tcColWidth, 1, 1) = "2.8cm"
       Vsp.TableCell(tcColWidth, 1, 2) = "7.2cm"
       Vsp.TableCell(tcColWidth, 1, 3) = "3.1cm"
       Vsp.TableCell(tcColWidth, 1, 4) = "1.0cm"
       Vsp.TableCell(tcColWidth, 1, 5) = "3.2cm"
       Vsp.TableCell(tcColWidth, 1, 6) = "1.1cm"
   
       TempStr = FormatNumber(Jnl.AMOUNT)
       IntPart = Mid(TempStr, 1, InStr(1, TempStr, ".") - 1)
       DecimalPart = Mid(TempStr, InStr(1, TempStr, ".") + 1)

      If Jnl.TRNTYP = 0 Then
         Vsp.TableCell(tcAlign, 1, 1) = taCenterMiddle
         Vsp.TableCell(tcText, 1, 1) = Jnl.ACCNUM
         
         Vsp.TableCell(tcAlign, 1, 2) = taLeftMiddle
         Vsp.TableCell(tcText, 1, 2) = Jnl.ACCNAM
      
         Vsp.TableCell(tcAlign, 1, 3) = taRightMiddle
         Vsp.TableCell(tcText, 1, 3) = IntPart
         Vsp.TableCell(tcAlign, 1, 4) = taLeftMiddle
         Vsp.TableCell(tcText, 1, 4) = DecimalPart
      Else
         Vsp.TableCell(tcAlign, 1, 1) = taCenterMiddle
         Vsp.TableCell(tcText, 1, 1) = "     " & Jnl.ACCNUM
         
         Vsp.TableCell(tcAlign, 1, 2) = taLeftMiddle
         Vsp.TableCell(tcText, 1, 2) = "     " & Jnl.ACCNAM
      
         Vsp.TableCell(tcAlign, 1, 5) = taRightMiddle
         Vsp.TableCell(tcText, 1, 5) = IntPart
         Vsp.TableCell(tcAlign, 1, 6) = taLeftMiddle
         Vsp.TableCell(tcText, 1, 6) = DecimalPart
      End If

       Vsp.TableBorder = Border
       Vsp.EndTable
       '===============
   
      '===============
      i = i + 1
      
      Vsp.StartTable
      Vsp.TableCell(tcCols) = 6
      Vsp.TableCell(tcRows) = 1
      Vsp.TableCell(tcFontSize) = 14

       Vsp.TableCell(tcRowHeight) = RowHeight
       Vsp.TableCell(tcColWidth, 1, 1) = "2.8cm"
       Vsp.TableCell(tcColWidth, 1, 2) = "7.2cm"
       Vsp.TableCell(tcColWidth, 1, 3) = "3.1cm"
       Vsp.TableCell(tcColWidth, 1, 4) = "1.0cm"
       Vsp.TableCell(tcColWidth, 1, 5) = "3.2cm"
       Vsp.TableCell(tcColWidth, 1, 6) = "1.1cm"
   
      If Jnl.TRNTYP = 0 Then
         Vsp.TableCell(tcAlign, 1, 2) = taLeftMiddle
         Vsp.TableCell(tcText, 1, 2) = Jnl.DESCRP
      Else
         Vsp.TableCell(tcAlign, 1, 2) = taLeftMiddle
         Vsp.TableCell(tcText, 1, 2) = "     " & Jnl.DESCRP
      End If

       Vsp.TableBorder = Border
       Vsp.EndTable
       '===============
       
       If (i Mod 20) = 0 Then
         Vsp.NewPage
       End If
   Next Jnl
   m_EndItem = True

   LeftRow = MaxRow - i

   For j = 1 To LeftRow
      Vsp.StartTable
      Vsp.TableCell(tcCols) = 6
      Vsp.TableCell(tcRows) = 1
      Vsp.TableCell(tcFontSize) = 14

       Vsp.TableCell(tcRowHeight) = RowHeight
       Vsp.TableCell(tcColWidth, 1, 1) = "2.8cm"
       Vsp.TableCell(tcColWidth, 1, 2) = "7.2cm"
       Vsp.TableCell(tcColWidth, 1, 3) = "3.1cm"
       Vsp.TableCell(tcColWidth, 1, 4) = "1.0cm"
       Vsp.TableCell(tcColWidth, 1, 5) = "3.2cm"
       Vsp.TableCell(tcColWidth, 1, 6) = "1.1cm"

       Vsp.TableBorder = Border
       Vsp.EndTable
   Next j
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
Dim DocNum As String
Dim ReportType As Long
Static AlreadyStart As Long
Static AlreadyNewPage As Long
Dim Jnl As CGLJnl

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

   Set m_TempGLJnl = New CGLJnl
   m_TempGLJnl.VOUCHER = mcolParam("DOCUMENT_NO")
   Call m_TempGLJnl.QueryData(1, Rs, iCount)
   While Not Rs.EOF
      Set Jnl = New CGLJnl
      Call Jnl.PopulateFromRS(1, Rs)
      Call m_TempCollection.Add(Jnl)
      Set Jnl = Nothing
      
      Rs.MoveNext
   Wend
   
   Vsp.StartDoc

   Call GenerateJvItem(m_TempGLJnl, tbNone, alngX)
   
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
      Call GenerateJVFooter(m_TempGLJnl, tbNone, alngX)
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
   
   If mcolParam("PICTURE_FLAG") = "Y" Then
      Vsp.DrawPicture mcolParam("BACK_GROUND"), Vsp.MarginLeft, Vsp.MarginTop, m_ImageWidth, m_ImageHeight, 10
   End If
   
   If mblnHeader Then
      Call printHeader
   End If
   
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


