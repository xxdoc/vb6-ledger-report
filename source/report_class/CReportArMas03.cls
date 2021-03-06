VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CReportArMas03"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'User Group Report
Implements CReportInterface

Private Const MODULE_NAME = "CReportArMas03"
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
Private m_CustomerTypes As Collection

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
   Set m_CustomerTypes = New Collection
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
   Set m_CustomerTypes = Nothing
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
   
   Call m_Headers1.ClearField
   
   Vsp.FontSize = HEADER_SIZE
   Vsp.FontBold = True
   Vsp.TableBorder = tbAll
   
   Set Cf = New CReportField
      
    Set Cf = Nothing
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
   Vsp.MarginBottom = 300
   Vsp.MarginFooter = 300
   Vsp.MarginHeader = 300
   Vsp.MarginLeft = 0
   Vsp.MarginRight = 0
   Vsp.MarginTop = 300
   Vsp.FontName = "AngsanaUPC"
   Vsp.FontSize = DETAIL_SIZE
      
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
Dim PrevKey1 As String
Dim PrevKey2 As String
Dim St As CIsTab
Dim Ar As CARMas
Dim tempY As Double
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
      

    Set Ar = New CARMas
    Ar.FROM_CUSTOMER_CODE = mcolParam("FROM_CUSTOMER_CODE")
   Ar.TO_CUSTOMER_CODE = mcolParam("TO_CUSTOMER_CODE")
   
   Ar.FROM_SALE_CODE = mcolParam("FROM_SALE_CODE")
   Ar.TO_SALE_CODE = mcolParam("TO_SALE_CODE")
   
   Call Ar.QueryData(1, Rs, iCount)
      
   Vsp.FontBold = False
   Vsp.FontSize = 14
   
   i = 0
   j = 0
   
   Vsp.TextAlign = taLeftBottom
   
   While Not Rs.EOF
      If i > 0 Then
         Vsp.NewPage
      End If
      i = i + 1
      
      Call Ar.PopulateFromRS(1, Rs)
      
      Vsp.DrawPicture mcolParam("BACK_GROUND"), "0.0cm", "0.4cm", "20.5cm", "28cm", 10
   
      Vsp.StartTable
      Vsp.TableCell(tcCols) = 2
      Vsp.TableCell(tcRows) = 1
      Vsp.TableCell(tcRowHeight, 1) = 1.1 * 567
      Vsp.TableCell(tcColWidth, 1, 1) = "16.1cm"
      Vsp.TableCell(tcColWidth, 1, 2) = "2.8cm"
      Vsp.TableCell(tcAlign, 1, 2) = taCenterBottom
      'Vsp.TableCell(tcFontSize, 1, 2) = 16
      Vsp.TableCell(tcText, 1, 2) = " " & Ar.CUSCOD
      Vsp.TableBorder = tbNone
      Vsp.EndTable
      
      Vsp.StartTable
      Vsp.TableCell(tcCols) = 1
      Vsp.TableCell(tcRows) = 1
      Vsp.TableCell(tcRowHeight, 1) = 1.7 * 567
      Vsp.TableCell(tcColWidth, 1, 1) = "18.9cm"
      Vsp.TableCell(tcAlign, 1, 1) = taCenterBottom
      Vsp.TableCell(tcText, 1, 1) = ""
      Vsp.TableBorder = tbNone
      Vsp.EndTable
      
      tempY = Vsp.CurrentY
      
      Vsp.StartTable
      Vsp.TableCell(tcCols) = 2
      Vsp.TableCell(tcRows) = 1
      Vsp.TableCell(tcRowHeight, 1) = 0.6 * 567
      Vsp.TableCell(tcColWidth, 1, 1) = "2.5cm"
      Vsp.TableCell(tcColWidth, 1, 2) = "16.4cm"
      
      Vsp.TableCell(tcAlign, 1, 1) = taCenterBottom
      Vsp.TableCell(tcAlign, 1, 2) = taLeftBottom
      Vsp.TableCell(tcText, 1, 1) = ""
      Vsp.TableCell(tcText, 1, 2) = Ar.PRENAM & " " & Ar.CUSNAM
      Vsp.TableBorder = tbNone
      Vsp.EndTable
      
      Vsp.StartTable
      Vsp.TableCell(tcCols) = 2
      Vsp.TableCell(tcRows) = 1
      Vsp.TableCell(tcRowHeight, 1) = 0.8 * 567
      Vsp.TableCell(tcColWidth, 1, 1) = "2cm"
      Vsp.TableCell(tcColWidth, 1, 2) = "16.9cm"
      
      Vsp.TableCell(tcAlign, 1, 1) = taCenterBottom
      Vsp.TableCell(tcAlign, 1, 2) = taLeftBottom
      Vsp.TableCell(tcText, 1, 1) = ""
      Vsp.TableCell(tcText, 1, 2) = Ar.ADDR01 & " " & Ar.ADDR02
      Vsp.TableBorder = tbNone
      Vsp.EndTable
      
      Vsp.CurrentY = tempY
      
      Vsp.StartTable
      Vsp.TableCell(tcCols) = 2
      Vsp.TableCell(tcRows) = 1
      Vsp.TableCell(tcRowHeight, 1) = 1.4 * 567
      Vsp.TableCell(tcColWidth, 1, 1) = "15.2cm"
      Vsp.TableCell(tcColWidth, 1, 2) = "3.7cm"
      Vsp.TableCell(tcAlign, 1, 1) = taCenterTop
      Vsp.TableCell(tcAlign, 1, 2) = taCenterTop
      Vsp.TableCell(tcText, 1, 1) = ""
      Vsp.TableCell(tcText, 1, 2) = Ar.TELNUM
      Vsp.TableBorder = tbNone
      Vsp.EndTable

      
      Vsp.StartTable
      Vsp.TableCell(tcCols) = 1
      Vsp.TableCell(tcRows) = 1
      Vsp.TableCell(tcRowHeight, 1) = 13.2 * 567
      Vsp.TableCell(tcColWidth, 1, 1) = "18.9cm"
      Vsp.TableCell(tcAlign, 1, 1) = taCenterBottom
      Vsp.TableCell(tcText, 1, 1) = ""
      Vsp.TableBorder = tbNone
      Vsp.EndTable
      
      Vsp.StartTable
      Vsp.TableCell(tcCols) = 4
      Vsp.TableCell(tcRows) = 1
      Vsp.TableCell(tcRowHeight, 1) = 0.5 * 567
      Vsp.TableCell(tcColWidth, 1, 1) = "4.4cm"
      Vsp.TableCell(tcColWidth, 1, 2) = "2.2cm"
      Vsp.TableCell(tcColWidth, 1, 3) = "1.3cm"
      Vsp.TableCell(tcColWidth, 1, 4) = "11.0cm"
      Vsp.TableCell(tcAlign, 1, 1) = taCenterMiddle
      Vsp.TableCell(tcAlign, 1, 2) = taCenterMiddle
      Vsp.TableCell(tcAlign, 1, 3) = taCenterMiddle
      Vsp.TableCell(tcAlign, 1, 4) = taCenterMiddle
      
      Vsp.TableCell(tcText, 1, 1) = ""
      Vsp.TableCell(tcText, 1, 2) = FormatNumber(Ar.CRLINE)
      Vsp.TableCell(tcText, 1, 3) = Ar.PAYTRM
      
      Vsp.TableBorder = tbNone
      Vsp.EndTable
      
      
      Rs.MoveNext
      
   Wend
   
   Vsp.EndDoc
   
   'Call CloseExportFile(Vsp)
   
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
   
   
'   Vsp.FontSize = TITLE_SIZE
'   Vsp.FontBold = True
'   Vsp.TextAlign = taCenterBottom
'   Vsp.Paragraph = MODULE_DESC
'   Vsp.Paragraph = ""
'   If mblnHeader Then
'      Call printHeader
'   End If
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
