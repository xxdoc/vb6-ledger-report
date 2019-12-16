VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CReportCostProducts"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'User Group Report
Implements CReportInterface

Private Const MODULE_NAME = "CReportCostProducts"
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

Private ColCustomerStock As Collection
Private ColCustomerStockSum As Collection
Private ColCostProducts As Collection
Private ColProvinceMap As Collection

Private FreeFlag As Boolean
Private SaleFlag As Boolean
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
   
   Set ColCustomerStock = New Collection
   Set ColCustomerStockSum = New Collection
   Set ColCostProducts = New Collection
   Set ColProvinceMap = New Collection
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
   
   Set ColCustomerStock = Nothing
   Set ColCustomerStockSum = Nothing
   Set ColCostProducts = Nothing
   Set ColProvinceMap = Nothing
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
Dim TempStr1 As String
Dim TempStr1_1  As String
Dim Cf As CReportField
Dim HeadCf As CReportField
Dim j As Long
Dim TempStr As String

Dim FromDate As Date
Dim ToDate As Date
   
   Call GetFirstLastDate(DateSerial(mcolParam("TO_YEAR_NO") - 543, mcolParam("TO_MONTH_ID"), 1), FromDate, ToDate)
   FromDate = DateSerial(mcolParam("FROM_YEAR_NO") - 543, mcolParam("FROM_MONTH_ID"), 1)
   
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
   
      Call Cf.SetFieldValue(10, "^", "No", "<")
      Call m_Headers1.AddField(Cf)
   
      Call Cf.SetFieldValue(15, "^", "Year", "<")
      Call m_Headers1.AddField(Cf)
   
      Call Cf.SetFieldValue(13, "^", "Month", "<")
      Call m_Headers1.AddField(Cf)
   
      Call Cf.SetFieldValue(25, "^", "CUST CODE", "<")
      Call m_Headers1.AddField(Cf)
   
      Call Cf.SetFieldValue(50, "^", "CUST NAME", "<")
      Call m_Headers1.AddField(Cf)
   
      Call Cf.SetFieldValue(25, "^", "PROVINCE", "<")
      Call m_Headers1.AddField(Cf)
   
      Call Cf.SetFieldValue(20, "^", "MGP CODE", "<")
      Call m_Headers1.AddField(Cf)
   
      Call Cf.SetFieldValue(15, "^", "BI CODE", ">")
      Call m_Headers1.AddField(Cf)
   
      Call Cf.SetFieldValue(40, "^", "ITEM NAME", "<")
      Call m_Headers1.AddField(Cf)
   
      Call Cf.SetFieldValue(20, "^", "QTY", ">")
      Call m_Headers1.AddField(Cf)
   
      Call Cf.SetFieldValue(20, "^", "SALEPRICE", ">")
      Call m_Headers1.AddField(Cf)
   
      Call Cf.SetFieldValue(20, "^", "NET SALES", ">")
      Call m_Headers1.AddField(Cf)
   
      Call Cf.SetFieldValue(15, "^", "FreeGoods", ">")
      Call m_Headers1.AddField(Cf)
   
      Call Cf.SetFieldValue(15, "^", "TypeCus", ">")
      Call m_Headers1.AddField(Cf)
      
   Set Cf = Nothing
   
  Call m_Headers1.GetString(1, TempStr1, TempStr1_1)
   strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
   strPrint = TempStr1_1
   Call Vsp.AddTable(strFormat, "", strPrint)
            
   Vsp.TableBorder = tbdT
   Vsp.FontBold = blnBold
   Vsp.FontSize = iSize
End Sub

Private Function initDoc() As Boolean
Dim strDate As String
Dim TempStr1 As String
Dim FromDate As Date
Dim ToDate As Date
      
   Call GetFirstLastDate(DateSerial(mcolParam("TO_YEAR_NO") - 543, mcolParam("TO_MONTH_ID"), 1), FromDate, ToDate)
   FromDate = DateSerial(mcolParam("FROM_YEAR_NO") - 543, mcolParam("FROM_MONTH_ID"), 1)
   
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
   Vsp.MarginRight = 500
   Vsp.MarginTop = 300
   Vsp.FontName = "AngsanaUPC"
   Vsp.FontSize = DETAIL_SIZE
   
   If mcolParam("SPLIT_SALE") = "Y" Then
      Call LoadCostProductStcrd(ColCustomerStock, FromDate, ToDate, mcolParam("FROM_CUSTOMER_CODE"), mcolParam("TO_CUSTOMER_CODE"), mcolParam("FROM_STOCK_CODE"), mcolParam("TO_STOCK_CODE"), mcolParam("FROM_SALE_CODE"), mcolParam("TO_SALE_CODE"), 2)
   End If
   
   Call LoadCostProductStcrd(ColCustomerStock, FromDate, ToDate, mcolParam("FROM_CUSTOMER_CODE"), mcolParam("TO_CUSTOMER_CODE"), mcolParam("FROM_STOCK_CODE"), mcolParam("TO_STOCK_CODE"), mcolParam("FROM_SALE_CODE"), mcolParam("TO_SALE_CODE"))
    
   Call LoadCostProducts(ColCostProducts)
   
   Call LoadProvinceMap(Nothing, ColProvinceMap)
   
   Call GenerateListCollection(ColCustomerStock, ColCustomerStockSum)
      
   MODULE_DESC = glbCompanyName & vbCrLf & mcolParam("REPORT_NAME") & _
                                       vbCrLf & "จากงวด " & mcolParam("FROM_MONTH_ID") & "/" & mcolParam("FROM_YEAR_NO") & "   " & " ถึงงวด " & mcolParam("TO_MONTH_ID") & "/" & mcolParam("TO_YEAR_NO") & " " & " จากลูกค้า " & EmptyToString(mcolParam("FROM_CUSTOMER_CODE"), "N/A") & " ถึงลูกค้า " & EmptyToString(mcolParam("TO_CUSTOMER_CODE"), "N/A") & " " & " จากสินค้า " & EmptyToString(mcolParam("FROM_STOCK_CODE"), "N/A") & " ถึงสินค้า " & EmptyToString(mcolParam("TO_STOCK_CODE"), "N/A")
   
   Call SetReportConfig(Vsp, mcolParam("REPORT_KEY"))
   
   mdblWid = Vsp.PageWidth - Vsp.MarginLeft - Vsp.MarginRight
   If Not glbDatabaseMngr.GetServerDateTime(strDate, glbErrorLog) Then
      mstrErrMsg = "Error GetDateTime Error."
      Exit Function
   End If
   mdteDate = InternalDateToDate(strDate)
   initDoc = True
End Function
Private Function genDoc(isPreview As Boolean) As Boolean
On Error GoTo ErrHandler
Dim RName As String
Dim i As Long
Dim j As Long
Dim K As Long
Dim strFormat As String
Dim alngX() As Long
Dim IsOK As Boolean
Dim Amt As Double
Dim HeadCf As CReportField
Dim BodyCf As CReportField
Dim TempStr1 As String
Dim TempStr2 As String
Dim TempStr3 As String
Dim Total1(100) As Double
Dim Total2(100) As Double
Dim Total3(100) As Double
Dim Total4(100) As Double

Dim Total5(100) As Double
Dim Total6(100) As Double
Dim Total7(100) As Double
Dim Total8(100) As Double
Dim Total9(100) As Double
Dim Total10(100) As Double

Dim amtSystmp As Double
Dim Amt2 As Double
Dim sumFreeByUser As Double

Dim TempDataStr2 As CCostProducts
Dim Str2 As CPromotionConfig

Dim iCount As Long
Dim TempStr As String
Dim PrevKey1 As String
Dim PrevKey2 As String

Dim PrevKey3 As String
Dim PrevKey4 As String

Dim FromDate  As Date
Dim ToDate  As Date
Dim SumAmount As Double
Dim SumPrice As Double
Dim TempDataStr As CStcrd
Dim Str  As CStcrd
Dim HaveSaleFlag As Boolean
Dim HaveFreeFlag As Boolean
Dim HaveHeader As Boolean
   RName = "genDoc"
'-----------------------------------------------------------------------------------------------------
'                                             Query Here
'-----------------------------------------------------------------------------------------------------
   
      
   If Not initDoc Then Exit Function
   
   Vsp.Preview = isPreview
'-----------------------------------------------------------------------------------------------------
'                                         Main Operation Here
'-----------------------------------------------------------------------------------------------------
   
   For j = 1 To UBound(Total1)
      Total1(j) = 0
      Total2(j) = 0
      Total3(j) = 0
      Total4(j) = 0
   Next j
   
   Call StartExportFile(Vsp)
   Vsp.StartDoc
   
   Set BodyCf = New CReportField
   i = 0
   
   If i = 0 Then
      GenerateItemNewPage
   End If
      
   For Each Str In ColCustomerStockSum
      Set TempDataStr2 = GetObject("CCostProducts", ColCostProducts, Trim(Str.STKCOD), False)
      Call m_Details1.ClearField
      '===Head
      K = 0
      SumAmount = 0
      SumPrice = 0
      
      PrevKey3 = Str.SLMCOD
      
      If mcolParam("SPLIT_SALE") = "Y" Then
         If PrevKey4 <> Str.SLMCOD Then
'            If Str.CUSCOD = "" Then
               Vsp = Str.SLMNAM & " (" & (Str.SLMCOD) & ")"
               PrevKey4 = PrevKey3
               i = 0
'            End If
         End If
      End If

      If mcolParam("SHOW_INFOMATION") = "Y" Then
         If Not TempDataStr2 Is Nothing Then
            i = i + 1
            K = K + 1
            Set HeadCf = m_Headers1.Fields(K)
            TempStr = i
            Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr, HeadCf.BodyAlign)
            Call m_Details1.AddField(BodyCf)
                  
            K = K + 1
            Set HeadCf = m_Headers1.Fields(K)
            TempStr = Year(Str.DOCDAT)
            Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr, HeadCf.BodyAlign)
            Call m_Details1.AddField(BodyCf)
                  
            K = K + 1
            Set HeadCf = m_Headers1.Fields(K)
            TempStr = Format(Month(Str.DOCDAT), "00")
            Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr, HeadCf.BodyAlign)
            Call m_Details1.AddField(BodyCf)
                  
            K = K + 1
            Set HeadCf = m_Headers1.Fields(K)
            TempStr = Str.CUSCOD
            Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr, HeadCf.BodyAlign)
            Call m_Details1.AddField(BodyCf)
                  
            K = K + 1
            Set HeadCf = m_Headers1.Fields(K)
            TempStr = Str.CUSNAM
            Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr, HeadCf.BodyAlign)
            Call m_Details1.AddField(BodyCf)
            
            K = K + 1
            Set HeadCf = m_Headers1.Fields(K)
            TempStr = GetProvince(Str.ADDR02)
            Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr, HeadCf.BodyAlign)
            Call m_Details1.AddField(BodyCf)
      
            K = K + 1
            Set HeadCf = m_Headers1.Fields(K)
            TempStr = Str.STKCOD
            Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr, HeadCf.BodyAlign)
            Call m_Details1.AddField(BodyCf)
      
            K = K + 1
            Set HeadCf = m_Headers1.Fields(K)
            TempStr = ""
            Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr, HeadCf.BodyAlign)
            Call m_Details1.AddField(BodyCf)
      
            K = K + 1
            Set HeadCf = m_Headers1.Fields(K)
            TempStr = Str.STKDES
            Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr, HeadCf.BodyAlign)
            Call m_Details1.AddField(BodyCf)
            '-----------------------------------------------------------------------------------------------------------------------------
                  
            K = K + 1
            Set HeadCf = m_Headers1.Fields(K)
            Amt = Str.TRNQTY
            TempStr = FormatNumberToNull(Amt, 0)
            Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr, HeadCf.BodyAlign)
            Call m_Details1.AddField(BodyCf)
      
            K = K + 1
            Set HeadCf = m_Headers1.Fields(K)
            If Not TempDataStr2 Is Nothing Then
               TempStr = TempDataStr2.COST_PRODUCT
               SumAmount = TempDataStr2.COST_PRODUCT
            Else
               TempStr = ""
            End If
            Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr, HeadCf.BodyAlign)
            Call m_Details1.AddField(BodyCf)
      
            K = K + 1
            Set HeadCf = m_Headers1.Fields(K)
            If Str.TRNQTY > 0 And SumAmount > 0 Then
                  TempStr = FormatNumberToNull(Str.TRNQTY * SumAmount)
            Else
            TempStr = ""
            End If
            Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr, HeadCf.BodyAlign)
            Call m_Details1.AddField(BodyCf)
      
            K = K + 1
            Set HeadCf = m_Headers1.Fields(K)
            TempStr = FormatNumberToNull(Str.AMOUNT, 0)
            Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr, HeadCf.BodyAlign)
            Call m_Details1.AddField(BodyCf)
      
            K = K + 1
            Set HeadCf = m_Headers1.Fields(K)
            TempStr = ""
            Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr, HeadCf.BodyAlign)
            Call m_Details1.AddField(BodyCf)
         
            Call m_Details1.GetString(1, TempStr1, TempStr2)
            strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
            Call Vsp.AddTable(strFormat, "", TempStr2)
         
         End If
      Else
      
         i = i + 1
         K = K + 1
         Set HeadCf = m_Headers1.Fields(K)
         TempStr = i
         Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr, HeadCf.BodyAlign)
         Call m_Details1.AddField(BodyCf)
               
         K = K + 1
         Set HeadCf = m_Headers1.Fields(K)
         TempStr = Year(Str.DOCDAT)
         Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr, HeadCf.BodyAlign)
         Call m_Details1.AddField(BodyCf)
               
         K = K + 1
         Set HeadCf = m_Headers1.Fields(K)
         TempStr = Format(Month(Str.DOCDAT), "00")
         Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr, HeadCf.BodyAlign)
         Call m_Details1.AddField(BodyCf)
               
         K = K + 1
         Set HeadCf = m_Headers1.Fields(K)
         TempStr = Str.CUSCOD
         Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr, HeadCf.BodyAlign)
         Call m_Details1.AddField(BodyCf)
               
         K = K + 1
         Set HeadCf = m_Headers1.Fields(K)
         TempStr = Str.CUSNAM
         Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr, HeadCf.BodyAlign)
         Call m_Details1.AddField(BodyCf)
         
         K = K + 1
         Set HeadCf = m_Headers1.Fields(K)
         TempStr = GetProvince(Str.ADDR02)
         Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr, HeadCf.BodyAlign)
         Call m_Details1.AddField(BodyCf)
   
         K = K + 1
         Set HeadCf = m_Headers1.Fields(K)
         TempStr = Str.STKCOD
         Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr, HeadCf.BodyAlign)
         Call m_Details1.AddField(BodyCf)
   
         K = K + 1
         Set HeadCf = m_Headers1.Fields(K)
         TempStr = ""
         Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr, HeadCf.BodyAlign)
         Call m_Details1.AddField(BodyCf)
   
         K = K + 1
         Set HeadCf = m_Headers1.Fields(K)
         TempStr = Str.STKDES
         Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr, HeadCf.BodyAlign)
         Call m_Details1.AddField(BodyCf)
         '-----------------------------------------------------------------------------------------------------------------------------
               
         K = K + 1
         Set HeadCf = m_Headers1.Fields(K)
         Amt = Str.TRNQTY
         TempStr = FormatNumberToNull(Amt, 0)
         Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr, HeadCf.BodyAlign)
         Call m_Details1.AddField(BodyCf)
   
         K = K + 1
         Set HeadCf = m_Headers1.Fields(K)
         If Not TempDataStr2 Is Nothing Then
            TempStr = TempDataStr2.COST_PRODUCT
            SumAmount = TempDataStr2.COST_PRODUCT
         Else
            TempStr = ""
         End If
         Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr, HeadCf.BodyAlign)
         Call m_Details1.AddField(BodyCf)
   
         K = K + 1
         Set HeadCf = m_Headers1.Fields(K)
         If Str.TRNQTY > 0 And SumAmount > 0 Then
               TempStr = FormatNumberToNull(Str.TRNQTY * SumAmount)
         Else
         TempStr = ""
         End If
         Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr, HeadCf.BodyAlign)
         Call m_Details1.AddField(BodyCf)
   
         K = K + 1
         Set HeadCf = m_Headers1.Fields(K)
         TempStr = FormatNumberToNull(Str.AMOUNT, 0)
         Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr, HeadCf.BodyAlign)
         Call m_Details1.AddField(BodyCf)
   
         K = K + 1
         Set HeadCf = m_Headers1.Fields(K)
         TempStr = ""
         Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr, HeadCf.BodyAlign)
         Call m_Details1.AddField(BodyCf)
            
         Call m_Details1.GetString(1, TempStr1, TempStr2)
         strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
         Call Vsp.AddTable(strFormat, "", TempStr2)
      
      End If
   Next Str
      
   Set BodyCf = Nothing
   Vsp.EndDoc
   Call CloseExportFile(Vsp)
   
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
      ElseIf ((j - 2) Mod 3) = 2 Then
         If mcolParam("ONLY_AMOUNT") = "N" And mcolParam("ONLY_PRICE") = "N" Then
            TempStr = FormatNumberToNull(MyDiff(Tot(j + 1), Tot(j - 1)))
         Else
            TempStr = FormatNumberToNull(Tot(j))
         End If
      ElseIf j > 2 Then
         TempStr = FormatNumberToNull(Tot(j))
      End If
      
      Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.Align, TempStr, HeadCf.BodyAlign)
      Call Details.AddField(BodyCf)
   Next HeadCf
      
   Set BodyCf = Nothing
End Sub

Private Sub GenerateHeader(Txt1 As String, Details As CFieldList)
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
   If Vsp.PageCount > 1 Then
      Call GenerateItemNewPage
   End If
End Sub
Private Sub GenerateItemNewPage()
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
Private Sub GenerateListCollection(SourceCollection As Collection, TempCollection As Collection)
Dim Str As CStcrd
Dim TempStr As CStcrd
   For Each Str In SourceCollection
'      If Str.SLMCOD = "22" And Str.CUSCOD = "10-036" And Str.STKCOD = "V160-1" And (Format(Month(Str.DOCDAT), "00") & "/" & Year(Str.DOCDAT) = "12/2017") Then
'         Debug.Print
'      End If
      
      Set TempStr = GetObject("CStcrd", TempCollection, Trim(Str.CUSCOD & "-" & Str.STKCOD & "-" & Format(Month(Str.DOCDAT), "00") & "/" & Year(Str.DOCDAT)), False)
      If TempStr Is Nothing Then
         Set TempStr = New CStcrd
         TempStr.CUSCOD = Str.CUSCOD
         TempStr.CUSNAM = Str.CUSNAM
         TempStr.STKCOD = Str.STKCOD
         TempStr.STKDES = Str.STKDES
         TempStr.DOCDAT = Str.DOCDAT
         TempStr.ADDR02 = Str.ADDR02
         TempStr.SLMCOD = Str.SLMCOD
         TempStr.SLMNAM = Str.SLMNAM
         If Str.RECTYP = 3 Then
            TempStr.UNITPR = Str.UNITPR
            If Val(Str.NETVAL) > 0 Then
               TempStr.TRNQTY = Str.TRNQTY   'ตัวไม่ฟรี
            Else
               TempStr.AMOUNT = Str.TRNQTY   'ตัวฟรี
            End If
            
            TempStr.NETVAL = Str.NETVAL
         ElseIf Str.RECTYP = 5 Then
            TempStr.UNITPR = -1 * Str.UNITPR
            If Val(Str.NETVAL) > 0 Then
               TempStr.TRNQTY = -1 * Str.TRNQTY   'ตัวไม่ฟรี
            Else
               TempStr.AMOUNT = -1 * Str.TRNQTY   'ตัวฟรี
            End If
            TempStr.NETVAL = -1 * Str.NETVAL
         End If
         Call TempCollection.Add(TempStr, Trim(Str.CUSCOD & "-" & Str.STKCOD & "-" & Format(Month(Str.DOCDAT), "00") & "/" & Year(Str.DOCDAT)))
      Else
         If Str.RECTYP = 3 Then
            TempStr.UNITPR = TempStr.UNITPR + Str.UNITPR
            If Val(Str.NETVAL) > 0 Then
               TempStr.TRNQTY = TempStr.TRNQTY + Str.TRNQTY    'ตัวไม่ฟรี
            Else
               TempStr.AMOUNT = TempStr.AMOUNT + Str.TRNQTY    'ตัวฟรี
            End If
            
            TempStr.NETVAL = TempStr.NETVAL + Str.NETVAL
         Else
            TempStr.UNITPR = TempStr.UNITPR - Str.UNITPR
            If Val(Str.NETVAL) > 0 Then
               TempStr.TRNQTY = TempStr.TRNQTY - Str.TRNQTY    'ตัวไม่ฟรี
            Else
               TempStr.AMOUNT = TempStr.AMOUNT - Str.TRNQTY    'ตัวฟรี
            End If
            TempStr.NETVAL = TempStr.NETVAL - Str.NETVAL
         End If
      End If
   Next Str
End Sub
Private Function CalculateFreeFlag(Amt As Double) As String
   If Amt > 0 Then
      If SaleFlag Then
         CalculateFreeFlag = "N"
      End If
   Else
      If FreeFlag Then
         CalculateFreeFlag = "Y"
      End If
   End If
End Function

Private Function GetProvince(TempAddress As String) As String
Dim mProvince As CProvinceMap
   GetProvince = ""
   For Each mProvince In ColProvinceMap
      If InStr(1, TempAddress, mProvince.KEY_SEARCH) > 0 Then
         GetProvince = mProvince.KEY_MAP
         Exit For
      End If
   Next mProvince
   
End Function
