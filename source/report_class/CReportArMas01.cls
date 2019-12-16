VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CReportArMas01"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'User Group Report
Implements CReportInterface

Private Const MODULE_NAME = "CReportArMas01"
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
Private m_CUS_PIG_TYPE  As Collection
Private ColCustomerStockVac As Collection
Private ColCustomerStockNonVac As Collection

Private ColProvinceMap As Collection
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
   Set m_CUS_PIG_TYPE = New Collection
   Set ColCustomerStockVac = New Collection
   Set ColCustomerStockNonVac = New Collection
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
   Set m_CustomerTypes = Nothing
   Set m_CUS_PIG_TYPE = Nothing
   Set ColCustomerStockVac = Nothing
   Set ColCustomerStockNonVac = Nothing
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
   
   If mcolParam("SHOW_NO") = "Y" Then
      Call Cf.SetFieldValue(8, "^", "NO.", "^")
      Call m_Headers1.AddField(Cf)
   End If
   Call Cf.SetFieldValue(15, "^", "รหัสลูกค้า", "<")
   Call m_Headers1.AddField(Cf)
   Call Cf.SetFieldValue(50, "^", "ชื่อลูกค้า", "<")
   Call m_Headers1.AddField(Cf)
   
   If mcolParam("SHOW_ADDRESS") = "Y" Then
      Call Cf.SetFieldValue(65, "^", "ที่อยู่", "<")
      Call m_Headers1.AddField(Cf)
   End If
   
   If mcolParam("SHOW_PROVINCE") = "Y" Then
      Call Cf.SetFieldValue(65, "^", "จังหวัด", "<")
      Call m_Headers1.AddField(Cf)
   End If
   
   If mcolParam("SHOW_TEL") = "Y" Then
      Call Cf.SetFieldValue(45, "^", "โทรศัพท์", "<")
      Call m_Headers1.AddField(Cf)
   End If
   
   If mcolParam("SHOW_CREDIT") = "Y" Then
      Call Cf.SetFieldValue(10, "^", "เครดิต", "<")
      Call m_Headers1.AddField(Cf)
   End If
   
'   If mcolParam("SHOW_SALE") = "Y" Then
'      Call Cf.SetFieldValue(10, "^", "พข", "<")
'      Call m_Headers1.AddField(Cf)
'   End If
   
   If mcolParam("SHOW_AREA") = "Y" Then
      Call Cf.SetFieldValue(10, "^", "เขต", "<")
      Call m_Headers1.AddField(Cf)
   End If
   
   If mcolParam("SHOW_LIMIT") = "Y" Then
      Call Cf.SetFieldValue(15, "^", "วงเงิน", ">")
      Call m_Headers1.AddField(Cf)
   End If
   '
   If mcolParam("SHOW_CONTRACT") = "Y" Then
      Call Cf.SetFieldValue(20, "^", "ชื่อผู้ติดต่อ", "<")
      Call m_Headers1.AddField(Cf)
   End If
   
   If mcolParam("SHOW_CONTRACT") = "Y" Then
      Call Cf.SetFieldValue(20, "^", "พนักงานขาย", "<")
      Call m_Headers1.AddField(Cf)
   End If
   
    If mcolParam("SHOW_PIGDATA") = "Y" Then
      Call Cf.SetFieldValue(20, "^", "หมูพันธุ์", ">", "SUM_FLAG")
      Call m_Headers1.AddField(Cf)
      
      Call Cf.SetFieldValue(20, "^", "หมูขุน", ">", "SUM_FLAG")
      Call m_Headers1.AddField(Cf)
      
      Call Cf.SetFieldValue(20, "^", "ลูกหมู", ">", "SUM_FLAG")
      Call m_Headers1.AddField(Cf)
   End If
   
    If mcolParam("SHOW_VAC_NONVAC") = "Y" Then
      Call Cf.SetFieldValue(20, "^", "วัคซีน", ">", "SUM_FLAG")
      Call m_Headers1.AddField(Cf)
      
      Call Cf.SetFieldValue(20, "^", "อื่นๆ", ">", "SUM_FLAG")
      Call m_Headers1.AddField(Cf)
   End If
   
   Call m_Headers1.GetString(1, TempStr1, TempStr1_1)
   strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
   strPrint = TempStr1_1
   Call Vsp.AddTable(strFormat, "", strPrint)
    
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
Dim FromDate As Date
Dim ToDate As Date
Dim ToPaidDate As String
Dim FROM_STOCK_CODE As String
Dim TO_STOCK_CODE As String


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
   
   FromDate = mcolParam("FROM_DOC_DATE")
   ToDate = mcolParam("TO_DOC_DATE")
   FROM_STOCK_CODE = "V000-0"
   TO_STOCK_CODE = "V999-9"
      
   Call LoadCustomerType(Nothing, m_CustomerTypes)
   Call LoadCusPigType(m_CUS_PIG_TYPE)
   
   Call SetReportConfig(Vsp, mcolParam("REPORT_KEY"))

   Call LoadSumSaleCustomerStcrdVac(ColCustomerStockVac, FromDate, ToDate, mcolParam("FROM_CUSTOMER_CODE"), mcolParam("TO_CUSTOMER_CODE"), FROM_STOCK_CODE, TO_STOCK_CODE, mcolParam("FROM_SALE_CODE"), mcolParam("TO_SALE_CODE"))
   Call LoadSumSaleCustomerStcrdNonVac(ColCustomerStockNonVac, FromDate, ToDate, mcolParam("FROM_CUSTOMER_CODE"), mcolParam("TO_CUSTOMER_CODE"), FROM_STOCK_CODE, TO_STOCK_CODE, mcolParam("FROM_SALE_CODE"), mcolParam("TO_SALE_CODE"))
   
   Call LoadProvinceMap(Nothing, ColProvinceMap)
   
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
Dim currentKey1 As String
Dim St As CIsTab
Dim Ar As CARMas
Dim tempCusPigType As CCusPigType
Dim tempStcrdIV As CStcrd
Dim tempStcrdSR As CStcrd
Dim tempStcrdDN As CStcrd
Dim sumTempStcrdVac As Double
Dim sumTempStcrdNonVac As Double

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
   If mcolParam("CUSTOMER_TYPE") > 0 Then
         Set St = GetItemFromListIndex(m_CustomerTypes, mcolParam("CUSTOMER_TYPE"))
         Ar.CUSTYP = St.TYPCOD
      Else
         Ar.CUSTYP = ""
      End If
'   Ar.OrderBy = mcolParam("ORDER_BY")
'   Ar.OrderType = mcolParam("ORDER_TYPE")
   
   Call Ar.QueryData(1, Rs, iCount)
      
   i = 0
   j = 0
PrevKey1 = ""
   While Not Rs.EOF
         Call Ar.PopulateFromRS(1, Rs)
         
         If mcolParam("SHOW_SALE") = "Y" Then
            currentKey1 = Ar.SLMCOD
            If PrevKey1 <> currentKey1 And i > 0 Then
               Vsp.FontBold = True
               Call GenerateFooter("รวม", PrevKey1, "", Total1, m_Details1, 0)
               Call m_Details1.GetString(2, TempStr1, TempStr2)
               strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
               Call Vsp.AddTable(strFormat, "", TempStr2)
               Vsp.FontBold = False
      
               For j = 1 To UBound(Total1)
                   Total1(j) = 0
               Next j
               Vsp = Ar.SLMNAM & "(" & Ar.SLMCOD & ")"
               Call printHeader
            ElseIf i = 0 Then
                Vsp = Ar.SLMNAM & "(" & Ar.SLMCOD & ")"
            End If
         End If
      i = i + 1
      PrevKey1 = Ar.SLMCOD
      
      Call m_Details1.ClearField
'''      If i = 1 Then
'''         Call printHeader
'''      End If
      
      j = 0
      
      If mcolParam("SHOW_NO") = "Y" Then
         j = j + 1
         Set HeadCf = m_Headers1.Fields(j)
         TempStr = i
         Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
         Call m_Details1.AddField(BodyCf)
      End If
      
      j = j + 1
      Set HeadCf = m_Headers1.Fields(j)
      TempStr = Ar.CUSCOD
      Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
      Call m_Details1.AddField(BodyCf)
      
      j = j + 1
      Set HeadCf = m_Headers1.Fields(j)
      If mcolParam("SHOW_PREFIX") = "Y" Then
         TempStr = Ar.PRENAM & " " & Ar.CUSNAM
      Else
         TempStr = Ar.CUSNAM
      End If
      Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
      Call m_Details1.AddField(BodyCf)
      
      If mcolParam("SHOW_ADDRESS") = "Y" Then
         j = j + 1
         Set HeadCf = m_Headers1.Fields(j)
         TempStr = Ar.ADDR01 & " " & Ar.ADDR02
         Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
         Call m_Details1.AddField(BodyCf)
      End If
      
      If mcolParam("SHOW_PROVINCE") = "Y" Then
         j = j + 1
         Set HeadCf = m_Headers1.Fields(j)
         TempStr = GetProvince(Ar.ADDR02)
         Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
         Call m_Details1.AddField(BodyCf)
      End If
      
      If mcolParam("SHOW_TEL") = "Y" Then
         j = j + 1
         Set HeadCf = m_Headers1.Fields(j)
         TempStr = Ar.TELNUM
         Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
         Call m_Details1.AddField(BodyCf)
      End If
      
      If mcolParam("SHOW_CREDIT") = "Y" Then
         j = j + 1
         Set HeadCf = m_Headers1.Fields(j)
         TempStr = Ar.PAYTRM
         Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
         Call m_Details1.AddField(BodyCf)
      End If
      
'      If mcolParam("SHOW_SALE") = "Y" Then
'         j = j + 1
'         Set HeadCf = m_Headers1.Fields(j)
'         TempStr = Ar.SLMCOD
'         Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
'         Call m_Details1.AddField(BodyCf)
'      End If
      
      If mcolParam("SHOW_AREA") = "Y" Then
         j = j + 1
         Set HeadCf = m_Headers1.Fields(j)
         TempStr = Ar.AREACOD
         Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
         Call m_Details1.AddField(BodyCf)
      End If
      
      If mcolParam("SHOW_LIMIT") = "Y" Then
         j = j + 1
         Set HeadCf = m_Headers1.Fields(j)
         TempStr = FormatNumberToNull(Ar.CRLINE, 0)
         Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
         Call m_Details1.AddField(BodyCf)
         Total1(j) = Total1(j) + Ar.CRLINE
      End If
      
      If mcolParam("SHOW_CONTRACT") = "Y" Then
         j = j + 1
         Set HeadCf = m_Headers1.Fields(j)
         TempStr = Ar.CONTRACT
         Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
         Call m_Details1.AddField(BodyCf)
      End If
      
      If mcolParam("SHOW_CONTRACT") = "Y" Then
         j = j + 1
         Set HeadCf = m_Headers1.Fields(j)
         TempStr = Ar.SLMNAM & "-" & Ar.SLMCOD
         Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
         Call m_Details1.AddField(BodyCf)
      End If
      
      Set tempCusPigType = GetObject("CCusPigType", m_CUS_PIG_TYPE, Ar.CUSCOD, False)
         If Not tempCusPigType Is Nothing Then
                If mcolParam("SHOW_PIGDATA") = "Y" Then
                     j = j + 1
                     Set HeadCf = m_Headers1.Fields(j)
                     TempStr = FormatNumberToNull(tempCusPigType.CUS_PIG_TYPE_BREED, 0)
                     Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
                     Call m_Details1.AddField(BodyCf)
                     Total1(j) = Total1(j) + tempCusPigType.CUS_PIG_TYPE_BREED
                     
                      j = j + 1
                     Set HeadCf = m_Headers1.Fields(j)
                     TempStr = FormatNumberToNull(tempCusPigType.CUS_PIG_TYPE_KHUN, 0)
                     Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
                     Call m_Details1.AddField(BodyCf)
                      Total1(j) = Total1(j) + tempCusPigType.CUS_PIG_TYPE_KHUN
                     
                      j = j + 1
                     Set HeadCf = m_Headers1.Fields(j)
                     TempStr = FormatNumberToNull(tempCusPigType.CUS_PIG_TYPE_PIGGY, 0)
                     Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
                     Call m_Details1.AddField(BodyCf)
                     Total1(j) = Total1(j) + tempCusPigType.CUS_PIG_TYPE_PIGGY
                End If
         Else
             If mcolParam("SHOW_PIGDATA") = "Y" Then
                     j = j + 1
                     Set HeadCf = m_Headers1.Fields(j)
                     TempStr = ""
                     Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
                     Call m_Details1.AddField(BodyCf)
                     
                     
                     j = j + 1
                     Set HeadCf = m_Headers1.Fields(j)
                     TempStr = ""
                     Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
                     Call m_Details1.AddField(BodyCf)
                  
                     j = j + 1
                     Set HeadCf = m_Headers1.Fields(j)
                     TempStr = ""
                     Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
                     Call m_Details1.AddField(BodyCf)
                End If
         End If
      Set tempCusPigType = Nothing
      
      'getSumVac
      Set tempStcrdIV = GetObject("CStcrd", ColCustomerStockVac, Ar.CUSCOD & "-3")
      Set tempStcrdSR = GetObject("CStcrd", ColCustomerStockVac, Ar.CUSCOD & "-4")
      Set tempStcrdDN = GetObject("CStcrd", ColCustomerStockVac, Ar.CUSCOD & "-5")
      sumTempStcrdVac = tempStcrdIV.NETVAL - tempStcrdSR.NETVAL + tempStcrdDN.NETVAL
    
      
      'getSumNonVac
      Set tempStcrdIV = GetObject("CStcrd", ColCustomerStockNonVac, Ar.CUSCOD & "-3")
      Set tempStcrdSR = GetObject("CStcrd", ColCustomerStockNonVac, Ar.CUSCOD & "-4")
      Set tempStcrdDN = GetObject("CStcrd", ColCustomerStockNonVac, Ar.CUSCOD & "-5")
      sumTempStcrdNonVac = (tempStcrdIV.NETVAL - tempStcrdSR.NETVAL + tempStcrdDN.NETVAL) - sumTempStcrdVac
      
      If mcolParam("SHOW_VAC_NONVAC") = "Y" Then
           j = j + 1
           Set HeadCf = m_Headers1.Fields(j)
           TempStr = FormatNumberToNull(sumTempStcrdVac, 0)
           Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
           Call m_Details1.AddField(BodyCf)
           Total1(j) = Total1(j) + sumTempStcrdVac
           
            j = j + 1
           Set HeadCf = m_Headers1.Fields(j)
           TempStr = FormatNumberToNull(sumTempStcrdNonVac, 0)
           Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
           Call m_Details1.AddField(BodyCf)
           Total1(j) = Total1(j) + sumTempStcrdNonVac
      End If
      Set tempCusPigType = Nothing
      
      Call m_Details1.GetString(1, TempStr1, TempStr2)
      strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
      Call Vsp.AddTable(strFormat, "", TempStr2)
      
      Rs.MoveNext
   Wend
   
   If mcolParam("SHOW_SALE") = "Y" Then
      Vsp.FontBold = True
      Call GenerateFooter("รวม", PrevKey1, "", Total1, m_Details1, 0)
      Call m_Details1.GetString(2, TempStr1, TempStr2)
      strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
      Call Vsp.AddTable(strFormat, "", TempStr2)
      Vsp.FontBold = False
   End If
   
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
'      ElseIf j = 3 Then
'         TempStr = Txt3
'      ElseIf (j = 4) Then
'         TempStr = ""
      ElseIf HeadCf.TempKey = "SUM_FLAG" Then
         TempStr = FormatNumberToNull(Tot(j), 0)
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
