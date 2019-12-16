VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CReportAP006_1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'User Group Report
Implements CReportInterface

Private Const MODULE_NAME = "CReportAP002"
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
Private m_Left As Double
Private m_YearSeqs As Collection

Private m_SupplierTypes As Collection
Private m_StockGroups As Collection
Private m_PaidAmounts As Collection
Private m_PaidAllAmounts As Collection
Private m_MaxMins As Collection
Private m_ApSupAmounts As Collection
Private m_ApSupPaids As Collection
Private m_DbnCdnByBills As Collection

Private HaveHeader As Boolean

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
   Set m_StockGroups = New Collection
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
   Set m_StockGroups = Nothing
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
Dim SumHeader As Double

Dim TempFromdate As Date
Dim TempToDate As Date
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
   
   SumHeader = 0
   '===
   If mcolParam("NO_SHOW_BILL") = "N" And mcolParam("SUMMARY_MODE") = "N" Then
      Call Cf.SetFieldValue(6, "^", "รายละเอียด", "^")
      Call m_Headers0.AddField(Cf)
   
      Call Cf.SetFieldValue(3, "^", "เลขที่เอกสาร", "^")
      Call m_Headers1.AddField(Cf)
      
      Call Cf.SetFieldValue(3, "^", "วันที่", "^")
      Call m_Headers1.AddField(Cf)
   Else
      Call Cf.SetFieldValue(6, "^", "รายละเอียด", "^")
      Call m_Headers0.AddField(Cf)
   
      Call Cf.SetFieldValue(6, "^", "ผู้ขาย", "<")
      Call m_Headers1.AddField(Cf)
   End If
   
   TempDate = mcolParam("FROM_DOC_DATE")
   TempFromdate = mcolParam("FROM_DOC_DATE")
   TempToDate = mcolParam("TO_DOC_DATE")
   
   While TempDate <= mcolParam("TO_DOC_DATE")
      Call GetFirstLastDate(TempDate, TempFromdate, TempToDate)
      If mcolParam("SUMMARY_COLUMN") = "N" Then
         Call Cf.SetFieldValue(4, "^", Day(TempFromdate) & "-" & Day(TempToDate) & "/" & Format(Month(TempFromdate), "00") & "/" & (Year(TempFromdate) + 543), "<")
         Call m_Headers0.AddField(Cf)
      End If
      TempDate = DateAdd("M", 1, TempDate)
   Wend
   Call Cf.SetFieldValue(6, "^", "รวม", "<")
   Call m_Headers0.AddField(Cf)
   
   
   
   TempDate = mcolParam("FROM_DOC_DATE")
   TempFromdate = mcolParam("FROM_DOC_DATE")
   TempToDate = mcolParam("TO_DOC_DATE")
   
   While TempDate <= mcolParam("TO_DOC_DATE")
      Call GetFirstLastDate(TempDate, TempFromdate, TempToDate)
      If mcolParam("SUMMARY_COLUMN") = "N" Then
         Call Cf.SetFieldValue(2, "^", "จำนวน", ">", "SUM_FLAG")
         Call m_Headers1.AddField(Cf)
         Call Cf.SetFieldValue(2, "^", "ยอดเงิน", ">", "SUM_FLAG")
         Call m_Headers1.AddField(Cf)
      End If
      TempDate = DateAdd("M", 1, TempDate)
   Wend
   
   Call Cf.SetFieldValue(2, "^", "จำนวน", ">", "SUM_FLAG")
   Call m_Headers1.AddField(Cf)
   Call Cf.SetFieldValue(2, "^", "ราคา", ">", "AVG_FLAG")
   Call m_Headers1.AddField(Cf)
   Call Cf.SetFieldValue(2, "^", "ยอดเงิน", ">", "SUM_FLAG")
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
   Vsp.MarginBottom = 300
   Vsp.MarginFooter = 300
   Vsp.MarginHeader = 300
   Vsp.MarginLeft = 300
   Vsp.MarginRight = 300
   Vsp.MarginTop = 300
   Vsp.FontName = "AngsanaUPC"
   Vsp.FontSize = DETAIL_SIZE
   
   Call LoadSupplierType(Nothing, m_SupplierTypes)
   
   FromDate = "จากวันที่ " & EmptyToString(DateToStringExtEx2(mcolParam("FROM_DOC_DATE")), "N/A")
   ToDate = " ถึง วันที่ " & EmptyToString(DateToStringExtEx2(mcolParam("TO_DOC_DATE")), "N/A")
   
   MODULE_DESC = glbCompanyName & vbCrLf & mcolParam("REPORT_NAME") & _
                                       vbCrLf & FromDate & ToDate & "  " & ToPaidDate & _
                                       " จากรหัส RM " & EmptyToString(mcolParam("FROM_STOCK_CODE"), "N/A") & " ถึงรหัส RM " & EmptyToString(mcolParam("TO_STOCK_CODE"), "N/A")
   
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
Dim PrevKey3 As String
Dim PrevKey4 As String
Dim PrevKey5 As String

Dim Stc As CStcrd

Dim St As CIsTab
Dim Ap As CAPMas
Dim Width As Double
Dim Apt As CApTrn
Dim TempRs As ADODB.Recordset
Dim TempApt As CApTrn
Dim TempApRcIt As CAPRcIt
Dim HaveValue As Boolean
Dim Balance As Double

Dim TempFromdate As Date
Dim TempToDate As Date
Dim TempDate As Date

Dim SumAmount As Double
Dim SumPrice As Double
Dim TempAmount As Double
Dim TempPrice As Double

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
   
   
   mblnHeader = False
   Call StartExportFile(Vsp)
   Vsp.StartDoc
   Call printHeader
   mblnHeader = True
   
   
      i = 0
      Set BodyCf = New CReportField
      
      Set Stc = New CStcrd
      Stc.FROM_DOC_DATE = mcolParam("FROM_DOC_DATE")
      Stc.TO_DOC_DATE = mcolParam("TO_DOC_DATE")
      Stc.FROM_SUPPLIER_CODE = mcolParam("FROM_SUPPLIER_CODE")
      Stc.TO_SUPPLIER_CODE = mcolParam("TO_SUPPLIER_CODE")
      Stc.FROM_STOCK_CODE = mcolParam("FROM_STOCK_CODE")
      Stc.TO_STOCK_CODE = mcolParam("TO_STOCK_CODE")
      If mcolParam("SUPPLIER_TYPE") > 0 Then
         Set St = GetItemFromListIndex(m_SupplierTypes, mcolParam("SUPPLIER_TYPE"))
         Stc.SUPTYP = St.TYPCOD
      ElseIf Len(mcolParam("SUPPLIER_TYPE_SET")) > 0 Then
         Stc.SUPTYPSET = mcolParam("SUPPLIER_TYPE_SET")
      Else
         Stc.SUPTYP = ""
      End If
      
      Call Stc.QueryData(25, Rs, iCount)
      
      
      i = 0
      While Not Rs.EOF
         
         Call Stc.PopulateFromRS(25, Rs)
         
         If (PrevKey1 <> Stc.SUPCOD) And (i > 0) Then
            If mcolParam("NO_SHOW_BILL") = "N" Then
               Call GenerateFooter(PrevKey2, "", "", Total2, m_Details1, 0)
            Else
               Call GenerateFooter(PrevKey4 & "(" & PrevKey2 & ")", "", "", Total2, m_Details1, 0)
            End If
            If mcolParam("SUMMARY_MODE") = "N" Then
               Call m_Details1.GetString(2, TempStr1, TempStr2)
               strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
               Call Vsp.AddTable(strFormat, "", TempStr2)
            End If
            
            If mcolParam("SUMMARY_MODE") = "N" Then
               Call GenerateFooter("รวม " & PrevKey1, "", "", Total1, m_Details1, 0)
            Else
               Call GenerateFooter(PrevKey3, "", "", Total1, m_Details1, 0)
            End If
            Call m_Details1.GetString(2, TempStr1, TempStr2)
            strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
            Vsp.FontBold = True
            Call Vsp.AddTable(strFormat, "", TempStr2)
            Vsp.FontBold = False
               
            For j = 1 To UBound(Total1)
               Total1(j) = 0
               Total2(j) = 0
            Next j
            
            If mcolParam("SUMMARY_MODE") = "N" Then
               Vsp = Stc.SUPNAM & " ( " & Stc.SUPCOD & " )"
               If mcolParam("NO_SHOW_BILL") = "N" Then
                  Vsp = "           " & Stc.STKDES & " ( " & Stc.STKCOD & " )"
               End If
            End If
         ElseIf (PrevKey2 <> Stc.STKCOD) And (i > 0) Then
            If mcolParam("NO_SHOW_BILL") = "N" Then
               Call GenerateFooter(PrevKey2, "", "", Total2, m_Details1, 0)
            Else
               Call GenerateFooter(PrevKey4 & "(" & PrevKey2 & ")", "", "", Total2, m_Details1, 0)
            End If
            If mcolParam("SUMMARY_MODE") = "N" Then
               Call m_Details1.GetString(2, TempStr1, TempStr2)
               strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
               Call Vsp.AddTable(strFormat, "", TempStr2)
            End If
            For j = 1 To UBound(Total1)
               Total2(j) = 0
            Next j
            
            If mcolParam("SUMMARY_MODE") = "N" Then
               If mcolParam("NO_SHOW_BILL") = "N" Then
                  Vsp = "           " & Stc.STKDES & " ( " & Stc.STKCOD & " )"
               End If
            End If
         ElseIf i = 0 Then
            If mcolParam("SUMMARY_MODE") = "N" Then
               Vsp = Stc.SUPNAM & " ( " & Stc.SUPCOD & " )"
               If mcolParam("NO_SHOW_BILL") = "N" Then
                  Vsp = "           " & Stc.STKDES & " ( " & Stc.STKCOD & " )"
               End If
            End If
         End If
         
         PrevKey1 = Stc.SUPCOD
         PrevKey2 = Stc.STKCOD
        
         PrevKey3 = Stc.SUPNAM
         PrevKey4 = Stc.STKDES
         
         Call m_Details1.ClearField
         i = i + 1
         
         Set HeadCf = m_Headers1.Fields(1)
         TempStr = Stc.DOCNUM
         Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
        Call m_Details1.AddField(BodyCf)
        
         j = 1
         
         If mcolParam("NO_SHOW_BILL") = "N" And mcolParam("SUMMARY_MODE") = "N" Then
            j = j + 1
            Set HeadCf = m_Headers1.Fields(j)
            TempStr = DateToStringExtEx2(Stc.DOCDAT)
            Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
            Call m_Details1.AddField(BodyCf)
         End If
         
         SumAmount = 0
         SumPrice = 0
         
         TempDate = mcolParam("FROM_DOC_DATE")
         TempFromdate = mcolParam("FROM_DOC_DATE")
         TempToDate = mcolParam("TO_DOC_DATE")
         
         While TempDate <= mcolParam("TO_DOC_DATE")
            Call GetFirstLastDate(TempDate, TempFromdate, TempToDate)
            
            TempAmount = 0
            TempPrice = 0
            If TempFromdate <= Stc.DOCDAT And Stc.DOCDAT <= TempToDate Then
               TempAmount = Stc.TRNQTY
               TempPrice = Stc.NETVAL
               If Stc.POSOPR = "7" Then 'ลดหนี้
                  TempAmount = TempAmount * -1
                  TempPrice = TempPrice * -1
               End If
            End If
                  
            If mcolParam("SUMMARY_COLUMN") = "N" Then
               j = j + 1
               Set HeadCf = m_Headers1.Fields(j)
               TempStr = FormatNumberToNull(TempAmount, 0)
               Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
               Call m_Details1.AddField(BodyCf)
               Total3(j) = Total3(j) + TempAmount
               Total1(j) = Total1(j) + TempAmount
               Total2(j) = Total2(j) + TempAmount
            End If
            SumAmount = SumAmount + TempAmount
            
            If mcolParam("SUMMARY_COLUMN") = "N" Then
               j = j + 1
               Set HeadCf = m_Headers1.Fields(j)
               TempStr = FormatNumberToNull(TempPrice, 0)
               Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
               Call m_Details1.AddField(BodyCf)
               Total3(j) = Total3(j) + TempPrice
               Total1(j) = Total1(j) + TempPrice
               Total2(j) = Total2(j) + TempPrice
            End If
            
            SumPrice = SumPrice + TempPrice
            
            TempDate = DateAdd("M", 1, TempDate)
         Wend
         
          j = j + 1
         Set HeadCf = m_Headers1.Fields(j)
         TempStr = FormatNumberToNull(SumAmount, 0)
         Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
         Call m_Details1.AddField(BodyCf)
         
         Total3(j) = Total3(j) + SumAmount
         Total1(j) = Total1(j) + SumAmount
         Total2(j) = Total2(j) + SumAmount
         
         j = j + 1
         Set HeadCf = m_Headers1.Fields(j)
         TempStr = FormatNumberToNull(MyDiff(SumPrice, SumAmount), 0)
         Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
         Call m_Details1.AddField(BodyCf)
         
         j = j + 1
         Set HeadCf = m_Headers1.Fields(j)
         TempStr = FormatNumberToNull(SumPrice, 0)
         Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
         Call m_Details1.AddField(BodyCf)
         
         Total3(j) = Total3(j) + SumPrice
         Total1(j) = Total1(j) + SumPrice
         Total2(j) = Total2(j) + SumPrice
         
         If mcolParam("SUMMARY_MODE") = "N" Then
            If mcolParam("NO_SHOW_BILL") = "N" Then
               Call m_Details1.GetString(1, TempStr1, TempStr2)
               strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
               Call Vsp.AddTable(strFormat, "", TempStr2)
            End If
         End If
         
         Rs.MoveNext
      Wend
      
      If mcolParam("NO_SHOW_BILL") = "N" Then
         Call GenerateFooter(PrevKey2, "", "", Total2, m_Details1, 0)
      Else
         Call GenerateFooter(PrevKey4 & "(" & PrevKey2 & ")", "", "", Total2, m_Details1, 0)
      End If
      If mcolParam("SUMMARY_MODE") = "N" Then
         Call m_Details1.GetString(2, TempStr1, TempStr2)
         strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
         Call Vsp.AddTable(strFormat, "", TempStr2)
      End If
      
      If mcolParam("SUMMARY_MODE") = "N" Then
         Call GenerateFooter("รวม " & PrevKey1, "", "", Total1, m_Details1, 0)
      Else
         Call GenerateFooter(PrevKey3, "", "", Total1, m_Details1, 0)
      End If
      Call m_Details1.GetString(2, TempStr1, TempStr2)
      strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
      Vsp.FontBold = True
      Call Vsp.AddTable(strFormat, "", TempStr2)
      Vsp.FontBold = False
                  
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

   Call Details.ClearField
   Set BodyCf = New CReportField
   
   j = 0
   For Each HeadCf In m_Headers1.Fields
      j = j + 1
      If j = 1 Then
         TempStr = Txt1
      ElseIf j > 1 Then
         If HeadCf.TempKey = "AVG_FLAG" Then
            TempStr = FormatNumberToNull(MyDiff(Tot(j + 1), Tot(j - 1)), 0)
         ElseIf HeadCf.TempKey = "SUM_FLAG" Then
            TempStr = FormatNumberToNull(Tot(j), 0)
         Else
            TempStr = ""
         End If
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
