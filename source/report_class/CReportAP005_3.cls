VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CReportAP005_3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'User Group Report
Implements CReportInterface

Private Const MODULE_NAME = "CReportAP005"
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
Private m_ChqnumDocDate As Collection
Private m_DocumentCancel As Collection
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
   Set m_ChqnumDocDate = New Collection
   Set m_DocumentCancel = New Collection
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
   Set m_ChqnumDocDate = Nothing
   Set m_DocumentCancel = Nothing
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
   
   Call m_Headers0.ClearField
   Call m_Headers1.ClearField
   
   Set Cf = New CReportField
   
   If mcolParam("SHOW_SUMMARY") = "N" Then
      Call Cf.SetFieldValue(1, "^", "NO.", "^")
      Call m_Headers1.AddField(Cf)
   Else
      Call Cf.SetFieldValue(0, "^", "NO.", "^")
      Call m_Headers1.AddField(Cf)
   End If
   If mcolParam("SHOW_SUMMARY") = "N" Then
      Call Cf.SetFieldValue(3, "^", "เลขที่เช็ค", "^")
      Call m_Headers1.AddField(Cf)
   Else
      Call Cf.SetFieldValue(0, "^", "เลขที่เช็ค", "^")
      Call m_Headers1.AddField(Cf)
   End If
   
   Call Cf.SetFieldValue(3, "^", "วันที่ดิวเช็ค", "^")
   Call m_Headers1.AddField(Cf)
   
   If mcolParam("SHOW_SUMMARY") = "N" Then
      Call Cf.SetFieldValue(7, "^", "ชื่อผู้ขาย", "<")
      Call m_Headers1.AddField(Cf)
   Else
      Call Cf.SetFieldValue(0, "^", "ชื่อผู้ขาย", "<")
      Call m_Headers1.AddField(Cf)
   End If
   If mcolParam("SHOW_SUMMARY") = "N" Then
      Call Cf.SetFieldValue(4, "^", "ประเภท", "<")
      Call m_Headers1.AddField(Cf)
   Else
      Call Cf.SetFieldValue(0, "^", "ประเภท", "<")
      Call m_Headers1.AddField(Cf)
   End If
   If mcolParam("SHOW_SUMMARY") = "N" Then
      Call Cf.SetFieldValue(3, "^", "ใบสำคัญจ่าย", "^")
      Call m_Headers1.AddField(Cf)
   Else
      Call Cf.SetFieldValue(0, "^", "ใบสำคัญจ่าย", "^")
      Call m_Headers1.AddField(Cf)
   End If
   Call Cf.SetFieldValue(2, "^", "จำนวนเงิน", ">")
   Call m_Headers1.AddField(Cf)
   If mcolParam("SHOW_SUMMARY") = "N" Then
      Call Cf.SetFieldValue(1, "^", "เครดิต", "^")
      Call m_Headers1.AddField(Cf)
      Call Cf.SetFieldValue(2, "^", "วันที่ส่งของ", ">")
      Call m_Headers1.AddField(Cf)
   Else
      Call Cf.SetFieldValue(0, "^", "เครดิต", "^")
      Call m_Headers1.AddField(Cf)
      Call Cf.SetFieldValue(0, "^", "วันที่ส่งของ", ">")
      Call m_Headers1.AddField(Cf)
   End If
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
Dim PrintDate As String
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
   Vsp.MarginRight = 500
   Vsp.MarginTop = 300
   Vsp.FontName = "AngsanaUPC"
   Vsp.FontSize = DETAIL_SIZE
      
   Call LoadSupplierType(Nothing, m_SupplierTypes)
   
   FromDate = "จากวันที่ " & EmptyToString(DateToStringExtEx2(mcolParam("FROM_CHECK_DATE")), "N/A")
   ToDate = " ถึง วันที่ " & EmptyToString(DateToStringExtEx2(mcolParam("TO_CHECK_DATE")), "N/A")
   PrintDate = "    วันที่พิมพ์ " & EmptyToString(DateToStringExtEx2(mcolParam("PRINT_DATE")), "N/A")
   ToPaidDate = " วันที่รับชำระ " & EmptyToString(DateToStringExtEx2(mcolParam("TO_PAY_DATE")), "N/A")
   
   MODULE_DESC = glbCompanyName & vbCrLf & mcolParam("REPORT_NAME") & _
                                       vbCrLf & FromDate & ToDate & ToPaidDate & PrintDate & " " & mcolParam("FROM_DOCUMENT_NO") & "-" & mcolParam("TO_DOCUMENT_NO") & "," & mcolParam("FROM_DOCUMENT_NO1") & "-" & mcolParam("TO_DOCUMENT_NO1")
   
   Call SetReportConfig(Vsp, mcolParam("REPORT_KEY"))
   Call LoadBktChqnumDocDate(Nothing, m_ChqnumDocDate, , , mcolParam("FROM_CHECK_DATE"), mcolParam("TO_CHECK_DATE"))
   Call LoadAllDocumentCancel(Nothing, m_DocumentCancel)
   
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

Private Function AdjustCreditGroup(Ug As CBkTrn) As Long
   If mcolParam("PRINT_DATE") <= 0 Then
      Ug.Credit = Abs(DateDiff("D", mcolParam("FROM_CHECK_DATE"), Ug.CHQDAT))
   Else
      Ug.Credit = Abs(DateDiff("D", mcolParam("PRINT_DATE"), Ug.CHQDAT))
   End If

   If Ug.Credit <= 15 Then
      AdjustCreditGroup = 15
   ElseIf Ug.Credit <= 30 Then
      AdjustCreditGroup = 30
   Else
      AdjustCreditGroup = 60
   End If
   
   Ug.Credit = AdjustCreditGroup
End Function
Private Function GetKey(Ug As CBkTrn) As String
   'GetKey = Trim(Ug.SUPCOD & "-" & Month(Ug.ChqDat) & "-" & Year(Ug.ChqDat))  '& "-" & Ug.Credit
   GetKey = Trim(Ug.CHQDAT)
End Function
Private Function GetKey2(Ug As CBkTrn) As String
   GetKey2 = Trim(Format(Month(Ug.CHQDAT), "00") & "/" & Format(Year(Ug.CHQDAT) + 543, "0000"))
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
Dim Ug As CBkTrn
Dim St As CIsTab
Dim TempDocDate As Date
Dim CancelFlag As Boolean
Dim Dc As CDocumentCancel
   
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
      
      Set Ug = New CBkTrn
      Ug.SUPCOD = mcolParam("SUPPLIER_CODE")
      Ug.FROM_CHQDAT = mcolParam("FROM_CHECK_DATE")
      Ug.TO_CHQDAT = mcolParam("TO_CHECK_DATE")
      Ug.TO_PAY_DATE = mcolParam("TO_PAY_DATE")
      If mcolParam("SUPPLIER_TYPE") > 0 Then
         Set St = GetItemFromListIndex(m_SupplierTypes, mcolParam("SUPPLIER_TYPE"))
         Ug.SUPTYP = St.TYPCOD
      Else
         Ug.SUPTYP = ""
      End If
      Ug.OrderBy = 1
      Ug.OrderType = mcolParam("ORDER_TYPE")
      Call Ug.QueryData(3, Rs, iCount)
      i = 0
      
      If Not Rs.EOF Then
         Call Ug.PopulateFromRS(3, Rs)
         Call AdjustCreditGroup(Ug)
         PrevKey1 = GetKey(Ug)
         PrevKey2 = GetKey2(Ug)
      End If
      
      While Not Rs.EOF
         Call Ug.PopulateFromRS(3, Rs)
         Call AdjustCreditGroup(Ug)
         
         If Ug.VOUCHER >= "PV0000000" Then
            'debug.print
         End If
         
         If (Ug.VOUCHER >= mcolParam("FROM_DOCUMENT_NO") And Ug.VOUCHER <= mcolParam("TO_DOCUMENT_NO")) Or (Ug.VOUCHER >= mcolParam("FROM_DOCUMENT_NO1") And Ug.VOUCHER <= mcolParam("TO_DOCUMENT_NO1")) Then
            i = i + 1
            
            
            If PrevKey2 <> GetKey2(Ug) And i > 1 Then
               If mcolParam("SHOW_SUMMARY") = "N" Then
                  Call GenerateFooter("รวม", "", "", Total1, m_Details1, 0)
                  Call m_Details1.GetString(2, TempStr1, TempStr2)
                  strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
                  Vsp.FontBold = True
                  Call Vsp.AddTable(strFormat, "", TempStr2)
                  Vsp.FontBold = False
               Else
                  Call GenerateFooter("", "", PrevKey1, Total1, m_Details1, 0)
                  Call m_Details1.GetString(2, TempStr1, TempStr2)
                  strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
                  'Vsp.FontBold = True
                  Call Vsp.AddTable(strFormat, "", TempStr2)
                  Vsp.FontBold = False
               End If
               
               Call GenerateFooter("รวม", "", PrevKey2, Total2, m_Details1, 0)
               Call m_Details1.GetString(2, TempStr1, TempStr2)
               strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
               Vsp.FontBold = True
               Call Vsp.AddTable(strFormat, "", TempStr2)
               Vsp.FontBold = False
               
               PrevKey1 = GetKey(Ug)
               PrevKey2 = GetKey2(Ug)
               For j = 1 To UBound(Total1)
                  Total1(j) = 0
                  Total2(j) = 0
               Next j
            
            ElseIf PrevKey1 <> GetKey(Ug) And i > 1 Then
               If mcolParam("SHOW_SUMMARY") = "N" Then
                  Call GenerateFooter("รวม", "", "", Total1, m_Details1, 0)
                  Call m_Details1.GetString(2, TempStr1, TempStr2)
                  strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
                  Vsp.FontBold = True
                  Call Vsp.AddTable(strFormat, "", TempStr2)
                  Vsp.FontBold = False
               Else
                  Call GenerateFooter("", "", PrevKey1, Total1, m_Details1, 0)
                  Call m_Details1.GetString(2, TempStr1, TempStr2)
                  strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
                  'Vsp.FontBold = True
                  Call Vsp.AddTable(strFormat, "", TempStr2)
                  Vsp.FontBold = False
               End If
               
               PrevKey1 = GetKey(Ug)
               For j = 1 To UBound(Total1)
                  Total1(j) = 0
               Next j
            End If
            
            Call m_Details1.ClearField
   
            Set HeadCf = m_Headers1.Fields(1)
            TempStr = i
            Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
            Call m_Details1.AddField(BodyCf)
            
            
            CancelFlag = False
            Set Dc = GetObject("CDocumentCancel", m_DocumentCancel, Trim(Ug.VOUCHER), False)
            If Not (Dc Is Nothing) Then
               CancelFlag = True
            End If
            Set HeadCf = m_Headers1.Fields(2)
            If CancelFlag Then
               TempStr = "***" & Ug.CHQNUM
            Else
               TempStr = Ug.CHQNUM
            End If
            Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
            Call m_Details1.AddField(BodyCf)
   
            Set HeadCf = m_Headers1.Fields(3)
            TempStr = DateToStringExtEx2(Ug.CHQDAT)
            Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
            Call m_Details1.AddField(BodyCf)
   
            Set HeadCf = m_Headers1.Fields(4)
            If Len(Ug.SUPNAM) > 0 Then
               TempStr = Ug.SUPNAM
            Else
               TempStr = Ug.REMARK
            End If
            Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
            Call m_Details1.AddField(BodyCf)
   
            Set HeadCf = m_Headers1.Fields(5)
            TempStr = Ug.TYPDES
            Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
            Call m_Details1.AddField(BodyCf)
            
            Set HeadCf = m_Headers1.Fields(6)
            If CancelFlag Then
               TempStr = "***" & Ug.VOUCHER
            Else
               TempStr = Ug.VOUCHER
            End If
            Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
            Call m_Details1.AddField(BodyCf)
   
            Set HeadCf = m_Headers1.Fields(7)
            TempStr = FormatNumber(Ug.AMOUNT)
            Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
            Call m_Details1.AddField(BodyCf)
            
            If Not (CancelFlag) Then
               Total1(7) = Total1(7) + Ug.AMOUNT
               Total2(7) = Total2(7) + Ug.AMOUNT
               Total3(7) = Total3(7) + Ug.AMOUNT
            End If
            
            Set HeadCf = m_Headers1.Fields(8)
            TempStr = CalculateCreditCompare(Ug.CHQNUM, Ug.CHQDAT, TempDocDate)
            Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
            Call m_Details1.AddField(BodyCf)
            
            Set HeadCf = m_Headers1.Fields(9)
            TempStr = DateToStringExtEx2(TempDocDate)
            Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
            Call m_Details1.AddField(BodyCf)
            
            If mcolParam("SHOW_SUMMARY") = "N" Then
               Call m_Details1.GetString(1, TempStr1, TempStr2)
               strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
               Call Vsp.AddTable(strFormat, "", TempStr2)
            End If
         Else
            'debug.print
            
         End If
         Rs.MoveNext
      Wend
      
      If mcolParam("SHOW_SUMMARY") = "N" Then
         Call GenerateFooter("รวม", "", "", Total1, m_Details1, 0)
         Call m_Details1.GetString(2, TempStr1, TempStr2)
         strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
         Vsp.FontBold = True
         Call Vsp.AddTable(strFormat, "", TempStr2)
         Vsp.FontBold = False
      Else
         Call GenerateFooter("", "", PrevKey1, Total1, m_Details1, 0)
         Call m_Details1.GetString(2, TempStr1, TempStr2)
         strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
         'Vsp.FontBold = True
         Call Vsp.AddTable(strFormat, "", TempStr2)
         Vsp.FontBold = False
      End If
      Call GenerateFooter("รวม", "", PrevKey2, Total2, m_Details1, 0)
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
   
   If Rs.State = adStateOpen Then
      Rs.Close
   End If
   Set Rs = Nothing
   Set Ug = Nothing
   
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
      ElseIf j = 7 Then
         TempStr = FormatNumber(Tot(j))
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
   Call Vsp.AddTable(strFormat, "", "|เวลา|: " & Right$(DateToStringExtEx(mdteDate), 8))
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
Private Function CalculateCreditCompare(CHQNUM As String, CHQDAT As Date, DOCDAT As Date) As Double
Dim TempBkTrn As CBkTrn
   Set TempBkTrn = GetObject("CBkTrn", m_ChqnumDocDate, CHQNUM, False)
   If Not (TempBkTrn Is Nothing) And TempBkTrn.DOCDAT > 0 Then
      CalculateCreditCompare = DateDiff("D", TempBkTrn.DOCDAT, CHQDAT)
      DOCDAT = TempBkTrn.DOCDAT
   Else
      CalculateCreditCompare = 0
   End If
End Function



