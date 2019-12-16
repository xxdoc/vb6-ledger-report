VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CReportAP016_4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'User Group Report
Implements CReportInterface

Private Const MODULE_NAME = "CReportAP016_4"
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
Private m_Headers2 As CFieldList
Private m_Details1 As CFieldList
Private m_Details2 As CFieldList
Private m_Left As Double
Private m_YearSeqs As Collection

Private m_ExpenseTypes As Collection
Private m_PartGroup As Collection
Private collSupAddr As Collection
Private m_SupplierTypes As Collection

Private Sub Class_Initialize()
   Set mcolParam = New Collection
   Set m_TempCol = New Collection

   Set m_ExpenseTypes = New Collection
   Set m_PartGroup = New Collection

   Set m_Headers0 = New CFieldList
   Set m_Headers1 = New CFieldList
   Set m_Headers2 = New CFieldList
   Set m_Details1 = New CFieldList
   Set m_Details2 = New CFieldList
   Set m_YearSeqs = New Collection
   Set collSupAddr = New Collection
   Set m_SupplierTypes = New Collection
   
End Sub
Private Sub Class_Terminate()
   Call ClearParam
   Set mcolParam = Nothing
   Set Vsp = Nothing
   Set m_TempCol = Nothing

   Set m_ExpenseTypes = Nothing
   Set m_PartGroup = Nothing

   Set m_Headers0 = Nothing
   Set m_Headers1 = Nothing
   Set m_Headers2 = Nothing
   Set m_Details1 = Nothing
   Set m_YearSeqs = Nothing
   Set m_Details2 = Nothing
   Set collSupAddr = Nothing
   Set m_SupplierTypes = Nothing
End Sub
Private Function GetItemFromListIndex(TempCol As Collection, Ind As Long) As Object
   Set GetItemFromListIndex = TempCol(Ind)
End Function
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
'Dim Sgi As CSGroupItem
'Dim TempSgi As CSGroupItem
Dim PrevID As Long
Dim TempStr1 As String
Dim TempStr1_1 As String
Dim Cf As CReportField
Dim ExportWidth As Long
Dim HeadCf As CReportField
Dim j As Long
Dim TempStr As String

   tbdT = Vsp.TableBorder
   blnBold = Vsp.FontBold
   iSize = Vsp.FontSize

   'Vsp.FontSize = HEADER_SIZE
   Vsp.FontBold = True
   Vsp.TableBorder = tbAll

   Call m_Headers0.ClearField
   Call m_Headers1.ClearField
   Call m_Headers2.ClearField

   Set Cf = New CReportField

   Call Cf.SetFieldValue(2, "^", "วันที่เอกสาร", "^")
   Call m_Headers1.AddField(Cf)
   Call Cf.SetFieldValue(2, "^", "RO", "<")
   Call m_Headers1.AddField(Cf)
   Call Cf.SetFieldValue(8, "^", "RM", "<")
   Call m_Headers1.AddField(Cf)
   Call Cf.SetFieldValue(2, "^", "@", ">")
   Call m_Headers1.AddField(Cf)
   Call Cf.SetFieldValue(2, "^", "จำนวน", ">")
   Call m_Headers1.AddField(Cf)
   Call Cf.SetFieldValue(2, "^", "มูลค่า", ">")
   Call m_Headers1.AddField(Cf)
   Set Cf = Nothing
   
   Call m_Headers1.GetString(1, TempStr1, TempStr1_1)
   strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
   strPrint = TempStr1_1
   Call Vsp.AddTable(strFormat, "", strPrint)
   
   Call m_Headers0.GetString(1, TempStr1, TempStr1_1)
   strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
   strPrint = TempStr1_1
   Call Vsp.AddTable(strFormat, "", strPrint)
   
   Vsp.TableBorder = tbdT
   Vsp.FontBold = blnBold
   Vsp.FontSize = iSize
   Vsp.TableBorder = tbNone
'   Vsp.TableBorder = tbAll
   
End Sub

Private Function initDoc() As Boolean
Dim strDate As String
Dim TempStr1 As String
Dim TempStr2 As String
Dim TempStr3 As String
Dim FromDate As String
Dim ToDate As String

   mstrErrMsg = ""
   mblnHeader = True
   mblnNewPage = True
   mblnEndPage = True
   Vsp.PaperSize = pprA4 '
   Vsp.ORIENTATION = orPortrait
   Vsp.MarginBottom = 600
   Vsp.MarginFooter = 600
   Vsp.MarginHeader = 600
   Vsp.MarginLeft = 600
   Vsp.MarginRight = 600
   Vsp.MarginTop = 600
   Vsp.FontName = "AngsanaUPC"
   Vsp.FontSize = DETAIL_SIZE
   
   
'   Call LoadSupplier(Nothing, collSupAddr, , , mcolParam("AMPHUR"), mcolParam("PROVINCE"))
   Call LoadSupplier(Nothing, collSupAddr)
   Call LoadSupplierType(Nothing, m_SupplierTypes)
   
   FromDate = "จาก วันที่ " & EmptyToString(DateToStringExtEx2(mcolParam("FROM_DATE")), "N/A")
   ToDate = " ถึง วันที่ " & EmptyToString(DateToStringExtEx2(mcolParam("TO_DATE")), "N/A")

   MODULE_DESC = glbCompanyName & vbCrLf & mcolParam("REPORT_NAME") & _
                                       vbCrLf & FromDate & "   " & ToDate
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
Dim O As Object
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
Dim Total1(100) As Double
Dim Total2(100) As Double
Dim iCount As Long
Dim TempStr As String
Dim TempID As Long

Dim D As CStcrd
Dim PrevKey1 As String
Dim currentKey1 As String

Dim TempAddr As CAPMas
Dim TempAddrSup As CAPMas
Dim TempCount As Long
Dim DocType As String
Dim St As CIsTab

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
   Next j
   
   Set D = New CStcrd

   D.FROM_DOC_DATE = mcolParam("FROM_DATE")
   D.TO_DOC_DATE = mcolParam("TO_DATE")
   D.SUPCOD = mcolParam("SUPPLIER_CODE")
   D.STKCOD = mcolParam("PART_NO")
   D.STKDES = PatchWildCard2(mcolParam("PART_DESC"))
'   D.RO_TYPE = PatchWildCard2(mcolParam("RO_TYPE"))
   If mcolParam("SUPPLIER_TYPE") > 0 Then
      Set St = GetItemFromListIndex(m_SupplierTypes, mcolParam("SUPPLIER_TYPE"))
      D.SUPTYP = St.TYPCOD
   Else
      D.SUPTYP = ""
   End If
   D.TEMP_ADDR1 = PatchWildCard2(mcolParam("AMPHUR"))
   D.TEMP_ADDR2 = PatchWildCard2(mcolParam("PROVINCE"))

   
   Call D.QueryData(1, Rs, iCount)
   
   Call StartExportFile(Vsp)
   Vsp.StartDoc
   
   i = 0
   K = 0
   Set BodyCf = New CReportField
   
   PrevKey1 = ""
   While Not Rs.EOF
      Call m_Details1.ClearField
      Call D.PopulateFromRS(1, Rs)
'      If D.SUPCOD = "ว-109" Then
'         Debug.Print D.SUPCOD
'      End If
      
      Set TempAddr = GetObject("CAPMas", collSupAddr, Trim(D.SUPCOD), False) 'Trim(Str(d.SUPCOD))
      If Not (TempAddr Is Nothing) Then
         currentKey1 = D.SUPCOD
         If PrevKey1 <> currentKey1 And i > 0 Then
             Vsp.FontBold = True
            Call GenerateFooter("รวม", "", PrevKey1, Total1, m_Details1, 0)
             
            Call m_Details1.GetString(2, TempStr1, TempStr2)
            strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
            Call Vsp.AddTable(strFormat, "", TempStr2)
            Vsp.FontBold = False
   
            For j = 1 To UBound(Total1)
                Total1(j) = 0
            Next j
            Vsp = D.SUPNAM & "(" & D.SUPCOD & ")"
            Vsp = D.ADDR01 & " " & D.ADDR02 & " " & D.TELNUM
         ElseIf i = 0 Then
             Vsp = D.SUPNAM & "(" & D.SUPCOD & ")"
             Vsp = D.ADDR01 & " " & D.ADDR02 & " " & D.TELNUM
         End If
         
         i = i + 1
         PrevKey1 = D.SUPCOD
         Call m_Details1.ClearField
         
         Set HeadCf = m_Headers1.Fields(1) 'วันที่เอกสาร
         TempStr = DateToStringExtEx2(D.DOCDAT)
         Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.Align, TempStr, HeadCf.BodyAlign)
         Call m_Details1.AddField(BodyCf)
         
         Set HeadCf = m_Headers1.Fields(2) 'RO
         TempStr = D.DOCNUM
         Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.Align, TempStr, HeadCf.BodyAlign)
         Call m_Details1.AddField(BodyCf)
         
         Set HeadCf = m_Headers1.Fields(3) 'RM
         TempStr = D.STKDES
         Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.Align, TempStr, HeadCf.BodyAlign)
         Call m_Details1.AddField(BodyCf)
                  
         Set HeadCf = m_Headers1.Fields(4) '@
         TempStr = FormatNumberToNull(D.TRNQTY, 2)
         Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.Align, TempStr, HeadCf.BodyAlign)
         Call m_Details1.AddField(BodyCf)
         
         Set HeadCf = m_Headers1.Fields(5)
         TempStr = FormatNumberToNull(D.UNITPR, 2) 'TX_AMOUNT
         Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.Align, TempStr, HeadCf.BodyAlign)
         Call m_Details1.AddField(BodyCf)
         Total1(5) = Total1(5) + D.UNITPR
         Total2(5) = Total2(5) + D.UNITPR
            
         Set HeadCf = m_Headers1.Fields(6) 'มุลค่า
         TempStr = FormatNumber(D.NETVAL - D.DISAMT)
         Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.Align, TempStr, HeadCf.BodyAlign)
         Call m_Details1.AddField(BodyCf)
         Total1(6) = Total1(6) + (D.NETVAL - D.DISAMT)
         Total2(6) = Total2(6) + (D.NETVAL - D.DISAMT)
         
         Call m_Details1.GetString(2, TempStr1, TempStr2)
         strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
         Call Vsp.AddTable(strFormat, "", TempStr2)
      End If
      Rs.MoveNext
   Wend
   
   Call GenerateFooter("รวม", "", PrevKey1, Total1, m_Details1, 0)
    Call m_Details1.GetString(2, TempStr1, TempStr2)
    strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
    Call Vsp.AddTable(strFormat, "", TempStr2)
    Vsp.FontBold = False
            
   Call GenerateFooter("สุทธิ", "", "", Total2, m_Details1, 0)
    Call m_Details1.GetString(2, TempStr1, TempStr2)
    strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
    Call Vsp.AddTable(strFormat, "", TempStr2)
    Vsp.FontBold = False
            
            
   Set BodyCf = Nothing
   Vsp.EndDoc
   Call CloseExportFile(Vsp)
   
   'Set RoColl = Nothing
   
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

Private Sub GenerateFooter(Txt1 As String, Txt2 As String, Txt3 As String, Tot() As Double, Details As CFieldList, BalAmount As Double)
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
      ElseIf j = 3 Then
         TempStr = Txt3
      ElseIf j = 5 Then
         TempStr = FormatNumber(Tot(j))
      ElseIf j = 6 Then
         TempStr = FormatNumber(Tot(j))
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
   Call Vsp.AddTable(strFormat, "", "|ผู้ใช้|: " & glbUser.USER_NAME)
   Vsp.CurrentY = dY(0)

   Vsp.FontSize = TITLE_SIZE
   Vsp.FontBold = True
   Vsp.TextAlign = taCenterBottom
   Vsp.Paragraph = MODULE_DESC
   Vsp.Paragraph = ""
   Vsp.FontSize = iSize
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
   Vsp.TableBorder = tbAll
End Sub

Private Sub VSP_StartDoc()
'This event occur when VSPrinter.StartDoc is used and used to initialize some information before generating printed document
End Sub
Private Sub GenerateHeader(Txt1 As String, Txt2 As String, Txt3 As String, Details As CFieldList)
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
         TempStr = Txt1
      ElseIf j = 2 Then
         TempStr = Txt2
      ElseIf j = 3 Then
         TempStr = Txt3
      Else
         TempStr = ""
      End If
      
      Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr, HeadCf.Align)
      Call Details.AddField(BodyCf)
   Next HeadCf
      
   Set BodyCf = Nothing
End Sub

