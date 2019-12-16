VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CReportGL01"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'User Group Report
Implements CReportInterface

Private Const MODULE_NAME = "CReportAR001"
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
   
   Call m_Headers0.ClearField
   Call m_Headers1.ClearField
   
   Vsp.FontSize = HEADER_SIZE
   Vsp.FontBold = True
   Vsp.TableBorder = tbAll
   
   Set Cf = New CReportField
     
   Call Cf.SetFieldValue(70, "^", "รายละเอียด", "^")
   Call m_Headers0.AddField(Cf)
     
   Call Cf.SetFieldValue(40, "^", "ยอดยกมาต้นงวด", "^")
   Call m_Headers0.AddField(Cf)
   
   Call Cf.SetFieldValue(40, "^", "ยอดเปลี่ยนแปลง", "^")
   Call m_Headers0.AddField(Cf)
   
   Call Cf.SetFieldValue(40, "^", "ยอดคงเหลือสิ้นงวด", "^")
   Call m_Headers0.AddField(Cf)
   
   Call Cf.SetFieldValue(15, "^", "เลขที่บัญชี", "^")
   Call m_Headers1.AddField(Cf)
   
   Call Cf.SetFieldValue(55, "^", "ชื่อบัญชี", "<")
   Call m_Headers1.AddField(Cf)
   
   Call Cf.SetFieldValue(20, "^", "เดบิต", ">")
   Call m_Headers1.AddField(Cf)
   Call Cf.SetFieldValue(20, "^", "เครดิต", ">")
   Call m_Headers1.AddField(Cf)
   Call Cf.SetFieldValue(20, "^", "เดบิต", ">")
   Call m_Headers1.AddField(Cf)
   Call Cf.SetFieldValue(20, "^", "เครดิต", ">")
   Call m_Headers1.AddField(Cf)
   Call Cf.SetFieldValue(20, "^", "เดบิต", ">")
   Call m_Headers1.AddField(Cf)
   Call Cf.SetFieldValue(20, "^", "เครดิต", ">")
   Call m_Headers1.AddField(Cf)
   
   Call m_Headers0.GetString(1, TempStr1, TempStr1_1)
   strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
   strPrint = TempStr1_1
   Call Vsp.AddTable(strFormat, "", strPrint)
   
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
   Vsp.MarginLeft = 300
   Vsp.MarginRight = 300
   Vsp.MarginTop = 300
   Vsp.FontName = "AngsanaUPC"
   Vsp.FontSize = DETAIL_SIZE
   
   MODULE_DESC = glbCompanyName & vbCrLf & mcolParam("REPORT_NAME") & _
                                       vbCrLf & "จากเดือน " & mcolParam("FROM_MONTH") & "ถึงเดือน " & mcolParam("TO_MONTH") & " ปี " & mcolParam("YEAR")
   
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

Dim GlAcc As CGLAcc
Dim TempSumMovement As Double
Dim HaveValue As Boolean

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
      

    Set GlAcc = New CGLAcc
    GlAcc.OrderBy = 2
    GlAcc.OrderType = 1
   Call GlAcc.QueryData(1, Rs, iCount)
      
   i = 0
   j = 0
   While Not Rs.EOF
      i = i + 1
      
      Call m_Details1.ClearField
      
      Call GlAcc.PopulateFromRS(1, Rs)
      
'      If i = 1 Then
'         Call printHeader
'      End If
      If GlAcc.ACCTYP = 0 Then
         HaveValue = False
         
         j = 0
         
'         If GlAcc.ACCNUM = "142-2200" Then
'            'debug.print
'         End If
         j = j + 1
         Set HeadCf = m_Headers1.Fields(j)
         TempStr = GlAcc.ACCNUM
         Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
         Call m_Details1.AddField(BodyCf)
         
         j = j + 1
         Set HeadCf = m_Headers1.Fields(j)
         TempStr = GlAcc.ACCNAM
         Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
         Call m_Details1.AddField(BodyCf)
         
         If (GlAcc.BEGCUR >= 0 And GlAcc.NATURE = 0) Or (GlAcc.BEGCUR < 0 And GlAcc.NATURE = 1) Then
            j = j + 1
            Set HeadCf = m_Headers1.Fields(j)
            TempStr = FormatNumber(Abs(GlAcc.BEGCUR))
            Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
            Call m_Details1.AddField(BodyCf)
            Total1(j) = Total1(j) + Abs(GlAcc.BEGCUR)
            
            j = j + 1
            Set HeadCf = m_Headers1.Fields(j)
            TempStr = FormatNumber(0)
            Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
            Call m_Details1.AddField(BodyCf)
         ElseIf GlAcc.BEGCUR < 0 And GlAcc.NATURE = 0 Or (GlAcc.BEGCUR >= 0 And GlAcc.NATURE = 1) Then
            j = j + 1
            Set HeadCf = m_Headers1.Fields(j)
            TempStr = FormatNumber(0)
            Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
            Call m_Details1.AddField(BodyCf)
            
            j = j + 1
            Set HeadCf = m_Headers1.Fields(j)
            TempStr = FormatNumber(Abs(GlAcc.BEGCUR))
            Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
            Call m_Details1.AddField(BodyCf)
            Total1(j) = Total1(j) + Abs(GlAcc.BEGCUR)
         End If
         If Round(GlAcc.BEGCUR, 2) <> 0 Then
            HaveValue = True
         End If
         
         TempSumMovement = 0
         TempSumMovement = GlAcc.BAL1CY
         TempSumMovement = TempSumMovement + GlAcc.BAL2CY
         TempSumMovement = TempSumMovement + GlAcc.BAL3CY
         TempSumMovement = TempSumMovement + GlAcc.BAL4CY
         TempSumMovement = TempSumMovement + GlAcc.BAL5CY
         TempSumMovement = TempSumMovement + GlAcc.BAL6CY
         TempSumMovement = TempSumMovement + GlAcc.BAL7CY
         TempSumMovement = TempSumMovement + GlAcc.BAL8CY
         TempSumMovement = TempSumMovement + GlAcc.BAL9CY
         TempSumMovement = TempSumMovement + GlAcc.BAL10CY
         TempSumMovement = TempSumMovement + GlAcc.BAL11CY
         TempSumMovement = TempSumMovement + GlAcc.BAL12CY
         If (Round(TempSumMovement, 2) <> 0) Then
            HaveValue = True
         End If
         
         If (TempSumMovement >= 0 And GlAcc.NATURE = 0) Or (TempSumMovement < 0 And GlAcc.NATURE = 1) Then
            j = j + 1
            Set HeadCf = m_Headers1.Fields(j)
            TempStr = FormatNumber(Abs(TempSumMovement))
            Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
            Call m_Details1.AddField(BodyCf)
            Total1(j) = Total1(j) + Abs(TempSumMovement)
            
            j = j + 1
            Set HeadCf = m_Headers1.Fields(j)
            TempStr = FormatNumber(0)
            Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
            Call m_Details1.AddField(BodyCf)
         ElseIf (TempSumMovement < 0 And GlAcc.NATURE = 0) Or (TempSumMovement >= 0 And GlAcc.NATURE = 1) Then
            j = j + 1
            Set HeadCf = m_Headers1.Fields(j)
            TempStr = FormatNumber(0)
            Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
            Call m_Details1.AddField(BodyCf)
            
            j = j + 1
            Set HeadCf = m_Headers1.Fields(j)
            TempStr = FormatNumber(Abs(TempSumMovement))
            Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
            Call m_Details1.AddField(BodyCf)
            Total1(j) = Total1(j) + Abs(TempSumMovement)
         End If
         
         If ((TempSumMovement + GlAcc.BEGCUR) >= 0 And GlAcc.NATURE = 0) Or ((TempSumMovement + GlAcc.BEGCUR) < 0 And GlAcc.NATURE = 1) Then
            j = j + 1
            Set HeadCf = m_Headers1.Fields(j)
            TempStr = FormatNumber(Abs(TempSumMovement + GlAcc.BEGCUR))
            Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
            Call m_Details1.AddField(BodyCf)
            Total1(j) = Total1(j) + Abs(TempSumMovement + GlAcc.BEGCUR)
            
            j = j + 1
            Set HeadCf = m_Headers1.Fields(j)
            TempStr = FormatNumber(0)
            Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
            Call m_Details1.AddField(BodyCf)
         ElseIf ((TempSumMovement + GlAcc.BEGCUR) < 0 And GlAcc.NATURE = 0) Or ((TempSumMovement + GlAcc.BEGCUR) >= 0 And GlAcc.NATURE = 1) Then
            j = j + 1
            Set HeadCf = m_Headers1.Fields(j)
            TempStr = FormatNumber(0)
            Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
            Call m_Details1.AddField(BodyCf)
            
            j = j + 1
            Set HeadCf = m_Headers1.Fields(j)
            TempStr = FormatNumber(Abs(TempSumMovement + GlAcc.BEGCUR))
            Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
            Call m_Details1.AddField(BodyCf)
            Total1(j) = Total1(j) + Abs(TempSumMovement + GlAcc.BEGCUR)
         End If
         
         If HaveValue Then
            Call m_Details1.GetString(1, TempStr1, TempStr2)
            strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
            Call Vsp.AddTable(strFormat, "", TempStr2)
         End If
      End If
      
      Rs.MoveNext
   Wend
   
   Call GenerateFooter("สุทธิ", "", "", Total1, m_Details1, 0)
   Call m_Details1.GetString(1, TempStr1, TempStr2)
   strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
   Vsp.FontBold = True
   Call Vsp.AddTable(strFormat, "", TempStr2)
   Vsp.FontBold = False
   
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
      Else
         TempStr = FormatNumber(Tot(j))
      End If
      
      Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr, HeadCf.BodyAlign)
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
'   Vsp.Paragraph = ""
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
Private Sub GetBalance()

End Sub
