VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CReportCom14"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'User Group Report
Implements CReportInterface

Private Const MODULE_NAME = "CReportCOM14"
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
Private TempRs As ADODB.Recordset

Private Const TITLE_SIZE = 14
Private Const HEADER_SIZE = 12
Private Const DETAIL_SIZE = 12
Private Const TITLE_FORMAT = "<32|<3|<5;"

Private m_HeaderFormat1 As String
Private m_HeaderFormat2 As String
Private m_HeaderText1 As String
Private m_HeaderText2 As String
Private m_Record01 As Collection
Private m_Record02 As Collection
Private m_Record03 As Collection
Private m_RecordMix As Collection
Private temp_RecordMix As CComRecord

Private m_Headers0 As CFieldList
Private m_Headers1 As CFieldList
Private m_Details1 As CFieldList

'Private m_TempCol As Collection
'Private FirstDate As Date
'Private LastDate As Date
'Private eachSLM As Collection
 'Private TempEachSLM As CCondiIncenSum
 Dim FromSaleID As String
 Dim ToSaleID As String

Private Sub Class_Initialize()
   Set mcolParam = New Collection
   
   Set m_Headers0 = New CFieldList
   Set m_Headers1 = New CFieldList
   Set m_Details1 = New CFieldList
   
   Set m_Record01 = New Collection
   Set m_Record02 = New Collection
   Set m_Record03 = New Collection
   Set m_RecordMix = New Collection
   Set temp_RecordMix = New CComRecord
   'Set m_TempCol = New Collection
   'Set TempEachSLM = New CCondiIncenSum
End Sub

Private Sub Class_Terminate()
   Call ClearParam
   Set mcolParam = Nothing
   Set Vsp = Nothing
   
   Set m_Headers0 = Nothing
   Set m_Headers1 = Nothing
   Set m_Details1 = Nothing
   
   Set m_Record01 = Nothing
   Set m_Record02 = Nothing
   Set m_Record03 = Nothing
   Set m_RecordMix = Nothing
   Set temp_RecordMix = Nothing
   'Set m_TempCol = Nothing
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
      
   Call Cf.SetFieldValue(1, "^", "No.", "^")
   Call m_Headers0.AddField(Cf)
   Call Cf.SetFieldValue(12, "^", "พนักงานขาย", "^")    '27
   Call m_Headers0.AddField(Cf)
    If mcolParam("COM1_MODE") = "Y" Then
      Call Cf.SetFieldValue(3, "^", "ค่าคอมขาย", ">")
      Call m_Headers0.AddField(Cf)
   End If
   If mcolParam("COM2_MODE") = "Y" Then
      Call Cf.SetFieldValue(3, "^", "ค่าคอมเก็บเงิน", ">")
      Call m_Headers0.AddField(Cf)
   End If
   If mcolParam("INCEN_MODE") = "Y" Then
      Call Cf.SetFieldValue(3, "^", "Incentive", ">")
      Call m_Headers0.AddField(Cf)
   End If
   Call Cf.SetFieldValue(3, "^", "รวม", ">")
   Call m_Headers0.AddField(Cf)
   
'   Call m_Headers0.GetString(1, TempStr1, TempStr1_1)
'   strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
'   strPrint = TempStr1_1
'   Call Vsp.AddTable(strFormat, "", strPrint)
   

   Call Cf.SetFieldValue(1, "^", "No.", "^")
   Call m_Headers1.AddField(Cf)
   Call Cf.SetFieldValue(1, "^", "รหัส", "^")
   Call m_Headers1.AddField(Cf)
   Call Cf.SetFieldValue(3, "^", "ชื่อพนักงานขาย", "<")
   Call m_Headers1.AddField(Cf)
   Call Cf.SetFieldValue(2, "^", "จากวันที่ส่ง", "^")
   Call m_Headers1.AddField(Cf)
   Call Cf.SetFieldValue(2, "^", "วันที่จ่าย", "^")
   Call m_Headers1.AddField(Cf)
   Call Cf.SetFieldValue(2, "^", "เขตการขาย", "^")
   Call m_Headers1.AddField(Cf)
   Call Cf.SetFieldValue(2, "^", "ประเภท", "^")
   Call m_Headers1.AddField(Cf)
   
   If mcolParam("COM1_MODE") = "Y" Then
      Call Cf.SetFieldValue(3, "^", "ค่าคอมขาย", ">")
      Call m_Headers1.AddField(Cf)
   End If
   If mcolParam("COM2_MODE") = "Y" Then
      Call Cf.SetFieldValue(3, "^", "ค่าคอมเก็บเงิน", ">")
      Call m_Headers1.AddField(Cf)
   End If
   If mcolParam("INCEN_MODE") = "Y" Then
      Call Cf.SetFieldValue(3, "^", "Incentive", ">")
      Call m_Headers1.AddField(Cf)
   End If
   Call Cf.SetFieldValue(3, "^", "รวม", ">")
   Call m_Headers1.AddField(Cf)

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
Dim YEAR_ID As Long

   mstrErrMsg = ""
   mblnHeader = True
   mblnNewPage = True
   mblnEndPage = True
   Vsp.PaperSize = pprA4 '
   Vsp.ORIENTATION = orLandscape
      'Vsp.ORIENTATION = orPortrait
   Vsp.MarginBottom = 300
   Vsp.MarginFooter = 300
   Vsp.MarginHeader = 300
   Vsp.MarginLeft = 300
   Vsp.MarginRight = 500
   Vsp.MarginTop = 300
   Vsp.FontName = "AngsanaUPC"
   Vsp.FontSize = DETAIL_SIZE
      
   MODULE_DESC = glbCompanyName & vbCrLf & mcolParam("REPORT_NAME")
   If mcolParam("COM1_MODE") = "Y" Then
      MODULE_DESC = MODULE_DESC & "    คอมขาย"
   End If
   If mcolParam("COM2_MODE") = "Y" Then
      MODULE_DESC = MODULE_DESC & "    คอมเก็บเงิน"
   End If
   If mcolParam("INCEN_MODE") = "Y" Then
      MODULE_DESC = MODULE_DESC & "    Incentive"
   End If
   MODULE_DESC = MODULE_DESC & vbCrLf & "วันที่จ่าย    " & EmptyToString(DateToStringExtEx2(mcolParam("FROM_DOC_DATE")), "N/A") & "      ถึง        " & EmptyToString(DateToStringExtEx2(mcolParam("TO_DOC_DATE")), "N/A")
   
  Dim bak_ID As Long
  Dim FromSaleName As String
  Dim ToSaleName As String
  
   bak_ID = InStr(1, mcolParam("FROM_SALE_NAME"), " ")
   If bak_ID > 0 Then
     FromSaleID = Left(mcolParam("FROM_SALE_NAME"), bak_ID - 1)
     FromSaleName = Mid(mcolParam("FROM_SALE_NAME"), bak_ID + 3)
   End If
   bak_ID = InStr(1, mcolParam("TO_SALE_NAME"), " ")
   If bak_ID > 0 Then
     ToSaleID = Left(mcolParam("TO_SALE_NAME"), bak_ID - 1)
     ToSaleName = Mid(mcolParam("TO_SALE_NAME"), bak_ID + 3)
   End If
   
   If Len(mcolParam("FROM_SALE_NAME")) > 0 Then
      MODULE_DESC = MODULE_DESC & "       พนักงานขาย  " & EmptyToString(FromSaleName, "N/A")  '
      MODULE_DESC = MODULE_DESC & "   ถึง    " & EmptyToString(ToSaleName, "N/A")
   End If
   
   Call LoadComRecordJoin(m_Record01, "", mcolParam("FROM_DOC_DATE"), mcolParam("TO_DOC_DATE")) ' all record
   Call LoadComRecordMix(m_RecordMix, m_Record01)

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
Dim Total3(100) As Double
Dim iCount As Long
Dim TempStr As String
Dim PrevKey1 As String
Dim PrevKey2 As String
Dim Stcrd As CStcrd

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
   
   Call StartExportFile(Vsp)
   Vsp.StartDoc
   i = 1  'นับ
   Set BodyCf = New CReportField
   For Each temp_RecordMix In m_RecordMix
   
               If temp_RecordMix.SLMCOD <> PrevKey1 And PrevKey1 <> "" Then
                     Call GenerateFooter("รวม", PrevKey2, Total1, m_Details1)       ' เช็ค haveValue ใน Call GenerateFooter **** เซลล์ต้องมีค่าถึงจะแสดงบรรทัดรวม
                     Vsp.FontBold = True
                     Call m_Details1.GetString(2, TempStr1, TempStr2)
                     strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
                     Call Vsp.AddTable(strFormat, "", TempStr2)
                     Vsp.FontBold = False
                     
                     For j = 1 To UBound(Total1)
                           Total1(j) = 0
                     Next j
                     
               End If
   
   
               Call m_Details1.ClearField
               j = 0

                j = j + 1
               Set HeadCf = m_Headers1.Fields(j)
               TempStr = i
               Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
               Call m_Details1.AddField(BodyCf)

                j = j + 1
               Set HeadCf = m_Headers1.Fields(j)
               TempStr = temp_RecordMix.SLMCOD
               Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
               Call m_Details1.AddField(BodyCf)

               j = j + 1
               Set HeadCf = m_Headers1.Fields(j)
               TempStr = temp_RecordMix.SLMNAME
               Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
               Call m_Details1.AddField(BodyCf)

                j = j + 1
               Set HeadCf = m_Headers1.Fields(j)
               TempStr = DateToStringExtEx2(temp_RecordMix.FROMDAT)
               Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
               Call m_Details1.AddField(BodyCf)

                j = j + 1
               Set HeadCf = m_Headers1.Fields(j)
               TempStr = DateToStringExtEx2(temp_RecordMix.TODAT)
               Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
               Call m_Details1.AddField(BodyCf)

               j = j + 1
               Set HeadCf = m_Headers1.Fields(j)
               TempStr = temp_RecordMix.AREA_NAME
               Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
               Call m_Details1.AddField(BodyCf)
               
               j = j + 1
               Set HeadCf = m_Headers1.Fields(j)
               TempStr = temp_RecordMix.GOODS_GROUP_NAME
               Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
               Call m_Details1.AddField(BodyCf)
               
               If mcolParam("COM1_MODE") = "Y" Then
               j = j + 1                                    ' คอมขาย
               Set HeadCf = m_Headers1.Fields(j)
               TempStr = FormatNumber(temp_RecordMix.COM_VALUE01)
               Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
               Call m_Details1.AddField(BodyCf)
               Total1(j) = Total1(j) + temp_RecordMix.COM_VALUE01    ' รายเซลล์
               Total2(j) = Total2(j) + temp_RecordMix.COM_VALUE01    'ทั้งหมด
               Total3(i) = Total3(i) + temp_RecordMix.COM_VALUE01    'แนวนอน
               End If
               
               If mcolParam("COM2_MODE") = "Y" Then
               j = j + 1                                    ' คอมเก็บเงิน
               Set HeadCf = m_Headers1.Fields(j)
               TempStr = FormatNumber(temp_RecordMix.COM_VALUE02)
               Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
               Call m_Details1.AddField(BodyCf)
               Total1(j) = Total1(j) + temp_RecordMix.COM_VALUE02    ' รายเซลล์
               Total2(j) = Total2(j) + temp_RecordMix.COM_VALUE02    'ทั้งหมด
               Total3(i) = Total3(i) + temp_RecordMix.COM_VALUE02    'แนวนอน
               End If
              
               If mcolParam("INCEN_MODE") = "Y" Then
                j = j + 1                                    ' incentive
               Set HeadCf = m_Headers1.Fields(j)
               TempStr = FormatNumber(temp_RecordMix.COM_VALUE03)
               Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
               Call m_Details1.AddField(BodyCf)
                Total1(j) = Total1(j) + temp_RecordMix.COM_VALUE03    ' รายเซลล์
               Total2(j) = Total2(j) + temp_RecordMix.COM_VALUE03    'ทั้งหมด
               Total3(i) = Total3(i) + temp_RecordMix.COM_VALUE03    'แนวนอน
               End If
               
                j = j + 1                                    ' รวม
               Set HeadCf = m_Headers1.Fields(j)
               TempStr = FormatNumber(Total3(i))
               Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
               Call m_Details1.AddField(BodyCf)
               Total1(j) = Total1(j) + Total3(i)    ' รายเซลล์
               Total2(j) = Total2(j) + Total3(i)   'ทั้งหมด
               Total3(i) = Total3(i) + Total3(i)    'แนวนอน
                  
                  Call m_Details1.GetString(1, TempStr1, TempStr2)
                  strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
                  Call Vsp.AddTable(strFormat, "", TempStr2)

               i = i + 1
              PrevKey1 = temp_RecordMix.SLMCOD
              PrevKey2 = temp_RecordMix.SLMNAME

   Next temp_RecordMix
   
                     Call GenerateFooter("รวม", PrevKey2, Total1, m_Details1)       ' เซลล์สุดท้าย
                     Vsp.FontBold = True
                     Call m_Details1.GetString(2, TempStr1, TempStr2)
                     strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
                     Call Vsp.AddTable(strFormat, "", TempStr2)
                     Vsp.FontBold = False

                     Call GenerateFooter("", "สุทธิ", Total2, m_Details1)       ' เซลล์สุดท้าย
                     Vsp.FontBold = True
                     Call m_Details1.GetString(2, TempStr1, TempStr2)
                     strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
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
      Set TempRs = Nothing
End Function

Private Sub GenerateHeader(Txt1 As String, Txt2 As String, Details1 As CFieldList)
Dim HeadCf As CReportField
Dim BodyCf As CReportField
Dim j As Long
Dim TempStr As String

   Call Details1.ClearField
   Set BodyCf = New CReportField
   
                Set HeadCf = m_Headers0.Fields(1)
                TempStr = Txt1
                Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
                Call m_Details1.AddField(BodyCf)
    
                Set HeadCf = m_Headers0.Fields(2)
                TempStr = Txt2
                Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
                Call m_Details1.AddField(BodyCf)
    
    
   Set BodyCf = Nothing
End Sub

Private Sub GenerateFooter(Txt1 As String, Txt2 As String, Tot() As Double, Details As CFieldList)
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
         TempStr = FormatNumber(Tot(8))
      ElseIf j = 4 Then
         TempStr = FormatNumber(Tot(9))
      ElseIf j = 5 Then
         TempStr = FormatNumber(Tot(10))
      ElseIf j = 6 Then
         TempStr = FormatNumber(Tot(11))
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

Public Sub LoadComRecordMix(Optional Cl_Mix2 As Collection = Nothing, Optional Cl_Mix1 As Collection = Nothing)
On Error GoTo ErrorHandler
Dim ArS As COESLM
Dim ItemCount As Long
Dim IsOK As Boolean
Dim D_01 As CComRecord
Dim D_02 As CComRecord
Dim D_Mix As CComRecord
Dim TempRs As ADODB.Recordset
Dim i As Long

   Set TempRs = New ADODB.Recordset
   
    If Not (Cl_Mix2 Is Nothing) Then
      Set Cl_Mix2 = Nothing
      Set Cl_Mix2 = New Collection
   End If
   
   Set ArS = New COESLM
   ArS.FROM_SALE_CODE = FromSaleID
   ArS.TO_SALE_CODE = ToSaleID
   Call glbDaily.QuerySale(ArS, TempRs, ItemCount, IsOK, glbErrorLog)
   
   While Not TempRs.EOF          ' sale
      Call m_Details1.ClearField
      Call ArS.PopulateFromRS(1, TempRs)
         
      For Each D_01 In Cl_Mix1
         If D_01.SLMCOD = ArS.SLMCOD Then
            Set D_Mix = New CComRecord
            Set D_02 = GetObject("CComRecord", Cl_Mix2, Trim(D_01.SLMCOD & "-" & D_01.FROMDAT & "-" & D_01.TODAT & "-" & D_01.MASTER_AREA_ID & "-" & D_01.GOODS_GROUP_ID), False)
             If D_02 Is Nothing Then
                  Set D_Mix = D_01
                  If D_01.COMTYP = "01" Then
                       D_Mix.COM_VALUE01 = D_01.COM_VALUE
                  ElseIf D_01.COMTYP = "02" Then
                       D_Mix.COM_VALUE02 = D_01.COM_VALUE
                  ElseIf D_01.COMTYP = "03" Then
                       D_Mix.COM_VALUE03 = D_01.COM_VALUE
                  End If
                   Call Cl_Mix2.Add(D_Mix, Trim(D_01.SLMCOD & "-" & D_01.FROMDAT & "-" & D_01.TODAT & "-" & D_01.MASTER_AREA_ID & "-" & D_01.GOODS_GROUP_ID))
                   Set D_Mix = Nothing
              Else
                     If D_01.COMTYP = "01" Then
                          D_02.COM_VALUE01 = D_01.COM_VALUE
                     ElseIf D_01.COMTYP = "02" Then
                          D_02.COM_VALUE02 = D_01.COM_VALUE
                     ElseIf D_01.COMTYP = "03" Then
                          D_02.COM_VALUE03 = D_01.COM_VALUE
                     End If
                     Set D_02 = Nothing
              End If
                
           End If
      Next D_01
                              
         TempRs.MoveNext                                                            ' วนเซลล์
   Wend

   Set TempRs = Nothing
   Set D_Mix = Nothing
   Set D_01 = Nothing
   Set D_02 = Nothing
   Set ArS = Nothing
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub
