VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CReportAR012"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'User Group Report
Implements CReportInterface

Private Const MODULE_NAME = "CReportAR012"
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

Private m_CustomerTypes As Collection
Private m_ChqnumDocDate As Collection
Private m_DocumentCancel As Collection

Dim TempCCheck As CCheckCancel
Dim TempCheckCancel As CCheckCancel
Dim m_ReceiveCancelitem As Collection
Dim ItemCount As Long
 Dim m_Rs As ADODB.Recordset
 
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
   Set m_CustomerTypes = New Collection
   Set m_ChqnumDocDate = New Collection
   Set m_DocumentCancel = New Collection
   
   Set TempCCheck = New CCheckCancel
   Set TempCheckCancel = New CCheckCancel
   Set m_ReceiveCancelitem = New Collection
   
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
   Set m_CustomerTypes = Nothing
   Set m_ChqnumDocDate = Nothing
   Set m_DocumentCancel = Nothing
   
   Set TempCCheck = Nothing
   Set TempCheckCancel = Nothing
   Set m_ReceiveCancelitem = Nothing
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
   Call Cf.SetFieldValue(9, "^", "", "^")
   Call m_Headers0.AddField(Cf)
   Call Cf.SetFieldValue(20, "^", "", "<")
   Call m_Headers0.AddField(Cf)
Else:
    Call Cf.SetFieldValue(2, "^", "", "^")
   Call m_Headers0.AddField(Cf)
   Call Cf.SetFieldValue(4, "^", "", "<")
   Call m_Headers0.AddField(Cf)
End If
   
   If mcolParam("SHOW_SUMMARY") = "N" Then
      Call Cf.SetFieldValue(1, "^", "NO.", "^")
      Call m_Headers1.AddField(Cf)
   Else
      Call Cf.SetFieldValue(0, "^", "NO.", "^")
      Call m_Headers1.AddField(Cf)
   End If
   If mcolParam("SHOW_SUMMARY") = "N" Then
      Call Cf.SetFieldValue(2, "^", "เลขที่เช็ค", "^")
      Call m_Headers1.AddField(Cf)
   Else
      Call Cf.SetFieldValue(0, "^", "เลขที่เช็ค", "^")
      Call m_Headers1.AddField(Cf)
   End If
   
   If mcolParam("SHOW_SUMMARY") = "N" Then
      Call Cf.SetFieldValue(2, "^", "วัน.หน้าเช็ค", "^")
      Call m_Headers1.AddField(Cf)
      Call Cf.SetFieldValue(2, "^", "ปมก.ผ่าน", ">")
      Call m_Headers1.AddField(Cf)
   Else
      Call Cf.SetFieldValue(0, "^", "เช็ค", "^")
      Call m_Headers1.AddField(Cf)
      If mcolParam("GROUP_BY_DUE_DATE") = "Y" Then
         Call Cf.SetFieldValue(2, "^", "สรุปประมาณการเช็คผ่านตามวันที่ผ่าน", "<")
      ElseIf mcolParam("GROUP_BY_CUSTOMER") = "Y" Then
         Call Cf.SetFieldValue(2, "^", "สรุปประมาณการเช็คผ่านตามลูกค้า", "<")
      ElseIf mcolParam("GROUP_BY_SALE") = "Y" Then
         Call Cf.SetFieldValue(2, "^", "สรุปประมาณการเช็คผ่านตามพนักงานขาย", "<")
     ElseIf mcolParam("GROUP_BY_SALE_AND_CUSTOMER") = "Y" Then
         Call Cf.SetFieldValue(2, "^", "สรุปประมาณการเช็คผ่านตามพนักงานขายและลูกค้า", "<")
      Else
         Call Cf.SetFieldValue(2, "^", "สรุปประมาณการเช็คผ่านตามวันที่ผ่าน", "<")
      End If
      Call m_Headers1.AddField(Cf)
   End If
       
    If mcolParam("SHOW_SUMMARY") = "N" Then
      Call Cf.SetFieldValue(2, "^", "วันที่ผ่าน", "<")
      Call m_Headers1.AddField(Cf)
   Else
      Call Cf.SetFieldValue(0, "^", "วันที่ผ่าน", "<")
      Call m_Headers1.AddField(Cf)
   End If
   
    If mcolParam("SHOW_SUMMARY") = "N" Then
      Call Cf.SetFieldValue(1, "^", "B.", "<")
      Call m_Headers1.AddField(Cf)
   Else
      Call Cf.SetFieldValue(0, "^", "ธนาคาร", "<")
      Call m_Headers1.AddField(Cf)
   End If
   
   If mcolParam("SHOW_SUMMARY") = "N" Then
      Call Cf.SetFieldValue(3, "^", "สาขา", "<")
      Call m_Headers1.AddField(Cf)
   Else
      Call Cf.SetFieldValue(0, "^", "สาขา", "<")
      Call m_Headers1.AddField(Cf)
   End If
   
   If mcolParam("SHOW_SUMMARY") = "N" Then
      Call Cf.SetFieldValue(2, "^", "รหัสลูกค้า", "<")
      Call m_Headers1.AddField(Cf)
   Else
      Call Cf.SetFieldValue(0, "^", "สาขา", "<")
      Call m_Headers1.AddField(Cf)
   End If
   
   If mcolParam("SHOW_SUMMARY") = "N" Then
      Call Cf.SetFieldValue(4, "^", "ชื่อลูกค้า", "<")
      Call m_Headers1.AddField(Cf)
   Else
      Call Cf.SetFieldValue(0, "^", "ชื่อลูกค้า", "<")
      Call m_Headers1.AddField(Cf)
   End If
   
   If mcolParam("SHOW_SUMMARY") = "N" Then
      Call Cf.SetFieldValue(2, "^", "ใบสำคัญรับ", "^")
      Call m_Headers1.AddField(Cf)
   Else
      Call Cf.SetFieldValue(0, "^", "ใบสำคัญรับ", "^")
      Call m_Headers1.AddField(Cf)
   End If
   
   Call Cf.SetFieldValue(2, "^", "จำนวนเงิน", ">")
   Call m_Headers1.AddField(Cf)
   
   If mcolParam("SHOW_SUMMARY") = "N" Then
      Call Cf.SetFieldValue(2, "^", "วันที่ส่งของ", ">")
      Call m_Headers1.AddField(Cf)
      
      Call Cf.SetFieldValue(1, "^", "C", "^")
      Call m_Headers1.AddField(Cf)
   Else
      Call Cf.SetFieldValue(0, "^", "วันที่ส่งของ", ">")
      Call m_Headers1.AddField(Cf)
      
      Call Cf.SetFieldValue(2, "^", "เครดิตหัวเฉลี่ย", "^")
      Call m_Headers1.AddField(Cf)
   End If
   
   If mcolParam("SHOW_SUMMARY") = "N" Then
      Call Cf.SetFieldValue(3, "^", "พนักงานขาย", "<")
      Call m_Headers1.AddField(Cf)
   Else
      Call Cf.SetFieldValue(0, "^", "พนักงานขาย", "<")
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

Dim iCount As Long

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
      
   Call LoadCustomerType(Nothing, m_CustomerTypes)
   
   'โหลด error ตรงนี้ล่ะ เรียกไป modloaddata     ----------->    Set Ari2 = GetARRcpItemEx(m_ReceiveAllAmounts, ARt.DOCNUM)
   Call TempCheckCancel.QueryData(m_Rs, iCount)
   Call LoadCheckCancel(m_ReceiveCancelitem, TempCheckCancel.CHECK_NO, m_Rs)

   FromDate = "จากวันที่ " & EmptyToString(DateToStringExtEx2(mcolParam("FROM_CHECK_DATE")), "N/A")
   ToDate = " ถึง วันที่ " & EmptyToString(DateToStringExtEx2(mcolParam("TO_CHECK_DATE")), "N/A")
   
   MODULE_DESC = glbCompanyName & vbCrLf & mcolParam("REPORT_NAME") & _
                                       vbCrLf & FromDate & ToDate & " จากลูกค้า " & EmptyToString(mcolParam("FROM_CUSTOMER_CODE"), "N/A") & " ถึงลูกค้า " & EmptyToString(mcolParam("TO_CUSTOMER_CODE"), "N/A") & _
                                       " จากพนักงานขาย " & EmptyToString(mcolParam("FROM_SALE_CODE"), "N/A") & " ถึงพนักงานขาย " & EmptyToString(mcolParam("TO_SALE_CODE"), "N/A")
                                       
   If mcolParam("GROUP_BY_CUSTOMER") = "Y" Then
      MODULE_DESC = MODULE_DESC & " รวมตามลูกค้า"
   ElseIf mcolParam("GROUP_BY_SALE") = "Y" Then
      MODULE_DESC = MODULE_DESC & " รวมตามพนักงานขาย"
   ElseIf mcolParam("GROUP_BY_SALE_AND_CUSTOMER") = "Y" Then
      MODULE_DESC = MODULE_DESC & " รวมตามพนักงานขายและลูกค้า"
   Else
      MODULE_DESC = MODULE_DESC & " รวมตามวันที่ดิวเช็ค"
   End If
   
   Call SetReportConfig(Vsp, mcolParam("REPORT_KEY"))
   Call LoadBktChqnumDocDateAR(Nothing, m_ChqnumDocDate, mcolParam("FROM_CHECK_DATE"), mcolParam("TO_CHECK_DATE")) 'ใช้วันที่ GETDAT เนื่องจากว่าพอรับเช็คจริงต้องได้เงินในวันที่ตาม CHQDAT แต่ต้องรอเครียร์ริ่งอีกจึงเปลี่ยนเป็น GETDAT
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
Private Function GetKey(Ug As CBkTrn) As String
   If mcolParam("GROUP_BY_DUE_DATE") = "Y" Then
      GetKey = Trim(Ug.GETDAT)
   ElseIf mcolParam("GROUP_BY_CUSTOMER") = "Y" Then
      GetKey = Trim(Ug.CUSCOD)
   ElseIf mcolParam("GROUP_BY_SALE") = "Y" Then
      GetKey = Trim(Ug.SLMCOD)
  ElseIf mcolParam("GROUP_BY_SALE_AND_CUSTOMER") = "Y" Then
      GetKey = Trim(Ug.CUSCOD)   ' <------------ ยังไม่ได้แก้ GetKey  เป็น ลูกค้า อ่ะ
Else
      GetKey = Trim(Ug.GETDAT)
   End If
End Function
Private Function GetKey2(Ug As CBkTrn) As String
   If mcolParam("GROUP_BY_DUE_DATE") = "Y" Then
      GetKey2 = Trim(Format(Month(Ug.GETDAT), "00") & "/" & Format(Year(Ug.GETDAT) + 543, "0000"))
   ElseIf mcolParam("GROUP_BY_CUSTOMER") = "Y" Then
      GetKey2 = Trim(Ug.CUSCOD)
   ElseIf mcolParam("GROUP_BY_SALE") = "Y" Then
      GetKey2 = Trim(Ug.SLMCOD)
  ElseIf mcolParam("GROUP_BY_SALE_AND_CUSTOMER") = "Y" Then
      GetKey2 = Trim(Ug.SLMCOD)       ' <------------ ยังไม่ได้แก้ GetKey อันนี้ล่ะ PrevKey2 เก็บ sale
   Else
      GetKey2 = Trim(Format(Month(Ug.GETDAT), "00") & "/" & Format(Year(Ug.GETDAT) + 543, "0000"))
   End If
   
End Function
Private Function GetKey3(Ug As CBkTrn) As String
   If mcolParam("GROUP_BY_DUE_DATE") = "Y" Then
      GetKey3 = Trim(Ug.GETDAT)
   ElseIf mcolParam("GROUP_BY_CUSTOMER") = "Y" Then
      GetKey3 = Trim(Ug.CUSNAM) & " ( " & Trim(Ug.CUSCOD) & ")"
   ElseIf mcolParam("GROUP_BY_SALE") = "Y" Then
      GetKey3 = Trim(Ug.SLMNAM) & " ( " & Trim(Ug.SLMCOD) & ")"
   ElseIf mcolParam("GROUP_BY_SALE_AND_CUSTOMER") = "Y" Then
      GetKey3 = Trim(Ug.CUSNAM) & " ( " & Trim(Ug.CUSCOD) & ")"  ' <------------ ยังไม่ได้แก้ GetKey  // ไม่ต้องก็ได้นะความจริง ใช้ 2 กะ 3 เอา
   Else
      GetKey3 = Trim(Ug.GETDAT)
   End If
   
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
Dim Ug As CBkTrn
Dim St As CIsTab
Dim TempDocDate As Date
Dim CancelFlag As Boolean
Dim Dc As CDocumentCancel

Dim TempCCheck As CCheckCancel
Dim FoundInCheckCancel As Boolean

Dim CountDue1 As Double
Dim CountDue2 As Double
Dim SumDue1 As Double
Dim SumDue2 As Double
Dim SumDue3 As Double

 
   RName = "genDoc"
'-----------------------------------------------------------------------------------------------------
'                                             Query Here
'-----------------------------------------------------------------------------------------------------
   
  Set Rs = New ADODB.Recordset
  Set m_Rs = New ADODB.Recordset
  Set TempCheckCancel = New CCheckCancel
  Set m_ReceiveCancelitem = New Collection
  
  Set TempCCheck = New CCheckCancel
   
   If Not initDoc Then Exit Function
   
   Vsp.Preview = isPreview
'-----------------------------------------------------------------------------------------------------
'                                         Main Operation Here
'-----------------------------------------------------------------------------------------------------
SumDue1 = 0
SumDue2 = 0
SumDue3 = 0
CountDue1 = 0
CountDue2 = 0

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
      Ug.FROM_CUSTOMER_CODE = mcolParam("FROM_CUSTOMER_CODE")
      Ug.TO_CUSTOMER_CODE = mcolParam("TO_CUSTOMER_CODE")
      Ug.FROM_SALE_CODE = mcolParam("FROM_SALE_CODE")
      Ug.TO_SALE_CODE = mcolParam("TO_SALE_CODE")
      Ug.FROM_GETDAT = mcolParam("FROM_CHECK_DATE")
      Ug.TO_GETDAT = mcolParam("TO_CHECK_DATE")
      If mcolParam("CUSTOMER_TYPE") > 0 Then
         Set St = GetItemFromListIndex(m_CustomerTypes, mcolParam("CUSTOMER_TYPE"))
         Ug.CUSTYP = St.TYPCOD
      Else
         Ug.CUSTYP = ""
      End If
      If mcolParam("GROUP_BY_DUE_DATE") = "Y" Then
         Ug.OrderBy = 1
      ElseIf mcolParam("GROUP_BY_CUSTOMER") = "Y" Then
         Ug.OrderBy = 2
      ElseIf mcolParam("GROUP_BY_SALE") = "Y" Then
         Ug.OrderBy = 3
       ElseIf mcolParam("GROUP_BY_SALE_AND_CUSTOMER") = "Y" Then
         Ug.OrderBy = 4
         Else
         Ug.OrderBy = 1
      End If
      
      Ug.OrderType = mcolParam("ORDER_TYPE")
      Call Ug.QueryData(6, Rs, iCount)
      i = 0
      
      If Not Rs.EOF Then
         Call Ug.PopulateFromRS(6, Rs)
         PrevKey1 = GetKey(Ug)
         PrevKey2 = GetKey2(Ug)              ' เก็บรหัสลูกค้า
         PrevKey3 = GetKey3(Ug)
      End If
      
      While Not Rs.EOF
         Call Ug.PopulateFromRS(6, Rs)
         
         Set TempCCheck = GetCheckCancelitem(m_ReceiveCancelitem, Ug.CHQNUM, False)
         If (TempCCheck Is Nothing) Then
                        FoundInCheckCancel = False
         Else
                        FoundInCheckCancel = True
          End If

         
 If FoundInCheckCancel = False And (Ug.VOUCHER >= mcolParam("FROM_DOCUMENT_NO") And Ug.VOUCHER <= mcolParam("TO_DOCUMENT_NO")) Or (Ug.VOUCHER >= mcolParam("FROM_DOCUMENT_NO1") And Ug.VOUCHER <= mcolParam("TO_DOCUMENT_NO1")) Then
            i = i + 1
            
            If PrevKey2 <> GetKey2(Ug) And i > 1 And GetKey(Ug) <> GetKey2(Ug) Then     ' PrevKey2 = ชื่อลูกค้า
                                    If mcolParam("SHOW_SUMMARY") = "N" Then
                                       Call GenerateFooter("รวม", "", PrevKey1, Total1, m_Details1, 0)                         ' รวมย่อยสุด RE เดียวกัน
                                       Call m_Details1.GetString(2, TempStr1, TempStr2)
                                       strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
                                       Vsp.FontBold = True
                                       Call Vsp.AddTable(strFormat, "", TempStr2)
                                       Vsp.FontBold = False
                                    Else
                                       Call GenerateFooter("", "", PrevKey3, Total1, m_Details1, 0)
                                       Call m_Details1.GetString(2, TempStr1, TempStr2)
                                       strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
                                       'Vsp.FontBold = True
                                       Call Vsp.AddTable(strFormat, "", TempStr2)
                                       Vsp.FontBold = False
                                    End If
               
                     Call GenerateFooter("รวม", "", "* " & PrevKey2, Total2, m_Details1, 0)             ' รวมเดือนเดียวกัน  ,,, อันนี้ล่ะ ที่ sale เดียวกัน
                     Call m_Details1.GetString(2, TempStr1, TempStr2)
                     strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
                     Vsp.FontBold = True
                     Call Vsp.AddTable(strFormat, "", TempStr2)
                     Vsp.FontBold = False

                     
                     For j = 1 To UBound(Total1)
                        Total1(j) = 0
                        Total2(j) = 0
                     Next j
                    SumDue1 = 0
                     CountDue1 = 0
                     SumDue2 = 0
                     CountDue2 = 0
                     
            
            ElseIf PrevKey1 <> GetKey(Ug) And i > 1 Then
                            If mcolParam("SHOW_SUMMARY") = "N" Then
                               Call GenerateFooter("รวม", "", PrevKey1, Total1, m_Details1, 0)
                               Call m_Details1.GetString(2, TempStr1, TempStr2)
                               strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
                               Vsp.FontBold = True
                               Call Vsp.AddTable(strFormat, "", TempStr2)
                               Vsp.FontBold = False
                            Else
                               Call GenerateFooter("", "", PrevKey3, Total1, m_Details1, 0)                       ' เมื่อติ๊ก พนักงานขาย PrevKey3 = ชื่อพนักงานขาย ,, แต่ถ้าติ๊กลูกค้า  = ขื่อลูกค้า
                               Call m_Details1.GetString(2, TempStr1, TempStr2)
                               strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
                               'Vsp.FontBold = True
                               Call Vsp.AddTable(strFormat, "", TempStr2)
                               Vsp.FontBold = False
                            End If
               
                        For j = 1 To UBound(Total1)
                           Total1(j) = 0
                        Next j
                        CountDue1 = 0
                         SumDue1 = 0
            End If
            
                  If CountDue2 = 0 And mcolParam("GROUP_BY_SALE_AND_CUSTOMER") = "Y" Then
                      Call GenerateHeader(" ( " & GetKey2(Ug) & " )   " & Ug.SLMNAM, m_Details1)             ' รวมเดือนเดียวกัน  ,,, อันนี้ล่ะ ที่ sale เดียวกัน
                                        Call m_Details1.GetString(2, TempStr1, TempStr2)
                                        strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
                                        Vsp.FontBold = True
                                        Call Vsp.AddTable(strFormat, "", TempStr2)
                                        Vsp.FontBold = False
                  End If
            
            Call m_Details1.ClearField
   
            PrevKey1 = GetKey(Ug)
            PrevKey2 = GetKey2(Ug)          'เก็บรหัสลูกค้า
            PrevKey3 = GetKey3(Ug)
            
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
            TempStr = DateToStringExtEx2(Ug.GETDAT)
            Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
            Call m_Details1.AddField(BodyCf)
            
            Set HeadCf = m_Headers1.Fields(5)
            TempStr = DateToStringExtEx2(Ug.PAYINDAT)
            Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
            Call m_Details1.AddField(BodyCf)
            
            Set HeadCf = m_Headers1.Fields(6)
            TempStr = Left(Ug.BNKCOD, 3)
            Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
            Call m_Details1.AddField(BodyCf)
            
            Set HeadCf = m_Headers1.Fields(7)
            TempStr = Left(Ug.BRANCH, 14)
            Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
            Call m_Details1.AddField(BodyCf)
    
            Set HeadCf = m_Headers1.Fields(8)
            TempStr = Ug.CUSCOD
            Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
            Call m_Details1.AddField(BodyCf)
            
            Set HeadCf = m_Headers1.Fields(9)
            If Len(Ug.CUSNAM) > 0 Then
               TempStr = Ug.CUSNAM
            Else
               TempStr = Ug.REMARK
            End If
            TempStr = Left(TempStr, 20)
            Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
            Call m_Details1.AddField(BodyCf)
   
            Set HeadCf = m_Headers1.Fields(10)
            If CancelFlag Then
               TempStr = "***" & Ug.VOUCHER
            Else
               TempStr = Ug.VOUCHER
            End If
            Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
            Call m_Details1.AddField(BodyCf)
   
            Set HeadCf = m_Headers1.Fields(11)
            TempStr = FormatNumber(Ug.AMOUNT)
            Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
            Call m_Details1.AddField(BodyCf)
             CountDue1 = CountDue1 + 1
            CountDue2 = CountDue2 + 1
            If Not (CancelFlag) And FoundInCheckCancel = False Then
               Total1(11) = Total1(11) + Ug.AMOUNT
               Total2(11) = Total2(11) + Ug.AMOUNT
               Total3(11) = Total3(11) + Ug.AMOUNT
            End If
            
            Set HeadCf = m_Headers1.Fields(12)
            TempStr = CalculateCreditCompare(Ug.CHQNUM, Ug.CHQDAT, TempDocDate)
            TempStr = DateToStringExtEx2(TempDocDate)
            Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
            Call m_Details1.AddField(BodyCf)
                        
            Set HeadCf = m_Headers1.Fields(13)
            TempStr = CalculateCreditCompare(Ug.CHQNUM, Ug.CHQDAT, TempDocDate)
            Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
            Call m_Details1.AddField(BodyCf)
            SumDue1 = SumDue1 + TempStr
            SumDue2 = SumDue2 + TempStr
            SumDue3 = SumDue3 + TempStr

             If CountDue1 <> 0 Then
               Total1(13) = Round(SumDue1 / CountDue1)
               Total2(13) = Round(SumDue2 / CountDue2)
               Total3(13) = Round(SumDue3 / i)
                End If
            
            Set HeadCf = m_Headers1.Fields(14)
            TempStr = Left(Ug.SLMNAM, 15)
            Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
            Call m_Details1.AddField(BodyCf)
                                                         
            If mcolParam("SHOW_SUMMARY") = "N" And FoundInCheckCancel = False Then
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
         Call GenerateFooter("รวม", "", PrevKey1, Total1, m_Details1, 0)            ' รวม RE เดียวกัน ครั้งสุดท้าย
         Call m_Details1.GetString(2, TempStr1, TempStr2)
         strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
         Vsp.FontBold = True
         Call Vsp.AddTable(strFormat, "", TempStr2)
         Vsp.FontBold = False
      Else
         Call GenerateFooter("", "", PrevKey3, Total1, m_Details1, 0)                'เมื่อติ๊กพนักงานขาย ชื่อพนักงานขายคนสุดท้าย
         Call m_Details1.GetString(2, TempStr1, TempStr2)
         strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
         'Vsp.FontBold = True
         Call Vsp.AddTable(strFormat, "", TempStr2)
         Vsp.FontBold = False
      End If
      
      If PrevKey1 <> PrevKey2 Then
         Call GenerateFooter("รวม", "", "* " & PrevKey2, Total2, m_Details1, 0)
         Call m_Details1.GetString(2, TempStr1, TempStr2)
         strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
         Vsp.FontBold = True
         Call Vsp.AddTable(strFormat, "", TempStr2)
         Vsp.FontBold = False
      End If
      
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

Private Sub GenerateHeader(Txt1 As String, Details As CFieldList)
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
      Else:
         TempStr = ""
      End If
      
      Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.Align, TempStr, HeadCf.BodyAlign)
      Call Details.AddField(BodyCf)
   Next HeadCf
      
   Set BodyCf = Nothing
End Sub

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
      ElseIf j = 4 Then
         TempStr = Txt3
      ElseIf j = 11 Then
         TempStr = FormatNumber(Tot(j))
      ElseIf j = 13 Then
         TempStr = Tot(j)
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






