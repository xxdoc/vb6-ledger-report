VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CReportAR005_2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'User Group Report
Implements CReportInterface

Private Const MODULE_NAME = "CReportAR005"
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
Private SaleWidth As Double

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

Private m_ReceiveAmounts As Collection
Private m_ReceiveAllAmounts As Collection
Private m_MaxMins As Collection
Private m_DbnCdnByBills As Collection
Private m_RealCreditColl As Collection
Private m_CnAllAmounts As Collection
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
   Set m_YearSeqs = New Collection
   Set m_ReceiveAmounts = New Collection
   Set m_MaxMins = New Collection
   Set m_ReceiveAllAmounts = New Collection
   Set m_DbnCdnByBills = New Collection
   Set m_RealCreditColl = New Collection
   Set m_CnAllAmounts = New Collection
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
   Set m_YearSeqs = Nothing
   Set m_Headers1_1 = Nothing
   Set m_Headers1_2 = Nothing
   Set m_Headers1_3 = Nothing
   Set m_Headers1_4 = Nothing
   Set m_Details2 = Nothing
   Set m_ReceiveAmounts = Nothing
   Set m_MaxMins = Nothing
   Set m_ReceiveAllAmounts = Nothing
   Set m_DbnCdnByBills = Nothing
   Set m_RealCreditColl = Nothing
   Set m_CnAllAmounts = Nothing
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
Dim FromDate As Date
Dim ToDate As Date
Dim TempDate As Date
Dim TempWidth As Double
   
   tbdT = Vsp.TableBorder
   blnBold = Vsp.FontBold
   iSize = Vsp.FontSize
   
   Vsp.FontSize = HEADER_SIZE
   Vsp.FontBold = True
   Vsp.TableBorder = tbAll
   
   Call m_Headers0.ClearField
   Call m_Headers1.ClearField
   
   Set Cf = New CReportField
      
   '===
   SaleWidth = 0
   If mcolParam("SUMMARY_MODE") = "Y" Or mcolParam("SALE_AND_CUSTOMER") = "Y" Then
      Call Cf.SetFieldValue(0, "^", "NO.", "^")
      Call m_Headers0.AddField(Cf)
   Else
      Call Cf.SetFieldValue(1, "^", "NO.", "^")
      Call m_Headers0.AddField(Cf)
   End If
   If mcolParam("SUMMARY_MODE") = "Y" Then
      Call Cf.SetFieldValue(2, "<", "", "<")
      Call m_Headers0.AddField(Cf)
      Call Cf.SetFieldValue(4, "<", "", "<")
      Call m_Headers0.AddField(Cf)
   ElseIf mcolParam("SALE_AND_CUSTOMER") = "Y" Then
      Call Cf.SetFieldValue(2, "<", "", "<")
      Call m_Headers0.AddField(Cf)
      Call Cf.SetFieldValue(6, "<", "", "<")
      Call m_Headers0.AddField(Cf)
      
      SaleWidth = SaleWidth + 6
   Else
      Call Cf.SetFieldValue(2, "^", "�Ţ���", "^")
      Call m_Headers0.AddField(Cf)
      Call Cf.SetFieldValue(2, "^", "", "^")
      Call m_Headers0.AddField(Cf)
      SaleWidth = SaleWidth + 2
   End If
   
   If mcolParam("SUMMARY_MODE") = "Y" Or mcolParam("SALE_AND_CUSTOMER") = "Y" Then
      Call Cf.SetFieldValue(0, "^", "", "^")
      Call m_Headers0.AddField(Cf)
   Else
      Call Cf.SetFieldValue(2, "^", "", "^")
      Call m_Headers0.AddField(Cf)
      SaleWidth = SaleWidth + 2
   End If
   
   FromDate = mcolParam("FROM_MONTH")
   ToDate = mcolParam("TO_MONTH")
   Call GetFirstLastDate(mcolParam("FROM_MONTH"), FromDate, TempDate)
   Call GetFirstLastDate(mcolParam("TO_MONTH"), TempDate, ToDate)
      
   TempWidth = 0
   While (FromDate < ToDate)
      TempWidth = TempWidth + 2
      SaleWidth = SaleWidth + 2
      FromDate = DateAdd("M", 1, FromDate)
   Wend
   Call Cf.SetFieldValue(TempWidth + 4, "^", "��ǧ DUE", ">")
   Call m_Headers0.AddField(Cf)
   SaleWidth = SaleWidth + 4
   
   Call Cf.SetFieldValue(2, "^", "", "^")
   Call m_Headers0.AddField(Cf)
   SaleWidth = SaleWidth + 2
   
   If mcolParam("SUMMARY_MODE") = "Y" Or mcolParam("SALE_AND_CUSTOMER") = "Y" Then
      Call Cf.SetFieldValue(0, "^", "", "^")
      Call m_Headers0.AddField(Cf)
   Else
      Call Cf.SetFieldValue(2, "^", "�Թ", "^")
      Call m_Headers0.AddField(Cf)
      SaleWidth = SaleWidth + 2
   End If
   
   Call Cf.SetFieldValue(2, "^", "�����˵�", "^")
   Call m_Headers0.AddField(Cf)
   SaleWidth = SaleWidth + 2
   '===
   
   If mcolParam("SUMMARY_MODE") = "Y" Or mcolParam("SALE_AND_CUSTOMER") = "Y" Then
      Call Cf.SetFieldValue(0, "^", "NO.", "^")
      Call m_Headers1.AddField(Cf)
   Else
      Call Cf.SetFieldValue(1, "^", "NO.", "^")
      Call m_Headers1.AddField(Cf)
   End If
   
   If mcolParam("SUMMARY_MODE") = "Y" Then
      Call Cf.SetFieldValue(2, "<", "����", "<")
      Call m_Headers1.AddField(Cf)
      Call Cf.SetFieldValue(4, "<", "��ѡ�ҹ���", "<")
      Call m_Headers1.AddField(Cf)
   ElseIf mcolParam("SALE_AND_CUSTOMER") = "Y" Then
      Call Cf.SetFieldValue(2, "<", "����", "<")
      Call m_Headers1.AddField(Cf)
      Call Cf.SetFieldValue(6, "<", "�١���", "<")
      Call m_Headers1.AddField(Cf)
   Else
      Call Cf.SetFieldValue(2, "^", "��觢ͧ", "^")
      Call m_Headers1.AddField(Cf)
      Call Cf.SetFieldValue(2, "^", "�ѹ�����", "^")
      Call m_Headers1.AddField(Cf)
   End If
   
   If mcolParam("SUMMARY_MODE") = "Y" Or mcolParam("SALE_AND_CUSTOMER") = "Y" Then
      Call Cf.SetFieldValue(0, "^", "", "^")
      Call m_Headers1.AddField(Cf)
   Else
      Call Cf.SetFieldValue(2, "^", "�ú��˹�", "^")
      Call m_Headers1.AddField(Cf)
   End If
   
   FromDate = mcolParam("FROM_MONTH")
   ToDate = mcolParam("TO_MONTH")
   Call GetFirstLastDate(mcolParam("FROM_MONTH"), FromDate, TempDate)
   Call GetFirstLastDate(mcolParam("TO_MONTH"), TempDate, ToDate)
      
   Call Cf.SetFieldValue(2, "^", "< " & Format(Month(FromDate), "00") & "/" & Format(Year(FromDate) + 543, "00"), ">")
   Call m_Headers1.AddField(Cf)
      
   While (FromDate < ToDate)
      Call Cf.SetFieldValue(2, "^", Format(Month(FromDate), "00") & "/" & Format(Year(FromDate) + 543, "00"), ">")
      Call m_Headers1.AddField(Cf)
      FromDate = DateAdd("M", 1, FromDate)
   Wend
    
    FromDate = DateAdd("M", -1, FromDate)
    Call Cf.SetFieldValue(2, "^", Format(Month(FromDate), "00") & "/" & Format(Year(FromDate) + 543, "00") & ">", ">")
   Call m_Headers1.AddField(Cf)
    
   Call Cf.SetFieldValue(2, "^", "���˹���ҧ", ">")
   Call m_Headers1.AddField(Cf)
   
   If mcolParam("SUMMARY_MODE") = "Y" Or mcolParam("SALE_AND_CUSTOMER") = "Y" Then
      Call Cf.SetFieldValue(0, "^", "(�ѹ)", "^")
      Call m_Headers1.AddField(Cf)
   Else
      Call Cf.SetFieldValue(2, "^", "", "^")
      Call m_Headers1.AddField(Cf)
   End If
   
   Call Cf.SetFieldValue(2, "^", "", "^")
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
   Vsp.MarginBottom = 700
   Vsp.MarginFooter = 700
   Vsp.MarginHeader = 700
   Vsp.MarginLeft = 700
   Vsp.MarginRight = 700
   Vsp.MarginTop = 700
   Vsp.FontName = "AngsanaUPC"
   Vsp.FontSize = DETAIL_SIZE
   
   Call LoadReceiveAmountByBill(Nothing, m_ReceiveAmounts, -1, mcolParam("TO_PAY_DATE"))
   Call LoadReceiveAmountByBill(Nothing, m_ReceiveAllAmounts, -1, -1)
   
   Call LoadCustomerType(Nothing, m_CustomerTypes)
   Call LoadARCNAmountByBill(Nothing, m_CnAllAmounts, DateAdd("D", 1, mcolParam("TO_DOC_DATE")), -1, -1, mcolParam("TO_DOC_DATE"))         '��ͧ�ǡ�ʹ¡������ �ʹ CN �������ͧ�����ѧ CN �� �͡��� LINK ��㹪�ǧ¡��
   
   FromDate = "�ҡ�ѹ��� " & EmptyToString(DateToStringExtEx2(mcolParam("FROM_DOC_DATE")), "N/A")
   ToDate = " �֧ �ѹ��� " & EmptyToString(DateToStringExtEx2(mcolParam("TO_DOC_DATE")), "N/A")
   ToPaidDate = "    �ѹ�������Թ�֧�ѹ��� " & EmptyToString(DateToStringExtEx2(mcolParam("TO_PAY_DATE")), "N/A") & "    �ѹ������� " & EmptyToString(DateToStringExtEx2(mcolParam("PRINT_DATE")), "N/A")
   
   'glbEnterPrise.ENTERPRISE_NAME & " , " & glbEnterPrise.BRANCH_NAME
   MODULE_DESC = glbCompanyName & vbCrLf & mcolParam("REPORT_NAME") & _
                                       vbCrLf & FromDate & ToDate & "  " & ToPaidDate & " �ҡ�١��� " & EmptyToString(mcolParam("FROM_CUSTOMER_CODE"), "N/A") & " �֧�١��� " & EmptyToString(mcolParam("TO_CUSTOMER_CODE"), "N/A")
   
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

Private Sub GenerateDetail(alngX() As Long, i As Long, ARt As CARTrn, Tot1() As Double, Tot2() As Double, Tot3() As Double, Optional HaveHeader As Boolean = False, Optional PrevKey1 As String, Optional PrevKey4 As String)
Dim HeadCf As CReportField
Dim TempStr As String
Dim TempStr1 As String
Dim TempStr2 As String
Dim BodyCf As CReportField
Dim strFormat As String
Dim Ari As CARRcIt
Dim Ari2 As CARRcIt
Dim Apt1 As CARTrn
Dim ARt2 As CARTrn
Dim DueCount As Long
Dim MM As CMaxMin
Dim j As Long
Dim Sum As Double
Dim PrintDate As Date
Dim PaidAmount As Double
Dim PaidBalance As Double '��������¡��
Dim TempAmt As Double
Dim TempRealCredit As CRealCredit
Dim Width  As Double
Dim FromDate As Date
Dim ToDate As Date
Dim TempDate As Date
Dim TempFromdate As Date
Dim TempToDate As Date
Dim NewDueDate As Date
   
   Set BodyCf = New CReportField
   
   If mcolParam("PRINT_DATE") <= 0 Then
      PrintDate = Now
   Else
      PrintDate = mcolParam("PRINT_DATE")
   End If
   NewDueDate = DateAdd("D", mcolParam("CREDIT_DAY"), ARt.DOCDAT)
   DueCount = DateDiff("D", NewDueDate, PrintDate)
   
   Set Ari = GetARRcpItem(m_ReceiveAmounts, ARt.DOCNUM)
   Set Ari2 = GetARRcpItemEx(m_ReceiveAllAmounts, ARt.DOCNUM)
   
'   Set Apt1 = GetAPTrn(m_DbnCdnByBills, Apt.DocNum)
   Set ARt2 = GetARTrn(m_CnAllAmounts, ARt.DOCNUM)
   '�Ѻ��������¡����ҡѺ APT.PAYAMT - �ʹ�Ѻ���з������ͧ��Ӥѭ����
   PaidBalance = ARt.RCVAMT - Ari2.RCVAMT - ARt2.AMOUNT
   If (ARt.AMOUNT) > (Ari.RCVAMT + PaidBalance) Then  '�Ѻ�ѧ���ú
'      If Art.RECTYP = 5 Then 'Ŵ˹��
'         Art.AMOUNT = -1 * Art.AMOUNT
'      End If
      
      '-----------------------------      '�ʴ�����١���
      If (HaveHeader) And PrevKey1 <> Trim(ARt.CUSCOD) And (mcolParam("INCLUDE_CUSTOMER_MODE") = "Y") And mcolParam("SUMMARY_MODE") = "N" Then
         Call GenerateFooter("��� ", PrevKey1, "", Tot3, m_Details1, 0)   '�������١���
         Call m_Details1.GetString(2, TempStr1, TempStr2)
         strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
         Vsp.FontBold = True
         Call Vsp.AddTable(strFormat, "", TempStr2)
         Vsp.FontBold = False
         
         HaveHeader = False
         For j = 1 To UBound(Tot3)
            Tot3(j) = 0
         Next j
      ElseIf (HaveHeader) And mcolParam("SALE_AND_CUSTOMER") = "Y" And PrevKey1 <> Trim(ARt.CUSCOD) Then
         Call GenerateFooter("", PrevKey1, PrevKey4, Tot3, m_Details1, 0)   '�������١���
         Call m_Details1.GetString(2, TempStr1, TempStr2)
         strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
         
         Call Vsp.AddTable(strFormat, "", TempStr2)
         Vsp.FontBold = False
         
         HaveHeader = False
         For j = 1 To UBound(Tot3)
            Tot3(j) = 0
         Next j
      End If
         
      If Not (HaveHeader) And PrevKey1 <> Trim(ARt.CUSCOD) And (mcolParam("INCLUDE_CUSTOMER_MODE") = "Y") And mcolParam("SUMMARY_MODE") = "N" Then
         Call m_Details1.ClearField
         Width = m_Headers0.Fields(1).Width + m_Headers0.Fields(2).Width + m_Headers0.Fields(3).Width
         TempStr = ARt.CUSCOD
         Call BodyCf.SetFieldValue(Width, "^", TempStr)
         Call m_Details1.AddField(BodyCf)
         
         Width = SaleWidth - m_Headers0.Fields(3).Width
         TempStr = ARt.CUSNAM
         Call BodyCf.SetFieldValue(Width, "<", TempStr)
         Call m_Details1.AddField(BodyCf)
   
         Call m_Details1.GetString(1, TempStr1, TempStr2)
         strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
         Call Vsp.AddTable(strFormat, "", TempStr2)
         HaveHeader = True
      ElseIf Not (HaveHeader) And PrevKey1 <> Trim(ARt.CUSCOD) And (mcolParam("SALE_AND_CUSTOMER") = "Y") And mcolParam("SUMMARY_MODE") = "N" Then
         HaveHeader = True
      End If
      PrevKey1 = Trim(ARt.CUSCOD)
      PrevKey4 = Trim(ARt.CUSNAM)
      '-----------------------------      '�ʴ�����١���
      
      
      Call m_Details1.ClearField
      Set HeadCf = m_Headers1.Fields(1)
      TempStr = i
      Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
      Call m_Details1.AddField(BodyCf)
   
      Set HeadCf = m_Headers1.Fields(2)
      TempStr = ARt.DOCNUM
      Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
      Call m_Details1.AddField(BodyCf)

      Set HeadCf = m_Headers1.Fields(3)
      TempStr = DateToStringExtEx2(ARt.DOCDAT)
      Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
      Call m_Details1.AddField(BodyCf)
      
      Set HeadCf = m_Headers1.Fields(4)
      TempStr = DateToStringExtEx2(NewDueDate)
      Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
      Call m_Details1.AddField(BodyCf)
      
      j = 5
      Sum = 0
      
      FromDate = mcolParam("FROM_MONTH")
      ToDate = mcolParam("TO_MONTH")
      Call GetFirstLastDate(mcolParam("FROM_MONTH"), FromDate, TempDate)
      Call GetFirstLastDate(mcolParam("TO_MONTH"), TempDate, ToDate)
         
      If (NewDueDate < FromDate) Then
         TempAmt = (ARt.AMOUNT) - (Ari.RCVAMT + PaidBalance)
         If ARt.RECTYP = 5 Then 'Ŵ˹��
            TempAmt = -1 * TempAmt
         End If
         TempStr = FormatNumber(TempAmt, "-")
         
         Tot1(j) = Tot1(j) + TempAmt
         Tot2(j) = Tot2(j) + TempAmt
         Tot3(j) = Tot3(j) + TempAmt
         Sum = Sum + TempAmt
      Else
         TempStr = FormatNumber(0, "-")
      End If
      
      Set HeadCf = m_Headers1.Fields(j)
      Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
      Call m_Details1.AddField(BodyCf)
      
      While (FromDate < ToDate)
         j = j + 1
         Call GetFirstLastDate(FromDate, TempFromdate, TempToDate)
         If (TempFromdate <= NewDueDate) And (NewDueDate <= TempToDate) Then
            TempAmt = (ARt.AMOUNT) - (Ari.RCVAMT + PaidBalance)
            If ARt.RECTYP = 5 Then 'Ŵ˹��
               TempAmt = -1 * TempAmt
            End If
            TempStr = FormatNumber(TempAmt, "-")
            
            Tot1(j) = Tot1(j) + TempAmt
            Tot2(j) = Tot2(j) + TempAmt
            Tot3(j) = Tot3(j) + TempAmt
            Sum = Sum + TempAmt
         Else
            TempStr = FormatNumber(0, "-")
         End If
         
         Set HeadCf = m_Headers1.Fields(j)
         Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
         Call m_Details1.AddField(BodyCf)
         
         FromDate = DateAdd("M", 1, FromDate)
      Wend
      
      j = j + 1
      If (NewDueDate > ToDate) Then
         TempAmt = (ARt.AMOUNT) - (Ari.RCVAMT + PaidBalance)
         If ARt.RECTYP = 5 Then 'Ŵ˹��
            TempAmt = -1 * TempAmt
         End If
         TempStr = FormatNumber(TempAmt, "-")
         
         Tot1(j) = Tot1(j) + TempAmt
         Tot2(j) = Tot2(j) + TempAmt
         Tot3(j) = Tot3(j) + TempAmt
         Sum = Sum + TempAmt
      Else
         TempStr = FormatNumber(0, "-")
      End If
      Set HeadCf = m_Headers1.Fields(j)
      Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
      Call m_Details1.AddField(BodyCf)
         
      j = j + 1
      Set HeadCf = m_Headers1.Fields(j)
      TempStr = FormatNumber(Sum, "-")
      Tot1(j) = Tot1(j) + TempAmt
      Tot2(j) = Tot2(j) + TempAmt
      Tot3(j) = Tot3(j) + TempAmt
      Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
      Call m_Details1.AddField(BodyCf)

      j = j + 1
      Set HeadCf = m_Headers1.Fields(j)
      TempStr = FormatNumberInt(DueCount)
      Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
      Call m_Details1.AddField(BodyCf)
      
      j = j + 1
      Set HeadCf = m_Headers1.Fields(j)
      TempStr = ""
      Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
      Call m_Details1.AddField(BodyCf)
      
      If mcolParam("SUMMARY_MODE") = "N" Then
            If mcolParam("SALE_AND_CUSTOMER") = "N" Then
            Call m_Details1.GetString(1, TempStr1, TempStr2)
            strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
            Call Vsp.AddTable(strFormat, "", TempStr2)
         End If
      End If
   Else
      i = i - 1
   End If
   Set BodyCf = Nothing
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
Dim Total1(100) As Double
Dim Total2(100) As Double
Dim Total3(100) As Double
Dim iCount As Long
Dim TempStr As String
Dim PrevKey1 As String
Dim PrevKey2 As String
Dim PrevKey3 As String
Dim PrevKey4 As String
Dim St As CIsTab
Dim Ar As COESLM
Dim Width As Double
Dim ARt As CARTrn
Dim TempRs As ADODB.Recordset
Dim HaveHeader As Boolean
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
      i = 0
      Set BodyCf = New CReportField
      
      Set Ar = New COESLM
      Ar.FROM_SALE_CODE = mcolParam("FROM_SALE_CODE")
      Ar.TO_SALE_CODE = mcolParam("TO_SALE_CODE")
      Ar.OrderBy = mcolParam("ORDER_TYPE")
      Ar.OrderType = mcolParam("ORDER_BY")
      Call glbDaily.QuerySale(Ar, Rs, iCount, IsOK, glbErrorLog)
      
      While Not Rs.EOF
         For j = 1 To UBound(Total1)
            Total1(j) = 0
         Next j

         Call m_Details1.ClearField
         Call Ar.PopulateFromRS(1, Rs)
         
         If True Then  '���ʹ��ҧ����
            '===Head
            If mcolParam("SUMMARY_MODE") = "N" Then
               Width = m_Headers0.Fields(1).Width + m_Headers0.Fields(2).Width
               TempStr = Ar.SLMCOD
               Call BodyCf.SetFieldValue(Width, "^", TempStr)
               Call m_Details1.AddField(BodyCf)
   
               Width = SaleWidth
               TempStr = Ar.SLMNAM
               Call BodyCf.SetFieldValue(Width, "<", TempStr)
               Call m_Details1.AddField(BodyCf)
   
               Call m_Details1.GetString(1, TempStr1, TempStr2)
               strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
               Call Vsp.AddTable(strFormat, "", TempStr2)
            End If
            PrevKey2 = Ar.SLMCOD
            PrevKey3 = Ar.SLMNAM
            
            '===

            '=== Detail
            Set ARt = New CARTrn
            ARt.SLMCOD = Ar.SLMCOD
            ARt.FROM_DOC_DATE = mcolParam("FROM_DOC_DATE")
            ARt.TO_DOC_DATE = mcolParam("TO_DOC_DATE")
            
             If mcolParam("CUSTOMER_TYPE") > 0 Then
               Set St = GetItemFromListIndex(m_CustomerTypes, mcolParam("CUSTOMER_TYPE"))
               ARt.CUSTYP = St.TYPCOD
            Else
               ARt.CUSTYP = ""
            End If
            
            ARt.FROM_CUSTOMER_CODE = mcolParam("FROM_CUSTOMER_CODE")
            ARt.TO_CUSTOMER_CODE = mcolParam("TO_CUSTOMER_CODE")
            
            ARt.RECTYP = ""
            ARt.RecTypeSet = "('3', '4', '5')"
            If mcolParam("INCLUDE_CUSTOMER_MODE") = "Y" Or mcolParam("SALE_AND_CUSTOMER") = "Y" Or mcolParam("SUMMARY_MODE") = "Y" Then
               ARt.OrderBy = 3
            Else
               ARt.OrderBy = 1
            End If
            ARt.OrderType = 1
            Call ARt.QueryData(1, TempRs, iCount)
            
            HaveHeader = False
            
            While Not TempRs.EOF
               i = i + 1
               Call ARt.PopulateFromRS(1, TempRs)
               
               Call GenerateDetail(alngX(), i, ARt, Total1, Total2, Total3, HaveHeader, PrevKey1, PrevKey4)
                     
               TempRs.MoveNext
            Wend
            If HaveHeader And mcolParam("INCLUDE_CUSTOMER_MODE") = "Y" And mcolParam("SUMMARY_MODE") = "N" Then
               Call GenerateFooter("��� ", PrevKey1, "", Total3, m_Details1, 0)   '�������١���
               Call m_Details1.GetString(2, TempStr1, TempStr2)
               strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
               Vsp.FontBold = True
               Call Vsp.AddTable(strFormat, "", TempStr2)
               Vsp.FontBold = False
               
               HaveHeader = False
               For j = 1 To UBound(Total3)
                  Total3(j) = 0
               Next j
            ElseIf HaveHeader And mcolParam("SALE_AND_CUSTOMER") = "Y" Then
               Call GenerateFooter("", PrevKey1, PrevKey4, Total3, m_Details1, 0)   '�������١���
               Call m_Details1.GetString(2, TempStr1, TempStr2)
               strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
               Call Vsp.AddTable(strFormat, "", TempStr2)
               Vsp.FontBold = False
               
               HaveHeader = False
               For j = 1 To UBound(Total3)
                  Total3(j) = 0
               Next j
            End If
            
            Set ARt = Nothing
            
            If mcolParam("SUMMARY_MODE") = "Y" Then
               Call GenerateFooter("", PrevKey2, PrevKey3, Total1, m_Details1, 0)   '��������ѡ�ҹ���
               Call m_Details1.GetString(2, TempStr1, TempStr2)
               strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
               Call Vsp.AddTable(strFormat, "", TempStr2)
            Else
               Call GenerateFooter("���", PrevKey2, "", Total1, m_Details1, 0)   '��������ѡ�ҹ���
               Call m_Details1.GetString(2, TempStr1, TempStr2)
               strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
               Vsp.FontBold = True
               Call Vsp.AddTable(strFormat, "", TempStr2)
               Vsp.FontBold = False
            End If
            '===
          End If

         Rs.MoveNext
      Wend
      
      If mcolParam("SUMMARY_MODE") = "Y" Then
         Call GenerateFooter("", "�ط��", "", Total2, m_Details1, 0)
      Else
         Call GenerateFooter("�ط��", "", "", Total2, m_Details1, 0)
      End If
      Call m_Details1.GetString(2, TempStr1, TempStr2)
      strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
      Vsp.FontBold = True
      Call Vsp.AddTable(strFormat, "", TempStr2)
      Vsp.FontBold = False
      
      Set Ar = Nothing
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
      ElseIf j = 2 Then
         TempStr = Txt2
      ElseIf j = 3 Then
         TempStr = Txt3
      ElseIf j = (m_Headers1.Fields.Count) Or (j = (m_Headers1.Fields.Count - 1) And (mcolParam("SUMMARY_MODE") = "N")) Then
         TempStr = ""
      Else
         TempStr = FormatNumber(Tot(j), "-")
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
   Call Vsp.AddTable(strFormat, "", "|˹�ҷ��|: " & Vsp.CurrentPage)
   Call Vsp.AddTable(strFormat, "", "|�ѹ|: " & DateToStringExt(mdteDate))
   Call Vsp.AddTable(strFormat, "", "|����|: " & Right$(DateToStringExtEx(Now), 8))
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

