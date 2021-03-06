VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CReportAP008"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'User Group Report
Implements CReportInterface

Private Const MODULE_NAME = "CReportAP007"
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
Private m_ApSupAmounts As Collection
Private m_ApSupAmountCns As Collection
Private m_ApSupAmountCashs As Collection

Private m_ApSupAmountTo As Collection
Private m_ApSupAmountToCn As Collection

Private m_ApSupPaids As Collection
Private m_ApSupPaidCns As Collection
Private m_ApSupPaidTo As Collection
Private m_ApSupPaidToCn As Collection

Private m_PaidAmounts As Collection
Private m_PaidAllAmounts As Collection
Private m_PaidAllAmountCns As Collection

Private m_MaxMins As Collection
Private m_DbnCdnByBills As Collection

Private m_5 As Collection
Private m_6 As Collection
Private m_7 As Collection
Private m_8 As Collection
Private m_9 As Collection
Private m_10 As Collection
Private m_11 As Collection

Private SupplierGroupColl As Collection
Private GroupTypeColl As Collection

Private HaveValue As Boolean
Private HaveValueSum As Boolean
Private FirstDate As Date
Private LastDate As Date

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
   Set m_ApSupAmountTo = New Collection
   Set m_ApSupPaidTo = New Collection
   Set m_ApSupAmountCns = New Collection
   Set m_ApSupPaidCns = New Collection
   Set m_PaidAllAmountCns = New Collection
   Set m_ApSupAmountToCn = New Collection
   Set m_ApSupPaidToCn = New Collection
   Set m_ApSupAmountCashs = New Collection
   
   Set SupplierGroupColl = New Collection
   Set GroupTypeColl = New Collection
   
   Set m_5 = New Collection
   Set m_6 = New Collection
   Set m_7 = New Collection
   Set m_8 = New Collection
   Set m_9 = New Collection
   Set m_10 = New Collection
   Set m_11 = New Collection
   
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
   Set m_ApSupAmountTo = Nothing
   Set m_ApSupPaidTo = Nothing
   Set m_ApSupAmountCns = Nothing
   Set m_ApSupPaidCns = Nothing
   Set m_PaidAllAmountCns = Nothing
   Set m_ApSupAmountToCn = Nothing
   Set m_ApSupPaidToCn = Nothing
   Set m_ApSupAmountCashs = Nothing
   
   Set SupplierGroupColl = Nothing
   Set GroupTypeColl = Nothing
   
   Set m_5 = Nothing
   Set m_6 = Nothing
   Set m_7 = Nothing
   Set m_8 = Nothing
   Set m_9 = Nothing
   Set m_10 = Nothing
   Set m_11 = Nothing
   
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
      
   Call Cf.SetFieldValue(1, "^", "N", "^")
   Call m_Headers0.AddField(Cf)
   Call Cf.SetFieldValue(4, "^", "�����������", "<")
   Call m_Headers0.AddField(Cf)
   Call Cf.SetFieldValue(14, "^", "��������´����", "<")
   Call m_Headers0.AddField(Cf)
   Call Cf.SetFieldValue(8, "^", "����", "<")
   Call m_Headers0.AddField(Cf)
   Call Cf.SetFieldValue(4, "^", "POST", "<")
   Call m_Headers0.AddField(Cf)
   '===
   Call Cf.SetFieldValue(1, "^", "NO.", "^")
   Call m_Headers1.AddField(Cf)
   Call Cf.SetFieldValue(4, "^", "���ͼ����", "<")
   Call m_Headers1.AddField(Cf)
   Call Cf.SetFieldValue(2, "^", "�ʹ˹��¡��", ">")
   Call m_Headers1.AddField(Cf)
   Call Cf.SetFieldValue(2, "^", "����/����˹��", ">")
   Call m_Headers1.AddField(Cf)
   Call Cf.SetFieldValue(2, "^", "����/Ŵ˹��", ">")
   Call m_Headers1.AddField(Cf)
   Call Cf.SetFieldValue(2, "^", "�ʹ˹��¡�", ">")
   Call m_Headers1.AddField(Cf)
   Call Cf.SetFieldValue(2, "^", "�ʹ���¨�ԧ DUE㹪�ǧ", ">")
   Call m_Headers1.AddField(Cf)
   Call Cf.SetFieldValue(2, "^", "�ʹ���� POST DUE��ѧ��ǧ", ">")
   Call m_Headers1.AddField(Cf)
   Call Cf.SetFieldValue(2, "^", "˹�餧�����", ">")
   Call m_Headers1.AddField(Cf)
   Call Cf.SetFieldValue(2, "^", "DUE 㹪�ǧ", ">")
   Call m_Headers1.AddField(Cf)
   Call Cf.SetFieldValue(2, "^", "�Թʴ 㹪�ǧ", ">")
   Call m_Headers1.AddField(Cf)
   Call Cf.SetFieldValue(2, "^", "�͡���¨��� 㹪�ǧ", ">")
   Call m_Headers1.AddField(Cf)
   Call Cf.SetFieldValue(2, "^", "DUE ��ѧ��ǧ", ">")
   Call m_Headers1.AddField(Cf)
   Call Cf.SetFieldValue(2, "^", "���¡�͹��ǧDUE㹪�ǧ", ">")
   Call m_Headers1.AddField(Cf)
   Call Cf.SetFieldValue(2, "^", "���¡�͹��ǧDUE��ѧ��ǧ", ">")
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
'Dim ToPaidDate As String
   
   'Call GetFirstLastDate(DateSerial(mcolParam("YEAR_NO") - 543, mcolParam("MONTH_ID"), 1), FirstDate, LastDate)
   
   FirstDate = mcolParam("FROM_DOC_DATE")
   LastDate = mcolParam("TO_DOC_DATE")
   
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
   
   Call LoadSupplierType(Nothing, m_SupplierTypes)
   
   'Call LoadPaidAmountByBill(Nothing, m_PaidAmounts, -1, mcolParam("TO_PAY_DATE"))
   'Call LoadPaidAmountByBill(Nothing, m_PaidAllAmounts, -1, -1)
   'Call LoadDueDateInterval1(Nothing, m_MaxMins)
   
   '-------------------------------- ¡�� --------------------------------
   Call LoadAPAmountBySup(Nothing, m_ApSupAmounts, -1, DateAdd("D", -1, FirstDate), "('3','4')")    '�ʹ���� �֧�ѹ���
   Call LoadAPAmountBySup(Nothing, m_ApSupAmountCns, -1, DateAdd("D", -1, FirstDate), "('5')")     '�ʹ���� �֧�ѹ���       Ŵ˹��
   
   Call LoadAPAmountBySup(Nothing, m_ApSupAmountCashs, FirstDate, LastDate, "('9')")    '�ʹ�����Թʴ �֧�ѹ���
   
   Call LoadPaidAmountBySup(Nothing, m_ApSupPaids, -1, DateAdd("D", -1, FirstDate), -1, -1, "('3','4')")           '�ʹ���¶֧�ѹ���
   Call LoadPaidAmountBySup(Nothing, m_ApSupPaidCns, -1, DateAdd("D", -1, FirstDate), -1, -1, "('5')")           '�ʹ���¶֧�ѹ���    Ŵ˹��
   
   Call LoadPaidAmountBySup(Nothing, m_PaidAllAmounts, -1, -1, -1, DateAdd("D", -1, FirstDate), "('3','4')")         '�ʹ�Ѵ��ŷ������ͧ��������ѹ����˹�
   Call LoadPaidAmountBySup(Nothing, m_PaidAllAmountCns, -1, -1, -1, DateAdd("D", -1, FirstDate), "('5')")          '�ʹ�Ѵ��ŷ������ͧ��������ѹ����˹� Ŵ˹��
   '-------------------------------- ¡�� --------------------------------
   
   '-------------------------------- ����͹��� --------------------------------
   Call LoadAPAmountBySup(Nothing, m_ApSupAmountTo, FirstDate, LastDate, "('3','4')")     '�ʹ���� �֧�ѹ���
   Call LoadAPAmountBySup(Nothing, m_ApSupAmountToCn, FirstDate, LastDate, "('5')")     'Ŵ˹�� �֧�ѹ���
   
   Call LoadPaidAmountBySup(Nothing, m_ApSupPaidTo, FirstDate, LastDate, -1, -1, "('3','4')")           '�ʹ���¶֧�ѹ���
   Call LoadPaidAmountBySup(Nothing, m_ApSupPaidToCn, FirstDate, LastDate, -1, -1, "('5')")           '�ʹ���¶֧�ѹ���
   '-------------------------------- ����͹��� --------------------------------
   
   '-------------------------------- ������ --------------------------------
   Call LoadBktChqAmountBySup(Nothing, m_5, , , FirstDate, LastDate)                                                                                      '�ʹ���¨�ԧ DUE ��͹4
   Call LoadBktChqAmountBySup(Nothing, m_6, , LastDate, DateAdd("D", 1, LastDate))                                                                        '�ʹ���� POST DUE ��ѧ��͹4
   
   Call LoadBktChqAmountBySup(Nothing, m_8, FirstDate, LastDate, FirstDate, LastDate)                                                       '���� 4 DUE 4
   Call LoadBktChqAmountBySup(Nothing, m_9, FirstDate, LastDate, DateAdd("D", 1, LastDate))                                        '���� 4 DUE��ѧ 4
   
   Call LoadBktChqAmountBySup(Nothing, m_10, -1, DateAdd("D", -1, FirstDate), FirstDate, LastDate)                              '���¡�͹ 4 DUE 4
   Call LoadBktChqAmountBySup(Nothing, m_11, -1, DateAdd("D", -1, FirstDate), DateAdd("D", 1, LastDate), -1)                             '���¡�͹ 4 DUE��ѧ 4
   '-------------------------------- ������ --------------------------------
   
   Call LoadSupplierGroup(Nothing, SupplierGroupColl, , , mcolParam("DATA_TYPE_ID"))
   Call LoadGroupType(Nothing, GroupTypeColl)
   
   'Call LoadDbnCdnByBill(Nothing, m_DbnCdnByBills, -1, mcolParam("TO_PAY_DATE"))
   FromDate = "�ҡ�ѹ��� " & EmptyToString(DateToStringExtEx2(mcolParam("FROM_DOC_DATE")), "N/A")
   ToDate = " �֧ �ѹ��� " & EmptyToString(DateToStringExtEx2(mcolParam("TO_DOC_DATE")), "N/A")
'   ToPaidDate = "    �ѹ�������Թ�֧�ѹ��� " & EmptyToString(DateToStringExtEx2(mcolParam("TO_PAY_DATE")), "N/A")
      
   'glbEnterPrise.ENTERPRISE_NAME & " , " & glbEnterPrise.BRANCH_NAME
   MODULE_DESC = glbCompanyName & vbCrLf & mcolParam("REPORT_NAME") & _
                                       vbCrLf & FromDate & "   " & ToDate & " " & mcolParam("DATA_TYPE_NAME")
   
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
Dim St As CIsTab
Dim Ap As CAPMas
Dim Width As Double
Dim Apt As CApTrn
Dim TempRs As ADODB.Recordset
Dim TempApt As CApTrn
Dim TempAptCn As CApTrn
Dim TempAptCash As CApTrn

Dim TempApRcIt As CAPRcIt
Dim TempApRcItCn As CAPRcIt

Dim TempAptTo As CApTrn
Dim TempAptToCn As CApTrn

Dim TempApRcItTo As CAPRcIt
Dim TempApRcItToCn As CAPRcIt

Dim TempApBktTrn As CBkTrn

Dim TempApRcIt3 As CAPRcIt
Dim TempApRcIt3Cn As CAPRcIt

Dim BalanceAmount As Double

Dim TempPaidCheck As Double
Dim TempPaidAmount  As Double

Dim SupGroup As CSupplierGroup
Dim GroupType As CGroupType
Dim GroupTypeItem As CColumnValue
Dim TempColumnValue As CColumnValue
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
   
      Set BodyCf = New CReportField
      i = 0
      For Each St In m_SupplierTypes
         '===Head
         Call m_Details1.ClearField
         Width = m_Headers0.Fields(1).Width
         TempStr = ""
         Call BodyCf.SetFieldValue(Width, "^", TempStr)
         Call m_Details1.AddField(BodyCf)
         
         Width = m_Headers0.Fields(2).Width
         TempStr = St.TYPCOD
         Call BodyCf.SetFieldValue(Width, "^", TempStr)
         Call m_Details1.AddField(BodyCf)

         Width = m_Headers0.Fields(3).Width
         TempStr = St.TYPDES
         Call BodyCf.SetFieldValue(Width, "<", TempStr)
         Call m_Details1.AddField(BodyCf)
         
         Width = m_Headers0.Fields(4).Width
         TempStr = ""
         Call BodyCf.SetFieldValue(Width, "<", TempStr)
         Call m_Details1.AddField(BodyCf)
         
         Width = m_Headers0.Fields(5).Width
         TempStr = ""
         Call BodyCf.SetFieldValue(Width, "<", TempStr)
         Call m_Details1.AddField(BodyCf)
         
         '===

         Set BodyCf = New CReportField
   
         Set Ap = New CAPMas
         Ap.FROM_SUPPLIER_CODE = mcolParam("FROM_SUPPLIER_CODE")
         Ap.TO_SUPPLIER_CODE = mcolParam("TO_SUPPLIER_CODE")
         Ap.SUPTYP = St.TYPCOD
         Call glbDaily.QuerySupplier(Ap, Rs, iCount, IsOK, glbErrorLog)
         
         If iCount > 0 Then
            Call m_Details1.GetString(1, TempStr1, TempStr2)
            strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
            Call Vsp.AddTable(strFormat, "", TempStr2)
         End If
         
         HaveValueSum = False
         
         While Not Rs.EOF
            i = i + 1
            Call m_Details1.ClearField
            Call Ap.PopulateFromRS(1, Rs)
            Set TempApt = GetAPTrn(m_ApSupAmounts, Ap.SUPCOD)
            Set TempAptCn = GetAPTrn(m_ApSupAmountCns, Ap.SUPCOD)
            
            Set TempApRcIt = GetAPRcpItem(m_ApSupPaids, Ap.SUPCOD)
            Set TempApRcItCn = GetAPRcpItem(m_ApSupPaidCns, Ap.SUPCOD)
               
            Set TempApRcIt3 = GetAPRcpItemEx(m_PaidAllAmounts, Ap.SUPCOD)
            Set TempApRcIt3Cn = GetAPRcpItemEx(m_PaidAllAmountCns, Ap.SUPCOD)
               
            Set SupGroup = GetObject("CSupplierGroup", SupplierGroupColl, Trim(Ap.SUPCOD), False)
            If Not (SupGroup Is Nothing) Then
               Set GroupType = GetObject("CGroupType", GroupTypeColl, Trim(SupGroup.GROUP_TYPE_CODE))
            Else
               Set GroupType = GetObject("CGroupType", GroupTypeColl, "99-9999")
            End If
            
            HaveValue = False
            
            '===Head
            Set HeadCf = m_Headers1.Fields(1)
            TempStr = i
            Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
            Call m_Details1.AddField(BodyCf)
            
            Set HeadCf = m_Headers1.Fields(2)
            TempStr = Left(Ap.SUPNAM, 30)
            Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
            Call m_Details1.AddField(BodyCf)
            
            BalanceAmount = (TempApt.AMOUNT - TempAptCn.AMOUNT) - (TempApt.PAYAMT - TempAptCn.PAYAMT) - (TempApRcIt.PAYAMT - TempApRcItCn.PAYAMT) + (TempApRcIt3.PAYAMT - TempApRcIt3Cn.PAYAMT)
            
            K = 2
            K = K + 1
            Set HeadCf = m_Headers1.Fields(K)
            TempStr = FormatNumber(BalanceAmount, "")
            Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
            Call m_Details1.AddField(BodyCf)
            
            Total1(K) = Total1(K) + BalanceAmount
            Total2(K) = Total2(K) + BalanceAmount
            If Round((BalanceAmount), 2) <> 0 Then
               HaveValue = True
            End If
            Call SumValue(GroupType.ItemAmountColl, Trim(Str(K)), BalanceAmount)
            
            Set TempAptTo = GetAPTrn(m_ApSupAmountTo, Ap.SUPCOD)
            Set TempAptToCn = GetAPTrn(m_ApSupAmountToCn, Ap.SUPCOD)
            
            Set TempApRcItTo = GetAPRcpItem(m_ApSupPaidTo, Ap.SUPCOD)
            Set TempApRcItToCn = GetAPRcpItem(m_ApSupPaidToCn, Ap.SUPCOD)
                     
            K = K + 1
            Set HeadCf = m_Headers1.Fields(K)
            TempStr = FormatNumber(TempAptTo.AMOUNT, "")
            'TempStr = FormatNumber(TempAptTo.AMOUNT - TempAptToCn.AMOUNT, "")
            Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
            Call m_Details1.AddField(BodyCf)
            
'            Total1(K) = Total1(K) + TempAptTo.AMOUNT - TempAptToCn.AMOUNT
'            Total2(K) = Total2(K) + TempAptTo.AMOUNT - TempAptToCn.AMOUNT
            Total1(K) = Total1(K) + TempAptTo.AMOUNT
            Total2(K) = Total2(K) + TempAptTo.AMOUNT
'            If Round((TempAptTo.AMOUNT - TempAptToCn.AMOUNT), 2) <> 0 Then
'               HaveValue = True
'            End If
            If Round((TempAptTo.AMOUNT), 2) <> 0 Then
               HaveValue = True
            End If
            Call SumValue(GroupType.ItemAmountColl, Trim(Str(K)), TempAptTo.AMOUNT)
            
            K = K + 1
            Set HeadCf = m_Headers1.Fields(K)
            TempStr = FormatNumber(TempApRcItTo.PAYAMT - TempApRcItToCn.PAYAMT + TempAptToCn.AMOUNT, "")
            Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
            Call m_Details1.AddField(BodyCf)
            
            Total1(K) = Total1(K) + TempApRcItTo.PAYAMT - TempApRcItToCn.PAYAMT + TempAptToCn.AMOUNT
            Total2(K) = Total2(K) + TempApRcItTo.PAYAMT - TempApRcItToCn.PAYAMT + TempAptToCn.AMOUNT
            If Round((TempApRcItTo.PAYAMT - TempApRcItToCn.PAYAMT + TempAptToCn.AMOUNT), 2) <> 0 Then
               HaveValue = True
            End If
            TempPaidCheck = Round((TempApRcItTo.PAYAMT - TempApRcItToCn.PAYAMT + TempAptToCn.AMOUNT), 2)
            Call SumValue(GroupType.ItemAmountColl, Trim(Str(K)), TempApRcItTo.PAYAMT - TempApRcItToCn.PAYAMT + TempAptToCn.AMOUNT)
            
            K = K + 1
            Set HeadCf = m_Headers1.Fields(K)
            TempStr = FormatNumber(BalanceAmount + (TempAptTo.AMOUNT - TempAptToCn.AMOUNT) - (TempApRcItTo.PAYAMT - TempApRcItToCn.PAYAMT), "")
            Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
            Call m_Details1.AddField(BodyCf)
            
            Total1(K) = Total1(K) + (BalanceAmount) + (TempAptTo.AMOUNT - TempAptToCn.AMOUNT) - (TempApRcItTo.PAYAMT - TempApRcItToCn.PAYAMT)
            Total2(K) = Total2(K) + (BalanceAmount) + (TempAptTo.AMOUNT - TempAptToCn.AMOUNT) - (TempApRcItTo.PAYAMT - TempApRcItToCn.PAYAMT)
            If Round(((BalanceAmount) + (TempAptTo.AMOUNT - TempAptToCn.AMOUNT) - (TempApRcItTo.PAYAMT - TempApRcItToCn.PAYAMT)), 2) <> 0 Then
               HaveValue = True
            End If
            Call SumValue(GroupType.ItemAmountColl, Trim(Str(K)), (BalanceAmount) + (TempAptTo.AMOUNT - TempAptToCn.AMOUNT) - (TempApRcItTo.PAYAMT - TempApRcItToCn.PAYAMT))
            
            Set TempApBktTrn = GetBkTrn(m_5, Ap.SUPCOD)
            Set TempAptCash = GetAPTrn(m_ApSupAmountCashs, Ap.SUPCOD)
            
            K = K + 1
            Set HeadCf = m_Headers1.Fields(K)
            TempStr = FormatNumber(TempApBktTrn.AMOUNT + TempAptCash.CSHPAY, "")
            Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
            Call m_Details1.AddField(BodyCf)
            
            Total1(K) = Total1(K) + TempApBktTrn.AMOUNT + TempAptCash.CSHPAY
            Total2(K) = Total2(K) + TempApBktTrn.AMOUNT + TempAptCash.CSHPAY
            If Round((TempApBktTrn.AMOUNT + TempAptCash.CSHPAY), 2) <> 0 Then
               HaveValue = True
            End If
            Call SumValue(GroupType.ItemAmountColl, Trim(Str(K)), TempApBktTrn.AMOUNT + TempAptCash.CSHPAY)
            
            Set TempApBktTrn = GetBkTrn(m_6, Ap.SUPCOD)
            K = K + 1
            Set HeadCf = m_Headers1.Fields(K)
            TempStr = FormatNumber(TempApBktTrn.AMOUNT, "")
            Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
            Call m_Details1.AddField(BodyCf)
            
            Total1(K) = Total1(K) + TempApBktTrn.AMOUNT
            Total2(K) = Total2(K) + TempApBktTrn.AMOUNT
            If Round((TempApBktTrn.AMOUNT), 2) <> 0 Then
               HaveValue = True
            End If
            Call SumValue(GroupType.ItemAmountColl, Trim(Str(K)), TempApBktTrn.AMOUNT)
            
            '7 =  4 + 6
            K = K + 1
            Set HeadCf = m_Headers1.Fields(K)
            TempStr = FormatNumber(BalanceAmount + (TempAptTo.AMOUNT - TempAptToCn.AMOUNT) - (TempApRcItTo.PAYAMT - TempApRcItToCn.PAYAMT) + TempApBktTrn.AMOUNT, "")
            Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
            Call m_Details1.AddField(BodyCf)
            
            Total1(K) = Total1(K) + BalanceAmount + (TempAptTo.AMOUNT - TempAptToCn.AMOUNT) - (TempApRcItTo.PAYAMT - TempApRcItToCn.PAYAMT) + TempApBktTrn.AMOUNT
            Total2(K) = Total2(K) + BalanceAmount + (TempAptTo.AMOUNT - TempAptToCn.AMOUNT) - (TempApRcItTo.PAYAMT - TempApRcItToCn.PAYAMT) + TempApBktTrn.AMOUNT
            If Round((BalanceAmount + (TempAptTo.AMOUNT - TempAptToCn.AMOUNT) - (TempApRcItTo.PAYAMT - TempApRcItToCn.PAYAMT) + TempApBktTrn.AMOUNT), 2) <> 0 Then
               HaveValue = True
            End If
            Call SumValue(GroupType.ItemAmountColl, Trim(Str(K)), BalanceAmount + (TempAptTo.AMOUNT - TempAptToCn.AMOUNT) - (TempApRcItTo.PAYAMT - TempApRcItToCn.PAYAMT) + TempApBktTrn.AMOUNT)
            
           '8
            Set TempApBktTrn = GetBkTrn(m_8, Ap.SUPCOD)
            K = K + 1
            Set HeadCf = m_Headers1.Fields(K)
            TempStr = FormatNumber(TempApBktTrn.AMOUNT, "")
            Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
            Call m_Details1.AddField(BodyCf)
            
            Total1(K) = Total1(K) + TempApBktTrn.AMOUNT
            Total2(K) = Total2(K) + TempApBktTrn.AMOUNT
            If Round((TempApBktTrn.AMOUNT), 2) <> 0 Then
               HaveValue = True
            End If
            TempPaidAmount = 0
            TempPaidAmount = TempPaidAmount + TempApBktTrn.AMOUNT
            Call SumValue(GroupType.ItemAmountColl, Trim(Str(K)), TempApBktTrn.AMOUNT)
            
            '9
            K = K + 1
            Set HeadCf = m_Headers1.Fields(K)
            TempStr = FormatNumber(TempAptCash.CSHPAY, "")
            Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
            Call m_Details1.AddField(BodyCf)
            
            Total1(K) = Total1(K) + TempAptCash.CSHPAY
            Total2(K) = Total2(K) + TempAptCash.CSHPAY
            If Round((TempAptCash.CSHPAY), 2) <> 0 Then
               HaveValue = True
            End If
            TempPaidAmount = TempPaidAmount + TempAptCash.CSHPAY
            Call SumValue(GroupType.ItemAmountColl, Trim(Str(K)), TempAptCash.CSHPAY)
            
            '----------------------------------------------------------------------------------------
            K = K + 1
            Set HeadCf = m_Headers1.Fields(K)
            TempStr = FormatNumber(TempAptCash.INTPAY, "")
            Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
            Call m_Details1.AddField(BodyCf)
            
            Total1(K) = Total1(K) + TempAptCash.INTPAY
            Total2(K) = Total2(K) + TempAptCash.INTPAY
            If Round((TempAptCash.INTPAY), 2) <> 0 Then
               HaveValue = True
            End If
            Call SumValue(GroupType.ItemAmountColl, Trim(Str(K)), TempAptCash.INTPAY)
            '-------------------------------------------------------------------------------------
            
            '10
            Set TempApBktTrn = GetBkTrn(m_9, Ap.SUPCOD)
            TempPaidAmount = TempPaidAmount + TempApBktTrn.AMOUNT
            K = K + 1
            Set HeadCf = m_Headers1.Fields(K)
            'If Round(TempPaidCheck, 2) = Round(TempPaidAmount, 2) Then
               TempStr = FormatNumber(TempApBktTrn.AMOUNT, "")
               Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
               Call m_Details1.AddField(BodyCf)
               
               Total1(K) = Total1(K) + TempApBktTrn.AMOUNT
               Total2(K) = Total2(K) + TempApBktTrn.AMOUNT
               If Round((TempApBktTrn.AMOUNT), 2) <> 0 Then
                  HaveValue = True
               End If
               Call SumValue(GroupType.ItemAmountColl, Trim(Str(K)), TempApBktTrn.AMOUNT)
'            Else
'               TempStr = FormatNumber(TempApBktTrn.AMOUNT - (TempPaidAmount - TempPaidCheck), "")
'               Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
'               Call m_Details1.AddField(BodyCf)
'
'               Total1(K) = Total1(K) + TempApBktTrn.AMOUNT - (TempPaidAmount - TempPaidCheck)
'               Total2(K) = Total2(K) + TempApBktTrn.AMOUNT - (TempPaidAmount - TempPaidCheck)
'               If Round((TempApBktTrn.AMOUNT - (TempPaidAmount - TempPaidCheck)), 2) <> 0 Then
'                  HaveValue = True
'               End If
'            End If
            '11
            Set TempApBktTrn = GetBkTrn(m_10, Ap.SUPCOD)
            K = K + 1
            Set HeadCf = m_Headers1.Fields(K)
            TempStr = FormatNumber(TempApBktTrn.AMOUNT, "")
            Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
            Call m_Details1.AddField(BodyCf)
            
            Total1(K) = Total1(K) + TempApBktTrn.AMOUNT
            Total2(K) = Total2(K) + TempApBktTrn.AMOUNT
            If Round((TempApBktTrn.AMOUNT), 2) <> 0 Then
               HaveValue = True
            End If
            Call SumValue(GroupType.ItemAmountColl, Trim(Str(K)), TempApBktTrn.AMOUNT)
            
            '12
            Set TempApBktTrn = GetBkTrn(m_11, Ap.SUPCOD)
            K = K + 1
            Set HeadCf = m_Headers1.Fields(K)
            TempStr = FormatNumber(TempApBktTrn.AMOUNT, "")
            Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
            Call m_Details1.AddField(BodyCf)
            
            Total1(K) = Total1(K) + TempApBktTrn.AMOUNT
            Total2(K) = Total2(K) + TempApBktTrn.AMOUNT
            If Round((TempApBktTrn.AMOUNT), 2) <> 0 Then
               HaveValue = True
            End If
            Call SumValue(GroupType.ItemAmountColl, Trim(Str(K)), TempApBktTrn.AMOUNT)
            
            If HaveValue Then
               Call m_Details1.GetString(1, TempStr1, TempStr2)
               strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
               Call Vsp.AddTable(strFormat, "", TempStr2)
               
               HaveValueSum = True
            Else
               i = i - 1
            End If
            '===
   
            Rs.MoveNext
         Wend
         
         If HaveValueSum Then
            Call GenerateFooter("���", "", "", Total1, m_Details1)
            Call m_Details1.GetString(2, TempStr1, TempStr2)
            strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
            Call Vsp.AddTable(strFormat, "", TempStr2)
         End If
         For j = 1 To UBound(Total1)
            Total1(j) = 0
         Next j
         
         Set Ap = Nothing
      Next St
      Set BodyCf = Nothing

      Call GenerateFooter("�ط��", "", "", Total2, m_Details1)
      Call m_Details1.GetString(2, TempStr1, TempStr2)
      strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
      Vsp.FontBold = True
      Call Vsp.AddTable(strFormat, "", TempStr2)
      Vsp.FontBold = False
            
      Vsp = ""
      Vsp = ""
      
      For Each GroupType In GroupTypeColl
         Call GenerateFooterSummary("", GroupType.GROUP_TYPE_NAME, GroupType.ItemAmountColl, m_Details1)
         Call m_Details1.GetString(2, TempStr1, TempStr2)
         strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
         Vsp.FontBold = True
         Call Vsp.AddTable(strFormat, "", TempStr2)
         Vsp.FontBold = False
      Next GroupType
      
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

Private Sub GenerateFooter(Txt1 As String, Txt2 As String, Txt3 As String, Tot() As Double, Details As CFieldList)
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
      ElseIf j > 2 Then
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
Private Sub SumValue(Coll As Collection, KEY As String, Value As Double)
Dim GroupTypeItem As CColumnValue
Dim TempColumnValue  As CColumnValue
   
   Set GroupTypeItem = GetObject("CColumnValue", Coll, KEY, False)
   If GroupTypeItem Is Nothing Then
      Set TempColumnValue = New CColumnValue
      TempColumnValue.KEY = KEY
      TempColumnValue.Value = Value
      Call Coll.Add(TempColumnValue, KEY)
   Else
      GroupTypeItem.Value = GroupTypeItem.Value + Value
   End If
End Sub
Private Sub GenerateFooterSummary(Txt1 As String, Txt2 As String, Coll As Collection, Details As CFieldList)
Dim HeadCf As CReportField
Dim BodyCf As CReportField
Dim j As Long
Dim TempStr As String
Dim ColumnValue As CColumnValue
   
   
   Call Details.ClearField
   Set BodyCf = New CReportField
   
   j = 0
   For Each HeadCf In m_Headers1.Fields
      j = j + 1
      If j = 1 Then
         TempStr = Txt1
      ElseIf j = 2 Then
         TempStr = Txt2
      ElseIf j > 2 Then
         Set ColumnValue = GetObject("CColumnValue", Coll, Trim(Str(j)), False)
         If ColumnValue Is Nothing Then
            TempStr = FormatNumber(0)
         Else
            TempStr = FormatNumber(ColumnValue.Value)
         End If
      End If
      
      Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.Align, TempStr, HeadCf.BodyAlign)
      Call Details.AddField(BodyCf)
   Next HeadCf
      
   Set BodyCf = Nothing
End Sub
