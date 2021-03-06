VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CReportCom07"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'User Group Report
Implements CReportInterface

Private Const MODULE_NAME = "CReportCOM07"
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
Private m_StcrdBySale As Collection
Private m_AreaCod As Collection
Private m_AreaFromIV As Collection
Private m_ConditionCommiss1 As Collection
Private m_ConditionCommiss2 As Collection
Private m_ConditionCommiss3 As Collection
Private m_ConditionCommiss4 As Collection

'Private m_cusFromArea(20) As Collection
Private nonidentifyArea As Boolean
Dim Stb As CCommissMasterArea
Private m_MinusStkcod As Collection
Private m_IVincomplete0 As Collection
Private m_IVincomplete12 As Collection
Private m_REsumIV As Collection
Private m_ReDocdat As Collection
Private temp_Area As CCommissionCustomerArea
Private temp_Area2 As CCommissionCustomerArea
Public REsumIV As CARRcIt
Private tempComDonStkcod As CComDonStk
Private m_ComDonStkcod As Collection
Dim L As Long
  Dim FromSaleID As String
  Dim ToSaleID As String

'Private havevalue As Boolean

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
   Set m_StcrdBySale = New Collection
   Set m_AreaCod = New Collection
   Set m_AreaFromIV = New Collection
   Set m_ConditionCommiss1 = New Collection
   Set m_ConditionCommiss2 = New Collection
   Set m_ConditionCommiss3 = New Collection
   Set m_ConditionCommiss4 = New Collection
   Set Stb = New CCommissMasterArea
   Set m_MinusStkcod = New Collection
   Set m_IVincomplete0 = New Collection
   Set m_IVincomplete12 = New Collection
   Set m_REsumIV = New Collection
   Set m_ReDocdat = New Collection
   Set temp_Area = New CCommissionCustomerArea
   Set temp_Area2 = New CCommissionCustomerArea
   Set REsumIV = New CARRcIt
      Set m_ComDonStkcod = New Collection
   Set tempComDonStkcod = New CComDonStk
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
   Set m_StcrdBySale = Nothing
   Set m_AreaCod = Nothing
   Set m_AreaFromIV = Nothing
   Set m_ConditionCommiss1 = Nothing
   Set m_ConditionCommiss2 = Nothing
   Set m_ConditionCommiss3 = Nothing
   Set m_ConditionCommiss4 = Nothing
   Set Stb = Nothing
   Set m_MinusStkcod = Nothing
   Set m_IVincomplete0 = Nothing
   Set m_IVincomplete12 = Nothing
   Set m_REsumIV = Nothing
   Set m_ReDocdat = Nothing
   Set temp_Area = Nothing
   Set temp_Area2 = Nothing
   Set REsumIV = Nothing
      Set m_ComDonStkcod = Nothing
   Set tempComDonStkcod = Nothing
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
'Dim TempRs As ADODB.Recordset
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
      
   Call Cf.SetFieldValue(1, "^", "", "^")
   Call m_Headers0.AddField(Cf)
   Call Cf.SetFieldValue(22, "^", "", "<")    '27
   Call m_Headers0.AddField(Cf)
      
      
If mcolParam("SUMMARY_MODE") = "N" Then
      Call Cf.SetFieldValue(1, "^", "NO.", "^")
      Call m_Headers1.AddField(Cf)
      
      Call Cf.SetFieldValue(2, "^", "INVOICE", "^")
      Call m_Headers1.AddField(Cf)
     
      Call Cf.SetFieldValue(2, "^", "�����١���", "^")
      Call m_Headers1.AddField(Cf)
      
      Call Cf.SetFieldValue(5, "^", "�����١���", "<")
      Call m_Headers1.AddField(Cf)
      
      Call Cf.SetFieldValue(5, "^", "�����Թ���", "<")
      Call m_Headers1.AddField(Cf)
Else
      Call Cf.SetFieldValue(0, "^", "NO.", "^")
      Call m_Headers1.AddField(Cf)
      
      Call Cf.SetFieldValue(0, "^", "�����١���", "^")
      Call m_Headers1.AddField(Cf)
      
      Call Cf.SetFieldValue(2, "^", "����", "<")
      Call m_Headers1.AddField(Cf)
      
      Call Cf.SetFieldValue(6, "^", "��ѡ�ҹ���", "<")
      Call m_Headers1.AddField(Cf)
      
      Call Cf.SetFieldValue(0, "^", "�����Թ���", "<")
      Call m_Headers1.AddField(Cf)
End If

   Call Cf.SetFieldValue(2, "^", "�ʹ���", ">")
   Call m_Headers1.AddField(Cf)
      
 If mcolParam("SUMMARY_MODE") = "N" Then
   Call Cf.SetFieldValue(2, "^", "�ѹ���Թ���", "^")
   Call m_Headers1.AddField(Cf)
   
   Call Cf.SetFieldValue(2, "^", "�ѹ�Ѻ����", "^")
   Call m_Headers1.AddField(Cf)
   
   Call Cf.SetFieldValue(2, "^", "�Ţ��������", "^")
   Call m_Headers1.AddField(Cf)
 Else
    Call Cf.SetFieldValue(0, "^", "�ѹ���Թ���", "^")
   Call m_Headers1.AddField(Cf)
 
    Call Cf.SetFieldValue(0, "^", "�ѹ����Ѻ����", "^")
   Call m_Headers1.AddField(Cf)
   
    Call Cf.SetFieldValue(0, "^", "�Ţ��������", "^")
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
Dim FromDate As String
Dim ToDate As String
Dim ToPaidDate As String
Dim YEAR_ID As Long

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
   Vsp.MarginRight = 500
   Vsp.MarginTop = 300
   Vsp.FontName = "AngsanaUPC"
   Vsp.FontSize = DETAIL_SIZE

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

   MODULE_DESC = glbCompanyName & vbCrLf & mcolParam("REPORT_NAME") & _
                             vbCrLf & "�ѹ������Թ���   " & EmptyToString(DateToStringExtEx2(mcolParam("FROM_DOC_DATE")), "N/A") & "    �֧    " & EmptyToString(DateToStringExtEx2(mcolParam("TO_DOC_DATE")), "N/A") & "      �ѹ����Ѻ����    " & EmptyToString(DateToStringExtEx2(mcolParam("FROM_CMPL_DATE")), "N/A") & "    �֧    " & EmptyToString(DateToStringExtEx2(mcolParam("TO_CMPL_DATE")), "N/A")
      If Len(mcolParam("FROM_SALE_NAME")) > 0 Then
              MODULE_DESC = MODULE_DESC & vbCrLf & "    ��ѡ�ҹ���  " & EmptyToString(FromSaleName, "N/A")
      End If
       If Len(mcolParam("TO_SALE_NAME")) > 0 Then
              MODULE_DESC = MODULE_DESC & "   �֧    " & EmptyToString(ToSaleName, "N/A")
      End If
      
      Call LoadMinusIV(m_MinusStkcod, mcolParam("FROM_DOC_DATE"), mcolParam("TO_DOC_DATE"))
      Call LoadComDonStk(m_ComDonStkcod, mcolParam("FROM_DOC_DATE"), mcolParam("TO_DOC_DATE"))
      Call LoadREDocDat(m_ReDocdat, mcolParam("FROM_CMPL_DATE"), mcolParam("TO_CMPL_DATE")) ' ��ͧ��������
      Call LoadREsumIV(Nothing, m_REsumIV)

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
Dim PrevKey3 As String  '����ࢵ
Dim PrevKey4 As String    'ࢵ��â��
Dim St As CCommissMasterArea
Dim ArS As COESLM
Dim Stcrd As CStcrd

Dim haveValueSale As Boolean
Dim haveData As Boolean
Dim haveData_ByArea As Boolean
Dim DueCount As Long

Dim TempCConditionCommiss As CConditionCommission
Dim m_runConditionCommiss As CConditionCommission
Dim PercentNum1 As Double
Dim PercentSum As Double
Dim tempCusArea As CCommissionCustomerArea
Dim toCMPLdate As Date
Dim tempMinusStkcod As CComMinusStk
Dim tempREdoc As CARTrn
Dim NETVAL As Double
Dim strNETVAL As String
Dim PayCompleted  As Boolean
Dim CMPLDAT As Date
Dim RECOD As String
Dim temnetval As Long
Dim CorrectStkcod As Boolean
Dim fixedRE As Boolean
Dim showfromDate As Boolean
   
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
   i = 1  '�Ѻ�.��÷Ѵ����������
   Set BodyCf = New CReportField
       
      Set Stcrd = New CStcrd
      Stcrd.FROM_DOC_DATE = mcolParam("FROM_DOC_DATE")
      Stcrd.TO_DOC_DATE = mcolParam("TO_DOC_DATE")
      If Len(mcolParam("FROM_SALE_NAME")) > 0 Then
         Stcrd.FROM_SALE_CODE = FromSaleID
         Stcrd.TO_SALE_CODE = ToSaleID
      End If
      Call Stcrd.QueryData(17, Rs, iCount)

      While Not Rs.EOF


      Call Stcrd.PopulateFromRS(17, Rs)
      ' �ӹǳ��ǹŴ��͹
      Set tempMinusStkcod = GetMinusCommiss(m_MinusStkcod, Stcrd.DOCDAT & "-" & Stcrd.DOCNUM & "-" & Stcrd.STKCOD, False)
      If (tempMinusStkcod Is Nothing) Then   '�����㹤���礪�� = �������ǹŴ
               NETVAL = Stcrd.NETVAL
               strNETVAL = FormatNumber(NETVAL)
      Else:
               NETVAL = (Stcrd.NETVAL + Val(tempMinusStkcod.MINUS_AMOUNT))
               strNETVAL = "*" & FormatNumber(NETVAL)
      End If
      
       fixedRE = False
      temnetval = temnetval + NETVAL
      
'      If Stcrd.DOCNUM = "IV0042847" Then
'         'debug.print
'      End If
      
      Set REsumIV = GetARRcpItem(m_REsumIV, Stcrd.DOCNUM, False)                     ' /*      ' <----------- �����鹵���� �����ҡ�ٷء IV
      If Not (REsumIV Is Nothing) Then
        If ((REsumIV.RCVAMT - NETVAL) >= 0) Then
            Set tempREdoc = GetREDocDat(m_ReDocdat, Stcrd.DOCNUM, False)         'REsumIV.RCPNUM
            If Not (tempREdoc Is Nothing) Then
                CMPLDAT = tempREdoc.DOCDAT
                RECOD = REsumIV.RCPNUM
                fixedRE = True
            End If
         End If
      End If
      
      If fixedRE = False Then
         RECOD = ""
         CMPLDAT = -1
      End If
                  
          If mcolParam("FROM_DOC_DATE") < 0 And mcolParam("TO_DOC_DATE") < 0 Then
            showfromDate = False
         Else
            showfromDate = True
          End If
          
         If fixedRE = True Or showfromDate Then
                  
                 If (Stcrd.AREACOD <> PrevKey2 Or PrevKey3 <> Stcrd.SLMNAM) And PrevKey2 <> "" Then
                    If mcolParam("SUMMARY_MODE") = "N" Then                                                         ' ࢵ��鹷��
                       Call GenerateFooter("���", "", PrevKey4, Total2, m_Details1)
                       Vsp.FontBold = True
'                     Else
'                       Call GenerateFooter("", "", PrevKey4, Total2, m_Details1)
'                       Vsp.FontBold = False
'                     End If
            '        K = K + 1
                    Call m_Details1.GetString(2, TempStr1, TempStr2)
                    strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
                    Call Vsp.AddTable(strFormat, "", TempStr2)
                    Vsp.FontBold = False

                     For j = 1 To UBound(Total2)
                         Total2(j) = 0
                     Next j
                     End If
               End If


               If Stcrd.SLMCOD <> PrevKey1 And PrevKey1 <> "" Then         '!
                      If mcolParam("SUMMARY_MODE") = "N" Then
                        Call GenerateFooter("���", PrevKey1, PrevKey3, Total1, m_Details1)       ' �� haveValue � Call GenerateFooter **** �����ͧ�դ�Ҷ֧���ʴ���÷Ѵ���
                        Vsp.FontBold = True
                     Else
                        Call GenerateFooter("", PrevKey1, PrevKey3, Total1, m_Details1)       ' �� haveValue � Call GenerateFooter **** �����ͧ�դ�Ҷ֧���ʴ���÷Ѵ���
                        Vsp.FontBold = False
                     End If

                     Call m_Details1.GetString(2, TempStr1, TempStr2)
                     strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
                     Call Vsp.AddTable(strFormat, "", TempStr2)
                     Vsp.FontBold = False
                     
                       For j = 1 To UBound(Total1)
                           Total1(j) = 0
                        Next j
                End If
                
               If mcolParam("SUMMARY_MODE") = "N" Then
                   If Stcrd.SLMCOD <> PrevKey1 Then
                        Call GenerateHeader(Stcrd.SLMCOD, Stcrd.SLMNAM, m_Details1)
                        Call m_Details1.GetString(1, TempStr1, TempStr2)
                        strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
                        Vsp.FontBold = True
                        Call Vsp.AddTable(strFormat, "", TempStr2)
                        Vsp.FontBold = False
                   End If

                     If Stcrd.AREACOD <> PrevKey2 Or PrevKey3 <> Stcrd.SLMNAM Then   'PrevKey2 <> ""
                        Call GenerateHeader("ࢵ : ", Stcrd.AREANAM, m_Details1)
                        Call m_Details1.GetString(1, TempStr1, TempStr2)
                        strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
                        Vsp.FontBold = True
                        Call Vsp.AddTable(strFormat, "", TempStr2)
                        Vsp.FontBold = False
                     End If
               End If

               Call m_Details1.ClearField
               j = 0
               haveData = False
            
                j = j + 1
               Set HeadCf = m_Headers1.Fields(j)
               TempStr = i
               Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
               Call m_Details1.AddField(BodyCf)
   
                j = j + 1
               Set HeadCf = m_Headers1.Fields(j)
               TempStr = Stcrd.DOCNUM
               Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
               Call m_Details1.AddField(BodyCf)

                j = j + 1
               Set HeadCf = m_Headers1.Fields(j)
               TempStr = Stcrd.CUSCOD
               Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
               Call m_Details1.AddField(BodyCf)

                 j = j + 1
               Set HeadCf = m_Headers1.Fields(j)
               TempStr = Left(Stcrd.CUSNAM, 30)
               Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
               Call m_Details1.AddField(BodyCf)
      
               j = j + 1
               Set HeadCf = m_Headers1.Fields(j)
               TempStr = Left(Stcrd.STKDES, 25)
               Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
               Call m_Details1.AddField(BodyCf)

               j = j + 1
               Set HeadCf = m_Headers1.Fields(j)
                 TempStr = strNETVAL
               Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
               Call m_Details1.AddField(BodyCf)
'               Total3(j) = Total3(j) + NETVAL
'               Total2(j) = Total2(j) + NETVAL
'               Total1(j) = Total1(j) + NETVAL
            
                j = j + 1
               Set HeadCf = m_Headers1.Fields(j)
               TempStr = DateToStringExtEx2(Stcrd.DOCDAT)                                                                                 ' �ѹ������Թ���
               Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
               Call m_Details1.AddField(BodyCf)
               
                j = j + 1
               Set HeadCf = m_Headers1.Fields(j)
               If CMPLDAT > 0 Then
                      TempStr = DateToStringExtEx2(CMPLDAT)                                                                                 ' �ѹ����Ѻ����
               Else:
                     TempStr = ""
               End If
               Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
               Call m_Details1.AddField(BodyCf)
               
                j = j + 1
               Set HeadCf = m_Headers1.Fields(j)
               TempStr = RECOD                                                                                 ' RE
               Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
               Call m_Details1.AddField(BodyCf)
               

               
               If mcolParam("SUMMARY_MODE") = "N" Then
                     If mcolParam("FROM_CMPL_DATE") > 0 And mcolParam("TO_CMPL_DATE") > 0 And CMPLDAT > 0 Then
                          Call m_Details1.GetString(1, TempStr1, TempStr2)
                          strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
                          Call Vsp.AddTable(strFormat, "", TempStr2)
                          i = i + 1   '�Ѻ��÷Ѵ����������
                        Total3(6) = Total3(6) + NETVAL
                        Total2(6) = Total2(6) + NETVAL
                        Total1(6) = Total1(6) + NETVAL
                     ElseIf mcolParam("FROM_CMPL_DATE") <= 0 And mcolParam("TO_CMPL_DATE") <= 0 Then
                           Call m_Details1.GetString(1, TempStr1, TempStr2)
                          strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
                          Call Vsp.AddTable(strFormat, "", TempStr2)
                          i = i + 1   '�Ѻ��÷Ѵ����������
                        Total3(6) = Total3(6) + NETVAL
                        Total2(6) = Total2(6) + NETVAL
                        Total1(6) = Total1(6) + NETVAL
                     End If
               End If
               
                PrevKey1 = Stcrd.SLMCOD
                PrevKey3 = Stcrd.SLMNAM
                PrevKey2 = Stcrd.AREACOD
                PrevKey4 = Stcrd.AREANAM
          End If
          
                Rs.MoveNext                                                                                            ' ǹ��
         Wend


                    If mcolParam("SUMMARY_MODE") = "N" Then                                                         ' ࢵ��鹷��
                             Call GenerateFooter("���", "", PrevKey4, Total2, m_Details1)
                             Vsp.FontBold = True
                    
                  '        K = K + 1
                          Call m_Details1.GetString(2, TempStr1, TempStr2)
                          strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
                          Call Vsp.AddTable(strFormat, "", TempStr2)
                          Vsp.FontBold = False
      
                           For j = 1 To UBound(Total2)
                               Total2(j) = 0
                           Next j
                      End If

                If mcolParam("SUMMARY_MODE") = "N" Then
                  Call GenerateFooter("���", PrevKey1, PrevKey3, Total1, m_Details1)       ' �� haveValue � Call GenerateFooter **** �����ͧ�դ�Ҷ֧���ʴ���÷Ѵ���
                  Vsp.FontBold = True
               Else
                  Call GenerateFooter("", PrevKey1, PrevKey3, Total1, m_Details1)       ' �� haveValue � Call GenerateFooter **** �����ͧ�դ�Ҷ֧���ʴ���÷Ѵ���
                  Vsp.FontBold = False
               End If

               Call m_Details1.GetString(2, TempStr1, TempStr2)
               strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
               Call Vsp.AddTable(strFormat, "", TempStr2)
               Vsp.FontBold = False
            If mcolParam("SUMMARY_MODE") = "N" Then
              Call GenerateFooter("�ط��", "", "", Total3, m_Details1)
            Else
             Call GenerateFooter("", "", "�ط��", Total3, m_Details1)
            End If
            Call m_Details1.GetString(2, TempStr1, TempStr2)
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
               
               If TempRs.State = adStateOpen Then
                  TempRs.Close
               End If
               Set TempRs = Nothing
               
               'debug.print temnetval
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
         TempStr = ""
      ElseIf j = 2 Then
         TempStr = Txt1
      ElseIf j = 3 Then
         TempStr = Txt2
      ElseIf j = 4 Then
         TempStr = Txt3
     ElseIf j = 5 Then
         TempStr = ""
      ElseIf j = 7 Then
         TempStr = ""
      ElseIf j = 8 Then
         TempStr = ""
      ElseIf j = 9 Then
         TempStr = ""
      Else
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


