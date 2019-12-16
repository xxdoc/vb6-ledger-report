VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CReportAP009_1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'User Group Report
Implements CReportInterface

Private Const MODULE_NAME = "CReportAP009_1"
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
'1
Private m_ApSupAmounts As Collection
Private m_ApSupAmountCns As Collection
Private m_ApSupAmountCashs As Collection
Private m_ApSupPaids As Collection
Private m_ApSupPaidCns As Collection
Private m_PaidAllAmounts As Collection
Private m_PaidAllAmountCns As Collection
Private m_ApSupAmountTo As Collection
Private m_ApSupAmountToCn As Collection
Private m_ApSupPaidTo As Collection
Private m_ApSupPaidToCn As Collection
Private m_5 As Collection
Private m_6 As Collection
'1
'2
Private m_ApSupAmounts2 As Collection
Private m_ApSupAmountCns2 As Collection
Private m_ApSupAmountCashs2 As Collection
Private m_ApSupPaids2 As Collection
Private m_ApSupPaidCns2 As Collection
Private m_PaidAllAmounts2 As Collection
Private m_PaidAllAmountCns2 As Collection
Private m_ApSupAmountTo2 As Collection
Private m_ApSupAmountToCn2 As Collection
Private m_ApSupPaidTo2 As Collection
Private m_ApSupPaidToCn2 As Collection
Private m_52 As Collection
Private m_62 As Collection
'2
'3
Private m_ApSupAmounts3 As Collection
Private m_ApSupAmountCns3 As Collection
Private m_ApSupAmountCashs3 As Collection
Private m_ApSupPaids3 As Collection
Private m_ApSupPaidCns3 As Collection
Private m_PaidAllAmounts3 As Collection
Private m_PaidAllAmountCns3 As Collection
Private m_ApSupAmountTo3 As Collection
Private m_ApSupAmountToCn3 As Collection
Private m_ApSupPaidTo3 As Collection
Private m_ApSupPaidToCn3 As Collection
Private m_53 As Collection
Private m_63 As Collection
'3
'4
Private m_ApSupAmounts4 As Collection
Private m_ApSupAmountCns4 As Collection
Private m_ApSupAmountCashs4 As Collection
Private m_ApSupPaids4 As Collection
Private m_ApSupPaidCns4 As Collection
Private m_PaidAllAmounts4 As Collection
Private m_PaidAllAmountCns4 As Collection
Private m_ApSupAmountTo4 As Collection
Private m_ApSupAmountToCn4 As Collection
Private m_ApSupPaidTo4 As Collection
Private m_ApSupPaidToCn4 As Collection
Private m_54 As Collection
Private m_64 As Collection
'4
'5
Private m_ApSupAmounts5 As Collection
Private m_ApSupAmountCns5 As Collection
Private m_ApSupAmountCashs5 As Collection
Private m_ApSupPaids5 As Collection
Private m_ApSupPaidCns5 As Collection
Private m_PaidAllAmounts5 As Collection
Private m_PaidAllAmountCns5 As Collection
Private m_ApSupAmountTo5 As Collection
Private m_ApSupAmountToCn5 As Collection
Private m_ApSupPaidTo5 As Collection
Private m_ApSupPaidToCn5 As Collection
Private m_55 As Collection
Private m_65 As Collection
'5
'6
Private m_ApSupAmounts6 As Collection
Private m_ApSupAmountCns6 As Collection
Private m_ApSupAmountCashs6 As Collection
Private m_ApSupPaids6 As Collection
Private m_ApSupPaidCns6 As Collection
Private m_PaidAllAmounts6 As Collection
Private m_PaidAllAmountCns6 As Collection
Private m_ApSupAmountTo6 As Collection
Private m_ApSupAmountToCn6 As Collection
Private m_ApSupPaidTo6 As Collection
Private m_ApSupPaidToCn6 As Collection
Private m_56 As Collection
Private m_66 As Collection
'6
'7
Private m_ApSupAmounts7 As Collection
Private m_ApSupAmountCns7 As Collection
Private m_ApSupAmountCashs7 As Collection
Private m_ApSupPaids7 As Collection
Private m_ApSupPaidCns7 As Collection
Private m_PaidAllAmounts7 As Collection
Private m_PaidAllAmountCns7 As Collection
Private m_ApSupAmountTo7 As Collection
Private m_ApSupAmountToCn7 As Collection
Private m_ApSupPaidTo7 As Collection
Private m_ApSupPaidToCn7 As Collection
Private m_57 As Collection
Private m_67 As Collection
'7
'8
Private m_ApSupAmounts8 As Collection
Private m_ApSupAmountCns8 As Collection
Private m_ApSupAmountCashs8 As Collection
Private m_ApSupPaids8 As Collection
Private m_ApSupPaidCns8 As Collection
Private m_PaidAllAmounts8 As Collection
Private m_PaidAllAmountCns8 As Collection
Private m_ApSupAmountTo8 As Collection
Private m_ApSupAmountToCn8 As Collection
Private m_ApSupPaidTo8 As Collection
Private m_ApSupPaidToCn8 As Collection
Private m_58 As Collection
Private m_68 As Collection
'8
'9
Private m_ApSupAmounts9 As Collection
Private m_ApSupAmountCns9 As Collection
Private m_ApSupAmountCashs9 As Collection
Private m_ApSupPaids9 As Collection
Private m_ApSupPaidCns9 As Collection
Private m_PaidAllAmounts9 As Collection
Private m_PaidAllAmountCns9 As Collection
Private m_ApSupAmountTo9 As Collection
Private m_ApSupAmountToCn9 As Collection
Private m_ApSupPaidTo9 As Collection
Private m_ApSupPaidToCn9 As Collection
Private m_59 As Collection
Private m_69 As Collection
'9
'10
Private m_ApSupAmounts10 As Collection
Private m_ApSupAmountCns10 As Collection
Private m_ApSupAmountCashs10 As Collection
Private m_ApSupPaids10 As Collection
Private m_ApSupPaidCns10 As Collection
Private m_PaidAllAmounts10 As Collection
Private m_PaidAllAmountCns10 As Collection
Private m_ApSupAmountTo10 As Collection
Private m_ApSupAmountToCn10 As Collection
Private m_ApSupPaidTo10 As Collection
Private m_ApSupPaidToCn10 As Collection
Private m_510 As Collection
Private m_610 As Collection
'10
'11
Private m_ApSupAmounts11 As Collection
Private m_ApSupAmountCns11 As Collection
Private m_ApSupAmountCashs11 As Collection
Private m_ApSupPaids11 As Collection
Private m_ApSupPaidCns11 As Collection
Private m_PaidAllAmounts11 As Collection
Private m_PaidAllAmountCns11 As Collection
Private m_ApSupAmountTo11 As Collection
Private m_ApSupAmountToCn11 As Collection
Private m_ApSupPaidTo11 As Collection
Private m_ApSupPaidToCn11 As Collection
Private m_511 As Collection
Private m_611 As Collection
'11
'12
Private m_ApSupAmounts12 As Collection
Private m_ApSupAmountCns12 As Collection
Private m_ApSupAmountCashs12 As Collection
Private m_ApSupPaids12 As Collection
Private m_ApSupPaidCns12 As Collection
Private m_PaidAllAmounts12 As Collection
Private m_PaidAllAmountCns12 As Collection
Private m_ApSupAmountTo12 As Collection
Private m_ApSupAmountToCn12 As Collection
Private m_ApSupPaidTo12 As Collection
Private m_ApSupPaidToCn12 As Collection
Private m_512 As Collection
Private m_612 As Collection
'12

Private Total1(100) As Double
Private Total2(100) As Double
Private Total3(100) As Double
Private Total12(100) As Double
Private Total22(100) As Double
Private Total13(100) As Double
Private Total23(100) As Double
Private Total14(100) As Double
Private Total24(100) As Double
Private Total15(100) As Double
Private Total25(100) As Double
Private Total16(100) As Double
Private Total26(100) As Double
Private Total17(100) As Double
Private Total27(100) As Double
Private Total18(100) As Double
Private Total28(100) As Double
Private Total19(100) As Double
Private Total29(100) As Double
Private Total110(100) As Double
Private Total210(100) As Double
Private Total111(100) As Double
Private Total211(100) As Double
Private Total112(100) As Double
Private Total212(100) As Double

Private SupplierGroupColl As Collection
Private Supplier As Collection
Private HaveValue As Boolean
Private HaveValueSum As Boolean
Private FirstDate As Date
Private LastDate As Date
Private RunMonth As Long

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
   
   '1
   Set m_ApSupAmounts = New Collection
   Set m_ApSupAmountCns = New Collection
   Set m_ApSupAmountCashs = New Collection
   Set m_ApSupPaids = New Collection
   Set m_ApSupPaidCns = New Collection
   Set m_PaidAllAmounts = New Collection
   Set m_PaidAllAmountCns = New Collection
   Set m_ApSupAmountTo = New Collection
   Set m_ApSupAmountToCn = New Collection
   Set m_ApSupPaidTo = New Collection
   Set m_ApSupPaidToCn = New Collection
   Set m_5 = New Collection
   Set m_6 = New Collection
   '1
   
   '2
   Set m_ApSupAmounts2 = New Collection
   Set m_ApSupAmountCns2 = New Collection
   Set m_ApSupAmountCashs2 = New Collection
   Set m_ApSupPaids2 = New Collection
   Set m_ApSupPaidCns2 = New Collection
   Set m_PaidAllAmounts2 = New Collection
   Set m_PaidAllAmountCns2 = New Collection
   Set m_ApSupAmountTo2 = New Collection
   Set m_ApSupAmountToCn2 = New Collection
   Set m_ApSupPaidTo2 = New Collection
   Set m_ApSupPaidToCn2 = New Collection
   Set m_52 = New Collection
   Set m_62 = New Collection
   '2
   
   '3
   Set m_ApSupAmounts3 = New Collection
   Set m_ApSupAmountCns3 = New Collection
   Set m_ApSupAmountCashs3 = New Collection
   Set m_ApSupPaids3 = New Collection
   Set m_ApSupPaidCns3 = New Collection
   Set m_PaidAllAmounts3 = New Collection
   Set m_PaidAllAmountCns3 = New Collection
   Set m_ApSupAmountTo3 = New Collection
   Set m_ApSupAmountToCn3 = New Collection
   Set m_ApSupPaidTo3 = New Collection
   Set m_ApSupPaidToCn3 = New Collection
   Set m_53 = New Collection
   Set m_63 = New Collection
   '3
   
   '4
   Set m_ApSupAmounts4 = New Collection
   Set m_ApSupAmountCns4 = New Collection
   Set m_ApSupAmountCashs4 = New Collection
   Set m_ApSupPaids4 = New Collection
   Set m_ApSupPaidCns4 = New Collection
   Set m_PaidAllAmounts4 = New Collection
   Set m_PaidAllAmountCns4 = New Collection
   Set m_ApSupAmountTo4 = New Collection
   Set m_ApSupAmountToCn4 = New Collection
   Set m_ApSupPaidTo4 = New Collection
   Set m_ApSupPaidToCn4 = New Collection
   Set m_54 = New Collection
   Set m_64 = New Collection
   '4
   
   '5
   Set m_ApSupAmounts5 = New Collection
   Set m_ApSupAmountCns5 = New Collection
   Set m_ApSupAmountCashs5 = New Collection
   Set m_ApSupPaids5 = New Collection
   Set m_ApSupPaidCns5 = New Collection
   Set m_PaidAllAmounts5 = New Collection
   Set m_PaidAllAmountCns5 = New Collection
   Set m_ApSupAmountTo5 = New Collection
   Set m_ApSupAmountToCn5 = New Collection
   Set m_ApSupPaidTo5 = New Collection
   Set m_ApSupPaidToCn5 = New Collection
   Set m_55 = New Collection
   Set m_65 = New Collection
   '5
   
   '6
   Set m_ApSupAmounts6 = New Collection
   Set m_ApSupAmountCns6 = New Collection
   Set m_ApSupAmountCashs6 = New Collection
   Set m_ApSupPaids6 = New Collection
   Set m_ApSupPaidCns6 = New Collection
   Set m_PaidAllAmounts6 = New Collection
   Set m_PaidAllAmountCns6 = New Collection
   Set m_ApSupAmountTo6 = New Collection
   Set m_ApSupAmountToCn6 = New Collection
   Set m_ApSupPaidTo6 = New Collection
   Set m_ApSupPaidToCn6 = New Collection
   Set m_56 = New Collection
   Set m_66 = New Collection
   '6
   
   '7
   Set m_ApSupAmounts7 = New Collection
   Set m_ApSupAmountCns7 = New Collection
   Set m_ApSupAmountCashs7 = New Collection
   Set m_ApSupPaids7 = New Collection
   Set m_ApSupPaidCns7 = New Collection
   Set m_PaidAllAmounts7 = New Collection
   Set m_PaidAllAmountCns7 = New Collection
   Set m_ApSupAmountTo7 = New Collection
   Set m_ApSupAmountToCn7 = New Collection
   Set m_ApSupPaidTo7 = New Collection
   Set m_ApSupPaidToCn7 = New Collection
   Set m_57 = New Collection
   Set m_67 = New Collection
   '7
   
   '8
   Set m_ApSupAmounts8 = New Collection
   Set m_ApSupAmountCns8 = New Collection
   Set m_ApSupAmountCashs8 = New Collection
   Set m_ApSupPaids8 = New Collection
   Set m_ApSupPaidCns8 = New Collection
   Set m_PaidAllAmounts8 = New Collection
   Set m_PaidAllAmountCns8 = New Collection
   Set m_ApSupAmountTo8 = New Collection
   Set m_ApSupAmountToCn8 = New Collection
   Set m_ApSupPaidTo8 = New Collection
   Set m_ApSupPaidToCn8 = New Collection
   Set m_58 = New Collection
   Set m_68 = New Collection
   '8
   
   '9
   Set m_ApSupAmounts9 = New Collection
   Set m_ApSupAmountCns9 = New Collection
   Set m_ApSupAmountCashs9 = New Collection
   Set m_ApSupPaids9 = New Collection
   Set m_ApSupPaidCns9 = New Collection
   Set m_PaidAllAmounts9 = New Collection
   Set m_PaidAllAmountCns9 = New Collection
   Set m_ApSupAmountTo9 = New Collection
   Set m_ApSupAmountToCn9 = New Collection
   Set m_ApSupPaidTo9 = New Collection
   Set m_ApSupPaidToCn9 = New Collection
   Set m_59 = New Collection
   Set m_69 = New Collection
   '9
   
   '10
   Set m_ApSupAmounts10 = New Collection
   Set m_ApSupAmountCns10 = New Collection
   Set m_ApSupAmountCashs10 = New Collection
   Set m_ApSupPaids10 = New Collection
   Set m_ApSupPaidCns10 = New Collection
   Set m_PaidAllAmounts10 = New Collection
   Set m_PaidAllAmountCns10 = New Collection
   Set m_ApSupAmountTo10 = New Collection
   Set m_ApSupAmountToCn10 = New Collection
   Set m_ApSupPaidTo10 = New Collection
   Set m_ApSupPaidToCn10 = New Collection
   Set m_510 = New Collection
   Set m_610 = New Collection
   '10
   
   '11
   Set m_ApSupAmounts11 = New Collection
   Set m_ApSupAmountCns11 = New Collection
   Set m_ApSupAmountCashs11 = New Collection
   Set m_ApSupPaids11 = New Collection
   Set m_ApSupPaidCns11 = New Collection
   Set m_PaidAllAmounts11 = New Collection
   Set m_PaidAllAmountCns11 = New Collection
   Set m_ApSupAmountTo11 = New Collection
   Set m_ApSupAmountToCn11 = New Collection
   Set m_ApSupPaidTo11 = New Collection
   Set m_ApSupPaidToCn11 = New Collection
   Set m_511 = New Collection
   Set m_611 = New Collection
   '11
   
   '12
   Set m_ApSupAmounts12 = New Collection
   Set m_ApSupAmountCns12 = New Collection
   Set m_ApSupAmountCashs12 = New Collection
   Set m_ApSupPaids12 = New Collection
   Set m_ApSupPaidCns12 = New Collection
   Set m_PaidAllAmounts12 = New Collection
   Set m_PaidAllAmountCns12 = New Collection
   Set m_ApSupAmountTo12 = New Collection
   Set m_ApSupAmountToCn12 = New Collection
   Set m_ApSupPaidTo12 = New Collection
   Set m_ApSupPaidToCn12 = New Collection
   Set m_512 = New Collection
   Set m_612 = New Collection
   '12
   
   Set SupplierGroupColl = New Collection
   Set Supplier = New Collection
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
   
   '1
   Set m_ApSupAmounts = Nothing
   Set m_ApSupAmountCns = Nothing
   Set m_ApSupAmountCashs = Nothing
   Set m_ApSupPaids = Nothing
   Set m_ApSupPaidCns = Nothing
   Set m_PaidAllAmounts = Nothing
   Set m_PaidAllAmountCns = Nothing
   Set m_ApSupAmountTo = Nothing
   Set m_ApSupAmountToCn = Nothing
   Set m_ApSupPaidTo = Nothing
   Set m_ApSupPaidToCn = Nothing
   Set m_5 = Nothing
   Set m_6 = Nothing
   '1
   
   '2
   Set m_ApSupAmounts2 = Nothing
   Set m_ApSupAmountCns2 = Nothing
   Set m_ApSupAmountCashs2 = Nothing
   Set m_ApSupPaids2 = Nothing
   Set m_ApSupPaidCns2 = Nothing
   Set m_PaidAllAmounts2 = Nothing
   Set m_PaidAllAmountCns2 = Nothing
   Set m_ApSupAmountTo2 = Nothing
   Set m_ApSupAmountToCn2 = Nothing
   Set m_ApSupPaidTo2 = Nothing
   Set m_ApSupPaidToCn2 = Nothing
   Set m_52 = Nothing
   Set m_62 = Nothing
   '2
   
   '3
   Set m_ApSupAmounts3 = Nothing
   Set m_ApSupAmountCns3 = Nothing
   Set m_ApSupAmountCashs3 = Nothing
   Set m_ApSupPaids3 = Nothing
   Set m_ApSupPaidCns3 = Nothing
   Set m_PaidAllAmounts3 = Nothing
   Set m_PaidAllAmountCns3 = Nothing
   Set m_ApSupAmountTo3 = Nothing
   Set m_ApSupAmountToCn3 = Nothing
   Set m_ApSupPaidTo3 = Nothing
   Set m_ApSupPaidToCn3 = Nothing
   Set m_53 = Nothing
   Set m_63 = Nothing
   '3
   
   '4
   Set m_ApSupAmounts4 = Nothing
   Set m_ApSupAmountCns4 = Nothing
   Set m_ApSupAmountCashs4 = Nothing
   Set m_ApSupPaids4 = Nothing
   Set m_ApSupPaidCns4 = Nothing
   Set m_PaidAllAmounts4 = Nothing
   Set m_PaidAllAmountCns4 = Nothing
   Set m_ApSupAmountTo4 = Nothing
   Set m_ApSupAmountToCn4 = Nothing
   Set m_ApSupPaidTo4 = Nothing
   Set m_ApSupPaidToCn4 = Nothing
   Set m_54 = Nothing
   Set m_64 = Nothing
   '4
   
   '5
   Set m_ApSupAmounts5 = Nothing
   Set m_ApSupAmountCns5 = Nothing
   Set m_ApSupAmountCashs5 = Nothing
   Set m_ApSupPaids5 = Nothing
   Set m_ApSupPaidCns5 = Nothing
   Set m_PaidAllAmounts5 = Nothing
   Set m_PaidAllAmountCns5 = Nothing
   Set m_ApSupAmountTo5 = Nothing
   Set m_ApSupAmountToCn5 = Nothing
   Set m_ApSupPaidTo5 = Nothing
   Set m_ApSupPaidToCn5 = Nothing
   Set m_55 = Nothing
   Set m_65 = Nothing
   '5
   
      '6
   Set m_ApSupAmounts6 = Nothing
   Set m_ApSupAmountCns6 = Nothing
   Set m_ApSupAmountCashs6 = Nothing
   Set m_ApSupPaids6 = Nothing
   Set m_ApSupPaidCns6 = Nothing
   Set m_PaidAllAmounts6 = Nothing
   Set m_PaidAllAmountCns6 = Nothing
   Set m_ApSupAmountTo6 = Nothing
   Set m_ApSupAmountToCn6 = Nothing
   Set m_ApSupPaidTo6 = Nothing
   Set m_ApSupPaidToCn6 = Nothing
   Set m_56 = Nothing
   Set m_66 = Nothing
   '6
   
      '7
   Set m_ApSupAmounts7 = Nothing
   Set m_ApSupAmountCns7 = Nothing
   Set m_ApSupAmountCashs7 = Nothing
   Set m_ApSupPaids7 = Nothing
   Set m_ApSupPaidCns7 = Nothing
   Set m_PaidAllAmounts7 = Nothing
   Set m_PaidAllAmountCns7 = Nothing
   Set m_ApSupAmountTo7 = Nothing
   Set m_ApSupAmountToCn7 = Nothing
   Set m_ApSupPaidTo7 = Nothing
   Set m_ApSupPaidToCn7 = Nothing
   Set m_57 = Nothing
   Set m_67 = Nothing
   '7
   
      '8
   Set m_ApSupAmounts8 = Nothing
   Set m_ApSupAmountCns8 = Nothing
   Set m_ApSupAmountCashs8 = Nothing
   Set m_ApSupPaids8 = Nothing
   Set m_ApSupPaidCns8 = Nothing
   Set m_PaidAllAmounts8 = Nothing
   Set m_PaidAllAmountCns8 = Nothing
   Set m_ApSupAmountTo8 = Nothing
   Set m_ApSupAmountToCn8 = Nothing
   Set m_ApSupPaidTo8 = Nothing
   Set m_ApSupPaidToCn8 = Nothing
   Set m_58 = Nothing
   Set m_68 = Nothing
   '8
   
      '9
   Set m_ApSupAmounts9 = Nothing
   Set m_ApSupAmountCns9 = Nothing
   Set m_ApSupAmountCashs9 = Nothing
   Set m_ApSupPaids9 = Nothing
   Set m_ApSupPaidCns9 = Nothing
   Set m_PaidAllAmounts9 = Nothing
   Set m_PaidAllAmountCns9 = Nothing
   Set m_ApSupAmountTo9 = Nothing
   Set m_ApSupAmountToCn9 = Nothing
   Set m_ApSupPaidTo9 = Nothing
   Set m_ApSupPaidToCn9 = Nothing
   Set m_59 = Nothing
   Set m_69 = Nothing
   '9
   
      '10
   Set m_ApSupAmounts10 = Nothing
   Set m_ApSupAmountCns10 = Nothing
   Set m_ApSupAmountCashs10 = Nothing
   Set m_ApSupPaids10 = Nothing
   Set m_ApSupPaidCns10 = Nothing
   Set m_PaidAllAmounts10 = Nothing
   Set m_PaidAllAmountCns10 = Nothing
   Set m_ApSupAmountTo10 = Nothing
   Set m_ApSupAmountToCn10 = Nothing
   Set m_ApSupPaidTo10 = Nothing
   Set m_ApSupPaidToCn10 = Nothing
   Set m_510 = Nothing
   Set m_610 = Nothing
   '10
   
      '11
   Set m_ApSupAmounts11 = Nothing
   Set m_ApSupAmountCns11 = Nothing
   Set m_ApSupAmountCashs11 = Nothing
   Set m_ApSupPaids11 = Nothing
   Set m_ApSupPaidCns11 = Nothing
   Set m_PaidAllAmounts11 = Nothing
   Set m_PaidAllAmountCns11 = Nothing
   Set m_ApSupAmountTo11 = Nothing
   Set m_ApSupAmountToCn11 = Nothing
   Set m_ApSupPaidTo11 = Nothing
   Set m_ApSupPaidToCn11 = Nothing
   Set m_511 = Nothing
   Set m_611 = Nothing
   '11
   
      '12
   Set m_ApSupAmounts12 = Nothing
   Set m_ApSupAmountCns12 = Nothing
   Set m_ApSupAmountCashs12 = Nothing
   Set m_ApSupPaids12 = Nothing
   Set m_ApSupPaidCns12 = Nothing
   Set m_PaidAllAmounts12 = Nothing
   Set m_PaidAllAmountCns12 = Nothing
   Set m_ApSupAmountTo12 = Nothing
   Set m_ApSupAmountToCn12 = Nothing
   Set m_ApSupPaidTo12 = Nothing
   Set m_ApSupPaidToCn12 = Nothing
   Set m_512 = Nothing
   Set m_612 = Nothing
   '12
   
   Set SupplierGroupColl = Nothing
   Set Supplier = Nothing
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
   Call Cf.SetFieldValue(4, "^", "ประเภทผู้ขาย", "<")
   Call m_Headers0.AddField(Cf)
   Call Cf.SetFieldValue(16, "^", "รายละเอียดชำระ", "<")
   Call m_Headers0.AddField(Cf)
   '===
   Call Cf.SetFieldValue(1, "^", "NO.", "^")
   Call m_Headers1.AddField(Cf)
   Call Cf.SetFieldValue(4, "^", "ชื่อผู้ขาย", "<")
   Call m_Headers1.AddField(Cf)
   Call Cf.SetFieldValue(2, "^", "ยอดหนี้ยกมา", ">")
   Call m_Headers1.AddField(Cf)
   Call Cf.SetFieldValue(2, "^", "ซื้อ/เพิ่มหนี้", ">")
   Call m_Headers1.AddField(Cf)
   Call Cf.SetFieldValue(2, "^", "จ่าย/ลดหนี้", ">")
   Call m_Headers1.AddField(Cf)
   Call Cf.SetFieldValue(2, "^", "ยอดหนี้ยกไป", ">")
   Call m_Headers1.AddField(Cf)
   Call Cf.SetFieldValue(2, "^", "ยอดจ่ายจริง", ">")
   Call m_Headers1.AddField(Cf)
   Call Cf.SetFieldValue(2, "^", "ยอดจ่ายPOST", ">")
   Call m_Headers1.AddField(Cf)
   Call Cf.SetFieldValue(2, "^", "หนี้คงเหลือ", ">")
   Call m_Headers1.AddField(Cf)
   Call Cf.SetFieldValue(2, "^", "ซื้อ-จ่าย", ">")
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
'Dim FromDate As String
'Dim ToDate As String
'Dim ToPaidDate As String
   RunMonth = Val(mcolParam("RUN_MONTH"))
   
   Call GetFirstLastDate(DateSerial(mcolParam("YEAR_NO") - 543, mcolParam("MONTH_ID"), 1), FirstDate, LastDate)
   
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
   
   Call LoadSupplier(Nothing, Supplier, , mcolParam("SUPPLIER_CODE"))
   Call LoadSupplierGroup(Nothing, SupplierGroupColl, , mcolParam("SUPPLIER_CODE"), mcolParam("DATA_TYPE_ID"))
   Call AddCollection(SupplierGroupColl, Supplier)
   
   'เดือน 1
   '-------------------------------- ยกมา --------------------------------
   Call LoadAPAmountBySup(Nothing, m_ApSupAmounts, -1, DateAdd("D", -1, FirstDate), "('3','4')")    'ยอดซื้อ ถึงวันที่
   Call LoadAPAmountBySup(Nothing, m_ApSupAmountCns, -1, DateAdd("D", -1, FirstDate), "('5')")     'ยอดซื้อ ถึงวันที่       ลดหนี้
   Call LoadAPAmountBySup(Nothing, m_ApSupAmountCashs, FirstDate, LastDate, "('9')")    'ยอดจ่ายเงินสด ถึงวันที่
   Call LoadPaidAmountBySup(Nothing, m_ApSupPaids, -1, DateAdd("D", -1, FirstDate), -1, -1, "('3','4')")           'ยอดจ่ายถึงวันที่
   Call LoadPaidAmountBySup(Nothing, m_ApSupPaidCns, -1, DateAdd("D", -1, FirstDate), -1, -1, "('5')")           'ยอดจ่ายถึงวันที่    ลดหนี้
   Call LoadPaidAmountBySup(Nothing, m_PaidAllAmounts, -1, -1, -1, DateAdd("D", -1, FirstDate), "('3','4')")         'ยอดตัดบิลทั้งหมดของบิลภายในวันที่กำหนด
   Call LoadPaidAmountBySup(Nothing, m_PaidAllAmountCns, -1, -1, -1, DateAdd("D", -1, FirstDate), "('5')")          'ยอดตัดบิลทั้งหมดของบิลภายในวันที่กำหนด ลดหนี้
   '-------------------------------- ยกมา --------------------------------
   '-------------------------------- เคลื่อนไหว --------------------------------
   Call LoadAPAmountBySup(Nothing, m_ApSupAmountTo, FirstDate, LastDate, "('3','4')")     'ยอดซื้อ ถึงวันที่
   Call LoadAPAmountBySup(Nothing, m_ApSupAmountToCn, FirstDate, LastDate, "('5')")     'ลดหนี้ ถึงวันที่
   Call LoadPaidAmountBySup(Nothing, m_ApSupPaidTo, FirstDate, LastDate, -1, -1, "('3','4')")           'ยอดจ่ายถึงวันที่
   Call LoadPaidAmountBySup(Nothing, m_ApSupPaidToCn, FirstDate, LastDate, -1, -1, "('5')")           'ยอดจ่ายถึงวันที่
   '-------------------------------- เคลื่อนไหว --------------------------------
   '-------------------------------- จ่ายเช็ค --------------------------------
   Call LoadBktChqAmountBySup(Nothing, m_5, , , FirstDate, LastDate)                                                                                      'ยอดจ่ายจริง DUE เดือน4
   Call LoadBktChqAmountBySup(Nothing, m_6, , LastDate, DateAdd("D", 1, LastDate))                                                                        'ยอดจ่าย POST DUE หลังเดือน4
   '-------------------------------- จ่ายเช็ค --------------------------------
   'เดือน 1
   
   If (RunMonth >= 2) Then
   'เดือน 2
      FirstDate = DateAdd("M", 1, FirstDate)
      Call GetFirstLastDate(FirstDate, FirstDate, LastDate)
      '-------------------------------- ยกมา --------------------------------
      Call LoadAPAmountBySup(Nothing, m_ApSupAmounts2, -1, DateAdd("D", -1, FirstDate), "('3','4')")    'ยอดซื้อ ถึงวันที่
      Call LoadAPAmountBySup(Nothing, m_ApSupAmountCns2, -1, DateAdd("D", -1, FirstDate), "('5')")     'ยอดซื้อ ถึงวันที่       ลดหนี้
      Call LoadAPAmountBySup(Nothing, m_ApSupAmountCashs2, FirstDate, LastDate, "('9')")    'ยอดจ่ายเงินสด ถึงวันที่
      Call LoadPaidAmountBySup(Nothing, m_ApSupPaids2, -1, DateAdd("D", -1, FirstDate), -1, -1, "('3','4')")           'ยอดจ่ายถึงวันที่
      Call LoadPaidAmountBySup(Nothing, m_ApSupPaidCns2, -1, DateAdd("D", -1, FirstDate), -1, -1, "('5')")           'ยอดจ่ายถึงวันที่    ลดหนี้
      Call LoadPaidAmountBySup(Nothing, m_PaidAllAmounts2, -1, -1, -1, DateAdd("D", -1, FirstDate), "('3','4')")         'ยอดตัดบิลทั้งหมดของบิลภายในวันที่กำหนด
      Call LoadPaidAmountBySup(Nothing, m_PaidAllAmountCns2, -1, -1, -1, DateAdd("D", -1, FirstDate), "('5')")          'ยอดตัดบิลทั้งหมดของบิลภายในวันที่กำหนด ลดหนี้
      '-------------------------------- ยกมา --------------------------------
      '-------------------------------- เคลื่อนไหว --------------------------------
      Call LoadAPAmountBySup(Nothing, m_ApSupAmountTo2, FirstDate, LastDate, "('3','4')")     'ยอดซื้อ ถึงวันที่
      Call LoadAPAmountBySup(Nothing, m_ApSupAmountToCn2, FirstDate, LastDate, "('5')")     'ลดหนี้ ถึงวันที่
      Call LoadPaidAmountBySup(Nothing, m_ApSupPaidTo2, FirstDate, LastDate, -1, -1, "('3','4')")           'ยอดจ่ายถึงวันที่
      Call LoadPaidAmountBySup(Nothing, m_ApSupPaidToCn2, FirstDate, LastDate, -1, -1, "('5')")           'ยอดจ่ายถึงวันที่
      '-------------------------------- เคลื่อนไหว --------------------------------
      '-------------------------------- จ่ายเช็ค --------------------------------
      Call LoadBktChqAmountBySup(Nothing, m_52, , , FirstDate, LastDate)                                                                                      'ยอดจ่ายจริง DUE เดือน4
      Call LoadBktChqAmountBySup(Nothing, m_62, , LastDate, DateAdd("D", 1, LastDate))                                                                        'ยอดจ่าย POST DUE หลังเดือน4
      '-------------------------------- จ่ายเช็ค --------------------------------
   'เดือน 2
   End If
   
   If (RunMonth >= 3) Then
   'เดือน 3
      FirstDate = DateAdd("M", 1, FirstDate)
      Call GetFirstLastDate(FirstDate, FirstDate, LastDate)
      '-------------------------------- ยกมา --------------------------------
      Call LoadAPAmountBySup(Nothing, m_ApSupAmounts3, -1, DateAdd("D", -1, FirstDate), "('3','4')")    'ยอดซื้อ ถึงวันที่
      Call LoadAPAmountBySup(Nothing, m_ApSupAmountCns3, -1, DateAdd("D", -1, FirstDate), "('5')")     'ยอดซื้อ ถึงวันที่       ลดหนี้
      Call LoadAPAmountBySup(Nothing, m_ApSupAmountCashs3, FirstDate, LastDate, "('9')")    'ยอดจ่ายเงินสด ถึงวันที่
      Call LoadPaidAmountBySup(Nothing, m_ApSupPaids3, -1, DateAdd("D", -1, FirstDate), -1, -1, "('3','4')")           'ยอดจ่ายถึงวันที่
      Call LoadPaidAmountBySup(Nothing, m_ApSupPaidCns3, -1, DateAdd("D", -1, FirstDate), -1, -1, "('5')")           'ยอดจ่ายถึงวันที่    ลดหนี้
      Call LoadPaidAmountBySup(Nothing, m_PaidAllAmounts3, -1, -1, -1, DateAdd("D", -1, FirstDate), "('3','4')")         'ยอดตัดบิลทั้งหมดของบิลภายในวันที่กำหนด
      Call LoadPaidAmountBySup(Nothing, m_PaidAllAmountCns3, -1, -1, -1, DateAdd("D", -1, FirstDate), "('5')")          'ยอดตัดบิลทั้งหมดของบิลภายในวันที่กำหนด ลดหนี้
      '-------------------------------- ยกมา --------------------------------
      '-------------------------------- เคลื่อนไหว --------------------------------
      Call LoadAPAmountBySup(Nothing, m_ApSupAmountTo3, FirstDate, LastDate, "('3','4')")     'ยอดซื้อ ถึงวันที่
      Call LoadAPAmountBySup(Nothing, m_ApSupAmountToCn3, FirstDate, LastDate, "('5')")     'ลดหนี้ ถึงวันที่
      Call LoadPaidAmountBySup(Nothing, m_ApSupPaidTo3, FirstDate, LastDate, -1, -1, "('3','4')")           'ยอดจ่ายถึงวันที่
      Call LoadPaidAmountBySup(Nothing, m_ApSupPaidToCn3, FirstDate, LastDate, -1, -1, "('5')")           'ยอดจ่ายถึงวันที่
      '-------------------------------- เคลื่อนไหว --------------------------------
      '-------------------------------- จ่ายเช็ค --------------------------------
      Call LoadBktChqAmountBySup(Nothing, m_53, , , FirstDate, LastDate)                                                                                      'ยอดจ่ายจริง DUE เดือน4
      Call LoadBktChqAmountBySup(Nothing, m_63, , LastDate, DateAdd("D", 1, LastDate))                                                                        'ยอดจ่าย POST DUE หลังเดือน4
      '-------------------------------- จ่ายเช็ค --------------------------------
   'เดือน 3
   End If
   
   If (RunMonth >= 4) Then
   'เดือน 4
      FirstDate = DateAdd("M", 1, FirstDate)
      Call GetFirstLastDate(FirstDate, FirstDate, LastDate)
      '-------------------------------- ยกมา --------------------------------
      Call LoadAPAmountBySup(Nothing, m_ApSupAmounts4, -1, DateAdd("D", -1, FirstDate), "('3','4')")    'ยอดซื้อ ถึงวันที่
      Call LoadAPAmountBySup(Nothing, m_ApSupAmountCns4, -1, DateAdd("D", -1, FirstDate), "('5')")     'ยอดซื้อ ถึงวันที่       ลดหนี้
      Call LoadAPAmountBySup(Nothing, m_ApSupAmountCashs4, FirstDate, LastDate, "('9')")    'ยอดจ่ายเงินสด ถึงวันที่
      Call LoadPaidAmountBySup(Nothing, m_ApSupPaids4, -1, DateAdd("D", -1, FirstDate), -1, -1, "('3','4')")           'ยอดจ่ายถึงวันที่
      Call LoadPaidAmountBySup(Nothing, m_ApSupPaidCns4, -1, DateAdd("D", -1, FirstDate), -1, -1, "('5')")           'ยอดจ่ายถึงวันที่    ลดหนี้
      Call LoadPaidAmountBySup(Nothing, m_PaidAllAmounts4, -1, -1, -1, DateAdd("D", -1, FirstDate), "('3','4')")         'ยอดตัดบิลทั้งหมดของบิลภายในวันที่กำหนด
      Call LoadPaidAmountBySup(Nothing, m_PaidAllAmountCns4, -1, -1, -1, DateAdd("D", -1, FirstDate), "('5')")          'ยอดตัดบิลทั้งหมดของบิลภายในวันที่กำหนด ลดหนี้
      '-------------------------------- ยกมา --------------------------------
      '-------------------------------- เคลื่อนไหว --------------------------------
      Call LoadAPAmountBySup(Nothing, m_ApSupAmountTo4, FirstDate, LastDate, "('3','4')")     'ยอดซื้อ ถึงวันที่
      Call LoadAPAmountBySup(Nothing, m_ApSupAmountToCn4, FirstDate, LastDate, "('5')")     'ลดหนี้ ถึงวันที่
      Call LoadPaidAmountBySup(Nothing, m_ApSupPaidTo4, FirstDate, LastDate, -1, -1, "('3','4')")           'ยอดจ่ายถึงวันที่
      Call LoadPaidAmountBySup(Nothing, m_ApSupPaidToCn4, FirstDate, LastDate, -1, -1, "('5')")           'ยอดจ่ายถึงวันที่
      '-------------------------------- เคลื่อนไหว --------------------------------
      '-------------------------------- จ่ายเช็ค --------------------------------
      Call LoadBktChqAmountBySup(Nothing, m_54, , , FirstDate, LastDate)                                                                                      'ยอดจ่ายจริง DUE เดือน4
      Call LoadBktChqAmountBySup(Nothing, m_64, , LastDate, DateAdd("D", 1, LastDate))                                                                        'ยอดจ่าย POST DUE หลังเดือน4
      '-------------------------------- จ่ายเช็ค --------------------------------
   'เดือน 4
   End If
   
   If (RunMonth >= 5) Then
   'เดือน 5
      FirstDate = DateAdd("M", 1, FirstDate)
      Call GetFirstLastDate(FirstDate, FirstDate, LastDate)
      '-------------------------------- ยกมา --------------------------------
      Call LoadAPAmountBySup(Nothing, m_ApSupAmounts5, -1, DateAdd("D", -1, FirstDate), "('3','4')")    'ยอดซื้อ ถึงวันที่
      Call LoadAPAmountBySup(Nothing, m_ApSupAmountCns5, -1, DateAdd("D", -1, FirstDate), "('5')")     'ยอดซื้อ ถึงวันที่       ลดหนี้
      Call LoadAPAmountBySup(Nothing, m_ApSupAmountCashs5, FirstDate, LastDate, "('9')")    'ยอดจ่ายเงินสด ถึงวันที่
      Call LoadPaidAmountBySup(Nothing, m_ApSupPaids5, -1, DateAdd("D", -1, FirstDate), -1, -1, "('3','4')")           'ยอดจ่ายถึงวันที่
      Call LoadPaidAmountBySup(Nothing, m_ApSupPaidCns5, -1, DateAdd("D", -1, FirstDate), -1, -1, "('5')")           'ยอดจ่ายถึงวันที่    ลดหนี้
      Call LoadPaidAmountBySup(Nothing, m_PaidAllAmounts5, -1, -1, -1, DateAdd("D", -1, FirstDate), "('3','4')")         'ยอดตัดบิลทั้งหมดของบิลภายในวันที่กำหนด
      Call LoadPaidAmountBySup(Nothing, m_PaidAllAmountCns5, -1, -1, -1, DateAdd("D", -1, FirstDate), "('5')")          'ยอดตัดบิลทั้งหมดของบิลภายในวันที่กำหนด ลดหนี้
      '-------------------------------- ยกมา --------------------------------
      '-------------------------------- เคลื่อนไหว --------------------------------
      Call LoadAPAmountBySup(Nothing, m_ApSupAmountTo5, FirstDate, LastDate, "('3','4')")     'ยอดซื้อ ถึงวันที่
      Call LoadAPAmountBySup(Nothing, m_ApSupAmountToCn5, FirstDate, LastDate, "('5')")     'ลดหนี้ ถึงวันที่
      Call LoadPaidAmountBySup(Nothing, m_ApSupPaidTo5, FirstDate, LastDate, -1, -1, "('3','4')")           'ยอดจ่ายถึงวันที่
      Call LoadPaidAmountBySup(Nothing, m_ApSupPaidToCn5, FirstDate, LastDate, -1, -1, "('5')")           'ยอดจ่ายถึงวันที่
      '-------------------------------- เคลื่อนไหว --------------------------------
      '-------------------------------- จ่ายเช็ค --------------------------------
      Call LoadBktChqAmountBySup(Nothing, m_55, , , FirstDate, LastDate)                                                                                      'ยอดจ่ายจริง DUE เดือน4
      Call LoadBktChqAmountBySup(Nothing, m_65, , LastDate, DateAdd("D", 1, LastDate))                                                                        'ยอดจ่าย POST DUE หลังเดือน4
      '-------------------------------- จ่ายเช็ค --------------------------------
   'เดือน 5
   End If
   If (RunMonth >= 6) Then
   'เดือน 6
      FirstDate = DateAdd("M", 1, FirstDate)
      Call GetFirstLastDate(FirstDate, FirstDate, LastDate)
      '-------------------------------- ยกมา --------------------------------
      Call LoadAPAmountBySup(Nothing, m_ApSupAmounts6, -1, DateAdd("D", -1, FirstDate), "('3','4')")    'ยอดซื้อ ถึงวันที่
      Call LoadAPAmountBySup(Nothing, m_ApSupAmountCns6, -1, DateAdd("D", -1, FirstDate), "('5')")     'ยอดซื้อ ถึงวันที่       ลดหนี้
      Call LoadAPAmountBySup(Nothing, m_ApSupAmountCashs6, FirstDate, LastDate, "('9')")    'ยอดจ่ายเงินสด ถึงวันที่
      Call LoadPaidAmountBySup(Nothing, m_ApSupPaids6, -1, DateAdd("D", -1, FirstDate), -1, -1, "('3','4')")           'ยอดจ่ายถึงวันที่
      Call LoadPaidAmountBySup(Nothing, m_ApSupPaidCns6, -1, DateAdd("D", -1, FirstDate), -1, -1, "('5')")           'ยอดจ่ายถึงวันที่    ลดหนี้
      Call LoadPaidAmountBySup(Nothing, m_PaidAllAmounts6, -1, -1, -1, DateAdd("D", -1, FirstDate), "('3','4')")         'ยอดตัดบิลทั้งหมดของบิลภายในวันที่กำหนด
      Call LoadPaidAmountBySup(Nothing, m_PaidAllAmountCns6, -1, -1, -1, DateAdd("D", -1, FirstDate), "('5')")          'ยอดตัดบิลทั้งหมดของบิลภายในวันที่กำหนด ลดหนี้
      '-------------------------------- ยกมา --------------------------------
      '-------------------------------- เคลื่อนไหว --------------------------------
      Call LoadAPAmountBySup(Nothing, m_ApSupAmountTo6, FirstDate, LastDate, "('3','4')")     'ยอดซื้อ ถึงวันที่
      Call LoadAPAmountBySup(Nothing, m_ApSupAmountToCn6, FirstDate, LastDate, "('5')")     'ลดหนี้ ถึงวันที่
      Call LoadPaidAmountBySup(Nothing, m_ApSupPaidTo6, FirstDate, LastDate, -1, -1, "('3','4')")           'ยอดจ่ายถึงวันที่
      Call LoadPaidAmountBySup(Nothing, m_ApSupPaidToCn6, FirstDate, LastDate, -1, -1, "('5')")           'ยอดจ่ายถึงวันที่
      '-------------------------------- เคลื่อนไหว --------------------------------
      '-------------------------------- จ่ายเช็ค --------------------------------
      Call LoadBktChqAmountBySup(Nothing, m_56, , , FirstDate, LastDate)                                                                                      'ยอดจ่ายจริง DUE เดือน4
      Call LoadBktChqAmountBySup(Nothing, m_66, , LastDate, DateAdd("D", 1, LastDate))                                                                        'ยอดจ่าย POST DUE หลังเดือน4
      '-------------------------------- จ่ายเช็ค --------------------------------
   'เดือน 6
   End If
   
   If (RunMonth >= 7) Then
   'เดือน 7
      FirstDate = DateAdd("M", 1, FirstDate)
      Call GetFirstLastDate(FirstDate, FirstDate, LastDate)
      '-------------------------------- ยกมา --------------------------------
      Call LoadAPAmountBySup(Nothing, m_ApSupAmounts7, -1, DateAdd("D", -1, FirstDate), "('3','4')")    'ยอดซื้อ ถึงวันที่
      Call LoadAPAmountBySup(Nothing, m_ApSupAmountCns7, -1, DateAdd("D", -1, FirstDate), "('5')")     'ยอดซื้อ ถึงวันที่       ลดหนี้
      Call LoadAPAmountBySup(Nothing, m_ApSupAmountCashs7, FirstDate, LastDate, "('9')")    'ยอดจ่ายเงินสด ถึงวันที่
      Call LoadPaidAmountBySup(Nothing, m_ApSupPaids7, -1, DateAdd("D", -1, FirstDate), -1, -1, "('3','4')")           'ยอดจ่ายถึงวันที่
      Call LoadPaidAmountBySup(Nothing, m_ApSupPaidCns7, -1, DateAdd("D", -1, FirstDate), -1, -1, "('5')")           'ยอดจ่ายถึงวันที่    ลดหนี้
      Call LoadPaidAmountBySup(Nothing, m_PaidAllAmounts7, -1, -1, -1, DateAdd("D", -1, FirstDate), "('3','4')")         'ยอดตัดบิลทั้งหมดของบิลภายในวันที่กำหนด
      Call LoadPaidAmountBySup(Nothing, m_PaidAllAmountCns7, -1, -1, -1, DateAdd("D", -1, FirstDate), "('5')")          'ยอดตัดบิลทั้งหมดของบิลภายในวันที่กำหนด ลดหนี้
      '-------------------------------- ยกมา --------------------------------
      '-------------------------------- เคลื่อนไหว --------------------------------
      Call LoadAPAmountBySup(Nothing, m_ApSupAmountTo7, FirstDate, LastDate, "('3','4')")     'ยอดซื้อ ถึงวันที่
      Call LoadAPAmountBySup(Nothing, m_ApSupAmountToCn7, FirstDate, LastDate, "('5')")     'ลดหนี้ ถึงวันที่
      Call LoadPaidAmountBySup(Nothing, m_ApSupPaidTo7, FirstDate, LastDate, -1, -1, "('3','4')")           'ยอดจ่ายถึงวันที่
      Call LoadPaidAmountBySup(Nothing, m_ApSupPaidToCn7, FirstDate, LastDate, -1, -1, "('5')")           'ยอดจ่ายถึงวันที่
      '-------------------------------- เคลื่อนไหว --------------------------------
      '-------------------------------- จ่ายเช็ค --------------------------------
      Call LoadBktChqAmountBySup(Nothing, m_57, , , FirstDate, LastDate)                                                                                      'ยอดจ่ายจริง DUE เดือน4
      Call LoadBktChqAmountBySup(Nothing, m_67, , LastDate, DateAdd("D", 1, LastDate))                                                                        'ยอดจ่าย POST DUE หลังเดือน4
      '-------------------------------- จ่ายเช็ค --------------------------------
   'เดือน 7
   End If
   
   If (RunMonth >= 8) Then
   'เดือน 8
      FirstDate = DateAdd("M", 1, FirstDate)
      Call GetFirstLastDate(FirstDate, FirstDate, LastDate)
      '-------------------------------- ยกมา --------------------------------
      Call LoadAPAmountBySup(Nothing, m_ApSupAmounts8, -1, DateAdd("D", -1, FirstDate), "('3','4')")    'ยอดซื้อ ถึงวันที่
      Call LoadAPAmountBySup(Nothing, m_ApSupAmountCns8, -1, DateAdd("D", -1, FirstDate), "('5')")     'ยอดซื้อ ถึงวันที่       ลดหนี้
      Call LoadAPAmountBySup(Nothing, m_ApSupAmountCashs8, FirstDate, LastDate, "('9')")    'ยอดจ่ายเงินสด ถึงวันที่
      Call LoadPaidAmountBySup(Nothing, m_ApSupPaids8, -1, DateAdd("D", -1, FirstDate), -1, -1, "('3','4')")           'ยอดจ่ายถึงวันที่
      Call LoadPaidAmountBySup(Nothing, m_ApSupPaidCns8, -1, DateAdd("D", -1, FirstDate), -1, -1, "('5')")           'ยอดจ่ายถึงวันที่    ลดหนี้
      Call LoadPaidAmountBySup(Nothing, m_PaidAllAmounts8, -1, -1, -1, DateAdd("D", -1, FirstDate), "('3','4')")         'ยอดตัดบิลทั้งหมดของบิลภายในวันที่กำหนด
      Call LoadPaidAmountBySup(Nothing, m_PaidAllAmountCns8, -1, -1, -1, DateAdd("D", -1, FirstDate), "('5')")          'ยอดตัดบิลทั้งหมดของบิลภายในวันที่กำหนด ลดหนี้
      '-------------------------------- ยกมา --------------------------------
      '-------------------------------- เคลื่อนไหว --------------------------------
      Call LoadAPAmountBySup(Nothing, m_ApSupAmountTo8, FirstDate, LastDate, "('3','4')")     'ยอดซื้อ ถึงวันที่
      Call LoadAPAmountBySup(Nothing, m_ApSupAmountToCn8, FirstDate, LastDate, "('5')")     'ลดหนี้ ถึงวันที่
      Call LoadPaidAmountBySup(Nothing, m_ApSupPaidTo8, FirstDate, LastDate, -1, -1, "('3','4')")           'ยอดจ่ายถึงวันที่
      Call LoadPaidAmountBySup(Nothing, m_ApSupPaidToCn8, FirstDate, LastDate, -1, -1, "('5')")           'ยอดจ่ายถึงวันที่
      '-------------------------------- เคลื่อนไหว --------------------------------
      '-------------------------------- จ่ายเช็ค --------------------------------
      Call LoadBktChqAmountBySup(Nothing, m_58, , , FirstDate, LastDate)                                                                                      'ยอดจ่ายจริง DUE เดือน4
      Call LoadBktChqAmountBySup(Nothing, m_68, , LastDate, DateAdd("D", 1, LastDate))                                                                        'ยอดจ่าย POST DUE หลังเดือน4
      '-------------------------------- จ่ายเช็ค --------------------------------
   'เดือน 8
   End If
   
   If (RunMonth >= 9) Then
   'เดือน 9
      FirstDate = DateAdd("M", 1, FirstDate)
      Call GetFirstLastDate(FirstDate, FirstDate, LastDate)
      '-------------------------------- ยกมา --------------------------------
      Call LoadAPAmountBySup(Nothing, m_ApSupAmounts9, -1, DateAdd("D", -1, FirstDate), "('3','4')")    'ยอดซื้อ ถึงวันที่
      Call LoadAPAmountBySup(Nothing, m_ApSupAmountCns9, -1, DateAdd("D", -1, FirstDate), "('5')")     'ยอดซื้อ ถึงวันที่       ลดหนี้
      Call LoadAPAmountBySup(Nothing, m_ApSupAmountCashs9, FirstDate, LastDate, "('9')")    'ยอดจ่ายเงินสด ถึงวันที่
      Call LoadPaidAmountBySup(Nothing, m_ApSupPaids9, -1, DateAdd("D", -1, FirstDate), -1, -1, "('3','4')")           'ยอดจ่ายถึงวันที่
      Call LoadPaidAmountBySup(Nothing, m_ApSupPaidCns9, -1, DateAdd("D", -1, FirstDate), -1, -1, "('5')")           'ยอดจ่ายถึงวันที่    ลดหนี้
      Call LoadPaidAmountBySup(Nothing, m_PaidAllAmounts9, -1, -1, -1, DateAdd("D", -1, FirstDate), "('3','4')")         'ยอดตัดบิลทั้งหมดของบิลภายในวันที่กำหนด
      Call LoadPaidAmountBySup(Nothing, m_PaidAllAmountCns9, -1, -1, -1, DateAdd("D", -1, FirstDate), "('5')")          'ยอดตัดบิลทั้งหมดของบิลภายในวันที่กำหนด ลดหนี้
      '-------------------------------- ยกมา --------------------------------
      '-------------------------------- เคลื่อนไหว --------------------------------
      Call LoadAPAmountBySup(Nothing, m_ApSupAmountTo9, FirstDate, LastDate, "('3','4')")     'ยอดซื้อ ถึงวันที่
      Call LoadAPAmountBySup(Nothing, m_ApSupAmountToCn9, FirstDate, LastDate, "('5')")     'ลดหนี้ ถึงวันที่
      Call LoadPaidAmountBySup(Nothing, m_ApSupPaidTo9, FirstDate, LastDate, -1, -1, "('3','4')")           'ยอดจ่ายถึงวันที่
      Call LoadPaidAmountBySup(Nothing, m_ApSupPaidToCn9, FirstDate, LastDate, -1, -1, "('5')")           'ยอดจ่ายถึงวันที่
      '-------------------------------- เคลื่อนไหว --------------------------------
      '-------------------------------- จ่ายเช็ค --------------------------------
      Call LoadBktChqAmountBySup(Nothing, m_59, , , FirstDate, LastDate)                                                                                      'ยอดจ่ายจริง DUE เดือน4
      Call LoadBktChqAmountBySup(Nothing, m_69, , LastDate, DateAdd("D", 1, LastDate))                                                                        'ยอดจ่าย POST DUE หลังเดือน4
      '-------------------------------- จ่ายเช็ค --------------------------------
   'เดือน 9
   End If
   
   If (RunMonth >= 10) Then
   'เดือน 10
      FirstDate = DateAdd("M", 1, FirstDate)
      Call GetFirstLastDate(FirstDate, FirstDate, LastDate)
      '-------------------------------- ยกมา --------------------------------
      Call LoadAPAmountBySup(Nothing, m_ApSupAmounts10, -1, DateAdd("D", -1, FirstDate), "('3','4')")    'ยอดซื้อ ถึงวันที่
      Call LoadAPAmountBySup(Nothing, m_ApSupAmountCns10, -1, DateAdd("D", -1, FirstDate), "('5')")     'ยอดซื้อ ถึงวันที่       ลดหนี้
      Call LoadAPAmountBySup(Nothing, m_ApSupAmountCashs10, FirstDate, LastDate, "('9')")    'ยอดจ่ายเงินสด ถึงวันที่
      Call LoadPaidAmountBySup(Nothing, m_ApSupPaids10, -1, DateAdd("D", -1, FirstDate), -1, -1, "('3','4')")           'ยอดจ่ายถึงวันที่
      Call LoadPaidAmountBySup(Nothing, m_ApSupPaidCns10, -1, DateAdd("D", -1, FirstDate), -1, -1, "('5')")           'ยอดจ่ายถึงวันที่    ลดหนี้
      Call LoadPaidAmountBySup(Nothing, m_PaidAllAmounts10, -1, -1, -1, DateAdd("D", -1, FirstDate), "('3','4')")         'ยอดตัดบิลทั้งหมดของบิลภายในวันที่กำหนด
      Call LoadPaidAmountBySup(Nothing, m_PaidAllAmountCns10, -1, -1, -1, DateAdd("D", -1, FirstDate), "('5')")          'ยอดตัดบิลทั้งหมดของบิลภายในวันที่กำหนด ลดหนี้
      '-------------------------------- ยกมา --------------------------------
      '-------------------------------- เคลื่อนไหว --------------------------------
      Call LoadAPAmountBySup(Nothing, m_ApSupAmountTo10, FirstDate, LastDate, "('3','4')")     'ยอดซื้อ ถึงวันที่
      Call LoadAPAmountBySup(Nothing, m_ApSupAmountToCn10, FirstDate, LastDate, "('5')")     'ลดหนี้ ถึงวันที่
      Call LoadPaidAmountBySup(Nothing, m_ApSupPaidTo10, FirstDate, LastDate, -1, -1, "('3','4')")           'ยอดจ่ายถึงวันที่
      Call LoadPaidAmountBySup(Nothing, m_ApSupPaidToCn10, FirstDate, LastDate, -1, -1, "('5')")           'ยอดจ่ายถึงวันที่
      '-------------------------------- เคลื่อนไหว --------------------------------
      '-------------------------------- จ่ายเช็ค --------------------------------
      Call LoadBktChqAmountBySup(Nothing, m_510, , , FirstDate, LastDate)                                                                                      'ยอดจ่ายจริง DUE เดือน4
      Call LoadBktChqAmountBySup(Nothing, m_610, , LastDate, DateAdd("D", 1, LastDate))                                                                        'ยอดจ่าย POST DUE หลังเดือน4
      '-------------------------------- จ่ายเช็ค --------------------------------
   'เดือน 10
   End If
   
   If (RunMonth >= 11) Then
   'เดือน 11
      FirstDate = DateAdd("M", 1, FirstDate)
      Call GetFirstLastDate(FirstDate, FirstDate, LastDate)
      '-------------------------------- ยกมา --------------------------------
      Call LoadAPAmountBySup(Nothing, m_ApSupAmounts11, -1, DateAdd("D", -1, FirstDate), "('3','4')")    'ยอดซื้อ ถึงวันที่
      Call LoadAPAmountBySup(Nothing, m_ApSupAmountCns11, -1, DateAdd("D", -1, FirstDate), "('5')")     'ยอดซื้อ ถึงวันที่       ลดหนี้
      Call LoadAPAmountBySup(Nothing, m_ApSupAmountCashs11, FirstDate, LastDate, "('9')")    'ยอดจ่ายเงินสด ถึงวันที่
      Call LoadPaidAmountBySup(Nothing, m_ApSupPaids11, -1, DateAdd("D", -1, FirstDate), -1, -1, "('3','4')")           'ยอดจ่ายถึงวันที่
      Call LoadPaidAmountBySup(Nothing, m_ApSupPaidCns11, -1, DateAdd("D", -1, FirstDate), -1, -1, "('5')")           'ยอดจ่ายถึงวันที่    ลดหนี้
      Call LoadPaidAmountBySup(Nothing, m_PaidAllAmounts11, -1, -1, -1, DateAdd("D", -1, FirstDate), "('3','4')")         'ยอดตัดบิลทั้งหมดของบิลภายในวันที่กำหนด
      Call LoadPaidAmountBySup(Nothing, m_PaidAllAmountCns11, -1, -1, -1, DateAdd("D", -1, FirstDate), "('5')")          'ยอดตัดบิลทั้งหมดของบิลภายในวันที่กำหนด ลดหนี้
      '-------------------------------- ยกมา --------------------------------
      '-------------------------------- เคลื่อนไหว --------------------------------
      Call LoadAPAmountBySup(Nothing, m_ApSupAmountTo11, FirstDate, LastDate, "('3','4')")     'ยอดซื้อ ถึงวันที่
      Call LoadAPAmountBySup(Nothing, m_ApSupAmountToCn11, FirstDate, LastDate, "('5')")     'ลดหนี้ ถึงวันที่
      Call LoadPaidAmountBySup(Nothing, m_ApSupPaidTo11, FirstDate, LastDate, -1, -1, "('3','4')")           'ยอดจ่ายถึงวันที่
      Call LoadPaidAmountBySup(Nothing, m_ApSupPaidToCn11, FirstDate, LastDate, -1, -1, "('5')")           'ยอดจ่ายถึงวันที่
      '-------------------------------- เคลื่อนไหว --------------------------------
      '-------------------------------- จ่ายเช็ค --------------------------------
      Call LoadBktChqAmountBySup(Nothing, m_511, , , FirstDate, LastDate)                                                                                      'ยอดจ่ายจริง DUE เดือน4
      Call LoadBktChqAmountBySup(Nothing, m_611, , LastDate, DateAdd("D", 1, LastDate))                                                                        'ยอดจ่าย POST DUE หลังเดือน4
      '-------------------------------- จ่ายเช็ค --------------------------------
   'เดือน 11
   End If
   
   If (RunMonth >= 12) Then
   'เดือน 12
      FirstDate = DateAdd("M", 1, FirstDate)
      Call GetFirstLastDate(FirstDate, FirstDate, LastDate)
      '-------------------------------- ยกมา --------------------------------
      Call LoadAPAmountBySup(Nothing, m_ApSupAmounts12, -1, DateAdd("D", -1, FirstDate), "('3','4')")    'ยอดซื้อ ถึงวันที่
      Call LoadAPAmountBySup(Nothing, m_ApSupAmountCns12, -1, DateAdd("D", -1, FirstDate), "('5')")     'ยอดซื้อ ถึงวันที่       ลดหนี้
      Call LoadAPAmountBySup(Nothing, m_ApSupAmountCashs12, FirstDate, LastDate, "('9')")    'ยอดจ่ายเงินสด ถึงวันที่
      Call LoadPaidAmountBySup(Nothing, m_ApSupPaids12, -1, DateAdd("D", -1, FirstDate), -1, -1, "('3','4')")           'ยอดจ่ายถึงวันที่
      Call LoadPaidAmountBySup(Nothing, m_ApSupPaidCns12, -1, DateAdd("D", -1, FirstDate), -1, -1, "('5')")           'ยอดจ่ายถึงวันที่    ลดหนี้
      Call LoadPaidAmountBySup(Nothing, m_PaidAllAmounts12, -1, -1, -1, DateAdd("D", -1, FirstDate), "('3','4')")         'ยอดตัดบิลทั้งหมดของบิลภายในวันที่กำหนด
      Call LoadPaidAmountBySup(Nothing, m_PaidAllAmountCns12, -1, -1, -1, DateAdd("D", -1, FirstDate), "('5')")          'ยอดตัดบิลทั้งหมดของบิลภายในวันที่กำหนด ลดหนี้
      '-------------------------------- ยกมา --------------------------------
      '-------------------------------- เคลื่อนไหว --------------------------------
      Call LoadAPAmountBySup(Nothing, m_ApSupAmountTo12, FirstDate, LastDate, "('3','4')")     'ยอดซื้อ ถึงวันที่
      Call LoadAPAmountBySup(Nothing, m_ApSupAmountToCn12, FirstDate, LastDate, "('5')")     'ลดหนี้ ถึงวันที่
      Call LoadPaidAmountBySup(Nothing, m_ApSupPaidTo12, FirstDate, LastDate, -1, -1, "('3','4')")           'ยอดจ่ายถึงวันที่
      Call LoadPaidAmountBySup(Nothing, m_ApSupPaidToCn12, FirstDate, LastDate, -1, -1, "('5')")           'ยอดจ่ายถึงวันที่
      '-------------------------------- เคลื่อนไหว --------------------------------
      '-------------------------------- จ่ายเช็ค --------------------------------
      Call LoadBktChqAmountBySup(Nothing, m_512, , , FirstDate, LastDate)                                                                                      'ยอดจ่ายจริง DUE เดือน4
      Call LoadBktChqAmountBySup(Nothing, m_612, , LastDate, DateAdd("D", 1, LastDate))                                                                        'ยอดจ่าย POST DUE หลังเดือน4
      '-------------------------------- จ่ายเช็ค --------------------------------
   'เดือน 12
   End If
   
   MODULE_DESC = glbCompanyName & vbCrLf & mcolParam("REPORT_NAME") & _
                                       vbCrLf & "งวด " & mcolParam("MONTH_ID") & "/" & mcolParam("YEAR_NO") & " เดือนเคลื่อนไหว " & mcolParam("RUN_MONTH") & " เดือน " & mcolParam("DATA_TYPE_NAME")
   
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

Dim TempValue As Double

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
'----------------------------------------------------------------------------------------------------

   Call StartExportFile(Vsp)
   Vsp.StartDoc
   
   Set BodyCf = New CReportField
      
   i = 0
   For Each SupGroup In SupplierGroupColl
'      If SupGroup.GROUP_TYPE_CODE <= 0 Then
'         'debug.print
'      End If
      If (i >= 0 And PrevKey1 <> Trim(Str(SupGroup.GROUP_TYPE_CODE))) Then
            
         If i > 0 Then
            Call GenerateFooterSupplierGroup(PrevKey2)
         End If
         
         Call m_Details1.ClearField
         Width = m_Headers0.Fields(1).Width
         TempStr = ""
         Call BodyCf.SetFieldValue(Width, "^", TempStr)
         Call m_Details1.AddField(BodyCf)
         
         Width = m_Headers0.Fields(2).Width
         If SupGroup.GROUP_TYPE_CODE > 0 Then
            TempStr = SupGroup.GROUP_TYPE_CODE
         Else
            TempStr = "ยังไม่ได้สร้างกลุ่ม"
         End If
         Call BodyCf.SetFieldValue(Width, "^", TempStr)
         Call m_Details1.AddField(BodyCf)
         
         Width = m_Headers0.Fields(3).Width
         TempStr = SupGroup.GROUP_TYPE_NAME
         Call BodyCf.SetFieldValue(Width, "<", TempStr)
         Call m_Details1.AddField(BodyCf)

         If mcolParam("SUMMARY_FLAG") = "N" Then
            Call m_Details1.GetString(1, TempStr1, TempStr2)
            strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
            Vsp.FontBold = True
            Call Vsp.AddTable(strFormat, "", TempStr2)
            Vsp.FontBold = False
         End If
      End If
      
      HaveValueSum = False
      
      PrevKey1 = SupGroup.GROUP_TYPE_CODE
      PrevKey2 = SupGroup.GROUP_TYPE_NAME
      
      i = i + 1
      
      ' ของเดือนแรกที่เลือก 1
      Call m_Details1.ClearField
            
      Set TempApt = GetAPTrn(m_ApSupAmounts, SupGroup.SUPPLIER_CODE)
      Set TempAptCn = GetAPTrn(m_ApSupAmountCns, SupGroup.SUPPLIER_CODE)
      
      Set TempApRcIt = GetAPRcpItem(m_ApSupPaids, SupGroup.SUPPLIER_CODE)
      Set TempApRcItCn = GetAPRcpItem(m_ApSupPaidCns, SupGroup.SUPPLIER_CODE)
         
      Set TempApRcIt3 = GetAPRcpItemEx(m_PaidAllAmounts, SupGroup.SUPPLIER_CODE)
      Set TempApRcIt3Cn = GetAPRcpItemEx(m_PaidAllAmountCns, SupGroup.SUPPLIER_CODE)
               
      HaveValue = False
      FirstDate = DateSerial(mcolParam("YEAR_NO") - 543, mcolParam("MONTH_ID"), 1)
       '===Head
       Set HeadCf = m_Headers1.Fields(1)
       TempStr = Format(Month(FirstDate), "00") & "/" & Format(Year(FirstDate) + 543, "00")
       Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
       Call m_Details1.AddField(BodyCf)
       
       Set Ap = GetObject("CAPMas", Supplier, SupGroup.SUPPLIER_CODE)
       Set HeadCf = m_Headers1.Fields(2)
      TempStr = Left(Ap.SUPCOD & "/" & Ap.SUPNAM, 30)
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
       Total3(K) = Total3(K) + BalanceAmount
       If Round((BalanceAmount), 2) <> 0 Then
          HaveValue = True
       End If
       
       Set TempAptTo = GetAPTrn(m_ApSupAmountTo, SupGroup.SUPPLIER_CODE)
       Set TempAptToCn = GetAPTrn(m_ApSupAmountToCn, SupGroup.SUPPLIER_CODE)
          
       Set TempApRcItTo = GetAPRcpItem(m_ApSupPaidTo, SupGroup.SUPPLIER_CODE)
       Set TempApRcItToCn = GetAPRcpItem(m_ApSupPaidToCn, SupGroup.SUPPLIER_CODE)
                
       K = K + 1
       Set HeadCf = m_Headers1.Fields(K)
       TempStr = FormatNumber(TempAptTo.AMOUNT, "")
       Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
       Call m_Details1.AddField(BodyCf)
       
       Total1(K) = Total1(K) + TempAptTo.AMOUNT
       Total3(K) = Total3(K) + TempAptTo.AMOUNT
       If Round((TempAptTo.AMOUNT), 2) <> 0 Then
          HaveValue = True
       End If
       
       K = K + 1
       Set HeadCf = m_Headers1.Fields(K)
       TempStr = FormatNumber(TempApRcItTo.PAYAMT - TempApRcItToCn.PAYAMT + TempAptToCn.AMOUNT, "")
       Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
       Call m_Details1.AddField(BodyCf)
       
       Total1(K) = Total1(K) + TempApRcItTo.PAYAMT - TempApRcItToCn.PAYAMT + TempAptToCn.AMOUNT
       Total3(K) = Total3(K) + TempApRcItTo.PAYAMT - TempApRcItToCn.PAYAMT + TempAptToCn.AMOUNT
       If Round((TempApRcItTo.PAYAMT - TempApRcItToCn.PAYAMT + TempAptToCn.AMOUNT), 2) <> 0 Then
          HaveValue = True
       End If
       TempPaidCheck = Round((TempApRcItTo.PAYAMT - TempApRcItToCn.PAYAMT + TempAptToCn.AMOUNT), 2)
       
       K = K + 1
       Set HeadCf = m_Headers1.Fields(K)
       TempStr = FormatNumber(BalanceAmount + (TempAptTo.AMOUNT - TempAptToCn.AMOUNT) - (TempApRcItTo.PAYAMT - TempApRcItToCn.PAYAMT), "")
       Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
       Call m_Details1.AddField(BodyCf)
       
       Total1(K) = Total1(K) + (BalanceAmount) + (TempAptTo.AMOUNT - TempAptToCn.AMOUNT) - (TempApRcItTo.PAYAMT - TempApRcItToCn.PAYAMT)
       Total3(K) = Total3(K) + (BalanceAmount) + (TempAptTo.AMOUNT - TempAptToCn.AMOUNT) - (TempApRcItTo.PAYAMT - TempApRcItToCn.PAYAMT)
       If Round(((BalanceAmount) + (TempAptTo.AMOUNT - TempAptToCn.AMOUNT) - (TempApRcItTo.PAYAMT - TempApRcItToCn.PAYAMT)), 2) <> 0 Then
          HaveValue = True
       End If
       
       Set TempApBktTrn = GetBkTrn(m_5, SupGroup.SUPPLIER_CODE)
       Set TempAptCash = GetAPTrn(m_ApSupAmountCashs, SupGroup.SUPPLIER_CODE)
       
       K = K + 1
       Set HeadCf = m_Headers1.Fields(K)
       TempStr = FormatNumber(TempApBktTrn.AMOUNT + TempAptCash.CSHPAY, "")
       Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
       Call m_Details1.AddField(BodyCf)
       
       Total1(K) = Total1(K) + TempApBktTrn.AMOUNT + TempAptCash.CSHPAY
       Total3(K) = Total3(K) + TempApBktTrn.AMOUNT + TempAptCash.CSHPAY
       If Round((TempApBktTrn.AMOUNT + TempAptCash.CSHPAY), 2) <> 0 Then
          HaveValue = True
       End If
       
       Set TempApBktTrn = GetBkTrn(m_6, SupGroup.SUPPLIER_CODE)
       K = K + 1
       Set HeadCf = m_Headers1.Fields(K)
       TempStr = FormatNumber(TempApBktTrn.AMOUNT, "")
       Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
       Call m_Details1.AddField(BodyCf)
       
       Total1(K) = Total1(K) + TempApBktTrn.AMOUNT
       Total3(K) = Total3(K) + TempApBktTrn.AMOUNT
       If Round((TempApBktTrn.AMOUNT), 2) <> 0 Then
          HaveValue = True
       End If
       
       '7 =  4 + 6
       K = K + 1
       Set HeadCf = m_Headers1.Fields(K)
       TempStr = FormatNumber(BalanceAmount + (TempAptTo.AMOUNT - TempAptToCn.AMOUNT) - (TempApRcItTo.PAYAMT - TempApRcItToCn.PAYAMT) + TempApBktTrn.AMOUNT, "")
       Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
       Call m_Details1.AddField(BodyCf)
       
       Total1(K) = Total1(K) + BalanceAmount + (TempAptTo.AMOUNT - TempAptToCn.AMOUNT) - (TempApRcItTo.PAYAMT - TempApRcItToCn.PAYAMT) + TempApBktTrn.AMOUNT
       Total3(K) = Total3(K) + BalanceAmount + (TempAptTo.AMOUNT - TempAptToCn.AMOUNT) - (TempApRcItTo.PAYAMT - TempApRcItToCn.PAYAMT) + TempApBktTrn.AMOUNT
       If Round((BalanceAmount + (TempAptTo.AMOUNT - TempAptToCn.AMOUNT) - (TempApRcItTo.PAYAMT - TempApRcItToCn.PAYAMT) + TempApBktTrn.AMOUNT), 2) <> 0 Then
          HaveValue = True
       End If
       
      Set TempApBktTrn = GetBkTrn(m_5, SupGroup.SUPPLIER_CODE)
       Set TempAptCash = GetAPTrn(m_ApSupAmountCashs, SupGroup.SUPPLIER_CODE)
       
       K = K + 1
       Set HeadCf = m_Headers1.Fields(K)
       TempStr = FormatNumber(TempAptTo.AMOUNT - TempApBktTrn.AMOUNT - TempAptCash.CSHPAY, "")
       Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
       Call m_Details1.AddField(BodyCf)
       
       Total1(K) = Total1(K) + TempAptTo.AMOUNT - TempApBktTrn.AMOUNT - TempAptCash.CSHPAY
       Total3(K) = Total3(K) + TempAptTo.AMOUNT - TempApBktTrn.AMOUNT - TempAptCash.CSHPAY
       If Round((TempAptTo.AMOUNT - TempApBktTrn.AMOUNT - TempAptCash.CSHPAY), 2) <> 0 Then
          HaveValue = True
       End If
       
       If HaveValue Then
         If mcolParam("SUMMARY_FLAG") = "N" And mcolParam("ONLY_SUPPLIER") = "N" Then
            Call m_Details1.GetString(1, TempStr1, TempStr2)
            strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
            Call Vsp.AddTable(strFormat, "", TempStr2)
          End If
          HaveValueSum = True
       Else
          i = i - 1
       End If
       ' ของเดือนแรกที่เลือก 1
       
      If (RunMonth >= 2) Then
         ' ของเดือนแรกที่เลือก 2
         
         FirstDate = DateAdd("M", 1, FirstDate)
         Call m_Details1.ClearField
               
         Set TempApt = GetAPTrn(m_ApSupAmounts2, SupGroup.SUPPLIER_CODE)
         Set TempAptCn = GetAPTrn(m_ApSupAmountCns2, SupGroup.SUPPLIER_CODE)
         
         Set TempApRcIt = GetAPRcpItem(m_ApSupPaids2, SupGroup.SUPPLIER_CODE)
         Set TempApRcItCn = GetAPRcpItem(m_ApSupPaidCns2, SupGroup.SUPPLIER_CODE)
            
         Set TempApRcIt3 = GetAPRcpItemEx(m_PaidAllAmounts2, SupGroup.SUPPLIER_CODE)
         Set TempApRcIt3Cn = GetAPRcpItemEx(m_PaidAllAmountCns2, SupGroup.SUPPLIER_CODE)
         
         HaveValue = False
               
          '===Head
          Set HeadCf = m_Headers1.Fields(1)
          TempStr = Format(Month(FirstDate), "00") & "/" & Format(Year(FirstDate) + 543, "00")
          Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
          Call m_Details1.AddField(BodyCf)
          
          Set Ap = GetObject("CAPMas", Supplier, SupGroup.SUPPLIER_CODE)
          Set HeadCf = m_Headers1.Fields(2)
          TempStr = Left(Ap.SUPCOD & "/" & Ap.SUPNAM, 30)
          Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
          Call m_Details1.AddField(BodyCf)
          
          BalanceAmount = (TempApt.AMOUNT - TempAptCn.AMOUNT) - (TempApt.PAYAMT - TempAptCn.PAYAMT) - (TempApRcIt.PAYAMT - TempApRcItCn.PAYAMT) + (TempApRcIt3.PAYAMT - TempApRcIt3Cn.PAYAMT)
          
          K = 2
          K = K + 1
          Set HeadCf = m_Headers1.Fields(K)
          TempStr = FormatNumber(BalanceAmount, "")
          Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
          Call m_Details1.AddField(BodyCf)
          
          Total12(K) = Total12(K) + BalanceAmount
          If Round((BalanceAmount), 2) <> 0 Then
             HaveValue = True
          End If
          
          Set TempAptTo = GetAPTrn(m_ApSupAmountTo2, SupGroup.SUPPLIER_CODE)
          Set TempAptToCn = GetAPTrn(m_ApSupAmountToCn2, SupGroup.SUPPLIER_CODE)
          
          Set TempApRcItTo = GetAPRcpItem(m_ApSupPaidTo2, SupGroup.SUPPLIER_CODE)
          Set TempApRcItToCn = GetAPRcpItem(m_ApSupPaidToCn2, SupGroup.SUPPLIER_CODE)
                   
          K = K + 1
          Set HeadCf = m_Headers1.Fields(K)
          TempStr = FormatNumber(TempAptTo.AMOUNT, "")
          Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
          Call m_Details1.AddField(BodyCf)
          
          Total12(K) = Total12(K) + TempAptTo.AMOUNT
          Total3(K) = Total3(K) + TempAptTo.AMOUNT
          If Round((TempAptTo.AMOUNT), 2) <> 0 Then
             HaveValue = True
          End If
          
          K = K + 1
          Set HeadCf = m_Headers1.Fields(K)
          TempStr = FormatNumber(TempApRcItTo.PAYAMT - TempApRcItToCn.PAYAMT + TempAptToCn.AMOUNT, "")
          Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
          Call m_Details1.AddField(BodyCf)
          
          Total12(K) = Total12(K) + TempApRcItTo.PAYAMT - TempApRcItToCn.PAYAMT + TempAptToCn.AMOUNT
          Total3(K) = Total3(K) + TempApRcItTo.PAYAMT - TempApRcItToCn.PAYAMT + TempAptToCn.AMOUNT
          If Round((TempApRcItTo.PAYAMT - TempApRcItToCn.PAYAMT + TempAptToCn.AMOUNT), 2) <> 0 Then
             HaveValue = True
          End If
          TempPaidCheck = Round((TempApRcItTo.PAYAMT - TempApRcItToCn.PAYAMT + TempAptToCn.AMOUNT), 2)
          
          K = K + 1
          Set HeadCf = m_Headers1.Fields(K)
          TempStr = FormatNumber(BalanceAmount + (TempAptTo.AMOUNT - TempAptToCn.AMOUNT) - (TempApRcItTo.PAYAMT - TempApRcItToCn.PAYAMT), "")
          Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
          Call m_Details1.AddField(BodyCf)
          
          Total12(K) = Total12(K) + (BalanceAmount) + (TempAptTo.AMOUNT - TempAptToCn.AMOUNT) - (TempApRcItTo.PAYAMT - TempApRcItToCn.PAYAMT)
          If Round(((BalanceAmount) + (TempAptTo.AMOUNT - TempAptToCn.AMOUNT) - (TempApRcItTo.PAYAMT - TempApRcItToCn.PAYAMT)), 2) <> 0 Then
             HaveValue = True
          End If
          
          Set TempApBktTrn = GetBkTrn(m_52, SupGroup.SUPPLIER_CODE)
          Set TempAptCash = GetAPTrn(m_ApSupAmountCashs2, SupGroup.SUPPLIER_CODE)
          
          K = K + 1
          Set HeadCf = m_Headers1.Fields(K)
          TempStr = FormatNumber(TempApBktTrn.AMOUNT + TempAptCash.CSHPAY, "")
          Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
          Call m_Details1.AddField(BodyCf)
          
          Total12(K) = Total12(K) + TempApBktTrn.AMOUNT + TempAptCash.CSHPAY
          Total3(K) = Total3(K) + TempApBktTrn.AMOUNT + TempAptCash.CSHPAY
          If Round((TempApBktTrn.AMOUNT + TempAptCash.CSHPAY), 2) <> 0 Then
             HaveValue = True
          End If
          
          Set TempApBktTrn = GetBkTrn(m_62, SupGroup.SUPPLIER_CODE)
          K = K + 1
          Set HeadCf = m_Headers1.Fields(K)
          TempStr = FormatNumber(TempApBktTrn.AMOUNT, "")
          Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
          Call m_Details1.AddField(BodyCf)
          
          Total12(K) = Total12(K) + TempApBktTrn.AMOUNT
          Total3(K) = TempApBktTrn.AMOUNT
          If Round((TempApBktTrn.AMOUNT), 2) <> 0 Then
             HaveValue = True
          End If
          
          '7 =  4 + 6
          K = K + 1
          Set HeadCf = m_Headers1.Fields(K)
          TempStr = FormatNumber(BalanceAmount + (TempAptTo.AMOUNT - TempAptToCn.AMOUNT) - (TempApRcItTo.PAYAMT - TempApRcItToCn.PAYAMT) + TempApBktTrn.AMOUNT, "")
          Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
          Call m_Details1.AddField(BodyCf)
          
          Total12(K) = Total12(K) + BalanceAmount + (TempAptTo.AMOUNT - TempAptToCn.AMOUNT) - (TempApRcItTo.PAYAMT - TempApRcItToCn.PAYAMT) + TempApBktTrn.AMOUNT
          Total3(K) = BalanceAmount + (TempAptTo.AMOUNT - TempAptToCn.AMOUNT) - (TempApRcItTo.PAYAMT - TempApRcItToCn.PAYAMT) + TempApBktTrn.AMOUNT
          If Round((BalanceAmount + (TempAptTo.AMOUNT - TempAptToCn.AMOUNT) - (TempApRcItTo.PAYAMT - TempApRcItToCn.PAYAMT) + TempApBktTrn.AMOUNT), 2) <> 0 Then
             HaveValue = True
          End If
          
         Set TempApBktTrn = GetBkTrn(m_52, SupGroup.SUPPLIER_CODE)
          Set TempAptCash = GetAPTrn(m_ApSupAmountCashs2, SupGroup.SUPPLIER_CODE)
          
          K = K + 1
          Set HeadCf = m_Headers1.Fields(K)
          TempStr = FormatNumber(TempAptTo.AMOUNT - TempApBktTrn.AMOUNT - TempAptCash.CSHPAY, "")
          Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
          Call m_Details1.AddField(BodyCf)
          
          Total12(K) = Total12(K) + TempAptTo.AMOUNT - TempApBktTrn.AMOUNT - TempAptCash.CSHPAY
          Total3(K) = Total3(K) + TempAptTo.AMOUNT - TempApBktTrn.AMOUNT - TempAptCash.CSHPAY
          If Round((TempAptTo.AMOUNT - TempApBktTrn.AMOUNT - TempAptCash.CSHPAY), 2) <> 0 Then
             HaveValue = True
          End If
          
          If HaveValue Then
            If mcolParam("SUMMARY_FLAG") = "N" And mcolParam("ONLY_SUPPLIER") = "N" Then
               Call m_Details1.GetString(1, TempStr1, TempStr2)
               strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
               Call Vsp.AddTable(strFormat, "", TempStr2)
             End If
             HaveValueSum = True
          End If
          ' ของเดือนแรกที่เลือก 2
      End If
      
      If (RunMonth >= 3) Then
         ' ของเดือนแรกที่เลือก 3
         
         FirstDate = DateAdd("M", 1, FirstDate)
         Call m_Details1.ClearField
               
         Set TempApt = GetAPTrn(m_ApSupAmounts3, SupGroup.SUPPLIER_CODE)
         Set TempAptCn = GetAPTrn(m_ApSupAmountCns3, SupGroup.SUPPLIER_CODE)
         
         Set TempApRcIt = GetAPRcpItem(m_ApSupPaids3, SupGroup.SUPPLIER_CODE)
         Set TempApRcItCn = GetAPRcpItem(m_ApSupPaidCns3, SupGroup.SUPPLIER_CODE)
            
         Set TempApRcIt3 = GetAPRcpItemEx(m_PaidAllAmounts3, SupGroup.SUPPLIER_CODE)
         Set TempApRcIt3Cn = GetAPRcpItemEx(m_PaidAllAmountCns3, SupGroup.SUPPLIER_CODE)
         
         HaveValue = False
               
          '===Head
          Set HeadCf = m_Headers1.Fields(1)
          TempStr = Format(Month(FirstDate), "00") & "/" & Format(Year(FirstDate) + 543, "00")
          Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
          Call m_Details1.AddField(BodyCf)
          
          Set Ap = GetObject("CAPMas", Supplier, SupGroup.SUPPLIER_CODE)
          Set HeadCf = m_Headers1.Fields(2)
          TempStr = Left(Ap.SUPCOD & "/" & Ap.SUPNAM, 30)
          Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
          Call m_Details1.AddField(BodyCf)
          
          BalanceAmount = (TempApt.AMOUNT - TempAptCn.AMOUNT) - (TempApt.PAYAMT - TempAptCn.PAYAMT) - (TempApRcIt.PAYAMT - TempApRcItCn.PAYAMT) + (TempApRcIt3.PAYAMT - TempApRcIt3Cn.PAYAMT)
          
          K = 2
          K = K + 1
          Set HeadCf = m_Headers1.Fields(K)
          TempStr = FormatNumber(BalanceAmount, "")
          Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
          Call m_Details1.AddField(BodyCf)
          
          Total13(K) = Total13(K) + BalanceAmount
          If Round((BalanceAmount), 2) <> 0 Then
             HaveValue = True
          End If
          
          Set TempAptTo = GetAPTrn(m_ApSupAmountTo3, SupGroup.SUPPLIER_CODE)
          Set TempAptToCn = GetAPTrn(m_ApSupAmountToCn3, SupGroup.SUPPLIER_CODE)
          
          Set TempApRcItTo = GetAPRcpItem(m_ApSupPaidTo3, SupGroup.SUPPLIER_CODE)
          Set TempApRcItToCn = GetAPRcpItem(m_ApSupPaidToCn3, SupGroup.SUPPLIER_CODE)
                   
          K = K + 1
          Set HeadCf = m_Headers1.Fields(K)
          TempStr = FormatNumber(TempAptTo.AMOUNT, "")
          Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
          Call m_Details1.AddField(BodyCf)
          
          Total13(K) = Total13(K) + TempAptTo.AMOUNT
          Total3(K) = Total3(K) + TempAptTo.AMOUNT
          If Round((TempAptTo.AMOUNT), 2) <> 0 Then
             HaveValue = True
          End If
          
          K = K + 1
          Set HeadCf = m_Headers1.Fields(K)
          TempStr = FormatNumber(TempApRcItTo.PAYAMT - TempApRcItToCn.PAYAMT + TempAptToCn.AMOUNT, "")
          Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
          Call m_Details1.AddField(BodyCf)
          
          Total13(K) = Total13(K) + TempApRcItTo.PAYAMT - TempApRcItToCn.PAYAMT + TempAptToCn.AMOUNT
          Total3(K) = Total3(K) + TempApRcItTo.PAYAMT - TempApRcItToCn.PAYAMT + TempAptToCn.AMOUNT
          If Round((TempApRcItTo.PAYAMT - TempApRcItToCn.PAYAMT + TempAptToCn.AMOUNT), 2) <> 0 Then
             HaveValue = True
          End If
          TempPaidCheck = Round((TempApRcItTo.PAYAMT - TempApRcItToCn.PAYAMT + TempAptToCn.AMOUNT), 2)
          
          K = K + 1
          Set HeadCf = m_Headers1.Fields(K)
          TempStr = FormatNumber(BalanceAmount + (TempAptTo.AMOUNT - TempAptToCn.AMOUNT) - (TempApRcItTo.PAYAMT - TempApRcItToCn.PAYAMT), "")
          Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
          Call m_Details1.AddField(BodyCf)
          
          Total13(K) = Total13(K) + (BalanceAmount) + (TempAptTo.AMOUNT - TempAptToCn.AMOUNT) - (TempApRcItTo.PAYAMT - TempApRcItToCn.PAYAMT)
          If Round(((BalanceAmount) + (TempAptTo.AMOUNT - TempAptToCn.AMOUNT) - (TempApRcItTo.PAYAMT - TempApRcItToCn.PAYAMT)), 2) <> 0 Then
             HaveValue = True
          End If
          
          Set TempApBktTrn = GetBkTrn(m_53, SupGroup.SUPPLIER_CODE)
          Set TempAptCash = GetAPTrn(m_ApSupAmountCashs3, SupGroup.SUPPLIER_CODE)
          
          K = K + 1
          Set HeadCf = m_Headers1.Fields(K)
          TempStr = FormatNumber(TempApBktTrn.AMOUNT + TempAptCash.CSHPAY, "")
          Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
          Call m_Details1.AddField(BodyCf)
          
          Total13(K) = Total13(K) + TempApBktTrn.AMOUNT + TempAptCash.CSHPAY
          Total3(K) = Total3(K) + TempApBktTrn.AMOUNT + TempAptCash.CSHPAY
          If Round((TempApBktTrn.AMOUNT + TempAptCash.CSHPAY), 2) <> 0 Then
             HaveValue = True
          End If
          
          Set TempApBktTrn = GetBkTrn(m_63, SupGroup.SUPPLIER_CODE)
          K = K + 1
          Set HeadCf = m_Headers1.Fields(K)
          TempStr = FormatNumber(TempApBktTrn.AMOUNT, "")
          Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
          Call m_Details1.AddField(BodyCf)
          
          Total13(K) = Total13(K) + TempApBktTrn.AMOUNT
          Total3(K) = TempApBktTrn.AMOUNT
          If Round((TempApBktTrn.AMOUNT), 2) <> 0 Then
             HaveValue = True
          End If
          
          '7 =  4 + 6
          K = K + 1
          Set HeadCf = m_Headers1.Fields(K)
          TempStr = FormatNumber(BalanceAmount + (TempAptTo.AMOUNT - TempAptToCn.AMOUNT) - (TempApRcItTo.PAYAMT - TempApRcItToCn.PAYAMT) + TempApBktTrn.AMOUNT, "")
          Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
          Call m_Details1.AddField(BodyCf)
          
          Total13(K) = Total13(K) + BalanceAmount + (TempAptTo.AMOUNT - TempAptToCn.AMOUNT) - (TempApRcItTo.PAYAMT - TempApRcItToCn.PAYAMT) + TempApBktTrn.AMOUNT
          Total3(K) = BalanceAmount + (TempAptTo.AMOUNT - TempAptToCn.AMOUNT) - (TempApRcItTo.PAYAMT - TempApRcItToCn.PAYAMT) + TempApBktTrn.AMOUNT
          If Round((BalanceAmount + (TempAptTo.AMOUNT - TempAptToCn.AMOUNT) - (TempApRcItTo.PAYAMT - TempApRcItToCn.PAYAMT) + TempApBktTrn.AMOUNT), 2) <> 0 Then
             HaveValue = True
          End If
          
         Set TempApBktTrn = GetBkTrn(m_53, SupGroup.SUPPLIER_CODE)
          Set TempAptCash = GetAPTrn(m_ApSupAmountCashs3, SupGroup.SUPPLIER_CODE)
          
          K = K + 1
          Set HeadCf = m_Headers1.Fields(K)
          TempStr = FormatNumber(TempAptTo.AMOUNT - TempApBktTrn.AMOUNT - TempAptCash.CSHPAY, "")
          Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
          Call m_Details1.AddField(BodyCf)
          
          Total13(K) = Total13(K) + TempAptTo.AMOUNT - TempApBktTrn.AMOUNT - TempAptCash.CSHPAY
          Total3(K) = Total3(K) + TempAptTo.AMOUNT - TempApBktTrn.AMOUNT - TempAptCash.CSHPAY
          If Round((TempAptTo.AMOUNT - TempApBktTrn.AMOUNT - TempAptCash.CSHPAY), 2) <> 0 Then
             HaveValue = True
          End If
          
          If HaveValue Then
            If mcolParam("SUMMARY_FLAG") = "N" And mcolParam("ONLY_SUPPLIER") = "N" Then
               Call m_Details1.GetString(1, TempStr1, TempStr2)
               strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
               Call Vsp.AddTable(strFormat, "", TempStr2)
             End If
             HaveValueSum = True
          End If
          ' ของเดือนแรกที่เลือก 3
      End If
      
      If (RunMonth >= 4) Then
         ' ของเดือนแรกที่เลือก 4
         
         FirstDate = DateAdd("M", 1, FirstDate)
         Call m_Details1.ClearField
               
         Set TempApt = GetAPTrn(m_ApSupAmounts4, SupGroup.SUPPLIER_CODE)
         Set TempAptCn = GetAPTrn(m_ApSupAmountCns4, SupGroup.SUPPLIER_CODE)
         
         Set TempApRcIt = GetAPRcpItem(m_ApSupPaids4, SupGroup.SUPPLIER_CODE)
         Set TempApRcItCn = GetAPRcpItem(m_ApSupPaidCns4, SupGroup.SUPPLIER_CODE)
            
         Set TempApRcIt3 = GetAPRcpItemEx(m_PaidAllAmounts4, SupGroup.SUPPLIER_CODE)
         Set TempApRcIt3Cn = GetAPRcpItemEx(m_PaidAllAmountCns4, SupGroup.SUPPLIER_CODE)
         
         HaveValue = False
               
          '===Head
          Set HeadCf = m_Headers1.Fields(1)
          TempStr = Format(Month(FirstDate), "00") & "/" & Format(Year(FirstDate) + 543, "00")
          Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
          Call m_Details1.AddField(BodyCf)
          
          Set Ap = GetObject("CAPMas", Supplier, SupGroup.SUPPLIER_CODE)
          Set HeadCf = m_Headers1.Fields(2)
          TempStr = Left(Ap.SUPCOD & "/" & Ap.SUPNAM, 30)
          Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
          Call m_Details1.AddField(BodyCf)
          
          BalanceAmount = (TempApt.AMOUNT - TempAptCn.AMOUNT) - (TempApt.PAYAMT - TempAptCn.PAYAMT) - (TempApRcIt.PAYAMT - TempApRcItCn.PAYAMT) + (TempApRcIt3.PAYAMT - TempApRcIt3Cn.PAYAMT)
          
          K = 2
          K = K + 1
          Set HeadCf = m_Headers1.Fields(K)
          TempStr = FormatNumber(BalanceAmount, "")
          Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
          Call m_Details1.AddField(BodyCf)
          
          Total14(K) = Total14(K) + BalanceAmount
          If Round((BalanceAmount), 2) <> 0 Then
             HaveValue = True
          End If
          
          Set TempAptTo = GetAPTrn(m_ApSupAmountTo4, SupGroup.SUPPLIER_CODE)
          Set TempAptToCn = GetAPTrn(m_ApSupAmountToCn4, SupGroup.SUPPLIER_CODE)
          
          Set TempApRcItTo = GetAPRcpItem(m_ApSupPaidTo4, SupGroup.SUPPLIER_CODE)
          Set TempApRcItToCn = GetAPRcpItem(m_ApSupPaidToCn4, SupGroup.SUPPLIER_CODE)
                   
          K = K + 1
          Set HeadCf = m_Headers1.Fields(K)
          TempStr = FormatNumber(TempAptTo.AMOUNT, "")
          Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
          Call m_Details1.AddField(BodyCf)
          
          Total14(K) = Total14(K) + TempAptTo.AMOUNT
          Total3(K) = Total3(K) + TempAptTo.AMOUNT
          If Round((TempAptTo.AMOUNT), 2) <> 0 Then
             HaveValue = True
          End If
          
          K = K + 1
          Set HeadCf = m_Headers1.Fields(K)
          TempStr = FormatNumber(TempApRcItTo.PAYAMT - TempApRcItToCn.PAYAMT + TempAptToCn.AMOUNT, "")
          Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
          Call m_Details1.AddField(BodyCf)
          
          Total14(K) = Total14(K) + TempApRcItTo.PAYAMT - TempApRcItToCn.PAYAMT + TempAptToCn.AMOUNT
          Total3(K) = Total3(K) + TempApRcItTo.PAYAMT - TempApRcItToCn.PAYAMT + TempAptToCn.AMOUNT
          If Round((TempApRcItTo.PAYAMT - TempApRcItToCn.PAYAMT + TempAptToCn.AMOUNT), 2) <> 0 Then
             HaveValue = True
          End If
          TempPaidCheck = Round((TempApRcItTo.PAYAMT - TempApRcItToCn.PAYAMT + TempAptToCn.AMOUNT), 2)
          
          K = K + 1
          Set HeadCf = m_Headers1.Fields(K)
          TempStr = FormatNumber(BalanceAmount + (TempAptTo.AMOUNT - TempAptToCn.AMOUNT) - (TempApRcItTo.PAYAMT - TempApRcItToCn.PAYAMT), "")
          Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
          Call m_Details1.AddField(BodyCf)
          
          Total14(K) = Total14(K) + (BalanceAmount) + (TempAptTo.AMOUNT - TempAptToCn.AMOUNT) - (TempApRcItTo.PAYAMT - TempApRcItToCn.PAYAMT)
          If Round(((BalanceAmount) + (TempAptTo.AMOUNT - TempAptToCn.AMOUNT) - (TempApRcItTo.PAYAMT - TempApRcItToCn.PAYAMT)), 2) <> 0 Then
             HaveValue = True
          End If
          
          Set TempApBktTrn = GetBkTrn(m_54, SupGroup.SUPPLIER_CODE)
          Set TempAptCash = GetAPTrn(m_ApSupAmountCashs4, SupGroup.SUPPLIER_CODE)
          
          K = K + 1
          Set HeadCf = m_Headers1.Fields(K)
          TempStr = FormatNumber(TempApBktTrn.AMOUNT + TempAptCash.CSHPAY, "")
          Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
          Call m_Details1.AddField(BodyCf)
          
          Total14(K) = Total14(K) + TempApBktTrn.AMOUNT + TempAptCash.CSHPAY
          Total3(K) = Total3(K) + TempApBktTrn.AMOUNT + TempAptCash.CSHPAY
          If Round((TempApBktTrn.AMOUNT + TempAptCash.CSHPAY), 2) <> 0 Then
             HaveValue = True
          End If
          
          Set TempApBktTrn = GetBkTrn(m_64, SupGroup.SUPPLIER_CODE)
          K = K + 1
          Set HeadCf = m_Headers1.Fields(K)
          TempStr = FormatNumber(TempApBktTrn.AMOUNT, "")
          Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
          Call m_Details1.AddField(BodyCf)
          
          Total14(K) = Total14(K) + TempApBktTrn.AMOUNT
          Total3(K) = TempApBktTrn.AMOUNT
          If Round((TempApBktTrn.AMOUNT), 2) <> 0 Then
             HaveValue = True
          End If
          
          '7 =  4 + 6
          K = K + 1
          Set HeadCf = m_Headers1.Fields(K)
          TempStr = FormatNumber(BalanceAmount + (TempAptTo.AMOUNT - TempAptToCn.AMOUNT) - (TempApRcItTo.PAYAMT - TempApRcItToCn.PAYAMT) + TempApBktTrn.AMOUNT, "")
          Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
          Call m_Details1.AddField(BodyCf)
          
          Total14(K) = Total14(K) + BalanceAmount + (TempAptTo.AMOUNT - TempAptToCn.AMOUNT) - (TempApRcItTo.PAYAMT - TempApRcItToCn.PAYAMT) + TempApBktTrn.AMOUNT
          Total3(K) = BalanceAmount + (TempAptTo.AMOUNT - TempAptToCn.AMOUNT) - (TempApRcItTo.PAYAMT - TempApRcItToCn.PAYAMT) + TempApBktTrn.AMOUNT
          If Round((BalanceAmount + (TempAptTo.AMOUNT - TempAptToCn.AMOUNT) - (TempApRcItTo.PAYAMT - TempApRcItToCn.PAYAMT) + TempApBktTrn.AMOUNT), 2) <> 0 Then
             HaveValue = True
          End If
          
         Set TempApBktTrn = GetBkTrn(m_54, SupGroup.SUPPLIER_CODE)
          Set TempAptCash = GetAPTrn(m_ApSupAmountCashs4, SupGroup.SUPPLIER_CODE)
          
          K = K + 1
          Set HeadCf = m_Headers1.Fields(K)
          TempStr = FormatNumber(TempAptTo.AMOUNT - TempApBktTrn.AMOUNT - TempAptCash.CSHPAY, "")
          Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
          Call m_Details1.AddField(BodyCf)
          
          Total14(K) = Total14(K) + TempAptTo.AMOUNT - TempApBktTrn.AMOUNT - TempAptCash.CSHPAY
          Total3(K) = Total3(K) + TempAptTo.AMOUNT - TempApBktTrn.AMOUNT - TempAptCash.CSHPAY
          If Round((TempAptTo.AMOUNT - TempApBktTrn.AMOUNT - TempAptCash.CSHPAY), 2) <> 0 Then
             HaveValue = True
          End If
          
          If HaveValue Then
            If mcolParam("SUMMARY_FLAG") = "N" And mcolParam("ONLY_SUPPLIER") = "N" Then
               Call m_Details1.GetString(1, TempStr1, TempStr2)
               strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
               Call Vsp.AddTable(strFormat, "", TempStr2)
             End If
             HaveValueSum = True
          End If
          ' ของเดือนแรกที่เลือก 4
      End If
      
      If (RunMonth >= 5) Then
         ' ของเดือนแรกที่เลือก 5
         
         FirstDate = DateAdd("M", 1, FirstDate)
         Call m_Details1.ClearField
               
         Set TempApt = GetAPTrn(m_ApSupAmounts5, SupGroup.SUPPLIER_CODE)
         Set TempAptCn = GetAPTrn(m_ApSupAmountCns5, SupGroup.SUPPLIER_CODE)
         
         Set TempApRcIt = GetAPRcpItem(m_ApSupPaids5, SupGroup.SUPPLIER_CODE)
         Set TempApRcItCn = GetAPRcpItem(m_ApSupPaidCns5, SupGroup.SUPPLIER_CODE)
            
         Set TempApRcIt3 = GetAPRcpItemEx(m_PaidAllAmounts5, SupGroup.SUPPLIER_CODE)
         Set TempApRcIt3Cn = GetAPRcpItemEx(m_PaidAllAmountCns5, SupGroup.SUPPLIER_CODE)
         
         HaveValue = False
               
          '===Head
          Set HeadCf = m_Headers1.Fields(1)
          TempStr = Format(Month(FirstDate), "00") & "/" & Format(Year(FirstDate) + 543, "00")
          Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
          Call m_Details1.AddField(BodyCf)
          
          Set Ap = GetObject("CAPMas", Supplier, SupGroup.SUPPLIER_CODE)
          Set HeadCf = m_Headers1.Fields(2)
          TempStr = Left(Ap.SUPCOD & "/" & Ap.SUPNAM, 30)
          Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
          Call m_Details1.AddField(BodyCf)
          
          BalanceAmount = (TempApt.AMOUNT - TempAptCn.AMOUNT) - (TempApt.PAYAMT - TempAptCn.PAYAMT) - (TempApRcIt.PAYAMT - TempApRcItCn.PAYAMT) + (TempApRcIt3.PAYAMT - TempApRcIt3Cn.PAYAMT)
          
          K = 2
          K = K + 1
          Set HeadCf = m_Headers1.Fields(K)
          TempStr = FormatNumber(BalanceAmount, "")
          Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
          Call m_Details1.AddField(BodyCf)
          
          Total15(K) = Total15(K) + BalanceAmount
          If Round((BalanceAmount), 2) <> 0 Then
             HaveValue = True
          End If
          
          Set TempAptTo = GetAPTrn(m_ApSupAmountTo5, SupGroup.SUPPLIER_CODE)
          Set TempAptToCn = GetAPTrn(m_ApSupAmountToCn5, SupGroup.SUPPLIER_CODE)
          
          Set TempApRcItTo = GetAPRcpItem(m_ApSupPaidTo5, SupGroup.SUPPLIER_CODE)
          Set TempApRcItToCn = GetAPRcpItem(m_ApSupPaidToCn5, SupGroup.SUPPLIER_CODE)
                   
          K = K + 1
          Set HeadCf = m_Headers1.Fields(K)
          TempStr = FormatNumber(TempAptTo.AMOUNT, "")
          Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
          Call m_Details1.AddField(BodyCf)
          
          Total15(K) = Total15(K) + TempAptTo.AMOUNT
          Total3(K) = Total3(K) + TempAptTo.AMOUNT
          If Round((TempAptTo.AMOUNT), 2) <> 0 Then
             HaveValue = True
          End If
          
          K = K + 1
          Set HeadCf = m_Headers1.Fields(K)
          TempStr = FormatNumber(TempApRcItTo.PAYAMT - TempApRcItToCn.PAYAMT + TempAptToCn.AMOUNT, "")
          Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
          Call m_Details1.AddField(BodyCf)
          
          Total15(K) = Total15(K) + TempApRcItTo.PAYAMT - TempApRcItToCn.PAYAMT + TempAptToCn.AMOUNT
          Total3(K) = Total3(K) + TempApRcItTo.PAYAMT - TempApRcItToCn.PAYAMT + TempAptToCn.AMOUNT
          If Round((TempApRcItTo.PAYAMT - TempApRcItToCn.PAYAMT + TempAptToCn.AMOUNT), 2) <> 0 Then
             HaveValue = True
          End If
          TempPaidCheck = Round((TempApRcItTo.PAYAMT - TempApRcItToCn.PAYAMT + TempAptToCn.AMOUNT), 2)
          
          K = K + 1
          Set HeadCf = m_Headers1.Fields(K)
          TempStr = FormatNumber(BalanceAmount + (TempAptTo.AMOUNT - TempAptToCn.AMOUNT) - (TempApRcItTo.PAYAMT - TempApRcItToCn.PAYAMT), "")
          Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
          Call m_Details1.AddField(BodyCf)
          
          Total15(K) = Total15(K) + (BalanceAmount) + (TempAptTo.AMOUNT - TempAptToCn.AMOUNT) - (TempApRcItTo.PAYAMT - TempApRcItToCn.PAYAMT)
          If Round(((BalanceAmount) + (TempAptTo.AMOUNT - TempAptToCn.AMOUNT) - (TempApRcItTo.PAYAMT - TempApRcItToCn.PAYAMT)), 2) <> 0 Then
             HaveValue = True
          End If
          
          Set TempApBktTrn = GetBkTrn(m_55, SupGroup.SUPPLIER_CODE)
          Set TempAptCash = GetAPTrn(m_ApSupAmountCashs5, SupGroup.SUPPLIER_CODE)
          
          K = K + 1
          Set HeadCf = m_Headers1.Fields(K)
          TempStr = FormatNumber(TempApBktTrn.AMOUNT + TempAptCash.CSHPAY, "")
          Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
          Call m_Details1.AddField(BodyCf)
          
          Total15(K) = Total15(K) + TempApBktTrn.AMOUNT + TempAptCash.CSHPAY
          Total3(K) = Total3(K) + TempApBktTrn.AMOUNT + TempAptCash.CSHPAY
          If Round((TempApBktTrn.AMOUNT + TempAptCash.CSHPAY), 2) <> 0 Then
             HaveValue = True
          End If
          
          Set TempApBktTrn = GetBkTrn(m_65, SupGroup.SUPPLIER_CODE)
          K = K + 1
          Set HeadCf = m_Headers1.Fields(K)
          TempStr = FormatNumber(TempApBktTrn.AMOUNT, "")
          Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
          Call m_Details1.AddField(BodyCf)
          
          Total15(K) = Total15(K) + TempApBktTrn.AMOUNT
          Total3(K) = TempApBktTrn.AMOUNT
          If Round((TempApBktTrn.AMOUNT), 2) <> 0 Then
             HaveValue = True
          End If
          
          '7 =  4 + 6
          K = K + 1
          Set HeadCf = m_Headers1.Fields(K)
          TempStr = FormatNumber(BalanceAmount + (TempAptTo.AMOUNT - TempAptToCn.AMOUNT) - (TempApRcItTo.PAYAMT - TempApRcItToCn.PAYAMT) + TempApBktTrn.AMOUNT, "")
          Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
          Call m_Details1.AddField(BodyCf)
          
          Total15(K) = Total15(K) + BalanceAmount + (TempAptTo.AMOUNT - TempAptToCn.AMOUNT) - (TempApRcItTo.PAYAMT - TempApRcItToCn.PAYAMT) + TempApBktTrn.AMOUNT
          Total3(K) = BalanceAmount + (TempAptTo.AMOUNT - TempAptToCn.AMOUNT) - (TempApRcItTo.PAYAMT - TempApRcItToCn.PAYAMT) + TempApBktTrn.AMOUNT
          If Round((BalanceAmount + (TempAptTo.AMOUNT - TempAptToCn.AMOUNT) - (TempApRcItTo.PAYAMT - TempApRcItToCn.PAYAMT) + TempApBktTrn.AMOUNT), 2) <> 0 Then
             HaveValue = True
          End If
          
         Set TempApBktTrn = GetBkTrn(m_55, SupGroup.SUPPLIER_CODE)
          Set TempAptCash = GetAPTrn(m_ApSupAmountCashs5, SupGroup.SUPPLIER_CODE)
          
          K = K + 1
          Set HeadCf = m_Headers1.Fields(K)
          TempStr = FormatNumber(TempAptTo.AMOUNT - TempApBktTrn.AMOUNT - TempAptCash.CSHPAY, "")
          Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
          Call m_Details1.AddField(BodyCf)
          
          Total15(K) = Total15(K) + TempAptTo.AMOUNT - TempApBktTrn.AMOUNT - TempAptCash.CSHPAY
          Total3(K) = Total3(K) + TempAptTo.AMOUNT - TempApBktTrn.AMOUNT - TempAptCash.CSHPAY
          If Round((TempAptTo.AMOUNT - TempApBktTrn.AMOUNT - TempAptCash.CSHPAY), 2) <> 0 Then
             HaveValue = True
          End If
          
          If HaveValue Then
            If mcolParam("SUMMARY_FLAG") = "N" And mcolParam("ONLY_SUPPLIER") = "N" Then
               Call m_Details1.GetString(1, TempStr1, TempStr2)
               strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
               Call Vsp.AddTable(strFormat, "", TempStr2)
             End If
             HaveValueSum = True
          End If
          ' ของเดือนแรกที่เลือก 5
      End If
      
      If (RunMonth >= 6) Then
         ' ของเดือนแรกที่เลือก 6
         
         FirstDate = DateAdd("M", 1, FirstDate)
         Call m_Details1.ClearField
               
         Set TempApt = GetAPTrn(m_ApSupAmounts6, SupGroup.SUPPLIER_CODE)
         Set TempAptCn = GetAPTrn(m_ApSupAmountCns6, SupGroup.SUPPLIER_CODE)
         
         Set TempApRcIt = GetAPRcpItem(m_ApSupPaids6, SupGroup.SUPPLIER_CODE)
         Set TempApRcItCn = GetAPRcpItem(m_ApSupPaidCns6, SupGroup.SUPPLIER_CODE)
            
         Set TempApRcIt3 = GetAPRcpItemEx(m_PaidAllAmounts6, SupGroup.SUPPLIER_CODE)
         Set TempApRcIt3Cn = GetAPRcpItemEx(m_PaidAllAmountCns6, SupGroup.SUPPLIER_CODE)
         
         HaveValue = False
               
          '===Head
          Set HeadCf = m_Headers1.Fields(1)
          TempStr = Format(Month(FirstDate), "00") & "/" & Format(Year(FirstDate) + 543, "00")
          Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
          Call m_Details1.AddField(BodyCf)
          
          Set Ap = GetObject("CAPMas", Supplier, SupGroup.SUPPLIER_CODE)
          Set HeadCf = m_Headers1.Fields(2)
          TempStr = Left(Ap.SUPCOD & "/" & Ap.SUPNAM, 30)
          Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
          Call m_Details1.AddField(BodyCf)
          
          BalanceAmount = (TempApt.AMOUNT - TempAptCn.AMOUNT) - (TempApt.PAYAMT - TempAptCn.PAYAMT) - (TempApRcIt.PAYAMT - TempApRcItCn.PAYAMT) + (TempApRcIt3.PAYAMT - TempApRcIt3Cn.PAYAMT)
          
          K = 2
          K = K + 1
          Set HeadCf = m_Headers1.Fields(K)
          TempStr = FormatNumber(BalanceAmount, "")
          Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
          Call m_Details1.AddField(BodyCf)
          
          Total16(K) = Total16(K) + BalanceAmount
          If Round((BalanceAmount), 2) <> 0 Then
             HaveValue = True
          End If
          
          Set TempAptTo = GetAPTrn(m_ApSupAmountTo6, SupGroup.SUPPLIER_CODE)
          Set TempAptToCn = GetAPTrn(m_ApSupAmountToCn6, SupGroup.SUPPLIER_CODE)
          
          Set TempApRcItTo = GetAPRcpItem(m_ApSupPaidTo6, SupGroup.SUPPLIER_CODE)
          Set TempApRcItToCn = GetAPRcpItem(m_ApSupPaidToCn6, SupGroup.SUPPLIER_CODE)
                   
          K = K + 1
          Set HeadCf = m_Headers1.Fields(K)
          TempStr = FormatNumber(TempAptTo.AMOUNT, "")
          Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
          Call m_Details1.AddField(BodyCf)
          
          Total16(K) = Total16(K) + TempAptTo.AMOUNT
          Total3(K) = Total3(K) + TempAptTo.AMOUNT
          If Round((TempAptTo.AMOUNT), 2) <> 0 Then
             HaveValue = True
          End If
          
          K = K + 1
          Set HeadCf = m_Headers1.Fields(K)
          TempStr = FormatNumber(TempApRcItTo.PAYAMT - TempApRcItToCn.PAYAMT + TempAptToCn.AMOUNT, "")
          Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
          Call m_Details1.AddField(BodyCf)
          
          Total16(K) = Total16(K) + TempApRcItTo.PAYAMT - TempApRcItToCn.PAYAMT + TempAptToCn.AMOUNT
          Total3(K) = Total3(K) + TempApRcItTo.PAYAMT - TempApRcItToCn.PAYAMT + TempAptToCn.AMOUNT
          If Round((TempApRcItTo.PAYAMT - TempApRcItToCn.PAYAMT + TempAptToCn.AMOUNT), 2) <> 0 Then
             HaveValue = True
          End If
          TempPaidCheck = Round((TempApRcItTo.PAYAMT - TempApRcItToCn.PAYAMT + TempAptToCn.AMOUNT), 2)
          
          K = K + 1
          Set HeadCf = m_Headers1.Fields(K)
          TempStr = FormatNumber(BalanceAmount + (TempAptTo.AMOUNT - TempAptToCn.AMOUNT) - (TempApRcItTo.PAYAMT - TempApRcItToCn.PAYAMT), "")
          Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
          Call m_Details1.AddField(BodyCf)
          
          Total16(K) = Total16(K) + (BalanceAmount) + (TempAptTo.AMOUNT - TempAptToCn.AMOUNT) - (TempApRcItTo.PAYAMT - TempApRcItToCn.PAYAMT)
          If Round(((BalanceAmount) + (TempAptTo.AMOUNT - TempAptToCn.AMOUNT) - (TempApRcItTo.PAYAMT - TempApRcItToCn.PAYAMT)), 2) <> 0 Then
             HaveValue = True
          End If
          
          Set TempApBktTrn = GetBkTrn(m_56, SupGroup.SUPPLIER_CODE)
          Set TempAptCash = GetAPTrn(m_ApSupAmountCashs6, SupGroup.SUPPLIER_CODE)
          
          K = K + 1
          Set HeadCf = m_Headers1.Fields(K)
          TempStr = FormatNumber(TempApBktTrn.AMOUNT + TempAptCash.CSHPAY, "")
          Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
          Call m_Details1.AddField(BodyCf)
          
          Total16(K) = Total16(K) + TempApBktTrn.AMOUNT + TempAptCash.CSHPAY
          Total3(K) = Total3(K) + TempApBktTrn.AMOUNT + TempAptCash.CSHPAY
          If Round((TempApBktTrn.AMOUNT + TempAptCash.CSHPAY), 2) <> 0 Then
             HaveValue = True
          End If
          
          Set TempApBktTrn = GetBkTrn(m_66, SupGroup.SUPPLIER_CODE)
          K = K + 1
          Set HeadCf = m_Headers1.Fields(K)
          TempStr = FormatNumber(TempApBktTrn.AMOUNT, "")
          Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
          Call m_Details1.AddField(BodyCf)
          
          Total16(K) = Total16(K) + TempApBktTrn.AMOUNT
          Total3(K) = TempApBktTrn.AMOUNT
          If Round((TempApBktTrn.AMOUNT), 2) <> 0 Then
             HaveValue = True
          End If
          
          '7 =  4 + 6
          K = K + 1
          Set HeadCf = m_Headers1.Fields(K)
          TempStr = FormatNumber(BalanceAmount + (TempAptTo.AMOUNT - TempAptToCn.AMOUNT) - (TempApRcItTo.PAYAMT - TempApRcItToCn.PAYAMT) + TempApBktTrn.AMOUNT, "")
          Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
          Call m_Details1.AddField(BodyCf)
          
          Total16(K) = Total16(K) + BalanceAmount + (TempAptTo.AMOUNT - TempAptToCn.AMOUNT) - (TempApRcItTo.PAYAMT - TempApRcItToCn.PAYAMT) + TempApBktTrn.AMOUNT
          Total3(K) = BalanceAmount + (TempAptTo.AMOUNT - TempAptToCn.AMOUNT) - (TempApRcItTo.PAYAMT - TempApRcItToCn.PAYAMT) + TempApBktTrn.AMOUNT
          If Round((BalanceAmount + (TempAptTo.AMOUNT - TempAptToCn.AMOUNT) - (TempApRcItTo.PAYAMT - TempApRcItToCn.PAYAMT) + TempApBktTrn.AMOUNT), 2) <> 0 Then
             HaveValue = True
          End If
          
         Set TempApBktTrn = GetBkTrn(m_56, SupGroup.SUPPLIER_CODE)
          Set TempAptCash = GetAPTrn(m_ApSupAmountCashs6, SupGroup.SUPPLIER_CODE)
          
          K = K + 1
          Set HeadCf = m_Headers1.Fields(K)
          TempStr = FormatNumber(TempAptTo.AMOUNT - TempApBktTrn.AMOUNT - TempAptCash.CSHPAY, "")
          Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
          Call m_Details1.AddField(BodyCf)
          
          Total16(K) = Total16(K) + TempAptTo.AMOUNT - TempApBktTrn.AMOUNT - TempAptCash.CSHPAY
          Total3(K) = Total3(K) + TempAptTo.AMOUNT - TempApBktTrn.AMOUNT - TempAptCash.CSHPAY
          If Round((TempAptTo.AMOUNT - TempApBktTrn.AMOUNT - TempAptCash.CSHPAY), 2) <> 0 Then
             HaveValue = True
          End If
          
          If HaveValue Then
            If mcolParam("SUMMARY_FLAG") = "N" And mcolParam("ONLY_SUPPLIER") = "N" Then
               Call m_Details1.GetString(1, TempStr1, TempStr2)
               strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
               Call Vsp.AddTable(strFormat, "", TempStr2)
             End If
             HaveValueSum = True
          End If
          ' ของเดือนแรกที่เลือก 6
      End If
      
      If (RunMonth >= 7) Then
         ' ของเดือนแรกที่เลือก 7

         FirstDate = DateAdd("M", 1, FirstDate)
         Call m_Details1.ClearField

         Set TempApt = GetAPTrn(m_ApSupAmounts7, SupGroup.SUPPLIER_CODE)
         Set TempAptCn = GetAPTrn(m_ApSupAmountCns7, SupGroup.SUPPLIER_CODE)

         Set TempApRcIt = GetAPRcpItem(m_ApSupPaids7, SupGroup.SUPPLIER_CODE)
         Set TempApRcItCn = GetAPRcpItem(m_ApSupPaidCns7, SupGroup.SUPPLIER_CODE)

         Set TempApRcIt3 = GetAPRcpItemEx(m_PaidAllAmounts7, SupGroup.SUPPLIER_CODE)
         Set TempApRcIt3Cn = GetAPRcpItemEx(m_PaidAllAmountCns7, SupGroup.SUPPLIER_CODE)

         HaveValue = False

          '===Head
          Set HeadCf = m_Headers1.Fields(1)
          TempStr = Format(Month(FirstDate), "00") & "/" & Format(Year(FirstDate) + 543, "00")
          Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
          Call m_Details1.AddField(BodyCf)

          Set Ap = GetObject("CAPMas", Supplier, SupGroup.SUPPLIER_CODE)
          Set HeadCf = m_Headers1.Fields(2)
          TempStr = Left(Ap.SUPCOD & "/" & Ap.SUPNAM, 30)
          Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
          Call m_Details1.AddField(BodyCf)

          BalanceAmount = (TempApt.AMOUNT - TempAptCn.AMOUNT) - (TempApt.PAYAMT - TempAptCn.PAYAMT) - (TempApRcIt.PAYAMT - TempApRcItCn.PAYAMT) + (TempApRcIt3.PAYAMT - TempApRcIt3Cn.PAYAMT)

          K = 2
          K = K + 1
          Set HeadCf = m_Headers1.Fields(K)
          TempStr = FormatNumber(BalanceAmount, "")
          Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
          Call m_Details1.AddField(BodyCf)

          Total17(K) = Total17(K) + BalanceAmount
          If Round((BalanceAmount), 2) <> 0 Then
             HaveValue = True
          End If

          Set TempAptTo = GetAPTrn(m_ApSupAmountTo7, SupGroup.SUPPLIER_CODE)
          Set TempAptToCn = GetAPTrn(m_ApSupAmountToCn7, SupGroup.SUPPLIER_CODE)

          Set TempApRcItTo = GetAPRcpItem(m_ApSupPaidTo7, SupGroup.SUPPLIER_CODE)
          Set TempApRcItToCn = GetAPRcpItem(m_ApSupPaidToCn7, SupGroup.SUPPLIER_CODE)

          K = K + 1
          Set HeadCf = m_Headers1.Fields(K)
          TempStr = FormatNumber(TempAptTo.AMOUNT, "")
          Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
          Call m_Details1.AddField(BodyCf)

          Total17(K) = Total17(K) + TempAptTo.AMOUNT
          Total3(K) = Total3(K) + TempAptTo.AMOUNT
          If Round((TempAptTo.AMOUNT), 2) <> 0 Then
             HaveValue = True
          End If

          K = K + 1
          Set HeadCf = m_Headers1.Fields(K)
          TempStr = FormatNumber(TempApRcItTo.PAYAMT - TempApRcItToCn.PAYAMT + TempAptToCn.AMOUNT, "")
          Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
          Call m_Details1.AddField(BodyCf)

          Total17(K) = Total17(K) + TempApRcItTo.PAYAMT - TempApRcItToCn.PAYAMT + TempAptToCn.AMOUNT
          Total3(K) = Total3(K) + TempApRcItTo.PAYAMT - TempApRcItToCn.PAYAMT + TempAptToCn.AMOUNT
          If Round((TempApRcItTo.PAYAMT - TempApRcItToCn.PAYAMT + TempAptToCn.AMOUNT), 2) <> 0 Then
             HaveValue = True
          End If
          TempPaidCheck = Round((TempApRcItTo.PAYAMT - TempApRcItToCn.PAYAMT + TempAptToCn.AMOUNT), 2)

          K = K + 1
          Set HeadCf = m_Headers1.Fields(K)
          TempStr = FormatNumber(BalanceAmount + (TempAptTo.AMOUNT - TempAptToCn.AMOUNT) - (TempApRcItTo.PAYAMT - TempApRcItToCn.PAYAMT), "")
          Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
          Call m_Details1.AddField(BodyCf)

          Total17(K) = Total17(K) + (BalanceAmount) + (TempAptTo.AMOUNT - TempAptToCn.AMOUNT) - (TempApRcItTo.PAYAMT - TempApRcItToCn.PAYAMT)
          If Round(((BalanceAmount) + (TempAptTo.AMOUNT - TempAptToCn.AMOUNT) - (TempApRcItTo.PAYAMT - TempApRcItToCn.PAYAMT)), 2) <> 0 Then
             HaveValue = True
          End If

          Set TempApBktTrn = GetBkTrn(m_57, SupGroup.SUPPLIER_CODE)
          Set TempAptCash = GetAPTrn(m_ApSupAmountCashs7, SupGroup.SUPPLIER_CODE)

          K = K + 1
          Set HeadCf = m_Headers1.Fields(K)
          TempStr = FormatNumber(TempApBktTrn.AMOUNT + TempAptCash.CSHPAY, "")
          Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
          Call m_Details1.AddField(BodyCf)

          Total17(K) = Total17(K) + TempApBktTrn.AMOUNT + TempAptCash.CSHPAY
          Total3(K) = Total3(K) + TempApBktTrn.AMOUNT + TempAptCash.CSHPAY
          If Round((TempApBktTrn.AMOUNT + TempAptCash.CSHPAY), 2) <> 0 Then
             HaveValue = True
          End If

          Set TempApBktTrn = GetBkTrn(m_67, SupGroup.SUPPLIER_CODE)
          K = K + 1
          Set HeadCf = m_Headers1.Fields(K)
          TempStr = FormatNumber(TempApBktTrn.AMOUNT, "")
          Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
          Call m_Details1.AddField(BodyCf)

          Total17(K) = Total17(K) + TempApBktTrn.AMOUNT
          Total3(K) = TempApBktTrn.AMOUNT
          If Round((TempApBktTrn.AMOUNT), 2) <> 0 Then
             HaveValue = True
          End If

          '7 =  4 + 6
          K = K + 1
          Set HeadCf = m_Headers1.Fields(K)
          TempStr = FormatNumber(BalanceAmount + (TempAptTo.AMOUNT - TempAptToCn.AMOUNT) - (TempApRcItTo.PAYAMT - TempApRcItToCn.PAYAMT) + TempApBktTrn.AMOUNT, "")
          Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
          Call m_Details1.AddField(BodyCf)

          Total17(K) = Total17(K) + BalanceAmount + (TempAptTo.AMOUNT - TempAptToCn.AMOUNT) - (TempApRcItTo.PAYAMT - TempApRcItToCn.PAYAMT) + TempApBktTrn.AMOUNT
          Total3(K) = BalanceAmount + (TempAptTo.AMOUNT - TempAptToCn.AMOUNT) - (TempApRcItTo.PAYAMT - TempApRcItToCn.PAYAMT) + TempApBktTrn.AMOUNT
          If Round((BalanceAmount + (TempAptTo.AMOUNT - TempAptToCn.AMOUNT) - (TempApRcItTo.PAYAMT - TempApRcItToCn.PAYAMT) + TempApBktTrn.AMOUNT), 2) <> 0 Then
             HaveValue = True
          End If

         Set TempApBktTrn = GetBkTrn(m_57, SupGroup.SUPPLIER_CODE)
          Set TempAptCash = GetAPTrn(m_ApSupAmountCashs7, SupGroup.SUPPLIER_CODE)

          K = K + 1
          Set HeadCf = m_Headers1.Fields(K)
          TempStr = FormatNumber(TempAptTo.AMOUNT - TempApBktTrn.AMOUNT - TempAptCash.CSHPAY, "")
          Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
          Call m_Details1.AddField(BodyCf)

          Total17(K) = Total17(K) + TempAptTo.AMOUNT - TempApBktTrn.AMOUNT - TempAptCash.CSHPAY
          Total3(K) = Total3(K) + TempAptTo.AMOUNT - TempApBktTrn.AMOUNT - TempAptCash.CSHPAY
          If Round((TempAptTo.AMOUNT - TempApBktTrn.AMOUNT - TempAptCash.CSHPAY), 2) <> 0 Then
             HaveValue = True
          End If

          If HaveValue Then
            If mcolParam("SUMMARY_FLAG") = "N" And mcolParam("ONLY_SUPPLIER") = "N" Then
               Call m_Details1.GetString(1, TempStr1, TempStr2)
               strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
               Call Vsp.AddTable(strFormat, "", TempStr2)
             End If
             HaveValueSum = True
          End If
          ' ของเดือนแรกที่เลือก 7
      End If

      If (RunMonth >= 8) Then
         ' ของเดือนแรกที่เลือก 8

         FirstDate = DateAdd("M", 1, FirstDate)
         Call m_Details1.ClearField

         Set TempApt = GetAPTrn(m_ApSupAmounts8, SupGroup.SUPPLIER_CODE)
         Set TempAptCn = GetAPTrn(m_ApSupAmountCns8, SupGroup.SUPPLIER_CODE)

         Set TempApRcIt = GetAPRcpItem(m_ApSupPaids8, SupGroup.SUPPLIER_CODE)
         Set TempApRcItCn = GetAPRcpItem(m_ApSupPaidCns8, SupGroup.SUPPLIER_CODE)

         Set TempApRcIt3 = GetAPRcpItemEx(m_PaidAllAmounts8, SupGroup.SUPPLIER_CODE)
         Set TempApRcIt3Cn = GetAPRcpItemEx(m_PaidAllAmountCns8, SupGroup.SUPPLIER_CODE)

         HaveValue = False

          '===Head
          Set HeadCf = m_Headers1.Fields(1)
          TempStr = Format(Month(FirstDate), "00") & "/" & Format(Year(FirstDate) + 543, "00")
          Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
          Call m_Details1.AddField(BodyCf)

          Set Ap = GetObject("CAPMas", Supplier, SupGroup.SUPPLIER_CODE)
          Set HeadCf = m_Headers1.Fields(2)
          TempStr = Left(Ap.SUPCOD & "/" & Ap.SUPNAM, 30)
          Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
          Call m_Details1.AddField(BodyCf)

          BalanceAmount = (TempApt.AMOUNT - TempAptCn.AMOUNT) - (TempApt.PAYAMT - TempAptCn.PAYAMT) - (TempApRcIt.PAYAMT - TempApRcItCn.PAYAMT) + (TempApRcIt3.PAYAMT - TempApRcIt3Cn.PAYAMT)

          K = 2
          K = K + 1
          Set HeadCf = m_Headers1.Fields(K)
          TempStr = FormatNumber(BalanceAmount, "")
          Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
          Call m_Details1.AddField(BodyCf)

          Total18(K) = Total18(K) + BalanceAmount
          If Round((BalanceAmount), 2) <> 0 Then
             HaveValue = True
          End If

          Set TempAptTo = GetAPTrn(m_ApSupAmountTo8, SupGroup.SUPPLIER_CODE)
          Set TempAptToCn = GetAPTrn(m_ApSupAmountToCn8, SupGroup.SUPPLIER_CODE)

          Set TempApRcItTo = GetAPRcpItem(m_ApSupPaidTo8, SupGroup.SUPPLIER_CODE)
          Set TempApRcItToCn = GetAPRcpItem(m_ApSupPaidToCn8, SupGroup.SUPPLIER_CODE)

          K = K + 1
          Set HeadCf = m_Headers1.Fields(K)
          TempStr = FormatNumber(TempAptTo.AMOUNT, "")
          Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
          Call m_Details1.AddField(BodyCf)

          Total18(K) = Total18(K) + TempAptTo.AMOUNT
          Total3(K) = Total3(K) + TempAptTo.AMOUNT
          If Round((TempAptTo.AMOUNT), 2) <> 0 Then
             HaveValue = True
          End If

          K = K + 1
          Set HeadCf = m_Headers1.Fields(K)
          TempStr = FormatNumber(TempApRcItTo.PAYAMT - TempApRcItToCn.PAYAMT + TempAptToCn.AMOUNT, "")
          Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
          Call m_Details1.AddField(BodyCf)

          Total18(K) = Total18(K) + TempApRcItTo.PAYAMT - TempApRcItToCn.PAYAMT + TempAptToCn.AMOUNT
          Total3(K) = Total3(K) + TempApRcItTo.PAYAMT - TempApRcItToCn.PAYAMT + TempAptToCn.AMOUNT
          If Round((TempApRcItTo.PAYAMT - TempApRcItToCn.PAYAMT + TempAptToCn.AMOUNT), 2) <> 0 Then
             HaveValue = True
          End If
          TempPaidCheck = Round((TempApRcItTo.PAYAMT - TempApRcItToCn.PAYAMT + TempAptToCn.AMOUNT), 2)

          K = K + 1
          Set HeadCf = m_Headers1.Fields(K)
          TempStr = FormatNumber(BalanceAmount + (TempAptTo.AMOUNT - TempAptToCn.AMOUNT) - (TempApRcItTo.PAYAMT - TempApRcItToCn.PAYAMT), "")
          Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
          Call m_Details1.AddField(BodyCf)

          Total18(K) = Total18(K) + (BalanceAmount) + (TempAptTo.AMOUNT - TempAptToCn.AMOUNT) - (TempApRcItTo.PAYAMT - TempApRcItToCn.PAYAMT)
          If Round(((BalanceAmount) + (TempAptTo.AMOUNT - TempAptToCn.AMOUNT) - (TempApRcItTo.PAYAMT - TempApRcItToCn.PAYAMT)), 2) <> 0 Then
             HaveValue = True
          End If
          
          Set TempApBktTrn = GetBkTrn(m_58, SupGroup.SUPPLIER_CODE)
          Set TempAptCash = GetAPTrn(m_ApSupAmountCashs8, SupGroup.SUPPLIER_CODE)

          K = K + 1
          Set HeadCf = m_Headers1.Fields(K)
          TempStr = FormatNumber(TempApBktTrn.AMOUNT + TempAptCash.CSHPAY, "")
          Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
          Call m_Details1.AddField(BodyCf)

          Total18(K) = Total18(K) + TempApBktTrn.AMOUNT + TempAptCash.CSHPAY
          Total3(K) = Total3(K) + TempApBktTrn.AMOUNT + TempAptCash.CSHPAY
          If Round((TempApBktTrn.AMOUNT + TempAptCash.CSHPAY), 2) <> 0 Then
             HaveValue = True
          End If

          Set TempApBktTrn = GetBkTrn(m_68, SupGroup.SUPPLIER_CODE)
          K = K + 1
          Set HeadCf = m_Headers1.Fields(K)
          TempStr = FormatNumber(TempApBktTrn.AMOUNT, "")
          Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
          Call m_Details1.AddField(BodyCf)

          Total18(K) = Total18(K) + TempApBktTrn.AMOUNT
          Total3(K) = TempApBktTrn.AMOUNT
          If Round((TempApBktTrn.AMOUNT), 2) <> 0 Then
             HaveValue = True
          End If

          '7 =  4 + 6
          K = K + 1
          Set HeadCf = m_Headers1.Fields(K)
          TempStr = FormatNumber(BalanceAmount + (TempAptTo.AMOUNT - TempAptToCn.AMOUNT) - (TempApRcItTo.PAYAMT - TempApRcItToCn.PAYAMT) + TempApBktTrn.AMOUNT, "")
          Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
          Call m_Details1.AddField(BodyCf)

          Total18(K) = Total18(K) + BalanceAmount + (TempAptTo.AMOUNT - TempAptToCn.AMOUNT) - (TempApRcItTo.PAYAMT - TempApRcItToCn.PAYAMT) + TempApBktTrn.AMOUNT
          Total3(K) = BalanceAmount + (TempAptTo.AMOUNT - TempAptToCn.AMOUNT) - (TempApRcItTo.PAYAMT - TempApRcItToCn.PAYAMT) + TempApBktTrn.AMOUNT
          If Round((BalanceAmount + (TempAptTo.AMOUNT - TempAptToCn.AMOUNT) - (TempApRcItTo.PAYAMT - TempApRcItToCn.PAYAMT) + TempApBktTrn.AMOUNT), 2) <> 0 Then
             HaveValue = True
          End If

         Set TempApBktTrn = GetBkTrn(m_58, SupGroup.SUPPLIER_CODE)
          Set TempAptCash = GetAPTrn(m_ApSupAmountCashs8, SupGroup.SUPPLIER_CODE)

          K = K + 1
          Set HeadCf = m_Headers1.Fields(K)
          TempStr = FormatNumber(TempAptTo.AMOUNT - TempApBktTrn.AMOUNT - TempAptCash.CSHPAY, "")
          Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
          Call m_Details1.AddField(BodyCf)

          Total18(K) = Total18(K) + TempAptTo.AMOUNT - TempApBktTrn.AMOUNT - TempAptCash.CSHPAY
          Total3(K) = Total3(K) + TempAptTo.AMOUNT - TempApBktTrn.AMOUNT - TempAptCash.CSHPAY
          If Round((TempAptTo.AMOUNT - TempApBktTrn.AMOUNT - TempAptCash.CSHPAY), 2) <> 0 Then
             HaveValue = True
          End If

          If HaveValue Then
            If mcolParam("SUMMARY_FLAG") = "N" And mcolParam("ONLY_SUPPLIER") = "N" Then
               Call m_Details1.GetString(1, TempStr1, TempStr2)
               strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
               Call Vsp.AddTable(strFormat, "", TempStr2)
             End If
             HaveValueSum = True
          End If
          ' ของเดือนแรกที่เลือก 8
      End If

      If (RunMonth >= 9) Then
         ' ของเดือนแรกที่เลือก 9

         FirstDate = DateAdd("M", 1, FirstDate)
         Call m_Details1.ClearField

         Set TempApt = GetAPTrn(m_ApSupAmounts9, SupGroup.SUPPLIER_CODE)
         Set TempAptCn = GetAPTrn(m_ApSupAmountCns9, SupGroup.SUPPLIER_CODE)

         Set TempApRcIt = GetAPRcpItem(m_ApSupPaids9, SupGroup.SUPPLIER_CODE)
         Set TempApRcItCn = GetAPRcpItem(m_ApSupPaidCns9, SupGroup.SUPPLIER_CODE)

         Set TempApRcIt3 = GetAPRcpItemEx(m_PaidAllAmounts9, SupGroup.SUPPLIER_CODE)
         Set TempApRcIt3Cn = GetAPRcpItemEx(m_PaidAllAmountCns9, SupGroup.SUPPLIER_CODE)

         HaveValue = False

          '===Head
          Set HeadCf = m_Headers1.Fields(1)
          TempStr = Format(Month(FirstDate), "00") & "/" & Format(Year(FirstDate) + 543, "00")
          Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
          Call m_Details1.AddField(BodyCf)

          Set Ap = GetObject("CAPMas", Supplier, SupGroup.SUPPLIER_CODE)
          Set HeadCf = m_Headers1.Fields(2)
          TempStr = Left(Ap.SUPCOD & "/" & Ap.SUPNAM, 30)
          Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
          Call m_Details1.AddField(BodyCf)

          BalanceAmount = (TempApt.AMOUNT - TempAptCn.AMOUNT) - (TempApt.PAYAMT - TempAptCn.PAYAMT) - (TempApRcIt.PAYAMT - TempApRcItCn.PAYAMT) + (TempApRcIt3.PAYAMT - TempApRcIt3Cn.PAYAMT)

          K = 2
          K = K + 1
          Set HeadCf = m_Headers1.Fields(K)
          TempStr = FormatNumber(BalanceAmount, "")
          Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
          Call m_Details1.AddField(BodyCf)

          Total19(K) = Total19(K) + BalanceAmount
          If Round((BalanceAmount), 2) <> 0 Then
             HaveValue = True
          End If

          Set TempAptTo = GetAPTrn(m_ApSupAmountTo9, SupGroup.SUPPLIER_CODE)
          Set TempAptToCn = GetAPTrn(m_ApSupAmountToCn9, SupGroup.SUPPLIER_CODE)

          Set TempApRcItTo = GetAPRcpItem(m_ApSupPaidTo9, SupGroup.SUPPLIER_CODE)
          Set TempApRcItToCn = GetAPRcpItem(m_ApSupPaidToCn9, SupGroup.SUPPLIER_CODE)

          K = K + 1
          Set HeadCf = m_Headers1.Fields(K)
          TempStr = FormatNumber(TempAptTo.AMOUNT, "")
          Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
          Call m_Details1.AddField(BodyCf)

          Total19(K) = Total19(K) + TempAptTo.AMOUNT
          Total3(K) = Total3(K) + TempAptTo.AMOUNT
          If Round((TempAptTo.AMOUNT), 2) <> 0 Then
             HaveValue = True
          End If

          K = K + 1
          Set HeadCf = m_Headers1.Fields(K)
          TempStr = FormatNumber(TempApRcItTo.PAYAMT - TempApRcItToCn.PAYAMT + TempAptToCn.AMOUNT, "")
          Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
          Call m_Details1.AddField(BodyCf)

          Total19(K) = Total19(K) + TempApRcItTo.PAYAMT - TempApRcItToCn.PAYAMT + TempAptToCn.AMOUNT
          Total3(K) = Total3(K) + TempApRcItTo.PAYAMT - TempApRcItToCn.PAYAMT + TempAptToCn.AMOUNT
          If Round((TempApRcItTo.PAYAMT - TempApRcItToCn.PAYAMT + TempAptToCn.AMOUNT), 2) <> 0 Then
             HaveValue = True
          End If
          TempPaidCheck = Round((TempApRcItTo.PAYAMT - TempApRcItToCn.PAYAMT + TempAptToCn.AMOUNT), 2)

          K = K + 1
          Set HeadCf = m_Headers1.Fields(K)
          TempStr = FormatNumber(BalanceAmount + (TempAptTo.AMOUNT - TempAptToCn.AMOUNT) - (TempApRcItTo.PAYAMT - TempApRcItToCn.PAYAMT), "")
          Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
          Call m_Details1.AddField(BodyCf)

          Total19(K) = Total19(K) + (BalanceAmount) + (TempAptTo.AMOUNT - TempAptToCn.AMOUNT) - (TempApRcItTo.PAYAMT - TempApRcItToCn.PAYAMT)
          If Round(((BalanceAmount) + (TempAptTo.AMOUNT - TempAptToCn.AMOUNT) - (TempApRcItTo.PAYAMT - TempApRcItToCn.PAYAMT)), 2) <> 0 Then
             HaveValue = True
          End If

          Set TempApBktTrn = GetBkTrn(m_59, SupGroup.SUPPLIER_CODE)
          Set TempAptCash = GetAPTrn(m_ApSupAmountCashs9, SupGroup.SUPPLIER_CODE)

          K = K + 1
          Set HeadCf = m_Headers1.Fields(K)
          TempStr = FormatNumber(TempApBktTrn.AMOUNT + TempAptCash.CSHPAY, "")
          Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
          Call m_Details1.AddField(BodyCf)

          Total19(K) = Total19(K) + TempApBktTrn.AMOUNT + TempAptCash.CSHPAY
          Total3(K) = Total3(K) + TempApBktTrn.AMOUNT + TempAptCash.CSHPAY
          If Round((TempApBktTrn.AMOUNT + TempAptCash.CSHPAY), 2) <> 0 Then
             HaveValue = True
          End If

          Set TempApBktTrn = GetBkTrn(m_69, SupGroup.SUPPLIER_CODE)
          K = K + 1
          Set HeadCf = m_Headers1.Fields(K)
          TempStr = FormatNumber(TempApBktTrn.AMOUNT, "")
          Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
          Call m_Details1.AddField(BodyCf)

          Total19(K) = Total19(K) + TempApBktTrn.AMOUNT
          Total3(K) = TempApBktTrn.AMOUNT
          If Round((TempApBktTrn.AMOUNT), 2) <> 0 Then
             HaveValue = True
          End If

          '7 =  4 + 6
          K = K + 1
          Set HeadCf = m_Headers1.Fields(K)
          TempStr = FormatNumber(BalanceAmount + (TempAptTo.AMOUNT - TempAptToCn.AMOUNT) - (TempApRcItTo.PAYAMT - TempApRcItToCn.PAYAMT) + TempApBktTrn.AMOUNT, "")
          Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
          Call m_Details1.AddField(BodyCf)

          Total19(K) = Total19(K) + BalanceAmount + (TempAptTo.AMOUNT - TempAptToCn.AMOUNT) - (TempApRcItTo.PAYAMT - TempApRcItToCn.PAYAMT) + TempApBktTrn.AMOUNT
          Total3(K) = BalanceAmount + (TempAptTo.AMOUNT - TempAptToCn.AMOUNT) - (TempApRcItTo.PAYAMT - TempApRcItToCn.PAYAMT) + TempApBktTrn.AMOUNT
          If Round((BalanceAmount + (TempAptTo.AMOUNT - TempAptToCn.AMOUNT) - (TempApRcItTo.PAYAMT - TempApRcItToCn.PAYAMT) + TempApBktTrn.AMOUNT), 2) <> 0 Then
             HaveValue = True
          End If

         Set TempApBktTrn = GetBkTrn(m_59, SupGroup.SUPPLIER_CODE)
          Set TempAptCash = GetAPTrn(m_ApSupAmountCashs9, SupGroup.SUPPLIER_CODE)

          K = K + 1
          Set HeadCf = m_Headers1.Fields(K)
          TempStr = FormatNumber(TempAptTo.AMOUNT - TempApBktTrn.AMOUNT - TempAptCash.CSHPAY, "")
          Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
          Call m_Details1.AddField(BodyCf)

          Total19(K) = Total19(K) + TempAptTo.AMOUNT - TempApBktTrn.AMOUNT - TempAptCash.CSHPAY
          Total3(K) = Total3(K) + TempAptTo.AMOUNT - TempApBktTrn.AMOUNT - TempAptCash.CSHPAY
          If Round((TempAptTo.AMOUNT - TempApBktTrn.AMOUNT - TempAptCash.CSHPAY), 2) <> 0 Then
             HaveValue = True
          End If

          If HaveValue Then
            If mcolParam("SUMMARY_FLAG") = "N" And mcolParam("ONLY_SUPPLIER") = "N" Then
               Call m_Details1.GetString(1, TempStr1, TempStr2)
               strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
               Call Vsp.AddTable(strFormat, "", TempStr2)
             End If
             HaveValueSum = True
          End If
          ' ของเดือนแรกที่เลือก 9
      End If

      If (RunMonth >= 10) Then
         ' ของเดือนแรกที่เลือก 10

         FirstDate = DateAdd("M", 1, FirstDate)
         Call m_Details1.ClearField

         Set TempApt = GetAPTrn(m_ApSupAmounts10, SupGroup.SUPPLIER_CODE)
         Set TempAptCn = GetAPTrn(m_ApSupAmountCns10, SupGroup.SUPPLIER_CODE)

         Set TempApRcIt = GetAPRcpItem(m_ApSupPaids10, SupGroup.SUPPLIER_CODE)
         Set TempApRcItCn = GetAPRcpItem(m_ApSupPaidCns10, SupGroup.SUPPLIER_CODE)

         Set TempApRcIt3 = GetAPRcpItemEx(m_PaidAllAmounts10, SupGroup.SUPPLIER_CODE)
         Set TempApRcIt3Cn = GetAPRcpItemEx(m_PaidAllAmountCns10, SupGroup.SUPPLIER_CODE)

         HaveValue = False

          '===Head
          Set HeadCf = m_Headers1.Fields(1)
          TempStr = Format(Month(FirstDate), "00") & "/" & Format(Year(FirstDate) + 543, "00")
          Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
          Call m_Details1.AddField(BodyCf)

          Set Ap = GetObject("CAPMas", Supplier, SupGroup.SUPPLIER_CODE)
          Set HeadCf = m_Headers1.Fields(2)
          TempStr = Left(Ap.SUPCOD & "/" & Ap.SUPNAM, 30)
          Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
          Call m_Details1.AddField(BodyCf)

          BalanceAmount = (TempApt.AMOUNT - TempAptCn.AMOUNT) - (TempApt.PAYAMT - TempAptCn.PAYAMT) - (TempApRcIt.PAYAMT - TempApRcItCn.PAYAMT) + (TempApRcIt3.PAYAMT - TempApRcIt3Cn.PAYAMT)

          K = 2
          K = K + 1
          Set HeadCf = m_Headers1.Fields(K)
          TempStr = FormatNumber(BalanceAmount, "")
          Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
          Call m_Details1.AddField(BodyCf)

          Total110(K) = Total110(K) + BalanceAmount
          If Round((BalanceAmount), 2) <> 0 Then
             HaveValue = True
          End If

          Set TempAptTo = GetAPTrn(m_ApSupAmountTo10, SupGroup.SUPPLIER_CODE)
          Set TempAptToCn = GetAPTrn(m_ApSupAmountToCn10, SupGroup.SUPPLIER_CODE)

          Set TempApRcItTo = GetAPRcpItem(m_ApSupPaidTo10, SupGroup.SUPPLIER_CODE)
          Set TempApRcItToCn = GetAPRcpItem(m_ApSupPaidToCn10, SupGroup.SUPPLIER_CODE)

          K = K + 1
          Set HeadCf = m_Headers1.Fields(K)
          TempStr = FormatNumber(TempAptTo.AMOUNT, "")
          Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
          Call m_Details1.AddField(BodyCf)

          Total110(K) = Total110(K) + TempAptTo.AMOUNT
          Total3(K) = Total3(K) + TempAptTo.AMOUNT
          If Round((TempAptTo.AMOUNT), 2) <> 0 Then
             HaveValue = True
          End If

          K = K + 1
          Set HeadCf = m_Headers1.Fields(K)
          TempStr = FormatNumber(TempApRcItTo.PAYAMT - TempApRcItToCn.PAYAMT + TempAptToCn.AMOUNT, "")
          Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
          Call m_Details1.AddField(BodyCf)

          Total110(K) = Total110(K) + TempApRcItTo.PAYAMT - TempApRcItToCn.PAYAMT + TempAptToCn.AMOUNT
          Total3(K) = Total3(K) + TempApRcItTo.PAYAMT - TempApRcItToCn.PAYAMT + TempAptToCn.AMOUNT
          If Round((TempApRcItTo.PAYAMT - TempApRcItToCn.PAYAMT + TempAptToCn.AMOUNT), 2) <> 0 Then
             HaveValue = True
          End If
          TempPaidCheck = Round((TempApRcItTo.PAYAMT - TempApRcItToCn.PAYAMT + TempAptToCn.AMOUNT), 2)

          K = K + 1
          Set HeadCf = m_Headers1.Fields(K)
          TempStr = FormatNumber(BalanceAmount + (TempAptTo.AMOUNT - TempAptToCn.AMOUNT) - (TempApRcItTo.PAYAMT - TempApRcItToCn.PAYAMT), "")
          Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
          Call m_Details1.AddField(BodyCf)

          Total110(K) = Total110(K) + (BalanceAmount) + (TempAptTo.AMOUNT - TempAptToCn.AMOUNT) - (TempApRcItTo.PAYAMT - TempApRcItToCn.PAYAMT)
          If Round(((BalanceAmount) + (TempAptTo.AMOUNT - TempAptToCn.AMOUNT) - (TempApRcItTo.PAYAMT - TempApRcItToCn.PAYAMT)), 2) <> 0 Then
             HaveValue = True
          End If

          Set TempApBktTrn = GetBkTrn(m_510, SupGroup.SUPPLIER_CODE)
          Set TempAptCash = GetAPTrn(m_ApSupAmountCashs10, SupGroup.SUPPLIER_CODE)

          K = K + 1
          Set HeadCf = m_Headers1.Fields(K)
          TempStr = FormatNumber(TempApBktTrn.AMOUNT + TempAptCash.CSHPAY, "")
          Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
          Call m_Details1.AddField(BodyCf)

          Total110(K) = Total110(K) + TempApBktTrn.AMOUNT + TempAptCash.CSHPAY
          Total3(K) = Total3(K) + TempApBktTrn.AMOUNT + TempAptCash.CSHPAY
          If Round((TempApBktTrn.AMOUNT + TempAptCash.CSHPAY), 2) <> 0 Then
             HaveValue = True
          End If

          Set TempApBktTrn = GetBkTrn(m_610, SupGroup.SUPPLIER_CODE)
          K = K + 1
          Set HeadCf = m_Headers1.Fields(K)
          TempStr = FormatNumber(TempApBktTrn.AMOUNT, "")
          Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
          Call m_Details1.AddField(BodyCf)

          Total110(K) = Total110(K) + TempApBktTrn.AMOUNT
          Total3(K) = TempApBktTrn.AMOUNT
          If Round((TempApBktTrn.AMOUNT), 2) <> 0 Then
             HaveValue = True
          End If

          '7 =  4 + 6
          K = K + 1
          Set HeadCf = m_Headers1.Fields(K)
          TempStr = FormatNumber(BalanceAmount + (TempAptTo.AMOUNT - TempAptToCn.AMOUNT) - (TempApRcItTo.PAYAMT - TempApRcItToCn.PAYAMT) + TempApBktTrn.AMOUNT, "")
          Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
          Call m_Details1.AddField(BodyCf)

          Total110(K) = Total110(K) + BalanceAmount + (TempAptTo.AMOUNT - TempAptToCn.AMOUNT) - (TempApRcItTo.PAYAMT - TempApRcItToCn.PAYAMT) + TempApBktTrn.AMOUNT
          Total3(K) = BalanceAmount + (TempAptTo.AMOUNT - TempAptToCn.AMOUNT) - (TempApRcItTo.PAYAMT - TempApRcItToCn.PAYAMT) + TempApBktTrn.AMOUNT
          If Round((BalanceAmount + (TempAptTo.AMOUNT - TempAptToCn.AMOUNT) - (TempApRcItTo.PAYAMT - TempApRcItToCn.PAYAMT) + TempApBktTrn.AMOUNT), 2) <> 0 Then
             HaveValue = True
          End If

         Set TempApBktTrn = GetBkTrn(m_510, SupGroup.SUPPLIER_CODE)
          Set TempAptCash = GetAPTrn(m_ApSupAmountCashs10, SupGroup.SUPPLIER_CODE)

          K = K + 1
          Set HeadCf = m_Headers1.Fields(K)
          TempStr = FormatNumber(TempAptTo.AMOUNT - TempApBktTrn.AMOUNT - TempAptCash.CSHPAY, "")
          Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
          Call m_Details1.AddField(BodyCf)

          Total110(K) = Total110(K) + TempAptTo.AMOUNT - TempApBktTrn.AMOUNT - TempAptCash.CSHPAY
          Total3(K) = Total3(K) + TempAptTo.AMOUNT - TempApBktTrn.AMOUNT - TempAptCash.CSHPAY
          If Round((TempAptTo.AMOUNT - TempApBktTrn.AMOUNT - TempAptCash.CSHPAY), 2) <> 0 Then
             HaveValue = True
          End If

          If HaveValue Then
            If mcolParam("SUMMARY_FLAG") = "N" And mcolParam("ONLY_SUPPLIER") = "N" Then
               Call m_Details1.GetString(1, TempStr1, TempStr2)
               strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
               Call Vsp.AddTable(strFormat, "", TempStr2)
             End If
             HaveValueSum = True
          End If
          ' ของเดือนแรกที่เลือก 10
      End If

      If (RunMonth >= 11) Then
         ' ของเดือนแรกที่เลือก 11

         FirstDate = DateAdd("M", 1, FirstDate)
         Call m_Details1.ClearField

         Set TempApt = GetAPTrn(m_ApSupAmounts11, SupGroup.SUPPLIER_CODE)
         Set TempAptCn = GetAPTrn(m_ApSupAmountCns11, SupGroup.SUPPLIER_CODE)

         Set TempApRcIt = GetAPRcpItem(m_ApSupPaids11, SupGroup.SUPPLIER_CODE)
         Set TempApRcItCn = GetAPRcpItem(m_ApSupPaidCns11, SupGroup.SUPPLIER_CODE)

         Set TempApRcIt3 = GetAPRcpItemEx(m_PaidAllAmounts11, SupGroup.SUPPLIER_CODE)
         Set TempApRcIt3Cn = GetAPRcpItemEx(m_PaidAllAmountCns11, SupGroup.SUPPLIER_CODE)

         HaveValue = False

          '===Head
          Set HeadCf = m_Headers1.Fields(1)
          TempStr = Format(Month(FirstDate), "00") & "/" & Format(Year(FirstDate) + 543, "00")
          Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
          Call m_Details1.AddField(BodyCf)

          Set Ap = GetObject("CAPMas", Supplier, SupGroup.SUPPLIER_CODE)
          Set HeadCf = m_Headers1.Fields(2)
          TempStr = Left(Ap.SUPCOD & "/" & Ap.SUPNAM, 30)
          Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
          Call m_Details1.AddField(BodyCf)

          BalanceAmount = (TempApt.AMOUNT - TempAptCn.AMOUNT) - (TempApt.PAYAMT - TempAptCn.PAYAMT) - (TempApRcIt.PAYAMT - TempApRcItCn.PAYAMT) + (TempApRcIt3.PAYAMT - TempApRcIt3Cn.PAYAMT)

          K = 2
          K = K + 1
          Set HeadCf = m_Headers1.Fields(K)
          TempStr = FormatNumber(BalanceAmount, "")
          Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
          Call m_Details1.AddField(BodyCf)

          Total111(K) = Total111(K) + BalanceAmount
          If Round((BalanceAmount), 2) <> 0 Then
             HaveValue = True
          End If

          Set TempAptTo = GetAPTrn(m_ApSupAmountTo11, SupGroup.SUPPLIER_CODE)
          Set TempAptToCn = GetAPTrn(m_ApSupAmountToCn11, SupGroup.SUPPLIER_CODE)

          Set TempApRcItTo = GetAPRcpItem(m_ApSupPaidTo11, SupGroup.SUPPLIER_CODE)
          Set TempApRcItToCn = GetAPRcpItem(m_ApSupPaidToCn11, SupGroup.SUPPLIER_CODE)

          K = K + 1
          Set HeadCf = m_Headers1.Fields(K)
          TempStr = FormatNumber(TempAptTo.AMOUNT, "")
          Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
          Call m_Details1.AddField(BodyCf)

          Total111(K) = Total111(K) + TempAptTo.AMOUNT
          Total3(K) = Total3(K) + TempAptTo.AMOUNT
          If Round((TempAptTo.AMOUNT), 2) <> 0 Then
             HaveValue = True
          End If

          K = K + 1
          Set HeadCf = m_Headers1.Fields(K)
          TempStr = FormatNumber(TempApRcItTo.PAYAMT - TempApRcItToCn.PAYAMT + TempAptToCn.AMOUNT, "")
          Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
          Call m_Details1.AddField(BodyCf)

          Total111(K) = Total111(K) + TempApRcItTo.PAYAMT - TempApRcItToCn.PAYAMT + TempAptToCn.AMOUNT
          Total3(K) = Total3(K) + TempApRcItTo.PAYAMT - TempApRcItToCn.PAYAMT + TempAptToCn.AMOUNT
          If Round((TempApRcItTo.PAYAMT - TempApRcItToCn.PAYAMT + TempAptToCn.AMOUNT), 2) <> 0 Then
             HaveValue = True
          End If
          TempPaidCheck = Round((TempApRcItTo.PAYAMT - TempApRcItToCn.PAYAMT + TempAptToCn.AMOUNT), 2)

          K = K + 1
          Set HeadCf = m_Headers1.Fields(K)
          TempStr = FormatNumber(BalanceAmount + (TempAptTo.AMOUNT - TempAptToCn.AMOUNT) - (TempApRcItTo.PAYAMT - TempApRcItToCn.PAYAMT), "")
          Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
          Call m_Details1.AddField(BodyCf)

          Total111(K) = Total111(K) + (BalanceAmount) + (TempAptTo.AMOUNT - TempAptToCn.AMOUNT) - (TempApRcItTo.PAYAMT - TempApRcItToCn.PAYAMT)
          If Round(((BalanceAmount) + (TempAptTo.AMOUNT - TempAptToCn.AMOUNT) - (TempApRcItTo.PAYAMT - TempApRcItToCn.PAYAMT)), 2) <> 0 Then
             HaveValue = True
          End If

          Set TempApBktTrn = GetBkTrn(m_511, SupGroup.SUPPLIER_CODE)
          Set TempAptCash = GetAPTrn(m_ApSupAmountCashs11, SupGroup.SUPPLIER_CODE)

          K = K + 1
          Set HeadCf = m_Headers1.Fields(K)
          TempStr = FormatNumber(TempApBktTrn.AMOUNT + TempAptCash.CSHPAY, "")
          Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
          Call m_Details1.AddField(BodyCf)

          Total111(K) = Total111(K) + TempApBktTrn.AMOUNT + TempAptCash.CSHPAY
          Total3(K) = Total3(K) + TempApBktTrn.AMOUNT + TempAptCash.CSHPAY
          If Round((TempApBktTrn.AMOUNT + TempAptCash.CSHPAY), 2) <> 0 Then
             HaveValue = True
          End If

          Set TempApBktTrn = GetBkTrn(m_611, SupGroup.SUPPLIER_CODE)
          K = K + 1
          Set HeadCf = m_Headers1.Fields(K)
          TempStr = FormatNumber(TempApBktTrn.AMOUNT, "")
          Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
          Call m_Details1.AddField(BodyCf)

          Total111(K) = Total111(K) + TempApBktTrn.AMOUNT
          Total3(K) = TempApBktTrn.AMOUNT
          If Round((TempApBktTrn.AMOUNT), 2) <> 0 Then
             HaveValue = True
          End If

          '7 =  4 + 6
          K = K + 1
          Set HeadCf = m_Headers1.Fields(K)
          TempStr = FormatNumber(BalanceAmount + (TempAptTo.AMOUNT - TempAptToCn.AMOUNT) - (TempApRcItTo.PAYAMT - TempApRcItToCn.PAYAMT) + TempApBktTrn.AMOUNT, "")
          Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
          Call m_Details1.AddField(BodyCf)

          Total111(K) = Total111(K) + BalanceAmount + (TempAptTo.AMOUNT - TempAptToCn.AMOUNT) - (TempApRcItTo.PAYAMT - TempApRcItToCn.PAYAMT) + TempApBktTrn.AMOUNT
          Total3(K) = BalanceAmount + (TempAptTo.AMOUNT - TempAptToCn.AMOUNT) - (TempApRcItTo.PAYAMT - TempApRcItToCn.PAYAMT) + TempApBktTrn.AMOUNT
          If Round((BalanceAmount + (TempAptTo.AMOUNT - TempAptToCn.AMOUNT) - (TempApRcItTo.PAYAMT - TempApRcItToCn.PAYAMT) + TempApBktTrn.AMOUNT), 2) <> 0 Then
             HaveValue = True
          End If

         Set TempApBktTrn = GetBkTrn(m_511, SupGroup.SUPPLIER_CODE)
          Set TempAptCash = GetAPTrn(m_ApSupAmountCashs11, SupGroup.SUPPLIER_CODE)

          K = K + 1
          Set HeadCf = m_Headers1.Fields(K)
          TempStr = FormatNumber(TempAptTo.AMOUNT - TempApBktTrn.AMOUNT - TempAptCash.CSHPAY, "")
          Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
          Call m_Details1.AddField(BodyCf)

          Total111(K) = Total111(K) + TempAptTo.AMOUNT - TempApBktTrn.AMOUNT - TempAptCash.CSHPAY
          Total3(K) = Total3(K) + TempAptTo.AMOUNT - TempApBktTrn.AMOUNT - TempAptCash.CSHPAY
          If Round((TempAptTo.AMOUNT - TempApBktTrn.AMOUNT - TempAptCash.CSHPAY), 2) <> 0 Then
             HaveValue = True
          End If

          If HaveValue Then
            If mcolParam("SUMMARY_FLAG") = "N" And mcolParam("ONLY_SUPPLIER") = "N" Then
               Call m_Details1.GetString(1, TempStr1, TempStr2)
               strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
               Call Vsp.AddTable(strFormat, "", TempStr2)
             End If
             HaveValueSum = True
          End If
          ' ของเดือนแรกที่เลือก 11
      End If

      If (RunMonth >= 12) Then
         ' ของเดือนแรกที่เลือก 12

         FirstDate = DateAdd("M", 1, FirstDate)
         Call m_Details1.ClearField

         Set TempApt = GetAPTrn(m_ApSupAmounts12, SupGroup.SUPPLIER_CODE)
         Set TempAptCn = GetAPTrn(m_ApSupAmountCns12, SupGroup.SUPPLIER_CODE)

         Set TempApRcIt = GetAPRcpItem(m_ApSupPaids12, SupGroup.SUPPLIER_CODE)
         Set TempApRcItCn = GetAPRcpItem(m_ApSupPaidCns12, SupGroup.SUPPLIER_CODE)

         Set TempApRcIt3 = GetAPRcpItemEx(m_PaidAllAmounts12, SupGroup.SUPPLIER_CODE)
         Set TempApRcIt3Cn = GetAPRcpItemEx(m_PaidAllAmountCns12, SupGroup.SUPPLIER_CODE)

         HaveValue = False

          '===Head
          Set HeadCf = m_Headers1.Fields(1)
          TempStr = Format(Month(FirstDate), "00") & "/" & Format(Year(FirstDate) + 543, "00")
          Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
          Call m_Details1.AddField(BodyCf)

          Set Ap = GetObject("CAPMas", Supplier, SupGroup.SUPPLIER_CODE)
          Set HeadCf = m_Headers1.Fields(2)
          TempStr = Left(Ap.SUPCOD & "/" & Ap.SUPNAM, 30)
          Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
          Call m_Details1.AddField(BodyCf)

          BalanceAmount = (TempApt.AMOUNT - TempAptCn.AMOUNT) - (TempApt.PAYAMT - TempAptCn.PAYAMT) - (TempApRcIt.PAYAMT - TempApRcItCn.PAYAMT) + (TempApRcIt3.PAYAMT - TempApRcIt3Cn.PAYAMT)

          K = 2
          K = K + 1
          Set HeadCf = m_Headers1.Fields(K)
          TempStr = FormatNumber(BalanceAmount, "")
          Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
          Call m_Details1.AddField(BodyCf)

          Total112(K) = Total112(K) + BalanceAmount
          If Round((BalanceAmount), 2) <> 0 Then
             HaveValue = True
          End If

          Set TempAptTo = GetAPTrn(m_ApSupAmountTo12, SupGroup.SUPPLIER_CODE)
          Set TempAptToCn = GetAPTrn(m_ApSupAmountToCn12, SupGroup.SUPPLIER_CODE)

          Set TempApRcItTo = GetAPRcpItem(m_ApSupPaidTo12, SupGroup.SUPPLIER_CODE)
          Set TempApRcItToCn = GetAPRcpItem(m_ApSupPaidToCn12, SupGroup.SUPPLIER_CODE)

          K = K + 1
          Set HeadCf = m_Headers1.Fields(K)
          TempStr = FormatNumber(TempAptTo.AMOUNT, "")
          Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
          Call m_Details1.AddField(BodyCf)

          Total112(K) = Total112(K) + TempAptTo.AMOUNT
          Total3(K) = Total3(K) + TempAptTo.AMOUNT
          If Round((TempAptTo.AMOUNT), 2) <> 0 Then
             HaveValue = True
          End If

          K = K + 1
          Set HeadCf = m_Headers1.Fields(K)
          TempStr = FormatNumber(TempApRcItTo.PAYAMT - TempApRcItToCn.PAYAMT + TempAptToCn.AMOUNT, "")
          Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
          Call m_Details1.AddField(BodyCf)

          Total112(K) = Total112(K) + TempApRcItTo.PAYAMT - TempApRcItToCn.PAYAMT + TempAptToCn.AMOUNT
          Total3(K) = Total3(K) + TempApRcItTo.PAYAMT - TempApRcItToCn.PAYAMT + TempAptToCn.AMOUNT
          If Round((TempApRcItTo.PAYAMT - TempApRcItToCn.PAYAMT + TempAptToCn.AMOUNT), 2) <> 0 Then
             HaveValue = True
          End If
          TempPaidCheck = Round((TempApRcItTo.PAYAMT - TempApRcItToCn.PAYAMT + TempAptToCn.AMOUNT), 2)

          K = K + 1
          Set HeadCf = m_Headers1.Fields(K)
          TempStr = FormatNumber(BalanceAmount + (TempAptTo.AMOUNT - TempAptToCn.AMOUNT) - (TempApRcItTo.PAYAMT - TempApRcItToCn.PAYAMT), "")
          Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
          Call m_Details1.AddField(BodyCf)

          Total112(K) = Total112(K) + (BalanceAmount) + (TempAptTo.AMOUNT - TempAptToCn.AMOUNT) - (TempApRcItTo.PAYAMT - TempApRcItToCn.PAYAMT)
          If Round(((BalanceAmount) + (TempAptTo.AMOUNT - TempAptToCn.AMOUNT) - (TempApRcItTo.PAYAMT - TempApRcItToCn.PAYAMT)), 2) <> 0 Then
             HaveValue = True
          End If

          Set TempApBktTrn = GetBkTrn(m_512, SupGroup.SUPPLIER_CODE)
          Set TempAptCash = GetAPTrn(m_ApSupAmountCashs12, SupGroup.SUPPLIER_CODE)

          K = K + 1
          Set HeadCf = m_Headers1.Fields(K)
          TempStr = FormatNumber(TempApBktTrn.AMOUNT + TempAptCash.CSHPAY, "")
          Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
          Call m_Details1.AddField(BodyCf)

          Total112(K) = Total112(K) + TempApBktTrn.AMOUNT + TempAptCash.CSHPAY
          Total3(K) = Total3(K) + TempApBktTrn.AMOUNT + TempAptCash.CSHPAY
          If Round((TempApBktTrn.AMOUNT + TempAptCash.CSHPAY), 2) <> 0 Then
             HaveValue = True
          End If

          Set TempApBktTrn = GetBkTrn(m_612, SupGroup.SUPPLIER_CODE)
          K = K + 1
          Set HeadCf = m_Headers1.Fields(K)
          TempStr = FormatNumber(TempApBktTrn.AMOUNT, "")
          Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
          Call m_Details1.AddField(BodyCf)

          Total112(K) = Total112(K) + TempApBktTrn.AMOUNT
          Total3(K) = TempApBktTrn.AMOUNT
          If Round((TempApBktTrn.AMOUNT), 2) <> 0 Then
             HaveValue = True
          End If

          '7 =  4 + 6
          K = K + 1
          Set HeadCf = m_Headers1.Fields(K)
          TempStr = FormatNumber(BalanceAmount + (TempAptTo.AMOUNT - TempAptToCn.AMOUNT) - (TempApRcItTo.PAYAMT - TempApRcItToCn.PAYAMT) + TempApBktTrn.AMOUNT, "")
          Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
          Call m_Details1.AddField(BodyCf)

          Total112(K) = Total112(K) + BalanceAmount + (TempAptTo.AMOUNT - TempAptToCn.AMOUNT) - (TempApRcItTo.PAYAMT - TempApRcItToCn.PAYAMT) + TempApBktTrn.AMOUNT
          Total3(K) = BalanceAmount + (TempAptTo.AMOUNT - TempAptToCn.AMOUNT) - (TempApRcItTo.PAYAMT - TempApRcItToCn.PAYAMT) + TempApBktTrn.AMOUNT
          If Round((BalanceAmount + (TempAptTo.AMOUNT - TempAptToCn.AMOUNT) - (TempApRcItTo.PAYAMT - TempApRcItToCn.PAYAMT) + TempApBktTrn.AMOUNT), 2) <> 0 Then
             HaveValue = True
          End If

         Set TempApBktTrn = GetBkTrn(m_512, SupGroup.SUPPLIER_CODE)
          Set TempAptCash = GetAPTrn(m_ApSupAmountCashs12, SupGroup.SUPPLIER_CODE)

          K = K + 1
          Set HeadCf = m_Headers1.Fields(K)
          TempStr = FormatNumber(TempAptTo.AMOUNT - TempApBktTrn.AMOUNT - TempAptCash.CSHPAY, "")
          Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
          Call m_Details1.AddField(BodyCf)

          Total112(K) = Total112(K) + TempAptTo.AMOUNT - TempApBktTrn.AMOUNT - TempAptCash.CSHPAY
          Total3(K) = Total3(K) + TempAptTo.AMOUNT - TempApBktTrn.AMOUNT - TempAptCash.CSHPAY
          If Round((TempAptTo.AMOUNT - TempApBktTrn.AMOUNT - TempAptCash.CSHPAY), 2) <> 0 Then
             HaveValue = True
          End If

          If HaveValue Then
            If mcolParam("SUMMARY_FLAG") = "N" And mcolParam("ONLY_SUPPLIER") = "N" Then
               Call m_Details1.GetString(1, TempStr1, TempStr2)
               strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
               Call Vsp.AddTable(strFormat, "", TempStr2)
             End If
             HaveValueSum = True
          End If
          ' ของเดือนแรกที่เลือก 12
      End If
      
      If (RunMonth >= 2) Then
         ' รวม
         Call m_Details1.ClearField
               
          '===Head
          Set HeadCf = m_Headers1.Fields(1)
          TempStr = "รวม"
          Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
          Call m_Details1.AddField(BodyCf)
          
          Set Ap = GetObject("CAPMas", Supplier, SupGroup.SUPPLIER_CODE)
          Set HeadCf = m_Headers1.Fields(2)
          TempStr = Left(Ap.SUPCOD & "/" & Ap.SUPNAM, 30)
          Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
          Call m_Details1.AddField(BodyCf)
          
          K = 2
          K = K + 1
          Set HeadCf = m_Headers1.Fields(K)
          TempStr = FormatNumber(Total3(K), "")
          Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
          Call m_Details1.AddField(BodyCf)
          Total2(K) = Total2(K) + Total3(K)
          If Round((Total3(K)), 2) <> 0 Then
             HaveValue = True
          End If
          
          K = K + 1
          Set HeadCf = m_Headers1.Fields(K)
          TempValue = Total3(K)
          TempStr = FormatNumber(TempValue, "")
          Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
          Call m_Details1.AddField(BodyCf)
          Total2(K) = Total2(K) + Total3(K)
          If Round((Total3(K)), 2) <> 0 Then
             HaveValue = True
          End If
          
          K = K + 1
          Set HeadCf = m_Headers1.Fields(K)
          TempValue = Total3(K)
          TempStr = FormatNumber(TempValue, "")
          Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
          Call m_Details1.AddField(BodyCf)
          Total2(K) = Total2(K) + Total3(K)
          If Round((Total3(K)), 2) <> 0 Then
             HaveValue = True
          End If
          
          K = K + 1
          Set HeadCf = m_Headers1.Fields(K)
          TempValue = Total3(K - 3) + Total3(K - 2) - Total3(K - 1)
          TempStr = FormatNumber(TempValue, "")
          Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
          Call m_Details1.AddField(BodyCf)
          Total2(K) = Total2(K) + Total3(K - 3) + Total3(K - 2) - Total3(K - 1)
          If Round((Total3(K - 3) + Total3(K - 2) - Total3(K - 1)), 2) <> 0 Then
             HaveValue = True
          End If
          
          K = K + 1
          Set HeadCf = m_Headers1.Fields(K)
          TempValue = Total3(K)
          TempStr = FormatNumber(TempValue, "")
          Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
          Call m_Details1.AddField(BodyCf)
          Total2(K) = Total2(K) + Total3(K)
          If Round((Total3(K)), 2) <> 0 Then
             HaveValue = True
          End If
          
          K = K + 1
          Set HeadCf = m_Headers1.Fields(K)
          TempValue = Total3(K)
          TempStr = FormatNumber(TempValue, "")
          Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
          Call m_Details1.AddField(BodyCf)
          Total2(K) = Total2(K) + Total3(K)
          If Round((Total3(K)), 2) <> 0 Then
             HaveValue = True
          End If
          
          '7 =  4 + 6
          K = K + 1
          Set HeadCf = m_Headers1.Fields(K)
          TempValue = Total3(K)
          TempStr = FormatNumber(TempValue, "")
          Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
          Call m_Details1.AddField(BodyCf)
          Total2(K) = Total2(K) + Total3(K)
          If Round((Total3(K)), 2) <> 0 Then
             HaveValue = True
          End If
          
          K = K + 1
          Set HeadCf = m_Headers1.Fields(K)
          TempValue = Total3(K)
          TempStr = FormatNumber(TempValue, "")
          Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
          Call m_Details1.AddField(BodyCf)
          Total2(K) = Total2(K) + Total3(K)
          If Round((Total3(K)), 2) <> 0 Then
             HaveValue = True
          End If
          
          If HaveValue Then
            If mcolParam("SUMMARY_FLAG") = "N" Then
               Vsp.FontBold = True
               Call m_Details1.GetString(1, TempStr1, TempStr2)
               strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
               Call Vsp.AddTable(strFormat, "", TempStr2)
               Vsp.FontBold = False
             End If
             HaveValueSum = True
          End If
          
          For j = 1 To UBound(Total3)
            Total3(j) = 0
         Next j
            
          ' รวม
      End If
   Next SupGroup
      
   If HaveValueSum Then
      Call GenerateFooterSupplierGroup(PrevKey2)
   End If
   
   Set BodyCf = Nothing

   Call GenerateFooter("สุทธิ", "", "", Total2, m_Details1)
   Call m_Details1.GetString(2, TempStr1, TempStr2)
   strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
   Vsp.FontBold = True
   Call Vsp.AddTable(strFormat, "", TempStr2)
   Vsp.FontBold = False
               
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
Private Sub GenerateFooterSupplierGroup(PrevKey2 As String)
Dim TempStr1 As String
Dim TempStr2 As String
Dim alngX() As Long
Dim strFormat As String
Dim j As Long
Dim TotalFooter(100) As Double
Dim K As Long

      FirstDate = DateSerial(mcolParam("YEAR_NO") - 543, mcolParam("MONTH_ID"), 1)
      Call GenerateFooter(Format(Month(FirstDate), "00") & "/" & Format(Year(FirstDate) + 543, "00"), PrevKey2, "", Total1, m_Details1)
      Vsp.FontBold = True
      Call m_Details1.GetString(2, TempStr1, TempStr2)
      strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
      Call Vsp.AddTable(strFormat, "", TempStr2)
      Vsp.FontBold = False
            
      K = 3
      TotalFooter(K) = Total1(K) 'ยอดหนี้ยกมาเซ็ตเป็นอันแรก
      K = K + 1
      TotalFooter(K) = TotalFooter(K) + Total1(K)
      K = K + 1
      TotalFooter(K) = TotalFooter(K) + Total1(K)
      K = K + 1
      'ยอดหนี้ยกไปเซ็ตเป็นอันสุดท้าย
      K = K + 1
      TotalFooter(K) = TotalFooter(K) + Total1(K)
      K = K + 1
      TotalFooter(K) = TotalFooter(K) + Total1(K)
      K = K + 1
      'ยอดหนี้คงเหลือเซ็ตเป็นอันสุดท้าย
      K = K + 1
      TotalFooter(K) = TotalFooter(K) + Total1(K)
      
      For j = 1 To UBound(Total1)
         Total1(j) = 0
      Next j
      
      If (RunMonth >= 2) Then
         ' ของเดือนแรกที่เลือก 2
         FirstDate = DateAdd("M", 1, FirstDate)
         Call GenerateFooter(Format(Month(FirstDate), "00") & "/" & Format(Year(FirstDate) + 543, "00"), "", "", Total12, m_Details1)
         Vsp.FontBold = True
         Call m_Details1.GetString(2, TempStr1, TempStr2)
         strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
         Call Vsp.AddTable(strFormat, "", TempStr2)
         Vsp.FontBold = False
         
         K = 3
         'TotalFooter(K) = Total1(K) 'ยอดหนี้ยกมาเซ็ตเป็นอันแรก
         K = K + 1
         TotalFooter(K) = TotalFooter(K) + Total12(K)
         K = K + 1
         TotalFooter(K) = TotalFooter(K) + Total12(K)
         K = K + 1
         TotalFooter(K) = Total12(K) 'ยอดหนี้ยกไปเซ็ตเป็นอันสุดท้าย
         K = K + 1
         TotalFooter(K) = TotalFooter(K) + Total12(K)
         K = K + 1
         TotalFooter(K) = TotalFooter(K) + Total12(K)
         K = K + 1
         TotalFooter(K) = Total12(K) 'ยอดหนี้คงเหลือเซ็ตเป็นอันสุดท้าย
         K = K + 1
         TotalFooter(K) = TotalFooter(K) + Total12(K)
         
         For j = 1 To UBound(Total12)
            Total12(j) = 0
         Next j
      End If
      
      If (RunMonth >= 3) Then
         ' ของเดือนแรกที่เลือก 3
         FirstDate = DateAdd("M", 1, FirstDate)
         Call GenerateFooter(Format(Month(FirstDate), "00") & "/" & Format(Year(FirstDate) + 543, "00"), "", "", Total13, m_Details1)
         Vsp.FontBold = True
         Call m_Details1.GetString(2, TempStr1, TempStr2)
         strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
         Call Vsp.AddTable(strFormat, "", TempStr2)
         Vsp.FontBold = False
         
         K = 3
         'TotalFooter(K) = Total1(K) 'ยอดหนี้ยกมาเซ็ตเป็นอันแรก
         K = K + 1
         TotalFooter(K) = TotalFooter(K) + Total13(K)
         K = K + 1
         TotalFooter(K) = TotalFooter(K) + Total13(K)
         K = K + 1
         TotalFooter(K) = Total13(K) 'ยอดหนี้ยกไปเซ็ตเป็นอันสุดท้าย
         K = K + 1
         TotalFooter(K) = TotalFooter(K) + Total13(K)
         K = K + 1
         TotalFooter(K) = TotalFooter(K) + Total13(K)
         K = K + 1
         TotalFooter(K) = Total13(K) 'ยอดหนี้คงเหลือเซ็ตเป็นอันสุดท้าย
         K = K + 1
         TotalFooter(K) = TotalFooter(K) + Total13(K)
         
         For j = 1 To UBound(Total13)
            Total13(j) = 0
         Next j
      End If
      
      If (RunMonth >= 4) Then
         ' ของเดือนแรกที่เลือก 4
         FirstDate = DateAdd("M", 1, FirstDate)
         Call GenerateFooter(Format(Month(FirstDate), "00") & "/" & Format(Year(FirstDate) + 543, "00"), "", "", Total14, m_Details1)
         Vsp.FontBold = True
         Call m_Details1.GetString(2, TempStr1, TempStr2)
         strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
         Call Vsp.AddTable(strFormat, "", TempStr2)
         Vsp.FontBold = False
         
         K = 3
         'TotalFooter(K) = Total1(K) 'ยอดหนี้ยกมาเซ็ตเป็นอันแรก
         K = K + 1
         TotalFooter(K) = TotalFooter(K) + Total14(K)
         K = K + 1
         TotalFooter(K) = TotalFooter(K) + Total14(K)
         K = K + 1
         TotalFooter(K) = Total14(K) 'ยอดหนี้ยกไปเซ็ตเป็นอันสุดท้าย
         K = K + 1
         TotalFooter(K) = TotalFooter(K) + Total14(K)
         K = K + 1
         TotalFooter(K) = TotalFooter(K) + Total14(K)
         K = K + 1
         TotalFooter(K) = Total14(K) 'ยอดหนี้คงเหลือเซ็ตเป็นอันสุดท้าย
         K = K + 1
         TotalFooter(K) = TotalFooter(K) + Total14(K)
         
         For j = 1 To UBound(Total14)
            Total14(j) = 0
         Next j
      End If
      
      If (RunMonth >= 5) Then
         ' ของเดือนแรกที่เลือก 5
         FirstDate = DateAdd("M", 1, FirstDate)
         Call GenerateFooter(Format(Month(FirstDate), "00") & "/" & Format(Year(FirstDate) + 543, "00"), "", "", Total15, m_Details1)
         Vsp.FontBold = True
         Call m_Details1.GetString(2, TempStr1, TempStr2)
         strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
         Call Vsp.AddTable(strFormat, "", TempStr2)
         Vsp.FontBold = False
         
         K = 3
         'TotalFooter(K) = Total1(K) 'ยอดหนี้ยกมาเซ็ตเป็นอันแรก
         K = K + 1
         TotalFooter(K) = TotalFooter(K) + Total15(K)
         K = K + 1
         TotalFooter(K) = TotalFooter(K) + Total15(K)
         K = K + 1
         TotalFooter(K) = Total15(K) 'ยอดหนี้ยกไปเซ็ตเป็นอันสุดท้าย
         K = K + 1
         TotalFooter(K) = TotalFooter(K) + Total15(K)
         K = K + 1
         TotalFooter(K) = TotalFooter(K) + Total15(K)
         K = K + 1
         TotalFooter(K) = Total15(K) 'ยอดหนี้คงเหลือเซ็ตเป็นอันสุดท้าย
         K = K + 1
         TotalFooter(K) = TotalFooter(K) + Total15(K)
         
         For j = 1 To UBound(Total15)
            Total15(j) = 0
         Next j
      End If
      
      If (RunMonth >= 6) Then
         ' ของเดือนแรกที่เลือก 6
         FirstDate = DateAdd("M", 1, FirstDate)
         Call GenerateFooter(Format(Month(FirstDate), "00") & "/" & Format(Year(FirstDate) + 543, "00"), "", "", Total16, m_Details1)
         Vsp.FontBold = True
         Call m_Details1.GetString(2, TempStr1, TempStr2)
         strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
         Call Vsp.AddTable(strFormat, "", TempStr2)
         Vsp.FontBold = False
         
         K = 3
         'TotalFooter(K) = Total1(K) 'ยอดหนี้ยกมาเซ็ตเป็นอันแรก
         K = K + 1
         TotalFooter(K) = TotalFooter(K) + Total16(K)
         K = K + 1
         TotalFooter(K) = TotalFooter(K) + Total16(K)
         K = K + 1
         TotalFooter(K) = Total16(K) 'ยอดหนี้ยกไปเซ็ตเป็นอันสุดท้าย
         K = K + 1
         TotalFooter(K) = TotalFooter(K) + Total16(K)
         K = K + 1
         TotalFooter(K) = TotalFooter(K) + Total16(K)
         K = K + 1
         TotalFooter(K) = Total16(K) 'ยอดหนี้คงเหลือเซ็ตเป็นอันสุดท้าย
         K = K + 1
         TotalFooter(K) = TotalFooter(K) + Total16(K)
         
         For j = 1 To UBound(Total16)
            Total16(j) = 0
         Next j
      End If
      
      If (RunMonth >= 7) Then
         ' ของเดือนแรกที่เลือก 7
         FirstDate = DateAdd("M", 1, FirstDate)
         Call GenerateFooter(Format(Month(FirstDate), "00") & "/" & Format(Year(FirstDate) + 543, "00"), "", "", Total17, m_Details1)
         Vsp.FontBold = True
         Call m_Details1.GetString(2, TempStr1, TempStr2)
         strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
         Call Vsp.AddTable(strFormat, "", TempStr2)
         Vsp.FontBold = False
         
         K = 3
         'TotalFooter(K) = Total1(K) 'ยอดหนี้ยกมาเซ็ตเป็นอันแรก
         K = K + 1
         TotalFooter(K) = TotalFooter(K) + Total17(K)
         K = K + 1
         TotalFooter(K) = TotalFooter(K) + Total17(K)
         K = K + 1
         TotalFooter(K) = Total17(K) 'ยอดหนี้ยกไปเซ็ตเป็นอันสุดท้าย
         K = K + 1
         TotalFooter(K) = TotalFooter(K) + Total17(K)
         K = K + 1
         TotalFooter(K) = TotalFooter(K) + Total17(K)
         K = K + 1
         TotalFooter(K) = Total17(K) 'ยอดหนี้คงเหลือเซ็ตเป็นอันสุดท้าย
         K = K + 1
         TotalFooter(K) = TotalFooter(K) + Total17(K)
         
         For j = 1 To UBound(Total17)
            Total17(j) = 0
         Next j
      End If
      
      If (RunMonth >= 8) Then
         ' ของเดือนแรกที่เลือก 8
         FirstDate = DateAdd("M", 1, FirstDate)
         Call GenerateFooter(Format(Month(FirstDate), "00") & "/" & Format(Year(FirstDate) + 543, "00"), "", "", Total18, m_Details1)
         Vsp.FontBold = True
         Call m_Details1.GetString(2, TempStr1, TempStr2)
         strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
         Call Vsp.AddTable(strFormat, "", TempStr2)
         Vsp.FontBold = False
         
         K = 3
         'TotalFooter(K) = Total1(K) 'ยอดหนี้ยกมาเซ็ตเป็นอันแรก
         K = K + 1
         TotalFooter(K) = TotalFooter(K) + Total18(K)
         K = K + 1
         TotalFooter(K) = TotalFooter(K) + Total18(K)
         K = K + 1
         TotalFooter(K) = Total18(K) 'ยอดหนี้ยกไปเซ็ตเป็นอันสุดท้าย
         K = K + 1
         TotalFooter(K) = TotalFooter(K) + Total18(K)
         K = K + 1
         TotalFooter(K) = TotalFooter(K) + Total18(K)
         K = K + 1
         TotalFooter(K) = Total18(K) 'ยอดหนี้คงเหลือเซ็ตเป็นอันสุดท้าย
         K = K + 1
         TotalFooter(K) = TotalFooter(K) + Total18(K)
         
         For j = 1 To UBound(Total18)
            Total18(j) = 0
         Next j
      End If
      
      If (RunMonth >= 9) Then
         ' ของเดือนแรกที่เลือก 9
         FirstDate = DateAdd("M", 1, FirstDate)
         Call GenerateFooter(Format(Month(FirstDate), "00") & "/" & Format(Year(FirstDate) + 543, "00"), "", "", Total19, m_Details1)
         Vsp.FontBold = True
         Call m_Details1.GetString(2, TempStr1, TempStr2)
         strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
         Call Vsp.AddTable(strFormat, "", TempStr2)
         Vsp.FontBold = False
         
         K = 3
         'TotalFooter(K) = Total1(K) 'ยอดหนี้ยกมาเซ็ตเป็นอันแรก
         K = K + 1
         TotalFooter(K) = TotalFooter(K) + Total19(K)
         K = K + 1
         TotalFooter(K) = TotalFooter(K) + Total19(K)
         K = K + 1
         TotalFooter(K) = Total19(K) 'ยอดหนี้ยกไปเซ็ตเป็นอันสุดท้าย
         K = K + 1
         TotalFooter(K) = TotalFooter(K) + Total19(K)
         K = K + 1
         TotalFooter(K) = TotalFooter(K) + Total19(K)
         K = K + 1
         TotalFooter(K) = Total19(K) 'ยอดหนี้คงเหลือเซ็ตเป็นอันสุดท้าย
         K = K + 1
         TotalFooter(K) = TotalFooter(K) + Total19(K)
         
         For j = 1 To UBound(Total19)
            Total19(j) = 0
         Next j
      End If
      
      If (RunMonth >= 10) Then
         ' ของเดือนแรกที่เลือก 10
         FirstDate = DateAdd("M", 1, FirstDate)
         Call GenerateFooter(Format(Month(FirstDate), "00") & "/" & Format(Year(FirstDate) + 543, "00"), "", "", Total110, m_Details1)
         Vsp.FontBold = True
         Call m_Details1.GetString(2, TempStr1, TempStr2)
         strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
         Call Vsp.AddTable(strFormat, "", TempStr2)
         Vsp.FontBold = False
         
         K = 3
         'TotalFooter(K) = Total1(K) 'ยอดหนี้ยกมาเซ็ตเป็นอันแรก
         K = K + 1
         TotalFooter(K) = TotalFooter(K) + Total110(K)
         K = K + 1
         TotalFooter(K) = TotalFooter(K) + Total110(K)
         K = K + 1
         TotalFooter(K) = Total110(K) 'ยอดหนี้ยกไปเซ็ตเป็นอันสุดท้าย
         K = K + 1
         TotalFooter(K) = TotalFooter(K) + Total110(K)
         K = K + 1
         TotalFooter(K) = TotalFooter(K) + Total110(K)
         K = K + 1
         TotalFooter(K) = Total110(K) 'ยอดหนี้คงเหลือเซ็ตเป็นอันสุดท้าย
         K = K + 1
         TotalFooter(K) = TotalFooter(K) + Total110(K)
         
         For j = 1 To UBound(Total110)
            Total110(j) = 0
         Next j
      End If
      
      If (RunMonth >= 11) Then
         ' ของเดือนแรกที่เลือก 11
         FirstDate = DateAdd("M", 1, FirstDate)
         Call GenerateFooter(Format(Month(FirstDate), "00") & "/" & Format(Year(FirstDate) + 543, "00"), "", "", Total111, m_Details1)
         Vsp.FontBold = True
         Call m_Details1.GetString(2, TempStr1, TempStr2)
         strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
         Call Vsp.AddTable(strFormat, "", TempStr2)
         Vsp.FontBold = False
         
         K = 3
         'TotalFooter(K) = Total1(K) 'ยอดหนี้ยกมาเซ็ตเป็นอันแรก
         K = K + 1
         TotalFooter(K) = TotalFooter(K) + Total111(K)
         K = K + 1
         TotalFooter(K) = TotalFooter(K) + Total111(K)
         K = K + 1
         TotalFooter(K) = Total111(K) 'ยอดหนี้ยกไปเซ็ตเป็นอันสุดท้าย
         K = K + 1
         TotalFooter(K) = TotalFooter(K) + Total111(K)
         K = K + 1
         TotalFooter(K) = TotalFooter(K) + Total111(K)
         K = K + 1
         TotalFooter(K) = Total111(K) 'ยอดหนี้คงเหลือเซ็ตเป็นอันสุดท้าย
         K = K + 1
         TotalFooter(K) = TotalFooter(K) + Total111(K)
         
         For j = 1 To UBound(Total111)
            Total111(j) = 0
         Next j
      End If
      
      If (RunMonth >= 12) Then
         ' ของเดือนแรกที่เลือก 12
         FirstDate = DateAdd("M", 1, FirstDate)
         Call GenerateFooter(Format(Month(FirstDate), "00") & "/" & Format(Year(FirstDate) + 543, "00"), "", "", Total112, m_Details1)
         Vsp.FontBold = True
         Call m_Details1.GetString(2, TempStr1, TempStr2)
         strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
         Call Vsp.AddTable(strFormat, "", TempStr2)
         Vsp.FontBold = False
         
         K = 3
         'TotalFooter(K) = Total1(K) 'ยอดหนี้ยกมาเซ็ตเป็นอันแรก
         K = K + 1
         TotalFooter(K) = TotalFooter(K) + Total112(K)
         K = K + 1
         TotalFooter(K) = TotalFooter(K) + Total112(K)
         K = K + 1
         TotalFooter(K) = Total112(K) 'ยอดหนี้ยกไปเซ็ตเป็นอันสุดท้าย
         K = K + 1
         TotalFooter(K) = TotalFooter(K) + Total112(K)
         K = K + 1
         TotalFooter(K) = TotalFooter(K) + Total112(K)
         K = K + 1
         TotalFooter(K) = Total112(K) 'ยอดหนี้คงเหลือเซ็ตเป็นอันสุดท้าย
         K = K + 1
         TotalFooter(K) = TotalFooter(K) + Total112(K)
         
         For j = 1 To UBound(Total112)
            Total112(j) = 0
         Next j
      End If
      
      Call GenerateFooter("รวม", PrevKey2, "", TotalFooter, m_Details1)
      Vsp.FontBold = True
      Call m_Details1.GetString(2, TempStr1, TempStr2)
      strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
      Call Vsp.AddTable(strFormat, "", TempStr2)
      Vsp.FontBold = False
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
         TempStr = Txt1
      ElseIf j = 2 Then
         TempStr = Txt2
      ElseIf j > 2 Then
         TempStr = FormatNumber(Tot(j), "")
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
'Private Sub SumValue(Coll As Collection, KEY As String, Value As Double)
'Dim GroupTypeItem As CColumnValue
'Dim TempColumnValue  As CColumnValue
'
'   Set GroupTypeItem = GetObject("CColumnValue", Coll, KEY, False)
'   If GroupTypeItem Is Nothing Then
'      Set TempColumnValue = New CColumnValue
'      TempColumnValue.KEY = KEY
'      TempColumnValue.Value = Value
'      Call Coll.Add(TempColumnValue, KEY)
'   Else
'      GroupTypeItem.Value = GroupTypeItem.Value + Value
'   End If
'End Sub
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
Private Sub AddCollection(Src As Collection, Des As Collection)
Dim Sg As CSupplierGroup
Dim Apm As CAPMas
   For Each Apm In Des
      Set Sg = GetObject("CSupplierGroup", Src, Trim(Apm.SUPCOD), False)
      If Sg Is Nothing Then
          Set Sg = New CSupplierGroup
          Sg.SUPPLIER_CODE = Apm.SUPCOD
          Sg.SUPPLIER_NAME = Apm.SUPNAM
          Call Src.Add(Sg, Trim(Sg.SUPPLIER_CODE))
      End If
   Next Apm
End Sub
