VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CReportAR013"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'User Group Report
Implements CReportInterface

Private Const MODULE_NAME = "CReportAR013"
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

 'Private m_CustomerTypes As Collection        'Private m_CustomerTypes As Collection
'1
Private m_ArCusAmounts As Collection      '  Private m_ArCusAmounts As Collection
Private m_ArCusAmountCns As Collection        '  Private m_ArCusAmountCns As Collection
Private m_CnAllAmounts As Collection       ' Private m_CnAllAmounts As Collection

Private m_ArCusAmountTo As Collection   ' Private m_ArCusAmountTo As Collection
Private m_ArCusAmountToCn As Collection   'Private m_ArCusAmountToCn As Collection

Private m_ArCusPaids As Collection   'Private m_ArCusPaids As Collection
Private m_ArCusPaidCns As Collection     'Private m_ArCusPaidCns As Collection
Private m_ArCusPaidTo As Collection       ' มาเอง ไม่รุอ่ะPrivate m_ArCusPaidTo As Collection
Private m_ArCusPaidToCn As Collection         'มาเองPrivate m_ArCusPaidToCn As Collection

Private m_PaidAmounts As Collection   'มาเองอ่ะ ไม่รุ
Private m_PaidAllAmounts As Collection
Private m_PaidAllAmountCns As Collection

Private m_5 As Collection
Private m_6 As Collection
'1
'2
Private m_ArCusAmounts2 As Collection
Private m_ArCusAmountCns2 As Collection
Private m_CnAllAmounts2 As Collection
Private m_ArCusAmountTo2 As Collection
Private m_ArCusAmountToCn2 As Collection
Private m_ArCusPaids2 As Collection
Private m_ArCusPaidCns2 As Collection
Private m_ArCusPaidTo2 As Collection
Private m_ArCusPaidToCn2 As Collection
Private m_PaidAllAmounts2 As Collection
Private m_PaidAllAmountCns2 As Collection
Private m_52 As Collection
Private m_62 As Collection
'2
'3
Private m_ArCusAmounts3 As Collection
Private m_ArCusAmountCns3 As Collection
Private m_CnAllAmounts3 As Collection
Private m_ArCusAmountTo3 As Collection
Private m_ArCusAmountToCn3 As Collection
Private m_ArCusPaids3 As Collection
Private m_ArCusPaidCns3 As Collection
Private m_ArCusPaidTo3 As Collection
Private m_ArCusPaidToCn3 As Collection
Private m_PaidAllAmounts3 As Collection
Private m_PaidAllAmountCns3 As Collection
Private m_53 As Collection
Private m_63 As Collection
'3
'4
Private m_ArCusAmounts4 As Collection
Private m_ArCusAmountCns4 As Collection
Private m_CnAllAmounts4 As Collection
Private m_ArCusAmountTo4 As Collection
Private m_ArCusAmountToCn4 As Collection
Private m_ArCusPaids4 As Collection
Private m_ArCusPaidCns4 As Collection
Private m_ArCusPaidTo4 As Collection
Private m_ArCusPaidToCn4 As Collection
Private m_PaidAllAmounts4 As Collection
Private m_PaidAllAmountCns4 As Collection
Private m_54 As Collection
Private m_64 As Collection
'4
'5
Private m_ArCusAmounts5 As Collection
Private m_ArCusAmountCns5 As Collection
Private m_CnAllAmounts5 As Collection
Private m_ArCusAmountTo5 As Collection
Private m_ArCusAmountToCn5 As Collection
Private m_ArCusPaids5 As Collection
Private m_ArCusPaidCns5 As Collection
Private m_ArCusPaidTo5 As Collection
Private m_ArCusPaidToCn5 As Collection
Private m_PaidAllAmounts5 As Collection
Private m_PaidAllAmountCns5 As Collection
Private m_55 As Collection
Private m_65 As Collection
'5
'6
Private m_ArCusAmounts6 As Collection
Private m_ArCusAmountCns6 As Collection
Private m_CnAllAmounts6 As Collection
Private m_ArCusAmountTo6 As Collection
Private m_ArCusAmountToCn6 As Collection
Private m_ArCusPaids6 As Collection
Private m_ArCusPaidCns6 As Collection
Private m_ArCusPaidTo6 As Collection
Private m_ArCusPaidToCn6 As Collection
Private m_PaidAllAmounts6 As Collection
Private m_PaidAllAmountCns6 As Collection
Private m_56 As Collection
Private m_66 As Collection
'6
'7
Private m_ArCusAmounts7 As Collection
Private m_ArCusAmountCns7 As Collection
Private m_CnAllAmounts7 As Collection
Private m_ArCusAmountTo7 As Collection
Private m_ArCusAmountToCn7 As Collection
Private m_ArCusPaids7 As Collection
Private m_ArCusPaidCns7 As Collection
Private m_ArCusPaidTo7 As Collection
Private m_ArCusPaidToCn7 As Collection
Private m_PaidAllAmounts7 As Collection
Private m_PaidAllAmountCns7 As Collection
Private m_57 As Collection
Private m_67 As Collection
'7
'8
Private m_ArCusAmounts8 As Collection
Private m_ArCusAmountCns8 As Collection
Private m_CnAllAmounts8 As Collection
Private m_ArCusAmountTo8 As Collection
Private m_ArCusAmountToCn8 As Collection
Private m_ArCusPaids8 As Collection
Private m_ArCusPaidCns8 As Collection
Private m_ArCusPaidTo8 As Collection
Private m_ArCusPaidToCn8 As Collection
Private m_PaidAllAmounts8 As Collection
Private m_PaidAllAmountCns8 As Collection
Private m_58 As Collection
Private m_68 As Collection
'8
'9
Private m_ArCusAmounts9 As Collection
Private m_ArCusAmountCns9 As Collection
Private m_CnAllAmounts9 As Collection
Private m_ArCusAmountTo9 As Collection
Private m_ArCusAmountToCn9 As Collection
Private m_ArCusPaids9 As Collection
Private m_ArCusPaidCns9 As Collection
Private m_ArCusPaidTo9 As Collection
Private m_ArCusPaidToCn9 As Collection
Private m_PaidAllAmounts9 As Collection
Private m_PaidAllAmountCns9 As Collection
Private m_59 As Collection
Private m_69 As Collection
'9
'10
Private m_ArCusAmounts10 As Collection
Private m_ArCusAmountCns10 As Collection
Private m_CnAllAmounts10 As Collection
Private m_ArCusAmountTo10 As Collection
Private m_ArCusAmountToCn10 As Collection
Private m_ArCusPaids10 As Collection
Private m_ArCusPaidCns10 As Collection
Private m_ArCusPaidTo10 As Collection
Private m_ArCusPaidToCn10 As Collection
Private m_PaidAllAmounts10 As Collection
Private m_PaidAllAmountCns10 As Collection
Private m_510 As Collection
Private m_610 As Collection
'10
'11
Private m_ArCusAmounts11 As Collection
Private m_ArCusAmountCns11 As Collection
Private m_CnAllAmounts11 As Collection
Private m_ArCusAmountTo11 As Collection
Private m_ArCusAmountToCn11 As Collection
Private m_ArCusPaids11 As Collection
Private m_ArCusPaidCns11 As Collection
Private m_ArCusPaidTo11 As Collection
Private m_ArCusPaidToCn11 As Collection
Private m_PaidAllAmounts11 As Collection
Private m_PaidAllAmountCns11 As Collection
Private m_511 As Collection
Private m_611 As Collection
'11
'12
Private m_ArCusAmounts12 As Collection
Private m_ArCusAmountCns12 As Collection
Private m_CnAllAmounts12 As Collection
Private m_ArCusPaids12 As Collection
Private m_ArCusPaidCns12 As Collection
Private m_PaidAllAmounts12 As Collection
Private m_PaidAllAmountCns12 As Collection
Private m_ArCusAmountTo12 As Collection
Private m_ArCusAmountToCn12 As Collection
Private m_ArCusPaidTo12 As Collection
Private m_ArCusPaidToCn12 As Collection
Private m_512 As Collection
Private m_612 As Collection
'12

Private Total1(100) As Double
Private Total2(100) As Double
Private Total3(100) As Double
Private Total12(100) As Double         ' มีแค่ 2ตัว
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

'Private CustomerGroupColl As Collection           'Private CustomerGroupColl As Collection
                             
Private m_SaleGroupCus As Collection         'Private Customer As Collection
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
 '  Set m_CustomerTypes = New Collection
   
   '1
   Set m_ArCusAmounts = New Collection
   Set m_ArCusAmountCns = New Collection
   Set m_CnAllAmounts = New Collection
   Set m_ArCusPaids = New Collection
   Set m_ArCusPaidCns = New Collection
   Set m_PaidAllAmounts = New Collection
   Set m_PaidAllAmountCns = New Collection
   Set m_ArCusAmountTo = New Collection
   Set m_ArCusAmountToCn = New Collection
   Set m_ArCusPaidTo = New Collection
   Set m_ArCusPaidToCn = New Collection
   Set m_5 = New Collection
   Set m_6 = New Collection
   '1
   
   '2
   Set m_ArCusAmounts2 = New Collection
   Set m_ArCusAmountCns2 = New Collection
   Set m_CnAllAmounts2 = New Collection
   Set m_ArCusPaids2 = New Collection
   Set m_ArCusPaidCns2 = New Collection
   Set m_PaidAllAmounts2 = New Collection
   Set m_PaidAllAmountCns2 = New Collection
   Set m_ArCusAmountTo2 = New Collection
   Set m_ArCusAmountToCn2 = New Collection
   Set m_ArCusPaidTo2 = New Collection
   Set m_ArCusPaidToCn2 = New Collection
   Set m_52 = New Collection
   Set m_62 = New Collection
   '2
   
   '3
   Set m_ArCusAmounts3 = New Collection
   Set m_ArCusAmountCns3 = New Collection
   Set m_CnAllAmounts3 = New Collection
   Set m_ArCusPaids3 = New Collection
   Set m_ArCusPaidCns3 = New Collection
   Set m_PaidAllAmounts3 = New Collection
   Set m_PaidAllAmountCns3 = New Collection
   Set m_ArCusAmountTo3 = New Collection
   Set m_ArCusAmountToCn3 = New Collection
   Set m_ArCusPaidTo3 = New Collection
   Set m_ArCusPaidToCn3 = New Collection
   Set m_53 = New Collection
   Set m_63 = New Collection
   '3
   
   '4
   Set m_ArCusAmounts4 = New Collection
   Set m_ArCusAmountCns4 = New Collection
   Set m_CnAllAmounts4 = New Collection
   Set m_ArCusPaids4 = New Collection
   Set m_ArCusPaidCns4 = New Collection
   Set m_PaidAllAmounts4 = New Collection
   Set m_PaidAllAmountCns4 = New Collection
   Set m_ArCusAmountTo4 = New Collection
   Set m_ArCusAmountToCn4 = New Collection
   Set m_ArCusPaidTo4 = New Collection
   Set m_ArCusPaidToCn4 = New Collection
   Set m_54 = New Collection
   Set m_64 = New Collection
   '4
   
   '5
   Set m_ArCusAmounts5 = New Collection
   Set m_ArCusAmountCns5 = New Collection
   Set m_CnAllAmounts5 = New Collection
   Set m_ArCusPaids5 = New Collection
   Set m_ArCusPaidCns5 = New Collection
   Set m_PaidAllAmounts5 = New Collection
   Set m_PaidAllAmountCns5 = New Collection
   Set m_ArCusAmountTo5 = New Collection
   Set m_ArCusAmountToCn5 = New Collection
   Set m_ArCusPaidTo5 = New Collection
   Set m_ArCusPaidToCn5 = New Collection
   Set m_55 = New Collection
   Set m_65 = New Collection
   '5
   
   '6
   Set m_ArCusAmounts6 = New Collection
   Set m_ArCusAmountCns6 = New Collection
   Set m_CnAllAmounts6 = New Collection
   Set m_ArCusPaids6 = New Collection
   Set m_ArCusPaidCns6 = New Collection
   Set m_PaidAllAmounts6 = New Collection
   Set m_PaidAllAmountCns6 = New Collection
   Set m_ArCusAmountTo6 = New Collection
   Set m_ArCusAmountToCn6 = New Collection
   Set m_ArCusPaidTo6 = New Collection
   Set m_ArCusPaidToCn6 = New Collection
   Set m_56 = New Collection
   Set m_66 = New Collection
   '6
   
   '7
   Set m_ArCusAmounts7 = New Collection
   Set m_ArCusAmountCns7 = New Collection
   Set m_CnAllAmounts7 = New Collection
   Set m_ArCusPaids7 = New Collection
   Set m_ArCusPaidCns7 = New Collection
   Set m_PaidAllAmounts7 = New Collection
   Set m_PaidAllAmountCns7 = New Collection
   Set m_ArCusAmountTo7 = New Collection
   Set m_ArCusAmountToCn7 = New Collection
   Set m_ArCusPaidTo7 = New Collection
   Set m_ArCusPaidToCn7 = New Collection
   Set m_57 = New Collection
   Set m_67 = New Collection
   '7
   
   '8
   Set m_ArCusAmounts8 = New Collection
   Set m_ArCusAmountCns8 = New Collection
   Set m_CnAllAmounts8 = New Collection
   Set m_ArCusPaids8 = New Collection
   Set m_ArCusPaidCns8 = New Collection
   Set m_PaidAllAmounts8 = New Collection
   Set m_PaidAllAmountCns8 = New Collection
   Set m_ArCusAmountTo8 = New Collection
   Set m_ArCusAmountToCn8 = New Collection
   Set m_ArCusPaidTo8 = New Collection
   Set m_ArCusPaidToCn8 = New Collection
   Set m_58 = New Collection
   Set m_68 = New Collection
   '8
   
   '9
   Set m_ArCusAmounts9 = New Collection
   Set m_ArCusAmountCns9 = New Collection
   Set m_CnAllAmounts9 = New Collection
   Set m_ArCusPaids9 = New Collection
   Set m_ArCusPaidCns9 = New Collection
   Set m_PaidAllAmounts9 = New Collection
   Set m_PaidAllAmountCns9 = New Collection
   Set m_ArCusAmountTo9 = New Collection
   Set m_ArCusAmountToCn9 = New Collection
   Set m_ArCusPaidTo9 = New Collection
   Set m_ArCusPaidToCn9 = New Collection
   Set m_59 = New Collection
   Set m_69 = New Collection
   '9
   
   '10
   Set m_ArCusAmounts10 = New Collection
   Set m_ArCusAmountCns10 = New Collection
   Set m_CnAllAmounts10 = New Collection
   Set m_ArCusPaids10 = New Collection
   Set m_ArCusPaidCns10 = New Collection
   Set m_PaidAllAmounts10 = New Collection
   Set m_PaidAllAmountCns10 = New Collection
   Set m_ArCusAmountTo10 = New Collection
   Set m_ArCusAmountToCn10 = New Collection
   Set m_ArCusPaidTo10 = New Collection
   Set m_ArCusPaidToCn10 = New Collection
   Set m_510 = New Collection
   Set m_610 = New Collection
   '10
   
   '11
   Set m_ArCusAmounts11 = New Collection
   Set m_ArCusAmountCns11 = New Collection
   Set m_CnAllAmounts11 = New Collection
   Set m_ArCusPaids11 = New Collection
   Set m_ArCusPaidCns11 = New Collection
   Set m_PaidAllAmounts11 = New Collection
   Set m_PaidAllAmountCns11 = New Collection
   Set m_ArCusAmountTo11 = New Collection
   Set m_ArCusAmountToCn11 = New Collection
   Set m_ArCusPaidTo11 = New Collection
   Set m_ArCusPaidToCn11 = New Collection
   Set m_511 = New Collection
   Set m_611 = New Collection
   '11
   
   '12
   Set m_ArCusAmounts12 = New Collection
   Set m_ArCusAmountCns12 = New Collection
   Set m_CnAllAmounts12 = New Collection
   Set m_ArCusPaids12 = New Collection
   Set m_ArCusPaidCns12 = New Collection
   Set m_PaidAllAmounts12 = New Collection
   Set m_PaidAllAmountCns12 = New Collection
   Set m_ArCusAmountTo12 = New Collection
   Set m_ArCusAmountToCn12 = New Collection
   Set m_ArCusPaidTo12 = New Collection
   Set m_ArCusPaidToCn12 = New Collection
   Set m_512 = New Collection
   Set m_612 = New Collection
   '12
   
   Set m_SaleGroupCus = New Collection
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
'   Set m_CustomerTypes = Nothing
   
   '1
   Set m_ArCusAmounts = Nothing
   Set m_ArCusAmountCns = Nothing
   Set m_CnAllAmounts = Nothing
   Set m_ArCusPaids = Nothing
   Set m_ArCusPaidCns = Nothing
   Set m_PaidAllAmounts = Nothing
   Set m_PaidAllAmountCns = Nothing
   Set m_ArCusAmountTo = Nothing
   Set m_ArCusAmountToCn = Nothing
   Set m_ArCusPaidTo = Nothing
   Set m_ArCusPaidToCn = Nothing
   Set m_5 = Nothing
   Set m_6 = Nothing
   '1
   
   '2
   Set m_ArCusAmounts2 = Nothing
   Set m_ArCusAmountCns2 = Nothing
   Set m_CnAllAmounts2 = Nothing
   Set m_ArCusPaids2 = Nothing
   Set m_ArCusPaidCns2 = Nothing
   Set m_PaidAllAmounts2 = Nothing
   Set m_PaidAllAmountCns2 = Nothing
   Set m_ArCusAmountTo2 = Nothing
   Set m_ArCusAmountToCn2 = Nothing
   Set m_ArCusPaidTo2 = Nothing
   Set m_ArCusPaidToCn2 = Nothing
   Set m_52 = Nothing
   Set m_62 = Nothing
   '2
   
   '3
   Set m_ArCusAmounts3 = Nothing
   Set m_ArCusAmountCns3 = Nothing
   Set m_CnAllAmounts3 = Nothing
   Set m_ArCusPaids3 = Nothing
   Set m_ArCusPaidCns3 = Nothing
   Set m_PaidAllAmounts3 = Nothing
   Set m_PaidAllAmountCns3 = Nothing
   Set m_ArCusAmountTo3 = Nothing
   Set m_ArCusAmountToCn3 = Nothing
   Set m_ArCusPaidTo3 = Nothing
   Set m_ArCusPaidToCn3 = Nothing
   Set m_53 = Nothing
   Set m_63 = Nothing
   '3
   
   '4
   Set m_ArCusAmounts4 = Nothing
   Set m_ArCusAmountCns4 = Nothing
   Set m_CnAllAmounts4 = Nothing
   Set m_ArCusPaids4 = Nothing
   Set m_ArCusPaidCns4 = Nothing
   Set m_PaidAllAmounts4 = Nothing
   Set m_PaidAllAmountCns4 = Nothing
   Set m_ArCusAmountTo4 = Nothing
   Set m_ArCusAmountToCn4 = Nothing
   Set m_ArCusPaidTo4 = Nothing
   Set m_ArCusPaidToCn4 = Nothing
   Set m_54 = Nothing
   Set m_64 = Nothing
   '4
   
   '5
   Set m_ArCusAmounts5 = Nothing
   Set m_ArCusAmountCns5 = Nothing
   Set m_CnAllAmounts5 = Nothing
   Set m_ArCusPaids5 = Nothing
   Set m_ArCusPaidCns5 = Nothing
   Set m_PaidAllAmounts5 = Nothing
   Set m_PaidAllAmountCns5 = Nothing
   Set m_ArCusAmountTo5 = Nothing
   Set m_ArCusAmountToCn5 = Nothing
   Set m_ArCusPaidTo5 = Nothing
   Set m_ArCusPaidToCn5 = Nothing
   Set m_55 = Nothing
   Set m_65 = Nothing
   '5
   
      '6
   Set m_ArCusAmounts6 = Nothing
   Set m_ArCusAmountCns6 = Nothing
   Set m_CnAllAmounts6 = Nothing
   Set m_ArCusPaids6 = Nothing
   Set m_ArCusPaidCns6 = Nothing
   Set m_PaidAllAmounts6 = Nothing
   Set m_PaidAllAmountCns6 = Nothing
   Set m_ArCusAmountTo6 = Nothing
   Set m_ArCusAmountToCn6 = Nothing
   Set m_ArCusPaidTo6 = Nothing
   Set m_ArCusPaidToCn6 = Nothing
   Set m_56 = Nothing
   Set m_66 = Nothing
   '6
   
      '7
   Set m_ArCusAmounts7 = Nothing
   Set m_ArCusAmountCns7 = Nothing
   Set m_CnAllAmounts7 = Nothing
   Set m_ArCusPaids7 = Nothing
   Set m_ArCusPaidCns7 = Nothing
   Set m_PaidAllAmounts7 = Nothing
   Set m_PaidAllAmountCns7 = Nothing
   Set m_ArCusAmountTo7 = Nothing
   Set m_ArCusAmountToCn7 = Nothing
   Set m_ArCusPaidTo7 = Nothing
   Set m_ArCusPaidToCn7 = Nothing
   Set m_57 = Nothing
   Set m_67 = Nothing
   '7
   
      '8
   Set m_ArCusAmounts8 = Nothing
   Set m_ArCusAmountCns8 = Nothing
   Set m_CnAllAmounts8 = Nothing
   Set m_ArCusPaids8 = Nothing
   Set m_ArCusPaidCns8 = Nothing
   Set m_PaidAllAmounts8 = Nothing
   Set m_PaidAllAmountCns8 = Nothing
   Set m_ArCusAmountTo8 = Nothing
   Set m_ArCusAmountToCn8 = Nothing
   Set m_ArCusPaidTo8 = Nothing
   Set m_ArCusPaidToCn8 = Nothing
   Set m_58 = Nothing
   Set m_68 = Nothing
   '8
   
      '9
   Set m_ArCusAmounts9 = Nothing
   Set m_ArCusAmountCns9 = Nothing
   Set m_CnAllAmounts9 = Nothing
   Set m_ArCusPaids9 = Nothing
   Set m_ArCusPaidCns9 = Nothing
   Set m_PaidAllAmounts9 = Nothing
   Set m_PaidAllAmountCns9 = Nothing
   Set m_ArCusAmountTo9 = Nothing
   Set m_ArCusAmountToCn9 = Nothing
   Set m_ArCusPaidTo9 = Nothing
   Set m_ArCusPaidToCn9 = Nothing
   Set m_59 = Nothing
   Set m_69 = Nothing
   '9
   
      '10
   Set m_ArCusAmounts10 = Nothing
   Set m_ArCusAmountCns10 = Nothing
   Set m_CnAllAmounts10 = Nothing
   Set m_ArCusPaids10 = Nothing
   Set m_ArCusPaidCns10 = Nothing
   Set m_PaidAllAmounts10 = Nothing
   Set m_PaidAllAmountCns10 = Nothing
   Set m_ArCusAmountTo10 = Nothing
   Set m_ArCusAmountToCn10 = Nothing
   Set m_ArCusPaidTo10 = Nothing
   Set m_ArCusPaidToCn10 = Nothing
   Set m_510 = Nothing
   Set m_610 = Nothing
   '10
   
      '11
   Set m_ArCusAmounts11 = Nothing
   Set m_ArCusAmountCns11 = Nothing
   Set m_CnAllAmounts11 = Nothing
   Set m_ArCusPaids11 = Nothing
   Set m_ArCusPaidCns11 = Nothing
   Set m_PaidAllAmounts11 = Nothing
   Set m_PaidAllAmountCns11 = Nothing
   Set m_ArCusAmountTo11 = Nothing
   Set m_ArCusAmountToCn11 = Nothing
   Set m_ArCusPaidTo11 = Nothing
   Set m_ArCusPaidToCn11 = Nothing
   Set m_511 = Nothing
   Set m_611 = Nothing
   '11
   
      '12
   Set m_ArCusAmounts12 = Nothing
   Set m_ArCusAmountCns12 = Nothing
   Set m_CnAllAmounts12 = Nothing
   Set m_ArCusPaids12 = Nothing
   Set m_ArCusPaidCns12 = Nothing
   Set m_PaidAllAmounts12 = Nothing
   Set m_PaidAllAmountCns12 = Nothing
   Set m_ArCusAmountTo12 = Nothing
   Set m_ArCusAmountToCn12 = Nothing
   Set m_ArCusPaidTo12 = Nothing
   Set m_ArCusPaidToCn12 = Nothing
   Set m_512 = Nothing
   Set m_612 = Nothing
   '12
   
 '  Set CustomerGroupColl = Nothing
   Set m_SaleGroupCus = Nothing
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
      
    Call Cf.SetFieldValue(2, "^", "N", "^")
   Call m_Headers0.AddField(Cf)
   Call Cf.SetFieldValue(5, "^", "ประเภทผู้ขาย", "<")
   Call m_Headers0.AddField(Cf)
   Call Cf.SetFieldValue(8, "^", "รายละเอียด", "<")
   Call m_Headers0.AddField(Cf)
   '===
   Call Cf.SetFieldValue(2, "^", "NO.", "^")
   Call m_Headers1.AddField(Cf)
   Call Cf.SetFieldValue(5, "^", "ชื่อลูกค้า", "<")
   Call m_Headers1.AddField(Cf)
   Call Cf.SetFieldValue(2, "^", "ยอดหนี้ยกมา", ">")
   Call m_Headers1.AddField(Cf)
   Call Cf.SetFieldValue(2, "^", "ขาย/เพิ่มหนี้", ">")
   Call m_Headers1.AddField(Cf)
   Call Cf.SetFieldValue(2, "^", "รับชำระ/ลดหนี้", ">")
   Call m_Headers1.AddField(Cf)
   Call Cf.SetFieldValue(2, "^", "ยอดหนี้ยกไป", ">")
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


   RunMonth = Val(mcolParam("RUN_MONTH"))
   Call GetFirstLastDate(DateSerial(mcolParam("YEAR_NO") - 543, mcolParam("MONTH_ID"), 1), FirstDate, LastDate)
   
   mstrErrMsg = ""
   mblnHeader = True
   mblnNewPage = True
   mblnEndPage = True
   Vsp.PaperSize = pprA4 '
   Vsp.ORIENTATION = orPortrait
   Vsp.MarginBottom = 700
   Vsp.MarginFooter = 700
   Vsp.MarginHeader = 700
   Vsp.MarginLeft = 700
   Vsp.MarginRight = 700
   Vsp.MarginTop = 700
   Vsp.FontName = "AngsanaUPC"
   Vsp.FontSize = DETAIL_SIZE
   
  '   Call LoadCustomerType(Nothing, m_CustomerTypes)

   'เดือน 1
   '-------------------------------- ยกมา --------------------------------
   Call LoadARAmountByCust2(Nothing, m_ArCusAmounts, -1, DateAdd("D", -1, FirstDate), "('3','4')")    'ยอดซื้อ ถึงวันที่
   Call LoadARAmountByCust2(Nothing, m_ArCusAmountCns, -1, DateAdd("D", -1, FirstDate), "('5')")     'ยอดซื้อ ถึงวันที่       ลดหนี้
  Call LoadARAmountByCust3(Nothing, m_CnAllAmounts, FirstDate, -1, -1, DateAdd("D", -1, FirstDate))
   Call LoadPaidAmountByCust2(Nothing, m_ArCusPaids, -1, DateAdd("D", -1, FirstDate), -1, -1, "('3','4')")           'ยอดจ่ายถึงวันที่
   Call LoadPaidAmountByCust2(Nothing, m_ArCusPaidCns, -1, DateAdd("D", -1, FirstDate), -1, -1, "('5')")           'ยอดจ่ายถึงวันที่    ลดหนี้
   Call LoadPaidAmountByCust2(Nothing, m_PaidAllAmounts, -1, -1, -1, DateAdd("D", -1, FirstDate), "('3','4')")         'ยอดตัดบิลทั้งหมดของบิลภายในวันที่กำหนด
   Call LoadPaidAmountByCust2(Nothing, m_PaidAllAmountCns, -1, -1, -1, DateAdd("D", -1, FirstDate), "('5')")          'ยอดตัดบิลทั้งหมดของบิลภายในวันที่กำหนด ลดหนี้
   '-------------------------------- ยกมา --------------------------------
   '-------------------------------- เคลื่อนไหว --------------------------------
   Call LoadARAmountByCust2(Nothing, m_ArCusAmountTo, FirstDate, LastDate, "('3','4')")     'ยอดขาย ถึงวันที่
   Call LoadARAmountByCust2(Nothing, m_ArCusAmountToCn, FirstDate, LastDate, "('5')")     'ลดหนี้ ถึงวันที่
   Call LoadPaidAmountByCust2(Nothing, m_ArCusPaidTo, FirstDate, LastDate, -1, -1, "('3','4')")           'ยอดซี้อถึงวันที่
   Call LoadPaidAmountByCust2(Nothing, m_ArCusPaidToCn, FirstDate, LastDate, -1, -1, "('5')")           'ถึงวันที่
   '-------------------------------- เคลื่อนไหว --------------------------------
   
   If (RunMonth >= 2) Then
   'เดือน 2
      FirstDate = DateAdd("M", 1, FirstDate)     ' เพิ่มเดือนอีก 1 เดือน
      Call GetFirstLastDate(FirstDate, FirstDate, LastDate)
      '-------------------------------- ยกมา --------------------------------
      Call LoadARAmountByCust2(Nothing, m_ArCusAmounts2, -1, DateAdd("D", -1, FirstDate), "('3','4')")    'ยอดซื้อ ถึงวันที่
      Call LoadARAmountByCust2(Nothing, m_ArCusAmountCns2, -1, DateAdd("D", -1, FirstDate), "('5')")     'ยอดซื้อ ถึงวันที่       ลดหนี้
      Call LoadARAmountByCust3(Nothing, m_CnAllAmounts2, FirstDate, -1, -1, DateAdd("D", -1, FirstDate))
      Call LoadPaidAmountByCust2(Nothing, m_ArCusPaids2, -1, DateAdd("D", -1, FirstDate), -1, -1, "('3','4')")           'ยอดจ่ายถึงวันที่
      Call LoadPaidAmountByCust2(Nothing, m_ArCusPaidCns2, -1, DateAdd("D", -1, FirstDate), -1, -1, "('5')")           'ยอดจ่ายถึงวันที่    ลดหนี้
      Call LoadPaidAmountByCust2(Nothing, m_PaidAllAmounts2, -1, -1, -1, DateAdd("D", -1, FirstDate), "('3','4')")         'ยอดตัดบิลทั้งหมดของบิลภายในวันที่กำหนด
      Call LoadPaidAmountByCust2(Nothing, m_PaidAllAmountCns2, -1, -1, -1, DateAdd("D", -1, FirstDate), "('5')")          'ยอดตัดบิลทั้งหมดของบิลภายในวันที่กำหนด ลดหนี้
      '-------------------------------- ยกมา --------------------------------
      '-------------------------------- เคลื่อนไหว --------------------------------
      Call LoadARAmountByCust2(Nothing, m_ArCusAmountTo2, FirstDate, LastDate, "('3','4')")     'ยอดซื้อ ถึงวันที่
      Call LoadARAmountByCust2(Nothing, m_ArCusAmountToCn2, FirstDate, LastDate, "('5')")     'ลดหนี้ ถึงวันที่
      Call LoadPaidAmountByCust2(Nothing, m_ArCusPaidTo2, FirstDate, LastDate, -1, -1, "('3','4')")           'ยอดจ่ายถึงวันที่
      Call LoadPaidAmountByCust2(Nothing, m_ArCusPaidToCn2, FirstDate, LastDate, -1, -1, "('5')")           'ยอดจ่ายถึงวันที่
      '-------------------------------- เคลื่อนไหว --------------------------------
   'เดือน 2
   End If
   
   If (RunMonth >= 3) Then
   'เดือน 3
      FirstDate = DateAdd("M", 1, FirstDate)
      Call GetFirstLastDate(FirstDate, FirstDate, LastDate)
      '-------------------------------- ยกมา --------------------------------
      Call LoadARAmountByCust2(Nothing, m_ArCusAmounts3, -1, DateAdd("D", -1, FirstDate), "('3','4')")    'ยอดซื้อ ถึงวันที่
      Call LoadARAmountByCust2(Nothing, m_ArCusAmountCns3, -1, DateAdd("D", -1, FirstDate), "('5')")     'ยอดซื้อ ถึงวันที่       ลดหนี้
      Call LoadARAmountByCust3(Nothing, m_CnAllAmounts3, FirstDate, -1, -1, DateAdd("D", -1, FirstDate))  '
      Call LoadPaidAmountByCust2(Nothing, m_ArCusPaids3, -1, DateAdd("D", -1, FirstDate), -1, -1, "('3','4')")           'ยอดจ่ายถึงวันที่
      Call LoadPaidAmountByCust2(Nothing, m_ArCusPaidCns3, -1, DateAdd("D", -1, FirstDate), -1, -1, "('5')")           'ยอดจ่ายถึงวันที่    ลดหนี้
      Call LoadPaidAmountByCust2(Nothing, m_PaidAllAmounts3, -1, -1, -1, DateAdd("D", -1, FirstDate), "('3','4')")         'ยอดตัดบิลทั้งหมดของบิลภายในวันที่กำหนด
      Call LoadPaidAmountByCust2(Nothing, m_PaidAllAmountCns3, -1, -1, -1, DateAdd("D", -1, FirstDate), "('5')")          'ยอดตัดบิลทั้งหมดของบิลภายในวันที่กำหนด ลดหนี้
      '-------------------------------- ยกมา --------------------------------
      '-------------------------------- เคลื่อนไหว --------------------------------
      Call LoadARAmountByCust2(Nothing, m_ArCusAmountTo3, FirstDate, LastDate, "('3','4')")     'ยอดซื้อ ถึงวันที่
      Call LoadARAmountByCust2(Nothing, m_ArCusAmountToCn3, FirstDate, LastDate, "('5')")     'ลดหนี้ ถึงวันที่
      Call LoadPaidAmountByCust2(Nothing, m_ArCusPaidTo3, FirstDate, LastDate, -1, -1, "('3','4')")           'ยอดจ่ายถึงวันที่
      Call LoadPaidAmountByCust2(Nothing, m_ArCusPaidToCn3, FirstDate, LastDate, -1, -1, "('5')")           'ยอดจ่ายถึงวันที่
      '-------------------------------- เคลื่อนไหว --------------------------------
   'เดือน 3
   End If
   
   If (RunMonth >= 4) Then
   'เดือน 4
      FirstDate = DateAdd("M", 1, FirstDate)
      Call GetFirstLastDate(FirstDate, FirstDate, LastDate)
      '-------------------------------- ยกมา --------------------------------
      Call LoadARAmountByCust2(Nothing, m_ArCusAmounts4, -1, DateAdd("D", -1, FirstDate), "('3','4')")    'ยอดซื้อ ถึงวันที่
      Call LoadARAmountByCust2(Nothing, m_ArCusAmountCns4, -1, DateAdd("D", -1, FirstDate), "('5')")     'ยอดซื้อ ถึงวันที่       ลดหนี้
      Call LoadARAmountByCust3(Nothing, m_CnAllAmounts4, FirstDate, -1, -1, DateAdd("D", -1, FirstDate))
      Call LoadPaidAmountByCust2(Nothing, m_ArCusPaids4, -1, DateAdd("D", -1, FirstDate), -1, -1, "('3','4')")           'ยอดจ่ายถึงวันที่
      Call LoadPaidAmountByCust2(Nothing, m_ArCusPaidCns4, -1, DateAdd("D", -1, FirstDate), -1, -1, "('5')")           'ยอดจ่ายถึงวันที่    ลดหนี้
      Call LoadPaidAmountByCust2(Nothing, m_PaidAllAmounts4, -1, -1, -1, DateAdd("D", -1, FirstDate), "('3','4')")         'ยอดตัดบิลทั้งหมดของบิลภายในวันที่กำหนด
      Call LoadPaidAmountByCust2(Nothing, m_PaidAllAmountCns4, -1, -1, -1, DateAdd("D", -1, FirstDate), "('5')")          'ยอดตัดบิลทั้งหมดของบิลภายในวันที่กำหนด ลดหนี้
      '-------------------------------- ยกมา --------------------------------
      '-------------------------------- เคลื่อนไหว --------------------------------
      Call LoadARAmountByCust2(Nothing, m_ArCusAmountTo4, FirstDate, LastDate, "('3','4')")     'ยอดซื้อ ถึงวันที่
      Call LoadARAmountByCust2(Nothing, m_ArCusAmountToCn4, FirstDate, LastDate, "('5')")     'ลดหนี้ ถึงวันที่
      Call LoadPaidAmountByCust2(Nothing, m_ArCusPaidTo4, FirstDate, LastDate, -1, -1, "('3','4')")           'ยอดจ่ายถึงวันที่
      Call LoadPaidAmountByCust2(Nothing, m_ArCusPaidToCn4, FirstDate, LastDate, -1, -1, "('5')")           'ยอดจ่ายถึงวันที่
      '-------------------------------- เคลื่อนไหว --------------------------------
   'เดือน 4
   End If
   
   If (RunMonth >= 5) Then
   'เดือน 5
      FirstDate = DateAdd("M", 1, FirstDate)
      Call GetFirstLastDate(FirstDate, FirstDate, LastDate)
      '-------------------------------- ยกมา --------------------------------
      Call LoadARAmountByCust2(Nothing, m_ArCusAmounts5, -1, DateAdd("D", -1, FirstDate), "('3','4')")    'ยอดซื้อ ถึงวันที่
      Call LoadARAmountByCust2(Nothing, m_ArCusAmountCns5, -1, DateAdd("D", -1, FirstDate), "('5')")     'ยอดซื้อ ถึงวันที่       ลดหนี้
      Call LoadARAmountByCust3(Nothing, m_CnAllAmounts5, FirstDate, -1, -1, DateAdd("D", -1, FirstDate))
      Call LoadPaidAmountByCust2(Nothing, m_ArCusPaids5, -1, DateAdd("D", -1, FirstDate), -1, -1, "('3','4')")           'ยอดจ่ายถึงวันที่
      Call LoadPaidAmountByCust2(Nothing, m_ArCusPaidCns5, -1, DateAdd("D", -1, FirstDate), -1, -1, "('5')")           'ยอดจ่ายถึงวันที่    ลดหนี้
      Call LoadPaidAmountByCust2(Nothing, m_PaidAllAmounts5, -1, -1, -1, DateAdd("D", -1, FirstDate), "('3','4')")         'ยอดตัดบิลทั้งหมดของบิลภายในวันที่กำหนด
      Call LoadPaidAmountByCust2(Nothing, m_PaidAllAmountCns5, -1, -1, -1, DateAdd("D", -1, FirstDate), "('5')")          'ยอดตัดบิลทั้งหมดของบิลภายในวันที่กำหนด ลดหนี้
      '-------------------------------- ยกมา --------------------------------
      '-------------------------------- เคลื่อนไหว --------------------------------
      Call LoadARAmountByCust2(Nothing, m_ArCusAmountTo5, FirstDate, LastDate, "('3','4')")     'ยอดซื้อ ถึงวันที่
      Call LoadARAmountByCust2(Nothing, m_ArCusAmountToCn5, FirstDate, LastDate, "('5')")     'ลดหนี้ ถึงวันที่
      Call LoadPaidAmountByCust2(Nothing, m_ArCusPaidTo5, FirstDate, LastDate, -1, -1, "('3','4')")           'ยอดจ่ายถึงวันที่
      Call LoadPaidAmountByCust2(Nothing, m_ArCusPaidToCn5, FirstDate, LastDate, -1, -1, "('5')")           'ยอดจ่ายถึงวันที่
      '-------------------------------- เคลื่อนไหว --------------------------------
   'เดือน 5
   End If
   
   
   If (RunMonth >= 6) Then
   'เดือน 6
      FirstDate = DateAdd("M", 1, FirstDate)
      Call GetFirstLastDate(FirstDate, FirstDate, LastDate)
      '-------------------------------- ยกมา --------------------------------
      Call LoadARAmountByCust2(Nothing, m_ArCusAmounts6, -1, DateAdd("D", -1, FirstDate), "('3','4')")    'ยอดซื้อ ถึงวันที่
      Call LoadARAmountByCust2(Nothing, m_ArCusAmountCns6, -1, DateAdd("D", -1, FirstDate), "('5')")     'ยอดซื้อ ถึงวันที่       ลดหนี้
      Call LoadARAmountByCust3(Nothing, m_CnAllAmounts6, FirstDate, -1, -1, DateAdd("D", -1, FirstDate))
      Call LoadPaidAmountByCust2(Nothing, m_ArCusPaids6, -1, DateAdd("D", -1, FirstDate), -1, -1, "('3','4')")           'ยอดจ่ายถึงวันที่
      Call LoadPaidAmountByCust2(Nothing, m_ArCusPaidCns6, -1, DateAdd("D", -1, FirstDate), -1, -1, "('5')")           'ยอดจ่ายถึงวันที่    ลดหนี้
      Call LoadPaidAmountByCust2(Nothing, m_PaidAllAmounts6, -1, -1, -1, DateAdd("D", -1, FirstDate), "('3','4')")         'ยอดตัดบิลทั้งหมดของบิลภายในวันที่กำหนด
      Call LoadPaidAmountByCust2(Nothing, m_PaidAllAmountCns6, -1, -1, -1, DateAdd("D", -1, FirstDate), "('5')")          'ยอดตัดบิลทั้งหมดของบิลภายในวันที่กำหนด ลดหนี้
      '-------------------------------- ยกมา --------------------------------
      '-------------------------------- เคลื่อนไหว --------------------------------
      Call LoadARAmountByCust2(Nothing, m_ArCusAmountTo6, FirstDate, LastDate, "('3','4')")     'ยอดซื้อ ถึงวันที่
      Call LoadARAmountByCust2(Nothing, m_ArCusAmountToCn6, FirstDate, LastDate, "('5')")     'ลดหนี้ ถึงวันที่
      Call LoadPaidAmountByCust2(Nothing, m_ArCusPaidTo6, FirstDate, LastDate, -1, -1, "('3','4')")           'ยอดจ่ายถึงวันที่
      Call LoadPaidAmountByCust2(Nothing, m_ArCusPaidToCn6, FirstDate, LastDate, -1, -1, "('5')")           'ยอดจ่ายถึงวันที่
      '-------------------------------- เคลื่อนไหว --------------------------------
   'เดือน 6
   End If
   
   If (RunMonth >= 7) Then
   'เดือน 7
      FirstDate = DateAdd("M", 1, FirstDate)
      Call GetFirstLastDate(FirstDate, FirstDate, LastDate)
      '-------------------------------- ยกมา --------------------------------
      Call LoadARAmountByCust2(Nothing, m_ArCusAmounts7, -1, DateAdd("D", -1, FirstDate), "('3','4')")    'ยอดซื้อ ถึงวันที่
      Call LoadARAmountByCust2(Nothing, m_ArCusAmountCns7, -1, DateAdd("D", -1, FirstDate), "('5')")     'ยอดซื้อ ถึงวันที่       ลดหนี้
     Call LoadARAmountByCust3(Nothing, m_CnAllAmounts7, FirstDate, -1, -1, DateAdd("D", -1, FirstDate))
      Call LoadPaidAmountByCust2(Nothing, m_ArCusPaids7, -1, DateAdd("D", -1, FirstDate), -1, -1, "('3','4')")           'ยอดจ่ายถึงวันที่
      Call LoadPaidAmountByCust2(Nothing, m_ArCusPaidCns7, -1, DateAdd("D", -1, FirstDate), -1, -1, "('5')")           'ยอดจ่ายถึงวันที่    ลดหนี้
      Call LoadPaidAmountByCust2(Nothing, m_PaidAllAmounts7, -1, -1, -1, DateAdd("D", -1, FirstDate), "('3','4')")         'ยอดตัดบิลทั้งหมดของบิลภายในวันที่กำหนด
      Call LoadPaidAmountByCust2(Nothing, m_PaidAllAmountCns7, -1, -1, -1, DateAdd("D", -1, FirstDate), "('5')")          'ยอดตัดบิลทั้งหมดของบิลภายในวันที่กำหนด ลดหนี้
      '-------------------------------- ยกมา --------------------------------
      '-------------------------------- เคลื่อนไหว --------------------------------
      Call LoadARAmountByCust2(Nothing, m_ArCusAmountTo7, FirstDate, LastDate, "('3','4')")     'ยอดซื้อ ถึงวันที่
      Call LoadARAmountByCust2(Nothing, m_ArCusAmountToCn7, FirstDate, LastDate, "('5')")     'ลดหนี้ ถึงวันที่
      Call LoadPaidAmountByCust2(Nothing, m_ArCusPaidTo7, FirstDate, LastDate, -1, -1, "('3','4')")           'ยอดจ่ายถึงวันที่
      Call LoadPaidAmountByCust2(Nothing, m_ArCusPaidToCn7, FirstDate, LastDate, -1, -1, "('5')")           'ยอดจ่ายถึงวันที่
      '-------------------------------- เคลื่อนไหว --------------------------------
   'เดือน 7
   End If
   
   If (RunMonth >= 8) Then
   'เดือน 8
      FirstDate = DateAdd("M", 1, FirstDate)
      Call GetFirstLastDate(FirstDate, FirstDate, LastDate)
      '-------------------------------- ยกมา --------------------------------
      Call LoadARAmountByCust2(Nothing, m_ArCusAmounts8, -1, DateAdd("D", -1, FirstDate), "('3','4')")    'ยอดซื้อ ถึงวันที่
      Call LoadARAmountByCust2(Nothing, m_ArCusAmountCns8, -1, DateAdd("D", -1, FirstDate), "('5')")     'ยอดซื้อ ถึงวันที่       ลดหนี้
      Call LoadARAmountByCust3(Nothing, m_CnAllAmounts8, FirstDate, -1, -1, DateAdd("D", -1, FirstDate))
      Call LoadPaidAmountByCust2(Nothing, m_ArCusPaids8, -1, DateAdd("D", -1, FirstDate), -1, -1, "('3','4')")           'ยอดจ่ายถึงวันที่
      Call LoadPaidAmountByCust2(Nothing, m_ArCusPaidCns8, -1, DateAdd("D", -1, FirstDate), -1, -1, "('5')")           'ยอดจ่ายถึงวันที่    ลดหนี้
      Call LoadPaidAmountByCust2(Nothing, m_PaidAllAmounts8, -1, -1, -1, DateAdd("D", -1, FirstDate), "('3','4')")         'ยอดตัดบิลทั้งหมดของบิลภายในวันที่กำหนด
      Call LoadPaidAmountByCust2(Nothing, m_PaidAllAmountCns8, -1, -1, -1, DateAdd("D", -1, FirstDate), "('5')")          'ยอดตัดบิลทั้งหมดของบิลภายในวันที่กำหนด ลดหนี้
      '-------------------------------- ยกมา --------------------------------
      '-------------------------------- เคลื่อนไหว --------------------------------
      Call LoadARAmountByCust2(Nothing, m_ArCusAmountTo8, FirstDate, LastDate, "('3','4')")     'ยอดซื้อ ถึงวันที่
      Call LoadARAmountByCust2(Nothing, m_ArCusAmountToCn8, FirstDate, LastDate, "('5')")     'ลดหนี้ ถึงวันที่
      Call LoadPaidAmountByCust2(Nothing, m_ArCusPaidTo8, FirstDate, LastDate, -1, -1, "('3','4')")           'ยอดจ่ายถึงวันที่
      Call LoadPaidAmountByCust2(Nothing, m_ArCusPaidToCn8, FirstDate, LastDate, -1, -1, "('5')")           'ยอดจ่ายถึงวันที่
      '-------------------------------- เคลื่อนไหว --------------------------------
   'เดือน 8
   End If
   
   If (RunMonth >= 9) Then
   'เดือน 9
      FirstDate = DateAdd("M", 1, FirstDate)
      Call GetFirstLastDate(FirstDate, FirstDate, LastDate)
      '-------------------------------- ยกมา --------------------------------
      Call LoadARAmountByCust2(Nothing, m_ArCusAmounts9, -1, DateAdd("D", -1, FirstDate), "('3','4')")    'ยอดซื้อ ถึงวันที่
      Call LoadARAmountByCust2(Nothing, m_ArCusAmountCns9, -1, DateAdd("D", -1, FirstDate), "('5')")     'ยอดซื้อ ถึงวันที่       ลดหนี้
     Call LoadARAmountByCust3(Nothing, m_CnAllAmounts9, FirstDate, -1, -1, DateAdd("D", -1, FirstDate))
      Call LoadPaidAmountByCust2(Nothing, m_ArCusPaids9, -1, DateAdd("D", -1, FirstDate), -1, -1, "('3','4')")           'ยอดจ่ายถึงวันที่
      Call LoadPaidAmountByCust2(Nothing, m_ArCusPaidCns9, -1, DateAdd("D", -1, FirstDate), -1, -1, "('5')")           'ยอดจ่ายถึงวันที่    ลดหนี้
      Call LoadPaidAmountByCust2(Nothing, m_PaidAllAmounts9, -1, -1, -1, DateAdd("D", -1, FirstDate), "('3','4')")         'ยอดตัดบิลทั้งหมดของบิลภายในวันที่กำหนด
      Call LoadPaidAmountByCust2(Nothing, m_PaidAllAmountCns9, -1, -1, -1, DateAdd("D", -1, FirstDate), "('5')")          'ยอดตัดบิลทั้งหมดของบิลภายในวันที่กำหนด ลดหนี้
      '-------------------------------- ยกมา --------------------------------
      '-------------------------------- เคลื่อนไหว --------------------------------
      Call LoadARAmountByCust2(Nothing, m_ArCusAmountTo9, FirstDate, LastDate, "('3','4')")     'ยอดซื้อ ถึงวันที่
      Call LoadARAmountByCust2(Nothing, m_ArCusAmountToCn9, FirstDate, LastDate, "('5')")     'ลดหนี้ ถึงวันที่
      Call LoadPaidAmountByCust2(Nothing, m_ArCusPaidTo9, FirstDate, LastDate, -1, -1, "('3','4')")           'ยอดจ่ายถึงวันที่
      Call LoadPaidAmountByCust2(Nothing, m_ArCusPaidToCn9, FirstDate, LastDate, -1, -1, "('5')")           'ยอดจ่ายถึงวันที่
      '-------------------------------- เคลื่อนไหว --------------------------------
   'เดือน 9
   End If
   
   If (RunMonth >= 10) Then
   'เดือน 10
      FirstDate = DateAdd("M", 1, FirstDate)
      Call GetFirstLastDate(FirstDate, FirstDate, LastDate)
      '-------------------------------- ยกมา --------------------------------
      Call LoadARAmountByCust2(Nothing, m_ArCusAmounts10, -1, DateAdd("D", -1, FirstDate), "('3','4')")    'ยอดซื้อ ถึงวันที่
      Call LoadARAmountByCust2(Nothing, m_ArCusAmountCns10, -1, DateAdd("D", -1, FirstDate), "('5')")     'ยอดซื้อ ถึงวันที่       ลดหนี้
      Call LoadARAmountByCust3(Nothing, m_CnAllAmounts10, FirstDate, -1, -1, DateAdd("D", -1, FirstDate))
      Call LoadPaidAmountByCust2(Nothing, m_ArCusPaids10, -1, DateAdd("D", -1, FirstDate), -1, -1, "('3','4')")           'ยอดจ่ายถึงวันที่
      Call LoadPaidAmountByCust2(Nothing, m_ArCusPaidCns10, -1, DateAdd("D", -1, FirstDate), -1, -1, "('5')")           'ยอดจ่ายถึงวันที่    ลดหนี้
      Call LoadPaidAmountByCust2(Nothing, m_PaidAllAmounts10, -1, -1, -1, DateAdd("D", -1, FirstDate), "('3','4')")         'ยอดตัดบิลทั้งหมดของบิลภายในวันที่กำหนด
      Call LoadPaidAmountByCust2(Nothing, m_PaidAllAmountCns10, -1, -1, -1, DateAdd("D", -1, FirstDate), "('5')")          'ยอดตัดบิลทั้งหมดของบิลภายในวันที่กำหนด ลดหนี้
      '-------------------------------- ยกมา --------------------------------
      '-------------------------------- เคลื่อนไหว --------------------------------
      Call LoadARAmountByCust2(Nothing, m_ArCusAmountTo10, FirstDate, LastDate, "('3','4')")     'ยอดซื้อ ถึงวันที่
      Call LoadARAmountByCust2(Nothing, m_ArCusAmountToCn10, FirstDate, LastDate, "('5')")     'ลดหนี้ ถึงวันที่
      Call LoadPaidAmountByCust2(Nothing, m_ArCusPaidTo10, FirstDate, LastDate, -1, -1, "('3','4')")           'ยอดจ่ายถึงวันที่
      Call LoadPaidAmountByCust2(Nothing, m_ArCusPaidToCn10, FirstDate, LastDate, -1, -1, "('5')")           'ยอดจ่ายถึงวันที่
      '-------------------------------- เคลื่อนไหว --------------------------------
   'เดือน 10
   End If
   
   If (RunMonth >= 11) Then
   'เดือน 11
      FirstDate = DateAdd("M", 1, FirstDate)
      Call GetFirstLastDate(FirstDate, FirstDate, LastDate)
      '-------------------------------- ยกมา --------------------------------
      Call LoadARAmountByCust2(Nothing, m_ArCusAmounts11, -1, DateAdd("D", -1, FirstDate), "('3','4')")    'ยอดซื้อ ถึงวันที่
      Call LoadARAmountByCust2(Nothing, m_ArCusAmountCns11, -1, DateAdd("D", -1, FirstDate), "('5')")     'ยอดซื้อ ถึงวันที่       ลดหนี้
      Call LoadARAmountByCust3(Nothing, m_CnAllAmounts11, FirstDate, -1, -1, DateAdd("D", -1, FirstDate))  'นที่
      Call LoadPaidAmountByCust2(Nothing, m_ArCusPaids11, -1, DateAdd("D", -1, FirstDate), -1, -1, "('3','4')")           'ยอดจ่ายถึงวันที่
      Call LoadPaidAmountByCust2(Nothing, m_ArCusPaidCns11, -1, DateAdd("D", -1, FirstDate), -1, -1, "('5')")           'ยอดจ่ายถึงวันที่    ลดหนี้
      Call LoadPaidAmountByCust2(Nothing, m_PaidAllAmounts11, -1, -1, -1, DateAdd("D", -1, FirstDate), "('3','4')")         'ยอดตัดบิลทั้งหมดของบิลภายในวันที่กำหนด
      Call LoadPaidAmountByCust2(Nothing, m_PaidAllAmountCns11, -1, -1, -1, DateAdd("D", -1, FirstDate), "('5')")          'ยอดตัดบิลทั้งหมดของบิลภายในวันที่กำหนด ลดหนี้
      '-------------------------------- ยกมา --------------------------------
      '-------------------------------- เคลื่อนไหว --------------------------------
      Call LoadARAmountByCust2(Nothing, m_ArCusAmountTo11, FirstDate, LastDate, "('3','4')")     'ยอดซื้อ ถึงวันที่
      Call LoadARAmountByCust2(Nothing, m_ArCusAmountToCn11, FirstDate, LastDate, "('5')")     'ลดหนี้ ถึงวันที่
      Call LoadPaidAmountByCust2(Nothing, m_ArCusPaidTo11, FirstDate, LastDate, -1, -1, "('3','4')")           'ยอดจ่ายถึงวันที่
      Call LoadPaidAmountByCust2(Nothing, m_ArCusPaidToCn11, FirstDate, LastDate, -1, -1, "('5')")           'ยอดจ่ายถึงวันที่
      '-------------------------------- เคลื่อนไหว --------------------------------
   'เดือน 11
   End If
   
   If (RunMonth >= 12) Then
   'เดือน 12
      FirstDate = DateAdd("M", 1, FirstDate)
      Call GetFirstLastDate(FirstDate, FirstDate, LastDate)
      '-------------------------------- ยกมา --------------------------------
      Call LoadARAmountByCust2(Nothing, m_ArCusAmounts12, -1, DateAdd("D", -1, FirstDate), "('3','4')")    'ยอดซื้อ ถึงวันที่
      Call LoadARAmountByCust2(Nothing, m_ArCusAmountCns12, -1, DateAdd("D", -1, FirstDate), "('5')")     'ยอดซื้อ ถึงวันที่       ลดหนี้
      Call LoadARAmountByCust3(Nothing, m_CnAllAmounts12, FirstDate, -1, -1, DateAdd("D", -1, FirstDate))
      Call LoadPaidAmountByCust2(Nothing, m_ArCusPaids12, -1, DateAdd("D", -1, FirstDate), -1, -1, "('3','4')")           'ยอดจ่ายถึงวันที่
      Call LoadPaidAmountByCust2(Nothing, m_ArCusPaidCns12, -1, DateAdd("D", -1, FirstDate), -1, -1, "('5')")           'ยอดจ่ายถึงวันที่    ลดหนี้
      Call LoadPaidAmountByCust2(Nothing, m_PaidAllAmounts12, -1, -1, -1, DateAdd("D", -1, FirstDate), "('3','4')")         'ยอดตัดบิลทั้งหมดของบิลภายในวันที่กำหนด
      Call LoadPaidAmountByCust2(Nothing, m_PaidAllAmountCns12, -1, -1, -1, DateAdd("D", -1, FirstDate), "('5')")          'ยอดตัดบิลทั้งหมดของบิลภายในวันที่กำหนด ลดหนี้
      '-------------------------------- ยกมา --------------------------------
      '-------------------------------- เคลื่อนไหว --------------------------------
      Call LoadARAmountByCust2(Nothing, m_ArCusAmountTo12, FirstDate, LastDate, "('3','4')")     'ยอดซื้อ ถึงวันที่
      Call LoadARAmountByCust2(Nothing, m_ArCusAmountToCn12, FirstDate, LastDate, "('5')")     'ลดหนี้ ถึงวันที่
      Call LoadPaidAmountByCust2(Nothing, m_ArCusPaidTo12, FirstDate, LastDate, -1, -1, "('3','4')")           'ยอดจ่ายถึงวันที่
      Call LoadPaidAmountByCust2(Nothing, m_ArCusPaidToCn12, FirstDate, LastDate, -1, -1, "('5')")           'ยอดจ่ายถึงวันที่
      '-------------------------------- เคลื่อนไหว --------------------------------
   'เดือน 12
   End If
   
   MODULE_DESC = glbCompanyName & vbCrLf & mcolParam("REPORT_NAME") & _
                                       vbCrLf & "งวด " & mcolParam("MONTH_ID") & "/" & mcolParam("YEAR_NO") & " เดือนเคลื่อนไหว " & mcolParam("RUN_MONTH") & " เดือน "
   
   If mcolParam("ONLY_CUSTOMER") = "Y" Then
               MODULE_DESC = MODULE_DESC & "  เฉพาะรวมตามลูกค้า"
   ElseIf mcolParam("SUMMARY_FLAG") = "Y" Then
              MODULE_DESC = MODULE_DESC & "  สรุป"
   End If
   
   
   
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
Dim Ar As CARMas ' Dim Ar As CArMas
Dim Width As Double
Dim ARt As CARTrn ' Dim Art As CArTrn
Dim TempRs As ADODB.Recordset

Dim TempARt As CARTrn     ' Dim TempARt As CArTrn
Dim TempARtCn As CARTrn       'Dim TempARtCn As CArTrn
Dim TempARtCash As CARTrn        'Dim TempARtCash As CArTrn

Dim TempCnAll As CARTrn   ' มาจากไหนไม่รุอ่ะ ??

Dim TempARRcIt As CARRcIt         'Dim TempARRcIt As CArRcIt
Dim TempARRcItCn As CARRcIt      'Dim TempARRcItCn As CArRcIt

Dim TempARtTo As CARTrn            'Dim TempARtTo As CArTrn
Dim TempARtToCn As CARTrn         'Dim TempARtToCn As CArTrn

Dim TempARRcItTo As CARRcIt        'Dim TempARRcItTo As CArRcIt
Dim TempARRcItToCn As CARRcIt      'Dim TempARRcItToCn As CArRcIt

Dim TempARBktTrn As CBkTrn               'Dim TempARBktTrn As CBkTrn

Dim TempARRcIt3 As CARRcIt                'Dim TempARRcIt3 As CArRcIt
Dim TempARRcIt3Cn As CARRcIt              'Dim TempARRcIt3Cn As CArRcIt

Dim BalanceAmount As Double

Dim TempPaidCheck As Double
Dim TempPaidAmount  As Double

Dim CusGroup As CCustomer        'CGroupType '       'CCustomer '
Dim GroupType As CCustomerType    'CCustomerType '
Dim GroupTypeItem As CColumnValue
Dim TempColumnValue As CColumnValue

Dim TempSum1 As Double
Dim TempSum2 As Double
Dim TempSum3 As Double
Dim TempSum4 As Double
Dim TempSum5 As Double
Dim TempSum6 As Double
Dim TempSum7 As Double
Dim TempSum8 As Double
Dim TempSum9 As Double

Dim TempValue As Double
Dim PrevKey3 As String
Dim PrevKey4 As String
Dim ArS As COESLM

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

   For j = 1 To UBound(Total1)
      Total1(j) = 0
      Total2(j) = 0
   Next j

   Call StartExportFile(Vsp)
   Vsp.StartDoc
   
      Set BodyCf = New CReportField
      i = 1

     Set ArS = New COESLM
      ArS.FROM_SALE_CODE = mcolParam("FROM_SALE_CODE")
      ArS.TO_SALE_CODE = mcolParam("TO_SALE_CODE")
      Call glbDaily.QuerySale(ArS, TempRs, iCount, IsOK, glbErrorLog)

   While Not TempRs.EOF          ' sale
   
                 Call m_Details1.ClearField
                 Call ArS.PopulateFromRS(1, TempRs)
                 
                 For j = 1 To UBound(Total3)
                    Total3(j) = 0
                 Next j
                   
                'Set BodyCf = New CReportField
                 Set HeadCf = m_Headers0.Fields(1)
                TempStr = "( " & ArS.SLMCOD & " )"                                          ' ????????????? sale
                Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
                Call m_Details1.AddField(BodyCf)
    
                Set HeadCf = m_Headers0.Fields(2)
                TempStr = ArS.SLMNAM
                Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
                Call m_Details1.AddField(BodyCf)
    
               Set HeadCf = m_Headers0.Fields(3)
                TempStr = ""
                Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
                Call m_Details1.AddField(BodyCf)
                
   If mcolParam("SUMMARY_FLAG") = "N" Then
                Call m_Details1.GetString(1, TempStr1, TempStr2)
                 strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
                    Vsp.FontBold = True
                    Call Vsp.AddTable(strFormat, "", TempStr2)
                    Vsp.FontBold = False
   End If
                PrevKey3 = ArS.SLMCOD
                PrevKey4 = ArS.SLMNAM
         
         '===


   
         Set Ar = New CARMas
         Ar.SLMCOD = ArS.SLMCOD
         Ar.FROM_CUSTOMER_CODE = mcolParam("FROM_CUSTOMER_CODE")
         Ar.TO_CUSTOMER_CODE = mcolParam("TO_CUSTOMER_CODE")
         Call glbDaily.QueryCustomer(Ar, Rs, iCount, IsOK, glbErrorLog)                   ' เลือกลูกค้ามาแต่ละคน
         
         HaveValueSum = False
         
   While Not Rs.EOF
            Call m_Details1.ClearField
            Call Ar.PopulateFromRS(1, Rs)
            
     ' ของเดือนแรกที่เลือก 1
            Call m_Details1.ClearField
            'debug.print Ar.CUSCOD
              Set TempARt = GetARTrn(m_ArCusAmounts, Ar.CUSCOD)      ' ar cod เป็น key  แต่ใน 1-16-1 หนึ่งคนแล้ววน 12 เดือน
            Set TempARtCn = GetARTrn(m_ArCusAmountCns, Ar.CUSCOD)
            
            Set TempCnAll = GetARTrn(m_CnAllAmounts, Ar.CUSCOD)
            
            Set TempARRcIt = GetARRcpItem(m_ArCusPaids, Ar.CUSCOD)
            Set TempARRcItCn = GetARRcpItem(m_ArCusPaidCns, Ar.CUSCOD)
               
            Set TempARRcIt3 = GetARRcpItemEx(m_PaidAllAmounts, Ar.CUSCOD)
            Set TempARRcIt3Cn = GetARRcpItemEx(m_PaidAllAmountCns, Ar.CUSCOD)
            
      HaveValue = False
      FirstDate = DateSerial(mcolParam("YEAR_NO") - 543, mcolParam("MONTH_ID"), 1)
      
            '===Head
            Set HeadCf = m_Headers1.Fields(1)
            TempStr = Format(Month(FirstDate), "00") & "/" & Format(Year(FirstDate) + 543, "00")
            Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
            Call m_Details1.AddField(BodyCf)
            
            Set HeadCf = m_Headers1.Fields(2)
            TempStr = Left(Ar.CUSNAM, 30) & "   (" & Ar.CUSCOD & ")"
            Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
            Call m_Details1.AddField(BodyCf)
            
            BalanceAmount = (TempARt.AMOUNT - TempARtCn.AMOUNT) - (TempARt.RCVAMT - TempARtCn.RCVAMT) - (TempARRcIt.RCVAMT - TempARRcItCn.RCVAMT) + (TempARRcIt3.RCVAMT - TempARRcIt3Cn.RCVAMT)
            BalanceAmount = BalanceAmount + TempCnAll.AMOUNT
                        
            K = 2
            K = K + 1    '3
            Set HeadCf = m_Headers1.Fields(K)
            TempStr = FormatNumber(BalanceAmount, "")
            Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
            Call m_Details1.AddField(BodyCf)
            
            Total1(K) = BalanceAmount             'T1 รวมของลูกค้า  T2 รวมสุทธิ  T3 รวม sale
            If Round((BalanceAmount), 2) <> 0 Then
               HaveValue = True
            End If
            
            Set TempARtTo = GetARTrn(m_ArCusAmountTo, Ar.CUSCOD)
            Set TempARtToCn = GetARTrn(m_ArCusAmountToCn, Ar.CUSCOD)
            
            Set TempARRcItTo = GetARRcpItem(m_ArCusPaidTo, Ar.CUSCOD)
            Set TempARRcItToCn = GetARRcpItem(m_ArCusPaidToCn, Ar.CUSCOD)
                     
            K = K + 1        '4
            Set HeadCf = m_Headers1.Fields(K)
            TempStr = FormatNumber(TempARtTo.AMOUNT, "")
            Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
            Call m_Details1.AddField(BodyCf)
            
            Total1(K) = Total1(K) + TempARtTo.AMOUNT
            Total3(K) = Total3(K) + TempARtTo.AMOUNT
            If Round((TempARtTo.AMOUNT), 2) <> 0 Then
               HaveValue = True
            End If
           
            
            K = K + 1    '5
            Set HeadCf = m_Headers1.Fields(K)
            TempStr = FormatNumber(TempARRcItTo.RCVAMT - TempARRcItToCn.RCVAMT + TempARtToCn.AMOUNT, "")
            Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
            Call m_Details1.AddField(BodyCf)
            
            Total1(K) = Total1(K) + TempARRcItTo.RCVAMT - TempARRcItToCn.RCVAMT + TempARtToCn.AMOUNT
            Total3(K) = Total3(K) + TempARRcItTo.RCVAMT - TempARRcItToCn.RCVAMT + TempARtToCn.AMOUNT
            If Round((TempARRcItTo.RCVAMT - TempARRcItToCn.RCVAMT + TempARtToCn.AMOUNT), 2) <> 0 Then
               HaveValue = True
            End If
            TempPaidCheck = Round((TempARRcItTo.RCVAMT - TempARRcItToCn.RCVAMT + TempARtToCn.AMOUNT), 2)
            
            K = K + 1    '6
            Set HeadCf = m_Headers1.Fields(K)
            TempStr = FormatNumber(BalanceAmount + (TempARtTo.AMOUNT - TempARtToCn.AMOUNT) - (TempARRcItTo.RCVAMT - TempARRcItToCn.RCVAMT), "")
            Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
            Call m_Details1.AddField(BodyCf)
            
            Total1(K) = Total1(K) + (BalanceAmount) + (TempARtTo.AMOUNT - TempARtToCn.AMOUNT) - (TempARRcItTo.RCVAMT - TempARRcItToCn.RCVAMT)
            If Round(((BalanceAmount) + (TempARtTo.AMOUNT - TempARtToCn.AMOUNT) - (TempARRcItTo.RCVAMT - TempARRcItToCn.RCVAMT)), 2) <> 0 Then
               HaveValue = True
            End If
            
       If HaveValue Then
         If mcolParam("SUMMARY_FLAG") = "N" And mcolParam("ONLY_CUSTOMER") = "N" Then
            Call m_Details1.GetString(1, TempStr1, TempStr2)
            strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
            Call Vsp.AddTable(strFormat, "", TempStr2)
          End If
          HaveValueSum = True
       End If
       ' ของเดือนแรกที่เลือก 1
       
      If (RunMonth >= 2) Then                        ' ///////////////////////////////////////////////////////////////////////////////////////////////////
         ' ของเดือนแรกที่เลือก 2
         
         FirstDate = DateAdd("M", 1, FirstDate)
         Call m_Details1.ClearField
               
            Set TempARt = GetARTrn(m_ArCusAmounts2, Ar.CUSCOD)      ' ar cod เป็น key  แต่ใน 1-16-1 หนึ่งคนแล้ววน 12 เดือน
            Set TempARtCn = GetARTrn(m_ArCusAmountCns2, Ar.CUSCOD)
            
            Set TempCnAll = GetARTrn(m_CnAllAmounts2, Ar.CUSCOD)
            
            Set TempARRcIt = GetARRcpItem(m_ArCusPaids2, Ar.CUSCOD)
            Set TempARRcItCn = GetARRcpItem(m_ArCusPaidCns2, Ar.CUSCOD)
               
            Set TempARRcIt3 = GetARRcpItemEx(m_PaidAllAmounts2, Ar.CUSCOD)
            Set TempARRcIt3Cn = GetARRcpItemEx(m_PaidAllAmountCns2, Ar.CUSCOD)
            
            HaveValue = False
            
            '===Head
            Set HeadCf = m_Headers1.Fields(1)
            TempStr = Format(Month(FirstDate), "00") & "/" & Format(Year(FirstDate) + 543, "00")
            Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
            Call m_Details1.AddField(BodyCf)
            
            Set HeadCf = m_Headers1.Fields(2)
            TempStr = Left(Ar.CUSNAM, 30) & "   (" & Ar.CUSCOD & ")"
            Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
            Call m_Details1.AddField(BodyCf)
            
            BalanceAmount = (TempARt.AMOUNT - TempARtCn.AMOUNT) - (TempARt.RCVAMT - TempARtCn.RCVAMT) - (TempARRcIt.RCVAMT - TempARRcItCn.RCVAMT) + (TempARRcIt3.RCVAMT - TempARRcIt3Cn.RCVAMT)
            BalanceAmount = BalanceAmount + TempCnAll.AMOUNT
                        
            K = 2
            K = K + 1    '3
            Set HeadCf = m_Headers1.Fields(K)
            TempStr = FormatNumber(BalanceAmount, "")
            Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
            Call m_Details1.AddField(BodyCf)
            
            If Round((BalanceAmount), 2) <> 0 Then
               HaveValue = True
            End If
            
            Set TempARtTo = GetARTrn(m_ArCusAmountTo2, Ar.CUSCOD)
            Set TempARtToCn = GetARTrn(m_ArCusAmountToCn2, Ar.CUSCOD)
            
            Set TempARRcItTo = GetARRcpItem(m_ArCusPaidTo2, Ar.CUSCOD)
            Set TempARRcItToCn = GetARRcpItem(m_ArCusPaidToCn2, Ar.CUSCOD)
                     
            K = K + 1        '4
            Set HeadCf = m_Headers1.Fields(K)
            TempStr = FormatNumber(TempARtTo.AMOUNT, "")
            Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
            Call m_Details1.AddField(BodyCf)
            
            Total1(K) = Total1(K) + TempARtTo.AMOUNT
            Total3(K) = Total3(K) + TempARtTo.AMOUNT
            If Round((TempARtTo.AMOUNT), 2) <> 0 Then
               HaveValue = True
            End If
           
            
            K = K + 1    '5
            Set HeadCf = m_Headers1.Fields(K)
            TempStr = FormatNumber(TempARRcItTo.RCVAMT - TempARRcItToCn.RCVAMT + TempARtToCn.AMOUNT, "")
            Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
            Call m_Details1.AddField(BodyCf)
            
            Total1(K) = Total1(K) + TempARRcItTo.RCVAMT - TempARRcItToCn.RCVAMT + TempARtToCn.AMOUNT
                     Total3(K) = Total3(K) + TempARRcItTo.RCVAMT - TempARRcItToCn.RCVAMT + TempARtToCn.AMOUNT
            If Round((TempARRcItTo.RCVAMT - TempARRcItToCn.RCVAMT + TempARtToCn.AMOUNT), 2) <> 0 Then
               HaveValue = True
            End If
            TempPaidCheck = Round((TempARRcItTo.RCVAMT - TempARRcItToCn.RCVAMT + TempARtToCn.AMOUNT), 2)
            
            K = K + 1
            Set HeadCf = m_Headers1.Fields(K)
            TempStr = FormatNumber(BalanceAmount + (TempARtTo.AMOUNT - TempARtToCn.AMOUNT) - (TempARRcItTo.RCVAMT - TempARRcItToCn.RCVAMT), "")
            Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
            Call m_Details1.AddField(BodyCf)
            
            Total1(K) = Total1(K) + (BalanceAmount) + (TempARtTo.AMOUNT - TempARtToCn.AMOUNT) - (TempARRcItTo.RCVAMT - TempARRcItToCn.RCVAMT)
            If Round(((BalanceAmount) + (TempARtTo.AMOUNT - TempARtToCn.AMOUNT) - (TempARRcItTo.RCVAMT - TempARRcItToCn.RCVAMT)), 2) <> 0 Then
               HaveValue = True
            End If
            
            If HaveValue Then
              If mcolParam("SUMMARY_FLAG") = "N" And mcolParam("ONLY_CUSTOMER") = "N" Then
                 Call m_Details1.GetString(1, TempStr1, TempStr2)
                 strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
                 Call Vsp.AddTable(strFormat, "", TempStr2)
               End If
               HaveValueSum = True
            End If
       
       End If       ' ของเดือนแรกที่เลือก 2
      
      
    If (RunMonth >= 3) Then                        ' ///////////////////////////////////////////////////////////////////////////////////////////////////
'         ' ของเดือนแรกที่เลือก 3

         FirstDate = DateAdd("M", 1, FirstDate)
         Call m_Details1.ClearField
               
            Set TempARt = GetARTrn(m_ArCusAmounts3, Ar.CUSCOD)      ' ar cod เป็น key  แต่ใน 1-16-1 หนึ่งคนแล้ววน 12 เดือน
            Set TempARtCn = GetARTrn(m_ArCusAmountCns3, Ar.CUSCOD)
            
            Set TempCnAll = GetARTrn(m_CnAllAmounts3, Ar.CUSCOD)
            
            Set TempARRcIt = GetARRcpItem(m_ArCusPaids3, Ar.CUSCOD)
            Set TempARRcItCn = GetARRcpItem(m_ArCusPaidCns3, Ar.CUSCOD)
               
            Set TempARRcIt3 = GetARRcpItemEx(m_PaidAllAmounts3, Ar.CUSCOD)
            Set TempARRcIt3Cn = GetARRcpItemEx(m_PaidAllAmountCns3, Ar.CUSCOD)
            
            HaveValue = False
            
            '===Head
            Set HeadCf = m_Headers1.Fields(1)
            TempStr = Format(Month(FirstDate), "00") & "/" & Format(Year(FirstDate) + 543, "00")
            Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
            Call m_Details1.AddField(BodyCf)
            
            Set HeadCf = m_Headers1.Fields(2)
            TempStr = Left(Ar.CUSNAM, 30) & "   (" & Ar.CUSCOD & ")"
            Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
            Call m_Details1.AddField(BodyCf)
            
            BalanceAmount = (TempARt.AMOUNT - TempARtCn.AMOUNT) - (TempARt.RCVAMT - TempARtCn.RCVAMT) - (TempARRcIt.RCVAMT - TempARRcItCn.RCVAMT) + (TempARRcIt3.RCVAMT - TempARRcIt3Cn.RCVAMT)
            BalanceAmount = BalanceAmount + TempCnAll.AMOUNT
                        
            K = 2
            K = K + 1    '3
            Set HeadCf = m_Headers1.Fields(K)
            TempStr = FormatNumber(BalanceAmount, "")
            Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
            Call m_Details1.AddField(BodyCf)
            
            If Round((BalanceAmount), 2) <> 0 Then
               HaveValue = True
            End If
            
            Set TempARtTo = GetARTrn(m_ArCusAmountTo3, Ar.CUSCOD)
            Set TempARtToCn = GetARTrn(m_ArCusAmountToCn3, Ar.CUSCOD)
            
            Set TempARRcItTo = GetARRcpItem(m_ArCusPaidTo3, Ar.CUSCOD)
            Set TempARRcItToCn = GetARRcpItem(m_ArCusPaidToCn3, Ar.CUSCOD)
                     
            K = K + 1        '4
            Set HeadCf = m_Headers1.Fields(K)
            TempStr = FormatNumber(TempARtTo.AMOUNT, "")
            Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
            Call m_Details1.AddField(BodyCf)
            
            Total1(K) = Total1(K) + TempARtTo.AMOUNT
            Total3(K) = Total3(K) + TempARtTo.AMOUNT
            If Round((TempARtTo.AMOUNT), 2) <> 0 Then
               HaveValue = True
            End If
           
            
            K = K + 1    '5
            Set HeadCf = m_Headers1.Fields(K)
            TempStr = FormatNumber(TempARRcItTo.RCVAMT - TempARRcItToCn.RCVAMT + TempARtToCn.AMOUNT, "")
            Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
            Call m_Details1.AddField(BodyCf)
            
            Total1(K) = Total1(K) + TempARRcItTo.RCVAMT - TempARRcItToCn.RCVAMT + TempARtToCn.AMOUNT
               Total3(K) = Total3(K) + TempARRcItTo.RCVAMT - TempARRcItToCn.RCVAMT + TempARtToCn.AMOUNT
            If Round((TempARRcItTo.RCVAMT - TempARRcItToCn.RCVAMT + TempARtToCn.AMOUNT), 2) <> 0 Then
               HaveValue = True
            End If
            TempPaidCheck = Round((TempARRcItTo.RCVAMT - TempARRcItToCn.RCVAMT + TempARtToCn.AMOUNT), 2)
            
            K = K + 1
            Set HeadCf = m_Headers1.Fields(K)
            TempStr = FormatNumber(BalanceAmount + (TempARtTo.AMOUNT - TempARtToCn.AMOUNT) - (TempARRcItTo.RCVAMT - TempARRcItToCn.RCVAMT), "")
            Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
            Call m_Details1.AddField(BodyCf)
            
            Total1(K) = Total1(K) + (BalanceAmount) + (TempARtTo.AMOUNT - TempARtToCn.AMOUNT) - (TempARRcItTo.RCVAMT - TempARRcItToCn.RCVAMT)
            If Round(((BalanceAmount) + (TempARtTo.AMOUNT - TempARtToCn.AMOUNT) - (TempARRcItTo.RCVAMT - TempARRcItToCn.RCVAMT)), 2) <> 0 Then
               HaveValue = True
            End If
            
            If HaveValue Then
              If mcolParam("SUMMARY_FLAG") = "N" And mcolParam("ONLY_CUSTOMER") = "N" Then
                 Call m_Details1.GetString(1, TempStr1, TempStr2)
                 strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
                 Call Vsp.AddTable(strFormat, "", TempStr2)
               End If
               HaveValueSum = True
            End If

'          ' ของเดือนแรกที่เลือก 3
    End If
    

     If (RunMonth >= 4) Then                        ' ///////////////////////////////////////////////////////////////////////////////////////////////////
'         ' ของเดือนแรกที่เลือก 4

         FirstDate = DateAdd("M", 1, FirstDate)
         Call m_Details1.ClearField
               
            Set TempARt = GetARTrn(m_ArCusAmounts4, Ar.CUSCOD)      ' ar cod เป็น key  แต่ใน 1-16-1 หนึ่งคนแล้ววน 12 เดือน
            Set TempARtCn = GetARTrn(m_ArCusAmountCns4, Ar.CUSCOD)
            
            Set TempCnAll = GetARTrn(m_CnAllAmounts4, Ar.CUSCOD)
            
            Set TempARRcIt = GetARRcpItem(m_ArCusPaids4, Ar.CUSCOD)
            Set TempARRcItCn = GetARRcpItem(m_ArCusPaidCns4, Ar.CUSCOD)
               
            Set TempARRcIt3 = GetARRcpItemEx(m_PaidAllAmounts4, Ar.CUSCOD)
            Set TempARRcIt3Cn = GetARRcpItemEx(m_PaidAllAmountCns4, Ar.CUSCOD)
            
            HaveValue = False
            
            '===Head
            Set HeadCf = m_Headers1.Fields(1)
            TempStr = Format(Month(FirstDate), "00") & "/" & Format(Year(FirstDate) + 543, "00")
            Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
            Call m_Details1.AddField(BodyCf)
            
            Set HeadCf = m_Headers1.Fields(2)
            TempStr = Left(Ar.CUSNAM, 30) & "   (" & Ar.CUSCOD & ")"
            Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
            Call m_Details1.AddField(BodyCf)
            
            BalanceAmount = (TempARt.AMOUNT - TempARtCn.AMOUNT) - (TempARt.RCVAMT - TempARtCn.RCVAMT) - (TempARRcIt.RCVAMT - TempARRcItCn.RCVAMT) + (TempARRcIt3.RCVAMT - TempARRcIt3Cn.RCVAMT)
            BalanceAmount = BalanceAmount + TempCnAll.AMOUNT
                        
            K = 2
            K = K + 1    '3
            Set HeadCf = m_Headers1.Fields(K)
            TempStr = FormatNumber(BalanceAmount, "")
            Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
            Call m_Details1.AddField(BodyCf)
            
            If Round((BalanceAmount), 2) <> 0 Then
               HaveValue = True
            End If
            
            Set TempARtTo = GetARTrn(m_ArCusAmountTo4, Ar.CUSCOD)
            Set TempARtToCn = GetARTrn(m_ArCusAmountToCn4, Ar.CUSCOD)
            
            Set TempARRcItTo = GetARRcpItem(m_ArCusPaidTo4, Ar.CUSCOD)
            Set TempARRcItToCn = GetARRcpItem(m_ArCusPaidToCn4, Ar.CUSCOD)
                     
            K = K + 1        '4
            Set HeadCf = m_Headers1.Fields(K)
            TempStr = FormatNumber(TempARtTo.AMOUNT, "")
            Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
            Call m_Details1.AddField(BodyCf)
            
            Total1(K) = Total1(K) + TempARtTo.AMOUNT
                     Total3(K) = Total3(K) + TempARtTo.AMOUNT
        
            If Round((TempARtTo.AMOUNT), 2) <> 0 Then
               HaveValue = True
            End If
           
            
            K = K + 1    '5
            Set HeadCf = m_Headers1.Fields(K)
            TempStr = FormatNumber(TempARRcItTo.RCVAMT - TempARRcItToCn.RCVAMT + TempARtToCn.AMOUNT, "")
            Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
            Call m_Details1.AddField(BodyCf)
            
            Total1(K) = Total1(K) + TempARRcItTo.RCVAMT - TempARRcItToCn.RCVAMT + TempARtToCn.AMOUNT
                Total3(K) = Total3(K) + TempARRcItTo.RCVAMT - TempARRcItToCn.RCVAMT + TempARtToCn.AMOUNT
            If Round((TempARRcItTo.RCVAMT - TempARRcItToCn.RCVAMT + TempARtToCn.AMOUNT), 2) <> 0 Then
               HaveValue = True
            End If
            TempPaidCheck = Round((TempARRcItTo.RCVAMT - TempARRcItToCn.RCVAMT + TempARtToCn.AMOUNT), 2)
            
            K = K + 1
            Set HeadCf = m_Headers1.Fields(K)
            TempStr = FormatNumber(BalanceAmount + (TempARtTo.AMOUNT - TempARtToCn.AMOUNT) - (TempARRcItTo.RCVAMT - TempARRcItToCn.RCVAMT), "")
            Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
            Call m_Details1.AddField(BodyCf)
            
            Total1(K) = Total1(K) + (BalanceAmount) + (TempARtTo.AMOUNT - TempARtToCn.AMOUNT) - (TempARRcItTo.RCVAMT - TempARRcItToCn.RCVAMT)
            If Round(((BalanceAmount) + (TempARtTo.AMOUNT - TempARtToCn.AMOUNT) - (TempARRcItTo.RCVAMT - TempARRcItToCn.RCVAMT)), 2) <> 0 Then
               HaveValue = True
            End If
            
            If HaveValue Then
              If mcolParam("SUMMARY_FLAG") = "N" And mcolParam("ONLY_CUSTOMER") = "N" Then
                 Call m_Details1.GetString(1, TempStr1, TempStr2)
                 strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
                 Call Vsp.AddTable(strFormat, "", TempStr2)
               End If
               HaveValueSum = True
            End If '
            
         ' ของเดือนแรกที่เลือก 4
      End If
'

      If (RunMonth >= 5) Then                        ' ///////////////////////////////////////////////////////////////////////////////////////////////////
'         ' ของเดือนแรกที่เลือก 5

         FirstDate = DateAdd("M", 1, FirstDate)
         Call m_Details1.ClearField
               
            Set TempARt = GetARTrn(m_ArCusAmounts5, Ar.CUSCOD)      ' ar cod เป็น key  แต่ใน 1-16-1 หนึ่งคนแล้ววน 12 เดือน
            Set TempARtCn = GetARTrn(m_ArCusAmountCns5, Ar.CUSCOD)
            
            Set TempCnAll = GetARTrn(m_CnAllAmounts5, Ar.CUSCOD)
            
            Set TempARRcIt = GetARRcpItem(m_ArCusPaids5, Ar.CUSCOD)
            Set TempARRcItCn = GetARRcpItem(m_ArCusPaidCns5, Ar.CUSCOD)
               
            Set TempARRcIt3 = GetARRcpItemEx(m_PaidAllAmounts5, Ar.CUSCOD)
            Set TempARRcIt3Cn = GetARRcpItemEx(m_PaidAllAmountCns5, Ar.CUSCOD)
            
            HaveValue = False
            
            '===Head
            Set HeadCf = m_Headers1.Fields(1)
            TempStr = Format(Month(FirstDate), "00") & "/" & Format(Year(FirstDate) + 543, "00")
            Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
            Call m_Details1.AddField(BodyCf)
            
            Set HeadCf = m_Headers1.Fields(2)
            TempStr = Left(Ar.CUSNAM, 30) & "   (" & Ar.CUSCOD & ")"
            Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
            Call m_Details1.AddField(BodyCf)
            
            BalanceAmount = (TempARt.AMOUNT - TempARtCn.AMOUNT) - (TempARt.RCVAMT - TempARtCn.RCVAMT) - (TempARRcIt.RCVAMT - TempARRcItCn.RCVAMT) + (TempARRcIt3.RCVAMT - TempARRcIt3Cn.RCVAMT)
            BalanceAmount = BalanceAmount + TempCnAll.AMOUNT
                        
            K = 2
            K = K + 1    '3
            Set HeadCf = m_Headers1.Fields(K)
            TempStr = FormatNumber(BalanceAmount, "")
            Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
            Call m_Details1.AddField(BodyCf)
            
            If Round((BalanceAmount), 2) <> 0 Then
               HaveValue = True
            End If
            
            Set TempARtTo = GetARTrn(m_ArCusAmountTo5, Ar.CUSCOD)
            Set TempARtToCn = GetARTrn(m_ArCusAmountToCn5, Ar.CUSCOD)
            
            Set TempARRcItTo = GetARRcpItem(m_ArCusPaidTo5, Ar.CUSCOD)
            Set TempARRcItToCn = GetARRcpItem(m_ArCusPaidToCn5, Ar.CUSCOD)
                     
            K = K + 1        '4
            Set HeadCf = m_Headers1.Fields(K)
            TempStr = FormatNumber(TempARtTo.AMOUNT, "")
            Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
            Call m_Details1.AddField(BodyCf)
            
            Total1(K) = Total1(K) + TempARtTo.AMOUNT
       '     Total2(K) = Total2(K) + TempARtTo.AMOUNT
            Total3(K) = Total3(K) + TempARtTo.AMOUNT
            If Round((TempARtTo.AMOUNT), 2) <> 0 Then
               HaveValue = True
            End If
           
            
            K = K + 1    '5
            Set HeadCf = m_Headers1.Fields(K)
            TempStr = FormatNumber(TempARRcItTo.RCVAMT - TempARRcItToCn.RCVAMT + TempARtToCn.AMOUNT, "")
            Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
            Call m_Details1.AddField(BodyCf)
            
            Total1(K) = Total1(K) + TempARRcItTo.RCVAMT - TempARRcItToCn.RCVAMT + TempARtToCn.AMOUNT
            Total3(K) = Total3(K) + TempARRcItTo.RCVAMT - TempARRcItToCn.RCVAMT + TempARtToCn.AMOUNT
            If Round((TempARRcItTo.RCVAMT - TempARRcItToCn.RCVAMT + TempARtToCn.AMOUNT), 2) <> 0 Then
               HaveValue = True
            End If
            TempPaidCheck = Round((TempARRcItTo.RCVAMT - TempARRcItToCn.RCVAMT + TempARtToCn.AMOUNT), 2)
            
            K = K + 1
            Set HeadCf = m_Headers1.Fields(K)
            TempStr = FormatNumber(BalanceAmount + (TempARtTo.AMOUNT - TempARtToCn.AMOUNT) - (TempARRcItTo.RCVAMT - TempARRcItToCn.RCVAMT), "")
            Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
            Call m_Details1.AddField(BodyCf)
            
            Total1(K) = Total1(K) + (BalanceAmount) + (TempARtTo.AMOUNT - TempARtToCn.AMOUNT) - (TempARRcItTo.RCVAMT - TempARRcItToCn.RCVAMT)
            If Round(((BalanceAmount) + (TempARtTo.AMOUNT - TempARtToCn.AMOUNT) - (TempARRcItTo.RCVAMT - TempARRcItToCn.RCVAMT)), 2) <> 0 Then
               HaveValue = True
            End If
            
            If HaveValue Then
              If mcolParam("SUMMARY_FLAG") = "N" And mcolParam("ONLY_CUSTOMER") = "N" Then
                 Call m_Details1.GetString(1, TempStr1, TempStr2)
                 strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
                 Call Vsp.AddTable(strFormat, "", TempStr2)
               End If
               HaveValueSum = True
            End If
            '          ' ของเดือนแรกที่เลือก 5
      End If
'
      If (RunMonth >= 6) Then                        ' ///////////////////////////////////////////////////////////////////////////////////////////////////
'         ' ของเดือนแรกที่เลือก 6
'
         FirstDate = DateAdd("M", 1, FirstDate)
         Call m_Details1.ClearField
               
            Set TempARt = GetARTrn(m_ArCusAmounts6, Ar.CUSCOD)      ' ar cod เป็น key  แต่ใน 1-16-1 หนึ่งคนแล้ววน 12 เดือน
            Set TempARtCn = GetARTrn(m_ArCusAmountCns6, Ar.CUSCOD)
            
            Set TempCnAll = GetARTrn(m_CnAllAmounts6, Ar.CUSCOD)
            
            Set TempARRcIt = GetARRcpItem(m_ArCusPaids6, Ar.CUSCOD)
            Set TempARRcItCn = GetARRcpItem(m_ArCusPaidCns6, Ar.CUSCOD)
               
            Set TempARRcIt3 = GetARRcpItemEx(m_PaidAllAmounts6, Ar.CUSCOD)
            Set TempARRcIt3Cn = GetARRcpItemEx(m_PaidAllAmountCns2, Ar.CUSCOD)
            
            HaveValue = False
            
            '===Head
            Set HeadCf = m_Headers1.Fields(1)
            TempStr = Format(Month(FirstDate), "00") & "/" & Format(Year(FirstDate) + 543, "00")
            Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
            Call m_Details1.AddField(BodyCf)
            
            Set HeadCf = m_Headers1.Fields(2)
            TempStr = Left(Ar.CUSNAM, 30) & "   (" & Ar.CUSCOD & ")"
            Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
            Call m_Details1.AddField(BodyCf)
            
            BalanceAmount = (TempARt.AMOUNT - TempARtCn.AMOUNT) - (TempARt.RCVAMT - TempARtCn.RCVAMT) - (TempARRcIt.RCVAMT - TempARRcItCn.RCVAMT) + (TempARRcIt3.RCVAMT - TempARRcIt3Cn.RCVAMT)
            BalanceAmount = BalanceAmount + TempCnAll.AMOUNT
                        
            K = 2
            K = K + 1    '3
            Set HeadCf = m_Headers1.Fields(K)
            TempStr = FormatNumber(BalanceAmount, "")
            Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
            Call m_Details1.AddField(BodyCf)
            
            If Round((BalanceAmount), 2) <> 0 Then
               HaveValue = True
            End If
            
            Set TempARtTo = GetARTrn(m_ArCusAmountTo6, Ar.CUSCOD)
            Set TempARtToCn = GetARTrn(m_ArCusAmountToCn6, Ar.CUSCOD)
            
            Set TempARRcItTo = GetARRcpItem(m_ArCusPaidTo6, Ar.CUSCOD)
            Set TempARRcItToCn = GetARRcpItem(m_ArCusPaidToCn6, Ar.CUSCOD)
                     
            K = K + 1        '4
            Set HeadCf = m_Headers1.Fields(K)
            TempStr = FormatNumber(TempARtTo.AMOUNT, "")
            Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
            Call m_Details1.AddField(BodyCf)
            
            Total1(K) = Total1(K) + TempARtTo.AMOUNT
              Total3(K) = Total3(K) + TempARtTo.AMOUNT
            If Round((TempARtTo.AMOUNT), 2) <> 0 Then
               HaveValue = True
            End If
           
            
            K = K + 1    '5
            Set HeadCf = m_Headers1.Fields(K)
            TempStr = FormatNumber(TempARRcItTo.RCVAMT - TempARRcItToCn.RCVAMT + TempARtToCn.AMOUNT, "")
            Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
            Call m_Details1.AddField(BodyCf)
            
            Total1(K) = Total1(K) + TempARRcItTo.RCVAMT - TempARRcItToCn.RCVAMT + TempARtToCn.AMOUNT
              Total3(K) = Total3(K) + TempARRcItTo.RCVAMT - TempARRcItToCn.RCVAMT + TempARtToCn.AMOUNT
            If Round((TempARRcItTo.RCVAMT - TempARRcItToCn.RCVAMT + TempARtToCn.AMOUNT), 2) <> 0 Then
               HaveValue = True
            End If
            TempPaidCheck = Round((TempARRcItTo.RCVAMT - TempARRcItToCn.RCVAMT + TempARtToCn.AMOUNT), 2)
            
            K = K + 1
            Set HeadCf = m_Headers1.Fields(K)
            TempStr = FormatNumber(BalanceAmount + (TempARtTo.AMOUNT - TempARtToCn.AMOUNT) - (TempARRcItTo.RCVAMT - TempARRcItToCn.RCVAMT), "")
            Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
            Call m_Details1.AddField(BodyCf)
            
            Total1(K) = Total1(K) + (BalanceAmount) + (TempARtTo.AMOUNT - TempARtToCn.AMOUNT) - (TempARRcItTo.RCVAMT - TempARRcItToCn.RCVAMT)
            If Round(((BalanceAmount) + (TempARtTo.AMOUNT - TempARtToCn.AMOUNT) - (TempARRcItTo.RCVAMT - TempARRcItToCn.RCVAMT)), 2) <> 0 Then
               HaveValue = True
            End If
            
            If HaveValue Then
              If mcolParam("SUMMARY_FLAG") = "N" And mcolParam("ONLY_CUSTOMER") = "N" Then
                 Call m_Details1.GetString(1, TempStr1, TempStr2)
                 strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
                 Call Vsp.AddTable(strFormat, "", TempStr2)
               End If
               HaveValueSum = True
            End If
'          ' ของเดือนแรกที่เลือก 6
      End If
      
      
      If (RunMonth >= 7) Then                        ' ///////////////////////////////////////////////////////////////////////////////////////////////////
'         ' ของเดือนแรกที่เลือก 7

         FirstDate = DateAdd("M", 1, FirstDate)
         Call m_Details1.ClearField
               
            Set TempARt = GetARTrn(m_ArCusAmounts7, Ar.CUSCOD)      ' ar cod เป็น key  แต่ใน 1-16-1 หนึ่งคนแล้ววน 12 เดือน
            Set TempARtCn = GetARTrn(m_ArCusAmountCns7, Ar.CUSCOD)
            
            Set TempCnAll = GetARTrn(m_CnAllAmounts7, Ar.CUSCOD)
            
            Set TempARRcIt = GetARRcpItem(m_ArCusPaids7, Ar.CUSCOD)
            Set TempARRcItCn = GetARRcpItem(m_ArCusPaidCns7, Ar.CUSCOD)
               
            Set TempARRcIt3 = GetARRcpItemEx(m_PaidAllAmounts7, Ar.CUSCOD)
            Set TempARRcIt3Cn = GetARRcpItemEx(m_PaidAllAmountCns7, Ar.CUSCOD)
            
            HaveValue = False
            
            '===Head
            Set HeadCf = m_Headers1.Fields(1)
            TempStr = Format(Month(FirstDate), "00") & "/" & Format(Year(FirstDate) + 543, "00")
            Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
            Call m_Details1.AddField(BodyCf)
            
            Set HeadCf = m_Headers1.Fields(2)
            TempStr = Left(Ar.CUSNAM, 30) & "   (" & Ar.CUSCOD & ")"
            Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
            Call m_Details1.AddField(BodyCf)
            
            BalanceAmount = (TempARt.AMOUNT - TempARtCn.AMOUNT) - (TempARt.RCVAMT - TempARtCn.RCVAMT) - (TempARRcIt.RCVAMT - TempARRcItCn.RCVAMT) + (TempARRcIt3.RCVAMT - TempARRcIt3Cn.RCVAMT)
            BalanceAmount = BalanceAmount + TempCnAll.AMOUNT
                        
            K = 2
            K = K + 1    '3
            Set HeadCf = m_Headers1.Fields(K)
            TempStr = FormatNumber(BalanceAmount, "")
            Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
            Call m_Details1.AddField(BodyCf)
            
            If Round((BalanceAmount), 2) <> 0 Then
               HaveValue = True
            End If
            
            Set TempARtTo = GetARTrn(m_ArCusAmountTo7, Ar.CUSCOD)
            Set TempARtToCn = GetARTrn(m_ArCusAmountToCn7, Ar.CUSCOD)
            
            Set TempARRcItTo = GetARRcpItem(m_ArCusPaidTo7, Ar.CUSCOD)
            Set TempARRcItToCn = GetARRcpItem(m_ArCusPaidToCn7, Ar.CUSCOD)
                     
            K = K + 1        '4
            Set HeadCf = m_Headers1.Fields(K)
            TempStr = FormatNumber(TempARtTo.AMOUNT, "")
            Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
            Call m_Details1.AddField(BodyCf)
            
            Total1(K) = Total1(K) + TempARtTo.AMOUNT
            Total3(K) = Total3(K) + TempARtTo.AMOUNT
            If Round((TempARtTo.AMOUNT), 2) <> 0 Then
               HaveValue = True
            End If
           
            
            K = K + 1    '5
            Set HeadCf = m_Headers1.Fields(K)
            TempStr = FormatNumber(TempARRcItTo.RCVAMT - TempARRcItToCn.RCVAMT + TempARtToCn.AMOUNT, "")
            Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
            Call m_Details1.AddField(BodyCf)
            
            Total1(K) = Total1(K) + TempARRcItTo.RCVAMT - TempARRcItToCn.RCVAMT + TempARtToCn.AMOUNT
       '     Total2(K) = Total2(K) + TempARRcItTo.RCVAMT - TempARRcItToCn.RCVAMT + TempARtToCn.AMOUNT
             Total3(K) = Total3(K) + TempARRcItTo.RCVAMT - TempARRcItToCn.RCVAMT + TempARtToCn.AMOUNT
            If Round((TempARRcItTo.RCVAMT - TempARRcItToCn.RCVAMT + TempARtToCn.AMOUNT), 2) <> 0 Then
               HaveValue = True
            End If
            TempPaidCheck = Round((TempARRcItTo.RCVAMT - TempARRcItToCn.RCVAMT + TempARtToCn.AMOUNT), 2)
            
            K = K + 1
            Set HeadCf = m_Headers1.Fields(K)
            TempStr = FormatNumber(BalanceAmount + (TempARtTo.AMOUNT - TempARtToCn.AMOUNT) - (TempARRcItTo.RCVAMT - TempARRcItToCn.RCVAMT), "")
            Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
            Call m_Details1.AddField(BodyCf)
            
            Total1(K) = Total1(K) + (BalanceAmount) + (TempARtTo.AMOUNT - TempARtToCn.AMOUNT) - (TempARRcItTo.RCVAMT - TempARRcItToCn.RCVAMT)
            If Round(((BalanceAmount) + (TempARtTo.AMOUNT - TempARtToCn.AMOUNT) - (TempARRcItTo.RCVAMT - TempARRcItToCn.RCVAMT)), 2) <> 0 Then
               HaveValue = True
            End If
            
            If HaveValue Then
              If mcolParam("SUMMARY_FLAG") = "N" And mcolParam("ONLY_CUSTOMER") = "N" Then
                 Call m_Details1.GetString(1, TempStr1, TempStr2)
                 strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
                 Call Vsp.AddTable(strFormat, "", TempStr2)
               End If
               HaveValueSum = True
            End If
            ' ของเดือนแรกที่เลือก 7
      End If
      
      
      If (RunMonth >= 8) Then                        ' ///////////////////////////////////////////////////////////////////////////////////////////////////
'         ' ของเดือนแรกที่เลือก 8

         FirstDate = DateAdd("M", 1, FirstDate)
         Call m_Details1.ClearField
               
            Set TempARt = GetARTrn(m_ArCusAmounts8, Ar.CUSCOD)      ' ar cod เป็น key  แต่ใน 1-16-1 หนึ่งคนแล้ววน 12 เดือน
            Set TempARtCn = GetARTrn(m_ArCusAmountCns8, Ar.CUSCOD)
            
            Set TempCnAll = GetARTrn(m_CnAllAmounts8, Ar.CUSCOD)
            
            Set TempARRcIt = GetARRcpItem(m_ArCusPaids8, Ar.CUSCOD)
            Set TempARRcItCn = GetARRcpItem(m_ArCusPaidCns8, Ar.CUSCOD)
               
            Set TempARRcIt3 = GetARRcpItemEx(m_PaidAllAmounts8, Ar.CUSCOD)
            Set TempARRcIt3Cn = GetARRcpItemEx(m_PaidAllAmountCns8, Ar.CUSCOD)
            
            HaveValue = False
            
            '===Head
            Set HeadCf = m_Headers1.Fields(1)
            TempStr = Format(Month(FirstDate), "00") & "/" & Format(Year(FirstDate) + 543, "00")
            Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
            Call m_Details1.AddField(BodyCf)
            
            Set HeadCf = m_Headers1.Fields(2)
            TempStr = Left(Ar.CUSNAM, 30) & "   (" & Ar.CUSCOD & ")"
            Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
            Call m_Details1.AddField(BodyCf)
            
            BalanceAmount = (TempARt.AMOUNT - TempARtCn.AMOUNT) - (TempARt.RCVAMT - TempARtCn.RCVAMT) - (TempARRcIt.RCVAMT - TempARRcItCn.RCVAMT) + (TempARRcIt3.RCVAMT - TempARRcIt3Cn.RCVAMT)
            BalanceAmount = BalanceAmount + TempCnAll.AMOUNT
                        
            K = 2
            K = K + 1    '3
            Set HeadCf = m_Headers1.Fields(K)
            TempStr = FormatNumber(BalanceAmount, "")
            Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
            Call m_Details1.AddField(BodyCf)
            
            If Round((BalanceAmount), 2) <> 0 Then
               HaveValue = True
            End If
            
            Set TempARtTo = GetARTrn(m_ArCusAmountTo8, Ar.CUSCOD)
            Set TempARtToCn = GetARTrn(m_ArCusAmountToCn8, Ar.CUSCOD)
            
            Set TempARRcItTo = GetARRcpItem(m_ArCusPaidTo8, Ar.CUSCOD)
            Set TempARRcItToCn = GetARRcpItem(m_ArCusPaidToCn8, Ar.CUSCOD)
                     
            K = K + 1        '4
            Set HeadCf = m_Headers1.Fields(K)
            TempStr = FormatNumber(TempARtTo.AMOUNT, "")
            Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
            Call m_Details1.AddField(BodyCf)
            
            Total1(K) = Total1(K) + TempARtTo.AMOUNT
        '    Total2(K) = Total2(K) + TempARtTo.AMOUNT
            Total3(K) = Total3(K) + TempARtTo.AMOUNT
            If Round((TempARtTo.AMOUNT), 2) <> 0 Then
               HaveValue = True
            End If
           
            
            K = K + 1    '5
            Set HeadCf = m_Headers1.Fields(K)
            TempStr = FormatNumber(TempARRcItTo.RCVAMT - TempARRcItToCn.RCVAMT + TempARtToCn.AMOUNT, "")
            Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
            Call m_Details1.AddField(BodyCf)
            
            Total1(K) = Total1(K) + TempARRcItTo.RCVAMT - TempARRcItToCn.RCVAMT + TempARtToCn.AMOUNT
            Total3(K) = Total3(K) + TempARRcItTo.RCVAMT - TempARRcItToCn.RCVAMT + TempARtToCn.AMOUNT
            If Round((TempARRcItTo.RCVAMT - TempARRcItToCn.RCVAMT + TempARtToCn.AMOUNT), 2) <> 0 Then
               HaveValue = True
            End If
            TempPaidCheck = Round((TempARRcItTo.RCVAMT - TempARRcItToCn.RCVAMT + TempARtToCn.AMOUNT), 2)
            
            K = K + 1
            Set HeadCf = m_Headers1.Fields(K)
            TempStr = FormatNumber(BalanceAmount + (TempARtTo.AMOUNT - TempARtToCn.AMOUNT) - (TempARRcItTo.RCVAMT - TempARRcItToCn.RCVAMT), "")
            Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
            Call m_Details1.AddField(BodyCf)
            
            Total1(K) = Total1(K) + (BalanceAmount) + (TempARtTo.AMOUNT - TempARtToCn.AMOUNT) - (TempARRcItTo.RCVAMT - TempARRcItToCn.RCVAMT)
            If Round(((BalanceAmount) + (TempARtTo.AMOUNT - TempARtToCn.AMOUNT) - (TempARRcItTo.RCVAMT - TempARRcItToCn.RCVAMT)), 2) <> 0 Then
               HaveValue = True
            End If
            
            If HaveValue Then
              If mcolParam("SUMMARY_FLAG") = "N" And mcolParam("ONLY_CUSTOMER") = "N" Then
                 Call m_Details1.GetString(1, TempStr1, TempStr2)
                 strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
                 Call Vsp.AddTable(strFormat, "", TempStr2)
               End If
               HaveValueSum = True
            End If
            
            '          ' ของเดือนแรกที่เลือก 8
      End If
'

      If (RunMonth >= 9) Then                        ' ///////////////////////////////////////////////////////////////////////////////////////////////////
'         ' ของเดือนแรกที่เลือก 9

         FirstDate = DateAdd("M", 1, FirstDate)
         Call m_Details1.ClearField
               
            Set TempARt = GetARTrn(m_ArCusAmounts9, Ar.CUSCOD)      ' ar cod เป็น key  แต่ใน 1-16-1 หนึ่งคนแล้ววน 12 เดือน
            Set TempARtCn = GetARTrn(m_ArCusAmountCns9, Ar.CUSCOD)
            
            Set TempCnAll = GetARTrn(m_CnAllAmounts9, Ar.CUSCOD)
            
            Set TempARRcIt = GetARRcpItem(m_ArCusPaids9, Ar.CUSCOD)
            Set TempARRcItCn = GetARRcpItem(m_ArCusPaidCns9, Ar.CUSCOD)
               
            Set TempARRcIt3 = GetARRcpItemEx(m_PaidAllAmounts9, Ar.CUSCOD)
            Set TempARRcIt3Cn = GetARRcpItemEx(m_PaidAllAmountCns9, Ar.CUSCOD)
            
            HaveValue = False
            
            '===Head
            Set HeadCf = m_Headers1.Fields(1)
            TempStr = Format(Month(FirstDate), "00") & "/" & Format(Year(FirstDate) + 543, "00")
            Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
            Call m_Details1.AddField(BodyCf)
            
            Set HeadCf = m_Headers1.Fields(2)
            TempStr = Left(Ar.CUSNAM, 30) & "   (" & Ar.CUSCOD & ")"
            Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
            Call m_Details1.AddField(BodyCf)
            
            BalanceAmount = (TempARt.AMOUNT - TempARtCn.AMOUNT) - (TempARt.RCVAMT - TempARtCn.RCVAMT) - (TempARRcIt.RCVAMT - TempARRcItCn.RCVAMT) + (TempARRcIt3.RCVAMT - TempARRcIt3Cn.RCVAMT)
            BalanceAmount = BalanceAmount + TempCnAll.AMOUNT
                        
            K = 2
            K = K + 1    '3
            Set HeadCf = m_Headers1.Fields(K)
            TempStr = FormatNumber(BalanceAmount, "")
            Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
            Call m_Details1.AddField(BodyCf)
            
            If Round((BalanceAmount), 2) <> 0 Then
               HaveValue = True
            End If
            
            Set TempARtTo = GetARTrn(m_ArCusAmountTo9, Ar.CUSCOD)
            Set TempARtToCn = GetARTrn(m_ArCusAmountToCn9, Ar.CUSCOD)
            
            Set TempARRcItTo = GetARRcpItem(m_ArCusPaidTo9, Ar.CUSCOD)
            Set TempARRcItToCn = GetARRcpItem(m_ArCusPaidToCn9, Ar.CUSCOD)
                     
            K = K + 1        '4
            Set HeadCf = m_Headers1.Fields(K)
            TempStr = FormatNumber(TempARtTo.AMOUNT, "")
            Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
            Call m_Details1.AddField(BodyCf)
            
            Total1(K) = Total1(K) + TempARtTo.AMOUNT
            Total3(K) = Total3(K) + TempARtTo.AMOUNT
            If Round((TempARtTo.AMOUNT), 2) <> 0 Then
               HaveValue = True
            End If
           
            
            K = K + 1    '5
            Set HeadCf = m_Headers1.Fields(K)
            TempStr = FormatNumber(TempARRcItTo.RCVAMT - TempARRcItToCn.RCVAMT + TempARtToCn.AMOUNT, "")
            Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
            Call m_Details1.AddField(BodyCf)
            
            Total1(K) = Total1(K) + TempARRcItTo.RCVAMT - TempARRcItToCn.RCVAMT + TempARtToCn.AMOUNT
             Total3(K) = Total3(K) + TempARRcItTo.RCVAMT - TempARRcItToCn.RCVAMT + TempARtToCn.AMOUNT
            If Round((TempARRcItTo.RCVAMT - TempARRcItToCn.RCVAMT + TempARtToCn.AMOUNT), 2) <> 0 Then
               HaveValue = True
            End If
            TempPaidCheck = Round((TempARRcItTo.RCVAMT - TempARRcItToCn.RCVAMT + TempARtToCn.AMOUNT), 2)
            
            K = K + 1
            Set HeadCf = m_Headers1.Fields(K)
            TempStr = FormatNumber(BalanceAmount + (TempARtTo.AMOUNT - TempARtToCn.AMOUNT) - (TempARRcItTo.RCVAMT - TempARRcItToCn.RCVAMT), "")
            Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
            Call m_Details1.AddField(BodyCf)
            
            Total1(K) = Total1(K) + (BalanceAmount) + (TempARtTo.AMOUNT - TempARtToCn.AMOUNT) - (TempARRcItTo.RCVAMT - TempARRcItToCn.RCVAMT)
            If Round(((BalanceAmount) + (TempARtTo.AMOUNT - TempARtToCn.AMOUNT) - (TempARRcItTo.RCVAMT - TempARRcItToCn.RCVAMT)), 2) <> 0 Then
               HaveValue = True
            End If
            
            If HaveValue Then
              If mcolParam("SUMMARY_FLAG") = "N" And mcolParam("ONLY_CUSTOMER") = "N" Then
                 Call m_Details1.GetString(1, TempStr1, TempStr2)
                 strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
                 Call Vsp.AddTable(strFormat, "", TempStr2)
               End If
               HaveValueSum = True
            End If
'          ' ของเดือนแรกที่เลือก 9
      End If
'
      If (RunMonth >= 10) Then                        ' ///////////////////////////////////////////////////////////////////////////////////////////////////
'         ' ของเดือนแรกที่เลือก 10

         FirstDate = DateAdd("M", 1, FirstDate)
         Call m_Details1.ClearField
               
            Set TempARt = GetARTrn(m_ArCusAmounts10, Ar.CUSCOD)      ' ar cod เป็น key  แต่ใน 1-16-1 หนึ่งคนแล้ววน 12 เดือน
            Set TempARtCn = GetARTrn(m_ArCusAmountCns10, Ar.CUSCOD)
            
            Set TempCnAll = GetARTrn(m_CnAllAmounts10, Ar.CUSCOD)
            
            Set TempARRcIt = GetARRcpItem(m_ArCusPaids10, Ar.CUSCOD)
            Set TempARRcItCn = GetARRcpItem(m_ArCusPaidCns10, Ar.CUSCOD)
               
            Set TempARRcIt3 = GetARRcpItemEx(m_PaidAllAmounts10, Ar.CUSCOD)
            Set TempARRcIt3Cn = GetARRcpItemEx(m_PaidAllAmountCns10, Ar.CUSCOD)
            
            HaveValue = False
            
            '===Head
            Set HeadCf = m_Headers1.Fields(1)
            TempStr = Format(Month(FirstDate), "00") & "/" & Format(Year(FirstDate) + 543, "00")
            Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
            Call m_Details1.AddField(BodyCf)
            
            Set HeadCf = m_Headers1.Fields(2)
            TempStr = Left(Ar.CUSNAM, 30) & "   (" & Ar.CUSCOD & ")"
            Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
            Call m_Details1.AddField(BodyCf)
            
            BalanceAmount = (TempARt.AMOUNT - TempARtCn.AMOUNT) - (TempARt.RCVAMT - TempARtCn.RCVAMT) - (TempARRcIt.RCVAMT - TempARRcItCn.RCVAMT) + (TempARRcIt3.RCVAMT - TempARRcIt3Cn.RCVAMT)
            BalanceAmount = BalanceAmount + TempCnAll.AMOUNT
                        
            K = 2
            K = K + 1    '3
            Set HeadCf = m_Headers1.Fields(K)
            TempStr = FormatNumber(BalanceAmount, "")
            Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
            Call m_Details1.AddField(BodyCf)
            
            If Round((BalanceAmount), 2) <> 0 Then
               HaveValue = True
            End If
            
            Set TempARtTo = GetARTrn(m_ArCusAmountTo10, Ar.CUSCOD)
            Set TempARtToCn = GetARTrn(m_ArCusAmountToCn10, Ar.CUSCOD)
            
            Set TempARRcItTo = GetARRcpItem(m_ArCusPaidTo10, Ar.CUSCOD)
            Set TempARRcItToCn = GetARRcpItem(m_ArCusPaidToCn10, Ar.CUSCOD)
                     
            K = K + 1        '4
            Set HeadCf = m_Headers1.Fields(K)
            TempStr = FormatNumber(TempARtTo.AMOUNT, "")
            Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
            Call m_Details1.AddField(BodyCf)
            
            Total1(K) = Total1(K) + TempARtTo.AMOUNT
             Total3(K) = Total3(K) + TempARtTo.AMOUNT
            If Round((TempARtTo.AMOUNT), 2) <> 0 Then
               HaveValue = True
            End If
           
            
            K = K + 1    '5
            Set HeadCf = m_Headers1.Fields(K)
            TempStr = FormatNumber(TempARRcItTo.RCVAMT - TempARRcItToCn.RCVAMT + TempARtToCn.AMOUNT, "")
            Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
            Call m_Details1.AddField(BodyCf)
            
            Total1(K) = Total1(K) + TempARRcItTo.RCVAMT - TempARRcItToCn.RCVAMT + TempARtToCn.AMOUNT
            Total3(K) = Total3(K) + TempARRcItTo.RCVAMT - TempARRcItToCn.RCVAMT + TempARtToCn.AMOUNT
            If Round((TempARRcItTo.RCVAMT - TempARRcItToCn.RCVAMT + TempARtToCn.AMOUNT), 2) <> 0 Then
               HaveValue = True
            End If
            TempPaidCheck = Round((TempARRcItTo.RCVAMT - TempARRcItToCn.RCVAMT + TempARtToCn.AMOUNT), 2)
            
            K = K + 1
            Set HeadCf = m_Headers1.Fields(K)
            TempStr = FormatNumber(BalanceAmount + (TempARtTo.AMOUNT - TempARtToCn.AMOUNT) - (TempARRcItTo.RCVAMT - TempARRcItToCn.RCVAMT), "")
            Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
            Call m_Details1.AddField(BodyCf)
            
            Total1(K) = Total1(K) + (BalanceAmount) + (TempARtTo.AMOUNT - TempARtToCn.AMOUNT) - (TempARRcItTo.RCVAMT - TempARRcItToCn.RCVAMT)
            If Round(((BalanceAmount) + (TempARtTo.AMOUNT - TempARtToCn.AMOUNT) - (TempARRcItTo.RCVAMT - TempARRcItToCn.RCVAMT)), 2) <> 0 Then
               HaveValue = True
            End If
            
            If HaveValue Then
              If mcolParam("SUMMARY_FLAG") = "N" And mcolParam("ONLY_CUSTOMER") = "N" Then
                 Call m_Details1.GetString(1, TempStr1, TempStr2)
                 strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
                 Call Vsp.AddTable(strFormat, "", TempStr2)
               End If
               HaveValueSum = True
            End If
'          ' ของเดือนแรกที่เลือก 10
      End If
'
      If (RunMonth >= 11) Then                        ' ///////////////////////////////////////////////////////////////////////////////////////////////////
'         ' ของเดือนแรกที่เลือก 11

         FirstDate = DateAdd("M", 1, FirstDate)
         Call m_Details1.ClearField
               
            Set TempARt = GetARTrn(m_ArCusAmounts11, Ar.CUSCOD)      ' ar cod เป็น key  แต่ใน 1-16-1 หนึ่งคนแล้ววน 12 เดือน
            Set TempARtCn = GetARTrn(m_ArCusAmountCns11, Ar.CUSCOD)
            
            Set TempCnAll = GetARTrn(m_CnAllAmounts11, Ar.CUSCOD)
            
            Set TempARRcIt = GetARRcpItem(m_ArCusPaids11, Ar.CUSCOD)
            Set TempARRcItCn = GetARRcpItem(m_ArCusPaidCns11, Ar.CUSCOD)
               
            Set TempARRcIt3 = GetARRcpItemEx(m_PaidAllAmounts11, Ar.CUSCOD)
            Set TempARRcIt3Cn = GetARRcpItemEx(m_PaidAllAmountCns11, Ar.CUSCOD)
            
            HaveValue = False
            
            '===Head
            Set HeadCf = m_Headers1.Fields(1)
            TempStr = Format(Month(FirstDate), "00") & "/" & Format(Year(FirstDate) + 543, "00")
            Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
            Call m_Details1.AddField(BodyCf)
            
            Set HeadCf = m_Headers1.Fields(2)
            TempStr = Left(Ar.CUSNAM, 30) & "   (" & Ar.CUSCOD & ")"
            Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
            Call m_Details1.AddField(BodyCf)
            
            BalanceAmount = (TempARt.AMOUNT - TempARtCn.AMOUNT) - (TempARt.RCVAMT - TempARtCn.RCVAMT) - (TempARRcIt.RCVAMT - TempARRcItCn.RCVAMT) + (TempARRcIt3.RCVAMT - TempARRcIt3Cn.RCVAMT)
            BalanceAmount = BalanceAmount + TempCnAll.AMOUNT
                        
            K = 2
            K = K + 1    '3
            Set HeadCf = m_Headers1.Fields(K)
            TempStr = FormatNumber(BalanceAmount, "")
            Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
            Call m_Details1.AddField(BodyCf)
            
            If Round((BalanceAmount), 2) <> 0 Then
               HaveValue = True
            End If
            
            Set TempARtTo = GetARTrn(m_ArCusAmountTo11, Ar.CUSCOD)
            Set TempARtToCn = GetARTrn(m_ArCusAmountToCn11, Ar.CUSCOD)
            
            Set TempARRcItTo = GetARRcpItem(m_ArCusPaidTo11, Ar.CUSCOD)
            Set TempARRcItToCn = GetARRcpItem(m_ArCusPaidToCn11, Ar.CUSCOD)
                     
            K = K + 1        '4
            Set HeadCf = m_Headers1.Fields(K)
            TempStr = FormatNumber(TempARtTo.AMOUNT, "")
            Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
            Call m_Details1.AddField(BodyCf)
            
            Total1(K) = Total1(K) + TempARtTo.AMOUNT
            Total3(K) = Total3(K) + TempARtTo.AMOUNT
            If Round((TempARtTo.AMOUNT), 2) <> 0 Then
               HaveValue = True
            End If
           
            
            K = K + 1    '5
            Set HeadCf = m_Headers1.Fields(K)
            TempStr = FormatNumber(TempARRcItTo.RCVAMT - TempARRcItToCn.RCVAMT + TempARtToCn.AMOUNT, "")
            Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
            Call m_Details1.AddField(BodyCf)
            
            Total1(K) = Total1(K) + TempARRcItTo.RCVAMT - TempARRcItToCn.RCVAMT + TempARtToCn.AMOUNT
            Total3(K) = Total3(K) + TempARRcItTo.RCVAMT - TempARRcItToCn.RCVAMT + TempARtToCn.AMOUNT
            If Round((TempARRcItTo.RCVAMT - TempARRcItToCn.RCVAMT + TempARtToCn.AMOUNT), 2) <> 0 Then
               HaveValue = True
            End If
            TempPaidCheck = Round((TempARRcItTo.RCVAMT - TempARRcItToCn.RCVAMT + TempARtToCn.AMOUNT), 2)
            
            K = K + 1
            Set HeadCf = m_Headers1.Fields(K)
            TempStr = FormatNumber(BalanceAmount + (TempARtTo.AMOUNT - TempARtToCn.AMOUNT) - (TempARRcItTo.RCVAMT - TempARRcItToCn.RCVAMT), "")
            Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
            Call m_Details1.AddField(BodyCf)
            
            Total1(K) = Total1(K) + (BalanceAmount) + (TempARtTo.AMOUNT - TempARtToCn.AMOUNT) - (TempARRcItTo.RCVAMT - TempARRcItToCn.RCVAMT)
            If Round(((BalanceAmount) + (TempARtTo.AMOUNT - TempARtToCn.AMOUNT) - (TempARRcItTo.RCVAMT - TempARRcItToCn.RCVAMT)), 2) <> 0 Then
               HaveValue = True
            End If
            
            If HaveValue Then
              If mcolParam("SUMMARY_FLAG") = "N" And mcolParam("ONLY_CUSTOMER") = "N" Then
                 Call m_Details1.GetString(1, TempStr1, TempStr2)
                 strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
                 Call Vsp.AddTable(strFormat, "", TempStr2)
               End If
               HaveValueSum = True
            End If
            '          ' ของเดือนแรกที่เลือก 11
      End If

      If (RunMonth >= 12) Then                        ' ///////////////////////////////////////////////////////////////////////////////////////////////////
'         ' ของเดือนแรกที่เลือก 12

         FirstDate = DateAdd("M", 1, FirstDate)
         Call m_Details1.ClearField
               
            Set TempARt = GetARTrn(m_ArCusAmounts12, Ar.CUSCOD)      ' ar cod เป็น key  แต่ใน 1-16-1 หนึ่งคนแล้ววน 12 เดือน
            Set TempARtCn = GetARTrn(m_ArCusAmountCns12, Ar.CUSCOD)
            
            Set TempCnAll = GetARTrn(m_CnAllAmounts12, Ar.CUSCOD)
            
            Set TempARRcIt = GetARRcpItem(m_ArCusPaids12, Ar.CUSCOD)
            Set TempARRcItCn = GetARRcpItem(m_ArCusPaidCns12, Ar.CUSCOD)
               
            Set TempARRcIt3 = GetARRcpItemEx(m_PaidAllAmounts12, Ar.CUSCOD)
            Set TempARRcIt3Cn = GetARRcpItemEx(m_PaidAllAmountCns12, Ar.CUSCOD)
            
            HaveValue = False
            
            '===Head
            Set HeadCf = m_Headers1.Fields(1)
            TempStr = Format(Month(FirstDate), "00") & "/" & Format(Year(FirstDate) + 543, "00")
            Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
            Call m_Details1.AddField(BodyCf)
            
            Set HeadCf = m_Headers1.Fields(2)
            TempStr = Left(Ar.CUSNAM, 30) & "   (" & Ar.CUSCOD & ")"
            Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
            Call m_Details1.AddField(BodyCf)
            
            BalanceAmount = (TempARt.AMOUNT - TempARtCn.AMOUNT) - (TempARt.RCVAMT - TempARtCn.RCVAMT) - (TempARRcIt.RCVAMT - TempARRcItCn.RCVAMT) + (TempARRcIt3.RCVAMT - TempARRcIt3Cn.RCVAMT)
            BalanceAmount = BalanceAmount + TempCnAll.AMOUNT
                        
            K = 2
            K = K + 1    '3
            Set HeadCf = m_Headers1.Fields(K)
            TempStr = FormatNumber(BalanceAmount, "")
            Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
            Call m_Details1.AddField(BodyCf)
            
            If Round((BalanceAmount), 2) <> 0 Then
               HaveValue = True
            End If
            
            Set TempARtTo = GetARTrn(m_ArCusAmountTo12, Ar.CUSCOD)
            Set TempARtToCn = GetARTrn(m_ArCusAmountToCn12, Ar.CUSCOD)
            
            Set TempARRcItTo = GetARRcpItem(m_ArCusPaidTo12, Ar.CUSCOD)
            Set TempARRcItToCn = GetARRcpItem(m_ArCusPaidToCn12, Ar.CUSCOD)
                     
            K = K + 1        '4
            Set HeadCf = m_Headers1.Fields(K)
            TempStr = FormatNumber(TempARtTo.AMOUNT, "")
            Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
            Call m_Details1.AddField(BodyCf)
            
            Total1(K) = Total1(K) + TempARtTo.AMOUNT
            Total3(K) = Total3(K) + TempARtTo.AMOUNT
            If Round((TempARtTo.AMOUNT), 2) <> 0 Then
               HaveValue = True
            End If
           
            
            K = K + 1    '5
            Set HeadCf = m_Headers1.Fields(K)
            TempStr = FormatNumber(TempARRcItTo.RCVAMT - TempARRcItToCn.RCVAMT + TempARtToCn.AMOUNT, "")
            Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
            Call m_Details1.AddField(BodyCf)
            
            Total1(K) = Total1(K) + TempARRcItTo.RCVAMT - TempARRcItToCn.RCVAMT + TempARtToCn.AMOUNT
              Total3(K) = Total3(K) + TempARRcItTo.RCVAMT - TempARRcItToCn.RCVAMT + TempARtToCn.AMOUNT
            If Round((TempARRcItTo.RCVAMT - TempARRcItToCn.RCVAMT + TempARtToCn.AMOUNT), 2) <> 0 Then
               HaveValue = True
            End If
            TempPaidCheck = Round((TempARRcItTo.RCVAMT - TempARRcItToCn.RCVAMT + TempARtToCn.AMOUNT), 2)
            
            K = K + 1
            Set HeadCf = m_Headers1.Fields(K)
            TempStr = FormatNumber(BalanceAmount + (TempARtTo.AMOUNT - TempARtToCn.AMOUNT) - (TempARRcItTo.RCVAMT - TempARRcItToCn.RCVAMT), "")
            Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
            Call m_Details1.AddField(BodyCf)
            
            Total1(K) = Total1(K) + (BalanceAmount) + (TempARtTo.AMOUNT - TempARtToCn.AMOUNT) - (TempARRcItTo.RCVAMT - TempARRcItToCn.RCVAMT)

            If Round(((BalanceAmount) + (TempARtTo.AMOUNT - TempARtToCn.AMOUNT) - (TempARRcItTo.RCVAMT - TempARRcItToCn.RCVAMT)), 2) <> 0 Then
               HaveValue = True
            End If
            
            If HaveValue Then
              If mcolParam("SUMMARY_FLAG") = "N" And mcolParam("ONLY_CUSTOMER") = "N" Then
                 Call m_Details1.GetString(1, TempStr1, TempStr2)
                 strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
                 Call Vsp.AddTable(strFormat, "", TempStr2)
               End If
               HaveValueSum = True
            End If
                      ' ของเดือนแรกที่เลือก 12
      End If
'
   
   

      
 '     If (RunMonth >= 2) Then
         ' รวม ลูกค้า
         Call m_Details1.ClearField
          HaveValue = False

          '===Head
         If mcolParam("SUMMARY_FLAG") = "N" And mcolParam("ONLY_CUSTOMER") = "Y" Then
               TempStr = i
          Else
               TempStr = "รวม"
          End If
         Set HeadCf = m_Headers1.Fields(1)
          Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
         Call m_Details1.AddField(BodyCf)
          'debug.print Ar.CUSCOD
          Set HeadCf = m_Headers1.Fields(2)
          TempStr = Left(Ar.CUSCOD & " / " & Ar.CUSNAM, 30)
          Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
          Call m_Details1.AddField(BodyCf)
          
          K = 2
          K = K + 1
          Set HeadCf = m_Headers1.Fields(K)
          TempStr = FormatNumber(Total1(K), "")
          Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
          Call m_Details1.AddField(BodyCf)
          Total3(K) = Total1(K) + Total3(K)     'พนักงานขาย
          If Round((Total1(K)), 2) <> 0 Then
             HaveValue = True
          End If
          
          K = K + 1
          Set HeadCf = m_Headers1.Fields(K)
          TempValue = Total1(K)
          TempStr = FormatNumber(TempValue, "")
          Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
          Call m_Details1.AddField(BodyCf)
          If Round((Total1(K)), 2) <> 0 Then
             HaveValue = True
          End If
          
          K = K + 1
          Set HeadCf = m_Headers1.Fields(K)
          TempValue = Total1(K)
          TempStr = FormatNumber(TempValue, "")
          Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
          Call m_Details1.AddField(BodyCf)
          If Round((Total1(K)), 2) <> 0 Then
             HaveValue = True
          End If
          
          K = K + 1
          Set HeadCf = m_Headers1.Fields(K)
          TempValue = Total1(K - 3) + Total1(K - 2) - Total1(K - 1)
          TempStr = FormatNumber(TempValue, "")
          Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
          Call m_Details1.AddField(BodyCf)
          Total3(K) = Total3(K - 3) + Total3(K - 2) - Total3(K - 1)    'พนักงานขาย
          If Round((Total1(K - 3) + Total1(K - 2) - Total1(K - 1)), 2) <> 0 Then
             HaveValue = True
          End If
          
          
          If HaveValue Then
                     If mcolParam("SUMMARY_FLAG") = "N" Then
                               If mcolParam("ONLY_CUSTOMER") = "N" Then
                              Vsp.FontBold = True
                              End If
                        Call m_Details1.GetString(1, TempStr1, TempStr2)
                        strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
                        Call Vsp.AddTable(strFormat, "", TempStr2)
                        Vsp.FontBold = False
                        i = i + 1
                      End If
                      HaveValueSum = True
                   End If
                   
                   For j = 1 To UBound(Total3)
                     Total1(j) = 0
                  Next j
  '    End If
      
      
Rs.MoveNext
 Wend
      

   
       Set Ar = Nothing
       
             Call GenerateFooter("* รวม", PrevKey3 & "  " & PrevKey4, "", Total3, m_Details1)                '
             Call m_Details1.GetString(2, TempStr1, TempStr2)
             strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
        If mcolParam("SUMMARY_FLAG") = "N" Then
             Vsp.FontBold = True
        Else
            Vsp.FontBold = False
         End If
             Call Vsp.AddTable(strFormat, "", TempStr2)
             Vsp.FontBold = False

               Total2(3) = Total2(3) + Total3(3)
               Total2(4) = Total2(4) + Total3(4)
               Total2(5) = Total2(5) + Total3(5)
               Total2(6) = Total2(6) + Total3(6)

         TempRs.MoveNext
      Wend

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
Dim Sg As CCustomer     'CustomerGroup
Dim Apm As CARMas
   For Each Apm In Des
      Set Sg = GetObject("CCustomer", Src, Trim(Apm.CUSCOD), False)
      If Sg Is Nothing Then
          Set Sg = New CCustomer
          Sg.CUSTOMER_CODE = Apm.CUSCOD
          Sg.CUSTOMER_NAME = Apm.CUSNAM
          Call Src.Add(Sg, Trim(Sg.CUSTOMER_CODE))
      End If
   Next Apm
End Sub




