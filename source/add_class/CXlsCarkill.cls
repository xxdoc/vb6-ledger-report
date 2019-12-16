VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CXlsCarkill"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'CXlsFarm

Public InvNo As String
Public ShipDate As Date
Public FeedNo As String
Public Quantity As String
Public Price_Bag1 As String

Public NetPrice As String
Public Trans As String
Public TotalAmount As String
Public Trans50 As String
Public Price_Bag50 As String

Public NetPayment As String
Public CN As String
Public DueDate As String
Public Today As String
Public OutStandingDay As String

Public Remarks As String
Public ROW As Long

'========================
Public FromDate As Date
Public ToDate As Date
Public SheetIndex As Long
Public SheetName As String
Public ShortName As String
Public DateIndex As Long
Public SumKilo As Double
Public SumNetpay As Double
Public m_Farm As Collection  ' ที่จะเก็บคลาส CXlsCarkill
Public m_DistanceFarm As Collection

Public sumFlag As String

Public mySheetName As String
Public sigmaKilo As Double
Public sigmaNetpay As Double
