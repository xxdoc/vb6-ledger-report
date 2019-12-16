VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CXlsFarm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'CXlsFarm

Public farmName1 As String
Public farmName2 As String
Public weekNum As Long
Public fromDateHead As String
Public toDateHead As String

Public FoodNum   As String
Public F_unitName   As String
Public F_unitValue   As Double
Public F_date   As String
Public F_cost   As Double
Public F_Value As Double
Public F_TRANS_PRICE As Double
Public F_destination As String

Public unit_multiply As Double

