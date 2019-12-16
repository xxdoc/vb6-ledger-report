VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CChartTotal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public SALE_ID As String
Public AREA_ID  As Long
Public GOODS_GROUP_ID As Long

Public m_Cus As Collection
'Public CUS_ID  As Long
Public TOTAL1_SUM As Double
Public TOTAL2_SUM As Double
Public MINUS As String

Public REAL_COM1 As Double                        ' จ่ายจริงคอม 1
Public REAL_COM2 As Double                       ' จริง เก็บเงิน
Public REAL_INCENTIVE As Double
