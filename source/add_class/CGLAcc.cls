VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CGLAcc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_Conn As ADODB.Connection

Public AddEditMode As SHOW_MODE_TYPE
Public QueryFlag As Long
Public OrderBy As Long
Public OrderType As Long
Public Flag As String
' End of generic data member



Public ACCNAM As String
Public ACCNUM As String

Public LEVEL As Long
Public GROUP As Long
Public ACCTYP As Long
Public BEGCUR As Double
Public BAL1CY As Double
Public BAL2CY As Double
Public BAL3CY As Double
Public BAL4CY As Double
Public BAL5CY As Double
Public BAL6CY As Double
Public BAL7CY As Double
Public BAL8CY As Double
Public BAL9CY As Double
Public BAL10CY As Double
Public BAL11CY As Double
Public BAL12CY As Double

Public NATURE As Long


Public FROM_VOUCHER_DATE As Date
Public TO_VOUCHER_DATE As Date

Public KEY_ID As Long
Public KEY_LOOKUP As String

Private Sub Class_Initialize()
   Set m_Conn = glbDatabaseMngr.DBConnection
   FROM_VOUCHER_DATE = -1
   TO_VOUCHER_DATE = -1
End Sub

'In this class we will not handle any error or any transaction
Public Function QueryData(Ind As Long, Rs As ADODB.Recordset, ItemCount As Long)
Dim SQL1 As String
Dim SQL2 As String
Dim SelectStr1 As String
Dim SelectStr2 As String
Dim FromStr As String
Dim WhereStr As String
Dim OrderByStr As String
Dim OrderTypeStr As String
Dim SubLen As Long
Dim NewStr As String
Dim TempDate As String
Dim GroupByStr As String
      SelectStr1 = "SELECT GL.*  "
      SelectStr2 = "SELECT COUNT(*)  "
   
   FromStr = "FROM  GLACC  GL "
         
   OrderByStr = ""
   OrderTypeStr = ""
   
   If OrderBy = 1 Then
      OrderByStr = "ORDER BY ACCNAM "
   ElseIf OrderBy = 2 Then
      OrderByStr = "ORDER BY ACCNUM "
   Else
      OrderByStr = "ORDER BY ACCNAM "
   End If
   
   If OrderType = 1 Then
      OrderTypeStr = "ASC "
   ElseIf OrderType = 2 Then
      OrderTypeStr = "DESC "
   Else
      OrderTypeStr = "ASC "
   End If
   
   SQL1 = SelectStr1 & FromStr & WhereStr & GroupByStr & OrderByStr & OrderTypeStr
   SQL2 = SelectStr2 & FromStr & WhereStr
   
   Call glbDatabaseMngr.GetRs(SQL1, SQL2, True, ItemCount, Rs, glbErrorLog)
End Function

Public Function PopulateFromRS(Ind As Long, Rs As ADODB.Recordset)
   
      ACCNAM = NVLS(Rs("ACCNAM"), "")
      ACCNUM = NVLS(Rs("ACCNUM"), "")
      
      LEVEL = NVLI(Rs("LEVEL"), -1)
      GROUP = NVLI(Rs("GROUP"), -1)
      ACCTYP = NVLI(Rs("ACCTYP"), -1)
      NATURE = NVLI(Rs("NATURE"), -1)
      
      BEGCUR = NVLD(Rs("BEGCUR"), 0)
      
      BAL1CY = NVLD(Rs("BAL1CY"), 0)
      BAL2CY = NVLD(Rs("BAL2CY"), 0)
      BAL3CY = NVLD(Rs("BAL3CY"), 0)
      BAL4CY = NVLD(Rs("BAL4CY"), 0)
      BAL5CY = NVLD(Rs("BAL5CY"), 0)
      BAL6CY = NVLD(Rs("BAL6CY"), 0)
      BAL7CY = NVLD(Rs("BAL7CY"), 0)
      BAL8CY = NVLD(Rs("BAL8CY"), 0)
      BAL9CY = NVLD(Rs("BAL9CY"), 0)
      BAL10CY = NVLD(Rs("BAL10CY"), 0)
      BAL11CY = NVLD(Rs("BAL11CY"), 0)
      BAL12CY = NVLD(Rs("BAL12CY"), 0)
      
      KEY_ID = Rs.Bookmark
      KEY_LOOKUP = ACCNUM
End Function



