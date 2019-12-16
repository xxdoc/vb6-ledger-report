VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CIsInfo"
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

Public THINAM As String
Public ADDR01  As String
Public ADDR02 As String
Public TELNUM As String

Private Sub Class_Initialize()
   Set m_Conn = glbDatabaseMngr.DBConnection
End Sub

Public Function PackAddress()
Dim TempStr As String

   If Len(ADDR02) > 0 Then
      PackAddress = ADDR01 & vbCrLf & PackAddress
   Else
      PackAddress = ADDR01
   End If
End Function

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

   SelectStr1 = "SELECT ART.THINAM "
   SelectStr2 = "SELECT COUNT(*)  "
   
   FromStr = " FROM ISINFO ART "
   
'   OrderByStr = " ORDER BY ART.THINAM "
'   OrderTypeStr = " ASC"
   
   SQL1 = SelectStr1 & FromStr & WhereStr & OrderByStr & OrderTypeStr
   SQL2 = SelectStr2 & FromStr & WhereStr
   
   Call glbDatabaseMngr.GetRs(SQL1, SQL2, False, ItemCount, Rs, glbErrorLog)
End Function

Public Function PopulateFromRS(Ind As Long, Rs As ADODB.Recordset)
   THINAM = Trim(NVLS(Rs("THINAM"), ""))
'   ADDR01 = Trim(NVLS(Rs("ADDR01"), ""))
'   ADDR02 = Trim(NVLS(Rs("ADDR02"), ""))
'   TELNUM = Trim(NVLS(Rs("TELNUM"), ""))
   
End Function
