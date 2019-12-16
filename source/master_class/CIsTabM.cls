VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CIsTab"
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

Public TABTYP As String
Public TYPCOD As String
Public SHORTNAME As String
Public TYPDES As String

Private Sub Class_Initialize()
   Set m_Conn = glbDatabaseMngr.DBConnection
End Sub

'In this class we will not handle any error or any transaction
Public Function QueryData(Rs As ADODB.Recordset, ItemCount As Long)
Dim SQL1 As String
Dim SQL2 As String
Dim SelectStr1 As String
Dim SelectStr2 As String
Dim FromStr As String
Dim WhereStr As String
Dim OrderByStr As String
Dim OrderTypeStr As String

   SelectStr1 = "SELECT UG.*  "
   SelectStr2 = "SELECT COUNT(*)  "
   
   FromStr = "FROM TABTYP UG "
   
   WhereStr = ""
   If TABTYP_ID > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (TABTYP_ID = " & TABTYP_ID & ") "
      Else
         WhereStr = "AND (TABTYP_ID = " & TABTYP_ID & ") "
      End If
   End If
   
   OrderByStr = ""
   OrderTypeStr = ""
   If OrderBy > 0 Then
      If OrderBy = 1 Then
         OrderByStr = "ORDER BY TABTYP_ID "
      ElseIf OrderBy = 2 Then
         OrderByStr = "ORDER BY TABTYP_NAME "
      Else
         OrderByStr = "ORDER BY TABTYP_ID "
      End If
      
      If OrderType = 1 Then
         OrderTypeStr = "ASC "
      ElseIf OrderType = 2 Then
         OrderTypeStr = "DESC "
      Else
         OrderTypeStr = "DESC "
      End If
   End If
   
   SQL1 = SelectStr1 & FromStr & WhereStr & OrderByStr & OrderTypeStr
   SQL2 = SelectStr2 & FromStr & WhereStr
   
   Call glbDatabaseMngr.GetRs(SQL1, SQL2, True, ItemCount, Rs, glbErrorLog)
End Function

Public Function PopulateFromRS(Ind As Long, Rs As ADODB.Recordset)
   TABTYP_ID = NVLI(Rs("TABTYP_ID"), -1)
   TABTYP_NAME = NVLS(Rs("TABTYP_NAME"), "")
   TABTYP_NO = NVLS(Rs("TABTYP_NO"), "")
   CONTINENT_ID = NVLI(Rs("CONTINENT_ID"), -1)
End Function


