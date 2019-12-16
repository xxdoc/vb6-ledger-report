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

Public SHORTNAM As String
Public TYPDES As String

Private Sub Class_Initialize()
   Set m_Conn = glbDatabaseMngr.DBConnection
End Sub

Public Function GetKey(Ind As Long) As String
   If Ind = 1 Then
      GetKey = TABTYP & "-" & TYPCOD
   End If
End Function

Public Function QueryData(Rs As ADODB.Recordset, ItemCount As Long, Optional Ind As Long = 1)
Dim SQL1 As String
Dim SQL2 As String
Dim SelectStr1 As String
Dim SelectStr2 As String
Dim FromStr As String
Dim WhereStr As String
Dim OrderByStr As String
Dim OrderTypeStr As String

   SelectStr1 = "SELECT TB.*  "
   SelectStr2 = "SELECT COUNT(*)  "
   
   FromStr = "FROM ISTAB TB "
   
 If Ind = 1 Then
   WhereStr = ""
   Else: Ind = 2
       WhereStr = " WHERE (TYPCOD >='01') AND (TYPCOD <= '10') "
   End If
   
   If Len(TABTYP) > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (TABTYP = '" & Trim(ChangeQuote(TABTYP)) & "') "
      Else
         WhereStr = WhereStr & "AND (TABTYP = '" & Trim(ChangeQuote(TABTYP)) & "') "
      End If
   End If
   
   If Len(TYPCOD) > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (TYPCOD = '" & Trim(ChangeQuote(TYPCOD)) & "') "
      Else
         WhereStr = WhereStr & "AND (TYPCOD = '" & Trim(ChangeQuote(TYPCOD)) & "') "
      End If
   End If
   
   OrderByStr = ""
   OrderTypeStr = ""
   
   If OrderBy = 1 Then
      OrderByStr = "ORDER BY TYPCOD "
   ElseIf OrderBy = 2 Then
      OrderByStr = "ORDER BY TYPDES "
   Else
      OrderByStr = "ORDER BY TYPCOD "
   End If
   
   If OrderType = 1 Then
      OrderTypeStr = "ASC "
   ElseIf OrderType = 2 Then
      OrderTypeStr = "DESC "
   Else
      OrderTypeStr = "ASC "
   End If
   
   SQL1 = SelectStr1 & FromStr & WhereStr & OrderByStr & OrderTypeStr
   SQL2 = SelectStr2 & FromStr & WhereStr
   
   Call glbDatabaseMngr.GetRs(SQL1, SQL2, True, ItemCount, Rs, glbErrorLog)
End Function
Public Function PopulateFromRS(Ind As Long, Rs As ADODB.Recordset)
   TABTYP = NVLS(Rs("TABTYP"), "")
   TYPCOD = NVLS(Rs("TYPCOD"), "")
   TYPDES = NVLS(Rs("TYPDES"), "")
   SHORTNAM = NVLS(Rs("SHORTNAM"), "")
End Function


