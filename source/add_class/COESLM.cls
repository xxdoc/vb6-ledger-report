VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "COESLM"
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

Public SLMCOD As String
Public SLMNAM  As String

Public FROM_SALE_CODE As String
Public TO_SALE_CODE As String

Public KEY_ID As Long
Public KEY_LOOKUP As String
Private Sub Class_Initialize()
   Set m_Conn = glbDatabaseMngr.DBConnection
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

   SelectStr1 = "SELECT OES.*  "
   SelectStr2 = "SELECT COUNT(*)  "
   
   FromStr = "FROM OESLM OES "
      
   If Trim(SLMCOD) <> "" Then
      If WildCard(SLMCOD, SubLen, NewStr) Then
         If WhereStr = "" Then
            WhereStr = "WHERE (OES.SLMCOD LIKE '" & ChangeQuote(Trim(SLMCOD)) & "')"
         Else
            WhereStr = WhereStr & "AND (OES.SLMCOD LIKE '" & ChangeQuote(Trim(SLMCOD)) & "')"
         End If
      Else
         If WhereStr = "" Then
            WhereStr = "WHERE (OES.SLMCOD = '" & ChangeQuote(Trim(SLMCOD)) & "')"
         Else
            WhereStr = WhereStr & "AND (OES.SLMCOD = '" & ChangeQuote(Trim(SLMCOD)) & "')"
         End If
      End If
   End If
   
   If Trim(SLMNAM) <> "" Then
      If WildCard(SLMNAM, SubLen, NewStr) Then
         If WhereStr = "" Then
            WhereStr = "WHERE (SLMNAM LIKE '" & ChangeQuote(Trim(SLMNAM)) & "')"
         Else
            WhereStr = WhereStr & "AND (SLMNAM LIKE '" & ChangeQuote(Trim(SLMNAM)) & "')"
         End If
      Else
         If WhereStr = "" Then
            WhereStr = "WHERE (SLMNAM = '" & ChangeQuote(Trim(SLMNAM)) & "')"
         Else
            WhereStr = WhereStr & "AND (SLMNAM = '" & ChangeQuote(Trim(SLMNAM)) & "')"
         End If
      End If
   End If
   
   If Len(FROM_SALE_CODE) > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (OES.SLMCOD >= '" & Trim(ChangeQuote(FROM_SALE_CODE)) & "') "
      Else
         WhereStr = WhereStr & "AND (OES.SLMCOD >= '" & Trim(ChangeQuote(FROM_SALE_CODE)) & "') "
      End If
   End If
   
   If Len(TO_SALE_CODE) > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (OES.SLMCOD <= '" & Trim(ChangeQuote(TO_SALE_CODE)) & "') "
      Else
         WhereStr = WhereStr & "AND (OES.SLMCOD <= '" & Trim(ChangeQuote(TO_SALE_CODE)) & "') "
      End If
   End If
   
   OrderByStr = ""
   OrderTypeStr = ""
   
   If OrderBy = 1 Then
      OrderByStr = "ORDER BY SLMCOD "
   ElseIf OrderBy = 2 Then
      OrderByStr = "ORDER BY SLMNAM "
   Else
      OrderByStr = "ORDER BY SLMCOD "
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
   SLMCOD = Trim(NVLS(Rs("SLMCOD"), ""))
   SLMNAM = Trim(NVLS(Rs("SLMNAM"), ""))
   
   KEY_LOOKUP = SLMCOD
   KEY_ID = Val(SLMCOD)
End Function


