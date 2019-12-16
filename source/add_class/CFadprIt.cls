VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CFadprIt"
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

Public FASCOD As String
Public DOCNUM As String
Public DOCDAT As Date
Public DPRVAL As Double

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
Dim TempDate As String
Dim GroupByStr As String

   If Ind = 1 Then
      SelectStr1 = "SELECT ART.*, DTOS(ART.DOCDAT) DOCDAT_STR "
      SelectStr2 = "SELECT COUNT(*)  "
      
      FromStr = "FROM FADPRIT ART "
   End If
   
   If Trim(FASCOD) <> "" Then
      If WildCard(FASCOD, SubLen, NewStr) Then
         If WhereStr = "" Then
            WhereStr = "WHERE (FASCOD LIKE '" & ChangeQuote(Trim(FASCOD)) & "')"
         Else
            WhereStr = WhereStr & "AND (FASCOD LIKE '" & ChangeQuote(Trim(FASCOD)) & "')"
         End If
      Else
         If WhereStr = "" Then
            WhereStr = "WHERE (FASCOD = '" & ChangeQuote(Trim(FASCOD)) & "')"
         Else
            WhereStr = WhereStr & "AND (FASCOD = '" & ChangeQuote(Trim(FASCOD)) & "')"
         End If
      End If
   End If
                                                                              
   GroupByStr = ""
   
   OrderByStr = ""
   OrderTypeStr = ""
   
   If Ind = 1 Then
      If OrderBy = 1 Then
         OrderByStr = "ORDER BY FASCOD "
      ElseIf OrderBy = 2 Then
         OrderByStr = "ORDER BY FASCOD "
      Else
         OrderByStr = "ORDER BY FASCOD "
      End If
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
   If Ind = 1 Then
      FASCOD = NVLS(Rs("FASCOD"), "")
      DOCNUM = NVLS(Rs("DOCNUM"), "")
      DOCDAT = InternalDateToDate(NVLS(Rs("DOCDAT_STR"), ""))
      DPRVAL = NVLD(Rs("DPRVAL"), 0)
   End If
End Function

