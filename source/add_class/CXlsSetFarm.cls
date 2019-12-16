VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CXlsSetFarm"
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

Public SET_FARM_ID As Long
Public MAIN_FARM_ID As Long
Public MAIN_FARM_NAME As String
Public XLS_UNIT_NAME As String
Public XLS_UNIT_ID As Long
Public SET_FARM_PRICE As Double

Private Sub Class_Initialize()
   Set m_Conn = glbDatabaseMngr.DBLegacyConnection
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
Dim TempDate As String

If Ind = 1 Then
   SelectStr1 = "SELECT SF.*  "
   SelectStr1 = SelectStr1 & ", XU.XLS_UNIT_NAME "
   SelectStr1 = SelectStr1 & ", MF.MAIN_FARM_NAME "
   SelectStr2 = "SELECT COUNT(*)  "
   
   FromStr = "FROM ((XLS_SET_FARM SF "
   FromStr = FromStr & "LEFT OUTER JOIN XLS_UNIT XU ON (SF.XLS_UNIT_ID = XU.XLS_UNIT_ID)) "
   FromStr = FromStr & "LEFT OUTER JOIN XLS_MAIN_FARM MF ON (SF.MAIN_FARM_ID = MF.MAIN_FARM_ID)) "
   
   WhereStr = ""
'ElseIf ind = 2 Then

End If
   
'   If OwnerEnterprise Then
'      ENTERPRISE_CODE = glbCompanyCode
'   End If
   
   If MAIN_FARM_NAME <> "" Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (MF.MAIN_FARM_NAME = '" & ChangeQuote(Trim(MAIN_FARM_NAME)) & "') "
      Else
         WhereStr = WhereStr & "AND (MF.MAIN_FARM_NAME = '" & ChangeQuote(Trim(MAIN_FARM_NAME)) & "') "
      End If
   End If

   If SET_FARM_ID > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (SF.SET_FARM_ID = " & ChangeQuote(Trim(SET_FARM_ID)) & ") "
      Else
         WhereStr = WhereStr & "AND (SF.SET_FARM_ID = " & ChangeQuote(Trim(SET_FARM_ID)) & ") "
      End If
   End If

    If XLS_UNIT_ID > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (SF.XLS_UNIT_ID = " & ChangeQuote(Trim(XLS_UNIT_ID)) & ") "
      Else
         WhereStr = WhereStr & "AND (SF.XLS_UNIT_ID = " & ChangeQuote(Trim(XLS_UNIT_ID)) & ") "
      End If
   End If
   
   OrderByStr = ""
   OrderTypeStr = ""

   If OrderBy = 1 Then
      OrderByStr = "ORDER BY MF.MAIN_FARM_NAME "
   ElseIf OrderBy = 2 Then
      OrderByStr = "ORDER BY SF.XLS_UNIT_ID "
   Else
      OrderByStr = "ORDER BY MF.MAIN_FARM_NAME  ASC , SF.XLS_UNIT_ID "
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
  ' 'debug.print SQL1
   Call glbDatabaseMngr.GetLegacyRs(SQL1, SQL2, True, ItemCount, Rs, glbErrorLog)
End Function

Public Function PopulateFromRS(Ind As Long, Rs As ADODB.Recordset)
If Ind = 1 Then
   SET_FARM_ID = NVLI(Rs("SET_FARM_ID"), -1)
   MAIN_FARM_NAME = NVLS(Rs("MAIN_FARM_NAME"), "")
   XLS_UNIT_NAME = NVLS(Rs("XLS_UNIT_NAME"), "")
   SET_FARM_PRICE = NVLD(Rs("SET_FARM_PRICE"), -1)
   MAIN_FARM_ID = NVLI(Rs("MAIN_FARM_ID"), -1)
   XLS_UNIT_ID = NVLI(Rs("XLS_UNIT_ID"), -1)
End If
End Function
Public Function AddEditData()
Dim ServerDateTime As String
Dim SQL1 As String

   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)

   If AddEditMode = SHOW_ADD Then
      SQL1 = "INSERT INTO XLS_SET_FARM " & _
                     "( " & _
                     "   MAIN_FARM_ID, " & _
                     "   SET_FARM_PRICE, " & _
                     "   XLS_UNIT_ID " & _
                     ") " & _
                     "VALUES " & _
                     "( " & _
                        "'" & ChangeQuote(Trim(MAIN_FARM_ID)) & "'," & _
                        "" & ChangeQuote(Trim(SET_FARM_PRICE)) & "," & _
                        "" & ChangeQuote(Trim(XLS_UNIT_ID)) & "" & _
                     ") "

   Else 'Show edit
      SQL1 = "UPDATE XLS_SET_FARM SET " & _
                     "   MAIN_FARM_ID = " & ChangeQuote(Trim(MAIN_FARM_ID)) & ", " & _
                     "   SET_FARM_PRICE = " & ChangeQuote(Trim(SET_FARM_PRICE)) & ", " & _
                     "   XLS_UNIT_ID = " & ChangeQuote(Trim(XLS_UNIT_ID)) & " " & _
                     "WHERE SET_FARM_ID = " & SET_FARM_ID
   End If
   
   If (AddEditMode = SHOW_ADD) Or (AddEditMode = SHOW_EDIT) Then
      m_Conn.Execute (SQL1)
   End If
End Function

Public Function DeleteData()
Dim SQL1 As String
      SQL1 = "DELETE FROM XLS_SET_FARM WHERE SET_FARM_ID = " & SET_FARM_ID
      m_Conn.Execute (SQL1)
End Function
