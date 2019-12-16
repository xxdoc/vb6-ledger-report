VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CXlsUnit"
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

Public XLS_UNIT_ID As Long
Public XLS_UNIT_NAME As String
Public XLS_UNIT_MULTIPLY As Long
Public XLS_UNIT_LIMIT As Long

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
   SelectStr1 = "SELECT CMS.*  "
   SelectStr2 = "SELECT COUNT(*)  "
   
   FromStr = "FROM XLS_UNIT CMS "
   
   WhereStr = ""
'ElseIf ind = 2 Then

End If

    If XLS_UNIT_ID > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (XLS_UNIT_ID = " & ChangeQuote(Trim(XLS_UNIT_ID)) & ") "
      Else
         WhereStr = WhereStr & "AND (XLS_UNIT_ID = " & ChangeQuote(Trim(XLS_UNIT_ID)) & ") "
      End If
   End If
   
   OrderByStr = ""
   OrderTypeStr = ""

   OrderByStr = "ORDER BY CMS.XLS_UNIT_NAME "
   OrderTypeStr = "ASC "
   
   SQL1 = SelectStr1 & FromStr & WhereStr & OrderByStr & OrderTypeStr
   SQL2 = SelectStr2 & FromStr & WhereStr
   'debug.print SQL1
   Call glbDatabaseMngr.GetLegacyRs(SQL1, SQL2, True, ItemCount, Rs, glbErrorLog)
End Function

Public Function PopulateFromRS(Ind As Long, Rs As ADODB.Recordset)
If Ind = 1 Then
   XLS_UNIT_ID = NVLI(Rs("XLS_UNIT_ID"), -1)
   XLS_UNIT_NAME = NVLS(Rs("XLS_UNIT_NAME"), "")
   XLS_UNIT_MULTIPLY = NVLI(Rs("XLS_UNIT_MULTIPLY"), -1)
   XLS_UNIT_LIMIT = NVLI(Rs("XLS_UNIT_LIMIT"), -1)
End If
End Function
Public Function AddEditData()
Dim ServerDateTime As String
Dim SQL1 As String

   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   
'   ENTERPRISE_CODE = glbCompanyCode
   
   If AddEditMode = SHOW_ADD Then
      SQL1 = "INSERT INTO XLS_UNIT " & _
                     "( " & _
                     "   XLS_UNIT_NAME, " & _
                     "   XLS_UNIT_MULTIPLY, " & _
                     "   XLS_UNIT_LIMIT " & _
                     ") " & _
                     "VALUES " & _
                     "( " & _
                        "'" & ChangeQuote(Trim(XLS_UNIT_NAME)) & "'," & _
                        "" & ChangeQuote(Trim(XLS_UNIT_MULTIPLY)) & "," & _
                        "" & ChangeQuote(Trim(XLS_UNIT_LIMIT)) & "" & _
                     ") "

   Else 'Show edit
      SQL1 = "UPDATE XLS_UNIT SET " & _
                     "   XLS_UNIT_NAME = '" & ChangeQuote(Trim(XLS_UNIT_NAME)) & "', " & _
                     "   XLS_UNIT_MULTIPLY = " & ChangeQuote(Trim(XLS_UNIT_MULTIPLY)) & ", " & _
                     "   XLS_UNIT_LIMIT = " & ChangeQuote(Trim(XLS_UNIT_LIMIT)) & " " & _
                     "WHERE XLS_UNIT_ID = " & XLS_UNIT_ID
   End If
   
   If (AddEditMode = SHOW_ADD) Or (AddEditMode = SHOW_EDIT) Then
      m_Conn.Execute (SQL1)
   End If
End Function

Public Function DeleteData()
Dim SQL1 As String
      SQL1 = "DELETE FROM XLS_UNIT WHERE XLS_UNIT_ID = " & XLS_UNIT_ID
      m_Conn.Execute (SQL1)
End Function
