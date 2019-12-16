VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CProvinceMap"
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

Public KEY_ID As Long
Public KEY_SEARCH As String
Public KEY_MAP As String

Private Sub Class_Initialize()
   Set m_Conn = glbDatabaseMngr.DBLegacyConnection
End Sub

'In this class we will not handle any error or any transaction
Public Function QueryData(Rs As ADODB.Recordset, ItemCount As Long, Optional OwnerEnterprise As Boolean = True)
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

   SelectStr1 = "SELECT UG.*  "
   SelectStr2 = "SELECT COUNT(*)  "
   
   FromStr = "FROM PROVINCE_MAP UG "
   
   WhereStr = ""
   
   
   If KEY_ID > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (KEY_ID = " & KEY_ID & ") "
      Else
         WhereStr = WhereStr & "AND (KEY_ID = " & KEY_ID & ") "
      End If
   End If
   
   If Trim(KEY_SEARCH) <> "" Then
      If WildCard(KEY_SEARCH, SubLen, NewStr) Then
         If WhereStr = "" Then
            WhereStr = "WHERE (UG.KEY_SEARCH LIKE '" & ChangeQuote(Trim(KEY_SEARCH)) & "')"
         Else
            WhereStr = WhereStr & "AND (UG.KEY_SEARCH LIKE '" & ChangeQuote(Trim(KEY_SEARCH)) & "')"
         End If
      Else
         If WhereStr = "" Then
            WhereStr = "WHERE (UG.KEY_SEARCH = '" & ChangeQuote(Trim(KEY_SEARCH)) & "')"
         Else
            WhereStr = WhereStr & "AND (UG.KEY_SEARCH = '" & ChangeQuote(Trim(KEY_SEARCH)) & "')"
         End If
      End If
   End If
   
   
   OrderByStr = ""
   OrderTypeStr = ""

   If OrderBy = 1 Then
      OrderByStr = "ORDER BY KEY_ID "
   ElseIf OrderBy = 2 Then
      OrderByStr = "ORDER BY KEY_SEARCH "
   Else
      OrderByStr = "ORDER BY KEY_ID "
   End If
   
   If OrderType = 1 Then
      OrderTypeStr = "ASC "
   ElseIf OrderType = 2 Then
      OrderTypeStr = "DESC "
   Else
      OrderTypeStr = "DESC "
   End If
   
   SQL1 = SelectStr1 & FromStr & WhereStr & OrderByStr & OrderTypeStr
   SQL2 = SelectStr2 & FromStr & WhereStr
   
   Call glbDatabaseMngr.GetLegacyRs(SQL1, SQL2, True, ItemCount, Rs, glbErrorLog)
End Function

Public Function PopulateFromRS(Ind As Long, Rs As ADODB.Recordset)
   KEY_ID = NVLI(Rs("KEY_ID"), -1)
   KEY_SEARCH = NVLS(Rs("KEY_SEARCH"), "")
   KEY_MAP = NVLS(Rs("KEY_MAP"), "")
End Function
Public Function AddEditData()
Dim ServerDateTime As String
Dim SQL1 As String
Dim ID As Long
   
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   
   If AddEditMode = SHOW_ADD Then
      SQL1 = "INSERT INTO PROVINCE_MAP " & _
                     "( " & _
                     "   KEY_SEARCH, " & _
                     "   KEY_MAP " & _
                     ") " & _
                     "VALUES " & _
                     "( " & _
                        "'" & ChangeQuote(Trim(KEY_SEARCH)) & "'," & _
                         "'" & ChangeQuote(Trim(KEY_MAP)) & "'" & _
                     ") "
         KEY_ID = ID
   Else 'Show edit
      SQL1 = "UPDATE PROVINCE_MAP SET " & _
                     "   KEY_SEARCH = '" & ChangeQuote(Trim(KEY_SEARCH)) & "', " & _
                     "   KEY_MAP = '" & ChangeQuote(Trim(KEY_MAP)) & "' " & _
                     "WHERE KEY_ID = " & KEY_ID
   End If
   
   If (AddEditMode = SHOW_ADD) Or (AddEditMode = SHOW_EDIT) Then
      m_Conn.Execute (SQL1)
   End If
End Function

Public Function DeleteData()
Dim SQL1 As String
      SQL1 = "DELETE FROM PROVINCE_MAP WHERE KEY_ID = " & KEY_ID
      m_Conn.Execute (SQL1)
End Function
