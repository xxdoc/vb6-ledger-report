VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CDataType"
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

Public DATA_TYPE_ID As Long
Public DATA_TYPE_NAME As String

Public ORDER_BY As Long
Public ORDER_TYPE  As Long
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
   
   SelectStr1 = "SELECT UG.*  "
   SelectStr2 = "SELECT COUNT(*)  "
   
   FromStr = "FROM DATA_TYPE UG "
   
   WhereStr = ""
      
   If DATA_TYPE_ID > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (DATA_TYPE_ID= " & DATA_TYPE_ID & ") "
      Else
         WhereStr = WhereStr & "AND (DATA_TYPE_ID = " & DATA_TYPE_ID & ") "
      End If
   End If
   
   OrderByStr = ""
   OrderTypeStr = ""
   
   OrderByStr = "ORDER BY DATA_TYPE_ID "
   
   If OrderType = 1 Then
      OrderTypeStr = "ASC "
   ElseIf OrderType = 2 Then
      OrderTypeStr = "DESC "
   Else
      OrderTypeStr = "ASC "
   End If
   
   SQL1 = SelectStr1 & FromStr & WhereStr & OrderByStr & OrderTypeStr
   SQL2 = SelectStr2 & FromStr & WhereStr
   
   Call glbDatabaseMngr.GetLegacyRs(SQL1, SQL2, True, ItemCount, Rs, glbErrorLog)
End Function

Public Function PopulateFromRS(Ind As Long, Rs As ADODB.Recordset)
   DATA_TYPE_ID = NVLI(Rs("DATA_TYPE_ID"), -1)
   DATA_TYPE_NAME = NVLS(Rs("DATA_TYPE_NAME"), "")
End Function
Public Function AddEditData()
Dim ServerDateTime As String
Dim SQL1 As String

   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   
   If AddEditMode = SHOW_ADD Then
      SQL1 = "INSERT INTO DATA_TYPE " & _
                     "( " & _
                     "   DATA_TYPE_NAME " & _
                     ") " & _
                     "VALUES " & _
                     "( " & _
                        "'" & ChangeQuote(Trim(DATA_TYPE_NAME)) & "'" & _
                     ") "
   Else 'Show edit
      SQL1 = "UPDATE DATA_TYPE SET " & _
                     "   DATA_TYPE_NAME = '" & ChangeQuote(Trim(DATA_TYPE_NAME)) & "' " & _
                     "WHERE DATA_TYPE_ID = " & DATA_TYPE_ID
   End If
   
   If (AddEditMode = SHOW_ADD) Or (AddEditMode = SHOW_EDIT) Then
      m_Conn.Execute (SQL1)
   End If
End Function
Public Function DeleteData()
Dim SQL1 As String
      SQL1 = "DELETE FROM DATA_TYPE WHERE DATA_TYPE_ID = " & DATA_TYPE_ID
      m_Conn.Execute (SQL1)
End Function
