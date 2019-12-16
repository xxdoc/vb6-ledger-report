VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CComboSubGroup"
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

Public COMBO_SUB_ID As Long
Public COMBO_SUB_NAME As String

Public ORDER_BY As Long
Public ORDER_TYPE  As Long

Public ItemAmountColl As Collection

Private Sub Class_Initialize()
   Set m_Conn = glbDatabaseMngr.DBLegacyConnection
   
   Set ItemAmountColl = New Collection
End Sub
'In this class we will not handle any error or any transaction
Public Function QueryData(Optional Ind As Long = 1, Optional Rs As ADODB.Recordset, Optional ItemCount As Long, Optional OwnerEnterprise As Boolean = True)
Dim SQL1 As String
Dim SQL2 As String
Dim SelectStr1 As String
Dim SelectStr2 As String
Dim FromStr As String
Dim WhereStr As String
Dim OrderByStr As String
Dim OrderTypeStr As String
Dim GroupByStr As String

   If Ind = 1 Then
      SelectStr1 = "SELECT UG.*  "
      SelectStr2 = "SELECT COUNT(*)  "
      
      FromStr = "FROM COMBO_SUB_GROUP UG "
   ElseIf Ind = 2 Then
      SelectStr1 = "SELECT MAX(UG.COMBO_SUB_ID) AS COMBO_SUB_ID "
      SelectStr2 = "SELECT COUNT(*)  "
      
      FromStr = "FROM COMBO_SUB_GROUP UG "
   End If
   
   WhereStr = ""
      
   If COMBO_SUB_ID > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (UG.COMBO_SUB_ID= " & COMBO_SUB_ID & ") "
      Else
         WhereStr = WhereStr & "AND (UG.COMBO_SUB_ID = " & COMBO_SUB_ID & ") "
      End If
   End If


   OrderByStr = ""
   OrderTypeStr = ""
   
   If Ind = 1 Then
      OrderByStr = "ORDER BY UG.COMBO_SUB_NAME "
      If OrderType = 1 Then
         OrderTypeStr = "ASC "
      ElseIf OrderType = 2 Then
         OrderTypeStr = "DESC "
      Else
         OrderTypeStr = "ASC "
      End If
   End If
   
   SQL1 = SelectStr1 & FromStr & WhereStr & GroupByStr & OrderByStr & OrderTypeStr
   SQL2 = SelectStr2 & FromStr & WhereStr
   
   Call glbDatabaseMngr.GetLegacyRs(SQL1, SQL2, True, ItemCount, Rs, glbErrorLog)
End Function

Public Function PopulateFromRS(Ind As Long, Rs As ADODB.Recordset)
If Ind = 1 Then
   COMBO_SUB_ID = NVLI(Rs("COMBO_SUB_ID"), -1)
   COMBO_SUB_NAME = NVLS(Rs("COMBO_SUB_NAME"), "")
ElseIf Ind = 2 Then
    COMBO_SUB_ID = NVLI(Rs("COMBO_SUB_ID"), -1)
  ' COMBO_SUB_NAME = NVLS(Rs("COMBO_SUB_NAME"), "")
End If
End Function
Public Function AddEditData()
Dim ServerDateTime As String
Dim SQL1 As String

   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   
   If AddEditMode = show_Add Then
      SQL1 = "INSERT INTO COMBO_SUB_GROUP " & _
                     "( " & _
                     "   COMBO_SUB_NAME " & _
                     ") " & _
                     "VALUES " & _
                     "( " & _
                        "'" & ChangeQuote(Trim(COMBO_SUB_NAME)) & "'" & _
                     ") "
   Else 'Show edit
      SQL1 = "UPDATE COMBO_SUB_GROUP SET " & _
                     "   COMBO_SUB_NAME = '" & ChangeQuote(Trim(COMBO_SUB_NAME)) & "' " & _
                     "WHERE COMBO_SUB_ID = " & COMBO_SUB_ID
   End If
   
   If (AddEditMode = show_Add) Or (AddEditMode = SHOW_EDIT) Then
      m_Conn.Execute (SQL1)
   End If
End Function
Public Function DeleteData()
Dim SQL1 As String
      SQL1 = "DELETE FROM COMBO_SUB_GROUP WHERE COMBO_SUB_ID = " & COMBO_SUB_ID
      m_Conn.Execute (SQL1)
      
      SQL1 = "DELETE FROM COMBO_SUB_GROUP_DETAIL WHERE COMBO_SUB_ID = " & COMBO_SUB_ID
      m_Conn.Execute (SQL1)
End Function
Private Sub Class_Terminate()
   Set ItemAmountColl = Nothing
End Sub
