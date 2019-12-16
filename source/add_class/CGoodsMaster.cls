VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CGoodsMaster"
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
' End of generic data member

Public GOODS_MASTER_ID As Long
Public GOODS_MASTER_CODE As String
Public GOODS_MASTER_NAME As String

Public ORDER_BY As Long
Public ORDER_TYPE  As Long

Private Sub Class_Initialize()
   Set m_Conn = glbDatabaseMngr.DBLegacyConnection
'   FROM_DATE = -1
'   TO_DATE = -1
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
      SelectStr1 = "SELECT GM.*  "
      SelectStr2 = "SELECT COUNT(*)  "
      
      FromStr = "FROM GOODS_MASTER GM "
'  ElseIf Ind = 2 Then
'      SelectStr1 = "SELECT MAX(CSY.YEAR_ID) AS YEAR_ID  "
'      SelectStr2 = "SELECT COUNT(*)  "
'
'      FromStr = "FROM AREA_YEAR CSY "
  End If

   
         If Trim(GOODS_MASTER_CODE) <> "" Then
            If WhereStr = "" Then
               WhereStr = "WHERE (GM.GOODS_MASTER_CODE= '" & GOODS_MASTER_CODE & "') "
            Else
               WhereStr = WhereStr & "AND (GM.GOODS_MASTER_CODE = '" & GOODS_MASTER_CODE & "') "
            End If
         End If
         
         If Trim(GOODS_MASTER_NAME) <> "" Then
            If WhereStr = "" Then
               WhereStr = "WHERE (GM.GOODS_MASTER_NAME= '" & GOODS_MASTER_NAME & "') "
            Else
               WhereStr = WhereStr & "AND (GM.GOODS_MASTER_NAME = '" & GOODS_MASTER_NAME & "') "
            End If
         End If
         
   
    If Ind = 1 Then
         If OrderBy = 1 Then
             OrderByStr = "ORDER BY GOODS_MASTER_CODE  "
          ElseIf OrderBy = 2 Then
             OrderByStr = "ORDER BY GOODS_MASTER_NAME "
          Else
             OrderByStr = "ORDER BY GOODS_MASTER_ID "
          End If
   End If
               
    If Ind = 1 Then
         If OrderType = 2 Then
             OrderTypeStr = "DESC "
         Else
           OrderTypeStr = "ASC "
         End If
    End If
   
   SQL1 = SelectStr1 & FromStr & WhereStr & OrderByStr & OrderTypeStr
   SQL2 = SelectStr2 & FromStr & WhereStr
    'debug.print SQL1
   Call glbDatabaseMngr.GetLegacyRs(SQL1, SQL2, True, ItemCount, Rs, glbErrorLog)
End Function

Public Function PopulateFromRS(Ind As Long, Rs As ADODB.Recordset)
  If Ind = 1 Then
      GOODS_MASTER_ID = NVLI(Rs("GOODS_MASTER_ID"), "-1")
      GOODS_MASTER_CODE = NVLS(Rs("GOODS_MASTER_CODE"), "")
      GOODS_MASTER_NAME = NVLS(Rs("GOODS_MASTER_NAME"), "")
'   ElseIf Ind = 2 Then
'      YEAR_ID = NVLI(Rs("YEAR_ID"), "-1")
   End If

End Function
Public Function AddEditData()
Dim ServerDateTime As String
Dim SQL1 As String

   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   
'   ENTERPRISE_CODE = glbCompanyCode
   
   If AddEditMode = SHOW_ADD Then
             SQL1 = "INSERT INTO GOODS_MASTER " & _
                              "( " & _
                              "   GOODS_MASTER_CODE, " & _
                              "   GOODS_MASTER_NAME " & _
                              ") " & _
                              "VALUES " & _
                              "( " & _
                               "'" & ChangeQuote(Trim(GOODS_MASTER_CODE)) & "'," & _
                               "'" & ChangeQuote(Trim(GOODS_MASTER_NAME)) & "' " & _
                              ") "
   Else 'Show edit
               SQL1 = "UPDATE GOODS_MASTER SET " & _
                           "   GOODS_MASTER_CODE = '" & ChangeQuote(Trim(GOODS_MASTER_CODE)) & "'," & _
                           "   GOODS_MASTER_NAME = '" & ChangeQuote(Trim(GOODS_MASTER_NAME)) & "'" & _
                           "   WHERE GOODS_MASTER_ID = " & GOODS_MASTER_ID
   End If
   
   If (AddEditMode = SHOW_ADD) Or (AddEditMode = SHOW_EDIT) Then
      m_Conn.Execute (SQL1)
   End If
End Function

Public Function DeleteData()
Dim SQL1 As String
      SQL1 = "DELETE FROM GOODS_MASTER WHERE GOODS_MASTER_ID = " & GOODS_MASTER_ID
      m_Conn.Execute (SQL1)
      SQL1 = "DELETE FROM GOODS_DETAIL WHERE GOODS_MASTER_ID = " & GOODS_MASTER_ID
      m_Conn.Execute (SQL1)
End Function




