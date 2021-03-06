VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CCommissMasterArea"
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

Public MASTER_AREA_ID As Long
Public MASTER_AREA_NAME As String
Public COMMISSION_CUS_NAME As String
Public COMMISSION_CUS_ID As String
Public COMMISSION_CUS_AREA_ID As Long
Public Flag As String

Public ImportExportItems As Collection

Public ORDER_BY As Long
Public ORDER_TYPE  As Long

'Public KEY_ID As Long
'Public KEY_LOOKUP As String
Private Sub Class_Initialize()
   Set m_Conn = glbDatabaseMngr.DBLegacyConnection
      Set ImportExportItems = New Collection
End Sub
'In this class we will not handle any error or any transaction
Public Function QueryData(Optional Ind As Long = 1, Optional Rs As ADODB.Recordset, Optional ItemCount As Long)
Dim SQL1 As String
Dim SQL2 As String
Dim SelectStr1 As String
Dim SelectStr2 As String
Dim FromStr As String
Dim WhereStr As String
Dim OrderByStr As String
Dim OrderTypeStr As String
   
   If Ind = 1 Then
      SelectStr1 = "SELECT CMA.*  "
      SelectStr2 = "SELECT COUNT(*)  "
      
      FromStr = "FROM COMMISSION_MASTER_AREA CMA "
      
 ElseIf Ind = 2 Then
            SelectStr1 = "SELECT CS.* , CSA.MASTER_AREA_NAME "             '**********
      SelectStr2 = "SELECT COUNT(*)  "
      
      FromStr = "FROM COMMISSION_CUS_AREA CS "
      FromStr = FromStr & "LEFT OUTER JOIN COMMISSION_MASTER_AREA CSA ON (CS.MASTER_AREA_ID = CSA.MASTER_AREA_ID) "
End If
         
   If MASTER_AREA_ID <> 0 Then
     WhereStr = "WHERE (MASTER_AREA_ID=" & ChangeQuote(Trim(MASTER_AREA_ID)) & ") "
   End If
   ' ���͵ç������ elseif �� ?
'   If Trim(MASTER_AREA_NAME) <> "" Then
'      If WhereStr = "" Then
'         WhereStr = "WHERE (CSY.MASTER_AREA_NAME = '" & MASTER_AREA_NAME & "') "
'      Else
'         WhereStr = WhereStr & "AND (CSY.MASTER_AREA_NAME = '" & MASTER_AREA_NAME & "') "
'      End If
'   End If
   If Ind = 1 Then
      OrderByStr = "ORDER BY MASTER_AREA_ID "
  ElseIf Ind = 2 Then
      OrderByStr = "ORDER BY CSA.MASTER_AREA_ID ,CS.COMMISSION_CUS_ID "
   End If
      OrderTypeStr = "ASC "

   
   SQL1 = SelectStr1 & FromStr & WhereStr & OrderByStr & OrderTypeStr
   SQL2 = SelectStr2 & FromStr & WhereStr
   
   Call glbDatabaseMngr.GetLegacyRs(SQL1, SQL2, True, ItemCount, Rs, glbErrorLog)
End Function

Public Function PopulateFromRS(Ind As Long, Rs As ADODB.Recordset)
If Ind = 1 Then
        MASTER_AREA_ID = NVLI(Rs("MASTER_AREA_ID"), -1)
         MASTER_AREA_NAME = NVLS(Rs("MASTER_AREA_NAME"), "")

         Flag = NVLS(Rs("FLAG"), -1)
ElseIf Ind = 2 Then
              COMMISSION_CUS_AREA_ID = NVLI(Rs("COMMISSION_CUS_AREA_ID"), -1)
        MASTER_AREA_ID = NVLI(Rs("MASTER_AREA_ID"), -1)
        
        MASTER_AREA_NAME = NVLS(Rs("MASTER_AREA_NAME"), "")
        COMMISSION_CUS_NAME = NVLS(Rs("COMMISSION_CUS_NAME"), "")
        COMMISSION_CUS_ID = NVLS(Rs("COMMISSION_CUS_ID"), "")
   End If
End Function
Public Function AddEditData()
Dim ServerDateTime As String
Dim SQL1 As String

   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   
'   ENTERPRISE_CODE = glbCompanyCode
   
   If AddEditMode = SHOW_ADD Then
             SQL1 = "INSERT INTO COMMISSION_MASTER_AREA " & _
                              "( " & _
                              "   MASTER_AREA_NAME, " & _
                              "   FLAG " & _
                              ") " & _
                              "VALUES " & _
                              "( " & _
                                 "'" & ChangeQuote(Trim(MASTER_AREA_NAME)) & "'," & _
                                 "'Y'" & _
                              ") "
   Else 'Show edit
         
               SQL1 = "UPDATE COMMISSION_MASTER_AREA SET " & _
                           "   MASTER_AREA_NAME = '" & ChangeQuote(Trim(MASTER_AREA_NAME)) & "'," & _
                           "   FLAG = '" & ChangeQuote(Trim(Flag)) & "'" & _
                           "   WHERE MASTER_AREA_ID= " & MASTER_AREA_ID
   End If
   
   If (AddEditMode = SHOW_ADD) Or (AddEditMode = SHOW_EDIT) Then
      m_Conn.Execute (SQL1)
   End If
End Function

Public Function DeleteData()
Dim SQL1 As String
      SQL1 = "DELETE FROM COMMISSION_MASTER_AREA WHERE MASTER_AREA_ID= " & MASTER_AREA_ID
      m_Conn.Execute (SQL1)
      
      SQL1 = "DELETE FROM COMMISSION_CUS_AREA WHERE MASTER_AREA_ID= " & MASTER_AREA_ID
      m_Conn.Execute (SQL1)
End Function

Private Sub Class_Terminate()
   Set ImportExportItems = Nothing
End Sub


