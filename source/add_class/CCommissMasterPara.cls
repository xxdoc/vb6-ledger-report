VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CCommissMasterPara"
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

Public MASTER_PARAMETER_ID As Long
Public MASTER_PARAMETER_NAME As String
Public MASTER_PARAMETER_VALUE As Double
'Public COMMISSION_CUS_NAME As String
'Public COMMISSION_CUS_ID As String
Public VALID_FROM As Date
Public VALID_TO As Date
Public MASTER_FROMTO_ID As Long
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
Dim GroupByStr As String
Dim TempStr As String
Dim NewStr As String
Dim SubLen As Long
Dim TempDate As String


   If Ind = 1 Then
      SelectStr1 = "SELECT NM.*  "
      SelectStr2 = "SELECT COUNT(*)  "
      
      FromStr = "FROM COMMISSION_MASTER_PARAMETER NM "
      
   ElseIf Ind = 2 Then
       SelectStr1 = "SELECT NM.* , MFT.* "             '**********
      SelectStr2 = "SELECT COUNT(*)  "

      FromStr = "FROM COMMISSION_CUS_PARAMETER NM "
      FromStr = FromStr & "LEFT OUTER JOIN MASTER_FROMTO2 MFT ON (NM.MASTER_FROMTO_ID = MFT.MASTER_FROMTO_ID ) "
      
      TempStr = "WHERE "
      TempDate = DateToStringIntLow(VALID_FROM)
      WhereStr = WhereStr & TempStr & " (MFT.VALID_FROM <= '" & ChangeQuote(Trim(TempDate)) & "')"
      TempStr = "AND "
      TempDate = DateToStringIntLow(VALID_TO)
      WhereStr = WhereStr & TempStr & " (MFT.VALID_TO >= '" & ChangeQuote(Trim(TempDate)) & "')"

   End If
 
    If MASTER_PARAMETER_ID > 0 Then
      If Len(WhereStr) > 0 Then
         TempStr = "AND "
      Else
         TempStr = "WHERE "
      End If
      WhereStr = WhereStr & TempStr & " (NM.MASTER_PARAMETER_ID = " & MASTER_PARAMETER_ID & ") " & vbCrLf
   End If
   
   If MASTER_FROMTO_ID > 0 Then
      If Len(WhereStr) > 0 Then
         TempStr = "AND "
      Else
         TempStr = "WHERE "
      End If
      WhereStr = WhereStr & TempStr & " (NM.MASTER_FROMTO_ID = " & MASTER_FROMTO_ID & ") " & vbCrLf
   End If
 
 
   If Ind = 1 Then
      OrderByStr = "ORDER BY MASTER_PARAMETER_ID "
   ElseIf Ind = 2 Then
      OrderByStr = "ORDER BY VALID_FROM "
   End If
   
      OrderTypeStr = "ASC "

   
   SQL1 = SelectStr1 & FromStr & WhereStr & OrderByStr & OrderTypeStr
   SQL2 = SelectStr2 & FromStr & WhereStr
   
   Call glbDatabaseMngr.GetLegacyRs(SQL1, SQL2, True, ItemCount, Rs, glbErrorLog)
End Function

Public Function PopulateFromRS(Ind As Long, Rs As ADODB.Recordset)
If Ind = 1 Then
        MASTER_PARAMETER_ID = NVLI(Rs("MASTER_PARAMETER_ID"), -1)
         MASTER_PARAMETER_NAME = NVLS(Rs("MASTER_PARAMETER_NAME"), "")
         MASTER_PARAMETER_VALUE = NVLD(Rs("MASTER_PARAMETER_VALUE"), -1)
         MASTER_FROMTO_ID = NVLI(Rs("MASTER_FROMTO_ID"), -1)
         Flag = NVLS(Rs("FLAG"), -1)
ElseIf Ind = 2 Then
        MASTER_PARAMETER_ID = NVLI(Rs("MASTER_PARAMETER_ID"), -1)
         MASTER_PARAMETER_NAME = NVLS(Rs("MASTER_PARAMETER_NAME"), "")
         MASTER_PARAMETER_VALUE = NVLD(Rs("MASTER_PARAMETER_VALUE"), -1)
                  MASTER_FROMTO_ID = NVLI(Rs("MASTER_FROMTO_ID"), -1)
         Flag = NVLS(Rs("FLAG"), -1)
   End If


End Function
Public Function AddEditData()
Dim ServerDateTime As String
Dim SQL1 As String

   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)

'   ENTERPRISE_CODE = glbCompanyCode
' Call glbDatabaseMngr.GetSeqID ()
   
   If AddEditMode = SHOW_ADD Then
             SQL1 = "INSERT INTO COMMISSION_MASTER_PARAMETER " & _
                              "( " & _
                                "   MASTER_FROMTO_ID, " & _
                              "   MASTER_PARAMETER_NAME, " & _
                               "   MASTER_PARAMETER_VALUE, " & _
                              "   FLAG " & _
                              ") " & _
                              "VALUES " & _
                              "( " & _
                                 "" & ChangeQuote(Trim(MASTER_FROMTO_ID)) & "," & _
                                 "'" & ChangeQuote(Trim(MASTER_PARAMETER_NAME)) & "'," & _
                                 "" & ChangeQuote(Trim(MASTER_PARAMETER_VALUE)) & "," & _
                                 "'" & ChangeQuote(Trim(Flag)) & "'" & _
                              ") "
   Else 'Show edit
         
               SQL1 = "UPDATE COMMISSION_MASTER_PARAMETER SET " & _
                           "   MASTER_PARAMETER_NAME = '" & ChangeQuote(Trim(MASTER_PARAMETER_NAME)) & "'," & _
                           "   MASTER_PARAMETER_VALUE = " & ChangeQuote(Trim(MASTER_PARAMETER_VALUE)) & "," & _
                           "   FLAG = '" & ChangeQuote(Trim(Flag)) & "'" & _
                           "   WHERE MASTER_PARAMETER_ID= " & MASTER_PARAMETER_ID

   End If
   
   If (AddEditMode = SHOW_ADD) Or (AddEditMode = SHOW_EDIT) Then
      m_Conn.Execute (SQL1)
   End If
End Function

Public Function DeleteData()
Dim SQL1 As String
      SQL1 = "DELETE FROM COMMISSION_MASTER_PARAMETER WHERE MASTER_PARAMETER_ID= " & MASTER_PARAMETER_ID
      m_Conn.Execute (SQL1)
End Function

Private Sub Class_Terminate()
   Set ImportExportItems = Nothing
End Sub


