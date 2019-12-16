VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CCommissionCustomerArea"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
' ฐานข้อมูล access
Private m_Conn As ADODB.Connection

Public AddEditMode As SHOW_MODE_TYPE
Public QueryFlag As Long
Public OrderBy As Long
Public OrderType As Long

' End of generic data member

Public COMMISSION_CUS_NAME As String
Public COMMISSION_CUS_ID As String
Public MASTER_AREA_ID As Long
Public COMMISSION_CUS_AREA_ID As Long
Public MASTER_AREA_NAME As String

Public FROM_DATE As Date
Public TO_DATE As Date
Public YEARNUM As String

Public ImportExportItems As Collection
Public YEAR_ID As Long
'Public GOODS_GROUP_ID As Long
'Public GOODS_MASTER_ID As Long

Public NONAREA As String

Private Sub Class_Initialize()
   Set m_Conn = glbDatabaseMngr.DBLegacyConnection
         Set ImportExportItems = New Collection
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

   If Ind = 1 Then            'สำหรับ เขตไหน มีข้อมูลอะไร Tab1
      SelectStr1 = "SELECT  DISTINCT CS.MASTER_AREA_ID, CSA.MASTER_AREA_NAME ,CSY.YEARNUM ,CSY.FROM_DATE, CSY.TO_DATE   "             '**********
      SelectStr2 = "SELECT COUNT(*)  "
      
      FromStr = "FROM (( COMMISSION_CUS_AREA CS "
      FromStr = FromStr & "LEFT OUTER JOIN AREA_YEAR CSY ON (CS.YEAR_ID = CSY.YEAR_ID)) "
      FromStr = FromStr & "LEFT OUTER JOIN COMMISSION_MASTER_AREA CSA ON (CS.MASTER_AREA_ID = CSA.MASTER_AREA_ID)) "

      WhereStr = "WHERE (CS.MASTER_AREA_ID > 0) "
  ElseIf Ind = 2 Then            ' สำหรับแต่ละบรรทัดของ Grid
     SelectStr1 = "SELECT CS.* "
      SelectStr2 = "SELECT COUNT(*)  "
      
      FromStr = "FROM COMMISSION_CUS_AREA CS "
   
   ElseIf Ind = 3 Then            'สำหรับ เขตไหน มีข้อมูลอะไร Tab1
      SelectStr1 = "SELECT CS.* , CSA.MASTER_AREA_NAME , CSY.YEARNUM ,CSY.FROM_DATE, CSY.TO_DATE   "             '**********
      SelectStr2 = "SELECT COUNT(*)  "

      FromStr = "FROM ( COMMISSION_CUS_AREA CS "
      FromStr = FromStr & "LEFT OUTER JOIN AREA_YEAR CSY ON (CS.YEAR_ID = CSY.YEAR_ID)) "
      FromStr = FromStr & "LEFT OUTER JOIN COMMISSION_MASTER_AREA CSA ON (CS.MASTER_AREA_ID = CSA.MASTER_AREA_ID) "
   ElseIf Ind = 4 Then            'สำหรับ เขตไหน มีข้อมูลอะไร Tab1
      SelectStr1 = "SELECT  CS.*, CSA.MASTER_AREA_NAME "             '**********
      SelectStr2 = "SELECT COUNT(*)  "
      
      FromStr = "FROM COMMISSION_CUS_AREA CS "
      FromStr = FromStr & "LEFT OUTER JOIN COMMISSION_MASTER_AREA CSA ON (CS.MASTER_AREA_ID = CSA.MASTER_AREA_ID) "

      WhereStr = "WHERE (CS.MASTER_AREA_ID > 0) "
   End If
      
If MASTER_AREA_ID > 0 Then
          If WhereStr = "" Then
            WhereStr = "WHERE (CS.MASTER_AREA_ID=" & ChangeQuote(Trim(MASTER_AREA_ID)) & ") "
         Else
            WhereStr = WhereStr & "AND (CS.MASTER_AREA_ID=" & ChangeQuote(Trim(MASTER_AREA_ID)) & ") "
         End If
End If
         
         
      If YEAR_ID > 0 Then
          If WhereStr = "" Then
            WhereStr = "WHERE (CS.YEAR_ID =" & ChangeQuote(Trim(YEAR_ID)) & ") "
         Else
            WhereStr = WhereStr & "AND (CS.YEAR_ID=" & ChangeQuote(Trim(YEAR_ID)) & ") "
         End If
      End If
         
    If COMMISSION_CUS_AREA_ID > 0 Then
      If WhereStr = "" Then
         WhereStr = "WHERE (CS.COMMISSION_CUS_AREA_ID = " & COMMISSION_CUS_AREA_ID & ") "
      Else
         WhereStr = WhereStr & "AND (CS.COMMISSION_CUS_AREA_ID = " & COMMISSION_CUS_AREA_ID & ") "
      End If
   End If
   
   
    If NONAREA = "Y" Then
      If WhereStr = "" Then
         WhereStr = "WHERE (CS.COMMISSION_CUS_AREA_ID = NULL ) "
      Else
         WhereStr = WhereStr & "AND (CS.COMMISSION_CUS_AREA_ID = NULL) "
      End If
   End If
   
   OrderByStr = ""
   OrderTypeStr = ""

    If Ind = 1 Then
         OrderByStr = "ORDER BY CS.MASTER_AREA_ID "
    ElseIf Ind = 2 Then
        OrderByStr = "ORDER BY CS.COMMISSION_CUS_ID "
    ElseIf Ind = 3 Then
        OrderByStr = "ORDER BY CS.COMMISSION_CUS_ID "
    ElseIf Ind = 4 Then
        OrderByStr = "ORDER BY CS.COMMISSION_CUS_ID "
    End If
         
         OrderTypeStr = "ASC "
   
   SQL1 = SelectStr1 & FromStr & WhereStr & OrderByStr & OrderTypeStr
   SQL2 = SelectStr2 & FromStr & WhereStr
'  'debug.print SQL1
  Call glbDatabaseMngr.GetLegacyRs(SQL1, SQL2, True, ItemCount, Rs, glbErrorLog)
End Function

Public Function PopulateFromRS(Ind As Long, Rs As ADODB.Recordset)
If Ind = 1 Then
'        COMMISSION_CUS_AREA_ID = NVLI(Rs("COMMISSION_CUS_AREA_ID"), -1)
        MASTER_AREA_NAME = NVLS(Rs("MASTER_AREA_NAME"), "")
        MASTER_AREA_ID = NVLI(Rs("MASTER_AREA_ID"), -1)

         YEARNUM = NVLS(Rs("YEARNUM"), "")
         FROM_DATE = InternalDateToDate(NVLS(Rs("FROM_DATE"), ""))
         TO_DATE = InternalDateToDate(NVLS(Rs("TO_DATE"), ""))
         
'         GOODS_MASTER_ID = NVLI(Rs("GOODS_MASTER_ID"), -1)
'         GOODS_GROUP_ID = NVLI(Rs("GOODS_GROUP_ID"), -1)
ElseIf Ind = 2 Then
        COMMISSION_CUS_AREA_ID = NVLI(Rs("COMMISSION_CUS_AREA_ID"), -1)
        MASTER_AREA_ID = NVLI(Rs("MASTER_AREA_ID"), -1)
        YEAR_ID = NVLI(Rs("YEAR_ID"), -1)
        COMMISSION_CUS_NAME = NVLS(Rs("COMMISSION_CUS_NAME"), "")
        COMMISSION_CUS_ID = NVLS(Rs("COMMISSION_CUS_ID"), "")
ElseIf Ind = 3 Then
        COMMISSION_CUS_AREA_ID = NVLI(Rs("COMMISSION_CUS_AREA_ID"), -1)
        MASTER_AREA_NAME = NVLS(Rs("MASTER_AREA_NAME"), "")
        MASTER_AREA_ID = NVLI(Rs("MASTER_AREA_ID"), -1)

        COMMISSION_CUS_NAME = NVLS(Rs("COMMISSION_CUS_NAME"), "")
        COMMISSION_CUS_ID = NVLS(Rs("COMMISSION_CUS_ID"), "")
         YEAR_ID = NVLI(Rs("YEAR_ID"), -1)
         YEARNUM = NVLS(Rs("YEARNUM"), "")
         FROM_DATE = InternalDateToDate(NVLS(Rs("FROM_DATE"), ""))
         TO_DATE = InternalDateToDate(NVLS(Rs("TO_DATE"), ""))
ElseIf Ind = 4 Then
        COMMISSION_CUS_AREA_ID = NVLI(Rs("COMMISSION_CUS_AREA_ID"), -1)
        MASTER_AREA_ID = NVLI(Rs("MASTER_AREA_ID"), -1)
        YEAR_ID = NVLI(Rs("YEAR_ID"), -1)
        COMMISSION_CUS_NAME = NVLS(Rs("COMMISSION_CUS_NAME"), "")
        COMMISSION_CUS_ID = NVLS(Rs("COMMISSION_CUS_ID"), "")
        MASTER_AREA_NAME = NVLS(Rs("MASTER_AREA_NAME"), "")
End If
End Function
Public Function AddEditData()
Dim ServerDateTime As String
Dim SQL1 As String

   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   
   If AddEditMode = SHOW_ADD Then

               SQL1 = "INSERT INTO COMMISSION_CUS_AREA " & _
                              "( " & _
                              "   MASTER_AREA_ID, " & _
                              "   YEAR_ID, " & _
                              "   COMMISSION_CUS_NAME, " & _
                              "   COMMISSION_CUS_ID " & _
                              ") " & _
                              "VALUES " & _
                              "( " & _
                                 "" & ChangeQuote(Trim(MASTER_AREA_ID)) & "," & _
                                 "" & ChangeQuote(Trim(YEAR_ID)) & "," & _
                                 "'" & ChangeQuote(Trim(COMMISSION_CUS_NAME)) & "'," & _
                                 "'" & ChangeQuote(Trim(COMMISSION_CUS_ID)) & "'" & _
                              ") "
   Else 'Show edit

               SQL1 = "UPDATE COMMISSION_CUS_AREA SET " & _
                           "   MASTER_AREA_ID = " & ChangeQuote(Trim(MASTER_AREA_ID)) & ", " & _
                           "   YEAR_ID = " & ChangeQuote(Trim(YEAR_ID)) & ", " & _
                           "   COMMISSION_CUS_NAME = '" & ChangeQuote(Trim(COMMISSION_CUS_NAME)) & "', " & _
                           "   COMMISSION_CUS_ID = '" & ChangeQuote(Trim(COMMISSION_CUS_ID)) & "' " & _
                           " WHERE COMMISSION_CUS_AREA_ID = " & COMMISSION_CUS_AREA_ID

   End If
   
   If (AddEditMode = SHOW_ADD) Or (AddEditMode = SHOW_EDIT) Then
      'debug.print SQL1
      m_Conn.Execute (SQL1)
   End If
End Function

Public Function DeleteData(Optional Ind As Long = 1)
Dim SQL1 As String

If Ind = 1 Then
      SQL1 = "DELETE FROM COMMISSION_CUS_AREA WHERE COMMISSION_CUS_AREA_ID = " & COMMISSION_CUS_AREA_ID
ElseIf Ind = 2 Then
      SQL1 = "DELETE FROM COMMISSION_CUS_AREA WHERE ( MASTER_AREA_ID = " & MASTER_AREA_ID & ") AND  ( YEAR_ID = " & YEAR_ID & ")"
End If
      m_Conn.Execute (SQL1)
End Function

Private Sub Class_Terminate()
   Set ImportExportItems = Nothing
End Sub
