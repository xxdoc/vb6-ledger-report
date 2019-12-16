VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CXlsEstimateSetting"
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

Public XLS_EST_SET_ID As Long
Public SHEET_1 As Long
Public ROW_1 As Long
Public COLLUMN_1 As String
Public SHEET_2 As Long
Public ROW_2 As Long
Public COLLUMN_2 As String
Public SHEET_3 As Long
Public ROW_3 As Long
Public COLLUMN_3 As String
Public SHEET_4 As Long
Public ROW_4 As Long
Public COLLUMN_4 As String
Public ROW_6 As Long
Public COLLUMN_6 As String
Public ROW_7 As Long
Public COLLUMN_7 As String

Public FROMDATAROW As Long
Public TODATAROW As Long
Public COLLUMNFOOD As String

Public SHEET_5 As Long
Public ROW_5 As Long
Public COLLUMN5_1 As String
Public COLLUMN5_2 As String
Public COLLUMN5_3 As String
Public COLLUMN5_4 As String
Public COLLUMN5_5 As String
Public COLLUMN5_6 As String
Public COLLUMN5_7 As String
Public COLLUMN6_1 As String
Public COLLUMN6_2 As String
Public COLLUMN6_3 As String
Public COLLUMN6_4 As String
Public COLLUMN6_5 As String
Public COLLUMN6_6 As String
Public COLLUMN6_7 As String

Public SHEETBB As Long
Public ROWBB As Long

Public KEY_ID As Long
Public KEY_LOOKUP As String

Private Sub Class_Initialize()
   Set m_Conn = glbDatabaseMngr.DBLegacyConnection
End Sub

'In this class we will not handle any error or any transaction
Public Function QueryData(Rs As ADODB.Recordset, ItemCount As Long)
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
   
   FromStr = "FROM XLS_EST_SET UG "
   
   WhereStr = ""
   
   If Trim(XLS_EST_SET_ID) <> -1 Then
      If WhereStr = "" Then
         WhereStr = "WHERE UG.XLS_EST_SET_ID = " & Trim(XLS_EST_SET_ID)
      Else
         WhereStr = WhereStr & "AND UG.XLS_EST_SET_ID = " & Trim(XLS_EST_SET_ID)
      End If
   End If
   
   OrderByStr = ""
   OrderTypeStr = ""

   If OrderBy = 1 Then
      OrderByStr = " ORDER BY UG.XLS_EST_SET_ID "
   ElseIf OrderBy = 2 Then
      OrderByStr = " ORDER BY UG.XLS_EST_SET_ID "
   Else
      OrderByStr = " ORDER BY UG.XLS_EST_SET_ID "
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
   XLS_EST_SET_ID = NVLI(Rs("XLS_EST_SET_ID"), -1)
   SHEET_1 = NVLI(Rs("SHEET_1"), -1)
   ROW_1 = NVLI(Rs("ROW_1"), -1)
   COLLUMN_1 = NVLS(Rs("COLLUMN_1"), "")
   SHEET_2 = NVLI(Rs("SHEET_2"), -1)
   ROW_2 = NVLI(Rs("ROW_2"), -1)
   COLLUMN_2 = NVLS(Rs("COLLUMN_2"), "")
   SHEET_3 = NVLI(Rs("SHEET_3"), -1)
   ROW_3 = NVLI(Rs("ROW_3"), -1)
   COLLUMN_3 = NVLS(Rs("COLLUMN_3"), "")
   SHEET_4 = NVLI(Rs("SHEET_4"), -1)
   ROW_4 = NVLI(Rs("ROW_4"), -1)
   COLLUMN_4 = NVLS(Rs("COLLUMN_4"), "")
   ROW_6 = NVLI(Rs("ROW_6"), -1)
   COLLUMN_6 = NVLS(Rs("COLLUMN_6"), "")
   ROW_7 = NVLI(Rs("ROW_7"), -1)
   COLLUMN_7 = NVLS(Rs("COLLUMN_7"), "")
   
   TODATAROW = NVLI(Rs("TODATAROW"), -1)
   FROMDATAROW = NVLI(Rs("FROMDATAROW"), -1)
   
   SHEET_5 = NVLI(Rs("SHEET_5"), -1)
   ROW_5 = NVLI(Rs("ROW_5"), -1)
   COLLUMN5_1 = NVLS(Rs("COLLUMN5_1"), "")
   COLLUMN5_2 = NVLS(Rs("COLLUMN5_2"), "")
   COLLUMN5_3 = NVLS(Rs("COLLUMN5_3"), "")
   COLLUMN5_4 = NVLS(Rs("COLLUMN5_4"), "")
   COLLUMN5_5 = NVLS(Rs("COLLUMN5_5"), "")
   COLLUMN5_6 = NVLS(Rs("COLLUMN5_6"), "")
   COLLUMN5_7 = NVLS(Rs("COLLUMN5_7"), "")
   
   COLLUMN6_1 = NVLS(Rs("COLLUMN6_1"), "")
   COLLUMN6_2 = NVLS(Rs("COLLUMN6_2"), "")
   COLLUMN6_3 = NVLS(Rs("COLLUMN6_3"), "")
   COLLUMN6_4 = NVLS(Rs("COLLUMN6_4"), "")
   COLLUMN6_5 = NVLS(Rs("COLLUMN6_5"), "")
   COLLUMN6_6 = NVLS(Rs("COLLUMN6_6"), "")
   COLLUMN6_7 = NVLS(Rs("COLLUMN6_7"), "")
   
   COLLUMNFOOD = NVLS(Rs("COLLUMNFOOD"), "")
   SHEETBB = NVLI(Rs("SHEETBB"), -1)
   ROWBB = NVLI(Rs("ROWBB"), -1)
   
End Function
Public Function AddEditData()
Dim ServerDateTime As String
Dim SQL1 As String
Dim ID As Long

   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   
   If AddEditMode = SHOW_ADD Then
      SQL1 = "INSERT INTO XLS_EST_SET , " & _
                     "( " & _
                     "   SHEET_1 , ROW_1 , COLLUMN_1 , " & _
                     "   SHEET_2 , ROW_2 , COLLUMN_2 , " & _
                     "   SHEET_3 , ROW_3 , COLLUMN_3 , " & _
                     "   SHEET_4 , ROW_4 , COLLUMN_4 , " & _
                     "   ROW_6 , COLLUMN_6 ,  ROW_7 , COLLUMN_7 , " & _
                     "   FROMDATAROW , TODATAROW , " & _
                     "   SHEETBB , ROWBB , " & _
                     "   COLLUMNFOOD " & _
                     "   SHEET_5 , ROW_5 , COLLUMN5_1 , COLLUMN5_2 , COLLUMN5_3 , COLLUMN5_4 , COLLUMN5_5 , COLLUMN5_6 , COLLUMN5_7 " & _
                     "   COLLUMN6_1 , COLLUMN6_2 , COLLUMN6_3 , COLLUMN6_4 , COLLUMN6_5 , COLLUMN6_6 , COLLUMN6_7 " & _
                     ") " & _
                     "VALUES " & _
                     "( " & _
                        SHEET_1 & "," & ROW_1 & ",'" & COLLUMN_1 & "'," & _
                        SHEET_2 & "," & ROW_2 & ",'" & COLLUMN_2 & "'," & _
                        SHEET_3 & "," & ROW_3 & ",'" & COLLUMN_3 & "'," & _
                        SHEET_4 & "," & ROW_4 & ",'" & COLLUMN_4 & "', " & _
                        ROW_6 & ",'" & COLLUMN_6 & "', " & ROW_7 & ",'" & COLLUMN_7 & "', " & FROMDATAROW & "," & TODATAROW & ", " & _
                        SHEETBB & "," & ROWBB & ", " & _
                        " '" & COLLUMNFOOD & "','" & _
                        SHEET_5 & "," & ROW_5 & ",'" & COLLUMN5_1 & "','" & COLLUMN5_2 & "','" & COLLUMN5_3 & "','" & COLLUMN5_4 & "','" & COLLUMN5_5 & "','" & COLLUMN5_6 & "','" & COLLUMN5_7 & "', " & _
                        " '" & COLLUMN6_1 & "','" & COLLUMN6_2 & "','" & COLLUMN6_3 & "','" & COLLUMN6_4 & "','" & COLLUMN6_5 & "','" & COLLUMN6_6 & "','" & COLLUMN6_7 & "' " & _
                     ") "
         XLS_EST_SET_ID = ID
   Else 'Show edit
      SQL1 = "UPDATE XLS_EST_SET SET " & _
                     "   SHEET_1 = " & SHEET_1 & ", " & _
                     "   ROW_1 = " & ROW_1 & ", " & _
                     "   COLLUMN_1 = '" & COLLUMN_1 & "', " & _
                     "   SHEET_2 = " & SHEET_2 & ", " & _
                     "   ROW_2 = " & ROW_2 & ", " & _
                     "   COLLUMN_2 = '" & COLLUMN_2 & "', " & _
                     "   SHEET_3 = " & SHEET_3 & ", " & _
                     "   ROW_3 = " & ROW_3 & ", " & _
                     "   COLLUMN_3 = '" & COLLUMN_3 & "', " & _
                     "   SHEET_4 = " & SHEET_4 & ", " & _
                     "   ROW_4 = " & ROW_4 & ", " & _
                     "   COLLUMN_4 = '" & COLLUMN_4 & "', " & "   ROW_6 = " & ROW_6 & ", " & "   COLLUMN_6 = '" & COLLUMN_6 & "', " & "   ROW_7 = " & ROW_7 & ", " & "   COLLUMN_7 = '" & COLLUMN_7 & "', " & _
                     "   FROMDATAROW = " & FROMDATAROW & ", " & _
                     "   TODATAROW = " & TODATAROW & ", " & _
                     "   SHEET_5 = " & SHEET_5 & ", " & _
                     "   ROW_5 = " & ROW_5 & ", " & _
                     "   SHEETBB = " & SHEETBB & ", " & _
                     "   ROWBB = " & ROWBB & ", " & _
                     "   COLLUMNFOOD = '" & COLLUMNFOOD & "', " & _
                     "   COLLUMN5_1 = '" & COLLUMN5_1 & "', " & "   COLLUMN5_2 = '" & COLLUMN5_2 & "', " & "   COLLUMN5_3 = '" & COLLUMN5_3 & "', " & "   COLLUMN5_4 = '" & COLLUMN5_4 & "', " & _
                     "   COLLUMN5_5 = '" & COLLUMN5_5 & "', " & "   COLLUMN5_6 = '" & COLLUMN5_6 & "', " & "   COLLUMN5_7 = '" & COLLUMN5_7 & "', " & _
                     "   COLLUMN6_1 = '" & COLLUMN6_1 & "', " & "   COLLUMN6_2 = '" & COLLUMN6_2 & "', " & "   COLLUMN6_3 = '" & COLLUMN6_3 & "', " & "   COLLUMN6_4 = '" & COLLUMN6_4 & "', " & _
                     "   COLLUMN6_5 = '" & COLLUMN6_5 & "', " & "   COLLUMN6_6 = '" & COLLUMN6_6 & "', " & "   COLLUMN6_7 = '" & COLLUMN6_7 & "' " & _
                     "WHERE XLS_EST_SET_ID = " & XLS_EST_SET_ID
   End If
   
   If (AddEditMode = SHOW_ADD) Or (AddEditMode = SHOW_EDIT) Then
      m_Conn.Execute (SQL1)
   End If
End Function

Public Function DeleteData()
Dim SQL1 As String
      SQL1 = "DELETE FROM XLS_EST_SET WHERE XLS_EST_SET_ID = " & XLS_EST_SET_ID
      m_Conn.Execute (SQL1)
End Function
