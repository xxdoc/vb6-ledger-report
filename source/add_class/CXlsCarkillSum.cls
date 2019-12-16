VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CXlsCarkillSum"
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

Public XLS_SUM_ID As Long
Public SUM_ROW As Long
Public OPERATOR_1 As String
Public P_COLUMN_1 As String
Public P_ROW_1 As String
Public OPERATOR_2 As String
Public P_COLUMN_2 As String
Public P_ROW_2 As String
Public OPERATOR_3 As String
Public P_COLUMN_3 As String
Public P_ROW_3 As String
Public OPERATOR_4 As String
Public P_COLUMN_4 As String
Public P_ROW_4 As String
Public OPERATOR_5 As String
Public P_COLUMN_5 As String
Public P_ROW_5 As String
Public HORIZONTAL_FLAG As String

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
   SelectStr1 = "SELECT XS.*  "
   SelectStr2 = "SELECT COUNT(*)  "
   
   FromStr = "FROM XLS_SUM XS "
   
   WhereStr = ""
End If

    If XLS_SUM_ID > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (XLS_SUM_ID = " & ChangeQuote(Trim(XLS_SUM_ID)) & ") "
      Else
         WhereStr = WhereStr & "AND (XLS_SUM_ID = " & ChangeQuote(Trim(XLS_SUM_ID)) & ") "
      End If
   End If
   
   OrderByStr = ""
   OrderTypeStr = ""

   OrderByStr = "ORDER BY XS.HORIZONTAL_FLAG DESC , XS.SUM_ROW "
   OrderTypeStr = "ASC "
   
   SQL1 = SelectStr1 & FromStr & WhereStr & OrderByStr & OrderTypeStr
   SQL2 = SelectStr2 & FromStr & WhereStr
   'debug.print SQL1
   Call glbDatabaseMngr.GetLegacyRs(SQL1, SQL2, True, ItemCount, Rs, glbErrorLog)
End Function

Public Function PopulateFromRS(Ind As Long, Rs As ADODB.Recordset)
  If Ind = 1 Then
    XLS_SUM_ID = NVLI(Rs("XLS_SUM_ID"), -1)
    SUM_ROW = NVLI(Rs("SUM_ROW"), -1)
    OPERATOR_1 = NVLS(Rs("OPERATOR_1"), "")
    P_COLUMN_1 = NVLS(Rs("P_COLUMN_1"), "")
    P_ROW_1 = NVLS(Rs("P_ROW_1"), "")
    OPERATOR_2 = NVLS(Rs("OPERATOR_2"), "")
    P_COLUMN_2 = NVLS(Rs("P_COLUMN_2"), "")
    P_ROW_2 = NVLS(Rs("P_ROW_2"), "")
    OPERATOR_3 = NVLS(Rs("OPERATOR_3"), "")
    P_COLUMN_3 = NVLS(Rs("P_COLUMN_3"), "")
    P_ROW_3 = NVLS(Rs("P_ROW_3"), "")
    OPERATOR_4 = NVLS(Rs("OPERATOR_4"), "")
    P_COLUMN_4 = NVLS(Rs("P_COLUMN_4"), "")
    P_ROW_4 = NVLS(Rs("P_ROW_4"), "")
    OPERATOR_5 = NVLS(Rs("OPERATOR_5"), "")
    P_COLUMN_5 = NVLS(Rs("P_COLUMN_5"), "")
    P_ROW_5 = NVLS(Rs("P_ROW_5"), "")
    HORIZONTAL_FLAG = NVLS(Rs("HORIZONTAL_FLAG"), "")
  End If
End Function
Public Function AddEditData()
Dim ServerDateTime As String
Dim SQL1 As String

   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
'   ENTERPRISE_CODE = glbCompanyCode
   
   If AddEditMode = SHOW_ADD Then
      SQL1 = "INSERT INTO XLS_SUM " & _
                     "( " & _
                     "   SUM_ROW, " & _
                     "   OPERATOR_1, " & "   P_COLUMN_1, " & "   P_ROW_1, " & _
                     "   OPERATOR_2, " & "   P_COLUMN_2, " & "   P_ROW_2, " & _
                     "   OPERATOR_3, " & "   P_COLUMN_3, " & "   P_ROW_3, " & _
                     "   OPERATOR_4, " & "   P_COLUMN_4, " & "   P_ROW_4, " & _
                     "   OPERATOR_5, " & "   P_COLUMN_5, " & "   P_ROW_5, " & _
                     "   HORIZONTAL_FLAG " & _
                     ") " & _
                     "VALUES " & _
                     "( " & _
                        "" & ChangeQuote(Trim(SUM_ROW)) & "," & _
                        "'" & ChangeQuote(Trim(OPERATOR_1)) & "'," & "'" & ChangeQuote(Trim(P_COLUMN_1)) & "'," & "'" & ChangeQuote(Trim(P_ROW_1)) & "'," & _
                        "'" & ChangeQuote(Trim(OPERATOR_2)) & "'," & "'" & ChangeQuote(Trim(P_COLUMN_2)) & "'," & "'" & ChangeQuote(Trim(P_ROW_2)) & "'," & _
                        "'" & ChangeQuote(Trim(OPERATOR_3)) & "'," & "'" & ChangeQuote(Trim(P_COLUMN_3)) & "'," & "'" & ChangeQuote(Trim(P_ROW_3)) & "'," & _
                        "'" & ChangeQuote(Trim(OPERATOR_4)) & "'," & "'" & ChangeQuote(Trim(P_COLUMN_4)) & "'," & "'" & ChangeQuote(Trim(P_ROW_4)) & "'," & _
                        "'" & ChangeQuote(Trim(OPERATOR_5)) & "'," & "'" & ChangeQuote(Trim(P_COLUMN_5)) & "'," & "'" & ChangeQuote(Trim(P_ROW_5)) & "', " & _
                        "'" & ChangeQuote(Trim(HORIZONTAL_FLAG)) & "' " & _
                     ") "

   Else 'Show edit
      SQL1 = "UPDATE XLS_SUM SET " & _
                     "   SUM_ROW = " & ChangeQuote(Trim(SUM_ROW)) & ", " & _
                     "   OPERATOR_1 = '" & ChangeQuote(Trim(OPERATOR_1)) & "', " & _
                     "   P_COLUMN_1 = '" & ChangeQuote(Trim(P_COLUMN_1)) & "', " & _
                     "   P_ROW_1 = '" & ChangeQuote(Trim(P_ROW_1)) & "', " & _
                     "   OPERATOR_2 = '" & ChangeQuote(Trim(OPERATOR_2)) & "', " & _
                     "   P_COLUMN_2 = '" & ChangeQuote(Trim(P_COLUMN_2)) & "', " & _
                     "   P_ROW_2 = '" & ChangeQuote(Trim(P_ROW_2)) & "', " & _
                     "   OPERATOR_3 = '" & ChangeQuote(Trim(OPERATOR_3)) & "', " & _
                     "   P_COLUMN_3 = '" & ChangeQuote(Trim(P_COLUMN_3)) & "', " & _
                     "   P_ROW_3 = '" & ChangeQuote(Trim(P_ROW_3)) & "', " & _
                     "   OPERATOR_4 = '" & ChangeQuote(Trim(OPERATOR_4)) & "', " & _
                     "   P_COLUMN_4 = '" & ChangeQuote(Trim(P_COLUMN_4)) & "', " & _
                     "   P_ROW_4 = '" & ChangeQuote(Trim(P_ROW_4)) & "', " & _
                     "   OPERATOR_5 = '" & ChangeQuote(Trim(OPERATOR_5)) & "', " & _
                     "   P_COLUMN_5 = '" & ChangeQuote(Trim(P_COLUMN_5)) & "', " & _
                     "   P_ROW_5 = '" & ChangeQuote(Trim(P_ROW_5)) & "', " & _
                     "   HORIZONTAL_FLAG = '" & ChangeQuote(Trim(HORIZONTAL_FLAG)) & "' " & _
                     "WHERE XLS_SUM_ID = " & XLS_SUM_ID
   End If
   
   If (AddEditMode = SHOW_ADD) Or (AddEditMode = SHOW_EDIT) Then
      m_Conn.Execute (SQL1)
   End If
End Function

Public Function DeleteData()
Dim SQL1 As String
      SQL1 = "DELETE FROM XLS_SUM WHERE XLS_SUM_ID = " & XLS_SUM_ID
      m_Conn.Execute (SQL1)
End Function
