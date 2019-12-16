VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CXlsFood"
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

Public XLS_FOOD_ID As Long
Public XLS_FOOD_CODE As String
Public XLS_FOOD_NAME As String
Public XLS_FOOD_COST As Double
Public XLS_UNIT_ID As Long
Public XLS_COL_BAG As String
Public XLS_COL_BULK As String

Public XLS_UNIT_NAME As String
Public XLS_UNIT_MULTIPLY As Long
Public XLS_UNIT_LIMIT As Long
Public XLS_DESTINATION As String

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
   SelectStr1 = "SELECT XF.*  "
   SelectStr1 = SelectStr1 & ", XU.XLS_UNIT_NAME, XU.XLS_UNIT_MULTIPLY , XU.XLS_UNIT_LIMIT "
   SelectStr2 = "SELECT COUNT(*)  "
   
   FromStr = "FROM XLS_FOOD XF "
   FromStr = FromStr & "LEFT OUTER JOIN XLS_UNIT XU ON (XF.XLS_UNIT_ID = XU.XLS_UNIT_ID) "
   
   WhereStr = ""
'ElseIf ind = 2 Then

End If
   
'   If OwnerEnterprise Then
'      ENTERPRISE_CODE = glbCompanyCode
'   End If
   
   If XLS_FOOD_CODE <> "" Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (XF.XLS_FOOD_CODE = '" & ChangeQuote(Trim(XLS_FOOD_CODE)) & "') "
      Else
         WhereStr = WhereStr & "AND (XF.XLS_FOOD_CODE = '" & ChangeQuote(Trim(XLS_FOOD_CODE)) & "') "
      End If
   End If

   If XLS_FOOD_ID > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (XF.XLS_FOOD_ID = " & ChangeQuote(Trim(XLS_FOOD_ID)) & ") "
      Else
         WhereStr = WhereStr & "AND (XF.XLS_FOOD_ID = " & ChangeQuote(Trim(XLS_FOOD_ID)) & ") "
      End If
   End If

    If XLS_UNIT_ID > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (XF.XLS_UNIT_ID = " & ChangeQuote(Trim(XLS_UNIT_ID)) & ") "
      Else
         WhereStr = WhereStr & "AND (XF.XLS_UNIT_ID = " & ChangeQuote(Trim(XLS_UNIT_ID)) & ") "
      End If
   End If
   
   OrderByStr = ""
   OrderTypeStr = ""

   If OrderBy = 1 Then
      OrderByStr = "ORDER BY XF.XLS_FOOD_CODE "
   ElseIf OrderBy = 2 Then
      OrderByStr = "ORDER BY XU.XLS_UNIT_NAME "
   Else
      OrderByStr = "ORDER BY XF.XLS_FOOD_CODE ASC , XF.XLS_UNIT_ID "
   End If

   If OrderType = 1 Then
      OrderTypeStr = "ASC "
   ElseIf OrderType = 2 Then
      OrderTypeStr = "DESC "
   Else
      OrderTypeStr = "ASC "
   End If
   
   SQL1 = SelectStr1 & FromStr & WhereStr & OrderByStr & OrderTypeStr
   SQL2 = SelectStr2 & FromStr & WhereStr
  ' 'debug.print SQL1
   Call glbDatabaseMngr.GetLegacyRs(SQL1, SQL2, True, ItemCount, Rs, glbErrorLog)
End Function

Public Function PopulateFromRS(Ind As Long, Rs As ADODB.Recordset)
If Ind = 1 Then
   XLS_FOOD_ID = NVLI(Rs("XLS_FOOD_ID"), -1)
   XLS_FOOD_CODE = NVLS(Rs("XLS_FOOD_CODE"), "")
   XLS_FOOD_NAME = NVLS(Rs("XLS_FOOD_NAME"), "")
   XLS_FOOD_COST = NVLD(Rs("XLS_FOOD_COST"), -1)
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
      SQL1 = "INSERT INTO XLS_FOOD " & _
                     "( " & _
                     "   XLS_FOOD_CODE, " & _
                     "   XLS_FOOD_NAME, " & _
                     "   XLS_FOOD_COST, " & _
                     "   XLS_UNIT_ID " & _
                     ") " & _
                     "VALUES " & _
                     "( " & _
                        "'" & ChangeQuote(Trim(XLS_FOOD_CODE)) & "'," & _
                        "'" & ChangeQuote(Trim(XLS_FOOD_NAME)) & "'," & _
                        "" & ChangeQuote(Trim(XLS_FOOD_COST)) & "," & _
                        "" & ChangeQuote(Trim(XLS_UNIT_ID)) & "" & _
                     ") "

   Else 'Show edit
      SQL1 = "UPDATE XLS_FOOD SET " & _
                     "   XLS_FOOD_CODE = '" & ChangeQuote(Trim(XLS_FOOD_CODE)) & "', " & _
                     "   XLS_FOOD_NAME = '" & ChangeQuote(Trim(XLS_FOOD_NAME)) & "', " & _
                     "   XLS_FOOD_COST = " & ChangeQuote(Trim(XLS_FOOD_COST)) & ", " & _
                     "   XLS_UNIT_ID = " & ChangeQuote(Trim(XLS_UNIT_ID)) & " " & _
                     "WHERE XLS_FOOD_ID = " & XLS_FOOD_ID
   End If
   
   If (AddEditMode = SHOW_ADD) Or (AddEditMode = SHOW_EDIT) Then
      m_Conn.Execute (SQL1)
   End If
End Function

Public Function DeleteData()
Dim SQL1 As String
      SQL1 = "DELETE FROM XLS_FOOD WHERE XLS_FOOD_ID = " & XLS_FOOD_ID
      m_Conn.Execute (SQL1)
End Function
