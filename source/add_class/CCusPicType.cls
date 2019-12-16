VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CCusPigType"
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

Public CUS_PIG_TYPE_CODE  As String
Public CUS_PIG_TYPE_NAME As String
Public CUS_PIG_TYPE_BREED As Long
Public CUS_PIG_TYPE_KHUN As Long
Public CUS_PIG_TYPE_PIGGY As Long
Public CUS_PIG_TYPE_YEAR As Long

Public ORDER_BY As Long
Public ORDER_TYPE  As Long

Public KEY_ID As Long
Public KEY_LOOKUP As String

Private Sub Class_Initialize()
   Set m_Conn = glbDatabaseMngr.DBLegacyConnection
End Sub

'In this class we will not handle any error or any transaction
Public Function QueryData(Ind As Long, Rs As ADODB.Recordset, ItemCount As Long, Optional OwnerEnterprise As Boolean = True)
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

If Ind = 1 Then
   SelectStr1 = "SELECT CPT.*  "
   SelectStr2 = "SELECT COUNT(*)  "
   
   FromStr = "FROM CUS_PIG_TYPE CPT "
   
   WhereStr = ""
   
End If
   
   If Trim(CUS_PIG_TYPE_CODE) <> "" Then
      If WildCard(CUS_PIG_TYPE_CODE, SubLen, NewStr) Then
         If WhereStr = "" Then
            WhereStr = "WHERE (CUS_PIG_TYPE_CODE LIKE '" & ChangeQuote(Trim(CUS_PIG_TYPE_CODE)) & "')"
         Else
            WhereStr = WhereStr & "AND (CUS_PIG_TYPE_CODE LIKE '" & ChangeQuote(Trim(CUS_PIG_TYPE_CODE)) & "')"
         End If
      Else
         If WhereStr = "" Then
            WhereStr = "WHERE (CUS_PIG_TYPE_CODE = '" & ChangeQuote(Trim(CUS_PIG_TYPE_CODE)) & "')"
         Else
            WhereStr = WhereStr & "AND (CUS_PIG_TYPE_CODE = '" & ChangeQuote(Trim(CUS_PIG_TYPE_CODE)) & "')"
         End If
      End If
   End If
   
     If Trim(CUS_PIG_TYPE_NAME) <> "" Then
      If WildCard(CUS_PIG_TYPE_NAME, SubLen, NewStr) Then
         If WhereStr = "" Then
            WhereStr = "WHERE (CUS_PIG_TYPE_NAME LIKE '" & ChangeQuote(Trim(CUS_PIG_TYPE_NAME)) & "')"
         Else
            WhereStr = WhereStr & "AND (CUS_PIG_TYPE_NAME LIKE '" & ChangeQuote(Trim(CUS_PIG_TYPE_NAME)) & "')"
         End If
      Else
         If WhereStr = "" Then
            WhereStr = "WHERE (CUS_PIG_TYPE_NAME = '" & ChangeQuote(Trim(CUS_PIG_TYPE_NAME)) & "')"
         Else
            WhereStr = WhereStr & "AND (CUS_PIG_TYPE_NAME = '" & ChangeQuote(Trim(CUS_PIG_TYPE_NAME)) & "')"
         End If
      End If
   End If
   
   If CUS_PIG_TYPE_BREED > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (CPT.CUS_PIG_TYPE_BREED= " & CUS_PIG_TYPE_BREED & ") "
      Else
         WhereStr = WhereStr & "AND (CPT.CUS_PIG_TYPE_BREED = " & CUS_PIG_TYPE_BREED & ") "
      End If
   End If
   
    If CUS_PIG_TYPE_KHUN > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (CPT.CUS_PIG_TYPE_KHUN= " & CUS_PIG_TYPE_KHUN & ") "
      Else
         WhereStr = WhereStr & "AND (CPT.CUS_PIG_TYPE_KHUN = " & CUS_PIG_TYPE_KHUN & ") "
      End If
   End If
   
    If CUS_PIG_TYPE_PIGGY > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (CPT.CUS_PIG_TYPE_PIGGY= " & CUS_PIG_TYPE_PIGGY & ") "
      Else
         WhereStr = WhereStr & "AND (CPT.CUS_PIG_TYPE_PIGGY = " & CUS_PIG_TYPE_PIGGY & ") "
      End If
   End If
   
   If CUS_PIG_TYPE_YEAR > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (CPT.CUS_PIG_TYPE_YEAR= " & CUS_PIG_TYPE_YEAR & ") "
      Else
         WhereStr = WhereStr & "AND (CPT.CUS_PIG_TYPE_YEAR = " & CUS_PIG_TYPE_YEAR & ") "
      End If
   End If
   
   OrderByStr = ""
   OrderTypeStr = ""
   
   If ORDER_BY = 1 Then
      OrderByStr = "ORDER BY CUS_PIG_TYPE_CODE "
   ElseIf ORDER_BY = 2 Then
      OrderByStr = "ORDER BY CPT.CUS_PIG_TYPE_NAME ASC,CPT.CUS_PIG_TYPE_CODE "
   Else
      OrderByStr = "ORDER BY CUS_PIG_TYPE_CODE "
   End If
   If ORDER_TYPE = 1 Then
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
   CUS_PIG_TYPE_CODE = NVLS(Rs("CUS_PIG_TYPE_CODE"), "")
   CUS_PIG_TYPE_NAME = NVLS(Rs("CUS_PIG_TYPE_NAME"), "")
   CUS_PIG_TYPE_BREED = NVLI(Rs("CUS_PIG_TYPE_BREED"), -1)
   CUS_PIG_TYPE_KHUN = NVLI(Rs("CUS_PIG_TYPE_KHUN"), -1)
   CUS_PIG_TYPE_PIGGY = NVLI(Rs("CUS_PIG_TYPE_PIGGY"), -1)
   CUS_PIG_TYPE_YEAR = NVLI(Rs("CUS_PIG_TYPE_YEAR"), -1)
End Function
Public Function AddEditData()
Dim ServerDateTime As String
Dim SQL1 As String
   
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   
   If AddEditMode = SHOW_ADD Then
      SQL1 = "INSERT INTO CUS_PIG_TYPE " & _
                     "( " & _
                     "   CUS_PIG_TYPE_CODE, " & _
                     "   CUS_PIG_TYPE_NAME, " & _
                     "   CUS_PIG_TYPE_BREED, " & _
                     "   CUS_PIG_TYPE_KHUN, " & _
                     "   CUS_PIG_TYPE_PIGGY, " & _
                     "   CUS_PIG_TYPE_YEAR " & _
                     ") " & _
                     "VALUES " & _
                     "( " & _
                       "'" & ChangeQuote(Trim(CUS_PIG_TYPE_CODE)) & "'" & "," & _
                        "'" & ChangeQuote(Trim(CUS_PIG_TYPE_NAME)) & "'," & _
                         "" & CUS_PIG_TYPE_BREED & "" & ", " & _
                         "" & CUS_PIG_TYPE_KHUN & "" & ", " & _
                         "" & CUS_PIG_TYPE_PIGGY & "" & ", " & _
                         "" & CUS_PIG_TYPE_YEAR & "" & _
                     ") "
   Else 'Show edit
      SQL1 = "UPDATE CUS_PIG_TYPE SET " & _
                     "   CUS_PIG_TYPE_NAME = '" & ChangeQuote(Trim(CUS_PIG_TYPE_NAME)) & "' " & ", " & _
                     "   CUS_PIG_TYPE_BREED = " & CUS_PIG_TYPE_BREED & ", " & _
                     "   CUS_PIG_TYPE_KHUN = " & CUS_PIG_TYPE_KHUN & ", " & _
                     "   CUS_PIG_TYPE_PIGGY = " & CUS_PIG_TYPE_PIGGY & ", " & _
                     "   CUS_PIG_TYPE_YEAR = " & CUS_PIG_TYPE_YEAR & _
                     " WHERE CUS_PIG_TYPE_CODE = '" & CUS_PIG_TYPE_CODE & "'"
   End If
   
   If (AddEditMode = SHOW_ADD) Or (AddEditMode = SHOW_EDIT) Then
      m_Conn.Execute (SQL1)
   End If
End Function
Public Function DeleteData()
Dim SQL1 As String
      SQL1 = "DELETE FROM CUS_PIG_TYPE WHERE CUS_PIG_TYPE_CODE = '" & CUS_PIG_TYPE_CODE & "'"
      m_Conn.Execute (SQL1)
End Function
