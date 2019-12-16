VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CSupplierGroup"
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

Public SUPPLIER_GROUP_ID As Long
Public GROUP_TYPE_CODE As Long
Public GROUP_TYPE_NAME As String
Public SUB_GROUP_TYPE_CODE As Long
Public SUB_GROUP_TYPE_NAME As String
Public SUPPLIER_CODE  As String
Public SUPPLIER_NAME As String
Public DATA_TYPE_ID As Long
Public DATA_TYPE_NAME As String

Public ORDER_BY As Long
Public ORDER_TYPE  As Long

Public KEY_ID As Long
Public KEY_LOOKUP As String

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
Dim SubLen As Long
Dim NewStr As String

   SelectStr1 = "SELECT UG.*  "
   SelectStr1 = SelectStr1 & ",GT.GROUP_TYPE_NAME "
   SelectStr1 = SelectStr1 & ",DT.DATA_TYPE_NAME "
   SelectStr1 = SelectStr1 & ",SGT.SUB_GROUP_TYPE_NAME "
   
   SelectStr2 = "SELECT COUNT(*)  "
   
   FromStr = "FROM (((SUPPLIER_GROUP UG "
   FromStr = FromStr & " LEFT OUTER JOIN GROUP_TYPE  GT ON (GT.GROUP_TYPE_CODE = UG.GROUP_TYPE_CODE)) "
   FromStr = FromStr & " LEFT OUTER JOIN DATA_TYPE DT ON (DT.DATA_TYPE_ID = UG.DATA_TYPE_ID))"
   FromStr = FromStr & " LEFT OUTER JOIN SUB_GROUP_TYPE SGT ON (SGT.SUB_GROUP_TYPE_CODE = UG.SUB_GROUP_TYPE_CODE))"
   
   WhereStr = ""
'   If OwnerEnterprise Then
'      ENTERPRISE_CODE = glbCompanyCode
'   End If
'
'   If Len(ENTERPRISE_CODE) > 0 Then
'      If Len(WhereStr) <= 0 Then
'         WhereStr = "WHERE (ENTERPRISE_CODE = '" & ChangeQuote(Trim(ENTERPRISE_CODE)) & "') "
'      Else
'         WhereStr = WhereStr & "AND (ENTERPRISE_CODE = '" & ChangeQuote(Trim(ENTERPRISE_CODE)) & "') "
'      End If
'   End If
   
   If SUPPLIER_GROUP_ID > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (SUPPLIER_GROUP_ID= " & SUPPLIER_GROUP_ID & ") "
      Else
         WhereStr = WhereStr & "AND (SUPPLIER_GROUP_ID = " & SUPPLIER_GROUP_ID & ") "
      End If
   End If
   
   If DATA_TYPE_ID > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (UG.DATA_TYPE_ID= " & DATA_TYPE_ID & ") "
      Else
         WhereStr = WhereStr & "AND (UG.DATA_TYPE_ID = " & DATA_TYPE_ID & ") "
      End If
   End If
   
   If GROUP_TYPE_CODE > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (UG.GROUP_TYPE_CODE= " & GROUP_TYPE_CODE & ") "
      Else
         WhereStr = WhereStr & "AND (UG.GROUP_TYPE_CODE = " & GROUP_TYPE_CODE & ") "
      End If
   End If
   
   If SUB_GROUP_TYPE_CODE > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (UG.SUB_GROUP_TYPE_CODE= " & SUB_GROUP_TYPE_CODE & ") "
      Else
         WhereStr = WhereStr & "AND (UG.SUB_GROUP_TYPE_CODE = " & SUB_GROUP_TYPE_CODE & ") "
      End If
   End If
   
   If Trim(SUPPLIER_CODE) <> "" Then
      If WildCard(SUPPLIER_CODE, SubLen, NewStr) Then
         If WhereStr = "" Then
            WhereStr = "WHERE (SUPPLIER_CODE LIKE '" & ChangeQuote(Trim(SUPPLIER_CODE)) & "')"
         Else
            WhereStr = WhereStr & "AND (SUPPLIER_CODE LIKE '" & ChangeQuote(Trim(SUPPLIER_CODE)) & "')"
         End If
      Else
         If WhereStr = "" Then
            WhereStr = "WHERE (SUPPLIER_CODE = '" & ChangeQuote(Trim(SUPPLIER_CODE)) & "')"
         Else
            WhereStr = WhereStr & "AND (SUPPLIER_CODE = '" & ChangeQuote(Trim(SUPPLIER_CODE)) & "')"
         End If
      End If
   End If
   
   OrderByStr = ""
   OrderTypeStr = ""
   
   If ORDER_BY = 1 Then
      OrderByStr = "ORDER BY SUPPLIER_CODE "
   Else
      OrderByStr = "ORDER BY DT.DATA_TYPE_NAME ASC,GT.GROUP_TYPE_NAME ASC,SGT.SUB_GROUP_TYPE_NAME ASC,SUPPLIER_CODE "
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
   SUPPLIER_GROUP_ID = NVLI(Rs("SUPPLIER_GROUP_ID"), -1)
   SUPPLIER_CODE = NVLS(Rs("SUPPLIER_CODE"), "")
   GROUP_TYPE_CODE = NVLI(Rs("GROUP_TYPE_CODE"), -1)
   GROUP_TYPE_NAME = NVLS(Rs("GROUP_TYPE_NAME"), "")
   SUB_GROUP_TYPE_CODE = NVLI(Rs("SUB_GROUP_TYPE_CODE"), -1)
   SUB_GROUP_TYPE_NAME = NVLS(Rs("SUB_GROUP_TYPE_NAME"), "")
   
   DATA_TYPE_ID = NVLI(Rs("DATA_TYPE_ID"), -1)
   DATA_TYPE_NAME = NVLS(Rs("DATA_TYPE_NAME"), "")
End Function
Public Function AddEditData()
Dim ServerDateTime As String
Dim SQL1 As String
   
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   
   If AddEditMode = SHOW_ADD Then
      SQL1 = "INSERT INTO SUPPLIER_GROUP " & _
                     "( " & _
                     "   DATA_TYPE_ID, " & _
                     "   SUPPLIER_CODE, " & _
                     "   GROUP_TYPE_CODE, " & _
                     "   SUB_GROUP_TYPE_CODE " & _
                     ") " & _
                     "VALUES " & _
                     "( " & _
                        DATA_TYPE_ID & "," & _
                        "'" & ChangeQuote(Trim(SUPPLIER_CODE)) & "'," & _
                         "" & GROUP_TYPE_CODE & "," & _
                         "" & SUB_GROUP_TYPE_CODE & "" & _
                     ") "
   Else 'Show edit
      SQL1 = "UPDATE SUPPLIER_GROUP SET " & _
                     "   DATA_TYPE_ID = " & DATA_TYPE_ID & ", " & _
                     "   GROUP_TYPE_CODE = " & GROUP_TYPE_CODE & ", " & _
                     "   SUPPLIER_CODE = '" & ChangeQuote(Trim(SUPPLIER_CODE)) & "' ," & _
                     "   SUB_GROUP_TYPE_CODE = " & SUB_GROUP_TYPE_CODE & " " & _
                     "WHERE SUPPLIER_GROUP_ID = " & SUPPLIER_GROUP_ID
   End If
   
   If (AddEditMode = SHOW_ADD) Or (AddEditMode = SHOW_EDIT) Then
      m_Conn.Execute (SQL1)
   End If
End Function
Public Function DeleteData()
Dim SQL1 As String
      SQL1 = "DELETE FROM SUPPLIER_GROUP WHERE SUPPLIER_GROUP_ID = " & SUPPLIER_GROUP_ID
      m_Conn.Execute (SQL1)
End Function
