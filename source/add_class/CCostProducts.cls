VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CCostProducts"
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

Public PRODUCT_ID As Long
Public PRODUCT_CODE As String
Public PRODUCT_NAME As String
Public COST_PRODUCT As Double

Public KEY_ID As Long
Public KEY_LOOKUP As String

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
Dim GroupByStr As String
Dim TempDate As String
Dim SubLen As Long
Dim NewStr As String

   If Ind = 1 Then
      SelectStr1 = "SELECT CPRO.* "
      SelectStr2 = "SELECT COUNT(*)  "
      
      FromStr = "FROM COST_PRODUCTS CPRO "
   End If
   
   If PRODUCT_ID > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (PRODUCT_ID= " & PRODUCT_ID & ") "
      Else
         WhereStr = WhereStr & "AND (PRODUCT_ID = " & PRODUCT_ID & ") "
      End If
   End If
   
   
   If Trim(PRODUCT_CODE) <> "" Then
      If WildCard(PRODUCT_CODE, SubLen, NewStr) Then
         If WhereStr = "" Then
            WhereStr = "WHERE (PRODUCT_CODE LIKE '" & ChangeQuote(Trim(PRODUCT_CODE)) & "')"
         Else
            WhereStr = WhereStr & "AND (PRODUCT_CODE LIKE '" & ChangeQuote(Trim(PRODUCT_CODE)) & "')"
         End If
      Else
         If WhereStr = "" Then
            WhereStr = "WHERE (PRODUCT_CODE = '" & ChangeQuote(Trim(PRODUCT_CODE)) & "')"
         Else
            WhereStr = WhereStr & "AND (PRODUCT_CODE = '" & ChangeQuote(Trim(PRODUCT_CODE)) & "')"
         End If
      End If
   End If
   
   If Trim(PRODUCT_NAME) <> "" Then
      If WildCard(PRODUCT_NAME, SubLen, NewStr) Then
         If WhereStr = "" Then
            WhereStr = "WHERE (PRODUCT_NAME LIKE '" & ChangeQuote(Trim(PRODUCT_NAME)) & "')"
         Else
            WhereStr = WhereStr & "AND (PRODUCT_NAME LIKE '" & ChangeQuote(Trim(PRODUCT_NAME)) & "')"
         End If
      Else
         If WhereStr = "" Then
            WhereStr = "WHERE (PRODUCT_NAME = '" & ChangeQuote(Trim(PRODUCT_NAME)) & "')"
         Else
            WhereStr = WhereStr & "AND (PRODUCT_NAME = '" & ChangeQuote(Trim(PRODUCT_NAME)) & "')"
         End If
      End If
   End If
   
   If COST_PRODUCT > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (CPRO.COST_PRODUCT = " & COST_PRODUCT & ") "
      Else
         WhereStr = WhereStr & "AND (CPRO.COST_PRODUCT = " & COST_PRODUCT & ") "
      End If
   End If
   
   GroupByStr = ""

   OrderByStr = ""
   OrderTypeStr = ""
   
   If Ind = 1 Then
      If OrderBy = 1 Then
         OrderByStr = "ORDER BY CPRO.PRODUCT_CODE  "
      Else
         OrderByStr = "ORDER BY CPRO.PRODUCT_NAME  "
      End If
   End If
   
   If OrderType = 1 Then
      OrderTypeStr = "ASC "
   ElseIf OrderType = 2 Then
      OrderTypeStr = "DESC "
   Else
      OrderTypeStr = "ASC "
   End If
   
   SQL1 = SelectStr1 & FromStr & WhereStr & GroupByStr & OrderByStr & OrderTypeStr
   SQL2 = SelectStr2 & FromStr & WhereStr
   
   Call glbDatabaseMngr.GetLegacyRs(SQL1, SQL2, True, ItemCount, Rs, glbErrorLog)
End Function

Public Function PopulateFromRS(Ind As Long, Rs As ADODB.Recordset)
   If Ind = 1 Then
      PRODUCT_ID = NVLI(Rs("PRODUCT_ID"), -1)
      PRODUCT_CODE = NVLS(Rs("PRODUCT_CODE"), "")
      PRODUCT_NAME = NVLS(Rs("PRODUCT_NAME"), "")
      COST_PRODUCT = NVLD(Rs("COST_PRODUCT"), 0)
   End If
End Function
Public Function AddEditData()
Dim ServerDateTime As String
Dim SQL1 As String
Dim ID As Long

   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   
   If AddEditMode = SHOW_ADD Then
      SQL1 = "INSERT INTO COST_PRODUCTS " & _
                     "( " & _
                     "   PRODUCT_CODE, " & _
                     "   PRODUCT_NAME, " & _
                     "   COST_PRODUCT " & _
                     ") " & _
                     "VALUES " & _
                     "( " & _
                        "'" & ChangeQuote(Trim(PRODUCT_CODE)) & "'" & "," & _
                        "'" & ChangeQuote(Trim(PRODUCT_NAME)) & "'" & "," & _
                        "" & COST_PRODUCT & "" & _
                     ") "
   Else
            SQL1 = "UPDATE COST_PRODUCTS SET " & _
                        "   PRODUCT_CODE = '" & ChangeQuote(Trim(PRODUCT_CODE)) & "' ," & _
                        "   PRODUCT_NAME = '" & ChangeQuote(Trim(PRODUCT_NAME)) & "' ," & _
                        "   COST_PRODUCT = " & COST_PRODUCT & " " & _
                        " WHERE PRODUCT_ID = " & PRODUCT_ID
   End If
   
   If (AddEditMode = SHOW_ADD) Or (AddEditMode = SHOW_EDIT) Then
      m_Conn.Execute (SQL1)
   End If
End Function

Public Function DeleteData()
Dim SQL1 As String
      SQL1 = "DELETE FROM COST_PRODUCTS WHERE PRODUCT_ID  = " & PRODUCT_ID
      m_Conn.Execute (SQL1)
End Function
