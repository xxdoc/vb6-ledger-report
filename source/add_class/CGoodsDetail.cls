VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CGoodsDetail"
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

' Public Flag As String
Public GOODS_DETAIL_ID As Long
Public STKCOD As String
Public STKDES As String
Public GOODS_MASTER_ID As Long
Public GOODS_GROUP_ID As Long

Public GOODS_GROUP_NAME As String
Public GOODS_GROUP_CODE As String
Public GOODS_MASTER_NAME As String
Public GOODS_MASTER_CODE As String

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
      SelectStr1 = "SELECT  GD.* , GM.GOODS_MASTER_CODE , GM.GOODS_MASTER_NAME , GG.GOODS_GROUP_CODE , GG.GOODS_GROUP_NAME   "             '**********
      SelectStr2 = "SELECT COUNT(*)  "
      
      FromStr = "FROM (( GOODS_DETAIL GD "
      FromStr = FromStr & "LEFT OUTER JOIN GOODS_MASTER GM ON (GD.GOODS_MASTER_ID = GM.GOODS_MASTER_ID)) "
      FromStr = FromStr & "LEFT OUTER JOIN GOODS_GROUP GG ON (GD.GOODS_GROUP_ID = GG.GOODS_GROUP_ID)) "
  
  ElseIf Ind = 2 Then
      SelectStr1 = "SELECT  GD.* , GM.GOODS_MASTER_NAME "             '**********
      SelectStr2 = "SELECT COUNT(*)  "
      
      FromStr = "FROM GOODS_DETAIL GD "
      FromStr = FromStr & "LEFT OUTER JOIN GOODS_MASTER GM ON (GD.GOODS_MASTER_ID = GM.GOODS_MASTER_ID) "

      WhereStr = "WHERE (GD.GOODS_MASTER_ID >0) "
  End If
   
         If GOODS_DETAIL_ID > 0 Then
            If WhereStr = "" Then
               WhereStr = "WHERE (GD.GOODS_DETAIL_ID=" & ChangeQuote(Trim(GOODS_DETAIL_ID)) & ") "
            Else
               WhereStr = WhereStr & "AND (GD.GOODS_DETAIL_ID=" & ChangeQuote(Trim(GOODS_DETAIL_ID)) & ") "
            End If
         End If
   
         If GOODS_MASTER_ID > 0 Then
            If WhereStr = "" Then
               WhereStr = "WHERE (GD.GOODS_MASTER_ID=" & ChangeQuote(Trim(GOODS_MASTER_ID)) & ") "
            Else
               WhereStr = WhereStr & "AND (GD.GOODS_MASTER_ID=" & ChangeQuote(Trim(GOODS_MASTER_ID)) & ") "
            End If
         End If
         
         If GOODS_GROUP_ID > 0 Then
            If WhereStr = "" Then
               WhereStr = "WHERE (GD.GOODS_GROUP_ID=" & ChangeQuote(Trim(GOODS_GROUP_ID)) & ") "
            Else
               WhereStr = WhereStr & "AND (GD.GOODS_GROUP_ID=" & ChangeQuote(Trim(GOODS_GROUP_ID)) & ") "
            End If
         End If
         
          If Len(STKCOD) > 0 Then
            If WhereStr = "" Then
               WhereStr = "WHERE (GD.STKCOD='" & ChangeQuote(Trim(STKCOD)) & "') "
            Else
               WhereStr = WhereStr & "AND (GD.STKCOD='" & ChangeQuote(Trim(STKCOD)) & "') "
            End If
         End If
         
'           'วันที่รับชำระ
'         If FROM_DATE > -1 Then                     ' จะเอาไว้ โหลดรีพอร์ต
'            TempDate = DateToStringIntLow(Trim(FROM_DATE))
'            If WhereStr = "" Then
'               WhereStr = "WHERE (CSY.FROM_DATE <= '" & ChangeQuote(TempDate) & "') "  ' เครื่องหมายจะตรงกันข้ามจากปรกติ เพราะต้องการให้ครอบครุม
'            Else
'               WhereStr = WhereStr & "AND (CSY.FROM_DATE <= '" & ChangeQuote(TempDate) & "') "
'            End If
'         End If
'
'         If TO_DATE > -1 Then
'            TempDate = DateToStringIntHi(Trim(TO_DATE))
'            If WhereStr = "" Then
'               WhereStr = "WHERE (CSY.TO_DATE >= '" & ChangeQuote(TempDate) & "') "
'            Else
'               WhereStr = WhereStr & "AND (CSY.TO_DATE >= '" & ChangeQuote(TempDate) & "') "
'            End If
'         End If
   
    If Ind = 1 Then
         If OrderBy = 2 Then
             OrderByStr = "ORDER BY GD.STKCOD  "
          ElseIf OrderBy = 3 Then
             OrderByStr = "ORDER BY GD.STKDES "
          Else
             OrderByStr = "ORDER BY GD.GOODS_GROUP_ID "
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
      GOODS_DETAIL_ID = NVLI(Rs("GOODS_DETAIL_ID"), "-1")
      STKCOD = NVLS(Rs("STKCOD"), "")
      STKDES = NVLS(Rs("STKDES"), "")
      GOODS_MASTER_ID = NVLI(Rs("GOODS_MASTER_ID"), "-1")
      GOODS_GROUP_ID = NVLI(Rs("GOODS_GROUP_ID"), "-1")
      
      GOODS_GROUP_NAME = NVLS(Rs("GOODS_GROUP_NAME"), "")
      GOODS_GROUP_CODE = NVLS(Rs("GOODS_GROUP_CODE"), "")
      GOODS_MASTER_NAME = NVLS(Rs("GOODS_MASTER_NAME"), "")
      GOODS_MASTER_CODE = NVLS(Rs("GOODS_MASTER_CODE"), "")
 ElseIf Ind = 2 Then
      GOODS_DETAIL_ID = NVLI(Rs("GOODS_DETAIL_ID"), "-1")
      STKCOD = NVLS(Rs("STKCOD"), "")
      STKDES = NVLS(Rs("STKDES"), -1)
      GOODS_MASTER_ID = NVLI(Rs("GOODS_MASTER_ID"), "-1")
      GOODS_GROUP_ID = NVLI(Rs("GOODS_GROUP_ID"), "-1")

      GOODS_MASTER_NAME = NVLS(Rs("GOODS_MASTER_NAME"), "")

   End If

End Function
Public Function AddEditData()
Dim ServerDateTime As String
Dim SQL1 As String

   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   
'   ENTERPRISE_CODE = glbCompanyCode
   
   If AddEditMode = SHOW_ADD Then
             SQL1 = "INSERT INTO GOODS_DETAIL " & _
                              "( " & _
                              "   STKCOD, " & _
                              "   STKDES, " & _
                              "   GOODS_MASTER_ID, " & _
                              "   GOODS_GROUP_ID " & _
                              ") " & _
                              "VALUES " & _
                              "( " & _
                                 "'" & ChangeQuote(Trim(STKCOD)) & "'," & _
                                 "'" & ChangeQuote(Trim(STKDES)) & "'," & _
                                 "" & ChangeQuote(Trim(GOODS_MASTER_ID)) & "," & _
                                 "" & ChangeQuote(Trim(GOODS_GROUP_ID)) & "" & _
                              ") "
   Else 'Show edit
         
               SQL1 = "UPDATE GOODS_DETAIL SET " & _
                           "   STKCOD = '" & ChangeQuote(Trim(STKCOD)) & "'," & _
                           "   STKDES = '" & ChangeQuote(Trim(STKDES)) & "'," & _
                           "   GOODS_MASTER_ID = " & ChangeQuote(Trim(GOODS_MASTER_ID)) & "," & _
                           "   GOODS_GROUP_ID = " & ChangeQuote(Trim(GOODS_GROUP_ID)) & "" & _
                           "   WHERE GOODS_DETAIL_ID = " & GOODS_DETAIL_ID

   End If
   
   If (AddEditMode = SHOW_ADD) Or (AddEditMode = SHOW_EDIT) Then
      m_Conn.Execute (SQL1)
   End If
End Function

Public Function DeleteData()
Dim SQL1 As String
      SQL1 = "DELETE FROM GOODS_DETAIL WHERE GOODS_DETAIL_ID = " & GOODS_DETAIL_ID
      m_Conn.Execute (SQL1)
'      SQL1 = "DELETE FROM COMMISSION_CUS_AREA WHERE YEAR_ID = " & YEAR_ID
'      m_Conn.Execute (SQL1)
End Function




