VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CCommissionChart"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public TableName As String
Public SequenceName As String
Public PKField As String

Public m_FieldList As Collection
Private m_Conn As ADODB.Connection

Public AddEditMode As SHOW_MODE_TYPE
Public Flag As String

Public COMMISSION_CHART_ID As Long
Public MASTER_FROMTO_ID As Long
Public PARENT_ID As Long
Public SALE_ID As String
Public AREA_NAME As String
Public CHILD_COUNT As Long
Public MASTER_AREA_ID  As Long
Public BUDGET As String

      Public MASTER_FROMTO_NO As String
     Public MASTER_FROMTO_DESC As String
      Public MASTER_FROMTO_TYPE As Long
      Public VALID_FROM As Date
      Public VALID_TO As Date
      
      Public GOODS_GROUP_ID As Long
      Public GOODS_GROUP_CODE As String
      Public GOODS_GROUP_NAME As String

Public P_SALE_ID As String

Public ORDER_BY As Long
Public ORDER_TYPE As Long

Private Sub Class_Initialize()
   Set m_FieldList = New Collection
   Set m_Conn = glbDatabaseMngr.DBLegacyConnection
   
   VALID_FROM = -1
   VALID_TO = -1
End Sub

Private Sub Class_Terminate()
   Set m_FieldList = Nothing
End Sub

Public Function AddEditData()
Dim ServerDateTime As String
Dim SQL1 As String

Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)

   If AddEditMode = SHOW_ADD Then
             SQL1 = "INSERT INTO COMMISSION_CHART " & _
                              "( " & _
                              "   MASTER_FROMTO_ID, " & _
                              "   PARENT_ID, " & _
                              "   SALE_ID, " & _
                              "   CHILD_COUNT, " & _
                              "   MASTER_AREA_ID, " & _
                              "   GOODS_GROUP_ID, " & _
                              "   BUDGET, " & _
                              "   FLAG " & _
                              ") " & _
                              "VALUES " & _
                              "( " & _
                                 "'" & ChangeQuote(Trim(MASTER_FROMTO_ID)) & "'," & _
                                 "'" & ChangeQuote(Trim(PARENT_ID)) & "'," & _
                                 "'" & ChangeQuote(Trim(SALE_ID)) & "'," & _
                                 "'" & ChangeQuote(Trim(CHILD_COUNT)) & "'," & _
                                 "" & ChangeQuote(Trim(MASTER_AREA_ID)) & "," & _
                                 "" & ChangeQuote(Trim(GOODS_GROUP_ID)) & "," & _
                                 "'" & ChangeQuote(Trim(BUDGET)) & "'," & _
                                 "'Y'" & _
                              ") "
   Else 'Show edit
         
               SQL1 = "UPDATE COMMISSION_CHART SET " & _
                           "   MASTER_FROMTO_ID = '" & ChangeQuote(Trim(MASTER_FROMTO_ID)) & "'," & _
                           "   PARENT_ID = '" & ChangeQuote(Trim(PARENT_ID)) & "'," & _
                           "   SALE_ID = '" & ChangeQuote(Trim(SALE_ID)) & "'," & _
                            "  CHILD_COUNT = '" & ChangeQuote(Trim(CHILD_COUNT)) & "'," & _
                           "   MASTER_AREA_ID = " & ChangeQuote(Trim(MASTER_AREA_ID)) & "," & _
                           "   GOODS_GROUP_ID = " & ChangeQuote(Trim(GOODS_GROUP_ID)) & "," & _
                           "   BUDGET = '" & ChangeQuote(Trim(BUDGET)) & "'," & _
                           "   FLAG = '" & ChangeQuote(Trim(Flag)) & "'" & _
                           "   WHERE COMMISSION_CHART_ID = " & COMMISSION_CHART_ID

   End If
   
   If (AddEditMode = SHOW_ADD) Or (AddEditMode = SHOW_EDIT) Then
      m_Conn.Execute (SQL1)
   End If
End Function

Public Sub DeleteData()
Dim SQL As String
Dim Tf As CTableField
   SQL = "DELETE FROM COMMISSION_CHART WHERE COMMISSION_CHART_ID =" & COMMISSION_CHART_ID
   m_Conn.Execute (SQL)
End Sub

Public Sub QueryData(Ind As Long, Rs As ADODB.Recordset, ItemCount As Long)
Dim SQL1 As String
Dim SQL2 As String
Dim SelectStr1 As String
Dim SelectStr2 As String
Dim FromStr As String
Dim WhereStr As String
Dim GroupByStr As String
Dim OrderByStr As String
Dim TempStr As String
Dim TempDate As String

   
   If Ind = 1 Then
      SelectStr1 = "SELECT EN.* "
      SelectStr1 = SelectStr1 & ",CMA.MASTER_AREA_NAME "
      SelectStr1 = SelectStr1 & ",GG.GOODS_GROUP_NAME ,GG.GOODS_GROUP_CODE "
      SelectStr2 = "SELECT COUNT(*) "
      
      FromStr = " FROM ((COMMISSION_CHART EN "
      FromStr = FromStr & "LEFT OUTER JOIN COMMISSION_MASTER_AREA CMA ON (CMA.MASTER_AREA_ID = EN.MASTER_AREA_ID)) "
      FromStr = FromStr & "LEFT OUTER JOIN GOODS_GROUP GG ON (EN.GOODS_GROUP_ID = GG.GOODS_GROUP_ID)) "
            
      WhereStr = ""
   ElseIf Ind = 2 Then
      SelectStr1 = "SELECT EN.* "
     SelectStr1 = SelectStr1 & ",CMA.MASTER_AREA_NAME "
     SelectStr1 = SelectStr1 & ", MID.* "
     SelectStr1 = SelectStr1 & ",GG.GOODS_GROUP_NAME ,GG.GOODS_GROUP_CODE "

      SelectStr2 = "SELECT COUNT(*) "
      
      FromStr = "FROM (((COMMISSION_CHART EN "
      FromStr = FromStr & "LEFT OUTER JOIN COMMISSION_MASTER_AREA CMA ON (CMA.MASTER_AREA_ID = EN.MASTER_AREA_ID)) "
      FromStr = FromStr & "LEFT OUTER JOIN MASTER_FROMTO MID ON (MID.MASTER_FROMTO_ID = EN.MASTER_FROMTO_ID)) "
      FromStr = FromStr & "LEFT OUTER JOIN GOODS_GROUP GG ON (EN.GOODS_GROUP_ID = GG.GOODS_GROUP_ID)) "

   
      WhereStr = "WHERE (EN.PARENT_ID = 0) "
   ElseIf Ind = 3 Then
      SelectStr1 = "SELECT EN.* "
      SelectStr1 = SelectStr1 & ",CMA.MASTER_AREA_NAME "
      SelectStr1 = SelectStr1 & ", EN2.SALE_ID AS P_SALE_ID "
      SelectStr1 = SelectStr1 & ", MID.* "
      SelectStr1 = SelectStr1 & ",GG.GOODS_GROUP_NAME ,GG.GOODS_GROUP_CODE "

      SelectStr2 = "SELECT COUNT(*) "
      
      FromStr = "FROM ((((COMMISSION_CHART EN "
      FromStr = FromStr & "LEFT OUTER JOIN  COMMISSION_CHART EN2 ON (EN2.COMMISSION_CHART_ID = EN.PARENT_ID)) "
      FromStr = FromStr & "LEFT OUTER JOIN COMMISSION_MASTER_AREA CMA ON (CMA.MASTER_AREA_ID = EN.MASTER_AREA_ID)) "
      FromStr = FromStr & "LEFT OUTER JOIN MASTER_FROMTO MID ON (MID.MASTER_FROMTO_ID = EN.MASTER_FROMTO_ID)) "
      FromStr = FromStr & "LEFT OUTER JOIN GOODS_GROUP GG ON (EN.GOODS_GROUP_ID = GG.GOODS_GROUP_ID)) "

      WhereStr = "WHERE (EN.PARENT_ID is not null) "
   ElseIf Ind = 4 Then
      SelectStr1 = "SELECT EN.* "
      SelectStr2 = "SELECT COUNT(*) "
      
      FromStr = "FROM COMMISSION_CHART EN "
   End If
   
   If MASTER_FROMTO_ID > 0 Then
      If Len(WhereStr) > 0 Then
         TempStr = "AND "
      Else
         TempStr = "WHERE "
      End If
      WhereStr = WhereStr & TempStr & " (EN.MASTER_FROMTO_ID = " & MASTER_FROMTO_ID & ") " & vbCrLf
   End If
   
   If COMMISSION_CHART_ID > 0 Then
      If Len(WhereStr) > 0 Then
         TempStr = "AND "
      Else
         TempStr = "WHERE "
      End If
      WhereStr = WhereStr & TempStr & " (EN.COMMISSION_CHART_ID = " & COMMISSION_CHART_ID & ") " & vbCrLf
   End If
   
   If Ind <> 1 And VALID_FROM > -1 Then                   ' จะเอาไว้ โหลดรีพอร์ต
      TempDate = DateToStringIntLow(Trim(VALID_FROM))
      If WhereStr = "" Then
         WhereStr = "WHERE (MID.VALID_FROM <= '" & ChangeQuote(TempDate) & "') "  ' เครื่องหมายจะตรงกันข้ามจากปรกติ เพราะต้องการให้ครอบครุม
      Else
         WhereStr = WhereStr & "AND (MID.VALID_FROM <= '" & ChangeQuote(TempDate) & "') "
      End If
   End If

   If Ind <> 1 And VALID_TO > -1 Then
      TempDate = DateToStringIntHi(Trim(VALID_TO))
      If WhereStr = "" Then
         WhereStr = "WHERE (MID.VALID_TO>= '" & ChangeQuote(TempDate) & "') "
      Else
         WhereStr = WhereStr & "AND (MID.VALID_TO >= '" & ChangeQuote(TempDate) & "') "
      End If
   End If
'
   If Ind = 1 Then
      OrderByStr = "ORDER BY EN.COMMISSION_CHART_ID "
   ElseIf Ind = 2 Then
      OrderByStr = "ORDER BY EN.PARENT_ID "
   ElseIf Ind = 4 Then
      OrderByStr = "ORDER BY EN.PARENT_ID ASC , SALE_ID "
  Else
      OrderByStr = "ORDER BY EN.COMMISSION_CHART_ID "
   End If
   
      If ORDER_TYPE = 1 Then
         OrderByStr = OrderByStr & " ASC "
      ElseIf ORDER_TYPE = 2 Then
         OrderByStr = OrderByStr & " ASC "
      Else
         OrderByStr = OrderByStr & " DESC "
      End If
         
   SQL1 = SelectStr1 & FromStr & WhereStr & GroupByStr & OrderByStr
   SQL2 = SelectStr2 & FromStr & WhereStr & GroupByStr
   ''debug.print SQL1
   Call glbDatabaseMngr.GetLegacyRs(SQL1, SQL2, True, ItemCount, Rs, glbErrorLog)
End Sub

Public Sub PopulateFromRS(Ind As Long, Rs As ADODB.Recordset)
   If Ind = 1 Then
      COMMISSION_CHART_ID = NVLI(Rs("COMMISSION_CHART_ID"), -1)
      MASTER_FROMTO_ID = NVLI(Rs("MASTER_FROMTO_ID"), -1)
      PARENT_ID = NVLI(Rs("PARENT_ID"), -1)
     SALE_ID = NVLS(Rs("SALE_ID"), "")
      CHILD_COUNT = NVLI(Rs("CHILD_COUNT"), -1)
     AREA_NAME = NVLS(Rs("MASTER_AREA_NAME"), "")
      BUDGET = NVLS(Rs("BUDGET"), "")
       MASTER_AREA_ID = NVLI(Rs("MASTER_AREA_ID"), -1)
       
       GOODS_GROUP_ID = NVLI(Rs("GOODS_GROUP_ID"), -1)
       GOODS_GROUP_CODE = NVLS(Rs("GOODS_GROUP_CODE"), "")
       GOODS_GROUP_NAME = NVLS(Rs("GOODS_GROUP_NAME"), "")
  ElseIf Ind = 2 Then
      COMMISSION_CHART_ID = NVLI(Rs("COMMISSION_CHART_ID"), -1)
 '     MASTER_FROMTO_ID = NVLI(Rs("MASTER_FROMTO_ID"), -1)
      PARENT_ID = NVLI(Rs("PARENT_ID"), -1)
      SALE_ID = NVLS(Rs("SALE_ID"), "")
      CHILD_COUNT = NVLI(Rs("CHILD_COUNT"), -1)
      AREA_NAME = NVLS(Rs("MASTER_AREA_NAME"), "")
      BUDGET = NVLS(Rs("BUDGET"), "")
       MASTER_AREA_ID = NVLI(Rs("MASTER_AREA_ID"), -1)
       
'       MASTER_FROMTO_ID = NVLI(Rs("MASTER_FROMTO_ID"), -1)
       MASTER_FROMTO_NO = NVLS(Rs("MASTER_FROMTO_NO"), "")
       MASTER_FROMTO_DESC = NVLS(Rs("MASTER_FROMTO_DESC"), "")
'       MASTER_FROMTO_TYPE = NVLI(Rs(" MASTER_FROMTO_TYPE "), -1)
       VALID_FROM = InternalDateToDate(NVLS(Rs("VALID_FROM"), ""))
       VALID_TO = InternalDateToDate(NVLS(Rs("VALID_TO"), ""))
  '     Flag = NVLS(Rs("Flag"), "")
       GOODS_GROUP_ID = NVLI(Rs("GOODS_GROUP_ID"), -1)
       GOODS_GROUP_CODE = NVLS(Rs("GOODS_GROUP_CODE"), "")
       GOODS_GROUP_NAME = NVLS(Rs("GOODS_GROUP_NAME"), "")
    ElseIf Ind = 3 Then
      COMMISSION_CHART_ID = NVLI(Rs("COMMISSION_CHART_ID"), -1)
 '     MASTER_FROMTO_ID = NVLI(Rs("MASTER_FROMTO_ID"), -1)
      PARENT_ID = NVLI(Rs("PARENT_ID"), -1)
      SALE_ID = NVLS(Rs("SALE_ID"), "")
      P_SALE_ID = NVLS(Rs("P_SALE_ID"), "")
      CHILD_COUNT = NVLI(Rs("CHILD_COUNT"), -1)
    AREA_NAME = NVLS(Rs("MASTER_AREA_NAME"), "")
      BUDGET = NVLS(Rs("BUDGET"), "")
       MASTER_AREA_ID = NVLI(Rs("MASTER_AREA_ID"), -1)
       
'       MASTER_FROMTO_ID = NVLI(Rs("MASTER_FROMTO_ID"), -1)
       MASTER_FROMTO_NO = NVLS(Rs("MASTER_FROMTO_NO"), "")
       MASTER_FROMTO_DESC = NVLS(Rs("MASTER_FROMTO_DESC"), "")
'       MASTER_FROMTO_TYPE = NVLI(Rs(" MASTER_FROMTO_TYPE "), -1)
       VALID_FROM = InternalDateToDate(NVLS(Rs("VALID_FROM"), ""))
       VALID_TO = InternalDateToDate(NVLS(Rs("VALID_TO"), ""))
  '     Flag = NVLS(Rs("Flag"), "")
       GOODS_GROUP_ID = NVLI(Rs("GOODS_GROUP_ID"), -1)
       GOODS_GROUP_CODE = NVLS(Rs("GOODS_GROUP_CODE"), "")
       GOODS_GROUP_NAME = NVLS(Rs("GOODS_GROUP_NAME"), "")
     ElseIf Ind = 4 Then
      Flag = NVLS(Rs("FLAG"), "")
      MASTER_FROMTO_ID = NVLI(Rs("MASTER_FROMTO_ID"), -1)
      PARENT_ID = NVLI(Rs("PARENT_ID"), -1)
     SALE_ID = NVLS(Rs("SALE_ID"), "")
      CHILD_COUNT = NVLI(Rs("CHILD_COUNT"), -1)
     MASTER_AREA_ID = NVLI(Rs("MASTER_AREA_ID"), -1)
      BUDGET = NVLS(Rs("BUDGET"), "")
       GOODS_GROUP_ID = NVLI(Rs("GOODS_GROUP_ID"), -1)
   End If
End Sub
Public Sub SetFieldValue(FieldName As String, Value As Variant)
Dim Tf As CTableField

   Set Tf = m_FieldList(FieldName)
   Tf.SetValue (Value)
End Sub

Public Function GetFieldValue(FieldName As String)
Dim Tf As CTableField

   Set Tf = m_FieldList(FieldName)
   GetFieldValue = Tf.GetValue
End Function
