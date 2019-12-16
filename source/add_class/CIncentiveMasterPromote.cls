VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CIncentiveMasterPromote"
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

Public ShowMode As SHOW_MODE_TYPE
Public QueryFlag As Long

Public Details As Collection
Public CommissionExs As Collection

Public KEY_ID As Long
Public KEY_LOOKUP As String

  Public MASTER_INCENTIVE_ID As Long
   Public MASTER_INCENTIVE_NO As String
   Public MASTER_INCENTIVE_DESC As String
   Public VALID_FROM As Date
   Public VALID_TO As Date
   Public Flag As String
  
   Public ORDER_BY As Long
   Public ORDER_TYPE As Long


Private Sub Class_Initialize()
   Set m_FieldList = New Collection
   Set m_Conn = glbDatabaseMngr.DBLegacyConnection
   Set Details = New Collection
   Set CommissionExs = New Collection
   
   VALID_FROM = -1
   VALID_TO = -1
   Call RegisterFields
End Sub

Private Sub Class_Terminate()
   Set m_FieldList = Nothing
   Set Details = Nothing
   Set CommissionExs = Nothing
End Sub

Private Sub AddField(FName As String, FType As FIELD_TYPE, FValue As Variant, FCat As FIELD_CAT)
Dim F As CTableField

   Set F = New CTableField
   Call F.InitField(FName, FType, FValue, FCat)
   Call m_FieldList.Add(F, FName)
   Set F = Nothing
End Sub

Public Function GetFieldValue(FieldName As String)
Dim Tf As CTableField

   Set Tf = m_FieldList(FieldName)
   GetFieldValue = Tf.GetValue
End Function

Public Sub SetFieldValue(FieldName As String, Value As Variant)
Dim Tf As CTableField

   Set Tf = m_FieldList(FieldName)
   Tf.SetValue (Value)
End Sub



Public Function AddEditData()
Dim ServerDateTime As String
Dim SQL1 As String

Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   
'   ENTERPRISE_CODE = glbCompanyCode
   
   If ShowMode = SHOW_ADD Then
             SQL1 = "INSERT INTO MASTER_INCENTIVE_PROMOTE " & _
                              "( " & _
                              "   MASTER_INCENTIVE_NO, " & _
                              "   MASTER_INCENTIVE_DESC, " & _
                              "   VALID_FROM, " & _
                              "   VALID_TO, " & _
                              "   FLAG " & _
                              ") " & _
                              "VALUES " & _
                              "( " & _
                                 "'" & ChangeQuote(Trim(MASTER_INCENTIVE_NO)) & "'," & _
                                 "'" & ChangeQuote(Trim(MASTER_INCENTIVE_DESC)) & "'," & _
                                 "'" & ChangeQuote(Trim(DateToStringInt(VALID_FROM))) & "'," & _
                                 "'" & ChangeQuote(Trim(DateToStringInt(VALID_TO))) & "'," & _
                                 "'Y'" & _
                              ") "
   Else 'Show edit
         
               SQL1 = "UPDATE MASTER_INCENTIVE_PROMOTE SET " & _
                           "   MASTER_INCENTIVE_NO = '" & ChangeQuote(Trim(MASTER_INCENTIVE_NO)) & "'," & _
                           "   MASTER_INCENTIVE_DESC = '" & ChangeQuote(Trim(MASTER_INCENTIVE_DESC)) & "'," & _
                            "  VALID_FROM = '" & ChangeQuote(Trim(DateToStringInt(VALID_FROM))) & "'," & _
                           "   VALID_TO = '" & ChangeQuote(Trim(DateToStringInt(VALID_TO))) & "'," & _
                           "   FLAG = '" & ChangeQuote(Trim(Flag)) & "'" & _
                           "   WHERE MASTER_INCENTIVE_ID = " & MASTER_INCENTIVE_ID

   End If
   
   If (ShowMode = SHOW_ADD) Or (ShowMode = SHOW_EDIT) Then
      m_Conn.Execute (SQL1)
   End If
End Function

Public Sub DeleteData()
Dim SQL As String
Dim Tf As CTableField
      
   SQL = "DELETE FROM MASTER_INCENTIVE_PROMOTE WHERE MASTER_INCENTIVE_ID = " & MASTER_INCENTIVE_ID
      m_Conn.Execute (SQL)
      
  SQL = "DELETE FROM INCENTIVE_PROMOTE WHERE MASTER_INCENTIVE_ID = " & MASTER_INCENTIVE_ID
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
Dim NewStr As String
Dim SubLen As Long
Dim TempDate As String
   
   If Ind = 1 Then
      SelectStr1 = "SELECT NM.* "
      SelectStr2 = "SELECT COUNT(*) "
      
      FromStr = "FROM MASTER_INCENTIVE_PROMOTE NM "
      
   End If
   
   If MASTER_INCENTIVE_ID > 0 Then
      If Len(WhereStr) > 0 Then
         TempStr = "AND "
      Else
         TempStr = "WHERE "
      End If
      WhereStr = WhereStr & TempStr & " (NM.MASTER_INCENTIVE_ID = " & MASTER_INCENTIVE_ID & ") " & vbCrLf
   End If
   
'   If MASTER_INCENTIVE_TYPE > 0 Then
'      If Len(WhereStr) > 0 Then
'         TempStr = "AND "
'      Else
'         TempStr = "WHERE "
'      End If
'      WhereStr = WhereStr & TempStr & " (NM.MASTER_INCENTIVE_TYPE = " & MASTER_INCENTIVE_TYPE & ") " & vbCrLf
'   End If
'
   If VALID_FROM > -1 Then
      TempDate = DateToStringIntLow(VALID_FROM)
      If Len(WhereStr) > 0 Then
         TempStr = "AND "
      Else
         TempStr = "WHERE "
      End If
      WhereStr = WhereStr & TempStr & " (NM.VALID_FROM >= '" & ChangeQuote(Trim(TempDate)) & "')"
   End If
   
   If VALID_TO > -1 Then
      TempDate = DateToStringIntHi(VALID_TO)
      If Len(WhereStr) > 0 Then
         TempStr = "AND "
      Else
         TempStr = "WHERE "
      End If
      WhereStr = WhereStr & TempStr & " (NM.VALID_TO <= '" & ChangeQuote(Trim(TempDate)) & "')"
   End If
   
   If Ind = 1 Then
      If ORDER_BY = 1 Then
         OrderByStr = "ORDER BY VALID_FROM "
      ElseIf ORDER_BY = 2 Then
         OrderByStr = "ORDER BY VALID_FROM "
      ElseIf ORDER_BY = 3 Then
         OrderByStr = "ORDER BY VALID_TO "
      Else
         OrderByStr = "ORDER BY MASTER_INCENTIVE_ID "
      End If
   End If
   
   If ORDER_TYPE = 1 Then
      OrderByStr = OrderByStr & " ASC "
   Else
      OrderByStr = OrderByStr & " DESC "
   End If
   
   SQL1 = SelectStr1 & FromStr & WhereStr & GroupByStr & OrderByStr
   SQL2 = SelectStr2 & FromStr & WhereStr & GroupByStr
   'debug.print SQL1
   Call glbDatabaseMngr.GetLegacyRs(SQL1, SQL2, True, ItemCount, Rs, glbErrorLog)
End Sub

Public Sub PopulateFromRS(Ind As Long, Rs As ADODB.Recordset)
   If Ind = 1 Then
      MASTER_INCENTIVE_ID = NVLI(Rs("MASTER_INCENTIVE_ID"), -1)
      MASTER_INCENTIVE_NO = NVLS(Rs("MASTER_INCENTIVE_NO"), "")
      MASTER_INCENTIVE_DESC = NVLS(Rs("MASTER_INCENTIVE_DESC"), "")
      VALID_FROM = InternalDateToDate(NVLS(Rs("VALID_FROM"), ""))
      VALID_TO = InternalDateToDate(NVLS(Rs("VALID_TO"), ""))
      Flag = NVLS(Rs("FLAG"), "")
      
      KEY_ID = NVLI(Rs("MASTER_INCENTIVE_ID"), -1)
      KEY_LOOKUP = NVLS(Rs("MASTER_INCENTIVE_NO"), "")
   End If
End Sub

Public Sub RegisterFields()
   TableName = "MASTER_INCENTIVE"
   SequenceName = "MASTER_INCENTIVE_SEQ"
   PKField = "MASTER_INCENTIVE_ID"
'
'   Call AddField("MASTER_INCENTIVE_ID", INT_TYPE, -1, ID_CAT)
'   Call AddField("MASTER_INCENTIVE_NO", STRING_TYPE, "", DATA_CAT)
'   Call AddField("MASTER_INCENTIVE_DESC", STRING_TYPE, "", DATA_CAT)
'   Call AddField("VALID_FROM", DATE_TYPE, -1, DATA_CAT)
'   Call AddField("VALID_TO", DATE_TYPE, -1, DATA_CAT)
'   Call AddField("MASTER_INCENTIVE_TYPE", INT_TYPE, -1, DATA_CAT)
'
'  Call AddField("FLAG", STRING_TYPE, "", DATA_CAT)
'      Call AddField("ORDER_BY", INT_TYPE, "", TEMP_CAT)
'   Call AddField("ORDER_TYPE", INT_TYPE, "", TEMP_CAT)
   
 
   
'   Call AddField("INCLUDE_SUB_FLAG", STRING_TYPE, "N", DATA_CAT)
'   Call AddField("INCLUDE_SUB_PERCENT", MONEY_TYPE, 0, DATA_CAT)
'   Call AddField("MULTIPLE_FLAG", STRING_TYPE, "N", DATA_CAT)
'   Call AddField("MULTIPLE_PERCENT", MONEY_TYPE, 0, DATA_CAT)
'   Call AddField("STEP_FLAG", STRING_TYPE, "N", DATA_CAT)
'   Call AddField("TIER_FLAG", STRING_TYPE, "N", DATA_CAT)
'
'   Call AddField("CREATE_DATE", DATE_TYPE, -1, CREATE_DATE_CAT)
'   Call AddField("MODIFY_DATE", DATE_TYPE, -1, MODIFY_DATE_CAT)
'   Call AddField("CREATE_BY", INT_TYPE, -1, CREATE_BY_CAT)
'   Call AddField("MODIFY_BY", INT_TYPE, -1, MODIFY_BY_CAT)
   
End Sub

'Public Function GetFieldValue(FieldName As String)
'Dim Tf As CTableField
'
'   Set Tf = m_FieldList(FieldName)
'   GetFieldValue = Tf.GetValue
'End Function
'
'Public Sub SetFieldValue(FieldName As String, Value As Variant)
'Dim Tf As CTableField
'
'   Set Tf = m_FieldList(FieldName)
'   Tf.SetValue (Value)
'End Sub
