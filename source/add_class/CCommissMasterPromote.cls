VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CCommissMasterPromote"
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

Public DetailsCom1 As Collection
Public DetailsCom2 As Collection
Public DetailsCom3 As Collection
Public CommissionExs As Collection

Public KEY_ID As Long
Public KEY_LOOKUP As String

  Public MASTER_Commiss_ID As Long
   Public MASTER_Commiss_NO As String
   Public MASTER_COMMISS_DESC As String
      Public SLM_ID As String
   Public SLM_NAME As String
      Public CUS_ID As String
   Public CUS_NAME As String
   Public VALID_FROM As Date
   Public VALID_TO As Date
   
   Public FromCMPLDat As Date
   Public ToCMPLDat As Date
   
   Public Flag As String
  
   Public ORDER_BY As Long
   Public ORDER_TYPE As Long


Private Sub Class_Initialize()
   Set m_FieldList = New Collection
   Set m_Conn = glbDatabaseMngr.DBLegacyConnection
      Set DetailsCom1 = New Collection
      Set DetailsCom2 = New Collection
         Set DetailsCom3 = New Collection
   Set CommissionExs = New Collection
   
   FromCMPLDat = -1
   ToCMPLDat = -1
   VALID_FROM = -1
    VALID_TO = -1
   
   Call RegisterFields
End Sub

Private Sub Class_Terminate()
   Set m_FieldList = Nothing
   Set DetailsCom1 = Nothing
      Set DetailsCom2 = Nothing
         Set DetailsCom3 = Nothing
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
             SQL1 = "INSERT INTO MASTER_COMMISS_PROMOTE " & _
                              "( " & _
                              "   MASTER_COMMISS_NO, " & _
                              "   MASTER_COMMISS_DESC, " & _
                              "   VALID_FROM, " & _
                              "   VALID_TO, " & _
                              "   SLM_ID, " & _
                              "  SLM_NAME, " & _
                              "  CUS_ID, " & _
                              "  CUS_NAME, " & _
                              "   FLAG " & _
                              ") " & _
                              "VALUES " & _
                              "( " & _
                                 "'" & ChangeQuote(Trim(MASTER_Commiss_NO)) & "'," & _
                                 "'" & ChangeQuote(Trim(MASTER_COMMISS_DESC)) & "'," & _
                                 "'" & ChangeQuote(Trim(DateToStringInt(VALID_FROM))) & "'," & _
                                 "'" & ChangeQuote(Trim(DateToStringInt(VALID_TO))) & "'," & _
                                  "'" & ChangeQuote(Trim(SLM_ID)) & "'," & _
                                 "'" & ChangeQuote(Trim(SLM_NAME)) & "'," & _
                                  "'" & ChangeQuote(Trim(CUS_ID)) & "'," & _
                                 "'" & ChangeQuote(Trim(CUS_NAME)) & "'," & _
                                 "'Y'" & _
                              ") "
   Else 'Show edit
         
               SQL1 = "UPDATE MASTER_COMMISS_PROMOTE SET " & _
                           "   MASTER_COMMISS_NO = '" & ChangeQuote(Trim(MASTER_Commiss_NO)) & "'," & _
                           "   MASTER_COMMISS_DESC = '" & ChangeQuote(Trim(MASTER_COMMISS_DESC)) & "'," & _
                           "   SLM_ID = '" & ChangeQuote(Trim(SLM_ID)) & "'," & _
                           "   SLM_NAME = '" & ChangeQuote(Trim(SLM_NAME)) & "'," & _
                           "   CUS_ID = '" & ChangeQuote(Trim(CUS_ID)) & "'," & _
                           "   CUS_NAME = '" & ChangeQuote(Trim(CUS_NAME)) & "'," & _
                           "  VALID_FROM = '" & ChangeQuote(Trim(DateToStringInt(VALID_FROM))) & "'," & _
                           "   VALID_TO = '" & ChangeQuote(Trim(DateToStringInt(VALID_TO))) & "'," & _
                           "   FLAG = '" & ChangeQuote(Trim(Flag)) & "'" & _
                           "   WHERE MASTER_COMMISS_ID = " & MASTER_Commiss_ID

   End If
   
   If (ShowMode = SHOW_ADD) Or (ShowMode = SHOW_EDIT) Then
      m_Conn.Execute (SQL1)
   End If
End Function

Public Sub DeleteData()
Dim SQL As String
Dim Tf As CTableField
      
      SQL = "DELETE FROM MASTER_COMMISS_PROMOTE WHERE MASTER_COMMISS_ID = " & MASTER_Commiss_ID
      m_Conn.Execute (SQL)
      SQL = "DELETE FROM MASTER_COMMISS_SUB_PROMOTE WHERE MASTER_COMMISS_ID = " & MASTER_Commiss_ID  ' ลบ Master sub และ ลูกด้วย
      m_Conn.Execute (SQL)
      SQL = "DELETE FROM COMMISS_PROMOTE WHERE MASTER_COMMISS_ID = " & MASTER_Commiss_ID
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
      
      FromStr = "FROM MASTER_COMMISS_PROMOTE NM "
   End If
   
   If MASTER_Commiss_ID > 0 Then
      If Len(WhereStr) > 0 Then
         TempStr = "AND "
      Else
         TempStr = "WHERE "
      End If
      WhereStr = WhereStr & TempStr & " (NM.MASTER_COMMISS_ID = " & MASTER_Commiss_ID & ") " & vbCrLf
   End If
   
'   If MASTER_COMMISS_TYPE > 0 Then
'      If Len(WhereStr) > 0 Then
'         TempStr = "AND "
'      Else
'         TempStr = "WHERE "
'      End If
'      WhereStr = WhereStr & TempStr & " (NM.MASTER_COMMISS_TYPE = " & MASTER_COMMISS_TYPE & ") " & vbCrLf
'   End If

   If FromCMPLDat > -1 Then
      TempDate = DateToStringIntLow(FromCMPLDat)
      If Len(WhereStr) > 0 Then
         TempStr = "AND "
      Else
         TempStr = "WHERE "
      End If
      WhereStr = WhereStr & TempStr & " (NM.VALID_FROM <= '" & ChangeQuote(Trim(TempDate)) & "')"  ' เครื่องหมายตรงกันข้าม ไว้โหลดรายงาน
   End If
   
   If ToCMPLDat > -1 Then
      TempDate = DateToStringIntHi(ToCMPLDat)
      If Len(WhereStr) > 0 Then
         TempStr = "AND "
      Else
         TempStr = "WHERE "
      End If
      WhereStr = WhereStr & TempStr & " (NM.VALID_TO >= '" & ChangeQuote(Trim(TempDate)) & "')"
   End If

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
         OrderByStr = "ORDER BY MASTER_COMMISS_ID "
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
      MASTER_Commiss_ID = NVLI(Rs("MASTER_COMMISS_ID"), -1)
      MASTER_Commiss_NO = NVLS(Rs("MASTER_COMMISS_NO"), "")
      MASTER_COMMISS_DESC = NVLS(Rs("MASTER_COMMISS_DESC"), "")
      SLM_ID = NVLS(Rs("SLM_ID"), "")
      SLM_NAME = NVLS(Rs("SLM_NAME"), "")
      CUS_ID = NVLS(Rs("CUS_ID"), "")
      CUS_NAME = NVLS(Rs("CUS_NAME"), "")
      VALID_FROM = InternalDateToDate(NVLS(Rs("VALID_FROM"), ""))
      VALID_TO = InternalDateToDate(NVLS(Rs("VALID_TO"), ""))
      Flag = NVLS(Rs("FLAG"), "")
      
      KEY_ID = NVLI(Rs("MASTER_COMMISS_ID"), -1)
      KEY_LOOKUP = NVLS(Rs("MASTER_COMMISS_NO"), "")
   End If
End Sub

Public Sub RegisterFields()
   TableName = "MASTER_COMMISS"
   SequenceName = "MASTER_COMMISS_SEQ"
   PKField = "MASTER_COMMISS_ID"
'
'   Call AddField("MASTER_COMMISS_ID", INT_TYPE, -1, ID_CAT)
'   Call AddField("MASTER_COMMISS_NO", STRING_TYPE, "", DATA_CAT)
'   Call AddField("MASTER_COMMISS_DESC", STRING_TYPE, "", DATA_CAT)
'   Call AddField("VALID_FROM", DATE_TYPE, -1, DATA_CAT)
'   Call AddField("VALID_TO", DATE_TYPE, -1, DATA_CAT)
'   Call AddField("MASTER_COMMISS_TYPE", INT_TYPE, -1, DATA_CAT)
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
