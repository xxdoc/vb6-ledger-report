VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CMasterFromToDetail"
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

Public MASTER_FROMTO_DETAIL_ID As Long
Public MASTER_FROMTO_ID As Long
Public ORDER_BY As Long
Public ORDER_TYPE As Long
Public VALID_FROM As Date
Public VALID_TO As Date
Public GP As Double
Public SLMCOD As String
Public SLMNAME As String
Public MASTER_PARAMETER_ID As Long
Public MASTER_PARAMETER_NAME As String
Public MASTER_PARAMETER_VALUE As Double
Public MASTER_FROMTO_NO As String
Public MASTER_FROMTO_DESC As String

Public ShowMode As SHOW_MODE_TYPE
Public Flag As String
Public QueryFlag As Long

Private Sub Class_Initialize()
   Set m_FieldList = New Collection
   Set m_Conn = glbDatabaseMngr.DBLegacyConnection
   VALID_TO = -1
   VALID_FROM = -1
   Call RegisterFields
End Sub


Private Sub Class_Terminate()
   Set m_FieldList = Nothing
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
'
'
'Public Sub AddEditData()
'Dim SQL As String
'
'Dim Tf As CTableField
'
'   Set Tf = m_FieldList(PKField)
'
'   Call PopulateInternalField(ShowMode, Me)
'
'   If ShowMode = SHOW_ADD Then
'      SQL = GenerateInsertSQL(Me)
'   ElseIf ShowMode = SHOW_EDIT Then
'      SQL = GenerateUpdateSQL(Me)
'   End If
'
'   Call m_ADOConn.Execute(SQL)
'End Sub


Public Function AddEditData()
Dim ServerDateTime As String
Dim SQL1 As String
Dim Tf As CTableField

  ' Set Tf = m_FieldList(PKField)
   Call PopulateInternalField(ShowMode, Me)
Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   
'   ENTERPRISE_CODE = glbCompanyCode
   
   If ShowMode = SHOW_ADD Then
             SQL1 = "INSERT INTO MASTER_FROMTO_DETAIL " & _
                              "( " & _
                              "   MASTER_FROMTO_ID, " & _
                              "   GP, " & _
                              "   MASTER_PARAMETER_ID, " & _
                              "   SLMCOD, " & _
                              "   SLMNAME, " & _
                              "   FLAG " & _
                              ") " & _
                              "VALUES " & _
                              "( " & _
                                 "" & ChangeQuote(Trim(MASTER_FROMTO_ID)) & "," & _
                                 "" & ChangeQuote(Trim(GP)) & "," & _
                                 "" & ChangeQuote(Trim(MASTER_PARAMETER_ID)) & "," & _
                                 " '" & ChangeQuote(Trim(SLMCOD)) & "'," & _
                                 " '" & ChangeQuote(Trim(SLMNAME)) & "' ," & _
                                 "'" & ChangeQuote(Trim(Flag)) & "'" & _
                              ") "
   Else 'Show edit
         
               SQL1 = "UPDATE MASTER_FROMTO_DETAIL SET " & _
                           "   MASTER_FROMTO_ID = " & ChangeQuote(Trim(MASTER_FROMTO_ID)) & "," & _
                           "   GP = " & ChangeQuote(Trim(GP)) & "," & _
                           "   MASTER_PARAMETER_ID = " & ChangeQuote(Trim(MASTER_PARAMETER_ID)) & "," & _
                           "   SLMCOD = '" & ChangeQuote(Trim(SLMCOD)) & "'," & _
                           "   SLMNAME = '" & ChangeQuote(Trim(SLMNAME)) & "'," & _
                           "   FLAG = '" & ChangeQuote(Trim(Flag)) & "'" & _
                           "   WHERE MASTER_FROMTO_DETAIL_ID = " & MASTER_FROMTO_DETAIL_ID
   End If
   
   If (ShowMode = SHOW_ADD) Or (ShowMode = SHOW_EDIT) Then
      m_Conn.Execute (SQL1)
   End If
End Function
Public Sub DeleteData()
Dim SQL As String
Dim Tf As CTableField

'   Set Tf = m_FieldList(PKField)
      
   SQL = "DELETE FROM MASTER_FROMTO_DETAIL WHERE  MASTER_FROMTO_DETAIL_ID = " & MASTER_FROMTO_DETAIL_ID
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

'
   If Ind = 1 Then
      SelectStr1 = "SELECT NM.* "
      SelectStr1 = SelectStr1 & ",MR.MASTER_PARAMETER_NAME ,MR.MASTER_PARAMETER_VALUE "
      SelectStr2 = "SELECT COUNT(*) "
      
      FromStr = "FROM (MASTER_FROMTO_DETAIL  NM "
       FromStr = FromStr & "LEFT OUTER JOIN COMMISSION_MASTER_PARAMETER MR ON (NM.MASTER_PARAMETER_ID = MR.MASTER_PARAMETER_ID)) "


   ElseIf Ind = 2 Then
      SelectStr1 = "SELECT NM.*,MFT.* "
      SelectStr2 = "SELECT COUNT(*) "
   
      FromStr = "FROM MASTER_FROMTO_DETAIL NM "
      FromStr = FromStr & "LEFT OUTER JOIN MASTER_FROMTO2 MFT ON (NM.MASTER_FROMTO_ID = MFT.MASTER_FROMTO_ID ) "
      
      TempStr = "WHERE "
      TempDate = DateToStringIntLow(VALID_FROM)
      WhereStr = WhereStr & TempStr & " (MFT.VALID_FROM <= '" & ChangeQuote(Trim(TempDate)) & "')"
      TempStr = "AND "
      TempDate = DateToStringIntLow(VALID_TO)
      WhereStr = WhereStr & TempStr & " (MFT.VALID_TO >= '" & ChangeQuote(Trim(TempDate)) & "')"
      
   ElseIf Ind = 3 Then
      SelectStr1 = "SELECT NM.* "
      SelectStr1 = SelectStr1 & ",MR.MASTER_PARAMETER_NAME ,MR.MASTER_PARAMETER_VALUE "
      SelectStr1 = SelectStr1 & ",MA.VALID_FROM ,MA.VALID_TO ,MA.MASTER_FROMTO_NO ,MA.MASTER_FROMTO_DESC "
      SelectStr2 = "SELECT COUNT(*) "
      
      FromStr = "FROM ((MASTER_FROMTO_DETAIL  NM "
       FromStr = FromStr & "LEFT OUTER JOIN COMMISSION_MASTER_PARAMETER MR ON (NM.MASTER_PARAMETER_ID = MR.MASTER_PARAMETER_ID)) "
          FromStr = FromStr & "LEFT OUTER JOIN MASTER_FROMTO2 MA ON (NM.MASTER_FROMTO_ID = MA.MASTER_FROMTO_ID)) "

   End If
   
   If MASTER_FROMTO_DETAIL_ID > 0 Then
      If Len(WhereStr) > 0 Then
         TempStr = "AND "
      Else
         TempStr = "WHERE "
      End If
      WhereStr = WhereStr & TempStr & " (NM.MASTER_FROMTO_DETAIL_ID = " & MASTER_FROMTO_DETAIL_ID & ") " & vbCrLf
   End If
   
   If MASTER_FROMTO_ID > 0 Then
      If Len(WhereStr) > 0 Then
         TempStr = "AND "
      Else
         TempStr = "WHERE "
      End If
      WhereStr = WhereStr & TempStr & " (NM.MASTER_FROMTO_ID = " & MASTER_FROMTO_ID & ") " & vbCrLf
   End If
   
     If Ind = 3 And VALID_FROM > -1 Then                  ' จะเอาไว้ โหลดรีพอร์ต
      TempDate = DateToStringIntLow(Trim(VALID_FROM))
      If WhereStr = "" Then
         WhereStr = "WHERE (MA.VALID_FROM <= '" & ChangeQuote(TempDate) & "') "  ' เครื่องหมายจะตรงกันข้ามจากปรกติ เพราะต้องการให้ครอบครุม
      Else
         WhereStr = WhereStr & "AND (MA.VALID_FROM <= '" & ChangeQuote(TempDate) & "') "
      End If
   End If

   If Ind = 3 And VALID_TO > -1 Then
      TempDate = DateToStringIntHi(Trim(VALID_TO))
      If WhereStr = "" Then
         WhereStr = "WHERE (MA.VALID_TO>= '" & ChangeQuote(TempDate) & "') "
      Else
         WhereStr = WhereStr & "AND (MA.VALID_TO >= '" & ChangeQuote(TempDate) & "') "
      End If
   End If
'   If MASTER_FROMTO_TYPE > 0 Then
'      If Len(WhereStr) > 0 Then
'         TempStr = "AND "
'      Else
'         TempStr = "WHERE "
'      End If
'      WhereStr = WhereStr & TempStr & " (MFT.MASTER_FROMTO_TYPE = " & MASTER_FROMTO_TYPE & ") " & vbCrLf
'   End If

   If Ind = 1 Then
         OrderByStr = "ORDER BY NM.SLMCOD "
   ElseIf Ind = 2 Then
          OrderByStr = "ORDER BY VALID_FROM "
   ElseIf Ind = 3 Then
          OrderByStr = "ORDER BY NM.SLMCOD "
   End If
   
   If ORDER_TYPE = 1 Then
      OrderByStr = OrderByStr & " ASC "
   Else
      OrderByStr = OrderByStr & " ASC "
   End If
   
   SQL1 = SelectStr1 & FromStr & WhereStr & GroupByStr & OrderByStr
   SQL2 = SelectStr2 & FromStr & WhereStr & GroupByStr
   'debug.print SQL1
   Call glbDatabaseMngr.GetLegacyRs(SQL1, SQL2, True, ItemCount, Rs, glbErrorLog)
End Sub

Public Sub PopulateFromRS(Ind As Long, Rs As ADODB.Recordset)
   If Ind = 1 Then
     MASTER_FROMTO_DETAIL_ID = NVLI(Rs("MASTER_FROMTO_DETAIL_ID"), -1)
      MASTER_FROMTO_ID = NVLI(Rs("MASTER_FROMTO_ID"), -1)
      MASTER_PARAMETER_ID = NVLI(Rs("MASTER_PARAMETER_ID"), -1)
     GP = NVLD(Rs("GP"), -1)
        Flag = NVLS(Rs("FLAG"), "")
        MASTER_PARAMETER_NAME = NVLS(Rs("MASTER_PARAMETER_NAME"), "")
       SLMCOD = NVLS(Rs("SLMCOD"), "")
        SLMNAME = NVLS(Rs("SLMNAME"), "")
         MASTER_PARAMETER_VALUE = NVLD(Rs("MASTER_PARAMETER_VALUE"), -1)
ElseIf Ind = 2 Then
     MASTER_FROMTO_DETAIL_ID = NVLI(Rs("MASTER_FROMTO_DETAIL_ID"), -1)
      MASTER_FROMTO_ID = NVLI(Rs("MASTER_FROMTO_ID"), -1)
      MASTER_PARAMETER_ID = NVLI(Rs("MASTER_PARAMETER_ID"), -1)
            SLMCOD = NVLS(Rs("SLMCOD"), "")
     GP = NVLD(Rs("GP"), -1)
        Flag = NVLS(Rs("FLAG"), "")
   ElseIf Ind = 3 Then
     MASTER_FROMTO_DETAIL_ID = NVLI(Rs("MASTER_FROMTO_DETAIL_ID"), -1)
      MASTER_FROMTO_ID = NVLI(Rs("MASTER_FROMTO_ID"), -1)
      MASTER_PARAMETER_ID = NVLI(Rs("MASTER_PARAMETER_ID"), -1)
              MASTER_PARAMETER_NAME = NVLS(Rs("MASTER_PARAMETER_NAME"), "")
         MASTER_PARAMETER_VALUE = NVLD(Rs("MASTER_PARAMETER_VALUE"), -1)
       SLMCOD = NVLS(Rs("SLMCOD"), "")
        SLMNAME = NVLS(Rs("SLMNAME"), "")
     GP = NVLD(Rs("GP"), -1)
           MASTER_FROMTO_NO = NVLS(Rs("MASTER_FROMTO_NO"), "")
      MASTER_FROMTO_DESC = NVLS(Rs("MASTER_FROMTO_DESC"), "")
      VALID_FROM = InternalDateToDate(NVLS(Rs("VALID_FROM"), ""))
      VALID_TO = InternalDateToDate(NVLS(Rs("VALID_TO"), ""))
        Flag = NVLS(Rs("FLAG"), "")
   End If

'      VALUE_FLAG = NVLS(Rs("VALUE_FLAG"), "")
'      AMOUNT_P_FLAG = NVLS(Rs("AMOUNT_P_FLAG"), "")
'      VALUE_P_FLAG = NVLS(Rs("VALUE_P_FLAG"), "")
'     TAGET_VALUE_FLAG = NVLS(Rs("TAGET_VALUE_FLAG"), "")
'
'      AMOUNT_X_FLAG = NVLS(Rs("AMOUNT_X_FLAG"), "")
'      VALUE_X_FLAG = NVLS(Rs("VALUE_X_FLAG"), "")
'      NO_X_FLAG = NVLS(Rs("NO_X_FLAG"), "")
'
'      GROUP_COM_DESC = NVLS(Rs("GROUP_COM_DESC"), "")
'
'   ElseIf Ind = 2 Then
'        MASTER_FROMTO_DETAIL_ID = NVLI(Rs("MASTER_FROMTO_DETAIL_ID"), -1)
'      MASTER_FROMTO_ID = NVLI(Rs("MASTER_FROMTO_ID"), -1)
'      GROUP_COM_ID = NVLI(Rs("GROUP_COM_ID"), -1)
'      MASTER_FROMTO_DETAIL_FROM = NVLD(Rs("MASTER_FROMTO_DETAIL_FROM"), -1)
'   MASTER_AREA_ID = NVLI(Rs("MASTER_AREA_ID"), -1)
'      MASTER_FROMTO_DETAIL_VALUE1 = NVLD(Rs("MASTER_FROMTO_DETAIL_VALUE1"), -1)
'      MASTER_FROMTO_DETAIL_VALUE2 = NVLD(Rs("MASTER_FROMTO_DETAIL_VALUE2"), -1)
'    MASTER_FROMTO_DETAIL_VALUE3 = NVLD(Rs("MASTER_FROMTO_DETAIL_VALUE3"), -1)
'
'      AMOUNT_FLAG = NVLS(Rs("AMOUNT_FLAG"), "")
'      VALUE_FLAG = NVLS(Rs("VALUE_FLAG"), "")
'      AMOUNT_P_FLAG = NVLS(Rs("AMOUNT_P_FLAG"), "")
'      VALUE_P_FLAG = NVLS(Rs("VALUE_P_FLAG"), "")
'     TAGET_VALUE_FLAG = NVLS(Rs("TAGET_VALUE_FLAG"), "")
'
'      AMOUNT_X_FLAG = NVLS(Rs("AMOUNT_X_FLAG"), "")
'      VALUE_X_FLAG = NVLS(Rs("VALUE_X_FLAG"), "")
'      NO_X_FLAG = NVLS(Rs("NO_X_FLAG"), "")
'
''      GROUP_COM_DESC = NVLS(Rs("GROUP_COM_DESC"), "")
'
'      '---------------------------------------------------------------------------------------------------------------
'    INCLUDE_SUB_FLAG = NVLS(Rs("INCLUDE_SUB_FLAG"), "N")
'      INCLUDE_SUB_PERCENT = NVLD(Rs("INCLUDE_SUB_PERCENT"), -1)
'
'     MULTIPLE_FLAG = NVLS(Rs("MULTIPLE_FLAG"), "N")
'    MULTIPLE_PERCENT = NVLD(Rs("MULTIPLE_PERCENT"), -1)
'
'    STEP_FLAG = NVLS(Rs("STEP_FLAG"), "N")
'    TIER_FLAG = NVLS(Rs("TIER_FLAG"), "N")
'
'      '---------------------------------------------------------------------------------------------------------------

End Sub
Public Sub RegisterFields()
   TableName = "MASTER_FROMTO_DETAIL"
   SequenceName = "MASTER_FROMTO_DETAIL_SEQ"
   PKField = "MASTER_FROMTO_DETAIL_ID"
'
'   Call AddField("MASTER_FROMTO_DETAIL_ID", INT_TYPE, -1, ID_CAT)
'   Call AddField("MASTER_FROMTO_ID", INT_TYPE, 0, DATA_CAT)
'   Call AddField("GROUP_COM_ID", INT_TYPE, -1, DATA_CAT)
'   Call AddField("MASTER_FROMTO_DETAIL_FROM", MONEY_TYPE, 0, DATA_CAT)
'   Call AddField("MASTER_AREA_ID", MONEY_TYPE, 0, DATA_CAT)
'   Call AddField("MASTER_FROMTO_DETAIL_VALUE1", MONEY_TYPE, 0, DATA_CAT)
'   Call AddField("MASTER_FROMTO_DETAIL_VALUE2", MONEY_TYPE, 0, DATA_CAT)
'   Call AddField("MASTER_FROMTO_DETAIL_VALUE3", MONEY_TYPE, 0, DATA_CAT)
'   Call AddField("MASTER_FROMTO_DETAIL_VALUE4", MONEY_TYPE, 0, DATA_CAT)
'
'   Call AddField("AMOUNT_FLAG", STRING_TYPE, "", DATA_CAT)
'   Call AddField("VALUE_FLAG", STRING_TYPE, "", DATA_CAT)
'   Call AddField("AMOUNT_P_FLAG", STRING_TYPE, "", DATA_CAT)
'   Call AddField("VALUE_P_FLAG", STRING_TYPE, "", DATA_CAT)
'   Call AddField("TAGET_VALUE_FLAG", STRING_TYPE, "", DATA_CAT)
'
'   Call AddField("NO_X_FLAG", STRING_TYPE, "", DATA_CAT)
'   Call AddField("AMOUNT_X_FLAG", STRING_TYPE, "", DATA_CAT)
'   Call AddField("VALUE_X_FLAG", STRING_TYPE, "", DATA_CAT)
'
'   Call AddField("CREATE_DATE", DATE_TYPE, -1, CREATE_DATE_CAT)
'   Call AddField("MODIFY_DATE", DATE_TYPE, -1, MODIFY_DATE_CAT)
'   Call AddField("CREATE_BY", INT_TYPE, -1, CREATE_BY_CAT)
'   Call AddField("MODIFY_BY", INT_TYPE, -1, MODIFY_BY_CAT)
'
'   Call AddField("ORDER_BY", INT_TYPE, "", TEMP_CAT)
'   Call AddField("ORDER_TYPE", INT_TYPE, "", TEMP_CAT)
'
'   Call AddField("VALID_FROM", DATE_TYPE, -1, TEMP_CAT)
'   Call AddField("VALID_TO", DATE_TYPE, -1, TEMP_CAT)
'   Call AddField("MASTER_FROMTO_TYPE", INT_TYPE, -1, TEMP_CAT)
'
'   Call AddField("INCLUDE_SUB_FLAG", STRING_TYPE, "N", TEMP_CAT)
'   Call AddField("INCLUDE_SUB_PERCENT", MONEY_TYPE, 0, TEMP_CAT)
'   Call AddField("MULTIPLE_FLAG", STRING_TYPE, "N", TEMP_CAT)
'   Call AddField("MULTIPLE_PERCENT", MONEY_TYPE, 0, TEMP_CAT)
'   Call AddField("STEP_FLAG", STRING_TYPE, "N", TEMP_CAT)
'   Call AddField("TIER_FLAG", STRING_TYPE, "N", TEMP_CAT)
'
'   Call AddField("GROUP_COM_DESC", STRING_TYPE, "", TEMP_CAT)
End Sub
