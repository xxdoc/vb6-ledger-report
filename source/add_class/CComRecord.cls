VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CComRecord"
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
Public RECORD_ID As Long
Public COMTYP As String
Public SLMCOD As String
Public SLMNAME As String
Public MASTER_AREA_ID  As Long
Public AREA_NAME As String
Public GOODS_GROUP_ID As Long
Public GOODS_GROUP_CODE As String
Public GOODS_GROUP_NAME As String
Public FROMDAT As Date
Public TODAT As Date
Public COM_VALUE As Double
Public COM_VALUE01 As Double
Public COM_VALUE02 As Double
Public COM_VALUE03 As Double
Public Flag As String

Public FROM_DOC_DATE As String
Public TO_DOC_DATE As String     'Date

Public ORDER_BY As Long
Public ORDER_TYPE As Long

Private Sub Class_Initialize()
   Set m_FieldList = New Collection
   Set m_Conn = glbDatabaseMngr.DBLegacyConnection
   
    FROM_DOC_DATE = ""
    TO_DOC_DATE = ""
   
End Sub

Private Sub Class_Terminate()
   Set m_FieldList = Nothing
End Sub

Public Function AddEditData()
Dim ServerDateTime As String
Dim SQL1 As String

Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)

   If AddEditMode = SHOW_ADD Then
             SQL1 = "INSERT INTO COM_RECORD " & _
                              "( " & _
                              "   COMTYP, " & _
                              "   SLMCOD, " & _
                              "   SLMNAME, " & _
                              "   MASTER_AREA_ID, " & _
                              "   GOODS_GROUP_ID, " & _
                              "   FROMDAT,  " & _
                              "   TODAT,  " & _
                              "   COM_VALUE, " & _
                              "   FLAG " & _
                              ") " & _
                              "VALUES " & _
                              "( " & _
                                 "'" & ChangeQuote(Trim(COMTYP)) & "'," & _
                                 "'" & ChangeQuote(Trim(SLMCOD)) & "'," & _
                                 "'" & ChangeQuote(Trim(SLMNAME)) & "'," & _
                                 "" & ChangeQuote(Trim(MASTER_AREA_ID)) & "," & _
                                 "" & ChangeQuote(Trim(GOODS_GROUP_ID)) & "," & _
                                 "'" & ChangeQuote(Trim(DateToStringInt(FROMDAT))) & "'," & _
                                 "'" & ChangeQuote(Trim(DateToStringInt(TODAT))) & "'," & _
                                 "" & ChangeQuote(Trim(COM_VALUE)) & "," & _
                                 "'Y'" & _
                              ") "
   Else 'Show edit
         
               SQL1 = "UPDATE COM_RECORD SET " & _
                           "   COMTYP = '" & ChangeQuote(Trim(COMTYP)) & "'," & _
                           "   SLMCOD = '" & ChangeQuote(Trim(SLMCOD)) & "'," & _
                            "  SLMNAME = '" & ChangeQuote(Trim(SLMNAME)) & "'," & _
                           "   MASTER_AREA_ID = " & ChangeQuote(Trim(MASTER_AREA_ID)) & "," & _
                           "   GOODS_GROUP_ID = " & ChangeQuote(Trim(GOODS_GROUP_ID)) & "," & _
                           "   FROMDAT = '" & ChangeQuote(Trim(DateToStringInt(FROMDAT))) & "', " & _
                           "   TODAT = '" & ChangeQuote(Trim(DateToStringInt(TODAT))) & "', " & _
                           "   COM_VALUE = '" & ChangeQuote(Trim(COM_VALUE)) & "'," & _
                           "   FLAG = '" & ChangeQuote(Trim(Flag)) & "'" & _
                           "   WHERE RECORD_ID = " & RECORD_ID

   End If
   
   If (AddEditMode = SHOW_ADD) Or (AddEditMode = SHOW_EDIT) Then
      m_Conn.Execute (SQL1)
   End If
End Function

Public Sub DeleteData()
Dim SQL As String
Dim Tf As CTableField
   SQL = "DELETE FROM COM_RECORD WHERE RECORD_ID =" & RECORD_ID
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
      
      FromStr = " FROM ((COM_RECORD EN "
      FromStr = FromStr & "LEFT OUTER JOIN COMMISSION_MASTER_AREA CMA ON (CMA.MASTER_AREA_ID = EN.MASTER_AREA_ID)) "
      FromStr = FromStr & "LEFT OUTER JOIN GOODS_GROUP GG ON (EN.GOODS_GROUP_ID = GG.GOODS_GROUP_ID)) "
      
   ElseIf Ind = 2 Then
      SelectStr1 = "SELECT EN.* "
      SelectStr2 = "SELECT COUNT(*) "
      
      FromStr = "FROM COM_RECORD EN "
   End If
   
    If SLMCOD <> "" Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (EN.SLMCOD = '" & ChangeQuote(Trim(SLMCOD)) & "') "
      Else
         WhereStr = WhereStr & "AND (EN.SLMCOD = '" & ChangeQuote(Trim(SLMCOD)) & "') "
      End If
   End If

   If FROM_DOC_DATE <> "" Then
      If WhereStr = "" Then
         WhereStr = "WHERE (EN.TODAT >= '" & ChangeQuote(FROM_DOC_DATE) & "') "
      Else
         WhereStr = WhereStr & "AND (EN.TODAT >= '" & ChangeQuote(FROM_DOC_DATE) & "') "
      End If
   End If

   If TO_DOC_DATE <> "" Then
      If WhereStr = "" Then
         WhereStr = "WHERE (EN.TODAT <= '" & ChangeQuote(TO_DOC_DATE) & "') "
      Else
         WhereStr = WhereStr & "AND (EN.TODAT <= '" & ChangeQuote(TO_DOC_DATE) & "') "
      End If
   End If
   
  If COMTYP <> "" Then
      If WhereStr = "" Then
         WhereStr = "WHERE (EN.ComTyp = '" & ChangeQuote(COMTYP) & "') "
      Else
         WhereStr = WhereStr & "AND (EN.ComTyp = '" & ChangeQuote(COMTYP) & "') "
      End If
   End If
  
   If Ind = 1 Then
      OrderByStr = "ORDER BY EN.SLMCOD ASC , EN.TODAT ASC , EN.MASTER_AREA_ID ASC , EN.GOODS_GROUP_ID"
  Else
      OrderByStr = "ORDER BY EN.SLMCOD ASC , EN.TODAT ASC , EN.MASTER_AREA_ID ASC , EN.GOODS_GROUP_ID"
   End If
   
      If ORDER_TYPE = 1 Then
         OrderByStr = OrderByStr & " ASC "
      ElseIf ORDER_TYPE = 2 Then
         OrderByStr = OrderByStr & " ASC "
      Else
         OrderByStr = OrderByStr & " ASC "
      End If
         
   SQL1 = SelectStr1 & FromStr & WhereStr & GroupByStr & OrderByStr
   SQL2 = SelectStr2 & FromStr & WhereStr & GroupByStr
   ''debug.print SQL1
   Call glbDatabaseMngr.GetLegacyRs(SQL1, SQL2, True, ItemCount, Rs, glbErrorLog)
End Sub

Public Sub PopulateFromRS(Ind As Long, Rs As ADODB.Recordset)
   If Ind = 1 Then
      RECORD_ID = NVLI(Rs("RECORD_ID"), -1)
      COMTYP = NVLS(Rs("COMTYP"), "")
      SLMCOD = NVLS(Rs("SLMCOD"), "")
      SLMNAME = NVLS(Rs("SLMNAME"), "")
      AREA_NAME = NVLS(Rs("MASTER_AREA_NAME"), "")
      MASTER_AREA_ID = NVLI(Rs("MASTER_AREA_ID"), -1)
      FROMDAT = InternalDateToDate(NVLS(Rs("FROMDAT"), ""))
      TODAT = InternalDateToDate(NVLS(Rs("TODAT"), ""))
      COM_VALUE = NVLD(Rs("COM_VALUE"), -1)
      GOODS_GROUP_ID = NVLI(Rs("GOODS_GROUP_ID"), -1)
      GOODS_GROUP_CODE = NVLS(Rs("GOODS_GROUP_CODE"), "")
      GOODS_GROUP_NAME = NVLS(Rs("GOODS_GROUP_NAME"), "")
  ElseIf Ind = 2 Then
      RECORD_ID = NVLI(Rs("RECORD_ID"), -1)
      COMTYP = NVLS(Rs("COMTYP"), "")
      SLMCOD = NVLS(Rs("SLMCOD"), "")
      SLMNAME = NVLS(Rs("SLMNAME"), "")
      MASTER_AREA_ID = NVLI(Rs("MASTER_AREA_ID"), -1)
      FROMDAT = InternalDateToDate(NVLS(Rs("FROMDAT"), ""))
      TODAT = InternalDateToDate(NVLS(Rs("TODAT"), ""))
      COM_VALUE = NVLD(Rs("COM_VALUE"), -1)
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
