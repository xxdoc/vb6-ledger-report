VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CComMasSubPromote"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
' ฐานข้อมูล access
Private m_Conn As ADODB.Connection

Public AddEditMode As SHOW_MODE_TYPE
Public QueryFlag As Long
Public OrderBy As Long
Public OrderType As Long

' End of generic data member
Public DetailsCom1 As Collection
Public DetailsCom2 As Collection
Public DetailsCom3 As Collection
Public CommissionExs As Collection

Public MASTER_COMMISS_SUB_PROMOTE_ID As Long
Public MASTER_Commiss_ID As Long
Public Flag As String
Public CREDIT_NAME As Long
Public CREDIT_DESC As String 'Double
Public CREDIT_TYP As String

Public SLM_ID As String 'Double
Public CUS_ID As String

Public VALID_FROM As Date
Public VALID_TO As Date
Public MASTER_Commiss_NO As String
Public FROM_CMPL_DATE As Date
Public TO_CMPL_DATE As Date

Public ORDER_BY As Long
Public ORDER_TYPE  As Long

Public ShowMode As SHOW_MODE_TYPE


Private Sub Class_Initialize()
   Set m_Conn = glbDatabaseMngr.DBLegacyConnection
   
      Set DetailsCom1 = New Collection
      Set DetailsCom2 = New Collection
         Set DetailsCom3 = New Collection
   Set CommissionExs = New Collection
   
   FROM_CMPL_DATE = -1
   TO_CMPL_DATE = -1
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
   
   If Ind = 1 Then            'สำหรับ ปีไหน มีข้อมูลอะไร Tab1
      SelectStr1 = "SELECT CS.* , CSY.MASTER_COMMISS_NO ,CSY.VALID_FROM, CSY.VALID_TO "             '**********
      SelectStr2 = "SELECT COUNT(*)  "
      
      FromStr = "FROM MASTER_COMMISS_SUB_PROMOTE CS "
      FromStr = FromStr & "LEFT OUTER JOIN MASTER_COMMISS_PROMOTE CSY ON (CS.MASTER_COMMISS_ID = CSY.MASTER_COMMISS_ID) "
  ElseIf Ind = 2 Then            ' สำหรับแต่ละบรรทัดของ Grid
     SelectStr1 = "SELECT CS.* "
      SelectStr2 = "SELECT COUNT(*)  "
      
      FromStr = "FROM MASTER_COMMISS_SUB_PROMOTE CS "
   ElseIf Ind = 3 Then             'สำหรับ ปีไหน มีข้อมูลอะไร Tab1
      SelectStr1 = "SELECT CS.* , CSY.MASTER_COMMISS_NO ,CSY.VALID_FROM, CSY.VALID_TO ,CSY.SLM_ID , CSY.CUS_ID "             '**********
      SelectStr2 = "SELECT COUNT(*)  "
      
      FromStr = "FROM MASTER_COMMISS_SUB_PROMOTE CS "
      FromStr = FromStr & "LEFT OUTER JOIN MASTER_COMMISS_PROMOTE CSY ON (CS.MASTER_COMMISS_ID = CSY.MASTER_COMMISS_ID) "
   ElseIf Ind = 4 Then             'สำหรับ
      SelectStr1 = "SELECT CS.* , CSY.MASTER_COMMISS_NO ,CSY.VALID_FROM, CSY.VALID_TO "             '**********
      SelectStr2 = "SELECT COUNT(*)  "
      
      FromStr = "FROM MASTER_COMMISS_SUB_PROMOTE CS "
      FromStr = FromStr & "LEFT OUTER JOIN MASTER_COMMISS_PROMOTE CSY ON (CS.MASTER_COMMISS_ID = CSY.MASTER_COMMISS_ID) "
End If

If MASTER_Commiss_ID > 0 Then
      WhereStr = "WHERE (CS.MASTER_COMMISS_ID=" & ChangeQuote(Trim(MASTER_Commiss_ID)) & ") "
End If
         
    If MASTER_COMMISS_SUB_PROMOTE_ID > 0 Then
      If WhereStr = "" Then
         WhereStr = "WHERE (CS.MASTER_COMMISS_SUB_PROMOTE_ID = " & MASTER_COMMISS_SUB_PROMOTE_ID & ") "
      Else
         WhereStr = WhereStr & "AND (CS.MASTER_COMMISS_SUB_PROMOTE_ID = " & MASTER_COMMISS_SUB_PROMOTE_ID & ") "
      End If
   End If
   
      If Trim(CREDIT_TYP) <> "" Then
      If WhereStr = "" Then
         WhereStr = "WHERE (CS.CREDIT_TYP = '" & CREDIT_TYP & "') "
      Else
         WhereStr = WhereStr & "AND (CS.CREDIT_TYP = '" & CREDIT_TYP & "') "
      End If
   End If
   
   'วันที่รับชำระ
   If FROM_CMPL_DATE > -1 Then                     ' จะเอาไว้ โหลดรีพอร์ต
      TempDate = DateToStringIntLow(Trim(FROM_CMPL_DATE))
      If WhereStr = "" Then
         WhereStr = "WHERE (CSY.VALID_FROM <= '" & ChangeQuote(TempDate) & "') "  ' เครื่องหมายจะตรงกันข้ามจากปรกติ เพราะต้องการให้ครอบครุม
      Else
         WhereStr = WhereStr & "AND (CSY.VALID_FROM <= '" & ChangeQuote(TempDate) & "') "
      End If
   End If

   If TO_CMPL_DATE > -1 Then
      TempDate = DateToStringIntHi(Trim(TO_CMPL_DATE))
      If WhereStr = "" Then
         WhereStr = "WHERE (CSY.VALID_TO >= '" & ChangeQuote(TempDate) & "') "
      Else
         WhereStr = WhereStr & "AND (CSY.VALID_TO >= '" & ChangeQuote(TempDate) & "') "
      End If
   End If
   
   
   OrderByStr = ""
   OrderTypeStr = ""

 If Ind = 1 Then
        OrderByStr = "ORDER BY CS.CREDIT_NAME "
 Else
         OrderByStr = "ORDER BY CS.CREDIT_NAME "
  End If
 
 
If Ind = 1 Then
        OrderTypeStr = "ASC "
ElseIf Ind = 3 Then
        OrderTypeStr = "DESC "
Else
         OrderTypeStr = "ASC "
End If
   
   SQL1 = SelectStr1 & FromStr & WhereStr & OrderByStr & OrderTypeStr
   SQL2 = SelectStr2 & FromStr & WhereStr
   ''debug.print SQL1
  Call glbDatabaseMngr.GetLegacyRs(SQL1, SQL2, True, ItemCount, Rs, glbErrorLog)
End Function

Public Function PopulateFromRS(Ind As Long, Rs As ADODB.Recordset)
If Ind = 1 Then
      MASTER_COMMISS_SUB_PROMOTE_ID = NVLI(Rs("MASTER_COMMISS_SUB_PROMOTE_ID"), -1)
         MASTER_Commiss_ID = NVLI(Rs("MASTER_COMMISS_ID"), -1)
         CREDIT_NAME = NVLI(Rs("CREDIT_NAME"), -1)
         CREDIT_DESC = NVLS(Rs("CREDIT_DESC"), "")
         CREDIT_TYP = NVLS(Rs("CREDIT_TYP"), "")
         
          VALID_FROM = InternalDateToDate(NVLS(Rs("VALID_FROM"), ""))
          VALID_TO = InternalDateToDate(NVLS(Rs("VALID_TO"), ""))

ElseIf Ind = 2 Then
     MASTER_COMMISS_SUB_PROMOTE_ID = NVLI(Rs("MASTER_COMMISS_SUB_PROMOTE_ID"), -1)
         MASTER_Commiss_ID = NVLI(Rs("MASTER_COMMISS_ID"), -1)
         CREDIT_NAME = NVLI(Rs("CREDIT_NAME"), -1)
         CREDIT_DESC = NVLS(Rs("CREDIT_DESC"), "")
         CREDIT_TYP = NVLS(Rs("CREDIT_TYP"), "")
          VALID_FROM = InternalDateToDate(NVLS(Rs("VALID_FROM"), ""))
          VALID_TO = InternalDateToDate(NVLS(Rs("VALID_TO"), ""))
          
ElseIf Ind = 3 Then
      MASTER_COMMISS_SUB_PROMOTE_ID = NVLI(Rs("MASTER_COMMISS_SUB_PROMOTE_ID"), -1)
         MASTER_Commiss_ID = NVLI(Rs("MASTER_COMMISS_ID"), -1)
         CREDIT_NAME = NVLI(Rs("CREDIT_NAME"), -1)
         CREDIT_DESC = NVLS(Rs("CREDIT_DESC"), "")
         CREDIT_TYP = NVLS(Rs("CREDIT_TYP"), "")
         MASTER_COMMISS_SUB_PROMOTE_ID = NVLI(Rs("MASTER_COMMISS_SUB_PROMOTE_ID"), -1)
          VALID_FROM = InternalDateToDate(NVLS(Rs("VALID_FROM"), ""))
          VALID_TO = InternalDateToDate(NVLS(Rs("VALID_TO"), ""))
          CUS_ID = NVLS(Rs("CUS_ID"), "")
          SLM_ID = NVLS(Rs("SLM_ID"), "")
ElseIf Ind = 5 Then
       MASTER_COMMISS_SUB_PROMOTE_ID = NVLI(Rs("MASTER_COMMISS_SUB_PROMOTE_ID"), -1)
         MASTER_Commiss_ID = NVLI(Rs("MASTER_COMMISS_ID"), -1)
         CREDIT_NAME = NVLI(Rs("CREDIT_NAME"), -1)
         CREDIT_DESC = NVLS(Rs("CREDIT_DESC"), "")
         CREDIT_TYP = NVLS(Rs("CREDIT_TYP"), "")
         
          VALID_FROM = InternalDateToDate(NVLS(Rs("VALID_FROM"), ""))
          VALID_TO = InternalDateToDate(NVLS(Rs("VALID_TO"), ""))
ElseIf Ind = 6 Then
       MASTER_COMMISS_SUB_PROMOTE_ID = NVLI(Rs("MASTER_COMMISS_SUB_PROMOTE_ID"), -1)
         MASTER_Commiss_ID = NVLI(Rs("MASTER_COMMISS_ID"), -1)
         CREDIT_NAME = NVLI(Rs("CREDIT_NAME"), -1)
         CREDIT_DESC = NVLS(Rs("CREDIT_DESC"), "")
         CREDIT_TYP = NVLS(Rs("CREDIT_TYP"), "")
          VALID_FROM = InternalDateToDate(NVLS(Rs("VALID_FROM"), ""))
          VALID_TO = InternalDateToDate(NVLS(Rs("VALID_TO"), ""))
End If

End Function
Public Function AddEditData()
Dim ServerDateTime As String
Dim SQL1 As String

   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
 '  ENTERPRISE_CODE = glbCompanyCode
   
   If ShowMode = SHOW_ADD Then

               SQL1 = "INSERT INTO MASTER_COMMISS_SUB_PROMOTE " & _
                              "( " & _
                              "   CREDIT_NAME, " & _
                              "   CREDIT_DESC, " & _
                               "   MASTER_COMMISS_ID, " & _
                               "   CREDIT_TYP, " & _
                                "   FLAG " & _
                              ") " & _
                              "VALUES " & _
                              "( " & _
                                 "" & ChangeQuote(Trim(CREDIT_NAME)) & "," & _
                                 "'" & ChangeQuote(Trim(CREDIT_DESC)) & "'," & _
                                  "" & ChangeQuote(Trim(MASTER_Commiss_ID)) & "," & _
                                  "'" & ChangeQuote(Trim(CREDIT_TYP)) & "'," & _
                                  "'" & ChangeQuote(Trim(Flag)) & "' )"

   Else 'Show edit
' 'debug.print MASTER_Commiss_ID
               SQL1 = "UPDATE MASTER_COMMISS_SUB_PROMOTE SET " & _
                           "   CREDIT_NAME = " & ChangeQuote(Trim(CREDIT_NAME)) & ", " & _
                           "   CREDIT_DESC = '" & ChangeQuote(Trim(CREDIT_DESC)) & "', " & _
                           "   MASTER_COMMISS_ID = " & ChangeQuote(Trim(MASTER_Commiss_ID)) & ", " & _
                           "   FLAG = '" & ChangeQuote(Trim(Flag)) & "' " & _
                           "WHERE MASTER_COMMISS_SUB_PROMOTE_ID = " & MASTER_COMMISS_SUB_PROMOTE_ID
   End If
   
   If (ShowMode = SHOW_ADD) Or (ShowMode = SHOW_EDIT) Then
      'debug.print SQL1
      m_Conn.Execute (SQL1)
   End If
End Function

Public Function DeleteData(Optional Ind As Long = 1)
Dim SQL1 As String

If Ind = 1 Then
      SQL1 = "DELETE FROM MASTER_COMMISS_SUB_PROMOTE WHERE MASTER_COMMISS_SUB_PROMOTE_ID = " & MASTER_COMMISS_SUB_PROMOTE_ID  ' ลบ Master sub และ ลูกด้วย
      m_Conn.Execute (SQL1)
      SQL1 = "DELETE FROM COMMISS_PROMOTE WHERE MASTER_COMMISS_SUB_PROMOTE_ID = " & MASTER_COMMISS_SUB_PROMOTE_ID
      m_Conn.Execute (SQL1)
End If

End Function


Private Sub Class_Terminate()
   Set DetailsCom1 = Nothing
      Set DetailsCom2 = Nothing
         Set DetailsCom3 = Nothing
   Set CommissionExs = Nothing
End Sub
