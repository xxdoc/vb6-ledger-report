VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CComDonStk"
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

Public COMDONSTK_ID As Long
Public MASTER_COMDONSTK_ID As Long
Public STKCOD As String
Public STKDES As String
Public Flag As String
Public VALID_FROM As Date
Public VALID_TO As Date
Public MASTER_COMDONSTK_NO As String

Public ORDER_BY As Long
Public ORDER_TYPE  As Long

Public ShowMode As SHOW_MODE_TYPE


Private Sub Class_Initialize()
   Set m_Conn = glbDatabaseMngr.DBLegacyConnection

   VALID_FROM = -1
   VALID_TO = -1
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
      SelectStr1 = "SELECT CS.* , CSY.MASTER_COMDONSTK_NO,CSY.VALID_FROM, CSY.VALID_TO "             '**********
      SelectStr2 = "SELECT COUNT(*)  "
      
      FromStr = "FROM COMDONSTK CS "
      FromStr = FromStr & "LEFT OUTER JOIN MASTER_COMDONSTK CSY ON (CS.MASTER_COMDONSTK_ID = CSY.MASTER_COMDONSTK_ID) "
  ElseIf Ind = 2 Then            ' สำหรับแต่ละบรรทัดของ Grid
     SelectStr1 = "SELECT CS.* "
      SelectStr2 = "SELECT COUNT(*)  "
      
      FromStr = "FROM COMDONSTK CS "
   ElseIf Ind = 3 Then             'สำหรับ ปีไหน มีข้อมูลอะไร Tab1
      SelectStr1 = "SELECT CS.* , CSY.MASTER_COMDONSTK_NO,CSY.VALID_FROM, CSY.VALID_TO "             '**********
      SelectStr2 = "SELECT COUNT(*)  "
      
      FromStr = "FROM COMDONSTK CS "
      FromStr = FromStr & "LEFT OUTER JOIN MASTER_COMDONSTK CSY ON (CS.MASTER_COMDONSTK_ID = CSY.MASTER_COMDONSTK_ID) "
   ElseIf Ind = 4 Then             'สำหรับ
      SelectStr1 = "SELECT CS.* , CSY.MASTER_COMDONSTK_NO,CSY.VALID_FROM, CSY.VALID_TO "             '**********
      SelectStr2 = "SELECT COUNT(*)  "
      
      FromStr = "FROM COMDONSTK CS "
      FromStr = FromStr & "LEFT OUTER JOIN MASTER_COMDONSTK CSY ON (CS.MASTER_COMDONSTK_ID = CSY.MASTER_COMDONSTK_ID) "
   ElseIf Ind = 5 Then             'สำหรับ
      SelectStr1 = "SELECT DISTINCT CS.STKCOD, CS.STKDES , CSY.MASTER_COMDONSTK_NO,CSY.VALID_FROM, CSY.VALID_TO "             '**********
      SelectStr2 = "SELECT COUNT(*)  "
      
      FromStr = "FROM COMDONSTK CS "
      FromStr = FromStr & "LEFT OUTER JOIN MASTER_COMDONSTK CSY ON (CS.MASTER_COMDONSTK_ID = CSY.MASTER_COMDONSTK_ID) "
     ElseIf Ind = 6 Then             'สำหรับ GROUP1
      SelectStr1 = "SELECT DISTINCT CS.STKCOD, CSY.MASTER_COMDONSTK_NO,CSY.VALID_FROM, CSY.VALID_TO  "             '**********
      SelectStr2 = "SELECT COUNT(*)  "
      
      FromStr = "FROM COMDONSTK CS "
      FromStr = FromStr & "LEFT OUTER JOIN MASTER_COMDONSTK CSY ON (CS.MASTER_COMDONSTK_ID = CSY.MASTER_COMDONSTK_ID) "
End If
      
If MASTER_COMDONSTK_ID <> 0 Then
      WhereStr = "WHERE (CS.MASTER_COMDONSTK_ID=" & ChangeQuote(Trim(MASTER_COMDONSTK_ID)) & ") "
End If
   
   If Trim(STKCOD) <> "" Then
      If WhereStr = "" Then
         WhereStr = "WHERE (CS.STKCOD = '" & STKCOD & "') "
      Else
         WhereStr = WhereStr & "AND (CS.STKCOD = '" & STKCOD & "') "
      End If
   End If
         
    If COMDONSTK_ID <> 0 Then
      If WhereStr = "" Then
         WhereStr = "WHERE (CS.COMDONSTK_ID = " & COMDONSTK_ID & ") "
      Else
         WhereStr = WhereStr & "AND (CS.COMDONSTK_ID = " & COMDONSTK_ID & ") "
      End If
   End If
   
   'วันที่รับชำระ
   If VALID_FROM > -1 Then                     ' จะเอาไว้ โหลดรีพอร์ต
      TempDate = DateToStringIntLow(Trim(VALID_FROM))
      If WhereStr = "" Then
         WhereStr = "WHERE (CSY.VALID_FROM <= '" & ChangeQuote(TempDate) & "') "  ' เครื่องหมายจะตรงกันข้ามจากปรกติ เพราะต้องการให้ครอบครุม
      Else
         WhereStr = WhereStr & "AND (CSY.VALID_FROM <= '" & ChangeQuote(TempDate) & "') "
      End If
   End If

   If VALID_TO > -1 Then
      TempDate = DateToStringIntHi(Trim(VALID_TO))
      If WhereStr = "" Then
         WhereStr = "WHERE (CSY.VALID_TO >= '" & ChangeQuote(TempDate) & "') "
      Else
         WhereStr = WhereStr & "AND (CSY.VALID_TO >= '" & ChangeQuote(TempDate) & "') "
      End If
   End If
   
   
   OrderByStr = ""
   OrderTypeStr = ""

 'If Ind = 1 Then
        OrderByStr = "ORDER BY CS.STKCOD "
 'End If
 
If Ind = 3 Then
        OrderTypeStr = "DESC "
Else
         OrderTypeStr = "ASC "
 End If
   
   SQL1 = SelectStr1 & FromStr & WhereStr & OrderByStr & OrderTypeStr
   SQL2 = SelectStr2 & FromStr & WhereStr
  ' 'debug.print SQL1
  Call glbDatabaseMngr.GetLegacyRs(SQL1, SQL2, True, ItemCount, Rs, glbErrorLog)
End Function

Public Function PopulateFromRS(Ind As Long, Rs As ADODB.Recordset)
If Ind = 1 Then
        COMDONSTK_ID = NVLI(Rs("COMDONSTK_ID"), -1)
         MASTER_COMDONSTK_ID = NVLI(Rs("MASTER_COMDONSTK_ID"), -1)
         STKCOD = NVLS(Rs("STKCOD"), "")
         STKDES = NVLS(Rs("STKDES"), "")
         MASTER_COMDONSTK_NO = NVLS(Rs("MASTER_COMDONSTK_NO"), "")
         
              VALID_FROM = InternalDateToDate(NVLS(Rs("VALID_FROM"), ""))
              VALID_TO = InternalDateToDate(NVLS(Rs("VALID_TO"), ""))
ElseIf Ind = 2 Then
       COMDONSTK_ID = NVLI(Rs("COMDONSTK_ID"), -1)
         MASTER_COMDONSTK_ID = NVLI(Rs("MASTER_COMDONSTK_ID"), -1)
                  STKCOD = NVLS(Rs("STKCOD"), "")
ElseIf Ind = 5 Then
         MASTER_COMDONSTK_NO = NVLS(Rs("MASTER_COMDONSTK_NO"), "")
         VALID_FROM = InternalDateToDate(NVLS(Rs("VALID_FROM"), ""))
         VALID_TO = InternalDateToDate(NVLS(Rs("VALID_TO"), ""))
         STKCOD = NVLS(Rs("STKCOD"), "")
         STKDES = NVLS(Rs("STKDES"), "")
ElseIf Ind = 6 Then
         MASTER_COMDONSTK_NO = NVLS(Rs("MASTER_COMDONSTK_NO"), "")
         VALID_FROM = InternalDateToDate(NVLS(Rs("VALID_FROM"), ""))
         VALID_TO = InternalDateToDate(NVLS(Rs("VALID_TO"), ""))
         STKCOD = NVLS(Rs("STKCOD"), "")
End If

End Function
Public Function AddEditData()
Dim ServerDateTime As String
Dim SQL1 As String

   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   
 '  ENTERPRISE_CODE = glbCompanyCode
   
   If ShowMode = SHOW_ADD Then

               SQL1 = "INSERT INTO COMDONSTK " & _
                              "( " & _
                              "   STKCOD, " & _
                              "   STKDES, " & _
                               "   MASTER_COMDONSTK_ID, " & _
                                "   FLAG " & _
                              ") " & _
                              "VALUES " & _
                              "( " & _
                                 "'" & ChangeQuote(Trim(STKCOD)) & "'," & _
                                 "'" & ChangeQuote(Trim(STKDES)) & "'," & _
                                  "" & ChangeQuote(Trim(MASTER_COMDONSTK_ID)) & "," & _
                                  "'" & ChangeQuote(Trim(Flag)) & "' )"
                             

   Else 'Show edit

               SQL1 = "UPDATE COMDONSTK SET " & _
                           "   STKCOD = '" & ChangeQuote(Trim(STKCOD)) & "', " & _
                           "   STKDES = '" & ChangeQuote(Trim(STKDES)) & "', " & _
                           "   MASTER_COMDONSTK_ID = " & ChangeQuote(Trim(MASTER_COMDONSTK_ID)) & ", " & _
                           "   FLAG = '" & ChangeQuote(Trim(Flag)) & "' " & _
                           "WHERE COMDONSTK_ID = " & COMDONSTK_ID
                           
'
   End If
   
   If (ShowMode = SHOW_ADD) Or (ShowMode = SHOW_EDIT) Then
      'debug.print SQL1
      m_Conn.Execute (SQL1)
   End If
End Function

Public Function DeleteData(Optional Ind As Long = 1)
Dim SQL1 As String

If Ind = 1 Then
      SQL1 = "DELETE FROM COMDONSTK WHERE COMDONSTK_ID = " & COMDONSTK_ID
ElseIf Ind = 2 Then
      SQL1 = "DELETE FROM COMDONSTK WHERE MASTER_COMDONSTK_ID = " & MASTER_COMDONSTK_ID
End If
      m_Conn.Execute (SQL1)
End Function


