VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CCondiIncenSum"
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

Public INCEN_SLM_FORSUM_ID As Long
Public FORSUM_TYP As String
Public YEAR_ID As Long
Public SLMCOD As String
Public SLMNAME As String
Public Flag As String

Public FROM_DATE As Date
Public TO_DATE As Date
Public YEARNUM As String
Public FROM_CMPL_DATE As Date
Public TO_CMPL_DATE As Date

Public ORDER_BY As Long
Public ORDER_TYPE  As Long
Public GOODS_GROUP_ID As Long

Private Sub Class_Initialize()
   Set m_Conn = glbDatabaseMngr.DBLegacyConnection
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
      SelectStr1 = "SELECT CS.* , CSY.YEARNUM ,CSY.FROM_DATE, CSY.TO_DATE "             '**********
      SelectStr2 = "SELECT COUNT(*)  "
      
      FromStr = "FROM INCEN_SLM_FORSUM CS "
      FromStr = FromStr & "LEFT OUTER JOIN COMMISSION_YEAR CSY ON (CS.YEAR_ID = CSY.YEAR_ID) "
  ElseIf Ind = 2 Then            ' สำหรับแต่ละบรรทัดของ Grid
     SelectStr1 = "SELECT CS.* "
      SelectStr2 = "SELECT COUNT(*)  "
      
      FromStr = "FROM INCEN_SLM_FORSUM CS "
   ElseIf Ind = 3 Then             'สำหรับ ปีไหน มีข้อมูลอะไร Tab1
      SelectStr1 = "SELECT CS.* , CSY.YEARNUM ,CSY.FROM_DATE, CSY.TO_DATE "             '**********
      SelectStr2 = "SELECT COUNT(*)  "
      
      FromStr = "FROM INCEN_SLM_FORSUM CS "
      FromStr = FromStr & "LEFT OUTER JOIN COMMISSION_YEAR CSY ON (CS.YEAR_ID = CSY.YEAR_ID) "
   ElseIf Ind = 4 Then             'สำหรับ
      SelectStr1 = "SELECT CS.* , CSY.YEARNUM ,CSY.FROM_DATE, CSY.TO_DATE "             '**********
      SelectStr2 = "SELECT COUNT(*)  "
      
      FromStr = "FROM INCEN_SLM_FORSUM CS "
      FromStr = FromStr & "LEFT OUTER JOIN COMMISSION_YEAR CSY ON (CS.YEAR_ID = CSY.YEAR_ID) "
   ElseIf Ind = 5 Then             'สำหรับ
      SelectStr1 = "SELECT DISTINCT CS.SLMCOD, CS.SLMNAME , CSY.YEARNUM ,CSY.FROM_DATE, CSY.TO_DATE "             '**********
      SelectStr2 = "SELECT COUNT(*)  "
      
      FromStr = "FROM INCEN_SLM_FORSUM CS "
      FromStr = FromStr & "LEFT OUTER JOIN COMMISSION_YEAR CSY ON (CS.YEAR_ID = CSY.YEAR_ID) "
     ElseIf Ind = 6 Then             'สำหรับ GROUP1
      SelectStr1 = "SELECT DISTINCT CS.SLMCOD, CSY.YEARNUM ,CSY.FROM_DATE, CSY.TO_DATE  "             '**********
      SelectStr2 = "SELECT COUNT(*)  "
      
      FromStr = "FROM INCEN_SLM_FORSUM CS "
      FromStr = FromStr & "LEFT OUTER JOIN COMMISSION_YEAR CSY ON (CS.YEAR_ID = CSY.YEAR_ID) "
End If
      
If YEAR_ID > 0 Then
      WhereStr = "WHERE (CS.YEAR_ID=" & ChangeQuote(Trim(YEAR_ID)) & ") "
End If
   
   If Trim(SLMCOD) <> "" Then
      If WhereStr = "" Then
         WhereStr = "WHERE (CS.SLMCOD = '" & SLMCOD & "') "
      Else
         WhereStr = WhereStr & "AND (CS.SLMCOD = '" & SLMCOD & "') "
      End If
   End If
   
   If Trim(FORSUM_TYP) <> "" Then
      If WhereStr = "" Then
         WhereStr = "WHERE (CS.FORSUM_TYP = '" & FORSUM_TYP & "') "
      Else
         WhereStr = WhereStr & "AND (CS.FORSUM_TYP = '" & FORSUM_TYP & "') "
      End If
   End If
         
    If INCEN_SLM_FORSUM_ID > 0 Then
      If WhereStr = "" Then
         WhereStr = "WHERE (CS.INCEN_SLM_FORSUM_ID = " & INCEN_SLM_FORSUM_ID & ") "
      Else
         WhereStr = WhereStr & "AND (CS.INCEN_SLM_FORSUM_ID = " & INCEN_SLM_FORSUM_ID & ") "
      End If
   End If
   
   'วันที่รับชำระ
   If FROM_CMPL_DATE > -1 Then                     ' จะเอาไว้ โหลดรีพอร์ต
      TempDate = DateToStringIntLow(Trim(FROM_CMPL_DATE))
      If WhereStr = "" Then
         WhereStr = "WHERE (CSY.FROM_DATE <= '" & ChangeQuote(TempDate) & "') "  ' เครื่องหมายจะตรงกันข้ามจากปรกติ เพราะต้องการให้ครอบครุม
      Else
         WhereStr = WhereStr & "AND (CSY.FROM_DATE <= '" & ChangeQuote(TempDate) & "') "
      End If
   End If

   If TO_CMPL_DATE > -1 Then
      TempDate = DateToStringIntHi(Trim(TO_CMPL_DATE))
      If WhereStr = "" Then
         WhereStr = "WHERE (CSY.TO_DATE >= '" & ChangeQuote(TempDate) & "') "
      Else
         WhereStr = WhereStr & "AND (CSY.TO_DATE >= '" & ChangeQuote(TempDate) & "') "
      End If
   End If
   
  If GOODS_GROUP_ID > 0 Then
      If WhereStr = "" Then
         WhereStr = "WHERE (CSY.GOODS_GROUP_ID = " & GOODS_GROUP_ID & ") "
      Else
         WhereStr = WhereStr & "AND (CSY.GOODS_GROUP_ID = " & GOODS_GROUP_ID & ") "
      End If
   End If
   
   OrderByStr = ""
   OrderTypeStr = ""

 If Ind = 1 Then
        OrderByStr = "ORDER BY CS.SLMCOD "
  Else
        OrderByStr = "ORDER BY CS.SLMCOD "
 End If
 
If Ind = 1 Then
        OrderTypeStr = "ASC "
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
        INCEN_SLM_FORSUM_ID = NVLI(Rs("INCEN_SLM_FORSUM_ID"), -1)
        FORSUM_TYP = NVLS(Rs("FORSUM_TYP"), "")
         YEAR_ID = NVLI(Rs("YEAR_ID"), -1)
         SLMCOD = NVLS(Rs("SLMCOD"), "")
         SLMNAME = NVLS(Rs("SLMNAME"), "")
         YEARNUM = NVLS(Rs("YEARNUM"), "")
         
              FROM_DATE = InternalDateToDate(NVLS(Rs("FROM_DATE"), ""))
              TO_DATE = InternalDateToDate(NVLS(Rs("TO_DATE"), ""))
              Flag = NVLS(Rs("FLAG"), "")
ElseIf Ind = 2 Then
       INCEN_SLM_FORSUM_ID = NVLI(Rs("INCEN_SLM_FORSUM_ID"), -1)
        FORSUM_TYP = NVLS(Rs("FORSUM_TYP"), "")
         YEAR_ID = NVLI(Rs("YEAR_ID"), -1)
                  SLMCOD = NVLS(Rs("SLMCOD"), "")
                  SLMNAME = NVLS(Rs("SLMNAME"), "")
                  Flag = NVLS(Rs("FLAG"), "")
ElseIf Ind = 5 Then
         YEARNUM = NVLS(Rs("YEARNUM"), "")
         FROM_DATE = InternalDateToDate(NVLS(Rs("FROM_DATE"), ""))
         TO_DATE = InternalDateToDate(NVLS(Rs("TO_DATE"), ""))
         SLMCOD = NVLS(Rs("SLMCOD"), "")
         SLMNAME = NVLS(Rs("SLMNAME"), "")
         Flag = NVLS(Rs("FLAG"), "")
ElseIf Ind = 6 Then
         YEARNUM = NVLS(Rs("YEARNUM"), "")
         FROM_DATE = InternalDateToDate(NVLS(Rs("FROM_DATE"), ""))
         TO_DATE = InternalDateToDate(NVLS(Rs("TO_DATE"), ""))
         SLMCOD = NVLS(Rs("SLMCOD"), "")
         Flag = NVLS(Rs("FLAG"), "")
End If

End Function
Public Function AddEditData()
Dim ServerDateTime As String
Dim SQL1 As String

   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   
 '  ENTERPRISE_CODE = glbCompanyCode
   
   If AddEditMode = SHOW_ADD Then

               SQL1 = "INSERT INTO INCEN_SLM_FORSUM " & _
                              "( " & _
                              "   SLMCOD, " & _
                              "   SLMNAME, " & _
                              "   FORSUM_TYP, " & _
                               "   YEAR_ID, " & _
                                "   FLAG " & _
                              ") " & _
                              "VALUES " & _
                              "( " & _
                                 "'" & ChangeQuote(Trim(SLMCOD)) & "'," & _
                                 "'" & ChangeQuote(Trim(SLMNAME)) & "'," & _
                                 "'" & ChangeQuote(Trim(FORSUM_TYP)) & "'," & _
                                  "" & ChangeQuote(Trim(YEAR_ID)) & "," & _
                                  " '" & ChangeQuote(Trim(Flag)) & "' )"
                           
   Else 'Show edit

               SQL1 = "UPDATE INCEN_SLM_FORSUM SET " & _
                           "   SLMCOD = '" & ChangeQuote(Trim(SLMCOD)) & "', " & _
                           "   SLMNAME = '" & ChangeQuote(Trim(SLMNAME)) & "', " & _
                           "   FORSUM_TYP = '" & ChangeQuote(Trim(FORSUM_TYP)) & "', " & _
                           "   YEAR_ID = " & ChangeQuote(Trim(YEAR_ID)) & ", " & _
                           "   FLAG = '" & ChangeQuote(Trim(Flag)) & "' " & _
                           "WHERE INCEN_SLM_FORSUM_ID = " & INCEN_SLM_FORSUM_ID

   End If
   
   If (AddEditMode = SHOW_ADD) Or (AddEditMode = SHOW_EDIT) Then
   '   'debug.print SQL1
      m_Conn.Execute (SQL1)
   End If
End Function

Public Function DeleteData(Optional Ind As Long = 1)
Dim SQL1 As String

If Ind = 1 Then
      SQL1 = "DELETE FROM INCEN_SLM_FORSUM WHERE INCEN_SLM_FORSUM_ID = " & INCEN_SLM_FORSUM_ID
ElseIf Ind = 2 Then
      SQL1 = "DELETE FROM COMMISSION_YEAR WHERE YEAR_ID = " & YEAR_ID
End If
      m_Conn.Execute (SQL1)
End Function
