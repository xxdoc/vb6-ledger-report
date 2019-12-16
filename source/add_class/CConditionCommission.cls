VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CConditionCommission"
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

Public COM_ID As Long
Public COMTYP As String
Public YEAR_ID As Long
Public STKCOD As String
Public STKDES As String
Public SLM_PERCENT As String 'Double
Public Flag As String

Public NUM_ONE As Double 'String
Public OPERATOR As String
Public NUM_TWO As Double 'String

Public FROM_DATE As Date
Public TO_DATE As Date
Public YEARNUM As String
Public GROUP1 As Double
Public FROM_CMPL_DATE As Date
Public TO_CMPL_DATE As Date
Public INCEN_CR As Long

Public GOODS_GROUP_ID As Long
Public GOODS_MASTER_ID As Long

Public ORDER_BY As Long
Public ORDER_TYPE  As Long


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
      SelectStr1 = "SELECT CS.* , CSY.YEARNUM ,CSY.FROM_DATE, CSY.TO_DATE , CSY.INCEN_CR , CSY.GOODS_MASTER_ID , CSY.GOODS_GROUP_ID "             '**********
      SelectStr2 = "SELECT COUNT(*)  "
      
      FromStr = "FROM COMMISSION CS "
      FromStr = FromStr & "LEFT OUTER JOIN COMMISSION_YEAR CSY ON (CS.YEAR_ID = CSY.YEAR_ID) "
  ElseIf Ind = 2 Then            ' สำหรับแต่ละบรรทัดของ Grid
     SelectStr1 = "SELECT CS.* "
      SelectStr2 = "SELECT COUNT(*)  "
      
      FromStr = "FROM COMMISSION CS "
   ElseIf Ind = 3 Then             'สำหรับ ปีไหน มีข้อมูลอะไร Tab1
      SelectStr1 = "SELECT CS.* , CSY.YEARNUM ,CSY.FROM_DATE, CSY.TO_DATE , CSY.INCEN_CR , CSY.GOODS_MASTER_ID , CSY.GOODS_GROUP_ID "             '**********
      SelectStr2 = "SELECT COUNT(*)  "
      
      FromStr = "FROM COMMISSION CS "
      FromStr = FromStr & "LEFT OUTER JOIN COMMISSION_YEAR CSY ON (CS.YEAR_ID = CSY.YEAR_ID) "
   ElseIf Ind = 4 Then             'สำหรับ
      SelectStr1 = "SELECT CS.* , CSY.YEARNUM ,CSY.FROM_DATE, CSY.TO_DATE , CSY.INCEN_CR "             '**********
      SelectStr2 = "SELECT COUNT(*)  "
      
      FromStr = "FROM COMMISSION CS "
      FromStr = FromStr & "LEFT OUTER JOIN COMMISSION_YEAR CSY ON (CS.YEAR_ID = CSY.YEAR_ID) "
   ElseIf Ind = 5 Then             'สำหรับ
      SelectStr1 = "SELECT DISTINCT CS.STKCOD, CS.STKDES , CSY.YEARNUM ,CSY.FROM_DATE, CSY.TO_DATE , CSY.INCEN_CR "             '**********
      SelectStr2 = "SELECT COUNT(*)  "
      
      FromStr = "FROM COMMISSION CS "
      FromStr = FromStr & "LEFT OUTER JOIN COMMISSION_YEAR CSY ON (CS.YEAR_ID = CSY.YEAR_ID) "
     ElseIf Ind = 6 Then             'สำหรับ GROUP1
      SelectStr1 = "SELECT DISTINCT CS.STKCOD, CS.GROUP1, CSY.YEARNUM ,CSY.FROM_DATE, CSY.TO_DATE , CSY.INCEN_CR "             '**********
      SelectStr2 = "SELECT COUNT(*)  "
      
      FromStr = "FROM COMMISSION CS "
      FromStr = FromStr & "LEFT OUTER JOIN COMMISSION_YEAR CSY ON (CS.YEAR_ID = CSY.YEAR_ID) "
End If
      
If YEAR_ID <> 0 Then
      WhereStr = "WHERE (CS.YEAR_ID=" & ChangeQuote(Trim(YEAR_ID)) & ") "
End If
   
   If Trim(STKCOD) <> "" Then
      If WhereStr = "" Then
         WhereStr = "WHERE (CS.STKCOD = '" & STKCOD & "') "
      Else
         WhereStr = WhereStr & "AND (CS.STKCOD = '" & STKCOD & "') "
      End If
   End If
   
   If Trim(COMTYP) <> "" Then
      If WhereStr = "" Then
         WhereStr = "WHERE (CS.COMTYP = '" & COMTYP & "') "
      Else
         WhereStr = WhereStr & "AND (CS.COMTYP = '" & COMTYP & "') "
      End If
   End If
         
   If GOODS_GROUP_ID > 0 Then
      If WhereStr = "" Then
         WhereStr = "WHERE (CSY.GOODS_GROUP_ID = " & GOODS_GROUP_ID & ") "
      Else
         WhereStr = WhereStr & "AND (CSY.GOODS_GROUP_ID = " & GOODS_GROUP_ID & ") "
      End If
   End If
   
  If COM_ID > 0 Then
      If WhereStr = "" Then
         WhereStr = "WHERE (CS.COM_ID = " & COM_ID & ") "
      Else
         WhereStr = WhereStr & "AND (CS.COM_ID = " & COM_ID & ") "
      End If
   End If
   
   'วันที่รับชำระ
   If FROM_CMPL_DATE > -1 Then                    ' จะเอาไว้ โหลดรีพอร์ต
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
   
   
   OrderByStr = ""
   OrderTypeStr = ""

 If Ind = 1 Then
        OrderByStr = "ORDER BY CS.GROUP1 ASC , CS.STKCOD ASC, CS.NUM_ONE DESC , CS.NUM_TWO "
 ElseIf Ind = 2 Then
        OrderByStr = "ORDER BY NUM_ONE "
 ElseIf Ind = 3 Then
        OrderByStr = "ORDER BY CS.NUM_ONE DESC, CS.NUM_TWO "
  ElseIf Ind = 4 Then
        OrderByStr = "ORDER BY CS.STKCOD "
  ElseIf Ind = 5 Then
        OrderByStr = "ORDER BY CS.STKCOD "
   ElseIf Ind = 6 Then
        OrderByStr = "ORDER BY CS.GROUP1 "
 End If
 
If Ind = 1 Then
        OrderTypeStr = "DESC "
ElseIf Ind = 3 Then
        OrderTypeStr = "ASC "
Else
         OrderTypeStr = "ASC "
 End If
   
   SQL1 = SelectStr1 & FromStr & WhereStr & OrderByStr & OrderTypeStr
   SQL2 = SelectStr2 & FromStr & WhereStr
 '  'debug.print SQL1
  Call glbDatabaseMngr.GetLegacyRs(SQL1, SQL2, True, ItemCount, Rs, glbErrorLog)
End Function

Public Function PopulateFromRS(Ind As Long, Rs As ADODB.Recordset)
If Ind = 1 Then
        COM_ID = NVLI(Rs("COM_ID"), -1)
        COMTYP = NVLS(Rs("COMTYP"), "")
         YEAR_ID = NVLI(Rs("YEAR_ID"), -1)
         STKCOD = NVLS(Rs("STKCOD"), "")
         STKDES = NVLS(Rs("STKDES"), "")
         SLM_PERCENT = NVLS(Rs("SLM_PERCENT"), "")
         YEARNUM = NVLS(Rs("YEARNUM"), "")
         Flag = NVLS(Rs("FLAG"), "")
              FROM_DATE = InternalDateToDate(NVLS(Rs("FROM_DATE"), ""))
              TO_DATE = InternalDateToDate(NVLS(Rs("TO_DATE"), ""))
              INCEN_CR = NVLI(Rs("INCEN_CR"), -1)

         GOODS_MASTER_ID = NVLI(Rs("GOODS_MASTER_ID"), -1)
         GOODS_GROUP_ID = NVLI(Rs("GOODS_GROUP_ID"), -1)

        NUM_ONE = NVLD(Rs("NUM_ONE"), -1)
        OPERATOR = NVLS(Rs("OPERATOR"), "")
        NUM_TWO = NVLD(Rs("NUM_TWO"), -1)
                  GROUP1 = NVLI(Rs("GROUP1"), -1)
ElseIf Ind = 2 Then
       COM_ID = NVLI(Rs("COM_ID"), -1)
        COMTYP = NVLS(Rs("COMTYP"), "")
         YEAR_ID = NVLI(Rs("YEAR_ID"), -1)
     NUM_ONE = NVLD(Rs("NUM_ONE"), -1)
        OPERATOR = NVLS(Rs("OPERATOR"), "")
        NUM_TWO = NVLD(Rs("NUM_TWO"), -1)
        SLM_PERCENT = NVLS(Rs("SLM_PERCENT"), "")
                  STKCOD = NVLS(Rs("STKCOD"), "")
                           STKDES = NVLS(Rs("STKDES"), "")
                           GROUP1 = NVLI(Rs("GROUP1"), -1)
                           Flag = NVLS(Rs("FLAG"), "")
ElseIf Ind = 5 Then
         YEARNUM = NVLS(Rs("YEARNUM"), "")
         FROM_DATE = InternalDateToDate(NVLS(Rs("FROM_DATE"), ""))
         TO_DATE = InternalDateToDate(NVLS(Rs("TO_DATE"), ""))
          INCEN_CR = NVLI(Rs("INCEN_CR"), -1)
         STKCOD = NVLS(Rs("STKCOD"), "")
         STKDES = NVLS(Rs("STKDES"), "")
ElseIf Ind = 6 Then
         YEARNUM = NVLS(Rs("YEARNUM"), "")
         FROM_DATE = InternalDateToDate(NVLS(Rs("FROM_DATE"), ""))
         TO_DATE = InternalDateToDate(NVLS(Rs("TO_DATE"), ""))
          INCEN_CR = NVLI(Rs("INCEN_CR"), -1)
         STKCOD = NVLS(Rs("STKCOD"), "")
         GROUP1 = NVLI(Rs("GROUP1"), -1)
End If

End Function
Public Function AddEditData()
Dim ServerDateTime As String
Dim SQL1 As String

   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   
 '  ENTERPRISE_CODE = glbCompanyCode
   
   If AddEditMode = SHOW_ADD Then

               SQL1 = "INSERT INTO COMMISSION " & _
                              "( " & _
                              "   NUM_ONE, " & _
                              "   OPERATOR, " & _
                               "   NUM_TWO, " & _
                              "   STKCOD, " & _
                              "   STKDES, " & _
                               "   SLM_PERCENT, " & _
                              "   COMTYP, " & _
                               "   YEAR_ID, " & _
                                "  GROUP1, " & _
                                "   FLAG " & _
                              ") " & _
                              "VALUES " & _
                              "( " & _
                                 "" & ChangeQuote(Trim(NUM_ONE)) & "," & _
                                 "'" & ChangeQuote(Trim(OPERATOR)) & "'," & _
                                  "" & ChangeQuote(Trim(NUM_TWO)) & "," & _
                                 "'" & ChangeQuote(Trim(STKCOD)) & "'," & _
                                 "'" & ChangeQuote(Trim(STKDES)) & "'," & _
                                  "'" & ChangeQuote(Trim(SLM_PERCENT)) & "'," & _
                                 "'" & ChangeQuote(Trim(COMTYP)) & "'," & _
                                  "" & ChangeQuote(Trim(YEAR_ID)) & "," & _
                                  "" & ChangeQuote(Trim(GROUP1)) & "," & " '" & ChangeQuote(Trim(Flag)) & "' )"
                             

   Else 'Show edit

               SQL1 = "UPDATE COMMISSION SET " & _
                           "   NUM_ONE = " & ChangeQuote(Trim(NUM_ONE)) & ", " & _
                           "   NUM_TWO = " & ChangeQuote(Trim(NUM_TWO)) & ", " & _
                           "   OPERATOR = '" & ChangeQuote(Trim(OPERATOR)) & "', " & _
                           "   STKCOD = '" & ChangeQuote(Trim(STKCOD)) & "', " & _
                           "   STKDES = '" & ChangeQuote(Trim(STKDES)) & "', " & _
                           "   SLM_PERCENT = '" & ChangeQuote(Trim(SLM_PERCENT)) & "', " & _
                           "   COMTYP = '" & ChangeQuote(Trim(COMTYP)) & "', " & _
                           "   YEAR_ID = " & ChangeQuote(Trim(YEAR_ID)) & ", " & _
                           "   GROUP1 = " & ChangeQuote(Trim(GROUP1)) & ", " & "   FLAG = '" & ChangeQuote(Trim(Flag)) & "' " & _
                           "WHERE COM_ID = " & COM_ID
                           
'
   End If
   
   If (AddEditMode = SHOW_ADD) Or (AddEditMode = SHOW_EDIT) Then
      'debug.print SQL1
      m_Conn.Execute (SQL1)
   End If
End Function

Public Function DeleteData(Optional Ind As Long = 1)
Dim SQL1 As String

If Ind = 1 Then
      SQL1 = "DELETE FROM COMMISSION WHERE COM_ID = " & COM_ID
ElseIf Ind = 2 Then
      SQL1 = "DELETE FROM COMMISSION WHERE YEAR_ID = " & YEAR_ID
End If
      m_Conn.Execute (SQL1)
End Function
