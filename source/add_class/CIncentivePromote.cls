VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CIncentivePromote"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
' �ҹ������ access
Private m_Conn As ADODB.Connection

Public AddEditMode As SHOW_MODE_TYPE
Public QueryFlag As Long
Public OrderBy As Long
Public OrderType As Long

' End of generic data member

Public INCENTIVE_PROMOTE_ID As Long
Public INCENTIVE_TYP As String
Public MASTER_INCENTIVE_ID As Long
Public STKCOD As String
Public STKDES As String
Public SLM_PERCENT As String 'Double
Public Flag As String

Public NUM_ONE As Double 'String
Public OPERATOR As String
Public NUM_TWO As Double 'String

Public VALID_FROM As Date
Public VALID_TO As Date
Public MASTER_INCENTIVE_NO As String
Public GROUP1 As Double
Public FROM_CMPL_DATE As Date
Public TO_CMPL_DATE As Date

Public ORDER_BY As Long
Public ORDER_TYPE  As Long

Public ShowMode As SHOW_MODE_TYPE


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
   
   If Ind = 1 Then            '����Ѻ ���˹ �բ��������� Tab1
      SelectStr1 = "SELECT CS.* , CSY.MASTER_INCENTIVE_NO ,CSY.VALID_FROM, CSY.VALID_TO "             '**********
      SelectStr2 = "SELECT COUNT(*)  "
      
      FromStr = "FROM INCENTIVE_PROMOTE CS "
      FromStr = FromStr & "LEFT OUTER JOIN MASTER_INCENTIVE_PROMOTE CSY ON (CS.MASTER_INCENTIVE_ID = CSY.MASTER_INCENTIVE_ID) "
  ElseIf Ind = 2 Then            ' ����Ѻ���к�÷Ѵ�ͧ Grid
     SelectStr1 = "SELECT CS.* "
      SelectStr2 = "SELECT COUNT(*)  "
      
      FromStr = "FROM INCENTIVE_PROMOTE CS "
   ElseIf Ind = 3 Then             '����Ѻ ���˹ �բ��������� Tab1
      SelectStr1 = "SELECT CS.* , CSY.MASTER_INCENTIVE_NO ,CSY.VALID_FROM, CSY.VALID_TO "             '**********
      SelectStr2 = "SELECT COUNT(*)  "
      
      FromStr = "FROM INCENTIVE_PROMOTE CS "
      FromStr = FromStr & "LEFT OUTER JOIN MASTER_INCENTIVE_PROMOTE CSY ON (CS.MASTER_INCENTIVE_ID = CSY.MASTER_INCENTIVE_ID) "
   ElseIf Ind = 4 Then             '����Ѻ
      SelectStr1 = "SELECT CS.* , CSY.MASTER_INCENTIVE_NO ,CSY.VALID_FROM, CSY.VALID_TO "             '**********
      SelectStr2 = "SELECT COUNT(*)  "
      
      FromStr = "FROM INCENTIVE_PROMOTE CS "
      FromStr = FromStr & "LEFT OUTER JOIN MASTER_INCENTIVE_PROMOTE CSY ON (CS.MASTER_INCENTIVE_ID = CSY.MASTER_INCENTIVE_ID) "
   ElseIf Ind = 5 Then             '����Ѻ
      SelectStr1 = "SELECT DISTINCT CS.STKCOD, CS.STKDES , CSY.MASTER_INCENTIVE_NO ,CSY.VALID_FROM, CSY.VALID_TO "             '**********
      SelectStr2 = "SELECT COUNT(*)  "
      
      FromStr = "FROM INCENTIVE_PROMOTE CS "
      FromStr = FromStr & "LEFT OUTER JOIN MASTER_INCENTIVE_PROMOTE CSY ON (CS.MASTER_INCENTIVE_ID = CSY.MASTER_INCENTIVE_ID) "
     ElseIf Ind = 6 Then             '����Ѻ GROUP1
      SelectStr1 = "SELECT DISTINCT CS.STKCOD, CS.GROUP1, CSY.MASTER_INCENTIVE_NO ,CSY.VALID_FROM, CSY.VALID_TO  "             '**********
      SelectStr2 = "SELECT COUNT(*)  "
      
      FromStr = "FROM INCENTIVE_PROMOTE CS "
      FromStr = FromStr & "LEFT OUTER JOIN MASTER_INCENTIVE_PROMOTE CSY ON (CS.MASTER_INCENTIVE_ID = CSY.MASTER_INCENTIVE_ID) "
End If
      
If MASTER_INCENTIVE_ID <> 0 Then
      WhereStr = "WHERE (CS.MASTER_INCENTIVE_ID=" & ChangeQuote(Trim(MASTER_INCENTIVE_ID)) & ") "
End If
   
   If Trim(STKCOD) <> "" Then
      If WhereStr = "" Then
         WhereStr = "WHERE (CS.STKCOD = '" & STKCOD & "') "
      Else
         WhereStr = WhereStr & "AND (CS.STKCOD = '" & STKCOD & "') "
      End If
   End If
   
   If Trim(INCENTIVE_TYP) <> "" Then
      If WhereStr = "" Then
         WhereStr = "WHERE (CS.INCENTIVE_TYP = '" & INCENTIVE_TYP & "') "
      Else
         WhereStr = WhereStr & "AND (CS.INCENTIVE_TYP = '" & INCENTIVE_TYP & "') "
      End If
   End If
         
    If INCENTIVE_PROMOTE_ID > 0 Then
      If WhereStr = "" Then
         WhereStr = "WHERE (CS.INCENTIVE_PROMOTE_ID = " & INCENTIVE_PROMOTE_ID & ") "
      Else
         WhereStr = WhereStr & "AND (CS.INCENTIVE_PROMOTE_ID = " & INCENTIVE_PROMOTE_ID & ") "
      End If
   End If
   
   '�ѹ����Ѻ����
   If FROM_CMPL_DATE > -1 Then                     ' �������� ��Ŵ�վ���
      TempDate = DateToStringIntLow(Trim(FROM_CMPL_DATE))
      If WhereStr = "" Then
         WhereStr = "WHERE (CSY.VALID_FROM <= '" & ChangeQuote(TempDate) & "') "  ' ����ͧ���¨еç�ѹ�����ҡ�á�� ���е�ͧ�������ͺ����
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
        OrderByStr = "ORDER BY CS.GROUP1 ASC , CS.STKCOD ASC, CS.NUM_ONE "
 ElseIf Ind = 2 Then
        OrderByStr = "ORDER BY NUM_ONE "
 ElseIf Ind = 3 Then
        OrderByStr = "ORDER BY CS.NUM_ONE "
  ElseIf Ind = 4 Then
        OrderByStr = "ORDER BY CS.STKCOD "
  ElseIf Ind = 5 Then
        OrderByStr = "ORDER BY CS.STKCOD "
   ElseIf Ind = 6 Then
        OrderByStr = "ORDER BY CS.GROUP1 "
 End If
 
If Ind = 3 Then
        OrderTypeStr = "ASC "
Else
         OrderTypeStr = "ASC "
 End If
   
   SQL1 = SelectStr1 & FromStr & WhereStr & OrderByStr & OrderTypeStr
   SQL2 = SelectStr2 & FromStr & WhereStr
   'debug.print SQL1
  Call glbDatabaseMngr.GetLegacyRs(SQL1, SQL2, True, ItemCount, Rs, glbErrorLog)
End Function

Public Function PopulateFromRS(Ind As Long, Rs As ADODB.Recordset)
If Ind = 1 Then
        INCENTIVE_PROMOTE_ID = NVLI(Rs("INCENTIVE_PROMOTE_ID"), -1)
        INCENTIVE_TYP = NVLS(Rs("INCENTIVE_TYP"), "")
         MASTER_INCENTIVE_ID = NVLI(Rs("MASTER_INCENTIVE_ID"), -1)
         STKCOD = NVLS(Rs("STKCOD"), "")
         STKDES = NVLS(Rs("STKDES"), "")
         SLM_PERCENT = NVLS(Rs("SLM_PERCENT"), "")
         MASTER_INCENTIVE_NO = NVLS(Rs("MASTER_INCENTIVE_NO"), "")
         
              VALID_FROM = InternalDateToDate(NVLS(Rs("VALID_FROM"), ""))
              VALID_TO = InternalDateToDate(NVLS(Rs("VALID_TO"), ""))
'         VALID_FROM = InternalDateToDate(NVLS(Rs("VALID_FROM"), ""))
'         VALID_TO = InternalDateToDate(NVLS(Rs("VALID_TO"), ""))

        NUM_ONE = NVLI(Rs("NUM_ONE"), -1)
        OPERATOR = NVLS(Rs("OPERATOR"), "")
        NUM_TWO = NVLI(Rs("NUM_TWO"), -1)
                  GROUP1 = NVLI(Rs("GROUP1"), -1)
ElseIf Ind = 2 Then
       INCENTIVE_PROMOTE_ID = NVLI(Rs("INCENTIVE_PROMOTE_ID"), -1)
        INCENTIVE_TYP = NVLS(Rs("INCENTIVE_TYP"), "")
         MASTER_INCENTIVE_ID = NVLI(Rs("MASTER_INCENTIVE_ID"), -1)
     NUM_ONE = NVLI(Rs("NUM_ONE"), -1)
        OPERATOR = NVLS(Rs("OPERATOR"), "")
        NUM_TWO = NVLI(Rs("NUM_TWO"), -1)
        SLM_PERCENT = NVLS(Rs("SLM_PERCENT"), "")
                  STKCOD = NVLS(Rs("STKCOD"), "")
                           GROUP1 = NVLI(Rs("GROUP1"), -1)
ElseIf Ind = 5 Then
         MASTER_INCENTIVE_NO = NVLS(Rs("MASTER_INCENTIVE_NO"), "")
         VALID_FROM = InternalDateToDate(NVLS(Rs("VALID_FROM"), ""))
         VALID_TO = InternalDateToDate(NVLS(Rs("VALID_TO"), ""))
         STKCOD = NVLS(Rs("STKCOD"), "")
         STKDES = NVLS(Rs("STKDES"), "")
ElseIf Ind = 6 Then
         MASTER_INCENTIVE_NO = NVLS(Rs("MASTER_INCENTIVE_NO"), "")
         VALID_FROM = InternalDateToDate(NVLS(Rs("VALID_FROM"), ""))
         VALID_TO = InternalDateToDate(NVLS(Rs("VALID_TO"), ""))
         STKCOD = NVLS(Rs("STKCOD"), "")
         GROUP1 = NVLI(Rs("GROUP1"), -1)
End If

End Function
Public Function AddEditData()
Dim ServerDateTime As String
Dim SQL1 As String

   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   
 '  ENTERPRISE_CODE = glbCompanyCode
   
   If ShowMode = SHOW_ADD Then

               SQL1 = "INSERT INTO INCENTIVE_PROMOTE " & _
                              "( " & _
                              "   NUM_ONE, " & _
                              "   OPERATOR, " & _
                               "   NUM_TWO, " & _
                              "   STKCOD, " & _
                              "   STKDES, " & _
                               "   SLM_PERCENT, " & _
                              "   INCENTIVE_TYP, " & _
                               "   MASTER_INCENTIVE_ID, " & _
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
                                 "'" & ChangeQuote(Trim(INCENTIVE_TYP)) & "'," & _
                                  "" & ChangeQuote(Trim(MASTER_INCENTIVE_ID)) & "," & _
                                  "" & ChangeQuote(Trim(GROUP1)) & "," & " '" & ChangeQuote(Trim(Flag)) & "' )"
                             

   Else 'Show edit

               SQL1 = "UPDATE INCENTIVE_PROMOTE SET " & _
                           "   NUM_ONE = " & ChangeQuote(Trim(NUM_ONE)) & ", " & _
                           "   NUM_TWO = " & ChangeQuote(Trim(NUM_TWO)) & ", " & _
                           "   OPERATOR = '" & ChangeQuote(Trim(OPERATOR)) & "', " & _
                           "   STKCOD = '" & ChangeQuote(Trim(STKCOD)) & "', " & _
                           "   STKDES = '" & ChangeQuote(Trim(STKDES)) & "', " & _
                           "   SLM_PERCENT = '" & ChangeQuote(Trim(SLM_PERCENT)) & "', " & _
                           "   INCENTIVE_TYP = '" & ChangeQuote(Trim(INCENTIVE_TYP)) & "', " & _
                           "   MASTER_INCENTIVE_ID = " & ChangeQuote(Trim(MASTER_INCENTIVE_ID)) & ", " & _
                           "   GROUP1 = " & ChangeQuote(Trim(GROUP1)) & ", " & "   FLAG = '" & ChangeQuote(Trim(Flag)) & "' " & _
                           "WHERE INCENTIVE_PROMOTE_ID = " & INCENTIVE_PROMOTE_ID
                           
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
      SQL1 = "DELETE FROM INCENTIVE_PROMOTE WHERE INCENTIVE_PROMOTE_ID = " & INCENTIVE_PROMOTE_ID
ElseIf Ind = 2 Then
      SQL1 = "DELETE FROM INCENTIVE_PROMOTE WHERE MASTER_INCENTIVE_ID = " & MASTER_INCENTIVE_ID
End If
      m_Conn.Execute (SQL1)
End Function


