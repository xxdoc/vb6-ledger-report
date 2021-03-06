VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CMasterRef"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_ADOConn As ADODB.Connection

Public ShowMode As SHOW_MODE_TYPE
Public Flag As String
Public QueryFlag As Long

Public KEY_ID As Long
Public KEY_CODE As String
Public KEY_NAME As String
Public MASTER_AREA As Long
Public ORDER_BY As Long
Public ORDER_TYPE As Long

Public PARENT_ID As Long
Public PARENT_EX_ID As Long            'LINK GO TO EMPLOYEE
Public PARENT_EX_ID2 As Long           ' LINK GO TO APAR_MAS
Public PARENT_EX_ID3 As Long
Public PARENT_EX_ID4 As Long
Public PARENT_EX_ID5 As Long
Public MASTER_FLAG As String
Public KEY_LOOKUP As String
Public APAR_CODE As String
Public EMP_CODE As String
Public INDEX_LINK As Long
Public SHORT_CODE  As String
Public FROM_LOCATION_ID As String
Public TO_LOCATION_ID As String

Public APAR_NAME As String
Public MasterRefDetails As Collection
Private Sub Class_Initialize()
   Set m_ADOConn = glbDatabaseMngr.DBConnection
   
   Set MasterRefDetails = New Collection
End Sub
Private Sub Class_Terminate()
   Set MasterRefDetails = Nothing
End Sub
Public Sub DeleteData()
Dim SQL As String
   
   SQL = "DELETE FROM MASTER_REF WHERE (KEY_ID = " & KEY_ID & ")"
   Call m_ADOConn.Execute(SQL)
End Sub
Public Sub QueryData(Ind As Long, Rs As ADODB.Recordset, ItemCount As Long, Optional MustCount As Boolean = False)
Dim SQL1 As String
Dim SQL2 As String
Dim SelectStr1 As String
Dim SelectStr2 As String
Dim FromStr As String
Dim WhereStr As String
Dim GroupByStr As String
Dim OrderByStr As String
Dim TempStr As String
   
   If Ind = 1 Then
      SelectStr1 = "SELECT MR.* "
      SelectStr1 = SelectStr1 & ",APM.APAR_CODE "
      SelectStr1 = SelectStr1 & ",EM.EMP_CODE "
      SelectStr2 = "SELECT COUNT(*) "
      
      FromStr = "FROM MASTER_REF MR "
      
      FromStr = FromStr & "LEFT OUTER JOIN APAR_MAS APM ON (MR.PARENT_EX_ID2 = APM.APAR_MAS_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN EMPLOYEE EM ON (MR.PARENT_EX_ID = EM.EMP_ID) "
      
      WhereStr = ""
   ElseIf Ind = 2 Then     'Patch_2007_03_09_1_jill
      SelectStr1 = "SELECT MR.PARENT_EX_ID2,MR.KEY_ID "
      SelectStr2 = "SELECT COUNT(*) "

      FromStr = "FROM MASTER_REF MR "
      
      TempStr = "WHERE "
      
      WhereStr = TempStr & " (SELECT COUNT(MR1.PARENT_EX_ID2) FROM MASTER_REF MR1 WHERE MR1.MASTER_AREA = 21 AND MR1.PARENT_EX_ID2 = MR.PARENT_EX_ID2 GROUP BY MR1.PARENT_EX_ID2) = 1 "
   ElseIf Ind = 5 Then  'CReportMaster001
      SelectStr1 = "SELECT MR.* "
      SelectStr2 = "SELECT COUNT(*) "
      
      FromStr = "FROM MASTER_REF MR "
      
      WhereStr = ""
   ElseIf Ind = 6 Then
      SelectStr1 = "SELECT DISTINCT MR.PARENT_EX_ID2 , APM.APAR_CODE, NM2.LONG_NAME APAR_NAME "
      SelectStr2 = "SELECT COUNT(*) "
      
      FromStr = "FROM MASTER_REF MR "
      FromStr = FromStr & "LEFT OUTER JOIN APAR_MAS APM ON (MR.PARENT_EX_ID2 = APM.APAR_MAS_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN APAR_NAME EN ON (APM.APAR_MAS_ID = EN.APAR_MAS_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN NAME NM2 ON (EN.NAME_ID = NM2.NAME_ID) "
      
      
      WhereStr = ""
   End If
   
   If KEY_ID > 0 Then
      If Len(WhereStr) > 0 Then
         TempStr = "AND "
      Else
         TempStr = "WHERE "
      End If
      WhereStr = TempStr & " (MR.KEY_ID = " & KEY_ID & ") " & vbCrLf
   End If
   
   If MASTER_AREA > 0 Then
      If Len(WhereStr) > 0 Then
         TempStr = "AND "
      Else
         TempStr = "WHERE "
      End If
      WhereStr = WhereStr & TempStr & " (MR.MASTER_AREA = " & MASTER_AREA & ") " & vbCrLf
   End If
   
   If PARENT_ID > 0 Then
      If Len(WhereStr) > 0 Then
         TempStr = "AND "
      Else
         TempStr = "WHERE "
      End If
      WhereStr = WhereStr & TempStr & " (MR.PARENT_ID = " & PARENT_ID & ") " & vbCrLf
   End If
   
   If PARENT_EX_ID > 0 Then
      If Len(WhereStr) > 0 Then
         TempStr = "AND "
      Else
         TempStr = "WHERE "
      End If
      WhereStr = WhereStr & TempStr & " (MR.PARENT_EX_ID = " & PARENT_EX_ID & ") " & vbCrLf
   End If
   
   If PARENT_EX_ID2 > 0 Then
      If Len(WhereStr) > 0 Then
         TempStr = "AND "
      Else
         TempStr = "WHERE "
      End If
      WhereStr = WhereStr & TempStr & " (MR.PARENT_EX_ID2 = " & PARENT_EX_ID2 & ") " & vbCrLf
   End If
   
   If PARENT_EX_ID3 > 0 Then
      If Len(WhereStr) > 0 Then
         TempStr = "AND "
      Else
         TempStr = "WHERE "
      End If
      WhereStr = WhereStr & TempStr & " (MR.PARENT_EX_ID3 = " & PARENT_EX_ID3 & ") " & vbCrLf
   End If
   
   If PARENT_EX_ID4 > 0 Then
      If Len(WhereStr) > 0 Then
         TempStr = "AND "
      Else
         TempStr = "WHERE "
      End If
      WhereStr = WhereStr & TempStr & " (MR.PARENT_EX_ID4 = " & PARENT_EX_ID4 & ") " & vbCrLf
   End If
   
   If PARENT_EX_ID5 > 0 Then
      If Len(WhereStr) > 0 Then
         TempStr = "AND "
      Else
         TempStr = "WHERE "
      End If
      WhereStr = WhereStr & TempStr & " (MR.PARENT_EX_ID5 = " & PARENT_EX_ID5 & ") " & vbCrLf
   End If
   
   If INDEX_LINK > 0 Then
      If Len(WhereStr) > 0 Then
         TempStr = "AND "
      Else
         TempStr = "WHERE "
      End If
      WhereStr = WhereStr & TempStr & " (MR.INDEX_LINK = " & INDEX_LINK & ") " & vbCrLf
   End If
   
   If Len(Trim(MASTER_FLAG)) > 0 Then
      If Len(WhereStr) > 0 Then
         TempStr = "AND "
      Else
         TempStr = "WHERE "
      End If
      WhereStr = WhereStr & TempStr & " (MR.MASTER_FLAG = '" & ChangeQuote(Trim(MASTER_FLAG)) & "')"
   End If
   
   If Len(FROM_LOCATION_ID) > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (MR.KEY_CODE >= '" & Trim(ChangeQuote(FROM_LOCATION_ID)) & "') "
      Else
         WhereStr = WhereStr & "AND (MR.KEY_CODE >= '" & Trim(ChangeQuote(FROM_LOCATION_ID)) & "') "
      End If
   End If
   
   If Len(TO_LOCATION_ID) > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (MR.KEY_CODE <= '" & Trim(ChangeQuote(TO_LOCATION_ID)) & "') "
      Else
         WhereStr = WhereStr & "AND (MR.KEY_CODE <= '" & Trim(ChangeQuote(TO_LOCATION_ID)) & "') "
      End If
   End If
   
   If Ind = 1 Then
      If ORDER_BY = 2 Then
         OrderByStr = "ORDER BY KEY_NAME "
      Else
         OrderByStr = "ORDER BY KEY_CODE "
      End If
   ElseIf Ind = 2 Then
      OrderByStr = "ORDER BY KEY_CODE "
   ElseIf Ind = 5 Then
      If ORDER_BY = 1 Then
         OrderByStr = "ORDER BY KEY_CODE "
      ElseIf ORDER_BY = 2 Then
         OrderByStr = "ORDER BY KEY_NAME "
      Else
         OrderByStr = "ORDER BY KEY_CODE "
      End If
   ElseIf Ind = 6 Then
      OrderByStr = "ORDER BY KEY_CODE "
   End If
   
   If ORDER_TYPE = 1 Then
      OrderByStr = OrderByStr & " ASC "
   ElseIf ORDER_TYPE = 2 Then
      OrderByStr = OrderByStr & " DESC "
   Else
      OrderByStr = OrderByStr & " ASC "
   End If
   
   SQL1 = SelectStr1 & FromStr & WhereStr & GroupByStr & OrderByStr
   SQL2 = SelectStr2 & FromStr & WhereStr & GroupByStr
   
   Call glbDatabaseMngr.GetRs(SQL1, SQL2, MustCount, ItemCount, Rs, glbErrorLog)
End Sub
Public Sub PopulateFromRS(Ind As Long, Rs As ADODB.Recordset)
   If Ind = 1 Then
      KEY_ID = NVLI(Rs("KEY_ID"), -1)
      KEY_CODE = NVLS(Rs("KEY_CODE"), "")
      KEY_NAME = NVLS(Rs("KEY_NAME"), "")
      MASTER_AREA = NVLI(Rs("MASTER_AREA"), -1)
      PARENT_ID = NVLI(Rs("PARENT_ID"), -1)
      PARENT_EX_ID = NVLI(Rs("PARENT_EX_ID"), -1)
      PARENT_EX_ID2 = NVLI(Rs("PARENT_EX_ID2"), -1)
      PARENT_EX_ID3 = NVLI(Rs("PARENT_EX_ID3"), -1)
      PARENT_EX_ID4 = NVLI(Rs("PARENT_EX_ID4"), -1)
      PARENT_EX_ID5 = NVLI(Rs("PARENT_EX_ID5"), -1)
      INDEX_LINK = NVLI(Rs("INDEX_LINK"), -1)
      
      MASTER_FLAG = NVLS(Rs("MASTER_FLAG"), "N")
      SHORT_CODE = NVLS(Rs("SHORT_CODE"), "")
      
      KEY_ID = NVLI(Rs("KEY_ID"), -1)
      KEY_LOOKUP = NVLS(Rs("KEY_CODE"), "")
      
      APAR_CODE = NVLS(Rs("APAR_CODE"), "")
      EMP_CODE = NVLS(Rs("EMP_CODE"), "")
   ElseIf Ind = 2 Then
      KEY_ID = NVLI(Rs("KEY_ID"), -1)
      PARENT_EX_ID2 = NVLI(Rs("PARENT_EX_ID2"), -1)
   ElseIf Ind = 5 Then
      KEY_ID = NVLI(Rs("KEY_ID"), -1)
      KEY_CODE = NVLS(Rs("KEY_CODE"), "")
      KEY_NAME = NVLS(Rs("KEY_NAME"), "")
   ElseIf Ind = 6 Then
      PARENT_EX_ID2 = NVLI(Rs("PARENT_EX_ID2"), -1)
      APAR_CODE = NVLS(Rs("APAR_CODE"), "")
      APAR_NAME = NVLS(Rs("APAR_NAME"), "")
      
      KEY_ID = NVLI(Rs("PARENT_EX_ID2"), -1)
      KEY_LOOKUP = NVLS(Rs("APAR_CODE"), "")
   End If
End Sub
Public Function AddEditData()
Dim ServerDateTime As String
Dim SQL1 As String
Dim ID As Long
Dim TempID1 As String
Dim TempID2 As String
Dim TempID3 As String
Dim TempID4  As String
Dim TempID5  As String
Dim TempID6  As String

   TempID1 = PARENT_ID
   If PARENT_ID <= 0 Then
      TempID1 = "NULL"
   End If
   
   TempID2 = PARENT_EX_ID
   If PARENT_EX_ID <= 0 Then
      TempID2 = "NULL"
   End If
   
   TempID3 = PARENT_EX_ID2
   If PARENT_EX_ID2 <= 0 Then
      TempID3 = "NULL"
   End If
   
   TempID4 = PARENT_EX_ID3
   If PARENT_EX_ID3 <= 0 Then
      TempID4 = "NULL"
   End If
   
   TempID5 = PARENT_EX_ID4
   If PARENT_EX_ID4 <= 0 Then
      TempID5 = "NULL"
   End If
   
   TempID6 = PARENT_EX_ID5
   If PARENT_EX_ID5 <= 0 Then
      TempID6 = "NULL"
   End If
   
   '�� Link ��� ����� FK ��觵͹�����Ѻ�������͡��ä�ѧ
      
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   
   If ShowMode = SHOW_ADD Then
      Call glbDatabaseMngr.GetSeqID("MASTER_REF_SEQ", ID, glbErrorLog)
      
      SQL1 = "INSERT INTO MASTER_REF " & _
                     "( " & _
                     "   KEY_ID, " & _
                     "   KEY_CODE, " & _
                     "   KEY_NAME, " & _
                     "   MASTER_AREA, " & _
                     "   PARENT_ID, " & _
                     "   PARENT_EX_ID, " & _
                     "   PARENT_EX_ID2, " & _
                     "   PARENT_EX_ID3, " & _
                     "   PARENT_EX_ID4, " & _
                     "   PARENT_EX_ID5, " & _
                     "   MASTER_FLAG, " & _
                     "   SHORT_CODE, " & _
                     "   INDEX_LINK, " & _
                     "   CREATE_DATE, " & _
                    "    CREATE_BY, " & _
                    "    MODIFY_DATE, " & _
                    "    MODIFY_BY " & _
                     ") "
         SQL1 = SQL1 & _
                     "VALUES " & _
                     "( " & _
                         ID & "," & _
                        "'" & ChangeQuote(Trim(KEY_CODE)) & "'," & _
                        "'" & ChangeQuote(Trim(KEY_NAME)) & "'," & _
                        MASTER_AREA & "," & _
                        TempID1 & "," & _
                        TempID2 & "," & _
                        TempID3 & "," & _
                        TempID4 & "," & _
                        TempID5 & "," & _
                        TempID6 & "," & _
                        "'" & ChangeQuote(Trim(MASTER_FLAG)) & "'," & _
                        "'" & ChangeQuote(Trim(SHORT_CODE)) & "'," & _
                        INDEX_LINK & "," & _
                         "'" & ServerDateTime & "'," & _
                         glbUser.USER_ID & "," & _
                         "'" & ServerDateTime & "'," & _
                         glbUser.USER_ID & _
                     ") "
         KEY_ID = ID
   Else 'Show edit
      SQL1 = "UPDATE MASTER_REF SET " & _
                     "    KEY_CODE = '" & ChangeQuote(Trim(KEY_CODE)) & "', " & _
                     "    KEY_NAME = '" & ChangeQuote(Trim(KEY_NAME)) & "', " & _
                     "    MASTER_AREA = " & MASTER_AREA & ", " & _
                     "    PARENT_ID = " & TempID1 & ", " & _
                     "    PARENT_EX_ID = " & TempID2 & ", " & _
                     "    PARENT_EX_ID2 = " & TempID3 & ", " & _
                     "    PARENT_EX_ID3 = " & TempID4 & ", " & _
                     "    PARENT_EX_ID4 = " & TempID5 & ", " & _
                     "    PARENT_EX_ID5 = " & TempID6 & ", " & _
                     "    MASTER_FLAG = '" & ChangeQuote(Trim(MASTER_FLAG)) & "', " & _
                     "    SHORT_CODE = '" & ChangeQuote(Trim(SHORT_CODE)) & "', " & _
                     "    INDEX_LINK = " & INDEX_LINK & ", " & _
                     "    MODIFY_DATE = '" & ServerDateTime & "', " & _
                     "    MODIFY_BY = " & glbUser.USER_ID & " " & _
                     "WHERE KEY_ID = " & KEY_ID
   End If
   
   If (ShowMode = SHOW_ADD) Or (ShowMode = SHOW_EDIT) Then
      m_ADOConn.Execute (SQL1)
   End If
End Function
