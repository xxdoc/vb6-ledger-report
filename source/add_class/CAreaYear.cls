VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CAreaYear"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_Conn As ADODB.Connection

Public AddEditMode As SHOW_MODE_TYPE
Public QueryFlag As Long
Public OrderBy As Long
Public OrderType As Long

' End of generic data member

Public YEAR_ID As Long
Public YEARNUM As String
Public FROM_DATE As Date
Public TO_DATE As Date
Public Flag As String

Public ORDER_BY As Long
Public ORDER_TYPE  As Long

'Public KEY_ID As Long
'Public KEY_LOOKUP As String
Private Sub Class_Initialize()
   Set m_Conn = glbDatabaseMngr.DBLegacyConnection
   FROM_DATE = -1
   TO_DATE = -1
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
   
  If Ind = 1 Then
      SelectStr1 = "SELECT CSY.*  "
      SelectStr2 = "SELECT COUNT(*)  "
      
      FromStr = "FROM AREA_YEAR CSY "
  ElseIf Ind = 2 Then
      SelectStr1 = "SELECT MAX(CSY.YEAR_ID) AS YEAR_ID  "
      SelectStr2 = "SELECT COUNT(*)  "
      
      FromStr = "FROM AREA_YEAR CSY "
  End If

      If YEAR_ID > 0 Then
        WhereStr = "WHERE (YEAR_ID=" & ChangeQuote(Trim(YEAR_ID)) & ") "
      End If
   
         If Trim(YEARNUM) <> "" Then
            If WhereStr = "" Then
               WhereStr = "WHERE (CSY.YEARNUM = '" & YEARNUM & "') "
            Else
               WhereStr = WhereStr & "AND (CSY.YEARNUM = '" & YEARNUM & "') "
            End If
         End If
         
           'วันที่รับชำระ
         If FROM_DATE > -1 Then                     ' จะเอาไว้ โหลดรีพอร์ต
            TempDate = DateToStringIntLow(Trim(FROM_DATE))
            If WhereStr = "" Then
               WhereStr = "WHERE (CSY.FROM_DATE <= '" & ChangeQuote(TempDate) & "') "  ' เครื่องหมายจะตรงกันข้ามจากปรกติ เพราะต้องการให้ครอบครุม
            Else
               WhereStr = WhereStr & "AND (CSY.FROM_DATE <= '" & ChangeQuote(TempDate) & "') "
            End If
         End If
      
         If TO_DATE > -1 Then
            TempDate = DateToStringIntHi(Trim(TO_DATE))
            If WhereStr = "" Then
               WhereStr = "WHERE (CSY.TO_DATE >= '" & ChangeQuote(TempDate) & "') "
            Else
               WhereStr = WhereStr & "AND (CSY.TO_DATE >= '" & ChangeQuote(TempDate) & "') "
            End If
         End If
   
    If Ind = 1 Then
         If OrderBy = 2 Then
             OrderByStr = "ORDER BY FROM_DATE  "
          ElseIf OrderBy = 3 Then
             OrderByStr = "ORDER BY TO_DATE "
          Else
             OrderByStr = "ORDER BY YEARNUM "
          End If
   End If
               
    If Ind = 1 Then
         If OrderType = 1 Then
             OrderTypeStr = "ASC "
         Else
           OrderTypeStr = "DESC "
         End If
    End If
   
   SQL1 = SelectStr1 & FromStr & WhereStr & OrderByStr & OrderTypeStr
   SQL2 = SelectStr2 & FromStr & WhereStr
 'debug.print SQL1
   Call glbDatabaseMngr.GetLegacyRs(SQL1, SQL2, True, ItemCount, Rs, glbErrorLog)
End Function

Public Function PopulateFromRS(Ind As Long, Rs As ADODB.Recordset)
  If Ind = 1 Then
      YEAR_ID = NVLI(Rs("YEAR_ID"), "-1")
      YEARNUM = NVLS(Rs("YEARNUM"), "")
      FROM_DATE = InternalDateToDate(NVLS(Rs("FROM_DATE"), ""))
      TO_DATE = InternalDateToDate(NVLS(Rs("TO_DATE"), ""))
      Flag = NVLS(Rs("FLAG"), -1)
   ElseIf Ind = 2 Then
      YEAR_ID = NVLI(Rs("YEAR_ID"), "-1")
   End If

End Function
Public Function AddEditData()
Dim ServerDateTime As String
Dim SQL1 As String

   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   
'   ENTERPRISE_CODE = glbCompanyCode
   
   If AddEditMode = SHOW_ADD Then
             SQL1 = "INSERT INTO AREA_YEAR " & _
                              "( " & _
                              "   YEARNUM, " & _
                              "   FROM_DATE, " & _
                              "   TO_DATE, " & _
                              "   FLAG " & _
                              ") " & _
                              "VALUES " & _
                              "( " & _
                                 "'" & ChangeQuote(Trim(YEARNUM)) & "'," & _
                                  "'" & ChangeQuote(Trim(DateToStringInt(FROM_DATE))) & "'," & _
                                   "'" & ChangeQuote(Trim(DateToStringInt(TO_DATE))) & "'," & _
                                 "'Y'" & _
                              ") "
   Else 'Show edit
         
               SQL1 = "UPDATE AREA_YEAR SET " & _
                           "   YEARNUM = '" & ChangeQuote(Trim(YEARNUM)) & "'," & _
                           "   FROM_DATE = '" & ChangeQuote(Trim(DateToStringInt(FROM_DATE))) & "'," & _
                           "   TO_DATE = '" & ChangeQuote(Trim(DateToStringInt(TO_DATE))) & "'," & _
                           "   FLAG = '" & ChangeQuote(Trim(Flag)) & "'" & _
                           "   WHERE YEAR_ID = " & YEAR_ID

   End If
   
   If (AddEditMode = SHOW_ADD) Or (AddEditMode = SHOW_EDIT) Then
      m_Conn.Execute (SQL1)
   End If
End Function

Public Function DeleteData()
Dim SQL1 As String
      SQL1 = "DELETE FROM AREA_YEAR WHERE YEAR_ID = " & YEAR_ID
      m_Conn.Execute (SQL1)
      SQL1 = "DELETE FROM COMMISSION_CUS_AREA WHERE YEAR_ID = " & YEAR_ID
      m_Conn.Execute (SQL1)
End Function




