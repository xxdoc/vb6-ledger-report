VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CAPRcpCq"
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
Public Flag As String
' End of generic data member

Public RCPNUM  As String
Public CHQNUM As String
Public PAYAMT As Double

Public CHQDAT As Date
Public SUPCOD As String
Public SUPNAM As String
Public TYPDES As String
Public VOUCHER As String
Public SUPTYP As String
Public Credit As Long

Public KEY_ID As Long
Public KEY_LOOKUP As String

Public FROM_CHECK_DATE As Date
Public TO_CHECK_DATE As Date

Private Sub Class_Initialize()
   Set m_Conn = glbDatabaseMngr.DBConnection
   FROM_CHECK_DATE = -1
   TO_CHECK_DATE = -1
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
Dim SubLen As Long
Dim NewStr As String
Dim TempDate As String
Dim GroupByStr As String

   If Ind = 1 Then
      SelectStr1 = "SELECT ART.* "
      SelectStr2 = "SELECT COUNT(*)  "
      
      FromStr = "FROM APRCPCQ ART "
      
      WhereStr = ""
   ElseIf Ind = 2 Then
      SelectStr1 = "SELECT ART.CHQNUM, ART.PAYAMT, DTOS(BKT.CHQDAT) CHQDAT, APM.SUPNAM, APM.SUPCOD, BKT.VOUCHER, IST.TYPDES "
      SelectStr2 = "SELECT COUNT(*)  "
      
      FromStr = "FROM APRCPCQ ART "
      FromStr = FromStr & "LEFT OUTER JOIN BKTRN BKT ON (BKT.CHQNUM = ART.CHQNUM) "
      FromStr = FromStr & "LEFT OUTER JOIN APMAS APM ON (BKT.CUSCOD = APM.SUPCOD) "
      FromStr = FromStr & "LEFT OUTER JOIN ISTAB IST ON ((APM.SUPTYP = IST.TYPCOD) AND (IST.TABTYP = '46')) "
      
      WhereStr = "WHERE (BKT.BKTRNTYP = 'QP') "
   End If
   
   If Trim(RCPNUM) <> "" Then
      If WildCard(RCPNUM, SubLen, NewStr) Then
         If WhereStr = "" Then
            WhereStr = "WHERE (RCPNUM LIKE '" & ChangeQuote(Trim(RCPNUM)) & "')"
         Else
            WhereStr = WhereStr & "AND (RCPNUM LIKE '" & ChangeQuote(Trim(RCPNUM)) & "')"
         End If
      Else
         If WhereStr = "" Then
            WhereStr = "WHERE (RCPNUM = '" & ChangeQuote(Trim(RCPNUM)) & "')"
         Else
            WhereStr = WhereStr & "AND (RCPNUM = '" & ChangeQuote(Trim(RCPNUM)) & "')"
         End If
      End If
   End If
                  
    If FROM_CHECK_DATE > -1 Then
      TempDate = DateToStringIntLow(Trim(FROM_CHECK_DATE))
      If WhereStr = "" Then
         WhereStr = "WHERE (DTOS(BKT.CHQDAT) >= '" & ChangeQuote(TempDate) & "')"
      Else
         WhereStr = WhereStr & "AND (DTOS(BKT.CHQDAT) >= '" & ChangeQuote(TempDate) & "')"
      End If
   End If

   If TO_CHECK_DATE > -1 Then
      TempDate = DateToStringIntHi(Trim(TO_CHECK_DATE))
      If WhereStr = "" Then
         WhereStr = "WHERE (DTOS(BKT.CHQDAT) <= '" & ChangeQuote(TempDate) & "')"
      Else
         WhereStr = WhereStr & "AND (DTOS(BKT.CHQDAT) <= '" & ChangeQuote(TempDate) & "')"
      End If
   End If
                  
   If Trim(SUPCOD) <> "" Then
      If WildCard(SUPCOD, SubLen, NewStr) Then
         If WhereStr = "" Then
            WhereStr = "WHERE (APM.SUPCOD LIKE '" & ChangeQuote(Trim(SUPCOD)) & "')"
         Else
            WhereStr = WhereStr & "AND (APM.SUPCOD LIKE '" & ChangeQuote(Trim(SUPCOD)) & "')"
         End If
      Else
         If WhereStr = "" Then
            WhereStr = "WHERE (APM.SUPCOD = '" & ChangeQuote(Trim(SUPCOD)) & "')"
         Else
            WhereStr = WhereStr & "AND (APM.SUPCOD = '" & ChangeQuote(Trim(SUPCOD)) & "')"
         End If
      End If
   End If
                  
   If Len(SUPTYP) > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (APM.SUPTYP = '" & Trim(ChangeQuote(SUPTYP)) & "') "
      Else
         WhereStr = WhereStr & "AND (APM.SUPTYP = '" & Trim(ChangeQuote(SUPTYP)) & "') "
      End If
   End If
   
   GroupByStr = ""
   
   OrderByStr = ""
   OrderTypeStr = ""
   
   If Ind = 1 Then
      If OrderBy = 1 Then
         OrderByStr = "ORDER BY RCPNUM "
      ElseIf OrderBy = 2 Then
         OrderByStr = "ORDER BY RCPNUM "
      Else
         OrderByStr = "ORDER BY RCPNUM "
      End If
   ElseIf Ind = 2 Then
      OrderByStr = "ORDER BY APM.SUPNAM ASC, BKT.CHQDAT  "
   End If
   
   If OrderType = 1 Then
      OrderTypeStr = "ASC "
   ElseIf OrderType = 2 Then
      OrderTypeStr = "DESC "
   Else
      OrderTypeStr = "ASC "
   End If
   
   SQL1 = SelectStr1 & FromStr & WhereStr & GroupByStr & OrderByStr & OrderTypeStr
   SQL2 = SelectStr2 & FromStr & WhereStr
   
   Call glbDatabaseMngr.GetRs(SQL1, SQL2, True, ItemCount, Rs, glbErrorLog)
End Function

Public Function PopulateFromRS(Ind As Long, Rs As ADODB.Recordset)
   If Ind = 1 Then
      RCPNUM = NVLS(Rs("RCPNUM"), "")
      CHQNUM = NVLS(Rs("CHQNUM"), "")
      PAYAMT = NVLD(Rs("PAYAMT"), 0)
   ElseIf Ind = 2 Then
'      RCPNUM = NVLS(Rs("RCPNUM"), "")
      CHQNUM = NVLS(Rs("CHQNUM"), "")
      PAYAMT = NVLD(Rs("PAYAMT"), 0)
      
      CHQDAT = InternalDateToDate(NVLS(Rs("CHQDAT"), ""))
      SUPCOD = NVLS(Rs("SUPCOD"), "")
      SUPNAM = NVLS(Rs("SUPNAM"), "")
      TYPDES = NVLS(Rs("TYPDES"), "")
      VOUCHER = NVLS(Rs("VOUCHER"), "")
   End If
End Function


