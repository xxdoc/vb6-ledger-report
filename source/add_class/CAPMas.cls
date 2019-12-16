VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CAPMas"
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

Public SUPCOD As String
Public SUPTYP As String
Public PRENAM  As String
Public SUPNAM  As String
Public ADDR01  As String
Public ADDR02 As String
Public TELNUM As String
Public CONTRACT As String
Public PAYTRM  As Double
Public PAYCOND As String

Public TYPDES As String
Public FROM_SUPPLIER_CODE As String
Public TO_SUPPLIER_CODE As String

Public AMPHUR As String
Public PROVINCE As String

Public KEY_ID As Long
Public KEY_LOOKUP As String

Private Sub Class_Initialize()
   Set m_Conn = glbDatabaseMngr.DBConnection
End Sub

Public Function PackAddress()
Dim TempStr As String

   If Len(ADDR02) > 0 Then
      PackAddress = ADDR01 & vbCrLf & PackAddress
   Else
      PackAddress = ADDR01
   End If
End Function

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
   
   'SelectStr1 = "SELECT APM.*, IST.*  "
   SelectStr1 = "SELECT APM.SUPCOD,APM.SUPTYP,APM.PRENAM,APM.SUPNAM,APM.ADDR01,APM.ADDR02,APM.TELNUM,APM.CONTRACT,APM.PAYTRM,APM.PAYCOND, IST.*  "
   
   SelectStr2 = "SELECT COUNT(*)  "
   
   FromStr = "FROM APMAS APM "
   FromStr = FromStr & "LEFT OUTER JOIN ISTAB IST ON (APM.SUPTYP = IST.TYPCOD) AND (IST.TABTYP = '46') "
   
   If Trim(SUPCOD) <> "" Then
      If WildCard(SUPCOD, SubLen, NewStr) Then
         If WhereStr = "" Then
            WhereStr = "WHERE (SUPCOD LIKE '" & ChangeQuote(Trim(SUPCOD)) & "')"
         Else
            WhereStr = WhereStr & "AND (SUPCOD LIKE '" & ChangeQuote(Trim(SUPCOD)) & "')"
         End If
      Else
         If WhereStr = "" Then
            WhereStr = "WHERE (SUPCOD = '" & ChangeQuote(Trim(SUPCOD)) & "')"
         Else
            WhereStr = WhereStr & "AND (SUPCOD = '" & ChangeQuote(Trim(SUPCOD)) & "')"
         End If
      End If
   End If
   
   If Trim(SUPNAM) <> "" Then
      If WildCard(SUPNAM, SubLen, NewStr) Then
         If WhereStr = "" Then
            WhereStr = "WHERE (SUPNAM LIKE '" & ChangeQuote(Trim(SUPNAM)) & "')"
         Else
            WhereStr = WhereStr & "AND (SUPNAM LIKE '" & ChangeQuote(Trim(SUPNAM)) & "')"
         End If
      Else
         If WhereStr = "" Then
            WhereStr = "WHERE (SUPNAM = '" & ChangeQuote(Trim(SUPNAM)) & "')"
         Else
            WhereStr = WhereStr & "AND (SUPNAM = '" & ChangeQuote(Trim(SUPNAM)) & "')"
         End If
      End If
   End If
   
   If Len(SUPTYP) > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (SUPTYP = '" & Trim(ChangeQuote(SUPTYP)) & "') "
      Else
         WhereStr = WhereStr & "AND (SUPTYP = '" & Trim(ChangeQuote(SUPTYP)) & "') "
      End If
   End If
   
   If Len(FROM_SUPPLIER_CODE) > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (SUPCOD >= '" & Trim(ChangeQuote(FROM_SUPPLIER_CODE)) & "') "
      Else
         WhereStr = WhereStr & "AND (SUPCOD >= '" & Trim(ChangeQuote(FROM_SUPPLIER_CODE)) & "') "
      End If
   End If
   
   If Len(TO_SUPPLIER_CODE) > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (SUPCOD <= '" & Trim(ChangeQuote(TO_SUPPLIER_CODE)) & "') "
      Else
         WhereStr = WhereStr & "AND (SUPCOD <= '" & Trim(ChangeQuote(TO_SUPPLIER_CODE)) & "') "
      End If
   End If
   
   If Len(AMPHUR) > 0 Then
        If WhereStr = "" Then
            WhereStr = "WHERE (ADDR01 LIKE '%" & ChangeQuote(Trim(AMPHUR)) & "%')"
         Else
            WhereStr = WhereStr & "AND (ADDR01 LIKE '%" & ChangeQuote(Trim(AMPHUR)) & "%')"
         End If
   End If
   
    If Len(PROVINCE) > 0 Then
        If WhereStr = "" Then
            WhereStr = "WHERE (ADDR01 LIKE '%" & ChangeQuote(Trim(PROVINCE)) & "%')"
         Else
            WhereStr = WhereStr & "AND (ADDR01 LIKE '%" & ChangeQuote(Trim(PROVINCE)) & "%')"
         End If
   End If
   
   OrderByStr = ""
   OrderTypeStr = ""
   
   If OrderBy = 1 Then
      OrderByStr = "ORDER BY SUPCOD "
   ElseIf OrderBy = 2 Then
      OrderByStr = "ORDER BY SUPNAM "
   Else
      OrderByStr = "ORDER BY SUPCOD "
   End If
   
   If OrderType = 1 Then
      OrderTypeStr = "ASC "
   ElseIf OrderType = 2 Then
      OrderTypeStr = "DESC "
   Else
      OrderTypeStr = "ASC "
   End If
   
   SQL1 = SelectStr1 & FromStr & WhereStr & OrderByStr & OrderTypeStr
   SQL2 = SelectStr2 & FromStr & WhereStr
   
   Call glbDatabaseMngr.GetRs(SQL1, SQL2, True, ItemCount, Rs, glbErrorLog)
End Function

Public Function PopulateFromRS(Ind As Long, Rs As ADODB.Recordset)
   SUPCOD = NVLS(Rs("SUPCOD"), "")
   
'   If Left(SUPCOD, 1) = "¿" Then
'      'debug.print
'   End If
   
   SUPTYP = NVLS(Rs("SUPTYP"), "")
   PRENAM = NVLS(Rs("PRENAM"), "")
   SUPNAM = NVLS(Rs("SUPNAM"), "")
   ADDR01 = NVLS(Rs("ADDR01"), "")
   ADDR02 = NVLS(Rs("ADDR02"), "")
   TELNUM = NVLS(Rs("TELNUM"), "")
   CONTRACT = NVLS(Rs("CONTRACT"), "")
   PAYTRM = NVLD(Rs("PAYTRM"), 0)
   PAYCOND = NVLS(Rs("PAYCOND"), "")
   
   TYPDES = NVLS(Rs("TYPDES"), "")
   
   KEY_LOOKUP = SUPCOD
   KEY_ID = Val(SUPCOD)
End Function

Public Sub UpdateSupplierType()
Dim SQL1 As String

   SQL1 = "UPDATE APMAS SET SUPTYP = '" & SUPTYP & "'" & " WHERE SUPCOD = '" & SUPCOD & "'"
   m_Conn.Execute (SQL1)
End Sub

