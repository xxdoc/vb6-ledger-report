VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CFaMas"
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

Public FASCOD As String
Public FASDES As String
Public FASTYP As String
Public FASGRP As String
Public DEPCOD As String
Public ACCNUM As String
Public ACCNUMDP As String
Public ACCNUMAC As String
Public FIRSTDAT As Date
Public LASTDAT As Date
Public METHOD As Long
Public RATE As Double
Public LIFE As Double
Public PURVAL As Double
Public SALVAGVAL As Double
Public ACDPRVAL As Double
Public BALVAL As Double
Public DPRVAL As Double

Public FROM_FASCOD_DATE As Date
Public TO_FASCOD_DATE As Date
Public ASSET_CODE_SET As String
Public FROM_BUY_DATE As Date
Public TO_BUY_DATE As Date

Public KEY_ID As Long
Public KEY_LOOKUP As String

Private Sub Class_Initialize()
   Set m_Conn = glbDatabaseMngr.DBConnection
   FROM_FASCOD_DATE = -1
   TO_FASCOD_DATE = -1
   FROM_BUY_DATE = -1
   TO_BUY_DATE = -1
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
      SelectStr1 = "SELECT ART.*, DTOS(ART.FIRSTDAT) FIRSTDAT_STR, DTOS(ART.LASTDAT) LASTDAT_STR, GLA.ACCNAM "
      SelectStr2 = "SELECT COUNT(*)  "
      
      FromStr = "FROM FAMAS ART "
      FromStr = FromStr & "LEFT OUTER JOIN GLACC GLA ON (ART.ACCNUM = GLA.ACCNUM) "
   End If
   
   If Trim(FASCOD) <> "" Then
      If WildCard(FASCOD, SubLen, NewStr) Then
         If WhereStr = "" Then
            WhereStr = "WHERE (FASCOD LIKE '" & ChangeQuote(Trim(FASCOD)) & "')"
         Else
            WhereStr = WhereStr & "AND (FASCOD LIKE '" & ChangeQuote(Trim(FASCOD)) & "')"
         End If
      Else
         If WhereStr = "" Then
            WhereStr = "WHERE (FASCOD = '" & ChangeQuote(Trim(FASCOD)) & "')"
         Else
            WhereStr = WhereStr & "AND (FASCOD = '" & ChangeQuote(Trim(FASCOD)) & "')"
         End If
      End If
   End If
                                       
   If Trim(ACCNUMAC) <> "" Then
      If WildCard(ACCNUMAC, SubLen, NewStr) Then
         If WhereStr = "" Then
            WhereStr = "WHERE (ART.ACCNUMAC LIKE '" & ChangeQuote(Trim(ACCNUMAC)) & "')"
         Else
            WhereStr = WhereStr & "AND (ART.ACCNUMAC LIKE '" & ChangeQuote(Trim(ACCNUMAC)) & "')"
         End If
      Else
         If WhereStr = "" Then
            WhereStr = "WHERE (ART.ACCNUMAC = '" & ChangeQuote(Trim(ACCNUMAC)) & "')"
         Else
            WhereStr = WhereStr & "AND (ART.ACCNUMAC = '" & ChangeQuote(Trim(ACCNUMAC)) & "')"
         End If
      End If
   End If
                                       
                                       
   If Trim(ACCNUMDP) <> "" Then
      If WildCard(ACCNUMDP, SubLen, NewStr) Then
         If WhereStr = "" Then
            WhereStr = "WHERE (ART.ACCNUMDP LIKE '" & ChangeQuote(Trim(ACCNUMDP)) & "')"
         Else
            WhereStr = WhereStr & "AND (ART.ACCNUMDP LIKE '" & ChangeQuote(Trim(ACCNUMDP)) & "')"
         End If
      Else
         If WhereStr = "" Then
            WhereStr = "WHERE (ART.ACCNUMDP = '" & ChangeQuote(Trim(ACCNUMDP)) & "')"
         Else
            WhereStr = WhereStr & "AND (ART.ACCNUMDP = '" & ChangeQuote(Trim(ACCNUMDP)) & "')"
         End If
      End If
   End If
   
   If Trim(ACCNUM) <> "" Then
      If WildCard(ACCNUM, SubLen, NewStr) Then
         If WhereStr = "" Then
            WhereStr = "WHERE (ART.ACCNUM LIKE '" & ChangeQuote(Trim(ACCNUM)) & "')"
         Else
            WhereStr = WhereStr & "AND (ART.ACCNUM LIKE '" & ChangeQuote(Trim(ACCNUM)) & "')"
         End If
      Else
         If WhereStr = "" Then
            WhereStr = "WHERE (ART.ACCNUM = '" & ChangeQuote(Trim(ACCNUM)) & "')"
         Else
            WhereStr = WhereStr & "AND (ART.ACCNUM = '" & ChangeQuote(Trim(ACCNUM)) & "')"
         End If
      End If
   End If
   
    If FROM_FASCOD_DATE > -1 Then
      TempDate = DateToStringIntLow(Trim(FROM_FASCOD_DATE))
      If WhereStr = "" Then
         WhereStr = "WHERE (DTOS(FIRSTDAT) >= '" & ChangeQuote(TempDate) & "')"
      Else
         WhereStr = WhereStr & "AND (DTOS(FIRSTDAT) >= '" & ChangeQuote(TempDate) & "')"
      End If
   End If

   If TO_FASCOD_DATE > -1 Then
      TempDate = DateToStringIntHi(Trim(TO_FASCOD_DATE))
      If WhereStr = "" Then
         WhereStr = "WHERE (DTOS(FIRSTDAT) <= '" & ChangeQuote(TempDate) & "')"
      Else
         WhereStr = WhereStr & "AND (DTOS(FIRSTDAT) <= '" & ChangeQuote(TempDate) & "')"
      End If
   End If
                                       
   GroupByStr = ""
   
   OrderByStr = ""
   OrderTypeStr = ""
   
   If Ind = 1 Then
      If OrderBy = 1 Then
         OrderByStr = "ORDER BY FASCOD "
      ElseIf OrderBy = 2 Then
         OrderByStr = "ORDER BY FASCOD "
      Else
         OrderByStr = "ORDER BY FASCOD "
      End If
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
      FASCOD = NVLS(Rs("FASCOD"), "")
      FASDES = NVLS(Rs("FASDES"), "")
      FASTYP = NVLS(Rs("FASTYP"), "")
      FASGRP = NVLS(Rs("FASGRP"), "")
      DEPCOD = NVLS(Rs("DEPCOD"), "")
      ACCNUM = NVLS(Rs("ACCNUM"), "")
      ACCNUMDP = NVLS(Rs("ACCNUMDP"), "")
      ACCNUMAC = NVLS(Rs("ACCNUMAC"), "")
      FIRSTDAT = InternalDateToDate(NVLS(Rs("FIRSTDAT_STR"), ""))
      LASTDAT = InternalDateToDate(NVLS(Rs("LASTDAT_STR"), ""))
      METHOD = NVLI(Rs("ACCNUMAC"), -1)
      RATE = NVLD(Rs("RATE"), 0)
      LIFE = NVLD(Rs("LIFE"), 0)
      PURVAL = NVLD(Rs("PURVAL"), 0)
      SALVAGVAL = NVLD(Rs("SALVAGVAL"), 0)
      ACDPRVAL = NVLD(Rs("ACDPRVAL"), 0)
      BALVAL = NVLD(Rs("BALVAL"), 0)
      DPRVAL = NVLD(Rs("DPRVAL"), 0)
   End If
End Function

Public Function PatchDprVal() As Boolean
Dim SQL As String

   SQL = "UPDATE FAMAS SET DPRVAL = " & DPRVAL & " WHERE FASCOD IN " & ASSET_CODE_SET
   m_Conn.Execute (SQL)

   SQL = "UPDATE FADPRIT SET DPRVAL = " & DPRVAL & " WHERE FASCOD IN " & ASSET_CODE_SET
   m_Conn.Execute (SQL)
End Function



