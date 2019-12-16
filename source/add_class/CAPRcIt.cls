VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CAPRcIt"
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

Public RCPNUM   As String
Public DOCNUM  As String
Public RECTYP As String
Public PAYAMT  As Double
Public VATAMT As Double

Public FROM_PAY_DATE As Date
Public TO_PAY_DATE As Date
Public FROM_DOC_DATE As Date
Public TO_DOC_DATE As Date

Public SUPCOD As String

Public KEY_ID As Long
Public KEY_LOOKUP As String

Public RecTypeSet As String

Private Sub Class_Initialize()
   Set m_Conn = glbDatabaseMngr.DBConnection
   FROM_PAY_DATE = -1
   TO_PAY_DATE = -1
   FROM_DOC_DATE = -1
   TO_DOC_DATE = -1
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
      
      FromStr = "FROM APRCPIT ART "
   ElseIf Ind = 2 Then
      SelectStr1 = "SELECT ART.DOCNUM, SUM(ART.PAYAMT) PAYAMT, SUM(ART.VATAMT) VATAMT "
      SelectStr2 = "SELECT COUNT(*)  "
      
      FromStr = "FROM APRCPIT ART "
      FromStr = FromStr & "LEFT OUTER JOIN APTRN APT ON (ART.RCPNUM = APT.DOCNUM) "
                     'WhereStr = "WHERE (APT.SUPCOD = 'Á-0002') AND (ART.RECTYP = '3') "
   ElseIf Ind = 3 Then
      SelectStr1 = "SELECT APT.SUPCOD, SUM(ART.PAYAMT) PAYAMT, SUM(ART.VATAMT) VATAMT "
      SelectStr2 = "SELECT COUNT(*)  "
      
      FromStr = "FROM APRCPIT ART "
      FromStr = FromStr & "LEFT OUTER JOIN APTRN APT ON (ART.RCPNUM = APT.DOCNUM) "
      If FROM_DOC_DATE > 0 Or TO_DOC_DATE > 0 Then
         FromStr = FromStr & "LEFT OUTER JOIN APTRN APT2 ON (ART.DOCNUM = APT2.DOCNUM) "
      End If
                       'WhereStr = "WHERE (APT.SUPCOD = 'Á-0002') AND (ART.RECTYP = '3') "
                       'WhereStr = "WHERE (APT2.SUPCOD = '´-0009')  "
   End If
   
   If Trim(DOCNUM) <> "" Then
      If WildCard(DOCNUM, SubLen, NewStr) Then
         If WhereStr = "" Then
            WhereStr = "WHERE (ART.DOCNUM LIKE '" & ChangeQuote(Trim(DOCNUM)) & "')"
         Else
            WhereStr = WhereStr & "AND (ART.DOCNUM LIKE '" & ChangeQuote(Trim(DOCNUM)) & "')"
         End If
      Else
         If WhereStr = "" Then
            WhereStr = "WHERE (ART.DOCNUM = '" & ChangeQuote(Trim(DOCNUM)) & "')"
         Else
            WhereStr = WhereStr & "AND (ART.DOCNUM = '" & ChangeQuote(Trim(DOCNUM)) & "')"
         End If
      End If
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
            
   If Trim(RECTYP) <> "" Then
      If WildCard(RECTYP, SubLen, NewStr) Then
         If WhereStr = "" Then
            WhereStr = "WHERE (RECTYP LIKE '" & ChangeQuote(Trim(RECTYP)) & "')"
         Else
            WhereStr = WhereStr & "AND (RECTYP LIKE '" & ChangeQuote(Trim(RECTYP)) & "')"
         End If
      Else
         If WhereStr = "" Then
            WhereStr = "WHERE (RECTYP = '" & ChangeQuote(Trim(RECTYP)) & "')"
         Else
            WhereStr = WhereStr & "AND (RECTYP = '" & ChangeQuote(Trim(RECTYP)) & "')"
         End If
      End If
   End If

    If FROM_PAY_DATE > -1 Then
      TempDate = DateToStringIntLow(Trim(FROM_PAY_DATE))
      If WhereStr = "" Then
         WhereStr = "WHERE (DTOS(APT.DOCDAT) >= '" & ChangeQuote(TempDate) & "')"
      Else
         WhereStr = WhereStr & "AND (DTOS(APT.DOCDAT) >= '" & ChangeQuote(TempDate) & "')"
      End If
   End If

   If TO_PAY_DATE > -1 Then
      TempDate = DateToStringIntHi(Trim(TO_PAY_DATE))
      If WhereStr = "" Then
         WhereStr = "WHERE ((DTOS(APT.DOCDAT) <= '" & ChangeQuote(TempDate) & "') OR (APT.DOCDAT IS NULL)) "
      Else
         WhereStr = WhereStr & "AND ((DTOS(APT.DOCDAT) <= '" & ChangeQuote(TempDate) & "') OR (APT.DOCDAT IS NULL)) "
      End If
   End If
   
   If FROM_DOC_DATE > -1 Then
      TempDate = DateToStringIntLow(Trim(FROM_DOC_DATE))
      If WhereStr = "" Then
         WhereStr = "WHERE (DTOS(APT2.DOCDAT) >= '" & ChangeQuote(TempDate) & "')"
      Else
         WhereStr = WhereStr & "AND (DTOS(APT2.DOCDAT) >= '" & ChangeQuote(TempDate) & "')"
      End If
   End If

   If TO_DOC_DATE > -1 Then
      TempDate = DateToStringIntHi(Trim(TO_DOC_DATE))
      If WhereStr = "" Then
         WhereStr = "WHERE ((DTOS(APT2.DOCDAT) <= '" & ChangeQuote(TempDate) & "') OR (APT2.DOCDAT IS NULL)) "
      Else
         WhereStr = WhereStr & "AND ((DTOS(APT2.DOCDAT) <= '" & ChangeQuote(TempDate) & "') OR (APT2.DOCDAT IS NULL)) "
      End If
   End If
   
   If Trim(RecTypeSet) <> "" Then
      If WhereStr = "" Then
         WhereStr = "WHERE (ART.RECTYP IN " & RecTypeSet & ") "
      Else
         WhereStr = WhereStr & "AND (ART.RECTYP IN " & RecTypeSet & ")"
      End If
   End If
   
   GroupByStr = ""
   If Ind = 2 Then
      GroupByStr = "GROUP BY  ART.DOCNUM "
   ElseIf Ind = 3 Then
      GroupByStr = "GROUP BY  APT.SUPCOD "
   End If
   
   OrderByStr = ""
   OrderTypeStr = ""
   
   If Ind = 1 Then
      If OrderBy = 1 Then
         OrderByStr = "ORDER BY RCPNUM ASC, DOCNUM "
      ElseIf OrderBy = 2 Then
         OrderByStr = "ORDER BY RCPNUM ASC, DOCNUM "
      Else
         OrderByStr = "ORDER BY RCPNUM ASC, DOCNUM "
      End If
   ElseIf Ind = 2 Then
      OrderByStr = "ORDER BY ART.DOCNUM "
   ElseIf Ind = 3 Then
      'OrderByStr = "ORDER BY APT2.SUPCOD "
      OrderByStr = "ORDER BY APT.SUPCOD "
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
      RECTYP = NVLS(Rs("RECTYP"), "")
      RCPNUM = NVLS(Rs("RCPNUM"), "")
      DOCNUM = NVLS(Rs("DOCNUM"), "")
      PAYAMT = NVLD(Rs("PAYAMT"), 0)
      VATAMT = NVLD(Rs("VATAMT"), 0)
   ElseIf Ind = 2 Then
      DOCNUM = NVLS(Rs("DOCNUM"), "")
      PAYAMT = NVLD(Rs("PAYAMT"), 0)
      VATAMT = NVLD(Rs("VATAMT"), 0)
   ElseIf Ind = 3 Then
      SUPCOD = NVLS(Rs("SUPCOD"), "")
      PAYAMT = NVLD(Rs("PAYAMT"), 0)
      VATAMT = NVLD(Rs("VATAMT"), 0)
   End If
End Function

