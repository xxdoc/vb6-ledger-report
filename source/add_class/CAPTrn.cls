VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CApTrn"
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

Public RECTYP As String
Public DOCNUM As String
Public DOCDAT  As Date
Public SUPCOD  As String
Public DUEDAT  As Date
Public AMOUNT  As Double
Public TOTAL  As Double
Public DISAMT As Double
Public PAYAMT As Double
Public CSHPAY  As Double
Public INTPAY As Double
Public PONUM As String
Public CMPLAPP As String
Public RCPNUM As String
Public REMAMT As Double

Public DESCRPT As String
Public FROM_DOC_DATE As Date
Public TO_DOC_DATE As Date
Public FROM_DUE_DATE As Date
Public TO_DUE_DATE As Date
Public FROM_CHEQUE_DATE As Date
Public TO_CHEQUE_DATE As Date
Public FROM_PAY_DATE As Date
Public TO_PAY_DATE As Date
Public SUPNAM As String
Public PRENAM As String
Public RecTypeSet As String

Public CHQNUM As String
Public CHQDAT As String

Public KEY_ID As Long
Public KEY_LOOKUP As String

Public ReceiptChecks As Collection
Public ReceiptItems As Collection
Public JournalItems As Collection
Public BankTrans As Collection

Private Sub Class_Initialize()
   Set m_Conn = glbDatabaseMngr.DBConnection
   Set ReceiptChecks = New Collection
   Set JournalItems = New Collection
   Set ReceiptItems = New Collection
   Set BankTrans = New Collection
   
   FROM_DOC_DATE = -1
   TO_DOC_DATE = -1
   FROM_DUE_DATE = -1
   TO_DUE_DATE = -1
   FROM_CHEQUE_DATE = -1
   TO_CHEQUE_DATE = -1
   FROM_PAY_DATE = -1
   TO_PAY_DATE = -1
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
      SelectStr1 = "SELECT APT.*, DTOS(APT.DUEDAT) DUE_DATE, DTOS(APT.DOCDAT) DOC_DATE, APM.SUPNAM, APM.PRENAM "
      SelectStr2 = "SELECT COUNT(*)  "
      
      FromStr = "FROM APTRN APT "
      FromStr = FromStr & "LEFT OUTER JOIN APMAS APM ON (APT.SUPCOD = APM.SUPCOD) "
   ElseIf Ind = 2 Then
      SelectStr1 = "SELECT APT.SUPCOD, SUM(APT.PAYAMT) PAYAMT, SUM(APT.TOTAL) TOTAL, SUM(APT.REMAMT) REMAMT  "
      SelectStr1 = SelectStr1 & ", SUM(APT.CSHPAY) CSHPAY,SUM(APT.INTPAY) INTPAY "
      
      SelectStr2 = "SELECT COUNT(*)  "
      
      FromStr = "FROM APTRN APT "
      
   ElseIf Ind = 3 Then
      SelectStr1 = "SELECT APT.PONUM, SUM(APT.TOTAL) TOTAL "
      SelectStr2 = "SELECT COUNT(*)  "
      
      FromStr = "FROM APTRN APT "
    ElseIf Ind = 4 Then
     SelectStr1 = "SELECT APR.RCPNUM,BKT.CHQNUM,APT.DOCNUM,DTOS(BKT.CHQDAT) CHQDAT, ARQ.PAYAMT AMOUNT,APT.SUPCOD,APM.SUPNAM,DTOS(APT.DOCDAT) DOC_DATE "
     SelectStr2 = "SELECT COUNT(*)  "
      
      FromStr = " FROM APTRN APT "
      FromStr = FromStr & "LEFT OUTER JOIN APRCPIT APR ON (APT.DOCNUM = APR.DOCNUM) "
      FromStr = FromStr & "LEFT OUTER JOIN APRCPCQ ARQ ON (ARQ.RCPNUM = APR.RCPNUM ) "
       FromStr = FromStr & "LEFT OUTER JOIN BKTRN BKT ON (ARQ.CHQNUM = BKT.CHQNUM) "
       
       FromStr = FromStr & "LEFT OUTER JOIN APTRN APT2 ON (APR.RCPNUM = APT2.DOCNUM) "
       FromStr = FromStr & "LEFT OUTER JOIN APMAS APM ON (APT.SUPCOD = APM.SUPCOD) "
   End If
   
   If Trim(DOCNUM) <> "" Then
      If WildCard(DOCNUM, SubLen, NewStr) Then
         If WhereStr = "" Then
            WhereStr = "WHERE (DOCNUM = '" & ChangeQuote(Trim(DOCNUM)) & "')"
         Else
            WhereStr = WhereStr & "AND (DOCNUM = '" & ChangeQuote(Trim(DOCNUM)) & "')"
         End If
      Else
         If WhereStr = "" Then
            WhereStr = "WHERE (DOCNUM = '" & ChangeQuote(Trim(DOCNUM)) & "')"
         Else
            WhereStr = WhereStr & "AND (DOCNUM = '" & ChangeQuote(Trim(DOCNUM)) & "')"
         End If
      End If
   End If
   
   If Trim(SUPCOD) <> "" Then
      If WildCard(SUPCOD, SubLen, NewStr) Then
         If WhereStr = "" Then
            WhereStr = "WHERE (APT.SUPCOD LIKE '" & ChangeQuote(Trim(SUPCOD)) & "')"
         Else
            WhereStr = WhereStr & "AND (APT.SUPCOD LIKE '" & ChangeQuote(Trim(SUPCOD)) & "')"
         End If
      Else
         If WhereStr = "" Then
            WhereStr = "WHERE (APT.SUPCOD = '" & ChangeQuote(Trim(SUPCOD)) & "')"
         Else
            WhereStr = WhereStr & "AND (APT.SUPCOD = '" & ChangeQuote(Trim(SUPCOD)) & "')"
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
            WhereStr = "WHERE (RECTYP = '" & RECTYP & "') "
         Else
            WhereStr = WhereStr & "AND (RECTYP = '" & RECTYP & "')"
         End If
      End If
   End If
                        
   If Trim(RecTypeSet) <> "" Then
      If WhereStr = "" Then
         WhereStr = "WHERE (APT.RECTYP IN " & RecTypeSet & ") "
      Else
         WhereStr = WhereStr & "AND (APT.RECTYP IN " & RecTypeSet & ")"
      End If
   End If

    If FROM_DUE_DATE > -1 Then
      TempDate = DateToStringIntLow(Trim(FROM_DUE_DATE))
      If WhereStr = "" Then
         WhereStr = "WHERE (DTOS(APT.DUEDAT) >= '" & ChangeQuote(TempDate) & "')"
      Else
         WhereStr = WhereStr & "AND (DTOS(APT.DUEDAT) >= '" & ChangeQuote(TempDate) & "')"
      End If
   End If

   If TO_DUE_DATE > -1 Then
      TempDate = DateToStringIntHi(Trim(TO_DUE_DATE))
      If WhereStr = "" Then
         WhereStr = "WHERE (DTOS(APT.DUEDAT) <= '" & ChangeQuote(TempDate) & "')"
      Else
         WhereStr = WhereStr & "AND (DTOS(APT.DUEDAT) <= '" & ChangeQuote(TempDate) & "')"
      End If
   End If
      
    If FROM_DOC_DATE > -1 Then
      TempDate = DateToStringIntLow(Trim(FROM_DOC_DATE))
      If WhereStr = "" Then
         WhereStr = "WHERE (DTOS(APT.DOCDAT) >= '" & ChangeQuote(TempDate) & "')"
      Else
         WhereStr = WhereStr & "AND (DTOS(APT.DOCDAT) >= '" & ChangeQuote(TempDate) & "')"
      End If
   End If

   If TO_DOC_DATE > -1 Then
      TempDate = DateToStringIntHi(Trim(TO_DOC_DATE))
      If WhereStr = "" Then
         WhereStr = "WHERE (DTOS(APT.DOCDAT) <= '" & ChangeQuote(TempDate) & "')"
      Else
         WhereStr = WhereStr & "AND (DTOS(APT.DOCDAT) <= '" & ChangeQuote(TempDate) & "')"
      End If
   End If
   
    If FROM_PAY_DATE > -1 Then
      TempDate = DateToStringIntLow(Trim(FROM_PAY_DATE))
      If WhereStr = "" Then
         WhereStr = "WHERE (DTOS(APT2.DOCDAT) >= '" & ChangeQuote(TempDate) & "')"
      Else
         WhereStr = WhereStr & "AND (DTOS(APT2.DOCDAT) >= '" & ChangeQuote(TempDate) & "')"
      End If
   End If

   If TO_PAY_DATE > -1 Then
      TempDate = DateToStringIntHi(Trim(TO_PAY_DATE))
      If WhereStr = "" Then
         WhereStr = "WHERE (DTOS(APT2.DOCDAT) <= '" & ChangeQuote(TempDate) & "')"
      Else
         WhereStr = WhereStr & "AND (DTOS(APT2.DOCDAT) <= '" & ChangeQuote(TempDate) & "')"
      End If
   End If
   
   If FROM_CHEQUE_DATE > -1 Then
      TempDate = DateToStringIntLow(Trim(FROM_CHEQUE_DATE))
      If WhereStr = "" Then
         WhereStr = "WHERE (DTOS(BKT.CHQDAT) >= '" & ChangeQuote(TempDate) & "')"
      Else
         WhereStr = WhereStr & "AND (DTOS(BKT.CHQDAT) >= '" & ChangeQuote(TempDate) & "')"
      End If
   End If

   If TO_CHEQUE_DATE > -1 Then
      TempDate = DateToStringIntHi(Trim(TO_CHEQUE_DATE))
      If WhereStr = "" Then
         WhereStr = "WHERE (DTOS(BKT.CHQDAT) <= '" & ChangeQuote(TempDate) & "')"
      Else
         WhereStr = WhereStr & "AND (DTOS(BKT.CHQDAT) <= '" & ChangeQuote(TempDate) & "')"
      End If
   End If
      
   GroupByStr = ""
   If Ind = 2 Then
      GroupByStr = "GROUP BY  APT.SUPCOD "
   ElseIf Ind = 3 Then
      GroupByStr = "GROUP BY  APT.PONUM "
   End If
   
   OrderByStr = ""
   OrderTypeStr = ""
   
   If Ind = 1 Then
      If OrderBy = 1 Then
         OrderByStr = "ORDER BY DUEDAT ASC, DOCNUM "
      ElseIf OrderBy = 2 Then
         OrderByStr = "ORDER BY DUEDAT ASC, DOCNUM "
      Else
         OrderByStr = "ORDER BY DUEDAT ASC, DOCNUM "
      End If
   ElseIf Ind = 2 Then
      OrderByStr = "ORDER BY APT.SUPCOD "
   ElseIf Ind = 3 Then
      OrderByStr = "ORDER BY APT.PONUM "
   ElseIf Ind = 4 Then
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
   'debug.print SQL1
   Call glbDatabaseMngr.GetRs(SQL1, SQL2, True, ItemCount, Rs, glbErrorLog)
End Function

Public Function PopulateFromRS(Ind As Long, Rs As ADODB.Recordset)
   If Ind = 1 Then
      RECTYP = NVLS(Rs("RECTYP"), "")
      DOCNUM = NVLS(Rs("DOCNUM"), "")
      DOCDAT = InternalDateToDate(NVLS(Rs("DOC_DATE"), ""))
      SUPCOD = NVLS(Rs("SUPCOD"), "")
      DUEDAT = InternalDateToDate(NVLS(Rs("DUE_DATE"), ""))
      AMOUNT = NVLD(Rs("AMOUNT"), 0)
      TOTAL = NVLD(Rs("TOTAL"), 0)
      SUPNAM = NVLS(Rs("SUPNAM"), "")
      PRENAM = NVLS(Rs("PRENAM"), "")
      PAYAMT = NVLD(Rs("PAYAMT"), 0)
      CMPLAPP = NVLS(Rs("CMPLAPP"), "")
      
      KEY_LOOKUP = SUPCOD
      KEY_ID = Val(SUPCOD)
   ElseIf Ind = 2 Then
      SUPCOD = NVLS(Rs("SUPCOD"), "")
      AMOUNT = NVLD(Rs("TOTAL"), 0)
      PAYAMT = NVLD(Rs("PAYAMT"), 0)
      'TOTAL = NVLD(Rs("TOTAL"), 0)
      CSHPAY = NVLD(Rs("CSHPAY"), 0)
      INTPAY = NVLD(Rs("INTPAY"), 0)
      REMAMT = NVLD(Rs("REMAMT"), 0)
   ElseIf Ind = 3 Then
      PONUM = NVLS(Rs("PONUM"), "")
      TOTAL = NVLD(Rs("TOTAL"), 0)
    ElseIf Ind = 4 Then
        DOCNUM = NVLS(Rs("DOCNUM"), "")
        CHQNUM = NVLS(Rs("CHQNUM"), "")
        SUPCOD = NVLS(Rs("SUPCOD"), "")
        CHQDAT = InternalDateToDate(NVLS(Rs("CHQDAT"), ""))
        AMOUNT = NVLD(Rs("AMOUNT"), 0)
        SUPNAM = NVLS(Rs("SUPNAM"), "")
        DOCDAT = InternalDateToDate(NVLS(Rs("DOC_DATE"), ""))
        RCPNUM = NVLS(Rs("RCPNUM"), "")
   End If
End Function

Private Sub Class_Terminate()
   Set ReceiptChecks = Nothing
   Set JournalItems = Nothing
   Set ReceiptItems = Nothing
   Set BankTrans = Nothing
End Sub
