VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CStmas"
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
Public KEY As String
' End of generic data member

Public addData As String
Public RECTYP As String
'Public DOCNUM As String
'Public DOCDAT  As Date
'Public SUPCOD  As String
'Public CUSCOD  As String
'Public PEOPLE  As String
Public STKCOD  As String
Public STKDES  As String
Public STKTYP  As String
'Public TRNQTY  As Double      'นน
'Public NETVAL  As Double      'ราคาสุทธิ
'Public DISAMT As Double        'ส่วนลด
'Public UNITPR As Double       'ราคาเฉลี่ย/หน่วย
'
'Public LOCCOD As String       'คลัง NO
'
'Public FROM_DOC_DATE As Date
'Public TO_DOC_DATE As Date
'Public FROM_DUE_DATE As Date
'Public TO_DUE_DATE As Date
'Public FROM_CMPL_DATE As Date
'Public TO_CMPL_DATE As Date
'Public FROM_SUPPLIER_CODE As String
'Public TO_SUPPLIER_CODE As String
'Public FROM_CUSTOMER_CODE As String
'Public TO_CUSTOMER_CODE  As String
'Public FROM_STOCK_CODE As String
'Public TO_STOCK_CODE As String

'Public RENUM As String
'Public FROM_SLM_DATE As Date
'Public TO_SLM_DATE As Date
'Public CMPLDAT As Date         'วันที่รับชำระ
'Public SLMNAM As String
'Public AREACOD As String
'Public SEQNUM As String
'Public TYPDES As String
'Public AREANAM As String

'Public SUPNAM As String
'Public SUPTYP As String
'Public SUPTYPSET As String
'Public SUPTYPNAM As String
'Public CUSNAM As String

'Public CHQNUM  As String
'Public CHQDAT As Date            'วันที่ DUE เช็ค
'Public GETDAT As Date            'วันที่ มารับเช็ค
'Public PAYAMT As Double

'Public FROM_SALE_CODE As String
'Public TO_SALE_CODE As String
'Public SLMCOD As String

Public KEY_ID As Long
Public KEY_LOOKUP As String
Public RecTypeSet As String

'Public eachSLM As Collection
'Public temp_eachSLM As CCondiIncenSum
'Public db2 As Boolean

Private Sub Class_Initialize()
   Set m_Conn = glbDatabaseMngr.DBConnection
   
'   FROM_DOC_DATE = -1
'   TO_DOC_DATE = -1
'   FROM_DUE_DATE = -1
'   TO_DUE_DATE = -1
'   FROM_CMPL_DATE = -1
'   TO_CMPL_DATE = -1
'   FROM_SLM_DATE = -1
'   TO_SLM_DATE = -1
   
End Sub

'In this class we will not handle any error or any transaction
Public Function QueryData(Ind As Long, Rs As ADODB.Recordset, ItemCount As Long, Optional db_num As Long = 1)
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
Dim i As Long
   
   If Ind = 1 Then
      SelectStr1 = "SELECT STM.* "
       SelectStr2 = "SELECT COUNT(*)  "
      
      FromStr = " FROM STMAS STM "
       WhereStr = " WHERE (STM.STKTYP = '0') "
  End If
   
   If Trim(STKDES) <> "" Then
      If WhereStr = "" Then
         WhereStr = "WHERE (STM.STKDES = '" & ChangeQuote(Trim(STKDES)) & "') "
      Else
         WhereStr = WhereStr & "AND (STM.STKDES = '" & ChangeQuote(Trim(STKDES)) & "') "
      End If
   End If
   
   If Trim(STKCOD) <> "" Then
      If WhereStr = "" Then
         WhereStr = "WHERE (STM.STKCOD = '" & ChangeQuote(Trim(STKCOD)) & "') "
      Else
         WhereStr = WhereStr & "AND (STM.STKCOD = '" & ChangeQuote(Trim(STKCOD)) & "') "
      End If
   End If

  If OrderBy = 1 Then
      OrderByStr = " ORDER BY  STM.STKDES    "
   ElseIf OrderBy = 2 Then
      OrderByStr = " ORDER BY  STM.STKDES   "
   Else
      OrderByStr = " ORDER BY  STM.STKCOD   "
  End If
   
   If OrderType = 1 Then
      OrderTypeStr = "ASC "
   ElseIf OrderType = 2 Then
      OrderTypeStr = "DESC "
   Else
      OrderTypeStr = "ASC "
   End If
   
  If Len(OrderByStr) <= 0 Then
   OrderTypeStr = ""
  End If
  
   SQL1 = SelectStr1 & FromStr & WhereStr & GroupByStr & OrderByStr & OrderTypeStr
   SQL2 = SelectStr2 & FromStr & WhereStr
   'debug.print SQL1
'   If db_num = 2 Then
'      Call glbDatabaseMngr.GetRs2(SQL1, SQL2, True, ItemCount, Rs, glbErrorLog)
'   Else
      Call glbDatabaseMngr.GetRs(SQL1, SQL2, True, ItemCount, Rs, glbErrorLog)
'   End If
   
End Function

Public Function PopulateFromRS(Ind As Long, Rs As ADODB.Recordset, Optional Count As Long)
   If Ind = 1 Then
      STKCOD = NVLS(Rs("STKCOD"), "")
      STKDES = NVLS(Rs("STKDES"), "")
      STKTYP = NVLS(Rs("STKTYP"), "")
      
      KEY_ID = Count                       ' สำหรับ lookup
      KEY_LOOKUP = NVLS(Rs("STKCOD"), "")
   End If
End Function

