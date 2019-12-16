VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsDaily"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private m_ModuleName As String
Private m_ADOConn As ADODB.Connection
Private m_Rs1 As ADODB.Recordset
Private m_Rs2 As ADODB.Recordset
Private m_Rs3 As ADODB.Recordset

Private Sub Class_Initialize()
   Set m_ADOConn = glbDatabaseMngr.DBConnection
   m_ModuleName = "clsDaily"
   Set m_Rs1 = New ADODB.Recordset
   Set m_Rs2 = New ADODB.Recordset
   Set m_Rs3 = New ADODB.Recordset
End Sub

Private Sub Class_Terminate()
   If m_Rs1.State = adStateOpen Then
      m_Rs1.Close
   End If
   Set m_Rs1 = Nothing

   If m_Rs2.State = adStateOpen Then
      m_Rs2.Close
   End If
   Set m_Rs2 = Nothing

   If m_Rs3.State = adStateOpen Then
      m_Rs3.Close
   End If
   Set m_Rs3 = Nothing
End Sub

Public Sub StartTransaction()
   m_ADOConn.BeginTrans
End Sub

Public Sub CommitTransaction()
   m_ADOConn.CommitTrans
End Sub

Public Sub RollbackTransaction()
   m_ADOConn.RollbackTrans
End Sub

Public Function QueryReportConfig(Ua As CReportConfig, Rs As ADODB.Recordset, ItemCount As Long, IsOK As Boolean, ErrorObj As clsErrorLog) As Boolean
On Error GoTo ErrorHandler
Dim RName As String
Dim SQL1 As String
Dim SQL2 As String
Dim SelectStr As String
Dim FromStr As String
Dim WhereStr As String
Dim OrderStr As String
Dim OrderType As String
Dim NewStr As String
Dim SubLen As Long
Dim iCount As Long

   RName = "QueryReportConfig"
   QueryReportConfig = False

   IsOK = True
   Call Ua.QueryData(Rs, ItemCount)
   
   QueryReportConfig = True
   Exit Function

ErrorHandler:
   ErrorObj.SystemErrorMsg = Err.DESCRIPTION
   ErrorObj.RoutineName = RName
   ErrorObj.ModuleName = m_ModuleName
   ErrorObj.ShowErrorLog (LOG_FILE_MSGBOX)

   QueryReportConfig = False
End Function

Public Function AddEditReportConfig(Ug As CReportConfig, IsOK As Boolean, AutoCommit As Boolean, ErrorObj As clsErrorLog)
On Error GoTo ErrorHandler
Dim RName As String
Dim HasBegin As Boolean

   RName = "AddEditReportConfig"
   AddEditReportConfig = False
   HasBegin = False

   If AutoCommit Then
      m_ADOConn.BeginTrans
   End If
   HasBegin = True

   Call Ug.AddEditData

   If AutoCommit Then
      Call m_ADOConn.CommitTrans
   End If
   HasBegin = False
   IsOK = True

   AddEditReportConfig = True
   Exit Function

ErrorHandler:
   If HasBegin Then
      If AutoCommit Then
         m_ADOConn.RollbackTrans
      End If
   End If

   ErrorObj.LocalErrorMsg = "Runtime error."
   ErrorObj.SystemErrorMsg = Err.DESCRIPTION
   ErrorObj.RoutineName = RName
   ErrorObj.ModuleName = m_ModuleName
   ErrorObj.ShowErrorLog (LOG_TO_FILE)

   IsOK = False
   AddEditReportConfig = False
End Function

Public Function QuerySupplier(Ua As CAPMas, Rs As ADODB.Recordset, ItemCount As Long, IsOK As Boolean, ErrorObj As clsErrorLog) As Boolean
On Error GoTo ErrorHandler
Dim RName As String
Dim SQL1 As String
Dim SQL2 As String
Dim SelectStr As String
Dim FromStr As String
Dim WhereStr As String
Dim OrderStr As String
Dim OrderType As String
Dim NewStr As String
Dim SubLen As Long
Dim iCount As Long

   RName = "QuerySupplier"
   QuerySupplier = False

   IsOK = True
   Call Ua.QueryData(1, Rs, ItemCount)
   
   QuerySupplier = True
   Exit Function

ErrorHandler:
   ErrorObj.SystemErrorMsg = Err.DESCRIPTION
   ErrorObj.RoutineName = RName
   ErrorObj.ModuleName = m_ModuleName
   ErrorObj.ShowErrorLog (LOG_FILE_MSGBOX)

   QuerySupplier = False
End Function
Public Function QueryDocket(Ua As CGLJnl, Rs As ADODB.Recordset, ItemCount As Long, IsOK As Boolean, ErrorObj As clsErrorLog) As Boolean
On Error GoTo ErrorHandler
Dim RName As String
Dim SQL1 As String
Dim SQL2 As String
Dim SelectStr As String
Dim FromStr As String
Dim WhereStr As String
Dim OrderStr As String
Dim OrderType As String
Dim NewStr As String
Dim SubLen As Long
Dim iCount As Long

   RName = "QueryDocket"
   QueryDocket = False

   IsOK = True
   Call Ua.QueryData(4, Rs, ItemCount)
   
   QueryDocket = True
   Exit Function

ErrorHandler:
   ErrorObj.SystemErrorMsg = Err.DESCRIPTION
   ErrorObj.RoutineName = RName
   ErrorObj.ModuleName = m_ModuleName
   ErrorObj.ShowErrorLog (LOG_FILE_MSGBOX)

   QueryDocket = False
End Function
Public Function QueryAsset(Ua As CFaMas, Rs As ADODB.Recordset, ItemCount As Long, IsOK As Boolean, ErrorObj As clsErrorLog) As Boolean
On Error GoTo ErrorHandler
Dim RName As String
Dim SQL1 As String
Dim SQL2 As String
Dim SelectStr As String
Dim FromStr As String
Dim WhereStr As String
Dim OrderStr As String
Dim OrderType As String
Dim NewStr As String
Dim SubLen As Long
Dim iCount As Long

   RName = "QueryAsset"
   QueryAsset = False

   IsOK = True
   Call Ua.QueryData(1, Rs, ItemCount)
   
   QueryAsset = True
   Exit Function

ErrorHandler:
   ErrorObj.SystemErrorMsg = Err.DESCRIPTION
   ErrorObj.RoutineName = RName
   ErrorObj.ModuleName = m_ModuleName
   ErrorObj.ShowErrorLog (LOG_FILE_MSGBOX)

   QueryAsset = False
End Function

Public Function QueryCustomer(Ua As CARMas, Rs As ADODB.Recordset, ItemCount As Long, IsOK As Boolean, ErrorObj As clsErrorLog, Optional db As Long = 1) As Boolean
On Error GoTo ErrorHandler
Dim RName As String
Dim SQL1 As String
Dim SQL2 As String
Dim SelectStr As String
Dim FromStr As String
Dim WhereStr As String
Dim OrderStr As String
Dim OrderType As String
Dim NewStr As String
Dim SubLen As Long
Dim iCount As Long

   RName = "QueryCustomer"
   QueryCustomer = False

   IsOK = True
   Call Ua.QueryData(1, Rs, ItemCount, db)
   
   QueryCustomer = True
   Exit Function

ErrorHandler:
   ErrorObj.SystemErrorMsg = Err.DESCRIPTION
   ErrorObj.RoutineName = RName
   ErrorObj.ModuleName = m_ModuleName
   ErrorObj.ShowErrorLog (LOG_FILE_MSGBOX)

   QueryCustomer = False
End Function
Public Function QuerySale(Ua As COESLM, Rs As ADODB.Recordset, ItemCount As Long, IsOK As Boolean, ErrorObj As clsErrorLog) As Boolean
On Error GoTo ErrorHandler
Dim RName As String
Dim SQL1 As String
Dim SQL2 As String
Dim SelectStr As String
Dim FromStr As String
Dim WhereStr As String
Dim OrderStr As String
Dim OrderType As String
Dim NewStr As String
Dim SubLen As Long
Dim iCount As Long

   RName = "QuerySale"
   QuerySale = False

   IsOK = True
   Call Ua.QueryData(1, Rs, ItemCount)
   
   QuerySale = True
   Exit Function

ErrorHandler:
   ErrorObj.SystemErrorMsg = Err.DESCRIPTION
   ErrorObj.RoutineName = RName
   ErrorObj.ModuleName = m_ModuleName
   ErrorObj.ShowErrorLog (LOG_FILE_MSGBOX)
   
   QuerySale = False
End Function

Public Function QueryCompany(Ua As CSCComp, Rs As ADODB.Recordset, ItemCount As Long, IsOK As Boolean, ErrorObj As clsErrorLog) As Boolean
On Error GoTo ErrorHandler
Dim RName As String
Dim SQL1 As String
Dim SQL2 As String
Dim SelectStr As String
Dim FromStr As String
Dim WhereStr As String
Dim OrderStr As String
Dim OrderType As String
Dim NewStr As String
Dim SubLen As Long
Dim iCount As Long

   RName = "QueryCompany"
   QueryCompany = False

   IsOK = True
   Call Ua.QueryData(1, Rs, ItemCount)
   
   QueryCompany = True
   Exit Function

ErrorHandler:
   ErrorObj.SystemErrorMsg = Err.DESCRIPTION
   ErrorObj.RoutineName = RName
   ErrorObj.ModuleName = m_ModuleName
   ErrorObj.ShowErrorLog (LOG_FILE_MSGBOX)

   QueryCompany = False
End Function

Public Function QueryAPTran(Ua As CApTrn, Rs As ADODB.Recordset, ItemCount As Long, IsOK As Boolean, ErrorObj As clsErrorLog) As Boolean
On Error GoTo ErrorHandler
Dim RName As String
Dim SQL1 As String
Dim SQL2 As String
Dim SelectStr As String
Dim FromStr As String
Dim WhereStr As String
Dim OrderStr As String
Dim OrderType As String
Dim NewStr As String
Dim SubLen As Long
Dim iCount As Long
Dim Jnl As CGLJnl
Dim Rcc As CAPRcpCq
Dim Rci As CAPRcIt
Dim Bkt As CBkTrn

   RName = "QueryAPTran"
   QueryAPTran = False

   IsOK = True
   Call Ua.QueryData(1, Rs, ItemCount)
   
   If Ua.QueryFlag = 1 Then
      Set Ua.JournalItems = Nothing
      Set Ua.JournalItems = New Collection
      Set Jnl = New CGLJnl
      Jnl.VOUCHER = Ua.DOCNUM
      Call Jnl.QueryData(1, m_Rs1, iCount)
      Set Jnl = Nothing
      While Not m_Rs1.EOF
         Set Jnl = New CGLJnl
         Call Jnl.PopulateFromRS(1, m_Rs1)
         Jnl.Flag = "I"
         Call Ua.JournalItems.Add(Jnl)
         Set Jnl = Nothing
         m_Rs1.MoveNext
      Wend
   
      Set Ua.ReceiptChecks = Nothing
      Set Ua.ReceiptChecks = New Collection
      Set Rcc = New CAPRcpCq
      Rcc.RCPNUM = Ua.DOCNUM
      Call Rcc.QueryData(1, m_Rs1, iCount)
      Set Rcc = Nothing
      While Not m_Rs1.EOF
         Set Rcc = New CAPRcpCq
         Call Rcc.PopulateFromRS(1, m_Rs1)
         Rcc.Flag = "I"
         Call Ua.ReceiptChecks.Add(Rcc)
         Set Rcc = Nothing
         m_Rs1.MoveNext
      Wend
   
      Set Ua.ReceiptItems = Nothing
      Set Ua.ReceiptItems = New Collection
      Set Rci = New CAPRcIt
      Rci.RCPNUM = Ua.DOCNUM
      Call Rci.QueryData(1, m_Rs1, iCount)
      Set Rci = Nothing
      While Not m_Rs1.EOF
         Set Rci = New CAPRcIt
         Call Rci.PopulateFromRS(1, m_Rs1)
         Rci.Flag = "I"
         Call Ua.ReceiptItems.Add(Rci)
         Set Rci = Nothing
         m_Rs1.MoveNext
      Wend
   
      Set Ua.BankTrans = Nothing
      Set Ua.BankTrans = New Collection
      Set Bkt = New CBkTrn
      Bkt.VOUCHER = Ua.DOCNUM
      Call Bkt.QueryData(1, m_Rs1, iCount)
      Set Bkt = Nothing
      While Not m_Rs1.EOF
         Set Bkt = New CBkTrn
         Call Bkt.PopulateFromRS(1, m_Rs1)
         Bkt.Flag = "I"
         Call Ua.BankTrans.Add(Bkt)
         Set Bkt = Nothing
         m_Rs1.MoveNext
      Wend
   End If
   
   QueryAPTran = True
   Exit Function

ErrorHandler:
   ErrorObj.SystemErrorMsg = Err.DESCRIPTION
   ErrorObj.RoutineName = RName
   ErrorObj.ModuleName = m_ModuleName
   ErrorObj.ShowErrorLog (LOG_FILE_MSGBOX)

   QueryAPTran = False
End Function

Public Function QueryARTran(Ua As CARTrn, Rs As ADODB.Recordset, ItemCount As Long, IsOK As Boolean, ErrorObj As clsErrorLog) As Boolean
On Error GoTo ErrorHandler
Dim RName As String
Dim SQL1 As String
Dim SQL2 As String
Dim SelectStr As String
Dim FromStr As String
Dim WhereStr As String
Dim OrderStr As String
Dim OrderType As String
Dim NewStr As String
Dim SubLen As Long
Dim iCount As Long
Dim Jnl As CGLJnl
Dim Rcc As CAPRcpCq
Dim Rci As CARRcIt
Dim Bkt As CBkTrn

   RName = "QueryARTran"
   QueryARTran = False

   IsOK = True
   Call Ua.QueryData(1, Rs, ItemCount)
   
   If Ua.QueryFlag = 1 Then
      Set Ua.JournalItems = Nothing
      Set Ua.JournalItems = New Collection
      Set Jnl = New CGLJnl
      Jnl.VOUCHER = Ua.DOCNUM
      Call Jnl.QueryData(1, m_Rs1, iCount)
      Set Jnl = Nothing
      While Not m_Rs1.EOF
         Set Jnl = New CGLJnl
         Call Jnl.PopulateFromRS(1, m_Rs1)
         Jnl.Flag = "I"
         Call Ua.JournalItems.Add(Jnl)
         Set Jnl = Nothing
         m_Rs1.MoveNext
      Wend
   
      Set Ua.ReceiptChecks = Nothing
      Set Ua.ReceiptChecks = New Collection
      Set Rcc = New CAPRcpCq
      Rcc.RCPNUM = Ua.DOCNUM
      Call Rcc.QueryData(1, m_Rs1, iCount)
      Set Rcc = Nothing
      While Not m_Rs1.EOF
         Set Rcc = New CAPRcpCq
         Call Rcc.PopulateFromRS(1, m_Rs1)
         Rcc.Flag = "I"
         Call Ua.ReceiptChecks.Add(Rcc)
         Set Rcc = Nothing
         m_Rs1.MoveNext
      Wend
   
      Set Ua.ReceiptItems = Nothing
      Set Ua.ReceiptItems = New Collection
      Set Rci = New CARRcIt
      Rci.RCPNUM = Ua.DOCNUM
      Call Rci.QueryData(1, m_Rs1, iCount)
      Set Rci = Nothing
      While Not m_Rs1.EOF
         Set Rci = New CARRcIt
         Call Rci.PopulateFromRS(1, m_Rs1)
         Rci.Flag = "I"
         Call Ua.ReceiptItems.Add(Rci)
         Set Rci = Nothing
         m_Rs1.MoveNext
      Wend
   
      Set Ua.BankTrans = Nothing
      Set Ua.BankTrans = New Collection
      Set Bkt = New CBkTrn
      Bkt.VOUCHER = Ua.DOCNUM
      Call Bkt.QueryData(1, m_Rs1, iCount)
      Set Bkt = Nothing
      While Not m_Rs1.EOF
         Set Bkt = New CBkTrn
         Call Bkt.PopulateFromRS(1, m_Rs1)
         Bkt.Flag = "I"
         Call Ua.BankTrans.Add(Bkt)
         Set Bkt = Nothing
         m_Rs1.MoveNext
      Wend
   End If

   QueryARTran = True
   Exit Function

ErrorHandler:
   ErrorObj.SystemErrorMsg = Err.DESCRIPTION
   ErrorObj.RoutineName = RName
   ErrorObj.ModuleName = m_ModuleName
   ErrorObj.ShowErrorLog (LOG_FILE_MSGBOX)

   QueryARTran = False
End Function

Public Function QueryDistinctGlJnl(Ua As CGLJnl, Rs As ADODB.Recordset, ItemCount As Long, IsOK As Boolean, ErrorObj As clsErrorLog) As Boolean
On Error GoTo ErrorHandler
Dim RName As String
Dim SQL1 As String
Dim SQL2 As String
Dim SelectStr As String
Dim FromStr As String
Dim WhereStr As String
Dim OrderStr As String
Dim OrderType As String
Dim NewStr As String
Dim SubLen As Long
Dim iCount As Long

   RName = "QueryDistinctGlJnl"
   QueryDistinctGlJnl = False

   IsOK = True
   Call Ua.QueryData(2, Rs, ItemCount)
   
   If Ua.QueryFlag = 1 Then
   End If
   
   QueryDistinctGlJnl = True
   Exit Function

ErrorHandler:
   ErrorObj.SystemErrorMsg = Err.DESCRIPTION
   ErrorObj.RoutineName = RName
   ErrorObj.ModuleName = m_ModuleName
   ErrorObj.ShowErrorLog (LOG_FILE_MSGBOX)

   QueryDistinctGlJnl = False
End Function
Public Function QueryRealCredit(Ua As CRealCredit, Rs As ADODB.Recordset, ItemCount As Long, IsOK As Boolean, ErrorObj As clsErrorLog) As Boolean
On Error GoTo ErrorHandler
Dim RName As String
Dim SQL1 As String
Dim SQL2 As String
Dim SelectStr As String
Dim FromStr As String
Dim WhereStr As String
Dim OrderStr As String
Dim OrderType As String
Dim NewStr As String
Dim SubLen As Long
Dim iCount As Long

   RName = "QueryRealCredit"
   QueryRealCredit = False

   IsOK = True
   Call Ua.QueryData(Rs, ItemCount)
   
   QueryRealCredit = True
   Exit Function

ErrorHandler:
   ErrorObj.SystemErrorMsg = Err.DESCRIPTION
   ErrorObj.RoutineName = RName
   ErrorObj.ModuleName = m_ModuleName
   ErrorObj.ShowErrorLog (LOG_FILE_MSGBOX)

   QueryRealCredit = False
End Function
Public Function AddEditRealCredit(Ug As CRealCredit, IsOK As Boolean, AutoCommit As Boolean, ErrorObj As clsErrorLog)
On Error GoTo ErrorHandler
Dim RName As String
Dim HasBegin As Boolean

   RName = "AddEditRealCredit"
   AddEditRealCredit = False
   HasBegin = False

   If AutoCommit Then
      m_ADOConn.BeginTrans
   End If
   HasBegin = True

   Call Ug.AddEditData

   If AutoCommit Then
      Call m_ADOConn.CommitTrans
   End If
   HasBegin = False
   IsOK = True

   AddEditRealCredit = True
   Exit Function

ErrorHandler:
   If HasBegin Then
      If AutoCommit Then
         m_ADOConn.RollbackTrans
      End If
   End If

   ErrorObj.LocalErrorMsg = "Runtime error."
   ErrorObj.SystemErrorMsg = Err.DESCRIPTION
   ErrorObj.RoutineName = RName
   ErrorObj.ModuleName = m_ModuleName
   ErrorObj.ShowErrorLog (LOG_TO_FILE)

   IsOK = False
   AddEditRealCredit = False
End Function
Public Function DeleteRealCredit(Ug As CRealCredit, IsOK As Boolean, AutoCommit As Boolean, ErrorObj As clsErrorLog)
On Error GoTo ErrorHandler
Dim RName As String
Dim SQL As String
Dim HasBegin As Boolean

   RName = "DeleteRealCredit"
   DeleteRealCredit = False
   HasBegin = False
   
   If AutoCommit Then
      m_ADOConn.BeginTrans
   End If
   HasBegin = True
   
   Call Ug.DeleteData
   
   If AutoCommit Then
      Call m_ADOConn.CommitTrans
   End If
   HasBegin = False
   IsOK = True

   DeleteRealCredit = True
   Exit Function
   
ErrorHandler:
   If HasBegin Then
      If AutoCommit Then
         m_ADOConn.RollbackTrans
      End If
   End If
   
   ErrorObj.LocalErrorMsg = "Runtime error."
   ErrorObj.SystemErrorMsg = Err.DESCRIPTION
   ErrorObj.RoutineName = RName
   ErrorObj.ModuleName = m_ModuleName
   ErrorObj.LocalErrorMsg = "ไม่สามารถลบข้อมูลได้เนื่องจากข้อมูลนี้ยังถูกอ้างถึงอยู่จากส่วนอื่น"
   ErrorObj.ShowErrorLog (LOG_TO_FILE)
   
   IsOK = False
   
   DeleteRealCredit = False
End Function

Public Function QueryComMaxSubG(Ua As CComboSubGroup, Rs As ADODB.Recordset, ItemCount As Long, IsOK As Boolean, ErrorObj As clsErrorLog) As Boolean
On Error GoTo ErrorHandler
Dim RName As String
Dim SQL1 As String
Dim SQL2 As String
Dim SelectStr As String
Dim FromStr As String
Dim WhereStr As String
Dim OrderStr As String
Dim OrderType As String
Dim NewStr As String
Dim SubLen As Long
Dim iCount As Long

   RName = "QueryComMaxSubG"
   QueryComMaxSubG = False

   IsOK = True
   Call Ua.QueryData(2, Rs, ItemCount)
   
   QueryComMaxSubG = True
   Exit Function

ErrorHandler:
   ErrorObj.SystemErrorMsg = Err.DESCRIPTION
   ErrorObj.RoutineName = RName
   ErrorObj.ModuleName = m_ModuleName
   ErrorObj.ShowErrorLog (LOG_FILE_MSGBOX)

   QueryComMaxSubG = False
End Function


Public Function QueryComboSubGroup(Ua As CComboSubGroup, Rs As ADODB.Recordset, ItemCount As Long, IsOK As Boolean, ErrorObj As clsErrorLog) As Boolean
On Error GoTo ErrorHandler
Dim RName As String
Dim SQL1 As String
Dim SQL2 As String
Dim SelectStr As String
Dim FromStr As String
Dim WhereStr As String
Dim OrderStr As String
Dim OrderType As String
Dim NewStr As String
Dim SubLen As Long
Dim iCount As Long

   RName = "QueryComboSubGroup"
   QueryComboSubGroup = False

   IsOK = True
   Call Ua.QueryData(1, Rs, ItemCount)
   
   QueryComboSubGroup = True
   Exit Function

ErrorHandler:
   ErrorObj.SystemErrorMsg = Err.DESCRIPTION
   ErrorObj.RoutineName = RName
   ErrorObj.ModuleName = m_ModuleName
   ErrorObj.ShowErrorLog (LOG_FILE_MSGBOX)

   QueryComboSubGroup = False
End Function
Public Function QueryGroupType(Ua As CGroupType, Rs As ADODB.Recordset, ItemCount As Long, IsOK As Boolean, ErrorObj As clsErrorLog) As Boolean
On Error GoTo ErrorHandler
Dim RName As String
Dim SQL1 As String
Dim SQL2 As String
Dim SelectStr As String
Dim FromStr As String
Dim WhereStr As String
Dim OrderStr As String
Dim OrderType As String
Dim NewStr As String
Dim SubLen As Long
Dim iCount As Long

   RName = "QueryGroupType"
   QueryGroupType = False

   IsOK = True
   Call Ua.QueryData(Rs, ItemCount)
   
   QueryGroupType = True
   Exit Function

ErrorHandler:
   ErrorObj.SystemErrorMsg = Err.DESCRIPTION
   ErrorObj.RoutineName = RName
   ErrorObj.ModuleName = m_ModuleName
   ErrorObj.ShowErrorLog (LOG_FILE_MSGBOX)

   QueryGroupType = False
End Function
Public Function DeleteXlsUnit(Ug As CXlsUnit, IsOK As Boolean, AutoCommit As Boolean, ErrorObj As clsErrorLog)
On Error GoTo ErrorHandler
Dim RName As String
Dim SQL As String
Dim HasBegin As Boolean

   RName = DeleteXlsUnit
   DeleteXlsUnit = False
   HasBegin = False
   
   If AutoCommit Then
      m_ADOConn.BeginTrans
   End If
   HasBegin = True
   
   Call Ug.DeleteData
   
   If AutoCommit Then
      Call m_ADOConn.CommitTrans
   End If
   HasBegin = False
   IsOK = True

   DeleteXlsUnit = True
   Exit Function
   
ErrorHandler:
   If HasBegin Then
      If AutoCommit Then
         m_ADOConn.RollbackTrans
      End If
   End If
   
   ErrorObj.LocalErrorMsg = "Runtime error."
   ErrorObj.SystemErrorMsg = Err.DESCRIPTION
   ErrorObj.RoutineName = RName
   ErrorObj.ModuleName = m_ModuleName
   ErrorObj.LocalErrorMsg = "ไม่สามารถลบข้อมูลได้เนื่องจากข้อมูลนี้ยังถูกอ้างถึงอยู่จากส่วนอื่น"
   ErrorObj.ShowErrorLog (LOG_TO_FILE)
   
   IsOK = False
   
   DeleteXlsUnit = False
End Function

Public Function AddEditXlsUnit(Ug As CXlsUnit, IsOK As Boolean, AutoCommit As Boolean, ErrorObj As clsErrorLog)
On Error GoTo ErrorHandler
Dim RName As String
Dim HasBegin As Boolean

   RName = AddEditXlsUnit
   AddEditXlsUnit = False
   HasBegin = False

   If AutoCommit Then
      m_ADOConn.BeginTrans
   End If
   HasBegin = True

   Call Ug.AddEditData

   If AutoCommit Then
      Call m_ADOConn.CommitTrans
   End If
   HasBegin = False
   IsOK = True
   
   AddEditXlsUnit = True
   Exit Function

ErrorHandler:
   If HasBegin Then
      If AutoCommit Then
         m_ADOConn.RollbackTrans
      End If
   End If

   ErrorObj.LocalErrorMsg = "Runtime error."
   ErrorObj.SystemErrorMsg = Err.DESCRIPTION
   ErrorObj.RoutineName = RName
   ErrorObj.ModuleName = m_ModuleName
   ErrorObj.ShowErrorLog (LOG_TO_FILE)

   IsOK = False
   AddEditXlsUnit = False
End Function

Public Function QueryXlsUnit(Ua As CXlsUnit, Rs As ADODB.Recordset, ItemCount As Long, IsOK As Boolean, ErrorObj As clsErrorLog) As Boolean
On Error GoTo ErrorHandler
Dim RName As String
Dim SQL1 As String
Dim SQL2 As String
Dim SelectStr As String
Dim FromStr As String
Dim WhereStr As String
Dim OrderStr As String
Dim OrderType As String
Dim NewStr As String
Dim SubLen As Long
Dim iCount As Long

   RName = QueryXlsUnit
   QueryXlsUnit = False

   IsOK = True
   Call Ua.QueryData(1, Rs, ItemCount)
   
   QueryXlsUnit = True
   Exit Function

ErrorHandler:
   ErrorObj.SystemErrorMsg = Err.DESCRIPTION
   ErrorObj.RoutineName = RName
   ErrorObj.ModuleName = m_ModuleName
   ErrorObj.ShowErrorLog (LOG_FILE_MSGBOX)

   QueryXlsUnit = False
End Function

Public Function QueryXlsMainFarm(Ua As CXlsMainfarm, Rs As ADODB.Recordset, ItemCount As Long, IsOK As Boolean, ErrorObj As clsErrorLog) As Boolean
On Error GoTo ErrorHandler
Dim RName As String
Dim SQL1 As String
Dim SQL2 As String
Dim SelectStr As String
Dim FromStr As String
Dim WhereStr As String
Dim OrderStr As String
Dim OrderType As String
Dim NewStr As String
Dim SubLen As Long
Dim iCount As Long

   RName = QueryXlsMainFarm
   QueryXlsMainFarm = False

   IsOK = True
   Call Ua.QueryData(1, Rs, ItemCount)
   
   QueryXlsMainFarm = True
   Exit Function

ErrorHandler:
   ErrorObj.SystemErrorMsg = Err.DESCRIPTION
   ErrorObj.RoutineName = RName
   ErrorObj.ModuleName = m_ModuleName
   ErrorObj.ShowErrorLog (LOG_FILE_MSGBOX)

   QueryXlsMainFarm = False
End Function

Public Function AddEditXlsMainFarm(Ug As CXlsMainfarm, IsOK As Boolean, AutoCommit As Boolean, ErrorObj As clsErrorLog)
On Error GoTo ErrorHandler
Dim RName As String
Dim HasBegin As Boolean

   RName = AddEditXlsMainFarm
   AddEditXlsMainFarm = False
   HasBegin = False

   If AutoCommit Then
      m_ADOConn.BeginTrans
   End If
   HasBegin = True

   Call Ug.AddEditData

   If AutoCommit Then
      Call m_ADOConn.CommitTrans
   End If
   HasBegin = False
   IsOK = True
   
   AddEditXlsMainFarm = True
   Exit Function

ErrorHandler:
   If HasBegin Then
      If AutoCommit Then
         m_ADOConn.RollbackTrans
      End If
   End If

   ErrorObj.LocalErrorMsg = "Runtime error."
   ErrorObj.SystemErrorMsg = Err.DESCRIPTION
   ErrorObj.RoutineName = RName
   ErrorObj.ModuleName = m_ModuleName
   ErrorObj.ShowErrorLog (LOG_TO_FILE)

   IsOK = False
   AddEditXlsMainFarm = False
End Function

Public Function DeleteXlsMainFarm(Ug As CXlsMainfarm, IsOK As Boolean, AutoCommit As Boolean, ErrorObj As clsErrorLog)
On Error GoTo ErrorHandler
Dim RName As String
Dim SQL As String
Dim HasBegin As Boolean

   RName = DeleteXlsMainFarm
   DeleteXlsMainFarm = False
   HasBegin = False
   
   If AutoCommit Then
      m_ADOConn.BeginTrans
   End If
   HasBegin = True
   
   Call Ug.DeleteData
   
   If AutoCommit Then
      Call m_ADOConn.CommitTrans
   End If
   HasBegin = False
   IsOK = True

   DeleteXlsMainFarm = True
   Exit Function
   
ErrorHandler:
   If HasBegin Then
      If AutoCommit Then
         m_ADOConn.RollbackTrans
      End If
   End If
   
   ErrorObj.LocalErrorMsg = "Runtime error."
   ErrorObj.SystemErrorMsg = Err.DESCRIPTION
   ErrorObj.RoutineName = RName
   ErrorObj.ModuleName = m_ModuleName
   ErrorObj.LocalErrorMsg = "ไม่สามารถลบข้อมูลได้เนื่องจากข้อมูลนี้ยังถูกอ้างถึงอยู่จากส่วนอื่น"
   ErrorObj.ShowErrorLog (LOG_TO_FILE)
   
   IsOK = False
   
   DeleteXlsMainFarm = False
End Function

Public Function QuerySubGroupType(Ua As CSubGroupType, Rs As ADODB.Recordset, ItemCount As Long, IsOK As Boolean, ErrorObj As clsErrorLog) As Boolean
On Error GoTo ErrorHandler
Dim RName As String
Dim SQL1 As String
Dim SQL2 As String
Dim SelectStr As String
Dim FromStr As String
Dim WhereStr As String
Dim OrderStr As String
Dim OrderType As String
Dim NewStr As String
Dim SubLen As Long
Dim iCount As Long

   RName = "QuerySubGroupType"
   QuerySubGroupType = False

   IsOK = True
   Call Ua.QueryData(Rs, ItemCount)
   
   QuerySubGroupType = True
   Exit Function

ErrorHandler:
   ErrorObj.SystemErrorMsg = Err.DESCRIPTION
   ErrorObj.RoutineName = RName
   ErrorObj.ModuleName = m_ModuleName
   ErrorObj.ShowErrorLog (LOG_FILE_MSGBOX)

   QuerySubGroupType = False
End Function
Public Function AddEditGroupType(Ug As CGroupType, IsOK As Boolean, AutoCommit As Boolean, ErrorObj As clsErrorLog)
On Error GoTo ErrorHandler
Dim RName As String
Dim HasBegin As Boolean

   RName = "AddEditGroupType"
   AddEditGroupType = False
   HasBegin = False

   If AutoCommit Then
      m_ADOConn.BeginTrans
   End If
   HasBegin = True

   Call Ug.AddEditData

   If AutoCommit Then
      Call m_ADOConn.CommitTrans
   End If
   HasBegin = False
   IsOK = True
   
   AddEditGroupType = True
   Exit Function

ErrorHandler:
   If HasBegin Then
      If AutoCommit Then
         m_ADOConn.RollbackTrans
      End If
   End If

   ErrorObj.LocalErrorMsg = "Runtime error."
   ErrorObj.SystemErrorMsg = Err.DESCRIPTION
   ErrorObj.RoutineName = RName
   ErrorObj.ModuleName = m_ModuleName
   ErrorObj.ShowErrorLog (LOG_TO_FILE)

   IsOK = False
   AddEditGroupType = False
End Function
Public Function AddEditSubGroupType(Ug As CSubGroupType, IsOK As Boolean, AutoCommit As Boolean, ErrorObj As clsErrorLog)
On Error GoTo ErrorHandler
Dim RName As String
Dim HasBegin As Boolean

   RName = "AddEditSubGroupType"
   AddEditSubGroupType = False
   HasBegin = False

   If AutoCommit Then
      m_ADOConn.BeginTrans
   End If
   HasBegin = True

   Call Ug.AddEditData

   If AutoCommit Then
      Call m_ADOConn.CommitTrans
   End If
   HasBegin = False
   IsOK = True
   
   AddEditSubGroupType = True
   Exit Function

ErrorHandler:
   If HasBegin Then
      If AutoCommit Then
         m_ADOConn.RollbackTrans
      End If
   End If

   ErrorObj.LocalErrorMsg = "Runtime error."
   ErrorObj.SystemErrorMsg = Err.DESCRIPTION
   ErrorObj.RoutineName = RName
   ErrorObj.ModuleName = m_ModuleName
   ErrorObj.ShowErrorLog (LOG_TO_FILE)

   IsOK = False
   AddEditSubGroupType = False
End Function
Public Function DeleteGroupType(Ug As CGroupType, IsOK As Boolean, AutoCommit As Boolean, ErrorObj As clsErrorLog)
On Error GoTo ErrorHandler
Dim RName As String
Dim SQL As String
Dim HasBegin As Boolean

   RName = "DeleteGroupType"
   DeleteGroupType = False
   HasBegin = False
   
   If AutoCommit Then
      m_ADOConn.BeginTrans
   End If
   HasBegin = True
   
   Call Ug.DeleteData
   
   If AutoCommit Then
      Call m_ADOConn.CommitTrans
   End If
   HasBegin = False
   IsOK = True

   DeleteGroupType = True
   Exit Function
   
ErrorHandler:
   If HasBegin Then
      If AutoCommit Then
         m_ADOConn.RollbackTrans
      End If
   End If
   
   ErrorObj.LocalErrorMsg = "Runtime error."
   ErrorObj.SystemErrorMsg = Err.DESCRIPTION
   ErrorObj.RoutineName = RName
   ErrorObj.ModuleName = m_ModuleName
   ErrorObj.LocalErrorMsg = "ไม่สามารถลบข้อมูลได้เนื่องจากข้อมูลนี้ยังถูกอ้างถึงอยู่จากส่วนอื่น"
   ErrorObj.ShowErrorLog (LOG_TO_FILE)
   
   IsOK = False
   
   DeleteGroupType = False
End Function
Public Function DeleteSubGroupType(Ug As CSubGroupType, IsOK As Boolean, AutoCommit As Boolean, ErrorObj As clsErrorLog)
On Error GoTo ErrorHandler
Dim RName As String
Dim SQL As String
Dim HasBegin As Boolean

   RName = "DeleteSubGroupType"
   DeleteSubGroupType = False
   HasBegin = False
   
   If AutoCommit Then
      m_ADOConn.BeginTrans
   End If
   HasBegin = True
   
   Call Ug.DeleteData
   
   If AutoCommit Then
      Call m_ADOConn.CommitTrans
   End If
   HasBegin = False
   IsOK = True

   DeleteSubGroupType = True
   Exit Function
   
ErrorHandler:
   If HasBegin Then
      If AutoCommit Then
         m_ADOConn.RollbackTrans
      End If
   End If
   
   ErrorObj.LocalErrorMsg = "Runtime error."
   ErrorObj.SystemErrorMsg = Err.DESCRIPTION
   ErrorObj.RoutineName = RName
   ErrorObj.ModuleName = m_ModuleName
   ErrorObj.LocalErrorMsg = "ไม่สามารถลบข้อมูลได้เนื่องจากข้อมูลนี้ยังถูกอ้างถึงอยู่จากส่วนอื่น"
   ErrorObj.ShowErrorLog (LOG_TO_FILE)
   
   IsOK = False
   
   DeleteSubGroupType = False
End Function

Public Function QueryPromotionPayCustomer(Ua As CPromotionPayCustom, Rs As ADODB.Recordset, ItemCount As Long, IsOK As Boolean, ErrorObj As clsErrorLog) As Boolean
On Error GoTo ErrorHandler
Dim RName As String
Dim SQL1 As String
Dim SQL2 As String
Dim SelectStr As String
Dim FromStr As String
Dim WhereStr As String
Dim OrderStr As String
Dim OrderType As String
Dim NewStr As String
Dim SubLen As Long
Dim iCount As Long

   RName = "QueryPromotionPayCustomer"
   QueryPromotionPayCustomer = False

   IsOK = True
   Call Ua.QueryData(1, Rs, ItemCount)
   
   QueryPromotionPayCustomer = True
   Exit Function

ErrorHandler:
   ErrorObj.SystemErrorMsg = Err.DESCRIPTION
   ErrorObj.RoutineName = RName
   ErrorObj.ModuleName = m_ModuleName
   ErrorObj.ShowErrorLog (LOG_FILE_MSGBOX)

   QueryPromotionPayCustomer = False
End Function

Public Function QueryCostProduct(Ua As CCostProducts, Rs As ADODB.Recordset, ItemCount As Long, IsOK As Boolean, ErrorObj As clsErrorLog) As Boolean
On Error GoTo ErrorHandler
Dim RName As String
Dim SQL1 As String
Dim SQL2 As String
Dim SelectStr As String
Dim FromStr As String
Dim WhereStr As String
Dim OrderStr As String
Dim OrderType As String
Dim NewStr As String
Dim SubLen As Long
Dim iCount As Long

   RName = "QueryCostProduct"
   QueryCostProduct = False

   IsOK = True
   Call Ua.QueryData(1, Rs, ItemCount)
   
   QueryCostProduct = True
   Exit Function

ErrorHandler:
   ErrorObj.SystemErrorMsg = Err.DESCRIPTION
   ErrorObj.RoutineName = RName
   ErrorObj.ModuleName = m_ModuleName
   ErrorObj.ShowErrorLog (LOG_FILE_MSGBOX)

   QueryCostProduct = False
End Function
Public Function QueryPromotionYear(Ua As CPromotionYear, Rs As ADODB.Recordset, ItemCount As Long, IsOK As Boolean, ErrorObj As clsErrorLog) As Boolean
On Error GoTo ErrorHandler
Dim RName As String
Dim SQL1 As String
Dim SQL2 As String
Dim SelectStr As String
Dim FromStr As String
Dim WhereStr As String
Dim OrderStr As String
Dim OrderType As String
Dim NewStr As String
Dim SubLen As Long
Dim iCount As Long

   RName = "QueryPromotionYear"
   QueryPromotionYear = False

   IsOK = True
   Call Ua.QueryData(1, Rs, ItemCount)
   
   QueryPromotionYear = True
   Exit Function

ErrorHandler:
   ErrorObj.SystemErrorMsg = Err.DESCRIPTION
   ErrorObj.RoutineName = RName
   ErrorObj.ModuleName = m_ModuleName
   ErrorObj.ShowErrorLog (LOG_FILE_MSGBOX)

   QueryPromotionYear = False
End Function
Public Function QuerySupplierGroup(Ua As CSupplierGroup, Rs As ADODB.Recordset, ItemCount As Long, IsOK As Boolean, ErrorObj As clsErrorLog) As Boolean
On Error GoTo ErrorHandler
Dim RName As String
Dim SQL1 As String
Dim SQL2 As String
Dim SelectStr As String
Dim FromStr As String
Dim WhereStr As String
Dim OrderStr As String
Dim OrderType As String
Dim NewStr As String
Dim SubLen As Long
Dim iCount As Long

   RName = "QuerySupplierGroup"
   QuerySupplierGroup = False

   IsOK = True
   Call Ua.QueryData(Rs, ItemCount)
   
   QuerySupplierGroup = True
   Exit Function

ErrorHandler:
   ErrorObj.SystemErrorMsg = Err.DESCRIPTION
   ErrorObj.RoutineName = RName
   ErrorObj.ModuleName = m_ModuleName
   ErrorObj.ShowErrorLog (LOG_FILE_MSGBOX)

   QuerySupplierGroup = False
End Function

Public Function AddEditSupplierGroup(Ug As CSupplierGroup, IsOK As Boolean, AutoCommit As Boolean, ErrorObj As clsErrorLog)
On Error GoTo ErrorHandler
Dim RName As String
Dim HasBegin As Boolean

   RName = "AddEditSupplierGroup"
   AddEditSupplierGroup = False
   HasBegin = False

   If AutoCommit Then
      m_ADOConn.BeginTrans
   End If
   HasBegin = True

   Call Ug.AddEditData

   If AutoCommit Then
      Call m_ADOConn.CommitTrans
   End If
   HasBegin = False
   IsOK = True
   
   AddEditSupplierGroup = True
   Exit Function

ErrorHandler:
   If HasBegin Then
      If AutoCommit Then
         m_ADOConn.RollbackTrans
      End If
   End If

   ErrorObj.LocalErrorMsg = "Runtime error."
   ErrorObj.SystemErrorMsg = Err.DESCRIPTION
   ErrorObj.RoutineName = RName
   ErrorObj.ModuleName = m_ModuleName
   ErrorObj.ShowErrorLog (LOG_TO_FILE)

   IsOK = False
   AddEditSupplierGroup = False
End Function

Public Function AddEditComboSup(Ug As CComboSubGroup, IsOK As Boolean, AutoCommit As Boolean, ErrorObj As clsErrorLog)
On Error GoTo ErrorHandler
Dim RName As String
Dim HasBegin As Boolean

   RName = "AddEditComboSup"
   AddEditComboSup = False
   HasBegin = False

   If AutoCommit Then
      m_ADOConn.BeginTrans
   End If
   HasBegin = True

   Call Ug.AddEditData

   If AutoCommit Then
      Call m_ADOConn.CommitTrans
   End If
   HasBegin = False
   IsOK = True
   
   AddEditComboSup = True
   Exit Function

ErrorHandler:
   If HasBegin Then
      If AutoCommit Then
         m_ADOConn.RollbackTrans
      End If
   End If

   ErrorObj.LocalErrorMsg = "Runtime error."
   ErrorObj.SystemErrorMsg = Err.DESCRIPTION
   ErrorObj.RoutineName = RName
   ErrorObj.ModuleName = m_ModuleName
   ErrorObj.ShowErrorLog (LOG_TO_FILE)

   IsOK = False
   AddEditComboSup = False
End Function

Public Function DeleteComboSup(Ug As CComboSubGroup, IsOK As Boolean, AutoCommit As Boolean, ErrorObj As clsErrorLog)
On Error GoTo ErrorHandler
Dim RName As String
Dim SQL As String
Dim HasBegin As Boolean

   RName = "DeleteComboSup"
   DeleteComboSup = False
   HasBegin = False
   
   If AutoCommit Then
      m_ADOConn.BeginTrans
   End If
   HasBegin = True
   
   Call Ug.DeleteData
   
   If AutoCommit Then
      Call m_ADOConn.CommitTrans
   End If
   HasBegin = False
   IsOK = True

   DeleteComboSup = True
   Exit Function
   
ErrorHandler:
   If HasBegin Then
      If AutoCommit Then
         m_ADOConn.RollbackTrans
      End If
   End If
   
   ErrorObj.LocalErrorMsg = "Runtime error."
   ErrorObj.SystemErrorMsg = Err.DESCRIPTION
   ErrorObj.RoutineName = RName
   ErrorObj.ModuleName = m_ModuleName
   ErrorObj.LocalErrorMsg = "ไม่สามารถลบข้อมูลได้เนื่องจากข้อมูลนี้ยังถูกอ้างถึงอยู่จากส่วนอื่น"
   ErrorObj.ShowErrorLog (LOG_TO_FILE)
   
   IsOK = False
   DeleteComboSup = False
End Function

Public Function AddEditComboSupDe(Ug As CComboSubGroupDe, IsOK As Boolean, AutoCommit As Boolean, ErrorObj As clsErrorLog)
On Error GoTo ErrorHandler
Dim RName As String
Dim HasBegin As Boolean

   RName = "AddEditComboSupDe"
   AddEditComboSupDe = False
   HasBegin = False

   If AutoCommit Then
      m_ADOConn.BeginTrans
   End If
   HasBegin = True

   Call Ug.AddEditData

   If AutoCommit Then
      Call m_ADOConn.CommitTrans
   End If
   HasBegin = False
   IsOK = True
   
   AddEditComboSupDe = True
   Exit Function

ErrorHandler:
   If HasBegin Then
      If AutoCommit Then
         m_ADOConn.RollbackTrans
      End If
   End If

   ErrorObj.LocalErrorMsg = "Runtime error."
   ErrorObj.SystemErrorMsg = Err.DESCRIPTION
   ErrorObj.RoutineName = RName
   ErrorObj.ModuleName = m_ModuleName
   ErrorObj.ShowErrorLog (LOG_TO_FILE)

   IsOK = False
   AddEditComboSupDe = False
End Function
Public Function DeleteComboSupDe(Ug As CComboSubGroupDe, IsOK As Boolean, AutoCommit As Boolean, ErrorObj As clsErrorLog)
On Error GoTo ErrorHandler
Dim RName As String
Dim SQL As String
Dim HasBegin As Boolean

   RName = "DeleteComboSupDe"
   DeleteComboSupDe = False
   HasBegin = False
   
   If AutoCommit Then
      m_ADOConn.BeginTrans
   End If
   HasBegin = True
   
   Call Ug.DeleteData
   
   If AutoCommit Then
      Call m_ADOConn.CommitTrans
   End If
   HasBegin = False
   IsOK = True

   DeleteComboSupDe = True
   Exit Function
   
ErrorHandler:
   If HasBegin Then
      If AutoCommit Then
         m_ADOConn.RollbackTrans
      End If
   End If
   
   ErrorObj.LocalErrorMsg = "Runtime error."
   ErrorObj.SystemErrorMsg = Err.DESCRIPTION
   ErrorObj.RoutineName = RName
   ErrorObj.ModuleName = m_ModuleName
   ErrorObj.LocalErrorMsg = "ไม่สามารถลบข้อมูลได้เนื่องจากข้อมูลนี้ยังถูกอ้างถึงอยู่จากส่วนอื่น"
   ErrorObj.ShowErrorLog (LOG_TO_FILE)
   
   IsOK = False
   DeleteComboSupDe = False
End Function

Public Function DeletePromotionPayCustomer(Ug As CPromotionPayCustom, IsOK As Boolean, AutoCommit As Boolean, ErrorObj As clsErrorLog)
On Error GoTo ErrorHandler
Dim RName As String
Dim SQL As String
Dim HasBegin As Boolean

   RName = "DeletePromotionPayCustomer"
   DeletePromotionPayCustomer = False
   HasBegin = False
   
   If AutoCommit Then
      m_ADOConn.BeginTrans
   End If
   HasBegin = True
   
   Call Ug.DeleteData
   
   If AutoCommit Then
      Call m_ADOConn.CommitTrans
   End If
   HasBegin = False
   IsOK = True

   DeletePromotionPayCustomer = True
   Exit Function
   
ErrorHandler:
   If HasBegin Then
      If AutoCommit Then
         m_ADOConn.RollbackTrans
      End If
   End If
   
   ErrorObj.LocalErrorMsg = "Runtime error."
   ErrorObj.SystemErrorMsg = Err.DESCRIPTION
   ErrorObj.RoutineName = RName
   ErrorObj.ModuleName = m_ModuleName
   ErrorObj.LocalErrorMsg = "ไม่สามารถลบข้อมูลได้เนื่องจากข้อมูลนี้ยังถูกอ้างถึงอยู่จากส่วนอื่น"
   ErrorObj.ShowErrorLog (LOG_TO_FILE)
   
   IsOK = False
   
   DeletePromotionPayCustomer = False
End Function

Public Function DeletePromotionYear(Ug As CPromotionYear, IsOK As Boolean, AutoCommit As Boolean, ErrorObj As clsErrorLog)
On Error GoTo ErrorHandler
Dim RName As String
Dim SQL As String
Dim HasBegin As Boolean

   RName = "DeletePromotionYear"
   DeletePromotionYear = False
   HasBegin = False
   
   If AutoCommit Then
      m_ADOConn.BeginTrans
   End If
   HasBegin = True
   
   Call Ug.DeleteData
   
   If AutoCommit Then
      Call m_ADOConn.CommitTrans
   End If
   HasBegin = False
   IsOK = True

   DeletePromotionYear = True
   Exit Function
   
ErrorHandler:
   If HasBegin Then
      If AutoCommit Then
         m_ADOConn.RollbackTrans
      End If
   End If
   
   ErrorObj.LocalErrorMsg = "Runtime error."
   ErrorObj.SystemErrorMsg = Err.DESCRIPTION
   ErrorObj.RoutineName = RName
   ErrorObj.ModuleName = m_ModuleName
   ErrorObj.LocalErrorMsg = "ไม่สามารถลบข้อมูลได้เนื่องจากข้อมูลนี้ยังถูกอ้างถึงอยู่จากส่วนอื่น"
   ErrorObj.ShowErrorLog (LOG_TO_FILE)
   
   IsOK = False
   
   DeletePromotionYear = False
End Function

Public Function DeleteCoustProducts(Ug As CCostProducts, IsOK As Boolean, AutoCommit As Boolean, ErrorObj As clsErrorLog)
On Error GoTo ErrorHandler
Dim RName As String
Dim SQL As String
Dim HasBegin As Boolean

   RName = "DeleteCoustProducts"
   DeleteCoustProducts = False
   HasBegin = False
   
   If AutoCommit Then
      m_ADOConn.BeginTrans
   End If
   HasBegin = True
   
   Call Ug.DeleteData
   
   If AutoCommit Then
      Call m_ADOConn.CommitTrans
   End If
   HasBegin = False
   IsOK = True

   DeleteCoustProducts = True
   Exit Function
   
ErrorHandler:
   If HasBegin Then
      If AutoCommit Then
         m_ADOConn.RollbackTrans
      End If
   End If
   
   ErrorObj.LocalErrorMsg = "Runtime error."
   ErrorObj.SystemErrorMsg = Err.DESCRIPTION
   ErrorObj.RoutineName = RName
   ErrorObj.ModuleName = m_ModuleName
   ErrorObj.LocalErrorMsg = "ไม่สามารถลบข้อมูลได้เนื่องจากข้อมูลนี้ยังถูกอ้างถึงอยู่จากส่วนอื่น"
   ErrorObj.ShowErrorLog (LOG_TO_FILE)
   
   IsOK = False
   
   DeleteCoustProducts = False
End Function

Public Function DeleteSupplierGroup(Ug As CSupplierGroup, IsOK As Boolean, AutoCommit As Boolean, ErrorObj As clsErrorLog)
On Error GoTo ErrorHandler
Dim RName As String
Dim SQL As String
Dim HasBegin As Boolean

   RName = "DeleteSupplierGroup"
   DeleteSupplierGroup = False
   HasBegin = False
   
   If AutoCommit Then
      m_ADOConn.BeginTrans
   End If
   HasBegin = True
   
   Call Ug.DeleteData
   
   If AutoCommit Then
      Call m_ADOConn.CommitTrans
   End If
   HasBegin = False
   IsOK = True

   DeleteSupplierGroup = True
   Exit Function
   
ErrorHandler:
   If HasBegin Then
      If AutoCommit Then
         m_ADOConn.RollbackTrans
      End If
   End If
   
   ErrorObj.LocalErrorMsg = "Runtime error."
   ErrorObj.SystemErrorMsg = Err.DESCRIPTION
   ErrorObj.RoutineName = RName
   ErrorObj.ModuleName = m_ModuleName
   ErrorObj.LocalErrorMsg = "ไม่สามารถลบข้อมูลได้เนื่องจากข้อมูลนี้ยังถูกอ้างถึงอยู่จากส่วนอื่น"
   ErrorObj.ShowErrorLog (LOG_TO_FILE)
   
   IsOK = False
   
   DeleteSupplierGroup = False
End Function
Public Function QueryCustomer2(Ua As CCustomer, Rs As ADODB.Recordset, ItemCount As Long, IsOK As Boolean, ErrorObj As clsErrorLog) As Boolean
On Error GoTo ErrorHandler
Dim RName As String
Dim SQL1 As String
Dim SQL2 As String
Dim SelectStr As String
Dim FromStr As String
Dim WhereStr As String
Dim OrderStr As String
Dim OrderType As String
Dim NewStr As String
Dim SubLen As Long
Dim iCount As Long

   RName = "QueryCustomer2"
   QueryCustomer2 = False

   IsOK = True
   Call Ua.QueryData(Rs, ItemCount)
   
   QueryCustomer2 = True
   Exit Function

ErrorHandler:
   ErrorObj.SystemErrorMsg = Err.DESCRIPTION
   ErrorObj.RoutineName = RName
   ErrorObj.ModuleName = m_ModuleName
   ErrorObj.ShowErrorLog (LOG_FILE_MSGBOX)

   QueryCustomer2 = False
End Function
Public Function QueryARTran2(Ua As CARTrn, Rs As ADODB.Recordset, ItemCount As Long, IsOK As Boolean, ErrorObj As clsErrorLog) As Boolean
On Error GoTo ErrorHandler
Dim RName As String
Dim SQL1 As String
Dim SQL2 As String
Dim SelectStr As String
Dim FromStr As String
Dim WhereStr As String
Dim OrderStr As String
Dim OrderType As String
Dim NewStr As String
Dim SubLen As Long
Dim iCount As Long

   RName = "QueryARTran2"
   QueryARTran2 = False

   IsOK = True
   Call Ua.QueryData(1, Rs, ItemCount)

   QueryARTran2 = True
   Exit Function

ErrorHandler:
   ErrorObj.SystemErrorMsg = Err.DESCRIPTION
   ErrorObj.RoutineName = RName
   ErrorObj.ModuleName = m_ModuleName
   ErrorObj.ShowErrorLog (LOG_FILE_MSGBOX)

   QueryARTran2 = False
End Function
Public Function QueryARMas(Ua As CARMas, Rs As ADODB.Recordset, ItemCount As Long, IsOK As Boolean, ErrorObj As clsErrorLog) As Boolean
On Error GoTo ErrorHandler
Dim RName As String
Dim SQL1 As String
Dim SQL2 As String
Dim SelectStr As String
Dim FromStr As String
Dim WhereStr As String
Dim OrderStr As String
Dim OrderType As String
Dim NewStr As String
Dim SubLen As Long
Dim iCount As Long

   RName = "QueryARMas"
   QueryARMas = False

   IsOK = True
   Call Ua.QueryData(1, Rs, ItemCount)

   QueryARMas = True
   Exit Function

ErrorHandler:
   ErrorObj.SystemErrorMsg = Err.DESCRIPTION
   ErrorObj.RoutineName = RName
   ErrorObj.ModuleName = m_ModuleName
   ErrorObj.ShowErrorLog (LOG_FILE_MSGBOX)

   QueryARMas = False
End Function
Public Function QueryCusPigType(Ua As CCusPigType, Rs As ADODB.Recordset, ItemCount As Long, IsOK As Boolean, ErrorObj As clsErrorLog) As Boolean
On Error GoTo ErrorHandler
Dim RName As String
Dim SQL1 As String
Dim SQL2 As String
Dim SelectStr As String
Dim FromStr As String
Dim WhereStr As String
Dim OrderStr As String
Dim OrderType As String
Dim NewStr As String
Dim SubLen As Long
Dim iCount As Long

   RName = "QueryCusPigType"
   QueryCusPigType = False

   IsOK = True
   Call Ua.QueryData(1, Rs, ItemCount)

   QueryCusPigType = True
   Exit Function

ErrorHandler:
   ErrorObj.SystemErrorMsg = Err.DESCRIPTION
   ErrorObj.RoutineName = RName
   ErrorObj.ModuleName = m_ModuleName
   ErrorObj.ShowErrorLog (LOG_FILE_MSGBOX)

   QueryCusPigType = False
End Function
Public Function AddEditBank(Ug As CBank, IsOK As Boolean, AutoCommit As Boolean, ErrorObj As clsErrorLog)
On Error GoTo ErrorHandler
Dim RName As String
Dim HasBegin As Boolean

   RName = "AddEditBank"
   AddEditBank = False
   HasBegin = False

   If AutoCommit Then
      m_ADOConn.BeginTrans
   End If
   HasBegin = True

   Call Ug.AddEditData

   If AutoCommit Then
      Call m_ADOConn.CommitTrans
   End If
   HasBegin = False
   IsOK = True
   
   AddEditBank = True
   Exit Function

ErrorHandler:
   If HasBegin Then
      If AutoCommit Then
         m_ADOConn.RollbackTrans
      End If
   End If

   ErrorObj.LocalErrorMsg = "Runtime error."
   ErrorObj.SystemErrorMsg = Err.DESCRIPTION
   ErrorObj.RoutineName = RName
   ErrorObj.ModuleName = m_ModuleName
   ErrorObj.ShowErrorLog (LOG_TO_FILE)

   IsOK = False
   AddEditBank = False
End Function
Public Function AddEditAnalyzeCustomer(Ug As CAnalyzeCustomer, IsOK As Boolean, AutoCommit As Boolean, ErrorObj As clsErrorLog)
On Error GoTo ErrorHandler
Dim RName As String
Dim HasBegin As Boolean

   RName = "AddEditAnalyzeCustomer"
   AddEditAnalyzeCustomer = False
   HasBegin = False

   If AutoCommit Then
      m_ADOConn.BeginTrans
   End If
   HasBegin = True

   Call Ug.AddEditData

   If AutoCommit Then
      Call m_ADOConn.CommitTrans
   End If
   HasBegin = False
   IsOK = True
   
   AddEditAnalyzeCustomer = True
   Exit Function

ErrorHandler:
   If HasBegin Then
      If AutoCommit Then
         m_ADOConn.RollbackTrans
      End If
   End If

   ErrorObj.LocalErrorMsg = "Runtime error."
   ErrorObj.SystemErrorMsg = Err.DESCRIPTION
   ErrorObj.RoutineName = RName
   ErrorObj.ModuleName = m_ModuleName
   ErrorObj.ShowErrorLog (LOG_TO_FILE)

   IsOK = False
   AddEditAnalyzeCustomer = False
End Function
Public Function AddEditBankCustomer(Ug As CBankCustomer, IsOK As Boolean, AutoCommit As Boolean, ErrorObj As clsErrorLog)
On Error GoTo ErrorHandler
Dim RName As String
Dim HasBegin As Boolean

   RName = "AddEditBankCustomer"
   AddEditBankCustomer = False
   HasBegin = False

   If AutoCommit Then
      m_ADOConn.BeginTrans
   End If
   HasBegin = True

   Call Ug.AddEditData

   If AutoCommit Then
      Call m_ADOConn.CommitTrans
   End If
   HasBegin = False
   IsOK = True
   
   AddEditBankCustomer = True
   Exit Function

ErrorHandler:
   If HasBegin Then
      If AutoCommit Then
         m_ADOConn.RollbackTrans
      End If
   End If

   ErrorObj.LocalErrorMsg = "Runtime error."
   ErrorObj.SystemErrorMsg = Err.DESCRIPTION
   ErrorObj.RoutineName = RName
   ErrorObj.ModuleName = m_ModuleName
   ErrorObj.ShowErrorLog (LOG_TO_FILE)

   IsOK = False
   AddEditBankCustomer = False
End Function
Public Function AddEditBankCredit(Ug As CBankCredit, IsOK As Boolean, AutoCommit As Boolean, ErrorObj As clsErrorLog)
On Error GoTo ErrorHandler
Dim RName As String
Dim HasBegin As Boolean

   RName = "AddEditBankCredit"
   AddEditBankCredit = False
   HasBegin = False

   If AutoCommit Then
      m_ADOConn.BeginTrans
   End If
   HasBegin = True

   Call Ug.AddEditData

   If AutoCommit Then
      Call m_ADOConn.CommitTrans
   End If
   HasBegin = False
   IsOK = True
   
   AddEditBankCredit = True
   Exit Function

ErrorHandler:
   If HasBegin Then
      If AutoCommit Then
         m_ADOConn.RollbackTrans
      End If
   End If

   ErrorObj.LocalErrorMsg = "Runtime error."
   ErrorObj.SystemErrorMsg = Err.DESCRIPTION
   ErrorObj.RoutineName = RName
   ErrorObj.ModuleName = m_ModuleName
   ErrorObj.ShowErrorLog (LOG_TO_FILE)

   IsOK = False
   AddEditBankCredit = False
End Function
Public Function AddEditTicket(Ug As CTicket, IsOK As Boolean, AutoCommit As Boolean, ErrorObj As clsErrorLog)
On Error GoTo ErrorHandler
Dim RName As String
Dim HasBegin As Boolean

   RName = "AddEditTicket"
   AddEditTicket = False
   HasBegin = False

   If AutoCommit Then
      m_ADOConn.BeginTrans
   End If
   HasBegin = True

   Call Ug.AddEditData

   If AutoCommit Then
      Call m_ADOConn.CommitTrans
   End If
   HasBegin = False
   IsOK = True
   
   AddEditTicket = True
   Exit Function

ErrorHandler:
   If HasBegin Then
      If AutoCommit Then
         m_ADOConn.RollbackTrans
      End If
   End If

   ErrorObj.LocalErrorMsg = "Runtime error."
   ErrorObj.SystemErrorMsg = Err.DESCRIPTION
   ErrorObj.RoutineName = RName
   ErrorObj.ModuleName = m_ModuleName
   ErrorObj.ShowErrorLog (LOG_TO_FILE)

   IsOK = False
   AddEditTicket = False
End Function
Public Function DeleteBank(Ug As CBank, IsOK As Boolean, AutoCommit As Boolean, ErrorObj As clsErrorLog)
On Error GoTo ErrorHandler
Dim RName As String
Dim SQL As String
Dim HasBegin As Boolean

   RName = "DeleteBank"
   DeleteBank = False
   HasBegin = False
   
   If AutoCommit Then
      m_ADOConn.BeginTrans
   End If
   HasBegin = True
   
   Call Ug.DeleteData
   
   If AutoCommit Then
      Call m_ADOConn.CommitTrans
   End If
   HasBegin = False
   IsOK = True

   DeleteBank = True
   Exit Function
   
ErrorHandler:
   If HasBegin Then
      If AutoCommit Then
         m_ADOConn.RollbackTrans
      End If
   End If
   
   ErrorObj.LocalErrorMsg = "Runtime error."
   ErrorObj.SystemErrorMsg = Err.DESCRIPTION
   ErrorObj.RoutineName = RName
   ErrorObj.ModuleName = m_ModuleName
   ErrorObj.LocalErrorMsg = "ไม่สามารถลบข้อมูลได้เนื่องจากข้อมูลนี้ยังถูกอ้างถึงอยู่จากส่วนอื่น"
   ErrorObj.ShowErrorLog (LOG_TO_FILE)
   
   IsOK = False
   
   DeleteBank = False
End Function
Public Function DeleteAnalyzeCustomer(Ug As CAnalyzeCustomer, IsOK As Boolean, AutoCommit As Boolean, ErrorObj As clsErrorLog)
On Error GoTo ErrorHandler
Dim RName As String
Dim SQL As String
Dim HasBegin As Boolean

   RName = "DeleteAnalyzeCustomer"
   DeleteAnalyzeCustomer = False
   HasBegin = False
   
   If AutoCommit Then
      m_ADOConn.BeginTrans
   End If
   HasBegin = True
   
   Call Ug.DeleteData
   
   If AutoCommit Then
      Call m_ADOConn.CommitTrans
   End If
   HasBegin = False
   IsOK = True

   DeleteAnalyzeCustomer = True
   Exit Function
   
ErrorHandler:
   If HasBegin Then
      If AutoCommit Then
         m_ADOConn.RollbackTrans
      End If
   End If
   
   ErrorObj.LocalErrorMsg = "Runtime error."
   ErrorObj.SystemErrorMsg = Err.DESCRIPTION
   ErrorObj.RoutineName = RName
   ErrorObj.ModuleName = m_ModuleName
   ErrorObj.LocalErrorMsg = "ไม่สามารถลบข้อมูลได้เนื่องจากข้อมูลนี้ยังถูกอ้างถึงอยู่จากส่วนอื่น"
   ErrorObj.ShowErrorLog (LOG_TO_FILE)
   
   IsOK = False
   
   DeleteAnalyzeCustomer = False
End Function
Public Function DeleteBankCredit(Ug As CBankCredit, IsOK As Boolean, AutoCommit As Boolean, ErrorObj As clsErrorLog)
On Error GoTo ErrorHandler
Dim RName As String
Dim SQL As String
Dim HasBegin As Boolean

   RName = "DeleteBankCredit"
   DeleteBankCredit = False
   HasBegin = False
   
   If AutoCommit Then
      m_ADOConn.BeginTrans
   End If
   HasBegin = True
   
   Call Ug.DeleteData
   
   If AutoCommit Then
      Call m_ADOConn.CommitTrans
   End If
   HasBegin = False
   IsOK = True

   DeleteBankCredit = True
   Exit Function
   
ErrorHandler:
   If HasBegin Then
      If AutoCommit Then
         m_ADOConn.RollbackTrans
      End If
   End If
   
   ErrorObj.LocalErrorMsg = "Runtime error."
   ErrorObj.SystemErrorMsg = Err.DESCRIPTION
   ErrorObj.RoutineName = RName
   ErrorObj.ModuleName = m_ModuleName
   ErrorObj.LocalErrorMsg = "ไม่สามารถลบข้อมูลได้เนื่องจากข้อมูลนี้ยังถูกอ้างถึงอยู่จากส่วนอื่น"
   ErrorObj.ShowErrorLog (LOG_TO_FILE)
   
   IsOK = False
   
   DeleteBankCredit = False
End Function
Public Function DeleteBankCustomer(Ug As CBankCustomer, IsOK As Boolean, AutoCommit As Boolean, ErrorObj As clsErrorLog)
On Error GoTo ErrorHandler
Dim RName As String
Dim SQL As String
Dim HasBegin As Boolean

   RName = "DeleteBankCustomer"
   DeleteBankCustomer = False
   HasBegin = False
   
   If AutoCommit Then
      m_ADOConn.BeginTrans
   End If
   HasBegin = True
   
   Call Ug.DeleteData
   
   If AutoCommit Then
      Call m_ADOConn.CommitTrans
   End If
   HasBegin = False
   IsOK = True

   DeleteBankCustomer = True
   Exit Function
   
ErrorHandler:
   If HasBegin Then
      If AutoCommit Then
         m_ADOConn.RollbackTrans
      End If
   End If
   
   ErrorObj.LocalErrorMsg = "Runtime error."
   ErrorObj.SystemErrorMsg = Err.DESCRIPTION
   ErrorObj.RoutineName = RName
   ErrorObj.ModuleName = m_ModuleName
   ErrorObj.LocalErrorMsg = "ไม่สามารถลบข้อมูลได้เนื่องจากข้อมูลนี้ยังถูกอ้างถึงอยู่จากส่วนอื่น"
   ErrorObj.ShowErrorLog (LOG_TO_FILE)
   
   IsOK = False
   
   DeleteBankCustomer = False
End Function
Public Function DeleteTicket(Ug As CTicket, IsOK As Boolean, AutoCommit As Boolean, ErrorObj As clsErrorLog)
On Error GoTo ErrorHandler
Dim RName As String
Dim SQL As String
Dim HasBegin As Boolean

   RName = "DeleteTicket"
   DeleteTicket = False
   HasBegin = False
   
   If AutoCommit Then
      m_ADOConn.BeginTrans
   End If
   HasBegin = True
   
   Call Ug.DeleteData
   
   If AutoCommit Then
      Call m_ADOConn.CommitTrans
   End If
   HasBegin = False
   IsOK = True

   DeleteTicket = True
   Exit Function
   
ErrorHandler:
   If HasBegin Then
      If AutoCommit Then
         m_ADOConn.RollbackTrans
      End If
   End If
   
   ErrorObj.LocalErrorMsg = "Runtime error."
   ErrorObj.SystemErrorMsg = Err.DESCRIPTION
   ErrorObj.RoutineName = RName
   ErrorObj.ModuleName = m_ModuleName
   ErrorObj.LocalErrorMsg = "ไม่สามารถลบข้อมูลได้เนื่องจากข้อมูลนี้ยังถูกอ้างถึงอยู่จากส่วนอื่น"
   ErrorObj.ShowErrorLog (LOG_TO_FILE)
   
   IsOK = False
   DeleteTicket = False
End Function
Public Function QueryBank(Ua As CBank, Rs As ADODB.Recordset, ItemCount As Long, IsOK As Boolean, ErrorObj As clsErrorLog) As Boolean
On Error GoTo ErrorHandler
Dim RName As String
Dim SQL1 As String
Dim SQL2 As String
Dim SelectStr As String
Dim FromStr As String
Dim WhereStr As String
Dim OrderStr As String
Dim OrderType As String
Dim NewStr As String
Dim SubLen As Long
Dim iCount As Long

   RName = "QueryBank"
   QueryBank = False

   IsOK = True
   Call Ua.QueryData(Rs, ItemCount)
   
   QueryBank = True
   Exit Function

ErrorHandler:
   ErrorObj.SystemErrorMsg = Err.DESCRIPTION
   ErrorObj.RoutineName = RName
   ErrorObj.ModuleName = m_ModuleName
   ErrorObj.ShowErrorLog (LOG_FILE_MSGBOX)

   QueryBank = False
End Function
Public Function QueryAnalyzeCustomer(Ua As CAnalyzeCustomer, Rs As ADODB.Recordset, ItemCount As Long, IsOK As Boolean, ErrorObj As clsErrorLog) As Boolean
On Error GoTo ErrorHandler
Dim RName As String
Dim SQL1 As String
Dim SQL2 As String
Dim SelectStr As String
Dim FromStr As String
Dim WhereStr As String
Dim OrderStr As String
Dim OrderType As String
Dim NewStr As String
Dim SubLen As Long
Dim iCount As Long

   RName = "QueryAnalyzeCustomer"
   QueryAnalyzeCustomer = False

   IsOK = True
   Call Ua.QueryData(Rs, ItemCount)
   
   QueryAnalyzeCustomer = True
   Exit Function

ErrorHandler:
   ErrorObj.SystemErrorMsg = Err.DESCRIPTION
   ErrorObj.RoutineName = RName
   ErrorObj.ModuleName = m_ModuleName
   ErrorObj.ShowErrorLog (LOG_FILE_MSGBOX)

   QueryAnalyzeCustomer = False
End Function
Public Function QueryBankCredit(Ua As CBankCredit, Rs As ADODB.Recordset, ItemCount As Long, IsOK As Boolean, ErrorObj As clsErrorLog) As Boolean
On Error GoTo ErrorHandler
Dim RName As String
Dim SQL1 As String
Dim SQL2 As String
Dim SelectStr As String
Dim FromStr As String
Dim WhereStr As String
Dim OrderStr As String
Dim OrderType As String
Dim NewStr As String
Dim SubLen As Long
Dim iCount As Long

   RName = "QueryBankCredit"
   QueryBankCredit = False

   IsOK = True
   Call Ua.QueryData(1, Rs, ItemCount)
   
   QueryBankCredit = True
   Exit Function

ErrorHandler:
   ErrorObj.SystemErrorMsg = Err.DESCRIPTION
   ErrorObj.RoutineName = RName
   ErrorObj.ModuleName = m_ModuleName
   ErrorObj.ShowErrorLog (LOG_FILE_MSGBOX)

   QueryBankCredit = False
End Function
Public Function QueryBankCustomer(Ua As CBankCustomer, Rs As ADODB.Recordset, ItemCount As Long, IsOK As Boolean, ErrorObj As clsErrorLog) As Boolean
On Error GoTo ErrorHandler
Dim RName As String
Dim SQL1 As String
Dim SQL2 As String
Dim SelectStr As String
Dim FromStr As String
Dim WhereStr As String
Dim OrderStr As String
Dim OrderType As String
Dim NewStr As String
Dim SubLen As Long
Dim iCount As Long

   RName = "QueryBankCustomer"
   QueryBankCustomer = False

   IsOK = True
   Call Ua.QueryData(Rs, ItemCount)
   
   QueryBankCustomer = True
   Exit Function

ErrorHandler:
   ErrorObj.SystemErrorMsg = Err.DESCRIPTION
   ErrorObj.RoutineName = RName
   ErrorObj.ModuleName = m_ModuleName
   ErrorObj.ShowErrorLog (LOG_FILE_MSGBOX)

   QueryBankCustomer = False
End Function
Public Function QueryTicket(Ua As CTicket, Rs As ADODB.Recordset, ItemCount As Long, IsOK As Boolean, ErrorObj As clsErrorLog) As Boolean
On Error GoTo ErrorHandler
Dim RName As String
Dim SQL1 As String
Dim SQL2 As String
Dim SelectStr As String
Dim FromStr As String
Dim WhereStr As String
Dim OrderStr As String
Dim OrderType As String
Dim NewStr As String
Dim SubLen As Long
Dim iCount As Long

   RName = "QueryTicket"
   QueryTicket = False

   IsOK = True
   Call Ua.QueryData(1, Rs, ItemCount)
   
   QueryTicket = True
   Exit Function

ErrorHandler:
   ErrorObj.SystemErrorMsg = Err.DESCRIPTION
   ErrorObj.RoutineName = RName
   ErrorObj.ModuleName = m_ModuleName
   ErrorObj.ShowErrorLog (LOG_FILE_MSGBOX)

   QueryTicket = False
End Function
Public Function AddEditCustomer(Ug As CCustomer, IsOK As Boolean, AutoCommit As Boolean, ErrorObj As clsErrorLog)
On Error GoTo ErrorHandler
Dim RName As String
Dim HasBegin As Boolean

   RName = "AddEditCustomer"
   AddEditCustomer = False
   HasBegin = False

   If AutoCommit Then
      m_ADOConn.BeginTrans
   End If
   HasBegin = True

   Call Ug.AddEditData

   If AutoCommit Then
      Call m_ADOConn.CommitTrans
   End If
   HasBegin = False
   IsOK = True
   
   AddEditCustomer = True
   Exit Function

ErrorHandler:
   If HasBegin Then
      If AutoCommit Then
         m_ADOConn.RollbackTrans
      End If
   End If

   ErrorObj.LocalErrorMsg = "Runtime error."
   ErrorObj.SystemErrorMsg = Err.DESCRIPTION
   ErrorObj.RoutineName = RName
   ErrorObj.ModuleName = m_ModuleName
   ErrorObj.ShowErrorLog (LOG_TO_FILE)

   IsOK = False
   AddEditCustomer = False
End Function
Public Function AddEditCusPigType(Ug As CCusPigType, IsOK As Boolean, AutoCommit As Boolean, ErrorObj As clsErrorLog)
On Error GoTo ErrorHandler
Dim RName As String
Dim HasBegin As Boolean

   RName = "AddEditCusPigType"
   AddEditCusPigType = False
   HasBegin = False

   If AutoCommit Then
      m_ADOConn.BeginTrans
   End If
   HasBegin = True

   Call Ug.AddEditData

   If AutoCommit Then
      Call m_ADOConn.CommitTrans
   End If
   HasBegin = False
   IsOK = True
   
   AddEditCusPigType = True
   Exit Function

ErrorHandler:
   If HasBegin Then
      If AutoCommit Then
         m_ADOConn.RollbackTrans
      End If
   End If

   ErrorObj.LocalErrorMsg = "Runtime error."
   ErrorObj.SystemErrorMsg = Err.DESCRIPTION
   ErrorObj.RoutineName = RName
   ErrorObj.ModuleName = m_ModuleName
   ErrorObj.ShowErrorLog (LOG_TO_FILE)

   IsOK = False
   AddEditCusPigType = False
End Function
Public Function DeleteCustomer(Ug As CCustomer, IsOK As Boolean, AutoCommit As Boolean, ErrorObj As clsErrorLog)
On Error GoTo ErrorHandler
Dim RName As String
Dim SQL As String
Dim HasBegin As Boolean

   RName = "DeleteCustomer"
   DeleteCustomer = False
   HasBegin = False
   
   If AutoCommit Then
      m_ADOConn.BeginTrans
   End If
   HasBegin = True
   
   Call Ug.DeleteData
   
   If AutoCommit Then
      Call m_ADOConn.CommitTrans
   End If
   HasBegin = False
   IsOK = True

   DeleteCustomer = True
   Exit Function
   
ErrorHandler:
   If HasBegin Then
      If AutoCommit Then
         m_ADOConn.RollbackTrans
      End If
   End If
   
   ErrorObj.LocalErrorMsg = "Runtime error."
   ErrorObj.SystemErrorMsg = Err.DESCRIPTION
   ErrorObj.RoutineName = RName
   ErrorObj.ModuleName = m_ModuleName
   ErrorObj.LocalErrorMsg = "ไม่สามารถลบข้อมูลได้เนื่องจากข้อมูลนี้ยังถูกอ้างถึงอยู่จากส่วนอื่น"
   ErrorObj.ShowErrorLog (LOG_TO_FILE)
   
   IsOK = False
   
   DeleteCustomer = False
End Function
Public Function DeleteCusPigType(Ug As CARMas, IsOK As Boolean, AutoCommit As Boolean, ErrorObj As clsErrorLog)
On Error GoTo ErrorHandler
Dim RName As String
Dim SQL As String
Dim HasBegin As Boolean
Dim tempUg As CCusPigType
Set tempUg = New CCusPigType

   RName = "DeleteCusPigType"
   DeleteCusPigType = False
   HasBegin = False
   
   If AutoCommit Then
      m_ADOConn.BeginTrans
   End If
   HasBegin = True
   tempUg.CUS_PIG_TYPE_CODE = Ug.CUSCOD
   Call tempUg.DeleteData
   
   If AutoCommit Then
      Call m_ADOConn.CommitTrans
   End If
   HasBegin = False
   IsOK = True

   DeleteCusPigType = True
   Exit Function
   
ErrorHandler:
   If HasBegin Then
      If AutoCommit Then
         m_ADOConn.RollbackTrans
      End If
   End If
   
   ErrorObj.LocalErrorMsg = "Runtime error."
   ErrorObj.SystemErrorMsg = Err.DESCRIPTION
   ErrorObj.RoutineName = RName
   ErrorObj.ModuleName = m_ModuleName
   ErrorObj.LocalErrorMsg = "ไม่สามารถลบข้อมูลได้เนื่องจากข้อมูลนี้ยังถูกอ้างถึงอยู่จากส่วนอื่น"
   ErrorObj.ShowErrorLog (LOG_TO_FILE)
   
   IsOK = False
   
   DeleteCusPigType = False
End Function
Public Function QueryCustomerType(Ua As CCustomerType, Rs As ADODB.Recordset, ItemCount As Long, IsOK As Boolean, ErrorObj As clsErrorLog) As Boolean
On Error GoTo ErrorHandler
Dim RName As String
Dim SQL1 As String
Dim SQL2 As String
Dim SelectStr As String
Dim FromStr As String
Dim WhereStr As String
Dim OrderStr As String
Dim OrderType As String
Dim NewStr As String
Dim SubLen As Long
Dim iCount As Long

   RName = "QueryCustomerType"
   QueryCustomerType = False

   IsOK = True
   Call Ua.QueryData(Rs, ItemCount)
   
   QueryCustomerType = True
   Exit Function

ErrorHandler:
   ErrorObj.SystemErrorMsg = Err.DESCRIPTION
   ErrorObj.RoutineName = RName
   ErrorObj.ModuleName = m_ModuleName
   ErrorObj.ShowErrorLog (LOG_FILE_MSGBOX)

   QueryCustomerType = False
End Function
Public Function AddEditCustomerType(Ug As CCustomerType, IsOK As Boolean, AutoCommit As Boolean, ErrorObj As clsErrorLog)
On Error GoTo ErrorHandler
Dim RName As String
Dim HasBegin As Boolean

   RName = "AddEditCustomerType"
   AddEditCustomerType = False
   HasBegin = False

   If AutoCommit Then
      m_ADOConn.BeginTrans
   End If
   HasBegin = True

   Call Ug.AddEditData

   If AutoCommit Then
      Call m_ADOConn.CommitTrans
   End If
   HasBegin = False
   IsOK = True
   
   AddEditCustomerType = True
   Exit Function

ErrorHandler:
   If HasBegin Then
      If AutoCommit Then
         m_ADOConn.RollbackTrans
      End If
   End If

   ErrorObj.LocalErrorMsg = "Runtime error."
   ErrorObj.SystemErrorMsg = Err.DESCRIPTION
   ErrorObj.RoutineName = RName
   ErrorObj.ModuleName = m_ModuleName
   ErrorObj.ShowErrorLog (LOG_TO_FILE)

   IsOK = False
   AddEditCustomerType = False
End Function
Public Function DeleteCustomerType(Ug As CCustomerType, IsOK As Boolean, AutoCommit As Boolean, ErrorObj As clsErrorLog)
On Error GoTo ErrorHandler
Dim RName As String
Dim SQL As String
Dim HasBegin As Boolean

   RName = "DeleteCustomerType"
   DeleteCustomerType = False
   HasBegin = False
   
   If AutoCommit Then
      m_ADOConn.BeginTrans
   End If
   HasBegin = True
   
   Call Ug.DeleteData
   
   If AutoCommit Then
      Call m_ADOConn.CommitTrans
   End If
   HasBegin = False
   IsOK = True

   DeleteCustomerType = True
   Exit Function
   
ErrorHandler:
   If HasBegin Then
      If AutoCommit Then
         m_ADOConn.RollbackTrans
      End If
   End If
   
   ErrorObj.LocalErrorMsg = "Runtime error."
   ErrorObj.SystemErrorMsg = Err.DESCRIPTION
   ErrorObj.RoutineName = RName
   ErrorObj.ModuleName = m_ModuleName
   ErrorObj.LocalErrorMsg = "ไม่สามารถลบข้อมูลได้เนื่องจากข้อมูลนี้ยังถูกอ้างถึงอยู่จากส่วนอื่น"
   ErrorObj.ShowErrorLog (LOG_TO_FILE)
   
   IsOK = False
   
   DeleteCustomerType = False
End Function


Public Function QueryCheckCancel(Ua As CCheckCancel, Rs As ADODB.Recordset, ItemCount As Long, IsOK As Boolean, ErrorObj As clsErrorLog) As Boolean
On Error GoTo ErrorHandler
Dim RName As String
Dim SQL1 As String
Dim SQL2 As String
Dim SelectStr As String
Dim FromStr As String
Dim WhereStr As String
Dim OrderStr As String
Dim OrderType As String
Dim NewStr As String
Dim SubLen As Long
Dim iCount As Long

   RName = "QueryCheckCancel"
   QueryCheckCancel = False

   IsOK = True
   Call Ua.QueryData(Rs, ItemCount)
   
   QueryCheckCancel = True
   Exit Function

ErrorHandler:
   ErrorObj.SystemErrorMsg = Err.DESCRIPTION
   ErrorObj.RoutineName = RName
   ErrorObj.ModuleName = m_ModuleName
   ErrorObj.ShowErrorLog (LOG_FILE_MSGBOX)

   QueryCheckCancel = False
End Function


Public Function AddEditCheckCancel(Ug As CCheckCancel, IsOK As Boolean, AutoCommit As Boolean, ErrorObj As clsErrorLog)
On Error GoTo ErrorHandler
Dim RName As String
Dim HasBegin As Boolean

   RName = "AddEditCheckCancel"
   AddEditCheckCancel = False
   HasBegin = False

   If AutoCommit Then
      m_ADOConn.BeginTrans
   End If
   HasBegin = True

   Call Ug.AddEditData

   If AutoCommit Then
      Call m_ADOConn.CommitTrans
   End If
   HasBegin = False
   IsOK = True

   AddEditCheckCancel = True
   Exit Function

ErrorHandler:
   If HasBegin Then
      If AutoCommit Then
         m_ADOConn.RollbackTrans
      End If
   End If

   ErrorObj.LocalErrorMsg = "Runtime error."
   ErrorObj.SystemErrorMsg = Err.DESCRIPTION
   ErrorObj.RoutineName = RName
   ErrorObj.ModuleName = m_ModuleName
   ErrorObj.ShowErrorLog (LOG_TO_FILE)

   IsOK = False
   AddEditCheckCancel = False
End Function
Public Function DeleteCheckCancel(Ug As CCheckCancel, IsOK As Boolean, AutoCommit As Boolean, ErrorObj As clsErrorLog)
On Error GoTo ErrorHandler
Dim RName As String
Dim SQL As String
Dim HasBegin As Boolean

   RName = "DeleteCheckCancel"
   DeleteCheckCancel = False
   HasBegin = False
   
   If AutoCommit Then
      m_ADOConn.BeginTrans
   End If
   HasBegin = True
   
   Call Ug.DeleteData
   
   If AutoCommit Then
      Call m_ADOConn.CommitTrans
   End If
   HasBegin = False
   IsOK = True

   DeleteCheckCancel = True
   Exit Function
   
ErrorHandler:
   If HasBegin Then
      If AutoCommit Then
         m_ADOConn.RollbackTrans
      End If
   End If
   
   ErrorObj.LocalErrorMsg = "Runtime error."
   ErrorObj.SystemErrorMsg = Err.DESCRIPTION
   ErrorObj.RoutineName = RName
   ErrorObj.ModuleName = m_ModuleName
   ErrorObj.LocalErrorMsg = "ไม่สามารถลบข้อมูลได้เนื่องจากข้อมูลนี้ยังถูกอ้างถึงอยู่จากส่วนอื่น"
   ErrorObj.ShowErrorLog (LOG_TO_FILE)
   
   IsOK = False
   
   DeleteCheckCancel = False
End Function

Public Function QueryDocumentCancel(Ua As CDocumentCancel, Rs As ADODB.Recordset, ItemCount As Long, IsOK As Boolean, ErrorObj As clsErrorLog) As Boolean
On Error GoTo ErrorHandler
Dim RName As String
Dim SQL1 As String
Dim SQL2 As String
Dim SelectStr As String
Dim FromStr As String
Dim WhereStr As String
Dim OrderStr As String
Dim OrderType As String
Dim NewStr As String
Dim SubLen As Long
Dim iCount As Long

   RName = "QueryDocumentCancel"
   QueryDocumentCancel = False

   IsOK = True
   Call Ua.QueryData(1, Rs, ItemCount)
   
   QueryDocumentCancel = True
   Exit Function

ErrorHandler:
   ErrorObj.SystemErrorMsg = Err.DESCRIPTION
   ErrorObj.RoutineName = RName
   ErrorObj.ModuleName = m_ModuleName
   ErrorObj.ShowErrorLog (LOG_FILE_MSGBOX)

   QueryDocumentCancel = False
End Function
Public Function AddEditDocumentCancel(Ug As CDocumentCancel, IsOK As Boolean, AutoCommit As Boolean, ErrorObj As clsErrorLog)
On Error GoTo ErrorHandler
Dim RName As String
Dim HasBegin As Boolean

   RName = "AddEditDocumentCancel"
   AddEditDocumentCancel = False
   HasBegin = False

   If AutoCommit Then
      m_ADOConn.BeginTrans
   End If
   HasBegin = True

   Call Ug.AddEditData

   If AutoCommit Then
      Call m_ADOConn.CommitTrans
   End If
   HasBegin = False
   IsOK = True

   AddEditDocumentCancel = True
   Exit Function

ErrorHandler:
   If HasBegin Then
      If AutoCommit Then
         m_ADOConn.RollbackTrans
      End If
   End If

   ErrorObj.LocalErrorMsg = "Runtime error."
   ErrorObj.SystemErrorMsg = Err.DESCRIPTION
   ErrorObj.RoutineName = RName
   ErrorObj.ModuleName = m_ModuleName
   ErrorObj.ShowErrorLog (LOG_TO_FILE)

   IsOK = False
   AddEditDocumentCancel = False
End Function
Public Function DeleteDocumentCancel(Ug As CDocumentCancel, IsOK As Boolean, AutoCommit As Boolean, ErrorObj As clsErrorLog)
On Error GoTo ErrorHandler
Dim RName As String
Dim SQL As String
Dim HasBegin As Boolean

   RName = "DeleteDocumentCancel"
   DeleteDocumentCancel = False
   HasBegin = False
   
   If AutoCommit Then
      m_ADOConn.BeginTrans
   End If
   HasBegin = True
   
   Call Ug.DeleteData
   
   If AutoCommit Then
      Call m_ADOConn.CommitTrans
   End If
   HasBegin = False
   IsOK = True

   DeleteDocumentCancel = True
   Exit Function
   
ErrorHandler:
   If HasBegin Then
      If AutoCommit Then
         m_ADOConn.RollbackTrans
      End If
   End If
   
   ErrorObj.LocalErrorMsg = "Runtime error."
   ErrorObj.SystemErrorMsg = Err.DESCRIPTION
   ErrorObj.RoutineName = RName
   ErrorObj.ModuleName = m_ModuleName
   ErrorObj.LocalErrorMsg = "ไม่สามารถลบข้อมูลได้เนื่องจากข้อมูลนี้ยังถูกอ้างถึงอยู่จากส่วนอื่น"
   ErrorObj.ShowErrorLog (LOG_TO_FILE)
   
   IsOK = False
   
   DeleteDocumentCancel = False
End Function

Public Function QueryDataType(Ua As CDataType, Rs As ADODB.Recordset, ItemCount As Long, IsOK As Boolean, ErrorObj As clsErrorLog) As Boolean
On Error GoTo ErrorHandler
Dim RName As String
Dim SQL1 As String
Dim SQL2 As String
Dim SelectStr As String
Dim FromStr As String
Dim WhereStr As String
Dim OrderStr As String
Dim OrderType As String
Dim NewStr As String
Dim SubLen As Long
Dim iCount As Long

   RName = "QueryDataType"
   QueryDataType = False

   IsOK = True
   Call Ua.QueryData(Rs, ItemCount)
   
   QueryDataType = True
   Exit Function

ErrorHandler:
   ErrorObj.SystemErrorMsg = Err.DESCRIPTION
   ErrorObj.RoutineName = RName
   ErrorObj.ModuleName = m_ModuleName
   ErrorObj.ShowErrorLog (LOG_FILE_MSGBOX)

   QueryDataType = False
End Function
Public Function AddEditDataType(Ug As CDataType, IsOK As Boolean, AutoCommit As Boolean, ErrorObj As clsErrorLog)
On Error GoTo ErrorHandler
Dim RName As String
Dim HasBegin As Boolean

   RName = "AddEditDataType"
   AddEditDataType = False
   HasBegin = False

   If AutoCommit Then
      m_ADOConn.BeginTrans
   End If
   HasBegin = True

   Call Ug.AddEditData

   If AutoCommit Then
      Call m_ADOConn.CommitTrans
   End If
   HasBegin = False
   IsOK = True
   
   AddEditDataType = True
   Exit Function

ErrorHandler:
   If HasBegin Then
      If AutoCommit Then
         m_ADOConn.RollbackTrans
      End If
   End If

   ErrorObj.LocalErrorMsg = "Runtime error."
   ErrorObj.SystemErrorMsg = Err.DESCRIPTION
   ErrorObj.RoutineName = RName
   ErrorObj.ModuleName = m_ModuleName
   ErrorObj.ShowErrorLog (LOG_TO_FILE)

   IsOK = False
   AddEditDataType = False
End Function
Public Function DeleteDataType(Ug As CDataType, IsOK As Boolean, AutoCommit As Boolean, ErrorObj As clsErrorLog)
On Error GoTo ErrorHandler
Dim RName As String
Dim SQL As String
Dim HasBegin As Boolean

   RName = "DeleteDataType"
   DeleteDataType = False
   HasBegin = False
   
   If AutoCommit Then
      m_ADOConn.BeginTrans
   End If
   HasBegin = True
   
   Call Ug.DeleteData
   
   If AutoCommit Then
      Call m_ADOConn.CommitTrans
   End If
   HasBegin = False
   IsOK = True

   DeleteDataType = True
   Exit Function
   
ErrorHandler:
   If HasBegin Then
      If AutoCommit Then
         m_ADOConn.RollbackTrans
      End If
   End If
   
   ErrorObj.LocalErrorMsg = "Runtime error."
   ErrorObj.SystemErrorMsg = Err.DESCRIPTION
   ErrorObj.RoutineName = RName
   ErrorObj.ModuleName = m_ModuleName
   ErrorObj.LocalErrorMsg = "ไม่สามารถลบข้อมูลได้เนื่องจากข้อมูลนี้ยังถูกอ้างถึงอยู่จากส่วนอื่น"
   ErrorObj.ShowErrorLog (LOG_TO_FILE)
   
   IsOK = False
   
   DeleteDataType = False
End Function
'Commission

Public Function QueryXlsSetting(Ua As CXlsEstimateSetting, Rs As ADODB.Recordset, ItemCount As Long, IsOK As Boolean, ErrorObj As clsErrorLog) As Boolean
On Error GoTo ErrorHandler
Dim RName As String
Dim SQL1 As String
Dim SQL2 As String
Dim SelectStr As String
Dim FromStr As String
Dim WhereStr As String
Dim OrderStr As String
Dim OrderType As String
Dim NewStr As String
Dim SubLen As Long
Dim iCount As Long

   RName = "QueryXlsSetting"
   QueryXlsSetting = False

   IsOK = True
   Call Ua.QueryData(Rs, ItemCount)
   
   QueryXlsSetting = True
   Exit Function

ErrorHandler:
   ErrorObj.SystemErrorMsg = Err.DESCRIPTION
   ErrorObj.RoutineName = RName
   ErrorObj.ModuleName = m_ModuleName
   ErrorObj.ShowErrorLog (LOG_FILE_MSGBOX)

   QueryXlsSetting = False
End Function

Public Function AddEditXlsSetting(Ug As CXlsEstimateSetting, IsOK As Boolean, AutoCommit As Boolean, ErrorObj As clsErrorLog)
On Error GoTo ErrorHandler
Dim RName As String
Dim HasBegin As Boolean

   RName = "AddEditXlsSetting"
   AddEditXlsSetting = False
   HasBegin = False

   If AutoCommit Then
      m_ADOConn.BeginTrans
   End If
   HasBegin = True

   Call Ug.AddEditData

   If AutoCommit Then
      Call m_ADOConn.CommitTrans
   End If
   HasBegin = False
   IsOK = True
   
   AddEditXlsSetting = True
   Exit Function

ErrorHandler:
   If HasBegin Then
      If AutoCommit Then
         m_ADOConn.RollbackTrans
      End If
   End If

   ErrorObj.LocalErrorMsg = "Runtime error."
   ErrorObj.SystemErrorMsg = Err.DESCRIPTION
   ErrorObj.RoutineName = RName
   ErrorObj.ModuleName = m_ModuleName
   ErrorObj.ShowErrorLog (LOG_TO_FILE)

   IsOK = False
   AddEditXlsSetting = False
End Function

Public Function DeleteXlsSetting(Ug As CXlsEstimateSetting, IsOK As Boolean, AutoCommit As Boolean, ErrorObj As clsErrorLog)
On Error GoTo ErrorHandler
Dim RName As String
Dim SQL As String
Dim HasBegin As Boolean

   RName = "DeleteXlsSetting"
   DeleteXlsSetting = False
   HasBegin = False
   
   If AutoCommit Then
      m_ADOConn.BeginTrans
   End If
   HasBegin = True
   
   Call Ug.DeleteData
   
   If AutoCommit Then
      Call m_ADOConn.CommitTrans
   End If
   HasBegin = False
   IsOK = True

   DeleteXlsSetting = True
   Exit Function
   
ErrorHandler:
   If HasBegin Then
      If AutoCommit Then
         m_ADOConn.RollbackTrans
      End If
   End If
   
   ErrorObj.LocalErrorMsg = "Runtime error."
   ErrorObj.SystemErrorMsg = Err.DESCRIPTION
   ErrorObj.RoutineName = RName
   ErrorObj.ModuleName = m_ModuleName
   ErrorObj.LocalErrorMsg = "ไม่สามารถลบข้อมูลได้เนื่องจากข้อมูลนี้ยังถูกอ้างถึงอยู่จากส่วนอื่น"
   ErrorObj.ShowErrorLog (LOG_TO_FILE)
   
   IsOK = False
   
   DeleteXlsSetting = False
End Function
'Commission

Public Function QueryPromoteType1(Ua As CCommissPromote, Rs As ADODB.Recordset, ItemCount As Long, IsOK As Boolean, ErrorObj As clsErrorLog) As Boolean
On Error GoTo ErrorHandler
Dim RName As String
Dim SQL1 As String
Dim SQL2 As String
Dim SelectStr As String
Dim FromStr As String
Dim WhereStr As String
Dim OrderStr As String
Dim OrderType As String
Dim NewStr As String
Dim SubLen As Long
Dim iCount As Long

   RName = "QueryPromoteType1"
   QueryPromoteType1 = False

   IsOK = True
   Call Ua.QueryData(2, Rs, ItemCount)
   
   QueryPromoteType1 = True
   Exit Function

ErrorHandler:
   ErrorObj.SystemErrorMsg = Err.DESCRIPTION
   ErrorObj.RoutineName = RName
   ErrorObj.ModuleName = m_ModuleName
   ErrorObj.ShowErrorLog (LOG_FILE_MSGBOX)

   QueryPromoteType1 = False
End Function

Public Function AddEditPromoteCommiss(Ug As CCommissPromote, IsOK As Boolean, AutoCommit As Boolean, ErrorObj As clsErrorLog)
On Error GoTo ErrorHandler
Dim RName As String
Dim HasBegin As Boolean

   RName = "AddEditPromoteCommiss"
   AddEditPromoteCommiss = False
   HasBegin = False

   If AutoCommit Then
      m_ADOConn.BeginTrans
   End If
   HasBegin = True

   Call Ug.AddEditData

   If AutoCommit Then
      Call m_ADOConn.CommitTrans
   End If
   HasBegin = False
   IsOK = True

   AddEditPromoteCommiss = True
   Exit Function

ErrorHandler:
   If HasBegin Then
      If AutoCommit Then
         m_ADOConn.RollbackTrans
      End If
   End If

   ErrorObj.LocalErrorMsg = "Runtime error."
   ErrorObj.SystemErrorMsg = Err.DESCRIPTION
   ErrorObj.RoutineName = RName
   ErrorObj.ModuleName = m_ModuleName
   ErrorObj.ShowErrorLog (LOG_TO_FILE)

   IsOK = False
   AddEditPromoteCommiss = False
End Function

Public Function AddEditPromotionPayCustomer(Ug As CPromotionPayCustom, IsOK As Boolean, AutoCommit As Boolean, ErrorObj As clsErrorLog)
On Error GoTo ErrorHandler
Dim RName As String
Dim HasBegin As Boolean

   RName = "AddEditPromotionPayCustomer"
   AddEditPromotionPayCustomer = False
   HasBegin = False

   If AutoCommit Then
      m_ADOConn.BeginTrans
   End If
   HasBegin = True

   Call Ug.AddEditData

   If AutoCommit Then
      Call m_ADOConn.CommitTrans
   End If
   HasBegin = False
   IsOK = True

   AddEditPromotionPayCustomer = True
   Exit Function

ErrorHandler:
   If HasBegin Then
      If AutoCommit Then
         m_ADOConn.RollbackTrans
      End If
   End If

   ErrorObj.LocalErrorMsg = "Runtime error."
   ErrorObj.SystemErrorMsg = Err.DESCRIPTION
   ErrorObj.RoutineName = RName
   ErrorObj.ModuleName = m_ModuleName
   ErrorObj.ShowErrorLog (LOG_TO_FILE)

   IsOK = False
   AddEditPromotionPayCustomer = False
End Function

Public Function AddEditPromotionYear(Ug As CPromotionYear, IsOK As Boolean, AutoCommit As Boolean, ErrorObj As clsErrorLog)
On Error GoTo ErrorHandler
Dim RName As String
Dim HasBegin As Boolean

   RName = "AddEditPromotionYear"
   AddEditPromotionYear = False
   HasBegin = False

   If AutoCommit Then
      m_ADOConn.BeginTrans
   End If
   HasBegin = True

   Call Ug.AddEditData

   If AutoCommit Then
      Call m_ADOConn.CommitTrans
   End If
   HasBegin = False
   IsOK = True

   AddEditPromotionYear = True
   Exit Function

ErrorHandler:
   If HasBegin Then
      If AutoCommit Then
         m_ADOConn.RollbackTrans
      End If
   End If

   ErrorObj.LocalErrorMsg = "Runtime error."
   ErrorObj.SystemErrorMsg = Err.DESCRIPTION
   ErrorObj.RoutineName = RName
   ErrorObj.ModuleName = m_ModuleName
   ErrorObj.ShowErrorLog (LOG_TO_FILE)

   IsOK = False
   AddEditPromotionYear = False
End Function
Public Function AddEditCostProducts(Ug As CCostProducts, IsOK As Boolean, AutoCommit As Boolean, ErrorObj As clsErrorLog)
On Error GoTo ErrorHandler
Dim RName As String
Dim HasBegin As Boolean

   RName = "AddEditCostProducts"
   AddEditCostProducts = False
   HasBegin = False

   If AutoCommit Then
      m_ADOConn.BeginTrans
   End If
   HasBegin = True

   Call Ug.AddEditData

   If AutoCommit Then
      Call m_ADOConn.CommitTrans
   End If
   HasBegin = False
   IsOK = True

   AddEditCostProducts = True
   Exit Function

ErrorHandler:
   If HasBegin Then
      If AutoCommit Then
         m_ADOConn.RollbackTrans
      End If
   End If

   ErrorObj.LocalErrorMsg = "Runtime error."
   ErrorObj.SystemErrorMsg = Err.DESCRIPTION
   ErrorObj.RoutineName = RName
   ErrorObj.ModuleName = m_ModuleName
   ErrorObj.ShowErrorLog (LOG_TO_FILE)

   IsOK = False
   AddEditCostProducts = False
End Function

Public Function AddEditConditionCommiss(Ug As CConditionCommission, IsOK As Boolean, AutoCommit As Boolean, ErrorObj As clsErrorLog)
On Error GoTo ErrorHandler
Dim RName As String
Dim HasBegin As Boolean

   RName = "AddEditConditionCommiss"
   AddEditConditionCommiss = False
   HasBegin = False

   If AutoCommit Then
      m_ADOConn.BeginTrans
   End If
   HasBegin = True

   Call Ug.AddEditData

   If AutoCommit Then
      Call m_ADOConn.CommitTrans
   End If
   HasBegin = False
   IsOK = True

   AddEditConditionCommiss = True
   Exit Function

ErrorHandler:
   If HasBegin Then
      If AutoCommit Then
         m_ADOConn.RollbackTrans
      End If
   End If

   ErrorObj.LocalErrorMsg = "Runtime error."
   ErrorObj.SystemErrorMsg = Err.DESCRIPTION
   ErrorObj.RoutineName = RName
   ErrorObj.ModuleName = m_ModuleName
   ErrorObj.ShowErrorLog (LOG_TO_FILE)

   IsOK = False
   AddEditConditionCommiss = False
End Function

Public Function AddEditCommissYear(Ug As CCommissYear, IsOK As Boolean, AutoCommit As Boolean, ErrorObj As clsErrorLog)
On Error GoTo ErrorHandler
Dim RName As String
Dim HasBegin As Boolean

   RName = "AddEditCommissYear"
   AddEditCommissYear = False
   HasBegin = False

   If AutoCommit Then
      m_ADOConn.BeginTrans
   End If
   HasBegin = True

   Call Ug.AddEditData

   If AutoCommit Then
      Call m_ADOConn.CommitTrans
   End If
   HasBegin = False
   IsOK = True

   AddEditCommissYear = True
   Exit Function

ErrorHandler:
   If HasBegin Then
      If AutoCommit Then
         m_ADOConn.RollbackTrans
      End If
   End If

   ErrorObj.LocalErrorMsg = "Runtime error."
   ErrorObj.SystemErrorMsg = Err.DESCRIPTION
   ErrorObj.RoutineName = RName
   ErrorObj.ModuleName = m_ModuleName
   ErrorObj.ShowErrorLog (LOG_TO_FILE)

   IsOK = False
   AddEditCommissYear = False
End Function
'===
Public Function QueryCommissYear(Ua As CCommissYear, Rs As ADODB.Recordset, ItemCount As Long, IsOK As Boolean, ErrorObj As clsErrorLog) As Boolean
On Error GoTo ErrorHandler
Dim RName As String
Dim SQL1 As String
Dim SQL2 As String
Dim SelectStr As String
Dim FromStr As String
Dim WhereStr As String
Dim OrderStr As String
Dim OrderType As String
Dim iCount As Long


   RName = "QueryCommissYear"
QueryCommissYear = False

   IsOK = True
   Call Ua.QueryData(1, Rs, ItemCount)

   QueryCommissYear = True
   Exit Function

ErrorHandler:
   ErrorObj.SystemErrorMsg = Err.DESCRIPTION
   ErrorObj.RoutineName = RName
   ErrorObj.ModuleName = m_ModuleName
   ErrorObj.ShowErrorLog (LOG_FILE_MSGBOX)

   QueryCommissYear = False
End Function


Public Function DeleteCommissYear(UID As Long, IsOK As Boolean, AutoCommit As Boolean, ErrorObj As clsErrorLog) As Boolean
On Error GoTo ErrorHandler
Dim RName As String
Dim SQL As String
Dim HasBegin As Boolean
Dim ItemCount As Long
Dim B As CCommissYear

   RName = "DeleteCommissYear"
   DeleteCommissYear = False
   HasBegin = False

   Set B = New CCommissYear

   If AutoCommit Then
      Call m_ADOConn.BeginTrans
   End If
   HasBegin = True

   B.YEAR_ID = UID
   Call B.DeleteData

   Set B = Nothing

   If AutoCommit Then
      Call m_ADOConn.CommitTrans
   End If
   HasBegin = False
   IsOK = True

   DeleteCommissYear = True
   Exit Function

ErrorHandler:
   If HasBegin Then
      If AutoCommit Then
         m_ADOConn.RollbackTrans
      End If
   End If

   ErrorObj.LocalErrorMsg = "Runtime error."
   ErrorObj.SystemErrorMsg = Err.DESCRIPTION
   ErrorObj.RoutineName = RName
   ErrorObj.ModuleName = m_ModuleName
   ErrorObj.LocalErrorMsg = "ไม่สามารถเปิดข้อมูลได้ เนื่องจากข้อมูลถูกอ้างถึงโดยส่วนอื่น"
   ErrorObj.ShowErrorLog (LOG_TO_FILE)

   IsOK = False

   DeleteCommissYear = True
End Function

Public Function QueryIncenForSum(Ua As CCondiIncenSum, Rs As ADODB.Recordset, ItemCount As Long, IsOK As Boolean, ErrorObj As clsErrorLog) As Boolean
On Error GoTo ErrorHandler
Dim RName As String
Dim SQL1 As String
Dim SQL2 As String
Dim SelectStr As String
Dim FromStr As String
Dim WhereStr As String
Dim OrderStr As String
Dim OrderType As String
Dim NewStr As String
Dim SubLen As Long
Dim iCount As Long

   RName = "QueryIncenForSum"
QueryIncenForSum = False

   IsOK = True
   Call Ua.QueryData(1, Rs, ItemCount)         ' CInventorydoc

   QueryIncenForSum = True
   Exit Function

ErrorHandler:
   ErrorObj.SystemErrorMsg = Err.DESCRIPTION
   ErrorObj.RoutineName = RName
   ErrorObj.ModuleName = m_ModuleName
   ErrorObj.ShowErrorLog (LOG_FILE_MSGBOX)

QueryIncenForSum = False
End Function
'==
Public Function QueryConditionCom(Ua As CConditionCommission, Rs As ADODB.Recordset, ItemCount As Long, IsOK As Boolean, ErrorObj As clsErrorLog) As Boolean
On Error GoTo ErrorHandler
Dim RName As String
Dim SQL1 As String
Dim SQL2 As String
Dim SelectStr As String
Dim FromStr As String
Dim WhereStr As String
Dim OrderStr As String
Dim OrderType As String
Dim NewStr As String
Dim SubLen As Long
Dim iCount As Long

   RName = "QueryConditionCom"
QueryConditionCom = False

   IsOK = True
   Call Ua.QueryData(1, Rs, ItemCount)         ' CInventorydoc

   QueryConditionCom = True
   Exit Function

ErrorHandler:
   ErrorObj.SystemErrorMsg = Err.DESCRIPTION
   ErrorObj.RoutineName = RName
   ErrorObj.ModuleName = m_ModuleName
   ErrorObj.ShowErrorLog (LOG_FILE_MSGBOX)

QueryConditionCom = False
End Function
Public Function DeleteIncenForSum(UID As Long, IsOK As Boolean, AutoCommit As Boolean, ErrorObj As clsErrorLog)
On Error GoTo ErrorHandler
Dim RName As String
Dim SQL As String
Dim HasBegin As Boolean
Dim ItemCount As Long
Dim B As CCondiIncenSum

   RName = "DeleteIncenForSum"
   DeleteIncenForSum = False
   HasBegin = False

   Set B = New CCondiIncenSum

   If AutoCommit Then
      m_ADOConn.BeginTrans
   End If
   HasBegin = True

   B.INCEN_SLM_FORSUM_ID = UID
   Call B.DeleteData(1)

   Set B = Nothing

   If AutoCommit Then
      Call m_ADOConn.CommitTrans
   End If
   HasBegin = False
   IsOK = True

   DeleteIncenForSum = True
   Exit Function

ErrorHandler:
   If HasBegin Then
      If AutoCommit Then
         m_ADOConn.RollbackTrans
      End If
   End If

   ErrorObj.LocalErrorMsg = "Runtime error."
   ErrorObj.SystemErrorMsg = Err.DESCRIPTION
   ErrorObj.RoutineName = RName
   ErrorObj.ModuleName = m_ModuleName
   ErrorObj.LocalErrorMsg = "ไม่สามารถลบข้อมูลได้เนื่องจากข้อมูลนี้ยังถูกอ้างถึงอยู่จากส่วนอื่น"
   ErrorObj.ShowErrorLog (LOG_TO_FILE)

   IsOK = False

   DeleteIncenForSum = True
End Function

Public Function DeleteConditionCom(UID As Long, IsOK As Boolean, AutoCommit As Boolean, ErrorObj As clsErrorLog)
On Error GoTo ErrorHandler
Dim RName As String
Dim SQL As String
Dim HasBegin As Boolean
Dim ItemCount As Long
Dim B As CConditionCommission

   RName = "DeleteConditionCom"
   DeleteConditionCom = False
   HasBegin = False

   Set B = New CConditionCommission

   If AutoCommit Then
      m_ADOConn.BeginTrans
   End If
   HasBegin = True

   B.COM_ID = UID
   Call B.DeleteData(1)

   Set B = Nothing

   If AutoCommit Then
      Call m_ADOConn.CommitTrans
   End If
   HasBegin = False
   IsOK = True

   DeleteConditionCom = True
   Exit Function

ErrorHandler:
   If HasBegin Then
      If AutoCommit Then
         m_ADOConn.RollbackTrans
      End If
   End If

   ErrorObj.LocalErrorMsg = "Runtime error."
   ErrorObj.SystemErrorMsg = Err.DESCRIPTION
   ErrorObj.RoutineName = RName
   ErrorObj.ModuleName = m_ModuleName
   ErrorObj.LocalErrorMsg = "ไม่สามารถลบข้อมูลได้เนื่องจากข้อมูลนี้ยังถูกอ้างถึงอยู่จากส่วนอื่น"
   ErrorObj.ShowErrorLog (LOG_TO_FILE)

   IsOK = False

   DeleteConditionCom = True
End Function

Public Function DeleteConditionComFromYear(UID As Long, IsOK As Boolean, AutoCommit As Boolean, ErrorObj As clsErrorLog)
On Error GoTo ErrorHandler
Dim RName As String
Dim SQL As String
Dim HasBegin As Boolean
Dim ItemCount As Long
Dim B As CConditionCommission

   RName = "DeleteConditionComFromYear"
   DeleteConditionComFromYear = False
   HasBegin = False

   Set B = New CConditionCommission

   If AutoCommit Then
      m_ADOConn.BeginTrans
   End If
   HasBegin = True

   B.YEAR_ID = UID
   Call B.DeleteData(2)

   Set B = Nothing

   If AutoCommit Then
      Call m_ADOConn.CommitTrans
   End If
   HasBegin = False
   IsOK = True

   DeleteConditionComFromYear = True
   Exit Function

ErrorHandler:
   If HasBegin Then
      If AutoCommit Then
         m_ADOConn.RollbackTrans
      End If
   End If

   ErrorObj.LocalErrorMsg = "Runtime error."
   ErrorObj.SystemErrorMsg = Err.DESCRIPTION
   ErrorObj.RoutineName = RName
   ErrorObj.ModuleName = m_ModuleName
   ErrorObj.LocalErrorMsg = "ไม่สามารถลบข้อมูลได้เนื่องจากข้อมูลนี้ยังถูกอ้างถึงอยู่จากส่วนอื่น"
   ErrorObj.ShowErrorLog (LOG_TO_FILE)

   IsOK = False

   DeleteConditionComFromYear = True
End Function

Public Function QueryComType1(Ua As CConditionCommission, Rs As ADODB.Recordset, ItemCount As Long, IsOK As Boolean, ErrorObj As clsErrorLog) As Boolean
On Error GoTo ErrorHandler
Dim RName As String
Dim SQL1 As String
Dim SQL2 As String
Dim SelectStr As String
Dim FromStr As String
Dim WhereStr As String
Dim OrderStr As String
Dim OrderType As String
Dim NewStr As String
Dim SubLen As Long
Dim iCount As Long

   RName = "QueryComType1"
   QueryComType1 = False

   IsOK = True
   Call Ua.QueryData(2, Rs, ItemCount)
   
   QueryComType1 = True
   Exit Function

ErrorHandler:
   ErrorObj.SystemErrorMsg = Err.DESCRIPTION
   ErrorObj.RoutineName = RName
   ErrorObj.ModuleName = m_ModuleName
   ErrorObj.ShowErrorLog (LOG_FILE_MSGBOX)

   QueryComType1 = False
End Function

Public Function DeleteCommissArea(UID As Long, IsOK As Boolean, AutoCommit As Boolean, ErrorObj As clsErrorLog) As Boolean
On Error GoTo ErrorHandler
Dim RName As String
Dim SQL As String
Dim HasBegin As Boolean
Dim ItemCount As Long
Dim B As CCommissMasterArea

   RName = "DeleteCommissArea"
   DeleteCommissArea = False
   HasBegin = False

   Set B = New CCommissMasterArea

   If AutoCommit Then
      Call m_ADOConn.BeginTrans
   End If
   HasBegin = True

   B.MASTER_AREA_ID = UID
   Call B.DeleteData

   Set B = Nothing

   If AutoCommit Then
      Call m_ADOConn.CommitTrans
   End If
   HasBegin = False
   IsOK = True

   DeleteCommissArea = True
   Exit Function

ErrorHandler:
   If HasBegin Then
      If AutoCommit Then
         m_ADOConn.RollbackTrans
      End If
   End If

   ErrorObj.LocalErrorMsg = "Runtime error."
   ErrorObj.SystemErrorMsg = Err.DESCRIPTION
   ErrorObj.RoutineName = RName
   ErrorObj.ModuleName = m_ModuleName
   ErrorObj.LocalErrorMsg = "ไม่สามารถเปิดข้อมูลได้ เนื่องจากข้อมูลถูกอ้างถึงโดยส่วนอื่น"
   ErrorObj.ShowErrorLog (LOG_TO_FILE)

   IsOK = False

   DeleteCommissArea = True
End Function

Public Function QueryCommissArea(Ua As CCommissMasterArea, Rs As ADODB.Recordset, ItemCount As Long, IsOK As Boolean, ErrorObj As clsErrorLog) As Boolean
On Error GoTo ErrorHandler
Dim RName As String
Dim OrderStr As String
Dim OrderType As String
Dim iCount As Long


   RName = "QueryCommissArea"
QueryCommissArea = False

   IsOK = True
   Call Ua.QueryData(1, Rs, ItemCount)

  QueryCommissArea = True
   Exit Function

ErrorHandler:
   ErrorObj.SystemErrorMsg = Err.DESCRIPTION
   ErrorObj.RoutineName = RName
   ErrorObj.ModuleName = m_ModuleName
   ErrorObj.ShowErrorLog (LOG_FILE_MSGBOX)

  QueryCommissArea = False
End Function


Public Function QueryComboSupDe(Ua As CComboSubGroupDe, Rs As ADODB.Recordset, ItemCount As Long, IsOK As Boolean, ErrorObj As clsErrorLog) As Boolean
On Error GoTo ErrorHandler
Dim RName As String
Dim OrderStr As String
Dim OrderType As String
Dim iCount As Long


   RName = "QueryComboSupDe"
QueryComboSupDe = False

   IsOK = True
   Call Ua.QueryData(Rs, ItemCount)

  QueryComboSupDe = True
   Exit Function

ErrorHandler:
   ErrorObj.SystemErrorMsg = Err.DESCRIPTION
   ErrorObj.RoutineName = RName
   ErrorObj.ModuleName = m_ModuleName
   ErrorObj.ShowErrorLog (LOG_FILE_MSGBOX)

  QueryComboSupDe = False
End Function

Public Function DeleteCusAreaCom(UID As Long, IsOK As Boolean, AutoCommit As Boolean, ErrorObj As clsErrorLog)
On Error GoTo ErrorHandler
Dim RName As String
Dim SQL As String
Dim HasBegin As Boolean
Dim ItemCount As Long
Dim B As CCommissionCustomerArea

   RName = "DeleteCusAreaCom"
   DeleteCusAreaCom = False
   HasBegin = False

   Set B = New CCommissionCustomerArea

   If AutoCommit Then
      m_ADOConn.BeginTrans
   End If
   HasBegin = True

   B.COMMISSION_CUS_AREA_ID = UID
   Call B.DeleteData(1)

   Set B = Nothing

   If AutoCommit Then
      Call m_ADOConn.CommitTrans
   End If
   HasBegin = False
   IsOK = True

   DeleteCusAreaCom = True
   Exit Function

ErrorHandler:
   If HasBegin Then
      If AutoCommit Then
         m_ADOConn.RollbackTrans
      End If
   End If

   ErrorObj.LocalErrorMsg = "Runtime error."
   ErrorObj.SystemErrorMsg = Err.DESCRIPTION
   ErrorObj.RoutineName = RName
   ErrorObj.ModuleName = m_ModuleName
   ErrorObj.LocalErrorMsg = "ไม่สามารถลบข้อมูลได้เนื่องจากข้อมูลนี้ยังถูกอ้างถึงอยู่จากส่วนอื่น"
   ErrorObj.ShowErrorLog (LOG_TO_FILE)

   IsOK = False

   DeleteCusAreaCom = True
End Function

Public Function DeleteCusAreaComFromMaster(UID As Long, IsOK As Boolean, AutoCommit As Boolean, ErrorObj As clsErrorLog)
On Error GoTo ErrorHandler
Dim RName As String
Dim SQL As String
Dim HasBegin As Boolean
Dim ItemCount As Long
Dim B As CCommissionCustomerArea

   RName = "DeleteCusAreaComFromMaster"
   DeleteCusAreaComFromMaster = False
   HasBegin = False

   Set B = New CCommissionCustomerArea

   If AutoCommit Then
      m_ADOConn.BeginTrans
   End If
   HasBegin = True

   B.MASTER_AREA_ID = UID
 '  B.YEAR_ID = YID
   Call B.DeleteData(2)

   Set B = Nothing

   If AutoCommit Then
      Call m_ADOConn.CommitTrans
   End If
   HasBegin = False
   IsOK = True

   DeleteCusAreaComFromMaster = True
   Exit Function

ErrorHandler:
   If HasBegin Then
      If AutoCommit Then
         m_ADOConn.RollbackTrans
      End If
   End If

   ErrorObj.LocalErrorMsg = "Runtime error."
   ErrorObj.SystemErrorMsg = Err.DESCRIPTION
   ErrorObj.RoutineName = RName
   ErrorObj.ModuleName = m_ModuleName
   ErrorObj.LocalErrorMsg = "ไม่สามารถลบข้อมูลได้เนื่องจากข้อมูลนี้ยังถูกอ้างถึงอยู่จากส่วนอื่น"
   ErrorObj.ShowErrorLog (LOG_TO_FILE)

   IsOK = False

   DeleteCusAreaComFromMaster = True
End Function

Public Function QueryCusAreaCom(Ua As CCommissionCustomerArea, Rs As ADODB.Recordset, ItemCount As Long, IsOK As Boolean, ErrorObj As clsErrorLog) As Boolean
On Error GoTo ErrorHandler
Dim RName As String
Dim OrderStr As String
Dim OrderType As String

   RName = "QueryCusAreaCom"
QueryCusAreaCom = False

   IsOK = True
   Call Ua.QueryData(1, Rs, ItemCount)         ' CInventorydoc

   QueryCusAreaCom = True
   Exit Function

ErrorHandler:
   ErrorObj.SystemErrorMsg = Err.DESCRIPTION
   ErrorObj.RoutineName = RName
   ErrorObj.ModuleName = m_ModuleName
   ErrorObj.ShowErrorLog (LOG_FILE_MSGBOX)

QueryCusAreaCom = False
End Function

Public Function QueryCusDetailAreaCom(Ua As CCommissionCustomerArea, Rs As ADODB.Recordset, ItemCount As Long, IsOK As Boolean, ErrorObj As clsErrorLog) As Boolean
On Error GoTo ErrorHandler
Dim RName As String
Dim OrderStr As String
Dim OrderType As String

   RName = "QueryCusDetailAreaCom"
   QueryCusDetailAreaCom = False

   IsOK = True
   Call Ua.QueryData(3, Rs, ItemCount)         ' CInventorydoc

   QueryCusDetailAreaCom = True
   Exit Function

ErrorHandler:
   ErrorObj.SystemErrorMsg = Err.DESCRIPTION
   ErrorObj.RoutineName = RName
   ErrorObj.ModuleName = m_ModuleName
   ErrorObj.ShowErrorLog (LOG_FILE_MSGBOX)

QueryCusDetailAreaCom = False
End Function

Public Function AddEditComMasterArea(Ug As CCommissMasterArea, IsOK As Boolean, AutoCommit As Boolean, ErrorObj As clsErrorLog)
On Error GoTo ErrorHandler
Dim RName As String
Dim HasBegin As Boolean

   RName = "AddEditComMasterArea"
   AddEditComMasterArea = False
   HasBegin = False

   If AutoCommit Then
      m_ADOConn.BeginTrans
   End If
   HasBegin = True

   Call Ug.AddEditData

   If AutoCommit Then
      Call m_ADOConn.CommitTrans
   End If
   HasBegin = False
   IsOK = True

   AddEditComMasterArea = True
   Exit Function

ErrorHandler:
   If HasBegin Then
      If AutoCommit Then
         m_ADOConn.RollbackTrans
      End If
   End If
      ErrorObj.LocalErrorMsg = "Runtime error."
   ErrorObj.SystemErrorMsg = Err.DESCRIPTION
   ErrorObj.RoutineName = RName
   ErrorObj.ModuleName = m_ModuleName
   ErrorObj.ShowErrorLog (LOG_TO_FILE)

   IsOK = False
   AddEditComMasterArea = False
End Function

   
Public Function AddEditCusArea(Ug As CCommissionCustomerArea, IsOK As Boolean, AutoCommit As Boolean, ErrorObj As clsErrorLog)
On Error GoTo ErrorHandler
Dim RName As String
Dim HasBegin As Boolean

   RName = "AddEditCusArea"
   AddEditCusArea = False
   HasBegin = False

   If AutoCommit Then
      m_ADOConn.BeginTrans
   End If
   HasBegin = True

   Call Ug.AddEditData

   If AutoCommit Then
      Call m_ADOConn.CommitTrans
   End If
   HasBegin = False
   IsOK = True

   AddEditCusArea = True
   Exit Function

ErrorHandler:
   If HasBegin Then
      If AutoCommit Then
         m_ADOConn.RollbackTrans
      End If
   End If

   ErrorObj.LocalErrorMsg = "Runtime error."
   ErrorObj.SystemErrorMsg = "  ลูกค้ารหัส'" & Ug.COMMISSION_CUS_ID & "' มีอยู่ในเขตการขายอื่นแล้ว กรุณาลบรหัสลูกค้าดังกล่าวในเขตการขายอื่นก่อน แล้วทำรายการใหม่  "
   ErrorObj.RoutineName = RName
   ErrorObj.ModuleName = m_ModuleName
   ErrorObj.ShowErrorLog (LOG_TO_FILE)

   IsOK = False
   AddEditCusArea = False
End Function

Public Function QueryCusAreaCom2(Ua As CCommissionCustomerArea, Rs As ADODB.Recordset, ItemCount As Long, IsOK As Boolean, ErrorObj As clsErrorLog) As Boolean
On Error GoTo ErrorHandler
Dim RName As String
Dim OrderStr As String
Dim OrderType As String

   RName = "QueryCusAreaCom2"
QueryCusAreaCom2 = False

   IsOK = True
   Call Ua.QueryData(2, Rs, ItemCount)         ' CInventorydoc

   QueryCusAreaCom2 = True
   Exit Function

ErrorHandler:
   ErrorObj.SystemErrorMsg = Err.DESCRIPTION
   ErrorObj.RoutineName = RName
   ErrorObj.ModuleName = m_ModuleName
   ErrorObj.ShowErrorLog (LOG_FILE_MSGBOX)

QueryCusAreaCom2 = False
End Function

Public Function QueryIVcredit(Ua As CComIVcredit, Cl As Collection, Rs As ADODB.Recordset, ItemCount As Long, IsOK As Boolean, ErrorObj As clsErrorLog) As Boolean
On Error GoTo ErrorHandler
Dim RName As String
Dim OrderStr As String
Dim OrderType As String
Dim TempData As CComIVcredit

   Set Rs = New ADODB.Recordset
   RName = "QueryIVcredit"
   QueryIVcredit = False

   IsOK = True
   Call Ua.QueryData(1, Rs, ItemCount)
   
     If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   
   While Not Rs.EOF
      Set TempData = New CComIVcredit
      Call TempData.PopulateFromRS(1, Rs)

      If Not (Cl Is Nothing) Then
         Call Cl.Add(TempData, Trim(TempData.IV_COD))   ' KEy
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   If ItemCount > 0 Then
        Call Rs.MoveFirst
   End If
   QueryIVcredit = True
   Exit Function

ErrorHandler:
   ErrorObj.SystemErrorMsg = Err.DESCRIPTION
   ErrorObj.RoutineName = RName
   ErrorObj.ModuleName = m_ModuleName
   ErrorObj.ShowErrorLog (LOG_FILE_MSGBOX)

QueryIVcredit = False
End Function

Public Function DeleteIVcredit(B As CComIVcredit, IsOK As Boolean, AutoCommit As Boolean, ErrorObj As clsErrorLog)
On Error GoTo ErrorHandler
Dim RName As String
Dim SQL As String
Dim HasBegin As Boolean
Dim ItemCount As Long

   RName = "DeleteIVcredit"
   DeleteIVcredit = False
   HasBegin = False

   If AutoCommit Then
      m_ADOConn.BeginTrans
   End If
   HasBegin = True

   Call B.DeleteData
   
   If AutoCommit Then
      Call m_ADOConn.CommitTrans
   End If
   HasBegin = False
   IsOK = True

   DeleteIVcredit = True
   Exit Function

ErrorHandler:
   If HasBegin Then
      If AutoCommit Then
         m_ADOConn.RollbackTrans
      End If
   End If

   ErrorObj.LocalErrorMsg = "Runtime error."
   ErrorObj.SystemErrorMsg = Err.DESCRIPTION
   ErrorObj.RoutineName = RName
   ErrorObj.ModuleName = m_ModuleName
   ErrorObj.LocalErrorMsg = "ไม่สามารถลบข้อมูลได้เนื่องจากข้อมูลนี้ยังถูกอ้างถึงอยู่จากส่วนอื่น"
   ErrorObj.ShowErrorLog (LOG_TO_FILE)

   IsOK = False

   DeleteIVcredit = True
End Function


Public Function AddEditIVcredit(Ug As CComIVcredit, IsOK As Boolean, AutoCommit As Boolean, ErrorObj As clsErrorLog)
On Error GoTo ErrorHandler
Dim RName As String
Dim HasBegin As Boolean

   RName = "AddEditIVcredit"
  AddEditIVcredit = False
   HasBegin = False

   If AutoCommit Then
      m_ADOConn.BeginTrans
   End If
   HasBegin = True

   Call Ug.AddEditData

   If AutoCommit Then
      Call m_ADOConn.CommitTrans
   End If
   HasBegin = False
   IsOK = True

   AddEditIVcredit = True
   Exit Function

ErrorHandler:
   If HasBegin Then
      If AutoCommit Then
         m_ADOConn.RollbackTrans
      End If
   End If

   ErrorObj.LocalErrorMsg = "Runtime error."
   ErrorObj.SystemErrorMsg = Err.DESCRIPTION
   ErrorObj.RoutineName = RName
   ErrorObj.ModuleName = m_ModuleName
   ErrorObj.ShowErrorLog (LOG_TO_FILE)

   IsOK = False
   AddEditIVcredit = False
End Function
Public Function QueryIVcenter(Ua As CComIVcenter, Cl As Collection, Rs As ADODB.Recordset, ItemCount As Long, IsOK As Boolean, ErrorObj As clsErrorLog) As Boolean
On Error GoTo ErrorHandler
Dim RName As String
Dim OrderStr As String
Dim OrderType As String
Dim TempData As CComIVcenter

   Set Rs = New ADODB.Recordset
   RName = "QueryIVcenter"
   QueryIVcenter = False

   IsOK = True
   Call Ua.QueryData(1, Rs, ItemCount)
   
     If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   
   While Not Rs.EOF
      Set TempData = New CComIVcenter
      Call TempData.PopulateFromRS(1, Rs)

      If Not (Cl Is Nothing) Then
         Call Cl.Add(TempData, Trim(TempData.IV_COD))   ' KEy
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   If ItemCount > 0 Then
        Call Rs.MoveFirst
   End If
   QueryIVcenter = True
   Exit Function

ErrorHandler:
   ErrorObj.SystemErrorMsg = Err.DESCRIPTION
   ErrorObj.RoutineName = RName
   ErrorObj.ModuleName = m_ModuleName
   ErrorObj.ShowErrorLog (LOG_FILE_MSGBOX)

QueryIVcenter = False
End Function

Public Function DeleteIVcenter(B As CComIVcenter, IsOK As Boolean, AutoCommit As Boolean, ErrorObj As clsErrorLog)
On Error GoTo ErrorHandler
Dim RName As String
Dim SQL As String
Dim HasBegin As Boolean
Dim ItemCount As Long

   RName = "DeleteIVcenter"
   DeleteIVcenter = False
   HasBegin = False

   If AutoCommit Then
      m_ADOConn.BeginTrans
   End If
   HasBegin = True

   Call B.DeleteData
   
   If AutoCommit Then
      Call m_ADOConn.CommitTrans
   End If
   HasBegin = False
   IsOK = True

   DeleteIVcenter = True
   Exit Function

ErrorHandler:
   If HasBegin Then
      If AutoCommit Then
         m_ADOConn.RollbackTrans
      End If
   End If

   ErrorObj.LocalErrorMsg = "Runtime error."
   ErrorObj.SystemErrorMsg = Err.DESCRIPTION
   ErrorObj.RoutineName = RName
   ErrorObj.ModuleName = m_ModuleName
   ErrorObj.LocalErrorMsg = "ไม่สามารถลบข้อมูลได้เนื่องจากข้อมูลนี้ยังถูกอ้างถึงอยู่จากส่วนอื่น"
   ErrorObj.ShowErrorLog (LOG_TO_FILE)

   IsOK = False

   DeleteIVcenter = True
End Function


Public Function AddEditIVcenter(Ug As CComIVcenter, IsOK As Boolean, AutoCommit As Boolean, ErrorObj As clsErrorLog)
On Error GoTo ErrorHandler
Dim RName As String
Dim HasBegin As Boolean

   RName = "AddEditIVcenter"
  AddEditIVcenter = False
   HasBegin = False

   If AutoCommit Then
      m_ADOConn.BeginTrans
   End If
   HasBegin = True

   Call Ug.AddEditData

   If AutoCommit Then
      Call m_ADOConn.CommitTrans
   End If
   HasBegin = False
   IsOK = True

   AddEditIVcenter = True
   Exit Function

ErrorHandler:
   If HasBegin Then
      If AutoCommit Then
         m_ADOConn.RollbackTrans
      End If
   End If

   ErrorObj.LocalErrorMsg = "Runtime error."
   ErrorObj.SystemErrorMsg = Err.DESCRIPTION
   ErrorObj.RoutineName = RName
   ErrorObj.ModuleName = m_ModuleName
   ErrorObj.ShowErrorLog (LOG_TO_FILE)

   IsOK = False
   AddEditIVcenter = False
End Function

Public Function AddEditCarkillSum(Ug As CXlsCarkillSum, IsOK As Boolean, AutoCommit As Boolean, ErrorObj As clsErrorLog)
On Error GoTo ErrorHandler
Dim RName As String
Dim HasBegin As Boolean

   RName = "AddEditCarkillSum"
    AddEditCarkillSum = False
   HasBegin = False

   If AutoCommit Then
      m_ADOConn.BeginTrans
   End If
   HasBegin = True

   Call Ug.AddEditData

   If AutoCommit Then
      Call m_ADOConn.CommitTrans
   End If

   HasBegin = False
   IsOK = True

   AddEditCarkillSum = True
   Exit Function

ErrorHandler:
   If HasBegin Then
      If AutoCommit Then
         m_ADOConn.RollbackTrans
      End If
   End If

   ErrorObj.LocalErrorMsg = "Runtime error."
   ErrorObj.SystemErrorMsg = Err.DESCRIPTION
   ErrorObj.RoutineName = RName
   ErrorObj.ModuleName = m_ModuleName
   ErrorObj.ShowErrorLog (LOG_TO_FILE)

   IsOK = False
   AddEditCarkillSum = False
End Function

Public Function QueryCarkillSum(Ua As CXlsCarkillSum, Rs As ADODB.Recordset, ItemCount As Long, IsOK As Boolean, ErrorObj As clsErrorLog) As Boolean
On Error GoTo ErrorHandler
Dim RName As String
Dim OrderStr As String
Dim OrderType As String
Dim iCount As Long

  RName = "QueryCarkillSum"
  QueryCarkillSum = False

   IsOK = True
   Call Ua.QueryData(1, Rs, ItemCount)

  QueryCarkillSum = True
   Exit Function

ErrorHandler:
   ErrorObj.SystemErrorMsg = Err.DESCRIPTION
   ErrorObj.RoutineName = RName
   ErrorObj.ModuleName = m_ModuleName
   ErrorObj.ShowErrorLog (LOG_FILE_MSGBOX)

  QueryCarkillSum = False
End Function

Public Function DeleteCarkillSum(Ug As CXlsCarkillSum, IsOK As Boolean, AutoCommit As Boolean, ErrorObj As clsErrorLog)
On Error GoTo ErrorHandler
Dim RName As String
Dim SQL As String
Dim HasBegin As Boolean

   RName = "DeleteCarkillSum"
   DeleteCarkillSum = False
   HasBegin = False
   
   If AutoCommit Then
      m_ADOConn.BeginTrans
   End If
   HasBegin = True
   
   Call Ug.DeleteData
   
   If AutoCommit Then
      Call m_ADOConn.CommitTrans
   End If
   HasBegin = False
   IsOK = True

   DeleteCarkillSum = True
   Exit Function
   
ErrorHandler:
   If HasBegin Then
      If AutoCommit Then
         m_ADOConn.RollbackTrans
      End If
   End If
   
   ErrorObj.LocalErrorMsg = "Runtime error."
   ErrorObj.SystemErrorMsg = Err.DESCRIPTION
   ErrorObj.RoutineName = RName
   ErrorObj.ModuleName = m_ModuleName
 '  ErrorObj.LocalErrorMsg = "ไม่สามารถลบข้อมูลได้เนื่องจากข้อมูลนี้ยังถูกอ้างถึงอยู่จากส่วนอื่น"
   ErrorObj.ShowErrorLog (LOG_TO_FILE)
   
   IsOK = False
   DeleteCarkillSum = False
End Function

Public Function AddEditCarkillFW(Ug As CXlsCarkillFW, IsOK As Boolean, AutoCommit As Boolean, ErrorObj As clsErrorLog)
On Error GoTo ErrorHandler
Dim RName As String
Dim HasBegin As Boolean

   RName = "AddEditCarkillFW"
    AddEditCarkillFW = False
   HasBegin = False

   If AutoCommit Then
      m_ADOConn.BeginTrans
   End If
   HasBegin = True

   Call Ug.AddEditData

   If AutoCommit Then
      Call m_ADOConn.CommitTrans
   End If

   HasBegin = False
   IsOK = True

   AddEditCarkillFW = True
   Exit Function

ErrorHandler:
   If HasBegin Then
      If AutoCommit Then
         m_ADOConn.RollbackTrans
      End If
   End If

   ErrorObj.LocalErrorMsg = "Runtime error."
   ErrorObj.SystemErrorMsg = Err.DESCRIPTION
   ErrorObj.RoutineName = RName
   ErrorObj.ModuleName = m_ModuleName
   ErrorObj.ShowErrorLog (LOG_TO_FILE)

   IsOK = False
   AddEditCarkillFW = False
End Function


Public Function QueryCarkillFW(Ua As CXlsCarkillFW, Rs As ADODB.Recordset, ItemCount As Long, IsOK As Boolean, ErrorObj As clsErrorLog) As Boolean
On Error GoTo ErrorHandler
Dim RName As String
Dim OrderStr As String
Dim OrderType As String
Dim iCount As Long

  RName = "QueryCarkillFW"
  QueryCarkillFW = False

   IsOK = True
   Call Ua.QueryData(1, Rs, ItemCount)

  QueryCarkillFW = True
   Exit Function

ErrorHandler:
   ErrorObj.SystemErrorMsg = Err.DESCRIPTION
   ErrorObj.RoutineName = RName
   ErrorObj.ModuleName = m_ModuleName
   ErrorObj.ShowErrorLog (LOG_FILE_MSGBOX)

  QueryCarkillFW = False
End Function

Public Function DeleteCarkillFW(Ug As CXlsCarkillFW, IsOK As Boolean, AutoCommit As Boolean, ErrorObj As clsErrorLog)
On Error GoTo ErrorHandler
Dim RName As String
Dim SQL As String
Dim HasBegin As Boolean

   RName = "DeleteCarkillFW"
   DeleteCarkillFW = False
   HasBegin = False
   
   If AutoCommit Then
      m_ADOConn.BeginTrans
   End If
   HasBegin = True
   
   Call Ug.DeleteData
   
   If AutoCommit Then
      Call m_ADOConn.CommitTrans
   End If
   HasBegin = False
   IsOK = True

   DeleteCarkillFW = True
   Exit Function
   
ErrorHandler:
   If HasBegin Then
      If AutoCommit Then
         m_ADOConn.RollbackTrans
      End If
   End If
   
   ErrorObj.LocalErrorMsg = "Runtime error."
   ErrorObj.SystemErrorMsg = Err.DESCRIPTION
   ErrorObj.RoutineName = RName
   ErrorObj.ModuleName = m_ModuleName
   ErrorObj.ShowErrorLog (LOG_TO_FILE)
   
   IsOK = False
   DeleteCarkillFW = False
End Function

Public Function QueryXlsSetFarm(Ua As CXlsSetFarm, Rs As ADODB.Recordset, ItemCount As Long, IsOK As Boolean, ErrorObj As clsErrorLog) As Boolean
On Error GoTo ErrorHandler
Dim RName As String
Dim OrderStr As String
Dim OrderType As String
Dim TempData As CXlsSetFarm

   Set Rs = New ADODB.Recordset
   RName = "QueryXlsSetFarm"
   QueryXlsSetFarm = False

   IsOK = True
   Call Ua.QueryData(1, Rs, ItemCount)
   
   QueryXlsSetFarm = True
   Exit Function

ErrorHandler:
   ErrorObj.SystemErrorMsg = Err.DESCRIPTION
   ErrorObj.RoutineName = RName
   ErrorObj.ModuleName = m_ModuleName
   ErrorObj.ShowErrorLog (LOG_FILE_MSGBOX)

QueryXlsSetFarm = False
End Function

Public Function DeleteXlsSetFarm(B As CXlsSetFarm, IsOK As Boolean, AutoCommit As Boolean, ErrorObj As clsErrorLog)
On Error GoTo ErrorHandler
Dim RName As String
Dim SQL As String
Dim HasBegin As Boolean
Dim ItemCount As Long

   RName = "DeleteXlsSetFarm"
   DeleteXlsSetFarm = False
   HasBegin = False

   If AutoCommit Then
      m_ADOConn.BeginTrans
   End If
   HasBegin = True

   Call B.DeleteData
   
   If AutoCommit Then
      Call m_ADOConn.CommitTrans
   End If
   HasBegin = False
   IsOK = True

   DeleteXlsSetFarm = True
   Exit Function

ErrorHandler:
   If HasBegin Then
      If AutoCommit Then
         m_ADOConn.RollbackTrans
      End If
   End If

   ErrorObj.LocalErrorMsg = "Runtime error."
   ErrorObj.SystemErrorMsg = Err.DESCRIPTION
   ErrorObj.RoutineName = RName
   ErrorObj.ModuleName = m_ModuleName
   ErrorObj.ShowErrorLog (LOG_TO_FILE)

   IsOK = False

   DeleteXlsSetFarm = True
End Function


Public Function AddEditXlsSetFarm(Ug As CXlsSetFarm, IsOK As Boolean, AutoCommit As Boolean, ErrorObj As clsErrorLog)
On Error GoTo ErrorHandler
Dim RName As String
Dim HasBegin As Boolean

   RName = "AddEditXlsSetFarm"
  AddEditXlsSetFarm = False
   HasBegin = False

   If AutoCommit Then
      m_ADOConn.BeginTrans
   End If
   HasBegin = True

   Call Ug.AddEditData

   If AutoCommit Then
      Call m_ADOConn.CommitTrans
   End If
   HasBegin = False
   IsOK = True

   AddEditXlsSetFarm = True
   Exit Function

ErrorHandler:
   If HasBegin Then
      If AutoCommit Then
         m_ADOConn.RollbackTrans
      End If
   End If

   ErrorObj.LocalErrorMsg = "Runtime error."
   ErrorObj.SystemErrorMsg = Err.DESCRIPTION
   ErrorObj.RoutineName = RName
   ErrorObj.ModuleName = m_ModuleName
   ErrorObj.ShowErrorLog (LOG_TO_FILE)

   IsOK = False
   AddEditXlsSetFarm = False
End Function

Public Function QueryXlsFood(Ua As CXlsFood, Rs As ADODB.Recordset, ItemCount As Long, IsOK As Boolean, ErrorObj As clsErrorLog) As Boolean
On Error GoTo ErrorHandler
Dim RName As String
Dim OrderStr As String
Dim OrderType As String
Dim TempData As CXlsFood

   Set Rs = New ADODB.Recordset
   RName = "QueryXlsFood"
   QueryXlsFood = False

   IsOK = True
   Call Ua.QueryData(1, Rs, ItemCount)
   
'     If Not (Cl Is Nothing) Then
'      Set Cl = Nothing
'      Set Cl = New Collection
'   End If
'
'   While Not Rs.EOF
'      Set TempData = New CXlsFood
'      Call TempData.PopulateFromRS(1, Rs)
'
'      If Not (Cl Is Nothing) Then
'         Call Cl.Add(TempData, Trim(TempData.XLS_FOOD_ID))    ' KEy
'      End If
'
'      Set TempData = Nothing
'      Rs.MoveNext
'   Wend
'
'   If ItemCount > 0 Then
'        Call Rs.MoveFirst
'   End If
   QueryXlsFood = True
   Exit Function

ErrorHandler:
   ErrorObj.SystemErrorMsg = Err.DESCRIPTION
   ErrorObj.RoutineName = RName
   ErrorObj.ModuleName = m_ModuleName
   ErrorObj.ShowErrorLog (LOG_FILE_MSGBOX)

QueryXlsFood = False
End Function

Public Function DeleteXlsFood(B As CXlsFood, IsOK As Boolean, AutoCommit As Boolean, ErrorObj As clsErrorLog)
On Error GoTo ErrorHandler
Dim RName As String
Dim SQL As String
Dim HasBegin As Boolean
Dim ItemCount As Long

   RName = "DeleteXlsFood"
   DeleteXlsFood = False
   HasBegin = False

   If AutoCommit Then
      m_ADOConn.BeginTrans
   End If
   HasBegin = True

   Call B.DeleteData
   
   If AutoCommit Then
      Call m_ADOConn.CommitTrans
   End If
   HasBegin = False
   IsOK = True

   DeleteXlsFood = True
   Exit Function

ErrorHandler:
   If HasBegin Then
      If AutoCommit Then
         m_ADOConn.RollbackTrans
      End If
   End If

   ErrorObj.LocalErrorMsg = "Runtime error."
   ErrorObj.SystemErrorMsg = Err.DESCRIPTION
   ErrorObj.RoutineName = RName
   ErrorObj.ModuleName = m_ModuleName
   ErrorObj.LocalErrorMsg = "ไม่สามารถลบข้อมูลได้เนื่องจากข้อมูลนี้ยังถูกอ้างถึงอยู่จากส่วนอื่น"
   ErrorObj.ShowErrorLog (LOG_TO_FILE)

   IsOK = False

   DeleteXlsFood = True
End Function


Public Function AddEditXlsFood(Ug As CXlsFood, IsOK As Boolean, AutoCommit As Boolean, ErrorObj As clsErrorLog)
On Error GoTo ErrorHandler
Dim RName As String
Dim HasBegin As Boolean

   RName = "AddEditXlsFood"
  AddEditXlsFood = False
   HasBegin = False

   If AutoCommit Then
      m_ADOConn.BeginTrans
   End If
   HasBegin = True

   Call Ug.AddEditData

   If AutoCommit Then
      Call m_ADOConn.CommitTrans
   End If
   HasBegin = False
   IsOK = True

   AddEditXlsFood = True
   Exit Function

ErrorHandler:
   If HasBegin Then
      If AutoCommit Then
         m_ADOConn.RollbackTrans
      End If
   End If

   ErrorObj.LocalErrorMsg = "Runtime error."
   ErrorObj.SystemErrorMsg = Err.DESCRIPTION
   ErrorObj.RoutineName = RName
   ErrorObj.ModuleName = m_ModuleName
   ErrorObj.ShowErrorLog (LOG_TO_FILE)

   IsOK = False
   AddEditXlsFood = False
End Function

Public Function AddEditMinusAmount(Ug As CComMinusStk, IsOK As Boolean, AutoCommit As Boolean, ErrorObj As clsErrorLog)
On Error GoTo ErrorHandler
Dim RName As String
Dim HasBegin As Boolean

   RName = "AddEditMinusAmount"
  AddEditMinusAmount = False
   HasBegin = False

   If AutoCommit Then
      m_ADOConn.BeginTrans
   End If
   HasBegin = True

   Call Ug.AddEditData

   If AutoCommit Then
      Call m_ADOConn.CommitTrans
   End If
   HasBegin = False
   IsOK = True

   AddEditMinusAmount = True
   Exit Function

ErrorHandler:
   If HasBegin Then
      If AutoCommit Then
         m_ADOConn.RollbackTrans
      End If
   End If

   ErrorObj.LocalErrorMsg = "Runtime error."
   ErrorObj.SystemErrorMsg = Err.DESCRIPTION
   ErrorObj.RoutineName = RName
   ErrorObj.ModuleName = m_ModuleName
   ErrorObj.ShowErrorLog (LOG_TO_FILE)

   IsOK = False
   AddEditMinusAmount = False
End Function

Public Function QueryMinusAmount(Ua As CComMinusStk, Cl As Collection, Rs As ADODB.Recordset, ItemCount As Long, IsOK As Boolean, ErrorObj As clsErrorLog) As Boolean
On Error GoTo ErrorHandler
Dim RName As String
Dim OrderStr As String
Dim OrderType As String
Dim TempData As CComMinusStk

   Set Rs = New ADODB.Recordset
   RName = "QueryMinusAmount"
   QueryMinusAmount = False

   IsOK = True
   Call Ua.QueryData(1, Rs, ItemCount)
   
     If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   
   While Not Rs.EOF
      Set TempData = New CComMinusStk
      Call TempData.PopulateFromRS(1, Rs)

      If Not (Cl Is Nothing) Then
         Call Cl.Add(TempData, Trim(TempData.IV_COD & "-" & TempData.STK_COD))   ' KEy
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   If Rs.RecordCount > 0 Then
      Call Rs.MoveFirst
   End If
   QueryMinusAmount = True
   Exit Function

ErrorHandler:
   ErrorObj.SystemErrorMsg = Err.DESCRIPTION
   ErrorObj.RoutineName = RName
   ErrorObj.ModuleName = m_ModuleName
   ErrorObj.ShowErrorLog (LOG_FILE_MSGBOX)

QueryMinusAmount = False
End Function

Public Function DeleteMinusAmount(B As CComMinusStk, IsOK As Boolean, AutoCommit As Boolean, ErrorObj As clsErrorLog)
On Error GoTo ErrorHandler
Dim RName As String
Dim SQL As String
Dim HasBegin As Boolean
Dim ItemCount As Long

   RName = "DeleteMinusAmount"
   DeleteMinusAmount = False
   HasBegin = False

   If AutoCommit Then
      m_ADOConn.BeginTrans
   End If
   HasBegin = True

   Call B.DeleteData
   
   If AutoCommit Then
      Call m_ADOConn.CommitTrans
   End If
   HasBegin = False
   IsOK = True

   DeleteMinusAmount = True
   Exit Function

ErrorHandler:
   If HasBegin Then
      If AutoCommit Then
         m_ADOConn.RollbackTrans
      End If
   End If

   ErrorObj.LocalErrorMsg = "Runtime error."
   ErrorObj.SystemErrorMsg = Err.DESCRIPTION
   ErrorObj.RoutineName = RName
   ErrorObj.ModuleName = m_ModuleName
   ErrorObj.LocalErrorMsg = "ไม่สามารถลบข้อมูลได้เนื่องจากข้อมูลนี้ยังถูกอ้างถึงอยู่จากส่วนอื่น"
   ErrorObj.ShowErrorLog (LOG_TO_FILE)

   IsOK = False

   DeleteMinusAmount = True
End Function

Public Function AddEditMasterFromTo(Ua As CMasterFromTo, IsOK As Boolean, AutoCommit As Boolean, ErrorObj As clsErrorLog)
On Error GoTo ErrorHandler
Dim RName As String
Dim HasBegin As Boolean
Dim i As Long
   RName = "AddEditMasterFromTo"
   AddEditMasterFromTo = False
   HasBegin = False

   If AutoCommit Then
      m_ADOConn.BeginTrans
      HasBegin = True
   End If
   
   Call Ua.AddEditData
   
   If AutoCommit Then
      Call m_ADOConn.CommitTrans
   End If
   
   HasBegin = False
   IsOK = True

   AddEditMasterFromTo = True
   Exit Function

ErrorHandler:
   If HasBegin Then
      If AutoCommit Then
         m_ADOConn.RollbackTrans
      End If
   End If

   ErrorObj.LocalErrorMsg = "Runtime error."
   ErrorObj.SystemErrorMsg = Err.DESCRIPTION
   ErrorObj.RoutineName = RName
   ErrorObj.ModuleName = m_ModuleName
   ErrorObj.ShowErrorLog (LOG_TO_FILE)

   IsOK = False
   AddEditMasterFromTo = False
End Function

Public Function QueryMasterFromTo(Ua As CMasterFromTo, Rs As ADODB.Recordset, ItemCount As Long, IsOK As Boolean, ErrorObj As clsErrorLog) As Boolean
On Error GoTo ErrorHandler
Dim RName As String
Dim NewStr As String
Dim SubLen As Long
Dim iCount As Long

   RName = "QueryMasterFromTo"
   QueryMasterFromTo = False

   IsOK = True
   
   Call Ua.QueryData(1, Rs, ItemCount)
   
   QueryMasterFromTo = True
   Exit Function

ErrorHandler:
   ErrorObj.SystemErrorMsg = Err.DESCRIPTION
   ErrorObj.RoutineName = RName
   ErrorObj.ModuleName = m_ModuleName
   ErrorObj.ShowErrorLog (LOG_FILE_MSGBOX)
   
   QueryMasterFromTo = False
End Function

Public Function DeleteMasterFromTo(Ua As CMasterFromTo, IsOK As Boolean, AutoCommit As Boolean, ErrorObj As clsErrorLog)
On Error GoTo ErrorHandler
Dim RName As String
Dim SQL As String
Dim HasBegin As Boolean
Dim ItemCount As Long

   RName = "DeleteMasterFromTo"
   DeleteMasterFromTo = False
   HasBegin = False

   If AutoCommit Then
      m_ADOConn.BeginTrans
   End If
   HasBegin = True
      
   Call Ua.DeleteData
      
   If AutoCommit Then
      Call m_ADOConn.CommitTrans
   End If
   HasBegin = False
   IsOK = True

   DeleteMasterFromTo = True
   Exit Function

ErrorHandler:
   If HasBegin Then
      If AutoCommit Then
         m_ADOConn.RollbackTrans
      End If
   End If

   ErrorObj.LocalErrorMsg = "Runtime error."
   ErrorObj.SystemErrorMsg = Err.DESCRIPTION
   ErrorObj.RoutineName = RName
   ErrorObj.ModuleName = m_ModuleName
   ErrorObj.ShowErrorLog (LOG_TO_FILE)

   ErrorObj.LocalErrorMsg = "ไม่สามารถลบข้อมูลได้เนื่องจากข้อมูลนี้ยังถูกอ้างถึงอยู่จากส่วนอื่น"
   IsOK = False

   DeleteMasterFromTo = True
End Function

Public Function AddEditIncenForSum(Ua As CCondiIncenSum, IsOK As Boolean, AutoCommit As Boolean, ErrorObj As clsErrorLog)
On Error GoTo ErrorHandler
Dim RName As String
Dim HasBegin As Boolean

   RName = "AddEditIncenForSum"
   AddEditIncenForSum = False
   HasBegin = False

   If AutoCommit Then
      m_ADOConn.BeginTrans
      HasBegin = True
   End If
   
   Call Ua.AddEditData
   
   If AutoCommit Then
      Call m_ADOConn.CommitTrans
   End If
   
   HasBegin = False
   IsOK = True

   AddEditIncenForSum = True
   Exit Function

ErrorHandler:
   If HasBegin Then
      If AutoCommit Then
         m_ADOConn.RollbackTrans
      End If
   End If

   ErrorObj.LocalErrorMsg = "Runtime error."
   ErrorObj.SystemErrorMsg = Err.DESCRIPTION
   ErrorObj.RoutineName = RName
   ErrorObj.ModuleName = m_ModuleName
   ErrorObj.ShowErrorLog (LOG_TO_FILE)

   IsOK = False
   AddEditIncenForSum = False
End Function

Public Function AddEditCommissionChart(Ua As CCommissionChart, IsOK As Boolean, AutoCommit As Boolean, ErrorObj As clsErrorLog)
On Error GoTo ErrorHandler
Dim RName As String
Dim HasBegin As Boolean

   RName = "AddEditCommissionChart"
   AddEditCommissionChart = False
   HasBegin = False

   If AutoCommit Then
      m_ADOConn.BeginTrans
      HasBegin = True
   End If
   
   Call Ua.AddEditData
   
   If AutoCommit Then
      Call m_ADOConn.CommitTrans
   End If
   
   HasBegin = False
   IsOK = True

   AddEditCommissionChart = True
   Exit Function

ErrorHandler:
   If HasBegin Then
      If AutoCommit Then
         m_ADOConn.RollbackTrans
      End If
   End If

   ErrorObj.LocalErrorMsg = "Runtime error."
   ErrorObj.SystemErrorMsg = Err.DESCRIPTION
   ErrorObj.RoutineName = RName
   ErrorObj.ModuleName = m_ModuleName
   ErrorObj.ShowErrorLog (LOG_TO_FILE)

   IsOK = False
   AddEditCommissionChart = False
End Function
Public Function QueryCommissionChart(Ua As CCommissionChart, Rs As ADODB.Recordset, ItemCount As Long, IsOK As Boolean, ErrorObj As clsErrorLog) As Boolean
On Error GoTo ErrorHandler
Dim RName As String
Dim SQL1 As String
Dim SQL2 As String
Dim SelectStr As String
Dim FromStr As String
Dim WhereStr As String
Dim OrderStr As String
Dim OrderType As String
Dim NewStr As String
Dim SubLen As Long
Dim iCount As Long

   RName = "QueryCommissionChart"
   QueryCommissionChart = False

   IsOK = True

   Call Ua.QueryData(1, Rs, ItemCount)
   
   QueryCommissionChart = True
   Exit Function

ErrorHandler:
   ErrorObj.SystemErrorMsg = Err.DESCRIPTION
   ErrorObj.RoutineName = RName
   ErrorObj.ModuleName = m_ModuleName
   ErrorObj.ShowErrorLog (LOG_FILE_MSGBOX)

   QueryCommissionChart = False
End Function
Public Function DeleteCommissionChart(Ua As CCommissionChart, IsOK As Boolean, AutoCommit As Boolean, ErrorObj As clsErrorLog)
On Error GoTo ErrorHandler
Dim RName As String
Dim SQL As String
Dim HasBegin As Boolean

   RName = "DeleteCommissionChart"
   DeleteCommissionChart = False
   HasBegin = False

   If AutoCommit Then
      m_ADOConn.BeginTrans
   End If
   HasBegin = True

   Call Ua.DeleteData

   If AutoCommit Then
      Call m_ADOConn.CommitTrans
   End If
   HasBegin = False
   IsOK = True

   DeleteCommissionChart = True
   Exit Function

ErrorHandler:
   If HasBegin Then
      If AutoCommit Then
         m_ADOConn.RollbackTrans
      End If
   End If

   ErrorObj.LocalErrorMsg = "Runtime error."
   ErrorObj.SystemErrorMsg = Err.DESCRIPTION
   ErrorObj.RoutineName = RName
   ErrorObj.ModuleName = m_ModuleName
   ErrorObj.ShowErrorLog (LOG_TO_FILE)

   ErrorObj.LocalErrorMsg = "ไม่สามารถลบข้อมูลได้เนื่องจากข้อมูลนี้ยังถูกอ้างถึงอยู่จากส่วนอื่น"
   IsOK = False

   DeleteCommissionChart = True
End Function
Public Function CopyCommissYear(Ug As CCommissYear, IsOK As Boolean, AutoCommit As Boolean, ErrorObj As clsErrorLog)
On Error GoTo ErrorHandler
Dim RName As String
Dim HasBegin As Boolean
Dim iCount As Long
Dim TempRs As ADODB.Recordset
Dim ItemCount As Long
Dim cm_Rs As ADODB.Recordset
Dim itemCountCm As Long
Dim InForSum_Rs As ADODB.Recordset
Dim ItemCountInForSum As Long
Dim YearMax As Long

   RName = "CopyCommissYear"
   CopyCommissYear = False
   HasBegin = False
   
   Set TempRs = New ADODB.Recordset
    Set cm_Rs = New ADODB.Recordset
     Set InForSum_Rs = New ADODB.Recordset

   Ug.QueryFlag = 1
   Call QueryCommissYear(Ug, TempRs, iCount, IsOK, ErrorObj)

   If Not TempRs.EOF Then
      Call Ug.PopulateFromRS(1, TempRs)
   End If

      Dim Cm As CConditionCommission
      Set Cm = New CConditionCommission    ' 5 คอมมิชชั่น
      Cm.YEAR_ID = Ug.YEAR_ID
      Cm.FROM_DATE = -1
      Cm.TO_DATE = -1
      Call Cm.QueryData(2, cm_Rs, itemCountCm)         '  หาลูกของต้นฉบับทั้งหมด
      
      Dim Ins As CCondiIncenSum
      Set Ins = New CCondiIncenSum   ' 5 คอมมิชชั่น
      Ins.YEAR_ID = Ug.YEAR_ID
      Ins.FROM_CMPL_DATE = -1
      Ins.TO_CMPL_DATE = -1
      Call Ins.QueryData(2, InForSum_Rs, ItemCountInForSum)         '  หาลูกของต้นฉบับทั้งหมด

      Ug.AddEditMode = SHOW_ADD
      Ug.YEAR_ID = -1
      Ug.YEARNUM = "***" & Ug.YEARNUM
      Call AddEditCommissYear(Ug, IsOK, False, ErrorObj)      ' เพิ่มตัวแม่
      YearMax = getCommissYearMax(Ug, IsOK, ErrorObj) ' หาตัวแม่มากสุด
      
      While Not cm_Rs.EOF          ' เพิ่มตัวลูก
            Call Cm.PopulateFromRS(2, cm_Rs)
             Cm.YEAR_ID = YearMax
            Cm.AddEditMode = SHOW_ADD
            Call Cm.AddEditData
             cm_Rs.MoveNext                                                                                            ' วนคน
      Wend
      
      While Not InForSum_Rs.EOF          ' เพิ่มตัวลูก
            Call Ins.PopulateFromRS(2, InForSum_Rs)
             Ins.YEAR_ID = YearMax
             Ins.AddEditMode = SHOW_ADD
            Call Ins.AddEditData
             InForSum_Rs.MoveNext                                                                                            ' วนคน
      Wend

   If AutoCommit Then
      m_ADOConn.CommitTrans
   End If
   HasBegin = False
   
   If TempRs.State = adStateOpen Then
      Call TempRs.Close
   End If
   Set TempRs = Nothing
   
   If cm_Rs.State = adStateOpen Then
      Call cm_Rs.Close
   End If
   Set cm_Rs = Nothing
   
   If InForSum_Rs.State = adStateOpen Then
      Call InForSum_Rs.Close
   End If
   Set InForSum_Rs = Nothing
   
   HasBegin = False
   IsOK = True

   CopyCommissYear = True
   Exit Function

ErrorHandler:
   If HasBegin Then
      If AutoCommit Then
         m_ADOConn.RollbackTrans
      End If
   End If
   
   ErrorObj.LocalErrorMsg = "Runtime error."
   ErrorObj.SystemErrorMsg = Err.DESCRIPTION
   ErrorObj.RoutineName = RName
   ErrorObj.ModuleName = m_ModuleName
   ErrorObj.ShowErrorLog (LOG_TO_FILE)

   IsOK = False
   CopyCommissYear = False
End Function

Public Function CopyAreaYear(Ug As CAreaYear, IsOK As Boolean, AutoCommit As Boolean, ErrorObj As clsErrorLog)
On Error GoTo ErrorHandler
Dim RName As String
Dim HasBegin As Boolean
Dim iCount As Long
Dim TempRs As ADODB.Recordset
Dim child_Rs As ADODB.Recordset
Dim childCount As Long
Dim YearMax As Long

   RName = "CopyAreaYear"
   CopyAreaYear = False
   HasBegin = False
   
   Set TempRs = New ADODB.Recordset
    Set child_Rs = New ADODB.Recordset

   Ug.QueryFlag = 1
   Ug.FROM_DATE = -1
   Ug.TO_DATE = -1
   Call QueryAreaYear(Ug, TempRs, iCount, IsOK, ErrorObj)

   If Not TempRs.EOF Then
      Call Ug.PopulateFromRS(1, TempRs)
   End If

      Dim Cm As CCommissionCustomerArea
      Set Cm = New CCommissionCustomerArea
      Cm.YEAR_ID = Ug.YEAR_ID
      Cm.FROM_DATE = -1
      Cm.TO_DATE = -1
      Call Cm.QueryData(2, child_Rs, childCount)         '  หาลูกของต้นฉบับทั้งหมด

      Ug.AddEditMode = SHOW_ADD
      Ug.YEAR_ID = -1
      Ug.YEARNUM = "***" & Ug.YEARNUM
      Call AddEditAreaYear(Ug, IsOK, False, ErrorObj)      ' เพิ่มตัวแม่
      YearMax = getAreaYearMax(Ug, IsOK, ErrorObj) ' หาตัวแม่มากสุด
      
      While Not child_Rs.EOF          ' เพิ่มตัวลูก
            Call Cm.PopulateFromRS(2, child_Rs)
             Cm.YEAR_ID = YearMax
            Cm.AddEditMode = SHOW_ADD
            Call Cm.AddEditData
             child_Rs.MoveNext
      Wend

   If AutoCommit Then
      m_ADOConn.CommitTrans
   End If
   HasBegin = False
   
   If child_Rs.State = adStateOpen Then
      Call child_Rs.Close
   End If
   Set child_Rs = Nothing
   
    If TempRs.State = adStateOpen Then
      Call TempRs.Close
   End If
   Set TempRs = Nothing
   
   HasBegin = False
   IsOK = True

   CopyAreaYear = True
   Exit Function

ErrorHandler:
   If HasBegin Then
      If AutoCommit Then
         m_ADOConn.RollbackTrans
      End If
   End If
   
   ErrorObj.LocalErrorMsg = "Runtime error."
   ErrorObj.SystemErrorMsg = Err.DESCRIPTION
   ErrorObj.RoutineName = RName
   ErrorObj.ModuleName = m_ModuleName
   ErrorObj.ShowErrorLog (LOG_TO_FILE)

   IsOK = False
   CopyAreaYear = False
End Function

Public Function QueryMaster2FromTo(Ua As CMaster2FromTo, Rs As ADODB.Recordset, ItemCount As Long, IsOK As Boolean, ErrorObj As clsErrorLog) As Boolean
On Error GoTo ErrorHandler
Dim RName As String
Dim SQL1 As String
Dim SQL2 As String
Dim SelectStr As String
Dim FromStr As String
Dim WhereStr As String
Dim OrderStr As String
Dim OrderType As String
Dim NewStr As String
Dim SubLen As Long
Dim iCount As Long

   RName = "QueryMaster2FromTo"
   QueryMaster2FromTo = False

   IsOK = True

   Call Ua.QueryData(1, Rs, ItemCount)
   
   If Ua.QueryFlag = 1 Then
      Dim Gr As CMasterFromToDetail
      Set Gr = New CMasterFromToDetail
      Gr.MASTER_FROMTO_ID = Ua.MASTER_FROMTO_ID
      Call Gr.QueryData(1, m_Rs1, iCount)
      Set Gr = Nothing

      Set Ua.Details = Nothing
      Set Ua.Details = New Collection

      While Not m_Rs1.EOF
         Set Gr = New CMasterFromToDetail
         Call Gr.PopulateFromRS(1, m_Rs1)

         Gr.Flag = "I"
         Call Ua.Details.Add(Gr)
         Set Gr = Nothing

         m_Rs1.MoveNext
      Wend

      Dim Gx As CCommissMasterPara
      Set Gx = New CCommissMasterPara
      Gx.MASTER_FROMTO_ID = Ua.MASTER_FROMTO_ID
      Call Gx.QueryData(1, m_Rs1, iCount)
      Set Gx = Nothing

      Set Ua.CommissionExs = Nothing
      Set Ua.CommissionExs = New Collection

      While Not m_Rs1.EOF
         Set Gx = New CCommissMasterPara
         Call Gx.PopulateFromRS(1, m_Rs1)

         Gx.Flag = "I"
         Call Ua.CommissionExs.Add(Gx)
         Set Gx = Nothing

         m_Rs1.MoveNext
      Wend

   End If
   
   QueryMaster2FromTo = True
   Exit Function

ErrorHandler:
   ErrorObj.SystemErrorMsg = Err.DESCRIPTION
   ErrorObj.RoutineName = RName
   ErrorObj.ModuleName = m_ModuleName
   ErrorObj.ShowErrorLog (LOG_FILE_MSGBOX)
   
   
   QueryMaster2FromTo = False
End Function

Public Function DeleteMaster2FromTo(Ua As CMaster2FromTo, IsOK As Boolean, AutoCommit As Boolean, ErrorObj As clsErrorLog)
On Error GoTo ErrorHandler
Dim RName As String
Dim SQL As String
Dim HasBegin As Boolean
Dim ItemCount As Long

   RName = "DeleteMaster2FromTo"
   DeleteMaster2FromTo = False
   HasBegin = False

   If AutoCommit Then
      m_ADOConn.BeginTrans
   End If
   HasBegin = True
      
   Call Ua.DeleteData
      
   If AutoCommit Then
      Call m_ADOConn.CommitTrans
   End If
   HasBegin = False
   IsOK = True

   DeleteMaster2FromTo = True
   Exit Function

ErrorHandler:
   If HasBegin Then
      If AutoCommit Then
         m_ADOConn.RollbackTrans
      End If
   End If

   ErrorObj.LocalErrorMsg = "Runtime error."
   ErrorObj.SystemErrorMsg = Err.DESCRIPTION
   ErrorObj.RoutineName = RName
   ErrorObj.ModuleName = m_ModuleName
   ErrorObj.ShowErrorLog (LOG_TO_FILE)

   ErrorObj.LocalErrorMsg = "ไม่สามารถลบข้อมูลได้เนื่องจากข้อมูลนี้ยังถูกอ้างถึงอยู่จากส่วนอื่น"
   IsOK = False

   DeleteMaster2FromTo = True
End Function


Public Function DeleteComboSubGroup(Ua As CComboSubGroup, IsOK As Boolean, AutoCommit As Boolean, ErrorObj As clsErrorLog)
On Error GoTo ErrorHandler
Dim RName As String
Dim SQL As String
Dim HasBegin As Boolean
Dim ItemCount As Long

   RName = "DeleteComboSubGroup"
   DeleteComboSubGroup = False
   HasBegin = False

   If AutoCommit Then
      m_ADOConn.BeginTrans
   End If
   HasBegin = True
      
   Call Ua.DeleteData
      
   If AutoCommit Then
      Call m_ADOConn.CommitTrans
   End If
   HasBegin = False
   IsOK = True

   DeleteComboSubGroup = True
   Exit Function

ErrorHandler:
   If HasBegin Then
      If AutoCommit Then
         m_ADOConn.RollbackTrans
      End If
   End If

   ErrorObj.LocalErrorMsg = "Runtime error."
   ErrorObj.SystemErrorMsg = Err.DESCRIPTION
   ErrorObj.RoutineName = RName
   ErrorObj.ModuleName = m_ModuleName
   ErrorObj.ShowErrorLog (LOG_TO_FILE)

   ErrorObj.LocalErrorMsg = "ไม่สามารถลบข้อมูลได้เนื่องจากข้อมูลนี้ยังถูกอ้างถึงอยู่จากส่วนอื่น"
   IsOK = False

   DeleteComboSubGroup = True
End Function


Public Function CopyMasterFromTo(Ug As CMasterFromTo, IsOK As Boolean, AutoCommit As Boolean, ErrorObj As clsErrorLog)
On Error GoTo ErrorHandler
Dim RName As String
Dim HasBegin As Boolean
Dim iCount As Long
Dim TempRs As ADODB.Recordset
Dim child_Rs As ADODB.Recordset
Dim childCount As Long
Dim MasterFromToMax As Long

   RName = "CopyMasterFromTo"
   CopyMasterFromTo = False
   HasBegin = False
   
   Set TempRs = New ADODB.Recordset
    Set child_Rs = New ADODB.Recordset

'   Ug.QueryFlag = 1
'   Ug.FROM_DATE = -1
'   Ug.TO_DATE = -1
   Call QueryMasterFromTo(Ug, TempRs, iCount, IsOK, glbErrorLog)

   If Not TempRs.EOF Then
      Call Ug.PopulateFromRS(1, TempRs)
   End If

      Dim Cm As CCommissionChart
      Set Cm = New CCommissionChart
      Cm.MASTER_FROMTO_ID = Ug.MASTER_FROMTO_ID
      Cm.VALID_FROM = -1
      Cm.VALID_TO = -1
      Cm.ORDER_TYPE = 1
      Call Cm.QueryData(4, child_Rs, childCount)         '  หาลูกของต้นฉบับทั้งหมด

      Ug.ShowMode = SHOW_ADD
      Ug.MASTER_FROMTO_ID = -1
      Ug.MASTER_FROMTO_NO = "***" & Ug.MASTER_FROMTO_NO
      Call AddEditMasterFromTo(Ug, IsOK, False, glbErrorLog)             ' เพิ่มตัวแม่
      MasterFromToMax = getMasterFromToMax(Ug, IsOK, ErrorObj)                          ' หาตัวแม่มากสุด
      
      While Not child_Rs.EOF          ' เพิ่มตัวลูก
            Call Cm.PopulateFromRS(4, child_Rs)
            Cm.MASTER_FROMTO_ID = MasterFromToMax
            Cm.AddEditMode = SHOW_ADD
            Cm.PARENT_ID = 0
            Call Cm.AddEditData
            child_Rs.MoveNext
      Wend

   If AutoCommit Then
      m_ADOConn.CommitTrans
   End If
   HasBegin = False
   
   If child_Rs.State = adStateOpen Then
      Call child_Rs.Close
   End If
   Set child_Rs = Nothing
   
    If TempRs.State = adStateOpen Then
      Call TempRs.Close
   End If
   Set TempRs = Nothing
   
   HasBegin = False
   IsOK = True

   CopyMasterFromTo = True
   Exit Function

ErrorHandler:
   If HasBegin Then
      If AutoCommit Then
         m_ADOConn.RollbackTrans
      End If
   End If
   
   ErrorObj.LocalErrorMsg = "Runtime error."
   ErrorObj.SystemErrorMsg = Err.DESCRIPTION
   ErrorObj.RoutineName = RName
   ErrorObj.ModuleName = m_ModuleName
   ErrorObj.ShowErrorLog (LOG_TO_FILE)

   IsOK = False
   CopyMasterFromTo = False
End Function
Public Function CopyMaster2FromTo(Ug As CMaster2FromTo, IsOK As Boolean, AutoCommit As Boolean, ErrorObj As clsErrorLog)
On Error GoTo ErrorHandler
Dim RName As String
Dim HasBegin As Boolean
Dim iCount As Long
Dim TempRs As ADODB.Recordset
Dim ItemCount As Long

   RName = "CopyMaster2FromTo"
   CopyMaster2FromTo = False
   HasBegin = False
   
   Set TempRs = New ADODB.Recordset

   Ug.QueryFlag = 1
   Call glbDaily.QueryMaster2FromTo(Ug, TempRs, iCount, IsOK, ErrorObj)
   
   If Not TempRs.EOF Then
      Call Ug.PopulateFromRS(1, TempRs)
   End If

   Ug.ShowMode = SHOW_ADD
   Ug.MASTER_FROMTO_ID = -1
   Ug.MASTER_FROMTO_NO = "***" & Ug.MASTER_FROMTO_NO
   
   Dim Det As CMasterFromToDetail
   For Each Det In Ug.Details
      Det.Flag = "A"
      Det.ShowMode = SHOW_ADD
   Next Det

   Dim Para As CCommissMasterPara
   For Each Para In Ug.CommissionExs
      Para.Flag = "A"
      Para.AddEditMode = SHOW_ADD
   Next Para

   If AutoCommit Then
      m_ADOConn.BeginTrans
   End If
   HasBegin = True
   
   Call AddEditMaster2FromTo(Ug, IsOK, False, ErrorObj)
   
   If AutoCommit Then
      m_ADOConn.CommitTrans
   End If
   HasBegin = False
   
   If TempRs.State = adStateOpen Then
      Call TempRs.Close
   End If
   Set TempRs = Nothing
   
   HasBegin = False
   IsOK = True

   CopyMaster2FromTo = True
   Exit Function

ErrorHandler:
   If HasBegin Then
      If AutoCommit Then
         m_ADOConn.RollbackTrans
      End If
   End If
   
   ErrorObj.LocalErrorMsg = "Runtime error."
   ErrorObj.SystemErrorMsg = Err.DESCRIPTION
   ErrorObj.RoutineName = RName
   ErrorObj.ModuleName = m_ModuleName
   ErrorObj.ShowErrorLog (LOG_TO_FILE)

   IsOK = False
   CopyMaster2FromTo = False
End Function

Public Function AddEditMaster2FromTo(Ua As CMaster2FromTo, IsOK As Boolean, AutoCommit As Boolean, ErrorObj As clsErrorLog)
On Error GoTo ErrorHandler
Dim RName As String
Dim HasBegin As Boolean
Dim i As Long
Dim MasFromTo2Max As Long

   RName = "AddEditMaster2FromTo"
   AddEditMaster2FromTo = False
   HasBegin = False

   If AutoCommit Then
      m_ADOConn.BeginTrans
      HasBegin = True
   End If
   
   Call Ua.AddEditData
   MasFromTo2Max = getMasFromTo2Max(Ua, IsOK, ErrorObj) ' หาตัวแม่มากสุด
   ' จากเดิมที่อ้างโดย Ua.MASTER_FROMTO_ID ต้องเปลี่ยนเป็นดึงค่าสุดท้ายที่ Add เพิ่มมา
   
   Dim Gr As CMasterFromToDetail
   For Each Gr In Ua.Details
      Gr.MASTER_FROMTO_ID = MasFromTo2Max
      If Gr.Flag = "A" Then
         Gr.ShowMode = SHOW_ADD
         Call Gr.AddEditData
      ElseIf Gr.Flag = "E" Then
         Gr.ShowMode = SHOW_EDIT
         Call Gr.AddEditData
      ElseIf Gr.Flag = "D" Then
         Call Gr.DeleteData
      End If

   Next Gr

   Dim Gx As CCommissMasterPara
   For Each Gx In Ua.CommissionExs
      Gx.MASTER_FROMTO_ID = MasFromTo2Max
      If Gx.Flag = "A" Then
         Gx.AddEditMode = SHOW_ADD
         Call Gx.AddEditData
      ElseIf Gx.Flag = "E" Then
         Gx.AddEditMode = SHOW_EDIT
         Call Gx.AddEditData
      ElseIf Gx.Flag = "D" Then
         Call Gx.DeleteData
      End If

   Next Gx
   
   If AutoCommit Then
      Call m_ADOConn.CommitTrans
   End If
   
   HasBegin = False
   IsOK = True

   AddEditMaster2FromTo = True
   Exit Function

ErrorHandler:
   If HasBegin Then
      If AutoCommit Then
         m_ADOConn.RollbackTrans
      End If
   End If

   ErrorObj.LocalErrorMsg = "Runtime error."
   ErrorObj.SystemErrorMsg = Err.DESCRIPTION
   ErrorObj.RoutineName = RName
   ErrorObj.ModuleName = m_ModuleName
   ErrorObj.ShowErrorLog (LOG_TO_FILE)

   IsOK = False
   AddEditMaster2FromTo = False
End Function
Public Function AddEditMasterIncentive(Ua As CIncentiveMasterPromote, IsOK As Boolean, AutoCommit As Boolean, ErrorObj As clsErrorLog)
On Error GoTo ErrorHandler
Dim RName As String
Dim HasBegin As Boolean
Dim i As Long
   RName = "AddEditMasterIncentive"
   AddEditMasterIncentive = False
   HasBegin = False

   If AutoCommit Then
      m_ADOConn.BeginTrans
      HasBegin = True
   End If
   
   Call Ua.AddEditData
   
   Dim Gr As CIncentivePromote
   For Each Gr In Ua.Details
      Gr.MASTER_INCENTIVE_ID = Ua.MASTER_INCENTIVE_ID
      If Gr.Flag = "A" Then
         Gr.ShowMode = SHOW_ADD
         Call Gr.AddEditData
      ElseIf Gr.Flag = "E" Then
         Gr.ShowMode = SHOW_EDIT
         Call Gr.AddEditData
      ElseIf Gr.Flag = "D" Then
         Call Gr.DeleteData
      End If

   Next Gr

'   Dim Gx As CIncentiveMasterPromote
'   For Each Gx In Ua.CommissionExs
'      Gx.MASTER_INCENTIVE_ID = Ua.MASTER_INCENTIVE_ID
'      If Gx.Flag = "A" Then
'         Gx.AddEditMode = SHOW_ADD
'         Call Gx.AddEditData
'      ElseIf Gx.Flag = "E" Then
'         Gx.AddEditMode = SHOW_EDIT
'         Call Gx.AddEditData
'      ElseIf Gx.Flag = "D" Then
'         Call Gx.DeleteData
'      End If
'
'   Next Gx
'
   If AutoCommit Then
      Call m_ADOConn.CommitTrans
   End If
   
   HasBegin = False
   IsOK = True

   AddEditMasterIncentive = True
   Exit Function

ErrorHandler:
   If HasBegin Then
      If AutoCommit Then
         m_ADOConn.RollbackTrans
      End If
   End If

   ErrorObj.LocalErrorMsg = "Runtime error."
   ErrorObj.SystemErrorMsg = Err.DESCRIPTION
   ErrorObj.RoutineName = RName
   ErrorObj.ModuleName = m_ModuleName
   ErrorObj.ShowErrorLog (LOG_TO_FILE)

   IsOK = False
   AddEditMasterIncentive = False
End Function

Public Function QueryMasterIncentive(Ua As CIncentiveMasterPromote, Rs As ADODB.Recordset, ItemCount As Long, IsOK As Boolean, ErrorObj As clsErrorLog) As Boolean
On Error GoTo ErrorHandler
Dim RName As String
Dim SQL1 As String
Dim SQL2 As String
Dim SelectStr As String
Dim FromStr As String
Dim WhereStr As String
Dim OrderStr As String
Dim OrderType As String
Dim NewStr As String
Dim SubLen As Long
Dim iCount As Long

   RName = "QueryMasterIncentive"
   QueryMasterIncentive = False

   IsOK = True

   Call Ua.QueryData(1, Rs, ItemCount)
   
   If Ua.QueryFlag = 1 Then
      Dim Gr As CIncentivePromote
      Set Gr = New CIncentivePromote
      Gr.MASTER_INCENTIVE_ID = Ua.MASTER_INCENTIVE_ID
      Gr.FROM_CMPL_DATE = -1
      Gr.TO_CMPL_DATE = -1
      Call Gr.QueryData(1, m_Rs1, iCount)
      Set Gr = Nothing

      Set Ua.Details = Nothing
      Set Ua.Details = New Collection

      While Not m_Rs1.EOF
         Set Gr = New CIncentivePromote
         Call Gr.PopulateFromRS(1, m_Rs1)

         Gr.Flag = "I"
         Call Ua.Details.Add(Gr)
         Set Gr = Nothing

         m_Rs1.MoveNext
      Wend


'      Dim Gx As CCommissMasterPara
'      Set Gx = New CCommissMasterPara
'      Gx.MASTER_INCENTIVE_ID = Ua.MASTER_INCENTIVE_ID
'      Call Gx.QueryData(1, m_Rs1, iCount)
'      Set Gx = Nothing
'
'      Set Ua.CommissionExs = Nothing
'      Set Ua.CommissionExs = New Collection
'
'      While Not m_Rs1.EOF
'         Set Gx = New CCommissMasterPara
'         Call Gx.PopulateFromRS(1, m_Rs1)
'
'         Gx.Flag = "I"
'         Call Ua.CommissionExs.Add(Gx)
'         Set Gx = Nothing
'
'         m_Rs1.MoveNext
'      Wend

   End If
   
   QueryMasterIncentive = True
   Exit Function

ErrorHandler:
   ErrorObj.SystemErrorMsg = Err.DESCRIPTION
   ErrorObj.RoutineName = RName
   ErrorObj.ModuleName = m_ModuleName
   ErrorObj.ShowErrorLog (LOG_FILE_MSGBOX)
   
   QueryMasterIncentive = False
End Function

'-----------------------------------------
Public Function DeleteMasterIncentive(Ua As CIncentiveMasterPromote, IsOK As Boolean, AutoCommit As Boolean, ErrorObj As clsErrorLog)
On Error GoTo ErrorHandler
Dim RName As String
Dim SQL As String
Dim HasBegin As Boolean
Dim ItemCount As Long

   RName = "DeleteMasterIncentive"
   DeleteMasterIncentive = False
   HasBegin = False

   If AutoCommit Then
      m_ADOConn.BeginTrans
   End If
   HasBegin = True
      
   Call Ua.DeleteData
      
   If AutoCommit Then
      Call m_ADOConn.CommitTrans
   End If
   HasBegin = False
   IsOK = True

   DeleteMasterIncentive = True
   Exit Function

ErrorHandler:
   If HasBegin Then
      If AutoCommit Then
         m_ADOConn.RollbackTrans
      End If
   End If

   ErrorObj.LocalErrorMsg = "Runtime error."
   ErrorObj.SystemErrorMsg = Err.DESCRIPTION
   ErrorObj.RoutineName = RName
   ErrorObj.ModuleName = m_ModuleName
   ErrorObj.ShowErrorLog (LOG_TO_FILE)

   ErrorObj.LocalErrorMsg = "ไม่สามารถลบข้อมูลได้เนื่องจากข้อมูลนี้ยังถูกอ้างถึงอยู่จากส่วนอื่น"
   IsOK = False

   DeleteMasterIncentive = True
End Function

Public Function DeleteMasterCommiss(UID As Long, IsOK As Boolean, AutoCommit As Boolean, ErrorObj As clsErrorLog)
On Error GoTo ErrorHandler
Dim RName As String
Dim SQL As String
Dim HasBegin As Boolean
Dim ItemCount As Long
Dim Ua As CCommissMasterPromote

   RName = "DeleteMasterCommiss"
   DeleteMasterCommiss = False
   HasBegin = False
   
   Set Ua = New CCommissMasterPromote
   
   If AutoCommit Then
      m_ADOConn.BeginTrans
   End If
   HasBegin = True
      
   Ua.MASTER_Commiss_ID = UID
   Call Ua.DeleteData
      
   If AutoCommit Then
      Call m_ADOConn.CommitTrans
   End If
   HasBegin = False
   IsOK = True

   DeleteMasterCommiss = True
   Exit Function

ErrorHandler:
   If HasBegin Then
      If AutoCommit Then
         m_ADOConn.RollbackTrans
      End If
   End If

   ErrorObj.LocalErrorMsg = "Runtime error."
   ErrorObj.SystemErrorMsg = Err.DESCRIPTION
   ErrorObj.RoutineName = RName
   ErrorObj.ModuleName = m_ModuleName
   ErrorObj.ShowErrorLog (LOG_TO_FILE)

   ErrorObj.LocalErrorMsg = "ไม่สามารถลบข้อมูลได้เนื่องจากข้อมูลนี้ยังถูกอ้างถึงอยู่จากส่วนอื่น"
   IsOK = False

   DeleteMasterCommiss = True
End Function

Public Function QueryMasterCommiss(Ua As CCommissMasterPromote, Rs As ADODB.Recordset, ItemCount As Long, IsOK As Boolean, ErrorObj As clsErrorLog) As Boolean
On Error GoTo ErrorHandler
Dim RName As String
Dim SQL1 As String
Dim SQL2 As String
Dim SelectStr As String
Dim FromStr As String
Dim WhereStr As String
Dim OrderStr As String
Dim OrderType As String
Dim NewStr As String
Dim SubLen As Long
Dim iCount As Long

   RName = "QueryMasterCommiss"
   QueryMasterCommiss = False

   IsOK = True

   Call Ua.QueryData(1, Rs, ItemCount)
   
   If Ua.QueryFlag = 1 Then
      Dim Gr As CComMasSubPromote
      
      Set Gr = New CComMasSubPromote
      Gr.MASTER_Commiss_ID = Ua.MASTER_Commiss_ID
      Gr.FROM_CMPL_DATE = -1
      Gr.TO_CMPL_DATE = -1
      Gr.CREDIT_TYP = "01"
      Call Gr.QueryData(1, m_Rs1, iCount)
      
       Set Gr = Nothing
      Set Ua.DetailsCom1 = Nothing
      Set Ua.DetailsCom1 = New Collection

      While Not m_Rs1.EOF
         Set Gr = New CComMasSubPromote
         Call Gr.PopulateFromRS(1, m_Rs1)                        ' นี่คือตัวเนื้อใน ต้อง Commission 4 เงื่อนไข

         Gr.Flag = "I"
         Call Ua.DetailsCom1.Add(Gr)
         Set Gr = Nothing

         m_Rs1.MoveNext
      Wend
      
      Set Gr = New CComMasSubPromote
      Gr.MASTER_Commiss_ID = Ua.MASTER_Commiss_ID
      Gr.FROM_CMPL_DATE = -1
      Gr.TO_CMPL_DATE = -1
      Gr.CREDIT_TYP = "02"
      Call Gr.QueryData(1, m_Rs2, iCount)
      
      Set Gr = Nothing
      Set Ua.DetailsCom2 = Nothing
      Set Ua.DetailsCom2 = New Collection

      While Not m_Rs2.EOF
         Set Gr = New CComMasSubPromote
         Call Gr.PopulateFromRS(1, m_Rs2)                        ' นี่คือตัวเนื้อใน ต้อง Commission 4 เงื่อนไข

        Gr.Flag = "I"
         Call Ua.DetailsCom2.Add(Gr)
        Set Gr = Nothing

         m_Rs2.MoveNext
      Wend

      Set Gr = New CComMasSubPromote
      Gr.MASTER_Commiss_ID = Ua.MASTER_Commiss_ID
      Gr.FROM_CMPL_DATE = -1
      Gr.TO_CMPL_DATE = -1
      Gr.CREDIT_TYP = "03"
      Call Gr.QueryData(1, m_Rs3, iCount)
      
      Set Gr = Nothing
      Set Ua.DetailsCom3 = Nothing
      Set Ua.DetailsCom3 = New Collection

      While Not m_Rs3.EOF
         Set Gr = New CComMasSubPromote
         Call Gr.PopulateFromRS(1, m_Rs3)                        ' นี่คือตัวเนื้อใน ต้อง Commission 4 เงื่อนไข

         Gr.Flag = "I"
         Call Ua.DetailsCom3.Add(Gr)
         Set Gr = Nothing

         m_Rs3.MoveNext
      Wend

   End If
    
   QueryMasterCommiss = True
   Exit Function

ErrorHandler:
   ErrorObj.SystemErrorMsg = Err.DESCRIPTION
   ErrorObj.RoutineName = RName
   ErrorObj.ModuleName = m_ModuleName
   ErrorObj.ShowErrorLog (LOG_FILE_MSGBOX)
   
   QueryMasterCommiss = False
End Function

Public Function AddEditMasterCommiss(Ua As CCommissMasterPromote, IsOK As Boolean, AutoCommit As Boolean, ErrorObj As clsErrorLog)
On Error GoTo ErrorHandler
Dim RName As String
Dim HasBegin As Boolean
Dim i As Long
   RName = "AddEditMasterCommiss"
   AddEditMasterCommiss = False
   HasBegin = False

   If AutoCommit Then
      m_ADOConn.BeginTrans
      HasBegin = True
   End If
   
   Call Ua.AddEditData
   
   Dim Gr As CComMasSubPromote
   For Each Gr In Ua.DetailsCom1

      Gr.MASTER_Commiss_ID = Ua.MASTER_Commiss_ID
      If Gr.Flag = "A" Then
         Gr.ShowMode = SHOW_ADD
         Call Gr.AddEditData
      ElseIf Gr.Flag = "E" Then
         Gr.ShowMode = SHOW_EDIT
         Call Gr.AddEditData
      ElseIf Gr.Flag = "D" Then
         Call Gr.DeleteData
      End If

   Next Gr

'   Dim Gx As CCommissMasterPromote
'   For Each Gx In Ua.CommissionExs
'      Gx.MASTER_Commiss_ID = Ua.MASTER_Commiss_ID
'      If Gx.Flag = "A" Then
'         Gx.AddEditMode = SHOW_ADD
'         Call Gx.AddEditData
'      ElseIf Gx.Flag = "E" Then
'         Gx.AddEditMode = SHOW_EDIT
'         Call Gx.AddEditData
'      ElseIf Gx.Flag = "D" Then
'         Call Gx.DeleteData
'      End If
'
'   Next Gx
'
   If AutoCommit Then
      Call m_ADOConn.CommitTrans
   End If
   
   HasBegin = False
   IsOK = True

   AddEditMasterCommiss = True
   Exit Function

ErrorHandler:
   If HasBegin Then
      If AutoCommit Then
         m_ADOConn.RollbackTrans
      End If
   End If

   ErrorObj.LocalErrorMsg = "Runtime error."
   ErrorObj.SystemErrorMsg = Err.DESCRIPTION
   ErrorObj.RoutineName = RName
   ErrorObj.ModuleName = m_ModuleName
   ErrorObj.ShowErrorLog (LOG_TO_FILE)

   IsOK = False
   AddEditMasterCommiss = False
End Function




Public Function DeleteMasterSubCommiss(UID As Long, IsOK As Boolean, AutoCommit As Boolean, ErrorObj As clsErrorLog, Optional Ind As Long = 1)
On Error GoTo ErrorHandler
Dim RName As String
Dim SQL As String
Dim HasBegin As Boolean
Dim ItemCount As Long
Dim Ua As CComMasSubPromote

   RName = "DeleteMasterSubCommiss"
   DeleteMasterSubCommiss = False
   HasBegin = False

   Set Ua = New CComMasSubPromote

   If AutoCommit Then
      m_ADOConn.BeginTrans
   End If
   HasBegin = True
      
   Ua.MASTER_COMMISS_SUB_PROMOTE_ID = UID
   Call Ua.DeleteData(Ind)
      
   If AutoCommit Then
      Call m_ADOConn.CommitTrans
   End If
   HasBegin = False
   IsOK = True

   DeleteMasterSubCommiss = True
   Exit Function

ErrorHandler:
   If HasBegin Then
      If AutoCommit Then
         m_ADOConn.RollbackTrans
      End If
   End If

   ErrorObj.LocalErrorMsg = "Runtime error."
   ErrorObj.SystemErrorMsg = Err.DESCRIPTION
   ErrorObj.RoutineName = RName
   ErrorObj.ModuleName = m_ModuleName
   ErrorObj.ShowErrorLog (LOG_TO_FILE)

   ErrorObj.LocalErrorMsg = "ไม่สามารถลบได้เนื่องจากข้อมูลถูกอ้างถึงโดยส่วนอื่น"
   IsOK = False

   DeleteMasterSubCommiss = True
End Function


Public Function QueryMasterSubCommiss(Ua As CComMasSubPromote, Rs As ADODB.Recordset, ItemCount As Long, IsOK As Boolean, ErrorObj As clsErrorLog) As Boolean
On Error GoTo ErrorHandler
Dim RName As String
Dim SQL1 As String
Dim SQL2 As String
Dim SelectStr As String
Dim FromStr As String
Dim WhereStr As String
Dim OrderStr As String
Dim OrderType As String
Dim NewStr As String
Dim SubLen As Long
Dim iCount As Long

   RName = "QueryMasterSubCommiss"
   QueryMasterSubCommiss = False

   IsOK = True
     Ua.FROM_CMPL_DATE = -1
      Ua.TO_CMPL_DATE = -1
   Call Ua.QueryData(1, Rs, ItemCount)
   
   If Ua.QueryFlag = 1 Then
      Dim Gr As CCommissPromote
      Set Gr = New CCommissPromote
      Gr.MASTER_COMMISS_SUB_PROMOTE_ID = Ua.MASTER_COMMISS_SUB_PROMOTE_ID
      Gr.FROM_CMPL_DATE = -1
      Gr.TO_CMPL_DATE = -1
      Gr.Commiss_TYP = "01"
      Call Gr.QueryData(1, m_Rs1, iCount)
      Set Gr = Nothing

      Set Ua.DetailsCom1 = Nothing
      Set Ua.DetailsCom1 = New Collection

      While Not m_Rs1.EOF
         Set Gr = New CCommissPromote
         Call Gr.PopulateFromRS(1, m_Rs1)                        ' นี่คือตัวเนื้อใน ต้อง Commission 4 เงื่อนไข

         Gr.Flag = "I"
         Call Ua.DetailsCom1.Add(Gr)
         Set Gr = Nothing
      m_Rs1.MoveNext
  Wend

      Set Gr = New CCommissPromote
      Gr.MASTER_COMMISS_SUB_PROMOTE_ID = Ua.MASTER_COMMISS_SUB_PROMOTE_ID
      Gr.FROM_CMPL_DATE = -1
      Gr.TO_CMPL_DATE = -1
      Gr.Commiss_TYP = "02"
      Call Gr.QueryData(1, m_Rs1, iCount)
      Set Gr = Nothing

      Set Ua.DetailsCom2 = Nothing
      Set Ua.DetailsCom2 = New Collection

      While Not m_Rs1.EOF
         Set Gr = New CCommissPromote
         Call Gr.PopulateFromRS(1, m_Rs1)                        ' นี่คือตัวเนื้อใน ต้อง Commission 4 เงื่อนไข

         Gr.Flag = "I"
         Call Ua.DetailsCom2.Add(Gr)
         Set Gr = Nothing
      m_Rs1.MoveNext
  Wend

      Set Gr = New CCommissPromote
      Gr.MASTER_COMMISS_SUB_PROMOTE_ID = Ua.MASTER_COMMISS_SUB_PROMOTE_ID
      Gr.FROM_CMPL_DATE = -1
      Gr.TO_CMPL_DATE = -1
      Gr.Commiss_TYP = "03"
      Call Gr.QueryData(1, m_Rs1, iCount)
      Set Gr = Nothing

      Set Ua.DetailsCom3 = Nothing
      Set Ua.DetailsCom3 = New Collection

      While Not m_Rs1.EOF
         Set Gr = New CCommissPromote
         Call Gr.PopulateFromRS(1, m_Rs1)                        ' นี่คือตัวเนื้อใน ต้อง Commission 4 เงื่อนไข

         Gr.Flag = "I"
         Call Ua.DetailsCom3.Add(Gr)
         Set Gr = Nothing
      m_Rs1.MoveNext
  Wend
  

   End If
   
   QueryMasterSubCommiss = True
   Exit Function

ErrorHandler:
   ErrorObj.SystemErrorMsg = Err.DESCRIPTION
   ErrorObj.RoutineName = RName
   ErrorObj.ModuleName = m_ModuleName
   ErrorObj.ShowErrorLog (LOG_FILE_MSGBOX)
   
   QueryMasterSubCommiss = False
End Function

Public Function AddEditMasterSubCommiss(Ua As CComMasSubPromote, IsOK As Boolean, AutoCommit As Boolean, ErrorObj As clsErrorLog)
On Error GoTo ErrorHandler
Dim RName As String
Dim HasBegin As Boolean
Dim i As Long
   RName = "AddEditMasterSubCommiss"
   AddEditMasterSubCommiss = False
   HasBegin = False

   If AutoCommit Then
      m_ADOConn.BeginTrans
      HasBegin = True
   End If
   
   Call Ua.AddEditData
   
   Dim Gr As CCommissPromote
   For Each Gr In Ua.DetailsCom1
      Gr.MASTER_COMMISS_SUB_PROMOTE_ID = Ua.MASTER_COMMISS_SUB_PROMOTE_ID
      If Gr.Flag = "A" Then
         Gr.ShowMode = SHOW_ADD
         Call Gr.AddEditData
      ElseIf Gr.Flag = "E" Then
         Gr.ShowMode = SHOW_EDIT
         Call Gr.AddEditData
      ElseIf Gr.Flag = "D" Then
         Call Gr.DeleteData
      End If

   Next Gr

'   Dim Gx As CCommissMasterPromote
'   For Each Gx In Ua.CommissionExs
'      Gx.MASTER_Commiss_ID = Ua.MASTER_Commiss_ID
'      If Gx.Flag = "A" Then
'         Gx.AddEditMode = SHOW_ADD
'         Call Gx.AddEditData
'      ElseIf Gx.Flag = "E" Then
'         Gx.AddEditMode = SHOW_EDIT
'         Call Gx.AddEditData
'      ElseIf Gx.Flag = "D" Then
'         Call Gx.DeleteData
'      End If
'
'   Next Gx
'
   If AutoCommit Then
      Call m_ADOConn.CommitTrans
   End If
   
   HasBegin = False
   IsOK = True

   AddEditMasterSubCommiss = True
   Exit Function

ErrorHandler:
   If HasBegin Then
      If AutoCommit Then
         m_ADOConn.RollbackTrans
      End If
   End If

   ErrorObj.LocalErrorMsg = "Runtime error."
   ErrorObj.SystemErrorMsg = Err.DESCRIPTION
   ErrorObj.RoutineName = RName
   ErrorObj.ModuleName = m_ModuleName
   ErrorObj.ShowErrorLog (LOG_TO_FILE)

   IsOK = False
   AddEditMasterSubCommiss = False
End Function

Public Function DeleteMasterComDonStk(Ua As CComDonStkMaster, IsOK As Boolean, AutoCommit As Boolean, ErrorObj As clsErrorLog)
On Error GoTo ErrorHandler
Dim RName As String
Dim SQL As String
Dim HasBegin As Boolean
Dim ItemCount As Long

   RName = "DeleteMasterComDonStk"
   DeleteMasterComDonStk = False
   HasBegin = False

   If AutoCommit Then
      m_ADOConn.BeginTrans
   End If
   HasBegin = True
      
   Call Ua.DeleteData
      
   If AutoCommit Then
      Call m_ADOConn.CommitTrans
   End If
   HasBegin = False
   IsOK = True

   DeleteMasterComDonStk = True
   Exit Function

ErrorHandler:
   If HasBegin Then
      If AutoCommit Then
         m_ADOConn.RollbackTrans
      End If
   End If

   ErrorObj.LocalErrorMsg = "Runtime error."
   ErrorObj.SystemErrorMsg = Err.DESCRIPTION
   ErrorObj.RoutineName = RName
   ErrorObj.ModuleName = m_ModuleName
   ErrorObj.ShowErrorLog (LOG_TO_FILE)

   ErrorObj.LocalErrorMsg = "ไม่สามารถลบข้อมูลได้เนื่องจากข้อมูลนี้ยังถูกอ้างถึงอยู่จากส่วนอื่น"
   IsOK = False

   DeleteMasterComDonStk = True
End Function

Public Function AddEditMasterComDonStk(Ua As CComDonStkMaster, IsOK As Boolean, AutoCommit As Boolean, ErrorObj As clsErrorLog)
On Error GoTo ErrorHandler
Dim RName As String
Dim HasBegin As Boolean
Dim i As Long
   RName = "AddEditMasterComDonStk"
   AddEditMasterComDonStk = False
   HasBegin = False

   If AutoCommit Then
      m_ADOConn.BeginTrans
      HasBegin = True
   End If
   
   Call Ua.AddEditData
   
   Dim Gr As CComDonStk
   For Each Gr In Ua.Details
      Gr.MASTER_COMDONSTK_ID = Ua.MASTER_COMDONSTK_ID
      If Gr.Flag = "A" Then
         Gr.ShowMode = SHOW_ADD
         Call Gr.AddEditData
      ElseIf Gr.Flag = "E" Then
         Gr.ShowMode = SHOW_EDIT
         Call Gr.AddEditData
      ElseIf Gr.Flag = "D" Then
         Call Gr.DeleteData
      End If

   Next Gr

'   Dim Gx As CComDonStkMaster
'   For Each Gx In Ua.CommissionExs
'      Gx.MASTER_COMDONSTK_ID = Ua.MASTER_COMDONSTK_ID
'      If Gx.Flag = "A" Then
'         Gx.AddEditMode = SHOW_ADD
'         Call Gx.AddEditData
'      ElseIf Gx.Flag = "E" Then
'         Gx.AddEditMode = SHOW_EDIT
'         Call Gx.AddEditData
'      ElseIf Gx.Flag = "D" Then
'         Call Gx.DeleteData
'      End If
'
'   Next Gx
'
   If AutoCommit Then
      Call m_ADOConn.CommitTrans
   End If
   
   HasBegin = False
   IsOK = True

   AddEditMasterComDonStk = True
   Exit Function

ErrorHandler:
   If HasBegin Then
      If AutoCommit Then
         m_ADOConn.RollbackTrans
      End If
   End If

   ErrorObj.LocalErrorMsg = "Runtime error."
   ErrorObj.SystemErrorMsg = Err.DESCRIPTION
   ErrorObj.RoutineName = RName
   ErrorObj.ModuleName = m_ModuleName
   ErrorObj.ShowErrorLog (LOG_TO_FILE)

   IsOK = False
   AddEditMasterComDonStk = False
End Function

Public Function QueryMasterComDonStk(Ua As CComDonStkMaster, Rs As ADODB.Recordset, ItemCount As Long, IsOK As Boolean, ErrorObj As clsErrorLog) As Boolean
On Error GoTo ErrorHandler
Dim RName As String
Dim SQL1 As String
Dim SQL2 As String
Dim SelectStr As String
Dim FromStr As String
Dim WhereStr As String
Dim OrderStr As String
Dim OrderType As String
Dim NewStr As String
Dim SubLen As Long
Dim iCount As Long

   RName = "QueryMasterComDonStk"
   QueryMasterComDonStk = False

   IsOK = True

   Call Ua.QueryData(1, Rs, ItemCount)
   
   If Ua.QueryFlag = 1 Then
      Dim Gr As CComDonStk
      Set Gr = New CComDonStk
      Gr.MASTER_COMDONSTK_ID = Ua.MASTER_COMDONSTK_ID
      Gr.VALID_FROM = -1
      Gr.VALID_TO = -1
      Call Gr.QueryData(1, m_Rs1, iCount)
      Set Gr = Nothing

      Set Ua.Details = Nothing
      Set Ua.Details = New Collection

      While Not m_Rs1.EOF
         Set Gr = New CComDonStk
         Call Gr.PopulateFromRS(1, m_Rs1)

         Gr.Flag = "I"
         Call Ua.Details.Add(Gr)
         Set Gr = Nothing

         m_Rs1.MoveNext
      Wend


'      Dim Gx As CCommissMasterPara
'      Set Gx = New CCommissMasterPara
'      Gx.MASTER_COMDONSTK_ID = Ua.MASTER_COMDONSTK_ID
'      Call Gx.QueryData(1, m_Rs1, iCount)
'      Set Gx = Nothing
'
'      Set Ua.CommissionExs = Nothing
'      Set Ua.CommissionExs = New Collection
'
'      While Not m_Rs1.EOF
'         Set Gx = New CCommissMasterPara
'         Call Gx.PopulateFromRS(1, m_Rs1)
'
'         Gx.Flag = "I"
'         Call Ua.CommissionExs.Add(Gx)
'         Set Gx = Nothing
'
'         m_Rs1.MoveNext
'      Wend

   End If
   
   QueryMasterComDonStk = True
   Exit Function

ErrorHandler:
   ErrorObj.SystemErrorMsg = Err.DESCRIPTION
   ErrorObj.RoutineName = RName
   ErrorObj.ModuleName = m_ModuleName
   ErrorObj.ShowErrorLog (LOG_FILE_MSGBOX)
   
   QueryMasterComDonStk = False
End Function

Public Function DeletePromoteType1(UID As Long, IsOK As Boolean, AutoCommit As Boolean, ErrorObj As clsErrorLog)
On Error GoTo ErrorHandler
Dim RName As String
Dim SQL As String
Dim HasBegin As Boolean
Dim ItemCount As Long
Dim B As CCommissPromote

   RName = "DeletePromoteType1"
   DeletePromoteType1 = False
   HasBegin = False

   Set B = New CCommissPromote

   If AutoCommit Then
      m_ADOConn.BeginTrans
   End If
   HasBegin = True

   B.Commiss_PROMOTE_ID = UID
   Call B.DeleteData(1)

   Set B = Nothing

   If AutoCommit Then
      Call m_ADOConn.CommitTrans
   End If
   HasBegin = False
   IsOK = True

   DeletePromoteType1 = True
   Exit Function

ErrorHandler:
   If HasBegin Then
      If AutoCommit Then
         m_ADOConn.RollbackTrans
      End If
   End If

   ErrorObj.LocalErrorMsg = "Runtime error."
   ErrorObj.SystemErrorMsg = Err.DESCRIPTION
   ErrorObj.RoutineName = RName
   ErrorObj.ModuleName = m_ModuleName
   ErrorObj.LocalErrorMsg = "ไม่สามารถลบข้อมูลได้เนื่องจากข้อมูลนี้ยังถูกอ้างถึงอยู่จากส่วนอื่น"
   ErrorObj.ShowErrorLog (LOG_TO_FILE)

   IsOK = False

   DeletePromoteType1 = True
End Function

Public Function AddEditAreaYear(Ug As CAreaYear, IsOK As Boolean, AutoCommit As Boolean, ErrorObj As clsErrorLog)
On Error GoTo ErrorHandler
Dim RName As String
Dim HasBegin As Boolean

   RName = "AddEditAreaYear"
   AddEditAreaYear = False
   HasBegin = False

   If AutoCommit Then
      m_ADOConn.BeginTrans
   End If
   HasBegin = True

   Call Ug.AddEditData

   If AutoCommit Then
      Call m_ADOConn.CommitTrans
   End If
   HasBegin = False
   IsOK = True

   AddEditAreaYear = True
   Exit Function

ErrorHandler:
   If HasBegin Then
      If AutoCommit Then
         m_ADOConn.RollbackTrans
      End If
   End If

   ErrorObj.LocalErrorMsg = "Runtime error."
   ErrorObj.SystemErrorMsg = Err.DESCRIPTION
   ErrorObj.RoutineName = RName
   ErrorObj.ModuleName = m_ModuleName
   ErrorObj.ShowErrorLog (LOG_TO_FILE)

   IsOK = False
   AddEditAreaYear = False
End Function
'===
Public Function QueryAreaYear(Ua As CAreaYear, Rs As ADODB.Recordset, ItemCount As Long, IsOK As Boolean, ErrorObj As clsErrorLog) As Boolean
On Error GoTo ErrorHandler
Dim RName As String
Dim SQL1 As String
Dim SQL2 As String
Dim SelectStr As String
Dim FromStr As String
Dim WhereStr As String
Dim OrderStr As String
Dim OrderType As String
Dim iCount As Long


   RName = "QueryAreaYear"
QueryAreaYear = False

   IsOK = True
   Call Ua.QueryData(1, Rs, ItemCount)

   QueryAreaYear = True
   Exit Function

ErrorHandler:
   ErrorObj.SystemErrorMsg = Err.DESCRIPTION
   ErrorObj.RoutineName = RName
   ErrorObj.ModuleName = m_ModuleName
   ErrorObj.ShowErrorLog (LOG_FILE_MSGBOX)

   QueryAreaYear = False
End Function

Public Function DeleteAreaYear(UID As Long, IsOK As Boolean, AutoCommit As Boolean, ErrorObj As clsErrorLog) As Boolean
On Error GoTo ErrorHandler
Dim RName As String
Dim SQL As String
Dim HasBegin As Boolean
Dim ItemCount As Long
Dim B As CAreaYear

   RName = "DeleteAreaYear"
   DeleteAreaYear = False
   HasBegin = False

   Set B = New CAreaYear

   If AutoCommit Then
      Call m_ADOConn.BeginTrans
   End If
   HasBegin = True

   B.YEAR_ID = UID
   Call B.DeleteData

   Set B = Nothing

   If AutoCommit Then
      Call m_ADOConn.CommitTrans
   End If
   HasBegin = False
   IsOK = True

   DeleteAreaYear = True
   Exit Function

ErrorHandler:
   If HasBegin Then
      If AutoCommit Then
         m_ADOConn.RollbackTrans
      End If
   End If

   ErrorObj.LocalErrorMsg = "Runtime error."
   ErrorObj.SystemErrorMsg = Err.DESCRIPTION
   ErrorObj.RoutineName = RName
   ErrorObj.ModuleName = m_ModuleName
   ErrorObj.LocalErrorMsg = "ไม่สามารถเปิดข้อมูลได้ เนื่องจากข้อมูลถูกอ้างถึงโดยส่วนอื่น"
   ErrorObj.ShowErrorLog (LOG_TO_FILE)

   IsOK = False

   DeleteAreaYear = True
End Function

Public Function DeleteGoodsGroup(UID As Long, IsOK As Boolean, AutoCommit As Boolean, ErrorObj As clsErrorLog) As Boolean
On Error GoTo ErrorHandler
Dim RName As String
Dim SQL As String
Dim HasBegin As Boolean
Dim ItemCount As Long
Dim B As CGoodsGroup

   RName = "DeleteGoodsGroup"
   DeleteGoodsGroup = False
   HasBegin = False

   Set B = New CGoodsGroup

   If AutoCommit Then
      Call m_ADOConn.BeginTrans
   End If
   HasBegin = True

   B.GOODS_GROUP_ID = UID
   Call B.DeleteData

   Set B = Nothing

   If AutoCommit Then
      Call m_ADOConn.CommitTrans
   End If
   HasBegin = False
   IsOK = True

   DeleteGoodsGroup = True
   Exit Function

ErrorHandler:
   If HasBegin Then
      If AutoCommit Then
         m_ADOConn.RollbackTrans
      End If
   End If

   ErrorObj.LocalErrorMsg = "Runtime error."
   ErrorObj.SystemErrorMsg = Err.DESCRIPTION
   ErrorObj.RoutineName = RName
   ErrorObj.ModuleName = m_ModuleName
   ErrorObj.ShowErrorLog (LOG_TO_FILE)

   IsOK = False

   DeleteGoodsGroup = True
End Function

Public Function QueryGoodsGroup(Ua As CGoodsGroup, Rs As ADODB.Recordset, ItemCount As Long, IsOK As Boolean, ErrorObj As clsErrorLog) As Boolean
On Error GoTo ErrorHandler
Dim RName As String
Dim OrderStr As String
Dim OrderType As String
Dim iCount As Long


   RName = "QueryGoodsGroup"
   QueryGoodsGroup = False

   IsOK = True
   Call Ua.QueryData(1, Rs, ItemCount)

  QueryGoodsGroup = True
   Exit Function

ErrorHandler:
   ErrorObj.SystemErrorMsg = Err.DESCRIPTION
   ErrorObj.RoutineName = RName
   ErrorObj.ModuleName = m_ModuleName
   ErrorObj.ShowErrorLog (LOG_FILE_MSGBOX)

  QueryGoodsGroup = False
End Function

Public Function AddEditGoodsGroup(Ug As CGoodsGroup, IsOK As Boolean, AutoCommit As Boolean, ErrorObj As clsErrorLog)
On Error GoTo ErrorHandler
Dim RName As String
Dim HasBegin As Boolean

   RName = "AddEditGoodsGroup"
   AddEditGoodsGroup = False
   HasBegin = False

   If AutoCommit Then
      m_ADOConn.BeginTrans
   End If
   HasBegin = True

   Call Ug.AddEditData

   If AutoCommit Then
      Call m_ADOConn.CommitTrans
   End If
   HasBegin = False
   IsOK = True

   AddEditGoodsGroup = True
   Exit Function

ErrorHandler:
   If HasBegin Then
      If AutoCommit Then
         m_ADOConn.RollbackTrans
      End If
   End If
   
   ErrorObj.LocalErrorMsg = "Runtime error."
   ErrorObj.SystemErrorMsg = Err.DESCRIPTION
   ErrorObj.RoutineName = RName
   ErrorObj.ModuleName = m_ModuleName
   ErrorObj.ShowErrorLog (LOG_TO_FILE)

   IsOK = False
   AddEditGoodsGroup = False
End Function

Public Function QueryComRecord(Ua As CComRecord, Rs As ADODB.Recordset, ItemCount As Long, IsOK As Boolean, ErrorObj As clsErrorLog) As Boolean
On Error GoTo ErrorHandler
Dim RName As String
Dim OrderStr As String
Dim OrderType As String
Dim iCount As Long


   RName = "QueryComRecord"
   QueryComRecord = False

   IsOK = True
   Call Ua.QueryData(2, Rs, ItemCount)

  QueryComRecord = True
   Exit Function

ErrorHandler:
   ErrorObj.SystemErrorMsg = Err.DESCRIPTION
   ErrorObj.RoutineName = RName
   ErrorObj.ModuleName = m_ModuleName
   ErrorObj.ShowErrorLog (LOG_FILE_MSGBOX)

  QueryComRecord = False
End Function

Public Function AddEditComRecord(Ug As CComRecord, IsOK As Boolean, AutoCommit As Boolean, ErrorObj As clsErrorLog)
On Error GoTo ErrorHandler
Dim RName As String
Dim HasBegin As Boolean

   RName = "AddEditComRecord"
   AddEditComRecord = False
   HasBegin = False

   If AutoCommit Then
      m_ADOConn.BeginTrans
   End If
   HasBegin = True

   Call Ug.AddEditData

   If AutoCommit Then
      Call m_ADOConn.CommitTrans
   End If
   HasBegin = False
   IsOK = True

   AddEditComRecord = True
   Exit Function

ErrorHandler:
   If HasBegin Then
      If AutoCommit Then
         m_ADOConn.RollbackTrans
      End If
   End If
   
   ErrorObj.LocalErrorMsg = "Runtime error."
   ErrorObj.SystemErrorMsg = Err.DESCRIPTION
   ErrorObj.RoutineName = RName
   ErrorObj.ModuleName = m_ModuleName
   ErrorObj.ShowErrorLog (LOG_TO_FILE)

   IsOK = False
   AddEditComRecord = False
End Function

Public Function AddEditGoodsMaster(Ug As CGoodsMaster, IsOK As Boolean, AutoCommit As Boolean, ErrorObj As clsErrorLog)
On Error GoTo ErrorHandler
Dim RName As String
Dim HasBegin As Boolean

   RName = "AddEditGoodsMaster"
   AddEditGoodsMaster = False
   HasBegin = False

   If AutoCommit Then
      m_ADOConn.BeginTrans
   End If
   HasBegin = True

   Call Ug.AddEditData

   If AutoCommit Then
      Call m_ADOConn.CommitTrans
   End If
   HasBegin = False
   IsOK = True

   AddEditGoodsMaster = True
   Exit Function

ErrorHandler:
   If HasBegin Then
      If AutoCommit Then
         m_ADOConn.RollbackTrans
      End If
   End If

   ErrorObj.LocalErrorMsg = "Runtime error."
   ErrorObj.SystemErrorMsg = Err.DESCRIPTION
   ErrorObj.RoutineName = RName
   ErrorObj.ModuleName = m_ModuleName
   ErrorObj.ShowErrorLog (LOG_TO_FILE)

   IsOK = False
   AddEditGoodsMaster = False
End Function

Public Function QueryGoodsMaster(Ua As CGoodsMaster, Rs As ADODB.Recordset, ItemCount As Long, IsOK As Boolean, ErrorObj As clsErrorLog) As Boolean
On Error GoTo ErrorHandler
Dim RName As String
Dim SQL1 As String
Dim SQL2 As String
Dim SelectStr As String
Dim FromStr As String
Dim WhereStr As String
Dim OrderStr As String
Dim OrderType As String
Dim iCount As Long


   RName = "QueryGoodsMaster"
   QueryGoodsMaster = False

   IsOK = True
   Call Ua.QueryData(1, Rs, ItemCount)

   QueryGoodsMaster = True
   Exit Function

ErrorHandler:
   ErrorObj.SystemErrorMsg = Err.DESCRIPTION
   ErrorObj.RoutineName = RName
   ErrorObj.ModuleName = m_ModuleName
   ErrorObj.ShowErrorLog (LOG_FILE_MSGBOX)

   QueryGoodsMaster = False
End Function

Public Function DeleteGoodsMaster(UID As Long, IsOK As Boolean, AutoCommit As Boolean, ErrorObj As clsErrorLog) As Boolean
On Error GoTo ErrorHandler
Dim RName As String
Dim SQL As String
Dim HasBegin As Boolean
Dim ItemCount As Long
Dim B As CGoodsMaster

   RName = "DeleteGoodsMaster"
   DeleteGoodsMaster = False
   HasBegin = False

   Set B = New CGoodsMaster
   
   If AutoCommit Then
      Call m_ADOConn.BeginTrans
   End If
   HasBegin = True

   B.GOODS_MASTER_ID = UID
   Call B.DeleteData

   Set B = Nothing

   If AutoCommit Then
      Call m_ADOConn.CommitTrans
   End If
   HasBegin = False
   IsOK = True

   DeleteGoodsMaster = True
   Exit Function

ErrorHandler:
   If HasBegin Then
      If AutoCommit Then
         m_ADOConn.RollbackTrans
      End If
   End If

   ErrorObj.LocalErrorMsg = "Runtime error."
   ErrorObj.SystemErrorMsg = Err.DESCRIPTION
   ErrorObj.RoutineName = RName
   ErrorObj.ModuleName = m_ModuleName
   ErrorObj.LocalErrorMsg = "ไม่สามารถเปิดข้อมูลได้ เนื่องจากข้อมูลถูกอ้างถึงโดยส่วนอื่น"
   ErrorObj.ShowErrorLog (LOG_TO_FILE)

   IsOK = False

   DeleteGoodsMaster = True
End Function


Public Function QueryGoodsEP(Ua As CStmas, Rs As ADODB.Recordset, ItemCount As Long, IsOK As Boolean, ErrorObj As clsErrorLog) As Boolean
On Error GoTo ErrorHandler
Dim RName As String
Dim SQL1 As String
Dim SQL2 As String
Dim SelectStr As String
Dim FromStr As String
Dim WhereStr As String
Dim OrderStr As String
Dim OrderType As String
Dim NewStr As String
Dim SubLen As Long
Dim iCount As Long

   RName = "QueryGoodsEP"
   QueryGoodsEP = False

   IsOK = True
   Call Ua.QueryData(1, Rs, ItemCount)

   QueryGoodsEP = True
   Exit Function

ErrorHandler:
   ErrorObj.SystemErrorMsg = Err.DESCRIPTION
   ErrorObj.RoutineName = RName
   ErrorObj.ModuleName = m_ModuleName
   ErrorObj.ShowErrorLog (LOG_FILE_MSGBOX)

   QueryGoodsEP = False
End Function

Public Function QueryGoodsDetail(Ua As CGoodsDetail, Rs As ADODB.Recordset, ItemCount As Long, IsOK As Boolean, ErrorObj As clsErrorLog) As Boolean
On Error GoTo ErrorHandler
Dim RName As String
Dim OrderStr As String
Dim OrderType As String

   RName = "QueryGoodsDetail"
   QueryGoodsDetail = False

   IsOK = True
   Call Ua.QueryData(1, Rs, ItemCount)         ' CInventorydoc

   QueryGoodsDetail = True
   Exit Function

ErrorHandler:
   ErrorObj.SystemErrorMsg = Err.DESCRIPTION
   ErrorObj.RoutineName = RName
   ErrorObj.ModuleName = m_ModuleName
   ErrorObj.ShowErrorLog (LOG_FILE_MSGBOX)

QueryGoodsDetail = False
End Function


Public Function DeleteGoodsDetail(Ug As CGoodsDetail, IsOK As Boolean, AutoCommit As Boolean, ErrorObj As clsErrorLog)
On Error GoTo ErrorHandler
Dim RName As String
Dim SQL As String
Dim HasBegin As Boolean

   RName = "DeleteGoodsDetail"
   DeleteGoodsDetail = False
   HasBegin = False

   If AutoCommit Then
      m_ADOConn.BeginTrans
   End If
   HasBegin = True

   Call Ug.DeleteData

   If AutoCommit Then
      Call m_ADOConn.CommitTrans
   End If
   HasBegin = False
   IsOK = True

   DeleteGoodsDetail = True
   Exit Function

ErrorHandler:
   If HasBegin Then
      If AutoCommit Then
         m_ADOConn.RollbackTrans
      End If
   End If

   ErrorObj.LocalErrorMsg = "Runtime error."
   ErrorObj.SystemErrorMsg = Err.DESCRIPTION
   ErrorObj.RoutineName = RName
   ErrorObj.ModuleName = m_ModuleName
   ErrorObj.LocalErrorMsg = "ไม่สามารถลบข้อมูลได้เนื่องจากข้อมูลนี้ยังถูกอ้างถึงอยู่จากส่วนอื่น"
   ErrorObj.ShowErrorLog (LOG_TO_FILE)

   IsOK = False
   DeleteGoodsDetail = False
End Function
'
Public Function AddEditGoodsDetail(Ug As CGoodsDetail, IsOK As Boolean, AutoCommit As Boolean, ErrorObj As clsErrorLog)
On Error GoTo ErrorHandler
Dim RName As String
Dim HasBegin As Boolean

   RName = "AddEditGoodsDetail"
   AddEditGoodsDetail = False
   HasBegin = False

   If AutoCommit Then
      m_ADOConn.BeginTrans
   End If
   HasBegin = True

   Call Ug.AddEditData

   If AutoCommit Then
      Call m_ADOConn.CommitTrans
   End If
   HasBegin = False
   IsOK = True

   AddEditGoodsDetail = True
   Exit Function

ErrorHandler:
   If HasBegin Then
      If AutoCommit Then
         m_ADOConn.RollbackTrans
      End If
   End If

   ErrorObj.LocalErrorMsg = "Runtime error."
   ErrorObj.SystemErrorMsg = Err.DESCRIPTION
   ErrorObj.RoutineName = RName
   ErrorObj.ModuleName = m_ModuleName
   ErrorObj.ShowErrorLog (LOG_TO_FILE)

   IsOK = False
   AddEditGoodsDetail = False
End Function
Public Function QueryProvinceMap(Ua As CProvinceMap, Rs As ADODB.Recordset, ItemCount As Long, IsOK As Boolean, ErrorObj As clsErrorLog, Optional db As Long = 1) As Boolean
On Error GoTo ErrorHandler
Dim RName As String
Dim SQL1 As String
Dim SQL2 As String
Dim SelectStr As String
Dim FromStr As String
Dim WhereStr As String
Dim OrderStr As String
Dim OrderType As String
Dim NewStr As String
Dim SubLen As Long
Dim iCount As Long
   
   RName = "QueryProvinceMap"
   QueryProvinceMap = False
   
   IsOK = True
   Call Ua.QueryData(Rs, ItemCount)
   
   QueryProvinceMap = True
   Exit Function

ErrorHandler:
   ErrorObj.SystemErrorMsg = Err.DESCRIPTION
   ErrorObj.RoutineName = RName
   ErrorObj.ModuleName = m_ModuleName
   ErrorObj.ShowErrorLog (LOG_FILE_MSGBOX)

   QueryProvinceMap = False
End Function
Public Function AddEditProvinceMap(Ug As CProvinceMap, IsOK As Boolean, AutoCommit As Boolean, ErrorObj As clsErrorLog)
On Error GoTo ProvinceMap
Dim RName As String
Dim HasBegin As Boolean

   RName = "AddEditProvinceMap"
   AddEditProvinceMap = False
   HasBegin = False

   If AutoCommit Then
      m_ADOConn.BeginTrans
   End If
   HasBegin = True

   Call Ug.AddEditData

   If AutoCommit Then
      Call m_ADOConn.CommitTrans
   End If
   HasBegin = False
   IsOK = True
   
   AddEditProvinceMap = True
   Exit Function

ProvinceMap:
   If HasBegin Then
      If AutoCommit Then
         m_ADOConn.RollbackTrans
      End If
   End If

   ErrorObj.LocalErrorMsg = "Runtime error."
   ErrorObj.SystemErrorMsg = Err.DESCRIPTION
   ErrorObj.RoutineName = RName
   ErrorObj.ModuleName = m_ModuleName
   ErrorObj.ShowErrorLog (LOG_TO_FILE)

   IsOK = False
   AddEditProvinceMap = False
End Function
Public Function DeleteProvinceMap(Ug As CProvinceMap, IsOK As Boolean, AutoCommit As Boolean, ErrorObj As clsErrorLog)
On Error GoTo ErrorHandler
Dim RName As String
Dim SQL As String
Dim HasBegin As Boolean

   RName = "DeleteProvinceMap"
   DeleteProvinceMap = False
   HasBegin = False
   
   If AutoCommit Then
      m_ADOConn.BeginTrans
   End If
   HasBegin = True
   
   Call Ug.DeleteData
   
   If AutoCommit Then
      Call m_ADOConn.CommitTrans
   End If
   HasBegin = False
   IsOK = True

   DeleteProvinceMap = True
   Exit Function
   
ErrorHandler:
   If HasBegin Then
      If AutoCommit Then
         m_ADOConn.RollbackTrans
      End If
   End If
   
   ErrorObj.LocalErrorMsg = "Runtime error."
   ErrorObj.SystemErrorMsg = Err.DESCRIPTION
   ErrorObj.RoutineName = RName
   ErrorObj.ModuleName = m_ModuleName
   ErrorObj.LocalErrorMsg = "ไม่สามารถลบข้อมูลได้เนื่องจากข้อมูลนี้ยังถูกอ้างถึงอยู่จากส่วนอื่น"
   ErrorObj.ShowErrorLog (LOG_TO_FILE)
   
   IsOK = False
   
   DeleteProvinceMap = False
End Function

