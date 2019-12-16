Attribute VB_Name = "modLoadData"
Option Explicit

Public Sub InitAccountNo1(C As ComboBox)
   C.Clear
   
   C.AddItem ("")
   C.ItemData(0) = 0
   
   C.AddItem ("260-7001")
   C.ItemData(1) = 1
   
   C.AddItem ("212-2200")
   C.ItemData(2) = 2

   C.AddItem ("211-1040")
   C.ItemData(3) = 3

   C.AddItem ("211-1150")
   C.ItemData(4) = 4
End Sub

Public Sub InitComDocType(C As ComboBox)
   C.Clear
   
   C.AddItem ("")
   C.ItemData(0) = 0
   
   C.AddItem (MapText("Commission ขาย"))
   C.ItemData(1) = 1
   
   C.AddItem (MapText("Commission เก็บเงิน"))
   C.ItemData(2) = 2
   
   C.AddItem (MapText("Incentive"))
   C.ItemData(3) = 3
   
   C.AddItem (MapText("ทุกเอกสาร"))
   C.ItemData(4) = 4
End Sub

Public Sub InitJournalType(C As ComboBox)
   C.Clear

   C.AddItem ("")
   C.ItemData(0) = 0
   
   C.AddItem ("กำหนดเอง")
   C.ItemData(1) = 1
   
   C.AddItem ("จากระบบ")
   C.ItemData(2) = 2
End Sub

Public Sub InitDocumentOrderBy(C As ComboBox)
   C.Clear

   C.AddItem ("")
   C.ItemData(0) = 0
   
   C.AddItem ("เลขที่เอกสาร")
   C.ItemData(1) = 1
   
   C.AddItem ("วันที่เอกสาร")
   C.ItemData(2) = 2
End Sub

Public Sub InitAssetOrderBy(C As ComboBox)
   C.Clear

   C.AddItem ("")
   C.ItemData(0) = 0
   
   C.AddItem ("รหัสสินทรัพย์")
   C.ItemData(1) = 1
End Sub

Public Sub InitCheckType(C As ComboBox)
   C.Clear

   C.AddItem ("")
   C.ItemData(0) = 0
   
   C.AddItem ("MP-UOB")
   C.ItemData(1) = 1

   C.AddItem ("MP-KTB")
   C.ItemData(2) = 2

   C.AddItem ("QMC-BKK")
   C.ItemData(3) = 3

   C.AddItem ("MGP-SCB")
   C.ItemData(4) = 4

   C.AddItem ("DTS-TFB")
   C.ItemData(5) = 5

   C.AddItem ("MH-TFB")
   C.ItemData(6) = 6
End Sub


Public Sub InitIntervalType(C As ComboBox)
   C.Clear

   C.AddItem ("")
   C.ItemData(0) = 0
   
   C.AddItem ("แบบ 1")
   C.ItemData(1) = 1
   
   C.AddItem ("แบบ 2")
   C.ItemData(2) = 2

   C.AddItem ("แบบ 3")
   C.ItemData(3) = 3
   
   C.AddItem ("แบบ 4")
   C.ItemData(4) = 4
   
End Sub

Public Sub InitCustomerOrderBy(C As ComboBox)
   C.Clear

   C.AddItem ("")
   C.ItemData(0) = 0
   
   C.AddItem ("รหัสลูกหนี้")
   C.ItemData(1) = 1
   
   C.AddItem ("ชื่อลูกหนี้")
   C.ItemData(2) = 2
End Sub
Public Sub InitCustomer2OrderBy(C As ComboBox)
   C.Clear

   C.AddItem ("")
   C.ItemData(0) = 0
   
   C.AddItem ("รหัสลูกค้า")
   C.ItemData(1) = 1
   
   C.AddItem ("ชื่อลูกค้า")
   C.ItemData(2) = 2
End Sub

Public Sub InitBankCreditOrderBy(C As ComboBox)
   C.Clear

   C.AddItem ("")
   C.ItemData(0) = 0
   
   C.AddItem ("วันที่")
   C.ItemData(1) = 1
   
   C.AddItem ("ชื่อธนาคาร")
   C.ItemData(2) = 2
   
   C.AddItem ("ลูกหนี้")
   C.ItemData(3) = 3
   
End Sub
Public Sub InitBankFeeType(C As ComboBox)
   C.Clear

   C.AddItem ("")
   C.ItemData(0) = 0
   
   C.AddItem ("บาท/การแพ็ค 1 ครั้ง")
   C.ItemData(1) = 1
   
   C.AddItem ("บาท/บิล")
   C.ItemData(2) = 2
   
End Sub
Public Sub InitTicketOrderBy(C As ComboBox)
   C.Clear

   C.AddItem ("")
   C.ItemData(0) = 0
   
   C.AddItem ("เลขที่ตั๋ว")
   C.ItemData(1) = 1
   
   C.AddItem ("วันที่")
   C.ItemData(2) = 2
   
   C.AddItem ("ลูกหนี้")
   C.ItemData(3) = 3
End Sub
Public Sub InitSupplierOrderBy(C As ComboBox)
   C.Clear

   C.AddItem ("")
   C.ItemData(0) = 0
   
   C.AddItem ("รหัสผู้จำหน่าย")
   C.ItemData(1) = 1
   
   C.AddItem ("ชื่อผู้จำหน่าย")
   C.ItemData(2) = 2
End Sub
Public Sub InitSaleOrderBy(C As ComboBox)
   C.Clear

   C.AddItem ("")
   C.ItemData(0) = 0
   
   C.AddItem ("รหัสพนักงานขาย")
   C.ItemData(1) = 1
   
   C.AddItem ("ชื่อพนักงานขาย")
   C.ItemData(2) = 2
End Sub

Public Sub InitAdjustType(C As ComboBox)
   C.Clear
   
   C.AddItem ("")
   C.ItemData(0) = 0
   
   C.AddItem (MapText("ปรับลด"))
   C.ItemData(1) = 1
   
   C.AddItem (MapText("ปรับเพิ่ม"))
   C.ItemData(2) = 2
End Sub

Public Sub LoadDBPath(C As ComboBox, Optional Cl As Collection = Nothing)
   C.Clear
   
   C.AddItem ("")
   C.ItemData(0) = 0
   
   C.AddItem ("D:\Express\Secure")
   C.ItemData(1) = 1
End Sub

Public Sub LoadSupplierType(C As ComboBox, Optional Cl As Collection = Nothing, Optional SupplierType As String = "")
On Error GoTo ErrorHandler
Dim D As CIsTab
Dim ItemCount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CIsTab
Dim i As Long
Dim SUPTYP As String

   If Val(SupplierType) <= 0 Then
      SUPTYP = ""
   Else
      SUPTYP = SupplierType
   End If
   
   Set D = New CIsTab
   Set Rs = New ADODB.Recordset

   
   D.TABTYP = "46"
   D.TYPCOD = SUPTYP
   Call D.QueryData(Rs, ItemCount)
   
   If Not (C Is Nothing) Then
      C.Clear
      i = 0
      C.AddItem ("")
   End If

   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   
'   If Not (Cl Is Nothing) Then
'      Set TempData = New CIsTab
'      TempData.TYPCOD = "00"
'      Call Cl.Add(TempData, TempData.TYPCOD)
'      Set TempData = Nothing
'   End If
   
   While Not Rs.EOF
      i = i + 1
      Set TempData = New CIsTab
      Call TempData.PopulateFromRS(1, Rs)

      If Not (C Is Nothing) Then
         C.AddItem (TempData.TYPDES & "(" & TempData.TYPCOD & ")")
'         C.ItemData(i) = Val(TempData.TYPCOD)
      End If
     'Set Cl = New Collection
      If Not (Cl Is Nothing) Then
         Call Cl.Add(TempData, TempData.TYPCOD)
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend

   Set Rs = Nothing
   Set D = Nothing
   Exit Sub

ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub

Public Sub LoadCustomerType(C As ComboBox, Optional Cl As Collection = Nothing, Optional CustomerType As String = "")
On Error GoTo ErrorHandler
Dim D As CIsTab
Dim ItemCount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CIsTab
Dim i As Long
Dim CUSTYP As String
   
   If Val(CustomerType) <= 0 Then
      CUSTYP = ""
   Else
      CUSTYP = CustomerType
   End If
   
   Set D = New CIsTab
   Set Rs = New ADODB.Recordset

   
   D.TABTYP = "45"
   D.TYPCOD = CUSTYP
   Call D.QueryData(Rs, ItemCount)

   If Not (C Is Nothing) Then
      C.Clear
      i = 0
      C.AddItem ("")
      C.ItemData(i) = 0
   End If

   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   
   If Not (Cl Is Nothing) Then
'      Set TempData = New CIsTab
'      TempData.TYPCOD = "AA"
'      Call Cl.Add(TempData, TempData.TYPCOD)
'      Set TempData = Nothing
   End If
   
   While Not Rs.EOF
      i = i + 1
      Set TempData = New CIsTab
      Call TempData.PopulateFromRS(1, Rs)

      If Not (C Is Nothing) Then
         C.AddItem (TempData.TYPDES)
         C.ItemData(i) = i
      End If

      If Not (Cl Is Nothing) Then
         Call Cl.Add(TempData, TempData.TYPCOD)
      End If

      Set TempData = Nothing
      Rs.MoveNext
   Wend

   Set Rs = Nothing
   Set D = Nothing
   Exit Sub

ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub

Public Sub LoadGLAcc(C As ComboBox, Optional Cl As Collection = Nothing, Optional ID As Long = -1)
On Error GoTo ErrorHandler
Dim D As CGLAcc
Dim ItemCount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CGLAcc
Dim TempData2  As CGLAcc
Dim i As Long
Dim j As Long

   Set D = New CGLAcc
   Set Rs = New ADODB.Recordset
   
   Call D.QueryData(-1, Rs, ItemCount)
   
   If Not (C Is Nothing) Then
      C.Clear
      i = 0
      C.AddItem ("")
   End If

   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   
   While Not Rs.EOF
      Set TempData = New CGLAcc
      Call TempData.PopulateFromRS(1, Rs)
   
      If Not (C Is Nothing) Then
         i = i + 1
         If Len(Trim(TempData.ACCNUM)) > 0 Then
            C.AddItem (TempData.ACCNAM)
            C.ItemData(i) = i
         End If
      End If
      
      If Not (Cl Is Nothing) Then
         If Len(Trim(TempData.ACCNUM)) > 0 Then
            j = j + 1
            Call Cl.Add(TempData, Trim(Str(j)))
         End If
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   Set Rs = Nothing
   Set D = Nothing
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub

Public Sub LoadCompany(C As ComboBox, Optional Cl As Collection = Nothing, Optional Database2 As Boolean = False, Optional Database3 As Boolean = False)
On Error GoTo ErrorHandler
Dim D As CSCComp
Dim ItemCount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CSCComp
Dim i As Long

   Set D = New CSCComp
   Set Rs = New ADODB.Recordset
   
   D.db2 = Database2
   D.db3 = Database3
   Call D.QueryData(1, Rs, ItemCount)

   If Not (C Is Nothing) Then
      C.Clear
      i = 0
      C.AddItem ("")
   End If

   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   While Not Rs.EOF
      i = i + 1
      Set TempData = New CSCComp
      Call TempData.PopulateFromRS(1, Rs)

      If Not (C Is Nothing) Then
         C.AddItem (TempData.COMPNAM)
      End If

      If Not (Cl Is Nothing) Then
         Call Cl.Add(TempData)
      End If

      Set TempData = Nothing
      Rs.MoveNext
   Wend

   Set Rs = Nothing
   Set D = Nothing
   Exit Sub

ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub
Public Sub LoadDueDateInterval1Bank(C As ComboBox, Optional Cl As Collection)
Dim MM As CMaxMin
   '===
   Set MM = New CMaxMin
   MM.MIN = -999999
   MM.MAX = 180
   Call Cl.Add(MM)
   Set MM = Nothing
   
   Set MM = New CMaxMin
   MM.MIN = 180
   MM.MAX = 99999999
   Call Cl.Add(MM)
   Set MM = Nothing
End Sub

Public Sub LoadDueDateInterval1(C As ComboBox, Optional Cl As Collection)
Dim MM As CMaxMin
   '===
   Set MM = New CMaxMin
   MM.MIN = -999999
   MM.MAX = -60
   Call Cl.Add(MM)
   Set MM = Nothing
   
   Set MM = New CMaxMin
   MM.MIN = -60
   MM.MAX = -30
   Call Cl.Add(MM)
   Set MM = Nothing

   Set MM = New CMaxMin
   MM.MIN = -30
   MM.MAX = 0
   Call Cl.Add(MM)
   Set MM = Nothing

   '===
   Set MM = New CMaxMin
   MM.MIN = 0
   MM.MAX = 15
   Call Cl.Add(MM)
   Set MM = Nothing

   Set MM = New CMaxMin
   MM.MIN = 15
   MM.MAX = 30
   Call Cl.Add(MM)
   Set MM = Nothing

   Set MM = New CMaxMin
   MM.MIN = 30
   MM.MAX = 60
   Call Cl.Add(MM)
   Set MM = Nothing

   Set MM = New CMaxMin
   MM.MIN = 60
   MM.MAX = 9999999
   Call Cl.Add(MM)
   Set MM = Nothing
End Sub

Public Sub LoadDueDateInterval2(C As ComboBox, Optional Cl As Collection)
Dim MM As CMaxMin
   '===
   Set MM = New CMaxMin
   MM.MIN = -999999
   MM.MAX = -60
   Call Cl.Add(MM)
   Set MM = Nothing
   
   Set MM = New CMaxMin
   MM.MIN = -60
   MM.MAX = -30
   Call Cl.Add(MM)
   Set MM = Nothing

   Set MM = New CMaxMin
   MM.MIN = -30
   MM.MAX = 0
   Call Cl.Add(MM)
   Set MM = Nothing

   '===
   Set MM = New CMaxMin
   MM.MIN = 0
   MM.MAX = 90
   Call Cl.Add(MM)
   Set MM = Nothing

   Set MM = New CMaxMin
   MM.MIN = 90
   MM.MAX = 180
   Call Cl.Add(MM)
   Set MM = Nothing

   Set MM = New CMaxMin
   MM.MIN = 180
   MM.MAX = 365
   Call Cl.Add(MM)
   Set MM = Nothing

   Set MM = New CMaxMin
   MM.MIN = 365
   MM.MAX = 9999999
   Call Cl.Add(MM)
   Set MM = Nothing
End Sub
Public Sub LoadDueDateInterval3(C As ComboBox, Optional Cl As Collection)
Dim MM As CMaxMin
   '===
   Set MM = New CMaxMin
   MM.MIN = -999999
   MM.MAX = -60
   Call Cl.Add(MM)
   Set MM = Nothing
   
   Set MM = New CMaxMin
   MM.MIN = -60
   MM.MAX = -30
   Call Cl.Add(MM)
   Set MM = Nothing

   Set MM = New CMaxMin
   MM.MIN = -30
   MM.MAX = 0
   Call Cl.Add(MM)
   Set MM = Nothing

   '===
   Set MM = New CMaxMin
   MM.MIN = 0
   MM.MAX = 30
   Call Cl.Add(MM)
   Set MM = Nothing

   Set MM = New CMaxMin
   MM.MIN = 30
   MM.MAX = 60
   Call Cl.Add(MM)
   Set MM = Nothing

   Set MM = New CMaxMin
   MM.MIN = 60
   MM.MAX = 90
   Call Cl.Add(MM)
   Set MM = Nothing

   Set MM = New CMaxMin
   MM.MIN = 90
   MM.MAX = 9999999
   Call Cl.Add(MM)
   Set MM = Nothing
End Sub

Public Sub LoadDueDateInterval5(C As ComboBox, Optional Cl As Collection)
Dim MM As CMaxMin
   '===
   Set MM = New CMaxMin
   MM.MIN = -999999
   MM.MAX = -120
   Call Cl.Add(MM)
   Set MM = Nothing
   
   Set MM = New CMaxMin
   MM.MIN = -120
   MM.MAX = -90
   Call Cl.Add(MM)
   Set MM = Nothing
   
   Set MM = New CMaxMin
   MM.MIN = -90
   MM.MAX = -60
   Call Cl.Add(MM)
   Set MM = Nothing
   
   Set MM = New CMaxMin
   MM.MIN = -60
   MM.MAX = -30
   Call Cl.Add(MM)
   Set MM = Nothing

   Set MM = New CMaxMin
   MM.MIN = -30
   MM.MAX = 0
   Call Cl.Add(MM)
   Set MM = Nothing
   
   '===
   Set MM = New CMaxMin
   MM.MIN = 0
   MM.MAX = 30
   Call Cl.Add(MM)
   Set MM = Nothing

   Set MM = New CMaxMin
   MM.MIN = 30
   MM.MAX = 60
   Call Cl.Add(MM)
   Set MM = Nothing

   Set MM = New CMaxMin
   MM.MIN = 60
   MM.MAX = 90
   Call Cl.Add(MM)
   Set MM = Nothing

   Set MM = New CMaxMin
   MM.MIN = 90
   MM.MAX = 9999999
   Call Cl.Add(MM)
   Set MM = Nothing
End Sub
Public Sub LoadDueDateInterval6(C As ComboBox, Optional Cl As Collection)
Dim MM As CMaxMin
   '===
   Set MM = New CMaxMin
   MM.MIN = -999999
   MM.MAX = 0
   Call Cl.Add(MM)
   Set MM = Nothing
   
   '===
   Set MM = New CMaxMin
   MM.MIN = 0
   MM.MAX = 9999999
   Call Cl.Add(MM)
   Set MM = Nothing

End Sub
Public Sub LoadDueDateInterval4(C As ComboBox, Optional Cl As Collection)
Dim MM As CMaxMin
   '===
   Set MM = New CMaxMin
   MM.MIN = 0
   MM.MAX = 15
   Call Cl.Add(MM)
   Set MM = Nothing

   Set MM = New CMaxMin
   MM.MIN = 15
   MM.MAX = 30
   Call Cl.Add(MM)
   Set MM = Nothing

   Set MM = New CMaxMin
   MM.MIN = 30
   MM.MAX = 60
   Call Cl.Add(MM)
   Set MM = Nothing

   Set MM = New CMaxMin
   MM.MIN = 60
   MM.MAX = 90
   Call Cl.Add(MM)
   Set MM = Nothing

   Set MM = New CMaxMin
   MM.MIN = 90
   MM.MAX = 120
   Call Cl.Add(MM)
   Set MM = Nothing
   
   Set MM = New CMaxMin
   MM.MIN = 120
   MM.MAX = 150
   Call Cl.Add(MM)
   Set MM = Nothing
   
   Set MM = New CMaxMin
   MM.MIN = 150
   MM.MAX = 180
   Call Cl.Add(MM)
   Set MM = Nothing
   
   Set MM = New CMaxMin
   MM.MIN = 180
   MM.MAX = 210
   Call Cl.Add(MM)
   
   Set MM = New CMaxMin
   MM.MIN = 210
   MM.MAX = 999999
   Call Cl.Add(MM)
   Set MM = Nothing
   Set MM = Nothing
   
End Sub
Public Sub LoadSaleDateInterval1(C As ComboBox, Optional Cl As Collection)
Dim MM As CMaxMin
   '===
   Set MM = New CMaxMin
   MM.MIN = -999999
   MM.MAX = 60
   Call Cl.Add(MM)
   Set MM = Nothing
   
   Set MM = New CMaxMin
   MM.MIN = 60
   MM.MAX = 120
   Call Cl.Add(MM)
   Set MM = Nothing

   Set MM = New CMaxMin
   MM.MIN = 120
   MM.MAX = 180
   Call Cl.Add(MM)
   Set MM = Nothing

   Set MM = New CMaxMin
   MM.MIN = 180
   MM.MAX = 9999999
   Call Cl.Add(MM)
   Set MM = Nothing
End Sub
Public Sub LoadSaleDateInterval2(C As ComboBox, Optional Cl As Collection)
Dim MM As CMaxMin
   '===
   Set MM = New CMaxMin
   MM.MIN = -999999
   MM.MAX = 60
   Call Cl.Add(MM)
   Set MM = Nothing
   
   Set MM = New CMaxMin
   MM.MIN = 60
   MM.MAX = 120
   Call Cl.Add(MM)
   Set MM = Nothing

   Set MM = New CMaxMin
   MM.MIN = 120
   MM.MAX = 9999
   Call Cl.Add(MM)
   Set MM = Nothing

   Set MM = New CMaxMin
   MM.MIN = 9999
   MM.MAX = 9999999
   Call Cl.Add(MM)
   Set MM = Nothing
End Sub
Public Sub LoadPaidAmountByBill(C As ComboBox, Optional Cl As Collection = Nothing, Optional FromDate As Date, Optional ToDate As Date)
On Error GoTo ErrorHandler
Dim D As CAPRcIt
Dim ItemCount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CAPRcIt
Dim i As Long
Dim Sum As Double
   
   Set D = New CAPRcIt
   Set Rs = New ADODB.Recordset
   
   D.FROM_PAY_DATE = FromDate
   D.TO_PAY_DATE = ToDate
   D.OrderBy = 1
   Call D.QueryData(2, Rs, ItemCount)
   
   If Not (C Is Nothing) Then
      C.Clear
      i = 0
      C.AddItem ("")
   End If

   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   While Not Rs.EOF
      i = i + 1
      Set TempData = New CAPRcIt
      Call TempData.PopulateFromRS(2, Rs)
   
      If Not (C Is Nothing) Then
      End If
      
      Sum = Sum + TempData.PAYAMT
      
      If Not (Cl Is Nothing) Then
         Call Cl.Add(TempData, TempData.DOCNUM)
      End If
      
      '''debug.print (TempData.DOCNUM & "        " & TempData.PAYAMT)
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   Set Rs = Nothing
   Set D = Nothing
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub
Public Sub LoadReceiveAmountByBill(C As ComboBox, Optional Cl As Collection = Nothing, Optional FromDate As Date, Optional ToDate As Date, Optional db As Long = 1)
On Error GoTo ErrorHandler
Dim D As CARRcIt
Dim ItemCount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CARRcIt
Dim i As Long

   Set D = New CARRcIt
   Set Rs = New ADODB.Recordset
   
   D.FROM_PAY_DATE = FromDate
   D.TO_PAY_DATE = ToDate
   D.OrderBy = 1
'   D.FROM_CUSTOMER_CODE = "12-009"
'   D.TO_CUSTOMER_CODE = "12-009"
   Call D.QueryData(2, Rs, ItemCount, db)
   
   If Not (C Is Nothing) Then
      C.Clear
      i = 0
      C.AddItem ("")
   End If
Dim tempSum As Double
   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   While Not Rs.EOF
      i = i + 1
      Set TempData = New CARRcIt
      Call TempData.PopulateFromRS(2, Rs)
   
      If Not (C Is Nothing) Then
      End If
      
      tempSum = tempSum + TempData.RCVAMT
      
      If Not (Cl Is Nothing) Then
         Call Cl.Add(TempData, TempData.DOCNUM)
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   Debug.Print tempSum
   
   Set Rs = Nothing
   Set D = Nothing
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub
Public Sub LoadChequeFuture(C As ComboBox, Optional Cl As Collection = Nothing, Optional FromChequeDate As Date, Optional ToChequeDate As Date, Optional FromCusCode As String, Optional ToCusCode As String)
On Error GoTo ErrorHandler
Dim D As CBkTrn
Dim ItemCount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CBkTrn
Dim i As Long

   Set D = New CBkTrn
   Set Rs = New ADODB.Recordset
   
   D.FROM_CUSTOMER_CODE = FromCusCode
   D.TO_CUSTOMER_CODE = ToCusCode
   D.FROM_GETDAT = FromChequeDate
   D.TO_GETDAT = ToChequeDate
   
   D.OrderBy = 1
   Call D.QueryData(6, Rs, ItemCount)
   
   If Not (C Is Nothing) Then
      C.Clear
      i = 0
      C.AddItem ("")
   End If

   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   
   While Not Rs.EOF
      i = i + 1
      Set TempData = New CBkTrn
      Call TempData.PopulateFromRS(6, Rs)
   
      If Not (C Is Nothing) Then
      End If
      
      If Not (Cl Is Nothing) Then
         Call Cl.Add(TempData)
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   Set Rs = Nothing
   Set D = Nothing
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub

Public Sub LoadAnalyzeCustomer(C As ComboBox, Optional Cl As Collection = Nothing)
On Error GoTo ErrorHandler
Dim D As CAnalyzeCustomer
Dim ItemCount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CAnalyzeCustomer
Dim i As Long

   Set D = New CAnalyzeCustomer
   Set Rs = New ADODB.Recordset

   Call D.QueryData(Rs, ItemCount)
   
   If Not (C Is Nothing) Then
      C.Clear
      i = 0
      C.AddItem ("")
   End If

   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   While Not Rs.EOF
      i = i + 1
      Set TempData = New CAnalyzeCustomer
      Call TempData.PopulateFromRS(1, Rs)
   
      If Not (C Is Nothing) Then
      End If
      
      If Not (Cl Is Nothing) Then
         Call Cl.Add(TempData, Trim(TempData.INVOICE))
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   Set Rs = Nothing
   Set D = Nothing
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub
Public Sub LoadAPAmountBySup(C As ComboBox, Optional Cl As Collection = Nothing, Optional FromDate As Date, Optional ToDate As Date, Optional RecTypeSet As String, Optional DOCNUM As String)
On Error GoTo ErrorHandler
Dim D As CApTrn
Dim ItemCount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CApTrn
Dim i As Long

   Set D = New CApTrn
   Set Rs = New ADODB.Recordset

   D.FROM_DOC_DATE = FromDate
   D.TO_DOC_DATE = ToDate
   D.OrderBy = 1
   D.RecTypeSet = RecTypeSet
'                        D.DOCNUM = "RR50030046"
'                        D.RecTypeSet = "('3', '4', '5')"
   D.DOCNUM = DOCNUM
'                        D.SUPCOD = "ค-106"
   Call D.QueryData(2, Rs, ItemCount)
   
   If Not (C Is Nothing) Then
      C.Clear
      i = 0
      C.AddItem ("")
   End If
   
   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   While Not Rs.EOF
      i = i + 1
      Set TempData = New CApTrn
      Call TempData.PopulateFromRS(2, Rs)
      
      If TempData.SUPCOD = "ส-0032" Then
         Debug.Print
      End If
      
      If Not (C Is Nothing) Then
      End If
      
      If Not (Cl Is Nothing) Then
         Call Cl.Add(TempData, TempData.SUPCOD)
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   Set Rs = Nothing
   Set D = Nothing
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub

Public Sub LoadFaDprIt(C As ComboBox, Optional Cl As Collection = Nothing, Optional FromDate As Date, Optional ToDate As Date)
On Error GoTo ErrorHandler
Dim D As CFadprIt
Dim ItemCount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CFadprIt
Dim i As Long

   Set D = New CFadprIt
   Set Rs = New ADODB.Recordset

'   D.FROM_DOC_DATE = FromDate
'   D.TO_DOC_DATE = ToDate
   D.OrderBy = 1
   Call D.QueryData(1, Rs, ItemCount)
   
   If Not (C Is Nothing) Then
      C.Clear
      i = 0
      C.AddItem ("")
   End If

   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   While Not Rs.EOF
      i = i + 1
      Set TempData = New CFadprIt
      Call TempData.PopulateFromRS(1, Rs)
   
      If Not (C Is Nothing) Then
      End If
      
      If Not (Cl Is Nothing) Then
'''debug.print TempData.FASCOD & "-" & DateToStringInt(TempData.DOCDAT)
         Call Cl.Add(TempData, TempData.FASCOD & "-" & DateToStringInt(TempData.DOCDAT))
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   Set Rs = Nothing
   Set D = Nothing
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub

Public Sub LoadARAmountByCust(C As ComboBox, Optional Cl As Collection = Nothing, Optional FromDate As Date, Optional ToDate As Date, Optional DocumentNo As String)
On Error GoTo ErrorHandler
Dim D As CARTrn
Dim ItemCount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CARTrn
Dim i As Long
   
   Set D = New CARTrn
   Set Rs = New ADODB.Recordset
   
   D.FROM_DOC_DATE = FromDate
   D.TO_DOC_DATE = ToDate
   D.OrderBy = 1
   D.DOCNUM = DocumentNo
   Call D.QueryData(2, Rs, ItemCount)
   
   If Not (C Is Nothing) Then
      C.Clear
      i = 0
      C.AddItem ("")
   End If

   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   While Not Rs.EOF
      i = i + 1
      Set TempData = New CARTrn
      Call TempData.PopulateFromRS(2, Rs)
   
      If Not (C Is Nothing) Then
      End If
      
      If Not (Cl Is Nothing) Then
         Call Cl.Add(TempData, TempData.CUSCOD)
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   Set Rs = Nothing
   Set D = Nothing
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub
Public Sub LoadARAmountByCust2(C As ComboBox, Optional Cl As Collection = Nothing, Optional FromDate As Date, Optional ToDate As Date, Optional RecTypeSet As String)
On Error GoTo ErrorHandler
Dim D As CARTrn
Dim ItemCount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CARTrn
Dim i As Long

   Set D = New CARTrn
   Set Rs = New ADODB.Recordset

   D.FROM_DOC_DATE = FromDate
   D.TO_DOC_DATE = ToDate
   D.OrderBy = 1
   D.RecTypeSet = RecTypeSet
   Call D.QueryData(5, Rs, ItemCount)
   
   If Not (C Is Nothing) Then
      C.Clear
      i = 0
      C.AddItem ("")
   End If
   
   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   While Not Rs.EOF
      i = i + 1
      Set TempData = New CARTrn
      Call TempData.PopulateFromRS(5, Rs)
      
      If Not (C Is Nothing) Then
      End If
      
      If Not (Cl Is Nothing) Then
         Call Cl.Add(TempData, TempData.CUSCOD)
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   Set Rs = Nothing
   Set D = Nothing
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub
Public Sub LoadARAmountBySale(C As ComboBox, Optional Cl As Collection = Nothing, Optional FromDate As Date, Optional ToDate As Date, Optional DocumentNo As String)
On Error GoTo ErrorHandler
Dim D As CARTrn
Dim ItemCount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CARTrn
Dim i As Long

   Set D = New CARTrn
   Set Rs = New ADODB.Recordset

   D.FROM_DOC_DATE = FromDate
   D.TO_DOC_DATE = ToDate
   D.OrderBy = 1
   D.DOCNUM = DocumentNo
   Call D.QueryData(4, Rs, ItemCount)
   
   If Not (C Is Nothing) Then
      C.Clear
      i = 0
      C.AddItem ("")
   End If

   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   While Not Rs.EOF
      i = i + 1
      Set TempData = New CARTrn
      Call TempData.PopulateFromRS(4, Rs)
   
      If Not (C Is Nothing) Then
      End If
      
      If Not (Cl Is Nothing) Then
         Call Cl.Add(TempData, TempData.SLMCOD)
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   Set Rs = Nothing
   Set D = Nothing
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub
Public Sub LoadPaidAmountBySup(C As ComboBox, Optional Cl As Collection = Nothing, Optional FromDate As Date = -1, Optional ToDate As Date = -1, Optional FromDocDate As Date = -1, Optional ToDocDate As Date = -1, Optional RecTypeSet As String)
On Error GoTo ErrorHandler
Dim D As CAPRcIt
Dim ItemCount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CAPRcIt
Dim i As Long
   
   Set D = New CAPRcIt
   Set Rs = New ADODB.Recordset
   
   D.FROM_PAY_DATE = FromDate
   D.TO_PAY_DATE = ToDate
   D.FROM_DOC_DATE = FromDocDate
   D.TO_DOC_DATE = ToDocDate
   D.OrderBy = 1
   D.RecTypeSet = RecTypeSet
                           'D.DOCNUM = "RR50030046"
   Call D.QueryData(3, Rs, ItemCount)
   
   If Not (C Is Nothing) Then
      C.Clear
      i = 0
      C.AddItem ("")
   End If

   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   While Not Rs.EOF
      i = i + 1
      Set TempData = New CAPRcIt
      Call TempData.PopulateFromRS(3, Rs)
   
      If Not (C Is Nothing) Then
      End If
      
      If Not (Cl Is Nothing) Then
         Call Cl.Add(TempData, TempData.SUPCOD)
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   Set Rs = Nothing
   Set D = Nothing
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub

Public Sub LoadPaidAmountByCust(C As ComboBox, Optional Cl As Collection = Nothing, Optional FromDate As Date, Optional ToDate As Date)
On Error GoTo ErrorHandler
Dim D As CARRcIt
Dim ItemCount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CARRcIt
Dim i As Long

   Set D = New CARRcIt
   Set Rs = New ADODB.Recordset
   
   D.FROM_PAY_DATE = FromDate
   D.TO_PAY_DATE = ToDate
   D.OrderBy = 1
   Call D.QueryData(3, Rs, ItemCount)
   
   If Not (C Is Nothing) Then
      C.Clear
      i = 0
      C.AddItem ("")
   End If

   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   While Not Rs.EOF
      i = i + 1
      Set TempData = New CARRcIt
      Call TempData.PopulateFromRS(3, Rs)
   
      If Not (C Is Nothing) Then
      End If
      
      If Not (Cl Is Nothing) Then
         If Len(TempData.CUSCOD) > 0 Then
            Call Cl.Add(TempData, TempData.CUSCOD)
         Else
            ''debug.print
         End If
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   Set Rs = Nothing
   Set D = Nothing
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub
Public Sub LoadPaidAmountByCust2(C As ComboBox, Optional Cl As Collection = Nothing, Optional FromDate As Date, Optional ToDate As Date, Optional FromDocDate As Date = -1, Optional ToDocDate As Date = -1, Optional RecTypeSet As String)
On Error GoTo ErrorHandler
Dim D As CARRcIt
Dim ItemCount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CARRcIt
Dim i As Long

   Set D = New CARRcIt
   Set Rs = New ADODB.Recordset
   
   D.FROM_PAY_DATE = FromDate
   D.TO_PAY_DATE = ToDate
   D.FROM_DOC_DATE = FromDocDate
   D.TO_DOC_DATE = ToDocDate
   D.OrderBy = 1
   D.RecTypeSet = RecTypeSet
   Call D.QueryData(5, Rs, ItemCount)
   
   If Not (C Is Nothing) Then
      C.Clear
      i = 0
      C.AddItem ("")
   End If

   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   While Not Rs.EOF
      i = i + 1
      Set TempData = New CARRcIt
      Call TempData.PopulateFromRS(5, Rs)
   
      If Not (C Is Nothing) Then
      End If
      
      If Not (Cl Is Nothing) Then
         If Len(TempData.CUSCOD) > 0 Then
            Call Cl.Add(TempData, TempData.CUSCOD)
         Else
            ''debug.print
         End If
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   Set Rs = Nothing
   Set D = Nothing
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub

Public Sub LoadPaidAmountBySale(C As ComboBox, Optional Cl As Collection = Nothing, Optional FromDate As Date, Optional ToDate As Date)
On Error GoTo ErrorHandler
Dim D As CARRcIt
Dim ItemCount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CARRcIt
Dim i As Long

   Set D = New CARRcIt
   Set Rs = New ADODB.Recordset
   
   D.FROM_PAY_DATE = FromDate
   D.TO_PAY_DATE = ToDate
   D.OrderBy = 1
   Call D.QueryData(4, Rs, ItemCount)
   
   If Not (C Is Nothing) Then
      C.Clear
      i = 0
      C.AddItem ("")
   End If

   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   While Not Rs.EOF
      i = i + 1
      Set TempData = New CARRcIt
      Call TempData.PopulateFromRS(4, Rs)
   
      If Not (C Is Nothing) Then
      End If
      
      If Not (Cl Is Nothing) Then
         If Len(TempData.SLMCOD) > 0 Then
            Call Cl.Add(TempData, TempData.SLMCOD)
         Else
            ''debug.print
         End If
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   Set Rs = Nothing
   Set D = Nothing
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub

Public Sub LoadDbnCdnByBill(C As ComboBox, Optional Cl As Collection = Nothing, Optional FromDate As Date, Optional ToDate As Date)
On Error GoTo ErrorHandler
Dim D As CApTrn
Dim ItemCount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CApTrn
Dim i As Long

   Set D = New CApTrn
   Set Rs = New ADODB.Recordset
   
   D.FROM_DOC_DATE = FromDate
   D.TO_DOC_DATE = ToDate
   D.RECTYP = 4
   D.OrderBy = 1
   Call D.QueryData(3, Rs, ItemCount)
   
   If Not (C Is Nothing) Then
      C.Clear
      i = 0
      C.AddItem ("")
   End If

   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   While Not Rs.EOF
      i = i + 1
      Set TempData = New CApTrn
      Call TempData.PopulateFromRS(3, Rs)
   
      If Not (C Is Nothing) Then
      End If
      
      If Not (Cl Is Nothing) Then
''debug.print TempData.PONUM
         Call Cl.Add(TempData, TempData.PONUM)
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   Set Rs = Nothing
   Set D = Nothing
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub

Public Function GetAPTrn(m_TempCol As Collection, TempKey As String, Optional HaveNew As Boolean = True) As CApTrn
On Error Resume Next
Dim Ei As CApTrn
Static TempEi As CApTrn

   Set Ei = m_TempCol(TempKey)
   If Ei Is Nothing And HaveNew Then
      If TempEi Is Nothing Then
         Set TempEi = New CApTrn
      End If
      Set GetAPTrn = TempEi
   Else
      Set GetAPTrn = Ei
   End If
End Function
Public Function GetARTrnNoNew(m_TempCol As Collection, TempKey As String) As CARTrn
On Error Resume Next
Dim Ei As CARTrn
Static TempEi As CARTrn

   Set Ei = m_TempCol(TempKey)
   If Ei Is Nothing Then
'      If TempEi Is Nothing Then
'         Set TempEi = New CARTrn
'      End If
      Set GetARTrnNoNew = TempEi
   Else
      Set GetARTrnNoNew = Ei
   End If
End Function

Public Function GetCheckCancelitem(m_TempCol As Collection, TempKey As String, Optional HaveNew As Boolean = True) As CCheckCancel
On Error Resume Next
Dim Ei As CCheckCancel
Static TempEi As CCheckCancel

   Set Ei = m_TempCol(TempKey)
    If Ei Is Nothing And HaveNew Then
                If TempEi Is Nothing Then
                   Set TempEi = New CCheckCancel
                End If
      Set GetCheckCancelitem = TempEi
   Else
      Set GetCheckCancelitem = Ei
   End If
End Function
Public Sub LoadBktChqCus(C As ComboBox, Optional Cl As Collection = Nothing, Optional FromGet As Date = -1, Optional ToGet As Date = -1, Optional FromChq As Date = -1, Optional ToChq As Date = -1, Optional FromTrn As Date = -1)
On Error GoTo ErrorHandler
Dim D As CBkTrn
Dim ItemCount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CBkTrn
Dim i As Long
Dim TempBk As CBkTrn
   
   Set D = New CBkTrn
   Set Rs = New ADODB.Recordset
   
   D.FROM_GETDAT = FromGet
   D.TO_GETDAT = ToGet
   D.FROM_CHQDAT = FromChq
   D.TO_CHQDAT = ToChq
   D.FROM_TRNDAT = FromTrn
   D.OrderBy = 1
   Call D.QueryData(9, Rs, ItemCount)  ' 9
   
   If Not (C Is Nothing) Then
      C.Clear
      i = 0
      C.AddItem ("")
   End If

   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   While Not Rs.EOF
      i = i + 1
      Set TempData = New CBkTrn
      Call TempData.PopulateFromRS(9, Rs)  '9
      
      '''debug.print (Trim(TempData.CHQNUM & "-" & TempData.DOCDAT))
      
      If Not (C Is Nothing) Then
      End If
      
      If Not (Cl Is Nothing) Then
         Set TempBk = GetObject("CBkTrn", Cl, Trim(TempData.CUSCOD), False)
         If TempBk Is Nothing Then
            Call Cl.Add(TempData, Trim(TempData.CUSCOD & "-" & TempData.SLMCOD))        ' ลูกค้าคนเดียวกัน sale คนเดียวกัน แต่เช็คคนละใบกัน คีย์ต้องเป็น ? "-" &  TempData.VOUCHER&
         End If
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   Set Rs = Nothing
   Set D = Nothing
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub

Public Sub LoadBktSame1_15_1(C As ComboBox, Optional Cl As Collection = Nothing, Optional FromCusCod As String = "", Optional ToCusCod As String = "", Optional FromSaleCod As String = "", Optional ToSaleCod As String = "", Optional FromGet As Date = -1)
On Error GoTo ErrorHandler
Dim D As CBkTrn
Dim ItemCount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CBkTrn
Dim i As Long
Dim TempBk As CBkTrn
   
   Set D = New CBkTrn
   Set Rs = New ADODB.Recordset
   
      D.FROM_CUSTOMER_CODE = FromCusCod
      D.TO_CUSTOMER_CODE = ToCusCod
      D.FROM_SALE_CODE = FromSaleCod
      D.TO_SALE_CODE = ToSaleCod
      D.FROM_GETDAT = FromGet
      D.OrderBy = 4
      D.CUSTYP = ""
   Call D.QueryData(10, Rs, ItemCount)      ' ต้องรวม
   
   If Not (C Is Nothing) Then
      C.Clear
      i = 0
      C.AddItem ("")
   End If

   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   While Not Rs.EOF
      i = i + 1
      Set TempData = New CBkTrn
      Call TempData.PopulateFromRS(10, Rs)
      
      ''debug.print (Trim(TempData.CUSCOD))
              TempData.Flag = "N"
      If Not (C Is Nothing) Then
      End If
      
      If Not (Cl Is Nothing) Then
         Set TempBk = GetObject("CBkTrn", Cl, Trim(TempData.CUSCOD), False)
         If TempBk Is Nothing Then
            Call Cl.Add(TempData, Trim(TempData.CUSCOD & "-" & TempData.SLMCOD))
         End If
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   Set Rs = Nothing
   Set D = Nothing
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub

Public Function GetBkTrn(m_TempCol As Collection, TempKey As String, Optional HaveNew As Boolean = True) As CBkTrn
On Error Resume Next
Dim Ei As CBkTrn
Static TempEi As CBkTrn
    
   Set Ei = m_TempCol(TempKey)
   If Ei Is Nothing And HaveNew Then
      If TempEi Is Nothing Then
         Set TempEi = New CBkTrn
      End If
      Set GetBkTrn = TempEi
   Else
      Set GetBkTrn = Ei
   End If
End Function
Public Function GetARTrn(m_TempCol As Collection, TempKey As String) As CARTrn
On Error Resume Next
Dim Ei As CARTrn
Static TempEi As CARTrn

   Set Ei = m_TempCol(TempKey)
   If Ei Is Nothing Then
      If TempEi Is Nothing Then
         Set TempEi = New CARTrn
      End If
      Set GetARTrn = TempEi
   Else
      Set GetARTrn = Ei
   End If
End Function
Public Function GetAnalyzeCustomer(m_TempCol As Collection, TempKey As String) As CAnalyzeCustomer
On Error Resume Next
Dim Ei As CAnalyzeCustomer
Static TempEi As CAnalyzeCustomer

   Set Ei = m_TempCol(TempKey)
   If Ei Is Nothing Then
'      If TempEi Is Nothing Then
         Set TempEi = New CAnalyzeCustomer
'      End If
      Set GetAnalyzeCustomer = TempEi
   Else
      Set GetAnalyzeCustomer = Ei
   End If
End Function
Public Sub LoadBank(C As ComboBox, Optional Cl As Collection = Nothing, Optional Bank As String = "")
On Error GoTo ErrorHandler
Dim D As CBank
Dim ItemCount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CBank
Dim i As Long
   
   Set D = New CBank
   Set Rs = New ADODB.Recordset
   
   Call D.QueryData(Rs, ItemCount)
   
   If Not (C Is Nothing) Then
      C.Clear
      i = 0
      C.AddItem ("")
   End If
   
   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   
   While Not Rs.EOF
      i = i + 1
      Set TempData = New CBank
      Call TempData.PopulateFromRS(1, Rs)

      If Not (C Is Nothing) Then
         C.AddItem (TempData.BANK_NAME)
         C.ItemData(i) = TempData.BANK_ID
      End If

      If Not (Cl Is Nothing) Then
         Call Cl.Add(TempData, Trim(Str(TempData.BANK_ID)))
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   Set Rs = Nothing
   Set D = Nothing
   Exit Sub

ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub
Public Sub LoadBankCustomer(C As ComboBox, Optional Cl As Collection = Nothing, Optional BankCustomer As String = "")
On Error GoTo ErrorHandler
Dim D As CBankCustomer
Dim ItemCount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CBankCustomer
Dim i As Long
   
   Set D = New CBankCustomer
   Set Rs = New ADODB.Recordset
   
   Call D.QueryData(Rs, ItemCount)
   
   If Not (C Is Nothing) Then
      C.Clear
      i = 0
      C.AddItem ("")
   End If
   
   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   
   While Not Rs.EOF
      i = i + 1
      Set TempData = New CBankCustomer
      Call TempData.PopulateFromRS(1, Rs)

      If Not (C Is Nothing) Then
         C.AddItem (TempData.CUSTOMER_NAME)
         C.ItemData(i) = TempData.CUSTOMER_ID
      End If

      If Not (Cl Is Nothing) Then
         Call Cl.Add(TempData, Trim(Str(TempData.CUSTOMER_ID)))
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   Set Rs = Nothing
   Set D = Nothing
   Exit Sub

ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub
Public Sub LoadBankCredit(C As ComboBox, Optional Cl As Collection = Nothing, Optional BankCredit As Long = -1, Optional BankCustomer As Long = -1, Optional ToDate As Date = -1, Optional TempDate As Date = -1, Optional Ind As Long)
On Error GoTo ErrorHandler
Dim D As CBankCredit
Dim ItemCount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CBankCredit
Dim i As Long
   
   Set D = New CBankCredit
   Set Rs = New ADODB.Recordset
   D.TO_DATE = ToDate
   D.BANK_ID = BankCredit
   D.CUSTOMER_ID = BankCustomer
   Call D.QueryData(Ind, Rs, ItemCount)
   
   If Not (C Is Nothing) Then
      C.Clear
      i = 0
      C.AddItem ("")
   End If
   
   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   
   While Not Rs.EOF
      i = i + 1
      Set TempData = New CBankCredit
      Call TempData.PopulateFromRS(Ind, Rs)

      If Not (C Is Nothing) Then
         C.AddItem (TempData.BANK_NAME)
         C.ItemData(i) = TempData.BANK_ID
      End If
      
      If BankCredit > 0 Then
         If TempDate < TempData.BANK_DATE_BROUGHT Then
               If Not (Cl Is Nothing) Then
                  Set Cl = New Collection
                  Call Cl.Add(TempData)
               End If
               TempDate = TempData.BANK_DATE_BROUGHT
         End If
      Else
            If Not (Cl Is Nothing) Then
                  Call Cl.Add(TempData)
            End If
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   Set Rs = Nothing
   Set D = Nothing
   Exit Sub

ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub
Public Sub LoadTicket(C As ComboBox, Optional Cl As Collection = Nothing, Optional Cl2 As Collection = Nothing, Optional Maxdate As Date = -1, Optional Mindate As Date = -1, Optional FromDate As Date, Optional ToDate As Date, Optional BankCustomer As Long = -1, Optional BankCredit As Long = -1)
On Error GoTo ErrorHandler
Dim D As CTicket
Dim ItemCount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CTicket
Dim i As Long
Dim Count As Long
Dim PrevKey1 As String
Dim PrevKey2 As Date
   
   Count = 0
   Set D = New CTicket
   Set Rs = New ADODB.Recordset
   D.TO_DATE = ToDate
   D.CUSTOMER_ID = BankCustomer
   D.BANK_ID = BankCredit
   D.MASTER_AREA = 1
   Call D.QueryData(1, Rs, ItemCount)
   
   If Not (C Is Nothing) Then
      C.Clear
      i = 0
      C.AddItem ("")
   End If
   
   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   
   If Not (Cl2 Is Nothing) Then
      Set Cl2 = Nothing
      Set Cl2 = New Collection
   End If
   
   While Not Rs.EOF
      i = i + 1
      Set TempData = New CTicket
      Call TempData.PopulateFromRS(1, Rs)
      
      If i = 1 Then
         Mindate = TempData.TICKET_DATE
      End If
      If Maxdate < TempData.TICKET_DATE_CHECK Then
         Maxdate = TempData.TICKET_DATE_CHECK
      End If
      If Not (Cl Is Nothing) Then
         TempData.TEMP_TYPE = "P"     'Pack
         Call Cl.Add(TempData, Trim(Str(TempData.TICKET_ID)))
      End If
      If PrevKey1 <> TempData.TICKET_INVOICE Then
         TempData.TEMP_COUNT = Count
         If Not (Cl2 Is Nothing) Then
            Call Cl2.Add(TempData, Trim(PrevKey1 & "-" & PrevKey2))
         End If
         Count = 0
      End If
      Count = Count + 1
      PrevKey1 = TempData.TICKET_INVOICE
      PrevKey2 = TempData.TICKET_DATE_CHECK
      
      If i = Rs.RecordCount Then
         TempData.TEMP_COUNT = Count
         If Not (Cl2 Is Nothing) Then
            Call Cl2.Add(TempData, Trim(PrevKey1 & "-" & PrevKey2))
         End If
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   
   Set Rs = Nothing
   Set D = Nothing
   Exit Sub

ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub
Public Sub LoadBudgetTicket(C As ComboBox, Optional Cl As Collection = Nothing, Optional FromDate As Date, Optional ToDate As Date, Optional BankCustomer As Long = -1, Optional BankCredit As Long = -1)
On Error GoTo ErrorHandler
Dim D As CTicket
Dim ItemCount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CTicket
Dim i As Long
Dim Count As Long
   
   Count = 0
   Set D = New CTicket
   Set Rs = New ADODB.Recordset
   D.TO_DATE = ToDate
   D.CUSTOMER_ID = BankCustomer
   D.BANK_ID = BankCredit
   D.MASTER_AREA = 2
   Call D.QueryData(1, Rs, ItemCount)
   
   If Not (C Is Nothing) Then
      C.Clear
      i = 0
      C.AddItem ("")
   End If
   
   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   
   While Not Rs.EOF
      i = i + 1
      Set TempData = New CTicket
      Call TempData.PopulateFromRS(1, Rs)

      If Not (Cl Is Nothing) Then
         TempData.TEMP_TYPE = "P"     'Pack
         Call Cl.Add(TempData, Trim(Str(TempData.TICKET_ID)))
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   
   Set Rs = Nothing
   Set D = Nothing
   Exit Sub

ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub
Public Sub LoadTicketMaxAmount(C As ComboBox, Optional Cl As Collection = Nothing)
On Error GoTo ErrorHandler
Dim D As CTicket
Dim ItemCount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CTicket
Dim i As Long
   
   Set D = New CTicket
   Set Rs = New ADODB.Recordset
   D.MASTER_AREA = 1
   Call D.QueryData(2, Rs, ItemCount)
   
   If Not (C Is Nothing) Then
      C.Clear
      i = 0
      C.AddItem ("")
   End If
   
   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   
   While Not Rs.EOF
      i = i + 1
      Set TempData = New CTicket
      Call TempData.PopulateFromRS(2, Rs)
      
      If Not (Cl Is Nothing) Then
         Call Cl.Add(TempData, Trim(TempData.TICKET_INVOICE))
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   Set Rs = Nothing
   Set D = Nothing
   Exit Sub

ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub
Public Sub LoadTicketReceive(C As ComboBox, Optional Cl As Collection = Nothing, Optional Cl2 As Collection, Optional FromDate As Date, Optional ToDate As Date, Optional BankCustomer As Long, Optional BankCredit As Long)
On Error GoTo ErrorHandler
Dim D As CTicket
Dim Check As CTicket
Dim ItemCount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CTicket
Dim i As Long
Dim TempDate As Long
   
   Set D = New CTicket
   Set Check = New CTicket
   Set Rs = New ADODB.Recordset
   D.TO_DATE_RECEIVE = ToDate
   D.CUSTOMER_ID = BankCustomer
   D.BANK_ID = BankCredit
   D.MASTER_AREA = 1
   Call D.QueryData(1, Rs, ItemCount)
   TempDate = 0
   
   If Not (C Is Nothing) Then
      C.Clear
      i = 0
      C.AddItem ("")
   End If
   
   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   
   While Not Rs.EOF
      i = i + 1
      Set TempData = New CTicket
      Call TempData.PopulateFromRS(1, Rs)
      
      If Not (Cl Is Nothing) Then
         TempDate = Weekday(TempData.TICKET_DATE_CHECK)
         If TempDate = 1 Then
            TempData.TICKET_DATE = DateAdd("D", 1, TempData.TICKET_DATE_CHECK)
         ElseIf TempDate = 7 Then
            TempData.TICKET_DATE = DateAdd("D", 2, TempData.TICKET_DATE_CHECK)
         Else
            TempData.TICKET_DATE = TempData.TICKET_DATE_CHECK
         End If
         TempData.TEMP_TYPE = "B"     'Bank
         Set Check = GetObject("CTicket", Cl2, Trim(TempData.TICKET_INVOICE & "-" & TempData.TICKET_DATE_CHECK), False)
         If Check Is Nothing Then
            TempData.TEMP_CHECK = "N"
         Else
            TempData.TEMP_COUNT = Check.TEMP_COUNT
            TempData.TEMP_CHECK = "Y"
         End If
         Call Cl.Add(TempData, Trim(Str(TempData.TICKET_ID)))
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   Set Rs = Nothing
   Set D = Nothing
   Exit Sub

ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub
Public Sub LoadBudgetTicketReceive(C As ComboBox, Optional Cl As Collection = Nothing, Optional FromDate As Date, Optional ToDate As Date, Optional BankCustomer As Long, Optional BankCredit As Long)
On Error GoTo ErrorHandler
Dim D As CTicket
Dim Check As CTicket
Dim ItemCount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CTicket
Dim i As Long
Dim TempDate As Long
   
   Set D = New CTicket
   Set Check = New CTicket
   Set Rs = New ADODB.Recordset
   D.TO_DATE_RECEIVE = ToDate
   D.CUSTOMER_ID = BankCustomer
   D.BANK_ID = BankCredit
   D.MASTER_AREA = 2
   Call D.QueryData(1, Rs, ItemCount)
   TempDate = 0
   
   If Not (C Is Nothing) Then
      C.Clear
      i = 0
      C.AddItem ("")
   End If
   
   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   
   While Not Rs.EOF
      i = i + 1
      Set TempData = New CTicket
      Call TempData.PopulateFromRS(1, Rs)
      
      If Not (Cl Is Nothing) Then
         TempDate = Weekday(TempData.TICKET_DATE_CHECK)
         If TempDate = 1 Then
            TempData.TICKET_DATE = DateAdd("D", 1, TempData.TICKET_DATE_CHECK)
         ElseIf TempDate = 7 Then
            TempData.TICKET_DATE = DateAdd("D", 2, TempData.TICKET_DATE_CHECK)
         Else
            TempData.TICKET_DATE = TempData.TICKET_DATE_CHECK
         End If
         TempData.TEMP_TYPE = "B"     'Bank
         Call Cl.Add(TempData, Trim(Str(TempData.TICKET_ID)))
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   Set Rs = Nothing
   Set D = Nothing
   Exit Sub

ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub
Public Sub LoadTicketClear(C As ComboBox, Optional Cl As Collection = Nothing, Optional FromDate As Date, Optional ToDate As Date, Optional BankCustomer As Long = -1, Optional BankCredit As Long = -1)
On Error GoTo ErrorHandler
Dim D As CTicket
Dim ItemCount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CTicket
Dim i As Long
Dim TempDate As Long
Dim PrevKey1 As Date
Dim PrevKey2 As String
   
   Set D = New CTicket
   Set Rs = New ADODB.Recordset
   D.TO_DATE_RECEIVE = ToDate
   D.CUSTOMER_ID = BankCustomer
   D.BANK_ID = BankCredit
   D.MASTER_AREA = 1
   Call D.QueryData(1, Rs, ItemCount)
   
   If Not (C Is Nothing) Then
      C.Clear
      i = 0
      C.AddItem ("")
   End If
   
   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   
   While Not Rs.EOF
      i = i + 1
      Set TempData = New CTicket
      Call TempData.PopulateFromRS(1, Rs)
      
      If Not (Cl Is Nothing) Then
         If TempData.TICKET_DATE_NEW > 0 Then
            TempDate = Weekday(TempData.TICKET_DATE_NEW)
            If TempDate = 1 Then
               TempData.TICKET_DATE = DateAdd("D", 3, TempData.TICKET_DATE_NEW)
            ElseIf TempDate = 5 Or TempDate = 6 Or TempDate = 7 Then
               TempData.TICKET_DATE = DateAdd("D", 4, TempData.TICKET_DATE_NEW)
            Else
               TempData.TICKET_DATE = DateAdd("D", 2, TempData.TICKET_DATE_NEW)
            End If
         Else
            TempDate = Weekday(TempData.TICKET_DATE_CHECK)
            If TempDate = 1 Then
               TempData.TICKET_DATE = DateAdd("D", 3, TempData.TICKET_DATE_CHECK)
            ElseIf TempDate = 5 Or TempDate = 6 Or TempDate = 7 Then
               TempData.TICKET_DATE = DateAdd("D", 4, TempData.TICKET_DATE_CHECK)
            Else
               TempData.TICKET_DATE = DateAdd("D", 2, TempData.TICKET_DATE_CHECK)
            End If
         End If
         TempData.TEMP_TYPE = "C"     'Clear
         If PrevKey1 <> TempData.TICKET_DATE Or PrevKey2 <> TempData.TICKET_INVOICE Then
            Call Cl.Add(TempData, Trim(Str(TempData.TICKET_ID)))
         End If
         If TempData.TICKET_DATE_NEW > 0 Then
            PrevKey1 = TempData.TICKET_DATE
            PrevKey2 = TempData.TICKET_INVOICE
         End If
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   Set Rs = Nothing
   Set D = Nothing
   Exit Sub

ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub
Public Sub LoadBudgetTicketClear(C As ComboBox, Optional Cl As Collection = Nothing, Optional FromDate As Date, Optional ToDate As Date, Optional BankCustomer As Long = -1, Optional BankCredit As Long = -1)
On Error GoTo ErrorHandler
Dim D As CTicket
Dim ItemCount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CTicket
Dim i As Long
Dim TempDate As Long
Dim PrevKey1 As Date
Dim PrevKey2 As String
   
   Set D = New CTicket
   Set Rs = New ADODB.Recordset
   D.TO_DATE_RECEIVE = ToDate
   D.CUSTOMER_ID = BankCustomer
   D.BANK_ID = BankCredit
   D.MASTER_AREA = 2
   Call D.QueryData(1, Rs, ItemCount)
   
   If Not (C Is Nothing) Then
      C.Clear
      i = 0
      C.AddItem ("")
   End If
   
   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   
   While Not Rs.EOF
      i = i + 1
      Set TempData = New CTicket
      Call TempData.PopulateFromRS(1, Rs)
      
      If Not (Cl Is Nothing) Then
         If TempData.TICKET_DATE_NEW > 0 Then
            TempData.TICKET_DATE = TempData.TICKET_DATE_NEW
         Else
            TempDate = Weekday(TempData.TICKET_DATE_CHECK)
            If TempDate = 1 Then
               TempData.TICKET_DATE = DateAdd("D", 3, TempData.TICKET_DATE_CHECK)
            ElseIf TempDate = 5 Or TempDate = 6 Or TempDate = 7 Then
               TempData.TICKET_DATE = DateAdd("D", 4, TempData.TICKET_DATE_CHECK)
            Else
               TempData.TICKET_DATE = DateAdd("D", 2, TempData.TICKET_DATE_CHECK)
            End If
         End If
         TempData.TEMP_TYPE = "C"     'Clear
         If PrevKey1 <> TempData.TICKET_DATE Or PrevKey2 <> TempData.TICKET_INVOICE Then
            Call Cl.Add(TempData, Trim(Str(TempData.TICKET_ID)))
         End If
         If TempData.TICKET_DATE_NEW > 0 Then
            PrevKey1 = TempData.TICKET_DATE
            PrevKey2 = TempData.TICKET_INVOICE
         End If
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   Set Rs = Nothing
   Set D = Nothing
   Exit Sub

ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub
Public Sub LoadTicketAmount(C As ComboBox, Optional Cl As Collection = Nothing)
On Error GoTo ErrorHandler
Dim D As CARTrn
Dim ItemCount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CARTrn
Dim i As Long
   
   Set D = New CARTrn
   Set Rs = New ADODB.Recordset
   
   Call D.QueryData(11, Rs, ItemCount)
   
   If Not (C Is Nothing) Then
      C.Clear
      i = 0
      C.AddItem ("")
   End If
   
   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   
   While Not Rs.EOF
      Set TempData = New CARTrn
      Call TempData.PopulateFromRS(11, Rs, i)
      
      Set D = GetARTrnNoNew(Cl, TempData.DOCNUM)
      If D Is Nothing Then
        Call Cl.Add(TempData, TempData.DOCNUM)
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   Set Rs = Nothing
   Set D = Nothing
   Exit Sub

ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub
Public Sub LoadTicketLookup(C As ComboBox, Optional Cl As Collection = Nothing)
On Error GoTo ErrorHandler
Dim D As CARTrn
Dim ItemCount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CARTrn
Dim i As Long
   
   Set D = New CARTrn
   Set Rs = New ADODB.Recordset
   
   Call D.QueryData(11, Rs, ItemCount)
   
   If Not (C Is Nothing) Then
      C.Clear
      i = 0
      C.AddItem ("")
   End If
   
   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   
   While Not Rs.EOF
      i = i + 1
      Set TempData = New CARTrn
      Call TempData.PopulateFromRS(11, Rs, i)
      If Not (C Is Nothing) Then
         C.AddItem (" วันที่ " & TempData.DOCDAT & ", จำนวนเงิน " & FormatNumber(TempData.AMOUNT) & " บาท")
         C.ItemData(i) = TempData.KEY_ID
      End If
      
      Set D = GetARTrnNoNew(Cl, TempData.KEY_ID)
      If D Is Nothing Then
        Call Cl.Add(TempData, Trim(Str(TempData.KEY_ID)))
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   Set Rs = Nothing
   Set D = Nothing
   Exit Sub

ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub
Public Sub LoadTicketData(C As ComboBox, Optional Cl As Collection = Nothing, Optional Cl1 As Collection, Optional Cl2 As Collection, Optional Cl3 As Collection, Optional Cl4 As Collection = Nothing, Optional Cl5 As Collection = Nothing, Optional Cl6 As Collection = Nothing, Optional FromDate As Date, Optional ToDate As Date)
On Error GoTo ErrorHandler
Dim TempData As CTicket
Dim TempData2 As CTicket
Dim TempData3 As CTicket
Dim TempData4 As CTicket
Dim TempData5 As CTicket
Dim TempData6 As CTicket
Dim Artrn As CARTrn
Dim i As Long

   Set TempData = New CTicket
   Set TempData2 = New CTicket
   Set TempData3 = New CTicket
   Set TempData4 = New CTicket
   Set TempData5 = New CTicket
   Set TempData6 = New CTicket
   
   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   
   While (FromDate <= ToDate)
      For Each TempData In Cl1
         Set Artrn = GetARTrn(Cl1, TempData.TICKET_INVOICE)
         If FromDate = TempData.TICKET_DATE Then
            If Not (Cl Is Nothing) Then
               Call Cl.Add(TempData)
            End If
         End If
      Next TempData
      
      For Each TempData4 In Cl4        'ประมาณการ
         Set Artrn = GetARTrn(Cl4, TempData4.TICKET_INVOICE)
         If FromDate = TempData4.TICKET_DATE Then
            If Not (Cl Is Nothing) Then
               Call Cl.Add(TempData4)
            End If
         End If
      Next TempData4
      
      For Each TempData2 In Cl2
         Set Artrn = GetARTrn(Cl2, TempData2.TICKET_INVOICE)
         If FromDate = TempData2.TICKET_DATE Then
               If Not (Cl Is Nothing) Then
                  Call Cl.Add(TempData2)
               End If
         End If
      Next TempData2
      
      For Each TempData5 In Cl5        'ประมาณการ
         Set Artrn = GetARTrn(Cl5, TempData5.TICKET_INVOICE)
         If FromDate = TempData5.TICKET_DATE Then
               If Not (Cl Is Nothing) Then
                  Call Cl.Add(TempData5)
               End If
         End If
      Next TempData5
      
      For Each TempData3 In Cl3
         Set Artrn = GetARTrn(Cl3, TempData3.TICKET_INVOICE)
         If FromDate = TempData3.TICKET_DATE Then
               If Not (Cl Is Nothing) Then
                  Call Cl.Add(TempData3)
               End If
         End If
      Next TempData3
      
      For Each TempData6 In Cl6     'ประมาณการ
         Set Artrn = GetARTrn(Cl6, TempData6.TICKET_INVOICE)
         If FromDate = TempData6.TICKET_DATE Then
               If Not (Cl Is Nothing) Then
                  Call Cl.Add(TempData6)
               End If
         End If
      Next TempData6
      
      FromDate = DateAdd("D", 1, FromDate)
   Wend
   
   Set TempData = Nothing
   Exit Sub

ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub
Public Sub LoadCusCode(C As ComboBox, Optional Cl As Collection = Nothing, Optional CustomerType As Long)
On Error GoTo ErrorHandler
Dim D As CCustomer
Dim ItemCount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CCustomer
Dim i As Long
   
   Set D = New CCustomer
   Set Rs = New ADODB.Recordset
   
   D.CUSTOMER_TYPE_ID = CustomerType
   Call D.QueryData(Rs, ItemCount)
   
   If Not (C Is Nothing) Then
      C.Clear
      i = 0
      C.AddItem ("")
   End If
   
   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   
   While Not Rs.EOF
      Set TempData = New CCustomer
      Call TempData.PopulateFromRS(1, Rs)
      Call Cl.Add(TempData)
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   Set Rs = Nothing
   Set D = Nothing
   Exit Sub

ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub
Public Sub LoadCusCodeAmount(C As ComboBox, Optional Cl As Collection = Nothing, Optional FromDate As Date, Optional ToDate As Date)
On Error GoTo ErrorHandler
Dim D As CARTrn
Dim ItemCount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CARTrn
Dim i As Long
   
   Set D = New CARTrn
   Set Rs = New ADODB.Recordset
   
    D.FROM_DOC_DATE = FromDate
    D.TO_DOC_DATE = ToDate
    D.OrderBy = 4
    D.OrderType = 1
    D.QueryFlag = -1
    D.RecTypeSet = "('3','4')"
   Call D.QueryData(1, Rs, ItemCount)
   
   If Not (C Is Nothing) Then
      C.Clear
      i = 0
      C.AddItem ("")
   End If
   
   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   
   While Not Rs.EOF
      Set TempData = New CARTrn
      Call TempData.PopulateFromRS(1, Rs)
      Set D = GetARTrnNoNew(Cl, TempData.CUSCOD & "-" & Format("00", Month(TempData.DOCDAT)) & Year(TempData.DOCDAT))
      If D Is Nothing Then
        Call Cl.Add(TempData, TempData.CUSCOD & "-" & Format("00", Month(TempData.DOCDAT)) & Year(TempData.DOCDAT))
      Else
        D.AMOUNT = D.AMOUNT + TempData.AMOUNT
      End If
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   Set Rs = Nothing
   Set D = Nothing
   Exit Sub

ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub

Public Sub LoadXlsUnit(C As ComboBox, Optional Cl As Collection = Nothing, Optional CustomerType As String = "")
On Error GoTo ErrorHandler
Dim D As CXlsUnit
Dim ItemCount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CXlsUnit
Dim i As Long
   
   Set D = New CXlsUnit
   Set Rs = New ADODB.Recordset
   
   Call D.QueryData(1, Rs, ItemCount)
   
   If Not (C Is Nothing) Then
      C.Clear
      i = 0
      C.AddItem ("")
   End If
   
   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   
   While Not Rs.EOF
      i = i + 1
      Set TempData = New CXlsUnit
      Call TempData.PopulateFromRS(1, Rs)

      If Not (C Is Nothing) Then
         C.AddItem (TempData.XLS_UNIT_NAME)
         C.ItemData(i) = TempData.XLS_UNIT_ID
      End If

      If Not (Cl Is Nothing) Then
         Call Cl.Add(TempData, Trim(Str(TempData.XLS_UNIT_ID)))
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
'   If Not (Cl Is Nothing) Then
'        Set TempData = New CCustomerType
'        TempData.CUSTOMER_TYPE_ID = 999999
'        TempData.CUSTOMER_TYPE_NAME = "อื่นๆ"
'        Call Cl.Add(TempData, Trim(Str(TempData.CUSTOMER_TYPE_ID)))
'        Set TempData = Nothing
'   End If
   
   Set Rs = Nothing
   Set D = Nothing
   Exit Sub

ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub

Public Sub LoadXlsNameFarm(C As ComboBox, Optional Cl As Collection = Nothing, Optional CustomerType As String = "")
On Error GoTo ErrorHandler
Dim D As CXlsMainfarm
Dim ItemCount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CXlsMainfarm
Dim i As Long
   
   Set D = New CXlsMainfarm
   Set Rs = New ADODB.Recordset
   
   Call D.QueryData(1, Rs, ItemCount)
   
   If Not (C Is Nothing) Then
      C.Clear
      i = 0
      C.AddItem ("")
   End If
   
   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   
   While Not Rs.EOF
      i = i + 1
      Set TempData = New CXlsMainfarm
      Call TempData.PopulateFromRS(1, Rs)

      If Not (C Is Nothing) Then
         C.AddItem (TempData.MAIN_FARM_NAME)
         C.ItemData(i) = TempData.MAIN_FARM_ID
      End If

      If Not (Cl Is Nothing) Then
         Call Cl.Add(TempData, Trim(Str(TempData.MAIN_FARM_ID)))
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   Set Rs = Nothing
   Set D = Nothing
   Exit Sub

ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub

Public Sub LoadCustomerType2(C As ComboBox, Optional Cl As Collection = Nothing, Optional CustomerType As String = "")
On Error GoTo ErrorHandler
Dim D As CCustomerType
Dim ItemCount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CCustomerType
Dim i As Long
   
   Set D = New CCustomerType
   Set Rs = New ADODB.Recordset
   
   Call D.QueryData(Rs, ItemCount)
   
   If Not (C Is Nothing) Then
      C.Clear
      i = 0
      C.AddItem ("")
   End If
   
   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   
   While Not Rs.EOF
      i = i + 1
      Set TempData = New CCustomerType
      Call TempData.PopulateFromRS(1, Rs)

      If Not (C Is Nothing) Then
         C.AddItem (TempData.CUSTOMER_TYPE_NAME)
         C.ItemData(i) = TempData.CUSTOMER_TYPE_ID
      End If

      If Not (Cl Is Nothing) Then
         Call Cl.Add(TempData, Trim(Str(TempData.CUSTOMER_TYPE_ID)))
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   If Not (Cl Is Nothing) Then
        Set TempData = New CCustomerType
        TempData.CUSTOMER_TYPE_ID = 999999
        TempData.CUSTOMER_TYPE_NAME = "อื่นๆ"
        Call Cl.Add(TempData, Trim(Str(TempData.CUSTOMER_TYPE_ID)))
        Set TempData = Nothing
   End If
   
   Set Rs = Nothing
   Set D = Nothing
   Exit Sub

ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub
Public Sub LoadCusPigType(Optional Cl As Collection = Nothing) '(Optional Cl As Collection = Nothing, Optional C As ComboBox, Optional CustomerType As String = "")
On Error GoTo ErrorHandler
Dim D As CCusPigType
Dim ItemCount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CCusPigType
Dim i As Long
   
   Set D = New CCusPigType
   Set Rs = New ADODB.Recordset
   
   Call D.QueryData(1, Rs, ItemCount)
   
'   If Not (C Is Nothing) Then
'      C.Clear
'      i = 0
'      C.AddItem ("")
'   End If
'   If Not (C Is Nothing) Then
'         C.AddItem (TempData.CUS_PIG_TYPE_NAME)
'         C.ItemData(i) = TempData.CUS_PIG_TYPE_CODE
'   End If
      
   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   
   While Not Rs.EOF
      i = i + 1
      Set TempData = New CCusPigType
      Call TempData.PopulateFromRS(1, Rs)
'Debug.Print TempData.CUS_PIG_TYPE_CODE
      If Not (Cl Is Nothing) Then
'         Call Cl.Add(TempData, Trim(TempData.CUS_PIG_TYPE_CODE) & "-" & Trim(Str(TempData.CUS_PIG_TYPE_YEAR)))
         Call Cl.Add(TempData, Trim(TempData.CUS_PIG_TYPE_CODE))
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   Set Rs = Nothing
   Set D = Nothing
   Exit Sub

ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub
Public Function GetFaDprIt(m_TempCol As Collection, TempKey As String) As CFadprIt
On Error Resume Next
Dim Ei As CFadprIt
Static TempEi As CFadprIt

   Set Ei = m_TempCol(TempKey)
   If Ei Is Nothing Then
      If TempEi Is Nothing Then
         Set TempEi = New CFadprIt
      End If
      Set GetFaDprIt = TempEi
   Else
      Set GetFaDprIt = Ei
   End If
End Function
Public Function GetGLJnl(m_TempCol As Collection, TempKey As String, Optional ReturnV As Boolean = True) As CGLJnl
On Error Resume Next
Dim Ei As CGLJnl
Static TempEi As CGLJnl

   Set Ei = m_TempCol(TempKey)
   If Ei Is Nothing Then
      If TempEi Is Nothing Then
         Set TempEi = New CGLJnl
      End If
      Set GetGLJnl = TempEi
      ReturnV = False
   Else
      Set GetGLJnl = Ei
      ReturnV = True
   End If
End Function
Public Function GetGLAcc(m_TempCol As Collection, TempKey As String) As CGLAcc
On Error Resume Next
Dim Ei As CGLAcc
Static TempEi As CGLAcc

   Set Ei = m_TempCol(TempKey)
   If Ei Is Nothing Then
      If TempEi Is Nothing Then
         Set TempEi = New CGLAcc
      End If
      Set GetGLAcc = TempEi
   Else
      Set GetGLAcc = Ei
   End If
End Function
Public Function GetAccountCode(m_TempCol As Collection, TempKey As String) As CAccountCode
On Error Resume Next
Dim Ei As CAccountCode
Static TempEi As CAccountCode

   Set Ei = m_TempCol(TempKey)
   If Ei Is Nothing Then
'      If TempEi Is Nothing Then
'         Set TempEi = New CAccountCode
'      End If
      Set GetAccountCode = TempEi
   Else
      Set GetAccountCode = Ei
   End If
End Function

Public Function GetAPRcpItem(m_TempCol As Collection, TempKey As String) As CAPRcIt
On Error Resume Next
Dim Ei As CAPRcIt
Static TempEi As CAPRcIt

   Set Ei = m_TempCol(TempKey)
   If Ei Is Nothing Then
      If TempEi Is Nothing Then
         Set TempEi = New CAPRcIt
      End If
      Set GetAPRcpItem = TempEi
   Else
      Set GetAPRcpItem = Ei
   End If
End Function

Public Function GetARRcpItem(m_TempCol As Collection, TempKey As String, Optional HaveNew As Boolean = True) As CARRcIt
On Error Resume Next
Dim Ei As CARRcIt
Static TempEi As CARRcIt

   Set Ei = m_TempCol(TempKey)
   If Ei Is Nothing And HaveNew Then
      If TempEi Is Nothing Then
         Set TempEi = New CARRcIt
      End If
      Set GetARRcpItem = TempEi
   Else
      Set GetARRcpItem = Ei
   End If
End Function

Public Function GetAPRcpItemEx(m_TempCol As Collection, TempKey As String) As CAPRcIt
On Error Resume Next
Dim Ei As CAPRcIt
Static TempEi As CAPRcIt

   Set Ei = m_TempCol(TempKey)
   If Ei Is Nothing Then
      If TempEi Is Nothing Then
         Set TempEi = New CAPRcIt
      End If
      Set GetAPRcpItemEx = TempEi
   Else
      Set GetAPRcpItemEx = Ei
   End If
End Function

Public Function GetARRcpItemEx(m_TempCol As Collection, TempKey As String) As CARRcIt
On Error Resume Next
Dim Ei As CARRcIt
Static TempEi As CARRcIt

   Set Ei = m_TempCol(TempKey)
   If Ei Is Nothing Then
      If TempEi Is Nothing Then
         Set TempEi = New CARRcIt
      End If
      Set GetARRcpItemEx = TempEi
   Else
      Set GetARRcpItemEx = Ei
   End If
End Function

Public Sub LoadGLJNLforAccountExcel(C As ComboBox, Optional Cl As Collection = Nothing, Optional FromDate As Date, Optional ToDate As Date)
On Error Resume Next
Dim D As CGLJnl
Dim ItemCount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CGLJnl
Dim i As Long
Dim j  As Long

   Set D = New CGLJnl
   Set Rs = New ADODB.Recordset
   
   D.FROM_VOUCHER_DATE = FromDate
   D.TO_VOUCHER_DATE = ToDate
   Call D.QueryData(3, Rs, ItemCount)
   
   If Not (C Is Nothing) Then
      C.Clear
      i = 0
      C.AddItem ("")
   End If

   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   
   j = 0
   
   While Not Rs.EOF
      i = i + 1
      Set TempData = New CGLJnl
      Call TempData.PopulateFromRS(3, Rs)
   
      If Not (C Is Nothing) Then
      End If
      
      If Not (Cl Is Nothing) Then
         Call Cl.Add(TempData, Trim(TempData.ACCNUM & "-" & TempData.TRNTYP))
'         If J = 0 Then
'            ''debug.print (Trim(TempData.ACCNUM))
'            J = 1
'         Else
'            J = 0
'         End If
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   Set Rs = Nothing
   Set D = Nothing
'   Exit Sub
   
'ErrorHandler:
'   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
'   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub
Public Sub LoadGLAccSearch(C As ComboBox, Optional Cl As Collection = Nothing)
On Error Resume Next
Dim D As CGLAcc
Dim ItemCount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CGLAcc
Dim i As Long
Dim j  As Long

   Set D = New CGLAcc
   Set Rs = New ADODB.Recordset
   
   Call D.QueryData(-1, Rs, ItemCount)
   
   If Not (C Is Nothing) Then
      C.Clear
      i = 0
      C.AddItem ("")
   End If

   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   
   j = 0
   
   While Not Rs.EOF
      i = i + 1
      Set TempData = New CGLAcc
      Call TempData.PopulateFromRS(-1, Rs)
   
      If Not (C Is Nothing) Then
      End If
      
      If Not (Cl Is Nothing) Then
         Call Cl.Add(TempData, Trim(TempData.ACCNUM))
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   Set Rs = Nothing
   Set D = Nothing
'   Exit Sub
'
'ErrorHandler:
'   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
'   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub


Public Sub LoadCheckCancel(Optional Cl As Collection = Nothing, Optional TempKey As String = "", Optional Rs As ADODB.Recordset)
On Error GoTo ErrorHandler
Dim D As CCheckCancel
Dim ItemCount As Long
Dim TempData As CCheckCancel
Dim i As Long
   
   Set D = New CCheckCancel

   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   While Not Rs.EOF
      i = i + 1
      Set TempData = New CCheckCancel
      Call TempData.PopulateFromRS(3, Rs)

      If Not (Cl Is Nothing) Then
         Call Cl.Add(TempData, TempData.CHECK_NO)
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   Set Rs = Nothing
   Set D = Nothing
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub

Public Sub LoadXlsCarkillSum(C As ComboBox, Optional Cl As Collection = Nothing)
On Error GoTo ErrorHandler
Dim D As CXlsCarkillSum
Dim ItemCount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CXlsCarkillSum
Dim i As Long

   Set D = New CXlsCarkillSum
   Set Rs = New ADODB.Recordset
   
   D.OrderBy = 1
   Call D.QueryData(1, Rs, ItemCount)
   
   If Not (C Is Nothing) Then
      C.Clear
      i = 0
      C.AddItem ("")
   End If

   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   
   While Not Rs.EOF
      i = i + 1
      Set TempData = New CXlsCarkillSum
      Call TempData.PopulateFromRS(1, Rs)
   
'      If Not (C Is Nothing) Then
'      End If
      
      If Not (Cl Is Nothing) Then
         Call Cl.Add(TempData)
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   Set Rs = Nothing
   Set D = Nothing
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub

Public Sub LoadXlsCarkillFW(C As ComboBox, Optional Cl As Collection = Nothing)
On Error GoTo ErrorHandler
Dim D As CXlsCarkillFW
Dim ItemCount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CXlsCarkillFW
Dim i As Long

   Set D = New CXlsCarkillFW
   Set Rs = New ADODB.Recordset
   
   D.OrderBy = 1
   Call D.QueryData(1, Rs, ItemCount)
   
   If Not (C Is Nothing) Then
      C.Clear
      i = 0
      C.AddItem ("")
   End If

   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   
   While Not Rs.EOF
      i = i + 1
      Set TempData = New CXlsCarkillFW
      Call TempData.PopulateFromRS(1, Rs)
   
'      If Not (C Is Nothing) Then
'      End If
      
      If Not (Cl Is Nothing) Then
         Call Cl.Add(TempData)
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   Set Rs = Nothing
   Set D = Nothing
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub

Public Sub LoadAccountCode(C As ComboBox, Optional Cl As Collection = Nothing)
On Error GoTo ErrorHandler
Dim D As CAccountCode
Dim ItemCount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CAccountCode
Dim i As Long

   Set D = New CAccountCode
   Set Rs = New ADODB.Recordset
   
   D.OrderBy = 1
   Call D.QueryData(Rs, ItemCount)
   
   If Not (C Is Nothing) Then
      C.Clear
      i = 0
      C.AddItem ("")
   End If

   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   While Not Rs.EOF
      i = i + 1
      Set TempData = New CAccountCode
      Call TempData.PopulateFromRS(1, Rs)
   
      If Not (C Is Nothing) Then
      End If
      
      If Not (Cl Is Nothing) Then
         Call Cl.Add(TempData)
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   Set Rs = Nothing
   Set D = Nothing
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub
Public Sub InitRealCreditOrderBy(C As ComboBox)
   C.Clear

   C.AddItem ("")
   C.ItemData(0) = 0
   
   C.AddItem ("วันที่ส่งสินค้า")
   C.ItemData(1) = 1
   
   C.AddItem ("INVOICE")
   C.ItemData(2) = 2
   
   C.AddItem ("รหัสสินค้า")
   C.ItemData(3) = 3
   
End Sub

Public Sub InitIVcenterOrderBy(C As ComboBox)
   C.Clear

   C.AddItem ("")
   C.ItemData(0) = 0
   
   C.AddItem ("วันที่ส่งสินค้า")
   C.ItemData(1) = 1
   
   C.AddItem ("INVOICE")
   C.ItemData(2) = 2
   
   C.AddItem ("พนักงานขาย")
   C.ItemData(3) = 3
   
   C.AddItem ("เขตการขาย")
   C.ItemData(4) = 4
End Sub
Public Sub InitSetFarmOrderBy(C As ComboBox)
   C.Clear

   C.AddItem ("")
   C.ItemData(0) = 0
   
   C.AddItem ("ชื่อฟาร์ม")
   C.ItemData(1) = 1
   
   C.AddItem ("หน่วย")
   C.ItemData(2) = 2
   
End Sub

Public Sub InitFoodNumOrderBy(C As ComboBox)
   C.Clear

   C.AddItem ("")
   C.ItemData(0) = 0
   
   C.AddItem ("เบอร์อาหาร")
   C.ItemData(1) = 1
   
   C.AddItem ("หน่วย")
   C.ItemData(2) = 2
   
End Sub
Public Sub InitSupplierGroupOrderBy(C As ComboBox)
   C.Clear

   C.AddItem ("")
   C.ItemData(0) = 0
   
   C.AddItem ("รหัสเจ้าหนี้")
   C.ItemData(1) = 1
   
End Sub

Public Sub InitPromotionPayCustomerOrderBy(C As ComboBox)
   C.Clear

   C.AddItem ("")
   C.ItemData(0) = 0
   
   C.AddItem ("วันที่เอกสาร")
   C.ItemData(1) = 1
   
   C.AddItem ("Sale")
   C.ItemData(2) = 2
   
   C.AddItem ("ลูกค้า")
   C.ItemData(3) = 3
   
   C.AddItem ("สินค้า")
   C.ItemData(4) = 4
   
End Sub

Public Sub InitCostProductsOrderBy(C As ComboBox)
   C.Clear

   C.AddItem ("")
   C.ItemData(0) = 0
   
   C.AddItem ("รหัสสินค้า")
   C.ItemData(1) = 1
   
   C.AddItem ("ชื่อสินค้า")
   C.ItemData(2) = 2
   
End Sub

Public Sub InitPromotionYearOrderBy(C As ComboBox)
   C.Clear

   C.AddItem ("")
   C.ItemData(0) = 0
   
   C.AddItem ("วันที่เอกสาร")
   C.ItemData(1) = 1
   
   C.AddItem ("ลูกค้า")
   C.ItemData(2) = 2
   
   C.AddItem ("สินค้า")
   C.ItemData(3) = 3
   
End Sub
Public Sub LoadRealCreditNo(C As ComboBox, Optional Cl As Collection = Nothing, Optional GroupByDocument As Boolean = True)
On Error GoTo ErrorHandler
Dim D As CRealCredit
Dim ItemCount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CRealCredit
Dim i As Long

   Set D = New CRealCredit
   Set Rs = New ADODB.Recordset
   
   D.OrderBy = 1
   D.PAID_FLAG = False
   Call D.QueryData(Rs, ItemCount)
   
   If Not (C Is Nothing) Then
      C.Clear
      i = 0
      C.AddItem ("")
   End If

   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   While Not Rs.EOF
      i = i + 1
      Set TempData = New CRealCredit
      Call TempData.PopulateFromRS(1, Rs)
   
      If Not (C Is Nothing) Then
      End If
      
      If Not (Cl Is Nothing) Then
         If GroupByDocument Then
            If Len(Trim(TempData.DOCUMENT_NO)) > 0 Then
               Call Cl.Add(TempData, Trim(TempData.DOCUMENT_NO))
            End If
         Else
            If Len(Trim(TempData.DOCUMENT_NO)) <= 0 Then
               Call Cl.Add(TempData, Trim(TempData.CUSTOMER_CODE))
            End If
         End If
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   Set Rs = Nothing
   Set D = Nothing
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub
Public Sub LoadBktChqAmountByCus(C As ComboBox, Optional Cl As Collection = Nothing, Optional FromCus As String = "", Optional ToCus As String = "", Optional FromSale As String = "", Optional ToSale As String = "", Optional FromGet As Date = -1, Optional ToGet As Date = -1, Optional FromChq As Date = -1, Optional ToChq As Date = -1, Optional FromTrn As Date = -1)
On Error GoTo ErrorHandler
Dim D As CBkTrn
Dim ItemCount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CBkTrn
Dim i As Long
   
   Set D = New CBkTrn
   Set Rs = New ADODB.Recordset

   D.FROM_GETDAT = FromGet
   D.TO_GETDAT = ToGet
   D.FROM_CHQDAT = FromChq
   D.TO_CHQDAT = ToChq
  D.FROM_TRNDAT = FromTrn
   D.OrderBy = 1
   D.FROM_CUSTOMER_CODE = FromCus
   D.TO_CUSTOMER_CODE = ToCus
   D.FROM_SALE_CODE = FromSale
   D.TO_SALE_CODE = ToSale
   Call D.QueryData(9, Rs, ItemCount)
   
   If Not (C Is Nothing) Then
      C.Clear
      i = 0
      C.AddItem ("")
   End If

   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   While Not Rs.EOF
      i = i + 1
      Set TempData = New CBkTrn
      Call TempData.PopulateFromRS(9, Rs)
   
      If Not (C Is Nothing) Then
      End If
      
      If Not (Cl Is Nothing) Then
         Call Cl.Add(TempData, TempData.CHQNUM)   '!!?
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   Set Rs = Nothing
   Set D = Nothing
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub
Public Sub LoadBktChqAmountBySup(C As ComboBox, Optional Cl As Collection = Nothing, Optional FromGet As Date = -1, Optional ToGet As Date = -1, Optional FromChq As Date = -1, Optional ToChq As Date = -1)
On Error GoTo ErrorHandler
Dim D As CBkTrn
Dim ItemCount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CBkTrn
Dim i As Long
   
   Set D = New CBkTrn
   Set Rs = New ADODB.Recordset

   D.FROM_GETDAT = FromGet
   D.TO_GETDAT = ToGet
   D.FROM_CHQDAT = FromChq
   D.TO_CHQDAT = ToChq
   D.OrderBy = 1
   Call D.QueryData(2, Rs, ItemCount)
   
   If Not (C Is Nothing) Then
      C.Clear
      i = 0
      C.AddItem ("")
   End If

   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   While Not Rs.EOF
      i = i + 1
      Set TempData = New CBkTrn
      Call TempData.PopulateFromRS(2, Rs)
   
      If Not (C Is Nothing) Then
      End If
      
      If Not (Cl Is Nothing) Then
         Call Cl.Add(TempData, TempData.CUSCOD)
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   Set Rs = Nothing
   Set D = Nothing
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub
Public Sub LoadSupplier(C As ComboBox, Optional Cl As Collection = Nothing, Optional SupplierType As String = "", Optional SupplierCode As String = "", Optional AMPHUR As String = "", Optional PROVINCE As String = "")
On Error GoTo ErrorHandler
Dim D As CAPMas
Dim ItemCount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CAPMas
Dim i As Long
Dim SUPTYP As String

   If Val(SupplierType) <= 0 Then
      SUPTYP = ""
   Else
      SUPTYP = SupplierType
   End If
   
   Set D = New CAPMas
   Set Rs = New ADODB.Recordset
   
   D.SUPTYP = SUPTYP
   D.SUPCOD = SupplierCode
   D.AMPHUR = AMPHUR
   D.PROVINCE = PROVINCE
   Call D.QueryData(1, Rs, ItemCount)
   
   If Not (C Is Nothing) Then
      C.Clear
      i = 0
      C.AddItem ("")
   End If

   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   
   While Not Rs.EOF
      i = i + 1
      Set TempData = New CAPMas
      Call TempData.PopulateFromRS(1, Rs)

      If Not (C Is Nothing) Then
         C.AddItem (TempData.SUPNAM & "(" & TempData.SUPCOD & ")")
'         C.ItemData(i) = Val(TempData.TYPCOD)
      End If
      
'      If Left(TempData.SUPCOD, 1) = "อ" And Right(TempData.SUPCOD, 3) = "021" Then
'         'debug.print Mid(TempData.SUPCOD, 1, 1) = "ฟ"
'         'debug.print Mid(TempData.SUPCOD, 2, 1) = " "
'
'         'debug.print Mid(TempData.SUPCOD, 3, 1) = "-"
'         'debug.print Mid(TempData.SUPCOD, 4, 1) = "0"
'         'debug.print Mid(TempData.SUPCOD, 5, 1) = "1"
'         'debug.print Mid(TempData.SUPCOD, 6, 1) = "0"
'      End If
'      If Asc(Mid(TempData.SUPCOD, 2, 1)) = 160 Then        'ถ้าแก้ตรงนี้ต้องไปแก้ตอน Load ข้อมูลที่มีเจ้าหนี้ทุกอันเลย ซึ่งยุ่งยากมาก
'         'debug.print TempData.SUPCOD
'         ''debug.print Left(TempData.SUPCOD, 1) & " " & Mid(TempData.SUPCOD, 3)
'         TempData.SUPCOD = Left(TempData.SUPCOD, 1) & " " & Mid(TempData.SUPCOD, 3)
'      End If
      
      
      If Not (Cl Is Nothing) Then
         Call Cl.Add(TempData, Trim(TempData.SUPCOD))
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend

   Set Rs = Nothing
   Set D = Nothing
   Exit Sub

ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub
Public Sub LoadSupplierGroup(C As ComboBox, Optional Cl As Collection = Nothing, Optional SupplierType As String = "", Optional SupplierCode As String = "", Optional DataType As Long = -1, Optional GroupType As Long = -1)
On Error GoTo ErrorHandler
Dim D As CSupplierGroup
Dim ItemCount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CSupplierGroup
Dim i As Long
   
   Set D = New CSupplierGroup
   Set Rs = New ADODB.Recordset
   
   D.SUPPLIER_CODE = SupplierCode
   D.DATA_TYPE_ID = DataType
   D.GROUP_TYPE_CODE = GroupType
   Call D.QueryData(Rs, ItemCount)
   
   If Not (C Is Nothing) Then
      C.Clear
      i = 0
      C.AddItem ("")
   End If
   
   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   
   While Not Rs.EOF
      i = i + 1
      Set TempData = New CSupplierGroup
      Call TempData.PopulateFromRS(1, Rs)

      If Not (C Is Nothing) Then
         C.AddItem (TempData.SUPPLIER_CODE)
'         C.ItemData(i) = Val(TempData.TYPCOD)
      End If

      If Not (Cl Is Nothing) Then
         Call Cl.Add(TempData, TempData.SUPPLIER_CODE)
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend

   Set Rs = Nothing
   Set D = Nothing
   Exit Sub

ErrorHandler:
   glbErrorLog.SystemErrorMsg = "รหัส " & TempData.SUPPLIER_CODE & " ซ้ำ"
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub

Public Sub LoadComboSupGroup(C As ComboBox, Optional Cl As Collection = Nothing, Optional COMBO_SUB_ID As Long = -1)
On Error GoTo ErrorHandler
Dim D As CComboSubGroupDe
Dim ItemCount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CComboSubGroupDe
Dim i As Long
   
   Set D = New CComboSubGroupDe
   Set Rs = New ADODB.Recordset
   
   D.COMBO_SUB_ID = COMBO_SUB_ID
   Call D.QueryData(Rs, ItemCount)
   
   If Not (C Is Nothing) Then
      C.Clear
      i = 0
      C.AddItem ("")
   End If
   
   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   
   While Not Rs.EOF
      i = i + 1
      Set TempData = New CComboSubGroupDe
      Call TempData.PopulateFromRS(1, Rs)

      If Not (C Is Nothing) Then
         C.AddItem (TempData.GROUP_TYPE_CODE)
'         C.ItemData(i) = Val(TempData.TYPCOD)
      End If

      If Not (Cl Is Nothing) Then
         Call Cl.Add(TempData, Str(TempData.GROUP_TYPE_CODE))
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend

   Set Rs = Nothing
   Set D = Nothing
   Exit Sub

ErrorHandler:
   glbErrorLog.SystemErrorMsg = "รหัส " & TempData.GROUP_TYPE_CODE & " ซ้ำ"
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub


Public Sub LoadGroupType(C As ComboBox, Optional Cl As Collection = Nothing, Optional SupplierType As String = "")
On Error GoTo ErrorHandler
Dim D As CGroupType
Dim ItemCount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CGroupType
Dim i As Long
   
   Set D = New CGroupType
   Set Rs = New ADODB.Recordset
   
   Call D.QueryData(Rs, ItemCount)
   
   If Not (C Is Nothing) Then
      C.Clear
      i = 0
      C.AddItem ("")
   End If
   
   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   
   While Not Rs.EOF
      i = i + 1
      Set TempData = New CGroupType
      Call TempData.PopulateFromRS(1, Rs)

      If Not (C Is Nothing) Then
         C.AddItem (TempData.GROUP_TYPE_NAME)
         C.ItemData(i) = TempData.GROUP_TYPE_CODE
      End If

      If Not (Cl Is Nothing) Then
         Call Cl.Add(TempData, Trim(Str(TempData.GROUP_TYPE_CODE)))
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   If Not (Cl Is Nothing) Then
        Set TempData = New CGroupType
        TempData.GROUP_TYPE_CODE = 999999
        TempData.GROUP_TYPE_NAME = "อื่นๆ"
        Call Cl.Add(TempData, Trim(Str(TempData.GROUP_TYPE_CODE)))    'Code สำหรับ อื่นๆ
        Set TempData = Nothing
   End If
   
   Set Rs = Nothing
   Set D = Nothing
   Exit Sub

ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub
Public Sub LoadSubGroupType(C As ComboBox, Optional Cl As Collection = Nothing, Optional SupplierType As String = "")
On Error GoTo ErrorHandler
Dim D As CSubGroupType
Dim ItemCount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CSubGroupType
Dim i As Long
   
   Set D = New CSubGroupType
   Set Rs = New ADODB.Recordset
   
   Call D.QueryData(Rs, ItemCount)
   
   If Not (C Is Nothing) Then
      C.Clear
      i = 0
      C.AddItem ("")
   End If
   
   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   
   While Not Rs.EOF
      i = i + 1
      Set TempData = New CSubGroupType
      Call TempData.PopulateFromRS(1, Rs)

      If Not (C Is Nothing) Then
         C.AddItem (TempData.SUB_GROUP_TYPE_NAME)
         C.ItemData(i) = TempData.SUB_GROUP_TYPE_CODE
      End If

      If Not (Cl Is Nothing) Then
         Call Cl.Add(TempData, Trim(Str(TempData.SUB_GROUP_TYPE_CODE)))
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   If Not (Cl Is Nothing) Then
        Set TempData = New CSubGroupType
        TempData.SUB_GROUP_TYPE_CODE = 999999
        TempData.SUB_GROUP_TYPE_NAME = "อื่นๆ"
        Call Cl.Add(TempData, Trim(Str(TempData.SUB_GROUP_TYPE_CODE)))    'Code สำหรับ อื่นๆ
        Set TempData = Nothing
   End If
   
   Set Rs = Nothing
   Set D = Nothing
   Exit Sub

ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub
Public Sub LoadSubGroupTypeData(C As ComboBox, Optional Cl As Collection = Nothing, Optional SupplierType As String = "")
On Error GoTo ErrorHandler
Dim D As CSubGroupType
Dim ItemCount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CSubGroupType
Dim i As Long
   
   Set D = New CSubGroupType
   Set Rs = New ADODB.Recordset
   
   Call D.QueryData(Rs, ItemCount)
   
   If Not (C Is Nothing) Then
      C.Clear
      i = 0
      C.AddItem ("")
   End If
   
   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   
   While Not Rs.EOF
      i = i + 1
      Set TempData = New CSubGroupType
      Call TempData.PopulateFromRS(1, Rs)

      If Not (C Is Nothing) Then
         C.AddItem (TempData.SUB_GROUP_TYPE_NAME)
         C.ItemData(i) = Val(TempData.SUB_GROUP_TYPE_CODE)
      End If

      If Not (Cl Is Nothing) Then
         Call Cl.Add(TempData, Trim(Str(TempData.SUB_GROUP_TYPE_CODE)))
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   Set Rs = Nothing
   Set D = Nothing
   Exit Sub

ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub

Public Sub LoadGroupTypeData(C As ComboBox, Optional Cl As Collection = Nothing, Optional SupplierType As String = "")
On Error GoTo ErrorHandler
Dim D As CGroupType
Dim ItemCount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CGroupType
Dim i As Long
   
   Set D = New CGroupType
   Set Rs = New ADODB.Recordset
   
   Call D.QueryData(Rs, ItemCount)
   
   If Not (C Is Nothing) Then
      C.Clear
      i = 0
      C.AddItem ("")
   End If
   
   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   
   While Not Rs.EOF
      i = i + 1
      Set TempData = New CGroupType
      Call TempData.PopulateFromRS(1, Rs)

      If Not (C Is Nothing) Then
         C.AddItem (TempData.GROUP_TYPE_NAME)
         C.ItemData(i) = Val(TempData.GROUP_TYPE_CODE)
      End If

      If Not (Cl Is Nothing) Then
         Call Cl.Add(TempData, Trim(Str(TempData.GROUP_TYPE_CODE)))
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   Set Rs = Nothing
   Set D = Nothing
   Exit Sub

ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub
Public Sub LoadComboSupTypeData(C As ComboBox, Optional Cl As Collection = Nothing, Optional SupplierType As String = "")
On Error GoTo ErrorHandler
Dim D As CComboSubGroup
Dim ItemCount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CComboSubGroup
Dim i As Long
   
   Set D = New CComboSubGroup
   Set Rs = New ADODB.Recordset
   
   Call D.QueryData(1, Rs, ItemCount)
   
   If Not (C Is Nothing) Then
      C.Clear
      i = 0
      C.AddItem ("")
   End If
   
   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   
   While Not Rs.EOF
      i = i + 1
      Set TempData = New CComboSubGroup
      Call TempData.PopulateFromRS(1, Rs)

      If Not (C Is Nothing) Then
         C.AddItem (TempData.COMBO_SUB_NAME)
         C.ItemData(i) = Val(TempData.COMBO_SUB_ID)
      End If

      If Not (Cl Is Nothing) Then
         Call Cl.Add(TempData, Trim(Str(TempData.COMBO_SUB_ID)))
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   Set Rs = Nothing
   Set D = Nothing
   Exit Sub

ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub
Public Sub LoadBktChqnumDocDate(C As ComboBox, Optional Cl As Collection = Nothing, Optional FromGet As Date = -1, Optional ToGet As Date = -1, Optional FromChq As Date = -1, Optional ToChq As Date = -1)
On Error GoTo ErrorHandler
Dim D As CBkTrn
Dim ItemCount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CBkTrn
Dim i As Long
Dim TempBk As CBkTrn
   
   Set D = New CBkTrn
   Set Rs = New ADODB.Recordset
   
   D.FROM_GETDAT = FromGet
   D.TO_GETDAT = ToGet
   D.FROM_CHQDAT = FromChq
   D.TO_CHQDAT = ToChq
   D.OrderBy = 1
   Call D.QueryData(4, Rs, ItemCount)
   
   If Not (C Is Nothing) Then
      C.Clear
      i = 0
      C.AddItem ("")
   End If

   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   While Not Rs.EOF
      i = i + 1
      Set TempData = New CBkTrn
      Call TempData.PopulateFromRS(4, Rs)
      
      '''debug.print (Trim(TempData.CHQNUM & "-" & TempData.DOCDAT))
      
      If Not (C Is Nothing) Then
      End If
      
      If Not (Cl Is Nothing) Then
         Set TempBk = GetObject("CBkTrn", Cl, Trim(TempData.CHQNUM), False)
         If TempBk Is Nothing Then
            Call Cl.Add(TempData, Trim(TempData.CHQNUM))
         End If
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   Set Rs = Nothing
   Set D = Nothing
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub
Public Sub LoadBktChqnumAmountBySupCode(C As ComboBox, Optional Cl As Collection = Nothing, Optional FromChq As Date = -1, Optional ToChq As Date = -1, Optional ToPaidDate As Date = -1)
On Error GoTo ErrorHandler
Dim D As CBkTrn
Dim ItemCount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CBkTrn
Dim i As Long
Dim TempBk As CBkTrn
   
   Set D = New CBkTrn
   Set Rs = New ADODB.Recordset
   
   D.TO_PAY_DATE = ToPaidDate
   D.FROM_CHQDAT = FromChq
   D.TO_CHQDAT = ToChq
   D.OrderBy = 1
   Call D.QueryData(8, Rs, ItemCount)
   
   If Not (C Is Nothing) Then
      C.Clear
      i = 0
      C.AddItem ("")
   End If

   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   While Not Rs.EOF
      i = i + 1
      Set TempData = New CBkTrn
      Call TempData.PopulateFromRS(8, Rs)
      
      '''debug.print (Trim(TempData.CHQNUM & "-" & TempData.DOCDAT))
      
      If Not (C Is Nothing) Then
      End If
'      If (TempData.SUPCOD) = "ป-0002" Then
'         ''debug.print
'      End If
      If Not (Cl Is Nothing) Then
         'Set TempBk = GetObject("CBkTrn", Cl, Trim(TempData.SUPCOD), False)
         If TempBk Is Nothing Then
            Call Cl.Add(TempData, Trim(TempData.SUPCOD))
         End If
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   Set Rs = Nothing
   Set D = Nothing
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub

Public Sub LoadBktChqnumDocDateAR(C As ComboBox, Optional Cl As Collection = Nothing, Optional FromGet As Date = -1, Optional ToGet As Date = -1, Optional FromChq As Date = -1, Optional ToChq As Date = -1)
On Error GoTo ErrorHandler
Dim D As CBkTrn
Dim ItemCount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CBkTrn
Dim i As Long
Dim TempBk As CBkTrn
   
   Set D = New CBkTrn
   Set Rs = New ADODB.Recordset
   
   D.FROM_GETDAT = FromGet
   D.TO_GETDAT = ToGet
   D.FROM_CHQDAT = FromChq
   D.TO_CHQDAT = ToChq
   D.OrderBy = 1
   Call D.QueryData(7, Rs, ItemCount)
   
   If Not (C Is Nothing) Then
      C.Clear
      i = 0
      C.AddItem ("")
   End If

   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   While Not Rs.EOF
      i = i + 1
      Set TempData = New CBkTrn
      Call TempData.PopulateFromRS(7, Rs)
      
      '''debug.print (Trim(TempData.CHQNUM & "-" & TempData.DOCDAT))
      
      If Not (C Is Nothing) Then
      End If
      
      If Not (Cl Is Nothing) Then
         Set TempBk = GetObject("CBkTrn", Cl, Trim(TempData.CHQNUM), False)
         If TempBk Is Nothing Then
            ''debug.print (TempData.CHQNUM)
            Call Cl.Add(TempData, Trim(TempData.CHQNUM))
         End If
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   Set Rs = Nothing
   Set D = Nothing
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub

Public Sub LoadAllDocumentCancel(C As ComboBox, Optional Cl As Collection = Nothing)
On Error GoTo ErrorHandler
Dim D As CDocumentCancel
Dim ItemCount As Long
Dim Rs As ADODB.Recordset
Dim i As Long
      
   Set D = New CDocumentCancel
   Set Rs = New ADODB.Recordset
   
   Call D.QueryData(1, Rs, ItemCount)
   
   If Not (C Is Nothing) Then
      C.Clear
      i = 0
      C.AddItem ("")
   End If

   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   While Not Rs.EOF
      i = i + 1
      Set D = New CDocumentCancel
      Call D.PopulateFromRS(1, Rs)
      
      '''debug.print (Trim(TempData.CHQNUM & "-" & TempData.DOCDAT))
      
      If Not (C Is Nothing) Then
      End If
      
      If Not (Cl Is Nothing) Then
         Call Cl.Add(D, Trim(D.DOCUMENT_NO))
      End If
      
      Set D = Nothing
      Rs.MoveNext
   Wend
   
   Set Rs = Nothing
   Set D = Nothing
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub
Public Sub LoadDistinctDocMonthYearFromReceipt(C As ComboBox, Optional Cl As Collection = Nothing, Optional FromDate As Date, Optional ToDate As Date, Optional RecTypeSet As String, Optional FromCustomerCode As String, Optional ToCustomerCode As String)
On Error GoTo ErrorHandler
Dim D As CARRcIt
Dim ItemCount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CARRcIt
Dim TempData1 As CARRcIt
Dim i As Long

   Set D = New CARRcIt
   Set Rs = New ADODB.Recordset
   
   D.FROM_PAY_DATE = FromDate
   D.TO_PAY_DATE = ToDate
   D.FROM_CUSTOMER_CODE = FromCustomerCode
   D.TO_CUSTOMER_CODE = ToCustomerCode
   D.RecTypeSet = RecTypeSet
   Call D.QueryData(6, Rs, ItemCount)
   
   If Not (C Is Nothing) Then
      C.Clear
      i = 0
      C.AddItem ("")
   End If

   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   While Not Rs.EOF
      i = i + 1
      Set TempData = New CARRcIt
      Call TempData.PopulateFromRS(6, Rs)
      
      If Not (Cl Is Nothing) Then
         Set TempData1 = GetObject("CARRcIt", Cl, Year(TempData.DOCDAT) & "-" & Format(Month(TempData.DOCDAT), "00"), False)
         If TempData1 Is Nothing Then
            '''debug.print Year(TempData.DOCDAT) & "-" & Format(Month(TempData.DOCDAT))
            Call Cl.Add(TempData, Year(TempData.DOCDAT) & "-" & Format(Month(TempData.DOCDAT), "00"))
         End If
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   Set Rs = Nothing
   Set D = Nothing
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub
Public Sub LoadSumMonthYearFromReceipt(C As ComboBox, Optional Cl As Collection = Nothing, Optional FromDate As Date, Optional ToDate As Date, Optional RecTypeSet As String, Optional FromCustomerCode As String, Optional ToCustomerCode As String)
On Error GoTo ErrorHandler
Dim D As CARRcIt
Dim ItemCount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CARRcIt
Dim TempData1 As CARRcIt
Dim i As Long

   Set D = New CARRcIt
   Set Rs = New ADODB.Recordset
   
   D.FROM_PAY_DATE = FromDate
   D.TO_PAY_DATE = ToDate
   D.FROM_CUSTOMER_CODE = FromCustomerCode
   D.TO_CUSTOMER_CODE = ToCustomerCode
   D.RecTypeSet = RecTypeSet
   Call D.QueryData(7, Rs, ItemCount)
   
   If Not (C Is Nothing) Then
      C.Clear
      i = 0
      C.AddItem ("")
   End If

   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   While Not Rs.EOF
      i = i + 1
      Set TempData = New CARRcIt
      Call TempData.PopulateFromRS(7, Rs)
      
      If Not (Cl Is Nothing) Then
         Set TempData1 = GetObject("CARRcIt", Cl, Year(TempData.PAYDAT) & "-" & Format(Month(TempData.PAYDAT), "00") & "-" & Year(TempData.DOCDAT) & "-" & Format(Month(TempData.DOCDAT), "00"), False)
         If TempData1 Is Nothing Then
            '''debug.print Year(TempData.PAYDAT) & "-" & Format(Month(TempData.PAYDAT), "00") & "-" & Year(TempData.DOCDAT) & "-" & Format(Month(TempData.DOCDAT), "00")
            Call Cl.Add(TempData, Year(TempData.PAYDAT) & "-" & Format(Month(TempData.PAYDAT), "00") & "-" & Year(TempData.DOCDAT) & "-" & Format(Month(TempData.DOCDAT), "00"))
         Else
            '''debug.print Year(TempData.PAYDAT) & "-" & Format(Month(TempData.PAYDAT), "00") & "-" & Year(TempData.DOCDAT) & "-" & Format(Month(TempData.DOCDAT), "00")
            TempData1.RCVAMT = TempData1.RCVAMT + TempData.RCVAMT
         End If
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   Set Rs = Nothing
   Set D = Nothing
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub

Public Sub LoadPromotionFree(Optional Cl As Collection = Nothing, Optional FromDate As Date, Optional ToDate As Date, Optional FromCustomerCode As String, Optional ToCustomerCode As String, Optional FromStockCode As String, Optional ToStockCode As String, Optional FromSaleCode As String, Optional ToSaleCode As String)
On Error GoTo ErrorHandler
Dim D As CPromotionConfig
Dim ItemCount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CPromotionConfig
Dim i As Long
Dim Sum As Double

   Set D = New CPromotionConfig
   Set Rs = New ADODB.Recordset
   
   D.FROM_PRO_DATE = FromDate
   D.TO_PRO_DATE = ToDate
   Call D.QueryData(1, Rs, ItemCount)
   
   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   While Not Rs.EOF
      i = i + 1
      Set TempData = New CPromotionConfig
      
      Call TempData.PopulateFromRS(1, Rs)
      
      If Not (Cl Is Nothing) Then
         Call Cl.Add(TempData, Trim(Format(TempData.MONTH_NUM, "00")) & "-" & Trim(Format(Val(TempData.YEAR_NUM) + 543, "0000")) & "-" & Trim(TempData.CUSTOMER_CODE) & "-" & Trim(TempData.STKCOD))
    End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   Set Rs = Nothing
   Set D = Nothing
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub

Public Sub LoadCostProducts(Optional Cl As Collection = Nothing, Optional FromDate As Date, Optional ToDate As Date, Optional FromCustomerCode As String, Optional ToCustomerCode As String, Optional FromStockCode As String, Optional ToStockCode As String, Optional FromSaleCode As String, Optional ToSaleCode As String)
On Error GoTo ErrorHandler
Dim D As CCostProducts
Dim ItemCount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CCostProducts
Dim i As Long
Dim Sum As Double

   Set D = New CCostProducts
   Set Rs = New ADODB.Recordset
   
'   D.FROM_PRO_DATE = FromDate
'   D.TO_PRO_DATE = ToDate
   Call D.QueryData(1, Rs, ItemCount)
   
   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   While Not Rs.EOF
      i = i + 1
      Set TempData = New CCostProducts
      
      Call TempData.PopulateFromRS(1, Rs)
      
      If Not (Cl Is Nothing) Then
         Call Cl.Add(TempData, Trim(TempData.PRODUCT_CODE))
    End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   Set Rs = Nothing
   Set D = Nothing
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub

Public Sub LoadPromotionYear(Optional Cl As Collection = Nothing, Optional FromDate As Date, Optional ToDate As Date, Optional FromCustomerCode As String, Optional ToCustomerCode As String, Optional FromStockCode As String, Optional ToStockCode As String, Optional FromSaleCode As String, Optional ToSaleCode As String)
On Error GoTo ErrorHandler
Dim D As CPromotionYear
Dim ItemCount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CPromotionYear
Dim i As Long
Dim Sum As Double

   Set D = New CPromotionYear
   Set Rs = New ADODB.Recordset
   
   D.FROM_PRO_DATE = FromDate
   D.TO_PRO_DATE = ToDate
   Call D.QueryData(2, Rs, ItemCount)
   
   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   While Not Rs.EOF
      i = i + 1
      Set TempData = New CPromotionYear
      
      Call TempData.PopulateFromRS(2, Rs)
   
      If Not (Cl Is Nothing) Then
         Call Cl.Add(TempData, Trim(TempData.CTMCODYEAR_PRO) & "-" & Trim(TempData.STKCODYEAR_PRO) & "-" & Trim(TempData.YYYY_MM))
    End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   Set Rs = Nothing
   Set D = Nothing
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub
Public Sub LoadStockPro(Cl As Collection)
On Error GoTo ErrorHandler
Dim D As CStmas
Dim ItemCount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CStmas
Dim i As Long
   
   Set D = New CStmas
   Set Rs = New ADODB.Recordset
   Call D.QueryData(1, Rs, ItemCount)
   
   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   
   While Not Rs.EOF
      i = i + 1
      Set TempData = New CStmas
      Call TempData.PopulateFromRS(1, Rs)

      If Not (Cl Is Nothing) Then
         Call Cl.Add(TempData, Trim(TempData.STKCOD))
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   Set Rs = Nothing
   Set D = Nothing
   Exit Sub

ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub

Public Sub LoadCustomerStcrd(Optional Cl As Collection = Nothing, Optional FromDate As Date, Optional ToDate As Date, Optional FromCustomerCode As String, Optional ToCustomerCode As String, Optional FromStockCode As String, Optional ToStockCode As String, Optional FromSaleCode As String, Optional ToSaleCode As String)
On Error GoTo ErrorHandler
Dim D As CStcrd
Dim ItemCount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CStcrd
Dim i As Long

   Set D = New CStcrd
   Set Rs = New ADODB.Recordset
   
   D.FROM_DOC_DATE = FromDate
   D.TO_DOC_DATE = ToDate
   D.FROM_CUSTOMER_CODE = FromCustomerCode
   D.TO_CUSTOMER_CODE = ToCustomerCode
   D.FROM_STOCK_CODE = FromStockCode
   D.TO_STOCK_CODE = ToStockCode
   D.FROM_SALE_CODE = FromSaleCode
   D.TO_SALE_CODE = ToSaleCode
   Call D.QueryData(3, Rs, ItemCount)
   
   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   While Not Rs.EOF
      i = i + 1
      Set TempData = New CStcrd
      Call TempData.PopulateFromRS(3, Rs)
      
      If Not (Cl Is Nothing) Then
         Call Cl.Add(TempData)
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   Set Rs = Nothing
   Set D = Nothing
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub
Public Sub LoadDistinctCustomerStcrd(Optional Cl As Collection = Nothing, Optional FromDate As Date, Optional ToDate As Date, Optional FromCustomerCode As String, Optional ToCustomerCode As String, Optional FromStockCode As String, Optional ToStockCode As String)
On Error GoTo ErrorHandler
Dim D As CStcrd
Dim ItemCount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CStcrd
Dim i As Long

   Set D = New CStcrd
   Set Rs = New ADODB.Recordset
   
   D.FROM_DOC_DATE = FromDate
   D.TO_DOC_DATE = ToDate
   D.FROM_CUSTOMER_CODE = FromCustomerCode
   D.TO_CUSTOMER_CODE = ToCustomerCode
   D.FROM_STOCK_CODE = FromStockCode
   D.TO_STOCK_CODE = ToStockCode
   Call D.QueryData(4, Rs, ItemCount)
   
   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   While Not Rs.EOF
      i = i + 1
      Set TempData = New CStcrd
      Call TempData.PopulateFromRS(4, Rs)
      
      If Not (Cl Is Nothing) Then
         Set D = GetObject("CStcrd", Cl, Trim(TempData.STKCOD & "-" & TempData.CUSCOD), False)
         If D Is Nothing Then
            Call Cl.Add(TempData, Trim(TempData.STKCOD & "-" & TempData.CUSCOD))
         End If
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   Set Rs = Nothing
   Set D = Nothing
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub
Public Sub LoadDistinctSaleCustomerStcrd(Optional Cl As Collection = Nothing, Optional FromDate As Date, Optional ToDate As Date, Optional FromCustomerCode As String, Optional ToCustomerCode As String, Optional FromStockCode As String, Optional ToStockCode As String, Optional FromSaleCode As String, Optional ToSaleCode As String, Optional OrderBy As Long = -1)
On Error GoTo ErrorHandler
Dim D As CStcrd
Dim ItemCount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CStcrd
Dim i As Long

   Set D = New CStcrd
   Set Rs = New ADODB.Recordset
   
   D.FROM_DOC_DATE = FromDate
   D.TO_DOC_DATE = ToDate
   D.FROM_CUSTOMER_CODE = FromCustomerCode
   D.TO_CUSTOMER_CODE = ToCustomerCode
   D.FROM_STOCK_CODE = FromStockCode
   D.TO_STOCK_CODE = ToStockCode
   D.FROM_SALE_CODE = FromSaleCode
   D.TO_SALE_CODE = ToSaleCode
   D.OrderBy = OrderBy
   Call D.QueryData(7, Rs, ItemCount)
   
   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   While Not Rs.EOF
      i = i + 1
      Set TempData = New CStcrd
      Call TempData.PopulateFromRS(7, Rs)
      
      If Not (Cl Is Nothing) Then
         Set D = GetObject("CStcrd", Cl, Trim(TempData.STKCOD & "-" & TempData.CUSCOD & "-" & TempData.SLMCOD), False)
         If D Is Nothing Then
            Call Cl.Add(TempData, Trim(TempData.STKCOD & "-" & TempData.CUSCOD & "-" & TempData.SLMCOD))
         End If
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   Set Rs = Nothing
   Set D = Nothing
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub

Public Sub LoadDistinctSaleCustomerStockStcrd(Optional Cl As Collection = Nothing, Optional FromDate As Date, Optional ToDate As Date, Optional FromSaleCode As String, Optional ToSaleCode As String, Optional FromCustomerCode As String, Optional ToCustomerCode As String, Optional FromStockCode As String, Optional ToStockCode As String, Optional OrderBy As Long = -1)
On Error GoTo ErrorHandler
Dim D As CPromotionPayCustom
Dim ItemCount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CPromotionPayCustom
Dim i As Long

   Set D = New CPromotionPayCustom
   Set Rs = New ADODB.Recordset
   
   D.FROM_PRO_DATE = FromDate
   D.TO_PRO_DATE = ToDate
   D.FROM_SALE_CODE = FromSaleCode
   D.TO_SALE_CODE = ToSaleCode
   D.FROM_CUSTOMER_CODE = FromCustomerCode
   D.TO_CUSTOMER_CODE = ToCustomerCode
   D.FROM_STOCK_CODE = FromStockCode
   D.TO_STOCK_CODE = ToStockCode
   D.OrderBy = OrderBy
   Call D.QueryData(2, Rs, ItemCount)
   
   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   While Not Rs.EOF
      i = i + 1
      Set TempData = New CPromotionPayCustom
      Call TempData.PopulateFromRS(2, Rs)
      
      If Not (Cl Is Nothing) Then
         Set D = GetObject("CPromotionPayCustom", Cl, Trim(TempData.SALECODE_PRO & "-" & TempData.CUSTOMERCODE_PRO & "-" & TempData.STKCOD_PRO), False)
         If D Is Nothing Then
            Call Cl.Add(TempData, Trim(TempData.SALECODE_PRO & "-" & TempData.CUSTOMERCODE_PRO & "-" & TempData.STKCOD_PRO))
         End If
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   Set Rs = Nothing
   Set D = Nothing
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub

Public Sub LoadDistinctStockCustomerStcrd(Optional Cl As Collection = Nothing, Optional FromDate As Date, Optional ToDate As Date, Optional FromStockCode As String, Optional ToStockCode As String, Optional FromCustomerCode As String, Optional ToCustomerCode As String, Optional OrderBy As Long = -1)
On Error GoTo ErrorHandler
Dim D As CPromotionPayCustom
Dim ItemCount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CPromotionPayCustom
Dim i As Long

   Set D = New CPromotionPayCustom
   Set Rs = New ADODB.Recordset
   
   D.FROM_PRO_DATE = FromDate
   D.TO_PRO_DATE = ToDate
   D.FROM_STOCK_CODE = FromStockCode
   D.TO_STOCK_CODE = ToStockCode
   D.FROM_CUSTOMER_CODE = FromCustomerCode
   D.TO_CUSTOMER_CODE = ToCustomerCode
   D.OrderBy = OrderBy
   Call D.QueryData(4, Rs, ItemCount)
   
   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   While Not Rs.EOF
      i = i + 1
      Set TempData = New CPromotionPayCustom
      Call TempData.PopulateFromRS(4, Rs)
      
      If Not (Cl Is Nothing) Then
         Set D = GetObject("CPromotionPayCustom", Cl, Trim(TempData.STKCOD_PRO & "-" & TempData.CUSTOMERCODE_PRO), False)
         If D Is Nothing Then
            Call Cl.Add(TempData, Trim(TempData.STKCOD_PRO & "-" & TempData.CUSTOMERCODE_PRO))
         End If
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   Set Rs = Nothing
   Set D = Nothing
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub

Public Sub LoadDistinctSaleCustomerSLMCOD(Optional Cl As Collection = Nothing, Optional FromDate As Date, Optional ToDate As Date, Optional FromCustomerCode As String, Optional ToCustomerCode As String, Optional FromStockCode As String, Optional ToStockCode As String, Optional FromSaleCode As String, Optional ToSaleCode As String)
On Error GoTo ErrorHandler
Dim D As CStcrd
Dim ItemCount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CStcrd
Dim i As Long

   Set D = New CStcrd
   Set Rs = New ADODB.Recordset
   
   D.FROM_DOC_DATE = FromDate
   D.TO_DOC_DATE = ToDate
   D.FROM_CUSTOMER_CODE = FromCustomerCode
   D.TO_CUSTOMER_CODE = ToCustomerCode
   D.FROM_STOCK_CODE = FromStockCode
   D.TO_STOCK_CODE = ToStockCode
   D.FROM_SALE_CODE = FromSaleCode
   D.TO_SALE_CODE = ToSaleCode
   Call D.QueryData(20, Rs, ItemCount)
   
   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   While Not Rs.EOF
      i = i + 1
      Set TempData = New CStcrd
      Call TempData.PopulateFromRS(20, Rs)
      
      If Not (Cl Is Nothing) Then
         Set D = GetObject("CStcrd", Cl, Trim(TempData.CUSCOD), False)
         If D Is Nothing Then
            Call Cl.Add(TempData, Trim(TempData.CUSCOD))
         End If
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   Set Rs = Nothing
   Set D = Nothing
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub
Public Sub LoadDistinctSaleCustomerStcrdCode(Optional Cl As Collection = Nothing, Optional FromDate As Date, Optional ToDate As Date, Optional FromCustomerCode As String, Optional ToCustomerCode As String, Optional FromStockCode As String, Optional ToStockCode As String, Optional FromSaleCode As String, Optional ToSaleCode As String)
On Error GoTo ErrorHandler
Dim D As CStcrd
Dim ItemCount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CStcrd
Dim i As Long

   Set D = New CStcrd
   Set Rs = New ADODB.Recordset
   
   D.FROM_DOC_DATE = FromDate
   D.TO_DOC_DATE = ToDate
   D.FROM_CUSTOMER_CODE = FromCustomerCode
   D.TO_CUSTOMER_CODE = ToCustomerCode
   D.FROM_STOCK_CODE = FromStockCode
   D.TO_STOCK_CODE = ToStockCode
   D.FROM_SALE_CODE = FromSaleCode
   D.TO_SALE_CODE = ToSaleCode
   Call D.QueryData(22, Rs, ItemCount)
   
   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   While Not Rs.EOF
      i = i + 1
      Set TempData = New CStcrd
      Call TempData.PopulateFromRS(22, Rs)
      
      If Not (Cl Is Nothing) Then
         Set D = GetObject("CStcrd", Cl, Trim(TempData.STKCOD), False)
         If D Is Nothing Then
            Call Cl.Add(TempData, Trim(TempData.STKCOD))
         End If
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   Set Rs = Nothing
   Set D = Nothing
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub
Public Sub LoadSaleCustomerStcrd(Optional Cl As Collection = Nothing, Optional FromDate As Date, Optional ToDate As Date, Optional FromCustomerCode As String, Optional ToCustomerCode As String, Optional FromStockCode As String, Optional ToStockCode As String, Optional FromSaleCode As String, Optional ToSaleCode As String)
On Error GoTo ErrorHandler
Dim D As CStcrd
Dim ItemCount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CStcrd
Dim i As Long

   Set D = New CStcrd
   Set Rs = New ADODB.Recordset
   
   D.FROM_DOC_DATE = FromDate
   D.TO_DOC_DATE = ToDate
   D.FROM_CUSTOMER_CODE = FromCustomerCode
   D.TO_CUSTOMER_CODE = ToCustomerCode
   D.FROM_STOCK_CODE = FromStockCode
   D.TO_STOCK_CODE = ToStockCode
   D.FROM_SALE_CODE = FromSaleCode
   D.TO_SALE_CODE = ToSaleCode
   Call D.QueryData(6, Rs, ItemCount)
   
   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   While Not Rs.EOF
      i = i + 1
      Set TempData = New CStcrd
      Call TempData.PopulateFromRS(6, Rs)
      
      If Not (Cl Is Nothing) Then
         Call Cl.Add(TempData)
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   Set Rs = Nothing
   Set D = Nothing
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub
Public Sub LoadCostProductStcrd(Optional Cl As Collection = Nothing, Optional FromDate As Date, Optional ToDate As Date, Optional FromCustomerCode As String, Optional ToCustomerCode As String, Optional FromStockCode As String, Optional ToStockCode As String, Optional FromSaleCode As String, Optional ToSaleCode As String, Optional OrderBy As Long = -1)
On Error GoTo ErrorHandler
Dim D As CStcrd
Dim ItemCount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CStcrd
Dim i As Long

   Set D = New CStcrd
   Set Rs = New ADODB.Recordset
   
   D.FROM_DOC_DATE = FromDate
   D.TO_DOC_DATE = ToDate
   D.FROM_CUSTOMER_CODE = FromCustomerCode
   D.TO_CUSTOMER_CODE = ToCustomerCode
   D.FROM_STOCK_CODE = FromStockCode
   D.TO_STOCK_CODE = ToStockCode
   D.FROM_SALE_CODE = FromSaleCode
   D.TO_SALE_CODE = ToSaleCode
   D.OrderBy = OrderBy
   Call D.QueryData(27, Rs, ItemCount)
   
   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   While Not Rs.EOF
      i = i + 1
      Set TempData = New CStcrd
      Call TempData.PopulateFromRS(27, Rs)
      
      If Not (Cl Is Nothing) Then
         Call Cl.Add(TempData)
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   Set Rs = Nothing
   Set D = Nothing
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub
Public Sub LoadSaleCustomerStockStcrd(Optional Cl As Collection = Nothing, Optional FromDate As Date, Optional ToDate As Date, Optional FromSaleCode As String, Optional ToSaleCode As String, Optional FromCustomerCode As String, Optional ToCustomerCode As String, Optional FromStockCode As String, Optional ToStockCode As String)
On Error GoTo ErrorHandler
Dim D As CPromotionPayCustom
Dim ItemCount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CPromotionPayCustom
Dim i As Long

   Set D = New CPromotionPayCustom
   Set Rs = New ADODB.Recordset
   
   D.FROM_PRO_DATE = FromDate
   D.TO_PRO_DATE = ToDate
   D.FROM_SALE_CODE = FromSaleCode
   D.TO_SALE_CODE = ToSaleCode
   D.FROM_CUSTOMER_CODE = FromCustomerCode
   D.TO_CUSTOMER_CODE = ToCustomerCode
   D.FROM_STOCK_CODE = FromStockCode
   D.TO_STOCK_CODE = ToStockCode
   Call D.QueryData(3, Rs, ItemCount)
   
   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   While Not Rs.EOF
      i = i + 1
      Set TempData = New CPromotionPayCustom
      Call TempData.PopulateFromRS(3, Rs)
      
      If Not (Cl Is Nothing) Then
         Call Cl.Add(TempData, Trim(TempData.SALECODE_PRO & "-" & TempData.CUSTOMERCODE_PRO & "-" & TempData.STKCOD_PRO & "-" & TempData.YYYY_MM))
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   Set Rs = Nothing
   Set D = Nothing
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub

Public Sub LoadStockCustomerStcrd(Optional Cl As Collection = Nothing, Optional FromDate As Date, Optional ToDate As Date, Optional FromStockCode As String, Optional ToStockCode As String, Optional FromCustomerCode As String, Optional ToCustomerCode As String)
On Error GoTo ErrorHandler
Dim D As CPromotionPayCustom
Dim ItemCount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CPromotionPayCustom
Dim i As Long

   Set D = New CPromotionPayCustom
   Set Rs = New ADODB.Recordset
   
   D.FROM_PRO_DATE = FromDate
   D.TO_PRO_DATE = ToDate
   D.FROM_STOCK_CODE = FromStockCode
   D.TO_STOCK_CODE = ToStockCode
   D.FROM_CUSTOMER_CODE = FromCustomerCode
   D.TO_CUSTOMER_CODE = ToCustomerCode
   Call D.QueryData(5, Rs, ItemCount)
   
   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   While Not Rs.EOF
      i = i + 1
      Set TempData = New CPromotionPayCustom
      Call TempData.PopulateFromRS(5, Rs)
      
      If Not (Cl Is Nothing) Then
         Call Cl.Add(TempData, Trim(TempData.STKCOD_PRO & "-" & TempData.CUSTOMERCODE_PRO & "-" & TempData.YYYY_MM))
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   Set Rs = Nothing
   Set D = Nothing
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub
Public Sub LoadDistinctStcrdCustomer(Optional Cl As Collection = Nothing, Optional FromDate As Date, Optional ToDate As Date, Optional FromCustomerCode As String, Optional ToCustomerCode As String, Optional FromStockCode As String, Optional ToStockCode As String, Optional FromSaleCode As String, Optional ToSaleCode As String)
On Error GoTo ErrorHandler
Dim D As CStcrd
Dim ItemCount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CStcrd
Dim i As Long

   Set D = New CStcrd
   Set Rs = New ADODB.Recordset
   
   D.FROM_DOC_DATE = FromDate
   D.TO_DOC_DATE = ToDate
   D.FROM_CUSTOMER_CODE = FromCustomerCode
   D.TO_CUSTOMER_CODE = ToCustomerCode
   D.FROM_STOCK_CODE = FromStockCode
   D.TO_STOCK_CODE = ToStockCode
   D.FROM_SALE_CODE = FromSaleCode
   D.TO_SALE_CODE = ToSaleCode
   Call D.QueryData(5, Rs, ItemCount)
   
   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   While Not Rs.EOF
      i = i + 1
      Set TempData = New CStcrd
      Call TempData.PopulateFromRS(5, Rs)
      
      If Not (Cl Is Nothing) Then
         Set D = GetObject("CStcrd", Cl, Trim(TempData.STKCOD & "-" & TempData.CUSCOD), False)
         If D Is Nothing Then
            Call Cl.Add(TempData, Trim(TempData.STKCOD & "-" & TempData.CUSCOD))
         End If
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   Set Rs = Nothing
   Set D = Nothing
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub

Public Sub LoadIVinDateStcrd(Optional Cl As Collection = Nothing, Optional FromDate As Date = -1, Optional ToDate As Date = -1)
On Error GoTo ErrorHandler
Dim D As CStcrd
Dim ItemCount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CStcrd
Dim i As Long

   Set D = New CStcrd
   Set Rs = New ADODB.Recordset
   
   D.FROM_DOC_DATE = FromDate
   D.TO_DOC_DATE = ToDate

   Call D.QueryData(10, Rs, ItemCount)
   
   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   While Not Rs.EOF
      i = i + 1
      Set TempData = New CStcrd
      Call TempData.PopulateFromRS(10, Rs)
      
      If Not (Cl Is Nothing) Then
         Call Cl.Add(TempData, Trim(TempData.DOCNUM))
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   Set Rs = Nothing
   Set D = Nothing
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub

Public Sub LoadIVExStcrd(Optional Cl As Collection = Nothing, Optional FromDate As Date = -1, Optional ToDate As Date = -1, Optional FromSLM As String = -1, Optional ToSLM As String = -1)
On Error GoTo ErrorHandler
Dim D As CStcrd
Dim ItemCount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CStcrd
Dim i As Long

   Set D = New CStcrd
   Set Rs = New ADODB.Recordset
   
   D.FROM_DOC_DATE = FromDate
   D.TO_DOC_DATE = ToDate
   D.FROM_SALE_CODE = FromSLM
   D.TO_SALE_CODE = ToSLM

   Call D.QueryData(18, Rs, ItemCount)
   
   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   While Not Rs.EOF
      i = i + 1
      Set TempData = New CStcrd
      Call TempData.PopulateFromRS(18, Rs)
      
      If Not (Cl Is Nothing) Then
         Call Cl.Add(TempData, Trim(TempData.DOCNUM))
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   Set Rs = Nothing
   Set D = Nothing
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub


'Public Function LoadEnterPriseName() As String
'On Error GoTo ErrorHandler
'Dim D As CIsInfo
'Dim ItemCount As Long
'Dim Rs As ADODB.Recordset
'Dim TempData As CIsInfo
'Dim I As Long
'
'   Set D = New CIsInfo
'   Set Rs = New ADODB.Recordset
'
'   Call D.QueryData(1, Rs, ItemCount)
'
'   While Not Rs.EOF
'      I = I + 1
'      Set TempData = New CIsInfo
'      Call TempData.PopulateFromRS(1, Rs)
'
'      LoadEnterPriseName = TempData.THINAM
'
'      Set TempData = Nothing
'      Rs.MoveNext
'   Wend
'
'   Set Rs = Nothing
'   Set D = Nothing
'   Exit Function
'
'ErrorHandler:
'   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
'   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
'End Function
Public Sub LoadARAmountByCust3(C As ComboBox, Optional Cl As Collection = Nothing, Optional FromDate As Date, Optional ToDate As Date, Optional FromInvDate As Date, Optional ToInvDate As Date)
On Error GoTo ErrorHandler
Dim D As CARTrn
Dim ItemCount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CARTrn
Dim i As Long

   Set D = New CARTrn
   Set Rs = New ADODB.Recordset

   D.FROM_DOC_DATE = FromDate
   D.TO_DOC_DATE = ToDate
   D.FROM_INV_DATE = FromInvDate
   D.TO_INV_DATE = ToInvDate
   D.OrderBy = 1
   Call D.QueryData(7, Rs, ItemCount)
   
   If Not (C Is Nothing) Then
      C.Clear
      i = 0
      C.AddItem ("")
   End If
   
   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   While Not Rs.EOF
      i = i + 1
      Set TempData = New CARTrn
      Call TempData.PopulateFromRS(7, Rs)
      
      If Not (C Is Nothing) Then
      End If
      
      If Not (Cl Is Nothing) Then
         Call Cl.Add(TempData, TempData.CUSCOD)
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   Set Rs = Nothing
   Set D = Nothing
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub
Public Sub LoadARCNAmountByBill(C As ComboBox, Optional Cl As Collection = Nothing, Optional FromDate As Date, Optional ToDate As Date, Optional FromInvDate As Date, Optional ToInvDate As Date, Optional db As Long = 1)
On Error GoTo ErrorHandler
Dim D As CARTrn
Dim ItemCount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CARTrn
Dim i As Long
   
   Set D = New CARTrn
   Set Rs = New ADODB.Recordset

   D.FROM_DOC_DATE = FromDate
   D.TO_DOC_DATE = ToDate
   D.FROM_INV_DATE = FromInvDate
   D.TO_INV_DATE = ToInvDate
   D.OrderBy = 1
   Call D.QueryData(8, Rs, ItemCount, False, db)

   
   If Not (C Is Nothing) Then
      C.Clear
      i = 0
      C.AddItem ("")
   End If
   
   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   While Not Rs.EOF
      i = i + 1
      Set TempData = New CARTrn
      Call TempData.PopulateFromRS(8, Rs)
      
      If Not (C Is Nothing) Then
      End If
      
      If Not (Cl Is Nothing) Then
         Call Cl.Add(TempData, TempData.DOCNUM)
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   Set Rs = Nothing
   Set D = Nothing
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub
Public Sub LoadARCNAmount(C As ComboBox, Optional Cl As Collection = Nothing, Optional FromDate As Date, Optional ToDate As Date, Optional FromInvDate As Date, Optional ToInvDate As Date, Optional FromCustomerCode As String, Optional ToCustomerCode As String)
On Error GoTo ErrorHandler
Dim D As CARTrn
Dim ItemCount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CARTrn
Dim i As Long
Dim TempData1   As CARTrn
   
   Set D = New CARTrn
   Set Rs = New ADODB.Recordset

   D.FROM_DOC_DATE = FromDate
   D.TO_DOC_DATE = ToDate
   D.FROM_INV_DATE = FromInvDate
   D.TO_INV_DATE = ToInvDate
   D.OrderBy = 1
   Call D.QueryData(9, Rs, ItemCount)
   
   If Not (C Is Nothing) Then
      C.Clear
      i = 0
      C.AddItem ("")
   End If
   
   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   While Not Rs.EOF
      i = i + 1
      Set TempData = New CARTrn
      Call TempData.PopulateFromRS(9, Rs)
      
      If Not (C Is Nothing) Then
      End If
      
      If Not (Cl Is Nothing) Then
         Call Cl.Add(TempData, "1")
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   Set Rs = Nothing
   Set D = Nothing
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub
Public Sub LoadReceiveAmount(C As ComboBox, Optional Cl As Collection = Nothing, Optional FromDate As Date, Optional ToDate As Date, Optional FromInvDate As Date, Optional ToInvDate As Date, Optional RecTypeSet As String, Optional FromCustomerCode As String, Optional ToCustomerCode As String)
On Error GoTo ErrorHandler
Dim D As CARRcIt
Dim ItemCount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CARRcIt
Dim i As Long
Dim TempData1 As CARRcIt

   Set D = New CARRcIt
   Set Rs = New ADODB.Recordset
   
   D.FROM_PAY_DATE = FromDate
   D.TO_PAY_DATE = ToDate
   D.FROM_DOC_DATE = FromInvDate
   D.TO_DOC_DATE = ToInvDate
   D.FROM_CUSTOMER_CODE = FromCustomerCode
   D.TO_CUSTOMER_CODE = ToCustomerCode
   D.RecTypeSet = RecTypeSet
   D.OrderBy = 1
   Call D.QueryData(8, Rs, ItemCount)
   
   If Not (C Is Nothing) Then
      C.Clear
      i = 0
      C.AddItem ("")
   End If

   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   While Not Rs.EOF
      i = i + 1
      Set TempData = New CARRcIt
      Call TempData.PopulateFromRS(8, Rs)
   
      If Not (C Is Nothing) Then
      End If
      
      If Not (Cl Is Nothing) Then
         Call Cl.Add(TempData, "1")
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   Set Rs = Nothing
   Set D = Nothing
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub
Public Sub LoadARAmount(C As ComboBox, Optional Cl As Collection = Nothing, Optional FromDate As Date, Optional ToDate As Date, Optional FromCustomerCode As String, Optional ToCustomerCode As String, Optional RecTypeSet As String)
On Error GoTo ErrorHandler
Dim D As CARTrn
Dim ItemCount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CARTrn
Dim i As Long
Dim TempData1   As CARTrn
   
   Set D = New CARTrn
   Set Rs = New ADODB.Recordset

   D.FROM_DOC_DATE = FromDate
   D.TO_DOC_DATE = ToDate
   D.OrderBy = 1
   D.FROM_CUSTOMER_CODE = FromCustomerCode
   D.TO_CUSTOMER_CODE = ToCustomerCode
   D.RecTypeSet = RecTypeSet
   Call D.QueryData(6, Rs, ItemCount)
   
   If Not (C Is Nothing) Then
      C.Clear
      i = 0
      C.AddItem ("")
   End If
   
   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   While Not Rs.EOF
      i = i + 1
      Set TempData = New CARTrn
      Call TempData.PopulateFromRS(6, Rs)
      
      If Not (C Is Nothing) Then
      End If
      
      If Not (Cl Is Nothing) Then
         Call Cl.Add(TempData, "1")
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   Set Rs = Nothing
   Set D = Nothing
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub
Public Sub LoadDataType(C As ComboBox, Optional Cl As Collection = Nothing)
On Error GoTo ErrorHandler
Dim D As CDataType
Dim ItemCount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CDataType
Dim i As Long
   
   Set D = New CDataType
   Set Rs = New ADODB.Recordset
   
   Call D.QueryData(Rs, ItemCount)
   
   If Not (C Is Nothing) Then
      C.Clear
      i = 0
      C.AddItem ("")
   End If
   
   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   
   While Not Rs.EOF
      i = i + 1
      Set TempData = New CDataType
      Call TempData.PopulateFromRS(1, Rs)

      If Not (C Is Nothing) Then
         C.AddItem (TempData.DATA_TYPE_NAME)
         C.ItemData(i) = Val(TempData.DATA_TYPE_ID)
      End If

      If Not (Cl Is Nothing) Then
         Call Cl.Add(TempData, TempData.DATA_TYPE_NAME)
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   Set Rs = Nothing
   Set D = Nothing
   Exit Sub

ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub
Public Function GetSupplier(m_TempCol As Collection, TempKey As String) As CAPMas
On Error Resume Next
Dim Ei As CAPMas
Static TempEi As CAPMas

   Set Ei = m_TempCol(TempKey)
   If Ei Is Nothing Then
   Else
      Set GetSupplier = Ei
   End If
End Function
Public Function GetSupplierGroup(m_TempCol As Collection, TempKey As String) As CSupplierGroup
On Error Resume Next
Dim Ei As CSupplierGroup
Static TempEi As CSupplierGroup

   Set Ei = m_TempCol(TempKey)
   If Ei Is Nothing Then
   Else
      Set GetSupplierGroup = Ei
   End If
End Function
Public Sub LoadAPAmountByReceiveCheque(C As ComboBox, Optional Cl As Collection = Nothing, Optional FromDate As Date, Optional ToDate As Date, Optional RecTypeSet As String, Optional FromChequeDate As Date, Optional ToChequeDate As Date)
On Error GoTo ErrorHandler
Dim D As CApTrn
Dim ItemCount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CApTrn
Dim i As Long
Dim TempApTrn As CApTrn

   Set D = New CApTrn
   Set Rs = New ADODB.Recordset

   D.FROM_PAY_DATE = FromDate
   D.TO_PAY_DATE = ToDate
   D.OrderBy = 1
   D.RecTypeSet = RecTypeSet
   D.FROM_CHEQUE_DATE = FromChequeDate
   D.TO_CHEQUE_DATE = ToChequeDate
'                        D.DOCNUM = "RR50030046"
'                        D.RecTypeSet = "('3', '4', '5')"
'                        D.SUPCOD = "ค-106"
   Call D.QueryData(4, Rs, ItemCount)
   
   If Not (C Is Nothing) Then
      C.Clear
      i = 0
      C.AddItem ("")
   End If
   
   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   While Not Rs.EOF
      i = i + 1
      Set TempData = New CApTrn
      Call TempData.PopulateFromRS(4, Rs)
      
'      If TempData.SUPCOD = "ฟ -002" Then
'         ''debug.print
'      End If
      
      If Not (C Is Nothing) Then
      End If
      
'      If TempData.AMOUNT > 0 Then
'        If Not (Cl Is Nothing) Then
'          Set TempApTrn = GetAPTrn(Cl, TempData.DOCNUM, False)
'          If (TempApTrn Is Nothing) Then
'              TempData.CHQNUM = TempData.CHQNUM & "(" & TempData.CHQDAT & "," & TempData.AMOUNT & ")"
'              Call Cl.Add(TempData, TempData.DOCNUM)
'          Else
'              TempApTrn.CHQNUM = TempApTrn.CHQNUM & " " & TempData.CHQNUM & "(" & TempData.CHQDAT & "," & TempData.AMOUNT & ")"
'              TempApTrn.AMOUNT = TempApTrn.AMOUNT + TempData.AMOUNT
'            End If
'        End If
'      End If
'      If TempData.RCPNUM = "PS5505230" Then
'         ''debug.print
'      End If
      
      If TempData.AMOUNT > 0 Then
        If Not (Cl Is Nothing) Then
          Set TempApTrn = GetAPTrn(Cl, TempData.RCPNUM, False)
          If (TempApTrn Is Nothing) Then
              TempData.DOCNUM = TempData.DOCNUM & "(" & TempData.DOCDAT & ")"
              TempData.CHQNUM = TempData.CHQNUM & "(" & TempData.CHQDAT & "," & TempData.AMOUNT & ")"
              Call Cl.Add(TempData, TempData.RCPNUM)
          Else
              If InStr(1, TempApTrn.DOCNUM, TempData.DOCNUM) <= 0 Then 'ไม่เจอ ไม่ซ้ำเพิ่ม
                  TempApTrn.DOCNUM = TempApTrn.DOCNUM & "," & TempData.DOCNUM & "(" & TempData.DOCDAT & ")"
              End If
               If InStr(1, TempApTrn.CHQNUM, TempData.CHQNUM) <= 0 Then 'ไม่เจอ ไม่ซ้ำเพิ่ม
                  TempApTrn.CHQNUM = TempApTrn.CHQNUM & " " & TempData.CHQNUM & "(" & TempData.CHQDAT & "," & TempData.AMOUNT & ")"
                  TempApTrn.AMOUNT = TempApTrn.AMOUNT + TempData.AMOUNT
              End If
            End If
        End If
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   Set Rs = Nothing
   Set D = Nothing
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub


Public Sub LoadArea(C As ComboBox, Optional Cl As Collection = Nothing, Optional Area As String = "")
On Error GoTo ErrorHandler
Dim D As CIsTab
Dim ItemCount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CIsTab
Dim i As Long
   
   Set D = New CIsTab
   Set Rs = New ADODB.Recordset
   D.TABTYP = "40"
   Call D.QueryData(Rs, ItemCount, 2)
   
   
   If Not (C Is Nothing) Then
      C.Clear
      i = 0
      C.AddItem ("")
   End If
   
   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   
   While Not Rs.EOF
      i = i + 1
      Set TempData = New CIsTab
      Call TempData.PopulateFromRS(1, Rs)

      If Not (C Is Nothing) Then
         C.AddItem (TempData.TYPCOD & "-" & TempData.TYPDES)
         C.ItemData(i) = TempData.TYPCOD
      '   C.List(i) = TempData.TYPCOD
   
      End If

      If Not (Cl Is Nothing) Then
         Call Cl.Add(TempData, Trim(Str(TempData.TYPCOD)))
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   Set Rs = Nothing
   Set D = Nothing
   Exit Sub

ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End
End Sub

Public Sub InitYearComOrderBy(C As ComboBox)
   C.Clear
   
   C.AddItem ("")
   C.ItemData(0) = 0
   
   C.AddItem (MapText("ปี"))
   C.ItemData(1) = 1
   
   C.AddItem (MapText("วันที่เริ่มต้น"))
   C.ItemData(2) = 2
   
   C.AddItem (MapText("วันที่สิ้นสุด"))
   C.ItemData(3) = 3
End Sub

Public Sub InitGoodsMasterOrderBy(C As ComboBox)
   C.Clear
   
   C.AddItem ("")
   C.ItemData(0) = 0
   
   C.AddItem (MapText("รหัสกลุ่ม"))
   C.ItemData(1) = 1
   
   C.AddItem (MapText("ชื่อกลุ่ม"))
   C.ItemData(2) = 2

End Sub

Public Sub LoadStkcodLookup(C As ComboBox, Optional Cl As Collection = Nothing, Optional Cl2 As Collection = Nothing, Optional DOCNUM As String, Optional findtxt As Boolean = True)
On Error GoTo ErrorHandler
Dim D As CStcrd
Dim ItemCount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CStcrd
Dim i As Long
   
   Set D = New CStcrd
   Set Rs = New ADODB.Recordset
   
   D.DOCNUM = DOCNUM                                ' ตัดว่า IV นี้ มีสินค้าตัวใดบ้าง
   Call D.QueryData(8, Rs, ItemCount)
   
   If Not (C Is Nothing) And findtxt Then  '
      C.Clear
      i = 0
      C.AddItem ("")
   End If
   
   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   
   While Not Rs.EOF
      i = i + 1
      Set TempData = New CStcrd
      Call TempData.PopulateFromRS(8, Rs, i)
      If Not (C Is Nothing) And findtxt Then  '
         C.AddItem (TempData.STKDES)
         C.ItemData(i) = i
      End If
      
      If i <> 0 Then
            Set D = GetStkcodNoNew(Cl, Trim(Str(i)))
            If D Is Nothing Then
                Call Cl.Add(TempData, Trim(Str(i)))             ' ถ้าแก้ตรงนี้งานจะเข้าตอน lookup
'                  If Not (Cl2 Is Nothing) Then
'                        Call Cl2.Add(TempData, Trim(TempData.DOCNUM))               ' ใช้คิวรี่ 8 ไม่ได้เพราะ IV จะซ้ำ
'                   End If
            End If
      End If
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   Set Rs = Nothing
   Set D = Nothing
   Exit Sub

ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub


Public Sub LoadIVsaleLookup(C As ComboBox, Optional outputSLMCOD As String = "", Optional Cl2 As Collection = Nothing, Optional DOCNUM As String, Optional findtxt As Boolean = True)
On Error GoTo ErrorHandler
Dim D As CStcrd
Dim ItemCount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CStcrd
Dim i As Long
   
   Set D = New CStcrd
   Set Rs = New ADODB.Recordset
   
   D.DOCNUM = DOCNUM                                ' ตัดว่า IV นี้ มีสินค้าตัวใดบ้าง
   Call D.QueryData(7, Rs, ItemCount)
   
   If Not (C Is Nothing) And findtxt Then  '
      C.Clear
      i = 0
      C.AddItem ("")
   End If
   
'   If Not (Cl Is Nothing) Then
'      Set Cl = Nothing
'      Set Cl = New Collection
'   End If

   outputSLMCOD = ""
   
   While Not Rs.EOF
      i = i + 1
      Set TempData = New CStcrd
      Call TempData.PopulateFromRS(7, Rs, i)
      If Not (C Is Nothing) And findtxt Then  '
         C.AddItem (Trim(TempData.SLMNAM))
         C.ItemData(i) = TempData.SLMCOD
      End If
      
      outputSLMCOD = TempData.SLMCOD
'      If i <> 0 Then
'            Set D = GetStkcodNoNew(Cl, Trim(Str(i)))
'            If D Is Nothing Then
'                Call Cl.Add(TempData, Trim(Str(i)))             ' ถ้าแก้ตรงนี้งานจะเข้าตอน lookup
''                  If Not (Cl2 Is Nothing) Then
''                        Call Cl2.Add(TempData, Trim(TempData.DOCNUM))               ' ใช้คิวรี่ 8 ไม่ได้เพราะ IV จะซ้ำ
''                   End If
'            End If
'      End If
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   Set Rs = Nothing
   Set D = Nothing
   Exit Sub

ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub


Public Sub LoadIVcustomerLookup(Optional outputCuscod As String = "", Optional Cl2 As Collection = Nothing, Optional DOCNUM As String, Optional findtxt As Boolean = True)
On Error GoTo ErrorHandler
Dim D As CStcrd
Dim ItemCount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CStcrd
Dim i As Long
   
   Set D = New CStcrd
   Set Rs = New ADODB.Recordset
   
   D.DOCNUM = DOCNUM                                ' ตัดว่า IV นี้ มีสินค้าตัวใดบ้าง
   Call D.QueryData(7, Rs, ItemCount)
   
'   If Not (C Is Nothing) And findtxt Then  '
'      C.Clear
'      i = 0
'      C.AddItem ("")
'   End If
   outputCuscod = ""
         
   While Not Rs.EOF
      i = i + 1
      Set TempData = New CStcrd
      Call TempData.PopulateFromRS(7, Rs, i)
'      If Not (C Is Nothing) And findtxt Then  '
'         C.AddItem (Trim(TempData.CUSNAM))
'         C.ItemData(i) = i
'      End If
      
      outputCuscod = TempData.CUSCOD
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   Set Rs = Nothing
   Set D = Nothing
   Exit Sub

ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub


Public Sub LoadIVfromCStcrd(C As ComboBox, Optional Cl As Collection = Nothing, Optional DOCNUM As String)
On Error GoTo ErrorHandler
Dim D As CStcrd
Dim ItemCount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CStcrd
Dim i As Long
   
   Set D = New CStcrd
   Set Rs = New ADODB.Recordset
   
   D.DOCNUM = DOCNUM
   Call D.QueryData(10, Rs, ItemCount)
   
   If Not (C Is Nothing) Then
      C.Clear
      i = 0
      C.AddItem ("")
   End If
   
   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   
   While Not Rs.EOF
      i = i + 1
      Set TempData = New CStcrd
      Call TempData.PopulateFromRS(10, Rs, i)
      If Not (C Is Nothing) Then
         C.AddItem (TempData.STKDES)
         C.ItemData(i) = i
      End If
      
      If i <> 0 Then
            Set D = GetStkcodNoNew(Cl, Trim(TempData.DOCNUM))
            If D Is Nothing Then
                Call Cl.Add(TempData, Trim(TempData.DOCNUM))             ' ถ้าแก้ตรงนี้งานจะเข้าตอน lookup
'                  If Not (Cl2 Is Nothing) Then
'                        Call Cl2.Add(TempData, Trim(TempData.DOCNUM))               ' ใช้คิวรี่ 8 ไม่ได้เพราะ IV จะซ้ำ
'                   End If
            End If
      End If
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   Set Rs = Nothing
   Set D = Nothing
   Exit Sub

ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub

Public Function GetStmasNoNew(m_TempCol As Collection, TempKey As String) As CStmas
On Error Resume Next
Dim Ei As CStmas
Static TempEi As CStmas

   Set Ei = m_TempCol(TempKey)
   If Ei Is Nothing Then
'      If TempEi Is Nothing Then
'         Set TempEi = New Cstcrd
'      End If
      Set GetStmasNoNew = TempEi
   Else
      Set GetStmasNoNew = Ei
   End If
End Function

Public Function GetStkcodNoNew(m_TempCol As Collection, TempKey As String) As CStcrd
On Error Resume Next
Dim Ei As CStcrd
Static TempEi As CStcrd

   Set Ei = m_TempCol(TempKey)
   If Ei Is Nothing Then
'      If TempEi Is Nothing Then
'         Set TempEi = New Cstcrd
'      End If
      Set GetStkcodNoNew = TempEi
   Else
      Set GetStkcodNoNew = Ei
   End If
End Function

Public Sub LoadStcrdBySale(C As ComboBox, Optional Cl As Collection = Nothing, Optional FromCMPLDate As Date, Optional toCMPLdate As Date, Optional AREACOD As String)
On Error GoTo ErrorHandler
Dim D As CStcrd
Dim ItemCount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CStcrd
Dim i As Long

   Set D = New CStcrd
   Set Rs = New ADODB.Recordset

   D.FROM_CMPL_DATE = FromCMPLDate
   D.TO_CMPL_DATE = toCMPLdate
   D.AREACOD = AREACOD
   Call D.QueryData(9, Rs, ItemCount)
   
   If Not (C Is Nothing) Then
      C.Clear
      i = 0
      C.AddItem ("")
   End If
   
   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   
   While Not Rs.EOF
      i = i + 1
      Set TempData = New CStcrd
      Call TempData.PopulateFromRS(9, Rs)
      
      If Not (C Is Nothing) Then
      End If
      
      If Not (Cl Is Nothing) Then
         Call Cl.Add(TempData, TempData.DOCNUM & "-" & TempData.SEQNUM)
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   Set Rs = Nothing
   Set D = Nothing
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub

Public Sub LoadSumStcrdMonth(C As ComboBox, Optional Cl As Collection = Nothing, Optional FromDocDate As Date, Optional ToDocDate As Date, Optional FromCMPLDate As Date, Optional toCMPLdate As Date)
On Error GoTo ErrorHandler
Dim D As CStcrd
Dim ItemCount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CStcrd
Dim stcrd_temp As CStcrd
Dim i As Long

   Set D = New CStcrd
   Set Rs = New ADODB.Recordset

   D.FROM_DOC_DATE = FromDocDate
   D.TO_DOC_DATE = ToDocDate
   D.FROM_SLM_DATE = FromCMPLDate
   D.TO_SLM_DATE = toCMPLdate
   Call D.QueryData(11, Rs, ItemCount)
   
   If ItemCount <= 0 Then            ' ติดต่อกับ db ปกติแล้วเป็น 0 ให้ติดต่อ db ที่เป็นรอง ดู
      D.FROM_DOC_DATE = FromDocDate
      D.TO_DOC_DATE = ToDocDate
      D.FROM_SLM_DATE = FromCMPLDate
      D.TO_SLM_DATE = toCMPLdate
      Call D.QueryData(11, Rs, ItemCount, 2)
   End If
   
   If Not (C Is Nothing) Then
      C.Clear
      i = 0
      C.AddItem ("")
   End If
   
   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   
   While Not Rs.EOF
      i = i + 1
      Set TempData = New CStcrd
      Call TempData.PopulateFromRS(11, Rs)
      
      If Not (C Is Nothing) Then
      End If
      
      If Not (Cl Is Nothing) Then
       '  If TempData.STKCOD = "V100-1" Then
       '     ''debug.print
       '  End If
         Set stcrd_temp = GetObject("CStcrd", Cl, TempData.STKCOD, False)
          If stcrd_temp Is Nothing Then  ' ถ้าไม่มีในคอเล็กชั่น คือ เพิ่มได้
              Call Cl.Add(TempData, TempData.STKCOD)
          Else
    
               stcrd_temp.TRNQTY = TempData.TRNQTY + stcrd_temp.TRNQTY
          End If
 
       '  ''debug.print TempData.STKCOD & " == " & TempData.TRNQTY
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   Set Rs = Nothing
   Set D = Nothing
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub

Public Sub LoadGoodsMaster(C As ComboBox, Optional Cl As Collection = Nothing, Optional GOODS_MASTER_ID As Long = 0)
On Error GoTo ErrorHandler
Dim D As CGoodsMaster
Dim ItemCount As Long
Dim iCount As Long
Dim Rs As ADODB.Recordset
Dim m_Rs1 As ADODB.Recordset
Dim TempData As CGoodsMaster
Dim i As Long
   
   Set D = New CGoodsMaster
   Set Rs = New ADODB.Recordset
   Set m_Rs1 = New ADODB.Recordset
   
   D.GOODS_MASTER_ID = GOODS_MASTER_ID
   Call D.QueryData(1, Rs, ItemCount)
   
   If Not (C Is Nothing) Then
      C.Clear
      i = 0
      C.AddItem ("")
   End If
   
   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   
   While Not Rs.EOF              ' แต่ของหยงไม่มีวนของแม่
      i = i + 1
      Set TempData = New CGoodsMaster
      Call TempData.PopulateFromRS(1, Rs)

      If Not (C Is Nothing) Then
         C.AddItem (TempData.GOODS_MASTER_CODE & " - " & TempData.GOODS_MASTER_NAME)
         C.ItemData(i) = TempData.GOODS_MASTER_ID                             ' param1
      End If

'      If Not (Cl Is Nothing) Then
'
'            ' If Ua.QueryFlag = 1 Then         ' เป็นตัวเช็คว่า ต้องการให้แอดตัวลูกในแม่ไหม ?   ถ้าไม่มีเนี่ย ตอนโหลดฟอร์มมันก็จะแอดลูกด้วย ซึ่งไม่จำเป็น
'                Dim Gse As CCommissionCustomerArea
'                Set Gse = New CCommissionCustomerArea
'                Gse.MASTER_AREA_ID = TempData.MASTER_AREA_ID
'                Call Gse.QueryData(1, m_Rs1, iCount)
'                Set Gse = Nothing
'
'                Set TempData.ImportExportItems = Nothing
'                Set TempData.ImportExportItems = New Collection
'
'                            While Not m_Rs1.EOF
'                               Set Gse = New CCommissionCustomerArea
'                               Call Gse.PopulateFromRS(1, m_Rs1)                ' ป๊อปค่าจากลูก แล้วแอดเข้าตัวแม่
'                               Call TempData.ImportExportItems.Add(Gse, Gse.COMMISSION_CUS_ID)
'                               Set Gse = Nothing
'
'                               m_Rs1.MoveNext
'                            Wend
'                 Call Cl.Add(TempData, Trim(Str(TempData.MASTER_AREA_ID)))             ' เก็บค่าตัวแม่ แล้วน่าจะแอด collecion ลูกในตัวแม่นั้นเลย
'            ' End If
'      End If
'
      Set TempData = Nothing
      Rs.MoveNext
   Wend
'
'   If Not (Cl Is Nothing) Then
'        Set TempData = New CCommissMasterArea
'        TempData.MASTER_AREA_ID = 999999
'        TempData.MASTER_AREA_NAME = "ยังไม่ระบุ"
'        Call Cl.Add(TempData, Trim(Str(TempData.MASTER_AREA_ID)))
'        Set TempData = Nothing
'   End If
   
   Set Rs = Nothing
   Set D = Nothing
    Exit Sub

ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub

Public Sub LoadColGoodsGroup(Optional Cl As Collection = Nothing)
On Error GoTo ErrorHandler
Dim D As CGoodsGroup
Dim ItemCount As Long
Dim iCount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CGoodsGroup
Dim i As Long
   
   Set D = New CGoodsGroup
   Set Rs = New ADODB.Recordset

   Call D.QueryData(1, Rs, ItemCount)

   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   
   While Not Rs.EOF              ' แต่ของหยงไม่มีวนของแม่
      i = i + 1
      Set TempData = New CGoodsGroup
      Call TempData.PopulateFromRS(1, Rs)

      If Not (Cl Is Nothing) Then
           Call Cl.Add(TempData, Trim(Str(TempData.GOODS_GROUP_ID)))             ' เก็บค่าตัวแม่ แล้วน่าจะแอด collecion ลูกในตัวแม่นั้นเลย
      End If

      Set TempData = Nothing
      Rs.MoveNext
   Wend

'   If Not (Cl Is Nothing) Then
'        Set TempData = New CCommissMasterArea
'        TempData.MASTER_AREA_ID = 999999
'        TempData.MASTER_AREA_NAME = "ยังไม่ระบุ"
'        Call Cl.Add(TempData, Trim(Str(TempData.MASTER_AREA_ID)))
'        Set TempData = Nothing
'   End If
   
   Set Rs = Nothing
   Set D = Nothing
    Exit Sub

ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub

Public Sub LoadGoodsGroup(C As ComboBox, Optional Cl As Collection = Nothing, Optional GOODS_GROUP_ID As Long = 0)
On Error GoTo ErrorHandler
Dim D As CGoodsGroup
Dim ItemCount As Long
Dim iCount As Long
Dim Rs As ADODB.Recordset
Dim m_Rs1 As ADODB.Recordset
Dim TempData As CGoodsGroup
Dim i As Long
   
   Set D = New CGoodsGroup
   Set Rs = New ADODB.Recordset
   Set m_Rs1 = New ADODB.Recordset
   
   D.GOODS_GROUP_ID = GOODS_GROUP_ID
   Call D.QueryData(1, Rs, ItemCount)
   
   If Not (C Is Nothing) Then
      C.Clear
      i = 0
      C.AddItem ("")
   End If
   
   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   
   While Not Rs.EOF              ' แต่ของหยงไม่มีวนของแม่
      i = i + 1
      Set TempData = New CGoodsGroup
      Call TempData.PopulateFromRS(1, Rs)

      If Not (C Is Nothing) Then
         C.AddItem (TempData.GOODS_GROUP_CODE & " - " & TempData.GOODS_GROUP_NAME)
         C.ItemData(i) = TempData.GOODS_GROUP_ID                             ' param1
      End If

'      If Not (Cl Is Nothing) Then
'
'            ' If Ua.QueryFlag = 1 Then         ' เป็นตัวเช็คว่า ต้องการให้แอดตัวลูกในแม่ไหม ?   ถ้าไม่มีเนี่ย ตอนโหลดฟอร์มมันก็จะแอดลูกด้วย ซึ่งไม่จำเป็น
'                Dimet Gse = New CCommissionCustomerArea
'                Gse As CCommissionCustomerArea
'                S Gse.MASTER_AREA_ID = TempData.MASTER_AREA_ID
'                Call Gse.QueryData(1, m_Rs1, iCount)
'                Set Gse = Nothing
'
'                Set TempData.ImportExportItems = Nothing
'                Set TempData.ImportExportItems = New Collection
'
'                            While Not m_Rs1.EOF
'                               Set Gse = New CCommissionCustomerArea
'                               Call Gse.PopulateFromRS(1, m_Rs1)                ' ป๊อปค่าจากลูก แล้วแอดเข้าตัวแม่
'                               Call TempData.ImportExportItems.Add(Gse, Gse.COMMISSION_CUS_ID)
'                               Set Gse = Nothing
'
'                               m_Rs1.MoveNext
'                            Wend
'                 Call Cl.Add(TempData, Trim(Str(TempData.MASTER_AREA_ID)))             ' เก็บค่าตัวแม่ แล้วน่าจะแอด collecion ลูกในตัวแม่นั้นเลย
'            ' End If
'      End If
'
      Set TempData = Nothing
      Rs.MoveNext
   Wend
'
'   If Not (Cl Is Nothing) Then
'        Set TempData = New CCommissMasterArea
'        TempData.MASTER_AREA_ID = 999999
'        TempData.MASTER_AREA_NAME = "ยังไม่ระบุ"
'        Call Cl.Add(TempData, Trim(Str(TempData.MASTER_AREA_ID)))
'        Set TempData = Nothing
'   End If
   
   Set Rs = Nothing
   Set D = Nothing
    Exit Sub

ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub

Public Sub LoadAreaCom(C As ComboBox, Optional Cl As Collection = Nothing, Optional MASTER_AREA_ID As Long = 0)
On Error GoTo ErrorHandler
Dim D As CCommissMasterArea
Dim ItemCount As Long
Dim iCount As Long
Dim Rs As ADODB.Recordset
Dim m_Rs1 As ADODB.Recordset
Dim TempData As CCommissMasterArea
Dim i As Long
   
   Set D = New CCommissMasterArea
   Set Rs = New ADODB.Recordset
   Set m_Rs1 = New ADODB.Recordset
   
   D.MASTER_AREA_ID = MASTER_AREA_ID
   Call D.QueryData(1, Rs, ItemCount)
   
   If Not (C Is Nothing) Then
      C.Clear
      i = 0
      C.AddItem ("")
   End If
   
   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   
   While Not Rs.EOF              ' แต่ของหยงไม่มีวนของแม่
      i = i + 1
      Set TempData = New CCommissMasterArea
      Call TempData.PopulateFromRS(1, Rs)

      If Not (C Is Nothing) Then
         C.AddItem (TempData.MASTER_AREA_ID & " - " & TempData.MASTER_AREA_NAME)
         C.ItemData(i) = TempData.MASTER_AREA_ID                             ' param1
      End If

      If Not (Cl Is Nothing) Then
       
            ' If Ua.QueryFlag = 1 Then         ' เป็นตัวเช็คว่า ต้องการให้แอดตัวลูกในแม่ไหม ?   ถ้าไม่มีเนี่ย ตอนโหลดฟอร์มมันก็จะแอดลูกด้วย ซึ่งไม่จำเป็น
                Dim Gse As CCommissionCustomerArea
                Set Gse = New CCommissionCustomerArea
                Gse.MASTER_AREA_ID = TempData.MASTER_AREA_ID
                Call Gse.QueryData(1, m_Rs1, iCount)
                Set Gse = Nothing
                
                Set TempData.ImportExportItems = Nothing
                Set TempData.ImportExportItems = New Collection
                
                            While Not m_Rs1.EOF
                               Set Gse = New CCommissionCustomerArea
                               Call Gse.PopulateFromRS(1, m_Rs1)                ' ป๊อปค่าจากลูก แล้วแอดเข้าตัวแม่
                               Call TempData.ImportExportItems.Add(Gse, Gse.COMMISSION_CUS_ID)
                               Set Gse = Nothing
                      
                               m_Rs1.MoveNext
                            Wend
                 Call Cl.Add(TempData, Trim(Str(TempData.MASTER_AREA_ID)))             ' เก็บค่าตัวแม่ แล้วน่าจะแอด collecion ลูกในตัวแม่นั้นเลย
            ' End If
      End If

      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   If Not (Cl Is Nothing) Then
        Set TempData = New CCommissMasterArea
        TempData.MASTER_AREA_ID = 999999
        TempData.MASTER_AREA_NAME = "ยังไม่ระบุ"
        Call Cl.Add(TempData, Trim(Str(TempData.MASTER_AREA_ID)))
        Set TempData = Nothing
   End If
   
   Set Rs = Nothing
   Set D = Nothing
    Exit Sub

ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub


Public Sub LoadAreaComReport(C As ComboBox, Optional Cl As Collection = Nothing, Optional YEAR_ID As Long = 0, Optional MASTER_AREA_ID As Long = -1)
On Error GoTo ErrorHandler
Dim D As CCommissionCustomerArea
Dim ItemCount As Long
Dim iCount As Long
Dim Rs As ADODB.Recordset
Dim m_Rs1 As ADODB.Recordset
Dim TempData As CCommissionCustomerArea
Dim i As Long
   
   
   Set D = New CCommissionCustomerArea
   Set Rs = New ADODB.Recordset
   Set m_Rs1 = New ADODB.Recordset
   
   D.YEAR_ID = YEAR_ID
   D.MASTER_AREA_ID = MASTER_AREA_ID
   Call D.QueryData(1, Rs, ItemCount)
   
   If Not (C Is Nothing) Then
      C.Clear
      i = 0
      C.AddItem ("")
   End If
   
   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   
   While Not Rs.EOF              ' แต่ของหยงไม่มีวนของแม่
      i = i + 1
      Set TempData = New CCommissionCustomerArea
      Call TempData.PopulateFromRS(1, Rs)

      If Not (C Is Nothing) Then
         C.AddItem (TempData.MASTER_AREA_ID & " - " & TempData.MASTER_AREA_NAME)
         C.ItemData(i) = TempData.MASTER_AREA_ID                             ' param1
      End If

      If Not (Cl Is Nothing) Then
       
            ' If Ua.QueryFlag = 1 Then         ' เป็นตัวเช็คว่า ต้องการให้แอดตัวลูกในแม่ไหม ?   ถ้าไม่มีเนี่ย ตอนโหลดฟอร์มมันก็จะแอดลูกด้วย ซึ่งไม่จำเป็น
                Dim Gse As CCommissionCustomerArea
                Set Gse = New CCommissionCustomerArea
                Gse.MASTER_AREA_ID = TempData.MASTER_AREA_ID
                Gse.YEAR_ID = YEAR_ID
                Call Gse.QueryData(3, m_Rs1, iCount)
                Set Gse = Nothing
                
                Set TempData.ImportExportItems = Nothing
                Set TempData.ImportExportItems = New Collection
                
                            While Not m_Rs1.EOF
                               Set Gse = New CCommissionCustomerArea
                               Call Gse.PopulateFromRS(3, m_Rs1)                ' ป๊อปค่าจากลูก แล้วแอดเข้าตัวแม่
                               Call TempData.ImportExportItems.Add(Gse, Gse.COMMISSION_CUS_ID)
                               Set Gse = Nothing
                      
                               m_Rs1.MoveNext
                            Wend
                 Call Cl.Add(TempData, Trim(Str(TempData.MASTER_AREA_ID)))             ' เก็บค่าตัวแม่ แล้วน่าจะแอด collecion ลูกในตัวแม่นั้นเลย
            ' End If
      End If

      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   If Not (Cl Is Nothing) Then
        Set TempData = New CCommissionCustomerArea
        TempData.MASTER_AREA_ID = 999999
        TempData.MASTER_AREA_NAME = "อื่นๆ"
        Call Cl.Add(TempData, Trim(Str(TempData.MASTER_AREA_ID)))
        Set TempData = Nothing
   End If
   
   Set Rs = Nothing
   Set D = Nothing
    Exit Sub

ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub

Public Function GetCheckCommiss(m_TempCol As Collection, TempKey As String, Optional HaveNew As Boolean = True) As CConditionCommission
On Error Resume Next
Dim Ei As CConditionCommission
Static TempEi As CConditionCommission

   Set Ei = m_TempCol(TempKey)
   ''debug.print TempKey
    If Ei Is Nothing And HaveNew Then
                If TempEi Is Nothing Then
                   Set TempEi = New CConditionCommission
                End If
      Set GetCheckCommiss = TempEi
   Else
      Set GetCheckCommiss = Ei
   End If
End Function
Public Function GetComMasSubPro(m_TempCol As Collection, TempKey As String, Optional HaveNew As Boolean = True) As CComMasSubPromote
On Error Resume Next
Dim Ei As CComMasSubPromote
Static TempEi As CComMasSubPromote

   Set Ei = m_TempCol(TempKey)
    If Ei Is Nothing And HaveNew Then
                If TempEi Is Nothing Then
                   Set TempEi = New CComMasSubPromote
                End If
      Set GetComMasSubPro = TempEi
   Else
      Set GetComMasSubPro = Ei
   End If
End Function

Public Function GetComMasPro(m_TempCol As Collection, TempKey As String, Optional HaveNew As Boolean = True) As CCommissMasterPromote
On Error Resume Next
Dim Ei As CCommissMasterPromote
Static TempEi As CCommissMasterPromote

   Set Ei = m_TempCol(TempKey)
    If Ei Is Nothing And HaveNew Then
                If TempEi Is Nothing Then
                   Set TempEi = New CCommissMasterPromote
                End If
      Set GetComMasPro = TempEi
   Else
      Set GetComMasPro = Ei
   End If
End Function

Public Function GetCheckIncenPro(m_TempCol As Collection, TempKey As String, Optional HaveNew As Boolean = True) As CIncentivePromote
On Error Resume Next
Dim Ei As CIncentivePromote
Static TempEi As CIncentivePromote

   Set Ei = m_TempCol(TempKey)
  ' ''debug.print TempKey
    If Ei Is Nothing And HaveNew Then
                If TempEi Is Nothing Then
                   Set TempEi = New CIncentivePromote
                End If
      Set GetCheckIncenPro = TempEi
   Else
      Set GetCheckIncenPro = Ei
   End If
End Function

Public Sub LoadCommission(ComType As String, Optional Cl As Collection = Nothing, Optional FromCMPLDat As Date = -1, Optional ToCMPLDat As Date = -1, Optional GOODS_GROUP_ID As Long = -1)
On Error GoTo ErrorHandler
Dim D As CConditionCommission
Dim ItemCount As Long
Dim TempData As CConditionCommission
Dim Rs As ADODB.Recordset
Dim i As Long

   Set D = New CConditionCommission
   Set Rs = New ADODB.Recordset
   
 '  D.YEARNUM = Year(FromCMPLDat)
   D.FROM_CMPL_DATE = FromCMPLDat
   D.TO_CMPL_DATE = ToCMPLDat
   D.COMTYP = ComType
   D.GOODS_GROUP_ID = GOODS_GROUP_ID
   Call D.QueryData(3, Rs, ItemCount)

   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   
   While Not Rs.EOF
      i = i + 1
      Set TempData = New CConditionCommission
      Call TempData.PopulateFromRS(1, Rs)

      If Not (Cl Is Nothing) Then
         Call Cl.Add(TempData, Trim(TempData.NUM_ONE))   ' KEy
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   Set Rs = Nothing
   Set D = Nothing
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub

Public Sub LoadComPro(ComType As String, Optional Cl As Collection = Nothing, Optional FromCMPLDat As Date = -1, Optional ToCMPLDat As Date = -1)
On Error GoTo ErrorHandler
Dim D As CComMasSubPromote
Dim ItemCount As Long
Dim TempData As CComMasSubPromote
Dim Rs As ADODB.Recordset
Dim m_Rs1 As ADODB.Recordset
Dim Gr As CCommissPromote
Dim iCount As Long
Dim i As Long

   Set D = New CComMasSubPromote
   Set Rs = New ADODB.Recordset
      Set m_Rs1 = New ADODB.Recordset
   
 '  D.YEARNUM = Year(FromCMPLDat)
D.FROM_CMPL_DATE = FromCMPLDat
D.TO_CMPL_DATE = ToCMPLDat
   D.CREDIT_TYP = ComType
   Call D.QueryData(3, Rs, ItemCount)

   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   
   While Not Rs.EOF
      i = i + 1
      Set TempData = New CComMasSubPromote
      Call TempData.PopulateFromRS(3, Rs)

      Set Gr = New CCommissPromote
      Gr.MASTER_COMMISS_SUB_PROMOTE_ID = TempData.MASTER_COMMISS_SUB_PROMOTE_ID
      Gr.FROM_CMPL_DATE = -1
      Gr.TO_CMPL_DATE = -1
      Gr.Commiss_TYP = "01"
      Call Gr.QueryData(1, m_Rs1, iCount)
      Set Gr = Nothing

      Set TempData.DetailsCom1 = Nothing
      Set TempData.DetailsCom1 = New Collection

      While Not m_Rs1.EOF
         Set Gr = New CCommissPromote
         Call Gr.PopulateFromRS(1, m_Rs1)                        ' นี่คือตัวเนื้อใน ต้อง Commission 4 เงื่อนไข

         Gr.Flag = "I"
         Call TempData.DetailsCom1.Add(Gr)
         Set Gr = Nothing
      m_Rs1.MoveNext
  Wend

      Set Gr = New CCommissPromote
      Gr.MASTER_COMMISS_SUB_PROMOTE_ID = TempData.MASTER_COMMISS_SUB_PROMOTE_ID
      Gr.FROM_CMPL_DATE = -1
      Gr.TO_CMPL_DATE = -1
      Gr.Commiss_TYP = "02"
      Call Gr.QueryData(1, m_Rs1, iCount)
      Set Gr = Nothing

      Set TempData.DetailsCom2 = Nothing
      Set TempData.DetailsCom2 = New Collection

      While Not m_Rs1.EOF
         Set Gr = New CCommissPromote
         Call Gr.PopulateFromRS(1, m_Rs1)                        ' นี่คือตัวเนื้อใน ต้อง Commission 4 เงื่อนไข

         Gr.Flag = "I"
         Call TempData.DetailsCom2.Add(Gr)
         Set Gr = Nothing
      m_Rs1.MoveNext
  Wend

      Set Gr = New CCommissPromote
      Gr.MASTER_COMMISS_SUB_PROMOTE_ID = TempData.MASTER_COMMISS_SUB_PROMOTE_ID
      Gr.FROM_CMPL_DATE = -1
      Gr.TO_CMPL_DATE = -1
      Gr.Commiss_TYP = "03"
      Call Gr.QueryData(1, m_Rs1, iCount)
      Set Gr = Nothing

      Set TempData.DetailsCom3 = Nothing
      Set TempData.DetailsCom3 = New Collection

      While Not m_Rs1.EOF
         Set Gr = New CCommissPromote
         Call Gr.PopulateFromRS(1, m_Rs1)                        ' นี่คือตัวเนื้อใน ต้อง Commission 4 เงื่อนไข

         Gr.Flag = "I"
         Call TempData.DetailsCom3.Add(Gr)
         Set Gr = Nothing
      m_Rs1.MoveNext
  Wend
      
      
      If Not (Cl Is Nothing) Then
         Call Cl.Add(TempData, Trim(TempData.SLM_ID & "-" & TempData.CUS_ID & "-" & TempData.CREDIT_NAME))   ' KEy
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   Set Rs = Nothing
   Set D = Nothing
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub

Public Sub LoadComNamePro(Optional Cl As Collection = Nothing, Optional FromCMPLDat As Date = -1, Optional ToCMPLDat As Date = -1)
On Error GoTo ErrorHandler
Dim D As CCommissMasterPromote
Dim ItemCount As Long
Dim TempData As CCommissMasterPromote
Dim Rs As ADODB.Recordset
Dim i As Long

   Set D = New CCommissMasterPromote
   Set Rs = New ADODB.Recordset
   
 '  D.YEARNUM = Year(FromCMPLDat)
D.FromCMPLDat = FromCMPLDat
D.ToCMPLDat = ToCMPLDat
   Call D.QueryData(1, Rs, ItemCount)

   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   
   While Not Rs.EOF
      i = i + 1
      Set TempData = New CCommissMasterPromote
      Call TempData.PopulateFromRS(1, Rs)
      
      If Not (Cl Is Nothing) Then
         Call Cl.Add(TempData, Trim(TempData.SLM_ID & "-" & TempData.CUS_ID))   ' KEy
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend

   Set Rs = Nothing
   Set D = Nothing
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub


Public Sub LoadCommission04(Optional Cl As Collection = Nothing, Optional FromCMPLDat As Date = -1, Optional ToCMPLDat As Date = -1, Optional GOODS_GROUP_ID As Long = 1)
On Error GoTo ErrorHandler
Dim D As CConditionCommission
Dim ItemCount As Long
Dim TempData As CConditionCommission
Dim Rs As ADODB.Recordset
Dim i As Long

   Set D = New CConditionCommission
   Set Rs = New ADODB.Recordset
   
 '  D.YEARNUM = Year(FromCMPLDat)
   D.FROM_CMPL_DATE = FromCMPLDat
   D.TO_CMPL_DATE = ToCMPLDat
   D.COMTYP = "04"
   D.GOODS_GROUP_ID = GOODS_GROUP_ID
   Call D.QueryData(3, Rs, ItemCount)

   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   
   While Not Rs.EOF
      i = i + 1
      Set TempData = New CConditionCommission
      Call TempData.PopulateFromRS(1, Rs)

      If Not (Cl Is Nothing) Then
         Call Cl.Add(TempData, Trim(TempData.STKCOD))   ' KEy
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   Set Rs = Nothing
   Set D = Nothing
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub

Public Sub LoadCommission05(Optional Cl1 As Collection = Nothing, Optional FromCMPLDat As Date = -1, Optional ToCMPLDat As Date = -1, Optional GOODS_GROUP_ID As Long = 1)
On Error GoTo ErrorHandler
Dim D As CConditionCommission
Dim ItemCount As Long
Dim TempData As CConditionCommission
Dim Rs As ADODB.Recordset
Dim i As Long

   Set D = New CConditionCommission
   Set Rs = New ADODB.Recordset
   
   D.FROM_CMPL_DATE = FromCMPLDat
   D.TO_CMPL_DATE = ToCMPLDat
   D.COMTYP = "05"
   D.GOODS_GROUP_ID = GOODS_GROUP_ID
   Call D.QueryData(3, Rs, ItemCount)

   If Not (Cl1 Is Nothing) Then
      Set Cl1 = Nothing
      Set Cl1 = New Collection
   End If
   
   While Not Rs.EOF
      i = i + 1
      Set TempData = New CConditionCommission
      Call TempData.PopulateFromRS(1, Rs)

      If Not (Cl1 Is Nothing) Then
         Call Cl1.Add(TempData, Trim(TempData.STKCOD & "-" & TempData.NUM_ONE & "-" & TempData.NUM_TWO))    ' KEy
         
      End If

      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   Set Rs = Nothing
   Set D = Nothing
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub

Public Sub LoadCommission05_1(Optional Cl1 As Collection = Nothing, Optional FromCMPLDat As Date = -1, Optional ToCMPLDat As Date = -1, Optional STKCOD As String, Optional GOODS_GROUP_ID As Long = 1)
On Error GoTo ErrorHandler
Dim D As CConditionCommission
Dim ItemCount As Long
Dim TempData As CConditionCommission
Dim Rs As ADODB.Recordset
Dim i As Long

   Set D = New CConditionCommission
   Set Rs = New ADODB.Recordset
   
   D.FROM_CMPL_DATE = FromCMPLDat
   D.TO_CMPL_DATE = ToCMPLDat
   D.COMTYP = "05"
   D.STKCOD = STKCOD
   D.GOODS_GROUP_ID = GOODS_GROUP_ID
   Call D.QueryData(3, Rs, ItemCount)

   If Not (Cl1 Is Nothing) Then
      Set Cl1 = Nothing
      Set Cl1 = New Collection
   End If
   
'   If Not (Cl2 Is Nothing) Then
'      Set Cl2 = Nothing
'      Set Cl2 = New Collection
'   End If
   
   While Not Rs.EOF
      i = i + 1
      Set TempData = New CConditionCommission
      Call TempData.PopulateFromRS(1, Rs)

      If Not (Cl1 Is Nothing) Then
          Call Cl1.Add(TempData)    ' ไม่ต้องมี key เพราะเอาไปวนหาอย่างเดียว
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   Set Rs = Nothing
   Set D = Nothing
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub


Public Sub LoadCommission05_2(Optional Cl1 As Collection = Nothing, Optional FromCMPLDat As Date = -1, Optional ToCMPLDat As Date = -1, Optional STKCOD As String, Optional GOODS_GROUP_ID As Long = 1)
On Error GoTo ErrorHandler
Dim D As CConditionCommission
Dim ItemCount As Long
Dim TempData As CConditionCommission
Dim Rs As ADODB.Recordset
Dim i As Long

   Set D = New CConditionCommission
   Set Rs = New ADODB.Recordset
   
   D.FROM_CMPL_DATE = FromCMPLDat
   D.TO_CMPL_DATE = ToCMPLDat
   D.COMTYP = "05"
   D.GOODS_GROUP_ID = GOODS_GROUP_ID
   Call D.QueryData(5, Rs, ItemCount)

   If Not (Cl1 Is Nothing) Then
      Set Cl1 = Nothing
      Set Cl1 = New Collection
   End If
   
'   If Not (Cl2 Is Nothing) Then
'      Set Cl2 = Nothing
'      Set Cl2 = New Collection
'   End If
   
   While Not Rs.EOF
      i = i + 1
      Set TempData = New CConditionCommission
      Call TempData.PopulateFromRS(5, Rs)

      If Not (Cl1 Is Nothing) Then
          Call Cl1.Add(TempData, TempData.STKCOD)   ' ไม่ต้องมี key เพราะเอาไปวนหาอย่างเดียว
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   Set Rs = Nothing
   Set D = Nothing
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub


Public Sub LoadIncenPro05(Optional Cl1 As Collection = Nothing, Optional FromCMPLDat As Date = -1, Optional ToCMPLDat As Date = -1)
On Error GoTo ErrorHandler
Dim D As CIncentivePromote
Dim ItemCount As Long
Dim TempData As CIncentivePromote
Dim Rs As ADODB.Recordset
Dim i As Long

   Set D = New CIncentivePromote
   Set Rs = New ADODB.Recordset
   
   D.FROM_CMPL_DATE = FromCMPLDat
   D.TO_CMPL_DATE = ToCMPLDat
   Call D.QueryData(3, Rs, ItemCount)

   If Not (Cl1 Is Nothing) Then
      Set Cl1 = Nothing
      Set Cl1 = New Collection
   End If
   
   While Not Rs.EOF
      i = i + 1
      Set TempData = New CIncentivePromote
      Call TempData.PopulateFromRS(1, Rs)

      If Not (Cl1 Is Nothing) Then
         Call Cl1.Add(TempData, Trim(TempData.STKCOD & "-" & TempData.NUM_ONE))    ' KEy
      End If

      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   Set Rs = Nothing
   Set D = Nothing
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub

Public Sub LoadIncenPro05_1(Optional Cl1 As Collection = Nothing, Optional FromCMPLDat As Date = -1, Optional ToCMPLDat As Date = -1, Optional STKCOD As String)
On Error GoTo ErrorHandler
Dim D As CIncentivePromote
Dim ItemCount As Long
Dim TempData As CIncentivePromote
Dim Rs As ADODB.Recordset
Dim i As Long

   Set D = New CIncentivePromote
   Set Rs = New ADODB.Recordset
   
   D.FROM_CMPL_DATE = FromCMPLDat
   D.TO_CMPL_DATE = ToCMPLDat
   D.STKCOD = STKCOD
   Call D.QueryData(3, Rs, ItemCount)

   If Not (Cl1 Is Nothing) Then
      Set Cl1 = Nothing
      Set Cl1 = New Collection
   End If
   
'   If Not (Cl2 Is Nothing) Then
'      Set Cl2 = Nothing
'      Set Cl2 = New Collection
'   End If
   
   While Not Rs.EOF
      i = i + 1
      Set TempData = New CIncentivePromote
      Call TempData.PopulateFromRS(1, Rs)

      If Not (Cl1 Is Nothing) Then
          Call Cl1.Add(TempData)    ' ????????? key ????????????????????????
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   Set Rs = Nothing
   Set D = Nothing
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub


Public Sub LoadIncenPro05_2(Optional Cl1 As Collection = Nothing, Optional FromCMPLDat As Date = -1, Optional ToCMPLDat As Date = -1, Optional STKCOD As String)
On Error GoTo ErrorHandler
Dim D As CIncentivePromote
Dim ItemCount As Long
Dim TempData As CIncentivePromote
Dim Rs As ADODB.Recordset
Dim i As Long

   Set D = New CIncentivePromote
   Set Rs = New ADODB.Recordset
   
   D.FROM_CMPL_DATE = FromCMPLDat
   D.TO_CMPL_DATE = ToCMPLDat
   Call D.QueryData(5, Rs, ItemCount)

   If Not (Cl1 Is Nothing) Then
      Set Cl1 = Nothing
      Set Cl1 = New Collection
   End If
   
'   If Not (Cl2 Is Nothing) Then
'      Set Cl2 = Nothing
'      Set Cl2 = New Collection
'   End If
   
   While Not Rs.EOF
      i = i + 1
      Set TempData = New CIncentivePromote
      Call TempData.PopulateFromRS(5, Rs)

      If Not (Cl1 Is Nothing) Then
          Call Cl1.Add(TempData, TempData.STKCOD)   ' ????????? key ????????????????????????
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   Set Rs = Nothing
   Set D = Nothing
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub

Public Sub LoadCustomerLookup(C As ComboBox, Optional Cl As Collection = Nothing)
On Error GoTo ErrorHandler
Dim D As CARMas
Dim ItemCount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CARMas
Dim i As Long
   
   Set D = New CARMas
   Set Rs = New ADODB.Recordset
   
   D.OrderBy = 2
   Call D.QueryData(2, Rs, ItemCount)
   
   If Not (C Is Nothing) Then
      C.Clear
      i = 0
      C.AddItem ("")
   End If
   
   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   
   While Not Rs.EOF
      i = i + 1
      Set TempData = New CARMas
      Call TempData.PopulateFromRS(1, Rs, i)
      If Not (C Is Nothing) Then
         C.AddItem (TempData.CUSNAM)
         C.ItemData(i) = i
      End If
      
      If i <> 0 Then
            Set D = GetCustomerNoNew(Cl, Str(i))
            If D Is Nothing Then
              Call Cl.Add(TempData, Trim(Str(i)))
            End If
      End If
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   Set Rs = Nothing
   Set D = Nothing
   Exit Sub

ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub

Public Function GetCustomerNoNew(m_TempCol As Collection, TempKey As String) As CStcrd
On Error Resume Next
Dim Ei As CARMas
Static TempEi As CARMas

   Set Ei = m_TempCol(TempKey)
   If Ei Is Nothing Then
      Set GetCustomerNoNew = TempEi
   Else
      Set GetCustomerNoNew = Ei
   End If
End Function

Public Sub LoadCusFromAreaCom(C As ComboBox, AreaID As Long, Optional Cl As Collection = Nothing)
On Error GoTo ErrorHandler
Dim D As CCommissionCustomerArea
Dim ItemCount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CCommissionCustomerArea
Dim i As Long

   Set D = New CCommissionCustomerArea
   Set Rs = New ADODB.Recordset

   D.MASTER_AREA_ID = AreaID
   Call D.QueryData(2, Rs, ItemCount)
   
   If Not (C Is Nothing) Then
      C.Clear
      i = 0
      C.AddItem ("")
   End If
   
   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   
   While Not Rs.EOF
      i = i + 1
      Set TempData = New CCommissionCustomerArea
      Call TempData.PopulateFromRS(2, Rs)
      
      If Not (C Is Nothing) Then
      End If
      
      If Not (Cl Is Nothing) Then
         Call Cl.Add(TempData, TempData.COMMISSION_CUS_ID)
'         If TempData.COMMISSION_CUS_ID = "66-001" Then
'               ''debug.print TempData.COMMISSION_CUS_ID
'         End If
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   Set Rs = Nothing
   Set D = Nothing
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub
Public Sub LoadCusNonAreaCom(NonAreaFLag As String, YearID As Long, Optional Cl As Collection = Nothing)
On Error GoTo ErrorHandler
Dim D As CCommissionCustomerArea
Dim ItemCount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CCommissionCustomerArea
Dim i As Long

   Set D = New CCommissionCustomerArea
   Set Rs = New ADODB.Recordset

   D.YEAR_ID = YearID
   D.NONAREA = NonAreaFLag
   Call D.QueryData(4, Rs, ItemCount)
   
   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   
   While Not Rs.EOF
      i = i + 1
      Set TempData = New CCommissionCustomerArea
      Call TempData.PopulateFromRS(4, Rs)
      
      If Not (Cl Is Nothing) Then
         Call Cl.Add(TempData, Trim(TempData.COMMISSION_CUS_ID))
'         If TempData.COMMISSION_CUS_ID = "66-001" Then
'               ''debug.print TempData.COMMISSION_CUS_ID
'         End If
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   Set Rs = Nothing
   Set D = Nothing
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub

Public Sub LoadStkcodFromMasterID(C As ComboBox, goodsMaterID As Long, Optional Cl As Collection = Nothing)
On Error GoTo ErrorHandler
Dim D As CGoodsDetail
Dim ItemCount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CGoodsDetail
Dim i As Long

   Set D = New CGoodsDetail
   Set Rs = New ADODB.Recordset

   D.GOODS_MASTER_ID = goodsMaterID
   Call D.QueryData(1, Rs, ItemCount)
   
   If Not (C Is Nothing) Then
      C.Clear
      i = 0
      C.AddItem ("")
   End If
   
   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   
   While Not Rs.EOF
      i = i + 1
      Set TempData = New CGoodsDetail
      Call TempData.PopulateFromRS(1, Rs)
      
'      If Not (C Is Nothing) Then
'      End If
      
      If Not (Cl Is Nothing) Then
         Call Cl.Add(TempData, Trim(TempData.STKCOD) & "-" & Trim(TempData.GOODS_MASTER_ID))
'         If TempData.COMMISSION_CUS_ID = "66-001" Then
'               ''debug.print TempData.COMMISSION_CUS_ID
'         End If
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   Set Rs = Nothing
   Set D = Nothing
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub

Public Sub LoadStmasLookup(C As ComboBox, Optional Cl As Collection = Nothing)
On Error GoTo ErrorHandler
Dim D As CStmas
Dim ItemCount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CStmas
Dim i As Long
   
   Set D = New CStmas
   Set Rs = New ADODB.Recordset
   
  ' D.DOCNUM = DOCNUM                                ' ตัดว่า IV นี้ มีสินค้าตัวใดบ้าง
   Call D.QueryData(1, Rs, ItemCount)
   
   If Not (C Is Nothing) Then  '
      C.Clear
      i = 0
      C.AddItem ("")
   End If
   
   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   
   While Not Rs.EOF
      i = i + 1
      Set TempData = New CStmas
      Call TempData.PopulateFromRS(1, Rs, i)
      If Not (C Is Nothing) Then  '
         C.AddItem (TempData.STKDES)
         C.ItemData(i) = i
      End If
      
      If i <> 0 Then
            Set D = GetStmasNoNew(Cl, Trim(Str(i)))
            If D Is Nothing Then
                Call Cl.Add(TempData, Trim(Str(i)))             ' ถ้าแก้ตรงนี้งานจะเข้าตอน lookup
'                  If Not (Cl2 Is Nothing) Then
'                        Call Cl2.Add(TempData, Trim(TempData.DOCNUM))               ' ใช้คิวรี่ 8 ไม่ได้เพราะ IV จะซ้ำ
'                   End If
            End If
      End If
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   Set Rs = Nothing
   Set D = Nothing
   Exit Sub

ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub

Public Sub LoadGoodsDetailFromGM_GG(Optional Cl As Collection = Nothing, Optional GOODS_MASTER_ID As Long = -1, Optional GOODS_GROUP_ID As Long = -1)
On Error GoTo ErrorHandler
Dim D As CGoodsDetail
Dim ItemCount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CGoodsDetail
Dim i As Long

   Set D = New CGoodsDetail
   Set Rs = New ADODB.Recordset

   D.GOODS_MASTER_ID = GOODS_MASTER_ID
   D.GOODS_GROUP_ID = GOODS_GROUP_ID
   Call D.QueryData(2, Rs, ItemCount)
   
   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   
   While Not Rs.EOF
      i = i + 1
      Set TempData = New CGoodsDetail
      Call TempData.PopulateFromRS(2, Rs)
      
'      If Not (C Is Nothing) Then
'      End If
      
      If Not (Cl Is Nothing) Then
         Call Cl.Add(TempData, Trim(TempData.STKCOD))
'         If TempData.COMMISSION_CUS_ID = "66-001" Then
'               ''debug.print TempData.COMMISSION_CUS_ID
'         End If
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   Set Rs = Nothing
   Set D = Nothing
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub
Public Sub LoadGoodsDetailFromMaster(C As ComboBox, GOODS_MASTER_ID As Long, Optional Cl As Collection = Nothing)
On Error GoTo ErrorHandler
Dim D As CGoodsDetail
Dim ItemCount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CGoodsDetail
Dim i As Long

   Set D = New CGoodsDetail
   Set Rs = New ADODB.Recordset

   D.GOODS_MASTER_ID = GOODS_MASTER_ID
   Call D.QueryData(2, Rs, ItemCount)
   
   If Not (C Is Nothing) Then
      C.Clear
      i = 0
      C.AddItem ("")
   End If
   
   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   
   While Not Rs.EOF
      i = i + 1
      Set TempData = New CGoodsDetail
      Call TempData.PopulateFromRS(2, Rs)
      
'      If Not (C Is Nothing) Then
'      End If
      
      If Not (Cl Is Nothing) Then
         Call Cl.Add(TempData, Trim(TempData.GOODS_DETAIL_ID))
'         If TempData.COMMISSION_CUS_ID = "66-001" Then
'               ''debug.print TempData.COMMISSION_CUS_ID
'         End If
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   Set Rs = Nothing
   Set D = Nothing
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub

Public Sub LoadCusFromAreaNameCom(C As ComboBox, YearID As Long, Optional Cl As Collection = Nothing)
On Error GoTo ErrorHandler
Dim D As CCommissionCustomerArea
Dim ItemCount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CCommissionCustomerArea
Dim i As Long

   Set D = New CCommissionCustomerArea
   Set Rs = New ADODB.Recordset

   D.YEAR_ID = YearID
   Call D.QueryData(4, Rs, ItemCount)
   
   If Not (C Is Nothing) Then
      C.Clear
      i = 0
      C.AddItem ("")
   End If
   
   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   
   While Not Rs.EOF
      i = i + 1
      Set TempData = New CCommissionCustomerArea
      Call TempData.PopulateFromRS(4, Rs)
      
      If Not (C Is Nothing) Then
      End If
      
      If Not (Cl Is Nothing) Then
         Call Cl.Add(TempData, Trim(TempData.COMMISSION_CUS_ID))
'         If TempData.COMMISSION_CUS_ID = "66-001" Then
'               ''debug.print TempData.COMMISSION_CUS_ID
'         End If
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   Set Rs = Nothing
   Set D = Nothing
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub

Public Function GetCusAreaCom(m_TempCol As Collection, TempKey As String, Optional HaveNew As Boolean = True) As CCommissionCustomerArea
On Error Resume Next
Dim Ei As CCommissionCustomerArea
Static TempEi As CCommissionCustomerArea

   Set Ei = m_TempCol(TempKey)
  ' ''debug.print TempKey
    If Ei Is Nothing And HaveNew Then
                If TempEi Is Nothing Then
                   Set TempEi = New CCommissionCustomerArea
                End If
      Set GetCusAreaCom = TempEi
   Else
      Set GetCusAreaCom = Ei
   End If
End Function

Public Function GetStkcodGoodsDetail(m_TempCol As Collection, TempKey As String, Optional HaveNew As Boolean = True) As CGoodsDetail
On Error Resume Next
Dim Ei As CGoodsDetail
Static TempEi As CGoodsDetail

   Set Ei = m_TempCol(TempKey)
  ' ''debug.print TempKey
    If Ei Is Nothing And HaveNew Then
                If TempEi Is Nothing Then
                   Set TempEi = New CGoodsDetail
                End If
      Set GetStkcodGoodsDetail = TempEi
   Else
      Set GetStkcodGoodsDetail = Ei
   End If
End Function

Public Sub LoadMinusIV(Optional Cl As Collection = Nothing, Optional FromDocDat As Date = -1, Optional ToDocDat As Date = -1)
On Error GoTo ErrorHandler
Dim D As CComMinusStk
Dim ItemCount As Long
Dim TempData As CComMinusStk
Dim Rs As ADODB.Recordset
Dim i As Long

   Set D = New CComMinusStk
   Set Rs = New ADODB.Recordset
   
   D.FROM_DOC_DATE = FromDocDat
   D.TO_DOC_DATE = ToDocDat
   Call D.QueryData(1, Rs, ItemCount)   ' ดึงข้อมูลเพื่อเอาไปเก็บในคอเล็กชั่น

   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   
   While Not Rs.EOF
      i = i + 1
      Set TempData = New CComMinusStk
      Call TempData.PopulateFromRS(1, Rs)

      If Not (Cl Is Nothing) Then
'      If TempData.IV_COD = "IV0044465" Then
'         ''debug.print
'      End If
      
         Call Cl.Add(TempData, Trim(TempData.IV_DOCDAT & "-" & TempData.IV_COD & "-" & TempData.STK_COD))    ' KEy
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   Set Rs = Nothing
   Set D = Nothing
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub

Public Sub LoadComRecordJoin(Optional Cl As Collection = Nothing, Optional COMTYP As String = "", Optional FromDocDat As Date = -1, Optional ToDocDat As Date = -1)
On Error GoTo ErrorHandler
Dim D As CComRecord
Dim ItemCount As Long
Dim TempData As CComRecord
Dim Rs As ADODB.Recordset
Dim i As Long

   Set D = New CComRecord
   Set Rs = New ADODB.Recordset
   
   D.COMTYP = COMTYP
   D.FROM_DOC_DATE = DateToStringInt(FromDocDat)
   D.TO_DOC_DATE = DateToStringInt(ToDocDat)
   Call D.QueryData(1, Rs, ItemCount)   ' ดึงข้อมูลเพื่อเอาไปเก็บในคอเล็กชั่น

   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   
   While Not Rs.EOF
      i = i + 1
      Set TempData = New CComRecord
      Call TempData.PopulateFromRS(1, Rs)
      If Not (Cl Is Nothing) Then
         Call Cl.Add(TempData)       ' KEy      , Trim(TempData.SLMCOD & "-" & TempData.MASTER_AREA_ID & "-" & TempData.GOODS_GROUP_ID & "-" & TempData.COMTYP & "-" & TempData.TODAT)
      End If
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   Set Rs = Nothing
   Set D = Nothing
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub

Public Sub LoadComRecord(Optional Cl As Collection = Nothing, Optional COMTYP As String = "", Optional ToDocDat As Date = -1)
On Error GoTo ErrorHandler
Dim D As CComRecord
Dim ItemCount As Long
Dim TempData As CComRecord
Dim Rs As ADODB.Recordset
Dim i As Long

   Set D = New CComRecord
   Set Rs = New ADODB.Recordset
   
   D.COMTYP = COMTYP
   D.TO_DOC_DATE = DateToStringInt(ToDocDat)
   Call D.QueryData(2, Rs, ItemCount)   ' ดึงข้อมูลเพื่อเอาไปเก็บในคอเล็กชั่น

   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   
   While Not Rs.EOF
      i = i + 1
      Set TempData = New CComRecord
      Call TempData.PopulateFromRS(2, Rs)

      If Not (Cl Is Nothing) Then
'      If TempData.IV_COD = "IV0044465" Then
'         ''debug.print
'      End If
      
         Call Cl.Add(TempData, Trim(TempData.SLMCOD & "-" & TempData.MASTER_AREA_ID & "-" & TempData.GOODS_GROUP_ID))     ' KEy
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   Set Rs = Nothing
   Set D = Nothing
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub
Public Sub LoadIVcenter(Optional Cl As Collection = Nothing, Optional FromDocDat As Date = -1, Optional ToDocDat As Date = -1)
On Error GoTo ErrorHandler
Dim D As CComIVcenter
Dim ItemCount As Long
Dim TempData As CComIVcenter
Dim Rs As ADODB.Recordset
Dim i As Long

   Set D = New CComIVcenter
   Set Rs = New ADODB.Recordset
   
   D.FROM_DOC_DATE = FromDocDat
   D.TO_DOC_DATE = ToDocDat
   Call D.QueryData(1, Rs, ItemCount)   ' ดึงข้อมูลเพื่อเอาไปเก็บในคอเล็กชั่น

   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   
   While Not Rs.EOF
      i = i + 1
      Set TempData = New CComIVcenter
      Call TempData.PopulateFromRS(1, Rs)

      If Not (Cl Is Nothing) Then
         Call Cl.Add(TempData, Trim(TempData.IV_COD))     ' KEy
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   Set Rs = Nothing
   Set D = Nothing
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub


Public Sub LoadIVcredit(Optional Cl As Collection = Nothing, Optional FromDocDat As Date = -1, Optional ToDocDat As Date = -1)
On Error GoTo ErrorHandler
Dim D As CComIVcredit
Dim ItemCount As Long
Dim TempData As CComIVcredit
Dim Rs As ADODB.Recordset
Dim i As Long

   Set D = New CComIVcredit
   Set Rs = New ADODB.Recordset
   
   D.FROM_DOC_DATE = FromDocDat
   D.TO_DOC_DATE = ToDocDat
   Call D.QueryData(1, Rs, ItemCount)   ' ดึงข้อมูลเพื่อเอาไปเก็บในคอเล็กชั่น

   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   
   While Not Rs.EOF
      i = i + 1
      Set TempData = New CComIVcredit
      Call TempData.PopulateFromRS(1, Rs)

      If Not (Cl Is Nothing) Then
         Call Cl.Add(TempData, Trim(TempData.IV_COD))     ' KEy
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   Set Rs = Nothing
   Set D = Nothing
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub

Public Sub LoadAllSaleInIV(Optional Cl As Collection = Nothing, Optional FromDocDat As Date = -1, Optional ToDocDat As Date = -1)
On Error GoTo ErrorHandler
Dim D As CStcrd
Dim ItemCount As Long
Dim TempData As CStcrd
Dim Rs As ADODB.Recordset
Dim i As Long

   Set D = New CStcrd
   Set Rs = New ADODB.Recordset
   
   D.FROM_DOC_DATE = FromDocDat
   D.TO_DOC_DATE = ToDocDat
   Call D.QueryData(15, Rs, ItemCount)

   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   
   While Not Rs.EOF
      i = i + 1
      Set TempData = New CStcrd
      Call TempData.PopulateFromRS(15, Rs)

      If Not (Cl Is Nothing) Then
         Call Cl.Add(TempData, Trim(TempData.SLMCOD))    ' KEy
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   Set Rs = Nothing
   Set D = Nothing
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub

Public Sub LoadIncenForSum(Optional Cl As Collection = Nothing, Optional FromDocDat As Date = -1, Optional ToDocDat As Date = -1)
On Error GoTo ErrorHandler
Dim D As CCondiIncenSum
Dim ItemCount As Long
Dim TempData As CCondiIncenSum
Dim Rs As ADODB.Recordset
Dim i As Long

   Set D = New CCondiIncenSum
   Set Rs = New ADODB.Recordset
   
   D.FROM_CMPL_DATE = FromDocDat
   D.TO_CMPL_DATE = ToDocDat
   D.GOODS_GROUP_ID = 1
   Call D.QueryData(1, Rs, ItemCount)   ' ดึงข้อมูลเพื่อเอาไปเก็บในคอเล็กชั่น

   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   
   While Not Rs.EOF
      i = i + 1
      Set TempData = New CCondiIncenSum
      Call TempData.PopulateFromRS(1, Rs)

      If Not (Cl Is Nothing) Then
         Call Cl.Add(TempData, Trim(TempData.SLMCOD))     ' KEy
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   Set Rs = Nothing
   Set D = Nothing
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub

Public Sub LoadNewCus(Optional Cl As Collection = Nothing, Optional FromDocDat As Date = -1, Optional ToDocDat As Date = -1, Optional PEOPLE As String)
On Error GoTo ErrorHandler
Dim D As CStcrd
Dim ItemCount As Long
Dim TempData As CStcrd
Dim Rs As ADODB.Recordset
Dim i As Long

   Set D = New CStcrd
   Set Rs = New ADODB.Recordset
   
   D.FROM_DOC_DATE = FromDocDat
   D.TO_DOC_DATE = ToDocDat
   D.PEOPLE = PEOPLE
   Call D.QueryData(12, Rs, ItemCount)   ' ดึงข้อมูลเพื่อเอาไปเก็บในคอเล็กชั่น

   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   
   While Not Rs.EOF
      i = i + 1
      Set TempData = New CStcrd
      Call TempData.PopulateFromRS(12, Rs)

      If Not (Cl Is Nothing) Then
         Call Cl.Add(TempData, Trim(TempData.CUSCOD))      ' KEy
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   Set Rs = Nothing
   Set D = Nothing
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub

Public Function GetMinusCommiss(m_TempCol As Collection, TempKey As String, Optional HaveNew As Boolean = True) As CComMinusStk
On Error Resume Next
Dim Ei As CComMinusStk
Static TempEi As CComMinusStk

   Set Ei = m_TempCol(TempKey)
   ''debug.print TempKey
    If Ei Is Nothing And HaveNew Then
                If TempEi Is Nothing Then
                   Set TempEi = New CComMinusStk
                End If
      Set GetMinusCommiss = TempEi
   Else
      Set GetMinusCommiss = Ei
   End If
End Function

Public Function GetIVcenter(m_TempCol As Collection, TempKey As String, Optional HaveNew As Boolean = True) As CComIVcenter
On Error Resume Next
Dim Ei As CComIVcenter
Static TempEi As CComIVcenter

   Set Ei = m_TempCol(TempKey)
  ' ''debug.print TempKey
    If Ei Is Nothing And HaveNew Then
                If TempEi Is Nothing Then
                   Set TempEi = New CComIVcenter
                End If
      Set GetIVcenter = TempEi
   Else
      Set GetIVcenter = Ei
   End If
End Function

Public Function GetIVcredit(m_TempCol As Collection, TempKey As String, Optional HaveNew As Boolean = True) As CComIVcredit
On Error Resume Next
Dim Ei As CComIVcredit
Static TempEi As CComIVcredit

   Set Ei = m_TempCol(TempKey)
  ' ''debug.print TempKey
    If Ei Is Nothing And HaveNew Then
                If TempEi Is Nothing Then
                   Set TempEi = New CComIVcredit
                End If
      Set GetIVcredit = TempEi
   Else
      Set GetIVcredit = Ei
   End If
End Function


Public Sub InitComPayed(C As ComboBox)
   C.Clear
   
   C.AddItem ("ยังไม่ชำระ")
   C.ItemData(0) = 0
   
   C.AddItem ("แบ่งชำระ")
   C.ItemData(1) = 1
   
End Sub
Public Sub LoadREDocDat(Optional Cl As Collection = Nothing, Optional FromDocDat As Date = -1, Optional ToDocDat As Date = -1, Optional db2 As Boolean = False)
On Error GoTo ErrorHandler
Dim D As CARTrn
Dim ItemCount As Long
Dim TempData As CARTrn
Dim Rs As ADODB.Recordset
Dim i As Long

   Set D = New CARTrn
   Set Rs = New ADODB.Recordset
   
   D.FROM_CMPL_DATE = FromDocDat
   D.TO_CMPL_DATE = ToDocDat
   D.db2 = db2
   '                 D.SLMCOD = "02"
   Call D.QueryData(12, Rs, ItemCount)   ' ดึงข้อมูลเพื่อเอาไปเก็บในคอเล็กชั่น

'   If Not (Cl Is Nothing) Then
'      Set Cl = Nothing
'      Set Cl = New Collection
'   End If
   
   While Not Rs.EOF
      i = i + 1
      Set TempData = New CARTrn
      Call TempData.PopulateFromRS(13, Rs)
      
'      'debug.print TempData.DOCDAT & "-" & TempData.DOCNUM
      If Not (Cl Is Nothing) Then
         Call Cl.Add(TempData, Trim(TempData.DOCNUM))      ' KEy
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   Set Rs = Nothing
   Set D = Nothing
   Exit Sub
   
ErrorHandler:
   'debug.print TempData.DOCNUM
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub

Public Function GetREDocDat(m_TempCol As Collection, TempKey As String, Optional HaveNew As Boolean = True) As CARTrn
On Error Resume Next
Dim Ei As CARTrn
Static TempEi As CARTrn

   Set Ei = m_TempCol(TempKey)
   ''debug.print TempKey
    If Ei Is Nothing And HaveNew Then
                If TempEi Is Nothing Then
                   Set TempEi = New CARTrn
                End If
      Set GetREDocDat = TempEi
   Else
      Set GetREDocDat = Ei
   End If
End Function

Public Sub LoadCustomerPro(Cl As Collection)
On Error GoTo ErrorHandler
Dim D As CARMas
Dim ItemCount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CARMas
Dim i As Long

   Set D = New CARMas
   Set Rs = New ADODB.Recordset

   Call D.QueryData(1, Rs, ItemCount)

   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If

   While Not Rs.EOF
      i = i + 1
      Set TempData = New CARMas
      Call TempData.PopulateFromRS(1, Rs)

      If Not (Cl Is Nothing) Then
         Call Cl.Add(TempData, TempData.CUSCOD)
      End If

      Set TempData = Nothing
      Rs.MoveNext
   Wend

   Set Rs = Nothing
   Set D = Nothing
   Exit Sub

ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub

Public Sub LoadSalePro(Cl As Collection)
On Error GoTo ErrorHandler
Dim D As COESLM
Dim ItemCount As Long
Dim Rs As ADODB.Recordset
Dim TempData As COESLM
Dim i As Long

   Set D = New COESLM
   Set Rs = New ADODB.Recordset

   Call D.QueryData(1, Rs, ItemCount)

   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If

   While Not Rs.EOF
      i = i + 1
      Set TempData = New COESLM
      Call TempData.PopulateFromRS(1, Rs)

      If Not (Cl Is Nothing) Then
         Call Cl.Add(TempData, TempData.SLMCOD)
      End If

      Set TempData = Nothing
      Rs.MoveNext
   Wend

   Set Rs = Nothing
   Set D = Nothing
   Exit Sub

ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub
'
Public Sub LoadSale(C As ComboBox, Optional Cl As Collection = Nothing, Optional SaleCode As String = "")
On Error GoTo ErrorHandler
Dim D As COESLM
Dim ItemCount As Long
Dim Rs As ADODB.Recordset
Dim TempData As COESLM
Dim i As Long

   Set D = New COESLM
   Set Rs = New ADODB.Recordset

   Call D.QueryData(1, Rs, ItemCount)

   If Not (C Is Nothing) Then
      C.Clear
      i = 0
      C.AddItem ("")
   End If

   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If

   While Not Rs.EOF
      i = i + 1
      Set TempData = New COESLM
      Call TempData.PopulateFromRS(1, Rs)

      If Not (C Is Nothing) Then
          C.AddItem (TempData.SLMCOD & " - " & TempData.SLMNAM)
         C.ItemData(i) = i
      End If

      If Not (Cl Is Nothing) Then
         Call Cl.Add(TempData, TempData.SLMCOD)
      End If

      Set TempData = Nothing
      Rs.MoveNext
   Wend

   Set Rs = Nothing
   Set D = Nothing
   Exit Sub

ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub

Public Function GetSlm(m_TempCol As Collection, TempKey As String, Optional HaveNew As Boolean = True) As COESLM
On Error Resume Next
Dim Ei As COESLM
Static TempEi As COESLM

   Set Ei = m_TempCol(TempKey)
   ''debug.print TempKey
    If Ei Is Nothing And HaveNew Then
                If TempEi Is Nothing Then
                   Set TempEi = New COESLM
                End If
      Set GetSlm = TempEi
   Else
      Set GetSlm = Ei
   End If
End Function

Public Sub InitCommissionOrderBy(C As ComboBox)
   C.Clear
   
   C.AddItem ("")
   C.ItemData(0) = 0
   
   C.AddItem (MapText("เลขที่"))
   C.ItemData(1) = 1

   C.AddItem (MapText("วันที่เริ่มใช้"))
   C.ItemData(2) = 2
   
   C.AddItem (MapText("วันที่สิ้นสุด"))
   C.ItemData(3) = 3

End Sub

Public Sub LoadCommissionChart(C As ComboBox, Optional Cl As Collection = Nothing, Optional FK_ID As Long = -1)
On Error GoTo ErrorHandler
Dim D As CCommissionChart
Dim ItemCount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CCommissionChart
Dim i As Long
Dim S As COESLM
Dim m_SaleName As Collection
Dim saleName As String
   
   If FK_ID <= 0 Then
      Exit Sub
   End If
   Set D = New CCommissionChart
   Set Rs = New ADODB.Recordset
   Set S = New COESLM
   Set m_SaleName = New Collection
   
'   D.COMMISSION_CHART_ID = -1
   D.MASTER_FROMTO_ID = FK_ID
   D.ORDER_TYPE = 1
   Call D.QueryData(1, Rs, ItemCount)
   
    Call LoadSale(Nothing, m_SaleName)

   If Not (C Is Nothing) Then
      C.Clear
      i = 0
      C.AddItem ("")
   End If

   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   While Not Rs.EOF
      i = i + 1
      Set TempData = New CCommissionChart
      Call TempData.PopulateFromRS(1, Rs)
   
      If Not (C Is Nothing) Then
            Set S = GetSlm(m_SaleName, TempData.SALE_ID, False)
            If Not (S Is Nothing) Then
                       saleName = S.SLMNAM
            Else
                       saleName = ""
            End If
   
         C.AddItem (TempData.COMMISSION_CHART_ID & " - " & saleName)
         C.ItemData(i) = TempData.COMMISSION_CHART_ID
      End If
      
      If Not (Cl Is Nothing) Then
         Call Cl.Add(TempData, Trim(Str(TempData.COMMISSION_CHART_ID)))
      End If
      
      Set TempData = Nothing
       Rs.MoveNext
   Wend
   
   Set Rs = Nothing
   Set D = Nothing
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub

''Public Sub LoadMasterFromTo(C As ComboBox, Optional Cl As Collection = Nothing, Optional MasterFromToType As MASTER_COMMISSION_AREA, Optional KeyType As Long = -1)
''On Error GoTo ErrorHandler
''Dim ItemCount As Long
''Dim Rs As ADODB.Recordset
''Dim TempData As CMasterFromTo
''Dim I As Long
''Dim D As CMasterFromTo
''
''   Set Rs = Nothing
''   Set Rs = New ADODB.Recordset
''
''   Set D = New CMasterFromTo
''  D.MASTER_FROMTO_TYPE = MasterFromToType
''   Call D.QueryData(1, Rs, ItemCount)
''
''   If Not (C Is Nothing) Then
''      C.Clear
''      I = 0
''      C.AddItem ("")
''   End If
''
''   If Not (Cl Is Nothing) Then
''      Set Cl = Nothing
''      Set Cl = New Collection
''   End If
''   While Not Rs.EOF
''      I = I + 1
''      Set TempData = New CMasterFromTo
''      Call TempData.PopulateFromRS(1, Rs)
''
''      If Not (C Is Nothing) Then
''         C.AddItem (TempData.MASTERFROMTO_DESC)
''         C.ItemData(I) = TempData.MASTER_FROMTO_ID
''      End If
''
''      If Not (Cl Is Nothing) Then
''         Call Cl.Add(TempData, Trim(Str(TempData.MASTER_FROMTO_ID)))
''      End If
''
''      Set TempData = Nothing
''      Rs.MoveNext
''   Wend
''
''   If Rs.State = adStateOpen Then
''      Rs.Close
''   End If
''   Set Rs = Nothing
''   Set D = Nothing
''
''   Exit Sub
''
''ErrorHandler:
''   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
''   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
''End Sub

Public Sub LoadSaleLookup(C As ComboBox, Optional Cl As Collection = Nothing)
On Error GoTo ErrorHandler
Dim D As COESLM
Dim ItemCount As Long
Dim Rs As ADODB.Recordset
Dim TempData As COESLM
Dim i As Long
   
   Set D = New COESLM
   Set Rs = New ADODB.Recordset
   
   Call D.QueryData(1, Rs, ItemCount)
   
   If Not (C Is Nothing) Then
      C.Clear
      i = 0
      C.AddItem ("")
   End If
   
   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   
   While Not Rs.EOF
      i = i + 1
      Set TempData = New COESLM
      
'      If i = 20 Then
'         Debug.Print
'      End If
      
      Call TempData.PopulateFromRS(1, Rs)
      If Not (C Is Nothing) Then
         C.AddItem (Trim(TempData.SLMNAM))
         C.ItemData(i) = Str(TempData.SLMCOD)                        'รหัส Sale ที่เพิ่มใหม่ใน Express ไม่สามารถใส่เป็นตัวหนังสือได้ต้องใส่เป็นเลขเท่านั้น เช่น   02NE1 ไม่ได้ ต้องใส่เป็น 0201 หรือ 0202 เท่านั้น
        ' i = i - 1
         'C.ItemData(i) = i
      End If
      
      If i <> 0 Then
            Set D = GetSaleNoNew(Cl, Trim(TempData.SLMCOD))
            If D Is Nothing Then
              Call Cl.Add(TempData, Trim(Str(TempData.SLMCOD)))
              'Call Cl.Add(TempData, Trim(Str(i)))
             ' Call Cl.Add(TempData)
            End If
      End If
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   Set Rs = Nothing
   Set D = Nothing
   Exit Sub

ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub


Public Function GetSaleNoNew(m_TempCol As Collection, TempKey As String) As COESLM
On Error Resume Next
Dim Ei As COESLM
Static TempEi As COESLM

   Set Ei = m_TempCol(TempKey)
   If Ei Is Nothing Then
'      If TempEi Is Nothing Then
'         Set TempEi = New Cstcrd
'      End If
      Set GetSaleNoNew = TempEi
   Else
      Set GetSaleNoNew = Ei
   End If
End Function

Public Sub LoadcollTotal(Optional FROM_DOC_DATE As Date = -1, Optional TO_DOC_DATE As Date = -1, Optional collTotal1 As Collection, Optional coll_Minus As Collection, Optional Allflag As Boolean = True)
Dim CMPLDAT As Date
Dim PrevKey1 As String
Dim PrevKey2 As String
Dim PrevKey3 As Long
Dim PrevKey4 As String
Dim PrevKey5 As String
Dim toCMPLdate As Date
Dim NETVAL As Double
Dim Rs As ADODB.Recordset
Dim TempRs As ADODB.Recordset
Dim Total1 As Double
Dim Total2 As Double
Dim Total3 As Double
Dim Total4 As Double
Dim L As Long
Dim iCount As Long
Dim IsOK As Boolean
Dim totMinus As Double

Dim DueCount As Long
Dim strTest As String

Dim ArS As COESLM
Dim Stcrd As CStcrd
Dim TempCConditionCommiss As CConditionCommission
Dim tempMinusStkcod As CComMinusStk
Dim m_runConditionCommiss As CConditionCommission
Dim tempCusArea As CCommissionCustomerArea
Dim temp_Area As CCommissionCustomerArea
Dim temp_Area2 As CCommissionCustomerArea
Dim D As CChartTotal
Dim C As CChartTotal
Dim m_MinusStkcod As Collection
Dim m_ConditionCommiss4 As Collection
'Dim m_cusFromArea(20) As Collection
Dim m_IVincomplete0 As Collection
Dim m_IVincomplete12 As Collection
Dim PercentNum1 As Double
Dim m_AreaCod As Collection
Dim m_ComDonStkcod As Collection
Dim tempComDonStkcod As CComDonStk
Dim CorrectStkcod As Boolean

Dim m_CusE As Collection
Dim m_CusF As Collection
Dim E As CChartSubTotal
Dim F As CChartSubTotal
Dim YEAR_ID As Long

Dim m_ComMasPro As Collection
Dim m_ComPro1 As Collection
Dim comMasPro As CCommissMasterPromote
Dim comPro1 As CComMasSubPromote

Dim REsumIV As CARRcIt
Dim m_REsumIV As Collection
Dim m_ReDocdat As Collection
Dim tempREdoc As CARTrn

Dim m_IVcenter As Collection
Dim m_IVinArea As Collection
Dim tempIVcenter As CComIVcenter
Dim IVoldinArea As Boolean
Dim stcrd_IVinArea As CStcrd

Dim temp_GoodsGroup As CGoodsGroup
Dim temp_GoodsInGroup As CGoodsDetail
Dim m_GoodsGroup As Collection
Dim n_array As Long
Dim m_GoodsInGroup(10) As Collection
Dim num As Long
Dim GOODS_GROUP_ID As Long
Dim CorrectGroup As Boolean

   Set m_IVcenter = New Collection
   Set m_IVinArea = New Collection

   Set Rs = New ADODB.Recordset
   Set TempRs = New ADODB.Recordset
   Set ArS = New COESLM
   Set Stcrd = New CStcrd
   Set TempCConditionCommiss = New CConditionCommission
   Set tempMinusStkcod = New CComMinusStk
   Set m_runConditionCommiss = New CConditionCommission
   Set temp_Area = New CCommissionCustomerArea
   Set temp_Area2 = New CCommissionCustomerArea
   Set tempCusArea = New CCommissionCustomerArea
   Set D = New CChartTotal
   Set m_MinusStkcod = New Collection
   Set m_ConditionCommiss4 = New Collection

   Set m_IVincomplete0 = New Collection
   Set m_IVincomplete12 = New Collection
   Set m_AreaCod = New Collection
   Set collTotal1 = New Collection
   Set coll_Minus = New Collection
   Set m_ComDonStkcod = New Collection
   Set tempComDonStkcod = New CComDonStk
   
   Set m_CusE = New Collection
   Set m_CusF = New Collection
   Set E = New CChartSubTotal
   Set F = New CChartSubTotal
   
   Set m_ComMasPro = New Collection
   Set m_ComPro1 = New Collection
   Set comMasPro = New CCommissMasterPromote
   Set comPro1 = New CComMasSubPromote
   
   Set REsumIV = New CARRcIt
   Set m_REsumIV = New Collection
   Set m_ReDocdat = New Collection
   Set tempREdoc = New CARTrn
   
   Set temp_GoodsGroup = New CGoodsGroup
   Set m_GoodsGroup = New Collection
   For n_array = 0 To UBound(m_GoodsInGroup)
       Set m_GoodsInGroup(n_array) = New Collection
   Next n_array
   Set temp_GoodsInGroup = New CGoodsDetail
   
   Call LoadComNamePro(m_ComMasPro, FROM_DOC_DATE, TO_DOC_DATE)       'รายชื่อ เซลล์และลูกค้าคนไหนทีเข้าข่าย
   Call LoadComPro("01", m_ComPro1, FROM_DOC_DATE, TO_DOC_DATE)      'เก็บเงินธรรมดา
   
  Call LoadIVcenter(m_IVcenter, FROM_DOC_DATE, TO_DOC_DATE)
  Call LoadComDonStk(m_ComDonStkcod, FROM_DOC_DATE, TO_DOC_DATE)
  Call LoadCommission04(m_ConditionCommiss4, FROM_DOC_DATE, TO_DOC_DATE)
  Call LoadMinusIV(m_MinusStkcod, FROM_DOC_DATE, TO_DOC_DATE)

  Call LoadREsumIV(Nothing, m_REsumIV)
  Call LoadREDocDat(m_ReDocdat) ' ต้องเอามาหมด

  Call LoadYearId(YEAR_ID, TO_DOC_DATE, TO_DOC_DATE)   'FROM_DOC_DATE,
  Call LoadAreaComReport(Nothing, m_AreaCod, YEAR_ID)
  
  n_array = 0
  Call LoadColGoodsGroup(m_GoodsGroup)                ' มีประเภทไรบ้าง    ' ดึง ชื่อสินค้าที่อยู่ใน master นั่นน และกรุ๊ปนั่นนนน
  For Each temp_GoodsGroup In m_GoodsGroup
      Set TempCConditionCommiss = m_ConditionCommiss4.ITEM(1)   ' ในกลุ่มนั้นมีสินค้าอะไรบ้าง
      Call LoadGoodsDetailFromGM_GG(m_GoodsInGroup(n_array), TempCConditionCommiss.GOODS_MASTER_ID, temp_GoodsGroup.GOODS_GROUP_ID)                 ' มีประเภทไรบ้าง
      n_array = n_array + 1
 Next temp_GoodsGroup
      
   Set ArS = New COESLM
   Call glbDaily.QuerySale(ArS, TempRs, iCount, IsOK, glbErrorLog)

   While Not TempRs.EOF          ' sale

         Call ArS.PopulateFromRS(1, TempRs)

         PrevKey1 = ArS.SLMCOD
         PrevKey2 = ArS.SLMNAM
         
     L = 0              ' ???????????????? ????????????
     For Each temp_Area In m_AreaCod
     L = L + 1
     
      Set Stcrd = New CStcrd
      Stcrd.FROM_DOC_DATE = FROM_DOC_DATE
      Stcrd.TO_DOC_DATE = TO_DOC_DATE
      Stcrd.SLMCOD = ArS.SLMCOD
      Call Stcrd.QueryData(14, Rs, iCount)

      While Not Rs.EOF                     ' ?????????????????
               CorrectStkcod = False
               Call Stcrd.PopulateFromRS(14, Rs)
            
     If Allflag = False Then
            CorrectGroup = False
            Set temp_GoodsInGroup = GetObject("CGoodsGroup", m_GoodsInGroup(1), Trim(Stcrd.STKCOD))         ' ฮาร์ตโค๊ดอีกแล้ว
            If Not (temp_GoodsInGroup Is Nothing) Then
               CorrectGroup = True
            End If
     Else
               CorrectGroup = True
     End If
            
      Set tempComDonStkcod = GetComDonStk(m_ComDonStkcod, Trim(Stcrd.STKCOD), False)
      If (tempComDonStkcod Is Nothing) And CorrectGroup = True Then
         CorrectStkcod = True                               ' สินค้านี้ คิดค่าคอมได้
      End If

      Set tempCusArea = GetCusAreaCom(temp_Area.ImportExportItems, Stcrd.CUSCOD, False)
      If (Not (tempCusArea Is Nothing)) And CorrectStkcod Then                     ' ลูกค้านี้อยู่ในเขต
   
   If ArS.SLMCOD = "15" And tempCusArea.MASTER_AREA_ID = 13 And Stcrd.DOCNUM = "IV0045140" Then
   ''debug.print
End If
   
      ' คำนวณดู IV ก่อน
      Set tempIVcenter = GetIVcenter(m_IVcenter, Stcrd.DOCNUM, False)        '
      If (tempIVcenter Is Nothing) Then   'ไม่เจอในคอเล็คชั่น = ไม่ใช่สินค้าพิเศษ
               IVoldinArea = True
      Else:
               IVoldinArea = False
               ' ดังนั้นต้องเก็บ ไว้ในคอล
               Set stcrd_IVinArea = New CStcrd
               stcrd_IVinArea.AREACOD = Trim(Str(tempIVcenter.MASTER_AREA_ID))
               stcrd_IVinArea.AREANAM = tempIVcenter.MASTER_AREA_NAME
               stcrd_IVinArea.NETVAL = Stcrd.NETVAL
               stcrd_IVinArea.DOCNUM = Stcrd.DOCNUM
               stcrd_IVinArea.CUSNAM = Stcrd.CUSNAM
               stcrd_IVinArea.STKDES = Stcrd.STKDES
               stcrd_IVinArea.DOCDAT = Stcrd.DOCDAT
               stcrd_IVinArea.STKCOD = Stcrd.STKCOD
               stcrd_IVinArea.SLMCOD = Stcrd.SLMCOD
               stcrd_IVinArea.CUSCOD = Stcrd.CUSCOD
               Call m_IVinArea.Add(stcrd_IVinArea)
               Set stcrd_IVinArea = Nothing
      End If
 
 If IVoldinArea = True Then  ' ******
               If tempCusArea.MASTER_AREA_ID <> PrevKey3 Or PrevKey5 <> Stcrd.CUSCOD Then  '
                     If PrevKey3 <> 0 And Total1 <> 0 And Total2 <> 0 Then
                                 Set F = New CChartSubTotal
                                 F.SALE_ID = PrevKey1
                                 F.AREA_ID = PrevKey3
                                 F.CUS_ID = PrevKey5
                                 F.TOTAL1_AMOUNT = Total1              'รวมของลูกค้า 1.ยอดจริง
                                 F.TOTAL2_AMOUNT = Total2              'รวมของลูกค้า 2.ยอดประเมิน
                                 Call m_CusF.Add(F) '& "-" & D.CUS_ID
                                 Set F = Nothing
                     End If
                End If
            
           If tempCusArea.MASTER_AREA_ID <> PrevKey3 Then
                     If PrevKey3 <> 0 And Total1 <> 0 And Total2 <> 0 Then
                                 Set D = New CChartTotal
                                 D.SALE_ID = PrevKey1
                                 D.AREA_ID = PrevKey3
                                 D.TOTAL1_SUM = Total3                      ' รวมของเซลล์ 1ยอดจริง
                                 D.TOTAL2_SUM = Total4                        ' รวมของเซลล์ 1ยอดประเมิน
                                 Set D.m_Cus = New Collection
                                 Set D.m_Cus = m_CusF
                                 Call collTotal1.Add(D, Trim(D.SALE_ID & "-" & D.AREA_ID)) '& "-" & D.CUS_ID
                                 Set D = Nothing
                                 Set m_CusF = Nothing
                                 Set m_CusF = New Collection
                     End If
                      Total1 = 0
                     Total2 = 0
                     Total3 = 0
                      Total4 = 0
               End If
               
               If ArS.SLMCOD = PrevKey1 And tempCusArea.MASTER_AREA_ID <> PrevKey3 And totMinus <> 0 Then         ' !
                     Set C = New CChartTotal                                                     ' ตรงนี้ต้องเช็ค   ArS.SLMCOD=PrevKey1 and temp_Area.MASTER_AREA_ID= PrevKey3
                     C.SALE_ID = PrevKey1  ' ArS.SLMCOD
                     C.AREA_ID = PrevKey3  'temp_Area.MASTER_AREA_ID           ' แก้ Prev3 แระ
                     C.MINUS = Val(totMinus)
                      Call coll_Minus.Add(C, Trim(C.SALE_ID & "-" & C.AREA_ID))
                      Set C = Nothing
                      totMinus = 0
'
'If PrevKey1 = "15" And PrevKey3 = 13 Then
'strTest = strTest & "  , " & Stcrd.DOCNUM
'End If
'
                End If

               Set tempMinusStkcod = GetMinusCommiss(m_MinusStkcod, Stcrd.DOCDAT & "-" & Stcrd.DOCNUM & "-" & Stcrd.STKCOD, False)
               ' ''debug.print tempCusArea.MASTER_AREA_ID & " __ " & PrevKey3
               If (tempMinusStkcod Is Nothing) Then
                        NETVAL = Stcrd.NETVAL                     ' สินค้านี้ไม่มีส่วนลด
               Else:
                        NETVAL = (Stcrd.NETVAL + Val(tempMinusStkcod.MINUS_AMOUNT))
If PrevKey1 = "15" And PrevKey3 = 13 Then
strTest = strTest & "  , " & tempMinusStkcod.MINUS_AMOUNT
End If
                        If ArS.SLMCOD = PrevKey1 Then ' And (tempCusArea.MASTER_AREA_ID = PrevKey3 Or PrevKey3 = 0) Then        ' !
                                 totMinus = totMinus + tempMinusStkcod.MINUS_AMOUNT

''                          Else
''                                 Set C = New CChartTotal                                                     ' ตรงนี้ต้องเช็ค   ArS.SLMCOD=PrevKey1 and temp_Area.MASTER_AREA_ID= PrevKey3
''                                 C.SALE_ID = PrevKey1  ' ArS.SLMCOD
''                                 C.AREA_ID = temp_Area.MASTER_AREA_ID '  PrevKey3          ' แก้ Prev3 แระ
''                                 C.MINUS = Val(totMinus)
''                                  Call coll_Minus.Add(C, Trim(C.SALE_ID & "-" & C.AREA_ID))
''                                  Set C = Nothing
''                                  totMinus = 0
                          End If
                          
                          
            End If

               Total1 = Total1 + NETVAL
               Total3 = Total3 + NETVAL
            
               Set m_runConditionCommiss = New CConditionCommission

               Set TempCConditionCommiss = GetCheckCommiss(m_ConditionCommiss4, Stcrd.STKCOD, False)
               If (TempCConditionCommiss Is Nothing) Then
                             PercentNum1 = (100 / 100)
               Else
                             PercentNum1 = (TempCConditionCommiss.SLM_PERCENT / 100)
               End If
               
               ' ยอดประเมิน วนลูป >= ว่าเข้าเงื่อนไขไหน
               Total2 = Total2 + (NETVAL * PercentNum1)
               Total4 = Total4 + (NETVAL * PercentNum1)
              
       ' เป็น sale และ ลูกค้าที่มีในเงื่อนไขพิเศษ
                Set comMasPro = GetObject("CCommissMasterPromote", m_ComMasPro, Trim(ArS.SLMCOD & "-" & Stcrd.CUSCOD))
                If Not (comMasPro Is Nothing) Then                                       ' เป็นลูกค้า AND ที่ใช้เงื่อนไขพิเศษ
                    ' คิดวันที่จ่ายตังค์
                    Set REsumIV = GetARRcpItem(m_REsumIV, Stcrd.DOCNUM, False)
                    If Not (REsumIV Is Nothing) Then                       ' มี RE ที่จ่าย IV นี้
                              ' ''debug.print REsumIV.RCPNUM
                              Set tempREdoc = GetREDocDat(m_ReDocdat, Stcrd.DOCNUM, False)
                              If Not (tempREdoc Is Nothing) Then
                                  CMPLDAT = tempREdoc.DOCDAT      ' ดึงวันที่ RE ออกมา
                              Else
                                 CMPLDAT = -1
                              End If
                              
                              If ((REsumIV.RCVAMT - NETVAL) >= 0) And CMPLDAT > 0 Then   'จ่ายครบยอด และ มีวันที่จ่าย
                                    DueCount = DateDiff("D", Stcrd.DOCDAT, CMPLDAT)  'ต้องคิดตรงนี้เพื่อไปคำนวณค่า
                              End If
                              
                     End If       '
               End If    ' เป็นลูกค้าที่ใช้เงื่อนไขพิเศษ
               
               
                 PrevKey3 = tempCusArea.MASTER_AREA_ID
                 Set temp_Area2 = GetAreaCom(m_AreaCod, tempCusArea.MASTER_AREA_ID)
                 PrevKey4 = temp_Area2.MASTER_AREA_NAME '!
                 PrevKey5 = Stcrd.CUSCOD
                 
 End If
       End If  '****** IVinArea
                Rs.MoveNext
         Wend
Next temp_Area
         
         For Each stcrd_IVinArea In m_IVinArea
         If Val(stcrd_IVinArea.AREACOD) = PrevKey3 And stcrd_IVinArea.SLMCOD = PrevKey1 Then
'                          ' คำนวณส่วนลดก่อน
'               Set tempMinusStkcod = GetMinusCommiss(m_MinusStkcod, stcrd_IVinArea.DOCDAT & "-" & stcrd_IVinArea.DOCNUM & "-" & stcrd_IVinArea.STKCOD, False)
'               If (tempMinusStkcod Is Nothing) Then   'ไม่เจอในคอเล็คชั่น = ไม่ใช่สินค้าพิเศษ
'                        NETVAL = stcrd_IVinArea.NETVAL
'               Else:
'                        NETVAL = (stcrd_IVinArea.NETVAL + Val(tempMinusStkcod.MINUS_AMOUNT))
'               End If
'               temnetval = temnetval + NETVAL

               '--------------- เหมือนข้างบนเด๊ะ stcrd_IVinArea<>Stcrd
                 '**X1
         
             If stcrd_IVinArea.SLMCOD = PrevKey1 And (stcrd_IVinArea.AREACOD <> PrevKey3 Or PrevKey5 <> stcrd_IVinArea.CUSCOD) Then  '
                     If PrevKey3 <> 0 And Total1 <> 0 And Total2 <> 0 Then
                                 Set F = New CChartSubTotal
                                 F.SALE_ID = PrevKey1
                                 F.AREA_ID = PrevKey3
                                 F.CUS_ID = PrevKey5
                                 F.TOTAL1_AMOUNT = Total1              'รวมของลูกค้า 1.ยอดจริง
                                 F.TOTAL2_AMOUNT = Total2              'รวมของลูกค้า 2.ยอดประเมิน
                                 Call m_CusF.Add(F) '& "-" & D.CUS_ID
                                 Set F = Nothing
                     End If
                End If
            
           If stcrd_IVinArea.SLMCOD = PrevKey1 And stcrd_IVinArea.AREACOD <> PrevKey3 Then
                     If PrevKey3 <> 0 And Total1 <> 0 And Total2 <> 0 Then
                                 Set D = New CChartTotal
                                 D.SALE_ID = PrevKey1
                                 D.AREA_ID = PrevKey3
                                 D.TOTAL1_SUM = Total3                      ' รวมของเซลล์ 1ยอดจริง
                                 D.TOTAL2_SUM = Total4                        ' รวมของเซลล์ 1ยอดประเมิน
                                 Set D.m_Cus = New Collection
                                 Set D.m_Cus = m_CusF
                                 Call collTotal1.Add(D, Trim(D.SALE_ID & "-" & D.AREA_ID)) '& "-" & D.CUS_ID
                                 Set D = Nothing
                                 Set m_CusF = Nothing
                                 Set m_CusF = New Collection
                     End If
                      Total1 = 0
                     Total2 = 0
                     Total3 = 0
                      Total4 = 0
               End If
               
               If stcrd_IVinArea.SLMCOD = PrevKey1 And stcrd_IVinArea.AREACOD <> PrevKey3 And totMinus <> 0 Then           ' !
                     Set C = New CChartTotal                                                     ' ตรงนี้ต้องเช็ค   ArS.SLMCOD=PrevKey1 and temp_Area.MASTER_AREA_ID= PrevKey3
                     C.SALE_ID = PrevKey1  ' ArS.SLMCOD
                     C.AREA_ID = PrevKey3  'temp_Area.MASTER_AREA_ID           ' แก้ Prev3 แระ
                     C.MINUS = Val(totMinus)
                      Call coll_Minus.Add(C, Trim(C.SALE_ID & "-" & C.AREA_ID))
                      Set C = Nothing
                      totMinus = 0
'   If Stcrd.SLMCOD = "15" And temp_Area.MASTER_AREA_ID = 13 Then
'      strTest = strTest & "  , " & Stcrd.DOCNUM
'   End If
                End If
Set tempMinusStkcod = GetMinusCommiss(m_MinusStkcod, stcrd_IVinArea.DOCDAT & "-" & stcrd_IVinArea.DOCNUM & "-" & stcrd_IVinArea.STKCOD, False)
               ' ''debug.print tempCusArea.MASTER_AREA_ID & " __ " & PrevKey3
               If (tempMinusStkcod Is Nothing) Then
                        NETVAL = stcrd_IVinArea.NETVAL                     ' สินค้านี้ไม่มีส่วนลด
               Else:
                        NETVAL = (stcrd_IVinArea.NETVAL + Val(tempMinusStkcod.MINUS_AMOUNT))
If PrevKey1 = "15" And PrevKey3 = 13 Then
strTest = strTest & "  , " & tempMinusStkcod.MINUS_AMOUNT
End If
                        If stcrd_IVinArea.SLMCOD = PrevKey1 Then 'And (stcrd_IVinArea.AREACOD = PrevKey3 Or PrevKey3 = 0) Then         ' !
                                 totMinus = totMinus + tempMinusStkcod.MINUS_AMOUNT

''                          Else
''                                 Set C = New CChartTotal                                                     ' ตรงนี้ต้องเช็ค   ArS.SLMCOD=PrevKey1 and temp_Area.MASTER_AREA_ID= PrevKey3
''                                 C.SALE_ID = PrevKey1  ' ArS.SLMCOD
''                                 C.AREA_ID = temp_Area.MASTER_AREA_ID  '    PrevKey3       ' แก้ Prev3 แระ
''                                 C.MINUS = Val(totMinus)
''                                  Call coll_Minus.Add(C, Trim(C.SALE_ID & "-" & C.AREA_ID))
''                                  Set C = Nothing
''                                  totMinus = 0
                          End If
            End If
            
            Total1 = Total1 + NETVAL
               Total3 = Total3 + NETVAL
            
               Set m_runConditionCommiss = New CConditionCommission

               Set TempCConditionCommiss = GetCheckCommiss(m_ConditionCommiss4, stcrd_IVinArea.STKCOD, False)
               If (TempCConditionCommiss Is Nothing) Then
                             PercentNum1 = (100 / 100)
               Else
                             PercentNum1 = (TempCConditionCommiss.SLM_PERCENT / 100)
               End If
               
               ' ยอดประเมิน วนลูป >= ว่าเข้าเงื่อนไขไหน
               Total2 = Total2 + (NETVAL * PercentNum1)
               Total4 = Total4 + (NETVAL * PercentNum1)
              
       ' เป็น sale และ ลูกค้าที่มีในเงื่อนไขพิเศษ
                Set comMasPro = GetObject("CCommissMasterPromote", m_ComMasPro, Trim(ArS.SLMCOD & "-" & stcrd_IVinArea.CUSCOD))
                If Not (comMasPro Is Nothing) Then                                       ' เป็นลูกค้า AND ที่ใช้เงื่อนไขพิเศษ
                    ' คิดวันที่จ่ายตังค์
                    Set REsumIV = GetARRcpItem(m_REsumIV, stcrd_IVinArea.DOCNUM, False)
                    If Not (REsumIV Is Nothing) Then                       ' มี RE ที่จ่าย IV นี้
                              ' ''debug.print REsumIV.RCPNUM
                              Set tempREdoc = GetREDocDat(m_ReDocdat, stcrd_IVinArea.DOCNUM, False)
                              If Not (tempREdoc Is Nothing) Then
                                  CMPLDAT = tempREdoc.DOCDAT      ' ดึงวันที่ RE ออกมา
                              Else
                                 CMPLDAT = -1
                              End If
                              
                              If ((REsumIV.RCVAMT - NETVAL) >= 0) And CMPLDAT > 0 Then   'จ่ายครบยอด และ มีวันที่จ่าย
                                    DueCount = DateDiff("D", Stcrd.DOCDAT, CMPLDAT)  'ต้องคิดตรงนี้เพื่อไปคำนวณค่า
                              End If
                              
                     End If       '
               End If    ' เป็นลูกค้าที่ใช้เงื่อนไขพิเศษ
               
                 PrevKey3 = stcrd_IVinArea.AREACOD
                 Set temp_Area2 = GetAreaCom(m_AreaCod, stcrd_IVinArea.AREACOD)
                 PrevKey4 = temp_Area2.MASTER_AREA_NAME '!
                 PrevKey5 = Stcrd.CUSCOD
               '-----------------------------------------
               
         End If  ' ต่อท้าย เซลล์คนเดียว เขตเดียวกัน
         Next stcrd_IVinArea
         
 
    
         If PrevKey3 <> 0 And Total1 <> 0 And Total2 <> 0 Then                   ' ???????????????????
            Set F = New CChartSubTotal
            F.SALE_ID = PrevKey1
            F.AREA_ID = PrevKey3
            F.CUS_ID = PrevKey5            ' ??????????
            F.TOTAL1_AMOUNT = Total1
            F.TOTAL2_AMOUNT = Total2
            Call m_CusF.Add(F) '& "-" & D.CUS_ID
            Set F = Nothing
            
          Set D = New CChartTotal
          D.SALE_ID = PrevKey1
          D.AREA_ID = PrevKey3
          D.TOTAL1_SUM = Total3
          D.TOTAL2_SUM = Total4
          Set D.m_Cus = New Collection
          Set D.m_Cus = m_CusF
          Call collTotal1.Add(D, Trim(D.SALE_ID & "-" & D.AREA_ID)) '& "-" & D.CUS_ID
          Set D = Nothing
          Set m_CusF = Nothing
          Set m_CusF = New Collection
         
         Set C = New CChartTotal                                                     ' ตรงนี้ต้องเช็ค   ArS.SLMCOD=PrevKey1 and temp_Area.MASTER_AREA_ID= PrevKey3
         C.SALE_ID = PrevKey1  ' ArS.SLMCOD
         C.AREA_ID = PrevKey3  'temp_Area.MASTER_AREA_ID
         C.MINUS = Val(totMinus)
         Call coll_Minus.Add(C, Trim(C.SALE_ID & "-" & C.AREA_ID))
         Set C = Nothing
         totMinus = 0

'   If PrevKey1 = "15" And PrevKey3 = 13 Then
'      strTest = strTest & "  , " & Stcrd.DOCNUM
'   End If
         
   End If
   
         Total1 = 0
         Total2 = 0
         Total3 = 0
         Total4 = 0
         
           PrevKey3 = 0
           PrevKey4 = ""
           PrevKey5 = ""
           
   TempRs.MoveNext                                                            ' ???????
Wend


   Set ArS = Nothing
   Set Stcrd = Nothing
   Set TempCConditionCommiss = Nothing
   Set tempMinusStkcod = Nothing
   Set m_runConditionCommiss = Nothing
   Set temp_Area = Nothing
    Set temp_Area2 = Nothing
   Set m_ConditionCommiss4 = Nothing
   Set m_IVincomplete0 = Nothing
   Set m_IVincomplete12 = Nothing
   Set m_AreaCod = Nothing
   Set tempCusArea = Nothing
   Set D = Nothing
End Sub

Public Sub LoadcollTotalV(Optional FROM_DOC_DATE As Date = -1, Optional TO_DOC_DATE As Date = -1, Optional collTotal1 As Collection, Optional coll_Minus As Collection, Optional Allflag As Boolean = True)
Dim CMPLDAT As Date
Dim PrevKey1 As String
Dim PrevKey2 As String
Dim PrevKey3 As Long
Dim PrevKey4 As String
Dim PrevKey5 As String
Dim PrevKey6 As Long
Dim toCMPLdate As Date
Dim NETVAL As Double
Dim Rs As ADODB.Recordset
Dim TempRs As ADODB.Recordset
Dim Total1 As Double
Dim Total2 As Double
Dim Total3 As Double
Dim Total4 As Double
Dim L As Long
Dim iCount As Long
Dim IsOK As Boolean
Dim totMinus As Double

Dim DueCount As Long
Dim strTest As String

Dim ArS As COESLM
Dim Stcrd As CStcrd
Dim TempCConditionCommiss As CConditionCommission
Dim tempMinusStkcod As CComMinusStk
Dim m_runConditionCommiss As CConditionCommission
Dim tempCusArea As CCommissionCustomerArea
Dim temp_Area As CCommissionCustomerArea
Dim temp_Area2 As CCommissionCustomerArea
Dim D As CChartTotal
Dim C As CChartTotal
Dim m_MinusStkcod As Collection
Dim m_ConditionCommiss4 As Collection
'Dim m_cusFromArea(20) As Collection
Dim m_IVincomplete0 As Collection
Dim m_IVincomplete12 As Collection
Dim PercentNum1 As Double
Dim m_AreaCod As Collection
Dim m_ComDonStkcod As Collection
Dim tempComDonStkcod As CComDonStk
Dim CorrectStkcod As Boolean
Dim CorrectGroup As Boolean

Dim m_CusE As Collection
Dim m_CusF As Collection
Dim E As CChartSubTotal
Dim F As CChartSubTotal
Dim YEAR_ID As Long

Dim m_ComMasPro As Collection
Dim m_ComPro1 As Collection
Dim comMasPro As CCommissMasterPromote
Dim comPro1 As CComMasSubPromote

Dim REsumIV As CARRcIt
Dim m_REsumIV As Collection
Dim m_ReDocdat As Collection
Dim tempREdoc As CARTrn

Dim m_IVcenter As Collection
Dim m_IVinArea As Collection
Dim tempIVcenter As CComIVcenter
Dim IVoldinArea As Boolean
Dim stcrd_IVinArea As CStcrd

Dim temp_GoodsGroup As CGoodsGroup
Dim temp_GoodsInGroup As CGoodsDetail
Dim m_GoodsGroup As Collection
Dim n_array As Long
Dim m_GoodsInGroup(10) As Collection
Dim num As Long
Dim GOODS_GROUP_ID As Long

   Set m_IVcenter = New Collection
   Set m_IVinArea = New Collection

   Set Rs = New ADODB.Recordset
   Set TempRs = New ADODB.Recordset
   Set ArS = New COESLM
   Set Stcrd = New CStcrd
   Set TempCConditionCommiss = New CConditionCommission
   Set tempMinusStkcod = New CComMinusStk
   Set m_runConditionCommiss = New CConditionCommission
   Set temp_Area = New CCommissionCustomerArea
   Set temp_Area2 = New CCommissionCustomerArea
   Set tempCusArea = New CCommissionCustomerArea
   Set D = New CChartTotal
   Set m_MinusStkcod = New Collection
   Set m_ConditionCommiss4 = New Collection

   Set m_IVincomplete0 = New Collection
   Set m_IVincomplete12 = New Collection
   Set m_AreaCod = New Collection
   Set collTotal1 = New Collection
   Set coll_Minus = New Collection
   Set m_ComDonStkcod = New Collection
   Set tempComDonStkcod = New CComDonStk
   
   Set m_CusE = New Collection
   Set m_CusF = New Collection
   Set E = New CChartSubTotal
   Set F = New CChartSubTotal
   
   Set m_ComMasPro = New Collection
   Set m_ComPro1 = New Collection
   Set comMasPro = New CCommissMasterPromote
   Set comPro1 = New CComMasSubPromote
   
   Set REsumIV = New CARRcIt
   Set m_REsumIV = New Collection
   Set m_ReDocdat = New Collection
   Set tempREdoc = New CARTrn
   
   Set temp_GoodsGroup = New CGoodsGroup
   Set m_GoodsGroup = New Collection
   For n_array = 0 To UBound(m_GoodsInGroup)
      Set m_GoodsInGroup(n_array) = New Collection
   Next n_array
   Set temp_GoodsInGroup = New CGoodsDetail
   
   Call LoadComNamePro(m_ComMasPro, FROM_DOC_DATE, TO_DOC_DATE)       'รายชื่อ เซลล์และลูกค้าคนไหนทีเข้าข่าย
   Call LoadComPro("01", m_ComPro1, FROM_DOC_DATE, TO_DOC_DATE)      'เก็บเงินธรรมดา
   
  Call LoadIVcenter(m_IVcenter, FROM_DOC_DATE, TO_DOC_DATE)
  Call LoadComDonStk(m_ComDonStkcod, FROM_DOC_DATE, TO_DOC_DATE)
  Call LoadCommission04(m_ConditionCommiss4, FROM_DOC_DATE, TO_DOC_DATE, 2)  ' ต้องมาแก้ defail
  Call LoadMinusIV(m_MinusStkcod, FROM_DOC_DATE, TO_DOC_DATE)

  Call LoadREsumIV(Nothing, m_REsumIV)
  Call LoadREDocDat(m_ReDocdat) ' ต้องเอามาหมด

  Call LoadYearId(YEAR_ID, TO_DOC_DATE, TO_DOC_DATE)   'FROM_DOC_DATE,
  Call LoadAreaComReport(Nothing, m_AreaCod, YEAR_ID)
      
   n_array = 0
   Call LoadColGoodsGroup(m_GoodsGroup)                ' มีประเภทไรบ้าง    ' ดึง ชื่อสินค้าที่อยู่ใน master นั่นน และกรุ๊ปนั่นนนน
  For Each temp_GoodsGroup In m_GoodsGroup
      Set TempCConditionCommiss = m_ConditionCommiss4.ITEM(1)   ' ในกลุ่มนั้นมีสินค้าอะไรบ้าง
      Call LoadGoodsDetailFromGM_GG(m_GoodsInGroup(n_array), TempCConditionCommiss.GOODS_MASTER_ID, temp_GoodsGroup.GOODS_GROUP_ID)                 ' มีประเภทไรบ้าง
      n_array = n_array + 1
Next temp_GoodsGroup
   
   Set ArS = New COESLM
   Call glbDaily.QuerySale(ArS, TempRs, iCount, IsOK, glbErrorLog)

   While Not TempRs.EOF          ' sale

         Call ArS.PopulateFromRS(1, TempRs)

         PrevKey1 = ArS.SLMCOD
         PrevKey2 = ArS.SLMNAM
         
     L = 0              ' ???????????????? ????????????
     For Each temp_Area In m_AreaCod
     L = L + 1
     
 ' For num = 1 To n_array
     num = 1
     
      Set Stcrd = New CStcrd
      Stcrd.FROM_DOC_DATE = FROM_DOC_DATE
      Stcrd.TO_DOC_DATE = TO_DOC_DATE
      Stcrd.SLMCOD = ArS.SLMCOD
      Call Stcrd.QueryData(14, Rs, iCount)

      While Not Rs.EOF                     ' ?????????????????

               CorrectStkcod = False
               Call Stcrd.PopulateFromRS(14, Rs)
            
      Set tempComDonStkcod = GetComDonStk(m_ComDonStkcod, Trim(Stcrd.STKCOD), False)
      If (tempComDonStkcod Is Nothing) Then
         CorrectStkcod = True                               ' สินค้านี้ คิดค่าคอมได้
      End If

      Set tempCusArea = GetCusAreaCom(temp_Area.ImportExportItems, Stcrd.CUSCOD, False)
      If (Not (tempCusArea Is Nothing)) And CorrectStkcod Then                     ' ลูกค้านี้อยู่ในเขต
   
   If ArS.SLMCOD = "15" And tempCusArea.MASTER_AREA_ID = 13 And Stcrd.DOCNUM = "IV0045140" Then
   ''debug.print
End If
   
      ' คำนวณดู IV ก่อน
      Set tempIVcenter = GetIVcenter(m_IVcenter, Stcrd.DOCNUM, False)        '
      If (tempIVcenter Is Nothing) Then   'ไม่เจอในคอเล็คชั่น = ไม่ใช่สินค้าพิเศษ
               IVoldinArea = True
      Else:
               IVoldinArea = False
               ' ดังนั้นต้องเก็บ ไว้ในคอล
               Set stcrd_IVinArea = New CStcrd
               stcrd_IVinArea.AREACOD = Trim(Str(tempIVcenter.MASTER_AREA_ID))
               stcrd_IVinArea.AREANAM = tempIVcenter.MASTER_AREA_NAME
               stcrd_IVinArea.NETVAL = Stcrd.NETVAL
               stcrd_IVinArea.DOCNUM = Stcrd.DOCNUM
               stcrd_IVinArea.CUSNAM = Stcrd.CUSNAM
               stcrd_IVinArea.STKDES = Stcrd.STKDES
               stcrd_IVinArea.DOCDAT = Stcrd.DOCDAT
               stcrd_IVinArea.STKCOD = Stcrd.STKCOD
               stcrd_IVinArea.SLMCOD = Stcrd.SLMCOD
               stcrd_IVinArea.CUSCOD = Stcrd.CUSCOD
               Call m_IVinArea.Add(stcrd_IVinArea)
               Set stcrd_IVinArea = Nothing
      End If
 
 If IVoldinArea = True Then  ' ******
 
         CorrectGroup = False

         If Allflag = False Then
              Set tempComDonStkcod = GetComDonStk(m_ComDonStkcod, Trim(Stcrd.STKCOD), False)
              If (tempComDonStkcod Is Nothing) Then
                    Set temp_GoodsInGroup = GetObject("CGoodsGroup", m_GoodsInGroup(num), Trim(Stcrd.STKCOD))
                    If Not (temp_GoodsInGroup Is Nothing) Then
                       GOODS_GROUP_ID = 1
                       CorrectGroup = True
'                     Else
'                        If num = 2 Then
'                              GOODS_GROUP_ID = 2
'                              CorrectGroup = True
'                         End If
                      End If
              End If
       Else
            GOODS_GROUP_ID = 2
            CorrectGroup = True      ' all = true
       End If

           
   If CorrectGroup = True Then
           
               If tempCusArea.MASTER_AREA_ID <> PrevKey3 Or PrevKey5 <> Stcrd.CUSCOD Then   '
                     If PrevKey3 <> 0 And Total1 <> 0 And Total2 <> 0 Then
                                 Set F = New CChartSubTotal
                                 F.SALE_ID = PrevKey1
                                 F.AREA_ID = PrevKey3
                                 F.CUS_ID = PrevKey5
                                 F.GOODS_GROUP_ID = PrevKey6
                                 F.TOTAL1_AMOUNT = Total1              'รวมของลูกค้า 1.ยอดจริง
                                 F.TOTAL2_AMOUNT = Total2              'รวมของลูกค้า 2.ยอดประเมิน
                                 Call m_CusF.Add(F) '& "-" & D.CUS_ID
                                 Set F = Nothing
                     End If
                End If
            
           If tempCusArea.MASTER_AREA_ID <> PrevKey3 Then
                     If PrevKey3 <> 0 And Total1 <> 0 And Total2 <> 0 Then
                                 Set D = New CChartTotal
                                 D.SALE_ID = PrevKey1
                                 D.AREA_ID = PrevKey3
                                 D.GOODS_GROUP_ID = PrevKey6
                                 D.TOTAL1_SUM = Total3                      ' รวมของเซลล์ 1ยอดจริง
                                 D.TOTAL2_SUM = Total4                        ' รวมของเซลล์ 1ยอดประเมิน
                                 Set D.m_Cus = New Collection
                                 Set D.m_Cus = m_CusF
                                 Call collTotal1.Add(D, Trim(D.SALE_ID & "-" & D.AREA_ID & "-" & D.GOODS_GROUP_ID))   '& "-" & D.CUS_ID
                                 Set D = Nothing
                                 Set m_CusF = Nothing
                                 Set m_CusF = New Collection
                     End If
                      Total1 = 0
                     Total2 = 0
                     Total3 = 0
                      Total4 = 0
               End If
               
               If ArS.SLMCOD = PrevKey1 And tempCusArea.MASTER_AREA_ID <> PrevKey3 And totMinus <> 0 Then          ' !
                     Set C = New CChartTotal                                                     ' ตรงนี้ต้องเช็ค   ArS.SLMCOD=PrevKey1 and temp_Area.MASTER_AREA_ID= PrevKey3
                     C.SALE_ID = PrevKey1  ' ArS.SLMCOD
                     C.AREA_ID = PrevKey3  'temp_Area.MASTER_AREA_ID           ' แก้ Prev3 แระ
                     C.MINUS = Val(totMinus)
                      C.GOODS_GROUP_ID = PrevKey6
                      Call coll_Minus.Add(C, Trim(C.SALE_ID & "-" & C.AREA_ID & "-" & C.GOODS_GROUP_ID))
                      Set C = Nothing
                      totMinus = 0
'
'If PrevKey1 = "15" And PrevKey3 = 13 Then
'strTest = strTest & "  , " & Stcrd.DOCNUM
'End If
'
                End If

               Set tempMinusStkcod = GetMinusCommiss(m_MinusStkcod, Stcrd.DOCDAT & "-" & Stcrd.DOCNUM & "-" & Stcrd.STKCOD, False)
               ' ''debug.print tempCusArea.MASTER_AREA_ID & " __ " & PrevKey3
               If (tempMinusStkcod Is Nothing) Then
                        NETVAL = Stcrd.NETVAL                     ' สินค้านี้ไม่มีส่วนลด
               Else:
                        NETVAL = (Stcrd.NETVAL + Val(tempMinusStkcod.MINUS_AMOUNT))
                        
If PrevKey1 = "15" And PrevKey3 = 13 Then
strTest = strTest & "  , " & tempMinusStkcod.MINUS_AMOUNT
End If
                        If ArS.SLMCOD = PrevKey1 Then ' And (tempCusArea.MASTER_AREA_ID = PrevKey3 Or PrevKey3 = 0) Then        ' !
                                 totMinus = totMinus + tempMinusStkcod.MINUS_AMOUNT
''                          Else
''                                 Set C = New CChartTotal                                                     ' ตรงนี้ต้องเช็ค   ArS.SLMCOD=PrevKey1 and temp_Area.MASTER_AREA_ID= PrevKey3
''                                 C.SALE_ID = PrevKey1  ' ArS.SLMCOD
''                                 C.AREA_ID = temp_Area.MASTER_AREA_ID '  PrevKey3          ' แก้ Prev3 แระ
''                                 C.MINUS = Val(totMinus)
''                                  Call coll_Minus.Add(C, Trim(C.SALE_ID & "-" & C.AREA_ID))
''                                  Set C = Nothing
''                                  totMinus = 0
                          End If
                   End If

               Total1 = Total1 + NETVAL
               Total3 = Total3 + NETVAL
            
               Set m_runConditionCommiss = New CConditionCommission

               Set TempCConditionCommiss = GetCheckCommiss(m_ConditionCommiss4, Stcrd.STKCOD, False)
               If (TempCConditionCommiss Is Nothing) Then
                             PercentNum1 = (100 / 100)
               Else
                             PercentNum1 = (TempCConditionCommiss.SLM_PERCENT / 100)
               End If
               
               ' ยอดประเมิน วนลูป >= ว่าเข้าเงื่อนไขไหน
               Total2 = Total2 + (NETVAL * PercentNum1)
               Total4 = Total4 + (NETVAL * PercentNum1)
              
       ' เป็น sale และ ลูกค้าที่มีในเงื่อนไขพิเศษ
                Set comMasPro = GetObject("CCommissMasterPromote", m_ComMasPro, Trim(ArS.SLMCOD & "-" & Stcrd.CUSCOD))
                If Not (comMasPro Is Nothing) Then                                       ' เป็นลูกค้า AND ที่ใช้เงื่อนไขพิเศษ
                    ' คิดวันที่จ่ายตังค์
                    Set REsumIV = GetARRcpItem(m_REsumIV, Stcrd.DOCNUM, False)
                    If Not (REsumIV Is Nothing) Then                       ' มี RE ที่จ่าย IV นี้
                              ' ''debug.print REsumIV.RCPNUM
                              Set tempREdoc = GetREDocDat(m_ReDocdat, Stcrd.DOCNUM, False)
                              If Not (tempREdoc Is Nothing) Then
                                  CMPLDAT = tempREdoc.DOCDAT      ' ดึงวันที่ RE ออกมา
                              Else
                                 CMPLDAT = -1
                              End If
                              
                              If ((REsumIV.RCVAMT - NETVAL) >= 0) And CMPLDAT > 0 Then   'จ่ายครบยอด และ มีวันที่จ่าย
                                    DueCount = DateDiff("D", Stcrd.DOCDAT, CMPLDAT)  'ต้องคิดตรงนี้เพื่อไปคำนวณค่า
                              End If
                              
                     End If       '
               End If    ' เป็นลูกค้าที่ใช้เงื่อนไขพิเศษ
               
               
                 PrevKey3 = tempCusArea.MASTER_AREA_ID
                 Set temp_Area2 = GetAreaCom(m_AreaCod, tempCusArea.MASTER_AREA_ID)
                 PrevKey4 = temp_Area2.MASTER_AREA_NAME '!
                 PrevKey5 = Stcrd.CUSCOD
                 PrevKey6 = GOODS_GROUP_ID
                 
             End If ' CorrectGroup = True
             
      End If  '****** IVinArea
      End If
                Rs.MoveNext
         Wend
         
'    Next num
Next temp_Area
         
         For Each stcrd_IVinArea In m_IVinArea
         
     For num = 0 To n_array                       ' เอากรุ๊ป 1 มาก่อน ค่อยกรุ๊ป 2
   
      If Allflag = False Then
            Set temp_GoodsInGroup = GetObject("CGoodsGroup", m_GoodsInGroup(num), Trim(stcrd_IVinArea.STKCOD))
            If Not (temp_GoodsInGroup Is Nothing) Then
               GOODS_GROUP_ID = 1
               CorrectGroup = True
'           Else
'               GOODS_GROUP_ID = 2         ' ไมใช่วัคซีนฮาร์ทโค๊ดเป็น 2 ไปเลย
            End If
      Else
               GOODS_GROUP_ID = 2
               CorrectGroup = True
      End If

               
         If CorrectGroup Then
         If Val(stcrd_IVinArea.AREACOD) = PrevKey3 And stcrd_IVinArea.SLMCOD = PrevKey1 Then
         
'                          ' คำนวณส่วนลดก่อน
'               Set tempMinusStkcod = GetMinusCommiss(m_MinusStkcod, stcrd_IVinArea.DOCDAT & "-" & stcrd_IVinArea.DOCNUM & "-" & stcrd_IVinArea.STKCOD, False)
'               If (tempMinusStkcod Is Nothing) Then   'ไม่เจอในคอเล็คชั่น = ไม่ใช่สินค้าพิเศษ
'                        NETVAL = stcrd_IVinArea.NETVAL
'               Else:
'                        NETVAL = (stcrd_IVinArea.NETVAL + Val(tempMinusStkcod.MINUS_AMOUNT))
'               End If
'               temnetval = temnetval + NETVAL

               '--------------- เหมือนข้างบนเด๊ะ stcrd_IVinArea<>Stcrd
                 '**X1
         
             If stcrd_IVinArea.SLMCOD = PrevKey1 And (stcrd_IVinArea.AREACOD <> PrevKey3 Or PrevKey5 <> stcrd_IVinArea.CUSCOD) Then  '
                     If PrevKey3 <> 0 And Total1 <> 0 And Total2 <> 0 Then
                                 Set F = New CChartSubTotal
                                 F.SALE_ID = PrevKey1
                                 F.AREA_ID = PrevKey3
                                 F.CUS_ID = PrevKey5
                                 F.GOODS_GROUP_ID = PrevKey6
                                 F.TOTAL1_AMOUNT = Total1              'รวมของลูกค้า 1.ยอดจริง
                                 F.TOTAL2_AMOUNT = Total2              'รวมของลูกค้า 2.ยอดประเมิน
                                 Call m_CusF.Add(F) '& "-" & D.CUS_ID
                                 Set F = Nothing
                     End If
                End If
            
           If stcrd_IVinArea.SLMCOD = PrevKey1 And stcrd_IVinArea.AREACOD <> PrevKey3 Then
                     If PrevKey3 <> 0 And Total1 <> 0 And Total2 <> 0 Then
                                 Set D = New CChartTotal
                                 D.SALE_ID = PrevKey1
                                 D.AREA_ID = PrevKey3
                                 D.GOODS_GROUP_ID = PrevKey6
                                 D.TOTAL1_SUM = Total3                      ' รวมของเซลล์ 1ยอดจริง
                                 D.TOTAL2_SUM = Total4                        ' รวมของเซลล์ 1ยอดประเมิน
                                 Set D.m_Cus = New Collection
                                 Set D.m_Cus = m_CusF
                                 Call collTotal1.Add(D, Trim(D.SALE_ID & "-" & D.AREA_ID & "-" & D.GOODS_GROUP_ID))  '& "-" & D.CUS_ID
                                 Set D = Nothing
                                 Set m_CusF = Nothing
                                 Set m_CusF = New Collection
                     End If
                      Total1 = 0
                     Total2 = 0
                     Total3 = 0
                      Total4 = 0
               End If
               
               If stcrd_IVinArea.SLMCOD = PrevKey1 And stcrd_IVinArea.AREACOD <> PrevKey3 And totMinus <> 0 Then           ' !
                     Set C = New CChartTotal                                                     ' ตรงนี้ต้องเช็ค   ArS.SLMCOD=PrevKey1 and temp_Area.MASTER_AREA_ID= PrevKey3
                     C.SALE_ID = PrevKey1  ' ArS.SLMCOD
                     C.AREA_ID = PrevKey3  'temp_Area.MASTER_AREA_ID           ' แก้ Prev3 แระ
                     C.GOODS_GROUP_ID = PrevKey6
                     C.MINUS = Val(totMinus)
                      Call coll_Minus.Add(C, Trim(C.SALE_ID & "-" & C.AREA_ID & "-" & C.GOODS_GROUP_ID))
                      Set C = Nothing
                      totMinus = 0
'   If Stcrd.SLMCOD = "15" And temp_Area.MASTER_AREA_ID = 13 Then
'      strTest = strTest & "  , " & Stcrd.DOCNUM
'   End If
                End If
Set tempMinusStkcod = GetMinusCommiss(m_MinusStkcod, stcrd_IVinArea.DOCDAT & "-" & stcrd_IVinArea.DOCNUM & "-" & stcrd_IVinArea.STKCOD, False)
               ' ''debug.print tempCusArea.MASTER_AREA_ID & " __ " & PrevKey3
               If (tempMinusStkcod Is Nothing) Then
                        NETVAL = stcrd_IVinArea.NETVAL                     ' สินค้านี้ไม่มีส่วนลด
               Else:
                        NETVAL = (stcrd_IVinArea.NETVAL + Val(tempMinusStkcod.MINUS_AMOUNT))
If PrevKey1 = "15" And PrevKey3 = 13 Then
strTest = strTest & "  , " & tempMinusStkcod.MINUS_AMOUNT
End If
                        If stcrd_IVinArea.SLMCOD = PrevKey1 Then 'And (stcrd_IVinArea.AREACOD = PrevKey3 Or PrevKey3 = 0) Then         ' !
                                 totMinus = totMinus + tempMinusStkcod.MINUS_AMOUNT

''                          Else
''                                 Set C = New CChartTotal                                                     ' ตรงนี้ต้องเช็ค   ArS.SLMCOD=PrevKey1 and temp_Area.MASTER_AREA_ID= PrevKey3
''                                 C.SALE_ID = PrevKey1  ' ArS.SLMCOD
''                                 C.AREA_ID = temp_Area.MASTER_AREA_ID  '    PrevKey3       ' แก้ Prev3 แระ
''                                 C.MINUS = Val(totMinus)
''                                  Call coll_Minus.Add(C, Trim(C.SALE_ID & "-" & C.AREA_ID))
''                                  Set C = Nothing
''                                  totMinus = 0
                          End If
            End If
            
               Total1 = Total1 + NETVAL
               Total3 = Total3 + NETVAL
            
               Set m_runConditionCommiss = New CConditionCommission

               Set TempCConditionCommiss = GetCheckCommiss(m_ConditionCommiss4, stcrd_IVinArea.STKCOD, False)
               If (TempCConditionCommiss Is Nothing) Then
                             PercentNum1 = (100 / 100)
               Else
                             PercentNum1 = (TempCConditionCommiss.SLM_PERCENT / 100)
               End If
               
               ' ยอดประเมิน วนลูป >= ว่าเข้าเงื่อนไขไหน
               Total2 = Total2 + (NETVAL * PercentNum1)
               Total4 = Total4 + (NETVAL * PercentNum1)
              
       ' เป็น sale และ ลูกค้าที่มีในเงื่อนไขพิเศษ
                Set comMasPro = GetObject("CCommissMasterPromote", m_ComMasPro, Trim(ArS.SLMCOD & "-" & stcrd_IVinArea.CUSCOD))
                If Not (comMasPro Is Nothing) Then                                       ' เป็นลูกค้า AND ที่ใช้เงื่อนไขพิเศษ
                    ' คิดวันที่จ่ายตังค์
                    Set REsumIV = GetARRcpItem(m_REsumIV, stcrd_IVinArea.DOCNUM, False)
                    If Not (REsumIV Is Nothing) Then                       ' มี RE ที่จ่าย IV นี้
                              ' ''debug.print REsumIV.RCPNUM
                              Set tempREdoc = GetREDocDat(m_ReDocdat, stcrd_IVinArea.DOCNUM, False)
                              If Not (tempREdoc Is Nothing) Then
                                  CMPLDAT = tempREdoc.DOCDAT      ' ดึงวันที่ RE ออกมา
                              Else
                                 CMPLDAT = -1
                              End If
                              
                              If ((REsumIV.RCVAMT - NETVAL) >= 0) And CMPLDAT > 0 Then   'จ่ายครบยอด และ มีวันที่จ่าย
                                    DueCount = DateDiff("D", Stcrd.DOCDAT, CMPLDAT)  'ต้องคิดตรงนี้เพื่อไปคำนวณค่า
                              End If
                              
                     End If       '
               End If    ' เป็นลูกค้าที่ใช้เงื่อนไขพิเศษ
               
                 PrevKey3 = stcrd_IVinArea.AREACOD
                 Set temp_Area2 = GetAreaCom(m_AreaCod, stcrd_IVinArea.AREACOD)
                 PrevKey4 = temp_Area2.MASTER_AREA_NAME '!
                 PrevKey5 = Stcrd.CUSCOD
                 PrevKey6 = GOODS_GROUP_ID
               '-----------------------------------------
               
         End If  ' ต่อท้าย เซลล์คนเดียว เขตเดียวกัน
         End If    '  CorrectStkcod
          
         Next num
         Next stcrd_IVinArea
         
'         Set temp_GoodsInGroup = GetObject("CGoodsGroup", m_GoodsInGroup(num), Trim(stcrd_IVinArea.STKCOD))
'            If Not (temp_GoodsInGroup Is Nothing) Then
'               PrevKey6 = temp_GoodsInGroup.GOODS_GROUP_ID
'           Else
'               PrevKey6 = 2         ' ไมใช่วัคซีนฮาร์ทโค๊ดเป็น 2 ไปเลย
'            End If

         If PrevKey3 <> 0 And Total1 <> 0 And Total2 <> 0 Then                   ' ???????????????????
            Set F = New CChartSubTotal
            F.SALE_ID = PrevKey1
            F.AREA_ID = PrevKey3
            F.CUS_ID = PrevKey5            ' ??????????
            F.GOODS_GROUP_ID = PrevKey6
            F.TOTAL1_AMOUNT = Total1
            F.TOTAL2_AMOUNT = Total2
            Call m_CusF.Add(F) '& "-" & D.CUS_ID
            Set F = Nothing
            
          Set D = New CChartTotal
          D.SALE_ID = PrevKey1
          D.AREA_ID = PrevKey3
          D.GOODS_GROUP_ID = PrevKey6
          D.TOTAL1_SUM = Total3
          D.TOTAL2_SUM = Total4
          Set D.m_Cus = New Collection
          Set D.m_Cus = m_CusF
          Call collTotal1.Add(D, Trim(D.SALE_ID & "-" & D.AREA_ID & "-" & D.GOODS_GROUP_ID))    '& "-" & D.CUS_ID
          Set D = Nothing
          Set m_CusF = Nothing
          Set m_CusF = New Collection
         
         Set C = New CChartTotal                                                     ' ตรงนี้ต้องเช็ค   ArS.SLMCOD=PrevKey1 and temp_Area.MASTER_AREA_ID= PrevKey3
         C.SALE_ID = PrevKey1  ' ArS.SLMCOD
         C.AREA_ID = PrevKey3  'temp_Area.MASTER_AREA_ID
         C.GOODS_GROUP_ID = PrevKey6
         C.MINUS = Val(totMinus)
         Call coll_Minus.Add(C, Trim(C.SALE_ID & "-" & C.AREA_ID & "-" & C.GOODS_GROUP_ID))
         Set C = Nothing
         totMinus = 0

'   If PrevKey1 = "15" And PrevKey3 = 13 Then
'      strTest = strTest & "  , " & Stcrd.DOCNUM
'   End If
   End If
   
         Total1 = 0
         Total2 = 0
         Total3 = 0
         Total4 = 0
         
           PrevKey3 = 0
           PrevKey4 = ""
           PrevKey5 = ""
           PrevKey6 = 0
           
   TempRs.MoveNext                                                            ' ???????
Wend


   Set ArS = Nothing
   Set Stcrd = Nothing
   Set TempCConditionCommiss = Nothing
   Set tempMinusStkcod = Nothing
   Set m_runConditionCommiss = Nothing
   Set temp_Area = Nothing
    Set temp_Area2 = Nothing
   Set m_ConditionCommiss4 = Nothing
   Set m_IVincomplete0 = Nothing
   Set m_IVincomplete12 = Nothing
   Set m_AreaCod = Nothing
   Set tempCusArea = Nothing
   Set D = Nothing
   Set temp_GoodsInGroup = Nothing
End Sub
Public Sub LoadSaleHead(Optional Cl As Collection = Nothing, Optional FromDocDat As Date = -1, Optional ToDocDat As Date = -1)
On Error GoTo ErrorHandler
Dim D As CCommissionChart
Dim ItemCount As Long
Dim TempData As CCommissionChart
Dim Rs As ADODB.Recordset
Dim i As Long

   Set D = New CCommissionChart
   Set Rs = New ADODB.Recordset
   
   D.VALID_FROM = FromDocDat
   D.VALID_TO = ToDocDat
   D.ORDER_TYPE = 1
   Call D.QueryData(2, Rs, ItemCount)   ' ดึงข้อมูลเพื่อเอาไปเก็บในคอเล็กชั่น

   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   
   While Not Rs.EOF
      i = i + 1
      Set TempData = New CCommissionChart
      Call TempData.PopulateFromRS(2, Rs)

      If Not (Cl Is Nothing) Then
         Call Cl.Add(TempData, Trim(Trim(TempData.SALE_ID & "-" & TempData.GOODS_GROUP_ID)))     ' KEy
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   Set Rs = Nothing
   Set D = Nothing
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub

Public Sub LoadSaleChartHead(Optional Cl As Collection = Nothing, Optional FromDocDat As Date = -1, Optional ToDocDat As Date = -1)
On Error GoTo ErrorHandler
Dim D As CCommissionChart
Dim ItemCount As Long
Dim TempData As CCommissionChart
Dim Rs As ADODB.Recordset
Dim i As Long

   Set D = New CCommissionChart
   Set Rs = New ADODB.Recordset
   
   D.VALID_FROM = FromDocDat
   D.VALID_TO = ToDocDat
   D.ORDER_TYPE = 1
   Call D.QueryData(2, Rs, ItemCount)   ' ดึงข้อมูลเพื่อเอาไปเก็บในคอเล็กชั่น

   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   
   While Not Rs.EOF
      i = i + 1
      Set TempData = New CCommissionChart
      Call TempData.PopulateFromRS(2, Rs)

      If Not (Cl Is Nothing) Then
         Call Cl.Add(TempData, Trim(TempData.SALE_ID & "-" & TempData.MASTER_AREA_ID))        ' KEy
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   Set Rs = Nothing
   Set D = Nothing
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub
Public Sub LoadSaleChartChildNotHead(Optional Cl As Collection = Nothing, Optional FromDocDat As Date = -1, Optional ToDocDat As Date = -1)
On Error GoTo ErrorHandler
Dim D As CCommissionChart
Dim ItemCount As Long
Dim TempData As CCommissionChart
Dim Rs As ADODB.Recordset
Dim i As Long

   Set D = New CCommissionChart
   Set Rs = New ADODB.Recordset
   
   D.VALID_FROM = FromDocDat
   D.VALID_TO = ToDocDat
   D.ORDER_TYPE = 1
   Call D.QueryData(3, Rs, ItemCount)   ' ดึงข้อมูลเพื่อเอาไปเก็บในคอเล็กชั่น

   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   
   While Not Rs.EOF
      i = i + 1
      Set TempData = New CCommissionChart
      Call TempData.PopulateFromRS(3, Rs)

      If Not (Cl Is Nothing) Then
    '     Call Cl.Add(TempData, Trim(TempData.PARENT_ID & "-" & TempData.SALE_ID))         ' KEy\
         Call Cl.Add(TempData, Trim(TempData.SALE_ID & "-" & TempData.MASTER_AREA_ID))           ' KEy
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   Set Rs = Nothing
   Set D = Nothing
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub

Public Sub LoadSaleChartChild(Optional Cl As Collection = Nothing, Optional FromDocDat As Date = -1, Optional ToDocDat As Date = -1)
On Error GoTo ErrorHandler
Dim D As CCommissionChart
Dim ItemCount As Long
Dim TempData As CCommissionChart
Dim Rs As ADODB.Recordset
Dim i As Long

   Set D = New CCommissionChart
   Set Rs = New ADODB.Recordset
   
   D.VALID_FROM = FromDocDat
   D.VALID_TO = ToDocDat
   D.ORDER_TYPE = 1
   Call D.QueryData(3, Rs, ItemCount)   ' ดึงข้อมูลเพื่อเอาไปเก็บในคอเล็กชั่น

   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   
   While Not Rs.EOF
      i = i + 1
      Set TempData = New CCommissionChart
      Call TempData.PopulateFromRS(3, Rs)

      If Not (Cl Is Nothing) Then
    '     Call Cl.Add(TempData, Trim(TempData.PARENT_ID & "-" & TempData.SALE_ID))         ' KEy\
         Call Cl.Add(TempData, Trim(TempData.P_SALE_ID & "-" & TempData.SALE_ID & "-" & TempData.MASTER_AREA_ID) & "-" & Trim(TempData.GOODS_GROUP_ID))          ' KEy
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   Set Rs = Nothing
   Set D = Nothing
   Exit Sub
   
ErrorHandler:
   'debug.print
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub

Public Function GetSaleChart(m_TempCol As Collection, TempKey As String, Optional HaveNew As Boolean = True) As CCommissionChart
On Error Resume Next
Dim Ei As CCommissionChart
Static TempEi As CCommissionChart

   Set Ei = m_TempCol(TempKey)
   '''debug.print TempKey
    If Ei Is Nothing And HaveNew Then
                If TempEi Is Nothing Then
                   Set TempEi = New CCommissionChart
                End If
      Set GetSaleChart = TempEi
   Else
      Set GetSaleChart = Ei
   End If
End Function

Public Function GetTotMinChart(m_TempCol As Collection, TempKey As String, Optional HaveNew As Boolean = True) As CChartTotal
On Error Resume Next
Dim Ei As CChartTotal
Static TempEi As CChartTotal

   Set Ei = m_TempCol(TempKey)
   ''debug.print TempKey
    If Ei Is Nothing And HaveNew Then
                If TempEi Is Nothing Then
                   Set TempEi = New CChartTotal
                End If
      Set GetTotMinChart = TempEi
   Else
      Set GetTotMinChart = Ei
   End If
End Function


Public Function GetAreaCom(m_TempCol As Collection, TempKey As String, Optional HaveNew As Boolean = True) As CCommissionCustomerArea
On Error Resume Next
Dim Ei As CCommissionCustomerArea
Static TempEi As CCommissionCustomerArea

   Set Ei = m_TempCol(TempKey)
   ''debug.print TempKey
    If Ei Is Nothing And HaveNew Then
                If TempEi Is Nothing Then
                   Set TempEi = New CCommissionCustomerArea
                End If
      Set GetAreaCom = TempEi
   Else
      Set GetAreaCom = Ei
   End If
End Function

Public Sub LoadREsumIV(C As ComboBox, Optional Cl As Collection = Nothing, Optional SaleCode As String = "", Optional db2 As Boolean)
On Error GoTo ErrorHandler
Dim D As CARRcIt
Dim ItemCount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CARRcIt
Dim i As Long

   Set D = New CARRcIt
   Set Rs = New ADODB.Recordset
   D.db2 = db2
   Call D.QueryData(9, Rs, ItemCount)

'   If Not (C Is Nothing) Then
'      C.Clear
'      i = 0
'      C.AddItem ("")
'   End If

   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If

   While Not Rs.EOF
      i = i + 1
      Set TempData = New CARRcIt
      Call TempData.PopulateFromRS(9, Rs)

      If Not (C Is Nothing) Then
          C.AddItem (TempData.DOCNUM)
      End If

      If Not (Cl Is Nothing) Then
         If TempData.DOCNUM = "IV0044548" Then
            ''debug.print
         End If
         Call Cl.Add(TempData, TempData.DOCNUM)
      End If

      Set TempData = Nothing
      Rs.MoveNext
   Wend

   Set Rs = Nothing
   Set D = Nothing
   Exit Sub

ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub


Public Sub LoadGroupIncentive(Optional Cl1 As Collection = Nothing, Optional FromCMPLDat As Date = -1, Optional ToCMPLDat As Date = -1, Optional GOODS_GROUP_ID As Long = 1)
On Error GoTo ErrorHandler
Dim D As CConditionCommission
Dim ItemCount As Long
Dim TempData As CConditionCommission
Dim Rs As ADODB.Recordset
Dim i As Long

   Set D = New CConditionCommission
   Set Rs = New ADODB.Recordset
   
   D.FROM_CMPL_DATE = FromCMPLDat
   D.TO_CMPL_DATE = ToCMPLDat
   D.COMTYP = "05"
   D.GOODS_GROUP_ID = GOODS_GROUP_ID
   Call D.QueryData(6, Rs, ItemCount)

   If Not (Cl1 Is Nothing) Then
      Set Cl1 = Nothing
      Set Cl1 = New Collection
   End If
   
'   If Not (Cl2 Is Nothing) Then
'      Set Cl2 = Nothing
'      Set Cl2 = New Collection
'   End If
   
   While Not Rs.EOF
      i = i + 1
      Set TempData = New CConditionCommission
      Call TempData.PopulateFromRS(6, Rs)

      If Not (Cl1 Is Nothing) Then
          Call Cl1.Add(TempData, TempData.STKCOD)   ' ไม่ต้องมี key เพราะเอาไปวนหาอย่างเดียว, TempData.STKCOD
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   Set Rs = Nothing
   Set D = Nothing
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub

Public Sub LoadGroupIncenPro(Optional Cl1 As Collection = Nothing, Optional FromCMPLDat As Date = -1, Optional ToCMPLDat As Date = -1)
On Error GoTo ErrorHandler
Dim D As CIncentivePromote
Dim ItemCount As Long
Dim TempData As CIncentivePromote
Dim Rs As ADODB.Recordset
Dim i As Long

   Set D = New CIncentivePromote
   Set Rs = New ADODB.Recordset
   
   D.FROM_CMPL_DATE = FromCMPLDat
   D.TO_CMPL_DATE = ToCMPLDat
   Call D.QueryData(6, Rs, ItemCount)

   If Not (Cl1 Is Nothing) Then
      Set Cl1 = Nothing
      Set Cl1 = New Collection
   End If
   
'   If Not (Cl2 Is Nothing) Then
'      Set Cl2 = Nothing
'      Set Cl2 = New Collection
'   End If
   
   While Not Rs.EOF
      i = i + 1
      Set TempData = New CIncentivePromote
      Call TempData.PopulateFromRS(6, Rs)

      If Not (Cl1 Is Nothing) Then
          Call Cl1.Add(TempData, TempData.STKCOD)   ' ไม่ต้องมี key เพราะเอาไปวนหาอย่างเดียว, TempData.STKCOD
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   Set Rs = Nothing
   Set D = Nothing
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub

Public Sub LoadCommissPara(C As ComboBox, Optional Cl As Collection = Nothing, Optional Parameter As Long = -1)
On Error GoTo ErrorHandler
Dim D As CCommissMasterPara
Dim ItemCount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CCommissMasterPara
Dim i As Long
   
   Set D = New CCommissMasterPara
   Set Rs = New ADODB.Recordset
   
   D.MASTER_FROMTO_ID = Parameter
   Call D.QueryData(1, Rs, ItemCount)
   
   If Not (C Is Nothing) Then
      C.Clear
      i = 0
      C.AddItem ("")
   End If
   
   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   
   While Not Rs.EOF
      i = i + 1
      Set TempData = New CCommissMasterPara
      Call TempData.PopulateFromRS(1, Rs)

      If Not (C Is Nothing) Then
         C.AddItem (TempData.MASTER_PARAMETER_NAME)
         C.ItemData(i) = TempData.MASTER_PARAMETER_ID
      End If

      If Not (Cl Is Nothing) Then
         Call Cl.Add(TempData, Trim(Str(TempData.MASTER_PARAMETER_ID)))
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   Set Rs = Nothing
   Set D = Nothing
   Exit Sub

ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub

Public Sub LoadGPwithGroup(Optional Cl As Collection = Nothing, Optional FromDocDat As Date = -1, Optional ToDocDat As Date = -1)
On Error GoTo ErrorHandler
Dim D As CMasterFromToDetail
Dim ItemCount As Long
Dim TempData As CMasterFromToDetail
Dim Rs As ADODB.Recordset
Dim i As Long

   Set D = New CMasterFromToDetail
   Set Rs = New ADODB.Recordset
   
   D.VALID_FROM = FromDocDat
   D.VALID_TO = ToDocDat
   D.ORDER_TYPE = 1
   Call D.QueryData(3, Rs, ItemCount)   ' ดึงข้อมูลเพื่อเอาไปเก็บในคอเล็กชั่น

   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   
   While Not Rs.EOF
      i = i + 1
      Set TempData = New CMasterFromToDetail
      Call TempData.PopulateFromRS(3, Rs)

      If Not (Cl Is Nothing) Then
         Call Cl.Add(TempData, Trim(TempData.SLMCOD))            ' KEy  ถ้าใส่เดือนครอบกัน จะ error
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   Set Rs = Nothing
   Set D = Nothing
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub


Public Sub Load3ComcollTotal(Optional FROM_DOC_DATE As Date = -1, Optional TO_DOC_DATE As Date = -1, Optional FROM_CMPL_DATE As Date = -1, Optional TO_CMPL_DATE As Date = -1, Optional collTotal3Com As Collection = Nothing)
Dim CMPLDAT As Date
Dim PrevKey1 As String
Dim PrevKey2 As String
Dim PrevKey3 As Long
Dim PrevKey4 As String
Dim PrevKey5 As String
Dim toCMPLdate As Date
Dim NETVAL As Double
Dim Rs As ADODB.Recordset
Dim TempRs As ADODB.Recordset
Dim Total1 As Double
Dim Total2 As Double
Dim Total3 As Double
Dim Total4 As Double
Dim Total5 As Double
Dim Total6 As Double
Dim L As Long
Dim iCount As Long
Dim IsOK As Boolean
Dim totMinus As Double

Dim ArS As COESLM
Dim Stcrd As CStcrd
Dim TempCConditionCommiss As CConditionCommission
Dim tempMinusStkcod As CComMinusStk
Dim m_runConditionCommiss As CConditionCommission
Dim tempCusArea As CCommissionCustomerArea
Dim temp_Area As CCommissionCustomerArea
Dim temp_Area2 As CCommissionCustomerArea
Dim D As CChartTotal
Dim C As CChartTotal
Dim m_MinusStkcod As Collection
Dim NEWCUSVAL As Double
Dim m_AreaCod As Collection
Dim m_saleChartHead As Collection
Dim m_saleChartChild As Collection
Dim m_SaleName As Collection
Dim m_GPwithGroup As Collection
Dim temp_Head As CCommissionChart
Dim temp_Child2 As CCommissionChart
Dim DueCount As Double
Dim DueCount2 As Double
Dim DueCount3 As Double
Dim m_ConditionCommiss1 As Collection
Dim m_ConditionCommiss2 As Collection
Dim m_ConditionCommiss3 As Collection
Dim m_ConditionCommiss4 As Collection
Dim m_ConditionCommiss5 As Collection
Dim m_ConditionCommiss5_1 As Collection
Dim m_ConditionCommiss5_2 As Collection
Dim m_ReDocdat As Collection
Dim m_REsumIV As Collection
Dim m_StkcodGroup As Collection
Dim GP As Double
Dim GROUP As Double
Dim temp_GPwithGroup As CMasterFromToDetail
Dim PercentNum1 As Double
Dim PercentSum As Double
Dim PercentNum2 As Double
Dim PercentSum2 As Double
Dim PercentNum3 As Double
Dim PercentSum3 As Double
Dim REsumIV As CARRcIt
Dim tempREdoc As CARTrn
Dim SixMonthFirst As Date
Dim SixMonthLast As Date
Dim CMPLFirstDate As Date
Dim CMPLLastDate As Date
Dim dayFirst As Date
Dim dayLast As Date
Dim SumTRNQTY As Double
Dim NUM_ONE As Double
Dim NUM_TWO As Double
Dim EnableIncentive As Boolean
Dim FlagNewCus   As Boolean
Dim CreditKey As Long

Dim m_NewCus As Collection
Dim m_runConditionCommiss2 As CConditionCommission
Dim m_AllNewCus As Collection
   Dim m_SumStcrd    As Collection
Dim SumStkcod  As CStcrd
Dim temp_Stcrd As CStcrd
Dim m_ComDonStkcod  As Collection
Dim tempComDonStkcod As CComDonStk
Dim CorrectStkcod As Boolean
Dim YEAR_ID As Long

Dim asAmount As Double
Dim MoreAmount As Double
Dim AMOUNT As Double
Dim Percent100 As Double
   
Dim m_ComPro2 As Collection
Dim m_ComPro3 As Collection
Dim m_ComPro4 As Collection
Dim m_ComMasPro As Collection
Dim temp_ComMasPro As CCommissMasterPromote
Dim TempComPro As CComMasSubPromote
Dim m_runComPro As CComMasSubPromote
Dim m_runPro As CCommissPromote
Dim m_IncenPro5 As Collection
Dim m_IncenPro5_1 As Collection
Dim m_IncenPro5_2 As Collection
Dim m_IncenProGroup As Collection
Dim m_runIncenPro2 As CIncentivePromote
Dim TempIncenPro As CIncentivePromote
Dim m_runIncenPro As CIncentivePromote
   
Dim m_IVinDateStcrd As Collection
Dim IVinDateStcrd As CStcrd

Dim firstInMonth As Date
Dim lastInMonth As Date

Dim m_IVcenter As Collection
Dim m_IVinArea As Collection
Dim tempIVcenter As CComIVcenter
Dim IVoldinArea As Boolean
Dim stcrd_IVinArea As CStcrd
Dim cr_condiCom As CConditionCommission
Dim NumCR As Long
Dim m_IVcredit As Collection
Dim tempIVcredit As CComIVcredit

Dim stcrd_mixdb As Collection
Dim stcrd_temp As CStcrd

   Set stcrd_mixdb = New Collection
   Set stcrd_temp = New CStcrd

   Set m_IVcredit = New Collection
   Set tempIVcredit = New CComIVcredit
   Set m_IVcenter = New Collection
   Set m_IVinArea = New Collection
   
   Set temp_Stcrd = New CStcrd
   Set SumStkcod = New CStcrd
   Set Rs = New ADODB.Recordset
   Set TempRs = New ADODB.Recordset
   Set REsumIV = New CARRcIt
   Set ArS = New COESLM
   Set Stcrd = New CStcrd
   Set TempCConditionCommiss = New CConditionCommission
   Set tempMinusStkcod = New CComMinusStk
   Set m_runConditionCommiss = New CConditionCommission
   Set temp_Area = New CCommissionCustomerArea
   Set temp_Area2 = New CCommissionCustomerArea
   Set tempCusArea = New CCommissionCustomerArea
   Set D = New CChartTotal
   Set m_MinusStkcod = New Collection
   Set m_ConditionCommiss1 = New Collection
   Set m_ConditionCommiss2 = New Collection
   Set m_ConditionCommiss3 = New Collection
   Set m_ConditionCommiss4 = New Collection
   Set m_ConditionCommiss5 = New Collection
   Set m_ConditionCommiss5_1 = New Collection
   Set m_ConditionCommiss5_2 = New Collection
   Set m_AreaCod = New Collection
   Set m_saleChartHead = New Collection
   Set m_saleChartChild = New Collection
   Set m_SaleName = New Collection
   Set m_GPwithGroup = New Collection
   Set temp_Head = New CCommissionChart
      Set temp_Child2 = New CCommissionChart
      Set m_ReDocdat = New Collection
      Set m_REsumIV = New Collection
      Set m_StkcodGroup = New Collection
     Set temp_GPwithGroup = New CMasterFromToDetail
     Set tempREdoc = New CARTrn
     Set m_NewCus = New Collection
     Set m_runConditionCommiss2 = New CConditionCommission
     Set m_AllNewCus = New Collection
      Set m_SumStcrd = New Collection
         Set m_ComDonStkcod = New Collection
   Set tempComDonStkcod = New CComDonStk
   
   Set m_ComPro2 = New Collection
   Set m_ComPro3 = New Collection
   Set m_ComPro4 = New Collection
   Set m_ComMasPro = New Collection
   Set temp_ComMasPro = New CCommissMasterPromote
   Set TempComPro = New CComMasSubPromote
  Set m_runComPro = New CComMasSubPromote
  Set m_runPro = New CCommissPromote
   
   Set m_IncenPro5 = New Collection
   Set m_IncenPro5_1 = New Collection
   Set m_IncenPro5_2 = New Collection
   Set m_IncenProGroup = New Collection
   Set m_runIncenPro2 = New CIncentivePromote
   Set TempIncenPro = New CIncentivePromote
   Set m_runIncenPro = New CIncentivePromote
   
   Set m_IVinDateStcrd = New Collection
   Set IVinDateStcrd = New CStcrd
   
  Call LoadREsumIV(Nothing, m_REsumIV)
  
Call GetFirstLastDate(FROM_DOC_DATE, firstInMonth, lastInMonth)
Call LoadCommission("01", m_ConditionCommiss1, firstInMonth, lastInMonth, 1)
Call LoadSaleChartHead(m_saleChartHead, firstInMonth, lastInMonth)                                   ' ไม่ใส่วัน มีปัญหา

Call LoadSaleChartChildNotHead(m_saleChartChild, firstInMonth, lastInMonth)
Call LoadSale(Nothing, m_SaleName)
Call LoadGPwithGroup(m_GPwithGroup, firstInMonth, lastInMonth)                        ' ไม่ใส่วัน มีปัญหา
Call LoadIVcenter(m_IVcenter, firstInMonth, lastInMonth)
Call LoadMinusIV(m_MinusStkcod, -1, lastInMonth)
Call LoadComDonStk(m_ComDonStkcod, firstInMonth, lastInMonth)
Call LoadIVinDateStcrd(m_IVinDateStcrd, firstInMonth, lastInMonth)
    
 Call LoadCommission("02", m_ConditionCommiss2, FROM_CMPL_DATE, TO_CMPL_DATE, 1)
 Call LoadCommission("03", m_ConditionCommiss3, FROM_CMPL_DATE, TO_CMPL_DATE, 1)
  Call LoadCommission04(m_ConditionCommiss4, FROM_CMPL_DATE, TO_CMPL_DATE)
  Call LoadIVcredit(m_IVcredit, -1, TO_CMPL_DATE)
  Call LoadREDocDat(m_ReDocdat, FROM_CMPL_DATE, TO_CMPL_DATE)
   Call LoadCommission05(m_ConditionCommiss5, FROM_CMPL_DATE, TO_CMPL_DATE, 2)   ' ไม่ใช่วัคซีน
   
      Set cr_condiCom = m_ConditionCommiss5.ITEM(1)
      NumCR = cr_condiCom.INCEN_CR
      
   Call LoadCommission05_2(m_ConditionCommiss5_2, FROM_CMPL_DATE, TO_CMPL_DATE)
   Call LoadGroupIncentive(m_StkcodGroup, FROM_CMPL_DATE, TO_CMPL_DATE)
   
    Call LoadComNamePro(m_ComMasPro, FROM_CMPL_DATE, TO_CMPL_DATE)       'รายชื่อ เซลล์และลูกค้าคนไหนทีเข้าข่าย
   Call LoadComPro("02", m_ComPro2, FROM_CMPL_DATE, TO_CMPL_DATE)      'เก็บเงินธรรมดา
   Call LoadComPro("03", m_ComPro3, FROM_CMPL_DATE, TO_CMPL_DATE)     'เก็บเงินพิเศษ

      Call LoadIncenPro05(m_IncenPro5, FROM_CMPL_DATE, TO_CMPL_DATE)  ' เงื่อนไข incentive
   Call LoadIncenPro05_2(m_IncenPro5_2, FROM_CMPL_DATE, TO_CMPL_DATE)  ' คัดมาเฉพาะสินค้า Incentive
   Call LoadGroupIncenPro(m_IncenProGroup, FROM_CMPL_DATE, TO_CMPL_DATE)
   
    Call LoadYearId(YEAR_ID, TO_CMPL_DATE, TO_CMPL_DATE)  ' เวลานี้ใช้ Area_Year ไหน
    Call LoadAreaComReport(Nothing, m_AreaCod, YEAR_ID)       ' โหลดไว้สำหรับ combo
    
     L = 1

   If TO_CMPL_DATE < 0 Then
         toCMPLdate = DateSerial(9999, 12, 31)
   Else:
          toCMPLdate = TO_CMPL_DATE
   End If
      
   Set ArS = New COESLM
   Call glbDaily.QuerySale(ArS, TempRs, iCount, IsOK, glbErrorLog)

   While Not TempRs.EOF          ' sale

         Call ArS.PopulateFromRS(1, TempRs)

         PrevKey1 = ArS.SLMCOD
         PrevKey2 = ArS.SLMNAM
         
         '===
           Set Stcrd = New CStcrd
      Stcrd.FROM_DOC_DATE = -1
      Stcrd.TO_DOC_DATE = -1
      Stcrd.SLMCOD = ArS.SLMCOD
      Call Stcrd.QueryData(13, Rs, iCount, 2) ' ติดต่อกับดาต้าเบสที่ 2
      While Not Rs.EOF
         Call Stcrd.PopulateFromRS(13, Rs)
            Set stcrd_temp = GetObject("CStcrd", stcrd_mixdb, Stcrd.DOCNUM & "-" & Stcrd.STKCOD, False)
            If stcrd_temp Is Nothing Then  ' ถ้าไม่มีในคอเล็กชั่น คือ เพิ่มได้
               Set stcrd_temp = New CStcrd
               stcrd_temp.TRNQTY = Stcrd.TRNQTY
            '   stcrd_temp.CMPLDAT = Stcrd.CMPLDAT
               stcrd_temp.NETVAL = Stcrd.NETVAL
               stcrd_temp.DOCNUM = Stcrd.DOCNUM
               stcrd_temp.CUSNAM = Stcrd.CUSNAM
               stcrd_temp.STKDES = Stcrd.STKDES
               stcrd_temp.DOCDAT = Stcrd.DOCDAT
               stcrd_temp.STKCOD = Stcrd.STKCOD
               stcrd_temp.SLMCOD = Stcrd.SLMCOD
               stcrd_temp.CUSCOD = Stcrd.CUSCOD
               stcrd_temp.UNITPR = Stcrd.UNITPR
               Call stcrd_mixdb.Add(stcrd_temp, stcrd_temp.DOCNUM & "-" & stcrd_temp.STKCOD)
               Set stcrd_temp = Nothing
            End If
         Rs.MoveNext                                                                                            ' วนคน    '  **** วนทุกบรรทัดของ Stcrd
      Wend
      
      Set Stcrd = New CStcrd
      Stcrd.FROM_DOC_DATE = -1
      Stcrd.TO_DOC_DATE = -1
      Stcrd.SLMCOD = ArS.SLMCOD
      Call Stcrd.QueryData(13, Rs, iCount)   ' ติดต่อกับดาต้าเบสที่ 1
      While Not Rs.EOF     '  **** วนทุกบรรทัดของ Stcrd
         Call Stcrd.PopulateFromRS(13, Rs)
            Set stcrd_temp = GetObject("CStcrd", stcrd_mixdb, Stcrd.DOCNUM & "-" & Stcrd.STKCOD, False)
            If stcrd_temp Is Nothing Then  ' ถ้าไม่มีในคอเล็กชั่น คือ เพิ่มได้
             Set stcrd_temp = New CStcrd
             stcrd_temp.TRNQTY = Stcrd.TRNQTY
      '       stcrd_temp.CMPLDAT = Stcrd.CMPLDAT
             stcrd_temp.NETVAL = Stcrd.NETVAL
             stcrd_temp.DOCNUM = Stcrd.DOCNUM
             stcrd_temp.CUSNAM = Stcrd.CUSNAM
             stcrd_temp.STKDES = Stcrd.STKDES
             stcrd_temp.DOCDAT = Stcrd.DOCDAT
             stcrd_temp.STKCOD = Stcrd.STKCOD
             stcrd_temp.SLMCOD = Stcrd.SLMCOD
             stcrd_temp.CUSCOD = Stcrd.CUSCOD
             stcrd_temp.UNITPR = Stcrd.UNITPR
             Call stcrd_mixdb.Add(stcrd_temp, stcrd_temp.DOCNUM & "-" & stcrd_temp.STKCOD)
             Set stcrd_temp = Nothing
            End If
         Rs.MoveNext                                                                                            ' วนคน    '  **** วนทุกบรรทัดของ Stcrd
      Wend
      

  L = 0              ' เปลี่ยนเซลล์ใหม่ ก็รอบเขตใหม่
  For Each temp_Area In m_AreaCod
  L = L + 1        ' ถ้าไม่ได้เลือก = ทุกเขต = วนเรื่อยๆ
        
    Set Stcrd = Nothing
    For Each Stcrd In stcrd_mixdb
    
      CorrectStkcod = False
      Set tempComDonStkcod = GetComDonStk(m_ComDonStkcod, Trim(Stcrd.STKCOD), False)
      If (tempComDonStkcod Is Nothing) Then                   ' มีในสินค้าที่ห้ามคิด com
         CorrectStkcod = True
      End If
      
      Set tempCusArea = GetCusAreaCom(temp_Area.ImportExportItems, Stcrd.CUSCOD, False)
      If (Not (tempCusArea Is Nothing)) And CorrectStkcod Then                  ' ในเขตนี้มีลูกค้า
   
         ' คำนวณดู IV ก่อน
         Set tempIVcenter = GetIVcenter(m_IVcenter, Stcrd.DOCNUM, False)        '
         If (tempIVcenter Is Nothing) Then   'ไม่เจอในคอเล็คชั่น = ไม่ใช่สินค้าพิเศษ
                  IVoldinArea = True
         Else:
                  IVoldinArea = False
                  ' ดังนั้นต้องเก็บ ไว้ในคอล
                  Set stcrd_IVinArea = New CStcrd
                  stcrd_IVinArea.AREACOD = Trim(Str(tempIVcenter.MASTER_AREA_ID))
                  stcrd_IVinArea.AREANAM = tempIVcenter.MASTER_AREA_NAME
                  stcrd_IVinArea.NETVAL = Stcrd.NETVAL
                  stcrd_IVinArea.DOCNUM = Stcrd.DOCNUM
                  stcrd_IVinArea.CUSNAM = Stcrd.CUSNAM
                  stcrd_IVinArea.STKDES = Stcrd.STKDES
                  stcrd_IVinArea.DOCDAT = Stcrd.DOCDAT
                  stcrd_IVinArea.STKCOD = Stcrd.STKCOD
                  stcrd_IVinArea.SLMCOD = Stcrd.SLMCOD
                  stcrd_IVinArea.CUSCOD = Stcrd.CUSCOD
                  Call m_IVinArea.Add(stcrd_IVinArea)
                  Set stcrd_IVinArea = Nothing
         End If
    
    If IVoldinArea = True Then  ' ******

               If tempCusArea.MASTER_AREA_ID <> PrevKey3 Then                        ' เปลี่ยนเขตเมื่อไหร่ จะเอาค่ามาคิด
                     If PrevKey3 <> 0 Then        'And Total1 <> 0 And Total2 <> 0
                        
                        Set temp_Head = GetSaleChart(m_saleChartHead, Trim(PrevKey1 & "-" & PrevKey3), False)               ' Head ก็ต้องมี key เป็นเขตด้วย
                        If Not (temp_Head Is Nothing) Then                                                              '  มีอยู่ใน Head ก็ดึงงบประมาณ head มา
                           If Val(temp_Head.BUDGET) <> 0 Then
                              DueCount = Round((Total2 / (Val(temp_Head.BUDGET) / 100)), 2)             '  คิดเป็น total ผลรวม  เอา Total2 มาคำนวณ %
                           Else
                              DueCount = 0
                           End If
                        Else
                               Set temp_Child2 = GetSaleChart(m_saleChartChild, Trim(PrevKey1 & "-" & PrevKey3), False)
                               If Not (temp_Child2 Is Nothing) Then
                                 If Val(temp_Child2.BUDGET) <> 0 Then
                                     DueCount = Round((Total2 / (Val(temp_Child2.BUDGET) / 100)), 2)
                                 Else
                                    DueCount = 0
                                 End If
                               Else
                                  DueCount = 0
                               End If
                         End If

                              Set TempCConditionCommiss = m_ConditionCommiss1.ITEM(1)
                              PercentSum = (TempCConditionCommiss.SLM_PERCENT / 100)
                                     'หาเปอร์เซ็นของ 100
                                    For Each m_runConditionCommiss In m_ConditionCommiss1
                                        If (m_runConditionCommiss.NUM_ONE) > 100 Then
                                             Percent100 = (m_runConditionCommiss.SLM_PERCENT / 100)
                                       End If
                                    Next m_runConditionCommiss
                              
                               For Each m_runConditionCommiss In m_ConditionCommiss1
                                  If (m_runConditionCommiss.NUM_ONE) >= DueCount Then
                                       PercentSum = (m_runConditionCommiss.SLM_PERCENT / 100)
                                 End If
                              Next m_runConditionCommiss

                            ' เกิน 100
                              AMOUNT = (Total2 * PercentSum)              ' เปลี่ยนจาก 1-->2

                              If DueCount > 100 Then    ' ???????? 100% ??????????????????? 100
                                    If Not (temp_Head Is Nothing) Then
                                      asAmount = 0
                                      MoreAmount = 0
                                      'Set m_runConditionCommiss = GetObject("CConditionCommission", m_ConditionCommiss1, "100")
                                      'asAmount = temp_Head.BUDGET * (m_runConditionCommiss.SLM_PERCENT / 100)
                                      asAmount = temp_Head.BUDGET * Percent100
   
                                      MoreAmount = Total2 - temp_Head.BUDGET
                                      MoreAmount = MoreAmount * PercentSum
   
                                       AMOUNT = MoreAmount + asAmount
                                    ElseIf Not (temp_Child2 Is Nothing) Then
                                      asAmount = 0
                                      MoreAmount = 0
                                      asAmount = temp_Child2.BUDGET * Percent100
   
                                      MoreAmount = Total2 - temp_Child2.BUDGET
                                      MoreAmount = MoreAmount * PercentSum
   
                                       AMOUNT = MoreAmount + asAmount
                                    Else
                                      AMOUNT = 0
                                    End If
                                  End If
                              
                                 Set temp_GPwithGroup = GetObject("CMasterFromToDetail", m_GPwithGroup, Trim(PrevKey1), False)   ' ของ
                                 If Not (temp_GPwithGroup Is Nothing) Then
                                     GP = temp_GPwithGroup.GP
                                     GROUP = temp_GPwithGroup.MASTER_PARAMETER_VALUE
                                 Else
                                     GP = 0
                                     GROUP = 1
                                End If
                                Total1 = AMOUNT * (GP / GROUP)      ' จ่ายจริง Com1 ขาย  (Total1 * PercentSum)
                               

                               
                                 Set D = New CChartTotal
                                 D.SALE_ID = PrevKey1
                                 D.AREA_ID = PrevKey3
                                 D.REAL_COM1 = Total1                        ' จ่ายจริงคอม 1
                                 D.REAL_COM2 = Total4                       ' จริง เก็บเงิน
                                D.REAL_INCENTIVE = Total5                  ' จริง Incentive
                                 Call collTotal3Com.Add(D, Trim(D.SALE_ID & "-" & D.AREA_ID))        ' 1
                                 Set D = Nothing
                     End If
                     Total1 = 0
                     Total2 = 0
                     Total3 = 0
                     Total4 = 0
                     Total5 = 0
                     Total6 = 0
             End If
               
'                If Stcrd.DOCNUM = "IV0044465" Then
'                  ''debug.print
'                  End If
              
               Set tempMinusStkcod = GetMinusCommiss(m_MinusStkcod, Stcrd.DOCDAT & "-" & Stcrd.DOCNUM & "-" & Stcrd.STKCOD, False)
               If (tempMinusStkcod Is Nothing) Then
                        NETVAL = Stcrd.NETVAL
               Else:
                        NETVAL = (Stcrd.NETVAL + Val(tempMinusStkcod.MINUS_AMOUNT))
               End If
                           
               Set TempCConditionCommiss = GetCheckCommiss(m_ConditionCommiss4, Stcrd.STKCOD, False)
               If (TempCConditionCommiss Is Nothing) Then
                             PercentNum1 = (100 / 100)
               Else
                             PercentNum1 = (TempCConditionCommiss.SLM_PERCENT / 100)   ' ยอดประเมินจะคิด 30 %
               End If
             
          '   If PrevKey5 <> Stcrd.DOCNUM Then   ' เพราะ IV ต้องซ้ำ ที่ไม่ซ้ำคือ IV+สินค้า
               Set IVinDateStcrd = GetObject(" CStcrd", m_IVinDateStcrd, Trim(Stcrd.DOCNUM), False)
               If Not (IVinDateStcrd Is Nothing) Then                          ' คอมขายรวมเฉพาะวันที่ IV อยู่ใน ที่ระบุ
                 '   ''debug.print Stcrd.DOCNUM & " - " & Stcrd.DOCDAT & ",,,"; IVinDateStcrd.DOCNUM & "-" & IVinDateStcrd.DOCDAT
                   Total1 = Total1 + NETVAL                                     ' ยอดจริงที่รวมส่วนลดแล้ว
                   Total2 = Total2 + (NETVAL * PercentNum1)             ' Total2 =ยอดประเมิน
               End If
          '   End If
           '      PrevKey5 = Stcrd.DOCNUM
                 
                '----------- ส่วนของ com2
               Set REsumIV = GetARRcpItem(m_REsumIV, Stcrd.DOCNUM, False)
               If Not (REsumIV Is Nothing) Then
                   Set tempREdoc = GetREDocDat(m_ReDocdat, Stcrd.DOCNUM, False)
                   If Not (tempREdoc Is Nothing) Then
                       CMPLDAT = tempREdoc.DOCDAT
                   Else
                      CMPLDAT = -1
                   End If
                  If ((REsumIV.RCVAMT - NETVAL) >= 0) And CMPLDAT >= FROM_CMPL_DATE And CMPLDAT <= TO_CMPL_DATE Then

                   DueCount2 = DateDiff("D", Stcrd.DOCDAT, CMPLDAT)
                    Set tempIVcredit = GetObject("CComIVcredit", m_IVcredit, Stcrd.DOCNUM, False)
                     If Not (tempIVcredit Is Nothing) Then
                         If tempIVcredit.CR_TYPE = "I" Then
                            DueCount2 = DueCount2 + tempIVcredit.CR_DATA
                         Else
                             DueCount2 = DueCount2 - tempIVcredit.CR_DATA
                         End If
                     End If
                     
                   Set TempCConditionCommiss = GetCheckCommiss(m_ConditionCommiss4, Stcrd.STKCOD, False)
                   If (TempCConditionCommiss Is Nothing) Then
                             PercentNum2 = (100 / 100)

                                             ' ถ้าเป็น เซล์และลูกค้าโปรโมต ใช้ comPro5
                                       Set temp_ComMasPro = GetComMasPro(m_ComMasPro, Trim(Stcrd.SLMCOD & "-" & Stcrd.CUSCOD), False)
                                       If Not (temp_ComMasPro Is Nothing) Then           ' หาโดยใช้  .SLMCOD .PEOPLE
                                              Set m_runComPro = m_ComPro2.ITEM(1)
                                              CreditKey = Val(m_runComPro.CREDIT_NAME)
                                              For Each m_runComPro In m_ComPro2                   ' เก็บเงินธรรมดา
                                                    If (Val(m_runComPro.CREDIT_NAME)) >= DueCount Then
                                                         CreditKey = Val(m_runComPro.CREDIT_NAME)    'จะได้ เครดิตเป็น Key
                                                    End If
                                               Next m_runComPro

                                               ' วน loop ของ YEAR_ID นี้ ,, 02(สินค้าธรรมดา) อยู่ในเครดิตไหน  แล้วมันจะได้ค่า MASTER_COMMISS_SUB_PROMOTE_ID เป็น Key
                                               Set TempComPro = GetComMasSubPro(m_ComPro2, Trim(Stcrd.SLMCOD & "-" & Stcrd.CUSCOD & "-" & CreditKey), False)
                                                If Not (TempComPro Is Nothing) Then
                                               PercentSum = (TempComPro.DetailsCom2.ITEM(1).SLM_PERCENT / 100)  'น้อยสุดป่ะ
                                               For Each m_runPro In TempComPro.DetailsCom2                    ' เก็บเงินธรรมดา
                                                     If (Val(m_runPro.NUM_ONE) + 7) >= DueCount Then
                                                             PercentSum = (m_runPro.SLM_PERCENT / 100)
                                                    End If
                                                Next m_runPro
                                                End If
                                       Else           ' ถ้าไม่ใช่สินค้าโปรโมต ให้ใช้ธรรมดา
                                                Set TempCConditionCommiss = m_ConditionCommiss2.ITEM(1)
                                                PercentSum2 = (TempCConditionCommiss.SLM_PERCENT / 100)
                                                For Each m_runConditionCommiss In m_ConditionCommiss2
                                                     If (m_runConditionCommiss.NUM_ONE + 7) >= DueCount2 Then
                                                          PercentSum2 = (m_runConditionCommiss.SLM_PERCENT / 100)
                                                    End If
                                                 Next m_runConditionCommiss
                                      End If
                  Else
                                 PercentNum2 = (TempCConditionCommiss.SLM_PERCENT / 100)   '
                                  ' ถ้าเป็น เซล์และลูกค้าโปรโมต ใช้ comPro6
                                 Set temp_ComMasPro = GetComMasPro(m_ComMasPro, Trim(Stcrd.SLMCOD & "-" & Stcrd.CUSCOD), False)
                                  If Not (temp_ComMasPro Is Nothing) Then           ' หาโดยใช้  .SLMCOD .PEOPLE
                                                      For Each m_runComPro In m_ComPro3                   ' เก็บเงินธรรมดา
                                                            If (Val(m_runComPro.CREDIT_NAME)) >= DueCount Then
                                                                 CreditKey = Val(m_runComPro.CREDIT_NAME)    'จะได้ เครดิตเป็น Key
                                                           End If
                                                      Next m_runComPro

                                                      ' วน loop ของ YEAR_ID นี้ ,, 02(สินค้าธรรมดา) อยู่ในเครดิตไหน  แล้วมันจะได้ค่า MASTER_COMMISS_SUB_PROMOTE_ID เป็น Key
                                                       Set TempComPro = GetComMasSubPro(m_ComPro3, Trim(Stcrd.SLMCOD & "-" & Stcrd.CUSCOD & "-" & CreditKey), False)
                                                       If Not (TempComPro Is Nothing) Then
                                                       PercentSum = (TempComPro.DetailsCom2.ITEM(1).SLM_PERCENT / 100)  'น้อยสุดป่ะ

                                                      For Each m_runPro In TempComPro.DetailsCom3                    ' เก็บเงินธรรมดา
                                                            If (Val(m_runPro.NUM_ONE) + 7) >= DueCount2 Then
                                                                     PercentSum = (m_runPro.SLM_PERCENT / 100)
                                                           End If
                                                      Next m_runPro
                                                      End If
                                     Else            ' ไม่ต้องคิดแบบโปรโมต

                                                      Set TempCConditionCommiss = m_ConditionCommiss3.ITEM(1)
                                                      PercentSum2 = (TempCConditionCommiss.SLM_PERCENT / 100)
                                                      For Each m_runConditionCommiss In m_ConditionCommiss3
                                                           If (m_runConditionCommiss.NUM_ONE + 7) >= DueCount2 Then
                                                                PercentSum2 = (m_runConditionCommiss.SLM_PERCENT / 100)
                                                          End If
                                                       Next m_runConditionCommiss
                                        End If  ' สินค้าพิเศษ sale&ลูกค้าโปรโมตหรือเปล่า
                  End If
                  
'                  If PrevKey1 = "16" Then
'                  ''debug.print ((NETVAL * PercentNum2) * PercentSum2)
'                  End If
                  Total3 = Total3 + (NETVAL * PercentNum2)   ' ยอดคิดค่าคอม * สินค้าพิเศษ 30% บางตัว
                  Total4 = Total4 + ((NETVAL * PercentNum2) * PercentSum2)

                     ' -------------------------- 3Incentive ในส่วนที่จ่ายครบ
'                          If Stcrd.DOCNUM = "IV0042761" Then
'                               ''debug.print
'                           End If
                     
          EnableIncentive = False                      '------------------------------------------
          DueCount = DateDiff("D", Stcrd.DOCDAT, CMPLDAT)
          Set tempIVcredit = GetObject("CComIVcredit", m_IVcredit, Stcrd.DOCNUM, False)
                 If Not (tempIVcredit Is Nothing) Then
                     If tempIVcredit.CR_TYPE = "I" Then
                        DueCount = DueCount + tempIVcredit.CR_DATA
                     Else
                         DueCount = DueCount - tempIVcredit.CR_DATA
                     End If
                 End If
          
           FlagNewCus = True

         Set m_runIncenPro2 = GetCheckIncenPro(m_IncenPro5_2, Stcrd.STKCOD, False)                 ' เงื่อนไขพิเศษ
         If (Not (m_runIncenPro2 Is Nothing)) Then        ' ต้องเป็นสินค้า Incentive ที่มีในเงื่อนไข
              Call LoadIncenPro05_1(m_IncenPro5_1, FROM_CMPL_DATE, TO_CMPL_DATE, Stcrd.STKCOD)   ' เงื่อนไข incentive
              '    โค๊ด if sumstkcod.TRNQTY  >= TempCConditionCommiss.NUM_ONE then ให้วนไป NUM_ONE ถัดไปของสินค้าตัวนั้น  ---- ต้องมี collection เก็บสินค้าใครมันหรือเปล่า
              '   จนกระทั้ง sumstkcod.TRNQTY <TempCConditionCommiss.NUM_ONE ก็ไม่ต้องเอาค่าใหม่ใส่ทับ และใช้ค่านั้นเป็น NUM_ONE
              NUM_ONE = 0                               ' น้อยสุด
              PercentSum3 = 0
              NUM_TWO = 0

           If DueCount <= NumCR Then    'EnableIncentive = True And
              For Each m_runIncenPro In m_IncenPro5_1
               If Stcrd.TRNQTY = 0 Then
                   Stcrd.TRNQTY = 1
               End If
                   If Val(NETVAL / Stcrd.TRNQTY) >= Val(m_runIncenPro.NUM_ONE) Then
                       NUM_ONE = m_runIncenPro.NUM_ONE    ' ใช้ m_runConditionCommiss.NUM_ONE เป็นคีย์ใน step3
                       EnableIncentive = True
                    End If
               Next m_runIncenPro
              End If  ' ถ้าระยะเวลาเก็บเงิน <= 97
              
               Set TempIncenPro = GetCheckIncenPro(m_IncenPro5, Trim(Stcrd.STKCOD & "-" & NUM_ONE), False)     '         Call Cl.Add(TempData, Trim(TempData.STKCOD & "-" & TempData.NUM_ONE & "-" & TempData.NUM_TWO))    ' KEy
               If Not (TempIncenPro Is Nothing) Then
                   ' ''debug.print Stcrd.DOCNUM
                   PercentSum3 = Val(TempIncenPro.SLM_PERCENT)
                   If TempIncenPro.OPERATOR = "Y" Then
                        FlagNewCus = False
                    End If
                End If
                   
'   ''debug.print Stcrd.SLMCOD & "-" & tempCusArea.MASTER_AREA_ID & "-" & Stcrd.DOCNUM & "-" & Stcrd.DOCDAT
'   ''debug.print NETVAL & "-" & Stcrd.TRNQTY & "  .." & PercentSum3
'   ''debug.print
            
    '------------------------------------------
    ElseIf m_runIncenPro2 Is Nothing Then
                  SumTRNQTY = 0
                  NUM_ONE = 0
                    Set m_runConditionCommiss2 = GetCheckCommiss(m_ConditionCommiss5_2, Stcrd.STKCOD, False)
                    If (Not (m_runConditionCommiss2 Is Nothing)) Then  ' เป็นสินค้า incentive
                        Set TempCConditionCommiss = GetCheckCommiss(m_StkcodGroup, Trim(Stcrd.STKCOD), False)
                        If Not (TempCConditionCommiss Is Nothing) Then
                             For Each m_runConditionCommiss In m_StkcodGroup
                                 If m_runConditionCommiss.GROUP1 = TempCConditionCommiss.GROUP1 Then
                                     Call GetFirstLastDate(DateSerial(Year(Stcrd.DOCDAT), Month(Stcrd.DOCDAT), 1), SixMonthFirst, SixMonthLast)
                                     Call GetFirstLastDate(DateSerial(Year(CMPLDAT), Month(CMPLDAT), 1), CMPLFirstDate, CMPLLastDate)
                                     Call LoadSumStcrdMonth(Nothing, m_SumStcrd, SixMonthFirst, SixMonthLast, CMPLFirstDate, CMPLLastDate)           ' โหลด sum TRNQTY groupby เดือนDOCDAT, สินค้า STKCOD     โดยดูผลรวมในเดือนนั้นๆ
                                     Set SumStkcod = GetStkcodNoNew(m_SumStcrd, m_runConditionCommiss.STKCOD)
                                     If Not (SumStkcod Is Nothing) Then
                                        SumTRNQTY = SumTRNQTY + SumStkcod.TRNQTY
                                     End If
                                 End If
                              Next m_runConditionCommiss
                       End If
                      
                             Call LoadCommission05_1(m_ConditionCommiss5_1, FROM_CMPL_DATE, TO_CMPL_DATE, Stcrd.STKCOD)
                             For Each m_runConditionCommiss In m_ConditionCommiss5_1
                                 If Val(SumTRNQTY) >= Val(m_runConditionCommiss.NUM_ONE) Then
                                     NUM_ONE = m_runConditionCommiss.NUM_ONE
                                     EnableIncentive = True
                                 End If
                             Next m_runConditionCommiss

                            PercentSum3 = 0
                            NUM_TWO = 0
                            If EnableIncentive = True And DueCount <= NumCR Then
                            For Each m_runConditionCommiss In m_ConditionCommiss5_1
                            If Stcrd.TRNQTY = 0 Then
                              Stcrd.TRNQTY = 1
                            End If
                                  If (NETVAL / Stcrd.TRNQTY) >= (m_runConditionCommiss.NUM_TWO) Then
                                     NUM_TWO = m_runConditionCommiss.NUM_TWO
                                    ' EnableNumTwo = True
                                 End If
                             Next m_runConditionCommiss
                    End If         ' 97
                    
                             Set TempCConditionCommiss = GetCheckCommiss(m_ConditionCommiss5, Trim(Stcrd.STKCOD & "-" & NUM_ONE & "-" & NUM_TWO), False)     '         Call Cl.Add(TempData, Trim(TempData.STKCOD & "-" & TempData.NUM_ONE & "-" & TempData.NUM_TWO))    ' KEy
                             If Not (TempCConditionCommiss Is Nothing) Then
                              '  ''debug.print Stcrd.DOCNUM
                              PercentSum3 = Val(TempCConditionCommiss.SLM_PERCENT)
                               If TempCConditionCommiss.OPERATOR = "Y" Then
                                        FlagNewCus = False
                               End If
                            End If
                            
'   ''debug.print Stcrd.SLMCOD & "-" & tempCusArea.MASTER_AREA_ID & "-" & Stcrd.DOCNUM & "-" & Stcrd.DOCDAT
'   ''debug.print NETVAL & "-" & Stcrd.TRNQTY & "  .." & PercentSum3
'   ''debug.print Total5
'   ''debug.print
 End If
End If

           If EnableIncentive = True Then
                 NEWCUSVAL = 0
                 If FlagNewCus = True And NETVAL >= 5000 And DueCount <= NumCR Then
                  Call GetFirstLastDate(DateSerial(Year(Stcrd.DOCDAT), Month(Stcrd.DOCDAT) - 6, 1), SixMonthFirst)
                  Call GetFirstLastDate(DateSerial(Year(Stcrd.DOCDAT), Month(Stcrd.DOCDAT) - 1, 1), , SixMonthLast)
                     If Year(SixMonthFirst) = Year(SixMonthLast) Then
                         Call LoadNewCus(m_NewCus, SixMonthFirst, SixMonthLast, Stcrd.CUSCOD)    ' ต้องดูย้อนหลังว่าลูกค้าคนนี้ย้อนไป 6 เดือนที่แล้ว มีใบสั่งซื้อไหม
                     Else
                        Call GetFirstLastDate(DateSerial(Year(SixMonthFirst), Month(1), 31), , dayLast)
                        Call LoadNewCus(m_NewCus, SixMonthFirst, dayLast, Stcrd.CUSCOD)          ' ปีเก่า เดือนนั้นๆ ถึงสิ้นปี
                        If m_NewCus.Count = 0 Then
                            Call GetFirstLastDate(DateSerial(Year(SixMonthLast), Month(12), 1), , dayFirst)
                            Call LoadNewCus(m_NewCus, dayFirst, SixMonthLast, Stcrd.CUSCOD)              ' ปีใหม่ ต้นปี ถึง เดือนนั้นๆ
                        End If
                     End If
                  Set temp_Stcrd = GetStkcodNoNew(m_AllNewCus, Stcrd.CUSCOD)
                  If (temp_Stcrd Is Nothing) Then

                        If m_NewCus.Count = 0 And temp_Stcrd Is Nothing Then
                           NEWCUSVAL = 300
                           Set temp_Stcrd = New CStcrd
                           temp_Stcrd.CUSCOD = Stcrd.CUSCOD
                           Call m_AllNewCus.Add(temp_Stcrd, Trim(Stcrd.CUSCOD))   ' 2
                           Set temp_Stcrd = Nothing
                        End If
                  End If
                  End If
                    Total5 = Total5 + NEWCUSVAL + (Val(Stcrd.TRNQTY) * PercentSum3)
          End If
              '   Else  ' ไม่ใช่สินค้า incentive แต่ต้องคิดลูกค้าใหม่
  If m_runIncenPro2 Is Nothing And m_runConditionCommiss2 Is Nothing And DueCount <= NumCR Then

'         If Stcrd.SLMCOD = "04" And tempCusArea.MASTER_AREA_ID = 5 Then
'              ''debug.print
'          End If

                     NEWCUSVAL = 300
                     Call GetFirstLastDate(DateSerial(Year(Stcrd.DOCDAT), Month(Stcrd.DOCDAT) - 6, 1), SixMonthFirst)
                     Call GetFirstLastDate(DateSerial(Year(Stcrd.DOCDAT), Month(Stcrd.DOCDAT) - 1, 1), , SixMonthLast)
                     If Year(SixMonthFirst) = Year(SixMonthLast) Then
                         Call LoadNewCus(m_NewCus, SixMonthFirst, SixMonthLast, Stcrd.CUSCOD)    ' ต้องดูย้อนหลังว่าลูกค้าคนนี้ย้อนไป 6 เดือนที่แล้ว มีใบสั่งซื้อไหม
                     Else
                        Call GetFirstLastDate(DateSerial(Year(SixMonthFirst), Month(1), 31), , dayLast)
                        Call LoadNewCus(m_NewCus, SixMonthFirst, dayLast, Stcrd.CUSCOD)          ' ปีเก่า เดือนนั้นๆ ถึงสิ้นปี
                        If m_NewCus.Count = 0 Then
                            Call GetFirstLastDate(DateSerial(Year(SixMonthLast), Month(12), 1), , dayFirst)
                            Call LoadNewCus(m_NewCus, dayFirst, SixMonthLast, Stcrd.CUSCOD)              ' ปีใหม่ ต้นปี ถึง เดือนนั้นๆ
                        End If
                     End If
                     Set temp_Stcrd = GetStkcodNoNew(m_AllNewCus, Stcrd.CUSCOD)
                     If m_NewCus.Count = 0 And temp_Stcrd Is Nothing And NETVAL > 5000 Then
                                  Total5 = Total5 + NEWCUSVAL
                                   Set temp_Stcrd = New CStcrd
                                    temp_Stcrd.CUSCOD = Stcrd.CUSCOD
                                    Call m_AllNewCus.Add(temp_Stcrd, Trim(Stcrd.CUSCOD))
                                    Set temp_Stcrd = Nothing
                     End If   '  เป็นลูกค้าใหม่ ที่ยังไม่มีในรายชื่อ ซื้อ >5000

'''debug.print Stcrd.SLMCOD & "-" & tempCusArea.MASTER_AREA_ID & "-" & Stcrd.DOCNUM & "-" & Stcrd.DOCDAT
'   ''debug.print NEWCUSVAL & "-"
'   ''debug.print Total5

 ' สินค้าตัวนี้ เป็นสินค้า Incentive
             '    =================  3Incentive ในส่วนที่จ่ายครบ

                End If   ' RE จ่าย IV ครบแล้ว และ RE อยู่ภายในวันที่นั้นๆ
              End If ' ดึง RE จาก Docnum
           '  =============== จบ Com2
 End If

                 PrevKey3 = tempCusArea.MASTER_AREA_ID
                 Set temp_Area2 = GetAreaCom(m_AreaCod, tempCusArea.MASTER_AREA_ID)
                 PrevKey4 = temp_Area2.MASTER_AREA_NAME '!
 End If
        End If  '****** IVinArea

   Next Stcrd
Next temp_Area

        For Each stcrd_IVinArea In m_IVinArea  ' ลูกค้าขายสด
         If Val(stcrd_IVinArea.AREACOD) = PrevKey3 And stcrd_IVinArea.SLMCOD = PrevKey1 Then
         '--------------- เหมือนข้างบนเด๊ะ stcrd_IVinArea<>Stcrd
             If stcrd_IVinArea.SLMCOD = PrevKey1 And stcrd_IVinArea.AREACOD <> PrevKey3 Then                        ' เปลี่ยนเขตเมื่อไหร่ จะเอาค่ามาคิด
                     If PrevKey3 <> 0 Then        'And Total1 <> 0 And Total2 <> 0
                        Set temp_Head = GetSaleChart(m_saleChartHead, Trim(PrevKey1 & "-" & PrevKey3), False)               ' Head ก็ต้องมี key เป็นเขตด้วย
                        If Not (temp_Head Is Nothing) Then                                                              '  มีอยู่ใน Head ก็ดึงงบประมาณ head มา
                           If Val(temp_Head.BUDGET) <> 0 Then
                              DueCount = Round((Total2 / (Val(temp_Head.BUDGET) / 100)), 2)             '  คิดเป็น total ผลรวม  เอา Total2 มาคำนวณ %
                           Else
                              DueCount = 0
                           End If
                        Else
                               Set temp_Child2 = GetSaleChart(m_saleChartChild, Trim(PrevKey1 & "-" & PrevKey3), False)
                               If Not (temp_Child2 Is Nothing) Then
                                 If Val(temp_Child2.BUDGET) <> 0 Then
                                  DueCount = Round((Total2 / (Val(temp_Child2.BUDGET) / 100)), 2)
                                 Else
                                    DueCount = 0
                                 End If
                               Else
                                  DueCount = 0
                               End If
                         End If

                              Set TempCConditionCommiss = m_ConditionCommiss1.ITEM(1)
                              PercentSum = (TempCConditionCommiss.SLM_PERCENT / 100)
                              
                                  'หาเปอร์เซ็นของ 100
                                    For Each m_runConditionCommiss In m_ConditionCommiss1
                                        If (m_runConditionCommiss.NUM_ONE) > 100 Then
                                             Percent100 = (m_runConditionCommiss.SLM_PERCENT / 100)
                                       End If
                                    Next m_runConditionCommiss
                                    
                               For Each m_runConditionCommiss In m_ConditionCommiss1
                                  If (m_runConditionCommiss.NUM_ONE) >= DueCount Then
                                       PercentSum = (m_runConditionCommiss.SLM_PERCENT / 100)
                                 End If
                              Next m_runConditionCommiss

                            ' เกิน 100
                              AMOUNT = (Total2 * PercentSum)

                              If DueCount > 100 Then    ' ???????? 100% ??????????????????? 100
                                    If Not (temp_Head Is Nothing) Then
                                      asAmount = 0
                                      MoreAmount = 0
                                  '    Set m_runConditionCommiss = GetObject("CConditionCommission", m_ConditionCommiss1, "100")
                                 '     asAmount = temp_Head.BUDGET * (m_runConditionCommiss.SLM_PERCENT / 100)
                                      asAmount = temp_Head.BUDGET * Percent100
   
                                      MoreAmount = Total2 - temp_Head.BUDGET
                                      MoreAmount = MoreAmount * PercentSum
   
                                       AMOUNT = MoreAmount + asAmount
                                    ElseIf Not (temp_Child2 Is Nothing) Then
                                      asAmount = 0
                                      MoreAmount = 0
                                    '  Set m_runConditionCommiss = GetObject("CConditionCommission", m_ConditionCommiss1, "100")
                                    '  asAmount = temp_Child2.BUDGET * (m_runConditionCommiss.SLM_PERCENT / 100)
                                       asAmount = temp_Child2.BUDGET * Percent100
   
                                      MoreAmount = Total2 - temp_Child2.BUDGET
                                      MoreAmount = MoreAmount * PercentSum
   
                                       AMOUNT = MoreAmount + asAmount
                                    Else
                                      AMOUNT = 0
                                    End If
                                  End If
                              
                                 Set temp_GPwithGroup = GetObject("CMasterFromToDetail", m_GPwithGroup, Trim(PrevKey1), False)   ' ของ
                                 If Not (temp_GPwithGroup Is Nothing) Then
                                     GP = temp_GPwithGroup.GP
                                     GROUP = temp_GPwithGroup.MASTER_PARAMETER_VALUE
                                 Else
                                     GP = 0
                                     GROUP = 1
                                End If
                                Total1 = AMOUNT * (GP / GROUP)      ' จ่ายจริง Com1 ขาย  (Total1 * PercentSum)

                                 Set D = New CChartTotal
                                 D.SALE_ID = PrevKey1
                                 D.AREA_ID = PrevKey3
                                 D.REAL_COM1 = Total1                        ' จ่ายจริงคอม 1
                                 D.REAL_COM2 = Total4                       ' จริง เก็บเงิน
                                D.REAL_INCENTIVE = Total5                  ' จริง Incentive
                                 Call collTotal3Com.Add(D, Trim(D.SALE_ID & "-" & D.AREA_ID))        ' 1
                                 Set D = Nothing
                     End If
                     Total1 = 0
                     Total2 = 0
                     Total3 = 0
                     Total4 = 0
                     Total5 = 0
                     Total6 = 0
             End If

'                If Stcrd.DOCNUM = "IV0044465" Then
'                  ''debug.print
'                  End If
               Set tempMinusStkcod = GetMinusCommiss(m_MinusStkcod, stcrd_IVinArea.DOCDAT & "-" & stcrd_IVinArea.DOCNUM & "-" & stcrd_IVinArea.STKCOD, False)
               If (tempMinusStkcod Is Nothing) Then
                        NETVAL = stcrd_IVinArea.NETVAL
               Else:
                        NETVAL = (stcrd_IVinArea.NETVAL + Val(tempMinusStkcod.MINUS_AMOUNT))
               End If
                           
               Set TempCConditionCommiss = GetCheckCommiss(m_ConditionCommiss4, stcrd_IVinArea.STKCOD, False)
               If (TempCConditionCommiss Is Nothing) Then
                             PercentNum1 = (100 / 100)
               Else
                             PercentNum1 = (TempCConditionCommiss.SLM_PERCENT / 100)   ' ยอดประเมินจะคิด 30 %
               End If
             
          '   If PrevKey5 <> Stcrd.DOCNUM Then   ' เพราะ IV ต้องซ้ำ ที่ไม่ซ้ำคือ IV+สินค้า
               Set IVinDateStcrd = GetObject(" CStcrd", m_IVinDateStcrd, Trim(stcrd_IVinArea.DOCNUM), False)
               If Not (IVinDateStcrd Is Nothing) Then                          ' คอมขายรวมเฉพาะวันที่ IV อยู่ใน ที่ระบุ
                 '   ''debug.print Stcrd.DOCNUM & " - " & Stcrd.DOCDAT & ",,,"; IVinDateStcrd.DOCNUM & "-" & IVinDateStcrd.DOCDAT
                   Total1 = Total1 + NETVAL                                     ' ยอดจริงที่รวมส่วนลดแล้ว
                   Total2 = Total2 + (NETVAL * PercentNum1)             ' Total2 =ยอดประเมิน
               End If
          '   End If
           '      PrevKey5 = Stcrd.DOCNUM
                 
                '----------- ส่วนของ com2
               Set REsumIV = GetARRcpItem(m_REsumIV, stcrd_IVinArea.DOCNUM, False)
               If Not (REsumIV Is Nothing) Then
                   Set tempREdoc = GetREDocDat(m_ReDocdat, stcrd_IVinArea.DOCNUM, False)
                   If Not (tempREdoc Is Nothing) Then
                       CMPLDAT = tempREdoc.DOCDAT
                   Else
                      CMPLDAT = -1
                   End If
                  If ((REsumIV.RCVAMT - NETVAL) >= 0) And CMPLDAT >= FROM_CMPL_DATE And CMPLDAT <= TO_CMPL_DATE Then

                   DueCount2 = DateDiff("D", stcrd_IVinArea.DOCDAT, CMPLDAT)
                    Set tempIVcredit = GetObject("CComIVcredit", m_IVcredit, stcrd_IVinArea.DOCNUM, False)
                     If Not (tempIVcredit Is Nothing) Then
                         If tempIVcredit.CR_TYPE = "I" Then
                            DueCount2 = DueCount2 + tempIVcredit.CR_DATA
                         Else
                             DueCount2 = DueCount2 - tempIVcredit.CR_DATA
                         End If
                     End If
                     
                   Set TempCConditionCommiss = GetCheckCommiss(m_ConditionCommiss4, stcrd_IVinArea.STKCOD, False)
                   If (TempCConditionCommiss Is Nothing) Then
                             PercentNum2 = (100 / 100)

                                             ' ถ้าเป็น เซล์และลูกค้าโปรโมต ใช้ comPro5
                                       Set temp_ComMasPro = GetComMasPro(m_ComMasPro, Trim(stcrd_IVinArea.SLMCOD & "-" & stcrd_IVinArea.CUSCOD), False)
                                       If Not (temp_ComMasPro Is Nothing) Then           ' หาโดยใช้  .SLMCOD .PEOPLE
                                              Set m_runComPro = m_ComPro2.ITEM(1)
                                              CreditKey = Val(m_runComPro.CREDIT_NAME)
                                              For Each m_runComPro In m_ComPro2                   ' เก็บเงินธรรมดา
                                                    If (Val(m_runComPro.CREDIT_NAME)) >= DueCount Then
                                                         CreditKey = Val(m_runComPro.CREDIT_NAME)    'จะได้ เครดิตเป็น Key
                                                    End If
                                               Next m_runComPro

                                               ' วน loop ของ YEAR_ID นี้ ,, 02(สินค้าธรรมดา) อยู่ในเครดิตไหน  แล้วมันจะได้ค่า MASTER_COMMISS_SUB_PROMOTE_ID เป็น Key
                                               Set TempComPro = GetComMasSubPro(m_ComPro2, Trim(stcrd_IVinArea.SLMCOD & "-" & stcrd_IVinArea.CUSCOD & "-" & CreditKey), False)
                                                If Not (TempComPro Is Nothing) Then
                                               PercentSum = (TempComPro.DetailsCom2.ITEM(1).SLM_PERCENT / 100)  'น้อยสุดป่ะ
                                               For Each m_runPro In TempComPro.DetailsCom2                    ' เก็บเงินธรรมดา
                                                     If (Val(m_runPro.NUM_ONE) + 7) >= DueCount2 Then
                                                             PercentSum = (m_runPro.SLM_PERCENT / 100)
                                                    End If
                                                Next m_runPro
                                                End If
                                       Else           ' ถ้าไม่ใช่สินค้าโปรโมต ให้ใช้ธรรมดา
                                                Set TempCConditionCommiss = m_ConditionCommiss2.ITEM(1)
                                                PercentSum2 = (TempCConditionCommiss.SLM_PERCENT / 100)
                                                For Each m_runConditionCommiss In m_ConditionCommiss2
                                                     If (m_runConditionCommiss.NUM_ONE + 7) >= DueCount2 Then
                                                          PercentSum2 = (m_runConditionCommiss.SLM_PERCENT / 100)
                                                    End If
                                                 Next m_runConditionCommiss
                                      End If
                  Else
                                 PercentNum2 = (TempCConditionCommiss.SLM_PERCENT / 100)   '
                                  ' ถ้าเป็น เซล์และลูกค้าโปรโมต ใช้ comPro6
                                 Set temp_ComMasPro = GetComMasPro(m_ComMasPro, Trim(stcrd_IVinArea.SLMCOD & "-" & stcrd_IVinArea.CUSCOD), False)
                                  If Not (temp_ComMasPro Is Nothing) Then           ' หาโดยใช้  .SLMCOD .PEOPLE
                                                      For Each m_runComPro In m_ComPro3                   ' เก็บเงินธรรมดา
                                                            If (Val(m_runComPro.CREDIT_NAME)) >= DueCount Then
                                                                 CreditKey = Val(m_runComPro.CREDIT_NAME)    'จะได้ เครดิตเป็น Key
                                                           End If
                                                      Next m_runComPro

                                                      ' วน loop ของ YEAR_ID นี้ ,, 02(สินค้าธรรมดา) อยู่ในเครดิตไหน  แล้วมันจะได้ค่า MASTER_COMMISS_SUB_PROMOTE_ID เป็น Key
                                                       Set TempComPro = GetComMasSubPro(m_ComPro3, Trim(stcrd_IVinArea.SLMCOD & "-" & stcrd_IVinArea.CUSCOD & "-" & CreditKey), False)
                                                       If Not (TempComPro Is Nothing) Then
                                                       PercentSum = (TempComPro.DetailsCom2.ITEM(1).SLM_PERCENT / 100)  'น้อยสุดป่ะ

                                                      For Each m_runPro In TempComPro.DetailsCom3                    ' เก็บเงินธรรมดา
                                                            If (Val(m_runPro.NUM_ONE) + 7) >= DueCount2 Then
                                                                     PercentSum = (m_runPro.SLM_PERCENT / 100)
                                                           End If
                                                      Next m_runPro
                                                      End If
                                     Else            ' ไม่ต้องคิดแบบโปรโมต

                                                      Set TempCConditionCommiss = m_ConditionCommiss3.ITEM(1)
                                                      PercentSum2 = (TempCConditionCommiss.SLM_PERCENT / 100)
                                                      For Each m_runConditionCommiss In m_ConditionCommiss3
                                                           If (m_runConditionCommiss.NUM_ONE + 7) >= DueCount2 Then
                                                                PercentSum2 = (m_runConditionCommiss.SLM_PERCENT / 100)
                                                          End If
                                                       Next m_runConditionCommiss
                                        End If  ' สินค้าพิเศษ sale&ลูกค้าโปรโมตหรือเปล่า
                  End If

                  Total3 = Total3 + (NETVAL * PercentNum2)   ' ยอดคิดค่าคอม * สินค้าพิเศษ 30% บางตัว
                  Total4 = Total4 + ((NETVAL * PercentNum2) * PercentSum2)

                     '  3Incentive ในส่วนที่จ่ายครบ
          EnableIncentive = False                      '------------------------------------------
          DueCount = DateDiff("D", stcrd_IVinArea.DOCDAT, CMPLDAT)
          Set tempIVcredit = GetObject("CComIVcredit", m_IVcredit, stcrd_IVinArea.DOCNUM, False)
                 If Not (tempIVcredit Is Nothing) Then
                     If tempIVcredit.CR_TYPE = "I" Then
                        DueCount = DueCount + tempIVcredit.CR_DATA
                     Else
                         DueCount = DueCount - tempIVcredit.CR_DATA
                     End If
                 End If
           FlagNewCus = True

         Set m_runIncenPro2 = GetCheckIncenPro(m_IncenPro5_2, stcrd_IVinArea.STKCOD, False)                 ' เงื่อนไขพิเศษ
         If (Not (m_runIncenPro2 Is Nothing)) Then        ' ต้องเป็นสินค้า Incentive ที่มีในเงื่อนไข
              Call LoadIncenPro05_1(m_IncenPro5_1, FROM_CMPL_DATE, TO_CMPL_DATE, stcrd_IVinArea.STKCOD)   ' เงื่อนไข incentive
              NUM_ONE = 0                               ' น้อยสุด
              PercentSum3 = 0
              NUM_TWO = 0

           If DueCount <= NumCR Then    'EnableIncentive = True And
              For Each m_runIncenPro In m_IncenPro5_1
                  If stcrd_IVinArea.TRNQTY = 0 Then
                   stcrd_IVinArea.TRNQTY = 1
                   End If
                   If Val(NETVAL / stcrd_IVinArea.TRNQTY) >= Val(m_runIncenPro.NUM_ONE) Then
                       NUM_ONE = m_runIncenPro.NUM_ONE    ' ใช้ m_runConditionCommiss.NUM_ONE เป็นคีย์ใน step3
                       EnableIncentive = True
                    End If
               Next m_runIncenPro
              End If  ' ถ้าระยะเวลาเก็บเงิน <= 97
              
               Set TempIncenPro = GetCheckIncenPro(m_IncenPro5, Trim(stcrd_IVinArea.STKCOD & "-" & NUM_ONE), False)     '         Call Cl.Add(TempData, Trim(TempData.STKCOD & "-" & TempData.NUM_ONE & "-" & TempData.NUM_TWO))    ' KEy
               If Not (TempIncenPro Is Nothing) Then
                   ' ''debug.print Stcrd.DOCNUM
                   PercentSum3 = Val(TempIncenPro.SLM_PERCENT)
                   If TempIncenPro.OPERATOR = "Y" Then
                        FlagNewCus = False
                    End If
                End If
     ElseIf m_runIncenPro2 Is Nothing Then
                  SumTRNQTY = 0
                  NUM_ONE = 0
                    Set m_runConditionCommiss2 = GetCheckCommiss(m_ConditionCommiss5_2, stcrd_IVinArea.STKCOD, False)
                    If (Not (m_runConditionCommiss2 Is Nothing)) Then  ' เป็นสินค้า incentive
                        Set TempCConditionCommiss = GetCheckCommiss(m_StkcodGroup, Trim(stcrd_IVinArea.STKCOD), False)
                        If Not (TempCConditionCommiss Is Nothing) Then
                             For Each m_runConditionCommiss In m_StkcodGroup
                                 If m_runConditionCommiss.GROUP1 = TempCConditionCommiss.GROUP1 Then
                                    Call GetFirstLastDate(DateSerial(Year(stcrd_IVinArea.DOCDAT), Month(stcrd_IVinArea.DOCDAT), 1), SixMonthFirst, SixMonthLast)
                                    Call GetFirstLastDate(DateSerial(Year(CMPLDAT), Month(CMPLDAT), 1), CMPLFirstDate, CMPLLastDate)
                                    Call LoadSumStcrdMonth(Nothing, m_SumStcrd, SixMonthFirst, SixMonthLast, CMPLFirstDate, CMPLLastDate)           ' โหลด sum TRNQTY groupby เดือนDOCDAT, สินค้า STKCOD     โดยดูผลรวมในเดือนนั้นๆ
                                    Set SumStkcod = GetStkcodNoNew(m_SumStcrd, m_runConditionCommiss.STKCOD)
                                    If Not (SumStkcod Is Nothing) Then
                                        SumTRNQTY = SumTRNQTY + SumStkcod.TRNQTY
                                    End If
                                 End If
                              Next m_runConditionCommiss
                       End If
                      
                             Call LoadCommission05_1(m_ConditionCommiss5_1, FROM_CMPL_DATE, TO_CMPL_DATE, stcrd_IVinArea.STKCOD)
                             For Each m_runConditionCommiss In m_ConditionCommiss5_1
                                 If Val(SumTRNQTY) >= Val(m_runConditionCommiss.NUM_ONE) Then
                                     NUM_ONE = m_runConditionCommiss.NUM_ONE
                                     EnableIncentive = True
                                 End If
                             Next m_runConditionCommiss

                            PercentSum3 = 0
                            NUM_TWO = 0
                            If EnableIncentive = True And DueCount <= NumCR Then
                            For Each m_runConditionCommiss In m_ConditionCommiss5_1
                              If stcrd_IVinArea.TRNQTY = 0 Then
                                  stcrd_IVinArea.TRNQTY = 1
                              End If
                                  If (NETVAL / stcrd_IVinArea.TRNQTY) >= (m_runConditionCommiss.NUM_TWO) Then
                                     NUM_TWO = m_runConditionCommiss.NUM_TWO
                                    ' EnableNumTwo = True
                                 End If
                             Next m_runConditionCommiss
                    End If         ' 97
                    
                             Set TempCConditionCommiss = GetCheckCommiss(m_ConditionCommiss5, Trim(stcrd_IVinArea.STKCOD & "-" & NUM_ONE & "-" & NUM_TWO), False)     '         Call Cl.Add(TempData, Trim(TempData.STKCOD & "-" & TempData.NUM_ONE & "-" & TempData.NUM_TWO))    ' KEy
                             If Not (TempCConditionCommiss Is Nothing) Then
                              '  ''debug.print Stcrd.DOCNUM
                              PercentSum3 = Val(TempCConditionCommiss.SLM_PERCENT)
                               If TempCConditionCommiss.OPERATOR = "Y" Then
                                        FlagNewCus = False
                               End If
                            End If
                End If
               End If

           If EnableIncentive = True Then
                 NEWCUSVAL = 0
                 If FlagNewCus = True And NETVAL >= 5000 Then
                  Call GetFirstLastDate(DateSerial(Year(stcrd_IVinArea.DOCDAT), Month(stcrd_IVinArea.DOCDAT) - 6, 1), SixMonthFirst)
                  Call GetFirstLastDate(DateSerial(Year(stcrd_IVinArea.DOCDAT), Month(stcrd_IVinArea.DOCDAT) - 1, 1), , SixMonthLast)
                     If Year(SixMonthFirst) = Year(SixMonthLast) Then
                         Call LoadNewCus(m_NewCus, SixMonthFirst, SixMonthLast, stcrd_IVinArea.CUSCOD)    ' ต้องดูย้อนหลังว่าลูกค้าคนนี้ย้อนไป 6 เดือนที่แล้ว มีใบสั่งซื้อไหม
                     Else
                        Call GetFirstLastDate(DateSerial(Year(SixMonthFirst), Month(1), 31), , dayLast)
                        Call LoadNewCus(m_NewCus, SixMonthFirst, dayLast, stcrd_IVinArea.CUSCOD)          ' ปีเก่า เดือนนั้นๆ ถึงสิ้นปี
                        If m_NewCus.Count = 0 Then
                            Call GetFirstLastDate(DateSerial(Year(SixMonthLast), Month(12), 1), , dayFirst)
                            Call LoadNewCus(m_NewCus, dayFirst, SixMonthLast, stcrd_IVinArea.CUSCOD)              ' ปีใหม่ ต้นปี ถึง เดือนนั้นๆ
                        End If
                     End If
                  Set temp_Stcrd = GetStkcodNoNew(m_AllNewCus, stcrd_IVinArea.CUSCOD)
                  If (temp_Stcrd Is Nothing) Then

                        If m_NewCus.Count = 0 And temp_Stcrd Is Nothing Then
                           NEWCUSVAL = 300
                           Set temp_Stcrd = New CStcrd
                           temp_Stcrd.CUSCOD = stcrd_IVinArea.CUSCOD
                           Call m_AllNewCus.Add(temp_Stcrd, Trim(stcrd_IVinArea.CUSCOD))   ' 2
                           Set temp_Stcrd = Nothing
                        End If
                  End If
                  End If
                    Total5 = Total5 + NEWCUSVAL + (Val(stcrd_IVinArea.TRNQTY) * PercentSum3)
          End If
              '   Else  ' ไม่ใช่สินค้า incentive แต่ต้องคิดลูกค้าใหม่
              If m_runIncenPro2 Is Nothing And m_runConditionCommiss2 Is Nothing Then
                     NEWCUSVAL = 300
                     Call GetFirstLastDate(DateSerial(Year(stcrd_IVinArea.DOCDAT), Month(stcrd_IVinArea.DOCDAT) - 6, 1), SixMonthFirst)
                     Call GetFirstLastDate(DateSerial(Year(stcrd_IVinArea.DOCDAT), Month(stcrd_IVinArea.DOCDAT) - 1, 1), , SixMonthLast)
                     If Year(SixMonthFirst) = Year(SixMonthLast) Then
                         Call LoadNewCus(m_NewCus, SixMonthFirst, SixMonthLast, stcrd_IVinArea.CUSCOD)    ' ต้องดูย้อนหลังว่าลูกค้าคนนี้ย้อนไป 6 เดือนที่แล้ว มีใบสั่งซื้อไหม
                     Else
                        Call GetFirstLastDate(DateSerial(Year(SixMonthFirst), Month(1), 31), , dayLast)
                        Call LoadNewCus(m_NewCus, SixMonthFirst, dayLast, stcrd_IVinArea.CUSCOD)          ' ปีเก่า เดือนนั้นๆ ถึงสิ้นปี
                        If m_NewCus.Count = 0 Then
                            Call GetFirstLastDate(DateSerial(Year(SixMonthLast), Month(12), 1), , dayFirst)
                            Call LoadNewCus(m_NewCus, dayFirst, SixMonthLast, stcrd_IVinArea.CUSCOD)              ' ปีใหม่ ต้นปี ถึง เดือนนั้นๆ
                        End If
                     End If
                     Set temp_Stcrd = GetStkcodNoNew(m_AllNewCus, stcrd_IVinArea.CUSCOD)
                     If m_NewCus.Count = 0 And temp_Stcrd Is Nothing And NETVAL > 5000 Then
                                  Total5 = Total5 + NEWCUSVAL
                                   Set temp_Stcrd = New CStcrd
                                    temp_Stcrd.CUSCOD = stcrd_IVinArea.CUSCOD
                                    Call m_AllNewCus.Add(temp_Stcrd, Trim(stcrd_IVinArea.CUSCOD))
                                    Set temp_Stcrd = Nothing
                     End If   '  เป็นลูกค้าใหม่ ที่ยังไม่มีในรายชื่อ ซื้อ >500
                       End If   ' RE จ่าย IV ครบแล้ว และ RE อยู่ภายในวันที่นั้นๆ
                     End If ' ดึง RE จาก Docnum
                 End If                  '  =============== จบ Com2

                 PrevKey3 = stcrd_IVinArea.AREACOD
                 Set temp_Area2 = GetAreaCom(m_AreaCod, stcrd_IVinArea.AREACOD)
                 PrevKey4 = temp_Area2.MASTER_AREA_NAME '!
         '-----------------------------------------
         End If  ' ต่อท้าย เซลล์คนเดียว เขตเดียวกัน
         Next stcrd_IVinArea
         
         
         If PrevKey3 <> 0 Then                    ' ??????????????????? Total1 <> 0 And And Total2 <> 0
            Set temp_Head = GetSaleChart(m_saleChartHead, PrevKey1, False)
            If Not (temp_Head Is Nothing) Then                                                              '  มีอยู่ใน Head ก็ดึงงบประมาณ head มา
               If Val(temp_Head.BUDGET) <> 0 Then
                  DueCount = Round((Total2 / (Val(temp_Head.BUDGET) / 100)), 2)             '  คิดเป็น total ผลรวม  เอา Total2 มาคำนวณ %
               Else
                  DueCount = 0
               End If
            Else
                   Set temp_Child2 = GetSaleChart(m_saleChartChild, Trim(PrevKey1 & "-" & PrevKey3), False)
                   If Not (temp_Child2 Is Nothing) Then
                     If Val(temp_Child2.BUDGET) <> 0 Then
                      DueCount = Round((Total2 / (Val(temp_Child2.BUDGET) / 100)), 2)
                     Else
                       DueCount = 0
                     End If
                   Else
                      DueCount = 0
                   End If
             End If
             
                  Set TempCConditionCommiss = m_ConditionCommiss1.ITEM(1)
                  PercentSum = (TempCConditionCommiss.SLM_PERCENT / 100)
                  
                  'หาเปอร์เซ็นของ 100
                  For Each m_runConditionCommiss In m_ConditionCommiss1
                      If (m_runConditionCommiss.NUM_ONE) > 100 Then
                           Percent100 = (m_runConditionCommiss.SLM_PERCENT / 100)
                     End If
                  Next m_runConditionCommiss
                  
                   For Each m_runConditionCommiss In m_ConditionCommiss1
                      If (m_runConditionCommiss.NUM_ONE) >= DueCount Then
                           PercentSum = (m_runConditionCommiss.SLM_PERCENT / 100)
                     End If
                  Next m_runConditionCommiss
                  
                              ' เกิน 100
'                              If PrevKey1 = "02" And PrevKey3 = 8 Then
'                                 ''debug.print
'                              End If
                              AMOUNT = (Total2 * PercentSum)

                              If DueCount > 100 Then    ' ???????? 100% ??????????????????? 100
                                    If Not (temp_Head Is Nothing) Then
                                      asAmount = 0
                                      MoreAmount = 0
                                   '   Set m_runConditionCommiss = GetObject("CConditionCommission", m_ConditionCommiss1, "100")
                                  '    asAmount = temp_Head.BUDGET * (m_runConditionCommiss.SLM_PERCENT / 100)
                                     asAmount = temp_Head.BUDGET * Percent100
    
                                      MoreAmount = Total2 - temp_Head.BUDGET
                                      MoreAmount = MoreAmount * PercentSum
   
                                       AMOUNT = MoreAmount + asAmount
                                    ElseIf Not (temp_Child2 Is Nothing) Then
                                      asAmount = 0
                                      MoreAmount = 0
                                  '    Set m_runConditionCommiss = GetObject("CConditionCommission", m_ConditionCommiss1, "100")
                                 '     asAmount = temp_Child2.BUDGET * (m_runConditionCommiss.SLM_PERCENT / 100)
                                      asAmount = temp_Child2.BUDGET * Percent100
   
                                      MoreAmount = Total2 - temp_Child2.BUDGET
                                      MoreAmount = MoreAmount * PercentSum
   
                                       AMOUNT = MoreAmount + asAmount
                                    Else
                                      AMOUNT = 0
                                    End If
                              End If
                  
                     Set temp_GPwithGroup = GetObject("CMasterFromToDetail", m_GPwithGroup, Trim(PrevKey1), False)   ' ของ
                     If Not (temp_GPwithGroup Is Nothing) Then
                         GP = temp_GPwithGroup.GP
                         GROUP = temp_GPwithGroup.MASTER_PARAMETER_VALUE
                     Else
                         GP = 0
                         GROUP = 1
                    End If
                  Total1 = AMOUNT * (GP / GROUP)     'Total1 = (Total1 * PercentSum) * (GP / GROUP)      ' จ่ายจริง Com1 ขาย
             
'                     If PrevKey1 = "18" And PrevKey3 = 7 Then
'                     ''debug.print
'                     End If
                     Set D = New CChartTotal
                     D.SALE_ID = PrevKey1
                     D.AREA_ID = PrevKey3
                     D.REAL_COM1 = Total1                        ' จ่ายจริงคอม 1
                     D.REAL_COM2 = Total4                       ' จริง เก็บเงิน
                      D.REAL_INCENTIVE = Total5                  ' จริง Incentive
                     Call collTotal3Com.Add(D, Trim(D.SALE_ID & "-" & D.AREA_ID))            ' 3
            Set D = Nothing
                      
          End If
         Total1 = 0
         Total2 = 0
            Total3 = 0
         Total4 = 0
            Total5 = 0
         Total6 = 0
         
           PrevKey3 = 0
           PrevKey4 = ""
           
            Set stcrd_mixdb = Nothing
            Set stcrd_mixdb = New Collection
           
   TempRs.MoveNext                                                            ' ???????
Wend

   Set ArS = Nothing
   Set Stcrd = Nothing
   Set TempCConditionCommiss = Nothing
   Set tempMinusStkcod = Nothing
   Set m_runConditionCommiss = Nothing
   Set temp_Area = Nothing
    Set temp_Area2 = Nothing
   Set m_ConditionCommiss4 = Nothing
   Set m_AreaCod = Nothing
   Set tempCusArea = Nothing
   Set D = Nothing
   Set m_ComPro2 = Nothing
   Set m_ComPro3 = Nothing
   Set m_ComPro4 = Nothing
   Set m_ComMasPro = Nothing
   Set temp_ComMasPro = Nothing
   Set m_IncenPro5 = Nothing
   Set m_IncenPro5_1 = Nothing
   Set m_IncenPro5_2 = Nothing
   Set m_IncenProGroup = Nothing
   Set m_runIncenPro2 = Nothing
   Set TempIncenPro = Nothing
   Set m_runIncenPro = Nothing
   Set m_IVcredit = Nothing
   Set tempIVcredit = Nothing
End Sub

Public Sub LoadComDonStk(Optional Cl As Collection = Nothing, Optional FromDocDat As Date = -1, Optional ToDocDat As Date = -1)
On Error GoTo ErrorHandler
Dim D As CComDonStk
Dim ItemCount As Long
Dim TempData As CComDonStk
Dim Rs As ADODB.Recordset
Dim i As Long

   Set D = New CComDonStk
   Set Rs = New ADODB.Recordset
   
   D.VALID_FROM = FromDocDat
   D.VALID_TO = ToDocDat
   Call D.QueryData(1, Rs, ItemCount)   ' ดึงข้อมูลเพื่อเอาไปเก็บในคอเล็กชั่น
   
   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   
   While Not Rs.EOF
      i = i + 1
      Set TempData = New CComDonStk
      Call TempData.PopulateFromRS(1, Rs)

      If Not (Cl Is Nothing) Then
         Call Cl.Add(TempData, Trim(TempData.STKCOD))    ' KEy
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   Set Rs = Nothing
   Set D = Nothing
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub

Public Sub LoadYearId(YearID As Long, Optional FromCMPLDat As Date = -1, Optional ToCMPLDat As Date = -1)
On Error GoTo ErrorHandler
Dim D As CAreaYear
Dim ItemCount As Long
Dim TempData As CAreaYear
Dim Rs As ADODB.Recordset
Dim i As Long

   Set D = New CAreaYear
   Set Rs = New ADODB.Recordset
   
   D.FROM_DATE = FromCMPLDat
   D.TO_DATE = ToCMPLDat
   Call D.QueryData(1, Rs, ItemCount)   ' ดึงข้อมูลเพื่อเอาไปเก็บในคอเล็กชั่น
   
   While Not Rs.EOF
   
      Set TempData = New CAreaYear
      Call TempData.PopulateFromRS(1, Rs)
      YearID = TempData.YEAR_ID
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   Set Rs = Nothing
   Set D = Nothing
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub


Public Function getCommissYearMax(Ua As CCommissYear, IsOK As Boolean, ErrorObj As clsErrorLog) As Long
On Error GoTo ErrorHandler
Dim RName As String
Dim iCount As Long
Dim Rs As ADODB.Recordset
Dim ItemCount As Long

   RName = "getCommissYearMax"
getCommissYearMax = False

   Set Rs = New ADODB.Recordset
   IsOK = True
   Ua.YEARNUM = ""
   Call Ua.QueryData(2, Rs, ItemCount)
   Call Ua.PopulateFromRS(2, Rs)
   getCommissYearMax = Ua.YEAR_ID
   Exit Function

ErrorHandler:
   ErrorObj.SystemErrorMsg = Err.DESCRIPTION
   ErrorObj.RoutineName = RName
   ErrorObj.ShowErrorLog (LOG_FILE_MSGBOX)
   getCommissYearMax = -1
End Function

Public Function getMasFromTo2Max(Ua As CMaster2FromTo, IsOK As Boolean, ErrorObj As clsErrorLog) As Long
On Error GoTo ErrorHandler
Dim RName As String
Dim iCount As Long
Dim Rs As ADODB.Recordset
Dim ItemCount As Long

   RName = "getMasFromTo2Max"
   getMasFromTo2Max = False

   Set Rs = New ADODB.Recordset
   IsOK = True
'   Ua.YEARNUM = ""
'   Ua.YEAR_ID = -1
   Ua.VALID_FROM = -1
   Ua.VALID_TO = -1
   Call Ua.QueryData(2, Rs, ItemCount)
   Call Ua.PopulateFromRS(2, Rs)
   getMasFromTo2Max = Ua.MASTER_FROMTO_ID
   Exit Function

ErrorHandler:
   ErrorObj.SystemErrorMsg = Err.DESCRIPTION
   ErrorObj.RoutineName = RName
   ErrorObj.ShowErrorLog (LOG_FILE_MSGBOX)
   getMasFromTo2Max = -1
End Function

Public Function getAreaYearMax(Ua As CAreaYear, IsOK As Boolean, ErrorObj As clsErrorLog) As Long
On Error GoTo ErrorHandler
Dim RName As String
Dim iCount As Long
Dim Rs As ADODB.Recordset
Dim ItemCount As Long

   RName = "getAreaYearMax"
   getAreaYearMax = False

   Set Rs = New ADODB.Recordset
   IsOK = True
   Ua.YEARNUM = ""
   Ua.YEAR_ID = -1
   Ua.FROM_DATE = -1
   Ua.TO_DATE = -1
   Call Ua.QueryData(2, Rs, ItemCount)
   Call Ua.PopulateFromRS(2, Rs)
   getAreaYearMax = Ua.YEAR_ID
   Exit Function

ErrorHandler:
   ErrorObj.SystemErrorMsg = Err.DESCRIPTION
   ErrorObj.RoutineName = RName
   ErrorObj.ShowErrorLog (LOG_FILE_MSGBOX)
   getAreaYearMax = -1
End Function

Public Function getMasterFromToMax(Ua As CMasterFromTo, IsOK As Boolean, ErrorObj As clsErrorLog) As Long
On Error GoTo ErrorHandler
Dim RName As String
Dim iCount As Long
Dim Rs As ADODB.Recordset
Dim ItemCount As Long

   RName = "getMasterFromToMax"
   getMasterFromToMax = False

   Set Rs = New ADODB.Recordset
   IsOK = True
'   Ua.YEARNUM = ""
   Ua.MASTER_FROMTO_TYPE = -1
   Ua.VALID_FROM = -1
   Ua.VALID_TO = -1
   Call Ua.QueryData(2, Rs, ItemCount)
   Call Ua.PopulateFromRS(2, Rs)
   getMasterFromToMax = Ua.MASTER_FROMTO_ID
   Exit Function

ErrorHandler:
   ErrorObj.SystemErrorMsg = Err.DESCRIPTION
   ErrorObj.RoutineName = RName
   ErrorObj.ShowErrorLog (LOG_FILE_MSGBOX)
   getMasterFromToMax = -1
End Function

Public Function GetComDonStk(m_TempCol As Collection, TempKey As String, Optional HaveNew As Boolean = True) As CComDonStk
On Error Resume Next
Dim Ei As CComDonStk
Static TempEi As CComDonStk

   Set Ei = m_TempCol(TempKey)
    If Ei Is Nothing And HaveNew Then
                If TempEi Is Nothing Then
                   Set TempEi = New CComDonStk
                End If
      Set GetComDonStk = TempEi
   Else
      Set GetComDonStk = Ei
   End If
End Function
Public Sub LoadSumSaleCustomerStcrd(Optional Cl As Collection = Nothing, Optional FromDate As Date, Optional ToDate As Date, Optional FromCustomerCode As String, Optional ToCustomerCode As String, Optional FromStockCode As String, Optional ToStockCode As String, Optional FromSaleCode As String, Optional ToSaleCode As String)
On Error GoTo ErrorHandler
Dim D As CStcrd
Dim ItemCount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CStcrd
Dim i As Long

   Set D = New CStcrd
   Set Rs = New ADODB.Recordset
   
   D.FROM_DOC_DATE = FromDate
   D.TO_DOC_DATE = ToDate
   D.FROM_CUSTOMER_CODE = FromCustomerCode
   D.TO_CUSTOMER_CODE = ToCustomerCode
   D.FROM_STOCK_CODE = FromStockCode
   D.TO_STOCK_CODE = ToStockCode
   D.FROM_SALE_CODE = FromSaleCode
   D.TO_SALE_CODE = ToSaleCode
   Call D.QueryData(21, Rs, ItemCount)
   
   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   While Not Rs.EOF
      i = i + 1
      Set TempData = New CStcrd
      Call TempData.PopulateFromRS(21, Rs)
      ''debug.print TempData.RECTYP
      If Not (Cl Is Nothing) Then
         Call Cl.Add(TempData, Trim(TempData.CUSCOD & "-" & TempData.STKCOD & "-" & TempData.RECTYP))
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   Set Rs = Nothing
   Set D = Nothing
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub
Public Sub LoadSumSaleCustomerStcrdVac(Optional Cl As Collection = Nothing, Optional FromDate As Date, Optional ToDate As Date, Optional FromCustomerCode As String, Optional ToCustomerCode As String, Optional FromStockCode As String, Optional ToStockCode As String, Optional FromSaleCode As String, Optional ToSaleCode As String)
On Error GoTo ErrorHandler
Dim D As CStcrd
Dim ItemCount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CStcrd
Dim i As Long

   Set D = New CStcrd
   Set Rs = New ADODB.Recordset
   
   D.FROM_DOC_DATE = FromDate
   D.TO_DOC_DATE = ToDate
   D.FROM_CUSTOMER_CODE = FromCustomerCode
   D.TO_CUSTOMER_CODE = ToCustomerCode
   D.FROM_STOCK_CODE = FromStockCode
   D.TO_STOCK_CODE = ToStockCode
   D.FROM_SALE_CODE = FromSaleCode
   D.TO_SALE_CODE = ToSaleCode
   Call D.QueryData(26, Rs, ItemCount)
   
   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   While Not Rs.EOF
      i = i + 1
      Set TempData = New CStcrd
      Call TempData.PopulateFromRS(26, Rs)
      ''debug.print TempData.RECTYP
      If Not (Cl Is Nothing) Then
         Call Cl.Add(TempData, Trim(TempData.CUSCOD) & "-" & Trim(TempData.RECTYP))
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   Set Rs = Nothing
   Set D = Nothing
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub
Public Sub LoadSumSaleCustomerStcrdNonVac(Optional Cl As Collection = Nothing, Optional FromDate As Date, Optional ToDate As Date, Optional FromCustomerCode As String, Optional ToCustomerCode As String, Optional FromStockCode As String, Optional ToStockCode As String, Optional FromSaleCode As String, Optional ToSaleCode As String)
On Error GoTo ErrorHandler
Dim D As CStcrd
Dim ItemCount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CStcrd
Dim i As Long

   Set D = New CStcrd
   Set Rs = New ADODB.Recordset
   
   D.FROM_DOC_DATE = FromDate
   D.TO_DOC_DATE = ToDate
   D.FROM_CUSTOMER_CODE = FromCustomerCode
   D.TO_CUSTOMER_CODE = ToCustomerCode
   D.FROM_STOCK_CODE = "" 'FromStockCode
   D.TO_STOCK_CODE = "" 'ToStockCode
   D.FROM_SALE_CODE = FromSaleCode
   D.TO_SALE_CODE = ToSaleCode
   Call D.QueryData(26, Rs, ItemCount)
   
   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   While Not Rs.EOF
      i = i + 1
      Set TempData = New CStcrd
      Call TempData.PopulateFromRS(26, Rs)
      ''debug.print TempData.RECTYP
      If Not (Cl Is Nothing) Then
         Call Cl.Add(TempData, Trim(TempData.CUSCOD) & "-" & Trim(TempData.RECTYP))
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   Set Rs = Nothing
   Set D = Nothing
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub


Public Sub LoadXlsSetting(Optional D As CXlsEstimateSetting = Nothing)
On Error GoTo ErrorHandler
Dim ItemCount As Long
Dim TempData As CXlsEstimateSetting
Dim Rs As ADODB.Recordset
Dim i As Long

   Set D = New CXlsEstimateSetting
   Set Rs = New ADODB.Recordset
   
'   D.VALID_FROM = FromDocDat
'   D.VALID_TO = ToDocDat
   D.XLS_EST_SET_ID = 1
   Call D.QueryData(Rs, ItemCount)   ' ดึงข้อมูลเพื่อเอาไปเก็บในคอเล็กชั่น

'   If Not (Cl Is Nothing) Then
'      Set Cl = Nothing
'      Set Cl = New Collection
'   End If
'
'   While Not Rs.EOF
'      i = i + 1
'      Set TempData = New CXlsEstimateSetting
      Call D.PopulateFromRS(1, Rs)
'
'      If Not (Cl Is Nothing) Then
'         Call Cl.Add(TempData, Trim(TempData.XLS_EST_SET_ID))            ' KEy  ถ้าใส่เดือนครอบกัน จะ error
'      End If
'
'      Set TempData = Nothing
'      Rs.MoveNext
'   Wend
   
'   Set Rs = Nothing
'   Set D = Nothing
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub

Public Sub LoadXlsFood(Optional Cl As Collection = Nothing)
On Error GoTo ErrorHandler
Dim D As CXlsFood
Dim ItemCount As Long
Dim TempData As CXlsFood
Dim Rs As ADODB.Recordset
Dim i As Long

   Set D = New CXlsFood
   Set Rs = New ADODB.Recordset
   
'   D.VALID_FROM = FromDocDat
'   D.VALID_TO = ToDocDat
'   D.ORDER_TYPE = 1
   Call D.QueryData(1, Rs, ItemCount)   ' ดึงข้อมูลเพื่อเอาไปเก็บในคอเล็กชั่น

   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   
   While Not Rs.EOF
      i = i + 1
      Set TempData = New CXlsFood
      Call TempData.PopulateFromRS(1, Rs)

      If Not (Cl Is Nothing) Then
        ' ''debug.print Trim(TempData.XLS_FOOD_CODE) & "-" & Trim(TempData.XLS_UNIT_NAME)
         Call Cl.Add(TempData, Trim(TempData.XLS_FOOD_CODE) & "-" & Trim(TempData.XLS_UNIT_NAME))            ' KEy  ถ้าใส่เดือนครอบกัน จะ error
         
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   Set Rs = Nothing
   Set D = Nothing
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub

Public Sub LoadXlsSetFarm(Optional Cl As Collection = Nothing)
On Error GoTo ErrorHandler
Dim D As CXlsSetFarm
Dim ItemCount As Long
Dim TempData As CXlsSetFarm
Dim Rs As ADODB.Recordset
Dim i As Long

   Set D = New CXlsSetFarm
   Set Rs = New ADODB.Recordset
   
   Call D.QueryData(1, Rs, ItemCount)   ' ดึงข้อมูลเพื่อเอาไปเก็บในคอเล็กชั่น

   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   
   While Not Rs.EOF
      i = i + 1
      Set TempData = New CXlsSetFarm
      Call TempData.PopulateFromRS(1, Rs)

      If Not (Cl Is Nothing) Then
        ' ''debug.print Trim(TempData.XLS_FOOD_CODE) & "-" & Trim(TempData.XLS_UNIT_NAME)
         Call Cl.Add(TempData, Trim(TempData.MAIN_FARM_NAME) & "-" & Trim(TempData.XLS_UNIT_NAME))            ' KEy  ถ้าใส่เดือนครอบกัน จะ error
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   Set Rs = Nothing
   Set D = Nothing
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub
Public Sub InitPrType(C As ComboBox)
   C.Clear
   
   C.AddItem ("")
   C.ItemData(0) = 0
   
   C.AddItem (MapText("RO สั่งซื้อวัตถุดิบ"))
   C.ItemData(1) = 100

   C.AddItem (MapText("RO สั่งซื้อวัสดุอุปกรณ์"))
   C.ItemData(2) = 101

   C.AddItem (MapText("RO สั่งซื้อ รับเข้าจ่ายออกวัสดุอุปกรณ์"))
   C.ItemData(3) = 102

   C.AddItem (MapText("RO สั่งซื้อทั่วไป"))
   C.ItemData(4) = 103
End Sub
'Public Sub LoadSupplierAddressGroupID(Cl As Collection, Optional SupplierID As Long = -1, Optional AmPhurSearch As String, Optional ProvinceSearch As String)
'On Error GoTo ErrorHandler
'Dim D As CAPMas
'Dim ItemCount As Long
'Dim Rs As ADODB.Recordset
'Dim TempData As CAPMas
'Dim i As Long
'
'   Set D = New CAPMas
'   Set Rs = New ADODB.Recordset
'
'  '''' D.ENTERPRISE_ID = -1
'   D.SUPCOD = SupplierID
'''''   D.A = PatchWildCard(AmPhurSearch)
'''''   D.PROVINCE = PatchWildCard(ProvinceSearch)
'
'   Call D.QueryData5(Rs, ItemCount)
'
'   If Not (Cl Is Nothing) Then
'      Set Cl = Nothing
'      Set Cl = New Collection
'   End If
'   While Not Rs.EOF
'      i = i + 1
'
'      Set TempData = New CAPMas
'      Call TempData.PopulateFromRS5(Rs)
'
'      Set D = GetObject("CAPMas", Cl, Trim(Str(TempData.SUPPLIER_ID)), False)
'      If D Is Nothing Then
'         Set D = New CAPMas
'         D.SUPPLIER_ID = TempData.SUPPLIER_ID
'         Call Cl.Add(D, Trim(Str(TempData.SUPPLIER_ID)))
'      End If
'
'      Call D.collSupAddr.Add(TempData)
'
'      Set TempData = Nothing
'      Rs.MoveNext
'   Wend
'
'   Set Rs = Nothing
'   Set D = Nothing
'   Exit Sub
'
'ErrorHandler:
'   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
'   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
'End Sub
Public Sub LoadProvinceMap(C As ComboBox, Optional Cl As Collection = Nothing)
On Error GoTo ErrorHandler
Dim D As CProvinceMap
Dim ItemCount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CProvinceMap
Dim i As Long
   
   Set D = New CProvinceMap
   Set Rs = New ADODB.Recordset
   
   Call D.QueryData(Rs, ItemCount)
   
   If Not (C Is Nothing) Then
      C.Clear
      i = 0
      C.AddItem ("")
   End If
   
   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   
   While Not Rs.EOF
      i = i + 1
      Set TempData = New CProvinceMap
      Call TempData.PopulateFromRS(1, Rs)

      If Not (C Is Nothing) Then
         C.AddItem (TempData.KEY_MAP)
         C.ItemData(i) = TempData.KEY_ID
      End If

      If Not (Cl Is Nothing) Then
         Call Cl.Add(TempData)
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   Set Rs = Nothing
   Set D = Nothing
   Exit Sub

ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub
Public Function GetProvince(TempAddress As String, ColProvinceMaps As Collection) As String
Dim mProvince As CProvinceMap
   GetProvince = ""
   For Each mProvince In ColProvinceMaps
      If InStr(1, TempAddress, mProvince.KEY_SEARCH) > 0 Then
         GetProvince = mProvince.KEY_MAP
         Exit For
      End If
   Next mProvince
End Function
