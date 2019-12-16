VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#2.0#0"; "THREED20.OCX"
Object = "{E684D8A3-716C-4E59-AA94-7144C04B0074}#1.1#0"; "GridEX20.ocx"
Begin VB.Form frmBankCredit 
   BackColor       =   &H80000000&
   ClientHeight    =   8520
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11910
   Icon            =   "frmBankCredit.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   8520
   ScaleWidth      =   11910
   StartUpPosition =   1  'CenterOwner
   WindowState     =   2  'Maximized
   Begin Threed.SSFrame SSFrame1 
      Height          =   8535
      Left            =   0
      TabIndex        =   14
      Top             =   0
      Width           =   13035
      _ExtentX        =   22992
      _ExtentY        =   15055
      _Version        =   131073
      PictureBackgroundStyle=   2
      Begin VB.ComboBox cboCustomerName 
         Height          =   315
         Left            =   6960
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   840
         Width           =   3015
      End
      Begin VB.ComboBox cboBankName 
         Height          =   315
         Left            =   1920
         Style           =   2  'Dropdown List
         TabIndex        =   0
         Top             =   840
         Width           =   2895
      End
      Begin prjLedgerReport.uctlDate uctlPutToDate 
         Height          =   375
         Left            =   6960
         TabIndex        =   3
         Top             =   1440
         Width           =   3855
         _ExtentX        =   6800
         _ExtentY        =   661
      End
      Begin prjLedgerReport.uctlDate uctlPutFromDate 
         Height          =   375
         Left            =   1920
         TabIndex        =   2
         Top             =   1440
         Width           =   3855
         _ExtentX        =   6800
         _ExtentY        =   661
      End
      Begin VB.ComboBox cboOrderBy 
         Height          =   315
         Left            =   1920
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   2040
         Width           =   2955
      End
      Begin VB.ComboBox cboOrderType 
         Height          =   315
         Left            =   6960
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   2040
         Width           =   2955
      End
      Begin Threed.SSPanel pnlHeader 
         Height          =   705
         Left            =   0
         TabIndex        =   15
         Top             =   0
         Width           =   11925
         _ExtentX        =   21034
         _ExtentY        =   1244
         _Version        =   131073
         PictureBackgroundStyle=   2
      End
      Begin GridEX20.GridEX GridEX1 
         Height          =   4935
         Left            =   240
         TabIndex        =   13
         Top             =   2640
         Width           =   12465
         _ExtentX        =   21987
         _ExtentY        =   8705
         Version         =   "2.0"
         BoundColumnIndex=   ""
         ReplaceColumnIndex=   ""
         TabKeyBehavior  =   1
         MethodHoldFields=   -1  'True
         AllowColumnDrag =   0   'False
         AllowEdit       =   0   'False
         BorderStyle     =   3
         GroupByBoxVisible=   0   'False
         DataMode        =   99
         HeaderFontName  =   "AngsanaUPC"
         HeaderFontBold  =   -1  'True
         HeaderFontSize  =   14.25
         HeaderFontWeight=   700
         FontSize        =   9.75
         BackColorBkg    =   16777215
         ColumnHeaderHeight=   480
         IntProp1        =   0
         IntProp2        =   0
         IntProp7        =   0
         ColumnsCount    =   2
         Column(1)       =   "frmBankCredit.frx":27A2
         Column(2)       =   "frmBankCredit.frx":286A
         FormatStylesCount=   5
         FormatStyle(1)  =   "frmBankCredit.frx":290E
         FormatStyle(2)  =   "frmBankCredit.frx":2A6A
         FormatStyle(3)  =   "frmBankCredit.frx":2B1A
         FormatStyle(4)  =   "frmBankCredit.frx":2BCE
         FormatStyle(5)  =   "frmBankCredit.frx":2CA6
         ImageCount      =   0
         PrinterProperties=   "frmBankCredit.frx":2D5E
      End
      Begin VB.Label lblCustomerName 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   375
         Left            =   5400
         TabIndex        =   21
         Top             =   840
         Width           =   1335
      End
      Begin VB.Label lblPutToDate 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   375
         Left            =   5640
         TabIndex        =   20
         Top             =   1440
         Width           =   1095
      End
      Begin VB.Label lblPutFromDate 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   435
         Left            =   480
         TabIndex        =   19
         Top             =   1440
         Width           =   1215
      End
      Begin Threed.SSCommand cmdClear 
         Height          =   525
         Left            =   11040
         TabIndex        =   7
         Top             =   1680
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdSearch 
         Height          =   525
         Left            =   11040
         TabIndex        =   6
         Top             =   1080
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmBankCredit.frx":2F36
         ButtonStyle     =   3
      End
      Begin VB.Label lblOrderBy 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   435
         Left            =   600
         TabIndex        =   18
         Top             =   2040
         Width           =   1095
      End
      Begin VB.Label lblOrderType 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   435
         Left            =   5400
         TabIndex        =   17
         Top             =   2040
         Width           =   1365
      End
      Begin VB.Label lblBankName 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   435
         Left            =   120
         TabIndex        =   16
         Top             =   840
         Width           =   1575
      End
      Begin Threed.SSCommand cmdDelete 
         Height          =   525
         Left            =   3420
         TabIndex        =   10
         Top             =   7830
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmBankCredit.frx":3250
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdAdd 
         Height          =   525
         Left            =   150
         TabIndex        =   8
         Top             =   7830
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmBankCredit.frx":356A
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdEdit 
         Height          =   525
         Left            =   1770
         TabIndex        =   9
         Top             =   7830
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   926
         _Version        =   131073
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdExit 
         Cancel          =   -1  'True
         Height          =   525
         Left            =   11040
         TabIndex        =   12
         Top             =   7830
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   926
         _Version        =   131073
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdOK 
         Height          =   525
         Left            =   9240
         TabIndex        =   11
         Top             =   7830
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmBankCredit.frx":3884
         ButtonStyle     =   3
      End
   End
End
Attribute VB_Name = "frmBankCredit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private m_HasActivate As Boolean
Private m_BankCredit As CBankCredit
Private m_TempBankCredit As CBankCredit
Private m_Rs As ADODB.Recordset
Public OKClick As Boolean
Private Sub cmdAdd_Click()
Dim ItemCount As Long
Dim OKClick As Boolean
   
   frmAddEditBankCredit.HeaderText = MapText("เพิ่มข้อมูลวงเงินธนาคาร")
   frmAddEditBankCredit.ShowMode = SHOW_ADD
   Load frmAddEditBankCredit
   frmAddEditBankCredit.Show 1
   
   OKClick = frmAddEditBankCredit.OKClick
   
   Unload frmAddEditBankCredit
   Set frmAddEditBankCredit = Nothing
   
   If OKClick Then
      Call QueryData(True)
   End If
End Sub
Private Sub cmdClear_Click()
   cboBankName.ListIndex = -1
   cboCustomerName = -1
   uctlPutFromDate.ShowDate = -1
   uctlPutToDate.ShowDate = -1
   cboOrderBy.ListIndex = -1
   cboOrderType.ListIndex = -1
End Sub
Private Sub cmdDelete_Click()
Dim IsOK As Boolean
Dim ItemCount As Long
Dim IsCanLock As Boolean
Dim ID As Long

   If Not VerifyGrid(GridEX1.Value(1)) Then
      Exit Sub
   End If
   
   ID = GridEX1.Value(1)
   
   If Not ConfirmDelete(GridEX1.Value(2) & " " & GridEX1.Value(3)) Then
      Exit Sub
   End If
   
   Call EnableForm(Me, False)
   m_BankCredit.BANK_NUMBER = ID
   If Not glbDaily.DeleteBankCredit(m_BankCredit, IsOK, True, glbErrorLog) Then
      m_BankCredit.BANK_NUMBER = -1
      glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
      Call EnableForm(Me, True)
      Exit Sub
   End If
   
   Call QueryData(True)
   Call EnableForm(Me, True)
End Sub
Private Sub cmdEdit_Click()
Dim IsOK As Boolean
Dim ItemCount As Long
Dim IsCanLock As Boolean
Dim ID As Long
Dim OKClick As Boolean
   
   If Not VerifyGrid(GridEX1.Value(1)) Then
      Exit Sub
   End If
   
   ID = GridEX1.Value(1)
   
   frmAddEditBankCredit.ID = ID
   frmAddEditBankCredit.HeaderText = MapText("แก้ไข้อมูลวงเงินธนาคาร")
   frmAddEditBankCredit.ShowMode = SHOW_EDIT
   Load frmAddEditBankCredit
   frmAddEditBankCredit.Show 1
   
   OKClick = frmAddEditBankCredit.OKClick
   
   Unload frmAddEditBankCredit
   Set frmAddEditBankCredit = Nothing
               
   If OKClick Then
      Call QueryData(True)
   End If

End Sub

Private Sub cmdOK_Click()
   OKClick = True
   Unload Me
End Sub

Private Sub cmdSearch_Click()
   Call QueryData(True)
End Sub

Private Sub Form_Activate()
   If Not m_HasActivate Then
      m_HasActivate = True
      
      Call LoadBank(cboBankName)
      Call LoadBankCustomer(cboCustomerName)
      Call InitBankCreditOrderBy(cboOrderBy)
      Call InitOrderType(cboOrderType)

      Call QueryData(True)
   End If
End Sub
Private Sub QueryData(Flag As Boolean)
Dim IsOK As Boolean
Dim ItemCount As Long
Dim Temp As Long

   If Flag Then
      Call EnableForm(Me, False)
      m_BankCredit.BANK_NUMBER = -1
      m_BankCredit.BANK_ID = cboBankName.ItemData(Minus2Zero(cboBankName.ListIndex))
      m_BankCredit.CUSTOMER_ID = cboCustomerName.ItemData(Minus2Zero(cboCustomerName.ListIndex))
      m_BankCredit.FROM_DATE = uctlPutFromDate.ShowDate
      m_BankCredit.TO_DATE = uctlPutToDate.ShowDate
      m_BankCredit.ORDER_BY = cboOrderBy.ItemData(Minus2Zero(cboOrderBy.ListIndex))
      m_BankCredit.ORDER_TYPE = cboOrderType.ItemData(Minus2Zero(cboOrderType.ListIndex))
      If m_BankCredit.ORDER_BY = 0 And m_BankCredit.ORDER_TYPE = 0 Then
            m_BankCredit.ORDER_BY = 1
            m_BankCredit.ORDER_TYPE = 2
      End If
      If Not glbDaily.QueryBankCredit(m_BankCredit, m_Rs, ItemCount, IsOK, glbErrorLog) Then
         glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
         Call EnableForm(Me, True)
         Exit Sub
      End If
   End If
   
   If Not IsOK Then
      glbErrorLog.ShowUserError
      Call EnableForm(Me, True)
      Exit Sub
   End If
   
   GridEX1.ItemCount = ItemCount
   GridEX1.Rebind
   
   Call EnableForm(Me, True)
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
   If Shift = 1 And KeyCode = DUMMY_KEY Then
      glbErrorLog.LocalErrorMsg = Me.Name
      glbErrorLog.ShowUserError
      KeyCode = 0
   ElseIf Shift = 0 And KeyCode = 116 Then
'      Call cmdSearch_Click
      KeyCode = 0
   ElseIf Shift = 0 And KeyCode = 115 Then
'      Call cmdClear_Click
      KeyCode = 0
   ElseIf Shift = 0 And KeyCode = 118 Then
      Call cmdAdd_Click
      KeyCode = 0
   ElseIf Shift = 0 And KeyCode = 117 Then
      Call cmdDelete_Click
      KeyCode = 0
   ElseIf Shift = 0 And KeyCode = 113 Then
      Call cmdOK_Click
      KeyCode = 0
   ElseIf Shift = 0 And KeyCode = 114 Then
      Call cmdEdit_Click
      KeyCode = 0
   ElseIf Shift = 0 And KeyCode = 121 Then
'      Call cmdPrint_Click
      KeyCode = 0
   End If
End Sub
Private Sub InitGrid()
Dim Col As JSColumn
Dim fmsTemp As JSFormatStyle

   GridEX1.Columns.Clear
   GridEX1.BackColor = GLB_GRID_COLOR
   GridEX1.BackColorHeader = GLB_GRIDHD_COLOR
   
   GridEX1.FormatStyles.Clear
   Set fmsTemp = GridEX1.FormatStyles.Add("N")
   fmsTemp.ForeColor = GLB_ALERT_COLOR
   
  Set Col = GridEX1.Columns.Add '1
  Col.Width = 0
  Col.Caption = MapText("ID")
   
   Set Col = GridEX1.Columns.Add '2
   Col.Width = 2000
   Col.Caption = MapText("ชื่อธนาคาร")
   
   Set Col = GridEX1.Columns.Add '3
   Col.Width = 2000
    Col.Caption = MapText("วงเงิน")
    Col.TextAlignment = jgexAlignRight
   
   Set Col = GridEX1.Columns.Add '4
   Col.Width = 1200
   Col.Caption = MapText("ดอกเบี้ย  (%)")
   Col.TextAlignment = jgexAlignRight
   
   Set Col = GridEX1.Columns.Add '5
   Col.Width = 1200
   Col.Caption = MapText("ค่าธรรมเนียม")
   Col.TextAlignment = jgexAlignRight
   
   Set Col = GridEX1.Columns.Add '6
   Col.Width = 1300
   Col.Caption = MapText("วันที่ยกมา")
   
   Set Col = GridEX1.Columns.Add '7
   Col.Width = 2200
   Col.Caption = MapText("ยอดเงินยกมา(หนี้คงเหลือ)")
   Col.TextAlignment = jgexAlignRight
   
   Set Col = GridEX1.Columns.Add '8
   Col.Width = 1800
   Col.Caption = MapText("จำนวนเงินรับมา  (%)")
   Col.TextAlignment = jgexAlignRight
   
   Set Col = GridEX1.Columns.Add '9
   Col.Width = 2500
   Col.Caption = MapText("ลูกหนี้")
   
   GridEX1.ItemCount = 0
End Sub
Private Sub InitFormLayout()
   pnlHeader.PictureBackground = LoadPicture(glbParameterObj.NormalForm1)
   SSFrame1.PictureBackground = LoadPicture(glbParameterObj.NormalForm1)
   
   Me.Caption = MapText("ข้อมูลวงเงินธนาคาร")
   pnlHeader.Caption = Me.Caption
   
   Call InitGrid
   
   Call InitNormalLabel(lblBankName, MapText("ชื่อธนาคาร"))
   Call InitNormalLabel(lblCustomerName, MapText("ชื่อลูกหนี้"))
   Call InitNormalLabel(lblPutFromDate, MapText("จากวันที่"))
   Call InitNormalLabel(lblPutToDate, MapText("ถึงวันที่"))
   Call InitNormalLabel(lblOrderBy, MapText("เรียงตาม"))
   Call InitNormalLabel(lblOrderType, MapText("เรียงจาก"))

   Me.Picture = LoadPicture(glbParameterObj.MainPicture)
   
   Call InitCombo(cboBankName)
   Call InitCombo(cboCustomerName)
   Call InitCombo(cboOrderBy)
   Call InitCombo(cboOrderType)
   
   pnlHeader.Font.Name = GLB_FONT
   pnlHeader.Font.Bold = True
   pnlHeader.Font.Size = 19
   
   cmdExit.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdOK.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdSearch.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdClear.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdAdd.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdEdit.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdDelete.Picture = LoadPicture(glbParameterObj.NormalButton1)
   
   Call InitMainButton(cmdExit, MapText("ยกเลิก (ESC)"))
   Call InitMainButton(cmdOK, MapText("ตกลง (F2)"))
   Call InitMainButton(cmdAdd, MapText("เพิ่ม (F7)"))
   Call InitMainButton(cmdEdit, MapText("แก้ไข (F3)"))
   Call InitMainButton(cmdDelete, MapText("ลบ (F6)"))
   Call InitMainButton(cmdSearch, MapText("ค้นหา (F5)"))
   Call InitMainButton(cmdClear, MapText("เคลียร์ (F4)"))
   
End Sub
Private Sub cmdExit_Click()
   OKClick = False
   Unload Me
End Sub

Private Sub Form_Load()
   
   Set m_BankCredit = New CBankCredit
   Set m_TempBankCredit = New CBankCredit
   Set m_Rs = New ADODB.Recordset

   Call InitFormLayout
   Call EnableForm(Me, True)
End Sub
Private Sub GridEX1_DblClick()
   Call cmdEdit_Click
End Sub
Private Sub GridEX1_RowFormat(RowBuffer As GridEX20.JSRowData)
   RowBuffer.RowStyle = RowBuffer.Value(2)
End Sub
Private Sub GridEX1_UnboundReadData(ByVal RowIndex As Long, ByVal Bookmark As Variant, ByVal Values As GridEX20.JSRowData)
On Error GoTo ErrorHandler
Dim RealIndex As Long

   glbErrorLog.ModuleName = Me.Name
   glbErrorLog.RoutineName = "UnboundReadData"

   If m_Rs Is Nothing Then
      Exit Sub
   End If

   If m_Rs.State <> adStateOpen Then
      Exit Sub
   End If

   If m_Rs.EOF Then
      Exit Sub
   End If

   If RowIndex <= 0 Then
      Exit Sub
   End If
   
   Call m_Rs.Move(RowIndex - 1, adBookmarkFirst)
   Call m_TempBankCredit.PopulateFromRS(1, m_Rs)
   
   Values(1) = m_TempBankCredit.BANK_NUMBER
   Values(2) = m_TempBankCredit.BANK_NAME
   Values(3) = FormatNumber(m_TempBankCredit.BANK_AMOUNT)
   Values(4) = m_TempBankCredit.BANK_INTEREST
   Values(5) = m_TempBankCredit.BANK_FEE_AMOUNT
   Values(6) = DateToStringExtEx2(m_TempBankCredit.BANK_DATE_BROUGHT)
   Values(7) = FormatNumber(m_TempBankCredit.BANK_AMOUNT_BROUGHT)
   Values(8) = m_TempBankCredit.BANK_GET_AMOUNT
   Values(9) = m_TempBankCredit.CUSTOMER_NAME
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub
Private Sub Form_Resize()
On Error Resume Next
   SSFrame1.Width = ScaleWidth
   SSFrame1.Height = ScaleHeight
   pnlHeader.Width = ScaleWidth
   GridEX1.Width = ScaleWidth - 2 * GridEX1.Left
   GridEX1.Height = ScaleHeight - GridEX1.Top - 620
   cmdAdd.Top = ScaleHeight - 580
   cmdEdit.Top = ScaleHeight - 580
   cmdDelete.Top = ScaleHeight - 580
   cmdOK.Top = ScaleHeight - 580
   cmdExit.Top = ScaleHeight - 580
   cmdExit.Left = ScaleWidth - cmdExit.Width - 50
   cmdOK.Left = cmdExit.Left - cmdOK.Width - 50
End Sub
Private Sub Label1_Click()

End Sub
