VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#2.0#0"; "THREED20.OCX"
Object = "{E684D8A3-716C-4E59-AA94-7144C04B0074}#1.1#0"; "GridEX20.ocx"
Begin VB.Form frmBudgetTicket 
   BackColor       =   &H80000000&
   ClientHeight    =   8520
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   12960
   Icon            =   "frmBudgetTicket.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   8520
   ScaleWidth      =   12960
   StartUpPosition =   1  'CenterOwner
   WindowState     =   2  'Maximized
   Begin Threed.SSFrame SSFrame1 
      Height          =   8535
      Left            =   0
      TabIndex        =   13
      Top             =   0
      Width           =   13035
      _ExtentX        =   22992
      _ExtentY        =   15055
      _Version        =   131073
      PictureBackgroundStyle=   2
      Begin VB.ComboBox cboBankName 
         Height          =   315
         Left            =   6960
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   1440
         Width           =   3015
      End
      Begin VB.ComboBox cboCustomerName 
         Height          =   315
         Left            =   1920
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   1440
         Width           =   3015
      End
      Begin prjLedgerReport.uctlDate uctlPutToDate 
         Height          =   375
         Left            =   6960
         TabIndex        =   1
         Top             =   840
         Width           =   3855
         _ExtentX        =   6800
         _ExtentY        =   661
      End
      Begin prjLedgerReport.uctlDate uctlPutFromDate 
         Height          =   375
         Left            =   1920
         TabIndex        =   0
         Top             =   840
         Width           =   3855
         _ExtentX        =   6800
         _ExtentY        =   661
      End
      Begin VB.ComboBox cboOrderType 
         Height          =   315
         Left            =   6960
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   1920
         Width           =   2955
      End
      Begin Threed.SSPanel pnlHeader 
         Height          =   705
         Left            =   0
         TabIndex        =   14
         Top             =   0
         Width           =   13005
         _ExtentX        =   22939
         _ExtentY        =   1244
         _Version        =   131073
         PictureBackgroundStyle=   2
      End
      Begin GridEX20.GridEX GridEX1 
         Height          =   5175
         Left            =   240
         TabIndex        =   12
         Top             =   2400
         Width           =   12465
         _ExtentX        =   21987
         _ExtentY        =   9128
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
         Column(1)       =   "frmBudgetTicket.frx":27A2
         Column(2)       =   "frmBudgetTicket.frx":286A
         FormatStylesCount=   5
         FormatStyle(1)  =   "frmBudgetTicket.frx":290E
         FormatStyle(2)  =   "frmBudgetTicket.frx":2A6A
         FormatStyle(3)  =   "frmBudgetTicket.frx":2B1A
         FormatStyle(4)  =   "frmBudgetTicket.frx":2BCE
         FormatStyle(5)  =   "frmBudgetTicket.frx":2CA6
         ImageCount      =   0
         PrinterProperties=   "frmBudgetTicket.frx":2D5E
      End
      Begin VB.Label lblBankName 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   255
         Left            =   5040
         TabIndex        =   19
         Top             =   1440
         Width           =   1695
      End
      Begin VB.Label lblCustomerName 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   375
         Left            =   240
         TabIndex        =   18
         Top             =   1440
         Width           =   1455
      End
      Begin VB.Label lblPutToDate 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   375
         Left            =   5640
         TabIndex        =   17
         Top             =   840
         Width           =   1095
      End
      Begin VB.Label lblPutFromDate 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   435
         Left            =   480
         TabIndex        =   16
         Top             =   840
         Width           =   1215
      End
      Begin Threed.SSCommand cmdClear 
         Height          =   525
         Left            =   11040
         TabIndex        =   6
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
         TabIndex        =   5
         Top             =   960
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmBudgetTicket.frx":2F36
         ButtonStyle     =   3
      End
      Begin VB.Label lblOrderType 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   435
         Left            =   5400
         TabIndex        =   15
         Top             =   1920
         Width           =   1365
      End
      Begin Threed.SSCommand cmdDelete 
         Height          =   525
         Left            =   3420
         TabIndex        =   9
         Top             =   7830
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmBudgetTicket.frx":3250
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdAdd 
         Height          =   525
         Left            =   150
         TabIndex        =   7
         Top             =   7830
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmBudgetTicket.frx":356A
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdEdit 
         Height          =   525
         Left            =   1770
         TabIndex        =   8
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
         TabIndex        =   11
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
         TabIndex        =   10
         Top             =   7830
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmBudgetTicket.frx":3884
         ButtonStyle     =   3
      End
   End
End
Attribute VB_Name = "frmBudgetTicket"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private m_HasActivate As Boolean
Private m_Ticket As CTicket
Private m_TempTicket As CTicket
Private m_Rs As ADODB.Recordset
Public OKClick As Boolean
Private Sub cmdAdd_Click()
Dim ItemCount As Long
Dim OKClick As Boolean
   
   frmAddEditBudgetTicket.HeaderText = MapText("เพิ่มข้อมูลประมาณการตั๋ว")
   frmAddEditBudgetTicket.ShowMode = SHOW_ADD
   Load frmAddEditBudgetTicket
   frmAddEditBudgetTicket.Show 1
   
   OKClick = frmAddEditBudgetTicket.OKClick
   
   Unload frmAddEditBudgetTicket
   Set frmAddEditBudgetTicket = Nothing
   
   If OKClick Then
      Call QueryData(True)
   End If
End Sub
Private Sub cmdClear_Click()
   cboCustomerName.ListIndex = -1
   cboBankName.ListIndex = -1
   uctlPutFromDate.ShowDate = -1
   uctlPutToDate.ShowDate = -1
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
   m_Ticket.TICKET_ID = ID
   If Not glbDaily.DeleteTicket(m_Ticket, IsOK, True, glbErrorLog) Then
      m_Ticket.TICKET_ID = -1
      glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
      m_Ticket.TICKET_NUMBER = ""
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
   
   frmAddEditBudgetTicket.ID = ID
   frmAddEditBudgetTicket.HeaderText = MapText("แก้ไขข้อมูลประมาณการตั๋ว")
   frmAddEditBudgetTicket.ShowMode = SHOW_EDIT
   Load frmAddEditBudgetTicket
   frmAddEditBudgetTicket.Show 1
   
   OKClick = frmAddEditBudgetTicket.OKClick
   
   Unload frmAddEditBudgetTicket
   Set frmAddEditBudgetTicket = Nothing
               
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
      m_Ticket.TICKET_ID = -1
      m_Ticket.CUSTOMER_ID = cboCustomerName.ItemData(Minus2Zero(cboCustomerName.ListIndex))
      m_Ticket.FROM_DATE = uctlPutFromDate.ShowDate
      m_Ticket.TO_DATE = uctlPutToDate.ShowDate
      m_Ticket.ORDER_TYPE = cboOrderType.ItemData(Minus2Zero(cboOrderType.ListIndex))
      m_Ticket.BANK_ID = cboBankName.ItemData(Minus2Zero(cboBankName.ListIndex))
      m_Ticket.MASTER_AREA = 2
      If m_Ticket.ORDER_TYPE = 0 Then
            m_Ticket.ORDER_TYPE = 2
      End If
      
      If Not glbDaily.QueryTicket(m_Ticket, m_Rs, ItemCount, IsOK, glbErrorLog) Then
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
   Col.Caption = MapText("วันที่")
   
   Set Col = GridEX1.Columns.Add '3
   Col.Width = 2500
   Col.Caption = MapText("จำนวนเงิน")
   Col.TextAlignment = jgexAlignRight

   Set Col = GridEX1.Columns.Add '4
   Col.Width = 2000
   Col.Caption = MapText("วันที่หน้า เช็ค")
   
   Set Col = GridEX1.Columns.Add '5
   Col.Width = 4000
   Col.Caption = MapText("ลูกหนี้")
   
   Set Col = GridEX1.Columns.Add '6
   Col.Width = 2500
   Col.Caption = MapText("ธนาคาร")
   
   GridEX1.ItemCount = 0
End Sub
Private Sub InitFormLayout()
   pnlHeader.PictureBackground = LoadPicture(glbParameterObj.NormalForm1)
   SSFrame1.PictureBackground = LoadPicture(glbParameterObj.NormalForm1)
   
   Me.Caption = MapText("ข้อมูลประมาณการตั๋ว")
   pnlHeader.Caption = Me.Caption
   
   Call InitGrid

   Call InitNormalLabel(lblBankName, MapText("ธนาคาร"))
   Call InitNormalLabel(lblCustomerName, MapText("ชื่อลูกหนี้"))
   Call InitNormalLabel(lblPutFromDate, MapText("จากวันที่"))
   Call InitNormalLabel(lblPutToDate, MapText("ถึงวันที่"))
   Call InitNormalLabel(lblOrderType, MapText("เรียงจาก"))
   Me.Picture = LoadPicture(glbParameterObj.MainPicture)
   
   Call InitCombo(cboBankName)
   Call InitCombo(cboCustomerName)
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
   
   Set m_Ticket = New CTicket
   Set m_TempTicket = New CTicket
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
   Call m_TempTicket.PopulateFromRS(1, m_Rs)
   
   Values(1) = m_TempTicket.TICKET_ID
   Values(2) = DateToStringExtEx2(m_TempTicket.TICKET_DATE)
   Values(3) = FormatNumber(m_TempTicket.TICKET_AMOUNT)
   Values(4) = DateToStringExtEx2(m_TempTicket.TICKET_DATE_CHECK)
   Values(5) = m_TempTicket.CUSTOMER_NAME
   Values(6) = m_TempTicket.BANK_NAME
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
