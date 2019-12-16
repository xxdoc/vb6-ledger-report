VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#2.0#0"; "THREED20.OCX"
Object = "{E684D8A3-716C-4E59-AA94-7144C04B0074}#1.1#0"; "GridEX20.ocx"
Begin VB.Form frmPromotionPayCustomer 
   BackColor       =   &H80000000&
   ClientHeight    =   9375
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   13620
   Icon            =   "frmPromotionPayCustomer.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   9375
   ScaleWidth      =   13620
   StartUpPosition =   1  'CenterOwner
   WindowState     =   2  'Maximized
   Begin Threed.SSFrame SSFrame1 
      Height          =   9375
      Left            =   0
      TabIndex        =   11
      Top             =   0
      Width           =   13635
      _ExtentX        =   24051
      _ExtentY        =   16536
      _Version        =   131073
      PictureBackgroundStyle=   2
      Begin VB.ComboBox cboOrderBy 
         Height          =   315
         Left            =   2130
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   2610
         Width           =   2985
      End
      Begin VB.ComboBox cboOrderType 
         Height          =   315
         Left            =   6480
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   2610
         Width           =   2955
      End
      Begin Threed.SSPanel pnlHeader 
         Height          =   705
         Left            =   0
         TabIndex        =   12
         Top             =   0
         Width           =   13605
         _ExtentX        =   23998
         _ExtentY        =   1244
         _Version        =   131073
         PictureBackgroundStyle=   2
      End
      Begin GridEX20.GridEX GridEX1 
         Height          =   4935
         Left            =   240
         TabIndex        =   10
         Top             =   3360
         Width           =   11505
         _ExtentX        =   20294
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
         Column(1)       =   "frmPromotionPayCustomer.frx":27A2
         Column(2)       =   "frmPromotionPayCustomer.frx":286A
         FormatStylesCount=   5
         FormatStyle(1)  =   "frmPromotionPayCustomer.frx":290E
         FormatStyle(2)  =   "frmPromotionPayCustomer.frx":2A6A
         FormatStyle(3)  =   "frmPromotionPayCustomer.frx":2B1A
         FormatStyle(4)  =   "frmPromotionPayCustomer.frx":2BCE
         FormatStyle(5)  =   "frmPromotionPayCustomer.frx":2CA6
         ImageCount      =   0
         PrinterProperties=   "frmPromotionPayCustomer.frx":2D5E
      End
      Begin prjLedgerReport.uctlTextBox txtSaleCode 
         Height          =   435
         Left            =   2130
         TabIndex        =   0
         Top             =   840
         Width           =   2895
         _ExtentX        =   5106
         _ExtentY        =   767
      End
      Begin prjLedgerReport.uctlTextBox txtCustomerCode 
         Height          =   435
         Left            =   2130
         TabIndex        =   17
         Top             =   1440
         Width           =   2895
         _ExtentX        =   5106
         _ExtentY        =   767
      End
      Begin prjLedgerReport.uctlTextBox txtStockCode 
         Height          =   435
         Left            =   2130
         TabIndex        =   18
         Top             =   2040
         Width           =   2895
         _ExtentX        =   5106
         _ExtentY        =   767
      End
      Begin prjLedgerReport.uctlDate uctlFromDate 
         Height          =   375
         Left            =   6480
         TabIndex        =   20
         Top             =   840
         Width           =   3855
         _ExtentX        =   6800
         _ExtentY        =   661
      End
      Begin prjLedgerReport.uctlDate uctlToDate 
         Height          =   375
         Left            =   6480
         TabIndex        =   22
         Top             =   1440
         Width           =   3855
         _ExtentX        =   6800
         _ExtentY        =   661
      End
      Begin VB.Label lblToDate 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   375
         Left            =   4800
         TabIndex        =   23
         Top             =   1440
         Width           =   1575
      End
      Begin VB.Label lblFromDate 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   375
         Left            =   4800
         TabIndex        =   21
         Top             =   840
         Width           =   1575
      End
      Begin VB.Label lblStockCode 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   435
         Left            =   600
         TabIndex        =   19
         Top             =   2040
         Width           =   1455
      End
      Begin VB.Label lblCustomerCode 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   435
         Left            =   600
         TabIndex        =   16
         Top             =   1440
         Width           =   1455
      End
      Begin Threed.SSCommand cmdClear 
         Height          =   525
         Left            =   11280
         TabIndex        =   4
         Top             =   1410
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdSearch 
         Height          =   525
         Left            =   11280
         TabIndex        =   3
         Top             =   840
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmPromotionPayCustomer.frx":2F36
         ButtonStyle     =   3
      End
      Begin VB.Label lblOrderBy 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   435
         Left            =   600
         TabIndex        =   15
         Top             =   2670
         Width           =   1455
      End
      Begin VB.Label lblOrderType 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   435
         Left            =   5010
         TabIndex        =   14
         Top             =   2670
         Width           =   1365
      End
      Begin VB.Label lblSaleCode 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   435
         Left            =   240
         TabIndex        =   13
         Top             =   840
         Width           =   1815
      End
      Begin Threed.SSCommand cmdDelete 
         Height          =   525
         Left            =   3540
         TabIndex        =   7
         Top             =   8430
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmPromotionPayCustomer.frx":3250
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdAdd 
         Height          =   525
         Left            =   270
         TabIndex        =   5
         Top             =   8430
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmPromotionPayCustomer.frx":356A
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdEdit 
         Height          =   525
         Left            =   1890
         TabIndex        =   6
         Top             =   8430
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   926
         _Version        =   131073
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdExit 
         Cancel          =   -1  'True
         Height          =   525
         Left            =   10215
         TabIndex        =   9
         Top             =   8430
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   926
         _Version        =   131073
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdOK 
         Height          =   525
         Left            =   8565
         TabIndex        =   8
         Top             =   8430
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmPromotionPayCustomer.frx":3884
         ButtonStyle     =   3
      End
   End
End
Attribute VB_Name = "frmPromotionPayCustomer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private m_HasActivate As Boolean
Private m_PromotionPayCustom As CPromotionPayCustom
Private m_TempPromotionPayCustom As CPromotionPayCustom
Private m_Rs As ADODB.Recordset

Public OKClick As Boolean
Public HeaderText As String

Private Sub cmdAdd_Click()
Dim ItemCount As Long
Dim OKClick As Boolean
   
   frmAddEditPromotionPay.HeaderText = MapText("เพิ่มข้อมูล Promotion")
   frmAddEditPromotionPay.ShowMode = SHOW_ADD
   Load frmAddEditPromotionPay
   frmAddEditPromotionPay.Show 1
   
   OKClick = frmAddEditPromotionPay.OKClick
   
   Unload frmAddEditPromotionPay
   Set frmAddEditPromotionPay = Nothing
   
   If OKClick Then
      Call QueryData(True)
   End If
End Sub

Private Sub cmdClear_Click()
   txtSaleCode.Text = ""
   txtCustomerCode.Text = ""
   txtStockCode.Text = ""
   uctlFromDate.ShowDate = -1
   uctlToDate.ShowDate = -1
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

   If Not ConfirmDelete("รหัสSale " & GridEX1.Value(3) & " ของวันที่ " & GridEX1.Value(2)) Then
      Exit Sub
   End If

   Call EnableForm(Me, False)
   m_PromotionPayCustom.PRO_ID = ID
   If Not glbDaily.DeletePromotionPayCustomer(m_PromotionPayCustom, IsOK, True, glbErrorLog) Then
      m_PromotionPayCustom.PRO_ID = -1
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

   frmAddEditPromotionPay.ID = ID
   frmAddEditPromotionPay.HeaderText = MapText("แก้ไข้อมูล Promotion")
   frmAddEditPromotionPay.ShowMode = SHOW_EDIT
   Load frmAddEditPromotionPay
   frmAddEditPromotionPay.Show 1

   OKClick = frmAddEditPromotionPay.OKClick

   Unload frmAddEditPromotionPay
   Set frmAddEditPromotionPay = Nothing

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
      
      Call InitPromotionPayCustomerOrderBy(cboOrderBy)
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
      
      m_PromotionPayCustom.PRO_ID = -1
      m_PromotionPayCustom.SALECODE_PRO = txtSaleCode.Text
      m_PromotionPayCustom.CUSTOMERCODE_PRO = txtCustomerCode.Text
      m_PromotionPayCustom.STKCOD_PRO = txtStockCode.Text
      m_PromotionPayCustom.FROM_PRO_DATE = uctlFromDate.ShowDate
      m_PromotionPayCustom.TO_PRO_DATE = uctlToDate.ShowDate

      m_PromotionPayCustom.OrderBy = cboOrderBy.ItemData(Minus2Zero(cboOrderBy.ListIndex))
      m_PromotionPayCustom.OrderType = cboOrderType.ItemData(Minus2Zero(cboOrderType.ListIndex))
      
      If Not glbDaily.QueryPromotionPayCustomer(m_PromotionPayCustom, m_Rs, ItemCount, IsOK, glbErrorLog) Then
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
   Col.Width = 1500
   Col.Caption = MapText("รหัส Sale")
   
   Set Col = GridEX1.Columns.Add '4
   Col.Width = 3500
   Col.Caption = MapText("Sale")
   
   Set Col = GridEX1.Columns.Add '5
   Col.Width = 1500
   Col.Caption = MapText("รหัสลูกค้า")
   
   Set Col = GridEX1.Columns.Add '6
   Col.Width = 3500
   Col.Caption = MapText("ชื่อลูกค้า")
   
   Set Col = GridEX1.Columns.Add '7
   Col.Width = 1500
   Col.Caption = MapText("รหัสสินค้า")
   
   Set Col = GridEX1.Columns.Add '8
   Col.Width = 3500
   Col.Caption = MapText("สินค้า")
   
   Set Col = GridEX1.Columns.Add '9
   Col.Width = 2500
   Col.Caption = MapText("ยอดส่วนลด")
   
   GridEX1.ItemCount = 0
End Sub
Private Sub InitFormLayout()
   pnlHeader.PictureBackground = LoadPicture(glbParameterObj.NormalForm1)
   SSFrame1.PictureBackground = LoadPicture(glbParameterObj.NormalForm1)
   
   Me.Caption = MapText("Promotion")
   pnlHeader.Caption = Me.Caption
   
   Call InitGrid
   
   Call InitNormalLabel(lblSaleCode, MapText("รหัสพนักงานขาย"))
   Call InitNormalLabel(lblCustomerCode, MapText("รหัสลูกค้า"))
   Call InitNormalLabel(lblStockCode, MapText("รหัสสินค้า"))
   Call InitNormalLabel(lblFromDate, MapText("วันที่"))
   Call InitNormalLabel(lblToDate, MapText("ถึงวันที่"))
   
   Call InitNormalLabel(lblOrderBy, MapText("เรียงตาม"))
   Call InitNormalLabel(lblOrderType, MapText("เรียงจาก"))
   
   Me.Picture = LoadPicture(glbParameterObj.MainPicture)
   
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
   
   Set m_PromotionPayCustom = New CPromotionPayCustom
   Set m_TempPromotionPayCustom = New CPromotionPayCustom
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
   Call m_TempPromotionPayCustom.PopulateFromRS(1, m_Rs)
   
   Values(1) = m_TempPromotionPayCustom.PRO_ID
   Values(2) = m_TempPromotionPayCustom.PRO_DATE
   Values(3) = m_TempPromotionPayCustom.SALECODE_PRO
   Values(4) = m_TempPromotionPayCustom.SALENAME_PRO
   Values(5) = m_TempPromotionPayCustom.CUSTOMERCODE_PRO
   Values(6) = m_TempPromotionPayCustom.CUSTOMERNAME_PRO
   Values(7) = m_TempPromotionPayCustom.STKCOD_PRO
   Values(8) = m_TempPromotionPayCustom.STKNAME_PRO
   Values(9) = m_TempPromotionPayCustom.AMOUNT_PRO

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
