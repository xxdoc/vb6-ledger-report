VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#2.0#0"; "THREED20.OCX"
Object = "{E684D8A3-716C-4E59-AA94-7144C04B0074}#1.1#0"; "GridEX20.ocx"
Begin VB.Form frmComMinusStk 
   BackColor       =   &H80000000&
   ClientHeight    =   8520
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   13035
   Icon            =   "frmComMinusStk.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   8520
   ScaleWidth      =   13035
   StartUpPosition =   1  'CenterOwner
   WindowState     =   2  'Maximized
   Begin Threed.SSFrame SSFrame1 
      Height          =   8535
      Left            =   0
      TabIndex        =   12
      Top             =   0
      Width           =   13035
      _ExtentX        =   22992
      _ExtentY        =   15055
      _Version        =   131073
      PictureBackgroundStyle=   2
      Begin prjLedgerReport.uctlTextBox txtIVNo 
         Height          =   375
         Left            =   2160
         TabIndex        =   0
         Top             =   1920
         Width           =   3855
         _ExtentX        =   6800
         _ExtentY        =   661
      End
      Begin VB.ComboBox cboOrderType 
         Height          =   315
         Left            =   7320
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   2400
         Width           =   3435
      End
      Begin VB.ComboBox cboOrderBy 
         Height          =   315
         Left            =   2160
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   2400
         Width           =   3795
      End
      Begin Threed.SSPanel pnlHeader 
         Height          =   705
         Left            =   0
         TabIndex        =   13
         Top             =   0
         Width           =   13725
         _ExtentX        =   24209
         _ExtentY        =   1244
         _Version        =   131073
         PictureBackgroundStyle=   2
      End
      Begin GridEX20.GridEX GridEX1 
         Height          =   4215
         Left            =   180
         TabIndex        =   6
         Top             =   2880
         Width           =   12465
         _ExtentX        =   21987
         _ExtentY        =   7435
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
         Column(1)       =   "frmComMinusStk.frx":27A2
         Column(2)       =   "frmComMinusStk.frx":286A
         FormatStylesCount=   5
         FormatStyle(1)  =   "frmComMinusStk.frx":290E
         FormatStyle(2)  =   "frmComMinusStk.frx":2A6A
         FormatStyle(3)  =   "frmComMinusStk.frx":2B1A
         FormatStyle(4)  =   "frmComMinusStk.frx":2BCE
         FormatStyle(5)  =   "frmComMinusStk.frx":2CA6
         ImageCount      =   0
         PrinterProperties=   "frmComMinusStk.frx":2D5E
      End
      Begin prjLedgerReport.uctlTextBox txtStkCodNo 
         Height          =   375
         Left            =   7320
         TabIndex        =   1
         Top             =   1920
         Width           =   3495
         _ExtentX        =   6165
         _ExtentY        =   661
      End
      Begin prjLedgerReport.uctlDate uctlFromDate 
         Height          =   375
         Left            =   2160
         TabIndex        =   18
         Top             =   960
         Width           =   3855
         _ExtentX        =   6800
         _ExtentY        =   661
      End
      Begin prjLedgerReport.uctlDate uctlToDate 
         Height          =   375
         Left            =   2160
         TabIndex        =   19
         Top             =   1440
         Width           =   3855
         _ExtentX        =   6800
         _ExtentY        =   661
      End
      Begin prjLedgerReport.uctlTextBox txtSlmCod 
         Height          =   375
         Left            =   7320
         TabIndex        =   22
         Top             =   1440
         Width           =   3495
         _ExtentX        =   5106
         _ExtentY        =   873
      End
      Begin prjLedgerReport.uctlTextBox txtTotalPrice 
         Height          =   375
         Left            =   9720
         TabIndex        =   24
         Top             =   7320
         Width           =   2895
         _ExtentX        =   5106
         _ExtentY        =   873
      End
      Begin VB.Label lblTotalPrice 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   435
         Left            =   7920
         TabIndex        =   25
         Top             =   7320
         Width           =   1755
      End
      Begin VB.Label lblSlmCod 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   6120
         TabIndex        =   23
         Top             =   1440
         Width           =   1065
      End
      Begin VB.Label lblToDate 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   375
         Left            =   0
         TabIndex        =   21
         Top             =   1440
         Width           =   2025
      End
      Begin VB.Label lblFromDate 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   435
         Left            =   120
         TabIndex        =   20
         Top             =   960
         Width           =   1905
      End
      Begin VB.Label lblStkCod 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   435
         Left            =   5760
         TabIndex        =   17
         Top             =   1920
         Width           =   1455
      End
      Begin VB.Label lblIVNo 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   435
         Left            =   480
         TabIndex        =   16
         Top             =   1920
         Width           =   1575
      End
      Begin VB.Label lblOrderType 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   435
         Left            =   5880
         TabIndex        =   15
         Top             =   2400
         Width           =   1365
      End
      Begin VB.Label lblOrderBy 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   435
         Left            =   600
         TabIndex        =   14
         Top             =   2400
         Width           =   1455
      End
      Begin Threed.SSCommand cmdSearch 
         Height          =   525
         Left            =   11160
         TabIndex        =   4
         Top             =   960
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmComMinusStk.frx":2F36
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdClear 
         Height          =   525
         Left            =   11160
         TabIndex        =   5
         Top             =   1680
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         ButtonStyle     =   3
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
         MouseIcon       =   "frmComMinusStk.frx":3250
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
         MouseIcon       =   "frmComMinusStk.frx":356A
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
         Left            =   9360
         TabIndex        =   10
         Top             =   7800
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmComMinusStk.frx":3884
         ButtonStyle     =   3
      End
   End
End
Attribute VB_Name = "frmComMinusStk"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Private m_HasActivate As Boolean
Private comMinusStk As CComMinusStk
Private temp_ComMinusStk As CComMinusStk
Private m_Rs As ADODB.Recordset
Private Rs As ADODB.Recordset
Private m_exitIVStkcod As Collection
Public OKClick As Boolean
Public HeaderText As String
Public SumMinus As Double


Private Sub ChkPaidFLag_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 13 Then
      CreateObject("WScript.Shell").SendKeys "{TAB}"
   End If
End Sub

Private Sub cmdAdd_Click()
Dim ItemCount As Long
Dim OKClick As Boolean
   
   frmAddEditComMinusStk.HeaderText = MapText("เพิ่มข้อมูลส่วนลด")
   frmAddEditComMinusStk.ShowMode = SHOW_ADD
   Set frmAddEditComMinusStk.m_exitIVStkcod = m_exitIVStkcod
   Load frmAddEditComMinusStk
   frmAddEditComMinusStk.Show 1
   
   OKClick = frmAddEditComMinusStk.OKClick
   
   Unload frmAddEditComMinusStk
   Set frmAddEditComMinusStk = Nothing
   
   If OKClick Then
      Call QueryData(True)
   End If
End Sub

Private Sub cmdClear_Click()
   txtIVNo.Text = ""
   txtStkCodNo.Text = ""
   txtSlmCod.Text = ""
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
   
   If Not ConfirmDelete(GridEX1.Value(3) & " - " & GridEX1.Value(5)) Then
      Exit Sub
   End If
   
   Call EnableForm(Me, False)
   comMinusStk.COM_MINUSSTK_ID = ID
   If Not glbDaily.DeleteMinusAmount(comMinusStk, IsOK, True, glbErrorLog) Then
      comMinusStk.COM_MINUSSTK_ID = -1
      glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
      Call EnableForm(Me, True)
      Exit Sub
   End If
   comMinusStk.COM_MINUSSTK_ID = -1
   comMinusStk.MINUS_COD = ""
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

   ID = Val(GridEX1.Value(1))
   
   frmAddEditComMinusStk.COM_MINUSSTK_ID = ID
   frmAddEditComMinusStk.HeaderText = MapText("แก้ไข้อมูลส่วนลด")
   frmAddEditComMinusStk.ShowMode = SHOW_EDIT
      Set frmAddEditComMinusStk.m_exitIVStkcod = m_exitIVStkcod
   Load frmAddEditComMinusStk
   frmAddEditComMinusStk.Show 1
   
   OKClick = frmAddEditComMinusStk.OKClick
   
   Unload frmAddEditComMinusStk
   Set frmAddEditComMinusStk = Nothing
               
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
      
      Call InitRealCreditOrderBy(cboOrderBy)
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
      
      comMinusStk.COM_MINUSSTK_ID = -1
      comMinusStk.IV_COD = txtIVNo.Text
      comMinusStk.STK_COD = txtStkCodNo.Text
      comMinusStk.SLMCOD = txtSlmCod.Text
      comMinusStk.OrderBy = cboOrderBy.ItemData(Minus2Zero(cboOrderBy.ListIndex))
      comMinusStk.OrderType = cboOrderType.ItemData(Minus2Zero(cboOrderType.ListIndex))
      comMinusStk.FROM_DOC_DATE = uctlFromDate.ShowDate
      comMinusStk.TO_DOC_DATE = uctlToDate.ShowDate
      comMinusStk.MINUS_COD = ""
      If Not glbDaily.QueryMinusAmount(comMinusStk, m_exitIVStkcod, m_Rs, ItemCount, IsOK, glbErrorLog) Then
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
   
   SumMinus = 0
 '  Set Rs = m_Rs
   If ItemCount > 0 Then
      While Not m_Rs.EOF
         Call comMinusStk.PopulateFromRS(1, m_Rs)
         SumMinus = SumMinus + comMinusStk.MINUS_AMOUNT
      m_Rs.MoveNext
      Wend
   End If
   
   GridEX1.ItemCount = ItemCount
   GridEX1.Rebind
   
   txtTotalPrice.Text = SumMinus
   
   Call EnableForm(Me, True)
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
   If Shift = 1 And KeyCode = DUMMY_KEY Then
      glbErrorLog.LocalErrorMsg = Me.Name
      glbErrorLog.ShowUserError
      KeyCode = 0
   ElseIf Shift = 0 And KeyCode = 116 Then
      Call cmdSearch_Click
      KeyCode = 0
   ElseIf Shift = 0 And KeyCode = 115 Then
      Call cmdClear_Click
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
   Col.Caption = "ID"

   Set Col = GridEX1.Columns.Add '2
   Col.Width = 2000
   Col.Caption = MapText("วันที่ส่งสินค้า")
      
   Set Col = GridEX1.Columns.Add '3
   Col.Width = 2000
   Col.TextAlignment = jgexAlignCenter
   Col.Caption = MapText("INVOICE")
   
   Set Col = GridEX1.Columns.Add '4
   Col.Width = 1000
   Col.TextAlignment = jgexAlignCenter
   Col.Caption = MapText("รหัส")
   
   Set Col = GridEX1.Columns.Add '5
   Col.Width = 3500
   Col.Caption = MapText("พนักงานขาย")
   
   Set Col = GridEX1.Columns.Add '6
   Col.Width = 2000
   Col.TextAlignment = jgexAlignCenter
   Col.Caption = MapText("รหัสสินค้า")
   
      Set Col = GridEX1.Columns.Add '7
   Col.Width = 3500
   Col.Caption = MapText("ชื่อสินค้า")
   
  Set Col = GridEX1.Columns.Add '8
   Col.Width = 2000
   Col.TextAlignment = jgexAlignRight
   Col.Caption = MapText("ส่วนลด")
   
   Set Col = GridEX1.Columns.Add '8
   Col.Width = 2000
   Col.TextAlignment = jgexAlignCenter
   Col.Caption = MapText("หมายเลข SR")
   
   GridEX1.ItemCount = 0
End Sub

Private Sub InitFormLayout()
   pnlHeader.PictureBackground = LoadPicture(glbParameterObj.NormalForm1)
   SSFrame1.PictureBackground = LoadPicture(glbParameterObj.NormalForm1)
   '   Me.Caption = HeaderText
      
   Me.Caption = MapText("ข้อมูลส่วนลดตามสินค้า")
   pnlHeader.Caption = Me.Caption
   
   Call InitGrid
   
   Call InitNormalLabel(lblFromDate, MapText("จากวันที่"))
   Call InitNormalLabel(lblToDate, MapText("ถึงวันที่"))
   
   Call InitNormalLabel(lblIVNo, MapText("INVOICE"))
   Call InitNormalLabel(lblSlmCod, MapText("รหัสเซลล์"))
   Call InitNormalLabel(lblStkCod, MapText("รหัสสินค้า"))
      
   Call InitNormalLabel(lblOrderBy, MapText("เรียงตาม"))
   Call InitNormalLabel(lblOrderType, MapText("เรียงจาก"))
   
   Call txtIVNo.SetTextLenType(TEXT_STRING, glbSetting.CODE_TYPE)
   Call txtStkCodNo.SetTextLenType(TEXT_STRING, glbSetting.CODE_TYPE)
   Call txtSlmCod.SetTextLenType(TEXT_STRING, glbSetting.CODE_TYPE)
   
   Call InitNormalLabel(lblTotalPrice, MapText("รวมส่วนลด"))

   Call txtTotalPrice.SetTextLenType(TEXT_FLOAT_MONEY, glbSetting.MONEY_TYPE)
   txtTotalPrice.Enabled = False
   
   Call InitCombo(cboOrderBy)
   Call InitCombo(cboOrderType)

   Me.Picture = LoadPicture(glbParameterObj.MainPicture)
   
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
      OKClick = False
      
   m_HasActivate = False
   Set comMinusStk = New CComMinusStk
   Set temp_ComMinusStk = New CComMinusStk
   Set m_Rs = New ADODB.Recordset
   Set Rs = New ADODB.Recordset
   Set m_exitIVStkcod = New Collection
   
      m_HasActivate = False
   Call InitFormLayout
   Call EnableForm(Me, True)
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set comMinusStk = Nothing
   Set temp_ComMinusStk = Nothing
   Set m_Rs = Nothing
   Set Rs = Nothing
   Set m_exitIVStkcod = Nothing
End Sub

Private Sub GridEX1_DblClick()
   Call cmdEdit_Click
End Sub
Private Sub GridEX1_RowFormat(RowBuffer As GridEX20.JSRowData)
   RowBuffer.RowStyle = RowBuffer.Value(4)
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

'   If m_Rs.EOF Then
'      Exit Sub
'   End If

   If RowIndex <= 0 Then
      Exit Sub
   End If

   Call m_Rs.Move(RowIndex - 1, adBookmarkFirst)
   Call temp_ComMinusStk.PopulateFromRS(1, m_Rs)
   
   Values(1) = temp_ComMinusStk.COM_MINUSSTK_ID
   Values(2) = temp_ComMinusStk.IV_DOCDAT
   Values(3) = temp_ComMinusStk.IV_COD
   Values(4) = temp_ComMinusStk.SLMCOD
   Values(5) = temp_ComMinusStk.SLMNAME
   Values(6) = temp_ComMinusStk.STK_COD
   Values(7) = temp_ComMinusStk.STK_NAME
   Values(8) = FormatNumber(temp_ComMinusStk.MINUS_AMOUNT)
   Values(9) = temp_ComMinusStk.MINUS_COD

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
   GridEX1.Height = ScaleHeight - GridEX1.Top - 1200
   cmdAdd.Top = ScaleHeight - 580
   cmdEdit.Top = ScaleHeight - 580
   cmdDelete.Top = ScaleHeight - 580
   cmdOK.Top = ScaleHeight - 580
   cmdExit.Top = ScaleHeight - 580
   cmdExit.Left = ScaleWidth - cmdExit.Width - 50
   cmdOK.Left = cmdExit.Left - cmdOK.Width - 50
   lblTotalPrice.Top = ScaleHeight - 1050
   txtTotalPrice.Top = ScaleHeight - 1050
   lblTotalPrice.Left = ScaleWidth - lblTotalPrice.Width - 500 - txtTotalPrice.Width - 2000
   txtTotalPrice.Left = ScaleWidth - txtTotalPrice.Width - 2000
End Sub

