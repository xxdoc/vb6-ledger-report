VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#2.0#0"; "THREED20.OCX"
Object = "{E684D8A3-716C-4E59-AA94-7144C04B0074}#1.1#0"; "GridEX20.ocx"
Begin VB.Form frmComIVcenter 
   BackColor       =   &H80000000&
   ClientHeight    =   8520
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   13800
   Icon            =   "frmComIVcenter.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   8520
   ScaleWidth      =   13800
   StartUpPosition =   1  'CenterOwner
   WindowState     =   2  'Maximized
   Begin Threed.SSFrame SSFrame1 
      Height          =   8535
      Left            =   0
      TabIndex        =   11
      Top             =   0
      Width           =   13755
      _ExtentX        =   24262
      _ExtentY        =   15055
      _Version        =   131073
      PictureBackgroundStyle=   2
      Begin VB.ComboBox cboAreaName 
         Height          =   315
         Left            =   7080
         Style           =   2  'Dropdown List
         TabIndex        =   21
         Top             =   1440
         Width           =   3855
      End
      Begin prjLedgerReport.uctlTextBox txtIVNo 
         Height          =   375
         Left            =   1680
         TabIndex        =   0
         Top             =   1440
         Width           =   3855
         _ExtentX        =   6800
         _ExtentY        =   661
      End
      Begin VB.ComboBox cboOrderType 
         Height          =   315
         Left            =   7080
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   1920
         Width           =   3915
      End
      Begin VB.ComboBox cboOrderBy 
         Height          =   315
         Left            =   1680
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   1920
         Width           =   3795
      End
      Begin Threed.SSPanel pnlHeader 
         Height          =   705
         Left            =   0
         TabIndex        =   12
         Top             =   0
         Width           =   13725
         _ExtentX        =   24209
         _ExtentY        =   1244
         _Version        =   131073
         PictureBackgroundStyle=   2
      End
      Begin GridEX20.GridEX GridEX1 
         Height          =   4815
         Left            =   180
         TabIndex        =   5
         Top             =   2880
         Width           =   13305
         _ExtentX        =   23469
         _ExtentY        =   8493
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
         Column(1)       =   "frmComIVcenter.frx":27A2
         Column(2)       =   "frmComIVcenter.frx":286A
         FormatStylesCount=   5
         FormatStyle(1)  =   "frmComIVcenter.frx":290E
         FormatStyle(2)  =   "frmComIVcenter.frx":2A6A
         FormatStyle(3)  =   "frmComIVcenter.frx":2B1A
         FormatStyle(4)  =   "frmComIVcenter.frx":2BCE
         FormatStyle(5)  =   "frmComIVcenter.frx":2CA6
         ImageCount      =   0
         PrinterProperties=   "frmComIVcenter.frx":2D5E
      End
      Begin prjLedgerReport.uctlDate uctlFromDate 
         Height          =   375
         Left            =   1680
         TabIndex        =   17
         Top             =   960
         Width           =   3855
         _ExtentX        =   6800
         _ExtentY        =   661
      End
      Begin prjLedgerReport.uctlDate uctlToDate 
         Height          =   375
         Left            =   7080
         TabIndex        =   18
         Top             =   960
         Width           =   3855
         _ExtentX        =   6800
         _ExtentY        =   661
      End
      Begin prjLedgerReport.uctlTextLookup uctlCommissionSale 
         Height          =   375
         Left            =   1680
         TabIndex        =   22
         Top             =   2400
         Width           =   6255
         _ExtentX        =   16325
         _ExtentY        =   661
      End
      Begin VB.Label lblSaleName 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   435
         Left            =   0
         TabIndex        =   23
         Top             =   2400
         Width           =   1575
      End
      Begin VB.Label lblToDate 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   375
         Left            =   5640
         TabIndex        =   20
         Top             =   960
         Width           =   1305
      End
      Begin VB.Label lblFromDate 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   435
         Left            =   240
         TabIndex        =   19
         Top             =   960
         Width           =   1305
      End
      Begin VB.Label lblArea 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   435
         Left            =   5520
         TabIndex        =   16
         Top             =   1440
         Width           =   1455
      End
      Begin VB.Label lblIVNo 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   435
         Left            =   120
         TabIndex        =   15
         Top             =   1440
         Width           =   1455
      End
      Begin VB.Label lblOrderType 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   435
         Left            =   5640
         TabIndex        =   14
         Top             =   1920
         Width           =   1365
      End
      Begin VB.Label lblOrderBy 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   435
         Left            =   240
         TabIndex        =   13
         Top             =   1920
         Width           =   1335
      End
      Begin Threed.SSCommand cmdSearch 
         Height          =   525
         Left            =   11880
         TabIndex        =   3
         Top             =   960
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmComIVcenter.frx":2F36
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdClear 
         Height          =   525
         Left            =   11880
         TabIndex        =   4
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
         TabIndex        =   8
         Top             =   7830
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmComIVcenter.frx":3250
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdAdd 
         Height          =   525
         Left            =   150
         TabIndex        =   6
         Top             =   7830
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmComIVcenter.frx":356A
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdEdit 
         Height          =   525
         Left            =   1770
         TabIndex        =   7
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
         Left            =   11880
         TabIndex        =   10
         Top             =   7800
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   926
         _Version        =   131073
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdOK 
         Height          =   525
         Left            =   10200
         TabIndex        =   9
         Top             =   7800
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmComIVcenter.frx":3884
         ButtonStyle     =   3
      End
   End
End
Attribute VB_Name = "frmComIVcenter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Private m_HasActivate As Boolean
Private ComIVcenter As CComIVcenter
Private temp_ComIVcenter As CComIVcenter
Private m_Rs As ADODB.Recordset
Private m_exitIVcenter As Collection
Public OKClick As Boolean
Public HeaderText As String

Private FtSaleColl As Collection
Private Sub ChkPaidFLag_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 13 Then
      CreateObject("WScript.Shell").SendKeys "{TAB}"
   End If
End Sub

Private Sub cmdAdd_Click()
Dim ItemCount As Long
Dim OKClick As Boolean
   
   frmAddEditComIVcenter.HeaderText = MapText("���������� IV �����Center")
   frmAddEditComIVcenter.ShowMode = SHOW_ADD
   Set frmAddEditComIVcenter.m_exitIVcenter = m_exitIVcenter
   Load frmAddEditComIVcenter
   frmAddEditComIVcenter.Show 1
   
   OKClick = frmAddEditComIVcenter.OKClick
   
   Unload frmAddEditComIVcenter
   Set frmAddEditComIVcenter = Nothing
   
   If OKClick Then
      Call QueryData(True)
   End If
End Sub

Private Sub cmdClear_Click()
   txtIVNo.Text = ""
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
   
   If Not ConfirmDelete(GridEX1.Value(3) & " - " & GridEX1.Value(6)) Then
      Exit Sub
   End If
   
   Call EnableForm(Me, False)
   ComIVcenter.COM_IV_CENTER_ID = ID
   If Not glbDaily.DeleteIVcenter(ComIVcenter, IsOK, True, glbErrorLog) Then
      ComIVcenter.COM_IV_CENTER_ID = -1
      glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
      Call EnableForm(Me, True)
      Exit Sub
   End If
   ComIVcenter.COM_IV_CENTER_ID = -1
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
   
   frmAddEditComIVcenter.COM_IV_CENTER_ID = ID
   frmAddEditComIVcenter.HeaderText = MapText("�������� IV �����Center")
   frmAddEditComIVcenter.ShowMode = SHOW_EDIT
      Set frmAddEditComIVcenter.m_exitIVcenter = m_exitIVcenter
   Load frmAddEditComIVcenter
   frmAddEditComIVcenter.Show 1
   
   OKClick = frmAddEditComIVcenter.OKClick
   
   Unload frmAddEditComIVcenter
   Set frmAddEditComIVcenter = Nothing
               
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
      
      Call InitIVcenterOrderBy(cboOrderBy)
      Call InitOrderType(cboOrderType)
      Call LoadAreaCom(cboAreaName)
            
      Call LoadSaleLookup(uctlCommissionSale.MyCombo, FtSaleColl) 'FtSaleColl, COMMISSION_TABLE
      Set uctlCommissionSale.MyCollection = FtSaleColl
      
      Call QueryData(True)
   End If
End Sub
Private Sub QueryData(Flag As Boolean)
Dim IsOK As Boolean
Dim ItemCount As Long
Dim Temp As Long

   If Flag Then
      Call EnableForm(Me, False)
      
'      ComIVcenter.COM_MINUSSTK_ID = -1
      ComIVcenter.IV_COD = txtIVNo.Text
       ComIVcenter.MASTER_AREA_ID = cboAreaName.ItemData(Minus2Zero(cboAreaName.ListIndex))
       ComIVcenter.SLMCOD = uctlCommissionSale.MyTextBox.Text
      ComIVcenter.ORDER_BY = cboOrderBy.ItemData(Minus2Zero(cboOrderBy.ListIndex))
      ComIVcenter.ORDER_TYPE = cboOrderType.ItemData(Minus2Zero(cboOrderType.ListIndex))
      ComIVcenter.FROM_DOC_DATE = uctlFromDate.ShowDate
      ComIVcenter.TO_DOC_DATE = uctlToDate.ShowDate
      If Not glbDaily.QueryIVcenter(ComIVcenter, m_exitIVcenter, m_Rs, ItemCount, IsOK, glbErrorLog) Then
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
   Col.Caption = MapText("�ѹ������Թ���")
      
   Set Col = GridEX1.Columns.Add '3
   Col.Width = 2000
   Col.Caption = MapText("INVOICE")
   
   Set Col = GridEX1.Columns.Add '4
   Col.Width = 700
   Col.Caption = MapText("����")
   
   Set Col = GridEX1.Columns.Add '5
   Col.Width = 3500
   Col.Caption = MapText("��ѡ�ҹ���")
   
      Set Col = GridEX1.Columns.Add '6
   Col.Width = 4500
   Col.Caption = MapText("ࢵ��â��")

'  Set Col = GridEX1.Columns.Add '4
'   Col.Width = 3000
'   Col.TextAlignment = jgexAlignRight
'   Col.Caption = MapText("��ǹŴ")
   
   GridEX1.ItemCount = 0
End Sub

Private Sub InitFormLayout()
   pnlHeader.PictureBackground = LoadPicture(glbParameterObj.NormalForm1)
   SSFrame1.PictureBackground = LoadPicture(glbParameterObj.NormalForm1)
    Me.Caption = HeaderText

   pnlHeader.Caption = Me.Caption
   
   Call InitGrid
   
   Call InitNormalLabel(lblFromDate, MapText("�ҡ�ѹ���"))
   Call InitNormalLabel(lblToDate, MapText("�֧�ѹ���"))
   
   Call InitNormalLabel(lblIVNo, MapText("INVOICE"))
   Call InitNormalLabel(lblArea, MapText("ࢵ��â��"))
   Call InitNormalLabel(lblSaleName, MapText("��ѡ�ҹ���"))
      
   Call InitNormalLabel(lblOrderBy, MapText("���§���"))
   Call InitNormalLabel(lblOrderType, MapText("���§�ҡ"))
   
   Call txtIVNo.SetTextLenType(TEXT_STRING, glbSetting.CODE_TYPE)
   
   Call InitCombo(cboOrderBy)
   Call InitCombo(cboOrderType)
   Call InitCombo(cboAreaName)

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
   
   Call InitMainButton(cmdExit, MapText("¡��ԡ (ESC)"))
   Call InitMainButton(cmdOK, MapText("��ŧ (F2)"))
   Call InitMainButton(cmdAdd, MapText("���� (F7)"))
   Call InitMainButton(cmdEdit, MapText("��� (F3)"))
   Call InitMainButton(cmdDelete, MapText("ź (F6)"))
   Call InitMainButton(cmdSearch, MapText("���� (F5)"))
   Call InitMainButton(cmdClear, MapText("������ (F4)"))
End Sub

Private Sub cmdExit_Click()
   OKClick = False
   Unload Me
End Sub

Private Sub Form_Load()
      OKClick = False
      
   m_HasActivate = False
   Set ComIVcenter = New CComIVcenter
   Set temp_ComIVcenter = New CComIVcenter
   Set m_Rs = New ADODB.Recordset
   Set m_exitIVcenter = New Collection
   Set FtSaleColl = New Collection
   
      m_HasActivate = False
   Call InitFormLayout
   Call EnableForm(Me, True)
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set ComIVcenter = Nothing
   Set temp_ComIVcenter = Nothing
   Set m_Rs = Nothing
   Set m_exitIVcenter = Nothing
   Set FtSaleColl = Nothing
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

   If m_Rs.EOF Then
      Exit Sub
   End If

   If RowIndex <= 0 Then
      Exit Sub
   End If

   Call m_Rs.Move(RowIndex - 1, adBookmarkFirst)
   Call temp_ComIVcenter.PopulateFromRS(1, m_Rs)
   
   Values(1) = temp_ComIVcenter.COM_IV_CENTER_ID
   Values(2) = temp_ComIVcenter.IV_DOCDAT
   Values(3) = temp_ComIVcenter.IV_COD
   Values(4) = temp_ComIVcenter.SLMCOD
   Values(5) = temp_ComIVcenter.SLMNAME
   Values(6) = temp_ComIVcenter.MASTER_AREA_NAME

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

