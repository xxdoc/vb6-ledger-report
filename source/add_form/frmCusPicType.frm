VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#2.0#0"; "THREED20.OCX"
Object = "{E684D8A3-716C-4E59-AA94-7144C04B0074}#1.1#0"; "GridEX20.ocx"
Begin VB.Form frmCusPigType 
   BackColor       =   &H80000000&
   ClientHeight    =   7875
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11910
   Icon            =   "frmCusPicType.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   7875
   ScaleWidth      =   11910
   StartUpPosition =   1  'CenterOwner
   WindowState     =   2  'Maximized
   Begin Threed.SSFrame SSFrame1 
      Height          =   7935
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   11955
      _ExtentX        =   21087
      _ExtentY        =   13996
      _Version        =   131073
      PictureBackgroundStyle=   2
      Begin VB.ComboBox cboOrderBy 
         Height          =   315
         Left            =   1680
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Top             =   1440
         Width           =   2955
      End
      Begin VB.ComboBox cboOrderType 
         Height          =   315
         Left            =   6050
         Style           =   2  'Dropdown List
         TabIndex        =   9
         Top             =   1440
         Width           =   2955
      End
      Begin Threed.SSPanel pnlHeader 
         Height          =   705
         Left            =   0
         TabIndex        =   5
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
         TabIndex        =   4
         Top             =   2160
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
         Column(1)       =   "frmCusPicType.frx":27A2
         Column(2)       =   "frmCusPicType.frx":286A
         FormatStylesCount=   5
         FormatStyle(1)  =   "frmCusPicType.frx":290E
         FormatStyle(2)  =   "frmCusPicType.frx":2A6A
         FormatStyle(3)  =   "frmCusPicType.frx":2B1A
         FormatStyle(4)  =   "frmCusPicType.frx":2BCE
         FormatStyle(5)  =   "frmCusPicType.frx":2CA6
         ImageCount      =   0
         PrinterProperties=   "frmCusPicType.frx":2D5E
      End
      Begin prjLedgerReport.uctlTextBox txtCustomerCode 
         Height          =   435
         Left            =   1680
         TabIndex        =   0
         Top             =   840
         Width           =   2895
         _ExtentX        =   5106
         _ExtentY        =   767
      End
      Begin prjLedgerReport.uctlTextBox txtCustomerName 
         Height          =   435
         Left            =   6050
         TabIndex        =   2
         Top             =   840
         Width           =   2895
         _ExtentX        =   5106
         _ExtentY        =   767
      End
      Begin VB.Label lblCustomerName 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   435
         Left            =   4530
         TabIndex        =   11
         Top             =   900
         Width           =   1455
      End
      Begin Threed.SSCommand cmdClear 
         Height          =   525
         Left            =   9960
         TabIndex        =   12
         Top             =   1440
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdSearch 
         Height          =   525
         Left            =   9960
         TabIndex        =   3
         Top             =   840
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmCusPicType.frx":2F36
         ButtonStyle     =   3
      End
      Begin VB.Label lblOrderBy 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   435
         Left            =   120
         TabIndex        =   10
         Top             =   1440
         Width           =   1455
      End
      Begin VB.Label lblOrderType 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   435
         Left            =   4530
         TabIndex        =   8
         Top             =   1440
         Width           =   1365
      End
      Begin VB.Label lblCustomerCode 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   435
         Left            =   0
         TabIndex        =   6
         Top             =   930
         Width           =   1575
      End
      Begin Threed.SSCommand cmdDelete 
         Height          =   525
         Left            =   3600
         TabIndex        =   15
         Top             =   7200
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmCusPicType.frx":3250
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdAdd 
         Height          =   525
         Left            =   240
         TabIndex        =   13
         Top             =   7200
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmCusPicType.frx":356A
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdEdit 
         Height          =   525
         Left            =   1920
         TabIndex        =   14
         Top             =   7200
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   926
         _Version        =   131073
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdExit 
         Cancel          =   -1  'True
         Height          =   525
         Left            =   10095
         TabIndex        =   17
         Top             =   7200
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   926
         _Version        =   131073
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdOK 
         Height          =   525
         Left            =   8445
         TabIndex        =   16
         Top             =   7200
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmCusPicType.frx":3884
         ButtonStyle     =   3
      End
   End
End
Attribute VB_Name = "frmCusPigType"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private m_HasActivate As Boolean
Private m_ARMas As CARMas
Private m_TempARMas As CARMas
Private m_Rs As ADODB.Recordset
Public OKClick As Boolean
Public HeaderText As String
Public m_CUS_PIG_TYPE As Collection

Private Sub cmdAdd_Click()
Dim ItemCount As Long
Dim OKClick As Boolean
   
   frmAddEditCusPigType.HeaderText = MapText("���������š�����١˹��")
   frmAddEditCusPigType.ShowMode = SHOW_ADD
   frmAddEditCusPigType.CUS_PIG_TYPE_CODE = GridEX1.Value(2)
   frmAddEditCusPigType.CUS_PIG_TYPE_NAME = GridEX1.Value(3)
   Load frmAddEditCusPigType
   frmAddEditCusPigType.Show 1
   
   OKClick = frmAddEditCusPigType.OKClick
   
   Unload frmAddEditCusPigType
   Set frmAddEditCusPigType = Nothing
   
   If OKClick Then
      Call QueryData(True)
   End If
End Sub

Private Sub cmdClear_Click()
   txtCustomerCode.Text = ""
   txtCustomerName.Text = ""
   cboOrderBy.ListIndex = -1
   cboOrderType.ListIndex = -1
End Sub

Private Sub cmdDelete_Click()
Dim IsOK As Boolean
Dim ItemCount As Long
Dim IsCanLock As Boolean
Dim ID As Long

   If Not VerifyGrid(GridEX1.Value(2)) Then
      Exit Sub
   End If
   
 '  ID = GridEX1.Value(2)
   
   If Not ConfirmDelete(GridEX1.Value(2) & " " & GridEX1.Value(3)) Then
      Exit Sub
   End If
   
   Call EnableForm(Me, False)
   'm_ARMas.KEY_ID = ID
   m_ARMas.CUSCOD = GridEX1.Value(2)
   If Not glbDaily.DeleteCusPigType(m_ARMas, IsOK, True, glbErrorLog) Then
        m_ARMas.CUSCOD = -1
        glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
      m_ARMas.CUSCOD = ""
      glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
      m_ARMas.CUSNAM = ""
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
   
   If Not VerifyGrid(GridEX1.Value(2)) Then
      Exit Sub
   End If
   


   frmAddEditCusPigType.HeaderText = MapText("��䢢����Ż��������")
   frmAddEditCusPigType.ShowMode = SHOW_EDIT
   frmAddEditCusPigType.CUS_PIG_TYPE_CODE = GridEX1.Value(2)
   frmAddEditCusPigType.CUS_PIG_TYPE_NAME = GridEX1.Value(3)
   Load frmAddEditCusPigType
   frmAddEditCusPigType.Show 1
   
   OKClick = frmAddEditCusPigType.OKClick
   
   Unload frmAddEditCusPigType
   Set frmAddEditCusPigType = Nothing
               
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
      
      Call InitCustomer2OrderBy(cboOrderBy)
      Call InitOrderType(cboOrderType)
      
      
      Call QueryData(True)
   End If
End Sub
Private Sub QueryData(Flag As Boolean)
Dim IsOK As Boolean
Dim ItemCount As Long
Dim Temp As Long
   Call LoadCusPigType(m_CUS_PIG_TYPE)
   If Flag Then
      Call EnableForm(Me, False)
      m_ARMas.CUSCOD = PatchWildCard(txtCustomerCode.Text)
      m_ARMas.CUSNAM = PatchWildCard(txtCustomerName.Text)
      m_ARMas.OrderBy = cboOrderBy.ItemData(Minus2Zero(cboOrderBy.ListIndex))
      m_ARMas.OrderType = cboOrderType.ItemData(Minus2Zero(cboOrderType.ListIndex))
      If Not glbDaily.QueryARMas(m_ARMas, m_Rs, ItemCount, IsOK, glbErrorLog) Then
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
   Col.Caption = MapText("�����١���")
   
   Set Col = GridEX1.Columns.Add '3
   Col.Width = 8000
    Col.Caption = MapText("�����١���")
    
    Set Col = GridEX1.Columns.Add '4
    Col.Width = 2000
    Col.Caption = MapText("�ӹǹ��پѹ���")
    
    Set Col = GridEX1.Columns.Add '5
    Col.Width = 2000
    Col.Caption = MapText("�ӹǹ��٢ع")
   
    Set Col = GridEX1.Columns.Add '6
    Col.Width = 2000
    Col.Caption = MapText("�ӹǹ�١���")
   GridEX1.ItemCount = 0
End Sub
Private Sub InitFormLayout()
   pnlHeader.PictureBackground = LoadPicture(glbParameterObj.NormalForm1)
   SSFrame1.PictureBackground = LoadPicture(glbParameterObj.NormalForm1)
   
   Me.Caption = MapText("�����Ũӹǹ������л�����")
   pnlHeader.Caption = Me.Caption
   
   Call InitGrid
   
   Call InitNormalLabel(lblCustomerCode, MapText("�����١���"))
   Call InitNormalLabel(lblCustomerName, MapText("�����١���"))
   Call InitNormalLabel(lblOrderBy, MapText("���§���"))
   Call InitNormalLabel(lblOrderType, MapText("���§�ҡ"))

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
   
   Call InitMainButton(cmdExit, MapText("¡��ԡ (ESC)"))
   Call InitMainButton(cmdOK, MapText("��ŧ (F2)"))
   Call InitMainButton(cmdAdd, MapText("���� (F7)"))
   Call InitMainButton(cmdEdit, MapText("��� (F3)"))
   Call InitMainButton(cmdDelete, MapText("ź (F6)"))
   Call InitMainButton(cmdSearch, MapText("���� (F5)"))
   Call InitMainButton(cmdClear, MapText("������ (F4)"))
   
'   cmdAdd.Enabled = False
   
End Sub
Private Sub cmdExit_Click()
   OKClick = False
   Unload Me
End Sub

Private Sub Form_Load()
   
   Set m_ARMas = New CARMas
   Set m_TempARMas = New CARMas
   Set m_Rs = New ADODB.Recordset
   Set m_CUS_PIG_TYPE = New Collection

   Call InitFormLayout
   Call EnableForm(Me, True)
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set m_CUS_PIG_TYPE = Nothing
End Sub

Private Sub GridEX1_DblClick()
   Call cmdEdit_Click
End Sub

Private Sub GridEX1_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
   Call cmdEdit_Click
End If
End Sub

Private Sub GridEX1_RowFormat(RowBuffer As GridEX20.JSRowData)
   RowBuffer.RowStyle = RowBuffer.Value(2)
End Sub
Private Sub GridEX1_UnboundReadData(ByVal RowIndex As Long, ByVal Bookmark As Variant, ByVal Values As GridEX20.JSRowData)
On Error GoTo ErrorHandler
Dim RealIndex As Long
Dim tempCusPigType As CCusPigType

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
   Call m_TempARMas.PopulateFromRS(1, m_Rs)
   
   Values(1) = ""
   Values(2) = m_TempARMas.CUSCOD
   Values(3) = m_TempARMas.CUSNAM

   Set tempCusPigType = GetObject("CCusPigType", m_CUS_PIG_TYPE, Trim(m_TempARMas.CUSCOD), False)
   
   If Not tempCusPigType Is Nothing Then
      Values(4) = tempCusPigType.CUS_PIG_TYPE_BREED
      Values(5) = tempCusPigType.CUS_PIG_TYPE_KHUN
      Values(6) = tempCusPigType.CUS_PIG_TYPE_PIGGY
   End If
   Set tempCusPigType = Nothing
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

