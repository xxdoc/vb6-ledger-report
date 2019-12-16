VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#2.0#0"; "THREED20.OCX"
Object = "{E684D8A3-716C-4E59-AA94-7144C04B0074}#1.1#0"; "GridEX20.ocx"
Begin VB.Form frmAddEditAreaInEP 
   BackColor       =   &H80000000&
   ClientHeight    =   8520
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11910
   Icon            =   "frmAddEditAreaInEP.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   8520
   ScaleWidth      =   11910
   StartUpPosition =   1  'CenterOwner
   WindowState     =   2  'Maximized
   Begin Threed.SSFrame SSFrame1 
      Height          =   8535
      Left            =   0
      TabIndex        =   9
      Top             =   0
      Width           =   11955
      _ExtentX        =   21087
      _ExtentY        =   15055
      _Version        =   131073
      PictureBackgroundStyle=   2
      Begin Threed.SSPanel pnlHeader 
         Height          =   705
         Left            =   0
         TabIndex        =   10
         Top             =   0
         Width           =   11925
         _ExtentX        =   21034
         _ExtentY        =   1244
         _Version        =   131073
         PictureBackgroundStyle=   2
      End
      Begin GridEX20.GridEX GridEX1 
         Height          =   5535
         Left            =   180
         TabIndex        =   5
         Top             =   2160
         Width           =   11505
         _ExtentX        =   20294
         _ExtentY        =   9763
         Version         =   "2.0"
         BoundColumnIndex=   ""
         ReplaceColumnIndex=   ""
         TabKeyBehavior  =   1
         MultiSelect     =   -1  'True
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
         Column(1)       =   "frmAddEditAreaInEP.frx":27A2
         Column(2)       =   "frmAddEditAreaInEP.frx":286A
         FormatStylesCount=   5
         FormatStyle(1)  =   "frmAddEditAreaInEP.frx":290E
         FormatStyle(2)  =   "frmAddEditAreaInEP.frx":2A6A
         FormatStyle(3)  =   "frmAddEditAreaInEP.frx":2B1A
         FormatStyle(4)  =   "frmAddEditAreaInEP.frx":2BCE
         FormatStyle(5)  =   "frmAddEditAreaInEP.frx":2CA6
         ImageCount      =   0
         PrinterProperties=   "frmAddEditAreaInEP.frx":2D5E
      End
      Begin prjLedgerReport.uctlTextBox txtCusCod 
         Height          =   375
         Left            =   2040
         TabIndex        =   0
         Top             =   1080
         Width           =   2775
         _ExtentX        =   4895
         _ExtentY        =   661
      End
      Begin prjLedgerReport.uctlTextBox txtCusName 
         Height          =   375
         Left            =   2040
         TabIndex        =   1
         Tag             =   "2"
         Top             =   1560
         Width           =   4695
         _ExtentX        =   8281
         _ExtentY        =   661
      End
      Begin Threed.SSCommand cmdClear 
         Height          =   525
         Left            =   9960
         TabIndex        =   4
         Top             =   1560
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
         Top             =   960
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmAddEditAreaInEP.frx":2F36
         ButtonStyle     =   3
      End
      Begin Threed.SSCheck ChkNonAreaFLag 
         Height          =   255
         Left            =   7560
         TabIndex        =   2
         Top             =   1560
         Visible         =   0   'False
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   450
         _Version        =   131073
         Caption         =   "SSCheck1"
      End
      Begin VB.Label lblCusName 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   495
         Left            =   120
         TabIndex        =   12
         Top             =   1560
         Width           =   1815
      End
      Begin VB.Label lblCusCod 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   435
         Left            =   240
         TabIndex        =   11
         Top             =   1080
         Width           =   1695
      End
      Begin Threed.SSCommand cmdAddMasArea 
         Height          =   525
         Left            =   8400
         TabIndex        =   7
         Top             =   7800
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdEdit 
         Height          =   525
         Left            =   120
         TabIndex        =   6
         Top             =   7830
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   926
         _Version        =   131073
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdOK 
         Height          =   525
         Left            =   10080
         TabIndex        =   8
         Top             =   7830
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmAddEditAreaInEP.frx":3250
         ButtonStyle     =   3
      End
   End
End
Attribute VB_Name = "frmAddEditAreaInEP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private m_HasActivate As Boolean
Private m_HasModify As Boolean
Private m_cusAcess As Collection
Private tempCusAcess As CCommissionCustomerArea

Private m_customerEP As CARMas
Public HeaderText As String

Private m_Rs As ADODB.Recordset
Private mm_Rs As ADODB.Recordset
Public OKClick As Boolean
Public YEAR_ID As Long
Public ShowMode As SHOW_MODE_TYPE

Public MASTER_AREA_ID As Long
Dim RowDelete As Long

Private Sub cmdClear_Click()
   txtCusCod.Text = ""
   txtCusName.Text = ""
End Sub
Private Sub cmdEdit_Click()
Dim IsOK As Boolean
Dim ItemCount As Long
Dim IsCanLock As Boolean
Dim ID As Long
Dim CUSCOD As String
Dim OKClick As Boolean
   
   If Not VerifyGrid(GridEX1.Value(1)) Then
      Exit Sub
   End If

   ID = GridEX1.Value(1)
   CUSCOD = GridEX1.Value(2)
      
    frmAddEditAreaInEP2.COMMISSION_CUS_AREA_ID = ID
    frmAddEditAreaInEP2.HeaderText = MapText("แก้ไขข้อมูลลูกค้า")
   frmAddEditAreaInEP2.YEAR_ID = YEAR_ID
   frmAddEditAreaInEP2.CUSCOD = CUSCOD
   Load frmAddEditAreaInEP2
   frmAddEditAreaInEP2.Show 1

   OKClick = frmAddEditAreaInEP2.OKClick

   Unload frmAddEditAreaInEP2
  Set frmAddEditAreaInEP2 = Nothing

   If OKClick Then
      Set m_cusAcess = Nothing
      Set m_cusAcess = New Collection
      Call LoadCusFromAreaNameCom(Nothing, YEAR_ID, m_cusAcess)
      Call QueryData(True)
   End If

End Sub

Private Sub cmdOK_Click()
   OKClick = True
   Unload Me
End Sub

Private Sub cmdSearch_Click()

   'ทุกครั้ง เพื่อเช็คว่าติ๊กหรือไม่ = ใช้การไม่ได้อยู่ดี เพราะลูกค้าจาก EP มาทั้งหมด
'      Set m_cusAcess = Nothing
'      Set m_cusAcess = New Collection
'      Call LoadCusNonAreaCom(Check2Flag(ChkNonAreaFLag.Value), YEAR_ID, m_cusAcess)

   Call QueryData(True)
End Sub

Private Sub Form_Activate()
   If Not m_HasActivate Then
      m_HasActivate = True
      
      Call LoadCusFromAreaNameCom(Nothing, YEAR_ID, m_cusAcess)

      If ShowMode = SHOW_EDIT Then
         m_customerEP.CUSCOD = ""
         m_customerEP.CUSNAM = ""
         m_customerEP.CUSTYP = ""
         m_customerEP.SLMCOD = ""
         Call QueryData(True)
      End If
      
   End If
End Sub
Private Sub QueryData(Flag As Boolean)
Dim IsOK As Boolean
Dim ItemCount As Long
Dim m_ItemCount As Long
Dim Temp As Long
                     
   Set m_customerEP = Nothing
   Set m_customerEP = New CARMas
   
   m_customerEP.CUSCOD = PatchWildCard(txtCusCod.Text)
   m_customerEP.CUSNAM = PatchWildCard(txtCusName.Text)

   If Not glbDaily.QueryCustomer(m_customerEP, m_Rs, ItemCount, IsOK, glbErrorLog) Then
      glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
      Call EnableForm(Me, True)
      Exit Sub
   End If
   Call m_customerEP.PopulateFromRS(1, m_Rs)
   
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
'      Call cmdAdd_Click
      KeyCode = 0
   ElseIf Shift = 0 And KeyCode = 117 Then
'      Call cmdDelete_Click
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

Private Sub ChkNonAreaFlag_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 13 Then
      CreateObject("WScript.Shell").SendKeys "{TAB}"
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
   Col.Caption = MapText("รหัสลูกค้า")
   
   Set Col = GridEX1.Columns.Add '3
   Col.Width = 5000
    Col.Caption = MapText("ชื่อลูกค้า")
   
   Set Col = GridEX1.Columns.Add '3
   Col.Width = 4000
    Col.Caption = MapText("เขตการขาย")
   
   GridEX1.ItemCount = 0
End Sub
Private Sub InitFormLayout()
   pnlHeader.PictureBackground = LoadPicture(glbParameterObj.NormalForm1)
   SSFrame1.PictureBackground = LoadPicture(glbParameterObj.NormalForm1)
   
   Me.Caption = MapText("ข้อมูลลูกของเขตการขาย")
   pnlHeader.Caption = Me.Caption
   
   Call InitNormalLabel(lblCusCod, MapText("รหัสลูกค้า"))
   Call InitNormalLabel(lblCusName, MapText("ชื่อลูกค้า"))
   Call txtCusCod.SetTextLenType(TEXT_STRING, glbSetting.CODE_TYPE)
   Call txtCusName.SetTextLenType(TEXT_STRING, glbSetting.NAME_TYPE)
   
   Call InitCheckBox(ChkNonAreaFLag, "ยังไม่ระบุเขต")
   ChkNonAreaFLag.Enabled = False
   
   Call InitGrid

   Me.Picture = LoadPicture(glbParameterObj.MainPicture)
   
   pnlHeader.Font.Name = GLB_FONT
   pnlHeader.Font.Bold = True
   pnlHeader.Font.Size = 19
   
   cmdOK.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdEdit.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdSearch.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdClear.Picture = LoadPicture(glbParameterObj.NormalButton1)
   
   Call InitMainButton(cmdOK, MapText("ตกลง (F2)"))
   Call InitMainButton(cmdEdit, MapText("แก้ไข (F3)"))
   Call InitMainButton(cmdSearch, MapText("ค้นหา (F5)"))
   Call InitMainButton(cmdClear, MapText("เคลียร์ (F4)"))
   
End Sub


Private Sub Form_Load()
      OKClick = False
         m_HasActivate = False
   m_HasModify = False
   
   Set m_cusAcess = New Collection
   Set tempCusAcess = New CCommissionCustomerArea
   Set m_customerEP = New CARMas
   Set m_Rs = New ADODB.Recordset
   Set mm_Rs = New ADODB.Recordset
   
         m_HasActivate = False
   Call InitFormLayout
   Call EnableForm(Me, True)
End Sub

Private Sub GridEX1_DblClick()
   Call cmdEdit_Click
End Sub
Private Sub GridEX1_RowFormat(RowBuffer As GridEX20.JSRowData)
   RowBuffer.RowStyle = "Y"         ' RowBuffer.Value(8)
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
   Call m_customerEP.PopulateFromRS(3, m_Rs)

   Set tempCusAcess = GetCusAreaCom(m_cusAcess, Trim(m_customerEP.CUSCOD), False)
   If Not (tempCusAcess Is Nothing) Then
      Values(1) = tempCusAcess.COMMISSION_CUS_AREA_ID
      Values(4) = tempCusAcess.MASTER_AREA_NAME
   Else
      Values(1) = -1
      Values(4) = "     -"
   End If
   Values(2) = m_customerEP.CUSCOD
   Values(3) = m_customerEP.CUSNAM
   
   RowDelete = RowIndex + 1
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
   GridEX1.Height = ScaleHeight - GridEX1.Top - 720
   cmdEdit.Top = ScaleHeight - 580
   cmdOK.Top = ScaleHeight - 580
   cmdOK.Left = ScaleWidth - cmdOK.Width - 50
End Sub

Private Sub txtAreaName_hasChange()
   m_HasModify = True
End Sub

Private Sub cboAreaName_Click()
 m_HasModify = True
End Sub
