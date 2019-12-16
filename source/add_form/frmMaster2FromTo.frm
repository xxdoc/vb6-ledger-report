VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#2.0#0"; "THREED20.OCX"
Object = "{E684D8A3-716C-4E59-AA94-7144C04B0074}#1.1#0"; "GridEX20.ocx"
Begin VB.Form frmMaster2FromTo 
   BackColor       =   &H80000000&
   ClientHeight    =   8520
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11910
   Icon            =   "frmMaster2FromTo.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   8520
   ScaleWidth      =   11910
   StartUpPosition =   1  'CenterOwner
   WindowState     =   2  'Maximized
   Begin Threed.SSFrame SSFrame1 
      Height          =   8535
      Left            =   0
      TabIndex        =   12
      Top             =   0
      Width           =   11955
      _ExtentX        =   21087
      _ExtentY        =   15055
      _Version        =   131073
      PictureBackgroundStyle=   2
      Begin VB.ComboBox cboOrderType 
         Height          =   315
         Left            =   6600
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   1920
         Width           =   2625
      End
      Begin VB.ComboBox cboOrderBy 
         Height          =   315
         Left            =   1560
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   1950
         Width           =   2985
      End
      Begin Threed.SSPanel pnlHeader 
         Height          =   705
         Left            =   0
         TabIndex        =   13
         Top             =   0
         Width           =   11925
         _ExtentX        =   21034
         _ExtentY        =   1244
         _Version        =   131073
         PictureBackgroundStyle=   2
      End
      Begin GridEX20.GridEX GridEX1 
         Height          =   4995
         Left            =   240
         TabIndex        =   6
         Top             =   2520
         Width           =   11400
         _ExtentX        =   20108
         _ExtentY        =   8811
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
         Column(1)       =   "frmMaster2FromTo.frx":000C
         Column(2)       =   "frmMaster2FromTo.frx":00D4
         FormatStylesCount=   5
         FormatStyle(1)  =   "frmMaster2FromTo.frx":0178
         FormatStyle(2)  =   "frmMaster2FromTo.frx":02D4
         FormatStyle(3)  =   "frmMaster2FromTo.frx":0384
         FormatStyle(4)  =   "frmMaster2FromTo.frx":0438
         FormatStyle(5)  =   "frmMaster2FromTo.frx":0510
         ImageCount      =   0
         PrinterProperties=   "frmMaster2FromTo.frx":05C8
      End
      Begin prjLedgerReport.uctlDate uctlFromDate 
         Height          =   375
         Left            =   1560
         TabIndex        =   0
         Top             =   840
         Width           =   3855
         _ExtentX        =   6800
         _ExtentY        =   661
      End
      Begin prjLedgerReport.uctlDate uctlToDate 
         Height          =   375
         Left            =   1550
         TabIndex        =   1
         Top             =   1400
         Width           =   3855
         _ExtentX        =   6800
         _ExtentY        =   661
      End
      Begin VB.Label lblFromDate 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   435
         Left            =   240
         TabIndex        =   17
         Top             =   1080
         Width           =   1185
      End
      Begin VB.Label lblOrderType 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   435
         Left            =   5400
         TabIndex        =   16
         Top             =   2040
         Width           =   1095
      End
      Begin VB.Label lblToDate 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   375
         Left            =   240
         TabIndex        =   15
         Top             =   1560
         Width           =   1185
      End
      Begin VB.Label lblOrderBy 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   435
         Left            =   120
         TabIndex        =   14
         Top             =   2040
         Width           =   1275
      End
      Begin Threed.SSCommand cmdSearch 
         Height          =   525
         Left            =   9750
         TabIndex        =   4
         Top             =   960
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdClear 
         Height          =   525
         Left            =   9750
         TabIndex        =   5
         Top             =   1560
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
         Left            =   10095
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
         Left            =   8445
         TabIndex        =   10
         Top             =   7830
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         ButtonStyle     =   3
      End
   End
End
Attribute VB_Name = "frmMaster2FromTo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Private m_HasActivate As Boolean
Private m_MasterFromTo As CMaster2FromTo
Private m_TempMasterFromTo As CMaster2FromTo
Private m_Rs As ADODB.Recordset
Private m_TableName As String

Public DocumentType As MASTER_COMMISSION_AREA
Public OKClick As Boolean
Public HeaderText As String
Private Sub cmdAdd_Click()
Dim ItemCount As Long
Dim OKClick As Boolean

      frmAddEditMasterFromTo.DocumentType = DocumentType
      frmAddEditMasterFromTo.HeaderText = MapText("เพิ่มข้อมูลกำหนดค่า GP รายเดือน")
      frmAddEditMasterFromTo.ShowMode = SHOW_ADD
      Load frmAddEditMasterFromTo
      frmAddEditMasterFromTo.Show 1

      OKClick = frmAddEditMasterFromTo.OKClick

      Unload frmAddEditMasterFromTo
      Set frmAddEditMasterFromTo = Nothing

 '  If OKClick Then
      Call QueryData(True)
 '  End If
End Sub

Private Sub cmdClear_Click()
   uctlFromDate.ShowDate = -1
   uctlToDate.ShowDate = -1
   cboOrderBy.ListIndex = -1
   cboOrderType.ListIndex = -1
End Sub

Private Sub cmdDelete_Click()
Dim IsOK As Boolean
Dim ItemCount As Long
Dim IsCanLock As Boolean
Dim ID As Long

   If Not cmdDelete.Enabled Then
      Exit Sub
   End If


   If Not VerifyGrid(GridEX1.Value(1)) Then
      Exit Sub
   End If
   ID = GridEX1.Value(1)
   
   If Not ConfirmDelete(GridEX1.Value(2)) Then
      Exit Sub
   End If

   Call EnableForm(Me, False)
   m_MasterFromTo.MASTER_FROMTO_ID = ID
   If Not glbDaily.DeleteMaster2FromTo(m_MasterFromTo, IsOK, True, glbErrorLog) Then
      m_MasterFromTo.BILLING_DOC_ID = -1
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
   
   ID = Val(GridEX1.Value(1))
   
'   If DocumentType = COMMISSION_CHART Then
'      frmCommissionChart.ID = ID
'      frmCommissionChart.DocumentType = DocumentType
'      frmCommissionChart.HeaderText = MapText("แก้ไขข้อมูล" & Comissiontype2Text(DocumentType))
'      frmCommissionChart.ShowMode = SHOW_EDIT
'      Load frmCommissionChart
'      frmCommissionChart.Show 1
'
'      OKClick = frmCommissionChart.OKClick
'
'      Unload frmCommissionChart
'      Set frmCommissionChart = Nothing
'   ElseIf DocumentType = COMMISSION_TABLE Or DocumentType = RETURN_TABLE Then
      frmAddEditMasterFromTo.ID = ID
      frmAddEditMasterFromTo.DocumentType = DocumentType
      frmAddEditMasterFromTo.HeaderText = MapText("แก้ไขข้อมูลกำหนดค่า GP รายเดือน")
      frmAddEditMasterFromTo.ShowMode = SHOW_EDIT
      Load frmAddEditMasterFromTo
      frmAddEditMasterFromTo.Show 1

      OKClick = frmAddEditMasterFromTo.OKClick

      Unload frmAddEditMasterFromTo
      Set frmAddEditMasterFromTo = Nothing
'   ElseIf DocumentType = COMMISSION_TABLE_EX Then
'      frmAddEditMasterFromToEx.ID = ID
'      frmAddEditMasterFromToEx.DocumentType = DocumentType
'      frmAddEditMasterFromToEx.HeaderText = MapText("แก้ไขข้อมูล")
'      frmAddEditMasterFromToEx.ShowMode = SHOW_EDIT
'      Load frmAddEditMasterFromToEx
'      frmAddEditMasterFromToEx.Show 1
'
'      OKClick = frmAddEditMasterFromToEx.OKClick
'
'      Unload frmAddEditMasterFromToEx
'      Set frmAddEditMasterFromToEx = Nothing
'   End If
'
  ' If OKClick Then
      Call QueryData(True)
  ' End If
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
      
      Call InitCommissionOrderBy(cboOrderBy)
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
      
       m_MasterFromTo.MASTER_FROMTO_ID = -1
       m_MasterFromTo.VALID_FROM = uctlFromDate.ShowDate
       m_MasterFromTo.VALID_TO = uctlToDate.ShowDate
       m_MasterFromTo.MASTER_FROMTO_TYPE = DocumentType
       m_MasterFromTo.ORDER_BY = cboOrderBy.ItemData(Minus2Zero(cboOrderBy.ListIndex))
       m_MasterFromTo.ORDER_TYPE = cboOrderType.ItemData(Minus2Zero(cboOrderType.ListIndex))
      If Not glbDaily.QueryMaster2FromTo(m_MasterFromTo, m_Rs, ItemCount, IsOK, glbErrorLog) Then
         glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
         Call EnableForm(Me, True)
         Exit Sub
      End If
      
      Call InitGrid
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
   Col.Width = 2115
   Col.Caption = MapText("เลขที่")
      
   Set Col = GridEX1.Columns.Add '3
   Col.Width = ScaleWidth - 5900
   Col.Caption = MapText("รายละเอียด")

   Set Col = GridEX1.Columns.Add '4
   Col.Width = 1500
   Col.Caption = MapText("วันที่เริ่มใช้")
   
   Set Col = GridEX1.Columns.Add '5
   Col.Width = 1500
   Col.Caption = MapText("วันที่สิ้นสุด")
      
   GridEX1.ItemCount = 0
End Sub

Private Sub InitFormLayout()
   pnlHeader.PictureBackground = LoadPicture(glbParameterObj.NormalForm1)
   SSFrame1.PictureBackground = LoadPicture(glbParameterObj.NormalForm1)
   Me.Caption = HeaderText
   
   Call InitGrid
   
   Call InitNormalLabel(lblFromDate, MapText("วันที่เริ่มใช้"))
   Call InitNormalLabel(lblToDate, MapText("วันที่สิ้นสุด"))
   Call InitNormalLabel(lblOrderBy, MapText("เรียงตาม"))
   Call InitNormalLabel(lblOrderType, MapText("เรียงจาก"))
   
   Call InitCombo(cboOrderBy)
   Call InitCombo(cboOrderType)
   
   Me.Picture = LoadPicture(glbParameterObj.MainPicture)
   
   pnlHeader.Font.Name = GLB_FONT
   pnlHeader.Font.Bold = True
   pnlHeader.Font.Size = 19
   pnlHeader.Caption = HeaderText
   pnlHeader.Caption = HeaderText
   
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
   m_TableName = "MASTER_FROMTO"
   m_HasActivate = False
   
   Set m_MasterFromTo = New CMaster2FromTo
   Set m_TempMasterFromTo = New CMaster2FromTo
   Set m_Rs = New ADODB.Recordset
      
   m_HasActivate = False
   Call InitFormLayout
   Call EnableForm(Me, True)
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set m_MasterFromTo = Nothing
End Sub
Private Sub GridEX1_ColResize(ByVal ColIndex As Integer, ByVal NewColWidth As Long, ByVal Cancel As GridEX20.JSRetBoolean)
   'debug.print ColIndex & " " & NewColWidth
End Sub
Private Sub GridEX1_DblClick()
   Call cmdEdit_Click
End Sub
Private Sub GridEX1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim oMenu As cPopupMenu
Dim lMenuChosen As Long
Dim TempID1 As Long
Dim Ms As CMaster2FromTo
Dim IsOK As Boolean
Dim OKClick As Boolean

   If GridEX1.ItemCount <= 0 Then
         Exit Sub
   End If
   
   TempID1 = GridEX1.Value(1)
   
   If Button = 2 Then
      Set oMenu = New cPopupMenu
      lMenuChosen = oMenu.Popup("COPY")
      If lMenuChosen = 0 Then
         Exit Sub
      End If
      Set oMenu = Nothing
   Else
      Exit Sub
   End If
   
   Call EnableForm(Me, False)
   If lMenuChosen = 1 Then
      Set Ms = New CMaster2FromTo
      Ms.MASTER_FROMTO_ID = TempID1
      Call glbDaily.CopyMaster2FromTo(Ms, IsOK, True, glbErrorLog)
      Call QueryData(True)
      Set Ms = Nothing
      
      Set Ms = Nothing
   End If
   
   Call EnableForm(Me, True)
End Sub
Private Sub GridEX1_RowFormat(RowBuffer As GridEX20.JSRowData)
   'RowBuffer.RowStyle = RowBuffer.Value(6)
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
   Call m_TempMasterFromTo.PopulateFromRS(1, m_Rs)

   Values(1) = m_TempMasterFromTo.MASTER_FROMTO_ID
   Values(2) = m_TempMasterFromTo.MASTER_FROMTO_NO
   Values(3) = m_TempMasterFromTo.MASTER_FROMTO_DESC
   Values(4) = DateToStringExtEx2(m_TempMasterFromTo.VALID_FROM)
   Values(5) = DateToStringExtEx2(m_TempMasterFromTo.VALID_TO)
   
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
Private Sub GridEX1_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = DUMMY_KEY Then
      Call cmdExit_Click
      KeyCode = 0
   ElseIf Shift = 0 And KeyCode = 113 Then
      Call cmdOK_Click
      KeyCode = 0
   End If
End Sub

