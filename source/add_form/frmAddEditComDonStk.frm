VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#2.0#0"; "THREED20.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Object = "{E684D8A3-716C-4E59-AA94-7144C04B0074}#1.1#0"; "GridEX20.ocx"
Begin VB.Form frmAddEditComDonStk 
   ClientHeight    =   8490
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11880
   Icon            =   "frmAddEditComDonStk.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   Moveable        =   0   'False
   NegotiateMenus  =   0   'False
   ScaleHeight     =   8490
   ScaleWidth      =   11880
   StartUpPosition =   1  'CenterOwner
   WindowState     =   2  'Maximized
   Begin Threed.SSFrame SSFrame1 
      Height          =   8520
      Left            =   0
      TabIndex        =   10
      Top             =   0
      Width           =   11895
      _ExtentX        =   20981
      _ExtentY        =   15028
      _Version        =   131073
      PictureBackgroundStyle=   2
      Begin MSComDlg.CommonDialog dlgAdd 
         Left            =   0
         Top             =   0
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin Threed.SSPanel pnlHeader 
         Height          =   705
         Left            =   0
         TabIndex        =   11
         Top             =   0
         Width           =   11925
         _ExtentX        =   21034
         _ExtentY        =   1244
         _Version        =   131073
         PictureBackgroundStyle=   2
      End
      Begin prjLedgerReport.uctlDate uctlFromDate 
         Height          =   375
         Left            =   3840
         TabIndex        =   2
         Top             =   1800
         Width           =   3855
         _ExtentX        =   6800
         _ExtentY        =   661
      End
      Begin prjLedgerReport.uctlTextBox txtMasIncentiveNo 
         Height          =   450
         Left            =   3840
         TabIndex        =   0
         Top             =   840
         Width           =   2895
         _ExtentX        =   5106
         _ExtentY        =   873
      End
      Begin prjLedgerReport.uctlDate uctlToDate 
         Height          =   375
         Left            =   3840
         TabIndex        =   3
         Top             =   2280
         Width           =   3855
         _ExtentX        =   6800
         _ExtentY        =   661
      End
      Begin GridEX20.GridEX GridEX1 
         Height          =   4725
         Left            =   120
         TabIndex        =   4
         Top             =   2880
         Width           =   11595
         _ExtentX        =   20452
         _ExtentY        =   8334
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
         Column(1)       =   "frmAddEditComDonStk.frx":27A2
         Column(2)       =   "frmAddEditComDonStk.frx":286A
         FormatStylesCount=   5
         FormatStyle(1)  =   "frmAddEditComDonStk.frx":290E
         FormatStyle(2)  =   "frmAddEditComDonStk.frx":2A6A
         FormatStyle(3)  =   "frmAddEditComDonStk.frx":2B1A
         FormatStyle(4)  =   "frmAddEditComDonStk.frx":2BCE
         FormatStyle(5)  =   "frmAddEditComDonStk.frx":2CA6
         ImageCount      =   0
         PrinterProperties=   "frmAddEditComDonStk.frx":2D5E
      End
      Begin prjLedgerReport.uctlTextBox txtMasterIncentiveDesc 
         Height          =   450
         Left            =   3840
         TabIndex        =   1
         Top             =   1320
         Width           =   6015
         _ExtentX        =   10610
         _ExtentY        =   794
      End
      Begin VB.Label lblMasIncentiveDesc 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   2040
         TabIndex        =   15
         Top             =   1320
         Width           =   1365
      End
      Begin Threed.SSCommand cmdExit 
         Cancel          =   -1  'True
         Height          =   525
         Left            =   10080
         TabIndex        =   9
         Top             =   7800
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   926
         _Version        =   131073
         ButtonStyle     =   3
      End
      Begin VB.Label lblFromDate 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   435
         Left            =   1680
         TabIndex        =   14
         Top             =   1800
         Width           =   1755
      End
      Begin VB.Label lblToDate 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   435
         Left            =   2160
         TabIndex        =   13
         Top             =   2280
         Width           =   1365
      End
      Begin VB.Label lblMasIncentiveNo 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   435
         Left            =   1680
         TabIndex        =   12
         Top             =   840
         Width           =   1755
      End
      Begin Threed.SSCommand cmdOK 
         Height          =   525
         Left            =   8400
         TabIndex        =   8
         Top             =   7800
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmAddEditComDonStk.frx":2F36
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdEdit 
         Height          =   525
         Left            =   1770
         TabIndex        =   6
         Top             =   7830
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   926
         _Version        =   131073
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdAdd 
         Height          =   525
         Left            =   150
         TabIndex        =   5
         Top             =   7830
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmAddEditComDonStk.frx":3250
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdDelete 
         Height          =   525
         Left            =   3420
         TabIndex        =   7
         Top             =   7830
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmAddEditComDonStk.frx":356A
         ButtonStyle     =   3
      End
   End
End
Attribute VB_Name = "frmAddEditComDonStk"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Const ROOT_TREE = "Root"
Private m_HasActivate As Boolean
Private m_HasModify As Boolean
Private m_Rs As ADODB.Recordset

Private m_PromoteYear As CComDonStkMaster
Private m_comDonStk As CComDonStk        ' incen
Private cm5_Rs As ADODB.Recordset    '' incentive

Public HeaderText As String
Public ShowMode As SHOW_MODE_TYPE
Public OKClick As Boolean

Public COMDONSTK_ID As Long
Public STKCOD As String
Public STKDES As String
Public MASTER_COMDONSTK_ID As Long
Public MASTER_COMDONSTK_NO As String
Public Flag As String

Public VALID_FROM As String
Public VALID_TO As String

Dim ItemCount As Long
Dim itemCountGrid5 As Long     ' incen

Private m_TableName As String
Private FileName As String
Private m_SumUnit As Double

Public Sub QueryData(Flag As Boolean)
Dim IsOK As Boolean

                     IsOK = True
                     If Flag Then
                            Call EnableForm(Me, False)
                            
                           If MASTER_COMDONSTK_ID = 0 Then
                               If Not glbDaily.QueryMasterComDonStk(m_PromoteYear, m_Rs, ItemCount, IsOK, glbErrorLog) Then
                                  glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
                                  Call EnableForm(Me, True)
                                  Exit Sub
                                End If
                                 Call m_PromoteYear.PopulateFromRS(1, m_Rs)
                                 MASTER_COMDONSTK_ID = m_PromoteYear.MASTER_COMDONSTK_ID
                            End If
                            
                            m_PromoteYear.MASTER_COMDONSTK_ID = MASTER_COMDONSTK_ID
                            m_PromoteYear.VALID_FROM = -1
                            m_PromoteYear.VALID_TO = -1
                            If Not glbDaily.QueryMasterComDonStk(m_PromoteYear, m_Rs, ItemCount, IsOK, glbErrorLog) Then
                                  glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
                                  Call EnableForm(Me, True)
                                  Exit Sub
                            End If
                     End If

                     If ItemCount > 0 Then
                        Call m_PromoteYear.PopulateFromRS(1, m_Rs)
                        txtMasIncentiveNo.Text = m_PromoteYear.MASTER_COMDONSTK_NO    'ยังอยู่ในโหมด edit
                        uctlFromDate.ShowDate = m_PromoteYear.VALID_FROM
                        uctlToDate.ShowDate = m_PromoteYear.VALID_TO
                        txtMasterIncentiveDesc.Text = m_PromoteYear.MASTER_COMDONSTK_DESC
                     End If

                  If Not IsOK Then
                     glbErrorLog.ShowUserError
                     Call EnableForm(Me, True)
                     Exit Sub
                  End If
   
      Call InitGrid4
      GridEX1.ItemCount = CountItem(m_PromoteYear.Details)
      GridEX1.Rebind
      
   Call EnableForm(Me, True)
End Sub

Private Function SaveData() As Boolean
Dim IsOK As Boolean

   If Not VerifyTextControl(lblMasIncentiveNo, txtMasIncentiveNo, False) Then
      Exit Function
   End If
   
   If Not VerifyTextControl(lblMasIncentiveDesc, txtMasterIncentiveDesc, False) Then
      Exit Function
   End If
   
   If Not VerifyDate(lblFromDate, uctlFromDate, False) Then
      Exit Function
   End If
   If Not VerifyDate(lblToDate, uctlToDate, False) Then
      Exit Function
   End If
   
   If Not CheckUniqueNs(EXPORT_UNIQUE, txtMasIncentiveNo.Text, MASTER_COMDONSTK_ID) Then
      glbErrorLog.LocalErrorMsg = MapText("มีข้อมูล") & " " & txtMasIncentiveNo.Text & " " & MapText("อยู่ในระบบแล้ว")
      glbErrorLog.ShowUserError
      Call txtMasIncentiveNo.SetFocus
      Exit Function
   End If
   
   If Not m_HasModify Then
      SaveData = True
      Exit Function
   End If

m_PromoteYear.ShowMode = ShowMode                        ' ตรงนี้ ตอนบันทึก
    m_PromoteYear.MASTER_COMDONSTK_ID = MASTER_COMDONSTK_ID
    m_PromoteYear.MASTER_COMDONSTK_NO = txtMasIncentiveNo.Text
    m_PromoteYear.VALID_FROM = uctlFromDate.ShowDate
    m_PromoteYear.VALID_TO = uctlToDate.ShowDate
    m_PromoteYear.MASTER_COMDONSTK_DESC = txtMasterIncentiveDesc.Text
   
   Call EnableForm(Me, False)
   
   If Not glbDaily.AddEditMasterComDonStk(m_PromoteYear, IsOK, True, glbErrorLog) Then
      glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
      SaveData = False
      Call EnableForm(Me, True)
      Exit Function
   End If
   
   If Not IsOK Then
      Call EnableForm(Me, True)
      glbErrorLog.ShowUserError
      Exit Function
   End If

   Call EnableForm(Me, True)
   SaveData = True
End Function

Private Sub cmdAdd_Click()    ' เมื่อกดเพิ่มแต่ละ case 3 case
Dim OKClick As Boolean
Dim IsOK As Boolean

If Not cmdAdd.Enabled Then
      Exit Sub
   End If
   
  If m_PromoteYear.MASTER_COMDONSTK_ID <= 0 Then
      glbErrorLog.LocalErrorMsg = MapText("กรุณาทำการบันทึกข้อมูลให้เรียบร้อยก่อน")
      glbErrorLog.ShowUserError
      Exit Sub
   End If
   
   Call EnableForm(Me, False)
   
   OKClick = False
   
      Set frmAddEditComTypeDonStk.ParentForm = Me
      Set frmAddEditComTypeDonStk.TempCollection = m_PromoteYear.Details
       frmAddEditComTypeDonStk.MASTER_COMDONSTK_ID = MASTER_COMDONSTK_ID
      frmAddEditComTypeDonStk.ShowMode = SHOW_ADD
      frmAddEditComTypeDonStk.HeaderText = MapText("เพิ่มสินค้าที่ไม่คิด Commercial #1")
      frmAddEditComTypeDonStk.itemCountGrid = itemCountGrid5
      Set frmAddEditComTypeDonStk.ParentForm = Me
      Load frmAddEditComTypeDonStk
     frmAddEditComTypeDonStk.Show 1
   
      OKClick = frmAddEditComTypeDonStk.OKClick
   
      Unload frmAddEditComTypeDonStk
      Set frmAddEditComTypeDonStk = Nothing

   
   If OKClick Then
      Call RefreshGrid
   End If

   If OKClick Then
      m_HasModify = True
   End If
End Sub

Private Sub cmdDelete_Click()
Dim ID1 As Long
Dim ID2 As Long

   If Not cmdDelete.Enabled Then
      Exit Sub
   End If
   
   If Not VerifyGrid(GridEX1.Value(1)) Then
      Exit Sub
   End If

   If Not ConfirmDelete(GridEX1.Value(3) & " " & GridEX1.Value(4)) Then
      Exit Sub
   End If

   ID2 = GridEX1.Value(2)
   ID1 = GridEX1.Value(1)

      If ID1 <= 0 Then
         m_PromoteYear.Details.Remove (ID2)
      Else
         m_PromoteYear.Details.ITEM(ID2).Flag = "D"
      End If
      m_HasModify = True
  
   Call RefreshGrid
End Sub

Private Sub cmdEdit_Click()
Dim IsOK As Boolean
Dim IsCanLock As Boolean
Dim ID As Long
Dim OKClick As Boolean

 If Not cmdEdit.Enabled Then
      Exit Sub
   End If
      
   If Not VerifyGrid(GridEX1.Value(1)) Then
      Exit Sub
   End If

   ID = Val(GridEX1.Value(2))   ' COMDONSTK_ID
   OKClick = False
   
         Set frmAddEditComTypeDonStk.ParentForm = Me
         Set frmAddEditComTypeDonStk.TempCollection = m_PromoteYear.Details
         frmAddEditComTypeDonStk.ID = ID                        ' ID ของ คอเล็คคชั้น
         frmAddEditComTypeDonStk.MASTER_COMDONSTK_ID = MASTER_COMDONSTK_ID
         frmAddEditComTypeDonStk.HeaderText = MapText("แก้ไขสินค้าที่ไม่คิด Commercial #1")
         frmAddEditComTypeDonStk.ShowMode = SHOW_EDIT
         frmAddEditComTypeDonStk.itemCountGrid = itemCountGrid5
         Load frmAddEditComTypeDonStk
         frmAddEditComTypeDonStk.Show 1
            
            OKClick = frmAddEditComTypeDonStk.OKClick
            
            Unload frmAddEditComTypeDonStk
            Set frmAddEditComTypeDonStk = Nothing

   If OKClick Then
      Call RefreshGrid
      m_HasModify = True
   End If
End Sub
Private Sub cmdOK_Click()
Dim oMenu As cPopupMenu
Dim lMenuChosen  As Long


   Set oMenu = New cPopupMenu
   lMenuChosen = oMenu.Popup("บันทึก", "-", "บันทึกและออก")
   If lMenuChosen = 0 Then
      Exit Sub
   End If
   
   If lMenuChosen = 1 Then
      If Not SaveData Then
         Exit Sub
      End If
      
      ShowMode = SHOW_EDIT
      MASTER_COMDONSTK_ID = m_PromoteYear.MASTER_COMDONSTK_ID
      m_PromoteYear.QueryFlag = 1
      QueryData (True)
      m_HasModify = False
      OKClick = True
   ElseIf lMenuChosen = 3 Then
      If Not SaveData Then
         Exit Sub
      End If
      
      OKClick = True
      Unload Me
   End If
End Sub

Private Sub Form_Activate()
   If Not m_HasActivate Then
      m_HasActivate = True
      Me.Refresh
      DoEvents
            
      Call EnableForm(Me, False)
      If (ShowMode = SHOW_EDIT) Or (ShowMode = SHOW_VIEW_ONLY) Then
         m_PromoteYear.QueryFlag = 1
         Call QueryData(True)
      ElseIf ShowMode = SHOW_ADD Then
         m_PromoteYear.QueryFlag = 0
         Call QueryData(False)
      End If
      
      Call EnableForm(Me, True)
      m_HasModify = False
   End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
Static InUsed As Long
   If InUsed = 1 Then
      Exit Sub
   End If
   InUsed = 1
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
'   ElseIf Shift = 0 And KeyCode = 121 Then
'      Call cmdSave_Click
'      KeyCode = 0
   End If
   InUsed = 0
End Sub

Private Sub Form_Unload(Cancel As Integer)
   If m_Rs.State = adStateOpen Then
      m_Rs.Close
   End If
   Set m_Rs = Nothing
   Set m_PromoteYear = Nothing

End Sub

Private Sub GridEX1_ColResize(ByVal ColIndex As Integer, ByVal NewColWidth As Long, ByVal Cancel As GridEX20.JSRetBoolean)
   'debug.print ColIndex & " " & NewColWidth
End Sub '

Private Sub InitGrid4()
Dim Col As JSColumn

   GridEX1.Columns.Clear
   GridEX1.BackColor = GLB_GRID_COLOR
   GridEX1.ItemCount = 0
   GridEX1.BackColorHeader = GLB_GRIDHD_COLOR
   GridEX1.ColumnHeaderFont.Bold = True
   GridEX1.ColumnHeaderFont.Name = GLB_FONT
   GridEX1.TabKeyBehavior = jgexControlNavigation
   
   Set Col = GridEX1.Columns.Add '1
   Col.Width = 0
   Col.Caption = "ID"
   
   Set Col = GridEX1.Columns.Add '1
   Col.Width = 0
   Col.Caption = "Realid"
   
   Set Col = GridEX1.Columns.Add '2
   Col.Width = 3700
   Col.Caption = MapText("เลขที่สินค้า")

   Set Col = GridEX1.Columns.Add '3
   Col.Width = 7350
   Col.Caption = MapText("ชื่อสินค้า")
   
End Sub

Private Sub InitFormLayout()
   pnlHeader.PictureBackground = LoadPicture(glbParameterObj.NormalForm1)
   SSFrame1.PictureBackground = LoadPicture(glbParameterObj.NormalForm1)
   
   Me.Caption = HeaderText
   pnlHeader.Caption = Me.Caption
'
   Call InitNormalLabel(lblMasIncentiveNo, MapText("เลขที่"))
      Call InitNormalLabel(lblMasIncentiveDesc, MapText("รายละเอียด"))
   Call InitNormalLabel(lblFromDate, MapText("เริ่มใช้วันที่"))
      Call InitNormalLabel(lblToDate, MapText("ถึงวันที่"))
      
      Call txtMasterIncentiveDesc.SetTextLenType(TEXT_STRING, glbSetting.DESC_TYPE)
  Call txtMasIncentiveNo.SetTextLenType(TEXT_STRING, glbSetting.CODE_TYPE)
   
   Me.Picture = LoadPicture(glbParameterObj.MainPicture)

   pnlHeader.Font.Name = GLB_FONT
   pnlHeader.Font.Bold = True
   pnlHeader.Font.Size = 19


   cmdOK.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdAdd.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdEdit.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdDelete.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdExit.Picture = LoadPicture(glbParameterObj.NormalButton1)
   
   Call InitMainButton(cmdExit, MapText("ยกเลิก (ESC)"))
   Call InitMainButton(cmdOK, MapText("ตกลง (F2)"))
   Call InitMainButton(cmdAdd, MapText("เพิ่ม (F7)"))
   Call InitMainButton(cmdEdit, MapText("แก้ไข (F3)"))
   Call InitMainButton(cmdDelete, MapText("ลบ (F6)"))

   Call InitGrid4
   
End Sub

Private Sub cmdExit_Click()
   If Not ConfirmExit(m_HasModify) Then
      Exit Sub
   End If

 '  OKClick = False
   Unload Me
End Sub

Private Sub Form_Load()
   OKClick = False
    Call InitFormLayout
    
   m_HasActivate = False
   m_HasModify = False
   Set m_Rs = New ADODB.Recordset
   Set m_PromoteYear = New CComDonStkMaster
End Sub

Private Sub GridEX1_DblClick()
   Call cmdEdit_Click
End Sub

Private Sub GridEX1_RowFormat(RowBuffer As GridEX20.JSRowData)
      RowBuffer.RowStyle = RowBuffer.Value(1)
End Sub

Private Sub GridEX1_UnboundReadData(ByVal RowIndex As Long, ByVal Bookmark As Variant, ByVal Values As GridEX20.JSRowData)
On Error GoTo ErrorHandler
Dim RealIndex As Long

   glbErrorLog.ModuleName = Me.Name
   glbErrorLog.RoutineName = "UnboundReadData"

      If m_PromoteYear.Details Is Nothing Then
         Exit Sub
      End If
   
      If RowIndex <= 0 Then
         Exit Sub
      End If
   
      Dim CR As CComDonStk
      If m_PromoteYear.Details.Count <= 0 Then
         Exit Sub
      End If
      Set CR = GetItem(m_PromoteYear.Details, RowIndex, RealIndex)
      If CR Is Nothing Then
         Exit Sub
      End If
 
      Values(1) = CR.COMDONSTK_ID
      Values(2) = RealIndex
      Values(3) = CR.STKCOD
      Values(4) = CR.STKDES

   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub

Private Sub txtMasIncentiveNo_Change()
   m_HasModify = True
End Sub
 Private Sub txtMasterIncentiveDesc_Change()
   m_HasModify = True
End Sub

Private Sub uctlToDate_HasChange()
   m_HasModify = True
End Sub

Private Sub uctlFromDate_HasChange()
   m_HasModify = True
End Sub

Public Sub RefreshGrid()
   GridEX1.ItemCount = CountItem(m_PromoteYear.Details)
   GridEX1.Rebind
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

