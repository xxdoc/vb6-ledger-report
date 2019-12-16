VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#2.0#0"; "THREED20.OCX"
Object = "{E684D8A3-716C-4E59-AA94-7144C04B0074}#1.1#0"; "GridEX20.ocx"
Begin VB.Form frmGoodsDetail 
   BackColor       =   &H80000000&
   ClientHeight    =   8520
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   18630
   Icon            =   "frmGoodsDetail.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   8520
   ScaleWidth      =   18630
   StartUpPosition =   1  'CenterOwner
   WindowState     =   2  'Maximized
   Begin Threed.SSFrame SSFrame1 
      Height          =   8535
      Left            =   0
      TabIndex        =   8
      Top             =   0
      Width           =   18675
      _ExtentX        =   32941
      _ExtentY        =   15055
      _Version        =   131073
      PictureBackgroundStyle=   2
      Begin Threed.SSPanel pnlHeader 
         Height          =   705
         Left            =   0
         TabIndex        =   9
         Top             =   0
         Width           =   18645
         _ExtentX        =   32888
         _ExtentY        =   1244
         _Version        =   131073
         PictureBackgroundStyle=   2
      End
      Begin GridEX20.GridEX GridEX1 
         Height          =   5295
         Left            =   180
         TabIndex        =   4
         Top             =   2400
         Width           =   18225
         _ExtentX        =   32147
         _ExtentY        =   9340
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
         Column(1)       =   "frmGoodsDetail.frx":27A2
         Column(2)       =   "frmGoodsDetail.frx":286A
         FormatStylesCount=   5
         FormatStyle(1)  =   "frmGoodsDetail.frx":290E
         FormatStyle(2)  =   "frmGoodsDetail.frx":2A6A
         FormatStyle(3)  =   "frmGoodsDetail.frx":2B1A
         FormatStyle(4)  =   "frmGoodsDetail.frx":2BCE
         FormatStyle(5)  =   "frmGoodsDetail.frx":2CA6
         ImageCount      =   0
         PrinterProperties=   "frmGoodsDetail.frx":2D5E
      End
      Begin prjLedgerReport.uctlTextBox txtMasterCode 
         Height          =   375
         Left            =   9960
         TabIndex        =   0
         Top             =   1440
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   661
      End
      Begin prjLedgerReport.uctlTextBox txtMasterName 
         Height          =   375
         Left            =   13080
         TabIndex        =   1
         Tag             =   "2"
         Top             =   1440
         Width           =   2535
         _ExtentX        =   8281
         _ExtentY        =   661
      End
      Begin prjLedgerReport.uctlTextBox txtSaveCode 
         Height          =   375
         Left            =   2160
         TabIndex        =   12
         Top             =   1200
         Width           =   3735
         _ExtentX        =   6588
         _ExtentY        =   661
      End
      Begin prjLedgerReport.uctlTextBox txtSaveName 
         Height          =   375
         Left            =   2160
         TabIndex        =   13
         Tag             =   "2"
         Top             =   1680
         Width           =   5295
         _ExtentX        =   9340
         _ExtentY        =   661
      End
      Begin VB.Label lblSaveCode 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   435
         Left            =   360
         TabIndex        =   15
         Top             =   1200
         Width           =   1695
      End
      Begin VB.Label lblSaveName 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   495
         Left            =   240
         TabIndex        =   14
         Top             =   1800
         Width           =   1815
      End
      Begin VB.Shape Shape1 
         BorderColor     =   &H8000000A&
         Height          =   975
         Left            =   8880
         Top             =   1200
         Width           =   9615
      End
      Begin VB.Label lblMasterCode 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   435
         Left            =   8160
         TabIndex        =   11
         Top             =   1440
         Width           =   1695
      End
      Begin Threed.SSCommand cmdClear 
         Height          =   525
         Left            =   17160
         TabIndex        =   3
         Top             =   1440
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   926
         _Version        =   131073
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdSearch 
         Height          =   525
         Left            =   15840
         TabIndex        =   2
         Top             =   1440
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmGoodsDetail.frx":2F36
         ButtonStyle     =   3
      End
      Begin VB.Label lblMasterName 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   375
         Left            =   11880
         TabIndex        =   10
         Top             =   1440
         Width           =   1095
      End
      Begin Threed.SSCommand cmdOK 
         Height          =   525
         Left            =   15000
         TabIndex        =   6
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
         TabIndex        =   5
         Top             =   7800
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   926
         _Version        =   131073
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdExit 
         Height          =   525
         Left            =   16800
         TabIndex        =   7
         Top             =   7800
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmGoodsDetail.frx":3250
         ButtonStyle     =   3
      End
   End
End
Attribute VB_Name = "frmGoodsDetail"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private m_HasActivate As Boolean
Private m_HasModify As Boolean
Private m_goodsDetail As Collection
Private tempGoodsDetail As CGoodsMaster
Private temp_getGoodsDetail As CGoodsDetail

Private m_stcrdEP As CStmas
Public HeaderText As String

Private m_Rs As ADODB.Recordset
Private mm_Rs As ADODB.Recordset
Public OKClick As Boolean
Public GOODS_MASTER_ID As Long
Public GOODS_MASTER_CODE As String
Public GOODS_MASTER_NAME As String

Public ShowMode As SHOW_MODE_TYPE

'Public GOODS_MASTER_ID As Long
Dim RowDelete As Long

Private Sub cmdClear_Click()
   txtMasterCode.Text = ""
   txtMasterName.Text = ""
End Sub
Private Sub cmdEdit_Click()
Dim IsOK As Boolean
Dim ItemCount As Long
Dim IsCanLock As Boolean
Dim ID As Long
Dim STKCOD As String
Dim OKClick As Boolean
   
   If GOODS_MASTER_ID <= 0 Then
      glbErrorLog.LocalErrorMsg = MapText("กรุณาทำการบันทึกข้อมูลให้เรียบร้อยก่อน")
      glbErrorLog.ShowUserError
      Exit Sub
   End If
   
   If Not VerifyGrid(GridEX1.Value(1)) Then
      Exit Sub
   End If

   ID = GridEX1.Value(1)
   STKCOD = GridEX1.Value(2)
      
   frmAddEditGoodsDetail.GOODS_DETAIL_ID = ID
   frmAddEditGoodsDetail.HeaderText = MapText("แก้ไขข้อมูลกลุ่มสินค้า")
   frmAddEditGoodsDetail.GOODS_MASTER_ID = GOODS_MASTER_ID
   frmAddEditGoodsDetail.STKCOD = STKCOD
   Load frmAddEditGoodsDetail
   frmAddEditGoodsDetail.Show 1

   OKClick = frmAddEditGoodsDetail.OKClick

   Unload frmAddEditGoodsDetail
   Set frmAddEditGoodsDetail = Nothing

   If OKClick Then
      Set m_goodsDetail = Nothing
      Set m_goodsDetail = New Collection
      Call LoadStkcodFromMasterID(Nothing, GOODS_MASTER_ID, m_goodsDetail)
      Call QueryData(True)
   End If

End Sub

Private Function SaveData() As Boolean
Dim IsOK As Boolean

   If Not VerifyTextControl(lblSaveCode, txtSaveCode, False) Then
      Exit Function
   End If
   
   If Not VerifyTextControl(lblSaveName, txtSaveName, False) Then
      Exit Function
   End If
   
   If Not m_HasModify Then
      SaveData = True
      Exit Function
   End If
   
'   If Not CheckUniqueNs(MASTER_FT_UNIQUE, txtMasterFromToNo.Text, ID) Then
'      glbErrorLog.LocalErrorMsg = MapText("มีข้อมูล") & " " & txtMasterFromToNo.Text & " " & MapText("อยู่ในระบบแล้ว")
'      glbErrorLog.ShowUserError
'      Call txtMasterFromToNo.SetFocus
'      Exit Function
'   End If
   
   tempGoodsDetail.AddEditMode = ShowMode
   tempGoodsDetail.GOODS_MASTER_ID = GOODS_MASTER_ID
   tempGoodsDetail.GOODS_MASTER_CODE = txtSaveCode.Text
   tempGoodsDetail.GOODS_MASTER_NAME = txtSaveName.Text

   
'   Call m_MasterFromTo.SetFieldValue("INCLUDE_SUB_FLAG", Check2Flag(ChkIncludeSub.Value))
'   Call m_MasterFromTo.SetFieldValue("INCLUDE_SUB_PERCENT", Val(txtIncludeSub.Text))
'   Call m_MasterFromTo.SetFieldValue("MULTIPLE_FLAG", CheckSSoptionToString(SSOption1.Value))
'   Call m_MasterFromTo.SetFieldValue("MULTIPLE_PERCENT", Val(txtValue1.Text))
'   Call m_MasterFromTo.SetFieldValue("STEP_FLAG", CheckSSoptionToString(SSOption2.Value))
'   Call m_MasterFromTo.SetFieldValue("TIER_FLAG", CheckSSoptionToString(SSOption3.Value))
      
   Call EnableForm(Me, False)
   
   If Not glbDaily.AddEditGoodsMaster(tempGoodsDetail, IsOK, True, glbErrorLog) Then
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

Private Sub cmdOK_Click()
Dim oMenu As cPopupMenu
Dim lMenuChosen  As Long
Set oMenu = New cPopupMenu
   lMenuChosen = oMenu.Popup("บันทึก", "-", "บันทึกและออกจากหน้าจอ")
   If lMenuChosen = 0 Then
      Exit Sub
   End If
   
   If lMenuChosen = 1 Then
      If Not SaveData Then
         Exit Sub
      End If
      
      ShowMode = SHOW_EDIT
         GOODS_MASTER_ID = tempGoodsDetail.GOODS_MASTER_ID
         tempGoodsDetail.QueryFlag = 1
         QueryData (True)
         m_HasModify = False
   ElseIf lMenuChosen = 3 Then
      If Not SaveData Then
         Exit Sub
      End If
      
      OKClick = True
      Unload Me
   End If
End Sub

Private Sub cmdExit_Click()
   If Not ConfirmExit(m_HasModify) Then
      Exit Sub
   End If
   
   OKClick = False
   Unload Me
End Sub

Private Sub cmdSearch_Click()

   'ทุกครั้ง เพื่อเช็คว่าติ๊กหรือไม่ = ใช้การไม่ได้อยู่ดี เพราะลูกค้าจาก EP มาทั้งหมด
'      Set m_goodsDetail = Nothing
'      Set m_goodsDetail = New Collection
'      Call LoadCusNonAreaCom(Check2Flag(ChkNonAreaFLag.Value), GOODS_MASTER_ID, m_goodsDetail)

   Call QueryData(True)
End Sub

Private Sub Form_Activate()
   If Not m_HasActivate Then
      m_HasActivate = True
      
      Call LoadStkcodFromMasterID(Nothing, GOODS_MASTER_ID, m_goodsDetail)

      If ShowMode = SHOW_EDIT Then
         txtSaveCode.Text = GOODS_MASTER_CODE
         txtSaveName.Text = GOODS_MASTER_NAME
'         m_stcrdEP.STKCOD = ""
'         m_stcrdEP.CUSNAM = ""
'         m_stcrdEP.CUSTYP = ""
'         m_stcrdEP.SLMCOD = ""
 '        Call QueryData(True)
      End If
      
      Call QueryData(True)
      
   End If
End Sub
Private Sub QueryData(Flag As Boolean)
Dim IsOK As Boolean
Dim ItemCount As Long
Dim m_ItemCount As Long
Dim Temp As Long
                     
   Set m_stcrdEP = Nothing
   Set m_stcrdEP = New CStmas
   
   m_stcrdEP.STKCOD = txtMasterCode.Text
   m_stcrdEP.STKDES = txtMasterName.Text

   If Not glbDaily.QueryGoodsEP(m_stcrdEP, m_Rs, ItemCount, IsOK, glbErrorLog) Then
      glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
      Call EnableForm(Me, True)
      Exit Sub
   End If
   Call m_stcrdEP.PopulateFromRS(1, m_Rs)
   
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
   Col.Caption = MapText("รหัสสินค้า")
   
   Set Col = GridEX1.Columns.Add '2
   Col.Width = 5000
   Col.Caption = MapText("สินค้า")
   
   Set Col = GridEX1.Columns.Add '3
   Col.Width = 4500
    Col.Caption = MapText("กลุ่ม")
   
   GridEX1.ItemCount = 0
End Sub
Private Sub InitFormLayout()
   pnlHeader.PictureBackground = LoadPicture(glbParameterObj.NormalForm1)
   SSFrame1.PictureBackground = LoadPicture(glbParameterObj.NormalForm1)
   
   Me.Caption = MapText("ข้อมูลการจัดกลุ่มสินค้า")
   pnlHeader.Caption = Me.Caption
   
   Call InitNormalLabel(lblMasterCode, MapText("รหัสสินค้า"))
   Call InitNormalLabel(lblMasterName, MapText("สินค้า"))
   Call txtMasterCode.SetTextLenType(TEXT_STRING, glbSetting.CODE_TYPE)
   Call txtMasterName.SetTextLenType(TEXT_STRING, glbSetting.NAME_TYPE)
   
   Call InitNormalLabel(lblSaveCode, MapText("รหัสกลุ่ม"))
   Call InitNormalLabel(lblSaveName, MapText("ชื่อกลุ่ม"))
   Call txtSaveCode.SetTextLenType(TEXT_STRING, glbSetting.CODE_TYPE)
   Call txtSaveName.SetTextLenType(TEXT_STRING, glbSetting.NAME_TYPE)
   
'   Call InitCheckBox(ChkNonAreaFLag, "ยังไม่ระบุเขต")
'   ChkNonAreaFLag.Enabled = False
   
   Call InitGrid

   Me.Picture = LoadPicture(glbParameterObj.MainPicture)
   
   pnlHeader.Font.Name = GLB_FONT
   pnlHeader.Font.Bold = True
   pnlHeader.Font.Size = 19
   
   cmdOK.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdExit.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdEdit.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdSearch.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdClear.Picture = LoadPicture(glbParameterObj.NormalButton1)
   
   Call InitMainButton(cmdOK, MapText("ตกลง (F2)"))
   Call InitMainButton(cmdExit, MapText("ยกเลิก (ESC)"))
   Call InitMainButton(cmdEdit, MapText("แก้ไข (F3)"))
   Call InitMainButton(cmdSearch, MapText("ค้นหา (F5)"))
   Call InitMainButton(cmdClear, MapText("เคลียร์ (F4)"))
   
End Sub


Private Sub Form_Load()
   OKClick = False
   m_HasActivate = False
   m_HasModify = False
   
   Set m_goodsDetail = New Collection
   Set tempGoodsDetail = New CGoodsMaster
   Set temp_getGoodsDetail = New CGoodsDetail
   Set m_stcrdEP = New CStmas
   Set m_Rs = New ADODB.Recordset
   Set mm_Rs = New ADODB.Recordset
   
    m_HasActivate = False
    
   Call InitFormLayout
   Call EnableForm(Me, True)
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set m_goodsDetail = Nothing
   Set tempGoodsDetail = Nothing
   Set temp_getGoodsDetail = Nothing
   Set m_stcrdEP = Nothing
   Set m_Rs = Nothing
   Set mm_Rs = Nothing
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
   Call m_stcrdEP.PopulateFromRS(1, m_Rs)

   Set temp_getGoodsDetail = GetStkcodGoodsDetail(m_goodsDetail, Trim(m_stcrdEP.STKCOD) & "-" & Trim(GOODS_MASTER_ID), False)
   If Not (temp_getGoodsDetail Is Nothing) Then
      Values(1) = temp_getGoodsDetail.GOODS_DETAIL_ID
      Values(4) = temp_getGoodsDetail.GOODS_GROUP_NAME
   Else
      Values(1) = -1
      Values(4) = "     -"
   End If
   Values(2) = m_stcrdEP.STKCOD
   Values(3) = m_stcrdEP.STKDES
   
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
   cmdExit.Top = ScaleHeight - 580
   cmdExit.Left = ScaleWidth - cmdOK.Width - 50
   cmdOK.Top = ScaleHeight - 580
   cmdOK.Left = ScaleWidth - cmdExit.Width - cmdOK.Width - 100
End Sub


Private Sub txtSaveCode_Change()
   m_HasModify = True
End Sub

Private Sub txtSaveName_Change()
    m_HasModify = True
End Sub
