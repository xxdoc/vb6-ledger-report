VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#2.0#0"; "THREED20.OCX"
Begin VB.Form frmAddEditGoodsDetail 
   BackColor       =   &H80000000&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   4095
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9225
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4095
   ScaleWidth      =   9225
   StartUpPosition =   1  'CenterOwner
   Begin Threed.SSFrame SSFrame1 
      Height          =   8325
      Left            =   0
      TabIndex        =   5
      Top             =   0
      Width           =   10905
      _ExtentX        =   19235
      _ExtentY        =   14684
      _Version        =   131073
      PictureBackgroundStyle=   2
      Begin VB.ComboBox cboGoodsGroup 
         Height          =   315
         Left            =   2115
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   2040
         Width           =   3375
      End
      Begin prjLedgerReport.uctlTextLookup uctlStkcodLookup 
         Height          =   375
         Left            =   2115
         TabIndex        =   0
         Top             =   1440
         Width           =   6255
         _ExtentX        =   11033
         _ExtentY        =   661
      End
      Begin Threed.SSPanel pnlHeader 
         Height          =   705
         Left            =   0
         TabIndex        =   6
         Top             =   0
         Width           =   9195
         _ExtentX        =   16219
         _ExtentY        =   1244
         _Version        =   131073
         PictureBackgroundStyle=   2
      End
      Begin VB.Label lblGoodsGroup 
         Alignment       =   1  'Right Justify
         Caption         =   "Label2"
         Height          =   435
         Left            =   240
         TabIndex        =   7
         Top             =   2040
         Width           =   1575
      End
      Begin VB.Label lblStkcod 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   375
         Left            =   240
         TabIndex        =   4
         Top             =   1440
         Width           =   1575
      End
      Begin Threed.SSCommand cmdExit 
         Cancel          =   -1  'True
         Height          =   525
         Left            =   4440
         TabIndex        =   3
         Top             =   2880
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   926
         _Version        =   131073
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdOK 
         Height          =   525
         Left            =   2640
         TabIndex        =   2
         Top             =   2880
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         ButtonStyle     =   3
      End
   End
End
Attribute VB_Name = "frmAddEditGoodsDetail"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private m_HasActivate As Boolean
Private m_HasModify As Boolean
Private m_Rs As ADODB.Recordset
Private itemGoodsDetail  As CGoodsDetail
Private item_getGoodsDetail As CGoodsDetail

Public ID As Long
Public OKClick As Boolean
Public ShowMode As SHOW_MODE_TYPE
Public HeaderText As String

Public GOODS_DETAIL_ID As Long
Public GOODS_MASTER_ID As Long
'Public YEAR_ID As Long

Public m_goodsDetail As Collection
Public m_Stmas As Collection
Public STKCOD As String

'Private Sub cboDataType_Click()
'   m_HasModify = True
'End Sub
Private Sub cboDataType_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
      CreateObject("WScript.Shell").SendKeys "{TAB}"
   End If
End Sub
Private Sub cmdOK_Click()
   If Not SaveData Then
      Exit Sub
   End If
   
   OKClick = True
Unload Me
End Sub
Private Sub QueryData(Flag As Boolean)
Dim IsOK As Boolean
Dim ItemCount As Long

   If Flag Then
      Call EnableForm(Me, False)
      
      itemGoodsDetail.GOODS_MASTER_ID = GOODS_MASTER_ID
      itemGoodsDetail.STKCOD = STKCOD
    '  itemGoodsDetail.GOODS_DETAIL_ID = GOODS_DETAIL_ID
      
      If Not glbDaily.QueryGoodsDetail(itemGoodsDetail, m_Rs, ItemCount, IsOK, glbErrorLog) Then
         glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
         Call EnableForm(Me, True)
         Exit Sub
      End If
   End If
   
   If ItemCount > 0 Then
      Call itemGoodsDetail.PopulateFromRS(1, m_Rs)
       cboGoodsGroup.ListIndex = IDToListIndex(cboGoodsGroup, itemGoodsDetail.GOODS_GROUP_ID)
   End If
   uctlStkcodLookup.MyTextBox.Text = STKCOD
   
   If Not IsOK Then
      glbErrorLog.ShowUserError
      Call EnableForm(Me, True)
      Exit Sub
   End If
   
   Call EnableForm(Me, True)
End Sub
Private Function SaveData() As Boolean
Dim IsOK As Boolean
Dim m_Lookup As CGoodsDetail
  Set m_Lookup = New CGoodsDetail
  

   If Not (uctlStkcodLookup.MyCombo.ItemData(Minus2Zero(uctlStkcodLookup.MyCombo.ListIndex)) > 0) Then
      Exit Function
   End If
   
'  If Not VerifyCombo(lblAreaName, cboGoodsGroup, False) Then
'      Exit Function
'   End If

   If Not m_HasModify Then
      SaveData = True
      Exit Function
   End If
   
   If cboGoodsGroup.ItemData(Minus2Zero(cboGoodsGroup.ListIndex)) > 0 Then
   
   Set item_getGoodsDetail = GetObject("", m_goodsDetail, Trim(uctlStkcodLookup.MyTextBox.Text))
   If Not (item_getGoodsDetail Is Nothing) Then
      glbErrorLog.LocalErrorMsg = MapText("คุณต้องการเปลี่ยนสินค้า : " & item_getGoodsDetail.STKCOD & " " & item_getGoodsDetail.STKDES)
      If glbErrorLog.AskMessage = vbNo Then
             Exit Function
      End If
   End If
      
      itemGoodsDetail.AddEditMode = ShowMode
      itemGoodsDetail.GOODS_MASTER_ID = GOODS_MASTER_ID
      itemGoodsDetail.STKCOD = uctlStkcodLookup.MyTextBox.Text
      itemGoodsDetail.STKDES = uctlStkcodLookup.MyCombo.Text
      itemGoodsDetail.GOODS_GROUP_ID = cboGoodsGroup.ItemData(Minus2Zero(cboGoodsGroup.ListIndex))

      Call EnableForm(Me, False)
      If Not glbDaily.AddEditGoodsDetail(itemGoodsDetail, IsOK, True, glbErrorLog) Then
         glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
         SaveData = False
         Call EnableForm(Me, True)
         Exit Function
      End If
   Else
         If Not ConfirmDelete(uctlStkcodLookup.MyCombo.Text & "  ในกลุ่มดังกล่าว") Then
            Exit Function
         End If

         If Not glbDaily.DeleteGoodsDetail(itemGoodsDetail, IsOK, True, glbErrorLog) Then
            itemGoodsDetail.GOODS_DETAIL_ID = ""
            glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
            itemGoodsDetail.STKCOD = ""
            glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
            Call EnableForm(Me, True)
            Exit Function
         End If
   End If
   
   If Not IsOK Then
      Call EnableForm(Me, True)
      glbErrorLog.ShowUserError
      Exit Function
   End If
   
   Call EnableForm(Me, True)
   SaveData = True
End Function
Private Sub Form_Activate()
   If Not m_HasActivate Then
      m_HasActivate = True
      Me.Refresh
      DoEvents
      
      Call LoadGoodsGroup(cboGoodsGroup)
      
      Call LoadStmasLookup(uctlStkcodLookup.MyCombo, m_Stmas)    'โหลด STK ....  หาสิ m_Stmasคืออะไร
      Set uctlStkcodLookup.MyCollection = m_Stmas

      Call LoadGoodsDetailFromMaster(Nothing, GOODS_MASTER_ID, m_goodsDetail)

      If GOODS_DETAIL_ID > 0 Then
         ShowMode = SHOW_EDIT
      Else
        ShowMode = SHOW_ADD
      End If
      
      Call QueryData(True)
      m_HasModify = False
   End If
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
'      Call cmdAdd_Click
      KeyCode = 0
   ElseIf Shift = 0 And KeyCode = 117 Then
'      Call cmdDelete_Click
      KeyCode = 0
   ElseIf Shift = 0 And KeyCode = 113 Then
      Call cmdOK_Click
      KeyCode = 0
   ElseIf Shift = 0 And KeyCode = 114 Then
'      Call cmdEdit_Click
      KeyCode = 0
   ElseIf Shift = 0 And KeyCode = 121 Then
'      Call cmdPrint_Click
      KeyCode = 0
   End If
End Sub

Private Sub InitFormLayout()
   pnlHeader.PictureBackground = LoadPicture(glbParameterObj.NormalForm1)
   SSFrame1.PictureBackground = LoadPicture(glbParameterObj.NormalForm1)
   
   Me.Caption = HeaderText
   pnlHeader.Caption = Me.Caption
   
   Call InitNormalLabel(lblStkCod, MapText("รหัส"))
   Call InitNormalLabel(lblGoodsGroup, MapText("ชื่อกลุ่ม"))
   Call InitCombo(cboGoodsGroup)

   Me.Picture = LoadPicture(glbParameterObj.MainPicture)
   
   pnlHeader.Font.Name = GLB_FONT
   pnlHeader.Font.Bold = True
   pnlHeader.Font.Size = 19
   
   cmdExit.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdOK.Picture = LoadPicture(glbParameterObj.NormalButton1)
   
   Call InitMainButton(cmdExit, MapText("ยกเลิก (ESC)"))
   Call InitMainButton(cmdOK, MapText("ตกลง (F2)"))
End Sub
Private Sub cmdExit_Click()
   If Not ConfirmExit(m_HasModify) Then
      Exit Sub
   End If
   
   OKClick = False
   Unload Me
End Sub
Private Sub Form_Load()
   Call EnableForm(Me, False)
   m_HasActivate = False
   Set m_Rs = New ADODB.Recordset
      
   Set itemGoodsDetail = New CGoodsDetail
   Set item_getGoodsDetail = New CGoodsDetail
   Set m_goodsDetail = New Collection
   Set m_Stmas = New Collection
   
   m_HasActivate = False
   Call InitFormLayout
   Call EnableForm(Me, True)
End Sub
Private Sub Form_Unload(Cancel As Integer)
   Set itemGoodsDetail = Nothing
   Set m_Stmas = Nothing
  Set m_goodsDetail = Nothing
  Set item_getGoodsDetail = Nothing
End Sub

Private Sub uctlStkcodLookup_HasChange()
   m_HasModify = True
End Sub

Private Sub uctlStkcodLookup_Change()
   m_HasModify = True
End Sub

Private Sub cboGoodsGroup_Click()
 m_HasModify = True
End Sub

