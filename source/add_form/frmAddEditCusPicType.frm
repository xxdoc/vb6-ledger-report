VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#2.0#0"; "THREED20.OCX"
Begin VB.Form frmAddEditCusPigType 
   BackColor       =   &H80000000&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   4560
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8370
   Icon            =   "frmAddEditCusPicType.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4560
   ScaleWidth      =   8370
   StartUpPosition =   1  'CenterOwner
   Begin Threed.SSFrame SSFrame1 
      Height          =   4485
      Left            =   0
      TabIndex        =   7
      Top             =   0
      Width           =   8385
      _ExtentX        =   14790
      _ExtentY        =   7911
      _Version        =   131073
      PictureBackgroundStyle=   2
      Begin Threed.SSPanel pnlHeader 
         Height          =   705
         Left            =   0
         TabIndex        =   8
         Top             =   0
         Width           =   8355
         _ExtentX        =   14737
         _ExtentY        =   1244
         _Version        =   131073
         PictureBackgroundStyle=   2
      End
      Begin prjLedgerReport.uctlTextBox txtName 
         Height          =   435
         Left            =   2340
         TabIndex        =   6
         Top             =   1440
         Width           =   5955
         _ExtentX        =   10504
         _ExtentY        =   767
      End
      Begin prjLedgerReport.uctlTextBox txtCode 
         Height          =   435
         Left            =   2340
         TabIndex        =   0
         Top             =   960
         Width           =   1875
         _ExtentX        =   5212
         _ExtentY        =   767
      End
      Begin prjLedgerReport.uctlTextBox txtPiggy 
         Height          =   435
         Left            =   2340
         TabIndex        =   3
         Top             =   2880
         Width           =   1875
         _ExtentX        =   5212
         _ExtentY        =   767
      End
      Begin prjLedgerReport.uctlTextBox txtKhun 
         Height          =   435
         Left            =   2340
         TabIndex        =   2
         Top             =   2400
         Width           =   1875
         _ExtentX        =   5212
         _ExtentY        =   767
      End
      Begin prjLedgerReport.uctlTextBox txtBreed 
         Height          =   435
         Left            =   2340
         TabIndex        =   1
         Top             =   1920
         Width           =   1875
         _ExtentX        =   5212
         _ExtentY        =   767
      End
      Begin VB.Label lblKhun 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   435
         Left            =   330
         TabIndex        =   13
         Top             =   2520
         Width           =   1935
      End
      Begin VB.Label lblPiggy 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   435
         Left            =   330
         TabIndex        =   12
         Top             =   3000
         Width           =   1935
      End
      Begin VB.Label lblBreed 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   435
         Left            =   330
         TabIndex        =   11
         Top             =   2040
         Width           =   1935
      End
      Begin VB.Label lblCode 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   255
         Left            =   330
         TabIndex        =   10
         Top             =   1080
         Width           =   1935
      End
      Begin VB.Label lblName 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   435
         Left            =   330
         TabIndex        =   9
         Top             =   1560
         Width           =   1935
      End
      Begin Threed.SSCommand cmdExit 
         Cancel          =   -1  'True
         Height          =   525
         Left            =   4920
         TabIndex        =   5
         Top             =   3600
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   926
         _Version        =   131073
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdOK 
         Height          =   525
         Left            =   3120
         TabIndex        =   4
         Top             =   3600
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmAddEditCusPicType.frx":27A2
         ButtonStyle     =   3
      End
   End
End
Attribute VB_Name = "frmAddEditCusPigType"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private m_HasActivate As Boolean
Private m_HasModify As Boolean
Private m_Rs As ADODB.Recordset
Private m_CusPigType As CCusPigType

Public ID As Long
Public OKClick As Boolean
Public ShowMode As SHOW_MODE_TYPE
Public HeaderText As String
Public CUS_PIG_TYPE_CODE As String
Public CUS_PIG_TYPE_NAME As String
Public CUS_PIG_TYPE_YEAR As Long

Private Sub cboDataType_Click()
   m_HasModify = True
End Sub

Private Sub cboDataType_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
      CreateObject("WScript.Shell").SendKeys "{TAB}"
   End If
End Sub
Private Sub cboCustomerTypeID_Click()
   m_HasModify = True
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
      
      m_CusPigType.CUS_PIG_TYPE_CODE = CUS_PIG_TYPE_CODE
      m_CusPigType.CUS_PIG_TYPE_YEAR = CUS_PIG_TYPE_YEAR
      
      m_CusPigType.QueryFlag = 1
      If Not glbDaily.QueryCusPigType(m_CusPigType, m_Rs, ItemCount, IsOK, glbErrorLog) Then
         glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
         Call EnableForm(Me, True)
         Exit Sub
      End If
   End If
   
   If ItemCount > 0 Then
      Call m_CusPigType.PopulateFromRS(1, m_Rs)
      
      txtCode.Text = m_CusPigType.CUS_PIG_TYPE_CODE
      txtName.Text = m_CusPigType.CUS_PIG_TYPE_NAME
      txtBreed.Text = m_CusPigType.CUS_PIG_TYPE_BREED
      txtKhun.Text = m_CusPigType.CUS_PIG_TYPE_KHUN
      txtPiggy.Text = m_CusPigType.CUS_PIG_TYPE_PIGGY
   End If
   
   If Not IsOK Then
      glbErrorLog.ShowUserError
      Call EnableForm(Me, True)
      Exit Sub
   End If
   
   Call EnableForm(Me, True)
End Sub
Private Function SaveData() As Boolean
Dim IsOK As Boolean
   
   If Not VerifyTextControl(lblName, txtName, False) Then
      Exit Function
   End If
   If Not VerifyTextControl(lblCode, txtCode, False) Then
      Exit Function
   End If
   
   If Not m_HasModify Then
      SaveData = True
      Exit Function
   End If
   
   m_CusPigType.AddEditMode = ShowMode
   m_CusPigType.CUS_PIG_TYPE_CODE = txtCode.Text
   m_CusPigType.CUS_PIG_TYPE_NAME = txtName.Text
   m_CusPigType.CUS_PIG_TYPE_BREED = Val(txtBreed.Text)
   m_CusPigType.CUS_PIG_TYPE_KHUN = Val(txtKhun.Text)
   m_CusPigType.CUS_PIG_TYPE_PIGGY = Val(txtPiggy.Text)
   'm_CusPigType.CUS_PIG_TYPE_YEAR = Year(Now)

   Call EnableForm(Me, False)
   If Not glbDaily.AddEditCusPigType(m_CusPigType, IsOK, True, glbErrorLog) Then
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
Private Sub Form_Activate()
   If Not m_HasActivate Then
      m_HasActivate = True
      Me.Refresh
      DoEvents

      
      If ShowMode = SHOW_EDIT Then
         Call QueryData(True)
      ElseIf ShowMode = SHOW_ADD Then
         ID = -1
      End If
      
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
   
   Call InitNormalLabel(lblCode, MapText("รหัสลูกค้า"))
   Call txtCode.SetKeySearch("CUS_PIG_TYPE_CODE")
   
   Call InitNormalLabel(lblName, MapText("ชื่อลูกค้า"))
   Call txtName.SetKeySearch("CUS_PIG_TYPE_NAME")
   
   Call InitNormalLabel(lblBreed, MapText("จำนวนหมูพันธุ์"))
   Call txtBreed.SetKeySearch("CUS_PIG_TYPE_BREED")
   
   Call InitNormalLabel(lblKhun, MapText("จำนวนหมูขุน"))
   Call txtKhun.SetKeySearch("CUSTOMER_NAME")
   
   Call InitNormalLabel(lblPiggy, MapText("จำนวนลูกหมู"))
   Call txtPiggy.SetKeySearch("CUSTOMER_NAME")
   
   Call txtCode.SetTextLenType(TEXT_STRING, glbSetting.CODE_TYPE)
   Call txtName.SetTextLenType(TEXT_STRING, glbSetting.NAME_TYPE)
   Call txtBreed.SetTextLenType(TEXT_FLOAT, glbSetting.DOUBLE_TYPE)
   Call txtKhun.SetTextLenType(TEXT_FLOAT, glbSetting.DOUBLE_TYPE)
   Call txtPiggy.SetTextLenType(TEXT_FLOAT, glbSetting.DOUBLE_TYPE)
 '  Call txtpRow1.SetTextLenType(TEXT_INTEGER, glbSetting.CODE_TYPE)
   
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
   
   Set m_CusPigType = New CCusPigType
   Set m_Rs = New ADODB.Recordset
   
   txtCode.Text = CUS_PIG_TYPE_CODE
   txtName.Text = CUS_PIG_TYPE_NAME
    txtCode.Enabled = False
   txtName.Enabled = False
   
   m_HasActivate = False
   Call InitFormLayout
   Call EnableForm(Me, True)
End Sub
Private Sub Form_Unload(Cancel As Integer)
   Set m_CusPigType = Nothing
End Sub

Private Sub txtBreed_Change()
   m_HasModify = True
End Sub

Private Sub txtCode_Change()
   m_HasModify = True
End Sub

Private Sub txtKhun_Change()
   m_HasModify = True
End Sub

Private Sub txtName_Change()
   m_HasModify = True
End Sub

Private Sub txtPiggy_Change()
   m_HasModify = True
End Sub
