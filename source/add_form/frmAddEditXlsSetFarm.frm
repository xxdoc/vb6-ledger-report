VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#2.0#0"; "THREED20.OCX"
Begin VB.Form frmAddEditXlsSetFarm 
   BackColor       =   &H80000000&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   4020
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7140
   Icon            =   "frmAddEditXlsSetFarm.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4020
   ScaleWidth      =   7140
   StartUpPosition =   1  'CenterOwner
   Begin Threed.SSFrame SSFrame1 
      Height          =   4005
      Left            =   0
      TabIndex        =   5
      Top             =   0
      Width           =   7185
      _ExtentX        =   12674
      _ExtentY        =   7064
      _Version        =   131073
      PictureBackgroundStyle=   2
      Begin VB.ComboBox cboXlsNameFarm 
         Height          =   315
         Left            =   2400
         Style           =   2  'Dropdown List
         TabIndex        =   0
         Top             =   1200
         Width           =   2955
      End
      Begin VB.ComboBox cboXlsUnit 
         Height          =   315
         Left            =   2400
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   1800
         Width           =   2955
      End
      Begin prjLedgerReport.uctlTextBox txtCost 
         Height          =   435
         Left            =   2400
         TabIndex        =   2
         Top             =   2280
         Width           =   2955
         _ExtentX        =   5212
         _ExtentY        =   767
      End
      Begin Threed.SSPanel pnlHeader 
         Height          =   705
         Left            =   0
         TabIndex        =   9
         Top             =   0
         Width           =   7155
         _ExtentX        =   12621
         _ExtentY        =   1244
         _Version        =   131073
         PictureBackgroundStyle=   2
      End
      Begin VB.Label lblNameFarm 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   255
         Left            =   330
         TabIndex        =   8
         Top             =   1200
         Width           =   1935
      End
      Begin VB.Label lblCost 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   435
         Left            =   360
         TabIndex        =   7
         Top             =   2280
         Width           =   1935
      End
      Begin VB.Label lblUnit 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   435
         Left            =   360
         TabIndex        =   6
         Top             =   1800
         Width           =   1935
      End
      Begin Threed.SSCommand cmdExit 
         Cancel          =   -1  'True
         Height          =   525
         Left            =   3720
         TabIndex        =   4
         Top             =   3120
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   926
         _Version        =   131073
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdOK 
         Height          =   525
         Left            =   1920
         TabIndex        =   3
         Top             =   3120
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmAddEditXlsSetFarm.frx":27A2
         ButtonStyle     =   3
      End
   End
End
Attribute VB_Name = "frmAddEditXlsSetFarm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private m_HasActivate As Boolean
Private m_HasModify As Boolean
Private m_Rs As ADODB.Recordset
Private XlsSetFarm As CXlsSetFarm
'Private XlsSetFarm As Collection

Public ID As Long
Public OKClick As Boolean
Public ShowMode As SHOW_MODE_TYPE
Public HeaderText As String
Private Sub cboUnit_Click()
   m_HasModify = True
End Sub

Private Sub cboUnit_KeyPress(KeyAscii As Integer)
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
      
      XlsSetFarm.SET_FARM_ID = ID
      XlsSetFarm.QueryFlag = 1
     If Not glbDaily.QueryXlsSetFarm(XlsSetFarm, m_Rs, ItemCount, IsOK, glbErrorLog) Then
         glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
         Call EnableForm(Me, True)
         Exit Sub
      End If
 End If
   
   If ItemCount > 0 Then
      Call XlsSetFarm.PopulateFromRS(1, m_Rs)
      txtCost.Text = XlsSetFarm.SET_FARM_PRICE
      cboXlsUnit.ListIndex = IDToListIndex(cboXlsUnit, XlsSetFarm.XLS_UNIT_ID)
      cboXlsNameFarm.ListIndex = IDToListIndex(cboXlsNameFarm, XlsSetFarm.MAIN_FARM_ID)
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

   If Not VerifyCombo(lblNameFarm, cboXlsNameFarm, False) Then
      Exit Function
   End If
   If Not VerifyTextControl(lblCost, txtCost, False) Then
      Exit Function
   End If
   If Not VerifyCombo(lblUnit, cboXlsUnit, False) Then
      Exit Function
   End If

'   If Not CheckUniqueNs(REAL_CREDIT_NO, txtDocumentNo.Text, ID) Then
'      glbErrorLog.LocalErrorMsg = MapText("�բ�����") & " " & txtDocumentNo.Text & " " & MapText("������к�����")
'      glbErrorLog.ShowUserError
'      Exit Function
'   End If
   
   'Call LoadARAmountByCust(Nothing, DocNoColl, -1, -1, txtDocumentNo.Text)
'   If DocNoColl.Count <= 0 Then
'      glbErrorLog.LocalErrorMsg = MapText("����բ����������Ţ�͡���") & " " & txtDocumentNo.Text & " " & MapText("��к�")
'      glbErrorLog.ShowUserError
'      Exit Function
'   End If
   
   If Not m_HasModify Then
      SaveData = True
      Exit Function
   End If
   
      XlsSetFarm.AddEditMode = ShowMode
      XlsSetFarm.SET_FARM_ID = ID
      XlsSetFarm.MAIN_FARM_ID = cboXlsNameFarm.ItemData(Minus2Zero(cboXlsNameFarm.ListIndex))
      XlsSetFarm.SET_FARM_PRICE = txtCost.Text
      XlsSetFarm.XLS_UNIT_ID = cboXlsUnit.ItemData(Minus2Zero(cboXlsUnit.ListIndex))

   Call EnableForm(Me, False)
   If Not glbDaily.AddEditXlsSetFarm(XlsSetFarm, IsOK, True, glbErrorLog) Then
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
      
      Call LoadXlsUnit(cboXlsUnit)
      Call LoadXlsNameFarm(cboXlsNameFarm)
      
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

   Call InitNormalLabel(lblCost, MapText("�Ҥ�"))
   Call txtCost.SetKeySearch("FOOD_COST")
   Call txtCost.SetTextLenType(TEXT_FLOAT_MONEY, glbSetting.MONEY_TYPE)
   
   Call InitNormalLabel(lblUnit, MapText("˹���"))
   Call InitCombo(cboXlsUnit)
   
   Call InitNormalLabel(lblNameFarm, MapText("���Ϳ����"))
   Call InitCombo(cboXlsNameFarm)

   Me.Picture = LoadPicture(glbParameterObj.MainPicture)
   
   pnlHeader.Font.Name = GLB_FONT
   pnlHeader.Font.Bold = True
   pnlHeader.Font.Size = 19
   
   cmdExit.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdOK.Picture = LoadPicture(glbParameterObj.NormalButton1)
   
   Call InitMainButton(cmdExit, MapText("¡��ԡ (ESC)"))
   Call InitMainButton(cmdOK, MapText("��ŧ (F2)"))
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
   
   Set XlsSetFarm = New CXlsSetFarm
   Set m_Rs = New ADODB.Recordset
   
   m_HasActivate = False
   Call InitFormLayout
   Call EnableForm(Me, True)
End Sub
Private Sub Form_Unload(Cancel As Integer)
   Set XlsSetFarm = Nothing
End Sub

Private Sub txtCost_Change()
   m_HasModify = True
End Sub
