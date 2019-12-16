VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#2.0#0"; "THREED20.OCX"
Begin VB.Form frmAddEditComEx 
   BackColor       =   &H80000000&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   3780
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6735
   Icon            =   "frmAddEditComEx.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3780
   ScaleWidth      =   6735
   StartUpPosition =   1  'CenterOwner
   Begin Threed.SSFrame SSFrame1 
      Height          =   3765
      Left            =   0
      TabIndex        =   4
      Top             =   0
      Width           =   6825
      _ExtentX        =   12039
      _ExtentY        =   6641
      _Version        =   131073
      PictureBackgroundStyle=   2
      Begin prjLedgerReport.uctlTextBox txtParaName 
         Height          =   435
         Left            =   2400
         TabIndex        =   0
         Top             =   1200
         Width           =   2925
         _extentx        =   5159
         _extenty        =   767
      End
      Begin Threed.SSPanel pnlHeader 
         Height          =   705
         Left            =   0
         TabIndex        =   5
         Top             =   0
         Width           =   6795
         _ExtentX        =   11986
         _ExtentY        =   1244
         _Version        =   131073
         PictureBackgroundStyle=   2
      End
      Begin prjLedgerReport.uctlTextBox txtParaValue 
         Height          =   435
         Left            =   2400
         TabIndex        =   1
         Top             =   1800
         Width           =   2925
         _extentx        =   5159
         _extenty        =   767
      End
      Begin VB.Label lblParaValue 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   435
         Left            =   240
         TabIndex        =   7
         Top             =   1800
         Width           =   1935
      End
      Begin VB.Label lblParaName 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   435
         Left            =   240
         TabIndex        =   6
         Top             =   1200
         Width           =   1935
      End
      Begin Threed.SSCommand cmdExit 
         Cancel          =   -1  'True
         Height          =   525
         Left            =   3360
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
         Left            =   1560
         TabIndex        =   2
         Top             =   2880
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmAddEditComEx.frx":27A2
         ButtonStyle     =   3
      End
   End
End
Attribute VB_Name = "frmAddEditComEx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Header As String
Public ShowMode As SHOW_MODE_TYPE
Private m_HasActivate As Boolean
Private m_HasModify As Boolean
Private m_Rs As ADODB.Recordset

Private MasterPara As CCommissMasterPara

Public HeaderText As String
Public ID As Long
Public OKClick As Boolean
Public TempCollection As Collection
Public ParentForm As Form

Public StepFlag  As Boolean
Public DocumentType As MASTER_COMMISSION_AREA
'Public MASTER_PARAMETER_ID As Long

Private Sub cmdExit_Click()
   If Not ConfirmExit(m_HasModify) Then
      Exit Sub
   End If
   
   OKClick = False
   Unload Me
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
      
      If ShowMode = SHOW_EDIT Then
         Dim BD As CCommissMasterPara
         
         Set BD = TempCollection.ITEM(ID)
          txtParaName.Text = BD.MASTER_PARAMETER_NAME
          txtParaValue.Text = BD.MASTER_PARAMETER_VALUE

      End If
   End If

   Call EnableForm(Me, True)
End Sub
Private Function SaveData() As Boolean
Dim IsOK As Boolean
   
   If Not VerifyTextControl(lblParaName, txtParaName, False) Then
      Exit Function
   End If
   
   If Not VerifyTextControl(lblParaValue, txtParaValue, False) Then
      Exit Function
   End If
   
   If Not m_HasModify Then
      SaveData = True
      Exit Function
   End If
   
    Dim BD As CCommissMasterPara
   If ShowMode = SHOW_ADD Then
      Set BD = New CCommissMasterPara
      BD.Flag = "A"
      Call TempCollection.Add(BD)
   Else
      Set BD = TempCollection.ITEM(ID)
      If BD.Flag <> "A" Then
         BD.Flag = "E"
      End If
   End If

 '  MasterPara.MASTER_PARAMETER_ID = MASTER_PARAMETER_ID
   BD.MASTER_PARAMETER_NAME = txtParaName.Text
   BD.MASTER_PARAMETER_VALUE = txtParaValue.Text
      
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
               ID = 0
         Call QueryData(True)
         'MASTER_PARAMETER_ID = -1
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
   Me.Picture = LoadPicture(glbParameterObj.MainPicture)
   pnlHeader.PictureBackground = LoadPicture(glbParameterObj.NormalForm1)
   SSFrame1.PictureBackground = LoadPicture(glbParameterObj.NormalForm1)
   
   Me.KeyPreview = True
   pnlHeader.Caption = HeaderText
   Me.BackColor = GLB_FORM_COLOR
   pnlHeader.BackColor = GLB_HEAD_COLOR
   SSFrame1.BackColor = GLB_FORM_COLOR
   
      pnlHeader.Font.Name = GLB_FONT
   pnlHeader.Font.Bold = True
   pnlHeader.Font.Size = 19
   pnlHeader.Caption = HeaderText
   
   Call InitNormalLabel(lblParaName, MapText("กลุ่ม"))
   Call txtParaName.SetTextLenType(TEXT_STRING, glbSetting.NAME_TYPE)
   
   Call InitNormalLabel(lblParaValue, MapText("ค่า"))
   Call txtParaValue.SetTextLenType(TEXT_FLOAT, glbSetting.DOUBLE_TYPE)

   
   cmdExit.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdOK.Picture = LoadPicture(glbParameterObj.NormalButton1)
   
   Call InitMainButton(cmdExit, MapText("ยกเลิก (ESC)"))
   Call InitMainButton(cmdOK, MapText("ตกลง (F2)"))
End Sub


Private Sub Form_Load()
      OKClick = False
   Call InitFormLayout
   m_HasActivate = False
   
   Set MasterPara = New CCommissMasterPara
   Set m_Rs = New ADODB.Recordset
   
    m_HasModify = False
 '  Call InitFormLayout
'   Call EnableForm(Me, True)
End Sub
Private Sub Form_Unload(Cancel As Integer)
    If m_Rs.State = adStateOpen Then
          m_Rs.Close
    End If
   Set m_Rs = Nothing
   
   Set MasterPara = Nothing
End Sub
Private Sub txtParaName_Change()
   m_HasModify = True
End Sub

Private Sub txtParaValue_Change()
   m_HasModify = True
End Sub
