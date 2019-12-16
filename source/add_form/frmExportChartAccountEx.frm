VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#2.0#0"; "THREED20.OCX"
Begin VB.Form frmExportChartAccountEx 
   BackColor       =   &H80000000&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   3075
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10605
   BeginProperty Font 
      Name            =   "AngsanaUPC"
      Size            =   14.25
      Charset         =   222
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmExportChartAccountEx.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3075
   ScaleWidth      =   10605
   StartUpPosition =   2  'CenterScreen
   Begin Threed.SSPanel pnlHeader 
      Height          =   615
      Left            =   0
      TabIndex        =   6
      Top             =   0
      Width           =   10605
      _ExtentX        =   18706
      _ExtentY        =   1085
      _Version        =   131073
      PictureBackgroundStyle=   2
   End
   Begin Threed.SSFrame SSFrame1 
      Height          =   2535
      Left            =   0
      TabIndex        =   7
      Top             =   600
      Width           =   10605
      _ExtentX        =   18706
      _ExtentY        =   4471
      _Version        =   131073
      PictureBackgroundStyle=   2
      Begin prjLedgerReport.uctlTextBox txtZipcode 
         Height          =   435
         Left            =   12450
         TabIndex        =   4
         Top             =   3270
         Width           =   2715
         _ExtentX        =   4789
         _ExtentY        =   767
      End
      Begin prjLedgerReport.uctlTextBox txtMainCode 
         Height          =   435
         Left            =   1950
         TabIndex        =   0
         Top             =   360
         Width           =   1665
         _ExtentX        =   2937
         _ExtentY        =   767
      End
      Begin prjLedgerReport.uctlTextBox txtSubCode 
         Height          =   435
         Left            =   1950
         TabIndex        =   1
         Top             =   840
         Width           =   7545
         _ExtentX        =   6800
         _ExtentY        =   767
      End
      Begin Threed.SSCommand cmdNext 
         Height          =   525
         Left            =   2880
         TabIndex        =   2
         Top             =   1560
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmExportChartAccountEx.frx":08CA
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdOK2 
         Cancel          =   -1  'True
         Height          =   525
         Left            =   4560
         TabIndex        =   3
         Top             =   1560
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   926
         _Version        =   131073
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdCancel 
         Height          =   525
         Left            =   6240
         TabIndex        =   5
         Top             =   1560
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   926
         _Version        =   131073
         ButtonStyle     =   3
      End
      Begin VB.Label lblSubCode 
         Alignment       =   1  'Right Justify
         Height          =   375
         Left            =   120
         TabIndex        =   9
         Top             =   960
         Width           =   1815
      End
      Begin VB.Label lblMainCode 
         Alignment       =   1  'Right Justify
         Height          =   375
         Left            =   60
         TabIndex        =   8
         Top             =   480
         Width           =   1815
      End
   End
End
Attribute VB_Name = "frmExportChartAccountEx"
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

Public HeaderText As String
Public ID As Long
Public OKClick As Boolean
Public TempCollection As Collection
Public ParentForm As Form

Private Sub cmdCancel_Click()
   If Not ConfirmExit(m_HasModify) Then
      Exit Sub
   End If
   
   OKClick = False
   Unload Me
End Sub
Private Sub InitFormLayout()
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
      
   Call InitNormalLabel(lblMainCode, "รหัสหลัก")
   Call InitNormalLabel(lblSubCode, "รหัสย่อย")
   
   Call txtMainCode.SetTextLenType(TEXT_STRING, glbSetting.DESC_TYPE)
   Call txtSubCode.SetTextLenType(TEXT_STRING, glbSetting.DESC_TYPE)
   
   cmdCancel.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdOK2.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdNext.Picture = LoadPicture(glbParameterObj.NormalButton1)
   
   Call InitMainButton(cmdCancel, MapText("ยกเลิก (ESC)"))
   Call InitMainButton(cmdOK2, MapText("ตกลง (F2)"))
   Call InitMainButton(cmdNext, MapText("ถัดไป"))
End Sub
Private Sub QueryData(Flag As Boolean)
Dim IsOK As Boolean
Dim ItemCount As Long

   If Flag Then
      Call EnableForm(Me, False)
            
      If ShowMode = SHOW_EDIT Then
         Dim D As CAccountCode
         Set D = TempCollection.Item(ID)

         txtMainCode.Text = D.MAIN_CODE
         txtSubCode.Text = D.SUB_CODE
      End If
   End If
   
   Call EnableForm(Me, True)
End Sub

Private Sub cmdNext_Click()
Dim NewID As Long

   If Not SaveData Then
      Exit Sub
   End If
   
   If ShowMode = SHOW_EDIT Then
      NewID = GetNextID(ID, TempCollection)
      If ID = NewID Then
         glbErrorLog.LocalErrorMsg = "ถึงเรคคอร์ดสุดท้ายแล้ว"
         glbErrorLog.ShowUserError
         
         Call ParentForm.RefreshGrid
         Exit Sub
      End If
      
      ID = NewID
   ElseIf ShowMode = SHOW_ADD Then
      txtMainCode.Text = ""
      txtSubCode.Text = ""
   End If
   
   txtMainCode.SetFocus
   
   Call QueryData(True)
   Call ParentForm.RefreshGrid
End Sub

Private Sub cmdOK2_Click()
   If Not SaveData Then
      Exit Sub
   End If
   
   OKClick = True
   Unload Me
End Sub

Private Function SaveData() As Boolean
Dim IsOK As Boolean
Dim RealIndex As Long
   
   SaveData = False
   If Not VerifyTextControl(lblMainCode, txtMainCode) Then
      Exit Function
   End If
   If Not VerifyTextControl(lblSubCode, txtSubCode) Then
      Exit Function
   End If
   
   If Not m_HasModify Then
      SaveData = True
      Exit Function
   End If
   
   Dim D As CAccountCode
   If ShowMode = SHOW_ADD Then
      Set D = New CAccountCode

      D.Flag = "A"
      Call TempCollection.Add(D)
   Else
      Set D = TempCollection.Item(ID)
      D.Flag = "E"
   End If

   D.MAIN_CODE = txtMainCode.Text
   D.SUB_CODE = txtSubCode.Text
   
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
         Call QueryData(False)
      End If
      
      m_HasModify = False
   End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
   If Shift = 1 And KeyCode = DUMMY_KEY Then
      glbErrorLog.LocalErrorMsg = Me.Name
      glbErrorLog.ShowUserError
   ElseIf Shift = 1 And KeyCode = 112 Then
      If glbUser.EXCEPTION_FLAG = "Y" Then
         glbUser.EXCEPTION_FLAG = "N"
      Else
         glbUser.EXCEPTION_FLAG = "Y"
      End If
   ElseIf Shift = 0 And KeyCode = 116 Then
'      Call cmdSearch_Click
   ElseIf Shift = 0 And KeyCode = 115 Then
'      Call cmdClear_Click
   ElseIf Shift = 0 And KeyCode = 118 Then
'      Call cmdAdd_Click
'   ElseIf Shift = 0 And KeyCode = 117 Then
'      Call cmdDelete_Click
   ElseIf Shift = 0 And KeyCode = 113 Then
      Call cmdOK2_Click
   ElseIf Shift = 0 And KeyCode = 114 Then
'      Call cmdEdit_Click
   ElseIf Shift = 0 And KeyCode = 121 Then
'      Call cmdPrint_Click
   End If
End Sub

Private Sub Form_Load()

   OKClick = False
   Call InitFormLayout
   
   m_HasActivate = False
   m_HasModify = False
   Set m_Rs = New ADODB.Recordset
End Sub
Private Sub Form_Unload(Cancel As Integer)
   If m_Rs.State = adStateOpen Then
      m_Rs.Close
   End If
   Set m_Rs = Nothing
End Sub
Private Sub txtMainCode_Change()
   m_HasModify = True
End Sub
Private Sub txtSubCode_Change()
   m_HasModify = True
End Sub
