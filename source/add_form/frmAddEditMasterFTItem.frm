VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#2.0#0"; "THREED20.OCX"
Begin VB.Form frmAddEditMasterFTItem 
   BackColor       =   &H80000000&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   4695
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9180
   BeginProperty Font 
      Name            =   "AngsanaUPC"
      Size            =   14.25
      Charset         =   222
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4695
   ScaleWidth      =   9180
   StartUpPosition =   2  'CenterScreen
   Begin Threed.SSPanel pnlHeader 
      Height          =   615
      Left            =   0
      TabIndex        =   4
      Top             =   0
      Width           =   9195
      _ExtentX        =   16219
      _ExtentY        =   1085
      _Version        =   131073
      PictureBackgroundStyle=   2
   End
   Begin Threed.SSFrame SSFrame1 
      Height          =   6255
      Left            =   0
      TabIndex        =   5
      Top             =   600
      Width           =   9225
      _ExtentX        =   16272
      _ExtentY        =   11033
      _Version        =   131073
      PictureBackgroundStyle=   2
      Begin VB.ComboBox cboParameter 
         Height          =   510
         Left            =   2280
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   2160
         Width           =   3855
      End
      Begin prjLedgerReport.uctlTextBox txtGp 
         Height          =   495
         Left            =   2280
         TabIndex        =   0
         Top             =   1440
         Width           =   2895
         _ExtentX        =   5106
         _ExtentY        =   873
      End
      Begin prjLedgerReport.uctlTextLookup uctlCommissionSale 
         Height          =   375
         Left            =   2280
         TabIndex        =   12
         Top             =   840
         Width           =   6255
         _ExtentX        =   11033
         _ExtentY        =   661
      End
      Begin VB.Label lblSaleName 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   435
         Left            =   240
         TabIndex        =   13
         Top             =   840
         Width           =   1935
      End
      Begin VB.Label lblParameter 
         Alignment       =   1  'Right Justify
         Height          =   375
         Left            =   600
         TabIndex        =   11
         Top             =   2160
         Width           =   1605
      End
      Begin VB.Label lblGp 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   435
         Left            =   120
         TabIndex        =   10
         Top             =   1440
         Width           =   1935
      End
      Begin VB.Label Label2 
         Height          =   315
         Left            =   4680
         TabIndex        =   9
         Top             =   5040
         Width           =   1755
      End
      Begin VB.Label Label1 
         Height          =   315
         Left            =   4680
         TabIndex        =   8
         Top             =   4560
         Width           =   1755
      End
      Begin VB.Label lblValue3 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   1080
         TabIndex        =   7
         Top             =   5040
         Width           =   1755
      End
      Begin VB.Label lblValue2 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   1200
         TabIndex        =   6
         Top             =   4560
         Width           =   1635
      End
      Begin Threed.SSCommand cmdOK 
         Height          =   525
         Left            =   2400
         TabIndex        =   2
         Top             =   3120
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdExit 
         Cancel          =   -1  'True
         Height          =   525
         Left            =   4200
         TabIndex        =   3
         Top             =   3120
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   926
         _Version        =   131073
         ButtonStyle     =   3
      End
   End
End
Attribute VB_Name = "frmAddEditMasterFTItem"
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

Private m_MasterFromToDetail As CMasterFromToDetail

Public HeaderText As String
Public ID As Long
Public MASTER_FROMTO_ID As Long
Public OKClick As Boolean
Public TempCollection As Collection
Public ParentForm As Form

Public StepFlag  As Boolean

Public DocumentType As MASTER_COMMISSION_AREA
Private FtSaleColl As Collection

Private Sub cboGroupCom_Click()
   m_HasModify = True
End Sub

Private Sub cmdExit_Click()
   If Not ConfirmExit(m_HasModify) Then
      Exit Sub
   End If
   
   OKClick = False
   Unload Me
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
   
   Call InitNormalLabel(lblParameter, MapText("�����"))
   Call InitNormalLabel(lblSaleName, MapText("��ѡ�ҹ���"))
   Call InitNormalLabel(lblGp, MapText("��� Gp"))
   
   Call txtGp.SetTextLenType(TEXT_FLOAT, glbSetting.DOUBLE_TYPE)
    Call InitCombo(cboParameter)
'   SSOption1.Value = True
'   SSOption4.Value = True
'
'   If StepFlag Then
'      lblValue3.Enabled = False
'      txtValue3.Enabled = False
'   End If
   
   cmdExit.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdOK.Picture = LoadPicture(glbParameterObj.NormalButton1)

   
   Call InitMainButton(cmdOK, MapText("��ŧ (F2)"))
   Call InitMainButton(cmdExit, MapText("¡��ԡ (ESC)"))

End Sub

Private Sub QueryData(Flag As Boolean)
Dim IsOK As Boolean
Dim ItemCount As Long

   If Flag Then
      Call EnableForm(Me, False)
      
      If ShowMode = SHOW_EDIT Then
         Dim Bd As CMasterFromToDetail
         
         Set Bd = TempCollection.ITEM(ID)
         txtGp.Text = Bd.GP
               cboParameter.ListIndex = IDToListIndex(cboParameter, Bd.MASTER_PARAMETER_ID)
                uctlCommissionSale.MyTextBox.Text = Bd.SLMCOD
         
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
         glbErrorLog.LocalErrorMsg = "�֧�ä�����ش��������"
         glbErrorLog.ShowUserError
         
         Call ParentForm.RefreshGrid
         Exit Sub
      End If
      
      ID = NewID
   ElseIf ShowMode = SHOW_ADD Then

      txtGp.Text = ""
'      txtValue2.Text = ""
'      txtValue3.Text = ""
   End If
   Call QueryData(True)
   
'   Call txtFrom.SetFocus
   
   Call ParentForm.RefreshGrid
End Sub

Private Sub cmdOK_Click()
   If Not SaveData Then
      Exit Sub
   End If
   
   OKClick = True
   Unload Me
End Sub
Private Function SaveData() As Boolean
Dim IsOK As Boolean
Dim RealIndex As Long

   If Not VerifyTextControl(lblGp, txtGp, False) Then
      Exit Function
   End If
   
   If Not Len(uctlCommissionSale.MyCombo.Text) > 0 Then
      Exit Function
   End If
   
   If Not VerifyCombo(lblParameter, cboParameter, False) Then
      Exit Function
   End If
   
   If Not m_HasModify Then
      SaveData = True
      Exit Function
   End If
   
   Dim Bd As CMasterFromToDetail
   If ShowMode = SHOW_ADD Then
      Set Bd = New CMasterFromToDetail
      Bd.Flag = "A"
      Call TempCollection.Add(Bd)
   Else
      Set Bd = TempCollection.ITEM(ID)
      If Bd.Flag <> "A" Then
         Bd.Flag = "E"
      End If
   End If

    Bd.GP = txtGp.Text
'    Bd.MASTER_AREA_ID = Left(cboAreaCod.Text, 2)
'    Bd.MASTER_AREA_NAME = Mid(cboAreaCod.Text, 5)
    Bd.SLMCOD = uctlCommissionSale.MyTextBox.Text
    Bd.SLMNAME = uctlCommissionSale.MyCombo.Text
    Bd.MASTER_PARAMETER_ID = cboParameter.ItemData(Minus2Zero(cboParameter.ListIndex))
    Bd.MASTER_PARAMETER_NAME = cboParameter.Text
'   Call BD.SetFieldValue("GROUP_COM_ID", cboGroupCom.ItemData(Minus2Zero(cboGroupCom.ListIndex)))
'   Call BD.SetFieldValue("GROUP_COM_DESC", cboGroupCom.Text)
   
   SaveData = True
End Function

Private Sub Form_Activate()
   If Not m_HasActivate Then
      m_HasActivate = True
      Me.Refresh
      DoEvents
            
       Call LoadSaleLookup(uctlCommissionSale.MyCombo, FtSaleColl) 'FtSaleColl, COMMISSION_TABLE
       Set uctlCommissionSale.MyCollection = FtSaleColl
       
       Call LoadCommissPara(cboParameter, , MASTER_FROMTO_ID)

      If ShowMode = SHOW_EDIT Then
         Call QueryData(True)
      ElseIf ShowMode = SHOW_ADD Then
         ID = 0
         Call QueryData(True)
      End If
      
      'm_HasModify = False
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

Private Sub Form_Load()

   OKClick = False
   Call InitFormLayout
   
   m_HasActivate = False
   m_HasModify = False
   Set m_Rs = New ADODB.Recordset
   Set m_MasterFromToDetail = New CMasterFromToDetail
   Set FtSaleColl = New Collection
   
End Sub

Private Sub Form_Unload(Cancel As Integer)
   If m_Rs.State = adStateOpen Then
      m_Rs.Close
   End If
   Set m_Rs = Nothing
   Set m_MasterFromToDetail = Nothing
   Set FtSaleColl = Nothing
End Sub

Private Sub txtGp_Change()
   m_HasModify = True
End Sub

Private Sub cboParameter_Click()
 m_HasModify = True
End Sub
