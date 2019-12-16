VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#2.0#0"; "THREED20.OCX"
Begin VB.Form frmAddEditComTypeDonStk 
   BackColor       =   &H80000000&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   3510
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11595
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3510
   ScaleWidth      =   11595
   StartUpPosition =   1  'CenterOwner
   Begin Threed.SSFrame SSFrame1 
      Height          =   8325
      Left            =   0
      TabIndex        =   2
      Top             =   -120
      Width           =   11985
      _ExtentX        =   21140
      _ExtentY        =   14684
      _Version        =   131073
      PictureBackgroundStyle=   2
      Begin Threed.SSPanel pnlHeader 
         Height          =   705
         Left            =   0
         TabIndex        =   3
         Top             =   120
         Width           =   11955
         _ExtentX        =   21087
         _ExtentY        =   1244
         _Version        =   131073
         PictureBackgroundStyle=   2
      End
      Begin prjLedgerReport.uctlTextLookup uctlStkCodStkdesLookup 
         Height          =   375
         Left            =   3360
         TabIndex        =   4
         Top             =   1440
         Width           =   7095
         _ExtentX        =   11033
         _ExtentY        =   661
      End
      Begin VB.Label lblStkcodStkdes 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   375
         Left            =   960
         TabIndex        =   5
         Top             =   1440
         Width           =   1935
      End
      Begin Threed.SSCommand cmdExit 
         Cancel          =   -1  'True
         Height          =   525
         Left            =   5280
         TabIndex        =   1
         Top             =   2400
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   926
         _Version        =   131073
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdOK 
         Height          =   525
         Left            =   3360
         TabIndex        =   0
         Top             =   2400
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         ButtonStyle     =   3
      End
   End
End
Attribute VB_Name = "frmAddEditComTypeDonStk"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private m_HasActivate As Boolean
Private m_HasModify As Boolean
Private m_Rs As ADODB.Recordset
Private ItemComDonStk  As CComDonStk

Public ID As Long
Public OKClick As Boolean
Public ShowMode As SHOW_MODE_TYPE
Public HeaderText As String

Public COMDONSTK_ID As Long
Public MASTER_COMDONSTK_ID As Long

Public m_comDonStk As Collection
Public StepFlag  As Boolean

Public ParentForm As Form
Public itemCountGrid As Long
Public TempCollection As Collection                     ' มี TempCollection  เต็มเลย ต้องเพิ่ม

Private Sub cboDataType_Click()
   m_HasModify = True
End Sub
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
         
         If ShowMode = SHOW_EDIT Then
            Dim Bd As CComDonStk
            Set Bd = TempCollection.ITEM(ID)
            uctlStkCodStkdesLookup.MyTextBox.Text = Bd.STKCOD
         End If
   End If
   
   Call EnableForm(Me, True)
End Sub
Private Function SaveData() As Boolean
Dim IsOK As Boolean
Dim m_Lookup As CComDonStk
  Set m_Lookup = New CComDonStk
   
   If Not (uctlStkCodStkdesLookup.MyCombo.ItemData(Minus2Zero(uctlStkCodStkdesLookup.MyCombo.ListIndex)) > 0) Then
      Exit Function
   End If

   If Not m_HasModify Then
      SaveData = True
      Exit Function
   End If
   
      Dim Bd As CComDonStk
   If ShowMode = SHOW_ADD Then
      Set Bd = New CComDonStk
      Bd.Flag = "A"
      Call TempCollection.Add(Bd)
   Else
      Set Bd = TempCollection.ITEM(ID)
      If Bd.Flag <> "A" Then
         Bd.Flag = "E"
      End If
   End If
   
   Bd.AddEditMode = ShowMode
   Bd.MASTER_COMDONSTK_ID = MASTER_COMDONSTK_ID
   Bd.STKCOD = uctlStkCodStkdesLookup.MyTextBox.Text
   Bd.STKDES = uctlStkCodStkdesLookup.MyCombo.Text

   SaveData = True
End Function
Private Sub Form_Activate()
   If Not m_HasActivate Then
      m_HasActivate = True
      Me.Refresh
      DoEvents
      
      Call LoadStkcodLookup(uctlStkCodStkdesLookup.MyCombo, m_comDonStk)    'โหลด STK ....  หาสิ m_comDonStkคืออะไร
      Set uctlStkCodStkdesLookup.MyCollection = m_comDonStk
      
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
   
   Call InitNormalLabel(lblStkcodStkdes, MapText("เลขที่สินค้า"))
      
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
      
   Set ItemComDonStk = New CComDonStk
   Set m_comDonStk = New Collection
   
   m_HasActivate = False
   Call InitFormLayout
   Call EnableForm(Me, True)
End Sub
Private Sub Form_Unload(Cancel As Integer)
   Set ItemComDonStk = Nothing
   Set m_comDonStk = Nothing
End Sub

Private Sub uctlStkCodStkdesLookup_HasChange()
   m_HasModify = True
End Sub
Private Sub uctlStkCodStkdesLookup_Change()
   m_HasModify = True
End Sub
