VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#2.0#0"; "THREED20.OCX"
Begin VB.Form frmAddEditCarkillFW 
   BackColor       =   &H80000000&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   3645
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7275
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3645
   ScaleWidth      =   7275
   StartUpPosition =   1  'CenterOwner
   Begin Threed.SSFrame SSFrame1 
      Height          =   3645
      Left            =   0
      TabIndex        =   5
      Top             =   0
      Width           =   7305
      _ExtentX        =   12885
      _ExtentY        =   6429
      _Version        =   131073
      PictureBackgroundStyle=   2
      Begin Threed.SSPanel pnlHeader 
         Height          =   705
         Left            =   0
         TabIndex        =   6
         Top             =   0
         Width           =   7275
         _ExtentX        =   12832
         _ExtentY        =   1244
         _Version        =   131073
         PictureBackgroundStyle=   2
      End
      Begin prjLedgerReport.uctlTextBox txtFwRow 
         Height          =   435
         Left            =   2880
         TabIndex        =   0
         Tag             =   "2"
         Top             =   1200
         Width           =   2835
         _ExtentX        =   5212
         _ExtentY        =   767
      End
      Begin Threed.SSCheck chkMain 
         Height          =   375
         Left            =   2880
         TabIndex        =   2
         Top             =   2280
         Width           =   2655
         _ExtentX        =   4683
         _ExtentY        =   661
         _Version        =   131073
         Caption         =   "SSCheck1"
         TripleState     =   -1  'True
      End
      Begin Threed.SSCheck chkPlusUpper 
         Height          =   375
         Left            =   2880
         TabIndex        =   1
         Top             =   1800
         Width           =   2655
         _ExtentX        =   4683
         _ExtentY        =   661
         _Version        =   131073
         Caption         =   "SSCheck1"
         TripleState     =   -1  'True
      End
      Begin VB.Label lblFwRow 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   435
         Left            =   840
         TabIndex        =   7
         Top             =   1200
         Width           =   1935
      End
      Begin Threed.SSCommand cmdExit 
         Cancel          =   -1  'True
         Height          =   525
         Left            =   3720
         TabIndex        =   4
         Top             =   2880
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   926
         _Version        =   131073
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdOK 
         Height          =   525
         Left            =   2040
         TabIndex        =   3
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
Attribute VB_Name = "frmAddEditCarkillFW"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private m_HasActivate As Boolean
Private m_HasModify As Boolean
Private m_Rs As ADODB.Recordset
Private ItemXlsFW  As CXlsCarkillFW

Public ID As Long
Public OKClick As Boolean
Public ShowMode As SHOW_MODE_TYPE
Public HeaderText As String

Public ParentForm As Form
Public itemCountGrid As Long

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
      
      ItemXlsFW.XLS_FORWARD_ID = ID
      If Not glbDaily.QueryCarkillFW(ItemXlsFW, m_Rs, ItemCount, IsOK, glbErrorLog) Then
         glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
         Call EnableForm(Me, True)
         Exit Sub
      End If
   End If
   
   If ItemCount > 0 Then
      Call ItemXlsFW.PopulateFromRS(1, m_Rs)
         txtFwRow.Text = ItemXlsFW.FW_ROW
         chkMain.Value = FlagToCheck(ItemXlsFW.MAIN_FLAG)
         chkPlusUpper.Value = FlagToCheck(ItemXlsFW.UPPER_FLAG)
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

   If Not VerifyTextControl(lblFwRow, txtFwRow, False) Then
      Exit Function
   End If

   If Not m_HasModify Then
      SaveData = True
      Exit Function
   End If
   
   ItemXlsFW.AddEditMode = ShowMode
   ItemXlsFW.FW_ROW = txtFwRow.Text
   ItemXlsFW.MAIN_FLAG = Check2Flag(chkMain.Value)
   ItemXlsFW.UPPER_FLAG = Check2Flag(chkPlusUpper.Value)
   
   Call EnableForm(Me, False)
   
   If Not glbDaily.AddEditCarkillFW(ItemXlsFW, IsOK, True, glbErrorLog) Then
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
   
   Call InitNormalLabel(lblFwRow, MapText("แถวที่"))
   Call txtFwRow.SetTextLenType(TEXT_INTEGER, glbSetting.CODE_TYPE)
   Call InitCheckBox(chkMain, "เป็นบรรทัดคงเหลือยกมา")
   Call InitCheckBox(chkPlusUpper, "คำนวนจากซื้อ-ขาย")
      
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
   OKClick = False
   Unload Me
End Sub
Private Sub Form_Load()
   Call EnableForm(Me, False)
   m_HasActivate = False
   Set m_Rs = New ADODB.Recordset
      
   Set ItemXlsFW = New CXlsCarkillFW
'   Set m_Stcrd = New Collection
   
   m_HasActivate = False
   Call InitFormLayout
   Call EnableForm(Me, True)
End Sub
Private Sub Form_Unload(Cancel As Integer)
   Set ItemXlsFW = Nothing
'   Set m_Stcrd = Nothing
End Sub

Private Sub txtFwRow_Change()
   m_HasModify = True
End Sub
Private Sub chkMain_Click(Value As Integer)
   m_HasModify = True
End Sub
Private Sub chkPlusUpper_Click(Value As Integer)
   m_HasModify = True
End Sub
