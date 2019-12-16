VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#2.0#0"; "THREED20.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmExportBill 
   BackColor       =   &H80000000&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   3480
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10605
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3480
   ScaleWidth      =   10605
   StartUpPosition =   1  'CenterOwner
   Begin Threed.SSFrame SSFrame1 
      Height          =   3525
      Left            =   0
      TabIndex        =   7
      Top             =   0
      Width           =   10665
      _ExtentX        =   18812
      _ExtentY        =   6218
      _Version        =   131073
      PictureBackgroundStyle=   2
      Begin MSComctlLib.ProgressBar prgProgress 
         Height          =   315
         Left            =   1860
         TabIndex        =   5
         Top             =   1950
         Width           =   8145
         _ExtentX        =   14367
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   1
      End
      Begin Threed.SSPanel pnlHeader 
         Height          =   705
         Left            =   0
         TabIndex        =   8
         Top             =   0
         Width           =   10635
         _ExtentX        =   18759
         _ExtentY        =   1244
         _Version        =   131073
         PictureBackgroundStyle=   2
      End
      Begin prjLedgerReport.uctlTextBox txtPercent 
         Height          =   465
         Left            =   1860
         TabIndex        =   6
         Top             =   2280
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   820
      End
      Begin MSComctlLib.ImageList ImageList1 
         Left            =   0
         Top             =   0
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   16
         ImageHeight     =   16
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   1
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmExportBill.frx":0000
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin prjLedgerReport.uctlTextBox txtFromInvNo 
         Height          =   465
         Left            =   1860
         TabIndex        =   0
         Top             =   840
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   820
      End
      Begin prjLedgerReport.uctlTextBox txtToInvNo 
         Height          =   465
         Left            =   1860
         TabIndex        =   1
         Top             =   1440
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   820
      End
      Begin Threed.SSCommand cmdStart 
         Height          =   525
         Left            =   1890
         TabIndex        =   2
         Top             =   2820
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         ButtonStyle     =   3
      End
      Begin VB.Label Label1 
         Caption         =   "Label1"
         Height          =   345
         Left            =   3600
         TabIndex        =   13
         Top             =   2400
         Width           =   1275
      End
      Begin VB.Label lblProgress 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   435
         Left            =   210
         TabIndex        =   12
         Top             =   2010
         Width           =   1575
      End
      Begin VB.Label lblPercent 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   435
         Left            =   210
         TabIndex        =   11
         Top             =   2430
         Width           =   1575
      End
      Begin VB.Label lblMasterName 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   435
         Left            =   210
         TabIndex        =   10
         Top             =   1500
         Width           =   1575
      End
      Begin VB.Label lblFileName 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   435
         Left            =   210
         TabIndex        =   9
         Top             =   1080
         Width           =   1575
      End
      Begin Threed.SSCommand cmdExit 
         Cancel          =   -1  'True
         Height          =   525
         Left            =   8535
         TabIndex        =   4
         Top             =   2820
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   926
         _Version        =   131073
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdOK 
         Height          =   525
         Left            =   6885
         TabIndex        =   3
         Top             =   2820
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         ButtonStyle     =   3
      End
   End
End
Attribute VB_Name = "frmExportBill"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private m_HasActivate As Boolean
Private m_HasModify As Boolean
Private m_Rs As ADODB.Recordset

Public ID As Long
Public OKClick As Boolean
Public ShowMode As SHOW_MODE_TYPE
Public HeaderText As String
Private Sub cmdOK_Click()
   OKClick = True
   Unload Me
End Sub
Private Sub Form_Activate()
Dim FromDate As Date
Dim ToDate As Date

   If Not m_HasActivate Then
      m_HasActivate = True
      Me.Refresh
      DoEvents
      
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
'   ElseIf Shift = 0 And KeyCode = 117 Then
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

Private Sub ResetStatus()
   prgProgress.MAX = 100
   prgProgress.MIN = 0
   prgProgress.Value = 0
   txtPercent.Text = 0
End Sub
Private Sub InitFormLayout()
   pnlHeader.PictureBackground = LoadPicture(glbParameterObj.NormalForm1)
   SSFrame1.PictureBackground = LoadPicture(glbParameterObj.NormalForm1)
   
   Me.Caption = HeaderText
   pnlHeader.Caption = "EXPORT ข้อมูล บิลขายไปยังสาขา"
   
   Me.Picture = LoadPicture(glbParameterObj.MainPicture)
   
   pnlHeader.Font.Name = GLB_FONT
   pnlHeader.Font.Bold = True
   pnlHeader.Font.Size = 19
   
   Call InitNormalLabel(lblFileName, "จาก INV NO")
   Call InitNormalLabel(lblMasterName, "ถึง INV NO")
   Call InitNormalLabel(lblProgress, "ความคืบหน้า")
   Call InitNormalLabel(lblPercent, "เปอร์เซนต์")
   Call InitNormalLabel(Label1, "%")

   Call txtPercent.SetTextLenType(TEXT_FLOAT, glbSetting.MONEY_TYPE)
   Call txtFromInvNo.SetTextLenType(TEXT_STRING, glbSetting.CODE_TYPE)
   Call txtToInvNo.SetTextLenType(TEXT_STRING, glbSetting.CODE_TYPE)
   txtPercent.Enabled = False
   
   cmdExit.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdOK.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdStart.Picture = LoadPicture(glbParameterObj.NormalButton1)
   
   Call InitMainButton(cmdExit, MapText("ยกเลิก (ESC)"))
   Call InitMainButton(cmdOK, MapText("ตกลง (F2)"))
   Call InitMainButton(cmdStart, MapText("เริ่ม"))
   
   Call ResetStatus
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
   
   m_HasActivate = False
   Call InitFormLayout
   Call EnableForm(Me, True)
End Sub
Private Sub cmdStart_Click()
Dim TempID As Long
Dim HasBegin As Boolean
   
   If Not VerifyTextControl(lblFileName, txtFromInvNo, False) Then
      Exit Sub
   End If
   If Not VerifyTextControl(lblMasterName, txtFromInvNo, False) Then
      Exit Sub
   End If
      
   Call EnableForm(Me, False)
   
   Call ExportBill
   
   Call EnableForm(Me, True)
End Sub
Private Sub ExportBill()
On Error GoTo ErrorHandler
Dim IsOK As Boolean
Dim Ar As CARTrn
Dim iCount As Long
Dim FileID As Long
Dim OldID As String
Dim i As Long

Dim LocationSave As String
   
   prgProgress.MAX = 100
   prgProgress.MIN = 0


   Set Ar = New CARTrn
   Ar.FROM_DOCNUM = txtFromInvNo.Text
   Ar.TO_DOCNUM = txtToInvNo.Text
   
   Call Ar.QueryData(10, m_Rs, iCount)
   
   LocationSave = "C:\ExportBranch\"
   
   If Dir(LocationSave, vbDirectory) = "" Then
      MkDir (LocationSave)
   End If
   
   LocationSave = LocationSave & Format(Year(Now), "0000") & Format(Month(Now), "00") & Format(Day(Now), "00")
   LocationSave = LocationSave & "_" & txtFromInvNo.Text
   LocationSave = LocationSave & "_" & txtToInvNo.Text
   LocationSave = LocationSave & ".txt"
   'LocationSave = "C:\1234.txt"
      
On Error GoTo XXX
   Call Kill(LocationSave)
XXX:


   FileID = FreeFile
   Open LocationSave For Append As #FileID
   
   If Not m_Rs.EOF Then
      Call Ar.GenerateArHeader(FileID, m_Rs)
      OldID = NVLS(m_Rs("DOCNUM"), "")
   End If
   
   prgProgress.MAX = 100
   prgProgress.MIN = 0
   prgProgress.Value = 0

   i = 0
   While Not m_Rs.EOF
      prgProgress.Value = MyDiff(i, iCount) * 100
      
      If OldID <> NVLS(m_Rs("DOCNUM"), "") Then
         Call Ar.GenerateArHeader(FileID, m_Rs)
         OldID = NVLS(m_Rs("DOCNUM"), "")
      End If

      'Generate detail here
      Call Ar.GenerateArDetail(FileID, m_Rs)
      m_Rs.MoveNext
      i = i + 1
   Wend
   Close #FileID

   Set Ar = Nothing
   prgProgress.Value = 100
   txtPercent.Text = 100

   Exit Sub

ErrorHandler:

   cmdStart.Enabled = True
   cmdExit.Enabled = True
   cmdOK.Enabled = True

   glbErrorLog.LocalErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowUserError
   Close #FileID
End Sub
