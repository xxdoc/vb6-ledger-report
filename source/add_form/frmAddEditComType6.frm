VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#2.0#0"; "THREED20.OCX"
Begin VB.Form frmAddEditComType6 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   4110
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9885
   Icon            =   "frmAddEditComType6.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   4110
   ScaleWidth      =   9885
   StartUpPosition =   1  'CenterOwner
   Begin Threed.SSFrame SSFrame1 
      Height          =   3765
      Left            =   0
      TabIndex        =   3
      Top             =   600
      Width           =   9855
      _ExtentX        =   17383
      _ExtentY        =   6641
      _Version        =   131073
      PictureBackgroundStyle=   2
      Begin prjLedgerReport.uctlTextLookup uctlCommissionSale 
         Height          =   375
         Left            =   2640
         TabIndex        =   0
         Top             =   960
         Width           =   6255
         _ExtentX        =   11033
         _ExtentY        =   661
      End
      Begin Threed.SSCheck chkCountIncen 
         Height          =   375
         Left            =   2640
         TabIndex        =   7
         Top             =   1560
         Width           =   2655
         _ExtentX        =   4683
         _ExtentY        =   661
         _Version        =   131073
         Caption         =   "SSCheck1"
         TripleState     =   -1  'True
      End
      Begin VB.Label lblSaleName 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   435
         Left            =   480
         TabIndex        =   6
         Top             =   960
         Width           =   1935
      End
      Begin VB.Label lblBudget 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   360
         TabIndex        =   5
         Top             =   2280
         Width           =   2115
      End
      Begin Threed.SSCommand cmdCancel 
         Cancel          =   -1  'True
         Height          =   525
         Left            =   4560
         TabIndex        =   2
         Top             =   2160
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   926
         _Version        =   131073
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdOK 
         Height          =   525
         Left            =   2880
         TabIndex        =   1
         Top             =   2160
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmAddEditComType6.frx":08CA
         ButtonStyle     =   3
      End
   End
   Begin Threed.SSPanel pnlHeader 
      Height          =   615
      Left            =   0
      TabIndex        =   4
      Top             =   0
      Width           =   9945
      _ExtentX        =   17542
      _ExtentY        =   1085
      _Version        =   131073
      PictureBackgroundStyle=   2
   End
End
Attribute VB_Name = "frmAddEditComType6"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Const MODULE_NAME = "frmAddEditCommissionChart"

Private HasActivate As Boolean
Private m_HasModify As Boolean
Public HeaderText As String
Public OKClick As Boolean
Public ID As Long
Public YEAR_ID As Long
Public ShowMode As SHOW_MODE_TYPE
Private m_Rs As ADODB.Recordset

Private m_IncenSum As CCondiIncenSum
Private m_IncenSums As Collection

Private m_SaleName As Collection
Dim TempSlm As COESLM

Private FtSaleColl As Collection
Private FtReturnColl As Collection


Private Sub cmdCancel_Click()
   If Not ConfirmExit(m_HasModify) Then
      Exit Sub
   End If
   OKClick = False
   Unload Me
End Sub

Private Sub cmdOK_Click()
On Error GoTo ErrorHandler
Dim IsOK As Boolean
   
   glbErrorLog.ModuleName = MODULE_NAME
   glbErrorLog.RoutineName = "Form_Activate"

'
   If Not (Len(uctlCommissionSale.MyCombo.Text) > 0) Then
      Exit Sub
   End If

   If Not m_HasModify Then
      Unload Me
      Exit Sub
   End If
   
 '  'debug.print m_IncenSum.ShowMode
   m_IncenSum.AddEditMode = ShowMode
   m_IncenSum.YEAR_ID = YEAR_ID
   m_IncenSum.SLMCOD = uctlCommissionSale.MyTextBox.Text
   m_IncenSum.SLMNAME = uctlCommissionSale.MyCombo.Text
   m_IncenSum.Flag = Check2Flag(chkCountIncen.Value)
    m_IncenSum.FORSUM_TYP = "06"
  Call EnableForm(Me, False)
   
   If Not glbDaily.AddEditIncenForSum(m_IncenSum, IsOK, True, glbErrorLog) Then
      glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
      Call EnableForm(Me, True)
      Exit Sub
   End If
   If Not IsOK Then
      glbErrorLog.ShowUserError
      Call EnableForm(Me, True)
      Exit Sub
   End If
   Call EnableForm(Me, True)
   
   OKClick = True
   Unload Me
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
   Call EnableForm(Me, True)
End Sub

Private Sub Form_Activate()
On Error GoTo ErrorHandler
Dim ItemCount As Long
Dim IsOK As Boolean

   glbErrorLog.ModuleName = MODULE_NAME
   
   glbErrorLog.RoutineName = "Form_Load"

   If Not HasActivate Then
      HasActivate = True
      Me.Refresh

      Call LoadSaleLookup(uctlCommissionSale.MyCombo, FtSaleColl) 'FtSaleColl, COMMISSION_TABLE
      Set uctlCommissionSale.MyCollection = FtSaleColl
      
      If (ShowMode = SHOW_EDIT) Or (ShowMode = SHOW_VIEW_ONLY) Then
         Call EnableForm(Me, False)
         m_IncenSum.INCEN_SLM_FORSUM_ID = ID
         m_IncenSum.YEAR_ID = YEAR_ID
         m_IncenSum.FROM_CMPL_DATE = -1
         m_IncenSum.TO_CMPL_DATE = -1
         If Not glbDaily.QueryIncenForSum(m_IncenSum, m_Rs, ItemCount, IsOK, glbErrorLog) Then
            glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
            Call EnableForm(Me, True)
            Exit Sub
         End If
         
         If ItemCount > 0 Then
            Call m_IncenSum.PopulateFromRS(1, m_Rs)
           uctlCommissionSale.MyTextBox.Text = m_IncenSum.SLMCOD
           chkCountIncen.Value = FlagToCheck(m_IncenSum.Flag)
         End If
         Call EnableForm(Me, True)
         m_HasModify = False
      End If
   End If
   
Call EnableForm(Me, True)
Exit Sub
   
ErrorHandler:
   Call EnableForm(Me, True)
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
   If Shift = 1 And KeyCode = DUMMY_KEY Then
      MsgBox Me.Name
   ElseIf Shift = 1 And KeyCode = 112 Then
      If glbUser.EXCEPTION_FLAG = "Y" Then
         glbUser.EXCEPTION_FLAG = "N"
      Else
         glbUser.EXCEPTION_FLAG = "Y"
      End If
   ElseIf Shift = 0 And KeyCode = 113 Then
      Call cmdOK_Click
   End If
End Sub
Private Sub Form_Load()
   Set m_Rs = New ADODB.Recordset
   Set m_IncenSum = New CCondiIncenSum
   Set m_SaleName = New Collection
   Set FtSaleColl = New Collection
   Set FtReturnColl = New Collection

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
   
   Call InitNormalLabel(lblSaleName, MapText("พนักงานขาย"))
   Call InitCheckBox(chkCountIncen, "นับยอด Incentive")
   
   cmdCancel.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdOK.Picture = LoadPicture(glbParameterObj.NormalButton1)
   
   Call InitMainButton(cmdOK, MapText("ตกลง (F2)"))
   Call InitMainButton(cmdCancel, MapText("ยกเลิก (ESC)"))
End Sub
Private Sub Form_Unload(Cancel As Integer)
   Set m_SaleName = Nothing
   Set m_IncenSum = Nothing
      Set FtSaleColl = Nothing
   Set FtReturnColl = Nothing
   If m_Rs.State = adStateOpen Then
      m_Rs.Close
   End If
   Set m_Rs = Nothing
End Sub

Private Sub uctlCommissionSale_HasChange()
   m_HasModify = True
End Sub

Private Sub uctlCommissionSale_Change()
   m_HasModify = True
End Sub

Private Sub chkCountIncen_Click(Value As Integer)
   m_HasModify = True
End Sub
