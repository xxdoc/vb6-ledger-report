VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#2.0#0"; "THREED20.OCX"
Begin VB.Form frmAddEditXlsUnit 
   BackColor       =   &H80000000&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   4440
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8025
   Icon            =   "frmAddEditXlsUnit.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4440
   ScaleWidth      =   8025
   StartUpPosition =   1  'CenterOwner
   Begin Threed.SSFrame SSFrame1 
      Height          =   4365
      Left            =   0
      TabIndex        =   5
      Top             =   0
      Width           =   8025
      _ExtentX        =   14155
      _ExtentY        =   7699
      _Version        =   131073
      PictureBackgroundStyle=   2
      Begin Threed.SSPanel pnlHeader 
         Height          =   705
         Left            =   0
         TabIndex        =   6
         Top             =   0
         Width           =   7995
         _ExtentX        =   14102
         _ExtentY        =   1244
         _Version        =   131073
         PictureBackgroundStyle=   2
      End
      Begin prjLedgerReport.uctlTextBox txtUnit 
         Height          =   435
         Left            =   2520
         TabIndex        =   0
         Top             =   1680
         Width           =   2925
         _ExtentX        =   5159
         _ExtentY        =   767
      End
      Begin prjLedgerReport.uctlTextBox txtMultiply 
         Height          =   435
         Left            =   2520
         TabIndex        =   1
         Top             =   2160
         Width           =   2925
         _ExtentX        =   5159
         _ExtentY        =   767
      End
      Begin prjLedgerReport.uctlTextBox txtLimit 
         Height          =   435
         Left            =   2520
         TabIndex        =   2
         Top             =   2640
         Width           =   2925
         _ExtentX        =   5159
         _ExtentY        =   767
      End
      Begin VB.Label lblLimit 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   435
         Left            =   360
         TabIndex        =   9
         Top             =   2640
         Width           =   1935
      End
      Begin Threed.SSCommand cmdOK 
         Height          =   525
         Left            =   2400
         TabIndex        =   3
         Top             =   3480
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmAddEditXlsUnit.frx":27A2
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdExit 
         Cancel          =   -1  'True
         Height          =   525
         Left            =   4080
         TabIndex        =   4
         Top             =   3480
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   926
         _Version        =   131073
         ButtonStyle     =   3
      End
      Begin VB.Label lblUnit 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   435
         Left            =   360
         TabIndex        =   8
         Top             =   1680
         Width           =   1935
      End
      Begin VB.Label lblMultiply 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   435
         Left            =   360
         TabIndex        =   7
         Top             =   2160
         Width           =   1935
      End
   End
End
Attribute VB_Name = "frmAddEditXlsUnit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private m_HasActivate As Boolean
Private m_HasModify As Boolean
Private m_Rs As ADODB.Recordset
Private xlsUnit As CXlsUnit

Public ID As Long
Public OKClick As Boolean
Public ShowMode As SHOW_MODE_TYPE
Public HeaderText As String
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
      
      xlsUnit.XLS_UNIT_ID = ID
      xlsUnit.QueryFlag = 1
      If Not glbDaily.QueryXlsUnit(xlsUnit, m_Rs, ItemCount, IsOK, glbErrorLog) Then
         glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
         Call EnableForm(Me, True)
         Exit Sub
      End If
   End If
   
   If ItemCount > 0 Then
      Call xlsUnit.PopulateFromRS(1, m_Rs)
      txtUnit.Text = xlsUnit.XLS_UNIT_NAME
      txtMultiply.Text = xlsUnit.XLS_UNIT_MULTIPLY
      txtLimit.Text = xlsUnit.XLS_UNIT_LIMIT
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
   
   If Not VerifyTextControl(lblUnit, txtUnit, False) Then
      Exit Function
   End If
   If Not VerifyTextControl(lblMultiply, txtMultiply, False) Then
      Exit Function
   End If
   If Not VerifyTextControl(lblLimit, txtLimit, False) Then
      Exit Function
   End If
   
'   If Not CheckUniqueNs(REAL_CREDIT_NO, txtGroupTypeCode.Text, ID) Then
'      glbErrorLog.LocalErrorMsg = MapText("มีข้อมูล") & " " & txtDocumentNo.Text & " " & MapText("อยู่ในระบบแล้ว")
'      glbErrorLog.ShowUserError
'      Exit Function
'   End If
   
   If Not m_HasModify Then
      SaveData = True
      Exit Function
   End If
   
   xlsUnit.AddEditMode = ShowMode
   xlsUnit.XLS_UNIT_ID = ID
   xlsUnit.XLS_UNIT_NAME = txtUnit.Text
   xlsUnit.XLS_UNIT_MULTIPLY = txtMultiply.Text
   xlsUnit.XLS_UNIT_LIMIT = txtLimit.Text
      
   Call EnableForm(Me, False)
   If Not glbDaily.AddEditXlsUnit(xlsUnit, IsOK, True, glbErrorLog) Then
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
   
   Call InitNormalLabel(lblUnit, MapText("หน่วย"))
   Call InitNormalLabel(lblMultiply, MapText("คูณ"))
   Call InitNormalLabel(lblLimit, MapText("ขีดจำกัดต่อคัน"))
   Call txtUnit.SetTextLenType(TEXT_STRING, glbSetting.NAME_TYPE)
   Call txtMultiply.SetTextLenType(TEXT_STRING, glbSetting.NAME_TYPE)
   Call txtLimit.SetTextLenType(TEXT_STRING, glbSetting.NAME_TYPE)
   
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
   
   Set xlsUnit = New CXlsUnit
   Set m_Rs = New ADODB.Recordset
   
   m_HasActivate = False
   Call InitFormLayout
   Call EnableForm(Me, True)
End Sub
Private Sub Form_Unload(Cancel As Integer)
   Set xlsUnit = Nothing
End Sub
Private Sub txtUnit_Change()
   m_HasModify = True
End Sub
Private Sub txtMultiply_Change()
   m_HasModify = True
End Sub
Private Sub txtLimit_Change()
   m_HasModify = True
End Sub
