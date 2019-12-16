VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#2.0#0"; "THREED20.OCX"
Begin VB.Form frmAddEditSubComBudget 
   BackColor       =   &H80000000&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   2985
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6735
   Icon            =   "frmAddEditSubComBudget.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2985
   ScaleWidth      =   6735
   StartUpPosition =   1  'CenterOwner
   Begin Threed.SSFrame SSFrame1 
      Height          =   3045
      Left            =   0
      TabIndex        =   5
      Top             =   0
      Width           =   6825
      _ExtentX        =   12039
      _ExtentY        =   5371
      _Version        =   131073
      PictureBackgroundStyle=   2
      Begin VB.ComboBox cboAreaCod 
         Height          =   315
         Left            =   2400
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   1320
         Width           =   2715
      End
      Begin VB.ComboBox cboSaleName 
         Height          =   315
         Left            =   2400
         Style           =   2  'Dropdown List
         TabIndex        =   0
         Top             =   840
         Width           =   2715
      End
      Begin Threed.SSPanel pnlHeader 
         Height          =   705
         Left            =   0
         TabIndex        =   6
         Top             =   0
         Width           =   6795
         _ExtentX        =   11986
         _ExtentY        =   1244
         _Version        =   131073
         PictureBackgroundStyle=   2
      End
      Begin prjLedgerReport.uctlTextBox txtBudget 
         Height          =   435
         Left            =   2400
         TabIndex        =   2
         Top             =   1800
         Width           =   2715
         _ExtentX        =   4789
         _ExtentY        =   767
      End
      Begin VB.Label lblSaleName 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   255
         Left            =   360
         TabIndex        =   9
         Top             =   840
         Width           =   1935
      End
      Begin VB.Label lblAreaCod 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   435
         Left            =   360
         TabIndex        =   8
         Top             =   1320
         Width           =   1935
      End
      Begin VB.Label lblBudget 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   435
         Left            =   360
         TabIndex        =   7
         Top             =   1800
         Width           =   1935
      End
      Begin Threed.SSCommand cmdExit 
         Cancel          =   -1  'True
         Height          =   525
         Left            =   3435
         TabIndex        =   4
         Top             =   2430
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   926
         _Version        =   131073
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdOK 
         Height          =   525
         Left            =   1785
         TabIndex        =   3
         Top             =   2430
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmAddEditSubComBudget.frx":27A2
         ButtonStyle     =   3
      End
   End
End
Attribute VB_Name = "frmAddEditSubComBudget"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private m_HasActivate As Boolean
Private m_HasModify As Boolean
Private m_Rs As ADODB.Recordset
Private supComBudget As CSupCombudget

Public SubID As Long
Public ID As Long
Public OKClick As Boolean
Public ShowMode As SHOW_MODE_TYPE
Public HeaderText As String
Public TempAddEditDataCollection As Collection
Public TempComBudgetCollection As Collection
Public m_SaleName As Collection
Public m_Area As Collection

Private Sub cboEnterpriseName_Click()
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
Dim TempSupComBudget As CSupCombudget
Set TempSupComBudget = New CSupCombudget

   Call EnableForm(Me, False)
   Set TempSupComBudget = TempAddEditDataCollection.ITEM(SubID)

   cboSaleName.ListIndex = IDToListIndex(cboSaleName, TempSupComBudget.SLM_ID)
   cboAreaCod.ListIndex = IDToListIndex(cboAreaCod, TempSupComBudget.MASTER_AREA_ID)
   txtBudget.Text = TempSupComBudget.BUDGET
   
   Call EnableForm(Me, True)
End Sub
Private Function SaveData() As Boolean
Dim IsOK As Boolean
Dim EnpAddress As CSupCombudget
Set EnpAddress = New CSupCombudget
   
   If Not VerifyCombo(lblSaleName, cboSaleName, False) Then
      Exit Function
   End If
   If Not VerifyCombo(lblAreaCod, cboAreaCod, False) Then
      Exit Function
   End If
   If Not VerifyTextControl(lblBudget, txtBudget, False) Then
      Exit Function
   End If
   
   If Not m_HasModify Then
      SaveData = True
      Exit Function
   End If
      
   Call EnableForm(Me, False)
   
   Dim TempComBudget As CCombudget
   Set TempComBudget = New CCombudget

   Set TempComBudget = GetObject("CSupCombudget", TempComBudgetCollection, Trim(cboSaleName.ItemData(Minus2Zero(cboSaleName.ListIndex))))
   If ShowMode = SHOW_ADD Then
      Set supComBudget = New CSupCombudget
      supComBudget.Flag = "A"
   
     Call TempAddEditDataCollection.Add(supComBudget)
   Else
      Set supComBudget = TempAddEditDataCollection.ITEM(SubID)
      If supComBudget.Flag <> "A" Then
         supComBudget.Flag = "E"
      End If
   End If
   
' สำหรับ Area ล่ะ
'   Set TempComBudget = GetObject("CSupCombudget", TempComBudgetCollection, Trim(cboSaleName.ItemData(Minus2Zero(cboSaleName.ListIndex))))
'   If ShowMode = SHOW_ADD Then
'      Set supComBudget = New CSupCombudget
'      supComBudget.Flag = "A"
'
'     Call TempAddEditDataCollection.Add(supComBudget)
'   Else
'      Set supComBudget = TempAddEditDataCollection.ITEM(SubID)
'      If supComBudget.Flag <> "A" Then
'         supComBudget.Flag = "E"
'      End If
'   End If
   
   supComBudget.COM_BUDGET_ID = cboSaleName.ItemData(Minus2Zero(cboSaleName.ListIndex))
'   supComBudget.ENTERPRISE_CODE = TempComBudget.ENTERPRISE_CODE
   supComBudget.COM_BUDGET_ID = cboAreaCod.ItemData(Minus2Zero(cboAreaCod.ListIndex))
   supComBudget.BUDGET = txtBudget.Text

   Call EnableForm(Me, True)
   SaveData = True
End Function
Private Sub Form_Activate()
   If Not m_HasActivate Then
      m_HasActivate = True
      Me.Refresh
      DoEvents
      
         Call LoadSale(cboSaleName, m_SaleName)       ' เอาเป็นตัวอย่างได้
         Call LoadArea(cboAreaCod)
      
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
   
   Call InitNormalLabel(lblSaleName, MapText("ชื่อพนักงานขาย"))
   Call InitNormalLabel(lblAreaCod, MapText("เขต"))
   Call InitNormalLabel(lblBudget, MapText("งบประมาณ"))
   
   Call txtBudget.SetTextLenType(TEXT_STRING, glbSetting.CODE_TYPE)
   
   Call InitCombo(cboSaleName)
   Call InitCombo(cboAreaCod)
   
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
   
   Set supComBudget = New CSupCombudget
   Set m_Rs = New ADODB.Recordset
   Set TempComBudgetCollection = New Collection
   
   m_HasActivate = False
   Call InitFormLayout
   Call EnableForm(Me, True)
End Sub
Private Sub Form_Unload(Cancel As Integer)
   Set supComBudget = Nothing
   Set TempComBudgetCollection = Nothing
End Sub

Private Sub txtBudget_Change()
   m_HasModify = True
End Sub

