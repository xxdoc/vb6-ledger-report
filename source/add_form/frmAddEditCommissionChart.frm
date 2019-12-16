VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#2.0#0"; "THREED20.OCX"
Begin VB.Form frmAddEditCommissionChart 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   4785
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9885
   Icon            =   "frmAddEditCommissionChart.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   4785
   ScaleWidth      =   9885
   StartUpPosition =   1  'CenterOwner
   Begin Threed.SSFrame SSFrame1 
      Height          =   4245
      Left            =   0
      TabIndex        =   6
      Top             =   540
      Width           =   9855
      _ExtentX        =   17383
      _ExtentY        =   7488
      _Version        =   131073
      PictureBackgroundStyle=   2
      Begin VB.ComboBox cboGoodsGroup 
         Height          =   315
         Left            =   2520
         Style           =   2  'Dropdown List
         TabIndex        =   12
         Top             =   2310
         Width           =   3855
      End
      Begin VB.ComboBox cboAreaCod 
         Height          =   315
         Left            =   2520
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   1800
         Width           =   3795
      End
      Begin VB.ComboBox cboParent 
         Height          =   315
         Left            =   2520
         Style           =   2  'Dropdown List
         TabIndex        =   0
         Top             =   840
         Width           =   3855
      End
      Begin prjLedgerReport.uctlTextBox txtBudget 
         Height          =   495
         Left            =   2520
         TabIndex        =   3
         Top             =   2760
         Width           =   2895
         _ExtentX        =   5106
         _ExtentY        =   873
      End
      Begin prjLedgerReport.uctlTextLookup uctlCommissionSale 
         Height          =   375
         Left            =   2520
         TabIndex        =   1
         Top             =   1320
         Width           =   6255
         _ExtentX        =   11033
         _ExtentY        =   661
      End
      Begin VB.Label lblGoodsGroup 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "AngsanaUPC"
            Size            =   14.25
            Charset         =   222
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   840
         TabIndex        =   13
         Top             =   2280
         Width           =   1605
      End
      Begin VB.Label lblAreaCod 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   435
         Left            =   480
         TabIndex        =   11
         Top             =   1800
         Width           =   1935
      End
      Begin VB.Label lblSaleName 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   435
         Left            =   480
         TabIndex        =   10
         Top             =   1320
         Width           =   1935
      End
      Begin VB.Label lblBudget 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   360
         TabIndex        =   9
         Top             =   2760
         Width           =   2115
      End
      Begin Threed.SSCommand cmdCancel 
         Cancel          =   -1  'True
         Height          =   525
         Left            =   4080
         TabIndex        =   5
         Top             =   3480
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   926
         _Version        =   131073
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdOK 
         Height          =   525
         Left            =   2400
         TabIndex        =   4
         Top             =   3480
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmAddEditCommissionChart.frx":08CA
         ButtonStyle     =   3
      End
      Begin VB.Label lblParent 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "AngsanaUPC"
            Size            =   14.25
            Charset         =   222
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   780
         TabIndex        =   7
         Top             =   810
         Width           =   1605
      End
   End
   Begin Threed.SSPanel pnlHeader 
      Height          =   615
      Left            =   0
      TabIndex        =   8
      Top             =   0
      Width           =   9945
      _ExtentX        =   17542
      _ExtentY        =   1085
      _Version        =   131073
      PictureBackgroundStyle=   2
   End
End
Attribute VB_Name = "frmAddEditCommissionChart"
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
Public FK_ID As Long
Public ShowMode As SHOW_MODE_TYPE
Private m_Rs As ADODB.Recordset

Private m_CommissionChart As CCommissionChart
Private m_CommissionCharts As Collection

Private m_SaleName As Collection
Dim TempSlm As COESLM
Private m_ParentName As Collection
Dim TempParent As COESLM

Private FtSaleColl As Collection
Private FtReturnColl As Collection
Private Sub cboParent_Click()
   m_HasModify = True
End Sub
Private Sub cboGoodsGroup_Click()
   m_HasModify = True
End Sub

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
   
   If Not VerifyCombo(lblParent, cboParent, True) Then
      Exit Sub
   End If

   If Not VerifyCombo(lblGoodsGroup, cboGoodsGroup, True) Then
      Exit Sub
   End If

   If Not (Len(uctlCommissionSale.MyCombo.Text) > 0) Then
      Exit Sub
   End If
   
   If Not VerifyTextControl(lblBudget, txtBudget, False) Then
      Exit Sub
   End If
 
   If Not m_HasModify Then
      Unload Me
      Exit Sub
   End If
   
 '  'debug.print m_CommissionChart.ShowMode
   m_CommissionChart.AddEditMode = ShowMode
   m_CommissionChart.MASTER_FROMTO_ID = FK_ID
'  'debug.print cboParent.ItemData(cboParent.ListIndex)
   If cboParent.ListIndex >= 0 Then
      m_CommissionChart.PARENT_ID = cboParent.ItemData(cboParent.ListIndex)
   Else
      m_CommissionChart.PARENT_ID = 0
   End If
  m_CommissionChart.GOODS_GROUP_ID = cboGoodsGroup.ItemData(cboGoodsGroup.ListIndex)

   m_CommissionChart.SALE_ID = uctlCommissionSale.MyTextBox.Text
    m_CommissionChart.MASTER_AREA_ID = Left(cboAreaCod.Text, 2)
   m_CommissionChart.BUDGET = txtBudget.Text
  Call EnableForm(Me, False)
   
   If Not glbDaily.AddEditCommissionChart(m_CommissionChart, IsOK, True, glbErrorLog) Then
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

      Call LoadAreaCom(cboAreaCod)
      Call LoadGoodsGroup(cboGoodsGroup)
      Call LoadCommissionChart(cboParent, m_CommissionCharts, FK_ID)
        
      Call LoadSaleLookup(uctlCommissionSale.MyCombo, FtSaleColl) 'FtSaleColl, COMMISSION_TABLE
      Set uctlCommissionSale.MyCollection = FtSaleColl
      
      If (ShowMode = SHOW_EDIT) Or (ShowMode = SHOW_VIEW_ONLY) Then
         Call EnableForm(Me, False)
         m_CommissionChart.COMMISSION_CHART_ID = ID
         m_CommissionChart.MASTER_FROMTO_ID = FK_ID
         If Not glbDaily.QueryCommissionChart(m_CommissionChart, m_Rs, ItemCount, IsOK, glbErrorLog) Then
            glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
            Call EnableForm(Me, True)
            Exit Sub
         End If
         
         If ItemCount > 0 Then
            Call m_CommissionChart.PopulateFromRS(1, m_Rs)
            If m_CommissionChart.PARENT_ID > 0 Then
                  cboParent.ListIndex = IDToListIndex(cboParent, m_CommissionChart.PARENT_ID)
            End If
            cboAreaCod.ListIndex = IDToListIndex(cboAreaCod, m_CommissionChart.MASTER_AREA_ID)
            cboGoodsGroup.ListIndex = IDToListIndex(cboGoodsGroup, m_CommissionChart.GOODS_GROUP_ID)
            uctlCommissionSale.MyTextBox.Text = m_CommissionChart.SALE_ID
            txtBudget.Text = m_CommissionChart.BUDGET
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
   Set m_CommissionChart = New CCommissionChart
   Set m_SaleName = New Collection
   Set TempSlm = New COESLM
   Set m_ParentName = New Collection
   Set TempParent = New COESLM
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
   
   Call InitNormalLabel(lblParent, MapText("ภายใต้"))
'   Call InitNormalLabel(lblEmployee, MapText("Manager"))
   Call InitNormalLabel(lblSaleName, MapText("พนักงานขาย"))
   Call InitNormalLabel(lblAreaCod, MapText("เขต"))
   Call InitNormalLabel(lblBudget, MapText("งบประมาณ"))
   Call InitNormalLabel(lblGoodsGroup, MapText("ประเภทสินค้า"))
'   Call InitNormalLabel(lblOrderID, MapText("ลำดับ"))
   
''   Call txtOrderID.SetTextLenType(TEXT_INTEGER, glbSetting.ID_TYPE)
   Call txtBudget.SetTextLenType(TEXT_INTEGER, glbSetting.ID_TYPE)
   
   Call InitCombo(cboParent)
'   Call InitCombo(cboSaleName)
   Call InitCombo(cboAreaCod)
   Call InitCombo(cboGoodsGroup)
   
   cmdCancel.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdOK.Picture = LoadPicture(glbParameterObj.NormalButton1)
   
   Call InitMainButton(cmdOK, MapText("ตกลง (F2)"))
   Call InitMainButton(cmdCancel, MapText("ยกเลิก (ESC)"))
End Sub
Private Sub Form_Unload(Cancel As Integer)
   Set m_SaleName = Nothing
   Set TempSlm = Nothing
   Set m_ParentName = Nothing
   Set TempParent = Nothing
   Set m_CommissionChart = Nothing
   If m_Rs.State = adStateOpen Then
      m_Rs.Close
   End If
   Set m_Rs = Nothing
   Set FtSaleColl = Nothing
   Set FtReturnColl = Nothing
End Sub

Private Sub cboAreaCod_Click()
   m_HasModify = True
End Sub

Private Sub txtBudget_Change()
   m_HasModify = True
End Sub

Private Sub uctlCommissionSale_HasChange()
   m_HasModify = True
End Sub

Private Sub uctlCommissionSale_Change()
   m_HasModify = True
End Sub
