VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#2.0#0"; "THREED20.OCX"
Begin VB.Form frmAddEditCarkillSum 
   BackColor       =   &H80000000&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   6045
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   12840
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6045
   ScaleWidth      =   12840
   StartUpPosition =   1  'CenterOwner
   Begin Threed.SSFrame SSFrame1 
      Height          =   8325
      Left            =   0
      TabIndex        =   21
      Top             =   0
      Width           =   12825
      _ExtentX        =   22622
      _ExtentY        =   14684
      _Version        =   131073
      PictureBackgroundStyle=   2
      Begin prjLedgerReport.uctlTextBox txtpRow1 
         Height          =   375
         Left            =   10800
         TabIndex        =   3
         Top             =   1680
         Width           =   975
         _ExtentX        =   5106
         _ExtentY        =   873
      End
      Begin Threed.SSPanel pnlHeader 
         Height          =   705
         Left            =   0
         TabIndex        =   22
         Top             =   0
         Width           =   12795
         _ExtentX        =   22569
         _ExtentY        =   1244
         _Version        =   131073
         PictureBackgroundStyle=   2
      End
      Begin prjLedgerReport.uctlTextBox txtpColumn1 
         Height          =   375
         Left            =   7800
         TabIndex        =   2
         Tag             =   "2"
         Top             =   1680
         Width           =   975
         _ExtentX        =   2566
         _ExtentY        =   661
      End
      Begin prjLedgerReport.uctlTextBox txtOperator1 
         Height          =   375
         Left            =   4920
         TabIndex        =   1
         Tag             =   "2"
         Top             =   1680
         Width           =   975
         _ExtentX        =   2566
         _ExtentY        =   661
      End
      Begin prjLedgerReport.uctlTextBox txtSumRow 
         Height          =   435
         Left            =   3000
         TabIndex        =   0
         Tag             =   "2"
         Top             =   1080
         Width           =   2835
         _ExtentX        =   5212
         _ExtentY        =   767
      End
      Begin prjLedgerReport.uctlTextBox txtpRow2 
         Height          =   375
         Left            =   10800
         TabIndex        =   6
         Top             =   2160
         Width           =   975
         _ExtentX        =   5106
         _ExtentY        =   873
      End
      Begin prjLedgerReport.uctlTextBox txtpColumn2 
         Height          =   375
         Left            =   7800
         TabIndex        =   5
         Tag             =   "2"
         Top             =   2160
         Width           =   975
         _ExtentX        =   2566
         _ExtentY        =   661
      End
      Begin prjLedgerReport.uctlTextBox txtOperator2 
         Height          =   375
         Left            =   4920
         TabIndex        =   4
         Tag             =   "2"
         Top             =   2160
         Width           =   975
         _ExtentX        =   2566
         _ExtentY        =   661
      End
      Begin prjLedgerReport.uctlTextBox txtpRow3 
         Height          =   375
         Left            =   10800
         TabIndex        =   9
         Top             =   2640
         Width           =   975
         _ExtentX        =   5106
         _ExtentY        =   873
      End
      Begin prjLedgerReport.uctlTextBox txtpColumn3 
         Height          =   375
         Left            =   7800
         TabIndex        =   8
         Tag             =   "2"
         Top             =   2640
         Width           =   975
         _ExtentX        =   2566
         _ExtentY        =   661
      End
      Begin prjLedgerReport.uctlTextBox txtOperator3 
         Height          =   375
         Left            =   4920
         TabIndex        =   7
         Tag             =   "2"
         Top             =   2640
         Width           =   975
         _ExtentX        =   2566
         _ExtentY        =   661
      End
      Begin prjLedgerReport.uctlTextBox txtpRow4 
         Height          =   375
         Left            =   10800
         TabIndex        =   12
         Top             =   3120
         Width           =   975
         _ExtentX        =   5106
         _ExtentY        =   873
      End
      Begin prjLedgerReport.uctlTextBox txtpColumn4 
         Height          =   375
         Left            =   7800
         TabIndex        =   11
         Tag             =   "2"
         Top             =   3120
         Width           =   975
         _ExtentX        =   2566
         _ExtentY        =   661
      End
      Begin prjLedgerReport.uctlTextBox txtOperator4 
         Height          =   375
         Left            =   4920
         TabIndex        =   10
         Tag             =   "2"
         Top             =   3120
         Width           =   975
         _ExtentX        =   2566
         _ExtentY        =   661
      End
      Begin prjLedgerReport.uctlTextBox txtpRow5 
         Height          =   375
         Left            =   10800
         TabIndex        =   15
         Top             =   3600
         Width           =   975
         _ExtentX        =   5106
         _ExtentY        =   873
      End
      Begin prjLedgerReport.uctlTextBox txtpColumn5 
         Height          =   375
         Left            =   7800
         TabIndex        =   14
         Tag             =   "2"
         Top             =   3600
         Width           =   975
         _ExtentX        =   2566
         _ExtentY        =   661
      End
      Begin prjLedgerReport.uctlTextBox txtOperator5 
         Height          =   375
         Left            =   4920
         TabIndex        =   13
         Tag             =   "2"
         Top             =   3600
         Width           =   975
         _ExtentX        =   2566
         _ExtentY        =   661
      End
      Begin Threed.SSCheck chkHorizontal 
         Height          =   375
         Left            =   4920
         TabIndex        =   16
         Top             =   4080
         Width           =   2655
         _ExtentX        =   4683
         _ExtentY        =   661
         _Version        =   131073
         Caption         =   "SSCheck1"
         TripleState     =   -1  'True
      End
      Begin VB.Label lblComment 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   375
         Left            =   3000
         TabIndex        =   38
         Top             =   4560
         Width           =   8775
      End
      Begin VB.Label lblpRow5 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   375
         Left            =   9000
         TabIndex        =   37
         Top             =   3600
         Width           =   1695
      End
      Begin VB.Label lblpColumn5 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   375
         Left            =   6120
         TabIndex        =   36
         Top             =   3600
         Width           =   1575
      End
      Begin VB.Label lblOperator5 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   375
         Left            =   3000
         TabIndex        =   35
         Top             =   3600
         Width           =   1815
      End
      Begin VB.Label lblpRow4 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   375
         Left            =   9000
         TabIndex        =   34
         Top             =   3120
         Width           =   1695
      End
      Begin VB.Label lblpColumn4 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   375
         Left            =   6120
         TabIndex        =   33
         Top             =   3120
         Width           =   1575
      End
      Begin VB.Label lblOperator4 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   375
         Left            =   3000
         TabIndex        =   32
         Top             =   3120
         Width           =   1815
      End
      Begin VB.Label lblpRow3 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   375
         Left            =   9000
         TabIndex        =   31
         Top             =   2640
         Width           =   1695
      End
      Begin VB.Label lblpColumn3 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   375
         Left            =   6120
         TabIndex        =   30
         Top             =   2640
         Width           =   1575
      End
      Begin VB.Label lblOperator3 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   375
         Left            =   3000
         TabIndex        =   29
         Top             =   2640
         Width           =   1815
      End
      Begin VB.Label lblpRow2 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   375
         Left            =   9000
         TabIndex        =   28
         Top             =   2160
         Width           =   1695
      End
      Begin VB.Label lblpColumn2 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   375
         Left            =   6120
         TabIndex        =   27
         Top             =   2160
         Width           =   1575
      End
      Begin VB.Label lblOperator2 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   375
         Left            =   3000
         TabIndex        =   26
         Top             =   2160
         Width           =   1815
      End
      Begin VB.Label lblSumRow 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   435
         Left            =   840
         TabIndex        =   25
         Top             =   1200
         Width           =   1935
      End
      Begin VB.Label lblOperator1 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   375
         Left            =   3000
         TabIndex        =   24
         Top             =   1680
         Width           =   1815
      End
      Begin VB.Label lblpColumn1 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   375
         Left            =   6120
         TabIndex        =   23
         Top             =   1680
         Width           =   1575
      End
      Begin VB.Label lblDetail 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   375
         Left            =   720
         TabIndex        =   19
         Top             =   1680
         Width           =   2055
      End
      Begin VB.Label lblpRow1 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   375
         Left            =   9000
         TabIndex        =   20
         Top             =   1680
         Width           =   1695
      End
      Begin Threed.SSCommand cmdExit 
         Cancel          =   -1  'True
         Height          =   525
         Left            =   6840
         TabIndex        =   18
         Top             =   5280
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   926
         _Version        =   131073
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdOK 
         Height          =   525
         Left            =   4920
         TabIndex        =   17
         Top             =   5280
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         ButtonStyle     =   3
      End
   End
End
Attribute VB_Name = "frmAddEditCarkillSum"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private m_HasActivate As Boolean
Private m_HasModify As Boolean
Private m_Rs As ADODB.Recordset
Private ItemXlsSum  As CXlsCarkillSum

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

Private Sub chkHorizontal_Click(Value As Integer)
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

   If Flag Then
      Call EnableForm(Me, False)
      
      ItemXlsSum.XLS_SUM_ID = ID
'      ItemXlsSum.QueryFlag = 1
'      ItemXlsSum.FROM_CMPL_DATE = -1
'      ItemXlsSum.TO_CMPL_DATE = -1
      If Not glbDaily.QueryCarkillSum(ItemXlsSum, m_Rs, ItemCount, IsOK, glbErrorLog) Then
         glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
         Call EnableForm(Me, True)
         Exit Sub
      End If
   End If
   
   If ItemCount > 0 Then
      Call ItemXlsSum.PopulateFromRS(1, m_Rs)
         txtSumRow.Text = ItemXlsSum.SUM_ROW
         txtOperator1.Text = ItemXlsSum.OPERATOR_1
         txtpColumn1.Text = ItemXlsSum.P_COLUMN_1
         txtpRow1.Text = ItemXlsSum.P_ROW_1
         txtOperator2.Text = ItemXlsSum.OPERATOR_2
         txtpColumn2.Text = ItemXlsSum.P_COLUMN_2
         txtpRow2.Text = ItemXlsSum.P_ROW_2
         txtOperator3.Text = ItemXlsSum.OPERATOR_3
         txtpColumn3.Text = ItemXlsSum.P_COLUMN_3
         txtpRow3.Text = ItemXlsSum.P_ROW_3
         txtOperator4.Text = ItemXlsSum.OPERATOR_4
         txtpColumn4.Text = ItemXlsSum.P_COLUMN_4
         txtpRow4.Text = ItemXlsSum.P_ROW_4
         txtOperator5.Text = ItemXlsSum.OPERATOR_5
         txtpColumn5.Text = ItemXlsSum.P_COLUMN_5
         txtpRow5.Text = ItemXlsSum.P_ROW_5
         chkHorizontal.Value = FlagToCheck(ItemXlsSum.HORIZONTAL_FLAG)
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


   If Not VerifyTextControl(lblSumRow, txtSumRow, False) Then
      Exit Function
   End If

  If Not VerifyTextControl(lblOperator1, txtOperator1, False) Then         'pColumn
      Exit Function
   End If
   
   If Not VerifyTextControl(lblpColumn1, txtpColumn1, False) Then         'pColumn
      Exit Function
   End If
   
     If Not VerifyTextControl(lblpRow1, txtpRow1, False) Then         'pColumn
      Exit Function
   End If

   If Not m_HasModify Then
      SaveData = True
      Exit Function
   End If
   
   ItemXlsSum.AddEditMode = ShowMode
   ItemXlsSum.SUM_ROW = txtSumRow.Text
   ItemXlsSum.OPERATOR_1 = txtOperator1.Text
   ItemXlsSum.P_COLUMN_1 = txtpColumn1.Text
   ItemXlsSum.P_ROW_1 = txtpRow1.Text
   ItemXlsSum.HORIZONTAL_FLAG = Check2Flag(chkHorizontal.Value)
   
   If txtOperator2.Text <> "" Then
      ItemXlsSum.OPERATOR_2 = txtOperator2.Text
      ItemXlsSum.P_COLUMN_2 = txtpColumn2.Text
      ItemXlsSum.P_ROW_2 = txtpRow2.Text
   End If

   If txtOperator3.Text <> "" Then
      ItemXlsSum.OPERATOR_3 = txtOperator3.Text
      ItemXlsSum.P_COLUMN_3 = txtpColumn3.Text
      ItemXlsSum.P_ROW_3 = txtpRow3.Text
   End If
   
   If txtOperator4.Text <> "" Then
      ItemXlsSum.OPERATOR_4 = txtOperator4.Text
      ItemXlsSum.P_COLUMN_4 = txtpColumn4.Text
      ItemXlsSum.P_ROW_4 = txtpRow4.Text
   End If
   
   If txtOperator5.Text <> "" Then
      ItemXlsSum.OPERATOR_5 = txtOperator5.Text
      ItemXlsSum.P_COLUMN_5 = txtpColumn5.Text
      ItemXlsSum.P_ROW_5 = txtpRow5.Text
   End If
   
   Call EnableForm(Me, False)
   
   If Not glbDaily.AddEditCarkillSum(ItemXlsSum, IsOK, True, glbErrorLog) Then
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
   
   Call InitNormalLabel(lblSumRow, MapText("แถวที่"))
   Call InitNormalLabel(lblDetail, MapText("ตั้งค่าตัวบวก"))
   Call InitNormalLabel(lblOperator1, MapText("เครื่องหมาย"))
   Call InitNormalLabel(lblpColumn1, MapText("*คอลัมป์"))
   Call InitNormalLabel(lblpRow1, MapText("แถวที่"))
   Call InitNormalLabel(lblOperator2, MapText("เครื่องหมาย"))
   Call InitNormalLabel(lblpColumn2, MapText("*คอลัมป์"))
   Call InitNormalLabel(lblpRow2, MapText("แถวที่"))
   Call InitNormalLabel(lblOperator3, MapText("เครื่องหมาย"))
   Call InitNormalLabel(lblpColumn3, MapText("*คอลัมป์"))
   Call InitNormalLabel(lblpRow3, MapText("แถวที่"))
   Call InitNormalLabel(lblOperator4, MapText("เครื่องหมาย"))
   Call InitNormalLabel(lblpColumn4, MapText("*คอลัมป์"))
   Call InitNormalLabel(lblpRow4, MapText("แถวที่"))
   Call InitNormalLabel(lblOperator5, MapText("เครื่องหมาย"))
   Call InitNormalLabel(lblpColumn5, MapText("*คอลัมป์"))
   Call InitNormalLabel(lblpRow5, MapText("แถวที่"))
   Call InitNormalLabel(lblComment, MapText("*คอลัมป์ : (-1) = คอลัมป์ก่อนหน้านี้ , 0 = คอลัมป์ปัจจุบัน , (+1) = คอลัมป์ถัดไป "))
   Call InitCheckBox(chkHorizontal, "รวมตามแนวนอน")
      
   Call txtSumRow.SetTextLenType(TEXT_INTEGER, glbSetting.CODE_TYPE)
   Call txtOperator1.SetTextLenType(TEXT_STRING, glbSetting.CODE_TYPE)
   Call txtpColumn1.SetTextLenType(TEXT_STRING, glbSetting.DOUBLE_TYPE)
   Call txtpRow1.SetTextLenType(TEXT_INTEGER, glbSetting.CODE_TYPE)
   Call txtOperator2.SetTextLenType(TEXT_STRING, glbSetting.CODE_TYPE)
   Call txtpColumn2.SetTextLenType(TEXT_STRING, glbSetting.DOUBLE_TYPE)
   Call txtpRow2.SetTextLenType(TEXT_INTEGER, glbSetting.CODE_TYPE)
   Call txtOperator3.SetTextLenType(TEXT_STRING, glbSetting.CODE_TYPE)
   Call txtpColumn3.SetTextLenType(TEXT_STRING, glbSetting.DOUBLE_TYPE)
   Call txtpRow3.SetTextLenType(TEXT_INTEGER, glbSetting.CODE_TYPE)
   Call txtOperator4.SetTextLenType(TEXT_STRING, glbSetting.CODE_TYPE)
   Call txtpColumn4.SetTextLenType(TEXT_STRING, glbSetting.DOUBLE_TYPE)
   Call txtpRow4.SetTextLenType(TEXT_INTEGER, glbSetting.CODE_TYPE)
   Call txtOperator5.SetTextLenType(TEXT_STRING, glbSetting.CODE_TYPE)
   Call txtpColumn5.SetTextLenType(TEXT_STRING, glbSetting.DOUBLE_TYPE)
   Call txtpRow5.SetTextLenType(TEXT_INTEGER, glbSetting.CODE_TYPE)
      
   Me.Picture = LoadPicture(glbParameterObj.MainPicture)
   
   pnlHeader.Font.Name = GLB_FONT
   pnlHeader.Font.Bold = True
   
   pnlHeader.Font.Size = 19
   
   cmdExit.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdOK.Picture = LoadPicture(glbParameterObj.NormalButton1)
 '  cmdNext.Picture = LoadPicture(glbParameterObj.NormalButton1)
   
   Call InitMainButton(cmdExit, MapText("ยกเลิก (ESC)"))
   Call InitMainButton(cmdOK, MapText("ตกลง (F2)"))
  ' Call InitMainButton(cmdNext, MapText("ถัดไป (F7)"))
End Sub
Private Sub cmdExit_Click()
'   If Not ConfirmExit(m_HasModify) Then
'      Exit Sub
'   End If
   
   OKClick = False
   Unload Me
End Sub
Private Sub Form_Load()
   Call EnableForm(Me, False)
   m_HasActivate = False
   Set m_Rs = New ADODB.Recordset
      
   Set ItemXlsSum = New CXlsCarkillSum
'   Set m_Stcrd = New Collection
   
   m_HasActivate = False
   Call InitFormLayout
   Call EnableForm(Me, True)
End Sub
Private Sub Form_Unload(Cancel As Integer)
   Set ItemXlsSum = Nothing
'   Set m_Stcrd = Nothing
End Sub
Private Sub txtSumRow_Change()
   m_HasModify = True
End Sub
Private Sub txtOperator1_Change()
   m_HasModify = True
End Sub
Private Sub txtpColumn1_Change()
   m_HasModify = True
End Sub
Private Sub txtpRow1_Change()
   m_HasModify = True
End Sub
Private Sub txtOperator2_Change()
   m_HasModify = True
End Sub
Private Sub txtpColumn2_Change()
   m_HasModify = True
End Sub
Private Sub txtpRow2_Change()
   m_HasModify = True
End Sub
Private Sub txtOperator3_Change()
   m_HasModify = True
End Sub
Private Sub txtpColumn3_Change()
   m_HasModify = True
End Sub
Private Sub txtpRow3_Change()
   m_HasModify = True
End Sub
Private Sub txtOperator4_Change()
   m_HasModify = True
End Sub
Private Sub txtpColumn4_Change()
   m_HasModify = True
End Sub
Private Sub txtpRow4_Change()
   m_HasModify = True
End Sub
Private Sub txtOperator5_Change()
   m_HasModify = True
End Sub
Private Sub txtpColumn5_Change()
   m_HasModify = True
End Sub
Private Sub txtpRow5_Change()
   m_HasModify = True
End Sub
