VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#2.0#0"; "THREED20.OCX"
Begin VB.Form frmAddEditBankCredit 
   BackColor       =   &H80000000&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   8235
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7170
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8235
   ScaleWidth      =   7170
   StartUpPosition =   1  'CenterOwner
   Begin Threed.SSFrame SSFrame1 
      Height          =   8805
      Left            =   -120
      TabIndex        =   17
      Top             =   0
      Width           =   7785
      _ExtentX        =   13732
      _ExtentY        =   15531
      _Version        =   131073
      PictureBackgroundStyle=   2
      Begin VB.ComboBox cboCustomerName 
         Height          =   315
         Left            =   1920
         Style           =   2  'Dropdown List
         TabIndex        =   8
         Top             =   6600
         Width           =   3375
      End
      Begin VB.ComboBox cboBankName 
         Height          =   315
         Left            =   1920
         Style           =   2  'Dropdown List
         TabIndex        =   0
         Top             =   1080
         Width           =   3375
      End
      Begin prjLedgerReport.uctlTextBox txtBankGetAmount 
         Height          =   495
         Left            =   1920
         TabIndex        =   7
         Top             =   5880
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   873
      End
      Begin prjLedgerReport.uctlTextBox txtBankFeeAmount 
         Height          =   495
         Left            =   1920
         TabIndex        =   4
         Top             =   3720
         Width           =   2895
         _ExtentX        =   5106
         _ExtentY        =   873
      End
      Begin VB.ComboBox cboBankFeeType 
         Height          =   315
         Left            =   1920
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   3120
         Width           =   3375
      End
      Begin prjLedgerReport.uctlTextBox txtBankAmountBrougt 
         Height          =   495
         Left            =   1920
         TabIndex        =   6
         Top             =   5160
         Width           =   2895
         _ExtentX        =   5106
         _ExtentY        =   873
      End
      Begin prjLedgerReport.uctlDate uctlPutDate 
         Height          =   375
         Left            =   1920
         TabIndex        =   5
         Top             =   4440
         Width           =   3855
         _ExtentX        =   6800
         _ExtentY        =   661
      End
      Begin prjLedgerReport.uctlTextBox txtBankInterest 
         Height          =   495
         Left            =   1920
         TabIndex        =   2
         Top             =   2400
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   873
      End
      Begin prjLedgerReport.uctlTextBox txtBankAmount 
         Height          =   495
         Left            =   1920
         TabIndex        =   1
         Top             =   1680
         Width           =   2895
         _ExtentX        =   5106
         _ExtentY        =   873
      End
      Begin Threed.SSPanel pnlHeader 
         Height          =   705
         Left            =   0
         TabIndex        =   18
         Top             =   0
         Width           =   7755
         _ExtentX        =   13679
         _ExtentY        =   1244
         _Version        =   131073
         PictureBackgroundStyle=   2
      End
      Begin VB.Label Label5 
         Caption         =   "Label5"
         Height          =   255
         Left            =   4920
         TabIndex        =   25
         Top             =   3840
         Width           =   735
      End
      Begin VB.Label lblCustomerName 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   375
         Left            =   360
         TabIndex        =   24
         Top             =   6600
         Width           =   1335
      End
      Begin VB.Label Label4 
         Caption         =   "Label4"
         Height          =   375
         Left            =   3240
         TabIndex        =   23
         Top             =   6000
         Width           =   1095
      End
      Begin VB.Label lblBankGetAmount 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   495
         Left            =   120
         TabIndex        =   22
         Top             =   6000
         Width           =   1695
      End
      Begin VB.Label Label3 
         Caption         =   "Label3"
         Height          =   495
         Left            =   3240
         TabIndex        =   21
         Top             =   2520
         Width           =   855
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Caption         =   "Label2"
         Height          =   495
         Left            =   4680
         TabIndex        =   20
         Top             =   5280
         Width           =   1095
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Caption         =   "Label1"
         Height          =   495
         Left            =   4680
         TabIndex        =   19
         Top             =   1800
         Width           =   1095
      End
      Begin VB.Label lblBankAmountBrougt 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   495
         Left            =   0
         TabIndex        =   11
         Top             =   5280
         Width           =   1695
      End
      Begin VB.Label lblPutDate 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   495
         Left            =   120
         TabIndex        =   12
         Top             =   4560
         Width           =   1575
      End
      Begin VB.Label lblBankFeeAmount 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   615
         Left            =   120
         TabIndex        =   13
         Top             =   3240
         Width           =   1575
      End
      Begin VB.Label lblBankInterest 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   615
         Left            =   120
         TabIndex        =   14
         Top             =   2520
         Width           =   1575
      End
      Begin VB.Label lblBankName 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   255
         Left            =   -240
         TabIndex        =   16
         Top             =   1080
         Width           =   1935
      End
      Begin VB.Label lblBankAmount 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   435
         Left            =   -240
         TabIndex        =   15
         Top             =   1800
         Width           =   1935
      End
      Begin Threed.SSCommand cmdExit 
         Cancel          =   -1  'True
         Height          =   525
         Left            =   4080
         TabIndex        =   10
         Top             =   7200
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   926
         _Version        =   131073
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdOK 
         Height          =   525
         Left            =   1920
         TabIndex        =   9
         Top             =   7200
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         ButtonStyle     =   3
      End
   End
End
Attribute VB_Name = "frmAddEditBankCredit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private m_HasActivate As Boolean
Private m_HasModify As Boolean
Private m_Rs As ADODB.Recordset
Private m_BankCredit As CBankCredit

Public ID As Long
Public OKClick As Boolean
Public ShowMode As SHOW_MODE_TYPE
Public HeaderText As String
Private Sub cboDataType_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
      CreateObject("WScript.Shell").SendKeys "{TAB}"
   End If
End Sub
Private Sub cboBankFeeType_Click()
   If cboBankFeeType.ListIndex = 1 Then
      Label5.Caption = " บาท"
   Else
      Label5.Caption = " %"
   End If
   m_HasModify = True
End Sub
Private Sub cboBankName_Click()
 m_HasModify = True
End Sub
Private Sub cboCustomerName_Click()
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
      
      m_BankCredit.BANK_NUMBER = ID
      m_BankCredit.QueryFlag = 1
      
      If Not glbDaily.QueryBankCredit(m_BankCredit, m_Rs, ItemCount, IsOK, glbErrorLog) Then
         glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
         Call EnableForm(Me, True)
         Exit Sub
      End If
   End If
   
   If ItemCount > 0 Then
      Call m_BankCredit.PopulateFromRS(1, m_Rs)
      cboBankName.ListIndex = IDToListIndex(cboBankName, m_BankCredit.BANK_ID)
      txtBankAmount.Text = m_BankCredit.BANK_AMOUNT
      txtBankInterest.Text = m_BankCredit.BANK_INTEREST
      cboBankFeeType.ListIndex = IDToListIndex(cboBankFeeType, m_BankCredit.BANK_FEE_TYPE)
      txtBankFeeAmount.Text = m_BankCredit.BANK_FEE_AMOUNT
      uctlPutDate.ShowDate = m_BankCredit.BANK_DATE_BROUGHT
      txtBankAmountBrougt.Text = m_BankCredit.BANK_AMOUNT_BROUGHT
      txtBankGetAmount.Text = m_BankCredit.BANK_GET_AMOUNT
      cboCustomerName.ListIndex = IDToListIndex(cboCustomerName, m_BankCredit.CUSTOMER_ID)
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

   If Not VerifyCombo(lblBankName, cboBankName, False) Then
      Exit Function
   End If
   If Not VerifyTextControl(lblBankAmount, txtBankAmount, False) Then
      Exit Function
   End If
   If Not VerifyTextControl(lblBankFeeAmount, txtBankFeeAmount, False) Then
      Exit Function
   End If
   If Not VerifyDate(lblPutDate, uctlPutDate, False) Then
      Exit Function
   End If
   If Not VerifyTextControl(lblBankAmountBrougt, txtBankAmountBrougt, False) Then
      Exit Function
   End If
    If Not VerifyTextControl(lblBankGetAmount, txtBankGetAmount, False) Then
      Exit Function
   End If
   If Not VerifyCombo(lblCustomerName, cboCustomerName, False) Then
      Exit Function
   End If
   
   If Not m_HasModify Then
      SaveData = True
      Exit Function
   End If
   
   m_BankCredit.AddEditMode = ShowMode
   m_BankCredit.BANK_NUMBER = ID
   m_BankCredit.BANK_ID = cboBankName.ItemData(Minus2Zero(cboBankName.ListIndex))
   m_BankCredit.BANK_AMOUNT = txtBankAmount.Text
   m_BankCredit.BANK_INTEREST = txtBankInterest.Text
   m_BankCredit.BANK_FEE_TYPE = cboBankFeeType.ItemData(Minus2Zero(cboBankFeeType.ListIndex))
   m_BankCredit.BANK_FEE_AMOUNT = txtBankFeeAmount.Text
   m_BankCredit.BANK_DATE_BROUGHT = uctlPutDate.ShowDate
   m_BankCredit.BANK_AMOUNT_BROUGHT = txtBankAmountBrougt.Text
   m_BankCredit.BANK_GET_AMOUNT = txtBankGetAmount.Text
   m_BankCredit.CUSTOMER_ID = cboCustomerName.ItemData(Minus2Zero(cboCustomerName.ListIndex))
   
   Call EnableForm(Me, False)
   If Not glbDaily.AddEditBankCredit(m_BankCredit, IsOK, True, glbErrorLog) Then
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
      
      Call LoadBank(cboBankName)
      Call LoadBankCustomer(cboCustomerName)
      Call InitBankFeeType(cboBankFeeType)
      
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
   
   Call InitNormalLabel(lblBankName, MapText("ธนาคาร"), RGB(255, 0, 0))
   
   Call InitNormalLabel(lblBankAmount, MapText("วงเงิน"), RGB(255, 0, 0))
   Call txtBankAmount.SetKeySearch("BANK_AMOUNT")
   
   Call InitNormalLabel(lblBankInterest, MapText("ดอกเบี้ย"))
   Call txtBankInterest.SetKeySearch("BANK_INTEREST")
   
   Call InitNormalLabel(lblBankFeeAmount, MapText("ค่าธรรมเนียม"), RGB(255, 0, 0))
   Call txtBankFeeAmount.SetKeySearch("BANK_FEE_AMOUNT")
   
   Call InitNormalLabel(lblPutDate, MapText("วันที่"), RGB(255, 0, 0))
   
   Call InitNormalLabel(lblBankAmountBrougt, MapText("ยอดเงินยกมา"), RGB(255, 0, 0))
   Call txtBankAmountBrougt.SetKeySearch("BANK_AMOUNT_BROUGT")
   
   Call InitNormalLabel(lblBankGetAmount, MapText("จำนวนเงินรับมา"), RGB(255, 0, 0))
   Call txtBankGetAmount.SetKeySearch("BANK_GET_AMOUNT")
   
   Call InitNormalLabel(lblCustomerName, MapText("ลูกหนี้"), RGB(255, 0, 0))

   Call InitNormalLabel(Label1, MapText("บาท"), RGB(255, 0, 0))
   Call InitNormalLabel(Label2, MapText("บาท"), RGB(255, 0, 0))
   Call InitNormalLabel(Label3, MapText("%"))
   Call InitNormalLabel(Label4, MapText("%"), RGB(255, 0, 0))
   Call InitNormalLabel(Label5, MapText(" "), RGB(255, 0, 0))

   Call txtBankAmount.SetTextLenType(TEXT_FLOAT_MONEY, glbSetting.MONEY_TYPE)
   Call txtBankInterest.SetTextLenType(TEXT_FLOAT_MONEY, glbSetting.MONEY_TYPE)
   Call txtBankFeeAmount.SetTextLenType(TEXT_FLOAT_MONEY, glbSetting.MONEY_TYPE)
   Call txtBankAmountBrougt.SetTextLenType(TEXT_FLOAT_MONEY, glbSetting.MONEY_TYPE)
   Call txtBankGetAmount.SetTextLenType(TEXT_FLOAT_MONEY, glbSetting.MONEY_TYPE)
   
   Call InitCombo(cboBankFeeType)
   Call InitCombo(cboCustomerName)
   Call InitCombo(cboBankName)

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
   
   Set m_BankCredit = New CBankCredit
   Set m_Rs = New ADODB.Recordset
   
   m_HasActivate = False
   Call InitFormLayout
   Call EnableForm(Me, True)
End Sub
Private Sub Form_Unload(Cancel As Integer)
   Set m_BankCredit = Nothing
End Sub
Private Sub txtBankName_Change()
   m_HasModify = True
End Sub
Private Sub txtBankAmount_Change()
   m_HasModify = True
End Sub
Private Sub txtBankInterest_Change()
   m_HasModify = True
End Sub
Private Sub txtBankFeeAmount_Change()
   m_HasModify = True
End Sub
Private Sub txtBankAmountBrougt_Change()
   m_HasModify = True
End Sub
Private Sub uctlPutDate_HasChange()
   m_HasModify = True
End Sub
Private Sub txtBankGetAmount_Change()
   m_HasModify = True
End Sub

