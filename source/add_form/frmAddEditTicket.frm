VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#2.0#0"; "THREED20.OCX"
Begin VB.Form frmAddEditTicket 
   BackColor       =   &H80000000&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   7590
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8475
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7590
   ScaleWidth      =   8475
   StartUpPosition =   1  'CenterOwner
   Begin Threed.SSFrame SSFrame1 
      Height          =   8325
      Left            =   0
      TabIndex        =   16
      Top             =   0
      Width           =   8625
      _ExtentX        =   15214
      _ExtentY        =   14684
      _Version        =   131073
      PictureBackgroundStyle=   2
      Begin VB.ComboBox cboBankID 
         Height          =   315
         Left            =   2160
         Style           =   2  'Dropdown List
         TabIndex        =   8
         Top             =   5400
         Width           =   3375
      End
      Begin prjLedgerReport.uctlDate uctlTicketDateNew 
         Height          =   375
         Left            =   2760
         TabIndex        =   25
         Top             =   3960
         Width           =   3855
         _ExtentX        =   6800
         _ExtentY        =   661
      End
      Begin prjLedgerReport.uctlTextBox txtTicketInterest 
         Height          =   495
         Left            =   2160
         TabIndex        =   22
         Top             =   5760
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   873
      End
      Begin prjLedgerReport.uctlTextBox txtTicketAmount 
         Height          =   495
         Left            =   2160
         TabIndex        =   20
         Top             =   2400
         Width           =   1695
         _ExtentX        =   5106
         _ExtentY        =   873
      End
      Begin prjLedgerReport.uctlTextLookup uctlTicketInvoiceLookup 
         Height          =   375
         Left            =   2160
         TabIndex        =   2
         Top             =   1920
         Width           =   6255
         _ExtentX        =   11033
         _ExtentY        =   661
      End
      Begin VB.ComboBox cboCustomerName 
         Height          =   315
         Left            =   2160
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Top             =   4920
         Width           =   3375
      End
      Begin prjLedgerReport.uctlDate uctlTicketDateGetCheck 
         Height          =   375
         Left            =   2160
         TabIndex        =   5
         Top             =   4440
         Width           =   3855
         _ExtentX        =   6800
         _ExtentY        =   661
      End
      Begin prjLedgerReport.uctlDate uctlTicketDateDue 
         Height          =   375
         Left            =   2160
         TabIndex        =   3
         Top             =   3000
         Width           =   3855
         _ExtentX        =   6800
         _ExtentY        =   661
      End
      Begin prjLedgerReport.uctlDate uctlTicketDate 
         Height          =   375
         Left            =   2160
         TabIndex        =   0
         Top             =   840
         Width           =   3855
         _ExtentX        =   6800
         _ExtentY        =   661
      End
      Begin prjLedgerReport.uctlDate uctlTicketDateCheck 
         Height          =   375
         Left            =   2160
         TabIndex        =   4
         Top             =   3480
         Width           =   3855
         _ExtentX        =   6800
         _ExtentY        =   661
      End
      Begin prjLedgerReport.uctlTextBox txtTicketNumber 
         Height          =   495
         Left            =   2160
         TabIndex        =   1
         Top             =   1320
         Width           =   2895
         _ExtentX        =   5106
         _ExtentY        =   873
      End
      Begin Threed.SSPanel pnlHeader 
         Height          =   705
         Left            =   0
         TabIndex        =   17
         Top             =   0
         Width           =   8535
         _ExtentX        =   15055
         _ExtentY        =   1244
         _Version        =   131073
         PictureBackgroundStyle=   2
      End
      Begin VB.Label lblBankID 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   375
         Left            =   120
         TabIndex        =   28
         Top             =   5400
         Width           =   1815
      End
      Begin Threed.SSCheck ChkCheckReceived 
         Height          =   375
         Left            =   4920
         TabIndex        =   27
         Top             =   2520
         Width           =   3375
         _ExtentX        =   5953
         _ExtentY        =   661
         _Version        =   131073
         Caption         =   "SSCheck1"
      End
      Begin VB.Label lblTicketDateNew 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   255
         Left            =   120
         TabIndex        =   26
         Top             =   3960
         Width           =   2535
      End
      Begin VB.Label Label2 
         Caption         =   "Label2"
         Height          =   255
         Left            =   3240
         TabIndex        =   24
         Top             =   5880
         Width           =   615
      End
      Begin VB.Label Label1 
         Caption         =   "Label1"
         Height          =   255
         Left            =   3960
         TabIndex        =   23
         Top             =   2520
         Width           =   615
      End
      Begin VB.Label lblTicketInterest 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   375
         Left            =   120
         TabIndex        =   21
         Top             =   5880
         Width           =   1815
      End
      Begin VB.Label lblTicketAmount 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   255
         Left            =   120
         TabIndex        =   19
         Top             =   2520
         Width           =   1815
      End
      Begin VB.Label lblCustomerName 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   375
         Left            =   240
         TabIndex        =   18
         Top             =   4920
         Width           =   1695
      End
      Begin VB.Label lblTicketDateGetCheck 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   375
         Left            =   240
         TabIndex        =   6
         Top             =   4440
         Width           =   1695
      End
      Begin VB.Label lblTicketDateCheck 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   375
         Left            =   360
         TabIndex        =   11
         Top             =   3480
         Width           =   1575
      End
      Begin VB.Label lblTicketDateDue 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   375
         Left            =   120
         TabIndex        =   12
         Top             =   3000
         Width           =   1815
      End
      Begin VB.Label lblTicketInvoice 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   375
         Left            =   360
         TabIndex        =   13
         Top             =   2040
         Width           =   1575
      End
      Begin VB.Label lblTicketDate 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   375
         Left            =   240
         TabIndex        =   15
         Top             =   960
         Width           =   1695
      End
      Begin VB.Label lblTicketNumber 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   315
         Left            =   240
         TabIndex        =   14
         Top             =   1440
         Width           =   1695
      End
      Begin Threed.SSCommand cmdExit 
         Cancel          =   -1  'True
         Height          =   525
         Left            =   4320
         TabIndex        =   10
         Top             =   6480
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   926
         _Version        =   131073
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdOK 
         Height          =   525
         Left            =   2160
         TabIndex        =   9
         Top             =   6480
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         ButtonStyle     =   3
      End
   End
End
Attribute VB_Name = "frmAddEditTicket"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private m_HasActivate As Boolean
Private m_HasModify As Boolean
Private m_Rs As ADODB.Recordset
Private m_Ticket As CTicket

Public ID As Long
Public OKClick As Boolean
Public ShowMode As SHOW_MODE_TYPE
Public HeaderText As String
Private m_Parts As Collection
Private m_MaxAmount As Collection
Private Sub ChkCheckReceived_Click(Value As Integer)
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
      
      m_Ticket.TICKET_ID = ID
      m_Ticket.QueryFlag = 1
      
      If Not glbDaily.QueryTicket(m_Ticket, m_Rs, ItemCount, IsOK, glbErrorLog) Then
         glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
         Call EnableForm(Me, True)
         Exit Sub
      End If
   End If
   
   If ItemCount > 0 Then
      Call m_Ticket.PopulateFromRS(1, m_Rs)
      uctlTicketDate.ShowDate = m_Ticket.TICKET_DATE
      txtTicketNumber.Text = m_Ticket.TICKET_NUMBER
      uctlTicketInvoiceLookup.MyTextBox.Text = m_Ticket.TICKET_INVOICE
      uctlTicketDateDue.ShowDate = m_Ticket.TICKET_DATE_DUE
      uctlTicketDateCheck.ShowDate = m_Ticket.TICKET_DATE_CHECK
      uctlTicketDateGetCheck.ShowDate = m_Ticket.TICKET_DATE_GET_CHECK
      cboCustomerName.ListIndex = IDToListIndex(cboCustomerName, m_Ticket.CUSTOMER_ID)
      txtTicketAmount.Text = m_Ticket.TICKET_AMOUNT
      txtTicketInterest.Text = m_Ticket.TICKET_INTEREST
      uctlTicketDateNew.ShowDate = m_Ticket.TICKET_DATE_NEW
      ChkCheckReceived.Value = m_Ticket.CHECK_RECEIVED
      cboBankID.ListIndex = IDToListIndex(cboBankID, m_Ticket.BANK_ID)
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
Dim m_Lookup As CARTrn
  Set m_Lookup = New CARTrn

   If Not VerifyDate(lblTicketDate, uctlTicketDate, False) Then
      Exit Function
   End If
   If Not VerifyTextControl(lblTicketNumber, txtTicketNumber, False) Then
      Exit Function
   End If
   If Not (uctlTicketInvoiceLookup.MyCombo.ItemData(Minus2Zero(uctlTicketInvoiceLookup.MyCombo.ListIndex)) > 0) Then
      Exit Function
   End If
   If Not VerifyDate(lblTicketDateDue, uctlTicketDateDue, False) Then
      Exit Function
   End If
   If Not VerifyDate(lblTicketDateCheck, uctlTicketDateCheck, False) Then
      Exit Function
   End If
   If Not VerifyDate(lblTicketDateGetCheck, uctlTicketDateGetCheck, False) Then
      Exit Function
   End If
   If Not VerifyCombo(lblCustomerName, cboCustomerName, False) Then
      Exit Function
   End If
   If Not VerifyCombo(lblBankID, cboBankID, False) Then
      Exit Function
   End If
   
   If Not m_HasModify Then
      SaveData = True
      Exit Function
   End If
   
   m_Ticket.AddEditMode = ShowMode
   m_Ticket.TICKET_ID = ID
   m_Ticket.TICKET_DATE = uctlTicketDate.ShowDate
   m_Ticket.TICKET_NUMBER = txtTicketNumber.Text
   m_Ticket.TICKET_INVOICE = uctlTicketInvoiceLookup.MyTextBox.Text
   m_Ticket.TICKET_DATE_DUE = uctlTicketDateDue.ShowDate
   m_Ticket.TICKET_DATE_CHECK = uctlTicketDateCheck.ShowDate
   m_Ticket.TICKET_DATE_GET_CHECK = uctlTicketDateGetCheck.ShowDate
    m_Ticket.CUSTOMER_ID = cboCustomerName.ItemData(Minus2Zero(cboCustomerName.ListIndex))
   m_Ticket.TICKET_AMOUNT = FormatNumber(txtTicketAmount.Text)
   m_Ticket.TICKET_INTEREST = txtTicketInterest.Text
   m_Ticket.TICKET_DATE_NEW = uctlTicketDateNew.ShowDate
   m_Ticket.CHECK_RECEIVED = ChkCheckReceived.Value
    m_Ticket.BANK_ID = cboBankID.ItemData(Minus2Zero(cboBankID.ListIndex))
    m_Ticket.MASTER_AREA = 1
   
   Call EnableForm(Me, False)
   If Not glbDaily.AddEditTicket(m_Ticket, IsOK, True, glbErrorLog) Then
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
      
      Call LoadTicketLookup(uctlTicketInvoiceLookup.MyCombo, m_Parts)
      Set uctlTicketInvoiceLookup.MyCollection = m_Parts
      
      Call LoadBank(cboBankID)
      Call LoadBankCustomer(cboCustomerName)
      Call LoadTicketMaxAmount(Nothing, m_MaxAmount)
      
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
   
   Call InitNormalLabel(lblTicketDate, MapText("วันที่"), RGB(255, 0, 0))
   
   Call InitNormalLabel(lblTicketNumber, MapText("ตั๋วเลขที่"), RGB(255, 0, 0))
   Call txtTicketNumber.SetKeySearch("TICKET_NUMBER")
   
   Call InitNormalLabel(lblTicketInvoice, MapText("INVOICE"), RGB(255, 0, 0))
   uctlTicketInvoiceLookup.MyCombo.Enabled = False
   
   Call InitNormalLabel(lblTicketDateDue, MapText("วันที่ตั๋วครบกำหนด"), RGB(255, 0, 0))
   Call InitNormalLabel(lblTicketDateCheck, MapText("วันที่หน้า เช็ค"), RGB(255, 0, 0))
   Call InitNormalLabel(lblTicketDateGetCheck, MapText("วันที่เก็บ เช็ค"), RGB(255, 0, 0))
   Call InitNormalLabel(lblTicketDateNew, MapText("วันที่ตั๋วครบกำหนด (พิเศษ)"))
   uctlTicketDateNew.TabStop = False
   Call InitCheckBox(ChkCheckReceived, MapText("ระบุจำนวนเงินหน้าเช็ค"))
   ChkCheckReceived.TabStop = False
   Call InitNormalLabel(lblCustomerName, MapText("ลูกหนี้"), RGB(255, 0, 0))
   Call InitNormalLabel(lblBankID, MapText("ธนาคาร"), RGB(255, 0, 0))
   Call InitNormalLabel(lblTicketAmount, MapText("จำนวนเงินหน้าเช็ค"))
   txtTicketAmount.TabStop = False
   Call InitNormalLabel(lblTicketInterest, MapText("ดอกเบี้ยใหม่"))
   txtTicketInterest.TabStop = False
   Call txtTicketNumber.SetTextLenType(TEXT_STRING, glbSetting.CODE_TYPE)
   Call InitCombo(cboCustomerName)
   Call InitCombo(cboBankID)
   
   Call InitNormalLabel(Label1, MapText("บาท"))
   Call InitNormalLabel(Label2, MapText("%"))

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
   
   Set m_Ticket = New CTicket
   Set m_Rs = New ADODB.Recordset
   Set m_Parts = New Collection
   Set m_MaxAmount = New Collection
   
   m_HasActivate = False
   Call InitFormLayout
   Call EnableForm(Me, True)
End Sub
Private Sub Form_Unload(Cancel As Integer)
   Set m_Ticket = Nothing
   Set m_Parts = Nothing
   Set m_MaxAmount = Nothing
End Sub
Private Sub txtTicketAmount_Change()
   Dim D As CARTrn
   Dim Check As CTicket
   Dim TempMaxAmount As Double
   Dim TempAmount As Double
   Set D = New CARTrn
   Set Check = New CTicket
   
   TempAmount = 0
   TempMaxAmount = 0
   Set D = GetObject("CARTrn", m_Parts, uctlTicketInvoiceLookup.MyCombo.ItemData(Minus2Zero(uctlTicketInvoiceLookup.MyCombo.ListIndex)), False)
   If Not D Is Nothing Then
      TempMaxAmount = CDbl(D.AMOUNT)
   End If
   If txtTicketAmount.Text = "" Then
      TempAmount = 0
   Else
      TempAmount = CDbl(txtTicketAmount.Text)
   End If
   
  Set Check = GetObject("CTicket", m_MaxAmount, Trim(D.DOCNUM), False)
   If Not Check Is Nothing Then
      TempAmount = TempAmount + Check.TICKET_AMOUNT
   End If
   
   If TempAmount > TempMaxAmount Then
      glbErrorLog.LocalErrorMsg = "จำนวนเงินที่ใส่ลงไป มีค่าเกินจากจำนวนเงินหน้าเช็ค"
      glbErrorLog.ShowUserError
      m_HasModify = False
      Exit Sub
   End If
m_HasModify = True
End Sub
Private Sub txtTicketInterest_Change()
m_HasModify = True
End Sub
Private Sub txtTicketNumber_Change()
   m_HasModify = True
End Sub
Private Sub uctlTicketDate_HasChange()
   m_HasModify = True
End Sub
Private Sub uctlTicketDateDue_HasChange()
   m_HasModify = True
End Sub
Private Sub uctlTicketDateCheck_HasChange()
   m_HasModify = True
End Sub
Private Sub uctlTicketDateGetCheck_HasChange()
   m_HasModify = True
End Sub
Private Sub cboCustomerName_Click()
   m_HasModify = True
End Sub
Private Sub cboBankID_Click()
   m_HasModify = True
End Sub
Private Sub uctlTicketDateNew_HasChange()
   m_HasModify = True
End Sub
Private Sub uctlTicketInvoiceLookup_Change()
   Dim D As CARTrn
   Set D = New CARTrn

   Set D = GetObject("CARTrn", m_Parts, uctlTicketInvoiceLookup.MyCombo.ItemData(Minus2Zero(uctlTicketInvoiceLookup.MyCombo.ListIndex)), False)
   If D Is Nothing Then
      uctlTicketDateCheck.ShowDate = -1
   Else
      uctlTicketDateCheck.ShowDate = D.DUEDAT
   End If
   m_HasModify = True
End Sub

