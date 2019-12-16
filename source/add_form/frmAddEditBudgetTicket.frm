VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#2.0#0"; "THREED20.OCX"
Begin VB.Form frmAddEditBudgetTicket 
   BackColor       =   &H80000000&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   4950
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5925
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4950
   ScaleWidth      =   5925
   StartUpPosition =   1  'CenterOwner
   Begin Threed.SSFrame SSFrame1 
      Height          =   5115
      Left            =   0
      TabIndex        =   9
      Top             =   0
      Width           =   6105
      _ExtentX        =   10769
      _ExtentY        =   9022
      _Version        =   131073
      PictureBackgroundStyle=   2
      Begin prjLedgerReport.uctlTextBox txtTicketDateCheck 
         Height          =   495
         Left            =   2160
         TabIndex        =   2
         Top             =   1920
         Width           =   1000
         _ExtentX        =   1773
         _ExtentY        =   873
      End
      Begin VB.ComboBox cboBankID 
         Height          =   315
         Left            =   2160
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   3840
         Width           =   3375
      End
      Begin prjLedgerReport.uctlTextBox txtTicketAmount 
         Height          =   495
         Left            =   2160
         TabIndex        =   1
         Top             =   1320
         Width           =   1695
         _ExtentX        =   5106
         _ExtentY        =   873
      End
      Begin VB.ComboBox cboCustomerName 
         Height          =   315
         Left            =   2160
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   3240
         Width           =   3375
      End
      Begin prjLedgerReport.uctlDate uctlTicketDate 
         Height          =   375
         Left            =   2160
         TabIndex        =   0
         Top             =   840
         Width           =   3615
         _ExtentX        =   6376
         _ExtentY        =   661
      End
      Begin prjLedgerReport.uctlDate uctlTicketDateCheck 
         Height          =   375
         Left            =   2160
         TabIndex        =   15
         Top             =   2520
         Width           =   3615
         _ExtentX        =   6376
         _ExtentY        =   661
      End
      Begin Threed.SSPanel pnlHeader 
         Height          =   705
         Left            =   0
         TabIndex        =   10
         Top             =   0
         Width           =   6015
         _ExtentX        =   10610
         _ExtentY        =   1244
         _Version        =   131073
         PictureBackgroundStyle=   2
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Caption         =   "Label2"
         Height          =   375
         Left            =   3240
         TabIndex        =   16
         Top             =   2040
         Width           =   375
      End
      Begin VB.Label lblBankID 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   375
         Left            =   120
         TabIndex        =   14
         Top             =   3840
         Width           =   1815
      End
      Begin VB.Label Label1 
         Caption         =   "Label1"
         Height          =   255
         Left            =   3960
         TabIndex        =   13
         Top             =   1440
         Width           =   615
      End
      Begin VB.Label lblTicketAmount 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   255
         Left            =   120
         TabIndex        =   12
         Top             =   1440
         Width           =   1815
      End
      Begin VB.Label lblCustomerName 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   375
         Left            =   240
         TabIndex        =   11
         Top             =   3240
         Width           =   1695
      End
      Begin VB.Label lblTicketDateCheck 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   375
         Left            =   360
         TabIndex        =   6
         Top             =   2040
         Width           =   1575
      End
      Begin VB.Label lblTicketDate 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   375
         Left            =   240
         TabIndex        =   8
         Top             =   960
         Width           =   1695
      End
      Begin Threed.SSCommand cmdExit 
         Cancel          =   -1  'True
         Height          =   525
         Left            =   3360
         TabIndex        =   7
         Top             =   4320
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   926
         _Version        =   131073
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdOK 
         Height          =   525
         Left            =   1320
         TabIndex        =   5
         Top             =   4320
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         ButtonStyle     =   3
      End
   End
End
Attribute VB_Name = "frmAddEditBudgetTicket"
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
      txtTicketAmount.Text = m_Ticket.TICKET_AMOUNT
      txtTicketDateCheck.Text = DateDiff("D", m_Ticket.TICKET_DATE, m_Ticket.TICKET_DATE_CHECK)
      uctlTicketDateCheck.ShowDate = m_Ticket.TICKET_DATE_CHECK
      cboCustomerName.ListIndex = IDToListIndex(cboCustomerName, m_Ticket.CUSTOMER_ID)
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
   If Not VerifyDate(lblTicketDateCheck, uctlTicketDateCheck, False) Then
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
   m_Ticket.TICKET_NUMBER = "ประมาณการ"
   m_Ticket.TICKET_AMOUNT = FormatNumber(txtTicketAmount.Text)
   m_Ticket.TICKET_DATE_CHECK = uctlTicketDateCheck.ShowDate
    m_Ticket.CUSTOMER_ID = cboCustomerName.ItemData(Minus2Zero(cboCustomerName.ListIndex))
    m_Ticket.BANK_ID = cboBankID.ItemData(Minus2Zero(cboBankID.ListIndex))
    m_Ticket.MASTER_AREA = 2
   
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
      
      Call LoadBank(cboBankID)
      Call LoadBankCustomer(cboCustomerName)
      
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
   
   Call InitNormalLabel(lblTicketDate, MapText("วันที่"))
   Call InitNormalLabel(lblTicketAmount, MapText("จำนวนเงิน ปมก."))
   Call InitNormalLabel(lblTicketDateCheck, MapText("วันที่หน้าเช็ค"))
   uctlTicketDateCheck.TabStop = False
   Call InitNormalLabel(lblCustomerName, MapText("ลูกหนี้"))
   Call InitNormalLabel(lblBankID, MapText("ธนาคาร"))
   Call InitCombo(cboCustomerName)
   Call InitCombo(cboBankID)
   
   Call InitNormalLabel(Label1, MapText("บาท"))
   Call InitNormalLabel(Label2, MapText("วัน"))

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
   
   m_HasActivate = False
   Call InitFormLayout
   Call EnableForm(Me, True)
End Sub
Private Sub Form_Unload(Cancel As Integer)
   Set m_Ticket = Nothing
   Set m_Parts = Nothing
End Sub
Private Sub txtTicketAmount_Change()
   m_HasModify = True
End Sub
Private Sub txtTicketDateCheck_Change()
   If Len(txtTicketDateCheck.Text) <> 0 Then
      uctlTicketDateCheck.ShowDate = DateAdd("D", txtTicketDateCheck.Text, uctlTicketDate.ShowDate)
   Else
      uctlTicketDateCheck.ShowDate = -1
   End If
   m_HasModify = True
End Sub
Private Sub uctlTicketDate_HasChange()
   m_HasModify = True
End Sub
Private Sub uctlTicketDateCheck_HasChange()
   m_HasModify = True
End Sub
Private Sub cboCustomerName_Click()
   m_HasModify = True
End Sub
Private Sub cboBankID_Click()
   m_HasModify = True
End Sub
