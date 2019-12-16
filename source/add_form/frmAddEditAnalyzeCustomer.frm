VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#2.0#0"; "THREED20.OCX"
Begin VB.Form frmAddEditAnalyzeCustomer 
   BackColor       =   &H80000000&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   3930
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5385
   Icon            =   "frmAddEditAnalyzeCustomer.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3930
   ScaleWidth      =   5385
   StartUpPosition =   1  'CenterOwner
   Begin Threed.SSFrame SSFrame1 
      Height          =   4005
      Left            =   0
      TabIndex        =   5
      Top             =   0
      Width           =   5505
      _ExtentX        =   9710
      _ExtentY        =   7064
      _Version        =   131073
      PictureBackgroundStyle=   2
      Begin prjLedgerReport.uctlDate uctlDatePayment 
         Height          =   375
         Left            =   1440
         TabIndex        =   1
         Top             =   2640
         Width           =   3855
         _ExtentX        =   6800
         _ExtentY        =   661
      End
      Begin prjLedgerReport.uctlDate uctlPutDate 
         Height          =   375
         Left            =   1440
         TabIndex        =   10
         Top             =   2040
         Width           =   3855
         _ExtentX        =   6800
         _ExtentY        =   661
      End
      Begin Threed.SSPanel pnlHeader 
         Height          =   705
         Left            =   0
         TabIndex        =   6
         Top             =   0
         Width           =   5475
         _ExtentX        =   9657
         _ExtentY        =   1244
         _Version        =   131073
         PictureBackgroundStyle=   2
      End
      Begin prjLedgerReport.uctlTextBox txtCustomerName 
         Height          =   435
         Left            =   1440
         TabIndex        =   0
         Top             =   960
         Width           =   3555
         _ExtentX        =   5212
         _ExtentY        =   767
      End
      Begin prjLedgerReport.uctlTextBox txtInvoice 
         Height          =   435
         Left            =   1440
         TabIndex        =   9
         Top             =   1440
         Width           =   3555
         _ExtentX        =   5212
         _ExtentY        =   767
      End
      Begin Threed.SSCommand cmdNext 
         Height          =   525
         Left            =   120
         TabIndex        =   2
         Top             =   3240
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         ButtonStyle     =   3
      End
      Begin VB.Label lblDatePayment 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   255
         Left            =   120
         TabIndex        =   12
         Top             =   2760
         Width           =   1215
      End
      Begin VB.Label lblPutDate 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   255
         Left            =   240
         TabIndex        =   11
         Top             =   2160
         Width           =   1095
      End
      Begin VB.Label lblInvoice 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   255
         Left            =   240
         TabIndex        =   8
         Top             =   1560
         Width           =   1095
      End
      Begin VB.Label lblCustomerName 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   315
         Left            =   240
         TabIndex        =   7
         Top             =   1080
         Width           =   1095
      End
      Begin Threed.SSCommand cmdExit 
         Cancel          =   -1  'True
         Height          =   525
         Left            =   3480
         TabIndex        =   4
         Top             =   3240
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   926
         _Version        =   131073
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdOK 
         Height          =   525
         Left            =   1800
         TabIndex        =   3
         Top             =   3240
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmAddEditAnalyzeCustomer.frx":27A2
         ButtonStyle     =   3
      End
   End
End
Attribute VB_Name = "frmAddEditAnalyzeCustomer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private m_HasActivate As Boolean
Private m_HasModify As Boolean
Private m_Rs As ADODB.Recordset
Private m_AnalyzeCustomer As CAnalyzeCustomer
Private m_Customer As CARTrn
Public ParentForm As Form

Public ID As Long
Public SubID As Long
Public haveData As Long
Public OKClick As Boolean
Public ShowMode As SHOW_MODE_TYPE
Public HeaderText As String
Public TempAddEditDataCollection As Collection
Public RunDataCollection As Collection
Private Sub cmdNext_Click()
Dim NewID As Long
   m_HasModify = True
   If Not SaveData Then
      Exit Sub
   End If

   NewID = GetNextID(SubID, RunDataCollection)
   If SubID = NewID Then
      glbErrorLog.LocalErrorMsg = "ถึงเรคคอร์ดสุดท้ายแล้ว"
      glbErrorLog.ShowUserError
   
      Call ParentForm.GridEX1.Rebind
      Exit Sub
   End If
   
   SubID = NewID
   
   Call ParentForm.GridEX1.Rebind

   Call QueryData(True)

   Call uctlDatePayment.SetFocus
   m_HasModify = False
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
   Set m_Customer = New CARTrn

   Set m_Customer = GetARTrn(RunDataCollection, Str(SubID))
   txtInvoice.Text = m_Customer.DOCNUM
   uctlPutDate.ShowDate = m_Customer.DOCDAT
   txtCustomerName.Text = m_Customer.CUSNAM

   Set m_AnalyzeCustomer = GetAnalyzeCustomer(TempAddEditDataCollection, Trim(m_Customer.DOCNUM))
   ID = m_AnalyzeCustomer.ANALYZE_CUSTOMER_ID
   If ID = 0 Then
      uctlDatePayment.ShowDate = -1
   Else
      uctlDatePayment.ShowDate = m_AnalyzeCustomer.DATE_OF_PAYMENT
   End If
   
    m_HasModify = False
   Call EnableForm(Me, True)
End Sub
Private Function SaveData() As Boolean
Dim IsOK As Boolean
   
   If Not VerifyDate(lblDatePayment, uctlDatePayment, False) Then
      Exit Function
   End If
   
   If Not m_HasModify Then
      SaveData = True
      Exit Function
   End If
   
   If ID <> 0 Then   'มีข้อมูลอยู่แล้ว
      m_AnalyzeCustomer.ANALYZE_CUSTOMER_ID = ID
      m_AnalyzeCustomer.AddEditMode = ShowMode
   Else                       'เพิ่มเข้าไปใหม่
      m_AnalyzeCustomer.AddEditMode = SHOW_ADD
   End If
   m_AnalyzeCustomer.INVOICE = txtInvoice.Text
   m_AnalyzeCustomer.DATE_OF_PAYMENT = uctlDatePayment.ShowDate
   
   Call EnableForm(Me, False)
   If Not glbDaily.AddEditAnalyzeCustomer(m_AnalyzeCustomer, IsOK, True, glbErrorLog) Then
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
      
      Call QueryData(True)
      
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
   
   Call InitNormalLabel(lblInvoice, MapText("เลขที่ INV."))
   Call txtInvoice.SetKeySearch("INVOICE")
   txtInvoice.Enabled = False
   
   Call InitNormalLabel(lblCustomerName, MapText("ชื่อลูกค้า"))
   Call txtCustomerName.SetKeySearch("CUSTOMER_NAME")
   txtCustomerName.Enabled = False
   
   Call txtInvoice.SetTextLenType(TEXT_STRING, glbSetting.CODE_TYPE)
   Call txtCustomerName.SetTextLenType(TEXT_STRING, glbSetting.NAME_TYPE)
   
   Call InitNormalLabel(lblPutDate, MapText("วันที่ขาย"))
   uctlPutDate.Enable = False
   Call InitNormalLabel(lblDatePayment, MapText("วันที่ชำระ"))
   
   Me.Picture = LoadPicture(glbParameterObj.MainPicture)
   
   pnlHeader.Font.Name = GLB_FONT
   pnlHeader.Font.Bold = True
   pnlHeader.Font.Size = 19
   
   cmdExit.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdOK.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdNext.Picture = LoadPicture(glbParameterObj.NormalButton1)
   
   Call InitMainButton(cmdExit, MapText("ยกเลิก (ESC)"))
   Call InitMainButton(cmdOK, MapText("ตกลง (F2)"))
   Call InitMainButton(cmdNext, MapText("ถัดไป"))
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
   
   Set m_AnalyzeCustomer = New CAnalyzeCustomer
   Set m_Customer = New CARTrn
   Set m_Rs = New ADODB.Recordset
   
   m_HasActivate = False
   Call InitFormLayout
   Call EnableForm(Me, True)
End Sub
Private Sub Form_Unload(Cancel As Integer)
   Set m_AnalyzeCustomer = Nothing
   Set m_Customer = Nothing
   Set m_Rs = Nothing
End Sub
Private Sub uctlDatePayment_HasChange()
   m_HasModify = True
End Sub
