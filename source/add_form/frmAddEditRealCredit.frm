VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#2.0#0"; "THREED20.OCX"
Begin VB.Form frmAddEditRealCredit 
   BackColor       =   &H80000000&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   4770
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11385
   Icon            =   "frmAddEditRealCredit.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4770
   ScaleWidth      =   11385
   StartUpPosition =   1  'CenterOwner
   Begin Threed.SSFrame SSFrame1 
      Height          =   4845
      Left            =   0
      TabIndex        =   6
      Top             =   0
      Width           =   11625
      _ExtentX        =   20505
      _ExtentY        =   8546
      _Version        =   131073
      PictureBackgroundStyle=   2
      Begin prjLedgerReport.uctlTextBox txtRealCredit 
         Height          =   435
         Left            =   2220
         TabIndex        =   2
         Top             =   2520
         Width           =   2925
         _ExtentX        =   5159
         _ExtentY        =   767
      End
      Begin Threed.SSPanel pnlHeader 
         Height          =   705
         Left            =   0
         TabIndex        =   7
         Top             =   0
         Width           =   11475
         _ExtentX        =   20241
         _ExtentY        =   1244
         _Version        =   131073
         PictureBackgroundStyle=   2
      End
      Begin prjLedgerReport.uctlTextBox txtDocumentNo 
         Height          =   435
         Left            =   2220
         TabIndex        =   1
         Top             =   2070
         Width           =   2955
         _ExtentX        =   5212
         _ExtentY        =   767
      End
      Begin prjLedgerReport.uctlTextBox txtCustomerCode 
         Height          =   450
         Left            =   2220
         TabIndex        =   0
         Top             =   1560
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   794
      End
      Begin prjLedgerReport.uctlTextBox txtCustomerName 
         Height          =   450
         Left            =   3660
         TabIndex        =   11
         Top             =   1560
         Width           =   4695
         _ExtentX        =   8281
         _ExtentY        =   794
      End
      Begin VB.Label lblCustomerCode 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   375
         Left            =   210
         TabIndex        =   10
         Top             =   1680
         Width           =   1935
      End
      Begin Threed.SSCheck chkPaidFlag 
         Height          =   375
         Left            =   4080
         TabIndex        =   3
         Top             =   3240
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   661
         _Version        =   131073
         Caption         =   "SSCheck1"
      End
      Begin VB.Label lblRealCredit 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   435
         Left            =   210
         TabIndex        =   9
         Top             =   2580
         Width           =   1935
      End
      Begin VB.Label lblDocumentNo 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   435
         Left            =   210
         TabIndex        =   8
         Top             =   2160
         Width           =   1935
      End
      Begin Threed.SSCommand cmdExit 
         Cancel          =   -1  'True
         Height          =   525
         Left            =   5715
         TabIndex        =   5
         Top             =   3750
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   926
         _Version        =   131073
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdOK 
         Height          =   525
         Left            =   4065
         TabIndex        =   4
         Top             =   3750
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmAddEditRealCredit.frx":27A2
         ButtonStyle     =   3
      End
   End
End
Attribute VB_Name = "frmAddEditRealCredit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private m_HasActivate As Boolean
Private m_HasModify As Boolean
Private m_Rs As ADODB.Recordset
Private m_RealCredit As CRealCredit

Public ID As Long
Public OKClick As Boolean
Public ShowMode As SHOW_MODE_TYPE
Public HeaderText As String
Private DocNoColl As Collection
Public CustomerColl As Collection
Private Sub chkPaidFlag_Click(Value As Integer)
   m_HasModify = True
End Sub

Private Sub ChkPaidFLag_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 13 Then
      CreateObject("WScript.Shell").SendKeys "{TAB}"
   End If
End Sub
'   m_HasModify = True
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
      
      m_RealCredit.ID = ID
      m_RealCredit.QueryFlag = 1
      If Not glbDaily.QueryRealCredit(m_RealCredit, m_Rs, ItemCount, IsOK, glbErrorLog) Then
         glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
         Call EnableForm(Me, True)
         Exit Sub
      End If
   End If
   
   If ItemCount > 0 Then
      Call m_RealCredit.PopulateFromRS(1, m_Rs)
      
      txtDocumentNo.Text = m_RealCredit.DOCUMENT_NO
      txtRealCredit.Text = m_RealCredit.REAL_CREDIT
      txtCustomerCode.Text = m_RealCredit.CUSTOMER_CODE
      txtCustomerName.Text = m_RealCredit.CUSTOMER_NAME
      chkPaidFlag.Value = FlagToCheck(m_RealCredit.PAID_FLAG)
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
Dim TempData As CARTrn

   If Not VerifyTextControl(lblCustomerCode, txtCustomerCode, False) Then
      Exit Function
   End If
   If Not VerifyTextControl(lblCustomerCode, txtCustomerName, False) Then
      Exit Function
   End If
   If Not VerifyTextControl(lblRealCredit, txtRealCredit, False) Then
      Exit Function
   End If
   
   If Not CheckUniqueNs(REAL_CREDIT_NO, txtDocumentNo.Text, ID) Then
      glbErrorLog.LocalErrorMsg = MapText("มีข้อมูล") & " " & txtDocumentNo.Text & " " & MapText("อยู่ในระบบแล้ว")
      glbErrorLog.ShowUserError
      Exit Function
   End If
   
   If Len(txtDocumentNo.Text) > 0 Then
      Call LoadARAmountByCust(Nothing, DocNoColl, -1, -1, txtDocumentNo.Text)
      If DocNoColl.Count <= 0 Then
         glbErrorLog.LocalErrorMsg = MapText("ไม่มีข้อมูลหมายเลขเอกสาร") & " " & txtDocumentNo.Text & " " & MapText("ในระบบ")
         glbErrorLog.ShowUserError
         Exit Function
      Else
         Set TempData = GetObject("CARTrn", DocNoColl, Trim(txtCustomerCode.Text), False)
         If TempData Is Nothing Then
            glbErrorLog.LocalErrorMsg = MapText("รหัสลูกค้าไม่ตรง")
            glbErrorLog.ShowUserError
            Exit Function
         End If
      End If
   End If
   
   If Not m_HasModify Then
      SaveData = True
      Exit Function
   End If
   
   m_RealCredit.AddEditMode = ShowMode
   m_RealCredit.ID = ID
   m_RealCredit.DOCUMENT_NO = txtDocumentNo.Text
   m_RealCredit.REAL_CREDIT = Val(txtRealCredit.Text)
   m_RealCredit.CUSTOMER_CODE = txtCustomerCode.Text
   m_RealCredit.CUSTOMER_NAME = txtCustomerName.Text
   
   m_RealCredit.PAID_FLAG = ("Y" = Check2Flag(chkPaidFlag.Value))
      
   Call EnableForm(Me, False)
   If Not glbDaily.AddEditRealCredit(m_RealCredit, IsOK, True, glbErrorLog) Then
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
      
      Call LoadCustomerPro(CustomerColl)
      
      If ShowMode = SHOW_EDIT Then
         Call QueryData(True)
      ElseIf ShowMode = SHOW_ADD Then
         ID = 0
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
   
   Call InitNormalLabel(lblDocumentNo, MapText("หมายเลขเอกสาร"))
   Call InitNormalLabel(lblRealCredit, MapText("เครดิตจริง (วัน)"))
   Call InitNormalLabel(lblCustomerCode, MapText("ลูกค้า"))
   
   Call txtDocumentNo.SetTextLenType(TEXT_STRING, glbSetting.NAME_LEN)
   Call txtRealCredit.SetTextLenType(TEXT_INTEGER, glbSetting.MONEY_TYPE)
   
   Call txtCustomerCode.SetTextLenType(TEXT_STRING, glbSetting.NAME_LEN)
   Call txtCustomerName.SetTextLenType(TEXT_STRING, glbSetting.NAME_LEN)
   
   Call txtCustomerCode.SetKeySearch("CUSTOMER_CODE")
   
   txtCustomerName.Enabled = False
   
   Me.Picture = LoadPicture(glbParameterObj.MainPicture)
   
   pnlHeader.Font.Name = GLB_FONT
   pnlHeader.Font.Bold = True
   pnlHeader.Font.Size = 19
   
   cmdExit.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdOK.Picture = LoadPicture(glbParameterObj.NormalButton1)
   
   Call InitCheckBox(chkPaidFlag, "จ่ายแล้ว")
   
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
   
   Set m_RealCredit = New CRealCredit
   Set m_Rs = New ADODB.Recordset
   Set DocNoColl = New Collection
   Set CustomerColl = New Collection
   
   m_HasActivate = False
   Call InitFormLayout
   Call EnableForm(Me, True)
End Sub
Private Sub Form_Unload(Cancel As Integer)
   Set m_RealCredit = Nothing
   Set DocNoColl = Nothing
   Set CustomerColl = Nothing
End Sub

Private Sub txtCustomerCode_Change()
Dim TempCustomer As CARMas
   m_HasModify = True
   Set TempCustomer = GetObject("CARMas", CustomerColl, txtCustomerCode.Text, False)
   If Not TempCustomer Is Nothing Then
      txtCustomerName.Text = TempCustomer.CUSNAM
   Else
      txtCustomerName.Text = ""
   End If
   
   m_HasModify = True
End Sub

Private Sub txtDocumentNo_Change()
   m_HasModify = True
End Sub

Private Sub txtRealCredit_Change()
   m_HasModify = True
End Sub
