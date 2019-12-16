VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#2.0#0"; "THREED20.OCX"
Begin VB.Form frmAddEditComIVcredit 
   BackColor       =   &H80000000&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   5220
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9990
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5220
   ScaleWidth      =   9990
   StartUpPosition =   1  'CenterOwner
   Begin Threed.SSFrame SSFrame1 
      Height          =   8325
      Left            =   0
      TabIndex        =   6
      Top             =   -120
      Width           =   10905
      _ExtentX        =   19235
      _ExtentY        =   14684
      _Version        =   131073
      PictureBackgroundStyle=   2
      Begin VB.ComboBox cboCRtype 
         Height          =   315
         Left            =   2520
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   3360
         Width           =   3135
      End
      Begin Threed.SSPanel pnlHeader 
         Height          =   705
         Left            =   0
         TabIndex        =   7
         Top             =   120
         Width           =   10875
         _ExtentX        =   19182
         _ExtentY        =   1244
         _Version        =   131073
         PictureBackgroundStyle=   2
      End
      Begin prjLedgerReport.uctlTextBox txtIVCod 
         Height          =   495
         Left            =   2520
         TabIndex        =   0
         Top             =   1200
         Width           =   2895
         _ExtentX        =   5106
         _ExtentY        =   873
      End
      Begin prjLedgerReport.uctlTextLookup uctlCommissionSale 
         Height          =   375
         Left            =   2520
         TabIndex        =   8
         Top             =   1800
         Width           =   6255
         _ExtentX        =   11033
         _ExtentY        =   661
      End
      Begin prjLedgerReport.uctlTextLookup uctlCustomerLookup 
         Height          =   375
         Left            =   2520
         TabIndex        =   10
         Top             =   2280
         Width           =   6255
         _ExtentX        =   11033
         _ExtentY        =   661
      End
      Begin prjLedgerReport.uctlTextBox txtCRdata 
         Height          =   375
         Left            =   2520
         TabIndex        =   2
         Top             =   3840
         Width           =   3135
         _ExtentX        =   5106
         _ExtentY        =   873
      End
      Begin prjLedgerReport.uctlTextBox txtNetval 
         Height          =   495
         Left            =   2520
         TabIndex        =   14
         Top             =   2760
         Width           =   2895
         _ExtentX        =   5106
         _ExtentY        =   873
      End
      Begin VB.Label lblNetval 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   435
         Left            =   360
         TabIndex        =   15
         Top             =   2760
         Width           =   1935
      End
      Begin VB.Label lblCRtype 
         Alignment       =   1  'Right Justify
         Caption         =   "Label2"
         Height          =   435
         Left            =   720
         TabIndex        =   13
         Top             =   3360
         Width           =   1575
      End
      Begin VB.Label lblCRdata 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   435
         Left            =   360
         TabIndex        =   12
         Top             =   3840
         Width           =   1935
      End
      Begin VB.Label lblCustomer 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   375
         Left            =   720
         TabIndex        =   11
         Top             =   2280
         Width           =   1575
      End
      Begin VB.Label lblSaleName 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   435
         Left            =   360
         TabIndex        =   9
         Top             =   1800
         Width           =   1935
      End
      Begin VB.Label lblIVCod 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   435
         Left            =   360
         TabIndex        =   5
         Top             =   1200
         Width           =   1935
      End
      Begin Threed.SSCommand cmdExit 
         Cancel          =   -1  'True
         Height          =   525
         Left            =   4080
         TabIndex        =   4
         Top             =   4560
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   926
         _Version        =   131073
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdOK 
         Height          =   525
         Left            =   2400
         TabIndex        =   3
         Top             =   4560
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         ButtonStyle     =   3
      End
   End
End
Attribute VB_Name = "frmAddEditComIVcredit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private m_HasActivate As Boolean
Private m_HasModify As Boolean
Private m_Rs As ADODB.Recordset
Private itemComIVcredit  As CComIVcredit

Public OKClick As Boolean
Public ShowMode As SHOW_MODE_TYPE
Public HeaderText As String

Public COM_IV_CREDIT_ID As Long
Public m_exitIVcredit As Collection
Private temp_ComIVcredit As CComIVcredit

Public lookupSLMCOD As String
Public lookupCUS_ID As String
Public m_IV4Date As Collection

Private FtSaleColl As Collection
Public m_Customer As Collection
Public tempSumIV  As CStcrd

Private Sub cboDataType_Click()
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
      
      itemComIVcredit.COM_IV_CREDIT_ID = COM_IV_CREDIT_ID
      If Not glbDaily.QueryIVcredit(itemComIVcredit, Nothing, m_Rs, ItemCount, IsOK, glbErrorLog) Then
         glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
         Call EnableForm(Me, True)
         Exit Sub
      End If
   End If
   
   If ItemCount > 0 Then
         Call itemComIVcredit.PopulateFromRS(1, m_Rs)
         txtIVCod.Text = itemComIVcredit.IV_COD
         uctlCommissionSale.MyTextBox.Text = itemComIVcredit.SLMCOD
         uctlCustomerLookup.MyTextBox.Text = itemComIVcredit.CUS_ID
         cboCRtype.ListIndex = IDToListIndex(cboCRtype, Code2AdjustType(itemComIVcredit.CR_TYPE))
         txtCRdata.Text = itemComIVcredit.CR_DATA
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
Dim D As CStcrd
Dim m_Lookup As CComIVcredit
  Set m_Lookup = New CComIVcredit

   If Not VerifyTextControl(lblIVCod, txtIVCod, False) Then
      Exit Function
   End If
   
   If Not VerifyTextControl(lblCRdata, txtCRdata, False) Then
      Exit Function
   End If
   
   If Not VerifyTextControl(lblSaleName, uctlCommissionSale.MyTextBox, False) Then
      Exit Function
   End If
   
  If Not VerifyTextControl(lblCustomer, uctlCustomerLookup.MyTextBox, False) Then
      Exit Function
   End If

   
  If Not VerifyCombo(lblCRtype, cboCRtype, False) Then
      Exit Function
   End If
   
  If ShowMode = SHOW_ADD Then
      Set temp_ComIVcredit = GetIVcredit(m_exitIVcredit, Trim(txtIVCod.Text), False)
      If Not (temp_ComIVcredit Is Nothing) Then
                If Not DuplicateData() Then
                    Exit Function
                End If
                Exit Function
      End If
 Else
      Set temp_ComIVcredit = GetIVcredit(m_exitIVcredit, Trim(txtIVCod.Text), False)
      If Not (temp_ComIVcredit Is Nothing) And Trim(txtIVCod.Text) <> itemComIVcredit.IV_COD Then
                If Not DuplicateData() Then
                    Exit Function
                End If
                Exit Function
      End If
 End If
   
   If Not m_HasModify Then
      SaveData = True
      Exit Function
   End If
   
   itemComIVcredit.AddEditMode = ShowMode
 '  itemComIVcredit.IV_DOCDAT = uctlDocDate.ShowDate
   itemComIVcredit.IV_COD = txtIVCod.Text
   itemComIVcredit.SLMCOD = uctlCommissionSale.MyTextBox.Text
   itemComIVcredit.SLMNAME = uctlCommissionSale.MyCombo.Text
      itemComIVcredit.CUS_ID = uctlCustomerLookup.MyTextBox.Text
   itemComIVcredit.CUS_NAME = uctlCustomerLookup.MyCombo.Text
   itemComIVcredit.CR_TYPE = AdjustType2Code(cboCRtype.ItemData(Minus2Zero(cboCRtype.ListIndex)))
   itemComIVcredit.CR_DATA = txtCRdata.Text

   
      Set D = GetObject("CStcrd", m_IV4Date, Trim(txtIVCod.Text), False)
      If Not (D Is Nothing) Then
          itemComIVcredit.IV_DOCDAT = D.DOCDAT
      End If
   
   Call EnableForm(Me, False)
   If Not glbDaily.AddEditIVcredit(itemComIVcredit, IsOK, True, glbErrorLog) Then
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
      
      Call InitAdjustType(cboCRtype)
      
      Call LoadSaleLookup(uctlCommissionSale.MyCombo, FtSaleColl) 'FtSaleColl, COMMISSION_TABLE
      Set uctlCommissionSale.MyCollection = FtSaleColl
      
      Call LoadIVfromCStcrd(Nothing, m_IV4Date)
      
       Call LoadCustomerLookup(uctlCustomerLookup.MyCombo, m_Customer)
      Set uctlCustomerLookup.MyCollection = m_Customer
      
      If ShowMode = SHOW_EDIT Then
         Call QueryData(True)
      ElseIf ShowMode = SHOW_ADD Then
         COM_IV_CREDIT_ID = -1
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
   
   Call InitNormalLabel(lblIVCod, MapText("INVOICE"))
   Call txtIVCod.SetTextLenType(TEXT_STRING, glbSetting.CODE_TYPE)
   If ShowMode = SHOW_ADD Then
      txtIVCod.Text = "IV"
   End If
   
   Call InitNormalLabel(lblCRdata, MapText("จำนวนเครดิต"))
   Call txtCRdata.SetTextLenType(TEXT_INTEGER, glbSetting.DESC_TYPE)
   
   Call InitNormalLabel(lblCustomer, MapText("ลูกค้า"))
   Call InitNormalLabel(lblSaleName, MapText("พนักงานขาย"))
   Call InitNormalLabel(lblCRtype, MapText("ประเภท"))
   Call InitNormalLabel(lblNetval, MapText("มูลค่า"))
      
   uctlCommissionSale.Enabled = False
   uctlCustomerLookup.Enabled = False
   txtNetval.Enabled = False
   
      Call InitCombo(cboCRtype)
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
   Set m_Rs = New ADODB.Recordset
      
   Set itemComIVcredit = New CComIVcredit
   Set m_Customer = New Collection
   Set m_IV4Date = New Collection
   Set FtSaleColl = New Collection
   Set tempSumIV = New CStcrd
   
   m_HasActivate = False
   Call InitFormLayout
   Call EnableForm(Me, True)
End Sub
Private Sub Form_Unload(Cancel As Integer)
   Set itemComIVcredit = Nothing
   Set m_Customer = Nothing
   Set m_IV4Date = Nothing
   Set FtSaleColl = Nothing
  Set tempSumIV = Nothing
End Sub
Private Sub txtMinusAmount_Change()
   m_HasModify = True
End Sub

Private Sub txtIVCod_Change()
      Call LoadIVsaleLookup(Nothing, lookupSLMCOD, Nothing, txtIVCod.Text, True)
      uctlCommissionSale.MyTextBox.Text = lookupSLMCOD
      
     Call LoadIVcustomerLookup(lookupCUS_ID, Nothing, txtIVCod.Text, True)
      uctlCustomerLookup.MyTextBox.Text = lookupCUS_ID
      
      Set tempSumIV = GetObject("Cstcrd", m_IV4Date, txtIVCod.Text, False)
      If Not (tempSumIV Is Nothing) Then
         txtNetval.Text = FormatNumber(tempSumIV.NETVAL)
      End If
      
   m_HasModify = True
End Sub
Private Sub uctlCommissionSale_HasChange()
   m_HasModify = True
End Sub

Private Sub uctlCommissionSale_Change()
   m_HasModify = True
End Sub

Private Sub uctlCustomerLookup_HasChange()
   m_HasModify = True
End Sub

Private Sub uctlCustomerLookup_Change()
   m_HasModify = True
End Sub

Private Sub cboCRtype_Click()
 m_HasModify = True
End Sub
Private Sub txtCRdata_Change()
   m_HasModify = True
End Sub
