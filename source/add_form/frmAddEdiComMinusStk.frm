VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#2.0#0"; "THREED20.OCX"
Begin VB.Form frmAddEditComMinusStk 
   BackColor       =   &H80000000&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   5205
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9945
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5205
   ScaleWidth      =   9945
   StartUpPosition =   1  'CenterOwner
   Begin Threed.SSFrame SSFrame1 
      Height          =   5205
      Left            =   0
      TabIndex        =   7
      Top             =   0
      Width           =   9945
      _ExtentX        =   17542
      _ExtentY        =   9181
      _Version        =   131073
      PictureBackgroundStyle=   2
      Begin prjLedgerReport.uctlTextBox txtMinusAmount 
         Height          =   495
         Left            =   2520
         TabIndex        =   2
         Top             =   3000
         Width           =   2895
         _ExtentX        =   5106
         _ExtentY        =   873
      End
      Begin Threed.SSPanel pnlHeader 
         Height          =   705
         Left            =   0
         TabIndex        =   8
         Top             =   0
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
      Begin prjLedgerReport.uctlTextLookup uctlStkCodStkdesLookup 
         Height          =   375
         Left            =   2520
         TabIndex        =   1
         Top             =   2400
         Width           =   6255
         _ExtentX        =   11033
         _ExtentY        =   661
      End
      Begin prjLedgerReport.uctlTextLookup uctlCommissionSale 
         Height          =   375
         Left            =   2520
         TabIndex        =   11
         Top             =   1800
         Width           =   6255
         _ExtentX        =   11033
         _ExtentY        =   661
      End
      Begin prjLedgerReport.uctlTextBox txtSRcod 
         Height          =   495
         Left            =   2520
         TabIndex        =   3
         Top             =   3600
         Width           =   2895
         _ExtentX        =   5106
         _ExtentY        =   873
      End
      Begin VB.Label lblSRcod 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   435
         Left            =   960
         TabIndex        =   13
         Top             =   3600
         Width           =   1455
      End
      Begin VB.Label lblSaleName 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   435
         Left            =   480
         TabIndex        =   12
         Top             =   1800
         Width           =   1935
      End
      Begin VB.Label lblMinus 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   435
         Left            =   480
         TabIndex        =   10
         Top             =   3000
         Width           =   1935
      End
      Begin VB.Label lblStk 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   435
         Left            =   480
         TabIndex        =   9
         Top             =   2400
         Width           =   1935
      End
      Begin VB.Label lblIVCod 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   435
         Left            =   480
         TabIndex        =   6
         Top             =   1200
         Width           =   1935
      End
      Begin Threed.SSCommand cmdExit 
         Cancel          =   -1  'True
         Height          =   525
         Left            =   4440
         TabIndex        =   5
         Top             =   4440
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   926
         _Version        =   131073
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdOK 
         Height          =   525
         Left            =   2640
         TabIndex        =   4
         Top             =   4440
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         ButtonStyle     =   3
      End
   End
End
Attribute VB_Name = "frmAddEditComMinusStk"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private m_HasActivate As Boolean
Private m_HasModify As Boolean
Private m_Rs As ADODB.Recordset
Private itemComMinus  As CComMinusStk

Public OKClick As Boolean
Public ShowMode As SHOW_MODE_TYPE
Public HeaderText As String

Public COM_MINUSSTK_ID As Long
Public m_exitIVStkcod As Collection
Private temp_ComMinusStk As CComMinusStk
Public lookupSLMCOD As String

Public m_Stcrd As Collection
Public m_Stcrd4Date As Collection
Private FtSaleColl As Collection

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
      
      itemComMinus.COM_MINUSSTK_ID = COM_MINUSSTK_ID
      If Not glbDaily.QueryMinusAmount(itemComMinus, Nothing, m_Rs, ItemCount, IsOK, glbErrorLog) Then
         glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
         Call EnableForm(Me, True)
         Exit Sub
      End If
   End If
   
   If ItemCount > 0 Then
         Call itemComMinus.PopulateFromRS(1, m_Rs)
         txtIVCod.Text = itemComMinus.IV_COD
         uctlStkCodStkdesLookup.MyTextBox.Text = itemComMinus.STK_COD
'         uctlStkCodStkdesLookup.MyCombo.Text = itemComMinus.STK_NAME
      '   uctlDocDate.ShowDate = itemComMinus.IV_DOCDAT
         txtMinusAmount.Text = itemComMinus.MINUS_AMOUNT
         txtSRcod.Text = itemComMinus.MINUS_COD
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
Dim m_Lookup As CComMinusStk
  Set m_Lookup = New CComMinusStk

   If Not VerifyTextControl(lblIVCod, txtIVCod, False) Then
      Exit Function
   End If
   
   If Not VerifyTextControl(lblMinus, txtMinusAmount, False) Then
      Exit Function
   End If
   
   If Not Len(uctlStkCodStkdesLookup.MyCombo.Text) > 0 Then
      Exit Function
   End If
   
   If Not VerifyTextControl(lblSaleName, uctlCommissionSale.MyTextBox, False) Then
      Exit Function
   End If
   
  If ShowMode = SHOW_ADD Then
      Set temp_ComMinusStk = GetMinusCommiss(m_exitIVStkcod, Trim(txtIVCod.Text & "-" & uctlStkCodStkdesLookup.MyTextBox.Text), False)
      If Not (temp_ComMinusStk Is Nothing) Then
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
   
   itemComMinus.AddEditMode = ShowMode
 '  itemComMinus.IV_DOCDAT = uctlDocDate.ShowDate
   itemComMinus.IV_COD = txtIVCod.Text
   itemComMinus.SLMCOD = uctlCommissionSale.MyTextBox.Text
   itemComMinus.SLMNAME = uctlCommissionSale.MyCombo.Text
   itemComMinus.STK_COD = uctlStkCodStkdesLookup.MyTextBox.Text
   itemComMinus.STK_NAME = uctlStkCodStkdesLookup.MyCombo.Text
   itemComMinus.MINUS_AMOUNT = txtMinusAmount.Text
   itemComMinus.MINUS_COD = txtSRcod.Text
   
      Set D = GetObject("CStcrd", m_Stcrd4Date, Trim(txtIVCod.Text), False)
      If Not (D Is Nothing) Then
          itemComMinus.IV_DOCDAT = D.DOCDAT
      End If
   
   Call EnableForm(Me, False)
   If Not glbDaily.AddEditMinusAmount(itemComMinus, IsOK, True, glbErrorLog) Then
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
      
      m_HasModify = False
            
      Call LoadStkcodLookup(uctlStkCodStkdesLookup.MyCombo, m_Stcrd)
      Set uctlStkCodStkdesLookup.MyCollection = m_Stcrd
      
       Call LoadIVfromCStcrd(Nothing, m_Stcrd4Date)
       
      Call LoadSaleLookup(uctlCommissionSale.MyCombo, FtSaleColl) 'FtSaleColl, COMMISSION_TABLE
      Set uctlCommissionSale.MyCollection = FtSaleColl
      
      If ShowMode = SHOW_EDIT Then
         Call QueryData(True)
      ElseIf ShowMode = SHOW_ADD Then
         COM_MINUSSTK_ID = -1
      End If
      
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
   
   Call InitNormalLabel(lblStk, MapText("เลขที่สินค้า"))

   Call InitNormalLabel(lblMinus, MapText("ส่วนลด"))
   Call txtMinusAmount.SetTextLenType(TEXT_STRING, glbSetting.DOUBLE_TYPE)

   Call InitNormalLabel(lblSRcod, MapText("หมายเลข SR"))
   Call txtSRcod.SetTextLenType(TEXT_STRING, glbSetting.CODE_TYPE)
   
   Call InitNormalLabel(lblSaleName, MapText("พนักงานขาย"))
   uctlCommissionSale.Enabled = False

   If ShowMode = SHOW_ADD Then
         txtIVCod.Text = "IV"
   End If


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
      
   Set itemComMinus = New CComMinusStk
   Set m_Stcrd = New Collection
   Set m_Stcrd4Date = New Collection
   Set FtSaleColl = New Collection
   
   m_HasActivate = False
   Call InitFormLayout
   Call EnableForm(Me, True)
End Sub
Private Sub Form_Unload(Cancel As Integer)
   Set itemComMinus = Nothing
   Set m_Stcrd = Nothing
   Set m_Stcrd4Date = Nothing
   Set FtSaleColl = Nothing
End Sub

Private Sub txtIVCod_LostFocus()
      Call LoadStkcodLookup(uctlStkCodStkdesLookup.MyCombo, m_Stcrd, Nothing, txtIVCod.Text, True)
      Set uctlStkCodStkdesLookup.MyCollection = m_Stcrd
      
      Call LoadIVsaleLookup(Nothing, lookupSLMCOD, Nothing, txtIVCod.Text, True)
      uctlCommissionSale.MyTextBox.Text = lookupSLMCOD

   m_HasModify = True
End Sub

Private Sub txtMinusAmount_Change()
   m_HasModify = True
End Sub
Private Sub uctlStkCodStkdesLookup_HasChange()
   m_HasModify = True
End Sub
Private Sub uctlCommissionSale_HasChange()
   m_HasModify = True
End Sub
Private Sub uctlStkCodStkdesLookup_Change()
   m_HasModify = True
End Sub
Private Sub uctlCommissionSale_Change()
   m_HasModify = True
End Sub
