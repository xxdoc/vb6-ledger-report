VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#2.0#0"; "THREED20.OCX"
Begin VB.Form frmAddEditPromotionYear 
   BackColor       =   &H80000000&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   5250
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9420
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5250
   ScaleWidth      =   9420
   StartUpPosition =   1  'CenterOwner
   Begin Threed.SSFrame SSFrame1 
      Height          =   8325
      Left            =   0
      TabIndex        =   4
      Top             =   0
      Width           =   10905
      _ExtentX        =   19235
      _ExtentY        =   14684
      _Version        =   131073
      PictureBackgroundStyle=   2
      Begin prjLedgerReport.uctlTextBox txtProAmount 
         Height          =   450
         Left            =   2400
         TabIndex        =   0
         Top             =   3120
         Width           =   2895
         _extentx        =   5106
         _extenty        =   794
      End
      Begin Threed.SSPanel pnlHeader 
         Height          =   705
         Left            =   0
         TabIndex        =   5
         Top             =   0
         Width           =   9435
         _ExtentX        =   16642
         _ExtentY        =   1244
         _Version        =   131073
         PictureBackgroundStyle=   2
      End
      Begin prjLedgerReport.uctlDate uctlProDate 
         Height          =   375
         Left            =   2400
         TabIndex        =   7
         Top             =   1320
         Width           =   3855
         _extentx        =   6800
         _extenty        =   661
      End
      Begin prjLedgerReport.uctlTextBox txtProCustomerCode 
         Height          =   450
         Left            =   2400
         TabIndex        =   10
         Top             =   1920
         Width           =   1455
         _extentx        =   2566
         _extenty        =   794
      End
      Begin prjLedgerReport.uctlTextBox txtProCustomerName 
         Height          =   450
         Left            =   3900
         TabIndex        =   11
         Top             =   1920
         Width           =   4695
         _extentx        =   8281
         _extenty        =   794
      End
      Begin prjLedgerReport.uctlTextBox txtProSTKCOD 
         Height          =   450
         Left            =   2400
         TabIndex        =   12
         Top             =   2520
         Width           =   1455
         _extentx        =   2566
         _extenty        =   794
      End
      Begin prjLedgerReport.uctlTextBox txtProSTKName 
         Height          =   450
         Left            =   3900
         TabIndex        =   13
         Top             =   2520
         Width           =   4695
         _extentx        =   8281
         _extenty        =   794
      End
      Begin VB.Label lblProSTKCOD 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   375
         Left            =   600
         TabIndex        =   6
         Top             =   2640
         Width           =   1575
      End
      Begin VB.Label lblProCustomerCode 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   375
         Left            =   600
         TabIndex        =   9
         Top             =   2040
         Width           =   1575
      End
      Begin VB.Label lblProDate 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   375
         Left            =   600
         TabIndex        =   8
         Top             =   1320
         Width           =   1575
      End
      Begin VB.Label lblProAmount 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   435
         Left            =   240
         TabIndex        =   3
         Top             =   3195
         Width           =   1935
      End
      Begin Threed.SSCommand cmdExit 
         Cancel          =   -1  'True
         Height          =   525
         Left            =   4560
         TabIndex        =   2
         Top             =   4080
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   926
         _Version        =   131073
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdOK 
         Height          =   525
         Left            =   2400
         TabIndex        =   1
         Top             =   4080
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         ButtonStyle     =   3
      End
   End
End
Attribute VB_Name = "frmAddEditPromotionYear"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private m_HasActivate As Boolean
Private m_HasModify As Boolean
Private m_Rs As ADODB.Recordset
Private ItemPromotion  As CPromotionYear

Public ID As Long
Public OKClick As Boolean
Public ShowMode As SHOW_MODE_TYPE
Public HeaderText As String

Public CustomerProColl As Collection
Public StockProColl As Collection

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

      ItemPromotion.PROYEAR_ID = ID
      ItemPromotion.QueryFlag = 1
      If Not glbDaily.QueryPromotionYear(ItemPromotion, m_Rs, ItemCount, IsOK, glbErrorLog) Then
         glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
         Call EnableForm(Me, True)
         Exit Sub
      End If
   End If

   If ItemCount > 0 Then
      Call ItemPromotion.PopulateFromRS(1, m_Rs)
      
      uctlProDate.ShowDate = ItemPromotion.DATEYEAR_PRO
      txtProCustomerCode.Text = ItemPromotion.CTMCODYEAR_PRO
      txtProCustomerName.Text = ItemPromotion.CTMNAMEYEAR_PRO
      txtProSTKCOD.Text = ItemPromotion.STKCODYEAR_PRO
      txtProSTKName.Text = ItemPromotion.STKNAMEYEAR_PRO
      txtProAmount.Text = ItemPromotion.AMOUNTYEAR_PRO

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
  
   If Not VerifyDate(lblProDate, uctlProDate, False) Then
      Exit Function
   End If
   If Not VerifyTextControl(lblProCustomerCode, txtProCustomerCode, False) Then
      Exit Function
   End If
   If Not VerifyTextControl(lblProCustomerCode, txtProCustomerName, False) Then
      Exit Function
   End If
   If Not VerifyTextControl(lblProSTKCOD, txtProSTKCOD, False) Then
      Exit Function
   End If
   If Not VerifyTextControl(lblProSTKCOD, txtProSTKName, False) Then
      Exit Function
   End If
   If Not VerifyTextControl(lblProAmount, txtProAmount, False) Then
      Exit Function
   End If

   If Not m_HasModify Then
      SaveData = True
      Exit Function
   End If

   ItemPromotion.AddEditMode = ShowMode
   ItemPromotion.DATEYEAR_PRO = Format(uctlProDate.ShowDate, "DD") & "/" & Format(uctlProDate.ShowDate, "MM") & "/" & Format(uctlProDate.ShowDate, "YYYY")
   ItemPromotion.CTMCODYEAR_PRO = txtProCustomerCode.Text
   ItemPromotion.CTMNAMEYEAR_PRO = txtProCustomerName.Text
   ItemPromotion.STKCODYEAR_PRO = txtProSTKCOD.Text
   ItemPromotion.STKNAMEYEAR_PRO = txtProSTKName.Text
   ItemPromotion.AMOUNTYEAR_PRO = Val(txtProAmount.Text)
   ItemPromotion.YYYY_MM = Year(uctlProDate.ShowDate) & "-" & Format(Month(uctlProDate.ShowDate), "00")
   
   Call EnableForm(Me, False)
   If Not glbDaily.AddEditPromotionYear(ItemPromotion, IsOK, True, glbErrorLog) Then
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
      
      Call LoadCustomerPro(CustomerProColl)
      Call LoadStockPro(StockProColl)
      
      If ShowMode = SHOW_EDIT Then
         Call QueryData(True)
      ElseIf ShowMode = SHOW_ADD Then
         uctlProDate.ShowDate = Now
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
   
   Call InitNormalLabel(lblProDate, MapText("วันที่"))
   Call InitNormalLabel(lblProCustomerCode, MapText("ลูกค้า"))
   Call InitNormalLabel(lblProSTKCOD, MapText("สินค้า"))
   Call InitNormalLabel(lblProAmount, MapText("จำนวน"))
   
   Call txtProCustomerCode.SetTextLenType(TEXT_STRING, glbSetting.CODE_TYPE)
   Call txtProCustomerName.SetTextLenType(TEXT_STRING, glbSetting.NAME_TYPE)
   Call txtProSTKCOD.SetTextLenType(TEXT_STRING, glbSetting.CODE_TYPE)
   Call txtProSTKName.SetTextLenType(TEXT_STRING, glbSetting.NAME_TYPE)
   Call txtProAmount.SetTextLenType(TEXT_STRING, glbSetting.CODE_TYPE)
   
   txtProCustomerName.Enabled = False
   txtProSTKName.Enabled = False

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
      
   Set ItemPromotion = New CPromotionYear

   Set CustomerProColl = New Collection
   Set StockProColl = New Collection
   
   m_HasActivate = False
   Call InitFormLayout
   Call EnableForm(Me, True)
End Sub
Private Sub Form_Unload(Cancel As Integer)
   Set ItemPromotion = Nothing

   Set CustomerProColl = Nothing
   Set StockProColl = Nothing
End Sub

Private Sub txtProAmount_Change()
   m_HasModify = True
End Sub

Private Sub txtProCustomerCode_Change()
Dim TempCustomer As CARMas
   m_HasModify = True
   Set TempCustomer = GetObject("CARMas", CustomerProColl, txtProCustomerCode.Text, False)
   If Not TempCustomer Is Nothing Then
      txtProCustomerName.Text = TempCustomer.CUSNAM
   Else
      txtProCustomerName.Text = ""
   End If
   
   m_HasModify = True
End Sub

Private Sub txtProCustomerName_Change()
   m_HasModify = True
End Sub

Private Sub txtProSTKCOD_Change()
Dim tempStcrd As CStmas
   m_HasModify = True
   If Len(txtProSTKCOD.Text) > 0 Then
      Set tempStcrd = GetObject("CStcrd", StockProColl, Trim(txtProSTKCOD.Text), False)
      If Not tempStcrd Is Nothing Then
         txtProSTKName.Text = tempStcrd.STKDES
      Else
         txtProSTKName.Text = ""
      End If
   End If
End Sub

Private Sub txtProSTKName_Change()
   m_HasModify = True
End Sub

Private Sub uctlProDate_HasChange()
   m_HasModify = True
End Sub
