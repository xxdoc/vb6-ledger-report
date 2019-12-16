VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#2.0#0"; "THREED20.OCX"
Begin VB.Form frmAddEditCostProducts 
   BackColor       =   &H80000000&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   3615
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9420
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3615
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
         Top             =   1680
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
      Begin prjLedgerReport.uctlTextBox txtProductCode 
         Height          =   450
         Left            =   2400
         TabIndex        =   7
         Top             =   1080
         Width           =   1455
         _extentx        =   2566
         _extenty        =   794
      End
      Begin prjLedgerReport.uctlTextBox txtProductName 
         Height          =   450
         Left            =   3900
         TabIndex        =   8
         Top             =   1080
         Width           =   4695
         _extentx        =   8281
         _extenty        =   794
      End
      Begin VB.Label lblProduct 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   375
         Left            =   600
         TabIndex        =   6
         Top             =   1200
         Width           =   1575
      End
      Begin VB.Label lblProAmount 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   435
         Left            =   240
         TabIndex        =   3
         Top             =   1755
         Width           =   1935
      End
      Begin Threed.SSCommand cmdExit 
         Cancel          =   -1  'True
         Height          =   525
         Left            =   4560
         TabIndex        =   2
         Top             =   2640
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
         Top             =   2640
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         ButtonStyle     =   3
      End
   End
End
Attribute VB_Name = "frmAddEditCostProducts"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private m_HasActivate As Boolean
Private m_HasModify As Boolean
Private m_Rs As ADODB.Recordset
Private ItemCostProducts  As CCostProducts

Public ID As Long
Public OKClick As Boolean
Public ShowMode As SHOW_MODE_TYPE
Public HeaderText As String

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

      ItemCostProducts.PRODUCT_ID = ID
      ItemCostProducts.QueryFlag = 1
      If Not glbDaily.QueryCostProduct(ItemCostProducts, m_Rs, ItemCount, IsOK, glbErrorLog) Then
         glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
         Call EnableForm(Me, True)
         Exit Sub
      End If
   End If

   If ItemCount > 0 Then
      Call ItemCostProducts.PopulateFromRS(1, m_Rs)
      
      txtProductCode.Text = ItemCostProducts.PRODUCT_CODE
      txtProductName.Text = ItemCostProducts.PRODUCT_NAME
      txtProAmount.Text = ItemCostProducts.COST_PRODUCT

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
  

   If Not VerifyTextControl(lblProduct, txtProductCode, False) Then
      Exit Function
   End If
   If Not VerifyTextControl(lblProduct, txtProductName, False) Then
      Exit Function
   End If
   If Not VerifyTextControl(lblProAmount, txtProAmount, False) Then
      Exit Function
   End If
   
   If Not CheckUniqueNs(COST_PRODUCTS_UNIQUE, txtProductCode.Text, ID, , 2) Then
      glbErrorLog.LocalErrorMsg = MapText("มีข้อมูล") & " " & txtProductCode.Text & " " & MapText("อยู่ในระบบแล้ว")
      glbErrorLog.ShowUserError
      Exit Function
   End If

   If Not m_HasModify Then
      SaveData = True
      Exit Function
   End If

   ItemCostProducts.AddEditMode = ShowMode
   ItemCostProducts.PRODUCT_CODE = txtProductCode.Text
   ItemCostProducts.PRODUCT_NAME = txtProductName.Text
   ItemCostProducts.COST_PRODUCT = Val(txtProAmount.Text)
   
   Call EnableForm(Me, False)
   If Not glbDaily.AddEditCostProducts(ItemCostProducts, IsOK, True, glbErrorLog) Then
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
      
      Call LoadStockPro(StockProColl)
      
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
   
   Call InitNormalLabel(lblProduct, MapText("สินค้า"))
   Call InitNormalLabel(lblProAmount, MapText("จำนวน"))
   
   Call txtProductCode.SetTextLenType(TEXT_STRING, glbSetting.CODE_TYPE)
   Call txtProductName.SetTextLenType(TEXT_STRING, glbSetting.NAME_TYPE)
   Call txtProAmount.SetTextLenType(TEXT_STRING, glbSetting.CODE_TYPE)
   
   txtProductName.Enabled = False

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
      
   Set ItemCostProducts = New CCostProducts
   Set StockProColl = New Collection
   
   m_HasActivate = False
   Call InitFormLayout
   Call EnableForm(Me, True)
End Sub
Private Sub Form_Unload(Cancel As Integer)
   Set ItemCostProducts = Nothing
   Set StockProColl = Nothing
End Sub

Private Sub txtProAmount_Change()
   m_HasModify = True
End Sub

Private Sub txtProductCode_Change()
Dim tempStcrd As CStmas
   m_HasModify = True
   If Len(txtProductCode.Text) > 0 Then
      Set tempStcrd = GetObject("CStcrd", StockProColl, Trim(txtProductCode.Text), False)
      If Not tempStcrd Is Nothing Then
         txtProductName.Text = tempStcrd.STKDES
      Else
         txtProductName.Text = ""
      End If
   End If
End Sub

Private Sub txtProductName_Change()
   m_HasModify = True
End Sub
