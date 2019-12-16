VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#2.0#0"; "THREED20.OCX"
Begin VB.Form frmAddEditMasterFTEx 
   BackColor       =   &H80000000&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   4455
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8100
   BeginProperty Font 
      Name            =   "AngsanaUPC"
      Size            =   14.25
      Charset         =   222
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmMaster2FromTo2.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4455
   ScaleWidth      =   8100
   StartUpPosition =   2  'CenterScreen
   Begin Threed.SSPanel pnlHeader 
      Height          =   615
      Left            =   0
      TabIndex        =   7
      Top             =   0
      Width           =   8235
      _ExtentX        =   14526
      _ExtentY        =   1085
      _Version        =   131073
      PictureBackgroundStyle=   2
   End
   Begin Threed.SSFrame SSFrame1 
      Height          =   3975
      Left            =   0
      TabIndex        =   8
      Top             =   600
      Width           =   8265
      _ExtentX        =   14579
      _ExtentY        =   7011
      _Version        =   131073
      PictureBackgroundStyle=   2
      Begin Threed.SSFrame SSFrame2 
         Height          =   975
         Left            =   0
         TabIndex        =   10
         Top             =   120
         Width           =   8175
         _ExtentX        =   14420
         _ExtentY        =   1720
         _Version        =   131073
         Caption         =   "กรุณาเลือกประเภท"
         Begin Threed.SSOption SSOption3 
            Height          =   375
            Left            =   5280
            TabIndex        =   2
            Top             =   360
            Width           =   2175
            _ExtentX        =   3836
            _ExtentY        =   661
            _Version        =   131073
            Caption         =   "SSOption1"
         End
         Begin Threed.SSOption SSOption2 
            Height          =   375
            Left            =   2880
            TabIndex        =   1
            Top             =   360
            Width           =   2175
            _ExtentX        =   3836
            _ExtentY        =   661
            _Version        =   131073
            Caption         =   "SSOption1"
         End
         Begin Threed.SSOption SSOption1 
            Height          =   375
            Left            =   720
            TabIndex        =   0
            Top             =   360
            Width           =   2175
            _ExtentX        =   3836
            _ExtentY        =   661
            _Version        =   131073
            Caption         =   "SSOption1"
         End
      End
      Begin VB.PictureBox txtPercent 
         Height          =   435
         Left            =   2160
         ScaleHeight     =   375
         ScaleWidth      =   1635
         TabIndex        =   3
         Top             =   2640
         Width           =   1695
      End
      Begin VB.PictureBox uctlProductLookup 
         Height          =   435
         Left            =   2160
         ScaleHeight     =   375
         ScaleWidth      =   5685
         TabIndex        =   14
         Top             =   1680
         Width           =   5745
      End
      Begin VB.PictureBox uctlAparMasLookup 
         Height          =   435
         Left            =   2160
         ScaleHeight     =   375
         ScaleWidth      =   5685
         TabIndex        =   15
         Top             =   1200
         Width           =   5745
      End
      Begin VB.PictureBox uctlSale 
         Height          =   435
         Left            =   2160
         ScaleHeight     =   375
         ScaleWidth      =   5685
         TabIndex        =   16
         Top             =   2160
         Width           =   5745
      End
      Begin VB.Label lblSale 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   240
         TabIndex        =   13
         Top             =   2220
         Width           =   1755
      End
      Begin VB.Label lblProduct 
         Alignment       =   1  'Right Justify
         Height          =   375
         Left            =   240
         TabIndex        =   12
         Top             =   1800
         Width           =   1725
      End
      Begin VB.Label lblAparMas 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   240
         TabIndex        =   11
         Top             =   1260
         Width           =   1755
      End
      Begin VB.Label lblPercent 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   240
         TabIndex        =   9
         Top             =   2760
         Width           =   1755
      End
      Begin Threed.SSCommand cmdNext 
         Height          =   525
         Left            =   1245
         TabIndex        =   4
         Top             =   3150
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmMaster2FromTo2.frx":08CA
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdOK 
         Height          =   525
         Left            =   2895
         TabIndex        =   5
         Top             =   3150
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmMaster2FromTo2.frx":0BE4
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdExit 
         Cancel          =   -1  'True
         Height          =   525
         Left            =   4545
         TabIndex        =   6
         Top             =   3150
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   926
         _Version        =   131073
         ButtonStyle     =   3
      End
   End
End
Attribute VB_Name = "frmAddEditMasterFTEx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Public Header As String
Public ShowMode As SHOW_MODE_TYPE
Private m_HasActivate As Boolean
Private m_HasModify As Boolean
Private m_Rs As ADODB.Recordset

Public HeaderText As String
Public ID As Long
Public OKClick As Boolean
Public TempCollection As Collection
Public ParentForm As Form

Private m_Products As Collection
Private Sub cmdExit_Click()
   If Not ConfirmExit(m_HasModify) Then
      Exit Sub
   End If
   
   OKClick = False
   Unload Me
End Sub

Private Sub InitFormLayout()
   Me.Picture = LoadPicture(glbParameterObj.MainPicture)
   pnlHeader.PictureBackground = LoadPicture(glbParameterObj.NormalForm1)
   SSFrame1.PictureBackground = LoadPicture(glbParameterObj.NormalForm1)

   Me.KeyPreview = True
   pnlHeader.Caption = HeaderText
   Me.BackColor = GLB_FORM_COLOR
   pnlHeader.BackColor = GLB_HEAD_COLOR
   SSFrame1.BackColor = GLB_FORM_COLOR
   
   pnlHeader.Font.Name = GLB_FONT
   pnlHeader.Font.Bold = True
   pnlHeader.Font.Size = 19
   pnlHeader.Caption = HeaderText
   
   Call InitNormalLabel(lblAparMas, MapText("ลูกค้า"))
   Call InitNormalLabel(lblProduct, MapText("สินค้า/วัตถุดิบ"))
   Call InitNormalLabel(lblSale, MapText("พนักงานขาย"))
   Call InitNormalLabel(lblPercent, MapText("%ส่วนแบ่ง"))
   
  ' txtPercent.TEXT_FLOAT_MONEY, glbSetting.MONEY_TYPE)
   
   Call InitOptionEx(SSOption1, "ลูกค้า")
   Call InitOptionEx(SSOption2, "สินค้า")
   Call InitOptionEx(SSOption3, "พนักงานขาย")
   
   SSOption1.Value = True
   SSOption2.Value = False
   SSOption3.Value = False
   
   cmdExit.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdOK.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdNext.Picture = LoadPicture(glbParameterObj.NormalButton1)
   
   Call InitMainButton(cmdOK, MapText("ตกลง (F2)"))
   Call InitMainButton(cmdExit, MapText("ยกเลิก (ESC)"))
   Call InitMainButton(cmdNext, MapText("ถัดไป"))
End Sub

Private Sub QueryData(Flag As Boolean)
Dim IsOK As Boolean
Dim ItemCount As Long

   If Flag Then
      Call EnableForm(Me, False)
      
      If ShowMode = SHOW_EDIT Then
         Dim BD As CMasterFromToEx
         
         Set BD = TempCollection.ITEM(ID)
         
         uctlAparMasLookup.MyCombo.ListIndex = IDToListIndex(uctlAparMasLookup.MyCombo, BD.CUSTOMER_ID)
         uctlProductLookup.MyCombo.ListIndex = IDToListIndex(uctlProductLookup.MyCombo, BD.PART_ITEM_ID)
         uctlSale.MyCombo.ListIndex = IDToListIndex(uctlSale.MyCombo, BD.EMP_ID)
         txtPercent.Text = BD.PERCENT
         
         If BD.CUSTOMER_ID > 0 Then
            SSOption1.Value = True
         End If
         If BD.PART_ITEM_ID > 0 Then
            SSOption2.Value = True
         End If
         If BD.EMP_ID > 0 Then
            SSOption3.Value = True
         End If
         
      End If
   End If
   
   Call EnableForm(Me, True)
End Sub

Private Sub cmdNext_Click()
Dim NewID As Long

   If Not SaveData Then
      Exit Sub
   End If
   
   If ShowMode = SHOW_EDIT Then
      NewID = GetNextID(ID, TempCollection)
      If ID = NewID Then
         glbErrorLog.LocalErrorMsg = "ถึงเรคคอร์ดสุดท้ายแล้ว"
         glbErrorLog.ShowUserError
         
         Call ParentForm.RefreshGrid
         Exit Sub
      End If
      
      ID = NewID
   ElseIf ShowMode = SHOW_ADD Then
      uctlAparMasLookup.MyCombo.ListIndex = -1
      uctlProductLookup.MyCombo.ListIndex = -1
      uctlSale.MyCombo.ListIndex = -1
      txtPercent.Text = ""
   End If
   Call QueryData(True)
   
   Call ParentForm.RefreshGrid
End Sub

Private Sub cmdOK_Click()
   If Not SaveData Then
      Exit Sub
   End If
   
   OKClick = True
   Unload Me
End Sub
Private Function SaveData() As Boolean
Dim IsOK As Boolean
Dim RealIndex As Long

   If Not m_HasModify Then
      SaveData = True
      Exit Function
   End If
   
   Dim BD As CMasterFromToEx
   If ShowMode = SHOW_ADD Then
      Set BD = New CMasterFromToEx
      BD.Flag = "A"
      Call TempCollection.Add(BD)
   Else
      Set BD = TempCollection.ITEM(ID)
      If BD.Flag <> "A" Then
         BD.Flag = "E"
      End If
   End If
   
   BD.CUSTOMER_ID = uctlAparMasLookup.MyCombo.ItemData(Minus2Zero(uctlAparMasLookup.MyCombo.ListIndex))
   BD.CUSTOMER_CODE = uctlAparMasLookup.MyTextBox.Text
   BD.CUSTOMER_NAME = uctlAparMasLookup.MyCombo.Text
   BD.PART_ITEM_ID = uctlProductLookup.MyCombo.ItemData(Minus2Zero(uctlProductLookup.MyCombo.ListIndex))
   BD.PART_NO = uctlProductLookup.MyTextBox.Text
   BD.PART_DESC = uctlProductLookup.MyCombo.Text
   BD.EMP_ID = uctlSale.MyCombo.ItemData(Minus2Zero(uctlSale.MyCombo.ListIndex))
   BD.SALE_CODE = uctlSale.MyTextBox.Text
   BD.SALE_NAME = uctlSale.MyCombo.Text
   
   BD.PERCENT = Val(txtPercent.Text)
   
   SaveData = True
End Function

Private Sub Form_Activate()
   If Not m_HasActivate Then
      m_HasActivate = True
      Me.Refresh
      DoEvents
      
      Call LoadApArMas(, uctlAparMasLookup.MyCombo)
      Set uctlAparMasLookup.MyCollection = m_CustomerColl
      
      Call LoadStockCode(uctlProductLookup.MyCombo, m_Products)
      Set uctlProductLookup.MyCollection = m_Products
      
      Call LoadEmployee(, uctlSale.MyCombo)
      Set uctlSale.MyCollection = m_EmployeeColl
      
      If ShowMode = SHOW_EDIT Then
         Call QueryData(True)
      ElseIf ShowMode = SHOW_ADD Then
         ID = 0
         Call QueryData(True)
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

Private Sub Form_Load()

   OKClick = False
   Call InitFormLayout
   
   m_HasActivate = False
   m_HasModify = False
   Set m_Products = New Collection
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set m_Products = Nothing
   
End Sub

Private Sub SSOption1_Click(Value As Integer)
   If SSOption1.Value = True Then
      uctlAparMasLookup.Enabled = True
      uctlProductLookup.Enabled = False
      uctlSale.Enabled = False
      txtPercent.Enabled = False
   ElseIf SSOption2.Value = True Then
      uctlAparMasLookup.Enabled = False
      uctlProductLookup.Enabled = True
      uctlSale.Enabled = False
      txtPercent.Enabled = False
   ElseIf SSOption3.Value = True Then
      uctlProductLookup.Enabled = False
      uctlAparMasLookup.Enabled = False
      uctlSale.Enabled = True
      txtPercent.Enabled = True
   End If
   m_HasModify = True
End Sub

Private Sub SSOption1_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      SendKeys ("{TAB}")
   End If
End Sub

Private Sub SSOption2_Click(Value As Integer)
   
   If SSOption1.Value = True Then
      uctlAparMasLookup.Enabled = True
      uctlProductLookup.Enabled = False
      uctlSale.Enabled = False
      txtPercent.Enabled = False
   ElseIf SSOption2.Value = True Then
      uctlAparMasLookup.Enabled = False
      uctlProductLookup.Enabled = True
      uctlSale.Enabled = False
      txtPercent.Enabled = False
   ElseIf SSOption3.Value = True Then
      uctlProductLookup.Enabled = False
      uctlAparMasLookup.Enabled = False
      uctlSale.Enabled = True
      txtPercent.Enabled = True
   End If
   m_HasModify = True
End Sub
Private Sub SSOption2_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      SendKeys ("{TAB}")
   End If
End Sub
Private Sub SSOption3_Click(Value As Integer)
   If SSOption1.Value = True Then
      uctlAparMasLookup.Enabled = True
      uctlProductLookup.Enabled = False
      uctlSale.Enabled = False
      txtPercent.Enabled = False
   ElseIf SSOption2.Value = True Then
      uctlAparMasLookup.Enabled = False
      uctlProductLookup.Enabled = True
      uctlSale.Enabled = False
      txtPercent.Enabled = False
   ElseIf SSOption3.Value = True Then
      uctlProductLookup.Enabled = False
      uctlAparMasLookup.Enabled = False
      uctlSale.Enabled = True
      txtPercent.Enabled = True
   End If
   m_HasModify = True
End Sub
Private Sub SSOption3_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      SendKeys ("{TAB}")
   End If
End Sub

Private Sub txtPercent_Change()
   m_HasModify = True
End Sub

Private Sub uctlApArMasLookup_Change()
   m_HasModify = True
End Sub
Private Sub uctlProductLookup_Change()
   m_HasModify = True
End Sub
Private Sub uctlSale_Change()
   m_HasModify = True
End Sub
