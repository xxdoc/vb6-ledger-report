VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#2.0#0"; "THREED20.OCX"
Begin VB.Form frmDrCr 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   7440
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8475
   Icon            =   "frmDrCr.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7440
   ScaleWidth      =   8475
   StartUpPosition =   2  'CenterScreen
   Begin Threed.SSPanel pnlHeader 
      Height          =   705
      Left            =   30
      TabIndex        =   12
      Top             =   -30
      Width           =   8475
      _ExtentX        =   14949
      _ExtentY        =   1244
      _Version        =   131073
      Caption         =   "uctlGLACC"
      PictureBackgroundStyle=   2
   End
   Begin Threed.SSFrame SSFrame1 
      Height          =   7470
      Left            =   0
      TabIndex        =   13
      Top             =   0
      Width           =   8505
      _ExtentX        =   15002
      _ExtentY        =   13176
      _Version        =   131073
      PictureBackgroundStyle=   2
      Begin prjLedgerReport.uctlTextBox txtDrCr1 
         Height          =   435
         Left            =   1950
         TabIndex        =   1
         Top             =   1320
         Width           =   1485
         _ExtentX        =   3466
         _ExtentY        =   767
      End
      Begin prjLedgerReport.uctlTextLookup uctlAccLookup1 
         Height          =   435
         Left            =   1950
         TabIndex        =   0
         Top             =   900
         Width           =   5355
         _ExtentX        =   9446
         _ExtentY        =   767
      End
      Begin Threed.SSPanel SSPanel1 
         Height          =   705
         Left            =   0
         TabIndex        =   14
         Top             =   -30
         Width           =   8475
         _ExtentX        =   14949
         _ExtentY        =   1244
         _Version        =   131073
         PictureBackgroundStyle=   2
      End
      Begin prjLedgerReport.uctlTextBox txtDrCr2 
         Height          =   435
         Left            =   1950
         TabIndex        =   3
         Top             =   2190
         Width           =   1485
         _ExtentX        =   3466
         _ExtentY        =   767
      End
      Begin prjLedgerReport.uctlTextLookup uctlAccLookup2 
         Height          =   435
         Left            =   1950
         TabIndex        =   2
         Top             =   1770
         Width           =   5355
         _ExtentX        =   9446
         _ExtentY        =   767
      End
      Begin prjLedgerReport.uctlTextBox txtDrCr3 
         Height          =   435
         Left            =   1950
         TabIndex        =   5
         Top             =   3060
         Width           =   1485
         _ExtentX        =   3466
         _ExtentY        =   767
      End
      Begin prjLedgerReport.uctlTextLookup uctlAccLookup3 
         Height          =   435
         Left            =   1950
         TabIndex        =   4
         Top             =   2640
         Width           =   5355
         _ExtentX        =   9446
         _ExtentY        =   767
      End
      Begin prjLedgerReport.uctlTextBox txtDrCr4 
         Height          =   435
         Left            =   1950
         TabIndex        =   7
         Top             =   4260
         Width           =   1485
         _ExtentX        =   3466
         _ExtentY        =   767
      End
      Begin prjLedgerReport.uctlTextLookup uctlAccLookup4 
         Height          =   435
         Left            =   1950
         TabIndex        =   6
         Top             =   3840
         Width           =   5355
         _ExtentX        =   9446
         _ExtentY        =   767
      End
      Begin prjLedgerReport.uctlTextBox txtDrCr5 
         Height          =   435
         Left            =   1950
         TabIndex        =   9
         Top             =   5130
         Width           =   1485
         _ExtentX        =   3466
         _ExtentY        =   767
      End
      Begin prjLedgerReport.uctlTextLookup uctlAccLookup5 
         Height          =   435
         Left            =   1950
         TabIndex        =   8
         Top             =   4710
         Width           =   5355
         _ExtentX        =   9446
         _ExtentY        =   767
      End
      Begin prjLedgerReport.uctlTextBox txtDrCr6 
         Height          =   435
         Left            =   1950
         TabIndex        =   11
         Top             =   6000
         Width           =   1485
         _ExtentX        =   3466
         _ExtentY        =   767
      End
      Begin prjLedgerReport.uctlTextLookup uctlAccLookup6 
         Height          =   435
         Left            =   1950
         TabIndex        =   10
         Top             =   5580
         Width           =   5355
         _ExtentX        =   9446
         _ExtentY        =   767
      End
      Begin VB.Label lblAccount4 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   315
         Left            =   90
         TabIndex        =   34
         Top             =   3930
         Width           =   1755
      End
      Begin VB.Label lblDrCrAmount4 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   495
         Left            =   90
         TabIndex        =   33
         Top             =   4380
         Width           =   1755
      End
      Begin VB.Label lblBath4 
         Caption         =   "Label1"
         Height          =   375
         Left            =   3450
         TabIndex        =   32
         Top             =   4350
         Width           =   1065
      End
      Begin VB.Label lblBath5 
         Caption         =   "Label1"
         Height          =   375
         Left            =   3450
         TabIndex        =   31
         Top             =   5220
         Width           =   1065
      End
      Begin VB.Label lblDrCrAmount5 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   495
         Left            =   90
         TabIndex        =   30
         Top             =   5250
         Width           =   1755
      End
      Begin VB.Label lblAccount5 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   315
         Left            =   90
         TabIndex        =   29
         Top             =   4800
         Width           =   1755
      End
      Begin VB.Label lblBath6 
         Caption         =   "Label1"
         Height          =   375
         Left            =   3450
         TabIndex        =   28
         Top             =   6090
         Width           =   1065
      End
      Begin VB.Label lblDrCrAmount6 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   495
         Left            =   90
         TabIndex        =   27
         Top             =   6120
         Width           =   1755
      End
      Begin VB.Label lblAccount6 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   315
         Left            =   90
         TabIndex        =   26
         Top             =   5670
         Width           =   1755
      End
      Begin VB.Label lblAccount3 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   315
         Left            =   90
         TabIndex        =   25
         Top             =   2730
         Width           =   1755
      End
      Begin VB.Label lblDrCrAmount3 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   495
         Left            =   90
         TabIndex        =   24
         Top             =   3180
         Width           =   1755
      End
      Begin VB.Label lblBath3 
         Caption         =   "Label1"
         Height          =   375
         Left            =   3450
         TabIndex        =   23
         Top             =   3120
         Width           =   1065
      End
      Begin VB.Label lblAccount2 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   315
         Left            =   90
         TabIndex        =   22
         Top             =   1860
         Width           =   1755
      End
      Begin VB.Label lblDrCrAmount2 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   495
         Left            =   90
         TabIndex        =   21
         Top             =   2310
         Width           =   1755
      End
      Begin VB.Label lblBath2 
         Caption         =   "Label1"
         Height          =   375
         Left            =   3450
         TabIndex        =   20
         Top             =   2250
         Width           =   1065
      End
      Begin VB.Label lblBath1 
         Caption         =   "Label1"
         Height          =   375
         Left            =   3450
         TabIndex        =   19
         Top             =   1380
         Width           =   1065
      End
      Begin VB.Label lblDrCrAmount1 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   495
         Left            =   90
         TabIndex        =   18
         Top             =   1440
         Width           =   1755
      End
      Begin VB.Label lblAccount1 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   315
         Left            =   90
         TabIndex        =   17
         Top             =   990
         Width           =   1755
      End
      Begin Threed.SSCommand cmdExit 
         Cancel          =   -1  'True
         Height          =   525
         Left            =   4252
         TabIndex        =   16
         Top             =   6660
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   926
         _Version        =   131073
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdOK 
         Height          =   525
         Left            =   2617
         TabIndex        =   15
         Top             =   6660
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         ButtonStyle     =   3
      End
   End
End
Attribute VB_Name = "frmDrCr"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const ROOT_TREE = "Root"

Private MustAsk As Boolean
Private m_HasActivate As Boolean
Private m_Rs  As ADODB.Recordset
Private m_TableName As String
Public ShowMode As SHOW_MODE_TYPE
Public HeaderText As String
Private m_MustAsk As Boolean

Private m_ReportControls As Collection
Private m_Texts As Collection
Private m_TextLookups As Collection
Private m_TextMoney As Collection
Private m_Dates As Collection
Private m_CheckBoxes As Collection
Private m_Labels As Collection
Private m_Labels2 As Collection
Private m_Combos As Collection

Private m_Glacc As Collection
Private C As CReportControl

Private m_ReportParams As Collection
Private m_FromDate As Date
Private m_ToDate As Date
Private m_DBPath As String

Private m_AccDrLookup(100) As uctlTextLookup
Private m_AccDrText(100) As uctlTextBox
Private m_AccCrLookup(100) As uctlTextLookup
Private m_AccCrText(100) As uctlTextBox
Private m_HasModify As Boolean

Public TempCollection As Collection
Public OKClick As Boolean

Private Sub InitFormLayout()
   pnlHeader.PictureBackground = LoadPicture(glbParameterObj.NormalForm1)
   SSFrame1.PictureBackground = LoadPicture(glbParameterObj.NormalForm1)
   
   Me.Caption = "กำหนดข้อมูลสมุดรายวัน"
   pnlHeader.Caption = Me.Caption
   
   Me.Picture = LoadPicture(glbParameterObj.MainPicture)
   Me.Caption = MapText("กำหนดข้อมูลสมุดรายวัน")

   pnlHeader.Font.Name = GLB_FONT
   pnlHeader.Font.Bold = True
   pnlHeader.Font.Size = 19

   cmdOK.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdExit.Picture = LoadPicture(glbParameterObj.NormalButton1)

   Call InitNormalLabel(lblAccount1, MapText("เลขที่บัญชี Dr 1"))
   Call InitNormalLabel(lblDrCrAmount1, MapText("จำนวน"))
   Call InitNormalLabel(lblBath1, MapText("บาท"))
   Call InitNormalLabel(lblAccount2, MapText("เลขที่บัญชี Dr 2"))
   Call InitNormalLabel(lblDrCrAmount2, MapText("จำนวน"))
   Call InitNormalLabel(lblBath2, MapText("บาท"))
   Call InitNormalLabel(lblAccount3, MapText("เลขที่บัญชี Dr 3"))
   Call InitNormalLabel(lblDrCrAmount3, MapText("จำนวน"))
   Call InitNormalLabel(lblBath3, MapText("บาท"))
   
   Call InitNormalLabel(lblAccount4, MapText("เลขที่บัญชี Cr 1"))
   Call InitNormalLabel(lblDrCrAmount4, MapText("จำนวน"))
   Call InitNormalLabel(lblBath4, MapText("บาท"))
   Call InitNormalLabel(lblAccount5, MapText("เลขที่บัญชี Cr 2"))
   Call InitNormalLabel(lblDrCrAmount5, MapText("จำนวน"))
   Call InitNormalLabel(lblBath5, MapText("บาท"))
   Call InitNormalLabel(lblAccount6, MapText("เลขที่บัญชี Cr 3"))
   Call InitNormalLabel(lblDrCrAmount6, MapText("จำนวน"))
   Call InitNormalLabel(lblBath6, MapText("บาท"))
   
   Call InitMainButton(cmdOK, MapText("พิมพ์ (F10)"))
   Call InitMainButton(cmdExit, MapText("ออก"))
   
   Call txtDrCr1.SetTextLenType(TEXT_FLOAT_MONEY, glbSetting.MONEY_TYPE)
   Call txtDrCr2.SetTextLenType(TEXT_FLOAT_MONEY, glbSetting.MONEY_TYPE)
   Call txtDrCr3.SetTextLenType(TEXT_FLOAT_MONEY, glbSetting.MONEY_TYPE)
   Call txtDrCr4.SetTextLenType(TEXT_FLOAT_MONEY, glbSetting.MONEY_TYPE)
   Call txtDrCr5.SetTextLenType(TEXT_FLOAT_MONEY, glbSetting.MONEY_TYPE)
   Call txtDrCr6.SetTextLenType(TEXT_FLOAT_MONEY, glbSetting.MONEY_TYPE)
   
   cmdExit.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdOK.Picture = LoadPicture(glbParameterObj.NormalButton1)
   
   Set m_AccDrLookup(0) = uctlAccLookup1
   Set m_AccDrText(0) = txtDrCr1
   Set m_AccDrLookup(1) = uctlAccLookup2
   Set m_AccDrText(1) = txtDrCr2
   Set m_AccDrLookup(2) = uctlAccLookup3
   Set m_AccDrText(2) = txtDrCr3

   Set m_AccCrLookup(3) = uctlAccLookup4
   Set m_AccCrText(3) = txtDrCr4
   Set m_AccCrLookup(4) = uctlAccLookup5
   Set m_AccCrText(4) = txtDrCr5
   Set m_AccCrLookup(5) = uctlAccLookup6
   Set m_AccCrText(5) = txtDrCr6

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

Private Sub cmdOK_Click()
Dim i As Long
Dim j As CGLJnl
Dim Sum1 As Double
Dim Sum2 As Double
   
   Sum1 = 0
   Sum2 = 0
   For i = 0 To 2
      If m_AccDrLookup(i).MyCombo.ListIndex > 0 Then
         Set j = New CGLJnl
         j.ACCNUM = m_AccDrLookup(i).MyTextBox.Text
         j.ACCNAM = m_AccDrLookup(i).MyCombo.Text
         j.AMOUNT = Val(m_AccDrText(i).Text)
         j.TRNTYP = 0
         Sum1 = Sum1 + j.AMOUNT
         Call TempCollection.Add(j)
         Set j = Nothing
      End If
   Next i

   For i = 3 To 5
      If m_AccCrLookup(i).MyCombo.ListIndex > 0 Then
         Set j = New CGLJnl
         j.ACCNUM = m_AccCrLookup(i).MyTextBox.Text
         j.ACCNAM = m_AccCrLookup(i).MyCombo.Text
         j.AMOUNT = Val(m_AccCrText(i).Text)
         j.TRNTYP = 1
         Sum2 = Sum2 + j.AMOUNT
         Call TempCollection.Add(j)
         Set j = Nothing
      End If
   Next i

   If CStr(Sum1) <> CStr(Sum2) Then
      glbErrorLog.LocalErrorMsg = "กรุณากรอกให้ผลรวมของ DEBIT = CREDIT"
      glbErrorLog.ShowUserError
      Exit Sub
   End If

   OKClick = True
   Unload Me
End Sub

Private Sub Form_Activate()
   If Not m_HasActivate Then
      Me.Refresh
      DoEvents
      
      Call LoadGLAcc(uctlAccLookup1.MyCombo, m_Glacc)
      Set uctlAccLookup1.MyCollection = m_Glacc
      
      Call LoadGLAcc(uctlAccLookup2.MyCombo, m_Glacc)
      Set uctlAccLookup2.MyCollection = m_Glacc
      
      Set uctlAccLookup3.MyCollection = m_Glacc
      Call LoadGLAcc(uctlAccLookup3.MyCombo, m_Glacc)
   
      Set uctlAccLookup4.MyCollection = m_Glacc
      Call LoadGLAcc(uctlAccLookup4.MyCombo, m_Glacc)
   
      Set uctlAccLookup5.MyCollection = m_Glacc
      Call LoadGLAcc(uctlAccLookup5.MyCombo, m_Glacc)
   
      Set uctlAccLookup6.MyCollection = m_Glacc
      Call LoadGLAcc(uctlAccLookup6.MyCombo, m_Glacc)
   End If
   
   m_HasActivate = True
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
   m_HasModify = False
    Set m_TextLookups = New Collection
   m_HasActivate = False
    Set m_TextMoney = New Collection
   Set m_Labels2 = New Collection
   Set m_Labels = New Collection
   Set m_ReportControls = New Collection
   Set m_Glacc = New Collection
   Call InitFormLayout
End Sub

Private Sub txtDrCr1_Change()
   m_HasModify = True
End Sub

Private Sub txtDrCr2_Change()
   m_HasModify = True
End Sub

Private Sub txtDrCr3_Change()
   m_HasModify = True
End Sub

Private Sub txtDrCr4_Change()
   m_HasModify = True
End Sub

Private Sub txtDrCr5_Change()
   m_HasModify = True
End Sub

Private Sub txtDrCr6_Change()
   m_HasModify = True
End Sub

Private Sub uctlAccLookup1_Change()
   m_HasModify = True
End Sub

Private Sub uctlAccLookup2_Change()
   m_HasModify = True
End Sub

Private Sub uctlAccLookup3_Change()
   m_HasModify = True
End Sub

Private Sub uctlAccLookup4_Change()
   m_HasModify = True
End Sub

Private Sub uctlAccLookup5_Change()
   m_HasModify = True
End Sub

Private Sub uctlAccLookup6_Change()
   m_HasModify = True
End Sub
