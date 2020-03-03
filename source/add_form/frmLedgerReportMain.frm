VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#2.0#0"; "THREED20.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmLedgerReportMain 
   ClientHeight    =   9495
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   13140
   Icon            =   "frmLedgerReportMain.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   9495
   ScaleWidth      =   13140
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin Threed.SSFrame SSFrame2 
      Height          =   9495
      Left            =   -120
      TabIndex        =   0
      Top             =   0
      Width           =   14775
      _ExtentX        =   26061
      _ExtentY        =   16748
      _Version        =   131073
      PictureBackgroundStyle=   2
      Begin Threed.SSFrame SSFrame3 
         Height          =   8055
         Left            =   6600
         TabIndex        =   9
         Top             =   720
         Width           =   6495
         _ExtentX        =   11456
         _ExtentY        =   14208
         _Version        =   131073
         PictureBackgroundStyle=   2
         Begin VB.ComboBox cboGeneric 
            BeginProperty Font 
               Name            =   "AngsanaUPC"
               Size            =   9
               Charset         =   222
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Index           =   0
            Left            =   2550
            Style           =   2  'Dropdown List
            TabIndex        =   10
            Top             =   540
            Visible         =   0   'False
            Width           =   3855
         End
         Begin prjLedgerReport.uctlTextBox txtGeneric 
            Height          =   435
            Index           =   0
            Left            =   2550
            TabIndex        =   11
            Top             =   930
            Visible         =   0   'False
            Width           =   3855
            _ExtentX        =   6800
            _ExtentY        =   767
         End
         Begin prjLedgerReport.uctlDate uctlGenericDate 
            Height          =   405
            Index           =   0
            Left            =   2550
            TabIndex        =   12
            Top             =   120
            Visible         =   0   'False
            Width           =   3825
            _ExtentX        =   5689
            _ExtentY        =   291
         End
         Begin Threed.SSCommand cmdAdd 
            Height          =   435
            Left            =   2520
            TabIndex        =   15
            Top             =   1860
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   767
            _Version        =   131073
            MousePointer    =   99
            MouseIcon       =   "frmLedgerReportMain.frx":08CA
            ButtonStyle     =   3
         End
         Begin VB.Label lblGeneric 
            Alignment       =   1  'Right Justify
            Height          =   375
            Index           =   0
            Left            =   120
            TabIndex        =   14
            Top             =   120
            Visible         =   0   'False
            Width           =   2355
         End
         Begin Threed.SSCheck chkGeneric 
            Height          =   465
            Index           =   0
            Left            =   2550
            TabIndex        =   13
            Top             =   1410
            Visible         =   0   'False
            Width           =   3765
            _ExtentX        =   6641
            _ExtentY        =   820
            _Version        =   131073
            Caption         =   "SSCheck1"
         End
      End
      Begin VB.Timer Timer1 
         Interval        =   1000
         Left            =   7920
         Top             =   8880
      End
      Begin VB.PictureBox Picture1 
         Height          =   765
         Left            =   8520
         ScaleHeight     =   705
         ScaleWidth      =   825
         TabIndex        =   1
         Top             =   8880
         Visible         =   0   'False
         Width           =   885
      End
      Begin Threed.SSPanel SSPanel1 
         Height          =   795
         Left            =   120
         TabIndex        =   5
         Top             =   0
         Width           =   6495
         _ExtentX        =   11456
         _ExtentY        =   1402
         _Version        =   131073
         BackStyle       =   1
         Begin Threed.SSCommand SSCommand1 
            Height          =   555
            Left            =   9660
            TabIndex        =   7
            Top             =   6390
            Width           =   1845
            _ExtentX        =   3254
            _ExtentY        =   979
            _Version        =   131073
            PictureFrames   =   1
            Picture         =   "frmLedgerReportMain.frx":0BE4
            Caption         =   "SSCommand1"
            ButtonStyle     =   3
         End
         Begin VB.Label lblDateTime 
            Alignment       =   2  'Center
            Caption         =   "Label1"
            Height          =   765
            Left            =   30
            TabIndex        =   6
            Top             =   0
            Width           =   6525
         End
      End
      Begin Threed.SSPanel pnlHeader 
         Height          =   735
         Left            =   6600
         TabIndex        =   8
         Top             =   -30
         Width           =   6645
         _ExtentX        =   11721
         _ExtentY        =   1296
         _Version        =   131073
         BackStyle       =   1
      End
      Begin MSComctlLib.TreeView trvMain 
         Height          =   7875
         Left            =   0
         TabIndex        =   18
         Top             =   800
         Width           =   6525
         _ExtentX        =   11509
         _ExtentY        =   13891
         _Version        =   393217
         Indentation     =   882
         LabelEdit       =   1
         Style           =   7
         ImageList       =   "ImageList1"
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "JasmineUPC"
            Size            =   14.25
            Charset         =   222
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSComctlLib.ImageList ImageList1 
         Left            =   5640
         Top             =   8760
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   16
         ImageHeight     =   16
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   12
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmLedgerReportMain.frx":1C74
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmLedgerReportMain.frx":254E
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmLedgerReportMain.frx":2E28
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmLedgerReportMain.frx":3702
               Key             =   ""
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmLedgerReportMain.frx":385C
               Key             =   ""
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmLedgerReportMain.frx":4136
               Key             =   ""
            EndProperty
            BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmLedgerReportMain.frx":4A10
               Key             =   ""
            EndProperty
            BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmLedgerReportMain.frx":4D2A
               Key             =   ""
            EndProperty
            BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmLedgerReportMain.frx":5604
               Key             =   ""
            EndProperty
            BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmLedgerReportMain.frx":5EDE
               Key             =   ""
            EndProperty
            BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmLedgerReportMain.frx":67B8
               Key             =   ""
            EndProperty
            BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmLedgerReportMain.frx":7492
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin Threed.SSCommand cmdPasswd2 
         Height          =   525
         Left            =   1680
         TabIndex        =   20
         Top             =   8880
         Width           =   1425
         _ExtentX        =   2514
         _ExtentY        =   926
         _Version        =   131073
         ButtonStyle     =   3
      End
      Begin VB.Label lblVersion 
         Caption         =   "Label1"
         Height          =   345
         Left            =   4800
         TabIndex        =   19
         Top             =   9000
         Width           =   3045
      End
      Begin Threed.SSCommand cmdExit 
         Height          =   525
         Left            =   3240
         TabIndex        =   17
         Top             =   8850
         Width           =   1425
         _ExtentX        =   2514
         _ExtentY        =   926
         _Version        =   131073
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdPasswd 
         Height          =   525
         Left            =   120
         TabIndex        =   16
         Top             =   8880
         Width           =   1425
         _ExtentX        =   2514
         _ExtentY        =   926
         _Version        =   131073
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdOK 
         Height          =   525
         Left            =   11460
         TabIndex        =   4
         Top             =   8850
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdConfig 
         Height          =   525
         Left            =   9750
         TabIndex        =   3
         Top             =   8850
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand SSCommand2 
         Height          =   525
         Left            =   6720
         TabIndex        =   2
         Top             =   8880
         Visible         =   0   'False
         Width           =   1875
         _ExtentX        =   3307
         _ExtentY        =   926
         _Version        =   131073
         Caption         =   "SSCommand2"
      End
   End
End
Attribute VB_Name = "frmLedgerReportMain"
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

Public HeaderText As String
Private m_MustAsk As Boolean
Private m_DrCrFlage As Boolean

Private m_ReportControls As Collection
Private m_Texts As Collection
Private m_TextLookups As Collection
Private m_Dates As Collection
Private m_CheckBoxes As Collection
Private m_Labels As Collection
Private m_Combos As Collection
Private m_ReportParams As Collection
Private m_FromDate As Date
Private m_ToDate As Date
Private m_DBPath As String
Private m_Journals As Collection      'Step 1

Private Sub cmdAdd_Click()
Dim OKClick As Boolean

  frmDrCr.ShowMode = SHOW_EDIT
   Set m_Journals = Nothing
   Set m_Journals = New Collection

   Set frmDrCr.TempCollection = m_Journals
   Load frmDrCr
   frmDrCr.Show 1
   OKClick = frmDrCr.OKClick
   Unload frmDrCr
   Set frmDatabaseSelect = Nothing

   If OKClick Then
   End If
End Sub

Private Sub cmdConfig_Click()
Dim ReportKey As String
Dim Rc As CReportConfig
Dim iCount As Long

   If trvMain.SelectedItem Is Nothing Then
      Exit Sub
   End If

   ReportKey = trvMain.SelectedItem.KEY
   
   Set Rc = New CReportConfig
   Rc.REPORT_KEY = ReportKey
   Rc.COMPUTER_NAME = glbParameterObj.ComputerName
   Call Rc.QueryData(m_Rs, iCount)
   
   If Not m_Rs.EOF Then
      Call Rc.PopulateFromRS(1, m_Rs)
      
      frmReportConfig.ShowMode = SHOW_EDIT
      frmReportConfig.ID = Rc.REPORT_CONFIG_ID
   Else
      frmReportConfig.ShowMode = SHOW_ADD
   End If

   frmReportConfig.ReportKey = ReportKey
   frmReportConfig.HeaderText = trvMain.SelectedItem.Text
   Load frmReportConfig
   frmReportConfig.Show 1
   
   Unload frmReportConfig
   Set frmReportConfig = Nothing
   
   Set Rc = Nothing
End Sub

Private Sub cmdExit_Click()
   Unload Me
End Sub

Private Sub cmdOK_Click()
Dim Report As CReportInterface
Dim SelectFlag As Boolean
Dim KEY As String
Dim Name As String
Dim C As CReportControl
Dim m_AP003 As CReportAP003

   KEY = trvMain.SelectedItem.KEY
   Name = trvMain.SelectedItem.Text
      
   SelectFlag = False
   
   If Not VerifyReportInput Then
      Exit Sub
   End If
   
   Set Report = New CReportInterface
   
   If Not (trvMain.SelectedItem Is Nothing) Then
      Call Report.AddParam(trvMain.SelectedItem.Text, "REPORT_TEXT")
   End If
   
   
   If KEY = ROOT_TREE & " 1-0-1" Then
'      If Not VerifyAccessRight("AP_REPORT_PRINT1", trvMaster.SelectedItem.Text) Then
'         Call EnableForm(Me, True)
'         Exit Sub
'      End If
      
      Set Report = New CReportArMas01
      SelectFlag = True
   ElseIf KEY = ROOT_TREE & " 1-0-1-1" Then
'      If Not VerifyAccessRight("AP_REPORT_PRINT1", trvMaster.SelectedItem.Text) Then
'         Call EnableForm(Me, True)
'         Exit Sub
'      End If

      Set Report = New CReportArMas02
      SelectFlag = True
   ElseIf KEY = ROOT_TREE & " 1-0-1-2" Then
'      If Not VerifyAccessRight("AP_REPORT_PRINT1", trvMaster.SelectedItem.Text) Then
'         Call EnableForm(Me, True)
'         Exit Sub
'      End If

      Set Report = New CReportArMas03
      Picture1.Picture = LoadPicture(glbParameterObj.MgpCustomerProfilePic)
      Call Report.AddParam(Picture1.Picture, "BACK_GROUND")
      
      SelectFlag = True
   ElseIf KEY = ROOT_TREE & " 1-0-2" Then
'      If Not VerifyAccessRight("AP_REPORT_PRINT1", trvMaster.SelectedItem.Text) Then
'         Call EnableForm(Me, True)
'         Exit Sub
'      End If

      Set Report = New CReportAR002
      SelectFlag = True
   ElseIf KEY = ROOT_TREE & " 1-0-2-1" Then
'      If Not VerifyAccessRight("AP_REPORT_PRINT1", trvMaster.SelectedItem.Text) Then
'         Call EnableForm(Me, True)
'         Exit Sub
'      End If

      Set Report = New CReportAR002_1
      SelectFlag = True
   ElseIf KEY = ROOT_TREE & " 1-0-4" Then
'      If Not VerifyAccessRight("AP_REPORT_PRINT1", trvMaster.SelectedItem.Text) Then
'         Call EnableForm(Me, True)
'         Exit Sub
'      End If

      Set Report = New CReportAsset001
      SelectFlag = True
'   ElseIf KEY = ROOT_TREE & " 1-0-5" Then
''      If Not VerifyAccessRight("AP_REPORT_PRINT1", trvMaster.SelectedItem.Text) Then
''         Call EnableForm(Me, True)
''         Exit Sub
''      End If
'
'      Set Report = New CReportAP005
'      SelectFlag = True
   ElseIf KEY = ROOT_TREE & " 1-0-5-1" Then
'      If Not VerifyAccessRight("AP_REPORT_PRINT1", trvMaster.SelectedItem.Text) Then
'         Call EnableForm(Me, True)
'         Exit Sub
'      End If
      
      Set Report = New CReportAP005_1
      SelectFlag = True
   ElseIf KEY = ROOT_TREE & " 1-0-5-2" Then
'      If Not VerifyAccessRight("AP_REPORT_PRINT1", trvMaster.SelectedItem.Text) Then
'         Call EnableForm(Me, True)
'         Exit Sub
'      End If

      Set Report = New CReportAP005_2
      SelectFlag = True
   ElseIf KEY = ROOT_TREE & " 1-0-5-3" Then
'      If Not VerifyAccessRight("AP_REPORT_PRINT1", trvMaster.SelectedItem.Text) Then
'         Call EnableForm(Me, True)
'         Exit Sub
'      End If

      Set Report = New CReportAP005_3
      SelectFlag = True
   ElseIf KEY = ROOT_TREE & " 1-0-5-4" Then
'      If Not VerifyAccessRight("AP_REPORT_PRINT1", trvMaster.SelectedItem.Text) Then
'         Call EnableForm(Me, True)
'         Exit Sub
'      End If

      Set Report = New CReportAP005_4
      SelectFlag = True
   ElseIf KEY = ROOT_TREE & " 1-0-6" Then
'      If Not VerifyAccessRight("AP_REPORT_PRINT1", trvMaster.SelectedItem.Text) Then
'         Call EnableForm(Me, True)
'         Exit Sub
'      End If

      Set Report = New CReportAR005
      SelectFlag = True
   ElseIf KEY = ROOT_TREE & " 1-0-6-1" Then
'      If Not VerifyAccessRight("AP_REPORT_PRINT1", trvMaster.SelectedItem.Text) Then
'         Call EnableForm(Me, True)
'         Exit Sub
'      End If

      Set Report = New CReportAR005_1
      SelectFlag = True
   ElseIf KEY = ROOT_TREE & " 1-0-6-2" Then
'      If Not VerifyAccessRight("AP_REPORT_PRINT1", trvMaster.SelectedItem.Text) Then
'         Call EnableForm(Me, True)
'         Exit Sub
'      End If

      Set Report = New CReportAR005_2
      SelectFlag = True
      
     ElseIf KEY = ROOT_TREE & " 1-0-6-3" Then
'      If Not VerifyAccessRight("AP_REPORT_PRINT1", trvMaster.SelectedItem.Text) Then
'         Call EnableForm(Me, True)
'         Exit Sub
'      End If
      Set Report = New CReportAR005_3
      SelectFlag = True
       ElseIf KEY = ROOT_TREE & " 1-0-6-4" Then
'      If Not VerifyAccessRight("AP_REPORT_PRINT1", trvMaster.SelectedItem.Text) Then
'         Call EnableForm(Me, True)
'         Exit Sub
'      End If
      Set Report = New CReportAR005_4
      SelectFlag = True
 ElseIf KEY = ROOT_TREE & " 1-0-6-5" Then
      Set Report = New CReportSaleApprove
      SelectFlag = True
   ElseIf KEY = ROOT_TREE & " 1-0-7" Then
'      If Not VerifyAccessRight("AP_REPORT_PRINT1", trvMaster.SelectedItem.Text) Then
'         Call EnableForm(Me, True)
'         Exit Sub
'      End If

      Set Report = New CReportAR006
      SelectFlag = True
      
   ElseIf KEY = ROOT_TREE & " 1-0-7-1" Then
      Set Report = New CReportAR013
      SelectFlag = True
      
    ElseIf KEY = ROOT_TREE & " 1-0-8" Then
'      If Not VerifyAccessRight("AP_REPORT_PRINT1", trvMaster.SelectedItem.Text) Then
'         Call EnableForm(Me, True)
'         Exit Sub
'      End If

      Set Report = New CReportAR007
      SelectFlag = True
      
     ElseIf KEY = ROOT_TREE & " 1-0-8-1" Then
'      If Not VerifyAccessRight("AP_REPORT_PRINT1", trvMaster.SelectedItem.Text) Then
'         Call EnableForm(Me, True)
'         Exit Sub
'      End If
    Set Report = New CReportAR007_1
    SelectFlag = True
      
   ElseIf KEY = ROOT_TREE & " 1-0-9" Then
'      If Not VerifyAccessRight("AP_REPORT_PRINT1", trvMaster.SelectedItem.Text) Then
'         Call EnableForm(Me, True)
'         Exit Sub
'      End If

      Set Report = New CReportAR008
      SelectFlag = True
   ElseIf KEY = ROOT_TREE & " 1-0-10" Then
'      If Not VerifyAccessRight("AP_REPORT_PRINT1", trvMaster.SelectedItem.Text) Then
'         Call EnableForm(Me, True)
'         Exit Sub
'      End If

      Set Report = New CReportAR009
      SelectFlag = True
   ElseIf KEY = ROOT_TREE & " 1-0-10-1" Then
'      If Not VerifyAccessRight("AP_REPORT_PRINT1", trvMaster.SelectedItem.Text) Then
'         Call EnableForm(Me, True)
'         Exit Sub
'      End If

      Set Report = New CReportAR009_1
      SelectFlag = True
   ElseIf KEY = ROOT_TREE & " 1-0-10-2" Then
'      If Not VerifyAccessRight("AP_REPORT_PRINT1", trvMaster.SelectedItem.Text) Then
'         Call EnableForm(Me, True)
'         Exit Sub
'      End If

      Set Report = New CReportAR009_2
      SelectFlag = True
      
      ElseIf KEY = ROOT_TREE & " 1-0-10-3" Then
'      If Not VerifyAccessRight("AP_REPORT_PRINT1", trvMaster.SelectedItem.Text) Then
'         Call EnableForm(Me, True)
'         Exit Sub
'      End If

      Set Report = New CReportAR009_3
      SelectFlag = True
   ElseIf KEY = ROOT_TREE & " 1-0-10-4" Then
'      If Not VerifyAccessRight("AP_REPORT_PRINT1", trvMaster.SelectedItem.Text) Then
'         Call EnableForm(Me, True)
'         Exit Sub
'      End If

      Set Report = New CReportAR009_10
      SelectFlag = True
   ElseIf KEY = ROOT_TREE & " 1-0-12" Then
'      If Not VerifyAccessRight("AP_REPORT_PRINT1", trvMaster.SelectedItem.Text) Then
'         Call EnableForm(Me, True)
'         Exit Sub
'      End If

      Set Report = New CReportAR011
      SelectFlag = True
   ElseIf KEY = ROOT_TREE & " 1-0-13" Then
'      If Not VerifyAccessRight("AP_REPORT_PRINT1", trvMaster.SelectedItem.Text) Then
'         Call EnableForm(Me, True)
'         Exit Sub
'      End If

      Set Report = New CReportAR011_1
      SelectFlag = True
   ElseIf KEY = ROOT_TREE & " 1-0-15-1" Then
      Set Report = New CReportAR012
      SelectFlag = True
   
   ElseIf KEY = ROOT_TREE & " 1-0-16" Then
'      If Not VerifyAccessRight("AP_REPORT_PRINT1", trvMaster.SelectedItem.Text) Then
'         Call EnableForm(Me, True)
'         Exit Sub
'      End If

      Set Report = New CReportAR016
      SelectFlag = True
   ElseIf KEY = ROOT_TREE & " 1-0-17" Then
      Set Report = New CReportAR017
      SelectFlag = True
   ElseIf KEY = ROOT_TREE & " 1-0-18" Then
      Set Report = New CReportAR018
      SelectFlag = True
   ElseIf KEY = ROOT_TREE & " 1-0-19" Then
      Set Report = New CReportAR019
      SelectFlag = True
   ElseIf KEY = ROOT_TREE & " 1-0-20" Then
      Set Report = New CReportAR020
      SelectFlag = True
   ElseIf KEY = ROOT_TREE & " 1-0-21" Then
      Set Report = New CReportCostProducts
      SelectFlag = True
   ElseIf KEY = ROOT_TREE & " 2-0-1" Then
'      If Not VerifyAccessRight("AP_REPORT_PRINT1", trvMaster.SelectedItem.Text) Then
'         Call EnableForm(Me, True)
'         Exit Sub
'      End If

      Set Report = New CReportAP001
      SelectFlag = True
   ElseIf KEY = ROOT_TREE & " 2-0-2" Then
'      If Not VerifyAccessRight("AP_REPORT_PRINT1", trvMaster.SelectedItem.Text) Then
'         Call EnableForm(Me, True)
'         Exit Sub
'      End If

      Set Report = New CReportAP002
      SelectFlag = True
   ElseIf KEY = ROOT_TREE & " 2-0-2-1" Then
'      If Not VerifyAccessRight("AP_REPORT_PRINT1", trvMaster.SelectedItem.Text) Then
'         Call EnableForm(Me, True)
'         Exit Sub
'      End If

      Set Report = New CReportAP002_1
      SelectFlag = True
   ElseIf KEY = ROOT_TREE & " 2-0-3" Then
'      If Not VerifyAccessRight("AP_REPORT_PRINT1", trvMaster.SelectedItem.Text) Then
'         Call EnableForm(Me, True)
'         Exit Sub
'      End If

      Set Report = New CReportAP004
      SelectFlag = True
   ElseIf KEY = ROOT_TREE & " 2-0-3-1" Then
      Set Report = New CReportAP004_1
      SelectFlag = True
    ElseIf KEY = ROOT_TREE & " 2-0-3-2" Then
      Set Report = New CReportAP004_2
      SelectFlag = True
   ElseIf KEY = ROOT_TREE & " 2-0-3-3" Then
      Set Report = New CReportAP004_3
      SelectFlag = True
   ElseIf KEY = ROOT_TREE & " 2-0-4" Then
'      If Not VerifyAccessRight("AP_REPORT_PRINT1", trvMaster.SelectedItem.Text) Then
'         Call EnableForm(Me, True)
'         Exit Sub
'      End If

      Set Report = New CReportAP002_Temp
      SelectFlag = True
   ElseIf KEY = ROOT_TREE & " 2-0-4-1" Then
'      If Not VerifyAccessRight("AP_REPORT_PRINT1", trvMaster.SelectedItem.Text) Then
'         Call EnableForm(Me, True)
'         Exit Sub
'      End If

      Set Report = New CReportAP002_2
      SelectFlag = True
   ElseIf KEY = ROOT_TREE & " 2-0-5" Then
'      If Not VerifyAccessRight("AP_REPORT_PRINT1", trvMaster.SelectedItem.Text) Then
'         Call EnableForm(Me, True)
'         Exit Sub
'      End If
      
      Set Report = New CReportAP006
      SelectFlag = True
   ElseIf KEY = ROOT_TREE & " 2-0-5-1" Then
'      If Not VerifyAccessRight("AP_REPORT_PRINT1", trvMaster.SelectedItem.Text) Then
'         Call EnableForm(Me, True)
'         Exit Sub
'      End If
      
      Set Report = New CReportAP006_1
      SelectFlag = True
   ElseIf KEY = ROOT_TREE & " 2-0-6" Then
'      If Not VerifyAccessRight("AP_REPORT_PRINT1", trvMaster.SelectedItem.Text) Then
'         Call EnableForm(Me, True)
'         Exit Sub
'      End If
      
      Set Report = New CReportAP007
      SelectFlag = True
   ElseIf KEY = ROOT_TREE & " 2-0-6-1" Then
'      If Not VerifyAccessRight("AP_REPORT_PRINT1", trvMaster.SelectedItem.Text) Then
'         Call EnableForm(Me, True)
'         Exit Sub
'      End If
      
      Set Report = New CReportAP009
      SelectFlag = True
   ElseIf KEY = ROOT_TREE & " 2-0-6-2" Then
'      If Not VerifyAccessRight("AP_REPORT_PRINT1", trvMaster.SelectedItem.Text) Then
'         Call EnableForm(Me, True)
'         Exit Sub
'      End If
      
      Set Report = New CReportAP009_1
      SelectFlag = True
   ElseIf KEY = ROOT_TREE & " 2-0-7" Then
'      If Not VerifyAccessRight("AP_REPORT_PRINT1", trvMaster.SelectedItem.Text) Then
'         Call EnableForm(Me, True)
'         Exit Sub
'      End If
      
      Set Report = New CReportAP008
      SelectFlag = True
   ElseIf KEY = ROOT_TREE & " 2-0-8" Then
      Set Report = New CReportAP016_4
      SelectFlag = True
   ElseIf KEY = ROOT_TREE & "-A" & " 5-0" Then
'      If Not VerifyAccessRight("AP_REPORT_PRINT1", trvMaster.SelectedItem.Text) Then
'         Call EnableForm(Me, True)
'         Exit Sub
'      End If

      Set Report = New CReportAP003
      Call Report.AddParam(1, "REPORT_TYPE")
      Call Report.AddParam(m_Journals, "JOURNAL")
      Call Report.AddParam(2, "JOURNAL_TYPE")
      Picture1.Picture = LoadPicture(glbParameterObj.PaidVocherPic)
      Call Report.AddParam(Picture1.Picture, "BACK_GROUND")
      SelectFlag = True
   ElseIf KEY = ROOT_TREE & "-A" & " 5-1" Then
      Set Report = New CReportAP003
      Call Report.AddParam(2, "REPORT_TYPE")
      Picture1.Picture = LoadPicture(glbParameterObj.PaidVocherPic)
      Call Report.AddParam(Picture1.Picture, "BACK_GROUND")
      Call Report.AddParam(m_Journals, "JOURNAL")                       '  Collection
      SelectFlag = True
   ElseIf KEY = ROOT_TREE & "-A" & " 5-2" Then
      Set Report = New CReportJV001
      Picture1.Picture = LoadPicture(glbParameterObj.JVVocherPic)
      Call Report.AddParam(Picture1.Picture, "BACK_GROUND")
      SelectFlag = True
   ElseIf KEY = ROOT_TREE & "-A" & " 5-3" Then
      Set Report = New CReportAR001
      Call Report.AddParam(1, "REPORT_TYPE")
      Call Report.AddParam(m_Journals, "JOURNAL")
      Call Report.AddParam(2, "JOURNAL_TYPE")
      'Picture1.Picture = LoadPicture(glbParameterObj.ReceiptVocherPic)
      Call Report.AddParam(Picture1.Picture, "BACK_GROUND")
      SelectFlag = True
   ElseIf KEY = ROOT_TREE & "-A" & " 5-4" Then
      Set Report = New CReportAR001
      Call Report.AddParam(2, "REPORT_TYPE")
      ''''''Picture1.Picture = LoadPicture(glbParameterObj.ReceiptVocherPic)
      Call Report.AddParam(Picture1.Picture, "BACK_GROUND")
      Call Report.AddParam(m_Journals, "JOURNAL")                       '  Collection
      SelectFlag = True
   ElseIf KEY = ROOT_TREE & "-B" & " 1-1" Then
      Set Report = New CReportCheq001
      SelectFlag = True

   ' commission เก็บเงิน
   ElseIf KEY = ROOT_TREE & " 6-0-1" Then
      Set Report = New CReportCom01
      SelectFlag = True
         
   ElseIf KEY = ROOT_TREE & " 6-0-2" Then
      Set Report = New CReportCom02
      SelectFlag = True
   
   ElseIf KEY = ROOT_TREE & " 6-0-3" Then
      Set Report = New CReportCom03
      SelectFlag = True
         
   ElseIf KEY = ROOT_TREE & " 6-0-4" Then
      Set Report = New CReportCom04
      SelectFlag = True
      
   ElseIf KEY = ROOT_TREE & " 6-0-5" Then
      Set Report = New CReportCom02_1
      SelectFlag = True
   
  ElseIf KEY = ROOT_TREE & " 6-0-6" Then
      Set Report = New CReportCom02_2
      SelectFlag = True
      
 ElseIf KEY = ROOT_TREE & " 6-0-7" Then
      Set Report = New CReportCom05
      SelectFlag = True
      
 ElseIf KEY = ROOT_TREE & " 6-0-8" Then
      Set Report = New CReportCom06
      SelectFlag = True
      
 ElseIf KEY = ROOT_TREE & " 6-0-9" Then
      Set Report = New CReportCom07
      SelectFlag = True
      
 ElseIf KEY = ROOT_TREE & " 6-0-10" Then
      Set Report = New CReportCom10
      SelectFlag = True
      
   ElseIf KEY = ROOT_TREE & " 6-0-11" Then
      Set Report = New CReportCom11
      SelectFlag = True
      
   ElseIf KEY = ROOT_TREE & " 6-0-12" Then
      Set Report = New CReportCom12
      SelectFlag = True
   
   ElseIf KEY = ROOT_TREE & " 6-0-13" Then
      Set Report = New CReportCom13
      SelectFlag = True
      
   ElseIf KEY = ROOT_TREE & " 6-0-14" Then
      Set Report = New CReportCom14
      SelectFlag = True
   ElseIf KEY = ROOT_TREE & "-C" & " 1-1" Then
      Set Report = New CReportGL01
      SelectFlag = True
   End If

   If SelectFlag Then
      If glbParameterObj.Temp = 0 Then
         glbParameterObj.UsedCount = glbParameterObj.UsedCount + 1
         glbParameterObj.Temp = 1
      End If
      
      Call FillReportInput(Report)
      Call Report.AddParam(Name, "REPORT_NAME")
      Call Report.AddParam(KEY, "REPORT_KEY")
      
      Set frmReport.ReportObject = Report
      frmReport.HeaderText = MapText("พิมพ์รายงาน")
      Load frmReport
      frmReport.Show 1

      Unload frmReport
      Set frmReport = Nothing
   End If
End Sub


Private Sub cmdPasswd_Click()
Dim lMenuChosen As Long
Dim oMenu As cPopupMenu
Dim OKClick As Boolean
Dim DBPath As String

   Set oMenu = New cPopupMenu
   lMenuChosen = oMenu.Popup("เปลี่ยนบริษัท", "ติดต่อ Database ที่ 2", "ติดต่อ Database ที่ 3", "Patch ค่าเสื่อม", "Export ข้อมูลกระดาษทำการ", "ประมาณการอาหาร / คุมวงเงิน", "ตาราง CREDIT ตามบิล", "ประเภทข้อมูล", "ประเภทกลุ่ม", "ประเภทกลุ่มย่อย", "รวมกลุ่มประเภทย่อย", "กลุ่มเจ้าหนี้", "เอกสารยกเลิก", "-", "Export ข้อมูลบิลขายไปให้สาขา", "-", "ประเภทลูกหนี้", "ลูกหนี้ตามแนวตั้ง", "วิเคราะห์อายุลูกหนี้", "-", "เช็คยกเลิก", "-", "ธนาคาร", "ลูกหนี้ธนาคาร", "วงเงินธนาคาร", "เลขที่ตั๋ว", "ประมาณการตั๋ว", "-", "เงื่อนไข Commission", "เงื่อนไข(พิเศษ) Commission", "เงื่อนไข(พิเศษ) Incentive", "ตั้งงบประมาณ Commercial #1 ", "สินค้าไม่คิด Commercial #1", "ตั้งค่าเขตการขาย ", "โปรแกรมส่วนลด ", "โปรแกรม IV กลุ่มขายสด", "โปรแกรมค่า GP", "โปรแกรมปรับเครดิตแต่ละ IV", "ตั้งค่าประเภทสินค้า", "ตั้งค่าประเภทหมูของลูกค้า")
   If lMenuChosen = 0 Then
      Exit Sub
   End If
   
   If lMenuChosen = 1 Then
      Call glbDatabaseMngr.DisConnectDatabase
      Call glbDatabaseMngr.ConnectDatabase(glbParameterObj.DBFile, "", "", glbErrorLog)
      
      frmDatabaseSelect.ShowMode = SHOW_EDIT
      Load frmDatabaseSelect
      frmDatabaseSelect.Show 1
   
      OKClick = frmDatabaseSelect.OKClick
      DBPath = frmDatabaseSelect.DBPath
      Unload frmDatabaseSelect
      Set frmDatabaseSelect = Nothing
     
      If OKClick Then
          Call glbDatabaseMngr.DisConnectDatabase
          Call glbDatabaseMngr.ConnectDatabase(DBPath, "", "", glbErrorLog)
          
          Me.Caption = glbCompanyName
          m_DBPath = DBPath
       Else
          Call glbDatabaseMngr.DisConnectDatabase
          Call glbDatabaseMngr.ConnectDatabase(m_DBPath, "", "", glbErrorLog)
       End If
    
    ElseIf lMenuChosen = 2 Then
      Call glbDatabaseMngr.DisConnectDatabase2
      Call glbDatabaseMngr.ConnectDatabase2(glbParameterObj.DBFile, "", "", glbErrorLog)
      
      frmDatabaseSelect.ShowMode = SHOW_EDIT
      frmDatabaseSelect.Database2 = True
      Load frmDatabaseSelect
      frmDatabaseSelect.Show 1
   
      OKClick = frmDatabaseSelect.OKClick
      DBPath = frmDatabaseSelect.DBPath
      Unload frmDatabaseSelect
      Set frmDatabaseSelect = Nothing
     
      If OKClick Then
         Call glbDatabaseMngr.DisConnectDatabase2
         Call glbDatabaseMngr.ConnectDatabase2(DBPath, "", "", glbErrorLog)
         
         Me.Caption = glbCompanyName
         m_DBPath = DBPath
       Else
         Call glbDatabaseMngr.DisConnectDatabase2
         Call glbDatabaseMngr.ConnectDatabase2(m_DBPath, "", "", glbErrorLog)
        End If
        
   ElseIf lMenuChosen = 3 Then
      Call glbDatabaseMngr.DisConnectDatabase3
      Call glbDatabaseMngr.ConnectDatabase3(glbParameterObj.DBFile, "", "", glbErrorLog)
      
      frmDatabaseSelect.ShowMode = SHOW_EDIT
      frmDatabaseSelect.Database3 = True
      Load frmDatabaseSelect
      frmDatabaseSelect.Show 1
   
      OKClick = frmDatabaseSelect.OKClick
      DBPath = frmDatabaseSelect.DBPath
      Unload frmDatabaseSelect
      Set frmDatabaseSelect = Nothing
     
      If OKClick Then
         Call glbDatabaseMngr.DisConnectDatabase3
         Call glbDatabaseMngr.ConnectDatabase3(DBPath, "", "", glbErrorLog)
         
         Me.Caption = glbCompanyName
         m_DBPath = DBPath
       Else
         Call glbDatabaseMngr.DisConnectDatabase3
         Call glbDatabaseMngr.ConnectDatabase2(m_DBPath, "", "", glbErrorLog)
        End If
      
   ElseIf lMenuChosen = 4 Then
      Dim Fa As CFaMas
      Set Fa = New CFaMas
      Fa.ASSET_CODE_SET = "('MM-06-05-018', 'MM-06-07-024', 'MM-06-01-072', 'MM-06-01-073', 'MM-04-08-062', 'MM-04-08-063', 'MM-05-07-108', 'MM-05-07-109', 'MM-05-08-108', 'MM-05-08-109') "
      Fa.DPRVAL = 0
      glbDaily.StartTransaction
      Call Fa.PatchDprVal
      glbDaily.CommitTransaction
      Set Fa = Nothing
      
   ElseIf lMenuChosen = 5 Then
      Load frmExportChartAccount
      frmExportChartAccount.Show 1
      
      Unload frmExportChartAccount
      Set frmExportChartAccount = Nothing
      
  ElseIf lMenuChosen = 6 Then             ' ประมาณการอาหารและคุมวงเงิน
      Set oMenu = New cPopupMenu
      lMenuChosen = oMenu.Popup("ประมาณการอาหาร", "ทำต่อจากคาร์กิล", "จัดจ่าย")
      If lMenuChosen = 0 Then
         Exit Sub
      End If
      
      If lMenuChosen = 1 Then
         Load frmXlsEstimate
         frmXlsEstimate.Show 1
         Unload frmXlsEstimate
         Set frmXlsEstimate = Nothing
      End If
      
       If lMenuChosen = 2 Then
         Load frmXlsCarkill
         frmXlsCarkill.Show 1
         Unload frmXlsCarkill
         Set frmXlsCarkill = Nothing
      End If
      
      If lMenuChosen = 3 Then
         Load frmXlsFoodPay
         frmXlsFoodPay.Show 1
         Unload frmXlsFoodPay
         Set frmXlsFoodPay = Nothing
      End If
      
   ElseIf lMenuChosen = 7 Then
      Load frmRealCredit
      frmRealCredit.Show 1
      
      Unload frmRealCredit
      Set frmRealCredit = Nothing
    ElseIf lMenuChosen = 8 Then 'ประเภทข้อมูล
      Load frmDataType
      frmDataType.Show 1
      
      Unload frmDataType
      Set frmDataType = Nothing
   ElseIf lMenuChosen = 9 Then
      Load frmGroupType
      frmGroupType.Show 1
      
      Unload frmGroupType
      Set frmGroupType = Nothing
   ElseIf lMenuChosen = 10 Then 'ประเภทกลุ่มย่อย
      Load frmSubGroupType
      frmSubGroupType.Show 1

      Unload frmSubGroupType
      Set frmSubGroupType = Nothing
      
    ElseIf lMenuChosen = 11 Then 'combo ประเภทกลุ่มย่อย
      frmComboSubGroup.HeaderText = MapText("รวมกลุ่มประเภทย่อย")
      Load frmComboSubGroup
      frmComboSubGroup.Show 1

      Unload frmComboSubGroup
      Set frmComboSubGroup = Nothing
      
   ElseIf lMenuChosen = 12 Then
      Load frmSupplierGroup
      frmSupplierGroup.Show 1
      
      Unload frmSupplierGroup
      Set frmSupplierGroup = Nothing
      
   ElseIf lMenuChosen = 13 Then                                                 ' 8 อันนี้ล่ะ ที่จะเพิ่มเข้ามาใน 1.15.1
      Load frmDocumentCancel
      frmDocumentCancel.Show 1
      
      Unload frmDocumentCancel
      Set frmDocumentCancel = Nothing
      
   ElseIf lMenuChosen = 15 Then
      Load frmExportBill
      frmExportBill.Show 1
      
      Unload frmExportBill
      Set frmExportBill = Nothing
   ElseIf lMenuChosen = 17 Then
      Load frmCustomerType
      frmCustomerType.Show 1
      
      Unload frmCustomerType
      Set frmCustomerType = Nothing
      
   ElseIf lMenuChosen = 18 Then
     Load frmCustomer
      frmCustomer.Show 1
      
     Unload frmCustomer
     Set frmCustomer = Nothing
     
   ElseIf lMenuChosen = 19 Then
     Load frmAnalyzeCustomer
      frmAnalyzeCustomer.Show 1

     Unload frmAnalyzeCustomer
     Set frmAnalyzeCustomer = Nothing
   
  ElseIf lMenuChosen = 21 Then
     Load frmCheckCancel
      frmCheckCancel.Show 1
     Unload frmCheckCancel
     Set frmCheckCancel = Nothing
   
   ElseIf lMenuChosen = 23 Then
      Load frmBank
      frmBank.Show 1

      Unload frmBank
      Set frmBank = Nothing
      
   ElseIf lMenuChosen = 24 Then
      Load frmBankCustomer
      frmBankCustomer.Show 1

      Unload frmBankCustomer
      Set frmBankCustomer = Nothing
     
   ElseIf lMenuChosen = 25 Then
      Load frmBankCredit
      frmBankCredit.Show 1

      Unload frmBankCredit
      Set frmBankCredit = Nothing
      
   ElseIf lMenuChosen = 26 Then
      Load frmTicket
      frmTicket.Show 1

      Unload frmTicket
      Set frmTicket = Nothing
      
   ElseIf lMenuChosen = 27 Then
      Load frmBudgetTicket
      frmBudgetTicket.Show 1

      Unload frmBudgetTicket
      Set frmBudgetTicket = Nothing
      
   'commission
   ElseIf lMenuChosen = 29 Then
      Load frmConditionCommission
      frmConditionCommission.Show 1

      Unload frmConditionCommission
      Set frmConditionCommission = Nothing
            
    ElseIf lMenuChosen = 30 Then        ' Com ระยะสั้น
       frmPromoteCommission.HeaderText = MapText("เงื่อนไข(พิเศษ) Commission")
      Load frmPromoteCommission
      frmPromoteCommission.Show 1

      Unload frmPromoteCommission
      Set frmPromoteCommission = Nothing
      
    ElseIf lMenuChosen = 31 Then            ' Incentive ระยะสั้น
       frmPromoteIncentive.HeaderText = MapText("เงื่อนไข(พิเศษ) Incentive")
      Load frmPromoteIncentive
      frmPromoteIncentive.Show 1

      Unload frmPromoteIncentive
      Set frmPromoteIncentive = Nothing
      
    ElseIf lMenuChosen = 32 Then
      frmMasterFromTo.HeaderText = MapText("ตั้งค่างบประมาณ commercial #1")
      Load frmMasterFromTo
      frmMasterFromTo.Show 1                           ' ตั้งค่างบประมาณ

      Unload frmMasterFromTo
      Set frmMasterFromTo = Nothing

   ElseIf lMenuChosen = 33 Then
     frmComDonStk.HeaderText = MapText("สินค้าไม่คิด commercial #1")
      Load frmComDonStk
      frmComDonStk.Show 1

      Unload frmComDonStk
      Set frmComDonStk = Nothing
      
   ElseIf lMenuChosen = 34 Then
      Load frmAreaMasterCom
      frmAreaMasterCom.Show 1

      Unload frmAreaMasterCom
      Set frmAreaMasterCom = Nothing

  ElseIf lMenuChosen = 35 Then
   '   frmComMinusStk.HeaderText = MapText("โปรแกรมส่วนลด")
      Load frmComMinusStk
      frmComMinusStk.Show 1

      Unload frmComMinusStk
      Set frmComMinusStk = Nothing
   
 ElseIf lMenuChosen = 36 Then
      frmComIVcenter.HeaderText = MapText("โปรแกรม IV เขตการขายCenter")
      Load frmComIVcenter
      frmComIVcenter.Show 1

      Unload frmComIVcenter
      Set frmComIVcenter = Nothing
      
   ElseIf lMenuChosen = 37 Then
      frmMaster2FromTo.HeaderText = MapText("โปรแกรมค่า GP")
      Load frmMaster2FromTo
      frmMaster2FromTo.Show 1

      Unload frmMaster2FromTo
      Set frmMaster2FromTo = Nothing
      
    ElseIf lMenuChosen = 38 Then
      frmIVcredit.HeaderText = MapText("โปรแกรมปรับเครดิตแต่ละ IV")
      Load frmIVcredit
      frmIVcredit.Show 1

      Unload frmIVcredit
      Set frmIVcredit = Nothing
   
   ElseIf lMenuChosen = 39 Then
      ' frmGoodsMasterCom.HeaderText = MapText("ตั้งค่าสินค้า")
      Load frmGoodsMasterCom
       frmGoodsMasterCom.Show 1

      Unload frmGoodsMasterCom
      Set frmGoodsMasterCom = Nothing
    ElseIf lMenuChosen = 40 Then
      frmCusPigType.HeaderText = MapText("ตั้งค่าประเภทหมูของลูกค้า")
      Load frmCusPigType
      frmCusPigType.Show 1

      Unload frmCusPigType
      Set frmCusPigType = Nothing
   End If
   
End Sub

Private Sub cmdPasswd2_Click()
Dim lMenuChosen As Long
Dim oMenu As cPopupMenu
Dim OKClick As Boolean
Dim DBPath As String

   Set oMenu = New cPopupMenu
   lMenuChosen = oMenu.Popup("โปรแกรม Promotion", "-", "โปรแกรมของแถมตามปี", "-", "MAP จังหวัด", "-", "โปรแกรม ต้นทุนสินค้า")
   If lMenuChosen = 0 Then
      Exit Sub
   End If
   
   If lMenuChosen = 1 Then
      frmPromotionPayCustomer.HeaderText = MapText("Promotion")
      Load frmPromotionPayCustomer
      frmPromotionPayCustomer.Show 1

      Unload frmPromotionPayCustomer
      Set frmPromotionPayCustomer = Nothing
      
   ElseIf lMenuChosen = 3 Then
      frmPromotionYear.HeaderText = MapText("Promotion Year")
      Load frmPromotionYear
      frmPromotionYear.Show 1

      Unload frmPromotionYear
      Set frmPromotionYear = Nothing
   ElseIf lMenuChosen = 5 Then
   
      Load frmProvinceMap
      frmProvinceMap.Show 1

      Unload frmProvinceMap
      Set frmProvinceMap = Nothing
      
   ElseIf lMenuChosen = 7 Then
      frmCostProducts.HeaderText = MapText("โปรแกรมต้นทุน")
      Load frmCostProducts
      frmCostProducts.Show 1

      Unload frmCostProducts
      Set frmCostProducts = Nothing
   End If
End Sub

Private Sub Form_Activate()
Dim OKClick As Boolean
Dim DBPath As String

   If m_HasActivate Then
      Exit Sub
   End If
   m_HasActivate = True

   Call EnableForm(Me, False)
   Call PatchDB

   frmDatabaseSelect.ShowMode = SHOW_EDIT
   Load frmDatabaseSelect
   frmDatabaseSelect.Show 1

   OKClick = frmDatabaseSelect.OKClick
   DBPath = frmDatabaseSelect.DBPath
   Unload frmDatabaseSelect
   Set frmDatabaseSelect = Nothing
  
  If OKClick Then
      m_DBPath = DBPath
      Call glbDatabaseMngr.DisConnectDatabase
      Call glbDatabaseMngr.ConnectDatabase(DBPath, "", "", glbErrorLog)
      
      Me.Caption = glbCompanyName
      
      'Call LoadSupplier(Nothing, m_SupplierColl)
   End If
   
   Call EnableForm(Me, True)
   
   Dim SumAll As Double
   Dim i As Long
   Dim Data(101) As Double
   
   
   If Not OKClick Then
      m_MustAsk = False
      Unload Me
   End If
End Sub

Private Sub Form_Load()
   m_MustAsk = True
   Call InitFormLayout
   Set m_Rs = New ADODB.Recordset
   
   Set m_ReportControls = New Collection
   Set m_Texts = New Collection
   Set m_Dates = New Collection
   Set m_Labels = New Collection
  Set m_TextLookups = New Collection
   Set m_Combos = New Collection
   Set m_ReportParams = New Collection
   Set m_CheckBoxes = New Collection

   Set m_Journals = New Collection                   'Step 2
End Sub

Private Sub InitFormLayout()
'   Call InitNormalLabel(lblUserName, MapText("ผู้ใช้ : "), RGB(0, 0, 255))
'   Call InitNormalLabel(lblUserGroup, MapText("กลุ่มผู้ใช้ : "), RGB(0, 0, 255))
   Call InitNormalLabel(lblVersion, MapText("เวอร์ชัน : ") & glbParameterObj.Version & " (Interbase) ", RGB(0, 0, 255))
   Call InitNormalLabel(lblDateTime, "", RGB(0, 0, 255))
   lblDateTime.FontSize = 30
   lblDateTime.BackStyle = 1
   
   lblDateTime.BackColor = RGB(255, 255, 255)
   
   SSFrame2.PictureBackground = LoadPicture(glbParameterObj.NormalForm1)
   SSFrame3.PictureBackground = LoadPicture(glbParameterObj.NormalForm1)
   
   cmdAdd.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdExit.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdPasswd.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdPasswd2.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdConfig.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdOK.Picture = LoadPicture(glbParameterObj.NormalButton1)
   
   Me.Caption = MapText("ระบบรายงานบัญชี Express")
   
   pnlHeader.Font.Name = GLB_FONT
   pnlHeader.Font.Bold = True
   pnlHeader.Font.Size = 19
      
   Call InitMainButton(cmdExit, MapText("ออก"))
   Call InitMainButton(cmdPasswd, MapText("โปรแกรม"))
   Call InitMainButton(cmdPasswd2, MapText("โปรแกรม 2"))
   Call InitMainButton(cmdOK, MapText("พิมพ์ (F10)"))
   Call InitMainButton(cmdConfig, MapText("ปรับค่า"))
   Call InitMainButton(cmdAdd, MapText("เพิ่ม"))

   Call InitMainTreeview
End Sub

Private Sub InitMainTreeview()
Dim Node As Node
Dim NewNodeID As String

   trvMain.Nodes.Clear
   trvMain.Font.Name = GLB_FONT
   trvMain.Font.Size = 14
   trvMain.Font.Bold = False
      
   Set Node = trvMain.Nodes.Add(, tvwFirst, ROOT_TREE, MapText("ระบบรายงานบัญชี Express"), 8, 8)
   Node.Expanded = True
   Node.Selected = True
   
   Set Node = trvMain.Nodes.Add(ROOT_TREE, tvwChild, ROOT_TREE & " 1-0", MapText("1. ข้อมูลสินทรัพย์"), 3, 3)
   Node.Expanded = False
   
   Set Node = trvMain.Nodes.Add(ROOT_TREE & " 1-0", tvwChild, ROOT_TREE & " 1-0-1", MapText("1.1. รายงานข้อมูลลูกหนี้"), 12, 11)
   Node.Expanded = False
   
   Set Node = trvMain.Nodes.Add(ROOT_TREE & " 1-0", tvwChild, ROOT_TREE & " 1-0-1-1", MapText("1.1.1. รายงานข้อมูลลูกหนี้ (จดหมาย)"), 12, 11)
   Node.Expanded = False
   
   Set Node = trvMain.Nodes.Add(ROOT_TREE & " 1-0", tvwChild, ROOT_TREE & " 1-0-1-2", MapText("1.1.2. รายงานข้อมูลลูกหนี้ (ใบประวัติลูกค้า)"), 12, 11)
   Node.Expanded = False
      
      Set Node = trvMain.Nodes.Add(ROOT_TREE & " 1-0", tvwChild, ROOT_TREE & " 1-0-2", MapText("1.2. รายงานวิเคราะห์อายุหนี้ ตามลูกหนี้"), 12, 11)
      Node.Expanded = False
      
      Set Node = trvMain.Nodes.Add(ROOT_TREE & " 1-0", tvwChild, ROOT_TREE & " 1-0-2-1", MapText("1.2.1. รายงานวิเคราะห์อายุหนี้ ตามลูกหนี้ ตามวันที่ขาย"), 12, 11)
      Node.Expanded = False
   
      Set Node = trvMain.Nodes.Add(ROOT_TREE & " 1-0", tvwChild, ROOT_TREE & " 1-0-4", MapText("1.4. รายงานค่าเสื่อมทรัพท์สิน"), 12, 11)
      Node.Expanded = False
   
'      Set Node = trvMain.Nodes.Add(ROOT_TREE & " 1-0", tvwChild, ROOT_TREE & " 1-0-5", MapText("1.5. รายงานเช็คจ่ายลงวันที่ล่วงหน้า"), 12, 11)
'      Node.Expanded = False
      
      Set Node = trvMain.Nodes.Add(ROOT_TREE & " 1-0", tvwChild, ROOT_TREE & " 1-0-5-1", MapText("1.5. รายงานเช็คจ่ายลงวันที่ล่วงหน้า (ตามรายชื่อเจ้าหนี้)"), 12, 11)
      Node.Expanded = False
      
      Set Node = trvMain.Nodes.Add(ROOT_TREE & " 1-0", tvwChild, ROOT_TREE & " 1-0-5-2", MapText("1.5.1 รายงานเช็คจ่ายลงวันที่ล่วงหน้า (แยกตามช่วงวันที่) "), 12, 11)
      Node.Expanded = False
      
      Set Node = trvMain.Nodes.Add(ROOT_TREE & " 1-0", tvwChild, ROOT_TREE & " 1-0-5-3", MapText("1.5.2 รายงานเช็คจ่ายลงวันที่ล่วงหน้า (แยกตามวันที่เช็ค) "), 12, 11)
      Node.Expanded = False
      
      Set Node = trvMain.Nodes.Add(ROOT_TREE & " 1-0", tvwChild, ROOT_TREE & " 1-0-5-4", MapText("1.5.3 รายงานเช็คจ่ายลงวันที่ล่วงหน้า (แยกตามวันที่เช็ค รหัสบัญชี) "), 12, 11)
      Node.Expanded = False
      
      Set Node = trvMain.Nodes.Add(ROOT_TREE & " 1-0", tvwChild, ROOT_TREE & " 1-0-6", MapText("1.6. รายงานวิเคราะห์อายุหนี้ ตามพนักงานขาย"), 12, 11)
      Node.Expanded = False
      
      Set Node = trvMain.Nodes.Add(ROOT_TREE & " 1-0", tvwChild, ROOT_TREE & " 1-0-6-1", MapText("1.6.1 รายงานวิเคราะห์อายุหนี้ ตามพนักงานขาย จากวันที่ขาย"), 12, 11)
      Node.Expanded = False
      
      Set Node = trvMain.Nodes.Add(ROOT_TREE & " 1-0", tvwChild, ROOT_TREE & " 1-0-6-2", MapText("1.6.2 รายงานวิเคราะห์อายุหนี้ ตามพนักงานขายตามเดือนดิว"), 12, 11)
      Node.Expanded = False
      
      Set Node = trvMain.Nodes.Add(ROOT_TREE & " 1-0", tvwChild, ROOT_TREE & " 1-0-6-3", MapText("1.6.3 รายงานวิเคราะห์อายุหนี้ ตามพนักงานขายตามเดือนดิว"), 12, 11)
      Node.Expanded = False
      
      Set Node = trvMain.Nodes.Add(ROOT_TREE & " 1-0", tvwChild, ROOT_TREE & " 1-0-6-4", MapText("1.6.4 รายงานวิเคราะห์อายุหนี้ ประมาณการรับ"), 12, 11)
      Node.Expanded = False
      
      Set Node = trvMain.Nodes.Add(ROOT_TREE & " 1-0", tvwChild, ROOT_TREE & " 1-0-6-5", MapText("1.6.5 แบบฟอร์มขออนุมัติขาย"), 12, 11)
      Node.Expanded = False
      
      Set Node = trvMain.Nodes.Add(ROOT_TREE & " 1-0", tvwChild, ROOT_TREE & " 1-0-7", MapText("1.7. รายงานสถานะลูกหนี้ช่วงวันที่"), 12, 11)
      Node.Expanded = False
      
     Set Node = trvMain.Nodes.Add(ROOT_TREE & " 1-0", tvwChild, ROOT_TREE & " 1-0-7-1", MapText("1.7.1 รายงานสถานะลูกหนี้ประจำงวด เรียงตามกลุ่มลูกหนี้ เคลื่อนไหว"), 12, 11)
      Node.Expanded = False                 ' copy 2-5-2
      
      Set Node = trvMain.Nodes.Add(ROOT_TREE & " 1-0", tvwChild, ROOT_TREE & " 1-0-8", MapText("1.8. รายงานยอดขาย ตามลูกหนี้ งวด แจงแจง เอกสาร"), 12, 11)
      Node.Expanded = False
      
      Set Node = trvMain.Nodes.Add(ROOT_TREE & " 1-0", tvwChild, ROOT_TREE & " 1-0-8-1", MapText("1.8.1. รายงานยอดขาย ตามลูกหนี้ แนวตั้ง"), 12, 11)
      Node.Expanded = False
      
      Set Node = trvMain.Nodes.Add(ROOT_TREE & " 1-0", tvwChild, ROOT_TREE & " 1-0-9", MapText("1.9. รายงานเปรียบเทียบยอดรับชำระ กับ ยอดขาย"), 12, 11)
      Node.Expanded = False
      
      Set Node = trvMain.Nodes.Add(ROOT_TREE & " 1-0", tvwChild, ROOT_TREE & " 1-0-10", MapText("1.10. รายงานยอดขาย ลูกค้า สินค้า แยกของแถม ตามเดือนปี"), 12, 11)
      Node.Expanded = False
      
      Set Node = trvMain.Nodes.Add(ROOT_TREE & " 1-0", tvwChild, ROOT_TREE & " 1-0-10-1", MapText("1.10.1. รายงานยอดขาย พนักงานขาย ลูกค้า สินค้า แยกของแถม ตามเดือนปี"), 12, 11)
      Node.Expanded = False
      
      Set Node = trvMain.Nodes.Add(ROOT_TREE & " 1-0", tvwChild, ROOT_TREE & " 1-0-10-2", MapText("1.10.2. รายงานยอดขาย พนักงานขาย ลูกค้า สินค้า แยกของแถม แนวตั้ง"), 12, 11)
      Node.Expanded = False
      
      Set Node = trvMain.Nodes.Add(ROOT_TREE & " 1-0", tvwChild, ROOT_TREE & " 1-0-10-3", MapText("1.10.3. รายงานยอดขาย พนักงานขาย สินค้า ลูกค้า แยกของแถม ตามเดือนปี"), 12, 11)
      Node.Expanded = False
      
      Set Node = trvMain.Nodes.Add(ROOT_TREE & " 1-0", tvwChild, ROOT_TREE & " 1-0-10-4", MapText("1.10.4. รายงานยอดขาย สินค้า ลูกค้า แยกของแถม ตามเดือนปี"), 12, 11)
      Node.Expanded = False
      
      Set Node = trvMain.Nodes.Add(ROOT_TREE & " 1-0", tvwChild, ROOT_TREE & " 1-0-12", MapText("1.12. รายงานแพ็คตั๋ว"), 12, 11)
      Node.Expanded = False
      
      Set Node = trvMain.Nodes.Add(ROOT_TREE & " 1-0", tvwChild, ROOT_TREE & " 1-0-13", MapText("1.12.1. รายงานแพ็คตั๋วธนาคาร"), 12, 11)
      Node.Expanded = False
      
      Set Node = trvMain.Nodes.Add(ROOT_TREE & " 1-0", tvwChild, ROOT_TREE & " 1-0-15-1", MapText("1.15.1 รายงานเช็ครับลงวันที่ล่วงหน้า"), 12, 11)
      Node.Expanded = False
      
      Set Node = trvMain.Nodes.Add(ROOT_TREE & " 1-0", tvwChild, ROOT_TREE & " 1-0-16", MapText("1.16 รายงานเปรียบเทียบลูกหนี้คงเหลือ"), 12, 11)
      Node.Expanded = False
      
      Set Node = trvMain.Nodes.Add(ROOT_TREE & " 1-0", tvwChild, ROOT_TREE & " 1-0-17", MapText("1.17 รายงานรับชำระ"), 12, 11)
      Node.Expanded = False
      
      Set Node = trvMain.Nodes.Add(ROOT_TREE & " 1-0", tvwChild, ROOT_TREE & " 1-0-18", MapText("1.18 รายงานจ่าย Promotion แสดงตามวันที่"), 12, 11)
      Node.Expanded = False
      
      Set Node = trvMain.Nodes.Add(ROOT_TREE & " 1-0", tvwChild, ROOT_TREE & " 1-0-19", MapText("1.19 รายงานจ่าย Promotion แยกตาม พนักงาน ลูกค้า สินค้า"), 12, 11)
      Node.Expanded = False
      
      Set Node = trvMain.Nodes.Add(ROOT_TREE & " 1-0", tvwChild, ROOT_TREE & " 1-0-20", MapText("1.20 รายงานจ่าย Promotion แยกตาม สินค้า ลูกค้า"), 12, 11)
      Node.Expanded = False
      
      Set Node = trvMain.Nodes.Add(ROOT_TREE & " 1-0", tvwChild, ROOT_TREE & " 1-0-21", MapText("1.21 รายงาน Monthly sales report sales out of MGP"), 12, 11)
      Node.Expanded = False
      
      Set Node = trvMain.Nodes.Add(ROOT_TREE, tvwChild, ROOT_TREE & " 2-0", MapText("2. ข้อมูลหนี้สิน"), 3, 3)
      Node.Expanded = False

      Set Node = trvMain.Nodes.Add(ROOT_TREE & " 2-0", tvwChild, ROOT_TREE & " 2-0-1", MapText("2.1. รายงานข้อมูลเจ้าหนี้"), 12, 11)
      Node.Expanded = False

      Set Node = trvMain.Nodes.Add(ROOT_TREE & " 2-0", tvwChild, ROOT_TREE & " 2-0-2", MapText("2.2. รายงานวิเคราะห์อายุหนี้ ตามเจ้าหนี้"), 12, 11)
      Node.Expanded = False
      
      Set Node = trvMain.Nodes.Add(ROOT_TREE & " 2-0", tvwChild, ROOT_TREE & " 2-0-2-1", MapText("2.2.1 รายงานวิเคราะห์อายุหนี้ ตามเจ้าหนี้ ประเภทข้อมูล"), 12, 11)
      Node.Expanded = False

      Set Node = trvMain.Nodes.Add(ROOT_TREE & " 2-0", tvwChild, ROOT_TREE & " 2-0-3", MapText("2.3. รายงานสรุปอายุหนี้ ตามเจ้าหนี้"), 12, 11)
      Node.Expanded = False
      
      Set Node = trvMain.Nodes.Add(ROOT_TREE & " 2-0", tvwChild, ROOT_TREE & " 2-0-3-1", MapText("2.3.1 รายงานสรุปอายุหนี้ ตามเจ้าหนี้ แยกตามช่วงวันที่ซื้อ"), 12, 11)
      Node.Expanded = False
      
      Set Node = trvMain.Nodes.Add(ROOT_TREE & " 2-0", tvwChild, ROOT_TREE & " 2-0-3-2", MapText("2.3.2 รายงานสรุปอายุหนี้ ตามเจ้าหนี้ แยกตามช่วงวันที่ซื้อ ตามกลุ่มเจ้าหนี้"), 12, 11)
      Node.Expanded = False
      
      Set Node = trvMain.Nodes.Add(ROOT_TREE & " 2-0", tvwChild, ROOT_TREE & " 2-0-3-3", MapText("2.3.3 รายงานสรุปอายุหนี้ ตามเจ้าหนี้ แยกตามช่วงวันที่ซื้อ ตามกลุ่มเจ้าหนี้ ประเภทย่อย"), 12, 11)
      Node.Expanded = False
      
      Set Node = trvMain.Nodes.Add(ROOT_TREE & " 2-0", tvwChild, ROOT_TREE & " 2-0-4", MapText("2.4. รายงานยอดเจ้าหนี้ตามบิล"), 12, 11)
      Node.Expanded = False
      
      Set Node = trvMain.Nodes.Add(ROOT_TREE & " 2-0", tvwChild, ROOT_TREE & " 2-0-4-1", MapText("2.4.1 รายงานยอดเจ้าหนี้ตามบิล แยกตามเดือน DUE"), 12, 11)
      Node.Expanded = False

      Set Node = trvMain.Nodes.Add(ROOT_TREE & " 2-0", tvwChild, ROOT_TREE & " 2-0-5", MapText("2.5. รายงานซื้อวัตถุดิบตามประเภทวัตถุดิบ (RM008)"), 12, 11)
      Node.Expanded = False
   
      Set Node = trvMain.Nodes.Add(ROOT_TREE & " 2-0", tvwChild, ROOT_TREE & " 2-0-5-1", MapText("2.5. รายงานซื้อวัตถุดิบตามประเภทวัตถุดิบ (RM008-1)"), 12, 11)
      Node.Expanded = False
      
      Set Node = trvMain.Nodes.Add(ROOT_TREE & " 2-0", tvwChild, ROOT_TREE & " 2-0-6", MapText("2.5. รายงานสถานะเจ้าหนี้ประจำงวด"), 12, 11)
      Node.Expanded = False
      
      Set Node = trvMain.Nodes.Add(ROOT_TREE & " 2-0", tvwChild, ROOT_TREE & " 2-0-6-1", MapText("2.5.1 รายงานสถานะเจ้าหนี้ประจำงวด เรียงตามกลุ่มเจ้าหนี้"), 12, 11)
      Node.Expanded = False
      
      Set Node = trvMain.Nodes.Add(ROOT_TREE & " 2-0", tvwChild, ROOT_TREE & " 2-0-6-2", MapText("2.5.2 รายงานสถานะเจ้าหนี้ประจำงวด เรียงตามกลุ่มเจ้าหนี้ เคลื่อนไหว"), 12, 11)
      Node.Expanded = False
      
      Set Node = trvMain.Nodes.Add(ROOT_TREE & " 2-0", tvwChild, ROOT_TREE & " 2-0-7", MapText("2.6. รายงานสถานะเจ้าหนี้ช่วงวันที่"), 12, 11)
      Node.Expanded = False
      
      Set Node = trvMain.Nodes.Add(ROOT_TREE & " 2-0", tvwChild, ROOT_TREE & " 2-0-8", MapText("2.7. รายงานใบรับสินค้าเรียงตามซัพพลายเออร์ แหล่งจำหน่าย"), 12, 11)
      Node.Expanded = False
      
   Set Node = trvMain.Nodes.Add(ROOT_TREE, tvwChild, ROOT_TREE & " 3-1", MapText("3. ข้อมูลทุน"), 3, 3)
   Node.Expanded = False
      
   Set Node = trvMain.Nodes.Add(ROOT_TREE, tvwChild, ROOT_TREE & " 4-0", MapText("4. ข้อมูลรายรับ"), 3, 3)
   Node.Expanded = False

   Set Node = trvMain.Nodes.Add(ROOT_TREE, tvwChild, ROOT_TREE & " 5-0", MapText("5. ข้อมูลรายจ่าย"), 3, 3)
   Node.Expanded = False

   Set Node = trvMain.Nodes.Add(, tvwFirst, ROOT_TREE & "-A", MapText("พิมพ์เอกสาร"), 8)
   Node.Expanded = True
   Node.Selected = True

   Set Node = trvMain.Nodes.Add(ROOT_TREE & "-A", tvwChild, ROOT_TREE & "-A" & " 5-0", MapText("1. ใบสำคัญจ่ายอื่น ๆ"), 12, 11)
   Node.Expanded = False

   Set Node = trvMain.Nodes.Add(ROOT_TREE & "-A", tvwChild, ROOT_TREE & "-A" & " 5-1", MapText("2. ใบสำคัญจ่ายชำระเจ้าหนี้"), 12, 11)
   Node.Expanded = False

   Set Node = trvMain.Nodes.Add(ROOT_TREE & "-A", tvwChild, ROOT_TREE & "-A" & " 5-2", MapText("3. ใบสำคัญการโอน"), 12, 11)
   Node.Expanded = False

   Set Node = trvMain.Nodes.Add(ROOT_TREE & "-A", tvwChild, ROOT_TREE & "-A" & " 5-3", MapText("4. ใบสำคัญรับอื่น ๆ"), 12, 11)
   Node.Expanded = False

   Set Node = trvMain.Nodes.Add(ROOT_TREE & "-A", tvwChild, ROOT_TREE & "-A" & " 5-4", MapText("5. ใบสำคัญรับชำระลูกหนี้"), 12, 11)
   Node.Expanded = False
   
    'commission
       Set Node = trvMain.Nodes.Add(ROOT_TREE, tvwChild, ROOT_TREE & " 6-0", MapText("6. commission "), 3, 3)
      Node.Expanded = False

      Set Node = trvMain.Nodes.Add(ROOT_TREE & " 6-0", tvwChild, ROOT_TREE & " 6-0-1", MapText("6.1. commission ขาย"), 12, 11)
      Node.Expanded = False
      
      Set Node = trvMain.Nodes.Add(ROOT_TREE & " 6-0", tvwChild, ROOT_TREE & " 6-0-2", MapText("6.2. commission เก็บเงิน"), 12, 11)
      Node.Expanded = False

      Set Node = trvMain.Nodes.Add(ROOT_TREE & " 6-0", tvwChild, ROOT_TREE & " 6-0-3", MapText("6.3. Incentive"), 12, 11)
      Node.Expanded = False
      
      Set Node = trvMain.Nodes.Add(ROOT_TREE & " 6-0", tvwChild, ROOT_TREE & " 6-0-4", MapText("6.4. สรุปค่าคอมมิชชั่น"), 12, 11)
      Node.Expanded = False

      Set Node = trvMain.Nodes.Add(ROOT_TREE & " 6-0", tvwChild, ROOT_TREE & " 6-0-5", MapText("6.5. ยอดคอมมิชชั่นขาย"), 12, 11)
      Node.Expanded = False
      
     Set Node = trvMain.Nodes.Add(ROOT_TREE & " 6-0", tvwChild, ROOT_TREE & " 6-0-9", MapText("6.6. ใบเช็คยอดคอมมิชชั่นขาย จาก Express"), 12, 11)
      Node.Expanded = False
      
      Set Node = trvMain.Nodes.Add(ROOT_TREE & " 6-0", tvwChild, ROOT_TREE & " 6-0-6", MapText("6.7. ยอดเฉพาะสินค้าที่ไม่คิด Commercial #1"), 12, 11)
      Node.Expanded = False

      Set Node = trvMain.Nodes.Add(ROOT_TREE & " 6-0", tvwChild, ROOT_TREE & " 6-0-7", MapText("6.8. ยอดสินค้า เฉพาะพนักงานขายที่คิดค่า Incentive"), 12, 11)
      Node.Expanded = False
      
      Set Node = trvMain.Nodes.Add(ROOT_TREE & " 6-0", tvwChild, ROOT_TREE & " 6-0-8", MapText("6.9. ประวัติลูกค้าย้อนหลัง"), 12, 11)
      Node.Expanded = False

      Set Node = trvMain.Nodes.Add(ROOT_TREE & " 6-0", tvwChild, ROOT_TREE & " 6-0-10", MapText("6.10. รายงานส่วนลดเรียงตามลูกค้า"), 12, 11)
      Node.Expanded = False
      
      Set Node = trvMain.Nodes.Add(ROOT_TREE & " 6-0", tvwChild, ROOT_TREE & " 6-0-11", MapText("6.11. รายงานจำนวนสินค้าแยกตามเขต ตามลูกค้า"), 12, 11)
      Node.Expanded = False
      
      Set Node = trvMain.Nodes.Add(ROOT_TREE & " 6-0", tvwChild, ROOT_TREE & " 6-0-12", MapText("6.12. รายงานจำนวนสินค้าแยกตามเขต ตามสินค้า"), 12, 11)
      Node.Expanded = False
      
      Set Node = trvMain.Nodes.Add(ROOT_TREE & " 6-0", tvwChild, ROOT_TREE & " 6-0-13", MapText("6.13. ยอดสินค้าราคาเฉลี่ยเรียงตามเซลล์"), 12, 11)
      Node.Expanded = False
      
      Set Node = trvMain.Nodes.Add(ROOT_TREE & " 6-0", tvwChild, ROOT_TREE & " 6-0-14", MapText("6.14. ใบสรุปรายการจ่ายจริง"), 12, 11)
      Node.Expanded = False
      
      
   Set Node = trvMain.Nodes.Add(, tvwFirst, ROOT_TREE & "-B", MapText("พิมพ์เช็ค"), 8)
   Node.Expanded = True
   Node.Selected = True

   Set Node = trvMain.Nodes.Add(ROOT_TREE & "-B", tvwChild, ROOT_TREE & "-B" & " 1-1", MapText("1. เช็คจ่าย"), 12, 11)
   Node.Expanded = False

   Set Node = trvMain.Nodes.Add(, tvwFirst, ROOT_TREE & "-C", MapText("รายงานงบ"), 8)
   Node.Expanded = True
   Node.Selected = True

   Set Node = trvMain.Nodes.Add(ROOT_TREE & "-C", tvwChild, ROOT_TREE & "-C" & " 1-1", MapText("1. รายงานงบทดลอง"), 12, 11)
   Node.Expanded = False
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
   If m_MustAsk Then
      glbErrorLog.LocalErrorMsg = MapText("ท่านต้องการออกจากโปรแกรมใช่หรือไม่")
      If glbErrorLog.AskMessage = vbYes Then
         Cancel = False
      Else
         Cancel = True
      End If
   Else
      Cancel = False
   End If
End Sub

Private Sub FillReportInput(R As CReportInterface)
Dim C As CReportControl

   For Each C In m_ReportControls
      If (C.ControlType = "C") Then
         If C.Param1 <> "" Then
            Call R.AddParam(m_Combos(C.ControlIndex).Text, C.Param1)
         End If
         
         If C.Param2 <> "" Then
            Call R.AddParam(m_Combos(C.ControlIndex).ItemData(Minus2Zero(m_Combos(C.ControlIndex).ListIndex)), C.Param2)
         End If
      End If
   
      If (C.ControlType = "CB") Then
         If C.Param1 <> "" Then
            Call R.AddParam(m_Combos(C.ControlIndex).Text, C.Param1)
         End If
         
         If C.Param2 <> "" Then
            Call R.AddParam(m_Combos(C.ControlIndex).ListIndex, C.Param2)
         End If
      End If
      
      If (C.ControlType = "T") Then
         If C.Param1 <> "" Then
            Call R.AddParam(m_Texts(C.ControlIndex).Text, C.Param1)
         End If
         
         If C.Param2 <> "" Then
            Call R.AddParam(m_Texts(C.ControlIndex).Text, C.Param2)
         End If
      End If
   
      If (C.ControlType = "CH") Then
         If C.Param1 <> "" Then
            Call R.AddParam(Check2Flag(m_CheckBoxes(C.ControlIndex).Value), C.Param1)
         End If
         
         If C.Param2 <> "" Then
            Call R.AddParam(Check2Flag(m_CheckBoxes(C.ControlIndex).Value), C.Param2)
         End If
      End If
      
      If (C.ControlType = "D") Then
         If C.Param1 <> "" Then
            Call R.AddParam(m_Dates(C.ControlIndex).ShowDate, C.Param1)
         End If
         
         If C.Param2 <> "" Then
            If m_Dates(C.ControlIndex).ShowDate <= 0 Then
               If C.Param2 = "TO_DATE" Then
                  m_Dates(C.ControlIndex).ShowDate = -1
               ElseIf C.Param2 = "FROM_DATE" Then
                  m_Dates(C.ControlIndex).ShowDate = -2
               End If
            End If
            If C.Param2 = "FROM_DATE" Then
               m_FromDate = m_Dates(C.ControlIndex).ShowDate
            ElseIf C.Param2 = "TO_DATE" Then
               m_ToDate = m_Dates(C.ControlIndex).ShowDate
            End If
            Call R.AddParam(m_Dates(C.ControlIndex).ShowDate, C.Param2)
         End If
      End If
   
   Next C
End Sub

Private Function VerifyReportInput() As Boolean
Dim C As CReportControl

   VerifyReportInput = False
   For Each C In m_ReportControls
      If (C.ControlType = "C") Then
         If Not VerifyCombo(Nothing, m_Combos(C.ControlIndex), C.AllowNull) Then
            Exit Function
         End If
      End If
   
      If (C.ControlType = "T") Then
         If Not VerifyTextControl(Nothing, m_Texts(C.ControlIndex), C.AllowNull) Then
            Exit Function
         End If
      End If
   
      If (C.ControlType = "D") Then
         If Not VerifyDate(Nothing, m_Dates(C.ControlIndex), C.AllowNull) Then
            Exit Function
         End If
      End If
   Next C
   VerifyReportInput = True
End Function
Private Function VerifyReportInputDrCr() As Boolean
Dim C As CReportControl

   VerifyReportInputDrCr = False

   For Each C In m_ReportControls
         If (C.ControlType = "D") Then
            If Not VerifyDate(Nothing, m_Dates(C.ControlIndex), C.AllowNull) Then
               Exit Function
            End If
         End If
   Next C
   VerifyReportInputDrCr = True
End Function
Private Sub LoadControl(ControlType As String, Width As Long, NullAllow As Boolean, TextMsg As String, Optional ComboLoadID As Long = -1, Optional Param1 As String = "", Optional Param2 As String = "", Optional OldLine As Boolean = False, Optional ToolTipText As String)
Dim CboIdx As Long
Dim TxtIdx As Long
Dim LblIdx2 As Long
Dim DateIdx As Long
Dim LblIdx As Long
Dim LkupIdx As Long
Dim C As CReportControl
Dim ChkIdx As Long

   CboIdx = m_Combos.Count + 1
   TxtIdx = m_Texts.Count + 1
   DateIdx = m_Dates.Count + 1
   LblIdx = m_Labels.Count + 1
   ChkIdx = m_CheckBoxes.Count + 1

   Set C = New CReportControl
   If ControlType = "L" Then
      Load lblGeneric(LblIdx)
      Call m_Labels.Add(lblGeneric(LblIdx))
      C.ControlIndex = LblIdx
      lblGeneric(LblIdx).ToolTipText = ToolTipText
   ElseIf ControlType = "C" Then
      Load cboGeneric(CboIdx)
      Call m_Combos.Add(cboGeneric(CboIdx))
      C.ControlIndex = CboIdx
      C.OldLine = OldLine
   ElseIf ControlType = "CB" Then
      Load cboGeneric(CboIdx)
      Call m_Combos.Add(cboGeneric(CboIdx))
      C.ControlIndex = CboIdx
      C.OldLine = OldLine
   ElseIf ControlType = "T" Then
      Load txtGeneric(TxtIdx)
      Call m_Texts.Add(txtGeneric(TxtIdx))
      C.ControlIndex = TxtIdx
      C.OldLine = OldLine
 
   ElseIf ControlType = "D" Then
      Load uctlGenericDate(DateIdx)
      Call m_Dates.Add(uctlGenericDate(DateIdx))
      C.ControlIndex = DateIdx
      
        If Param1 = "FROM_DOC_DATE" Or Param1 = "FROM_DATE" Then
         If m_FromDate > 0 Then
            uctlGenericDate(DateIdx).ShowDate = m_FromDate
         Else
            Call GetFirstLastDate(Now, m_FromDate, m_ToDate)
            uctlGenericDate(DateIdx).ShowDate = m_FromDate
         End If
      ElseIf Param1 = "TO_DOC_DATE" Or Param1 = "TO_DATE" Then
         If m_FromDate > 0 Then
            uctlGenericDate(DateIdx).ShowDate = m_ToDate
         Else
            Call GetFirstLastDate(Now, m_FromDate, m_ToDate)
            uctlGenericDate(DateIdx).ShowDate = m_ToDate
         End If
      ElseIf Param1 = "TO_PAY_DATE" Or Param1 = "PRINT_DATE" Or Param1 = "SENT_DATE" Then
          uctlGenericDate(DateIdx).ShowDate = Now
      ElseIf Param1 = "FROM_CHECK_DATE" Then
          uctlGenericDate(DateIdx).ShowDate = Now + 1
      End If
 ElseIf ControlType = "LU" Then
'         Load uctlGLACC(LkupIdx)
'         Call m_TextLookups.Add(uctlGLACC(LkupIdx))
'         C.ControlIndex = LkupIdx

   ElseIf ControlType = "CH" Then
      Load chkGeneric(ChkIdx)
      Call m_CheckBoxes.Add(chkGeneric(ChkIdx))
      Call InitCheckBox(chkGeneric(ChkIdx), TextMsg)
      C.ControlIndex = ChkIdx
      C.OldLine = OldLine
   End If

   C.AllowNull = NullAllow
   C.ControlType = ControlType
   C.Width = Width
   C.TextMsg = TextMsg
   C.Param1 = Param2
   C.Param2 = Param1
   C.ComboLoadID = ComboLoadID
   Call m_ReportControls.Add(C)
   Set C = Nothing
End Sub
Private Sub UnloadAllControl()
Dim i As Long
Dim j As Long

   i = m_Labels.Count
   While i > 0
      Call Unload(m_Labels(i))
      Call m_Labels.Remove(i)
      i = i - 1
   Wend
   
   i = m_Texts.Count
   While i > 0
      Call Unload(m_Texts(i))
      Call m_Texts.Remove(i)
      i = i - 1
   Wend

   i = m_Dates.Count
   While i > 0
      Call Unload(m_Dates(i))
      Call m_Dates.Remove(i)
      i = i - 1
   Wend

   i = m_Combos.Count
   While i > 0
      Call Unload(m_Combos(i))
      Call m_Combos.Remove(i)
      i = i - 1
   Wend
   
'   I = m_TextLookups.Count
'   While I > 0
'      Call Unload(m_TextLookups(I))
'      Call m_TextLookups.Remove(I)
'      I = I - 1
'   Wend
   
   i = m_CheckBoxes.Count
   While i > 0
      Call Unload(m_CheckBoxes(i))
      Call m_CheckBoxes.Remove(i)
      i = i - 1
   Wend
   
   Set m_ReportControls = Nothing
   Set m_ReportControls = New Collection
End Sub
Private Sub ShowControl()
Dim PrevTop As Long
Dim PrevLeft As Long
Dim PrevWidth As Long
Dim CurTop As Long
Dim CurLeft As Long
Dim CurWidth As Long
Dim C As CReportControl

   PrevTop = uctlGenericDate(0).Top
   PrevLeft = uctlGenericDate(0).Left
   PrevWidth = uctlGenericDate(0).Width
   
   For Each C In m_ReportControls
      If (C.ControlType = "C") Or (C.ControlType = "CB") Or (C.ControlType = "D") Or (C.ControlType = "T") Or (C.ControlType = "LU") Or (C.ControlType = "CH") Then
         If C.ControlType = "C" Or (C.ControlType = "CB") Then
            If C.OldLine Then
               m_Combos(C.ControlIndex).Left = PrevLeft + PrevWidth + 20
               m_Combos(C.ControlIndex).Top = PrevTop - m_Combos(C.ControlIndex - 1).Height
            Else
               m_Combos(C.ControlIndex).Left = PrevLeft
               m_Combos(C.ControlIndex).Top = PrevTop
            End If
            m_Combos(C.ControlIndex).Width = C.Width
            Call InitCombo(m_Combos(C.ControlIndex))
            m_Combos(C.ControlIndex).Visible = True
            
            CurTop = PrevTop
            CurLeft = PrevLeft
            CurWidth = PrevWidth
            
            PrevTop = m_Combos(C.ControlIndex).Top + m_Combos(C.ControlIndex).Height
            If C.OldLine Then
               PrevLeft = m_Combos(C.ControlIndex).Left - CurWidth - 20
            Else
               PrevLeft = m_Combos(C.ControlIndex).Left
            End If
            PrevWidth = C.Width
         ElseIf C.ControlType = "D" Then
            m_Dates(C.ControlIndex).Left = PrevLeft
            m_Dates(C.ControlIndex).Top = PrevTop
            m_Dates(C.ControlIndex).Width = C.Width
            m_Dates(C.ControlIndex).Visible = True
            
            CurTop = PrevTop
            CurLeft = PrevLeft
            CurWidth = PrevWidth
         
            PrevTop = m_Dates(C.ControlIndex).Top + m_Dates(C.ControlIndex).Height
            PrevLeft = m_Dates(C.ControlIndex).Left
            PrevWidth = C.Width
         ElseIf C.ControlType = "T" Then
            If C.OldLine Then
               m_Texts(C.ControlIndex).Left = PrevLeft + PrevWidth + 20
               m_Texts(C.ControlIndex).Top = PrevTop - txtGeneric(0).Height
               Call m_Texts(C.ControlIndex).SetTextLenType(TEXT_STRING, glbSetting.DESC_TYPE)
               m_Texts(C.ControlIndex).Visible = True
               m_Texts(C.ControlIndex).Width = C.Width
            Else
               m_Texts(C.ControlIndex).Left = PrevLeft
               m_Texts(C.ControlIndex).Top = PrevTop
               m_Texts(C.ControlIndex).Width = C.Width
               Call m_Texts(C.ControlIndex).SetTextLenType(TEXT_STRING, glbSetting.DESC_TYPE)
               m_Texts(C.ControlIndex).Visible = True
                              
               CurTop = PrevTop
               CurLeft = PrevLeft
               CurWidth = PrevWidth
               
               PrevTop = m_Texts(C.ControlIndex).Top + m_Texts(C.ControlIndex).Height
               PrevLeft = m_Texts(C.ControlIndex).Left
               PrevWidth = C.Width
            End If
         ElseIf C.ControlType = "LU" Then
            m_TextLookups(C.ControlIndex).Left = PrevLeft
            m_TextLookups(C.ControlIndex).Top = PrevTop
            m_TextLookups(C.ControlIndex).Width = C.Width
            m_TextLookups(C.ControlIndex).Visible = True
         
            CurTop = PrevTop
            CurLeft = PrevLeft
            CurWidth = PrevWidth
         
            PrevTop = m_TextLookups(C.ControlIndex).Top + m_TextLookups(C.ControlIndex).Height
            PrevLeft = m_TextLookups(C.ControlIndex).Left
            PrevWidth = C.Width
         ElseIf C.ControlType = "CH" Then
            If C.OldLine Then
               m_CheckBoxes(C.ControlIndex).Left = PrevLeft + PrevWidth + 20
'               m_CheckBoxes(C.ControlIndex).Left = PrevLeft + Len(m_CheckBoxes(C.ControlIndex).Caption) + 20
               m_CheckBoxes(C.ControlIndex).Top = PrevTop + 10 - m_CheckBoxes(C.ControlIndex - 1).Height
            Else
               m_CheckBoxes(C.ControlIndex).Left = PrevLeft
               m_CheckBoxes(C.ControlIndex).Top = PrevTop + 10
            End If

            m_CheckBoxes(C.ControlIndex).Width = C.Width
            m_CheckBoxes(C.ControlIndex).Visible = True
         
            CurTop = PrevTop
            CurLeft = PrevLeft
            CurWidth = PrevWidth
         
            PrevTop = m_CheckBoxes(C.ControlIndex).Top + m_CheckBoxes(C.ControlIndex).Height
             If C.OldLine Then
               PrevLeft = m_CheckBoxes(C.ControlIndex).Left - CurWidth - 20
            Else
               PrevLeft = m_CheckBoxes(C.ControlIndex).Left
            End If
            PrevWidth = C.Width
         End If
      
      Else 'Label
            m_Labels(C.ControlIndex).Left = lblGeneric(0).Left
            m_Labels(C.ControlIndex).Top = CurTop
            m_Labels(C.ControlIndex).Width = C.Width
            Call InitNormalLabel(m_Labels(C.ControlIndex), C.TextMsg)
            m_Labels(C.ControlIndex).Visible = True
      End If
   Next C
End Sub
'Private Sub ShowControl()
'Dim PrevTop As Long
'Dim PrevLeft As Long
'Dim PrevWidth As Long
'Dim CurTop As Long
'Dim CurLeft As Long
'Dim CurWidth As Long
'Dim C As CReportControl
'
'   PrevTop = uctlGenericDate(0).Top
'   PrevLeft = uctlGenericDate(0).Left
'   PrevWidth = uctlGenericDate(0).Width
'
'   For Each C In m_ReportControls
'      If (C.ControlType = "C") Or (C.ControlType = "CB") Or (C.ControlType = "D") Or (C.ControlType = "T") Or (C.ControlType = "CH") Or (C.ControlType = "LU") Then
'         If C.ControlType = "C" Then
'            m_Combos(C.ControlIndex).Left = PrevLeft
'            m_Combos(C.ControlIndex).Top = PrevTop
'            m_Combos(C.ControlIndex).Width = C.Width
'            Call InitCombo(m_Combos(C.ControlIndex))
'            m_Combos(C.ControlIndex).Visible = True
'
'            CurTop = PrevTop
'            CurLeft = PrevLeft
'            CurWidth = PrevWidth
'
'            PrevTop = m_Combos(C.ControlIndex).Top + m_Combos(C.ControlIndex).Height
'            PrevLeft = m_Combos(C.ControlIndex).Left
'            PrevWidth = C.Width
'         ElseIf C.ControlType = "CB" Then
'            m_Combos(C.ControlIndex).Left = PrevLeft
'            m_Combos(C.ControlIndex).Top = PrevTop
'            m_Combos(C.ControlIndex).Width = C.Width
'            Call InitCombo(m_Combos(C.ControlIndex))
'            m_Combos(C.ControlIndex).Visible = True
'
'            CurTop = PrevTop
'            CurLeft = PrevLeft
'            CurWidth = PrevWidth
'
'            PrevTop = m_Combos(C.ControlIndex).Top + m_Combos(C.ControlIndex).Height
'            PrevLeft = m_Combos(C.ControlIndex).Left
'            PrevWidth = C.Width
'          ElseIf C.ControlType = "CH" Then
'           m_CheckBoxes(C.ControlIndex).Left = PrevLeft
'           m_CheckBoxes(C.ControlIndex).Top = PrevTop
'           m_CheckBoxes(C.ControlIndex).Width = C.Width
'           Call InitCheckBox(m_CheckBoxes(C.ControlIndex), C.TextMsg)
'           m_CheckBoxes(C.ControlIndex).Visible = True
'
'            CurTop = PrevTop
'            CurLeft = PrevLeft
'            CurWidth = PrevWidth
'
'            PrevTop = m_CheckBoxes(C.ControlIndex).Top + m_CheckBoxes(C.ControlIndex).Height
'            PrevLeft = m_CheckBoxes(C.ControlIndex).Left
'            PrevWidth = C.Width
'         ElseIf C.ControlType = "D" Then
'            m_Dates(C.ControlIndex).Left = PrevLeft
'            m_Dates(C.ControlIndex).Top = PrevTop
'            m_Dates(C.ControlIndex).Width = C.Width
'            m_Dates(C.ControlIndex).Visible = True
'
'            CurTop = PrevTop
'            CurLeft = PrevLeft
'            CurWidth = PrevWidth
'
'            PrevTop = m_Dates(C.ControlIndex).Top + m_Dates(C.ControlIndex).Height
'            PrevLeft = m_Dates(C.ControlIndex).Left
'            PrevWidth = C.Width
'         ElseIf C.ControlType = "T" Then
'            m_Texts(C.ControlIndex).Left = PrevLeft
'            m_Texts(C.ControlIndex).Left = PrevLeft
'            m_Texts(C.ControlIndex).Top = PrevTop
'            m_Texts(C.ControlIndex).Width = C.Width
'
'
'            Call m_Texts(C.ControlIndex).SetTextLenType(TEXT_STRING, glbSetting.DESC_TYPE)
'            m_Texts(C.ControlIndex).Visible = True
'
'            CurTop = PrevTop
'            CurLeft = PrevLeft
'            CurWidth = PrevWidth
'
'            PrevTop = m_Texts(C.ControlIndex).Top + m_Texts(C.ControlIndex).Height
'            PrevLeft = m_Texts(C.ControlIndex).Left
'            PrevWidth = C.Width
'
'         ElseIf C.ControlType = "LU" Then
'            m_TextLookups(C.ControlIndex).Top = PrevTop
'            m_TextLookups(C.ControlIndex).Width = C.Width
'            m_TextLookups(C.ControlIndex).Visible = True
'
'            CurTop = PrevTop
'            CurLeft = PrevLeft
'            CurWidth = PrevWidth
'
'            PrevTop = m_TextLookups(C.ControlIndex).Top + m_TextLookups(C.ControlIndex).Height
'            PrevLeft = m_TextLookups(C.ControlIndex).Left
'            PrevWidth = C.Width
'
'         End If
'   Else 'Label
'
'            m_Labels(C.ControlIndex).Left = lblGeneric(0).Left
'            m_Labels(C.ControlIndex).Top = CurTop
'            m_Labels(C.ControlIndex).Width = C.Width
'
'            ''debug.print C.AllowNull
'            If C.AllowNull Then
'               Call InitNormalLabel(m_Labels(C.ControlIndex), C.TextMsg)
'            Else
'               Call InitNormalLabel(m_Labels(C.ControlIndex), C.TextMsg, RGB(0, 0, 255))
'            End If
'            m_Labels(C.ControlIndex).Visible = True
'
'   End If
'   Next C
'End Sub
Private Sub Form_Unload(Cancel As Integer)
   Set m_ReportControls = Nothing
   Set m_Texts = Nothing
   Set m_Dates = Nothing
   Set m_Labels = Nothing
   Set m_Combos = Nothing
   Set m_TextLookups = Nothing
   Set m_ReportParams = Nothing
   Set m_CheckBoxes = Nothing
   Set m_Rs = Nothing
   Set m_Journals = Nothing                      'Step 3
   Call ReleaseAll
End Sub
Private Sub Label1_Click()

End Sub


Private Sub SSCommand2_Click()
Dim Ap As CAPMas

   Set Ap = New CAPMas
   Ap.SUPCOD = "ล-0001"
   Ap.SUPTYP = "01"
   Call Ap.UpdateSupplierType
   Set Ap = Nothing
End Sub
Private Sub SSCommand3_Click()
Dim Bk As CBkTrn
   Set Bk = New CBkTrn
   'ใครสร้างปุ่มนี้หนะ
   Call glbDaily.StartTransaction
   Call Bk.DeleteAllData
   Call glbDaily.CommitTransaction
   Set Bk = Nothing
   'ใครสร้างปุ่มนี้หนะ จิวเพ่งเห็นวันที่ 26/10/2554 เลย ลบปุ่มในหน้า Design Form ออกก่อน
End Sub

Private Sub Timer1_Timer()
   Timer1.Enabled = False
   
   lblDateTime.Caption = "                                                    "
   lblDateTime.Caption = DateToStringExtEx3(Now)
   Timer1.Enabled = True
End Sub

Private Sub trvMain_NodeClick(ByVal Node As MSComctlLib.Node)
Static LastKey As String
Dim Status As Boolean
Dim ItemCount As Long
Dim QueryFlag As Boolean

   If LastKey = Node.KEY Then
      Exit Sub
   End If

   pnlHeader.Caption = Node.Text
   
   Status = True
   QueryFlag = False

   Call UnloadAllControl
   
   cmdAdd.Visible = False
   
   If Node.KEY = ROOT_TREE & " 1-0-1" Then
      Call InitReport1_0_1
   ElseIf Node.KEY = ROOT_TREE & " 1-0-1-1" Then
      Call InitReport1_0_1_1
   ElseIf Node.KEY = ROOT_TREE & " 1-0-1-2" Then
      Call InitReport1_0_1_2
   ElseIf Node.KEY = ROOT_TREE & " 1-0-2" Then
      Call InitReport1_0_2
   ElseIf Node.KEY = ROOT_TREE & " 1-0-2-1" Then
      Call InitReport1_0_2_1
   ElseIf Node.KEY = ROOT_TREE & " 1-0-4" Then
      Call InitReport1_0_4
'   ElseIf Node.KEY = ROOT_TREE & " 1-0-5" Then
'      Call InitReport1_0_5
   ElseIf Node.KEY = ROOT_TREE & " 1-0-5-1" Then
      Call InitReport1_0_5_1
   ElseIf Node.KEY = ROOT_TREE & " 1-0-5-2" Then
      Call InitReport1_0_5_2
   ElseIf Node.KEY = ROOT_TREE & " 1-0-5-3" Then
      Call InitReport1_0_5_3
   ElseIf Node.KEY = ROOT_TREE & " 1-0-5-4" Then
      Call InitReport1_0_5_4
   ElseIf Node.KEY = ROOT_TREE & " 1-0-6" Then
      Call InitReport1_0_6
   ElseIf Node.KEY = ROOT_TREE & " 1-0-6-1" Then
      Call InitReport1_0_6_1
   ElseIf Node.KEY = ROOT_TREE & " 1-0-6-2" Then
      Call InitReport1_0_6_2
   ElseIf Node.KEY = ROOT_TREE & " 1-0-6-3" Then
      Call InitReport1_0_6_3
   ElseIf Node.KEY = ROOT_TREE & " 1-0-6-4" Then
      Call InitReport1_0_6_4
   ElseIf Node.KEY = ROOT_TREE & " 1-0-6-5" Then
      Call InitReport1_0_6_5
   ElseIf Node.KEY = ROOT_TREE & " 1-0-7" Then
      Call InitReport1_0_7
  ElseIf Node.KEY = ROOT_TREE & " 1-0-7-1" Then
      Call InitReport1_0_7_1
   ElseIf Node.KEY = ROOT_TREE & " 1-0-8" Then
      Call InitReport1_0_8
   ElseIf Node.KEY = ROOT_TREE & " 1-0-8-1" Then
      Call InitReport1_0_8_1
   ElseIf Node.KEY = ROOT_TREE & " 1-0-9" Then
      Call InitReport1_0_9
   ElseIf Node.KEY = ROOT_TREE & " 1-0-10" Then
      Call InitReport1_0_10
   ElseIf Node.KEY = ROOT_TREE & " 1-0-10-1" Then
      Call InitReport1_0_10_1
   ElseIf Node.KEY = ROOT_TREE & " 1-0-10-2" Then
      Call InitReport1_0_10_2
  ElseIf Node.KEY = ROOT_TREE & " 1-0-10-3" Then
      Call InitReport1_0_10_3
   ElseIf Node.KEY = ROOT_TREE & " 1-0-10-4" Then
      Call InitReport1_0_10_4
   ElseIf Node.KEY = ROOT_TREE & " 1-0-11" Then
      Call InitReport1_0_11
   ElseIf Node.KEY = ROOT_TREE & " 1-0-12" Then
      Call InitReport1_0_12
   ElseIf Node.KEY = ROOT_TREE & " 1-0-13" Then
      Call InitReport1_0_12_1
   ElseIf Node.KEY = ROOT_TREE & " 1-0-15-1" Then
      Call InitReport1_0_15_1
   ElseIf Node.KEY = ROOT_TREE & " 1-0-16" Then
      Call InitReport1_0_16
   ElseIf Node.KEY = ROOT_TREE & " 1-0-17" Then
      Call InitReport1_0_17
   ElseIf Node.KEY = ROOT_TREE & " 1-0-18" Then
      Call InitReport1_0_18
   ElseIf Node.KEY = ROOT_TREE & " 1-0-19" Then
      Call InitReport1_0_19
   ElseIf Node.KEY = ROOT_TREE & " 1-0-20" Then
      Call InitReport1_0_20
   ElseIf Node.KEY = ROOT_TREE & " 1-0-21" Then
      Call InitReport1_0_21
   ElseIf Node.KEY = ROOT_TREE & " 2-0-1" Then
      Call InitReport3_1
   ElseIf Node.KEY = ROOT_TREE & " 2-0-2" Then
      Call InitReport3_2
   ElseIf Node.KEY = ROOT_TREE & " 2-0-2-1" Then
      Call InitReport2_0_4
   ElseIf Node.KEY = ROOT_TREE & " 2-0-3" Then
      Call InitReport3_2
   ElseIf Node.KEY = ROOT_TREE & " 2-0-3-1" Then
      Call InitReport3_2
   ElseIf Node.KEY = ROOT_TREE & " 2-0-3-2" Then
      Call InitReport2_0_3_2
   ElseIf Node.KEY = ROOT_TREE & " 2-0-3-3" Then
      Call InitReport2_0_3_3
   ElseIf Node.KEY = ROOT_TREE & " 2-0-4" Then
      Call InitReport2_0_4
   ElseIf Node.KEY = ROOT_TREE & " 2-0-4-1" Then
      Call InitReport2_0_4_1
   ElseIf Node.KEY = ROOT_TREE & " 2-0-5" Then
      Call InitReport2_0_5
   ElseIf Node.KEY = ROOT_TREE & " 2-0-5-1" Then
      Call InitReport2_0_5_1
   ElseIf Node.KEY = ROOT_TREE & " 2-0-6" Then
      Call InitReport2_0_6
   ElseIf Node.KEY = ROOT_TREE & " 2-0-6-1" Then
      Call InitReport2_0_6_1
   ElseIf Node.KEY = ROOT_TREE & " 2-0-6-2" Then
      Call InitReport2_0_6_2
   ElseIf Node.KEY = ROOT_TREE & " 2-0-7" Then
      Call InitReport2_0_7
   ElseIf Node.KEY = ROOT_TREE & " 2-0-8" Then
      Call InitReport2_0_8
   ElseIf Node.KEY = ROOT_TREE & "-A" & " 5-0" Then
      Call InitReportA_5_0
   ElseIf Node.KEY = ROOT_TREE & "-A" & " 5-1" Then
      Call InitReportA_5_1
      cmdAdd.Visible = True
   ElseIf Node.KEY = ROOT_TREE & "-A" & " 5-2" Then
      Call InitReportA_5_2
   ElseIf Node.KEY = ROOT_TREE & "-A" & " 5-3" Then
      Call InitReportA_5_3
   ElseIf Node.KEY = ROOT_TREE & "-A" & " 5-4" Then
      Call InitReportA_5_4
   ElseIf Node.KEY = ROOT_TREE & "-B" & " 1-1" Then
      Call InitReportB_1_1
 'commission
   ElseIf Node.KEY = ROOT_TREE & " 6-0-1" Then
      Call InitReport6_0_1
   ElseIf Node.KEY = ROOT_TREE & " 6-0-2" Then
     Call InitReport6_0_2
  ElseIf Node.KEY = ROOT_TREE & " 6-0-3" Then
     Call InitReport6_0_3
   ElseIf Node.KEY = ROOT_TREE & " 6-0-4" Then
      Call InitReport6_0_4

   ElseIf Node.KEY = ROOT_TREE & " 6-0-5" Then
     Call InitReport6_0_5
   ElseIf Node.KEY = ROOT_TREE & " 6-0-6" Then
      Call InitReport6_0_6
   ElseIf Node.KEY = ROOT_TREE & " 6-0-7" Then
       Call InitReport6_0_7
   ElseIf Node.KEY = ROOT_TREE & " 6-0-8" Then
      Call InitReport6_0_8
  ElseIf Node.KEY = ROOT_TREE & " 6-0-9" Then
      Call InitReport6_0_9
  ElseIf Node.KEY = ROOT_TREE & " 6-0-10" Then
      Call InitReport6_0_10                  ' รายงานส่วนลดตามรหัสลูกค้า
  ElseIf Node.KEY = ROOT_TREE & " 6-0-11" Then
     Call InitReport6_0_11
  ElseIf Node.KEY = ROOT_TREE & " 6-0-12" Then
     Call InitReport6_0_12
  ElseIf Node.KEY = ROOT_TREE & " 6-0-13" Then
      Call InitReport6_0_13
  ElseIf Node.KEY = ROOT_TREE & " 6-0-14" Then
      Call InitReport6_0_14
   ElseIf Node.KEY = ROOT_TREE & "-C" & " 1-1" Then
      Call InitReportC_1_1
   End If
   
End Sub

Private Sub InitReport3_1()
Dim C As CReportControl
Dim Top As Long
Dim Left As Long
Dim LabelWidth As Long
Dim Offset As Long

   Top = lblGeneric(0).Top
   Left = lblGeneric(0).Left
   LabelWidth = lblGeneric(0).Width
   Offset = 100
   
   '1 =============================
   Call LoadControl("T", txtGeneric(0).Width / 1.5, True, "", , "SUPPLIER_CODE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("รหัสผู้จำหน่าย"))
   
   '2 =============================
   Call LoadControl("T", txtGeneric(0).Width, True, "", , "SUPPLIER_NAME")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("ชื่อผู้จำหน่าย"))

   '3 =============================
   Call LoadControl("CB", cboGeneric(0).Width, True, "", 1, "SUPPLIER_TYPE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("ประเภทผู้จำหน่าย"))

   '3 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 2, "ORDER_BY")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("เรียงตาม"))

   '4 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 3, "ORDER_TYPE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("เรียงจาก"))
   
   Call ShowControl
   Call LoadComboData
End Sub
Private Sub InitReport2_0_4()
Dim C As CReportControl
Dim Top As Long
Dim Left As Long
Dim LabelWidth As Long
Dim Offset As Long

   Top = lblGeneric(0).Top
   Left = lblGeneric(0).Left
   LabelWidth = lblGeneric(0).Width
   Offset = 100
   
   '1 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, True, "", , "FROM_DOC_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("จากวันที่รับของ"))
   
   '1 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, True, "", , "TO_DOC_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("ถึงวันที่รับของ"))
   
   '1 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, True, "", , "FROM_DUE_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("จากวันครบดิว"))
   
   '1 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, True, "", , "TO_DUE_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("ถึงวันที่ครบดิว"))
   
   '1 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, True, "", , "TO_PAY_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("ถึงวันที่ชำระหนี้"))
   
   '1 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, True, "", , "PRINT_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("วันที่พิมพ์"))
   
   '1 =============================
   Call LoadControl("T", txtGeneric(0).Width / 1.5, True, "", , "FROM_SUPPLIER_CODE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("จากรหัสผู้จำหน่าย"))
   
   '2 =============================
   Call LoadControl("T", txtGeneric(0).Width / 1.5, True, "", , "TO_SUPPLIER_CODE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("ถึงรหัสผู้จำหน่าย"))

   '3 =============================
   Call LoadControl("CB", cboGeneric(0).Width, True, "", 1, "SUPPLIER_TYPE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("ประเภทผู้จำหน่าย"))

   '3 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 2, "ORDER_BY")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("เรียงตาม"))

   '4 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 3, "ORDER_TYPE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("เรียงจาก"))
   
   Call LoadControl("C", cboGeneric(0).Width, True, "", 4, "DATA_TYPE_ID", "DATA_TYPE_NAME")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("ประเภทข้อมูล"))
   
   Call LoadControl("C", cboGeneric(0).Width, True, "", 5, "GROUP_TYPE_CODE", "GROUP_TYPE_NAME")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("กลุ่มผู้จำหน่าย"))
   
   Call LoadControl("CH", chkGeneric(0).Width, True, "บันทึกลง EXCEL", 1, "SHOW_EXCEL")
   
   Call ShowControl
   Call LoadComboData
End Sub
Private Sub InitReport2_0_4_1()
Dim C As CReportControl
Dim Top As Long
Dim Left As Long
Dim LabelWidth As Long
Dim Offset As Long

   Top = lblGeneric(0).Top
   Left = lblGeneric(0).Left
   LabelWidth = lblGeneric(0).Width
   Offset = 100
   
   '1 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, True, "", , "FROM_DOC_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("จากวันที่รับของ"))
   
   '1 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, True, "", , "TO_DOC_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("ถึงวันที่รับของ"))
   
   '1 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, True, "", , "FROM_DUE_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("จากวันครบดิว"))
   
   '1 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, True, "", , "TO_DUE_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("ถึงวันที่ครบดิว"))
   
   '1 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, True, "", , "TO_PAY_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("ถึงวันที่ชำระหนี้"))
   
   '1 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, True, "", , "PRINT_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("วันที่พิมพ์"))
   
   '1 =============================
   Call LoadControl("T", txtGeneric(0).Width / 1.5, True, "", , "FROM_SUPPLIER_CODE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("จากรหัสผู้จำหน่าย"))
   
   '2 =============================
   Call LoadControl("T", txtGeneric(0).Width / 1.5, True, "", , "TO_SUPPLIER_CODE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("ถึงรหัสผู้จำหน่าย"))

   '3 =============================
   Call LoadControl("CB", cboGeneric(0).Width, True, "", 1, "SUPPLIER_TYPE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("ประเภทผู้จำหน่าย"))

   '3 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 2, "ORDER_BY")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("เรียงตาม"))

   '4 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 3, "ORDER_TYPE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("เรียงจาก"))
   
   Call LoadControl("C", cboGeneric(0).Width, False, "", 4, "DATA_TYPE_ID", "DATA_TYPE_NAME")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("ประเภทข้อมูล"))
   
'   Call LoadControl("C", cboGeneric(0).Width, True, "", 5, "GROUP_TYPE_CODE", "GROUP_TYPE_NAME")
'   Call LoadControl("L", lblGeneric(0).Width, True, MapText("กลุ่มผู้จำหน่าย"))
   
   '1 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, True, "", , "FROM_OVER_DUE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("จากวันเดือน OVER ดิว"))
   
   '1 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, True, "", , "TO_OVER_DUE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("ถึงวันเดือน OVER ดิว"))
   
   '1 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, True, "", , "FROM_DUE_DATE2")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("จากวันครบเดือนดิว"))
   
   '1 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, True, "", , "TO_DUE_DATE2")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("ถึงวันที่ครบดือนดิว"))
      
   Call LoadControl("T", txtGeneric(0).Width / 1.5, True, "", , "DECIMAL_AMOUNT")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("ทศนิยม (0)"))
   
   Call LoadControl("CH", chkGeneric(0).Width, True, "ไม่แสดงบิล", 1, "NOT_SHOW_BILL")
   Call LoadControl("CH", chkGeneric(0).Width, True, "แสดงเฉพาะกลุ่มผู้จำหน่าย", 1, "SHOW_ONLY_GROUP")
   
   Call ShowControl
   Call LoadComboData
End Sub

Private Sub InitReport3_2()
Dim C As CReportControl
Dim Top As Long
Dim Left As Long
Dim LabelWidth As Long
Dim Offset As Long

   Top = lblGeneric(0).Top
   Left = lblGeneric(0).Left
   LabelWidth = lblGeneric(0).Width
   Offset = 100
   
   '1 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, True, "", , "FROM_DOC_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("จากวันที่รับของ"))
   
   '1 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, True, "", , "TO_DOC_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("ถึงวันที่รับของ"))
   
   '1 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, True, "", , "FROM_DUE_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("จากวันครบดิว"))
   
   '1 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, True, "", , "TO_DUE_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("ถึงวันที่ครบดิว"))
   
   '1 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, True, "", , "TO_PAY_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("ถึงวันที่ชำระหนี้"))
   
   '1 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, True, "", , "PRINT_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("วันที่พิมพ์"))
   
   '1 =============================
   Call LoadControl("T", txtGeneric(0).Width / 1.5, True, "", , "FROM_SUPPLIER_CODE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("จากรหัสผู้จำหน่าย"))
   
   '2 =============================
   Call LoadControl("T", txtGeneric(0).Width / 1.5, True, "", , "TO_SUPPLIER_CODE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("ถึงรหัสผู้จำหน่าย"))
   
   '3 =============================
   Call LoadControl("CB", cboGeneric(0).Width, True, "", 1, "SUPPLIER_TYPE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("ประเภทผู้จำหน่าย"))

   '3 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 2, "ORDER_BY")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("เรียงตาม"))

   '4 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 3, "ORDER_TYPE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("เรียงจาก"))
   
   Call ShowControl
   Call LoadComboData
End Sub
Private Sub InitReport2_0_3_2()
Dim C As CReportControl
Dim Top As Long
Dim Left As Long
Dim LabelWidth As Long
Dim Offset As Long

   Top = lblGeneric(0).Top
   Left = lblGeneric(0).Left
   LabelWidth = lblGeneric(0).Width
   Offset = 100
   
   '1 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, True, "", , "FROM_DOC_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("จากวันที่รับของ"))
   
   '1 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, True, "", , "TO_DOC_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("ถึงวันที่รับของ"))
   
   '1 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, True, "", , "FROM_DUE_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("จากวันครบดิว"))
   
   '1 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, True, "", , "TO_DUE_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("ถึงวันที่ครบดิว"))
   
   '1 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, True, "", , "TO_PAY_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("ถึงวันที่ชำระหนี้"))
   
   '1 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, True, "", , "PRINT_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("วันที่พิมพ์"))
   
   '1 =============================
   Call LoadControl("T", txtGeneric(0).Width / 1.5, True, "", , "FROM_SUPPLIER_CODE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("จากรหัสผู้จำหน่าย"))
   
   '2 =============================
   Call LoadControl("T", txtGeneric(0).Width / 1.5, True, "", , "TO_SUPPLIER_CODE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("ถึงรหัสผู้จำหน่าย"))
       
   '3 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 2, "ORDER_BY")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("เรียงตาม"))

   '4 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 3, "ORDER_TYPE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("เรียงจาก"))
   
   Call LoadControl("C", cboGeneric(0).Width, False, "", 4, "DATA_TYPE_ID", "DATA_TYPE_NAME")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("ประเภทข้อมูล"))
   
   Call LoadControl("C", cboGeneric(0).Width, True, "", 5, "GROUP_TYPE_CODE", "GROUP_TYPE_NAME")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("กลุ่มผู้จำหน่าย"))
   
   Call ShowControl
   Call LoadComboData
End Sub
Private Sub InitReport2_0_3_3()
Dim C As CReportControl
Dim Top As Long
Dim Left As Long
Dim LabelWidth As Long
Dim Offset As Long

   Top = lblGeneric(0).Top
   Left = lblGeneric(0).Left
   LabelWidth = lblGeneric(0).Width
   Offset = 100
   
   '1 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, True, "", , "FROM_DOC_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("จากวันที่รับของ"))
   
   '1 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, True, "", , "TO_DOC_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("ถึงวันที่รับของ"))
   
   '1 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, True, "", , "FROM_DUE_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("จากวันครบดิว"))
   
   '1 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, True, "", , "TO_DUE_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("ถึงวันที่ครบดิว"))
   
   '1 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, True, "", , "TO_PAY_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("ถึงวันที่ชำระหนี้"))
   
   '1 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, True, "", , "PRINT_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("วันที่พิมพ์"))
   
   '1 =============================
   Call LoadControl("T", txtGeneric(0).Width / 1.5, True, "", , "FROM_SUPPLIER_CODE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("จากรหัสผู้จำหน่าย"))
   
   '2 =============================
   Call LoadControl("T", txtGeneric(0).Width / 1.5, True, "", , "TO_SUPPLIER_CODE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("ถึงรหัสผู้จำหน่าย"))
       
   '3 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 2, "ORDER_BY")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("เรียงตาม"))

   '4 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 3, "ORDER_TYPE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("เรียงจาก"))
   
   Call LoadControl("C", cboGeneric(0).Width, False, "", 4, "DATA_TYPE_ID", "DATA_TYPE_NAME")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("ประเภทข้อมูล"))
   
   Call LoadControl("C", cboGeneric(0).Width, True, "", 5, "GROUP_TYPE_CODE", "GROUP_TYPE_NAME")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("กลุ่มผู้จำหน่าย"))
   
      Call LoadControl("C", cboGeneric(0).Width, True, "", 6, "COMBO_SUB_ID", "COMBO_SUB_NAME")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("รวมกลุ่มประเภทย่อย"))
   
   Call LoadControl("CH", chkGeneric(0).Width, True, "สรุป", 1, "SUMMARIZE")
      Call LoadControl("CH", chkGeneric(0).Width, True, "แสดงชื่อเต็ม", 1, "SHOWNAME")
   
   Call ShowControl
   Call LoadComboData
End Sub
Private Sub InitReport1_0_5()
Dim C As CReportControl
Dim Top As Long
Dim Left As Long
Dim LabelWidth As Long
Dim Offset As Long

   Top = lblGeneric(0).Top
   Left = lblGeneric(0).Left
   LabelWidth = lblGeneric(0).Width
   Offset = 100
   
   '1 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, False, "", , "FROM_CHECK_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("จากวันที่ดิวเช็ค"))
   
   '1 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, True, "", , "TO_CHECK_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("ถึงวันที่ดิวเช็ค"))
      
   '1 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, True, "", , "PRINT_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("วันที่พิมพ์"))
   
   '1 =============================
'   Call LoadControl("T", txtGeneric(0).Width / 1.5, True, "", , "FROM_SUPPLIER_CODE")
'   Call LoadControl("L", lblGeneric(0).Width, True, MapText("จากรหัสผู้จำหน่าย"))
   
   '2 =============================
   Call LoadControl("T", txtGeneric(0).Width / 1.5, True, "", , "SUPPLIER_CODE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("รหัสผู้จำหน่าย"))

   '2 =============================
   Call LoadControl("T", txtGeneric(0).Width / 1.5, False, "", , "FROM_DOCUMENT_NO")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("จากรหัสใบสำคัญจ่าย"))
   
   '2 =============================
   Call LoadControl("T", txtGeneric(0).Width / 1.5, False, "", , "TO_DOCUMENT_NO")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("ถึงรหัสใบสำคัญจ่าย"))
   
   '2 =============================
   Call LoadControl("T", txtGeneric(0).Width / 1.5, False, "", , "FROM_DOCUMENT_NO1")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("หรือจากรหัสใบสำคัญจ่าย"))
   
   '2 =============================
   Call LoadControl("T", txtGeneric(0).Width / 1.5, False, "", , "TO_DOCUMENT_NO1")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("หรือถึงรหัสใบสำคัญจ่าย"))
   
   '3 =============================
   Call LoadControl("CB", cboGeneric(0).Width, True, "", 1, "SUPPLIER_TYPE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("ประเภทผู้จำหน่าย"))

   '3 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 2, "ORDER_BY")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("เรียงตาม"))

   '4 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 3, "ORDER_TYPE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("เรียงจาก"))
   
   Call ShowControl
   Call LoadComboData
End Sub
Private Sub InitReport1_0_1()
Dim C As CReportControl
Dim Top As Long
Dim Left As Long
Dim LabelWidth As Long
Dim Offset As Long

   Top = lblGeneric(0).Top
   Left = lblGeneric(0).Left
   LabelWidth = lblGeneric(0).Width
   Offset = 100
   
    '1 =============================
   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "FROM_CUSTOMER_CODE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("รหัสลูกหนี้"))
   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "TO_CUSTOMER_CODE", , True)

   '3 =============================
   Call LoadControl("CB", cboGeneric(0).Width, True, "", 1, "CUSTOMER_TYPE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("ประเภทลูกหนี้"))

   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "FROM_SALE_CODE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("รหัสพนักงานขาย"))
   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "TO_SALE_CODE", , True)
   
   '1 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, True, "", , "FROM_DOC_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("จากวันที่ขาย"))
   
   '1 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, True, "", , "TO_DOC_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("ถึงวันที่ขาย"))
   
'   '3 =============================
'   Call LoadControl("C", cboGeneric(0).Width, True, "", 2, "ORDER_BY")
'   Call LoadControl("L", lblGeneric(0).Width, True, MapText("เรียงตาม"))
'
'   '4 =============================
'   Call LoadControl("C", cboGeneric(0).Width, True, "", 3, "ORDER_TYPE")
'   Call LoadControl("L", lblGeneric(0).Width, True, MapText("เรียงจาก"))
   
   Call LoadControl("CH", chkGeneric(0).Width, True, "รวมตามพนักงานขาย", 1, "SHOW_SALE")
   
   Call LoadControl("CH", chkGeneric(0).Width, True, "บันทึกลง EXCEL", 1, "SHOW_EXCEL")
   Call LoadControl("CH", chkGeneric(0).Width, True, "NO", 1, "SHOW_NO")
   Call LoadControl("CH", chkGeneric(0).Width, True, "คำนำหน้าชื่อ", 1, "SHOW_PREFIX")
   Call LoadControl("CH", chkGeneric(0).Width, True, "ที่อยู่", 1, "SHOW_ADDRESS")
   Call LoadControl("CH", chkGeneric(0).Width, True, "จังหวัด", 1, "SHOW_PROVINCE")
   Call LoadControl("CH", chkGeneric(0).Width, True, "โทรศัพท์", 1, "SHOW_TEL")
   Call LoadControl("CH", chkGeneric(0).Width, True, "เครดิต", 1, "SHOW_CREDIT")
'   Call LoadControl("CH", chkGeneric(0).Width, True, "พนักงานขาย", 1, "SHOW_SALE")
   Call LoadControl("CH", chkGeneric(0).Width, True, "เขตการขาย", 1, "SHOW_AREA")
   Call LoadControl("CH", chkGeneric(0).Width, True, "วงเงิน", 1, "SHOW_LIMIT")
   Call LoadControl("CH", chkGeneric(0).Width, True, "ชื่อผู้ติดต่อ", 1, "SHOW_CONTRACT")
   Call LoadControl("CH", chkGeneric(0).Width, True, "ข้อมูลหมู", 1, "SHOW_PIGDATA")
   Call LoadControl("CH", chkGeneric(0).Width, True, "ข้อมูลยอดขาย", 1, "SHOW_VAC_NONVAC")
   
   
   
   Call ShowControl
   Call LoadComboData
End Sub
Private Sub InitReport1_0_1_1()
Dim C As CReportControl
Dim Top As Long
Dim Left As Long
Dim LabelWidth As Long
Dim Offset As Long

   Top = lblGeneric(0).Top
   Left = lblGeneric(0).Left
   LabelWidth = lblGeneric(0).Width
   Offset = 100
   
   '1 =============================
   Call LoadControl("T", txtGeneric(0).Width / 1.5, False, "", , "COLLUMN")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("จำนวนคอลัมน์"))
   
   Call LoadControl("T", txtGeneric(0).Width / 1.5, True, "", , "FONT_SIZE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("ขนาด FONT"))
   
   Call LoadControl("T", txtGeneric(0).Width / 1.5, True, "", , "ROW_HEIGHT")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("ส่วนสูงของแถว"))
   
   Call LoadControl("T", txtGeneric(0).Width / 1.5, True, "", , "FROM_CUSTOMER_CODE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("จากรหัสลูกหนี้"))
   
   '2 =============================
   Call LoadControl("T", txtGeneric(0).Width / 1.5, True, "", , "TO_CUSTOMER_CODE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("ถึงรหัสลูกหนี้"))

   '3 =============================
   Call LoadControl("CB", cboGeneric(0).Width, True, "", 1, "CUSTOMER_TYPE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("ประเภทลูกหนี้"))
   
   '3 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 2, "ORDER_BY")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("เรียงตาม"))

   '4 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 3, "ORDER_TYPE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("เรียงจาก"))
   
   Call LoadControl("CH", chkGeneric(0).Width, True, "บันทึกลง EXCEL", 1, "SHOW_EXCEL")
   
   Call ShowControl
   Call LoadComboData
End Sub
Private Sub InitReport1_0_1_2()
Dim C As CReportControl
Dim Top As Long
Dim Left As Long
Dim LabelWidth As Long
Dim Offset As Long

   Top = lblGeneric(0).Top
   Left = lblGeneric(0).Left
   LabelWidth = lblGeneric(0).Width
   Offset = 100
   
   '1 =============================
   Call LoadControl("T", txtGeneric(0).Width / 1.5, True, "", , "FROM_CUSTOMER_CODE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("จากรหัสลูกหนี้"))
   
   '2 =============================
   Call LoadControl("T", txtGeneric(0).Width / 1.5, True, "", , "TO_CUSTOMER_CODE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("ถึงรหัสลูกหนี้"))
   
   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "FROM_SALE_CODE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("รหัสพนักงานขาย"))
   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "TO_SALE_CODE", , True)
   
   Call ShowControl
   Call LoadComboData
End Sub

Private Sub InitReport1_0_2()
Dim C As CReportControl
Dim Top As Long
Dim Left As Long
Dim LabelWidth As Long
Dim Offset As Long

   Top = lblGeneric(0).Top
   Left = lblGeneric(0).Left
   LabelWidth = lblGeneric(0).Width
   Offset = 100
   
   '1 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, True, "", , "FROM_DOC_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("ถึงวันที่ขาย"))
   
   '1 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, True, "", , "TO_DOC_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("ถึงวันที่ขาย"))
   
   '1 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, True, "", , "TO_PAY_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("ถึงวันที่ชำระหนี้"))
   
   '1 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, True, "", , "PRINT_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("วันที่พิมพ์"))
   
   '1 =============================
   Call LoadControl("T", txtGeneric(0).Width / 1.5, True, "", , "FROM_CUSTOMER_CODE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("จากรหัสลูกหนี้"))
   
   '2 =============================
   Call LoadControl("T", txtGeneric(0).Width / 1.5, True, "", , "TO_CUSTOMER_CODE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("ถึงรหัสลูกหนี้"))

   '3 =============================
   Call LoadControl("CB", cboGeneric(0).Width \ 2, False, "", 2, "INTERVAL_TYPE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("ประเภทช่วง"))

   '3 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 3, "ORDER_BY")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("เรียงตาม"))

   '4 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 4, "ORDER_TYPE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("เรียงจาก"))
   
   Call LoadControl("T", txtGeneric(0).Width / 1.5, True, "", , "CREDIT")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("ระบบุเครดิต(วัน)"))
   
   Call LoadControl("CH", chkGeneric(0).Width, True, "ใช้เครดิตตามตารางจริง", 1, "REAL_CREDIT_FLAG")
   
   Call LoadControl("CH", chkGeneric(0).Width, True, "เครดิต 90 วัน(ถ้าไม่มีในตารางจริง)", 2, "NINETY_FLAG")
   
   Call LoadControl("CH", chkGeneric(0).Width, True, "สรุป", 2, "SUMMARY_FLAG")
   
   Call LoadControl("CH", chkGeneric(0).Width, True, "ลูกหนี้ BANK", 2, "CUSTOMER_BANK")
   
   Call ShowControl
   Call LoadComboData
End Sub
Private Sub InitReport1_0_16()
Dim C As CReportControl
Dim Top As Long
Dim Left As Long
Dim LabelWidth As Long
Dim Offset As Long

   Top = lblGeneric(0).Top
   Left = lblGeneric(0).Left
   LabelWidth = lblGeneric(0).Width
   Offset = 100
   
      '1 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, True, "", , "TO_PAY_DATE1")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("DB1 วันที่"))
   
   '1 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, True, "", , "TO_PAY_DATE2")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("DB2 วันที่"))
   
   '1 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, True, "", , "TO_PAY_DATE3")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("DB3 วันที่"))
   
   '1 =============================
   Call LoadControl("T", txtGeneric(0).Width / 1.5, True, "", , "FROM_CUSTOMER_CODE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("จากรหัสลูกหนี้"))
   
   '2 =============================
   Call LoadControl("T", txtGeneric(0).Width / 1.5, True, "", , "TO_CUSTOMER_CODE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("ถึงรหัสลูกหนี้"))

'   '3 =============================
'   Call LoadControl("C", cboGeneric(0).Width, True, "", 3, "ORDER_BY")
'   Call LoadControl("L", lblGeneric(0).Width, True, MapText("เรียงตาม"))
'
'   '4 =============================
'   Call LoadControl("C", cboGeneric(0).Width, True, "", 4, "ORDER_TYPE")
'   Call LoadControl("L", lblGeneric(0).Width, True, MapText("เรียงจาก"))

'   Call LoadControl("CH", chkGeneric(0).Width, True, "ใช้เครดิตตามตารางจริง", 1, "REAL_CREDIT_FLAG")
'
'   Call LoadControl("CH", chkGeneric(0).Width, True, "เครดิต 90 วัน(ถ้าไม่มีในตารางจริง)", 2, "NINETY_FLAG")
'
'   Call LoadControl("CH", chkGeneric(0).Width, True, "สรุป", 2, "SUMMARY_FLAG")
'
'   Call LoadControl("CH", chkGeneric(0).Width, True, "ลูกหนี้ BANK", 2, "CUSTOMER_BANK")

   Call ShowControl
   Call LoadComboData
End Sub

Private Sub InitReport1_0_2_1()
Dim C As CReportControl
Dim Top As Long
Dim Left As Long
Dim LabelWidth As Long
Dim Offset As Long

   Top = lblGeneric(0).Top
   Left = lblGeneric(0).Left
   LabelWidth = lblGeneric(0).Width
   Offset = 100
   
   '1 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, True, "", , "FROM_DOC_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("จากวันที่ขาย"))
   
   '1 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, True, "", , "TO_DOC_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("ถึงวันที่ขาย"))
   
   '1 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, True, "", , "TO_PAY_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("ถึงวันที่ชำระหนี้"))
   
   '1 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, True, "", , "PRINT_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("วันที่พิมพ์"))
   
   '1 =============================
   Call LoadControl("T", txtGeneric(0).Width / 1.5, True, "", , "FROM_CUSTOMER_CODE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("จากรหัสลูกหนี้"))
   
   '2 =============================
   Call LoadControl("T", txtGeneric(0).Width / 1.5, True, "", , "TO_CUSTOMER_CODE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("ถึงรหัสลูกหนี้"))

   Call LoadControl("CH", chkGeneric(0).Width, True, "สรุป", 2, "SUMMARY_FLAG")
   
   Call LoadControl("CH", chkGeneric(0).Width, True, "ลูกหนี้ BANK", 2, "CUSTOMER_BANK")
   
   Call ShowControl
   Call LoadComboData
End Sub

Private Sub InitReport1_0_4()
Dim C As CReportControl
Dim Top As Long
Dim Left As Long
Dim LabelWidth As Long
Dim Offset As Long

   Top = lblGeneric(0).Top
   Left = lblGeneric(0).Left
   LabelWidth = lblGeneric(0).Width
   Offset = 100
   
   '1 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, True, "", , "FROM_BUY_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("จากวันที่ซื้อ"))
   
   '1 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, True, "", , "TO_BUY_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("ถึงวันที่ซื้อ"))
      
   '1 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, False, "", , "PRINT_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("วันที่พิมพ์"))
      
   '1 =============================
   Call LoadControl("T", txtGeneric(0).Width / 1.5, True, "", , "FASCOD")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("รหัสทรัพย์สิน"))
      
   '1 =============================
   Call LoadControl("T", txtGeneric(0).Width / 1.5, True, "", , "ASSET_CODE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("รหัสบัญชีทรัพย์สิน"))
   
   '1 =============================
   Call LoadControl("T", txtGeneric(0).Width / 1.5, True, "", , "DPRC_CODE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("รหัสบัญชีค่าเสื่อม"))

   '3 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 1, "ORDER_BY")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("เรียงตาม"))

   '4 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 2, "ORDER_TYPE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("เรียงจาก"))
   
   Call ShowControl
   Call LoadComboData
End Sub

Private Sub InitReportA_5_2()
Dim C As CReportControl
Dim Top As Long
Dim Left As Long
Dim LabelWidth As Long
Dim Offset As Long

   Top = lblGeneric(0).Top
   Left = lblGeneric(0).Left
   LabelWidth = lblGeneric(0).Width
   Offset = 100

   '1 =============================
   Call LoadControl("T", txtGeneric(0).Width, False, "", , "DOCUMENT_NO")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("เลขที่ใบสำคัญโอน"))
 
   Call LoadControl("CH", chkGeneric(0).Width, True, "แสดงรูป", 1, "PICTURE_FLAG")
   
   Call ShowControl
   Call LoadComboData
End Sub

Private Sub InitReportA_5_3()
Dim C As CReportControl
Dim Top As Long
Dim Left As Long
Dim LabelWidth As Long
Dim Offset As Long

   Top = lblGeneric(0).Top
   Left = lblGeneric(0).Left
   LabelWidth = lblGeneric(0).Width
   Offset = 100

   '1 =============================
   Call LoadControl("T", txtGeneric(0).Width, False, "", , "DOCUMENT_NO")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("เลขที่ใบสำคัญรับ"))

'   '3 =============================
'   Call LoadControl("C", cboGeneric(0).Width, True, "", 1, "ACCOUNT_ID", "ACCOUNT_NO")
'   Call LoadControl("L", lblGeneric(0).Width, True, MapText("รหัสบัญชี"))

   '1 =============================
   Call LoadControl("T", txtGeneric(0).Width \ 2, False, "", , "ACCOUNT_NO")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("รหัสบัญชี"))
   
   '1 =============================
   Call LoadControl("T", txtGeneric(0).Width \ 2, False, "", , "PAY_FOR")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("ชำระค่า"))
   
   Call LoadControl("T", txtGeneric(0).Width \ 2, False, "", , "TOTAL_AMOUNT")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("จำนวนเงิน"))
   
   '4 =============================
   Call LoadControl("CH", cboGeneric(0).Width, True, "เงินโอน", 1, "TRANSFER_FLAG")
   
   '4 =============================
   Call LoadControl("CH", cboGeneric(0).Width, True, "แสดงรูป", 2, "PICTURE_FLAG")

   Call ShowControl
   Call LoadComboData
End Sub

Private Sub InitReportA_5_1()
Dim C As CReportControl
Dim Top As Long
Dim Left As Long
Dim LabelWidth As Long
Dim Offset As Long

   Top = lblGeneric(0).Top
   Left = lblGeneric(0).Left
   LabelWidth = lblGeneric(0).Width
   Offset = 100
   cmdAdd.Visible = True
   cmdAdd.Top = 3390
   '1 =============================
'   Call LoadControl("D", uctlGenericDate(0).Width, True, "", , "FROM_DOC_DATE")
'   Call LoadControl("L", lblGeneric(0).Width, True, MapText("จากวันที่เอกสาร"))
'
'   '1 =============================
'   Call LoadControl("D", uctlGenericDate(0).Width, True, "", , "TO_DOC_DATE")
'   Call LoadControl("L", lblGeneric(0).Width, True, MapText("ถึงวันที่เอกสาร"))

   '1 =============================
   Call LoadControl("T", txtGeneric(0).Width, False, "", , "DOCUMENT_NO")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("เลขที่ใบสำคัญจ่าย"))
 
'   '2  =============================
   Call LoadControl("T", txtGeneric(0).Width \ 2, False, "", , "ACCOUNT_NO")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("รหัสบัญชี"))
   
   '3 =============================
   Call LoadControl("C", cboGeneric(0).Width, False, "", 1, "JOURNAL_TYPE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("สมุดรายวันแบบ"))
   
   Call LoadControl("D", uctlGenericDate(0).Width, True, "", , "PRINT_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("วันที่พิมพ์"))
   
   Call LoadControl("CH", chkGeneric(0).Width, True, "แสดงรูป", 1, "PICTURE_FLAG")

'   '4 =============================
'   Call LoadControl("C", cboGeneric(0).Width, True, "", 2, "ORDER_TYPE")
'   Call LoadControl("L", lblGeneric(0).Width, True, MapText("เรียงจาก"))
   
 
   Call ShowControl
   Call LoadComboData
End Sub

Private Sub InitReportA_5_4()
Dim C As CReportControl
Dim Top As Long
Dim Left As Long
Dim LabelWidth As Long
Dim Offset As Long

   Top = lblGeneric(0).Top
   Left = lblGeneric(0).Left
   LabelWidth = lblGeneric(0).Width
   Offset = 100
   cmdAdd.Visible = True
   cmdAdd.Top = 3390 + cmdAdd.Height + 450
   
   '1 =============================
   Call LoadControl("T", txtGeneric(0).Width, False, "", , "DOCUMENT_NO")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("เลขที่ใบสำคัญรับ"))
 
   '1 =============================
   Call LoadControl("T", txtGeneric(0).Width \ 2, False, "", , "PAY_FOR")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("ชำระค่า"))

'   '2  =============================
   Call LoadControl("T", txtGeneric(0).Width \ 2, False, "", , "TOTAL_AMOUNT")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("จำนวนเงิน"))
   
   '3 =============================
   Call LoadControl("C", cboGeneric(0).Width, False, "", 1, "JOURNAL_TYPE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("สมุดรายวันแบบ"))
   
   '4 =============================
   Call LoadControl("CH", cboGeneric(0).Width, True, "เงินโอน", 1, "TRANSFER_FLAG")
      
   Call LoadControl("CH", chkGeneric(0).Width, True, "แสดงรูป", 2, "PICTURE_FLAG")
    
   Call ShowControl
   Call LoadComboData
End Sub

Private Sub InitReportB_1_1()
Dim C As CReportControl
Dim Top As Long
Dim Left As Long
Dim LabelWidth As Long
Dim Offset As Long

   Top = lblGeneric(0).Top
   Left = lblGeneric(0).Left
   LabelWidth = lblGeneric(0).Width
   Offset = 100
   
   '1 =============================
   Call LoadControl("T", txtGeneric(0).Width \ 2, False, "", , "CHECK_NO")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("เลขที่เช็ค"))
       
   '1 =============================
   Call LoadControl("T", txtGeneric(0).Width, True, "", , "PAYEE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("จ่ายให้"))
       
   '3 =============================
   Call LoadControl("C", cboGeneric(0).Width, False, "", 1, "CHECK_TYPE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("รูปแบบเช็ค"))
       
   Call LoadControl("CH", chkGeneric(0).Width, True, "ไม่ต้องลงวันที่", 2, "LINE_FLAG")
    
   Call ShowControl
   Call LoadComboData
End Sub
Private Sub InitReportC_1_1()
Dim C As CReportControl
Dim Top As Long
Dim Left As Long
Dim LabelWidth As Long
Dim Offset As Long

   Top = lblGeneric(0).Top
   Left = lblGeneric(0).Left
   LabelWidth = lblGeneric(0).Width
   Offset = 100
   
   '1 =============================
   Call LoadControl("T", txtGeneric(0).Width \ 2, False, "", , "FROM_MONTH")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("จากเดือน"))
       
   '1 =============================
   Call LoadControl("T", txtGeneric(0).Width, False, "", , "TO_MONTH")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("ถึงเดือน"))
    
   '3 =============================
   Call LoadControl("T", txtGeneric(0).Width, False, "", , "YEAR")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("ปี"))
    
   Call ShowControl
   Call LoadComboData
End Sub

Private Sub InitReportA_5_0()
Dim C As CReportControl
Dim Top As Long
Dim Left As Long
Dim LabelWidth As Long
Dim Offset As Long

   Top = lblGeneric(0).Top
   Left = lblGeneric(0).Left
   LabelWidth = lblGeneric(0).Width
   Offset = 100
   
'   '1 =============================
'   Call LoadControl("D", uctlGenericDate(0).Width, True, "", , "FROM_DOC_DATE")
'   Call LoadControl("L", lblGeneric(0).Width, True, MapText("จากวันที่เอกสาร"))
'
'   '1 =============================
'   Call LoadControl("D", uctlGenericDate(0).Width, True, "", , "TO_DOC_DATE")
'   Call LoadControl("L", lblGeneric(0).Width, True, MapText("ถึงวันที่เอกสาร"))

   '1 =============================
   Call LoadControl("T", txtGeneric(0).Width, False, "", , "DOCUMENT_NO")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("เลขที่ใบสำคัญจ่าย"))

   '1 =============================
   Call LoadControl("T", txtGeneric(0).Width \ 2, False, "", , "ACCOUNT_NO")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("รหัสบัญชี"))
   
'   '3 =============================

   Call LoadControl("D", uctlGenericDate(0).Width, True, "", , "PRINT_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("วันที่พิมพ์"))

'   Call LoadControl("C", cboGeneric(0).Width, True, "", 1, "ACCOUNT_ID", "ACCOUNT_NO")
'   Call LoadControl("L", lblGeneric(0).Width, True, MapText("เพื่อจ่ายให้"))

'   Call LoadControl("T", txtGeneric(0).Width \ 2, False, "", , "TOTAL_AMOUNT")
'   Call LoadControl("L", lblGeneric(0).Width, True, MapText("จำนวนเงิน"))

'   '4 =============================
'   Call LoadControl("C", cboGeneric(0).Width, True, "", 2, "ORDER_TYPE")
'   Call LoadControl("L", lblGeneric(0).Width, True, MapText("เรียงจาก"))
   
   '4 =============================
   Call LoadControl("CH", cboGeneric(0).Width, True, "แสดงรูป", 1, "PICTURE_FLAG")

   Call ShowControl
   Call LoadComboData
End Sub
Private Sub LoadComboData()
Dim C As CReportControl
Dim YEAR_ID As Long

'   Me.Refresh
'   DoEvents
'   Call EnableForm(Me, False)
   
   For Each C In m_ReportControls
      If (C.ControlType = "C") Or (C.ControlType = "CB") Then
         
         If trvMain.SelectedItem.KEY = ROOT_TREE & " 1-0-1" Then
            If C.ComboLoadID = 1 Then
               Call LoadCustomerType(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 2 Then
               'Call InitCustomerOrderBy(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 3 Then
               'Call InitOrderType(m_Combos(C.ControlIndex))
            End If
         End If
         
         If trvMain.SelectedItem.KEY = ROOT_TREE & " 1-0-1-1" Then
            If C.ComboLoadID = 1 Then
               Call LoadCustomerType(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 2 Then
               Call InitCustomerOrderBy(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 3 Then
               Call InitOrderType(m_Combos(C.ControlIndex))
            End If
         End If
         
         If trvMain.SelectedItem.KEY = ROOT_TREE & " 1-0-2" Then
            If C.ComboLoadID = 1 Then
               Call LoadCustomerType(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 2 Then
               Call InitIntervalType(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 3 Then
               Call InitCustomerOrderBy(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 4 Then
            
               Call InitOrderType(m_Combos(C.ControlIndex))
            End If
         End If
         
         If trvMain.SelectedItem.KEY = ROOT_TREE & " 1-0-4" Then
            If C.ComboLoadID = 1 Then
               Call InitAssetOrderBy(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 2 Then
               Call InitOrderType(m_Combos(C.ControlIndex))
            End If
         End If
      
         If trvMain.SelectedItem.KEY = ROOT_TREE & " 1-0-5" Or trvMain.SelectedItem.KEY = ROOT_TREE & " 1-0-5-1" Or trvMain.SelectedItem.KEY = ROOT_TREE & " 1-0-5-2" Or trvMain.SelectedItem.KEY = ROOT_TREE & " 1-0-5-3" Or trvMain.SelectedItem.KEY = ROOT_TREE & " 1-0-5-4" Then
            If C.ComboLoadID = 1 Then
               Call LoadSupplierType(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 2 Then
               Call InitSupplierOrderBy(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 3 Then
               Call InitOrderType(m_Combos(C.ControlIndex))
            End If
         End If
      
         If trvMain.SelectedItem.KEY = ROOT_TREE & " 1-0-6" Then
            If C.ComboLoadID = 1 Then
               Call InitIntervalType(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 2 Then
               Call LoadCustomerType(m_Combos(C.ControlIndex))
            End If
         End If
         
      If trvMain.SelectedItem.KEY = ROOT_TREE & " 1-0-6-1" Then
            If C.ComboLoadID = 1 Then
               Call LoadCustomerType(m_Combos(C.ControlIndex))
            End If
         End If
         
         If trvMain.SelectedItem.KEY = ROOT_TREE & " 1-0-6-2" Then
            If C.ComboLoadID = 1 Then
               Call LoadCustomerType(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 2 Then
               Call InitSaleOrderBy(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 3 Then
               Call InitOrderType(m_Combos(C.ControlIndex))
            End If
         End If
         
         If trvMain.SelectedItem.KEY = ROOT_TREE & " 1-0-6-3" Then
             If C.ComboLoadID = 1 Then
               Call LoadCustomerType(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 2 Then
               Call InitSaleOrderBy(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 3 Then
               Call InitOrderType(m_Combos(C.ControlIndex))
            End If
         End If
         
         If trvMain.SelectedItem.KEY = ROOT_TREE & " 1-0-6-4" Then
            If C.ComboLoadID = 1 Then
               Call InitThaiMonth(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 2 Then
               Call LoadCustomerType(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 3 Then
               Call InitSaleOrderBy(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 4 Then
               Call InitOrderType(m_Combos(C.ControlIndex))
            End If
         End If
         
         If trvMain.SelectedItem.KEY = ROOT_TREE & " 1-0-6-5" Then
              If C.ComboLoadID = 1 Then
               Call InitIntervalType(m_Combos(C.ControlIndex))
            End If
            
'            If C.ComboLoadID = 1 Then
'               Call InitCustomer2OrderBy(m_Combos(C.ControlIndex))
'            ElseIf C.ComboLoadID = 2 Then
'               Call InitOrderType(m_Combos(C.ControlIndex))
'            End If
         End If
         
        If trvMain.SelectedItem.KEY = ROOT_TREE & " 1-0-7-1" Then
            If C.ComboLoadID = 1 Then
                   Call InitThaiMonth(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 2 Then
                   Call LoadCustomerType(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 3 Then
                   Call LoadDataType(m_Combos(C.ControlIndex))
            End If
         End If
         
         
        If trvMain.SelectedItem.KEY = ROOT_TREE & " 1-0-8-1" Then
            If C.ComboLoadID = 1 Then
                Call LoadCustomerType2(m_Combos(C.ControlIndex))
            End If
        End If
         
         If trvMain.SelectedItem.KEY = ROOT_TREE & " 1-0-10" Or trvMain.SelectedItem.KEY = ROOT_TREE & " 1-0-10-1" Or trvMain.SelectedItem.KEY = ROOT_TREE & " 1-0-10-2" Or trvMain.SelectedItem.KEY = ROOT_TREE & " 1-0-10-3" Or trvMain.SelectedItem.KEY = ROOT_TREE & " 1-0-10-4" Or trvMain.SelectedItem.KEY = ROOT_TREE & " 1-0-11" Or trvMain.SelectedItem.KEY = ROOT_TREE & " 1-0-19" Or trvMain.SelectedItem.KEY = ROOT_TREE & " 1-0-20" Or trvMain.SelectedItem.KEY = ROOT_TREE & " 1-0-21" Then

            If C.ComboLoadID = 1 Then
               Call InitThaiMonth(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 2 Then
               Call InitThaiMonth(m_Combos(C.ControlIndex))
            End If
         End If
         
         If trvMain.SelectedItem.KEY = ROOT_TREE & " 1-0-13" Then
            If C.ComboLoadID = 1 Then
                Call LoadBank(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 2 Then
               Call LoadBankCustomer(m_Combos(C.ControlIndex))
            End If
        End If
         
         If trvMain.SelectedItem.KEY = ROOT_TREE & " 1-0-15-1" Then
            If C.ComboLoadID = 1 Then
               Call LoadCustomerType(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 2 Then
               Call InitCustomerOrderBy(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 3 Then
               Call InitOrderType(m_Combos(C.ControlIndex))
            End If
         End If
         
         If trvMain.SelectedItem.KEY = ROOT_TREE & " 2-0-1" Then
            If C.ComboLoadID = 1 Then
               Call LoadSupplierType(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 2 Then
               Call InitSupplierOrderBy(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 3 Then
               Call InitOrderType(m_Combos(C.ControlIndex))
            End If
         End If
      
         If trvMain.SelectedItem.KEY = ROOT_TREE & " 2-0-2" Then
            If C.ComboLoadID = 1 Then
               Call LoadSupplierType(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 2 Then
               Call InitSupplierOrderBy(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 3 Then
               Call InitOrderType(m_Combos(C.ControlIndex))
            End If
         End If
      
         If trvMain.SelectedItem.KEY = ROOT_TREE & " 2-0-3" Or trvMain.SelectedItem.KEY = ROOT_TREE & " 2-0-3-1" Then
            If C.ComboLoadID = 1 Then
               Call LoadSupplierType(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 2 Then
               Call InitSupplierOrderBy(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 3 Then
               Call InitOrderType(m_Combos(C.ControlIndex))
            End If
         End If
            
        If trvMain.SelectedItem.KEY = ROOT_TREE & " 2-0-3-2" Or trvMain.SelectedItem.KEY = ROOT_TREE & " 2-0-3-3" Then
            If C.ComboLoadID = 2 Then
               Call InitSupplierOrderBy(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 3 Then
               Call InitOrderType(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 4 Then
               Call LoadDataType(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 5 Then
               Call LoadGroupTypeData(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 6 Then
               Call LoadComboSupTypeData(m_Combos(C.ControlIndex))
            End If
         End If
         
         If trvMain.SelectedItem.KEY = ROOT_TREE & " 2-0-4" Or trvMain.SelectedItem.KEY = ROOT_TREE & " 2-0-4-1" Or trvMain.SelectedItem.KEY = ROOT_TREE & " 2-0-2-1" Then
            If C.ComboLoadID = 1 Then
               Call LoadSupplierType(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 2 Then
               Call InitSupplierOrderBy(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 3 Then
               Call InitOrderType(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 4 Then
               Call LoadDataType(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 5 Then
               Call LoadGroupTypeData(m_Combos(C.ControlIndex))
            End If
         End If
         
         If trvMain.SelectedItem.KEY = ROOT_TREE & " 2-0-5" Or trvMain.SelectedItem.KEY = ROOT_TREE & " 2-0-5-1" Then
            If C.ComboLoadID = 1 Then
               Call LoadSupplierType(m_Combos(C.ControlIndex))
            End If
         End If
         
          If trvMain.SelectedItem.KEY = ROOT_TREE & " 2-0-6" Or trvMain.SelectedItem.KEY = ROOT_TREE & " 2-0-6-1" Or trvMain.SelectedItem.KEY = ROOT_TREE & " 2-0-6-2" Then
            If C.ComboLoadID = 1 Then
               Call InitThaiMonth(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 2 Then
               Call LoadSupplierType(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 3 Then
               Call LoadDataType(m_Combos(C.ControlIndex))
            End If
         End If
         
          If trvMain.SelectedItem.KEY = ROOT_TREE & " 2-0-7" Then
            If C.ComboLoadID = 3 Then
               Call LoadDataType(m_Combos(C.ControlIndex))
            End If
         End If
         
         If trvMain.SelectedItem.KEY = ROOT_TREE & " 2-0-8" Then
            If C.ComboLoadID = 1 Then
               Call InitPrType(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 2 Then
               Call LoadSupplierType(m_Combos(C.ControlIndex))
            End If
         End If
         
         If trvMain.SelectedItem.KEY = ROOT_TREE & "-A" & " 5-0" Then
            If C.ComboLoadID = 1 Then
                  
            End If
         End If
      
         If trvMain.SelectedItem.KEY = ROOT_TREE & "-A" & " 5-1" Then
            If C.ComboLoadID = 1 Then
               Call InitJournalType(m_Combos(C.ControlIndex))
'            ElseIf C.ComboLoadID = 2 Then
'               Call InitOrderType(m_Combos(C.ControlIndex))
            End If
         End If
      
         If trvMain.SelectedItem.KEY = ROOT_TREE & "-A" & " 5-2" Then
'            If C.ComboLoadID = 1 Then
'               Call InitDocumentOrderBy(m_Combos(C.ControlIndex))
'            ElseIf C.ComboLoadID = 2 Then
'               Call InitOrderType(m_Combos(C.ControlIndex))
'            End If
         End If
      
        If trvMain.SelectedItem.KEY = ROOT_TREE & "-A" & " 5-3" Then
            If C.ComboLoadID = 1 Then
               Call InitAccountNo1(m_Combos(C.ControlIndex))
            End If
         End If
      
         If trvMain.SelectedItem.KEY = ROOT_TREE & "-A" & " 5-4" Then
            If C.ComboLoadID = 1 Then
               Call InitJournalType(m_Combos(C.ControlIndex))
            End If
         End If
      
         If trvMain.SelectedItem.KEY = ROOT_TREE & "-B" & " 1-1" Then
            If C.ComboLoadID = 1 Then
               Call InitCheckType(m_Combos(C.ControlIndex))
            End If
         End If
      
       ' commission เก็บเงิน
       If trvMain.SelectedItem.KEY = ROOT_TREE & " 6-0-1" Or trvMain.SelectedItem.KEY = ROOT_TREE & " 6-0-2" Or trvMain.SelectedItem.KEY = ROOT_TREE & " 6-0-3" Or trvMain.SelectedItem.KEY = ROOT_TREE & " 6-0-4" Or trvMain.SelectedItem.KEY = ROOT_TREE & " 6-0-5" Or trvMain.SelectedItem.KEY = ROOT_TREE & " 6-0-6" Or trvMain.SelectedItem.KEY = ROOT_TREE & " 6-0-13" Then
            If C.ComboLoadID = 1 Then
                  Call LoadAreaCom(m_Combos(C.ControlIndex))     ' modloaddata
            ElseIf C.ComboLoadID = 2 Then
                  Call LoadSale(m_Combos(C.ControlIndex))
            End If
        End If
        
      If trvMain.SelectedItem.KEY = ROOT_TREE & " 6-0-7" Then
            If C.ComboLoadID = 1 Then
                   Call InitThaiMonth(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 2 Then
                  Call LoadSale(m_Combos(C.ControlIndex))
'            ElseIf C.ComboLoadID = 3 Then
'                   Call LoadDataType(m_Combos(C.ControlIndex))
            End If
         End If
         
      If trvMain.SelectedItem.KEY = ROOT_TREE & " 6-0-8" Then
            If C.ComboLoadID = 1 Then
                   Call InitThaiMonth(m_Combos(C.ControlIndex))
            End If
         End If
      
      If trvMain.SelectedItem.KEY = ROOT_TREE & " 6-0-9" Then
            If C.ComboLoadID = 2 Then
                   Call LoadSale(m_Combos(C.ControlIndex))
            End If
         End If
      
     If trvMain.SelectedItem.KEY = ROOT_TREE & " 6-0-11" Then
            If C.ComboLoadID = 1 Then
                  Call LoadAreaCom(m_Combos(C.ControlIndex))     ' modloaddata
            End If
         End If
         
    If trvMain.SelectedItem.KEY = ROOT_TREE & " 6-0-12" Then
         If C.ComboLoadID = 1 Then
               Call LoadAreaCom(m_Combos(C.ControlIndex))     ' modloaddata
         End If
      End If
      
      If trvMain.SelectedItem.KEY = ROOT_TREE & " 6-0-14" Then
           If C.ComboLoadID = 1 Then
               Call LoadAreaCom(m_Combos(C.ControlIndex))     ' modloaddata
           ElseIf C.ComboLoadID = 2 Then
               Call InitComDocType(m_Combos(C.ControlIndex))     ' modloaddata
           ElseIf C.ComboLoadID = 3 Then
               Call LoadSale(m_Combos(C.ControlIndex))
           End If
        End If
        
      End If 'C.ControlType = "C"
   Next C
'   Call EnableForm(Me, True)
End Sub
Private Sub Form_Resize()
'On Error Resume Next
On Error GoTo ErrorHandler
   SSFrame2.Width = ScaleWidth
   SSFrame2.Height = ScaleHeight
   
   SSPanel1.Width = ScaleWidth
   
   If ScaleWidth > 0 Then
      trvMain.Width = ScaleWidth - SSFrame3.Width
   End If
   
   If ScaleHeight > 0 Then
      cmdPasswd.Top = ScaleHeight - cmdPasswd.Height - 100
      cmdPasswd2.Top = ScaleHeight - cmdPasswd2.Height - 100
      cmdExit.Top = cmdPasswd.Top
      cmdConfig.Top = cmdExit.Top
      cmdOK.Top = cmdExit.Top
      lblVersion.Top = cmdPasswd.Top + 100
   End If
   
   If ScaleWidth > 0 Then
      cmdOK.Left = ScaleWidth - cmdOK.Width - 40
      cmdConfig.Left = ScaleWidth - cmdOK.Width - 40 - cmdConfig.Width - 40
      
      trvMain.Height = cmdPasswd.Top - SSPanel1.Height - 100
      trvMain.Height = cmdPasswd2.Top - SSPanel1.Height - 100
      SSFrame3.Height = trvMain.Height
   
      SSFrame3.Left = trvMain.Width
      pnlHeader.Left = SSFrame3.Left
'      lblDateTime.Left = SSFrame3.Left
      lblDateTime.Width = SSFrame3.Width
   End If
   Exit Sub
ErrorHandler:
   glbErrorLog.LocalErrorMsg = "Eror"
   glbErrorLog.ShowUserError
End Sub

Private Sub InitReport1_0_6()
Dim C As CReportControl
Dim Top As Long
Dim Left As Long
Dim LabelWidth As Long
Dim Offset As Long

   Top = lblGeneric(0).Top
   Left = lblGeneric(0).Left
   LabelWidth = lblGeneric(0).Width
   Offset = 100
   
   '1 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, True, "", , "FROM_DOC_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("จากวันที่ขาย"))
   
   '1 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, True, "", , "TO_DOC_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("ถึงวันที่ขาย"))
   
   '1 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, True, "", , "TO_PAY_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("ถึงวันที่ชำระหนี้"))
   
   '1 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, True, "", , "PRINT_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("วันที่พิมพ์"))
   
   Call LoadControl("CB", cboGeneric(0).Width, True, "", 2, "CUSTOMER_TYPE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("ประเภทลูกหนี้"))
   
   '1 =============================
   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "FROM_CUSTOMER_CODE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("รหัสลูกหนี้"))
   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "TO_CUSTOMER_CODE", , True)

   '1 =============================
   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "FROM_SALE_CODE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("รหัสพนักงานขาย"))
   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "TO_SALE_CODE", , True)
   
   '3 =============================
   Call LoadControl("CB", cboGeneric(0).Width \ 2, False, "", 1, "INTERVAL_TYPE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("ประเภทช่วง"))
   
   Call LoadControl("T", txtGeneric(0).Width / 1.5, True, "", , "CREDIT")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("ระบบุเครดิต(วัน)"))
   
   Call LoadControl("T", txtGeneric(0).Width / 1.5, True, "", , "ZERO_STRING")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("ข้อความแสดงยอด 0"))
   
   Call LoadControl("CH", chkGeneric(0).Width, True, "แสดงตามตารางจริงเท่านั้น", 1, "REAL_ONLY_FLAG")
   
   Call LoadControl("CH", chkGeneric(0).Width, True, "ใช้เครดิตตามตารางจริง", 1, "REAL_CREDIT_FLAG")
   
   Call LoadControl("CH", chkGeneric(0).Width, True, "เครดิต 90 วัน(ถ้าไม่มีในตารางจริง)", 2, "NINETY_FLAG")
   
   Call LoadControl("CH", chkGeneric(0).Width, True, "พนักงานขายและลูกค้าและบิล", 2, "INCLUDE_CUSTOMER_MODE")
   
   Call LoadControl("CH", chkGeneric(0).Width, True, "พนักงานขายและลูกค้า", 2, "SALE_AND_CUSTOMER")
   
   Call LoadControl("CH", chkGeneric(0).Width / 2, True, "สรุป", 2, "SUMMARY_MODE")
   
   Call LoadControl("CH", chkGeneric(0).Width / 2, True, "แสดงจังหวัด", 2, "SHOW_PROVINCE", , True)
   
   Call LoadControl("CH", chkGeneric(0).Width / 2, True, "แสดงวงเงิน", 2, "SHOW_CRLINE")
   
   Call LoadControl("CH", chkGeneric(0).Width / 2, True, "ไม่แสดง SR", 2, "NOT_INCLUDE_SR", , True)
   
   Call LoadControl("CH", chkGeneric(0).Width / 2, True, "รับล่วงหน้า", 2, "SHOW_RCP_FUTURE")
   
   Call LoadControl("CH", chkGeneric(0).Width / 2, True, "แสดงเครดิต", 2, "SHOW_CREDIT", , True)
   
   Call LoadControl("CH", chkGeneric(0).Width, True, "เฉพาะที่เกินกำหนด", 2, "EXCEED_MODE")
   
   Call ShowControl
   Call LoadComboData
End Sub
Private Sub InitReport1_0_6_1()
Dim C As CReportControl
Dim Top As Long
Dim Left As Long
Dim LabelWidth As Long
Dim Offset As Long

   Top = lblGeneric(0).Top
   Left = lblGeneric(0).Left
   LabelWidth = lblGeneric(0).Width
   Offset = 100
   
   '1 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, True, "", , "FROM_DOC_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("จากวันที่ขาย"))
   
   '1 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, True, "", , "TO_DOC_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("ถึงวันที่ขาย"))
   
   '1 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, True, "", , "FROM_DUE_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("จากวันครบดิว"))
   
   '1 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, True, "", , "TO_DUE_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("ถึงวันที่ครบดิว"))
   
   '1 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, True, "", , "TO_PAY_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("ถึงวันที่ชำระหนี้"))
   
   '1 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, True, "", , "PRINT_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("วันที่พิมพ์"))
   
   Call LoadControl("CB", cboGeneric(0).Width, True, "", 1, "CUSTOMER_TYPE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("ประเภทลูกหนี้"))
   
   '1 =============================
   Call LoadControl("T", txtGeneric(0).Width / 1.5, True, "", , "FROM_CUSTOMER_CODE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("จากรหัสลูกหนี้"))

   '2 =============================
   Call LoadControl("T", txtGeneric(0).Width / 1.5, True, "", , "TO_CUSTOMER_CODE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("ถึงรหัสลูกหนี้"))
   
   '1 =============================
   Call LoadControl("T", txtGeneric(0).Width / 1.5, True, "", , "FROM_SALE_CODE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("จากรหัสพนักงานขาย"))
   
   '2 =============================
   Call LoadControl("T", txtGeneric(0).Width / 1.5, True, "", , "TO_SALE_CODE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("ถึงรหัสพนักงานขาย"))
   
   Call LoadControl("CH", chkGeneric(0).Width, True, "ไม่แสดงรายการลูกหนี้", 2, "NO_SHOW_CUSTOMER")
   Call LoadControl("CH", chkGeneric(0).Width, True, "ไม่แสดงรายการบิล", 2, "NO_SHOW_BILL")
   Call LoadControl("CH", chkGeneric(0).Width, True, "ไม่แสดงรายการสินค้า", 2, "NO_SHOW_ITEM")
   
   Call ShowControl
   Call LoadComboData
End Sub
Private Sub InitReport1_0_6_2()
Dim C As CReportControl
Dim Top As Long
Dim Left As Long
Dim LabelWidth As Long
Dim Offset As Long

   Top = lblGeneric(0).Top
   Left = lblGeneric(0).Left
   LabelWidth = lblGeneric(0).Width
   Offset = 100
   
   '1 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, True, "", , "FROM_DOC_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("จากวันที่ขาย"))
   
   '1 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, True, "", , "TO_DOC_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("ถึงวันที่ขาย"))
   
   '1 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, True, "", , "TO_PAY_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("ถึงวันที่ชำระหนี้"))
   
   '1 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, True, "", , "PRINT_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("วันที่พิมพ์"))
   
   Call LoadControl("CB", cboGeneric(0).Width, True, "", 1, "CUSTOMER_TYPE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("ประเภทลูกหนี้"))
   
   '1 =============================
   Call LoadControl("T", txtGeneric(0).Width / 1.5, True, "", , "FROM_CUSTOMER_CODE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("จากรหัสลูกหนี้"))

   '2 =============================
   Call LoadControl("T", txtGeneric(0).Width / 1.5, True, "", , "TO_CUSTOMER_CODE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("ถึงรหัสลูกหนี้"))

   '1 =============================
   Call LoadControl("T", txtGeneric(0).Width / 1.5, True, "", , "FROM_SALE_CODE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("จากรหัสพนักงานขาย"))
   
   '2 =============================
   Call LoadControl("T", txtGeneric(0).Width / 1.5, True, "", , "TO_SALE_CODE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("ถึงรหัสพนักงานขาย"))
   
   '3 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 2, "ORDER_BY")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("เรียงตาม"))

   '4 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 3, "ORDER_TYPE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("เรียงจาก"))
   
   '1 =============================
   Call LoadControl("T", txtGeneric(0).Width / 1.5, False, "", , "CREDIT_DAY")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("ยอด CREDIT (วัน)"))
   
   '1 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, True, "", , "FROM_MONTH")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("จากวันที่แสดง DUE"))
   
   '1 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, True, "", , "TO_MONTH")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("ถึงวันที่แสดง DUE"))
   
   Call LoadControl("CH", chkGeneric(0).Width, True, "พนักงานขายและลูกค้าและบิล", 2, "INCLUDE_CUSTOMER_MODE")
   
   Call LoadControl("CH", chkGeneric(0).Width, True, "พนักงานขายและลูกค้า", 2, "SALE_AND_CUSTOMER")
   
   Call LoadControl("CH", chkGeneric(0).Width, True, "สรุป", 2, "SUMMARY_MODE")
   
   Call ShowControl
   Call LoadComboData
End Sub

Private Sub InitReport1_0_6_3()                                                                         'แก้ไข 1-6-2-x
Dim C As CReportControl
Dim Top As Long
Dim Left As Long
Dim LabelWidth As Long
Dim Offset As Long

   Top = lblGeneric(0).Top
   Left = lblGeneric(0).Left
   LabelWidth = lblGeneric(0).Width
   Offset = 100
   
   '1 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, True, "", , "FROM_DOC_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("จากวันที่ขาย"))
   
   '1 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, True, "", , "TO_DOC_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("ถึงวันที่ขาย"))
   
   '1 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, True, "", , "TO_PAY_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("ถึงวันที่ชำระหนี้"))
   
   '1 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, True, "", , "PRINT_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("วันที่พิมพ์"))
   
   Call LoadControl("CB", cboGeneric(0).Width, True, "", 1, "CUSTOMER_TYPE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("ประเภทลูกหนี้"))
   
   '1 =============================
   Call LoadControl("T", txtGeneric(0).Width / 1.5, True, "", , "FROM_CUSTOMER_CODE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("จากรหัสลูกหนี้"))

   '2 =============================
   Call LoadControl("T", txtGeneric(0).Width / 1.5, True, "", , "TO_CUSTOMER_CODE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("ถึงรหัสลูกหนี้"))

   '1 =============================
   Call LoadControl("T", txtGeneric(0).Width / 1.5, True, "", , "FROM_SALE_CODE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("จากรหัสพนักงานขาย"))
   
   '2 =============================
   Call LoadControl("T", txtGeneric(0).Width / 1.5, True, "", , "TO_SALE_CODE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("ถึงรหัสพนักงานขาย"))
   
   '3 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 2, "ORDER_BY")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("เรียงตาม"))

   '4 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 3, "ORDER_TYPE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("เรียงจาก"))
   
   '1 =============================
   Call LoadControl("T", txtGeneric(0).Width / 1.5, False, "", , "CREDIT_DAY")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("ยอด CREDIT (วัน)"))
   
   '1 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, True, "", , "FROM_MONTH")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("จากวันที่แสดง DUE"))
   
   '1 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, True, "", , "TO_MONTH")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("ถึงวันที่แสดง DUE"))
   
   Call LoadControl("CH", chkGeneric(0).Width, True, "พนักงานขายและลูกค้าและบิล", 2, "INCLUDE_CUSTOMER_MODE")
   
   Call LoadControl("CH", chkGeneric(0).Width, True, "พนักงานขายและลูกค้า", 2, "SALE_AND_CUSTOMER")
   
   Call LoadControl("CH", chkGeneric(0).Width, True, "สรุป", 2, "SUMMARY_MODE")
   
   
   Call ShowControl
   Call LoadComboData
End Sub
Private Sub InitReport1_0_6_4()
Dim C As CReportControl
Dim Top As Long
Dim Left As Long
Dim LabelWidth As Long
Dim Offset As Long

   Top = lblGeneric(0).Top
   Left = lblGeneric(0).Left
   LabelWidth = lblGeneric(0).Width
   Offset = 100
   
   '1 =============================
   Call LoadControl("C", cboGeneric(0).Width \ 2, False, "", 1, "MONTH_ID", "MONTH_NAME")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("เดือน"))

   Call LoadControl("T", txtGeneric(0).Width / 2, False, "", , "YEAR_NO")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("ปี"))
   
   '1 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, True, "", , "FROM_DOC_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("จากวันที่ขาย"))
   
   '1 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, True, "", , "TO_DOC_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("ถึงวันที่ขาย"))
   
   '1 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, True, "", , "PRINT_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("วันที่พิมพ์"))
   
   Call LoadControl("CB", cboGeneric(0).Width, True, "", 2, "CUSTOMER_TYPE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("ประเภทลูกหนี้"))
   
   '1 =============================
   Call LoadControl("T", txtGeneric(0).Width / 1.5, True, "", , "FROM_CUSTOMER_CODE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("จากรหัสลูกหนี้"))

   '2 =============================
   Call LoadControl("T", txtGeneric(0).Width / 1.5, True, "", , "TO_CUSTOMER_CODE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("ถึงรหัสลูกหนี้"))

   '1 =============================
   Call LoadControl("T", txtGeneric(0).Width / 1.5, True, "", , "FROM_SALE_CODE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("จากรหัสพนักงานขาย"))
   
   '2 =============================
   Call LoadControl("T", txtGeneric(0).Width / 1.5, True, "", , "TO_SALE_CODE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("ถึงรหัสพนักงานขาย"))
   '3 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 3, "ORDER_BY")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("เรียงตาม"))
   Call LoadControl("C", cboGeneric(0).Width, True, "", 4, "ORDER_TYPE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("เรียงจาก"))
   
   Call LoadControl("T", txtGeneric(0).Width / 1.5, True, "", , "CREDIT")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("ระบบุเครดิต(วัน)"))
   
   Call LoadControl("CH", chkGeneric(0).Width, True, "แสดงลูกหนี้ทั้งหมด", 1, "SHOW_ALL")
   
   Call LoadControl("CH", chkGeneric(0).Width, True, "บันทึกลง EXCEL", 1, "SHOW_EXCEL")
   
   Call ShowControl
   Call LoadComboData
End Sub
Private Sub InitReport1_0_6_5()
Dim C As CReportControl
Dim Top As Long
Dim Left As Long
Dim LabelWidth As Long
Dim Offset As Long

   Top = lblGeneric(0).Top
   Left = lblGeneric(0).Left
   LabelWidth = lblGeneric(0).Width
   Offset = 100
   '1 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, True, "", , "FROM_DOC_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("จากวันที่ขาย"))
   
   '1 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, True, "", , "TO_DOC_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("ถึงวันที่ขาย"))
   
   '1 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, False, "", , "FROM_CHECK_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("จากวันที่มีผลเช็ค"))
   
   '1 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, True, "", , "TO_CHECK_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("ถึงวันที่มีผลเช็ค"))
   
   '1 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, True, "", , "TO_PAY_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("ถึงวันที่ชำระหนี้"))
   
   '1 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, True, "", , "PRINT_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("วันที่พิมพ์"))
   
      '1 =============================
   Call LoadControl("T", txtGeneric(0).Width / 1.5, False, "", , "CUSTOMER_CODE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("รหัสลูกค้า"))
   
   
   Call LoadControl("T", txtGeneric(0).Width / 1.5, True, "", , "CREDIT")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("ระบบุเครดิต(วัน)"))

      '1 =============================
   Call LoadControl("T", txtGeneric(0).Width / 1.5, True, "", , "WANT_BY_MORE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("ต้องการซื้อเพิ่ม(บาท)"))
   
      '1 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, True, "", , "SENT_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("วันที่ส่งของ(ซื้อเพิ่ม)"))
   
       '1 =============================
   Call LoadControl("T", txtGeneric(0).Width / 1.5, True, "", , "CHEQUE_RE_AMOUNT")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("ยอดเช็คคืน"))
   
   Call LoadControl("CH", chkGeneric(0).Width, True, "แสดงตามตารางจริงเท่านั้น", 1, "REAL_ONLY_FLAG")
   
   Call LoadControl("CH", chkGeneric(0).Width, True, "ใช้เครดิตตามตารางจริง", 1, "REAL_CREDIT_FLAG")
   
   Call LoadControl("CH", chkGeneric(0).Width, True, "เครดิต 90 วัน(ถ้าไม่มีในตารางจริง)", 2, "NINETY_FLAG")
   
   Call LoadControl("CH", chkGeneric(0).Width, True, "แสดงรายละเอียดเช็ครับล่วงหน้า", 1, "SHOW_DETAIL_RCP_FUTURE")
   
   Call ShowControl
   Call LoadComboData
End Sub
Private Sub InitReport1_0_7()
Dim C As CReportControl
Dim Top As Long
Dim Left As Long
Dim LabelWidth As Long
Dim Offset As Long
   
   Top = lblGeneric(0).Top
   Left = lblGeneric(0).Left
   LabelWidth = lblGeneric(0).Width
   Offset = 100
   
   '1 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, False, "", , "FROM_DOC_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("จากวันที่ขาย"))
   
   '1 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, False, "", , "TO_DOC_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("ถึงวันที่ขาย"))
   
   '1 =============================
   Call LoadControl("T", txtGeneric(0).Width / 1.5, True, "", , "FROM_CUSTOMER_CODE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("จากรหัสลูกหนี้"))

   '2 =============================
   Call LoadControl("T", txtGeneric(0).Width / 1.5, True, "", , "TO_CUSTOMER_CODE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("ถึงรหัสลูกหนี้"))
   
   Call ShowControl
   Call LoadComboData
End Sub
Private Sub InitReport1_0_8()
Dim C As CReportControl
Dim Top As Long
Dim Left As Long
Dim LabelWidth As Long
Dim Offset As Long
   
   Top = lblGeneric(0).Top
   Left = lblGeneric(0).Left
   LabelWidth = lblGeneric(0).Width
   Offset = 100
   
   '1 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, True, "", , "FROM_DOC_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("จากวันที่ขาย"))
   
   '1 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, True, "", , "TO_DOC_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("ถึงวันที่ขาย"))
   
   '1 =============================
   Call LoadControl("T", txtGeneric(0).Width / 1.5, True, "", , "FROM_CUSTOMER_CODE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("จากรหัสลูกหนี้"))

   '2 =============================
   Call LoadControl("T", txtGeneric(0).Width / 1.5, True, "", , "TO_CUSTOMER_CODE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("ถึงรหัสลูกหนี้"))
   
      '1 =============================
   Call LoadControl("T", txtGeneric(0).Width / 1.5, True, "", , "FROM_SALE_CODE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("จากรหัสพนักงานขาย"))
   
   '2 =============================
   Call LoadControl("T", txtGeneric(0).Width / 1.5, True, "", , "TO_SALE_CODE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("ถึงรหัสพนักงานขาย"))
   
      Call LoadControl("CH", chkGeneric(0).Width, True, "สรุป", 2, "SUMMARY_MODE")
   
   Call ShowControl
   Call LoadComboData
End Sub


Private Sub InitReport1_0_7_1()
Dim C As CReportControl
Dim Top As Long
Dim Left As Long
Dim LabelWidth As Long
Dim Offset As Long

   Top = lblGeneric(0).Top
   Left = lblGeneric(0).Left
   LabelWidth = lblGeneric(0).Width
   Offset = 100
   
   Call LoadControl("C", cboGeneric(0).Width \ 2, False, "", 1, "MONTH_ID", "MONTH_NAME")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("เดือน"))

   '2 =============================
   Call LoadControl("T", txtGeneric(0).Width / 2, False, "", , "YEAR_NO")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("ปี"))
   
   '2 =============================
   Call LoadControl("T", txtGeneric(0).Width / 2, False, "", , "RUN_MONTH")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("เดือนเคลื่อนไหว"))
   
   '1 =============================
   Call LoadControl("T", txtGeneric(0).Width / 1.5, True, "", , "FROM_CUSTOMER_CODE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("จากรหัสลูกหนี้"))

   '2 =============================
   Call LoadControl("T", txtGeneric(0).Width / 1.5, True, "", , "TO_CUSTOMER_CODE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("ถึงรหัสลูกหนี้"))
   
   Call LoadControl("T", txtGeneric(0).Width / 1.5, True, "", , "FROM_SALE_CODE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("จากพนักงานขาย"))

   Call LoadControl("T", txtGeneric(0).Width / 1.5, True, "", , "TO_SALE_CODE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("ถึงพนักงานขาย"))
   
   Call LoadControl("CH", chkGeneric(0).Width, True, "เฉพาะรายการรวมลูกค้า", 1, "ONLY_CUSTOMER")
   Call LoadControl("CH", chkGeneric(0).Width, True, "สรุป", 1, "SUMMARY_FLAG")
   
   Call ShowControl
   Call LoadComboData
End Sub


Private Sub InitReport1_0_8_1()
Dim C As CReportControl
Dim Top As Long
Dim Left As Long
Dim LabelWidth As Long
Dim Offset As Long
   
   Top = lblGeneric(0).Top
   Left = lblGeneric(0).Left
   LabelWidth = lblGeneric(0).Width
   Offset = 100
   
   '1 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, True, "", , "FROM_DOC_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("จากวันที่ขาย"))
   
   '1 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, True, "", , "TO_DOC_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("ถึงวันที่ขาย"))

    Call LoadControl("C", cboGeneric(0).Width, False, "", 1, "CUSTOMER_TYPE_ID", "CUSTOMER_TYPE_NAME")
    Call LoadControl("L", lblGeneric(0).Width, True, MapText("ประเภทลูกค้า"))
    
   
   Call ShowControl
   Call LoadComboData
End Sub
Private Sub InitReport1_0_9()
Dim C As CReportControl
Dim Top As Long
Dim Left As Long
Dim LabelWidth As Long
Dim Offset As Long
   
   Top = lblGeneric(0).Top
   Left = lblGeneric(0).Left
   LabelWidth = lblGeneric(0).Width
   Offset = 100
   
   '1 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, True, "", , "FROM_DOC_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("จากวันที่รับชำระ"))
   
   '1 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, True, "", , "TO_DOC_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("ถึงวันที่รับชำระ"))
      
   '1 =============================
   Call LoadControl("T", txtGeneric(0).Width / 1.5, False, "", , "YYYYMM")
   Call LoadControl("L", lblGeneric(0).Width, False, MapText("แจกแจงปีเดือน(YYYYMM)"))
      
   '1 =============================
   Call LoadControl("T", txtGeneric(0).Width / 1.5, True, "", , "FROM_CUSTOMER_CODE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("จากรหัสลูกหนี้"))

   '2 =============================
   Call LoadControl("T", txtGeneric(0).Width / 1.5, True, "", , "TO_CUSTOMER_CODE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("ถึงรหัสลูกหนี้"))
   
   Call ShowControl
   Call LoadComboData
End Sub
Private Sub InitReport1_0_10()
Dim C As CReportControl
Dim Top As Long
Dim Left As Long
Dim LabelWidth As Long
Dim Offset As Long
   
   Top = lblGeneric(0).Top
   Left = lblGeneric(0).Left
   LabelWidth = lblGeneric(0).Width
   Offset = 100
   
   Call LoadControl("C", cboGeneric(0).Width \ 2, False, "", 1, "FROM_MONTH_ID", "FROM_MONTH_NAME")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("จากเดือน"))

   '2 =============================
   Call LoadControl("T", txtGeneric(0).Width / 2, False, "", , "FROM_YEAR_NO")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("จากปี"))
   
   Call LoadControl("C", cboGeneric(0).Width \ 2, False, "", 2, "TO_MONTH_ID", "TO_MONTH_NAME")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("ถึงเดือน"))

   '2 =============================
   Call LoadControl("T", txtGeneric(0).Width / 2, False, "", , "TO_YEAR_NO")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("ถึงปี"))
   
   
   '1 =============================
   Call LoadControl("T", txtGeneric(0).Width / 1.5, True, "", , "FROM_CUSTOMER_CODE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("จากรหัสลูกหนี้"))

   '2 =============================
   Call LoadControl("T", txtGeneric(0).Width / 1.5, True, "", , "TO_CUSTOMER_CODE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("ถึงรหัสลูกหนี้"))
      
   '1 =============================
   Call LoadControl("T", txtGeneric(0).Width / 1.5, True, "", , "FROM_STOCK_CODE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("จากรหัสสินค้า"))

   '2 =============================
   Call LoadControl("T", txtGeneric(0).Width / 1.5, True, "", , "TO_STOCK_CODE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("ถึงรหัสสินค้า"))
   
   Call LoadControl("CH", chkGeneric(0).Width, True, "รวมของแถม", 1, "FREE_FLAG")
   Call LoadControl("CH", chkGeneric(0).Width, True, "รวมของขาย", 1, "SALE_FLAG")
   Call LoadControl("CH", chkGeneric(0).Width, True, "สรุป", 1, "SUMMARY_FLAG")
   
   Call LoadControl("CH", chkGeneric(0).Width, True, "แสดงเฉพาะจำนวน", 1, "ONLY_AMOUNT")
   
   Call ShowControl
   Call LoadComboData
End Sub
Private Sub InitReport1_0_10_1()
Dim C As CReportControl
Dim Top As Long
Dim Left As Long
Dim LabelWidth As Long
Dim Offset As Long
   
   Top = lblGeneric(0).Top
   Left = lblGeneric(0).Left
   LabelWidth = lblGeneric(0).Width
   Offset = 100
   
   Call LoadControl("C", cboGeneric(0).Width \ 2, False, "", 1, "FROM_MONTH_ID", "FROM_MONTH_NAME")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("จากเดือน"))

   '2 =============================
   Call LoadControl("T", txtGeneric(0).Width / 2, False, "", , "FROM_YEAR_NO")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("จากปี"))
   
   Call LoadControl("C", cboGeneric(0).Width \ 2, False, "", 2, "TO_MONTH_ID", "TO_MONTH_NAME")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("ถึงเดือน"))

   '2 =============================
   Call LoadControl("T", txtGeneric(0).Width / 2, False, "", , "TO_YEAR_NO")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("ถึงปี"))
   
   
   '1 =============================
   Call LoadControl("T", txtGeneric(0).Width / 1.5, True, "", , "FROM_CUSTOMER_CODE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("จากรหัสลูกหนี้"))

   '2 =============================
   Call LoadControl("T", txtGeneric(0).Width / 1.5, True, "", , "TO_CUSTOMER_CODE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("ถึงรหัสลูกหนี้"))
      
   '1 =============================
   Call LoadControl("T", txtGeneric(0).Width / 1.5, True, "", , "FROM_STOCK_CODE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("จากรหัสสินค้า"))

   '2 =============================
   Call LoadControl("T", txtGeneric(0).Width / 1.5, True, "", , "TO_STOCK_CODE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("ถึงรหัสสินค้า"))
   
   Call LoadControl("T", txtGeneric(0).Width / 1.5, True, "", , "FROM_SALE_CODE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("จากพนักงานขาย"))
   
   Call LoadControl("T", txtGeneric(0).Width / 1.5, True, "", , "TO_SALE_CODE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("ถึงพนักงานขาย"))
   
   Call LoadControl("CH", chkGeneric(0).Width, True, "รวมของแถม", 1, "FREE_FLAG")
   Call LoadControl("CH", chkGeneric(0).Width, True, "รวมของขาย", 1, "SALE_FLAG")
   Call LoadControl("CH", chkGeneric(0).Width, True, "สรุป", 1, "SUMMARY_FLAG")
   
   Call LoadControl("CH", chkGeneric(0).Width, True, "แสดงเฉพาะจำนวน", 1, "ONLY_AMOUNT")
   Call LoadControl("CH", chkGeneric(0).Width, True, "แสดงเฉพาะมูลค่า", 1, "ONLY_PRICE")
   
   Call ShowControl
   Call LoadComboData
End Sub
Private Sub InitReport1_0_10_2()
Dim C As CReportControl
Dim Top As Long
Dim Left As Long
Dim LabelWidth As Long
Dim Offset As Long
   
   Top = lblGeneric(0).Top
   Left = lblGeneric(0).Left
   LabelWidth = lblGeneric(0).Width
   Offset = 100
   
   Call LoadControl("C", cboGeneric(0).Width \ 2, False, "", 1, "FROM_MONTH_ID", "FROM_MONTH_NAME")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("จากเดือน"))

   '2 =============================
   Call LoadControl("T", txtGeneric(0).Width / 2, False, "", , "FROM_YEAR_NO")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("จากปี"))
   
   Call LoadControl("C", cboGeneric(0).Width \ 2, False, "", 2, "TO_MONTH_ID", "TO_MONTH_NAME")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("ถึงเดือน"))

   '2 =============================
   Call LoadControl("T", txtGeneric(0).Width / 2, False, "", , "TO_YEAR_NO")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("ถึงปี"))
   
   
   '1 =============================
   Call LoadControl("T", txtGeneric(0).Width / 1.5, True, "", , "FROM_CUSTOMER_CODE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("จากรหัสลูกหนี้"))

   '2 =============================
   Call LoadControl("T", txtGeneric(0).Width / 1.5, True, "", , "TO_CUSTOMER_CODE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("ถึงรหัสลูกหนี้"))
      
   '1 =============================
   Call LoadControl("T", txtGeneric(0).Width / 1.5, False, "", , "FROM_STOCK_CODE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("จากรหัสสินค้า"))

   '2 =============================
   Call LoadControl("T", txtGeneric(0).Width / 1.5, False, "", , "TO_STOCK_CODE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("ถึงรหัสสินค้า"))
   
'   Call LoadControl("T", txtGeneric(0).Width / 1.5, True, "", , "FROM_SALE_CODE")
'   Call LoadControl("L", lblGeneric(0).Width, True, MapText("จากพนักงานขาย"))
'
'   Call LoadControl("T", txtGeneric(0).Width / 1.5, True, "", , "TO_SALE_CODE")
'   Call LoadControl("L", lblGeneric(0).Width, True, MapText("ถึงพนักงานขาย"))
   
'   Call LoadControl("CH", chkGeneric(0).Width, True, "รวมของแถม", 1, "FREE_FLAG")
'   Call LoadControl("CH", chkGeneric(0).Width, True, "รวมของขาย", 1, "SALE_FLAG")
'   Call LoadControl("CH", chkGeneric(0).Width, True, "สรุป", 1, "SUMMARY_FLAG")
   
   Call LoadControl("CH", chkGeneric(0).Width, True, "แสดงเฉพาะจำนวน", 1, "ONLY_AMOUNT")
   Call LoadControl("CH", chkGeneric(0).Width, True, "แสดงเฉพาะมูลค่า", 1, "ONLY_PRICE")
   
   Call LoadControl("CH", chkGeneric(0).Width, True, "บันทึกลง EXCEL", 1, "SHOW_EXCEL")
   
   Call ShowControl
   Call LoadComboData
End Sub

Private Sub InitReport1_0_10_3()
Dim C As CReportControl
Dim Top As Long
Dim Left As Long
Dim LabelWidth As Long
Dim Offset As Long
   
   Top = lblGeneric(0).Top
   Left = lblGeneric(0).Left
   LabelWidth = lblGeneric(0).Width
   Offset = 100
   
   Call LoadControl("C", cboGeneric(0).Width \ 2, False, "", 1, "FROM_MONTH_ID", "FROM_MONTH_NAME")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("จากเดือน"))

   '2 =============================
   Call LoadControl("T", txtGeneric(0).Width / 2, False, "", , "FROM_YEAR_NO")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("จากปี"))
   
   Call LoadControl("C", cboGeneric(0).Width \ 2, False, "", 2, "TO_MONTH_ID", "TO_MONTH_NAME")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("ถึงเดือน"))

   '2 =============================
   Call LoadControl("T", txtGeneric(0).Width / 2, False, "", , "TO_YEAR_NO")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("ถึงปี"))
   
   
   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "FROM_CUSTOMER_CODE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("รหัสลูกหนี้"))
   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "TO_CUSTOMER_CODE", , True)
    
   '1 =============================
   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "FROM_STOCK_CODE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("รหัสสินค้า"))
   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "TO_STOCK_CODE", , True)
   
   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "FROM_SALE_CODE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("พนักงานขาย"))
   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "TO_SALE_CODE", , True)
   
   '1 =============================
   
   Call LoadControl("CH", chkGeneric(0).Width, True, "รวมของแถม", 1, "FREE_FLAG")
   Call LoadControl("CH", chkGeneric(0).Width, True, "รวมของขาย", 1, "SALE_FLAG")
   
   Call LoadControl("CH", chkGeneric(0).Width, True, "แสดงของขายของแถมรายสินค้าพนักงานขายสุทธิ", 1, "SHOW_SALE_AND_FREE")
   
   Call LoadControl("CH", chkGeneric(0).Width, True, "สรุป", 1, "SUMMARY_FLAG")
   
   Call LoadControl("CH", chkGeneric(0).Width, True, "แสดงเฉพาะจำนวน", 1, "ONLY_AMOUNT")
   Call LoadControl("CH", chkGeneric(0).Width, True, "แสดงเฉพาะมูลค่า", 1, "ONLY_PRICE")
   
   Call LoadControl("CH", chkGeneric(0).Width, True, "หักลบของแถมปลายปี", 1, "PROMOTION")
   
   Call ShowControl
   Call LoadComboData
End Sub

Private Sub InitReport1_0_10_4()
Dim C As CReportControl
Dim Top As Long
Dim Left As Long
Dim LabelWidth As Long
Dim Offset As Long
   
   Top = lblGeneric(0).Top
   Left = lblGeneric(0).Left
   LabelWidth = lblGeneric(0).Width
   Offset = 100
   
   Call LoadControl("C", cboGeneric(0).Width \ 2, False, "", 1, "FROM_MONTH_ID", "FROM_MONTH_NAME")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("จากเดือน"))

   '2 =============================
   Call LoadControl("T", txtGeneric(0).Width / 2, False, "", , "FROM_YEAR_NO")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("จากปี"))
   
   Call LoadControl("C", cboGeneric(0).Width \ 2, False, "", 2, "TO_MONTH_ID", "TO_MONTH_NAME")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("ถึงเดือน"))

   '2 =============================
   Call LoadControl("T", txtGeneric(0).Width / 2, False, "", , "TO_YEAR_NO")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("ถึงปี"))
   
   
   '1 =============================
   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "FROM_CUSTOMER_CODE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("รหัสลูกหนี้"))
   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "TO_CUSTOMER_CODE", , True)
    
   '1 =============================
   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "FROM_STOCK_CODE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("รหัสสินค้า"))
   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "TO_STOCK_CODE", , True)
   
   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "FROM_SALE_CODE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("พนักงานขาย"))
   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "TO_SALE_CODE", , True)
   
   Call LoadControl("CH", chkGeneric(0).Width, True, "รวมของแถม", 1, "FREE_FLAG")
   Call LoadControl("CH", chkGeneric(0).Width, True, "รวมของขาย", 1, "SALE_FLAG")
   
   Call LoadControl("CH", chkGeneric(0).Width, True, "แสดงของขายของแถมรายสินค้าพนักงานขายสุทธิ", 1, "SHOW_SALE_AND_FREE")
   
   Call LoadControl("CH", chkGeneric(0).Width, True, "สรุป", 1, "SUMMARY_FLAG")
   

   Call LoadControl("CH", chkGeneric(0).Width, True, "แสดงเฉพาะจำนวน", 1, "ONLY_AMOUNT")
   Call LoadControl("CH", chkGeneric(0).Width, True, "แสดงเฉพาะมูลค่า", 1, "ONLY_PRICE")
   
   Call LoadControl("CH", chkGeneric(0).Width, True, "หักลบของแถมปลายปี", 1, "PROMOTION")
   
   Call ShowControl
   Call LoadComboData
End Sub

Private Sub InitReport1_0_11()
Dim C As CReportControl
Dim Top As Long
Dim Left As Long
Dim LabelWidth As Long
Dim Offset As Long
   
   Top = lblGeneric(0).Top
   Left = lblGeneric(0).Left
   LabelWidth = lblGeneric(0).Width
   Offset = 100
   
   Call LoadControl("C", cboGeneric(0).Width \ 2, False, "", 1, "FROM_MONTH_ID", "FROM_MONTH_NAME")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("จากเดือน"))

   '2 =============================
   Call LoadControl("T", txtGeneric(0).Width / 2, False, "", , "FROM_YEAR_NO")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("จากปี"))
   
   Call LoadControl("C", cboGeneric(0).Width \ 2, False, "", 2, "TO_MONTH_ID", "TO_MONTH_NAME")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("ถึงเดือน"))

   '2 =============================
   Call LoadControl("T", txtGeneric(0).Width / 2, False, "", , "TO_YEAR_NO")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("ถึงปี"))
   
   
   '1 =============================
   Call LoadControl("T", txtGeneric(0).Width / 1.5, True, "", , "FROM_CUSTOMER_CODE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("จากรหัสลูกหนี้"))

   '2 =============================
   Call LoadControl("T", txtGeneric(0).Width / 1.5, True, "", , "TO_CUSTOMER_CODE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("ถึงรหัสลูกหนี้"))
      
   '1 =============================
   Call LoadControl("T", txtGeneric(0).Width / 1.5, True, "", , "FROM_STOCK_CODE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("จากรหัสสินค้า"))

   '2 =============================
   Call LoadControl("T", txtGeneric(0).Width / 1.5, True, "", , "TO_STOCK_CODE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("ถึงรหัสสินค้า"))
   
   Call LoadControl("T", txtGeneric(0).Width / 1.5, True, "", , "FROM_SALE_CODE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("จากพนักงานขาย"))
   
   Call LoadControl("T", txtGeneric(0).Width / 1.5, True, "", , "TO_SALE_CODE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("ถึงพนักงานขาย"))
   
   Call LoadControl("CH", chkGeneric(0).Width, True, "รวมของแถม", 1, "FREE_FLAG")
   Call LoadControl("CH", chkGeneric(0).Width, True, "รวมของขาย", 1, "SALE_FLAG")
   
   Call LoadControl("CH", chkGeneric(0).Width, True, "แสดงของขายของแถมรายสินค้าพนักงานขายสุทธิ", 1, "SHOW_SALE_AND_FREE")
   
   Call LoadControl("CH", chkGeneric(0).Width, True, "สรุป", 1, "SUMMARY_FLAG")
   
   Call LoadControl("CH", chkGeneric(0).Width, True, "แสดงเฉพาะจำนวน", 1, "ONLY_AMOUNT")
   Call LoadControl("CH", chkGeneric(0).Width, True, "แสดงเฉพาะมูลค่า", 1, "ONLY_PRICE")
   
   Call ShowControl
   Call LoadComboData
End Sub
Private Sub InitReport1_0_12()
Dim C As CReportControl
Dim Top As Long
Dim Left As Long
Dim LabelWidth As Long
Dim Offset As Long
   
   Top = lblGeneric(0).Top
   Left = lblGeneric(0).Left
   LabelWidth = lblGeneric(0).Width
   Offset = 100
   
   '1 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, True, "", , "FROM_DOC_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("จากวันที่ขาย"))
   
   '1 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, True, "", , "TO_DOC_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("ถึงวันที่ขาย"))
   
   '1 =============================
   Call LoadControl("T", txtGeneric(0).Width / 1.5, True, "", , "FROM_CUSTOMER_CODE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("จากรหัสลูกหนี้"))

   '2 =============================
   Call LoadControl("T", txtGeneric(0).Width / 1.5, True, "", , "TO_CUSTOMER_CODE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("ถึงรหัสลูกหนี้"))
   
   Call ShowControl
   Call LoadComboData
End Sub
Private Sub InitReport1_0_12_1()
Dim C As CReportControl
Dim Top As Long
Dim Left As Long
Dim LabelWidth As Long
Dim Offset As Long
   
   Top = lblGeneric(0).Top
   Left = lblGeneric(0).Left
   LabelWidth = lblGeneric(0).Width
   Offset = 100
   
   '1 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, True, "", , "FROM_DOC_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("จากวันที่"))
   
   '1 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, True, "", , "TO_DOC_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("ถึงวันที่"))
   
    Call LoadControl("C", cboGeneric(0).Width, False, "", 1, "BANK_ID", "BANK_NAME")
    Call LoadControl("L", lblGeneric(0).Width, True, MapText("ธนาคาร"))
    
   Call LoadControl("C", cboGeneric(0).Width, False, "", 2, "CUSTOMER_ID", "CUSTOMER_NAME")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("ลูกหนี้"))
    
   Call LoadControl("CH", chkGeneric(0).Width, True, "บันทึกลง EXCEL", 1, "SHOW_EXCEL")
   
   Call LoadControl("CH", chkGeneric(0).Width, True, "แสดงประมาณการ", 1, "SHOW_BUDGET")
   
   Call ShowControl
   Call LoadComboData
End Sub
Private Sub InitReport2_0_5()
Dim C As CReportControl
Dim Top As Long
Dim Left As Long
Dim LabelWidth As Long
Dim Offset As Long

   Top = lblGeneric(0).Top
   Left = lblGeneric(0).Left
   LabelWidth = lblGeneric(0).Width
   Offset = 100
   
   '1 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, True, "", , "FROM_DOC_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("จากวันที่รับของ"))
   
   '1 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, True, "", , "TO_DOC_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("ถึงวันที่รับของ"))
   
   '1 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, True, "", , "FROM_DUE_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("จากวันครบดิว"))
   
   '1 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, True, "", , "TO_DUE_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("ถึงวันที่ครบดิว"))
   
   '1 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, True, "", , "TO_PAY_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("ถึงวันที่ชำระหนี้"))
   
   '1 =============================
   Call LoadControl("T", txtGeneric(0).Width / 1.5, True, "", , "FROM_SUPPLIER_CODE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("จากรหัสผู้จำหน่าย"))
   
   '2 =============================
   Call LoadControl("T", txtGeneric(0).Width / 1.5, True, "", , "TO_SUPPLIER_CODE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("ถึงรหัสผู้จำหน่าย"))
   
   '3 =============================
   Call LoadControl("CB", cboGeneric(0).Width, True, "", 1, "SUPPLIER_TYPE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("ประเภทผู้จำหน่าย"))
               
   '2 =============================
   Call LoadControl("T", txtGeneric(0).Width / 1.5, True, "", , "SUPPLIER_TYPE_SET")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("ประเภทผู้จำหน่าย"))
   
    '1 =============================
   Call LoadControl("T", txtGeneric(0).Width / 1.5, True, "", , "FROM_STOCK_CODE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("จากรหัส RM"))
   
   '2 =============================
   Call LoadControl("T", txtGeneric(0).Width / 1.5, True, "", , "TO_STOCK_CODE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("ถึงรหัส  RM"))
   
   Call LoadControl("CH", chkGeneric(0).Width, True, "ไม่แสดงบิล", 1, "NO_SHOW_BILL")
   Call LoadControl("CH", chkGeneric(0).Width, True, "แสดงยอดค้าง", 1, "SHOW_SUM_PAID")
   
   Call ShowControl
   Call LoadComboData
End Sub
Private Sub InitReport2_0_5_1()
Dim C As CReportControl
Dim Top As Long
Dim Left As Long
Dim LabelWidth As Long
Dim Offset As Long

   Top = lblGeneric(0).Top
   Left = lblGeneric(0).Left
   LabelWidth = lblGeneric(0).Width
   Offset = 100
   
   '1 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, True, "", , "FROM_DOC_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("จากวันที่รับของ"))
   
   '1 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, True, "", , "TO_DOC_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("ถึงวันที่รับของ"))
   
   '1 =============================
   Call LoadControl("T", txtGeneric(0).Width / 1.5, True, "", , "FROM_SUPPLIER_CODE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("จากรหัสผู้จำหน่าย"))
   
   '2 =============================
   Call LoadControl("T", txtGeneric(0).Width / 1.5, True, "", , "TO_SUPPLIER_CODE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("ถึงรหัสผู้จำหน่าย"))
   
   '3 =============================
   Call LoadControl("CB", cboGeneric(0).Width, True, "", 1, "SUPPLIER_TYPE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("ประเภทผู้จำหน่าย"))
               
   '2 =============================
   Call LoadControl("T", txtGeneric(0).Width / 1.5, True, "", , "SUPPLIER_TYPE_SET")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("ประเภทผู้จำหน่าย"))
   
    '1 =============================
   Call LoadControl("T", txtGeneric(0).Width / 1.5, True, "", , "FROM_STOCK_CODE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("จากรหัส RM"))
   
   '2 =============================
   Call LoadControl("T", txtGeneric(0).Width / 1.5, True, "", , "TO_STOCK_CODE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("ถึงรหัส  RM"))
   
   Call LoadControl("CH", chkGeneric(0).Width, True, "ไม่แสดงบิล", 1, "NO_SHOW_BILL")
   Call LoadControl("CH", chkGeneric(0).Width, True, "สรุปใบปะหน้าผู้ขาย", 1, "SUMMARY_MODE")
   
   Call LoadControl("CH", chkGeneric(0).Width, True, "สรุปแนว Column", 1, "SUMMARY_COLUMN")
   
   
   Call ShowControl
   Call LoadComboData
End Sub

Private Sub InitReport2_0_6()
Dim C As CReportControl
Dim Top As Long
Dim Left As Long
Dim LabelWidth As Long
Dim Offset As Long

   Top = lblGeneric(0).Top
   Left = lblGeneric(0).Left
   LabelWidth = lblGeneric(0).Width
   Offset = 100
   
   Call LoadControl("C", cboGeneric(0).Width \ 2, False, "", 1, "MONTH_ID", "MONTH_NAME")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("เดือน"))

   '2 =============================
   Call LoadControl("T", txtGeneric(0).Width / 2, False, "", , "YEAR_NO")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("ปี"))
   
   '1 =============================
   Call LoadControl("T", txtGeneric(0).Width / 1.5, True, "", , "FROM_SUPPLIER_CODE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("จากรหัสผู้จำหน่าย"))
   
   '2 =============================
   Call LoadControl("T", txtGeneric(0).Width / 1.5, True, "", , "TO_SUPPLIER_CODE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("ถึงรหัสผู้จำหน่าย"))
   
   Call LoadControl("C", cboGeneric(0).Width, False, "", 3, "DATA_TYPE_ID", "DATA_TYPE_NAME")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("ประเภทข้อมูล"))
   
   '3 =============================
'   Call LoadControl("CB", cboGeneric(0).Width, True, "", 2, "SUPPLIER_TYPE")
'   Call LoadControl("L", lblGeneric(0).Width, True, MapText("ประเภทผู้จำหน่าย"))
               
'   '2 =============================
'   Call LoadControl("T", txtGeneric(0).Width / 1.5, True, "", , "SUPPLIER_TYPE_SET")
'   Call LoadControl("L", lblGeneric(0).Width, True, MapText("ประเภทผู้จำหน่าย"))
   
   Call ShowControl
   Call LoadComboData
End Sub
Private Sub InitReport2_0_6_1()
Dim C As CReportControl
Dim Top As Long
Dim Left As Long
Dim LabelWidth As Long
Dim Offset As Long

   Top = lblGeneric(0).Top
   Left = lblGeneric(0).Left
   LabelWidth = lblGeneric(0).Width
   Offset = 100
   
   Call LoadControl("C", cboGeneric(0).Width \ 2, False, "", 1, "MONTH_ID", "MONTH_NAME")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("เดือน"))

   '2 =============================
   Call LoadControl("T", txtGeneric(0).Width / 2, False, "", , "YEAR_NO")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("ปี"))
   
   Call LoadControl("C", cboGeneric(0).Width, False, "", 3, "DATA_TYPE_ID", "DATA_TYPE_NAME")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("ประเภทข้อมูล"))
   
   '3 =============================
'   Call LoadControl("CB", cboGeneric(0).Width, True, "", 2, "SUPPLIER_TYPE")
'   Call LoadControl("L", lblGeneric(0).Width, True, MapText("ประเภทผู้จำหน่าย"))
               
'   '2 =============================
'   Call LoadControl("T", txtGeneric(0).Width / 1.5, True, "", , "SUPPLIER_TYPE_SET")
'   Call LoadControl("L", lblGeneric(0).Width, True, MapText("ประเภทผู้จำหน่าย"))
   
   Call LoadControl("CH", chkGeneric(0).Width, True, "สรุป", 1, "SUMMARY_FLAG")
   
   Call ShowControl
   Call LoadComboData
End Sub
Private Sub InitReport2_0_6_2()
Dim C As CReportControl
Dim Top As Long
Dim Left As Long
Dim LabelWidth As Long
Dim Offset As Long

   Top = lblGeneric(0).Top
   Left = lblGeneric(0).Left
   LabelWidth = lblGeneric(0).Width
   Offset = 100
   
   Call LoadControl("C", cboGeneric(0).Width \ 2, False, "", 1, "MONTH_ID", "MONTH_NAME")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("เดือน"))

   '2 =============================
   Call LoadControl("T", txtGeneric(0).Width / 2, False, "", , "YEAR_NO")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("ปี"))
   
   '2 =============================
   Call LoadControl("T", txtGeneric(0).Width / 2, False, "", , "RUN_MONTH")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("เดือนเคลื่อนไหว"))
   '3 =============================
'   Call LoadControl("CB", cboGeneric(0).Width, True, "", 2, "SUPPLIER_TYPE")
'   Call LoadControl("L", lblGeneric(0).Width, True, MapText("ประเภทผู้จำหน่าย"))
               
'   '2 =============================
'   Call LoadControl("T", txtGeneric(0).Width / 1.5, True, "", , "SUPPLIER_TYPE_SET")
'   Call LoadControl("L", lblGeneric(0).Width, True, MapText("ประเภทผู้จำหน่าย"))
   
   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "SUPPLIER_CODE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("ผู้จำหน่าย"))
   
   Call LoadControl("C", cboGeneric(0).Width, False, "", 3, "DATA_TYPE_ID", "DATA_TYPE_NAME")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("ประเภทข้อมูล"))
   
   Call LoadControl("CH", chkGeneric(0).Width, True, "เฉพาะรายการรวมผู้จำหน่าย", 1, "ONLY_SUPPLIER")
   Call LoadControl("CH", chkGeneric(0).Width, True, "สรุป", 1, "SUMMARY_FLAG")
   
   Call ShowControl
   Call LoadComboData
End Sub

Private Sub InitReport2_0_7()
Dim C As CReportControl
Dim Top As Long
Dim Left As Long
Dim LabelWidth As Long
Dim Offset As Long

   Top = lblGeneric(0).Top
   Left = lblGeneric(0).Left
   LabelWidth = lblGeneric(0).Width
   Offset = 100
   
   '1 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, True, "", , "FROM_DOC_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("จากวันที่"))
   
   '1 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, True, "", , "TO_DOC_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("ถึงวันที่"))
   
   '1 =============================
   Call LoadControl("T", txtGeneric(0).Width / 1.5, True, "", , "FROM_SUPPLIER_CODE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("จากรหัสผู้จำหน่าย"))
   
   '2 =============================
   Call LoadControl("T", txtGeneric(0).Width / 1.5, True, "", , "TO_SUPPLIER_CODE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("ถึงรหัสผู้จำหน่าย"))
   
   Call LoadControl("C", cboGeneric(0).Width, False, "", 3, "DATA_TYPE_ID", "DATA_TYPE_NAME")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("ประเภทข้อมูล"))
   
   '3 =============================
'   Call LoadControl("CB", cboGeneric(0).Width, True, "", 2, "SUPPLIER_TYPE")
'   Call LoadControl("L", lblGeneric(0).Width, True, MapText("ประเภทผู้จำหน่าย"))
               
'   '2 =============================
'   Call LoadControl("T", txtGeneric(0).Width / 1.5, True, "", , "SUPPLIER_TYPE_SET")
'   Call LoadControl("L", lblGeneric(0).Width, True, MapText("ประเภทผู้จำหน่าย"))
   
   Call ShowControl
   Call LoadComboData
End Sub
Private Sub InitReport2_0_8()
Dim C As CReportControl
Dim Top As Long
Dim Left As Long
Dim LabelWidth As Long
Dim Offset As Long

   Top = lblGeneric(0).Top
   Left = lblGeneric(0).Left
   LabelWidth = lblGeneric(0).Width
   Offset = 100
   
   '1 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, True, "", , "FROM_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("จากวันที่"))
   
   '1 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, True, "", , "TO_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("ถึงวันที่"))
         
   '1 =============================
   Call LoadControl("T", txtGeneric(0).Width / 1.5, True, "", , "SUPPLIER_CODE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("รหัสผู้ขาย"))

   '1 =============================
   Call LoadControl("T", txtGeneric(0).Width / 1.5, True, "", , "PART_NO")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("รหัสสินค้า/วัตถุดิบ"))
   
   '1 =============================
   Call LoadControl("T", txtGeneric(0).Width / 1.5, True, "", , "PART_DESC")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("สินค้า/วัตถุดิบ"))
   
'    '1 =============================
'   Call LoadControl("C", cboGeneric(0).Width, True, "", 1, "RO_TYPE")
'   Call LoadControl("L", lblGeneric(0).Width, True, MapText("ประเภท RO"))
'
   '3 =============================
   Call LoadControl("CB", cboGeneric(0).Width, True, "", 2, "SUPPLIER_TYPE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("ประเภทผู้ขาย"))
   
   Call LoadControl("T", txtGeneric(0).Width / 1.5, True, "", , "AMPHUR")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("อำเภอ/เขต"))
   
   Call LoadControl("T", txtGeneric(0).Width / 1.5, True, "", , "PROVINCE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("จังหวัด"))
   
   Call ShowControl
   Call LoadComboData
End Sub
Private Sub InitReport1_0_5_1()
Dim C As CReportControl
Dim Top As Long
Dim Left As Long
Dim LabelWidth As Long
Dim Offset As Long

   Top = lblGeneric(0).Top
   Left = lblGeneric(0).Left
   LabelWidth = lblGeneric(0).Width
   Offset = 100
   
   '1 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, False, "", , "FROM_CHECK_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("จากวันที่ดิวเช็ค"))
   
   '1 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, True, "", , "TO_CHECK_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("ถึงวันที่ดิวเช็ค"))
      
   '1 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, True, "", , "PRINT_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("วันที่พิมพ์"))
   
   '1 =============================
'   Call LoadControl("T", txtGeneric(0).Width / 1.5, True, "", , "FROM_SUPPLIER_CODE")
'   Call LoadControl("L", lblGeneric(0).Width, True, MapText("จากรหัสผู้จำหน่าย"))
   
   '2 =============================
   Call LoadControl("T", txtGeneric(0).Width / 1.5, True, "", , "SUPPLIER_CODE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("รหัสผู้จำหน่าย"))

   '2 =============================
   Call LoadControl("T", txtGeneric(0).Width / 1.5, False, "", , "FROM_DOCUMENT_NO")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("จากรหัสใบสำคัญจ่าย"))
   
   '2 =============================
   Call LoadControl("T", txtGeneric(0).Width / 1.5, False, "", , "TO_DOCUMENT_NO")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("ถึงรหัสใบสำคัญจ่าย"))
   
   '2 =============================
   Call LoadControl("T", txtGeneric(0).Width / 1.5, False, "", , "FROM_DOCUMENT_NO1")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("หรือจากรหัสใบสำคัญจ่าย"))
   
   '2 =============================
   Call LoadControl("T", txtGeneric(0).Width / 1.5, False, "", , "TO_DOCUMENT_NO1")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("หรือถึงรหัสใบสำคัญจ่าย"))
   
   Call LoadControl("D", uctlGenericDate(0).Width, True, "", , "TO_PAY_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("ถึงวันที่ชำระหนี้"))
   
   '3 =============================
   Call LoadControl("CB", cboGeneric(0).Width, True, "", 1, "SUPPLIER_TYPE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("ประเภทผู้จำหน่าย"))

   '3 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 2, "ORDER_BY")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("เรียงตาม"))

   '4 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 3, "ORDER_TYPE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("เรียงจาก"))
      
   Call ShowControl
   Call LoadComboData
End Sub

Private Sub InitReport1_0_5_2()
Dim C As CReportControl
Dim Top As Long
Dim Left As Long
Dim LabelWidth As Long
Dim Offset As Long

   Top = lblGeneric(0).Top
   Left = lblGeneric(0).Left
   LabelWidth = lblGeneric(0).Width
   Offset = 100
   
   '1 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, False, "", , "FROM_CHECK_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("จากวันที่ดิวเช็ค"))
   
   Call LoadControl("T", txtGeneric(0).Width / 1.5, False, "", , "PERIOD_AMOUNT")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("ช่วงวันที่"))
      
   '1 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, True, "", , "PRINT_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("วันที่พิมพ์"))
   
   '1 =============================
'   Call LoadControl("T", txtGeneric(0).Width / 1.5, True, "", , "FROM_SUPPLIER_CODE")
'   Call LoadControl("L", lblGeneric(0).Width, True, MapText("จากรหัสผู้จำหน่าย"))
   
   '2 =============================
   Call LoadControl("T", txtGeneric(0).Width / 1.5, True, "", , "SUPPLIER_CODE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("รหัสผู้จำหน่าย"))

   '2 =============================
   Call LoadControl("T", txtGeneric(0).Width / 1.5, False, "", , "FROM_DOCUMENT_NO")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("จากรหัสใบสำคัญจ่าย"))
   
   '2 =============================
   Call LoadControl("T", txtGeneric(0).Width / 1.5, False, "", , "TO_DOCUMENT_NO")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("ถึงรหัสใบสำคัญจ่าย"))
   
   '2 =============================
   Call LoadControl("T", txtGeneric(0).Width / 1.5, False, "", , "FROM_DOCUMENT_NO1")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("หรือจากรหัสใบสำคัญจ่าย"))
   
   '2 =============================
   Call LoadControl("T", txtGeneric(0).Width / 1.5, False, "", , "TO_DOCUMENT_NO1")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("หรือถึงรหัสใบสำคัญจ่าย"))
   
   Call LoadControl("D", uctlGenericDate(0).Width, True, "", , "TO_PAY_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("ถึงวันที่ชำระหนี้"))
   
   '3 =============================
   Call LoadControl("CB", cboGeneric(0).Width, True, "", 1, "SUPPLIER_TYPE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("ประเภทผู้จำหน่าย"))
   
   '3 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 2, "ORDER_BY")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("เรียงตาม"))

   '4 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 3, "ORDER_TYPE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("เรียงจาก"))
      
   Call LoadControl("CH", chkGeneric(0).Width, True, "แสดงสรุป", 1, "SHOW_SUMMARY")
   
   Call ShowControl
   Call LoadComboData
End Sub
Private Sub InitReport1_0_5_3()
Dim C As CReportControl
Dim Top As Long
Dim Left As Long
Dim LabelWidth As Long
Dim Offset As Long

   Top = lblGeneric(0).Top
   Left = lblGeneric(0).Left
   LabelWidth = lblGeneric(0).Width
   Offset = 100
   
   '1 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, False, "", , "FROM_CHECK_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("จากวันที่ดิวเช็ค"))
   
   '1 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, True, "", , "TO_CHECK_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("ถึงวันที่ดิวเช็ค"))
      
   '1 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, True, "", , "PRINT_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("วันที่พิมพ์"))
   
   '1 =============================
'   Call LoadControl("T", txtGeneric(0).Width / 1.5, True, "", , "FROM_SUPPLIER_CODE")
'   Call LoadControl("L", lblGeneric(0).Width, True, MapText("จากรหัสผู้จำหน่าย"))
   
   '2 =============================
   Call LoadControl("T", txtGeneric(0).Width / 1.5, True, "", , "SUPPLIER_CODE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("รหัสผู้จำหน่าย"))

   '2 =============================
   Call LoadControl("T", txtGeneric(0).Width / 1.5, False, "", , "FROM_DOCUMENT_NO")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("จากรหัสใบสำคัญจ่าย"))
   
   '2 =============================
   Call LoadControl("T", txtGeneric(0).Width / 1.5, False, "", , "TO_DOCUMENT_NO")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("ถึงรหัสใบสำคัญจ่าย"))
   
   '2 =============================
   Call LoadControl("T", txtGeneric(0).Width / 1.5, False, "", , "FROM_DOCUMENT_NO1")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("หรือจากรหัสใบสำคัญจ่าย"))
   
   '2 =============================
   Call LoadControl("T", txtGeneric(0).Width / 1.5, False, "", , "TO_DOCUMENT_NO1")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("หรือถึงรหัสใบสำคัญจ่าย"))
   
   Call LoadControl("D", uctlGenericDate(0).Width, True, "", , "TO_PAY_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("ถึงวันที่ชำระหนี้"))
   
   '3 =============================
   Call LoadControl("CB", cboGeneric(0).Width, True, "", 1, "SUPPLIER_TYPE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("ประเภทผู้จำหน่าย"))

   '3 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 2, "ORDER_BY")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("เรียงตาม"))

   '4 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 3, "ORDER_TYPE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("เรียงจาก"))
   
   Call LoadControl("CH", chkGeneric(0).Width, True, "แสดงสรุป", 1, "SHOW_SUMMARY")
   
   Call ShowControl
   Call LoadComboData
End Sub
Private Sub InitReport1_0_5_4()
Dim C As CReportControl
Dim Top As Long
Dim Left As Long
Dim LabelWidth As Long
Dim Offset As Long

   Top = lblGeneric(0).Top
   Left = lblGeneric(0).Left
   LabelWidth = lblGeneric(0).Width
   Offset = 100
   
   '1 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, False, "", , "FROM_CHECK_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("จากวันที่ดิวเช็ค"))
   
   '1 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, True, "", , "TO_CHECK_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("ถึงวันที่ดิวเช็ค"))
      
   '1 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, True, "", , "PRINT_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("วันที่พิมพ์"))
   
   '1 =============================
'   Call LoadControl("T", txtGeneric(0).Width / 1.5, True, "", , "FROM_SUPPLIER_CODE")
'   Call LoadControl("L", lblGeneric(0).Width, True, MapText("จากรหัสผู้จำหน่าย"))
   
   '2 =============================
   Call LoadControl("T", txtGeneric(0).Width / 1.5, True, "", , "SUPPLIER_CODE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("รหัสผู้จำหน่าย"))

   '2 =============================
   Call LoadControl("T", txtGeneric(0).Width / 1.5, False, "", , "FROM_DOCUMENT_NO")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("จากรหัสใบสำคัญจ่าย"))
   
   '2 =============================
   Call LoadControl("T", txtGeneric(0).Width / 1.5, False, "", , "TO_DOCUMENT_NO")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("ถึงรหัสใบสำคัญจ่าย"))
   
   '2 =============================
   Call LoadControl("T", txtGeneric(0).Width / 1.5, False, "", , "FROM_DOCUMENT_NO1")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("หรือจากรหัสใบสำคัญจ่าย"))
   
   '2 =============================
   Call LoadControl("T", txtGeneric(0).Width / 1.5, False, "", , "TO_DOCUMENT_NO1")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("หรือถึงรหัสใบสำคัญจ่าย"))
   
   Call LoadControl("D", uctlGenericDate(0).Width, True, "", , "TO_PAY_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("ถึงวันที่ชำระหนี้"))
   
   '3 =============================
   Call LoadControl("CB", cboGeneric(0).Width, True, "", 1, "SUPPLIER_TYPE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("ประเภทผู้จำหน่าย"))
   
   '2 =============================
   Call LoadControl("T", txtGeneric(0).Width / 1.5, True, "", , "NOTE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("หมายเหตุ"))
   
   '2 =============================
   Call LoadControl("T", txtGeneric(0).Width / 1.5, True, "", , "ACCNUM")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("รหัสบัญชี"))
   
   '3 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 2, "ORDER_BY")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("เรียงตาม"))

   '4 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 3, "ORDER_TYPE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("เรียงจาก"))
   
   Call LoadControl("CH", chkGeneric(0).Width, True, "แสดงสรุป", 1, "SHOW_SUMMARY")
   
   Call ShowControl
   Call LoadComboData
End Sub
Private Sub InitReport1_0_15_1()
Dim C As CReportControl
Dim Top As Long
Dim Left As Long
Dim LabelWidth As Long
Dim Offset As Long

   Top = lblGeneric(0).Top
   Left = lblGeneric(0).Left
   LabelWidth = lblGeneric(0).Width
   Offset = 100
   
   '1 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, False, "", , "FROM_CHECK_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("จากวันที่มีผลเช็ค"))
   
   '1 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, True, "", , "TO_CHECK_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("ถึงวันที่มีผลเช็ค"))
      
   '2 =============================
   Call LoadControl("T", txtGeneric(0).Width / 1.5, False, "", , "FROM_DOCUMENT_NO")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("จากรหัสใบสำคัญรับ"))
   
   '2 =============================
   Call LoadControl("T", txtGeneric(0).Width / 1.5, False, "", , "TO_DOCUMENT_NO")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("ถึงรหัสใบสำคัญรับ"))
   
   '2 =============================
   Call LoadControl("T", txtGeneric(0).Width / 1.5, False, "", , "FROM_DOCUMENT_NO1")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("หรือจากรหัสใบสำคัญรับ"))
   
   '2 =============================
   Call LoadControl("T", txtGeneric(0).Width / 1.5, False, "", , "TO_DOCUMENT_NO1")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("หรือถึงรหัสใบสำคัญรับ"))
   
   '3 =============================
   Call LoadControl("CB", cboGeneric(0).Width, True, "", 1, "CUSTOMER_TYPE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("ประเภทลูกค้า"))

   Call LoadControl("T", txtGeneric(0).Width / 1.5, True, "", , "FROM_CUSTOMER_CODE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("จากลูกค้า"))
   
   Call LoadControl("T", txtGeneric(0).Width / 1.5, True, "", , "TO_CUSTOMER_CODE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("ถึงลูกค้า"))
   
   Call LoadControl("T", txtGeneric(0).Width / 1.5, True, "", , "FROM_SALE_CODE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("จากพนักงานขาย"))
   
   Call LoadControl("T", txtGeneric(0).Width / 1.5, True, "", , "TO_SALE_CODE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("ถึงพนักงานขาย"))
   
   '3 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 2, "ORDER_BY")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("เรียงตาม"))

   '4 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 3, "ORDER_TYPE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("เรียงจาก"))
   
   Call LoadControl("CH", chkGeneric(0).Width, True, "รวมตามวันที่ดิวเช็ค", 1, "GROUP_BY_DUE_DATE")
   Call LoadControl("CH", chkGeneric(0).Width, True, "รวมตามลูกค้า", 1, "GROUP_BY_CUSTOMER")
   Call LoadControl("CH", chkGeneric(0).Width, True, "รวมตามพนักงานขาย", 1, "GROUP_BY_SALE")
   
  Call LoadControl("CH", chkGeneric(0).Width, True, "รวมตามพนักงานขายและลูกค้า", 1, "GROUP_BY_SALE_AND_CUSTOMER")
   
   Call LoadControl("CH", chkGeneric(0).Width, True, "แสดงสรุป", 1, "SHOW_SUMMARY")
   
   Call ShowControl
   Call LoadComboData
End Sub
Private Sub InitReport6_0_1()
Dim C As CReportControl
Dim Top As Long
Dim Left As Long
Dim LabelWidth As Long
Dim Offset As Long
   
   Top = lblGeneric(0).Top
   Left = lblGeneric(0).Left
   LabelWidth = lblGeneric(0).Width
   Offset = 100
   
   '1 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, True, "", , "FROM_DOC_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("จากวันที่ส่งสินค้า"))

   
   '1 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, True, "", , "TO_DOC_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("ถึงวันที่ส่งสินค้า"))

      Call LoadControl("CH", chkGeneric(0).Width, True, "สรุป", 2, "SUMMARY_MODE")
      Call LoadControl("CH", chkGeneric(0).Width, True, "บันทึกลงฐานข้อมูล", 2, "SAVE_MODE")
      
   Call ShowControl
   Call LoadComboData
End Sub


Private Sub InitReport6_0_2()
Dim C As CReportControl
Dim Top As Long
Dim Left As Long
Dim LabelWidth As Long
Dim Offset As Long
   
   Top = lblGeneric(0).Top
   Left = lblGeneric(0).Left
   LabelWidth = lblGeneric(0).Width
   Offset = 100
   
   '1 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, True, "", , "FROM_DOC_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("จากวันที่ส่งสินค้า"))
   
   '1 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, True, "", , "TO_DOC_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("ถึงวันที่ส่งสินค้า"))
   
   '1 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, False, "", , "FROM_CMPL_DATE")
   Call LoadControl("L", lblGeneric(0).Width, False, MapText("จากวันที่รับชำระ"))
   
   '1 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, False, "", , "TO_CMPL_DATE")
   Call LoadControl("L", lblGeneric(0).Width, False, MapText("ถึงวันที่รับชำระ"))
   
   Call LoadControl("T", txtGeneric(0).Width / 1.5, True, "", , "FROM_CUSTOMER_CODE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("จากลูกค้า"))
   
   Call LoadControl("T", txtGeneric(0).Width / 1.5, True, "", , "TO_CUSTOMER_CODE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("ถึงลูกค้า"))
   
      '1 =============================
'   Call LoadControl("T", txtGeneric(0).Width / 1.5, True, "", , "FROM_SALE_CODE")
      Call LoadControl("CB", txtGeneric(0).Width, True, "", 2, "FROM_SALE_CODE", "FROM_SALE_NAME")
      Call LoadControl("L", lblGeneric(0).Width, True, MapText("จากรหัสพนักงานขาย"))

'   '2 =============================
'   Call LoadControl("T", txtGeneric(0).Width / 1.5, True, "", , "TO_SALE_CODE")
      Call LoadControl("CB", txtGeneric(0).Width, True, "", 2, "TO_SALE_CODE", "TO_SALE_NAME")
      Call LoadControl("L", lblGeneric(0).Width, True, MapText("จากรหัสพนักงานขาย"))
   
   Call LoadControl("CB", cboGeneric(0).Width, True, "", 1, "AREA_TYPE_ID", "AREA_TYPE_NAME")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("เขตการขาย"))
         
      Call LoadControl("CH", chkGeneric(0).Width, True, "แสดง Invoice ทั้งหมด", 2, "SHOWZERO_MODE")
      Call LoadControl("CH", chkGeneric(0).Width, True, "สรุป", 2, "SUMMARY_MODE")
      Call LoadControl("CH", chkGeneric(0).Width, True, "แสดงเฉพาะติดลบ", 2, "ONLY_MINUS")
      Call LoadControl("CH", chkGeneric(0).Width, True, "บันทึกลงฐานข้อมูล", 2, "SAVE_MODE")
   
   Call ShowControl
   Call LoadComboData
End Sub

Private Sub InitReport6_0_3()
Dim C As CReportControl
Dim Top As Long
Dim Left As Long
Dim LabelWidth As Long
Dim Offset As Long
   
   Top = lblGeneric(0).Top
   Left = lblGeneric(0).Left
   LabelWidth = lblGeneric(0).Width
   Offset = 100
   
   '1 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, True, "", , "FROM_DOC_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("จากวันที่ส่งสินค้า"))
   
   '1 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, True, "", , "TO_DOC_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("ถึงวันที่ส่งสินค้า"))
   
   '1 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, False, "", , "FROM_CMPL_DATE")
   Call LoadControl("L", lblGeneric(0).Width, False, MapText("จากวันที่รับชำระ"))
   
   '1 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, False, "", , "TO_CMPL_DATE")
   Call LoadControl("L", lblGeneric(0).Width, False, MapText("ถึงวันที่รับชำระ"))
   
      '1 =============================
'   Call LoadControl("T", txtGeneric(0).Width / 1.5, True, "", , "FROM_SALE_CODE")

   Call LoadControl("T", txtGeneric(0).Width / 1.5, True, "", , "FROM_CUSTOMER_CODE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("จากลูกค้า"))
   
   Call LoadControl("T", txtGeneric(0).Width / 1.5, True, "", , "TO_CUSTOMER_CODE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("ถึงลูกค้า"))
   
      Call LoadControl("CB", txtGeneric(0).Width, True, "", 2, "FROM_SALE_CODE", "FROM_SALE_NAME")
      Call LoadControl("L", lblGeneric(0).Width, True, MapText("จากรหัสพนักงานขาย"))
'   '2 =============================
'   Call LoadControl("T", txtGeneric(0).Width / 1.5, True, "", , "TO_SALE_CODE")
      Call LoadControl("CB", txtGeneric(0).Width, True, "", 2, "TO_SALE_CODE", "TO_SALE_NAME")
      Call LoadControl("L", lblGeneric(0).Width, True, MapText("จากรหัสพนักงานขาย"))
   
   Call LoadControl("CB", cboGeneric(0).Width, True, "", 1, "AREA_TYPE_ID", "AREA_TYPE_NAME")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("เขตการขาย"))
   
      '3 =============================
      
      Call LoadControl("CH", chkGeneric(0).Width, True, "แสดง Invoice ทั้งหมด", 2, "SHOWZERO_MODE")
      Call LoadControl("CH", chkGeneric(0).Width, True, "สรุป", 2, "SUMMARY_MODE")
      Call LoadControl("CH", chkGeneric(0).Width, True, "บันทึกลงฐานข้อมูล", 2, "SAVE_MODE")
   
   Call ShowControl
   Call LoadComboData
End Sub

Private Sub InitReport6_0_4()
Dim C As CReportControl
Dim Top As Long
Dim Left As Long
Dim LabelWidth As Long
Dim Offset As Long
   
   Top = lblGeneric(0).Top
   Left = lblGeneric(0).Left
   LabelWidth = lblGeneric(0).Width
   Offset = 100

   
   '1 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, True, "", , "FROM_CMPL_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("ประจำวันที่"))
   
   '1 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, True, "", , "TO_CMPL_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("ถึงวันที่"))
   
      Call LoadControl("CH", chkGeneric(0).Width, True, "สรุป", 2, "SUMMARY_MODE")
   
   Call ShowControl
End Sub

Private Sub InitReport6_0_5()           ' คอมขาย
Dim C As CReportControl
Dim Top As Long
Dim Left As Long
Dim LabelWidth As Long
Dim Offset As Long
   
   Top = lblGeneric(0).Top
   Left = lblGeneric(0).Left
   LabelWidth = lblGeneric(0).Width
   Offset = 100
   
   '1 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, False, "", , "FROM_DOC_DATE")  'บังคับใส่ค่า
   Call LoadControl("L", lblGeneric(0).Width, False, MapText("จากวันที่ส่งสินค้า"))  'ใส่สี
   
   
   '1 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, False, "", , "TO_DOC_DATE")
   Call LoadControl("L", lblGeneric(0).Width, False, MapText("ถึงวันที่ส่งสินค้า"))
   
'   '1 =============================
'   Call LoadControl("D", uctlGenericDate(0).Width, True, "", , "FROM_CMPL_DATE")
'   Call LoadControl("L", lblGeneric(0).Width, True, MapText("จากวันที่รับชำระ"))
'
'   '1 =============================
'   Call LoadControl("D", uctlGenericDate(0).Width, True, "", , "TO_CMPL_DATE")
'   Call LoadControl("L", lblGeneric(0).Width, True, MapText("ถึงวันที่รับชำระ"))
   
   Call LoadControl("T", txtGeneric(0).Width / 1.5, True, "", , "FROM_CUSTOMER_CODE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("จากลูกค้า"))
   
   Call LoadControl("T", txtGeneric(0).Width / 1.5, True, "", , "TO_CUSTOMER_CODE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("ถึงลูกค้า"))
   
      '1 =============================
      Call LoadControl("CB", txtGeneric(0).Width, True, "", 2, "FROM_SALE_CODE", "FROM_SALE_NAME")
      Call LoadControl("L", lblGeneric(0).Width, True, MapText("จากรหัสพนักงานขาย"))

'   '2 =============================
      Call LoadControl("CB", txtGeneric(0).Width, True, "", 2, "TO_SALE_CODE", "TO_SALE_NAME")
      Call LoadControl("L", lblGeneric(0).Width, True, MapText("จากรหัสพนักงานขาย"))

      '3 =============================
   Call LoadControl("CB", cboGeneric(0).Width, True, "", 1, "AREA_TYPE_ID", "AREA_TYPE_NAME")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("เขตการขาย"))

      Call LoadControl("CH", chkGeneric(0).Width, True, "สรุป", 2, "SUMMARY_MODE")
   
   Call ShowControl
   Call LoadComboData
End Sub

Private Sub InitReport6_0_6()           ' สินค้าไม่คิด Commercial #1
Dim C As CReportControl
Dim Top As Long
Dim Left As Long
Dim LabelWidth As Long
Dim Offset As Long
   
   Top = lblGeneric(0).Top
   Left = lblGeneric(0).Left
   LabelWidth = lblGeneric(0).Width
   Offset = 100
   
   '1 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, False, "", , "FROM_DOC_DATE")
   Call LoadControl("L", lblGeneric(0).Width, False, MapText("จากวันที่ส่งสินค้า"))
   
   '1 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, False, "", , "TO_DOC_DATE")
   Call LoadControl("L", lblGeneric(0).Width, False, MapText("ถึงวันที่ส่งสินค้า"))
   
   '1 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, True, "", , "FROM_CMPL_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("จากวันที่รับชำระ"))
   
   '1 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, True, "", , "TO_CMPL_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("ถึงวันที่รับชำระ"))
   
      '1 =============================
      Call LoadControl("CB", txtGeneric(0).Width, True, "", 2, "FROM_SALE_CODE", "FROM_SALE_NAME")
      Call LoadControl("L", lblGeneric(0).Width, True, MapText("จากรหัสพนักงานขาย"))

'   '2 =============================
      Call LoadControl("CB", txtGeneric(0).Width, True, "", 2, "TO_SALE_CODE", "TO_SALE_NAME")
      Call LoadControl("L", lblGeneric(0).Width, True, MapText("จากรหัสพนักงานขาย"))
  
   Call LoadControl("CB", cboGeneric(0).Width, True, "", 1, "AREA_TYPE_ID", "AREA_TYPE_NAME")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("เขตการขาย"))

   Call LoadControl("CH", chkGeneric(0).Width, True, "สรุป", 2, "SUMMARY_MODE")
   
   Call ShowControl
   Call LoadComboData
End Sub

Private Sub InitReport6_0_7()
Dim C As CReportControl
Dim Top As Long
Dim Left As Long
Dim LabelWidth As Long
Dim Offset As Long

   Top = lblGeneric(0).Top
   Left = lblGeneric(0).Left
   LabelWidth = lblGeneric(0).Width
   Offset = 100
   
   Call LoadControl("C", cboGeneric(0).Width \ 2, False, "", 1, "MONTH_ID", "MONTH_NAME")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("เดือน"))

   '2 =============================
   Call LoadControl("T", txtGeneric(0).Width / 2, False, "", , "YEAR_NO")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("ปี"))
   
   '1 =============================
   Call LoadControl("T", txtGeneric(0).Width / 1.5, True, "", , "FROM_STK_CODE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("จากรหัสสินค้า"))

   '2 =============================
   Call LoadControl("T", txtGeneric(0).Width / 1.5, True, "", , "TO_STK_CODE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("ถึงรหัสสินค้า"))
   
   Call LoadControl("CB", txtGeneric(0).Width, True, "", 2, "FROM_SALE_CODE", "FROM_SALE_NAME")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("จากพนักงานขาย"))

'   '2 =============================
   Call LoadControl("CB", txtGeneric(0).Width, True, "", 2, "TO_SALE_CODE", "TO_SALE_NAME")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("จากพนักงานขาย"))
   
   Call LoadControl("CH", chkGeneric(0).Width, True, "สรุป", 2, "SUMMARY_MODE")
   
   Call ShowControl
   Call LoadComboData
End Sub

Private Sub InitReport6_0_8()
Dim C As CReportControl
Dim Top As Long
Dim Left As Long
Dim LabelWidth As Long
Dim Offset As Long

   Top = lblGeneric(0).Top
   Left = lblGeneric(0).Left
   LabelWidth = lblGeneric(0).Width
   Offset = 100
   
   Call LoadControl("C", cboGeneric(0).Width \ 2, False, "", 1, "MONTH_ID", "MONTH_NAME")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("เดือน"))

   '2 =============================
   Call LoadControl("T", txtGeneric(0).Width / 2, False, "", , "YEAR_NO")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("ปี"))
   
   '2 =============================
   Call LoadControl("T", txtGeneric(0).Width / 2, False, "", , "RUN_MONTH")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("เดือนเคลื่อนไหว"))
   
   '1 =============================
   Call LoadControl("T", txtGeneric(0).Width / 1.5, True, "", , "FROM_CUSTOMER_CODE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("จากรหัสลูกค้า"))

   '2 =============================
   Call LoadControl("T", txtGeneric(0).Width / 1.5, True, "", , "TO_CUSTOMER_CODE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("ถึงรหัสลูกค้า"))
   
   Call LoadControl("T", txtGeneric(0).Width / 1.5, True, "", , "FROM_STK_CODE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("จากรหัสสินค้า"))

   Call LoadControl("T", txtGeneric(0).Width / 1.5, True, "", , "TO_STK_CODE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("ถึงรหัสสินค้า"))
   
   Call LoadControl("CH", chkGeneric(0).Width, True, "สรุป", 1, "SUMMARY_MODE")
   Call LoadControl("CH", chkGeneric(0).Width, True, "เฉพาะลูกค้าใหม่", 1, "NEWCUS_MODE")
   
   Call ShowControl
   Call LoadComboData
End Sub
Private Sub InitReport6_0_9()
Dim C As CReportControl
Dim Top As Long
Dim Left As Long
Dim LabelWidth As Long
Dim Offset As Long
   
   Top = lblGeneric(0).Top
   Left = lblGeneric(0).Left
   LabelWidth = lblGeneric(0).Width
   Offset = 100
   
   '1 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, True, "", , "FROM_DOC_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("จากวันที่ส่งสินค้า"))
   
   '1 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, True, "", , "TO_DOC_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("ถึงวันที่ส่งสินค้า"))

      '1 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, True, "", , "FROM_CMPL_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("จากวันที่รับชำระ"))
   
   '1 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, True, "", , "TO_CMPL_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("ถึงวันที่รับชำระ"))
   
      '1 =============================
      Call LoadControl("CB", txtGeneric(0).Width, True, "", 2, "FROM_SALE_CODE", "FROM_SALE_NAME")
      Call LoadControl("L", lblGeneric(0).Width, True, MapText("จากรหัสพนักงานขาย"))

'   '2 =============================
      Call LoadControl("CB", txtGeneric(0).Width, True, "", 2, "TO_SALE_CODE", "TO_SALE_NAME")
      Call LoadControl("L", lblGeneric(0).Width, True, MapText("จากรหัสพนักงานขาย"))

      '3 =============================
      Call LoadControl("CH", chkGeneric(0).Width, True, "สรุป", 2, "SUMMARY_MODE")   '
   
   Call ShowControl
   Call LoadComboData
End Sub

Private Sub InitReport6_0_10()
Dim C As CReportControl
Dim Top As Long
Dim Left As Long
Dim LabelWidth As Long
Dim Offset As Long
   
   Top = lblGeneric(0).Top
   Left = lblGeneric(0).Left
   LabelWidth = lblGeneric(0).Width
   Offset = 100
   
   '1 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, True, "", , "FROM_DOC_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("จากวันที่ส่งสินค้า"))
   
   Call LoadControl("D", uctlGenericDate(0).Width, True, "", , "TO_DOC_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("ถึงวันที่ส่งสินค้า"))
   
 '1 =============================
'   Call LoadControl("T", txtGeneric(0).Width / 1.5, True, "", , "FROM_GOODS_CODE")
'   Call LoadControl("L", lblGeneric(0).Width, True, MapText("จากรหัสสินค้า"))
'
'   Call LoadControl("T", txtGeneric(0).Width / 1.5, True, "", , "TO_GOODS_CODE")
'   Call LoadControl("L", lblGeneric(0).Width, True, MapText("ถึงรหัสสินค้า"))

   '1 =============================
   Call LoadControl("T", txtGeneric(0).Width / 1.5, True, "", , "FROM_CUSTOMER_CODE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("จากลูกค้า"))
   
   Call LoadControl("T", txtGeneric(0).Width / 1.5, True, "", , "TO_CUSTOMER_CODE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("ถึงลูกค้า"))
   
   '2 =============================
'   Call LoadControl("CB", txtGeneric(0).Width, True, "", 2, "FROM_SALE_CODE", "FROM_SALE_NAME")
'   Call LoadControl("L", lblGeneric(0).Width, True, MapText("จากรหัสพนักงานขาย"))
'
'   Call LoadControl("CB", txtGeneric(0).Width, True, "", 2, "TO_SALE_CODE", "TO_SALE_NAME")
'   Call LoadControl("L", lblGeneric(0).Width, True, MapText("จากรหัสพนักงานขาย"))
   
   '3 =============================
'  Call LoadControl("CH", chkGeneric(0).Width, True, "สรุป", 2, "SUMMARY_MODE")
  Call LoadControl("CH", chkGeneric(0).Width, True, "เฉพาะส่วนลดที่ยังไม่กำหนด SR", 2, "SHOWNONAME_MODE")
   
   Call ShowControl
   Call LoadComboData
End Sub
Private Sub InitReport6_0_11()
Dim C As CReportControl
Dim Top As Long
Dim Left As Long
Dim LabelWidth As Long
Dim Offset As Long
   
   Top = lblGeneric(0).Top
   Left = lblGeneric(0).Left
   LabelWidth = lblGeneric(0).Width
   Offset = 100
   
   '1 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, True, "", , "FROM_DOC_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("จากวันที่ส่งสินค้า"))
   
   Call LoadControl("D", uctlGenericDate(0).Width, True, "", , "TO_DOC_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("ถึงวันที่ส่งสินค้า"))
   
 '1 =============================
   Call LoadControl("T", txtGeneric(0).Width / 1.5, True, "", , "FROM_GOODS_CODE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("จากรหัสสินค้า"))

   Call LoadControl("T", txtGeneric(0).Width / 1.5, True, "", , "TO_GOODS_CODE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("ถึงรหัสสินค้า"))

   '1 =============================
   Call LoadControl("T", txtGeneric(0).Width / 1.5, True, "", , "FROM_CUSTOMER_CODE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("จากลูกค้า"))

   Call LoadControl("T", txtGeneric(0).Width / 1.5, True, "", , "TO_CUSTOMER_CODE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("ถึงลูกค้า"))
   
   '2 =============================
'   Call LoadControl("CB", txtGeneric(0).Width, True, "", 2, "FROM_SALE_CODE", "FROM_SALE_NAME")
'   Call LoadControl("L", lblGeneric(0).Width, True, MapText("จากรหัสพนักงานขาย"))
'
'   Call LoadControl("CB", txtGeneric(0).Width, True, "", 2, "TO_SALE_CODE", "TO_SALE_NAME")
'   Call LoadControl("L", lblGeneric(0).Width, True, MapText("จากรหัสพนักงานขาย"))
   
   '3 =============================
'  Call LoadControl("CH", chkGeneric(0).Width, True, "สรุป", 2, "SUMMARY_MODE")
'  Call LoadControl("CH", chkGeneric(0).Width, True, "เฉพาะส่วนลดที่ยังไม่กำหนด SR", 2, "SHOWNONAME_MODE")
   
   Call LoadControl("CB", cboGeneric(0).Width, True, "", 1, "AREA_TYPE_ID", "AREA_TYPE_NAME")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("เขตการขาย"))

   Call LoadControl("CH", chkGeneric(0).Width, True, "สรุป", 2, "SUMMARY_MODE")
   Call LoadControl("CH", chkGeneric(0).Width, True, "แสดงรวมลูกค้า", 2, "SHOWCUS_MODE")
   
   Call ShowControl
   Call LoadComboData
End Sub

Private Sub InitReport6_0_12()
Dim C As CReportControl
Dim Top As Long
Dim Left As Long
Dim LabelWidth As Long
Dim Offset As Long
   
   Top = lblGeneric(0).Top
   Left = lblGeneric(0).Left
   LabelWidth = lblGeneric(0).Width
   Offset = 100
   
   '1 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, True, "", , "FROM_DOC_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("จากวันที่ส่งสินค้า"))
   
   Call LoadControl("D", uctlGenericDate(0).Width, True, "", , "TO_DOC_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("ถึงวันที่ส่งสินค้า"))
   
 '1 =============================
   Call LoadControl("T", txtGeneric(0).Width / 1.5, True, "", , "FROM_GOODS_CODE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("จากรหัสสินค้า"))

   Call LoadControl("T", txtGeneric(0).Width / 1.5, True, "", , "TO_GOODS_CODE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("ถึงรหัสสินค้า"))

   '1 =============================
   Call LoadControl("T", txtGeneric(0).Width / 1.5, True, "", , "FROM_CUSTOMER_CODE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("จากลูกค้า"))

   Call LoadControl("T", txtGeneric(0).Width / 1.5, True, "", , "TO_CUSTOMER_CODE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("ถึงลูกค้า"))
   
   '2 =============================
'   Call LoadControl("CB", txtGeneric(0).Width, True, "", 2, "FROM_SALE_CODE", "FROM_SALE_NAME")
'   Call LoadControl("L", lblGeneric(0).Width, True, MapText("จากรหัสพนักงานขาย"))
'
'   Call LoadControl("CB", txtGeneric(0).Width, True, "", 2, "TO_SALE_CODE", "TO_SALE_NAME")
'   Call LoadControl("L", lblGeneric(0).Width, True, MapText("จากรหัสพนักงานขาย"))
   
   '3 =============================
'  Call LoadControl("CH", chkGeneric(0).Width, True, "สรุป", 2, "SUMMARY_MODE")
'  Call LoadControl("CH", chkGeneric(0).Width, True, "เฉพาะส่วนลดที่ยังไม่กำหนด SR", 2, "SHOWNONAME_MODE")
   
   Call LoadControl("CB", cboGeneric(0).Width, True, "", 1, "AREA_TYPE_ID", "AREA_TYPE_NAME")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("เขตการขาย"))

   Call LoadControl("CH", chkGeneric(0).Width, True, "สรุป", 2, "SUMMARY_MODE")
   
   Call ShowControl
   Call LoadComboData
End Sub

Private Sub InitReport6_0_13()           ' สินค้าไม่คิด Commercial #1
Dim C As CReportControl
Dim Top As Long
Dim Left As Long
Dim LabelWidth As Long
Dim Offset As Long
   
   Top = lblGeneric(0).Top
   Left = lblGeneric(0).Left
   LabelWidth = lblGeneric(0).Width
   Offset = 100
   
   '1 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, False, "", , "FROM_DOC_DATE")
   Call LoadControl("L", lblGeneric(0).Width, False, MapText("จากวันที่ส่งสินค้า"))
   
   '1 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, False, "", , "TO_DOC_DATE")
   Call LoadControl("L", lblGeneric(0).Width, False, MapText("ถึงวันที่ส่งสินค้า"))
      
    '1 =============================
   Call LoadControl("T", txtGeneric(0).Width / 1.5, True, "", , "FROM_GOODS_CODE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("จากรหัสสินค้า"))

   Call LoadControl("T", txtGeneric(0).Width / 1.5, True, "", , "TO_GOODS_CODE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("ถึงรหัสสินค้า"))

   '1 =============================
   Call LoadControl("T", txtGeneric(0).Width / 1.5, True, "", , "FROM_CUSTOMER_CODE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("จากลูกค้า"))

   Call LoadControl("T", txtGeneric(0).Width / 1.5, True, "", , "TO_CUSTOMER_CODE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("ถึงลูกค้า"))
      
      '1 =============================
      Call LoadControl("CB", txtGeneric(0).Width, True, "", 2, "FROM_SALE_CODE", "FROM_SALE_NAME")
      Call LoadControl("L", lblGeneric(0).Width, True, MapText("จากรหัสพนักงานขาย"))

'   '2 =============================
      Call LoadControl("CB", txtGeneric(0).Width, True, "", 2, "TO_SALE_CODE", "TO_SALE_NAME")
      Call LoadControl("L", lblGeneric(0).Width, True, MapText("จากรหัสพนักงานขาย"))
  
   Call LoadControl("CB", cboGeneric(0).Width, True, "", 1, "AREA_TYPE_ID", "AREA_TYPE_NAME")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("เขตการขาย"))

   Call LoadControl("CH", chkGeneric(0).Width, True, "สรุป", 2, "SUMMARY_MODE")
   
   Call ShowControl
   Call LoadComboData
End Sub

Private Sub InitReport6_0_14()           ' คอมขาย
Dim C As CReportControl
Dim Top As Long
Dim Left As Long
Dim LabelWidth As Long
Dim Offset As Long
   
   Top = lblGeneric(0).Top
   Left = lblGeneric(0).Left
   LabelWidth = lblGeneric(0).Width
   Offset = 100
   
   '1 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, False, "", , "FROM_DOC_DATE")  'บังคับใส่ค่า
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("จากวันที่จ่าย"))
   
   '1 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, False, "", , "TO_DOC_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("ถึงวันที่จ่าย"))
   
      '1 =============================
      Call LoadControl("CB", txtGeneric(0).Width, True, "", 3, "FROM_SALE_CODE", "FROM_SALE_NAME")
      Call LoadControl("L", lblGeneric(0).Width, True, MapText("จากรหัสพนักงานขาย"))

'   '2 =============================
      Call LoadControl("CB", txtGeneric(0).Width, True, "", 3, "TO_SALE_CODE", "TO_SALE_NAME")
      Call LoadControl("L", lblGeneric(0).Width, True, MapText("จากรหัสพนักงานขาย"))

   '3 =============================
'   Call LoadControl("CB", cboGeneric(0).Width, True, "", 1, "AREA_TYPE_ID", "AREA_TYPE_NAME")
'   Call LoadControl("L", lblGeneric(0).Width, True, MapText("เขตการขาย"))

   '4 =============================
   Call LoadControl("CH", chkGeneric(0).Width, True, "commission ขาย", 2, "COM1_MODE")
   Call LoadControl("CH", chkGeneric(0).Width, True, "commission เก็บเงิน", 2, "COM2_MODE")
   Call LoadControl("CH", chkGeneric(0).Width, True, "incentive", 2, "INCEN_MODE")
   
   Call ShowControl
   Call LoadComboData
End Sub
Private Sub InitReport1_0_17()
Dim C As CReportControl
Dim Top As Long
Dim Left As Long
Dim LabelWidth As Long
Dim Offset As Long

   Top = lblGeneric(0).Top
   Left = lblGeneric(0).Left
   LabelWidth = lblGeneric(0).Width
   Offset = 100
   
   '1 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, True, "", , "FROM_PAY_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("จากวันที่ชำระหนี้"))
   
   '1 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, True, "", , "TO_PAY_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("ถึงวันที่ชำระหนี้"))
   
   '1 =============================
   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "FROM_CUSTOMER_CODE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("รหัสลูกหนี้"))
   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "TO_CUSTOMER_CODE", , True)

   '1 =============================
   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "FROM_SALE_CODE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("รหัสพนักงานขาย"))
   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "TO_SALE_CODE", , True)
   
   Call ShowControl
   Call LoadComboData
End Sub

Private Sub InitReport1_0_18()
Dim C As CReportControl
Dim Top As Long
Dim Left As Long
Dim LabelWidth As Long
Dim Offset As Long

   Top = lblGeneric(0).Top
   Left = lblGeneric(0).Left
   LabelWidth = lblGeneric(0).Width
   Offset = 100
   
   '1 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, True, "", , "FROM_PAY_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("จากวันที่"))
   
   '1 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, True, "", , "TO_PAY_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("ถึงวันที่"))
   
   '1 =============================
   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "FROM_SALE_CODE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("รหัสพนักงานขาย"))
   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "TO_SALE_CODE", , True)
   
   '1 =============================
   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "FROM_CUSTOMER_CODE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("รหัสลูกค้า"))
   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "TO_CUSTOMER_CODE", , True)

   '1 =============================
   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "FROM_STOCK_CODE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("รหัสสินค้า"))
   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "TO_STOCK_CODE", , True)
'
   Call ShowControl
   Call LoadComboData
End Sub

Private Sub InitReport1_0_19()
Dim C As CReportControl
Dim Top As Long
Dim Left As Long
Dim LabelWidth As Long
Dim Offset As Long
   
   Top = lblGeneric(0).Top
   Left = lblGeneric(0).Left
   LabelWidth = lblGeneric(0).Width
   Offset = 100
   
   Call LoadControl("C", cboGeneric(0).Width \ 2, False, "", 1, "FROM_MONTH_ID", "FROM_MONTH_NAME")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("จากเดือน"))

   '2 =============================
   Call LoadControl("T", txtGeneric(0).Width / 2, False, "", , "FROM_YEAR_NO")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("จากปี"))
   
   Call LoadControl("C", cboGeneric(0).Width \ 2, False, "", 2, "TO_MONTH_ID", "TO_MONTH_NAME")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("ถึงเดือน"))

   '2 =============================
   Call LoadControl("T", txtGeneric(0).Width / 2, False, "", , "TO_YEAR_NO")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("ถึงปี"))
   
   '1 =============================
   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "FROM_SALE_CODE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("รหัสพนักงานขาย"))
   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "TO_SALE_CODE", , True)
   
   '1 =============================
   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "FROM_CUSTOMER_CODE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("รหัสลูกค้า"))
   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "TO_CUSTOMER_CODE", , True)
   
   '1 =============================
   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "FROM_STOCK_CODE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("รหัสสินค้า"))
   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "TO_STOCK_CODE", , True)

   Call LoadControl("CH", chkGeneric(0).Width, True, "สรุป", 1, "SUMMARY_FLAG")
   
'   Call LoadControl("CH", chkGeneric(0).Width, True, "แสดงเฉพาะจำนวน", 1, "ONLY_AMOUNT")
'   Call LoadControl("CH", chkGeneric(0).Width, True, "แสดงเฉพาะมูลค่า", 1, "ONLY_PRICE")
   
   Call ShowControl
   Call LoadComboData
End Sub

Private Sub InitReport1_0_20()
Dim C As CReportControl
Dim Top As Long
Dim Left As Long
Dim LabelWidth As Long
Dim Offset As Long
   
   Top = lblGeneric(0).Top
   Left = lblGeneric(0).Left
   LabelWidth = lblGeneric(0).Width
   Offset = 100
   
   Call LoadControl("C", cboGeneric(0).Width \ 2, False, "", 1, "FROM_MONTH_ID", "FROM_MONTH_NAME")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("จากเดือน"))

   '2 =============================
   Call LoadControl("T", txtGeneric(0).Width / 2, False, "", , "FROM_YEAR_NO")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("จากปี"))
   
   Call LoadControl("C", cboGeneric(0).Width \ 2, False, "", 2, "TO_MONTH_ID", "TO_MONTH_NAME")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("ถึงเดือน"))

   '2 =============================
   Call LoadControl("T", txtGeneric(0).Width / 2, False, "", , "TO_YEAR_NO")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("ถึงปี"))
   
   '1 =============================
   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "FROM_STOCK_CODE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("รหัสสินค้า"))
   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "TO_STOCK_CODE", , True)
   
   '1 =============================
   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "FROM_CUSTOMER_CODE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("รหัสลูกค้า"))
   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "TO_CUSTOMER_CODE", , True)

   Call LoadControl("CH", chkGeneric(0).Width, True, "สรุป", 1, "SUMMARY_FLAG")
   
   Call ShowControl
   Call LoadComboData
End Sub

Private Sub InitReport1_0_21()
Dim C As CReportControl
Dim Top As Long
Dim Left As Long
Dim LabelWidth As Long
Dim Offset As Long
   
   Top = lblGeneric(0).Top
   Left = lblGeneric(0).Left
   LabelWidth = lblGeneric(0).Width
   Offset = 100
   
   Call LoadControl("C", cboGeneric(0).Width \ 2, False, "", 1, "FROM_MONTH_ID", "FROM_MONTH_NAME")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("จากเดือน"))

   '2 =============================
   Call LoadControl("T", txtGeneric(0).Width / 2, False, "", , "FROM_YEAR_NO")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("จากปี"))
   
   Call LoadControl("C", cboGeneric(0).Width \ 2, False, "", 2, "TO_MONTH_ID", "TO_MONTH_NAME")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("ถึงเดือน"))

   '2 =============================
   Call LoadControl("T", txtGeneric(0).Width / 2, False, "", , "TO_YEAR_NO")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("ถึงปี"))
   
   
   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "FROM_CUSTOMER_CODE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("รหัสลูกหนี้"))
   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "TO_CUSTOMER_CODE", , True)
    
   '1 =============================
   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "FROM_STOCK_CODE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("รหัสสินค้า"))
   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "TO_STOCK_CODE", , True)
   
   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "FROM_SALE_CODE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("พนักงานขาย"))
   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "TO_SALE_CODE", , True)
   
   Call LoadControl("CH", chkGeneric(0).Width, True, "แยกตามลูกค้า", 1, "SPLIT_SALE")
   Call LoadControl("CH", chkGeneric(0).Width, True, "แสดงข้อมูลในตาราง", 1, "SHOW_INFOMATION")
   
   '1 =============================
   
   Call ShowControl
   Call LoadComboData
End Sub
