VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#2.0#0"; "THREED20.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmExportSupItem 
   BackColor       =   &H80000000&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   3480
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10605
   Icon            =   "frmExportSupItem.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3480
   ScaleWidth      =   10605
   StartUpPosition =   1  'CenterOwner
   Begin Threed.SSFrame SSFrame1 
      Height          =   3525
      Left            =   0
      TabIndex        =   5
      Top             =   0
      Width           =   10665
      _ExtentX        =   18812
      _ExtentY        =   6218
      _Version        =   131073
      PictureBackgroundStyle=   2
      Begin VB.ComboBox cboExportType 
         Height          =   315
         Left            =   6840
         Style           =   2  'Dropdown List
         TabIndex        =   14
         Top             =   1020
         Width           =   3615
      End
      Begin prjLedgerReport.uctlDate uctlFromDate 
         Height          =   405
         Left            =   1860
         TabIndex        =   12
         Top             =   1020
         Width           =   3855
         _ExtentX        =   6800
         _ExtentY        =   714
      End
      Begin MSComctlLib.ProgressBar prgProgress 
         Height          =   315
         Left            =   1860
         TabIndex        =   0
         Top             =   1950
         Width           =   8145
         _ExtentX        =   14367
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   1
      End
      Begin Threed.SSPanel pnlHeader 
         Height          =   705
         Left            =   30
         TabIndex        =   6
         Top             =   0
         Width           =   10635
         _ExtentX        =   18759
         _ExtentY        =   1244
         _Version        =   131073
         PictureBackgroundStyle=   2
      End
      Begin prjLedgerReport.uctlTextBox txtPercent 
         Height          =   465
         Left            =   1860
         TabIndex        =   1
         Top             =   2280
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   820
      End
      Begin prjLedgerReport.uctlDate uctlToDate 
         Height          =   405
         Left            =   1860
         TabIndex        =   13
         Top             =   1470
         Width           =   3855
         _ExtentX        =   6800
         _ExtentY        =   714
      End
      Begin MSComctlLib.ImageList ImageList1 
         Left            =   0
         Top             =   0
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   16
         ImageHeight     =   16
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   1
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmExportSupItem.frx":27A2
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin VB.Label lblExportType 
         Alignment       =   1  'Right Justify
         Caption         =   "lblOrderBy"
         Height          =   315
         Left            =   5880
         TabIndex        =   15
         Top             =   1080
         Width           =   885
      End
      Begin Threed.SSCommand cmdStart 
         Height          =   525
         Left            =   1890
         TabIndex        =   2
         Top             =   2820
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmExportSupItem.frx":307C
         ButtonStyle     =   3
      End
      Begin VB.Label Label1 
         Caption         =   "Label1"
         Height          =   345
         Left            =   3600
         TabIndex        =   11
         Top             =   2400
         Width           =   1275
      End
      Begin VB.Label lblProgress 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   435
         Left            =   210
         TabIndex        =   10
         Top             =   2010
         Width           =   1575
      End
      Begin VB.Label lblPercent 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   435
         Left            =   210
         TabIndex        =   9
         Top             =   2430
         Width           =   1575
      End
      Begin VB.Label lblMasterName 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   435
         Left            =   210
         TabIndex        =   8
         Top             =   1500
         Width           =   1575
      End
      Begin VB.Label lblFileName 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   435
         Left            =   210
         TabIndex        =   7
         Top             =   1080
         Width           =   1575
      End
      Begin Threed.SSCommand cmdExit 
         Cancel          =   -1  'True
         Height          =   525
         Left            =   8535
         TabIndex        =   4
         Top             =   2820
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   926
         _Version        =   131073
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdOK 
         Height          =   525
         Left            =   6885
         TabIndex        =   3
         Top             =   2820
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmExportSupItem.frx":3396
         ButtonStyle     =   3
      End
   End
End
Attribute VB_Name = "frmExportSupItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private m_HasActivate As Boolean
Private m_HasModify As Boolean
Private m_Rs As ADODB.Recordset

Public ID As Long
Public OKClick As Boolean
Public ShowMode As SHOW_MODE_TYPE
Public HeaderText As String
Private Sub cmdOK_Click()
   OKClick = True
   Unload Me
End Sub
Private Sub Form_Activate()
Dim FromDate As Date
Dim ToDate As Date

   If Not m_HasActivate Then
      m_HasActivate = True
      Me.Refresh
      DoEvents

      Call GetFirstLastDate(Now, FromDate, ToDate)
      uctlFromDate.ShowDate = FromDate
      uctlToDate.ShowDate = ToDate
      
      Call InitExportType(cboExportType)
      
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
'   ElseIf Shift = 0 And KeyCode = 117 Then
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
   ElseIf Shift = 0 And KeyCode = 123 Then
      Call AddMemoNote
      KeyCode = 0
   End If
End Sub

Private Sub ResetStatus()
   prgProgress.MAX = 100
   prgProgress.MIN = 0
   prgProgress.Value = 0
   txtPercent.Text = 0
End Sub
Private Sub InitFormLayout()
   pnlHeader.PictureBackground = LoadPicture(glbParameterObj.NormalForm1)
   SSFrame1.PictureBackground = LoadPicture(glbParameterObj.NormalForm1)
   
   Me.Caption = HeaderText
   pnlHeader.Caption = "EXPORT ข้อมูล ซัพพลายเออร์"
   
   Me.Picture = LoadPicture(glbParameterObj.MainPicture)
   
   pnlHeader.Font.Name = GLB_FONT
   pnlHeader.Font.Bold = True
   pnlHeader.Font.Size = 19
   
   Call InitNormalLabel(lblFileName, "จากวันที่")
   Call InitNormalLabel(lblMasterName, "ถึงวันที่")
   Call InitNormalLabel(lblProgress, "ความคืบหน้า")
   Call InitNormalLabel(lblPercent, "เปอร์เซนต์")
   Call InitNormalLabel(lblExportType, "ประเภท")
   Call InitNormalLabel(Label1, "%")

   Call txtPercent.SetTextLenType(TEXT_FLOAT, glbSetting.MONEY_TYPE)
   txtPercent.Enabled = False

   cmdExit.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdOK.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdStart.Picture = LoadPicture(glbParameterObj.NormalButton1)
   
   Call InitMainButton(cmdExit, MapText("ยกเลิก (ESC)"))
   Call InitMainButton(cmdOK, MapText("ตกลง (F2)"))
   Call InitMainButton(cmdStart, MapText("เริ่ม"))
   
   Call InitCombo(cboExportType)
   
   Call ResetStatus
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
   
   m_HasActivate = False
   Call InitFormLayout
   Call EnableForm(Me, True)
End Sub

Private Sub Form_Unload(Cancel As Integer)
   '
End Sub
Private Sub cmdStart_Click()
Dim TempID As Long
Dim HasBegin As Boolean
   
   If Not VerifyDate(lblFileName, uctlFromDate, False) Then
      Exit Sub
   End If
   If Not VerifyDate(lblMasterName, uctlToDate, False) Then
      Exit Sub
   End If
   If Not VerifyCombo(lblExportType, cboExportType, False) Then
      Exit Sub
   End If
      
   Call EnableForm(Me, False)
   
   TempID = cboExportType.ItemData(Minus2Zero(cboExportType.ListIndex))
   
   If TempID = 1 Then
      Call ExportSupplier
   ElseIf TempID = 2 Then
      Call ExportPartItem
   ElseIf TempID = 3 Then
      Call ExportSupPo
   ElseIf TempID = 4 Then
      Call ExportSupItem
   ElseIf TempID = 5 Then
      Call ExportSupCnDn
   End If
   
   Call EnableForm(Me, True)
End Sub
Private Sub ExportSupItem()
On Error GoTo ErrorHandler
Dim IsOK As Boolean
Dim Si As CSupItem
Dim iCount As Long
Dim FileID As Long
Dim OldID As Long
Dim I As Long

Dim LocationSave As String
   
   prgProgress.MAX = 100
   prgProgress.MIN = 0


   Set Si = New CSupItem
   Si.SUP_ITEM_ID = -1
   Si.FROM_DATE = uctlFromDate.ShowDate
   Si.TO_DATE = uctlToDate.ShowDate
   Si.DOCUMENT_TYPE_SET = "(100,101,102,103)"
   Call Si.QueryData(102, m_Rs, iCount)
   
   LocationSave = "C:\Export ไปยังสำนักงานใหญ่\"
   
   If Dir(LocationSave, vbDirectory) = "" Then
      MkDir (LocationSave)
   End If
   
   LocationSave = LocationSave & Format(Year(Now), "0000") & Format(Month(Now), "00") & Format(Day(Now), "00")
   LocationSave = LocationSave & "_" & Format(Year(uctlFromDate.ShowDate), "0000") & Format(Month(uctlFromDate.ShowDate), "00") & Format(Day(uctlFromDate.ShowDate), "00")
   LocationSave = LocationSave & "_" & Format(Year(uctlToDate.ShowDate), "0000") & Format(Month(uctlToDate.ShowDate), "00") & Format(Day(uctlToDate.ShowDate), "00")
   LocationSave = LocationSave & "_SupItem.txt"
   'LocationSave = "C:\1234.txt"
      
On Error GoTo XXX
   Call Kill(LocationSave)
XXX:


   FileID = FreeFile
   Open LocationSave For Append As #FileID
   
   If Not m_Rs.EOF Then
      Call Si.PopulateFromRS(102, m_Rs)
      
      Call Si.GenerateBDHeader(FileID)
      OldID = Si.DO_ID
   End If

   prgProgress.MAX = 100
   prgProgress.MIN = 0
   prgProgress.Value = 0

   I = 0
   While Not m_Rs.EOF
      prgProgress.Value = MyDiff(I, iCount) * 100

      Call Si.PopulateFromRS(102, m_Rs)
      If OldID <> Si.DO_ID Then
         Call Si.GenerateBDHeader(FileID)
         OldID = Si.DO_ID
      End If

      'Generate detail here
      Call Si.GenerateSupTailer(FileID)
      m_Rs.MoveNext
      I = I + 1
   Wend
   Close #FileID

   Set Si = Nothing
   prgProgress.Value = 100
   txtPercent.Text = 100

   Exit Sub

ErrorHandler:

   cmdStart.Enabled = True
   cmdExit.Enabled = True
   cmdOK.Enabled = True

   glbErrorLog.LocalErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowUserError
   Close #FileID
End Sub
Private Sub ExportSupplier()
On Error GoTo ErrorHandler
Dim IsOK As Boolean
Dim Si As CSupplier
Dim iCount As Long
Dim FileID As Long
Dim OldID As Long
Dim I As Long

Dim LocationSave As String
   
   prgProgress.MAX = 100
   prgProgress.MIN = 0

   
   Set Si = New CSupplier
   Si.SUPPLIER_ID = -1
   Si.FROM_DATE = uctlFromDate.ShowDate
   Si.TO_DATE = uctlToDate.ShowDate
   Call Si.QueryData2(m_Rs, iCount)
   
   LocationSave = "C:\Export ไปยังสำนักงานใหญ่\"
   
   If Dir(LocationSave, vbDirectory) = "" Then
      MkDir (LocationSave)
   End If
   
   LocationSave = LocationSave & Format(Year(Now), "0000") & Format(Month(Now), "00") & Format(Day(Now), "00")
   LocationSave = LocationSave & "_" & Format(Year(uctlFromDate.ShowDate), "0000") & Format(Month(uctlFromDate.ShowDate), "00") & Format(Day(uctlFromDate.ShowDate), "00")
   LocationSave = LocationSave & "_" & Format(Year(uctlToDate.ShowDate), "0000") & Format(Month(uctlToDate.ShowDate), "00") & Format(Day(uctlToDate.ShowDate), "00")
   LocationSave = LocationSave & "_Supplier.txt"
   'LocationSave = "C:\1234.txt"
      
On Error GoTo XXX
   Call Kill(LocationSave)
XXX:


   FileID = FreeFile
   Open LocationSave For Append As #FileID
   
'   If Not m_Rs.EOF Then
'      Call Si.PopulateFromRS(1, m_Rs)
'
'      Call Si.GenerateSPHeader(FileID)
'      'OldID = Si.DO_ID
'   End If

   prgProgress.MAX = 100
   prgProgress.MIN = 0
   prgProgress.Value = 0
   
   I = 0
   While Not m_Rs.EOF
      prgProgress.Value = MyDiff(I, iCount) * 100
      
      Call Si.PopulateFromRS(1, m_Rs)
      'If OldID <> Si.DO_ID Then
         Call Si.GenerateSPHeader(FileID)
         'OldID = Si.DO_ID
      'End If
      
      'Generate detail here
      'Call Si.GenerateSupTailer(FileID)
      m_Rs.MoveNext
      I = I + 1
   Wend
   Close #FileID
   
   Set Si = Nothing
   prgProgress.Value = 100
   txtPercent.Text = 100
   
   Exit Sub

ErrorHandler:

   cmdStart.Enabled = True
   cmdExit.Enabled = True
   cmdOK.Enabled = True

   glbErrorLog.LocalErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowUserError
   Close #FileID
End Sub
Private Sub ExportPartItem()
On Error GoTo ErrorHandler
Dim IsOK As Boolean
Dim Si As CPartItem
Dim iCount As Long
Dim FileID As Long
Dim OldID As Long
Dim I As Long

Dim LocationSave As String
   
   prgProgress.MAX = 100
   prgProgress.MIN = 0

   
   Set Si = New CPartItem
   Si.PART_ITEM_ID = -1
   Si.FROM_DATE = uctlFromDate.ShowDate
   Si.TO_DATE = uctlToDate.ShowDate
   Call Si.QueryData(1, m_Rs, iCount)
   
   LocationSave = "C:\Export ไปยังสำนักงานใหญ่\"
   
   If Dir(LocationSave, vbDirectory) = "" Then
      MkDir (LocationSave)
   End If
   
   LocationSave = LocationSave & Format(Year(Now), "0000") & Format(Month(Now), "00") & Format(Day(Now), "00")
   LocationSave = LocationSave & "_" & Format(Year(uctlFromDate.ShowDate), "0000") & Format(Month(uctlFromDate.ShowDate), "00") & Format(Day(uctlFromDate.ShowDate), "00")
   LocationSave = LocationSave & "_" & Format(Year(uctlToDate.ShowDate), "0000") & Format(Month(uctlToDate.ShowDate), "00") & Format(Day(uctlToDate.ShowDate), "00")
   LocationSave = LocationSave & "_PartItem.txt"
   'LocationSave = "C:\1234.txt"
      
On Error GoTo XXX
   Call Kill(LocationSave)
XXX:


   FileID = FreeFile
   Open LocationSave For Append As #FileID
   
'   If Not m_Rs.EOF Then
'      Call Si.PopulateFromRS(1, m_Rs)
'
'      Call Si.GenerateSPHeader(FileID)
'      'OldID = Si.DO_ID
'   End If

   prgProgress.MAX = 100
   prgProgress.MIN = 0
   prgProgress.Value = 0

   I = 0
   While Not m_Rs.EOF
      prgProgress.Value = MyDiff(I, iCount) * 100

      Call Si.PopulateFromRS(1, m_Rs)
      'If OldID <> Si.DO_ID Then
         Call Si.GeneratePIHeader(FileID)
         'OldID = Si.DO_ID
      'End If
      
      'Generate detail here
      'Call Si.GenerateSupTailer(FileID)
      m_Rs.MoveNext
      I = I + 1
   Wend
   Close #FileID

   Set Si = Nothing
   prgProgress.Value = 100
   txtPercent.Text = 100
   
   Exit Sub

ErrorHandler:

   cmdStart.Enabled = True
   cmdExit.Enabled = True
   cmdOK.Enabled = True

   glbErrorLog.LocalErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowUserError
   Close #FileID
End Sub
Private Sub ExportSupCnDn()
On Error GoTo ErrorHandler
Dim IsOK As Boolean
Dim Si As CReceiptItem
Dim iCount As Long
Dim FileID As Long
Dim OldID As Long
Dim I As Long

Dim LocationSave As String
   
   prgProgress.MAX = 100
   prgProgress.MIN = 0
   
   
   Set Si = New CReceiptItem
   Si.RECEIPT_ITEM_ID = -1
   Si.FROM_DOC_DATE = uctlFromDate.ShowDate
   Si.TO_DOC_DATE = uctlToDate.ShowDate
   Si.DocTypeSet = "(9,10,110)"
   Call Si.QueryData(105, m_Rs, iCount)
   
   LocationSave = "C:\Export ไปยังสำนักงานใหญ่\"
   
   If Dir(LocationSave, vbDirectory) = "" Then
      MkDir (LocationSave)
   End If
   
   LocationSave = LocationSave & Format(Year(Now), "0000") & Format(Month(Now), "00") & Format(Day(Now), "00")
   LocationSave = LocationSave & "_" & Format(Year(uctlFromDate.ShowDate), "0000") & Format(Month(uctlFromDate.ShowDate), "00") & Format(Day(uctlFromDate.ShowDate), "00")
   LocationSave = LocationSave & "_" & Format(Year(uctlToDate.ShowDate), "0000") & Format(Month(uctlToDate.ShowDate), "00") & Format(Day(uctlToDate.ShowDate), "00")
   LocationSave = LocationSave & "_CnDnRtItem.txt"
   'LocationSave = "C:\1234.txt"
      
On Error GoTo XXX
   Call Kill(LocationSave)
XXX:
   
   
   FileID = FreeFile
   Open LocationSave For Append As #FileID
   
   If Not m_Rs.EOF Then
      Call Si.PopulateFromRS(105, m_Rs)
      
      Call Si.GenerateBDHeader(FileID)
      OldID = Si.BILLING_DOC_ID
   End If
   
   prgProgress.MAX = 100
   prgProgress.MIN = 0
   prgProgress.Value = 0

   I = 0
   While Not m_Rs.EOF
      prgProgress.Value = MyDiff(I, iCount) * 100

      Call Si.PopulateFromRS(105, m_Rs)
      If OldID <> Si.BILLING_DOC_ID Then
         Call Si.GenerateBDHeader(FileID)
         OldID = Si.BILLING_DOC_ID
      End If

      'Generate detail here
      Call Si.GenerateRcpTailer(FileID)
      m_Rs.MoveNext
      I = I + 1
   Wend
   Close #FileID

   Set Si = Nothing
   prgProgress.Value = 100
   txtPercent.Text = 100

   Exit Sub

ErrorHandler:

   cmdStart.Enabled = True
   cmdExit.Enabled = True
   cmdOK.Enabled = True

   glbErrorLog.LocalErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowUserError
   Close #FileID
End Sub
Private Sub ExportSupPo()
On Error GoTo ErrorHandler
Dim IsOK As Boolean
Dim Si As CSupItem
Dim iCount As Long
Dim FileID As Long
Dim OldID As Long
Dim I As Long

Dim LocationSave As String
   
   prgProgress.MAX = 100
   prgProgress.MIN = 0
   
   Set Si = New CSupItem
   Si.SUP_ITEM_ID = -1
   Si.FROM_DATE = uctlFromDate.ShowDate
   Si.TO_DATE = uctlToDate.ShowDate
   Si.DOCUMENT_TYPE_SET = "(1000,1001,1002,1003)"
   Call Si.QueryData(102, m_Rs, iCount)
   
   LocationSave = "C:\Export ไปยังสำนักงานใหญ่\"
   
   If Dir(LocationSave, vbDirectory) = "" Then
      MkDir (LocationSave)
   End If
   
   LocationSave = LocationSave & Format(Year(Now), "0000") & Format(Month(Now), "00") & Format(Day(Now), "00")
   LocationSave = LocationSave & "_" & Format(Year(uctlFromDate.ShowDate), "0000") & Format(Month(uctlFromDate.ShowDate), "00") & Format(Day(uctlFromDate.ShowDate), "00")
   LocationSave = LocationSave & "_" & Format(Year(uctlToDate.ShowDate), "0000") & Format(Month(uctlToDate.ShowDate), "00") & Format(Day(uctlToDate.ShowDate), "00")
   LocationSave = LocationSave & "_SupPo.txt"
   'LocationSave = "C:\1234.txt"
      
On Error GoTo XXX
   Call Kill(LocationSave)
XXX:


   FileID = FreeFile
   Open LocationSave For Append As #FileID
   
   If Not m_Rs.EOF Then
      Call Si.PopulateFromRS(102, m_Rs)
      
      Call Si.GenerateBDHeader(FileID)
      OldID = Si.DO_ID
   End If

   prgProgress.MAX = 100
   prgProgress.MIN = 0
   prgProgress.Value = 0

   I = 0
   While Not m_Rs.EOF
      prgProgress.Value = MyDiff(I, iCount) * 100

      Call Si.PopulateFromRS(102, m_Rs)
      If OldID <> Si.DO_ID Then
         Call Si.GenerateBDHeader(FileID)
         OldID = Si.DO_ID
      End If

      'Generate detail here
      Call Si.GenerateSupTailer(FileID)
      m_Rs.MoveNext
      I = I + 1
   Wend
   Close #FileID

   Set Si = Nothing
   prgProgress.Value = 100
   txtPercent.Text = 100

   Exit Sub

ErrorHandler:

   cmdStart.Enabled = True
   cmdExit.Enabled = True
   cmdOK.Enabled = True

   glbErrorLog.LocalErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowUserError
   Close #FileID
End Sub
