VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#2.0#0"; "THREED20.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{E684D8A3-716C-4E59-AA94-7144C04B0074}#1.1#0"; "GridEX20.ocx"
Begin VB.Form frmExportChartAccount 
   BackColor       =   &H80000000&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   8400
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11805
   Icon            =   "frmExportChartAccount.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8400
   ScaleWidth      =   11805
   StartUpPosition =   1  'CenterOwner
   Begin Threed.SSFrame SSFrame1 
      Height          =   8445
      Left            =   0
      TabIndex        =   18
      Top             =   0
      Width           =   11865
      _ExtentX        =   20929
      _ExtentY        =   14896
      _Version        =   131073
      PictureBackgroundStyle=   2
      Begin prjLedgerReport.uctlDate uctlFromDate 
         Height          =   375
         Left            =   1800
         TabIndex        =   8
         Top             =   2280
         Width           =   3855
         _ExtentX        =   6800
         _ExtentY        =   661
      End
      Begin MSComctlLib.ProgressBar prgProgress 
         Height          =   315
         Left            =   1860
         TabIndex        =   16
         Top             =   6660
         Width           =   7905
         _ExtentX        =   13944
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   1
      End
      Begin Threed.SSPanel pnlHeader 
         Height          =   705
         Left            =   0
         TabIndex        =   19
         Top             =   0
         Width           =   11835
         _ExtentX        =   20876
         _ExtentY        =   1244
         _Version        =   131073
         PictureBackgroundStyle=   2
      End
      Begin prjLedgerReport.uctlTextBox txtPercent 
         Height          =   465
         Left            =   1860
         TabIndex        =   17
         Top             =   6990
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   820
      End
      Begin MSComDlg.CommonDialog dlgAdd 
         Left            =   11280
         Top             =   720
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin prjLedgerReport.uctlTextBox txtFileName 
         Height          =   435
         Left            =   1800
         TabIndex        =   0
         Top             =   810
         Width           =   6720
         _ExtentX        =   11853
         _ExtentY        =   767
      End
      Begin GridEX20.GridEX GridEX1 
         Height          =   3705
         Left            =   180
         TabIndex        =   24
         Top             =   2760
         Width           =   9615
         _ExtentX        =   16960
         _ExtentY        =   6535
         Version         =   "2.0"
         BoundColumnIndex=   ""
         ReplaceColumnIndex=   ""
         TabKeyBehavior  =   1
         MethodHoldFields=   -1  'True
         AllowColumnDrag =   0   'False
         AllowEdit       =   0   'False
         BorderStyle     =   3
         GroupByBoxVisible=   0   'False
         DataMode        =   99
         HeaderFontName  =   "AngsanaUPC"
         HeaderFontBold  =   -1  'True
         HeaderFontSize  =   14.25
         HeaderFontWeight=   700
         FontSize        =   9.75
         BackColorBkg    =   16777215
         ColumnHeaderHeight=   480
         IntProp1        =   0
         IntProp2        =   0
         IntProp7        =   0
         ColumnsCount    =   2
         Column(1)       =   "frmExportChartAccount.frx":27A2
         Column(2)       =   "frmExportChartAccount.frx":286A
         FormatStylesCount=   5
         FormatStyle(1)  =   "frmExportChartAccount.frx":290E
         FormatStyle(2)  =   "frmExportChartAccount.frx":2A6A
         FormatStyle(3)  =   "frmExportChartAccount.frx":2B1A
         FormatStyle(4)  =   "frmExportChartAccount.frx":2BCE
         FormatStyle(5)  =   "frmExportChartAccount.frx":2CA6
         ImageCount      =   0
         PrinterProperties=   "frmExportChartAccount.frx":2D5E
      End
      Begin prjLedgerReport.uctlTextBox txtCollumn 
         Height          =   435
         Left            =   1800
         TabIndex        =   4
         Top             =   1800
         Width           =   600
         _ExtentX        =   1905
         _ExtentY        =   767
      End
      Begin prjLedgerReport.uctlTextBox txtRow 
         Height          =   435
         Left            =   4200
         TabIndex        =   3
         Top             =   1320
         Width           =   600
         _ExtentX        =   1905
         _ExtentY        =   767
      End
      Begin prjLedgerReport.uctlTextBox txtSheet 
         Height          =   435
         Left            =   1800
         TabIndex        =   2
         Top             =   1320
         Width           =   600
         _ExtentX        =   1905
         _ExtentY        =   767
      End
      Begin prjLedgerReport.uctlTextBox txtCollumn2 
         Height          =   435
         Left            =   4200
         TabIndex        =   5
         Top             =   1800
         Width           =   600
         _ExtentX        =   1058
         _ExtentY        =   767
      End
      Begin prjLedgerReport.uctlTextBox txtCollumn3 
         Height          =   435
         Left            =   6480
         TabIndex        =   6
         Top             =   1800
         Width           =   720
         _ExtentX        =   1270
         _ExtentY        =   767
      End
      Begin prjLedgerReport.uctlDate uctlToDate 
         Height          =   375
         Left            =   7440
         TabIndex        =   9
         Top             =   2280
         Width           =   3855
         _ExtentX        =   6800
         _ExtentY        =   661
      End
      Begin prjLedgerReport.uctlTextBox txtCollumn4 
         Height          =   435
         Left            =   9000
         TabIndex        =   7
         Top             =   1800
         Width           =   600
         _ExtentX        =   1058
         _ExtentY        =   767
      End
      Begin prjLedgerReport.uctlTextBox txtRow2 
         Height          =   435
         Left            =   6480
         TabIndex        =   33
         Top             =   1320
         Width           =   720
         _ExtentX        =   1270
         _ExtentY        =   767
      End
      Begin prjLedgerReport.uctlTextBox txtTemptxt 
         Height          =   465
         Left            =   6120
         TabIndex        =   36
         Top             =   7680
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   820
      End
      Begin VB.Label lblName 
         Caption         =   "Label1"
         Height          =   345
         Left            =   5160
         TabIndex        =   37
         Top             =   7800
         Width           =   915
      End
      Begin Threed.SSCommand cmdOther 
         Height          =   525
         Left            =   3480
         TabIndex        =   35
         Top             =   7620
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         ButtonStyle     =   3
      End
      Begin VB.Label lblRow2 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   435
         Left            =   4920
         TabIndex        =   34
         Top             =   1440
         Width           =   1455
      End
      Begin VB.Label lblCollumn4 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   315
         Left            =   7680
         TabIndex        =   32
         Top             =   1920
         Width           =   1215
      End
      Begin VB.Label lblToDate 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   435
         Left            =   6360
         TabIndex        =   30
         Top             =   2400
         Width           =   975
      End
      Begin VB.Label lblFromDate 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   435
         Left            =   720
         TabIndex        =   31
         Top             =   2400
         Width           =   975
      End
      Begin VB.Label lblCollumn3 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   315
         Left            =   4920
         TabIndex        =   29
         Top             =   1920
         Width           =   1455
      End
      Begin VB.Label lblCollumn2 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   435
         Left            =   2880
         TabIndex        =   28
         Top             =   1920
         Width           =   1215
      End
      Begin Threed.SSCommand cmdDelete 
         Height          =   525
         Left            =   9960
         TabIndex        =   12
         Top             =   5640
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmExportChartAccount.frx":2F36
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdAdd 
         Height          =   525
         Left            =   9960
         TabIndex        =   10
         Top             =   3960
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmExportChartAccount.frx":3250
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdEdit 
         Height          =   525
         Left            =   9960
         TabIndex        =   11
         Top             =   4800
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   926
         _Version        =   131073
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdFileName 
         Height          =   435
         Left            =   8600
         TabIndex        =   1
         Top             =   810
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   767
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmExportChartAccount.frx":356A
         ButtonStyle     =   3
      End
      Begin VB.Label lblSheet 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   435
         Left            =   720
         TabIndex        =   27
         Top             =   1440
         Width           =   975
      End
      Begin VB.Label lblRow 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   435
         Left            =   2640
         TabIndex        =   26
         Top             =   1440
         Width           =   1455
      End
      Begin VB.Label lblCollumn 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   435
         Left            =   480
         TabIndex        =   25
         Top             =   1920
         Width           =   1215
      End
      Begin VB.Label lblFileName 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   435
         Left            =   120
         TabIndex        =   23
         Top             =   930
         Width           =   1575
      End
      Begin Threed.SSCommand cmdStart 
         Height          =   525
         Left            =   1860
         TabIndex        =   13
         Top             =   7620
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmExportChartAccount.frx":3884
         ButtonStyle     =   3
      End
      Begin VB.Label Label1 
         Caption         =   "Label1"
         Height          =   345
         Left            =   3600
         TabIndex        =   22
         Top             =   7110
         Width           =   1275
      End
      Begin VB.Label lblProgress 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   435
         Left            =   210
         TabIndex        =   21
         Top             =   6720
         Width           =   1575
      End
      Begin VB.Label lblPercent 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   435
         Left            =   210
         TabIndex        =   20
         Top             =   7140
         Width           =   1575
      End
      Begin Threed.SSCommand cmdExit 
         Height          =   525
         Left            =   9615
         TabIndex        =   15
         Top             =   7620
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   926
         _Version        =   131073
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdOK 
         Height          =   525
         Left            =   7965
         TabIndex        =   14
         Top             =   7620
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         ButtonStyle     =   3
      End
   End
End
Attribute VB_Name = "frmExportChartAccount"
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
Private MainCollection As Collection
Private SearchCollection As Collection
Private SearchNameCollection As Collection

Private m_ExcelApp As Object
Private m_ExcelSheet As Object
Private ConFigID As Long
Private Sub cmdAdd_Click()
Dim OKClick As Boolean
Dim lMenuChosen As Long
Dim oMenu As cPopupMenu

   If Not cmdAdd.Enabled Then
      Exit Sub
   End If
      
   OKClick = False
   frmExportChartAccountEx.HeaderText = "������¡������"
   frmExportChartAccountEx.ShowMode = SHOW_ADD
   Set frmExportChartAccountEx.ParentForm = Me
   Set frmExportChartAccountEx.TempCollection = MainCollection
   Load frmExportChartAccountEx
   frmExportChartAccountEx.Show 1
   
   OKClick = frmExportChartAccountEx.OKClick
   Unload frmExportChartAccountEx
   Set frmExportChartAccountEx = Nothing
   
   If OKClick Then
      m_HasModify = True
      GridEX1.ItemCount = CountItem(MainCollection)
      GridEX1.Rebind
   End If
   
End Sub
Private Sub cmdDelete_Click()
Dim ID1 As Long
Dim ID2 As Long

   If Not cmdDelete.Enabled Then
      Exit Sub
   End If

   If Not VerifyGrid(GridEX1.Value(1)) Then
      Exit Sub
   End If

   If Not ConfirmDelete(GridEX1.Value(3)) Then
      Exit Sub
   End If

   ID2 = GridEX1.Value(2)
   ID1 = GridEX1.Value(1)

   If ID1 <= 0 Then
      MainCollection.Remove (ID2)
   Else
      MainCollection.ITEM(ID2).Flag = "D"
   End If

   GridEX1.ItemCount = CountItem(MainCollection)
   GridEX1.Rebind
   m_HasModify = True
   
End Sub

Private Sub cmdEdit_Click()
Dim IsOK As Boolean
Dim ItemCount As Long
Dim IsCanLock As Boolean
Dim ID As Long
Dim OKClick As Boolean

      
   If Not VerifyGrid(GridEX1.Value(1)) Then
      Exit Sub
   End If

   ID = Val(GridEX1.Value(2))
   OKClick = False
      
   frmExportChartAccountEx.HeaderText = "�����¡������"
   frmExportChartAccountEx.ShowMode = SHOW_EDIT
   frmExportChartAccountEx.ID = ID
   Set frmExportChartAccountEx.ParentForm = Me
   Set frmExportChartAccountEx.TempCollection = MainCollection
   
   Load frmExportChartAccountEx
   frmExportChartAccountEx.Show 1
   
   OKClick = frmExportChartAccountEx.OKClick
   Unload frmExportChartAccountEx
   Set frmExportChartAccountEx = Nothing
   
   If OKClick Then
      m_HasModify = True
      GridEX1.ItemCount = CountItem(MainCollection)
      GridEX1.Rebind
   End If
   
End Sub

Private Sub cmdFileName_Click()
On Error Resume Next
Dim strDescription As String
   
   'edit the filter to support more image types
   dlgAdd.Filter = "Access Files (*.xls)|*.xls;"
   dlgAdd.DialogTitle = "Select access file to import"
   dlgAdd.ShowOpen
   If dlgAdd.FileName = "" Then
      Exit Sub
   End If
    
   txtFileName.Text = dlgAdd.FileName
   m_HasModify = True
End Sub

Private Sub cmdOK_Click()
Dim Ac As CAccountCode
   For Each Ac In MainCollection
      If Ac.Flag = "A" Then
         Ac.AddEditMode = SHOW_ADD
         Call Ac.AddEditData
      ElseIf Ac.Flag = "E" Then
         Ac.AddEditMode = SHOW_EDIT
         Call Ac.AddEditData
      ElseIf Ac.Flag = "D" Then
         Call Ac.DeleteData
      End If
   Next Ac
   
   OKClick = True
   Unload Me
End Sub

Private Sub cmdOther_Click()
Dim oMenu As cPopupMenu
Dim lMenuChosen As Long
Dim Ac As CAccountCode
Dim ItemCount As Long
Dim m_Rs As ADODB.Recordset
   Set oMenu = New cPopupMenu
   lMenuChosen = oMenu.Popup("�Ѵ�͡������ ��������ѷ (��������ͧ���� ���� ��������ͧ��ҷ���ͧ��¡��)")
   If lMenuChosen = 0 Then
      Exit Sub
   End If
   Set oMenu = Nothing
   
   Call EnableForm(Me, False)
   If lMenuChosen = 1 Then
      Set Ac = New CAccountCode
      Set m_Rs = New ADODB.Recordset
      
      txtTemptxt.Enabled = True
      
      If Not VerifyTextControl(lblName, txtTemptxt) Then
         Call EnableForm(Me, True)
         Exit Sub
      End If
   
      Ac.ENTERPRISE_CODE = txtTemptxt.Text
      Call Ac.QueryData(m_Rs, ItemCount, False)
      
      If ItemCount > 0 Then
         While Not m_Rs.EOF
            Call Ac.PopulateFromRS(1, m_Rs)
            Ac.AddEditMode = SHOW_ADD
            Call Ac.AddEditData
            m_Rs.MoveNext
         Wend
         glbErrorLog.LocalErrorMsg = MapText("COPY SUCCESS ��س� REBOOT �к�����")
         glbErrorLog.ShowUserError
      Else
         glbErrorLog.LocalErrorMsg = MapText("��辺�����Ţͧ����ѷ") & " " & txtTemptxt.Text & " " & MapText("��к�����")
         glbErrorLog.ShowUserError
      End If
      Set Ac = Nothing
   End If
   
   Call EnableForm(Me, True)
   
End Sub

Private Sub cmdStart_Click()
Dim TempID As Long
Dim HasBegin As Boolean
Dim MaxSheet As Long
Dim Ac As CAccountCode

   For Each Ac In MainCollection
      If Ac.Flag = "A" Then
         Ac.AddEditMode = SHOW_ADD
         Call Ac.AddEditData
      ElseIf Ac.Flag = "E" Then
         Ac.AddEditMode = SHOW_EDIT
         Call Ac.AddEditData
      ElseIf Ac.Flag = "D" Then
         Call Ac.DeleteData
      End If
   Next Ac
   
   If Not VerifyTextControl(lblFileName, txtFileName) Then
      Exit Sub
   End If
   
   If Not VerifyTextControl(lblCollumn, txtCollumn) Then
      Exit Sub
   End If
   
   If Not VerifyTextControl(lblCollumn2, txtCollumn2) Then
      Exit Sub
   End If
   
   If Not VerifyTextControl(lblCollumn3, txtCollumn3) Then
      Exit Sub
   End If
   
   If Not VerifyTextControl(lblCollumn4, txtCollumn4) Then
      Exit Sub
   End If
   
   If Not VerifyTextControl(lblRow, txtRow) Then
      Exit Sub
   End If
   
   If Not VerifyTextControl(lblRow2, txtRow2) Then
      Exit Sub
   End If
   
   If Not VerifyTextControl(lblSheet, txtSheet) Then
      Exit Sub
   End If
   
   Call SaveData
   
   Call EnableForm(Me, False)
   m_ExcelApp.Workbooks.Close
   m_ExcelApp.Workbooks.Open (txtFileName.Text)
   
   MaxSheet = m_ExcelApp.Sheets.Count
   
   If Val(txtSheet.Text) > MaxSheet Then
      Call MsgBox("��سҡ�͡������ �մ���١��ͧ���������ö�ҡ����  " & MaxSheet, vbOKOnly, PROJECT_NAME)
      Exit Sub
   End If
   
   Call ExportAccount
   
   DoEvents
   
'    m_ExcelApp.Workbooks.Saved = True
   
   
    Call m_ExcelApp.Workbooks.Close
   
'   m_ExcelApp.Workbooks.Close
   Call EnableForm(Me, True)
 
End Sub
Private Sub ExportAccount()
On Error GoTo ErrorHandler
Dim IsOK As Boolean
Dim iCount As Long
Dim i As Long
Dim TempNo As String
Dim TempName As String
Dim Gl As CGLJnl
Dim j As Long
Dim Ac As CAccountCode
Dim Debit As Double
Dim Credit As Double
Dim FindCode As Boolean
Dim MaxRow As Long
Dim GC As CGLAcc
Dim ReturnV As Boolean
Dim ReturnZ As Boolean
Dim Deposit As Double

   prgProgress.MAX = 100
   prgProgress.MIN = 0
   
   Call LoadGLJNLforAccountExcel(Nothing, SearchCollection, uctlFromDate.ShowDate, uctlToDate.ShowDate)
   Call LoadGLAccSearch(Nothing, SearchNameCollection)
   
   prgProgress.MAX = 100
   prgProgress.MIN = 0
   prgProgress.Value = 0

   i = 0
   
   Set m_ExcelSheet = m_ExcelApp.Sheets(Val(txtSheet.Text))
   
   j = 0
   
   iCount = Val(txtRow2.Text) - Val(txtRow.Text)
   While (j < iCount)
      j = j + 1
      prgProgress.Value = MyDiff(j, iCount) * 100
      txtPercent.Text = prgProgress.Value
      Me.Refresh
      
      TempNo = m_ExcelSheet.Cells(Val(txtRow.Text) + j - 1, Val(txtCollumn.Text)).Value   ' �֧����ҡ excell
      If TempNo <> "" Then
         
         If Trim(TempNo) = "212-1100" Then
            'debug.print ("")
         End If
         
         FindCode = False
         For Each Ac In MainCollection                                                                           ' ǹ�ҡ�������� access ��
            If Ac.MAIN_CODE = Trim(TempNo) Then
               FindCode = True
               Exit For
            End If
         Next Ac
         
         Set GC = GetGLAcc(SearchNameCollection, Trim(TempNo))                  ' �֧�Ҩҡ excell ���ҧ CGLAcc
         
         m_ExcelSheet.Cells(Val(txtRow.Text) + j - 1, Val(txtCollumn4.Text)).Value = GC.ACCNAM       ' �Թʴ-�����  ,, txtCollumn4.Text =��������Ǩ 21
         
         ReturnV = True
         ReturnZ = True
         
         Debit = 0
         Credit = 0
         Deposit = 0
         
         If Not (FindCode) Then
            Deposit = GC.BEGCUR
            If GC.GROUP = 2 Or GC.GROUP = 3 Then
               Deposit = Deposit * -1
            End If
            Set Gl = GetGLJnl(SearchCollection, Trim(TempNo) & "-0", ReturnV)
            Debit = Gl.AMOUNT
            Set Gl = GetGLJnl(SearchCollection, Trim(TempNo) & "-1", ReturnZ)
            Credit = Gl.AMOUNT
            If Deposit + Debit - Credit > 0 Then
               m_ExcelSheet.Cells(Val(txtRow.Text) + j - 1, Val(txtCollumn2.Text)).Value = Abs(Deposit + Debit - Credit)
            ElseIf Deposit + Debit - Credit < 0 Then
               m_ExcelSheet.Cells(Val(txtRow.Text) + j - 1, Val(txtCollumn3.Text)).Value = Abs(Deposit + Debit - Credit)
            End If
         Else
            Debit = 0
            Credit = 0
            Deposit = 0
            Call Recuresive(Trim(Ac.SUB_CODE), Debit, Credit, ReturnV, ReturnZ, Deposit)
            If ReturnV Then
               m_ExcelSheet.Cells(Val(txtRow.Text) + j - 1, Val(txtCollumn2.Text)).Font.colorindex = 1
            Else
               m_ExcelSheet.Cells(Val(txtRow.Text) + j - 1, Val(txtCollumn2.Text)).Font.colorindex = 3
            End If
            If ReturnZ Then
               m_ExcelSheet.Cells(Val(txtRow.Text) + j - 1, Val(txtCollumn3.Text)).Font.colorindex = 1
            Else
               m_ExcelSheet.Cells(Val(txtRow.Text) + j - 1, Val(txtCollumn3.Text)).Font.colorindex = 3
            End If
            If Deposit + Debit - Credit > 0 Then
               m_ExcelSheet.Cells(Val(txtRow.Text) + j - 1, Val(txtCollumn2.Text)).Value = Abs(Deposit + Debit - Credit)
            ElseIf Deposit + Debit - Credit < 0 Then
               m_ExcelSheet.Cells(Val(txtRow.Text) + j - 1, Val(txtCollumn3.Text)).Value = Abs(Deposit + Debit - Credit)
            End If
         End If
         If ReturnV Or ReturnZ Then
            m_ExcelSheet.Cells(Val(txtRow.Text) + j - 1, Val(txtCollumn4.Text)).Font.colorindex = 1
         Else
            m_ExcelSheet.Cells(Val(txtRow.Text) + j - 1, Val(txtCollumn4.Text)).Font.colorindex = 3
         End If
      End If
   Wend
   
   prgProgress.Value = 100
   txtPercent.Text = 100
      
   Exit Sub
   
ErrorHandler:
   glbErrorLog.LocalErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowUserError
End Sub
Private Sub Recuresive(Text As String, Debit As Double, Credit As Double, ReturnV As Boolean, ReturnZ As Boolean, Deposit As Double)
Dim Gl As CGLJnl
Dim TempNo As String
Dim Pos As Long
Dim ReturnX As Boolean
Dim Ac As CGLAcc
   
   Pos = InStr(1, Text, ",")
   If Pos = 0 Then
      TempNo = Text
      Text = ""
   Else
      TempNo = Left(Text, Pos - 1)
      Text = Mid(Text, Pos + 1, Len(Text) - Pos)
   End If
   
   Set Ac = GetGLAcc(SearchNameCollection, Trim(TempNo))
   
   Set Gl = GetGLJnl(SearchCollection, Trim(TempNo) & "-0", ReturnX)
   ReturnV = ReturnV And ReturnX
   Debit = Debit + Gl.AMOUNT
   Set Gl = GetGLJnl(SearchCollection, Trim(TempNo) & "-1")
   ReturnZ = ReturnZ And ReturnX
   Credit = Credit + Gl.AMOUNT
   
   If Ac.GROUP = 2 Or Ac.GROUP = 3 Then
      Deposit = Deposit - Ac.BEGCUR
   Else
      Deposit = Deposit + Ac.BEGCUR
   End If
   
   If Text <> "" Then
      Call Recuresive(Text, Debit, Credit, ReturnV, ReturnZ, Deposit)
   End If
End Sub

Private Sub Form_Activate()
Dim FromDate As Date
Dim ToDate As Date

   If Not m_HasActivate Then
      m_HasActivate = True
      Me.Refresh
      DoEvents
            
      m_HasModify = False
            
      Call LoadAccountCode(Nothing, MainCollection)
      
      Call QueryData
      
      GridEX1.ItemCount = CountItem(MainCollection)
      GridEX1.Rebind
      
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
      Call cmdAdd_Click
      KeyCode = 0
   ElseIf Shift = 0 And KeyCode = 117 Then
      Call cmdDelete_Click
      KeyCode = 0
   ElseIf Shift = 0 And KeyCode = 113 Then
      Call cmdOK_Click
      KeyCode = 0
   ElseIf Shift = 0 And KeyCode = 114 Then
      Call cmdEdit_Click
      KeyCode = 0
   ElseIf Shift = 0 And KeyCode = 121 Then
      'Call cmdPrint_Click
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
   pnlHeader.Caption = "Export �����żѧ�ѭ��"
   
   Me.Picture = LoadPicture(glbParameterObj.MainPicture)
   
   pnlHeader.Font.Name = GLB_FONT
   pnlHeader.Font.Bold = True
   pnlHeader.Font.Size = 19
   
   Call InitNormalLabel(lblName, "��¡��")
   
   Call InitNormalLabel(lblProgress, "�����׺˹��")
   Call InitNormalLabel(lblPercent, "����ૹ��")
   Call InitNormalLabel(Label1, "%")
   Call InitNormalLabel(lblFileName, "�������")
   
   Call InitNormalLabel(lblCollumn, "�����������")
   Call InitNormalLabel(lblRow, "�������")
   Call InitNormalLabel(lblRow2, "�Ǩ�")
   Call InitNormalLabel(lblSheet, "�մ")
   Call InitNormalLabel(lblCollumn2, "�������ഺԵ")
   Call InitNormalLabel(lblCollumn3, "��������ôԵ")
   Call InitNormalLabel(lblCollumn4, "��������Ǩ")
   
   Call InitNormalLabel(lblFromDate, "�ҡ�ѹ���")
   Call InitNormalLabel(lblToDate, "�֧�ѹ���")
   
   Call txtPercent.SetTextLenType(TEXT_FLOAT, glbSetting.MONEY_TYPE)
   txtPercent.Enabled = False
   Call txtFileName.SetTextLenType(TEXT_STRING, glbSetting.DESC_TYPE)
   'txtFileName.Enabled = False
   Call txtCollumn.SetTextLenType(TEXT_FLOAT, glbSetting.MONEY_TYPE)
   Call txtRow.SetTextLenType(TEXT_FLOAT, glbSetting.MONEY_TYPE)
   Call txtRow2.SetTextLenType(TEXT_FLOAT, glbSetting.MONEY_TYPE)
   Call txtSheet.SetTextLenType(TEXT_FLOAT, glbSetting.MONEY_TYPE)
   Call txtCollumn2.SetTextLenType(TEXT_FLOAT, glbSetting.MONEY_TYPE)
   Call txtCollumn3.SetTextLenType(TEXT_FLOAT, glbSetting.MONEY_TYPE)
   Call txtCollumn4.SetTextLenType(TEXT_FLOAT, glbSetting.MONEY_TYPE)
   Call txtTemptxt.SetTextLenType(TEXT_STRING, glbSetting.CODE_TYPE)
   
   txtTemptxt.Enabled = False
   
   cmdExit.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdOK.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdStart.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdFileName.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdAdd.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdEdit.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdDelete.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdOther.Picture = LoadPicture(glbParameterObj.NormalButton1)
   
   Call InitMainButton(cmdExit, MapText("¡��ԡ (ESC)"))
   Call InitMainButton(cmdOK, MapText("��ŧ (F2)"))
   Call InitMainButton(cmdStart, MapText("�����"))
   Call InitMainButton(cmdFileName, MapText("..."))
   Call InitMainButton(cmdAdd, MapText("���� (F7)"))
   Call InitMainButton(cmdEdit, MapText("��� (F3)"))
   Call InitMainButton(cmdDelete, MapText("ź (F6)"))
   
   Call InitMainButton(cmdOther, MapText("����"))
   
   Call InitGrid1
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
   Set MainCollection = New Collection
   Set SearchCollection = New Collection
   Set SearchNameCollection = New Collection
   
   Set m_ExcelApp = CreateObject("Excel.application")
   
   m_HasActivate = False
   Call InitFormLayout
   Call EnableForm(Me, True)
End Sub

Private Sub InitGrid1()
Dim Col As JSColumn

   GridEX1.Columns.Clear
   GridEX1.BackColor = GLB_GRID_COLOR
   GridEX1.ItemCount = 0
   GridEX1.BackColorHeader = GLB_GRIDHD_COLOR
   GridEX1.ColumnHeaderFont.Bold = True
   GridEX1.ColumnHeaderFont.Name = GLB_FONT
   GridEX1.TabKeyBehavior = jgexControlNavigation
   
   Set Col = GridEX1.Columns.Add '1
   Col.Width = 0
   Col.Caption = "ID"

   Set Col = GridEX1.Columns.Add '2
   Col.Width = 0
   Col.Caption = "Real ID"
   
   Set Col = GridEX1.Columns.Add '3
   Col.Width = 1500
   Col.Caption = MapText("������ѡ")
   
   Set Col = GridEX1.Columns.Add '4
   Col.Width = 12000
   Col.Caption = MapText("��������")
   
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set MainCollection = Nothing
   Set SearchCollection = Nothing
   Set SearchNameCollection = Nothing
   
   Call m_ExcelApp.Workbooks.Close
  ' Call m_ExcelApp.Close
End Sub
Private Sub GridEX1_ColResize(ByVal ColIndex As Integer, ByVal NewColWidth As Long, ByVal Cancel As GridEX20.JSRetBoolean)
   'debug.print ColIndex & " " & NewColWidth
End Sub
Private Sub GridEX1_DblClick()
   Call cmdEdit_Click
End Sub
Private Sub GridEX1_UnboundReadData(ByVal RowIndex As Long, ByVal Bookmark As Variant, ByVal Values As GridEX20.JSRowData)
On Error GoTo ErrorHandler
Dim RealIndex As Long

   glbErrorLog.ModuleName = Me.Name
   glbErrorLog.RoutineName = "UnboundReadData"

   If MainCollection Is Nothing Then
      Exit Sub
   End If

   If RowIndex <= 0 Then
      Exit Sub
   End If

   Dim CR As CAccountCode
   If MainCollection.Count <= 0 Then
      Exit Sub
   End If
   Set CR = GetItem(MainCollection, RowIndex, RealIndex)
   If CR Is Nothing Then
      Exit Sub
   End If

   Values(1) = CR.ACCOUNT_CODE_ID
   Values(2) = RealIndex
   Values(3) = CR.MAIN_CODE
   Values(4) = CR.SUB_CODE
      
Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub
Public Sub RefreshGrid()
   GridEX1.ItemCount = CountItem(MainCollection)
   GridEX1.Rebind
End Sub
Private Sub QueryData()
Dim AG As CAccountConfig
Dim ItemCount As Long
Dim Rs As ADODB.Recordset
   
   Set AG = New CAccountConfig
   Set Rs = New ADODB.Recordset
   
   AG.ACCOUNT_CONFIG_ID = -1
   Call AG.QueryData(Rs, ItemCount)
   If ItemCount > 0 Then
      Call AG.PopulateFromRS(1, Rs)
      ShowMode = SHOW_EDIT
      
      ConFigID = AG.ACCOUNT_CONFIG_ID
      txtSheet.Text = AG.SHEET
      txtRow.Text = AG.ROW
      txtRow2.Text = AG.ROW2
      txtCollumn.Text = AG.COLLUMN_CODE
      txtCollumn2.Text = AG.COLLUMN_DEBIT
      txtCollumn3.Text = AG.COLLUMN_CREDIT
      txtCollumn4.Text = AG.COLLUMN_CHECK
      uctlFromDate.ShowDate = AG.FROM_DATE
      uctlToDate.ShowDate = AG.TO_DATE
   Else
      ShowMode = SHOW_ADD
   End If
   
   If Rs.State = adStateOpen Then
      Rs.Close
   End If
   Set Rs = Nothing
End Sub
Private Sub SaveData()
Dim AG As CAccountConfig
Set AG = New CAccountConfig

   AG.AddEditMode = ShowMode
   AG.ACCOUNT_CONFIG_ID = ConFigID
   AG.SHEET = Val(txtSheet.Text)
   AG.ROW = Val(txtRow.Text)
   AG.ROW2 = Val(txtRow2.Text)
   AG.COLLUMN_CODE = Val(txtCollumn.Text)
   AG.COLLUMN_DEBIT = Val(txtCollumn2.Text)
   AG.COLLUMN_CREDIT = Val(txtCollumn3.Text)
   AG.COLLUMN_CHECK = Val(txtCollumn4.Text)
   AG.FROM_DATE = uctlFromDate.ShowDate
   AG.TO_DATE = uctlToDate.ShowDate
   
   Call AG.AddEditData
End Sub

