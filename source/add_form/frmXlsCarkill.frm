VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#2.0#0"; "THREED20.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmXlsCarkill 
   BackColor       =   &H80000000&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   9540
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   12000
   Icon            =   "frmXlsCarkill.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9540
   ScaleWidth      =   12000
   StartUpPosition =   1  'CenterOwner
   Begin Threed.SSFrame SSFrame1 
      Height          =   9525
      Left            =   0
      TabIndex        =   27
      Top             =   0
      Width           =   11985
      _ExtentX        =   21140
      _ExtentY        =   16801
      _Version        =   131073
      PictureBackgroundStyle=   2
      Begin prjLedgerReport.uctlDate uctlFromDate 
         Height          =   375
         Left            =   3000
         TabIndex        =   2
         Top             =   1560
         Width           =   3855
         _ExtentX        =   6800
         _ExtentY        =   661
      End
      Begin MSComctlLib.ProgressBar prgProgress 
         Height          =   315
         Left            =   1680
         TabIndex        =   25
         Top             =   7320
         Width           =   7905
         _ExtentX        =   13944
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   1
      End
      Begin Threed.SSPanel pnlHeader 
         Height          =   705
         Left            =   0
         TabIndex        =   28
         Top             =   0
         Width           =   11955
         _ExtentX        =   21087
         _ExtentY        =   1244
         _Version        =   131073
         PictureBackgroundStyle=   2
      End
      Begin prjLedgerReport.uctlTextBox txtPercent 
         Height          =   465
         Left            =   1680
         TabIndex        =   26
         Top             =   7800
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
         Left            =   3000
         TabIndex        =   1
         Top             =   960
         Width           =   6720
         _ExtentX        =   11853
         _ExtentY        =   767
      End
      Begin prjLedgerReport.uctlTextBox txtRow 
         Height          =   435
         Left            =   6000
         TabIndex        =   10
         Top             =   2760
         Width           =   600
         _ExtentX        =   1905
         _ExtentY        =   767
      End
      Begin prjLedgerReport.uctlTextBox txtSheet 
         Height          =   435
         Left            =   3960
         TabIndex        =   9
         Top             =   2760
         Width           =   600
         _ExtentX        =   1905
         _ExtentY        =   767
      End
      Begin prjLedgerReport.uctlTextBox txtPeriodDate 
         Height          =   435
         Left            =   3000
         TabIndex        =   3
         Top             =   2040
         Width           =   2640
         _ExtentX        =   1058
         _ExtentY        =   767
      End
      Begin prjLedgerReport.uctlTextBox txtRow2 
         Height          =   435
         Left            =   8160
         TabIndex        =   11
         Top             =   2760
         Width           =   600
         _ExtentX        =   1270
         _ExtentY        =   767
      End
      Begin prjLedgerReport.uctlTextBox txtTemptxt 
         Height          =   465
         Left            =   6120
         TabIndex        =   38
         Top             =   8760
         Visible         =   0   'False
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   820
      End
      Begin prjLedgerReport.uctlTextBox txtSheet6 
         Height          =   435
         Left            =   3960
         TabIndex        =   17
         Top             =   5520
         Width           =   600
         _ExtentX        =   1905
         _ExtentY        =   767
      End
      Begin prjLedgerReport.uctlTextBox txtRow6 
         Height          =   435
         Left            =   6000
         TabIndex        =   18
         Top             =   5520
         Width           =   600
         _ExtentX        =   1905
         _ExtentY        =   767
      End
      Begin prjLedgerReport.uctlTextBox txtColumn6 
         Height          =   435
         Left            =   8160
         TabIndex        =   19
         Top             =   5520
         Width           =   600
         _ExtentX        =   1270
         _ExtentY        =   767
      End
      Begin prjLedgerReport.uctlTextBox txtRow7 
         Height          =   435
         Left            =   6000
         TabIndex        =   23
         Top             =   6480
         Width           =   600
         _ExtentX        =   1905
         _ExtentY        =   767
      End
      Begin prjLedgerReport.uctlTextBox txtColumn7 
         Height          =   435
         Left            =   8160
         TabIndex        =   24
         Top             =   6480
         Width           =   600
         _ExtentX        =   1270
         _ExtentY        =   767
      End
      Begin prjLedgerReport.uctlTextBox txtColumnTotal 
         Height          =   435
         Left            =   6000
         TabIndex        =   16
         Top             =   4200
         Width           =   600
         _ExtentX        =   1905
         _ExtentY        =   767
      End
      Begin prjLedgerReport.uctlTextBox txtRowTotal 
         Height          =   435
         Left            =   10920
         TabIndex        =   48
         Top             =   7800
         Visible         =   0   'False
         Width           =   600
         _ExtentX        =   1905
         _ExtentY        =   767
      End
      Begin prjLedgerReport.uctlTextBox txtColumn8 
         Height          =   435
         Left            =   6000
         TabIndex        =   13
         Top             =   3240
         Width           =   600
         _ExtentX        =   1905
         _ExtentY        =   767
      End
      Begin prjLedgerReport.uctlTextBox txtRow8 
         Height          =   435
         Left            =   3960
         TabIndex        =   12
         Top             =   3240
         Width           =   600
         _ExtentX        =   1905
         _ExtentY        =   767
      End
      Begin prjLedgerReport.uctlTextBox txtFileOutName 
         Height          =   435
         Left            =   3000
         TabIndex        =   55
         Top             =   4920
         Width           =   6720
         _ExtentX        =   11853
         _ExtentY        =   767
      End
      Begin prjLedgerReport.uctlTextBox txtSheet7 
         Height          =   435
         Left            =   3960
         TabIndex        =   22
         Top             =   6480
         Width           =   600
         _ExtentX        =   1905
         _ExtentY        =   767
      End
      Begin prjLedgerReport.uctlTextBox txtColumn9 
         Height          =   435
         Left            =   6000
         TabIndex        =   15
         Top             =   3720
         Width           =   600
         _ExtentX        =   1905
         _ExtentY        =   767
      End
      Begin prjLedgerReport.uctlTextBox txtRow9 
         Height          =   435
         Left            =   3960
         TabIndex        =   14
         Top             =   3720
         Width           =   600
         _ExtentX        =   1905
         _ExtentY        =   767
      End
      Begin prjLedgerReport.uctlTextBox txtRow6_1 
         Height          =   435
         Left            =   6000
         TabIndex        =   20
         Top             =   6000
         Width           =   600
         _ExtentX        =   1905
         _ExtentY        =   767
      End
      Begin prjLedgerReport.uctlTextBox txtColumn6_2 
         Height          =   435
         Left            =   8160
         TabIndex        =   21
         Top             =   6000
         Width           =   600
         _ExtentX        =   1270
         _ExtentY        =   767
      End
      Begin VB.Label lblReal6_0 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   435
         Left            =   2760
         TabIndex        =   63
         Top             =   6000
         Width           =   1815
      End
      Begin VB.Label lblColumn6_2 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   435
         Left            =   6840
         TabIndex        =   62
         Top             =   6000
         Width           =   1215
      End
      Begin VB.Label lblRow6_1 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   435
         Left            =   4800
         TabIndex        =   61
         Top             =   6000
         Width           =   1095
      End
      Begin VB.Label lblSheetName 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   435
         Left            =   960
         TabIndex        =   60
         Top             =   3720
         Width           =   1815
      End
      Begin VB.Label lblColumn9 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   435
         Left            =   4800
         TabIndex        =   59
         Top             =   3720
         Width           =   1095
      End
      Begin VB.Label lblRow9 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   435
         Left            =   3000
         TabIndex        =   58
         Top             =   3720
         Width           =   855
      End
      Begin VB.Label lblSheet7 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   435
         Left            =   3000
         TabIndex        =   57
         Top             =   6480
         Width           =   855
      End
      Begin VB.Label lblFileOutName 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   435
         Left            =   600
         TabIndex        =   56
         Top             =   4920
         Width           =   2175
      End
      Begin Threed.SSCommand cmdFileOutName 
         Height          =   435
         Left            =   9840
         TabIndex        =   4
         Top             =   4920
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   767
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmXlsCarkill.frx":27A2
         ButtonStyle     =   3
      End
      Begin VB.Label lblRow8 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   435
         Left            =   3000
         TabIndex        =   54
         Top             =   3240
         Width           =   855
      End
      Begin VB.Label lblColumn8 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   435
         Left            =   4800
         TabIndex        =   53
         Top             =   3240
         Width           =   1095
      End
      Begin VB.Label lblShortName 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   435
         Left            =   960
         TabIndex        =   52
         Top             =   3240
         Width           =   1815
      End
      Begin VB.Label lblTotalData 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   435
         Left            =   2760
         TabIndex        =   51
         Top             =   4200
         Width           =   1815
      End
      Begin VB.Label lblColumnTotal 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   435
         Left            =   4800
         TabIndex        =   50
         Top             =   4200
         Width           =   1095
      End
      Begin VB.Label lblRowTotal 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   435
         Left            =   9960
         TabIndex        =   49
         Top             =   7800
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.Label lblRow7 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   435
         Left            =   4800
         TabIndex        =   47
         Top             =   6480
         Width           =   1095
      End
      Begin VB.Label lblColumn7 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   435
         Left            =   6840
         TabIndex        =   46
         Top             =   6480
         Width           =   1215
      End
      Begin VB.Label lblPrint7 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   435
         Left            =   960
         TabIndex        =   45
         Top             =   6480
         Width           =   1815
      End
      Begin VB.Label lblRow6 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   435
         Left            =   4800
         TabIndex        =   44
         Top             =   5520
         Width           =   1095
      End
      Begin VB.Label lblColumn6 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   435
         Left            =   6840
         TabIndex        =   43
         Top             =   5520
         Width           =   1215
      End
      Begin VB.Label lblSheet6 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   435
         Left            =   3000
         TabIndex        =   42
         Top             =   5520
         Width           =   855
      End
      Begin VB.Label lblReal6 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   435
         Left            =   960
         TabIndex        =   41
         Top             =   5520
         Width           =   1815
      End
      Begin VB.Label lblRunData 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   435
         Left            =   960
         TabIndex        =   40
         Top             =   2760
         Width           =   1815
      End
      Begin VB.Label lblName 
         Caption         =   "Label1"
         Height          =   345
         Left            =   5160
         TabIndex        =   39
         Top             =   8880
         Visible         =   0   'False
         Width           =   915
      End
      Begin Threed.SSCommand cmdSetting 
         Height          =   525
         Left            =   3240
         TabIndex        =   6
         Top             =   8760
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
         Left            =   6840
         TabIndex        =   37
         Top             =   2760
         Width           =   1215
      End
      Begin VB.Label lblPeriodDate 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   615
         Left            =   240
         TabIndex        =   36
         Top             =   2040
         Width           =   2535
      End
      Begin VB.Label lblFromDate 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   435
         Left            =   1800
         TabIndex        =   35
         Top             =   1560
         Width           =   975
      End
      Begin Threed.SSCommand cmdFileName 
         Height          =   435
         Left            =   9840
         TabIndex        =   0
         Top             =   960
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   767
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmXlsCarkill.frx":2ABC
         ButtonStyle     =   3
      End
      Begin VB.Label lblSheet 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   435
         Left            =   3000
         TabIndex        =   34
         Top             =   2760
         Width           =   855
      End
      Begin VB.Label lblRow 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   435
         Left            =   4800
         TabIndex        =   33
         Top             =   2760
         Width           =   1095
      End
      Begin VB.Label lblFileName 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   435
         Left            =   600
         TabIndex        =   32
         Top             =   960
         Width           =   2175
      End
      Begin Threed.SSCommand cmdStart 
         Height          =   525
         Left            =   1620
         TabIndex        =   5
         Top             =   8760
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmXlsCarkill.frx":2DD6
         ButtonStyle     =   3
      End
      Begin VB.Label Label1 
         Caption         =   "Label1"
         Height          =   345
         Left            =   3480
         TabIndex        =   31
         Top             =   7800
         Width           =   1275
      End
      Begin VB.Label lblProgress 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   435
         Left            =   -30
         TabIndex        =   30
         Top             =   7320
         Width           =   1575
      End
      Begin VB.Label lblPercent 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   435
         Left            =   -30
         TabIndex        =   29
         Top             =   7800
         Width           =   1575
      End
      Begin Threed.SSCommand cmdExit 
         Height          =   525
         Left            =   9960
         TabIndex        =   8
         Top             =   8700
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   926
         _Version        =   131073
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdOK 
         Height          =   525
         Left            =   8280
         TabIndex        =   7
         Top             =   8700
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         ButtonStyle     =   3
      End
   End
End
Attribute VB_Name = "frmXlsCarkill"
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
Dim MaxSheet As Long

Dim tempCarkill As CXlsCarkill
Private m_collCarkill(30) As Collection
Private tempCollCarkill As Collection
Private sheetHaveData As Collection
Dim tempSheet As CXlsCarkill
Private shortnameHaveData As Collection
Dim tempShortname As CXlsCarkill
Private col_XlsSum As Collection
Private col_XlsFW As Collection

Dim FirstDate As Date
Dim LastDate As Date
Dim D As String
Dim tempDateString As String
Dim TempFromdate  As Date
Dim TempToDate   As Date

Dim m_collPeriodDate As Collection
Dim TempPeriodDate As CXlsCarkill
Dim m_collRealFarm As Collection
Dim TempRealFarm As CXlsCarkill
Dim TempRealFarm2 As CXlsCarkill



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

Private Sub cmdFileOutName_Click()
On Error Resume Next
Dim strDescription As String

   'edit the filter to support more image types
   dlgAdd.Filter = "Access Files (*.xls)|*.xls;"
   dlgAdd.DialogTitle = "Select access file to import"
   dlgAdd.ShowOpen
   If dlgAdd.FileName = "" Then
      Exit Sub
   End If

   txtFileOutName.Text = dlgAdd.FileName
   m_HasModify = True
End Sub

Private Sub cmdOK_Click()
   If Not SaveData Then
      Exit Sub
   End If
   
   OKClick = True
 '  Unload Me
End Sub

Private Sub cmdsetting_Click()
Dim oMenu As cPopupMenu
Dim lMenuChosen As Long
'Dim Ac As CAccountCode
'Dim ItemCount As Long
'Dim m_Rs As ADODB.Recordset

   Set oMenu = New cPopupMenu
   lMenuChosen = oMenu.Popup("��駤�Һ�÷Ѵ�������", "��駤�Һ�÷Ѵ�������¡�")
   If lMenuChosen = 0 Then
      Exit Sub
   End If
   Set oMenu = Nothing

   Call EnableForm(Me, False)
   If lMenuChosen = 1 Then
      ' �Դ��������
      'frmCarkillSettingSum.ShowMode = SHOW_EDIT
      Load frmCarkillSettingSum
      frmCarkillSettingSum.Show 1

      Unload frmCarkillSettingSum
      Set frmCarkillSettingSum = Nothing
   End If

   Call EnableForm(Me, True)
    If lMenuChosen = 2 Then
      ' �Դ��������
      'frmCarkillSettingSum.ShowMode = SHOW_EDIT
      Load frmCarkillSetFW
      frmCarkillSetFW.Show 1

      Unload frmCarkillSetFW
      Set frmCarkillSetFW = Nothing
   End If

   Call EnableForm(Me, True)
   
End Sub

Private Sub cmdStart_Click()
Dim TempID As Long
Dim HasBegin As Boolean
Dim RunSheet As Long
Dim TempFromdate As Date
Dim TempToDate As Date
Dim Ac As CAccountCode
Dim DateCount  As Long
Dim runIndex As Long
Dim i As Long
   
   Call EnableForm(Me, False)

   m_ExcelApp.Workbooks.Close
   m_ExcelApp.Workbooks.Open (txtFileName.Text)
   
   MaxSheet = m_ExcelApp.Sheets.Count

   If Val(txtSheet.Text) > MaxSheet Then
      Call MsgBox("��سҡ�͡������ �մ���١��ͧ���������ö�ҡ����  " & MaxSheet, vbOKOnly, PROJECT_NAME)
      m_ExcelApp.Workbooks.Close
      Call EnableForm(Me, True)
      Exit Sub
   End If

   Call SaveData
   
' �Ҫش�ѹ���
   Set tempCarkill = New CXlsCarkill
   Set m_collPeriodDate = New Collection
   tempDateString = txtPeriodDate.Text
   TempFromdate = uctlFromDate.ShowDate
   i = 0
   
   While Len(tempDateString) > 0
      i = i + 1
      Call CalculateDatePeriod(tempDateString, DateCount)
      TempToDate = DateAdd("D", DateCount - 1, TempFromdate)
      
      tempCarkill.FromDate = TempFromdate
      tempCarkill.ToDate = TempToDate
      tempCarkill.DateIndex = i
      
      Call m_collPeriodDate.Add(tempCarkill, Str(i))                                 'Trim(tempCarkill.InvNo) & "-" & Trim(tempCarkill.FeedNo)
      Set tempCarkill = Nothing
      Set tempCarkill = New CXlsCarkill
      
       TempFromdate = DateAdd("D", 1, TempToDate)
   Wend
   ' -------- �� �Ҫش�ѹ���

   Call CalculateDate(txtPeriodDate.Text, DateCount)         ' 31 �ѹ
   FirstDate = uctlFromDate.ShowDate
   LastDate = DateAdd("D", DateCount - 1, uctlFromDate.ShowDate)
   
   For RunSheet = Val(txtSheet.Text) To MaxSheet                  ' ǹ�����շ�������� ���֧ �շ�ش����
      '�Ѵ���͡��÷Ѵ��������Ҥ��
      '1  ��� A  ��ҧ --> �������ѹ���㹪�ͧ B .. Ship date
      '                  --> ������ 15 㹪�ͧ A
      '2  ��Ҫ�ͧ A ��鹵鹴��� KRT = ������
      '    ��Ҫ�ͧ A ��鹵鹴��� 15 = ���  .......... ������� ��� sum 㹤��仴���
      Call GenCarkillRow(RunSheet)
   Next RunSheet

      '3  ��� �繪�ǧ�ѹ��� ������� D , �Ѻ �
      '4  ���� ��� �������駤�� �ǵ�駤�� �ͧ���Ъշ
      '5 ���Ѿ�����ͧ��ä�� �������� ��ǧ�ѹ��� 1 (, 2, 3, 4) �� ������� , Net Payment ������� �ͧ���п����
      For runIndex = Val(txtSheet.Text) To UBound(m_collCarkill)
         For Each tempCarkill In m_collCarkill(runIndex)                ' ǹ item ��硷������� Coll ,,, coll ˹�� = �շ
              Call SumByDate(tempCarkill, runIndex)
         Next tempCarkill
      Next runIndex
      
   Call GenDistanceFarm
   
'  'debug.print m_collCarkill(2).Count
'  'debug.print m_collPeriodDate.Count
'  'debug.print sheetHaveData.Count
   
   Call GenTotalInputSheet
   
   m_ExcelApp.Workbooks.Close
   m_ExcelApp.Workbooks.Open (txtFileOutName.Text)
   
   Call GenTotalOutputSheet
   Call GenTotalRealSheet
   'ǹ coll ��¹ ���շ��� ������ѹ��鹡���¹㹪շ�á����
   'Call GenTotalinSheet
   
   m_ExcelApp.Workbooks.Close

   Call EnableForm(Me, True)

End Sub
Private Sub GenCarkillRow(RunSheet As Long)
On Error GoTo ErrorHandler
Dim i As Long
Dim j As Long
Dim IsOK As Boolean
Dim iCount As Long
Dim ROW As Long
Dim collumn As Long

Dim PrevKey1 As String
Dim PrevKey2 As String
Dim PrevKey3 As Long
Dim PrevKey4 As Long
Dim PrevKey5 As Date
Dim tempShipDate As Date
Dim farmPeriodDate As CXlsCarkill
Dim TempPeriodDate As CXlsCarkill
Dim SumPeriodDate As CXlsCarkill
Dim DateIndex As Long
Dim ShortName As String
Dim BlankRow As Long

   prgProgress.MAX = 100
   prgProgress.MIN = 0
   prgProgress.Value = 0
   
   Set m_ExcelSheet = m_ExcelApp.Sheets(RunSheet)
   j = 0

   iCount = Val(txtRow2.Text) - Val(txtRow.Text)              ' ������ҷء��÷Ѵ�� 100%
   While (j < iCount)                                                      'And m_ExcelSheet.Cells(Val(txtRow.Text) + j - 1, "B").Value >= uctlFromDate.ShowDate
      collumn = 1
      j = j + 1
      
   If m_ExcelSheet.Cells(Val(txtRow.Text) + j - 1, "B").Value <> "" Then
       tempShipDate = StringSlashToDate(m_ExcelSheet.Cells(Val(txtRow.Text) + j - 1, "B").Value)
   Else
       tempShipDate = PrevKey5
   End If
 '   'debug.print tempShipDate
   
   If FirstDate <= tempShipDate And tempShipDate <= LastDate Then       ' �֧�ѹ����������示�÷Ѵ������
      While (collumn < 19)                                                                    'And m_ExcelSheet.Cells(Val(txtRow.Text) + j - 1, "F").Value <> ""
         
         prgProgress.Value = MyDiff(j, iCount) * 100                ' �������� �����ǹ=0 ����� 0 ������ٻẺ double
         txtPercent.Text = prgProgress.Value
         Me.Refresh

      If collumn = 1 Then
            If m_ExcelSheet.Cells(Val(txtRow.Text) + j - 1, "A").Value = "" Then
               tempCarkill.InvNo = PrevKey1                            ' ��Ҫ�ͧ A ��ҧ �����Ҫ�ͧ A �ͧ��÷Ѵ����
            Else
                tempCarkill.InvNo = m_ExcelSheet.Cells(Val(txtRow.Text) + j - 1, "A").Value    ' �֧����ҡ excell
            End If
            PrevKey1 = tempCarkill.InvNo
        ElseIf collumn = 3 Then
 '           'debug.print m_ExcelSheet.Cells(Val(txtRow.Text) + j - 1, "B").Value
            If m_ExcelSheet.Cells(Val(txtRow.Text) + j - 1, "B").Value = "" Then
               tempCarkill.ShipDate = PrevKey2                            ' ��Ҫ�ͧ A ��ҧ �����Ҫ�ͧ A �ͧ��÷Ѵ����
            Else
                tempCarkill.ShipDate = tempShipDate    ' �֧����ҡ excell
            End If
            PrevKey2 = tempCarkill.ShipDate
         ElseIf collumn = 4 Then
            tempCarkill.FeedNo = m_ExcelSheet.Cells(Val(txtRow.Text) + j - 1, "C").Value    ' �֧����ҡ excell
         ElseIf collumn = 5 Then
            tempCarkill.Quantity = m_ExcelSheet.Cells(Val(txtRow.Text) + j - 1, "D").Value    ' �֧����ҡ excell
            If Val(tempCarkill.Quantity) <= 350 Then
               tempCarkill.Quantity = Str(Val(tempCarkill.Quantity) * 30)
            End If
         ElseIf collumn = 6 Then
            tempCarkill.Price_Bag1 = m_ExcelSheet.Cells(Val(txtRow.Text) + j - 1, "E").Value    ' �֧����ҡ excell
         ElseIf collumn = 7 Then
            tempCarkill.NetPrice = m_ExcelSheet.Cells(Val(txtRow.Text) + j - 1, "F").Value   ' �֧����ҡ excell
         ElseIf collumn = 8 Then
            tempCarkill.Trans = m_ExcelSheet.Cells(Val(txtRow.Text) + j - 1, "G").Value    ' �֧����ҡ excell
         ElseIf collumn = 9 Then
            tempCarkill.TotalAmount = m_ExcelSheet.Cells(Val(txtRow.Text) + j - 1, "H").Value    ' �֧����ҡ excell
         ElseIf collumn = 10 Then
            tempCarkill.Trans50 = m_ExcelSheet.Cells(Val(txtRow.Text) + j - 1, "I").Value    ' �֧����ҡ excell
         ElseIf collumn = 11 Then
            tempCarkill.Price_Bag50 = m_ExcelSheet.Cells(Val(txtRow.Text) + j - 1, "J").Value    ' �֧����ҡ excell
         ElseIf collumn = 12 Then
            If LCase(tempCarkill.NetPrice) <> "total" Then
               tempCarkill.NetPayment = m_ExcelSheet.Cells(Val(txtRow.Text) + j - 1, "K").Value    ' �֧����ҡ excell
            Else
               tempCarkill.NetPayment = tempCarkill.Trans + tempCarkill.Trans50
            End If
         ElseIf collumn = 13 Then
            tempCarkill.CN = m_ExcelSheet.Cells(Val(txtRow.Text) + j - 1, "L").Value   ' �֧����ҡ excell
         ElseIf collumn = 14 Then
            tempCarkill.DueDate = m_ExcelSheet.Cells(Val(txtRow.Text) + j - 1, "M").Value    ' �֧����ҡ excell
         ElseIf collumn = 15 Then
            tempCarkill.Today = m_ExcelSheet.Cells(Val(txtRow.Text) + j - 1, "N").Value    ' �֧����ҡ excell
         ElseIf collumn = 16 Then
            tempCarkill.OutStandingDay = m_ExcelSheet.Cells(Val(txtRow.Text) + j - 1, "O").Value   ' �֧����ҡ excell
         ElseIf collumn = 17 Then
            tempCarkill.Remarks = m_ExcelSheet.Cells(Val(txtRow.Text) + j - 1, "P").Value    ' �֧����ҡ excell
         ElseIf collumn = 18 Then
            tempCarkill.DueDate = m_ExcelSheet.Cells(Val(txtRow.Text) + j - 1, "Q").Value    ' �֧����ҡ excell
         End If
         collumn = collumn + 1
      Wend
  End If
      PrevKey5 = tempShipDate
      
              ' ���Ţ��� Sheet ��� coll ���� ���ਹ ���ҧ�����ҹ��ҧ�շ   '������Ţ㹹�� �������㹿�����������¹�ѹ���
      Set tempSheet = GetObject("CXlscarkill", sheetHaveData, Str(RunSheet), False)
      If tempSheet Is Nothing Then
         Set tempSheet = New CXlsCarkill
         tempSheet.SheetIndex = RunSheet
         tempSheet.mySheetName = m_ExcelSheet.Cells(Val(txtRow9.Text), txtColumn9.Text).Value
         BlankRow = findBlankRow(Val(txtRow.Text))                                               ' �Ҩ�л�Ѻ��ا ����ͧǹ��� ������� coll ���� get ᷹
         tempSheet.ROW = BlankRow
         Call sheetHaveData.Add(tempSheet, Str(RunSheet))
         Set tempSheet = Nothing
      Else
         BlankRow = tempSheet.ROW
      End If
      
      If Not (tempCarkill Is Nothing) And tempCarkill.NetPrice <> "" Then                 'And LCase(tempCarkill.NetPrice) <> "total"
         tempCarkill.SheetIndex = RunSheet
         tempCarkill.ROW = BlankRow
         tempCarkill.SheetName = m_ExcelSheet.Cells(1, "A").Value
         For Each TempPeriodDate In m_collPeriodDate
               If TempPeriodDate.FromDate <= tempCarkill.ShipDate And tempCarkill.ShipDate <= TempPeriodDate.ToDate Then
                   DateIndex = TempPeriodDate.DateIndex
                   tempCarkill.DateIndex = TempPeriodDate.DateIndex
                   Exit For
               End If
          Next TempPeriodDate
          Call m_collCarkill(RunSheet).Add(tempCarkill)
         Set tempCarkill = Nothing
         Set tempCarkill = New CXlsCarkill
         
      If (PrevKey3 <> RunSheet Or DateIndex <> PrevKey4) And DateIndex <> 0 Then
         Set TempPeriodDate = GetObject("CXlsCarkill", m_collPeriodDate, Str(DateIndex))
            tempCarkill.SheetName = m_ExcelSheet.Cells(1, "A").Value
            tempCarkill.SheetIndex = RunSheet
            tempCarkill.DateIndex = DateIndex
            tempCarkill.ROW = BlankRow
            tempCarkill.ShortName = m_ExcelSheet.Cells(Val(txtRow8.Text), txtColumn8.Text).Value                      ' tempCarkill.SheetName
            tempCarkill.mySheetName = m_ExcelSheet.Cells(Val(txtRow9.Text), txtColumn9.Text).Value                      ' tempCarkill.SheetName
            ShortName = m_ExcelSheet.Cells(Val(txtRow8.Text), txtColumn8.Text).Value
            If TempPeriodDate.m_Farm Is Nothing Then
              Set TempPeriodDate.m_Farm = New Collection
            End If
            Call TempPeriodDate.m_Farm.Add(tempCarkill, Str(RunSheet))   '���������㹪�ǧ�ѹ������
            Set tempCarkill = Nothing
            Set tempCarkill = New CXlsCarkill
            
     ' ����Ѻ��� sum
            Set tempShortname = GetObject("CXlscarkill", shortnameHaveData, ShortName, False)
            If tempShortname Is Nothing Then
               Set tempShortname = New CXlsCarkill
               tempShortname.ShortName = ShortName
               Call shortnameHaveData.Add(tempShortname, ShortName)
               Set tempShortname = Nothing
            End If

         End If
         PrevKey3 = RunSheet
         PrevKey4 = DateIndex
      End If
  Wend
  
   prgProgress.Value = 100
   txtPercent.Text = 100
   Exit Sub

ErrorHandler:
   glbErrorLog.LocalErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowUserError
End Sub

Private Sub SumByDate(tempCollCarkill As CXlsCarkill, runIndex As Long)
On Error GoTo ErrorHandler
Dim i As Long
Dim DateCount As Long
Dim j As Long
Dim iCount As Long
Dim SumPeriodDate As CXlsCarkill
Dim TempPeriodDate As CXlsCarkill
Dim TempSumPeriod As CXlsCarkill
Dim TempDistancePeriod As CXlsCarkill
   
   prgProgress.MAX = 100
   prgProgress.MIN = 0
   prgProgress.Value = 0
   Set SumPeriodDate = New CXlsCarkill
   
   ' * tempCollCarkill ��� ����� m_Farm
   
      For Each TempPeriodDate In m_collPeriodDate
         If TempPeriodDate.FromDate <= tempCollCarkill.ShipDate And tempCollCarkill.ShipDate <= TempPeriodDate.ToDate Then
               
               Set SumPeriodDate = GetObject("CXlsCarkill", m_collPeriodDate, Str(TempPeriodDate.DateIndex), False)                     ' �����ͧ m_collPeriodDate
               Set TempPeriodDate = GetObject("CXlsCarkill", SumPeriodDate.m_Farm, Str(tempCollCarkill.SheetIndex), False)           ' SumPeriodDate = ����Ѻ�֧ ������ ��� sheet �Ҩҡ������
               
               TempPeriodDate.SumKilo = TempPeriodDate.SumKilo + Val(tempCollCarkill.Quantity / 2)
               TempPeriodDate.SumNetpay = TempPeriodDate.SumNetpay + Val(tempCollCarkill.NetPayment)
                       
          End If
      Next TempPeriodDate
      
         prgProgress.Value = MyDiff(j, iCount) * 100                ' �������� �����ǹ=0 ����� 0 ������ٻẺ double
         txtPercent.Text = prgProgress.Value
         Me.Refresh

   prgProgress.Value = 100
   txtPercent.Text = 100
   Exit Sub

ErrorHandler:
   glbErrorLog.LocalErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowUserError
End Sub

Private Sub GenDistanceFarm()
On Error GoTo ErrorHandler
Dim i As Long
Dim DateCount As Long
Dim j As Long
Dim iCount As Long
Dim SumPeriodDate As CXlsCarkill
Dim TempPeriodDate As CXlsCarkill
Dim TempSumPeriod As CXlsCarkill
Dim TempDistancePeriod As CXlsCarkill
   
   prgProgress.MAX = 100
   prgProgress.MIN = 0
   prgProgress.Value = 0
   Set SumPeriodDate = New CXlsCarkill

      For Each TempPeriodDate In m_collPeriodDate                             ' ����� m_collPeriodDate
         If Not (TempPeriodDate.m_Farm Is Nothing) Then
         For Each TempSumPeriod In TempPeriodDate.m_Farm               ' ����� m_Farm

            Set TempDistancePeriod = GetObject("CxlsCarkill", TempPeriodDate.m_DistanceFarm, TempSumPeriod.ShortName, False)
               If TempDistancePeriod Is Nothing Then
                  If TempPeriodDate.m_DistanceFarm Is Nothing Then
                       Set TempPeriodDate.m_DistanceFarm = New Collection
                  End If
                 TempSumPeriod.sumFlag = "N"
                 Call TempPeriodDate.m_DistanceFarm.Add(TempSumPeriod, TempSumPeriod.ShortName)    '���������㹪�ǧ�ѹ������
                 Set TempSumPeriod = Nothing
               Else
                  TempDistancePeriod.SumNetpay = TempDistancePeriod.SumNetpay + Val(TempSumPeriod.SumNetpay)
                  TempDistancePeriod.SumKilo = TempDistancePeriod.SumKilo + (Val(TempSumPeriod.SumKilo) / 2)             ' �ѹ��ͧ��� 2
                  Set TempDistancePeriod = Nothing
               End If
               
          Next TempSumPeriod
         End If
      Next TempPeriodDate

         prgProgress.Value = MyDiff(j, iCount) * 100                ' �������� �����ǹ=0 ����� 0 ������ٻẺ double
         txtPercent.Text = prgProgress.Value
         Me.Refresh

   prgProgress.Value = 100
   txtPercent.Text = 100
   Exit Sub

ErrorHandler:
   glbErrorLog.LocalErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowUserError
End Sub

Private Sub GenTotalInputSheet()
On Error GoTo ErrorHandler
Dim i As Long
Dim DateCount As Long
Dim j As Long
Dim iCount As Long
Dim startColumn As Long
Dim startColumn4Sum As Long
Dim DatePrintColumn As Long
Dim printColumn As Long
'Dim RowSetting As Long
Dim GroupPeriodDate As CXlsCarkill
Dim TempPeriodDate As CXlsCarkill
Dim ShortName As String
Dim DateRow As Long
Dim printFinish As Boolean

Dim Now_Kilo As String
Dim Str_Kilo As String
Dim Now_Net As String
Dim Str_Net As String
   
   prgProgress.MAX = 100
   prgProgress.MIN = 0
   prgProgress.Value = 0
   Set GroupPeriodDate = New CXlsCarkill

'   'debug.print m_collPeriodDate.Count
   startColumn = Val(txtColumnTotal.Text)
   startColumn4Sum = Val(txtColumnTotal.Text) + 2
   'RowSetting = Val(txtRowTotal.Text)
   printFinish = False
   DatePrintColumn = 0
   DateRow = 2
   
   For Each tempSheet In sheetHaveData
         Set m_ExcelSheet = m_ExcelApp.Sheets(tempSheet.SheetIndex)
         m_ExcelSheet.Cells(tempSheet.ROW + 2, startColumn).Value = "=" & txtColumn8.Text & txtRow8.Text                               ' tempCarkill.SheetName
         m_ExcelSheet.Cells(tempSheet.ROW, startColumn).Value = "�ѹ���"
         m_ExcelSheet.Cells(tempSheet.ROW + 1, startColumn).Value = "�.�. / �ҷ "
   Next tempSheet
   
  For Each tempSheet In sheetHaveData
      Set m_ExcelSheet = m_ExcelApp.Sheets(tempSheet.SheetIndex)
      Str_Kilo = "="
      Str_Net = "="

      For Each TempPeriodDate In m_collPeriodDate
            m_ExcelSheet.Cells(tempSheet.ROW, startColumn + 1).Value = DateToStringExtEx2(TempPeriodDate.FromDate) & " �֧ " & DateToStringExtEx2(TempPeriodDate.ToDate)
            m_ExcelSheet.Cells(tempSheet.ROW + 1, startColumn + 1).Value = "�.�. "
            m_ExcelSheet.Cells(tempSheet.ROW + 1, startColumn + 2).Value = "�ӹǹ�Թ"
            
            ' ��Ҥ�� column&row �ͧ �� ���Ѵ� excel ��ͧ (startColumn+m_collPeriodDate.Count-1) & (tempSheet.ROW + 1)
            ' now �� = column&row �ͧ ��
            ' tempSum =  m_ExcelSheet.Cells(,).Value
            ' m_ExcelSheet.Cells(,).Value = tempSum + now ��
                        
            Now_Kilo = number2Column(startColumn + 1) & Trim(Str(tempSheet.ROW + 2))
           ' Str_Kilo = Str_Kilo & "+" & number2Column(startColumn + m_collPeriodDate.Count - 1) & Trim(Str(tempSheet.ROW + 2))    ' �Ҩ���� = ������ҧ
            Str_Kilo = Str_Kilo & "+" & Now_Kilo
            m_ExcelSheet.Cells(tempSheet.ROW + 2, startColumn4Sum + (m_collPeriodDate.Count * 2) - 1).Value = Str_Kilo
            ' ��Ҥ�� column&row �ͧ �ӹǹ�Թ ���Ѵ� excel ��ͧ (startColumn+m_collPeriodDate.Count-1) & (tempSheet.ROW + 2)
            Now_Net = number2Column(startColumn + 2) & (tempSheet.ROW + 2)
          '  Str_Net = Str_Net & "+" & number2Column(startColumn + m_collPeriodDate.Count) & Trim(Str(tempSheet.ROW + 2))    ' �Ҩ���� = ������ҧ
            Str_Net = Str_Net & "+" & Now_Net
            m_ExcelSheet.Cells(tempSheet.ROW + 2, startColumn4Sum + (m_collPeriodDate.Count * 2)).Value = Str_Net
 
            startColumn = startColumn + 2                                                            ' ��ͧ����¹��ǧ�ѹ����͹
      Next TempPeriodDate
      
      m_ExcelSheet.Cells(tempSheet.ROW, startColumn + 1).Value = "��� " & FirstDate & "�֧ " & LastDate
      m_ExcelSheet.Cells(tempSheet.ROW + 1, startColumn + 1).Value = "�.�. "
      m_ExcelSheet.Cells(tempSheet.ROW + 1, startColumn + 2).Value = "�ӹǹ�Թ"

      startColumn = Val(txtColumnTotal.Text)
  Next tempSheet
  
    For Each TempPeriodDate In m_collPeriodDate
         If Not (TempPeriodDate.m_Farm Is Nothing) Then
           For Each tempCarkill In TempPeriodDate.m_Farm
               Set m_ExcelSheet = m_ExcelApp.Sheets(tempCarkill.SheetIndex)
               m_ExcelSheet.Cells(tempCarkill.ROW + 2, startColumn + 1).Value = tempCarkill.SumKilo
               m_ExcelSheet.Cells(tempCarkill.ROW + 2, startColumn + 2).Value = tempCarkill.SumNetpay

            
            Set TempRealFarm = GetObject("CXlsCarkill", m_collRealFarm, Trim(tempCarkill.mySheetName), False)               ' ����� Sum ����Ѻ�Դ��ԧ
            If (TempRealFarm Is Nothing) Then
               Set TempRealFarm = New CXlsCarkill
               TempRealFarm.mySheetName = tempCarkill.mySheetName
               TempRealFarm.SheetName = tempCarkill.SheetName
               TempRealFarm.mySheetName = tempCarkill.mySheetName
               TempRealFarm.ShortName = tempCarkill.ShortName
               TempRealFarm.SheetIndex = tempCarkill.SheetIndex
'               TempRealFarm.sigmaKilo = TempRealFarm.sigmaKilo + tempCarkill.SumKilo
'               TempRealFarm.sigmaNetpay = TempRealFarm.sigmaNetpay + tempCarkill.SumNetpay
               Call m_collRealFarm.Add(TempRealFarm, Trim(tempCarkill.mySheetName))
            End If
            
         '      Set TempRealFarm = GetObject("CXlsCarkill", TempRealFarm.m_Farm, tempCarkill.SheetIndex, False)               ' ����� Sum ����Ѻ�Դ��ԧ
         '      If TempRealFarm Is Nothing Then
                  If TempRealFarm.m_Farm Is Nothing Then
                    Set TempRealFarm.m_Farm = New Collection
                  End If
                  TempRealFarm2.ToDate = TempPeriodDate.ToDate
                  TempRealFarm2.FromDate = TempPeriodDate.FromDate
                  TempRealFarm2.SumKilo = tempCarkill.SumKilo
                  TempRealFarm2.SumNetpay = tempCarkill.SumNetpay
                  TempRealFarm2.mySheetName = tempCarkill.mySheetName
                  TempRealFarm2.SheetName = tempCarkill.SheetName
                  TempRealFarm2.mySheetName = tempCarkill.mySheetName
                  TempRealFarm2.SheetIndex = tempCarkill.SheetIndex
                  TempRealFarm2.ShortName = tempCarkill.ShortName
                  TempRealFarm2.DateIndex = tempCarkill.DateIndex
                  TempRealFarm.sigmaKilo = TempRealFarm.sigmaKilo + tempCarkill.SumKilo
                  TempRealFarm.sigmaNetpay = TempRealFarm.sigmaNetpay + tempCarkill.SumNetpay
                  Call TempRealFarm.m_Farm.Add(TempRealFarm2)           ' , Trim(Str(tempCarkill.SheetIndex))
                  Set TempRealFarm2 = Nothing
                  Set TempRealFarm2 = New CXlsCarkill
             '  End If
           
           Next tempCarkill
           startColumn = startColumn + 2                               ' ��ͧ����¹��ǧ�ѹ����͹
         End If
    Next TempPeriodDate
   
   startColumn = 0

    prgProgress.Value = MyDiff(j, iCount) * 100                   ' �������� �����ǹ=0 ����� 0 ������ٻẺ double
    txtPercent.Text = prgProgress.Value
    Me.Refresh

   prgProgress.Value = 100
   txtPercent.Text = 100
   Exit Sub

ErrorHandler:
   glbErrorLog.LocalErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowUserError
End Sub

Private Sub GenTotalRealSheet()
Dim A As Long
Dim B As Long
Dim ROW As Long
Dim initRow As Long
Dim Column As Long
Dim Minus_row As Long
Dim Minus_column As Long
Dim xlsFarmName As String

   Set m_ExcelSheet = m_ExcelApp.Sheets(Val(txtSheet6.Text))
   initRow = Val(txtRow6.Text)
   Column = Val(txtColumn6.Text)
   
      For Each TempPeriodDate In m_collPeriodDate
            m_ExcelSheet.Cells(initRow - 2, Column - 1 + TempPeriodDate.DateIndex + B).Value = DateToStringExtEx2(TempPeriodDate.FromDate) & " �֧ " & DateToStringExtEx2(TempPeriodDate.ToDate)
            m_ExcelSheet.Cells(initRow - 1, Column - 1 + TempPeriodDate.DateIndex + B).Value = "�.�. "
            m_ExcelSheet.Cells(initRow - 1, Column + TempPeriodDate.DateIndex + B).Value = "�ӹǹ�Թ"
            B = B + 1
      Next TempPeriodDate
      m_ExcelSheet.Cells(initRow - 2, Column + (m_collPeriodDate.Count * 2)).Value = "��� " & FirstDate & "�֧ " & LastDate
      m_ExcelSheet.Cells(initRow - 1, Column + (m_collPeriodDate.Count * 2)).Value = "�.�. "
      m_ExcelSheet.Cells(initRow - 1, Column + (m_collPeriodDate.Count * 2) + 1).Value = "�ӹǹ�Թ"

   ' �� get ��� ���к�÷Ѵ
  While (ROW <= 32)
      ROW = initRow + A
      ' 'debug.print ROW
      xlsFarmName = m_ExcelSheet.Cells(Val(ROW), Column - 2).Value
      Set TempRealFarm = GetObject("CXlsCarkill", m_collRealFarm, Trim(xlsFarmName), False)               ' ����� Sum ����Ѻ�Դ��ԧ
      If Not (TempRealFarm Is Nothing) Then
   
         For Each TempRealFarm2 In TempRealFarm.m_Farm
             m_ExcelSheet.Cells(Val(ROW), Column + (TempRealFarm2.DateIndex * 2) - 2).Value = TempRealFarm2.SumKilo
             m_ExcelSheet.Cells(Val(ROW), Column + (TempRealFarm2.DateIndex * 2) - 1).Value = TempRealFarm2.SumNetpay
         Next TempRealFarm2
         
         m_ExcelSheet.Cells(Val(ROW), Column + (m_collPeriodDate.Count * 2)).Value = TempRealFarm.sigmaKilo
         m_ExcelSheet.Cells(Val(ROW), Column + (m_collPeriodDate.Count * 2) + 1).Value = TempRealFarm.sigmaNetpay
      End If
     A = A + 1
   Wend
   
   ' ��ǹŴ
   Minus_row = Val(txtRow6_1.Text)
   Minus_column = Val(txtColumn6_2.Text)
   ROW = 0
   While (ROW <= 10)
       Minus_row = Minus_row + ROW
       xlsFarmName = m_ExcelSheet.Cells(Val(Minus_row), Minus_column).Value
       
      Set TempRealFarm = GetObject("CXlsCarkill", m_collRealFarm, Trim(xlsFarmName), False)               ' ����� Sum ����Ѻ�Դ��ԧ
        If Not (TempRealFarm Is Nothing) Then
           m_ExcelSheet.Cells(Minus_row, Column + (m_collPeriodDate.Count * 2)).Value = TempRealFarm.sigmaKilo
           m_ExcelSheet.Cells(Minus_row, Column + (m_collPeriodDate.Count * 2) + 1).Value = TempRealFarm.sigmaNetpay
        End If
   
      ROW = ROW + 1
   Wend

   

End Sub

Private Sub GenTotalOutputSheet()
On Error GoTo ErrorHandler
Dim i As Long
Dim DateCount As Long
Dim j As Long
Dim iCount As Long
Dim startColumn As Long
Dim DatePrintColumn As Long
Dim printColumn As Long
Dim RowSetting As Long
Dim GroupPeriodDate As CXlsCarkill
Dim TempPeriodDate As CXlsCarkill
Dim ShortName As String
Dim DateRow As Long
Dim printFinish As Boolean
Dim FirstDate As Date
Dim LastDate As Date
Dim Formula1 As String
Dim Formula2 As String
Dim Formula3 As String
Dim now_column As Long
Dim TempCarkillSum As CXlsCarkillSum
Dim TempCarkillFW As CXlsCarkillFW
Dim txtSum As String
Dim strForward As String
Dim T As Long

   prgProgress.MAX = 100
   prgProgress.MIN = 0
   prgProgress.Value = 0
   Set GroupPeriodDate = New CXlsCarkill

  ' startColumn = Val(txtColumnTotal.Text)
   RowSetting = Val(txtRowTotal.Text)
   printFinish = False
   DatePrintColumn = 0
   DateRow = 2
   
   startColumn = 0
   iCount = Val(txtRow2.Text) - Val(txtRow.Text)              ' ������ҷء��÷Ѵ�� 100%
   Set m_ExcelSheet = m_ExcelApp.Sheets(Val(txtSheet7.Text))
  
   '--- �ٵ� �������¡�    ,,, ǹ collumn
  For Each TempCarkillFW In col_XlsFW
     If TempCarkillFW.MAIN_FLAG = "Y" Then
         strForward = TempCarkillFW.FW_ROW
     End If
     DatePrintColumn = 0
     
     For Each TempPeriodDate In m_collPeriodDate
'        now_column = Val(txtColumn7.Text) + DatePrintColumn
'        txtSum = "=" & TempCarkillSum.OPERATOR_1 & number2Column(now_column - Val(TempCarkillSum.P_COLUMN_1)) & TempCarkillSum.P_ROW_1
        Formula3 = ""
        If TempCarkillFW.UPPER_FLAG = "Y" Then
          Formula1 = number2Column(Val(txtColumn7.Text) + DatePrintColumn) & Trim(Str(TempCarkillFW.FW_ROW - 2))
          Formula2 = number2Column(Val(txtColumn7.Text) + DatePrintColumn) & Trim(Str(TempCarkillFW.FW_ROW - 1))
          Formula3 = "+" & Formula1 & "-" & Formula2
        End If
          Formula3 = number2Column(Val(txtColumn7.Text) + DatePrintColumn - 1) & Trim(Str(TempCarkillFW.FW_ROW)) & Formula3
          m_ExcelSheet.Cells(TempCarkillFW.FW_ROW, Val(txtColumn7.Text) + DatePrintColumn).Value = "=" & Formula3
        
        DatePrintColumn = DatePrintColumn + 1
     Next TempPeriodDate

   ' sum1
   Formula3 = ""
   If TempCarkillFW.UPPER_FLAG = "Y" Then
      Formula1 = number2Column(Val(txtColumn7.Text) + DatePrintColumn) & Trim(Str(TempCarkillFW.FW_ROW - 2))
      Formula2 = number2Column(Val(txtColumn7.Text) + DatePrintColumn) & Trim(Str(TempCarkillFW.FW_ROW - 1))
      Formula3 = "+" & Formula1 & "-" & Formula2
   End If
     Formula3 = number2Column(Val(txtColumn7.Text) + DatePrintColumn - m_collPeriodDate.Count - 1) & Trim(Str(TempCarkillFW.FW_ROW)) & Formula3
     m_ExcelSheet.Cells(TempCarkillFW.FW_ROW, Val(txtColumn7.Text) + DatePrintColumn).Value = "=" & Formula3

    ' sum2
     m_ExcelSheet.Cells(TempCarkillFW.FW_ROW, Val(txtColumn7.Text) + DatePrintColumn + 1).Value = "=" & number2Column(Val(txtColumn7.Text) + DatePrintColumn) & Trim(TempCarkillFW.FW_ROW)

  Next TempCarkillFW
  '--- �� �������¡�
  
  
  DatePrintColumn = 0
  For Each tempSheet In sheetHaveData
      For Each TempPeriodDate In m_collPeriodDate
           If printFinish = False Then
               Set m_ExcelSheet = m_ExcelApp.Sheets(Val(txtSheet7.Text))                  ' ��¹�ѹ����˹�� print
               If DatePrintColumn = 0 Then
                  m_ExcelSheet.Cells(DateRow, Val(txtColumn7.Text)).Value = "�ѹ���ú��˹�����"
                  FirstDate = Left(DateToStringExtEx2(TempPeriodDate.FromDate), 2)
               End If
               
               ' �ç�����֧ �ѹ���  ��᷹
               m_ExcelSheet.Cells(DateRow + 1, Val(txtColumn7.Text) + DatePrintColumn).Value = Left(DateToStringExtEx2(TempPeriodDate.FromDate), 2) & "-" & DateToStringExtEx2(TempPeriodDate.ToDate)
               
               ' �������¡�� Row 4
               m_ExcelSheet.Cells(DateRow + 2, Val(txtColumn7.Text) + DatePrintColumn).Value = "=" & number2Column(Val(txtColumn7.Text) + DatePrintColumn - 1) & strForward
               
               '�͹ sum ��÷Ѵ 26
               'm_ExcelSheet.Cells(26, Val(txtColumn7.Text) + DatePrintColumn).Value = total26
               
               LastDate = DateToStringExtEx2(TempPeriodDate.ToDate)
               DatePrintColumn = DatePrintColumn + 1
            End If
      Next TempPeriodDate
     
         If printFinish = False Then
             ' �������¡�� Row 4 ,Sum1 ,, Sum2
             m_ExcelSheet.Cells(DateRow + 2, Val(txtColumn7.Text) + DatePrintColumn).Value = "=" & number2Column(Val(txtColumn7.Text)) & Trim(Str(DateRow + 2))
             m_ExcelSheet.Cells(DateRow + 2, Val(txtColumn7.Text) + DatePrintColumn + 1).Value = "=" & number2Column(Val(txtColumn7.Text) + DatePrintColumn) & strForward
             m_ExcelSheet.Cells(DateRow, Val(txtColumn7.Text) + DatePrintColumn + 1).Value = "˹�餧�����"
             m_ExcelSheet.Cells(DateRow + 1, Val(txtColumn7.Text) + DatePrintColumn).Value = "��� " & uctlFromDate.ShowDate & "-" & LastDate   ' uctlFromDate.ShowDate
             m_ExcelSheet.Cells(DateRow + 1, Val(txtColumn7.Text) + DatePrintColumn + 1).Value = "��� " & uctlFromDate.ShowDate & "-" & LastDate   '
             printFinish = True                                                                                                      ' �ú�ͺ�á������ print �ա
         End If
       '  startColumn = Val(txtColumnTotal.Text)
  Next tempSheet
  
      startColumn = 0
 '     iCount = Val(txtRow2.Text) - Val(txtRow.Text)              ' ������ҷء��÷Ѵ�� 100%
'      Set m_ExcelSheet = m_ExcelApp.Sheets(Val(txtSheet7.Text))

      For Each TempPeriodDate In m_collPeriodDate
        If Not (TempPeriodDate.m_DistanceFarm Is Nothing) Then
          For Each tempCarkill In TempPeriodDate.m_DistanceFarm
               While (j < iCount)                                                                               ' ǹ��÷Ѵ
                  j = j + 1
                  ShortName = m_ExcelSheet.Cells(Val(txtRow7.Text) + j - 1, 1).Value                      ' �֧���ͨҡ��ͧ�á�ͧ 4
                  Set GroupPeriodDate = GetObject("CXlsCarkill", TempPeriodDate.m_DistanceFarm, ShortName, False)
                  If Not (GroupPeriodDate Is Nothing) Then
                     If GroupPeriodDate.sumFlag = "N" Then
                     
                        ' ��÷Ѵ ����
                        m_ExcelSheet.Cells(Val(txtRow7.Text) + j - 1, Val(txtColumn7.Text) + GroupPeriodDate.DateIndex - 1).Value = GroupPeriodDate.SumNetpay
                     
                        ' ��÷Ѵ �ٵ�ź����
                        Formula1 = number2Column(Val(txtColumn7.Text) + GroupPeriodDate.DateIndex - 1) & Trim(Str(Val(txtRow7.Text) + j - 1))
                        Formula2 = number2Column(Val(txtColumn7.Text) + GroupPeriodDate.DateIndex - 1) & Trim(Str(Val(txtRow7.Text) + j))
                        Formula3 = number2Column(Val(txtColumn7.Text) + GroupPeriodDate.DateIndex - 2) & Trim(Str(Val(txtRow7.Text) + j + 1))
                        m_ExcelSheet.Cells(Val(txtRow7.Text) + j + 1, Val(txtColumn7.Text) + GroupPeriodDate.DateIndex - 1).Value = "=" & Formula3 & "+" & Formula1 & "-" & Formula2

                        Set tempShortname = GetObject("CXlsCarkill", shortnameHaveData, ShortName, False)
                        If Not (tempShortname Is Nothing) Then
                           tempShortname.SumNetpay = tempShortname.SumNetpay + GroupPeriodDate.SumNetpay
                           GroupPeriodDate.sumFlag = "Y"
                           Set tempShortname = Nothing
                           Set GroupPeriodDate = Nothing
                        End If
                     End If
                  End If
               Wend
          Next tempCarkill
          printColumn = printColumn + 1                               ' ��ͧ����¹��ǧ�ѹ����͹
          j = 0
        End If
      Next TempPeriodDate

   ' ��ͧ sum 1 = ��ͧ sum 2
         j = 0
        While (j < iCount)                                                                               ' ǹ��÷Ѵ
            j = j + 1
            ShortName = m_ExcelSheet.Cells(Val(txtRow7.Text) + j - 1, 1).Value                      ' �֧���ͨҡ��ͧ�á
            Set tempShortname = GetObject("CXlsCarkill", shortnameHaveData, ShortName, False)
            If Not (tempShortname Is Nothing) Then
'               If ShortName = "QMC" Then
'                  'debug.print ShortName
'               End If
               
               ' ��÷Ѵ㹪շ7 ��ͧ���
               m_ExcelSheet.Cells(Val(txtRow7.Text) + j - 1, Val(txtColumn7.Text) + m_collPeriodDate.Count).Value = tempShortname.SumNetpay
               
               ' ��÷Ѵ �ٵ�ź����
               Formula1 = number2Column(Val(txtColumn7.Text) + m_collPeriodDate.Count) & Trim(Str(Val(txtRow7.Text) + j - 1))
               Formula2 = number2Column(Val(txtColumn7.Text) + m_collPeriodDate.Count) & Trim(Str(Val(txtRow7.Text) + j))
               Formula3 = number2Column(Val(txtColumn7.Text) - 1) & Trim(Str(Val(txtRow7.Text) + j + 1))
               m_ExcelSheet.Cells(Val(txtRow7.Text) + j + 1, Val(txtColumn7.Text) + m_collPeriodDate.Count).Value = "=" & Formula3 & "+" & Formula1 & "-" & Formula2

               ' ��÷Ѵ㹪շ7 ��ͧ���2
               ' m_ExcelSheet.Cells(Val(txtRow7.Text) + j - 1, Val(txtColumn7.Text) + m_collPeriodDate.Count + 1).Value = tempShortname.SumNetpay
               
              ' ��÷Ѵ �ٵ�ź����
              ' Formula1 = number2Column(Val(txtColumn7.Text) + m_collPeriodDate.Count + 1) & Trim(Str(Val(txtRow7.Text) + j - 1))
              ' Formula2 = number2Column(Val(txtColumn7.Text) + m_collPeriodDate.Count + 1) & Trim(Str(Val(txtRow7.Text) + j))
               Formula3 = number2Column(Val(txtColumn7.Text) + m_collPeriodDate.Count) & Trim(Str(Val(txtRow7.Text) + j + 1))
               m_ExcelSheet.Cells(Val(txtRow7.Text) + j + 1, Val(txtColumn7.Text) + m_collPeriodDate.Count + 1).Value = "=" & Formula3             '& "+" & Formula1 & "-" & Formula2
                                       
            End If
         Wend

  '--- �ٵ����� 5 �ٵ�    ,,, ǹ collumn
  For Each TempCarkillSum In col_XlsSum
     DatePrintColumn = 1
     
     For Each TempPeriodDate In m_collPeriodDate
        now_column = Val(txtColumn7.Text) + DatePrintColumn - 1
        txtSum = "=" & TempCarkillSum.OPERATOR_1 & number2Column(now_column - Val(TempCarkillSum.P_COLUMN_1)) & TempCarkillSum.P_ROW_1
        If TempCarkillSum.P_ROW_2 <> "" Then
         txtSum = txtSum & TempCarkillSum.OPERATOR_2 & number2Column(now_column - Val(TempCarkillSum.P_COLUMN_2)) & TempCarkillSum.P_ROW_2
        End If
        If TempCarkillSum.P_ROW_3 <> "" Then
         txtSum = txtSum & TempCarkillSum.OPERATOR_3 & number2Column(now_column - Val(TempCarkillSum.P_COLUMN_3)) & TempCarkillSum.P_ROW_3
        End If
        If TempCarkillSum.P_ROW_4 <> "" Then
         txtSum = txtSum & TempCarkillSum.OPERATOR_4 & number2Column(now_column - Val(TempCarkillSum.P_COLUMN_4)) & TempCarkillSum.P_ROW_4
        End If
        If TempCarkillSum.P_ROW_5 <> "" Then
         txtSum = txtSum & TempCarkillSum.OPERATOR_5 & number2Column(now_column - Val(TempCarkillSum.P_COLUMN_5)) & TempCarkillSum.P_ROW_5
        End If
'
'        If TempCarkillSum.SUM_ROW = "40" Then
'      '   'debug.print
'        End If
        m_ExcelSheet.Cells(TempCarkillSum.SUM_ROW, now_column).Value = txtSum
                                                                                                                                                                                                  
        DatePrintColumn = DatePrintColumn + 1
     Next TempPeriodDate
     
     '' ����͹��� ��������ش����㹪ش
       txtSum = "=" & TempCarkillSum.OPERATOR_1 & number2Column(now_column - Val(TempCarkillSum.P_COLUMN_1)) & TempCarkillSum.P_ROW_1
        If TempCarkillSum.P_ROW_2 <> "" Then
         txtSum = txtSum & TempCarkillSum.OPERATOR_2 & number2Column(now_column - Val(TempCarkillSum.P_COLUMN_2)) & TempCarkillSum.P_ROW_2
        End If
        If TempCarkillSum.P_ROW_3 <> "" Then
         txtSum = txtSum & TempCarkillSum.OPERATOR_3 & number2Column(now_column - Val(TempCarkillSum.P_COLUMN_3)) & TempCarkillSum.P_ROW_3
        End If
        If TempCarkillSum.P_ROW_4 <> "" Then
         txtSum = txtSum & TempCarkillSum.OPERATOR_4 & number2Column(now_column - Val(TempCarkillSum.P_COLUMN_4)) & TempCarkillSum.P_ROW_4
        End If
        If TempCarkillSum.P_ROW_5 <> "" Then
         txtSum = txtSum & TempCarkillSum.OPERATOR_5 & number2Column(now_column - Val(TempCarkillSum.P_COLUMN_5)) & TempCarkillSum.P_ROW_5
        End If
        m_ExcelSheet.Cells(TempCarkillSum.SUM_ROW, now_column).Value = txtSum
        
        ' ___ �� sum1 �� Y (�ǹ͹)
        If TempCarkillSum.HORIZONTAL_FLAG = "Y" Then
           'sum1
           txtSum = "="
           For T = 0 To (m_collPeriodDate.Count - 1)
               txtSum = txtSum & "+" & number2Column(now_column - T) & TempCarkillSum.SUM_ROW
           Next T
           m_ExcelSheet.Cells(TempCarkillSum.SUM_ROW, now_column + 1).Value = txtSum
           
           m_ExcelSheet.Cells(TempCarkillSum.SUM_ROW, now_column + 2).Value = ""         'Sum2
        Else
            ' sum1
            now_column = now_column + 1
            txtSum = "=" & TempCarkillSum.OPERATOR_1 & number2Column(now_column - Val(TempCarkillSum.P_COLUMN_1)) & TempCarkillSum.P_ROW_1
            If TempCarkillSum.P_ROW_2 <> "" Then
             txtSum = txtSum & TempCarkillSum.OPERATOR_2 & number2Column(now_column - Val(TempCarkillSum.P_COLUMN_2)) & TempCarkillSum.P_ROW_2
            End If
            If TempCarkillSum.P_ROW_3 <> "" Then
             txtSum = txtSum & TempCarkillSum.OPERATOR_3 & number2Column(now_column - Val(TempCarkillSum.P_COLUMN_3)) & TempCarkillSum.P_ROW_3
            End If
            If TempCarkillSum.P_ROW_4 <> "" Then
             txtSum = txtSum & TempCarkillSum.OPERATOR_4 & number2Column(now_column - Val(TempCarkillSum.P_COLUMN_4)) & TempCarkillSum.P_ROW_4
            End If
            If TempCarkillSum.P_ROW_5 <> "" Then
             txtSum = txtSum & TempCarkillSum.OPERATOR_5 & number2Column(now_column - Val(TempCarkillSum.P_COLUMN_5)) & TempCarkillSum.P_ROW_5
            End If
            m_ExcelSheet.Cells(TempCarkillSum.SUM_ROW, now_column).Value = txtSum
            
            'sum2
            now_column = now_column + 1
            txtSum = "=" & TempCarkillSum.OPERATOR_1 & number2Column(now_column - Val(TempCarkillSum.P_COLUMN_1)) & TempCarkillSum.P_ROW_1
            If TempCarkillSum.P_ROW_2 <> "" Then
             txtSum = txtSum & TempCarkillSum.OPERATOR_2 & number2Column(now_column - Val(TempCarkillSum.P_COLUMN_2)) & TempCarkillSum.P_ROW_2
            End If
            If TempCarkillSum.P_ROW_3 <> "" Then
             txtSum = txtSum & TempCarkillSum.OPERATOR_3 & number2Column(now_column - Val(TempCarkillSum.P_COLUMN_3)) & TempCarkillSum.P_ROW_3
            End If
            If TempCarkillSum.P_ROW_4 <> "" Then
             txtSum = txtSum & TempCarkillSum.OPERATOR_4 & number2Column(now_column - Val(TempCarkillSum.P_COLUMN_4)) & TempCarkillSum.P_ROW_4
            End If
            If TempCarkillSum.P_ROW_5 <> "" Then
             txtSum = txtSum & TempCarkillSum.OPERATOR_5 & number2Column(now_column - Val(TempCarkillSum.P_COLUMN_5)) & TempCarkillSum.P_ROW_5
            End If
            m_ExcelSheet.Cells(TempCarkillSum.SUM_ROW, now_column).Value = txtSum
         
         End If
         
  Next TempCarkillSum
  '--- �� �ٵ�����

      prgProgress.Value = MyDiff(j, iCount) * 100                ' �������� �����ǹ=0 ����� 0 ������ٻẺ double
      txtPercent.Text = prgProgress.Value
      Me.Refresh

   prgProgress.Value = 100
   txtPercent.Text = 100
   Exit Sub

ErrorHandler:
   glbErrorLog.LocalErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowUserError
End Sub


Private Sub GenTotalinSheet()
On Error GoTo ErrorHandler
Dim i As Long
Dim DateCount As Long
Dim j As Long
Dim iCount As Long
Dim startColumn As Long
Dim DatePrintColumn As Long
Dim printColumn As Long
Dim RowSetting As Long
Dim GroupPeriodDate As CXlsCarkill
Dim TempPeriodDate As CXlsCarkill
Dim ShortName As String
Dim DateRow As Long
Dim printFinish As Boolean
Dim FirstDate As Date
Dim LastDate As Date
   
   prgProgress.MAX = 100
   prgProgress.MIN = 0
   prgProgress.Value = 0
   Set GroupPeriodDate = New CXlsCarkill

  ' 'debug.print m_collPeriodDate.Count
   startColumn = Val(txtColumnTotal.Text)
   RowSetting = Val(txtRowTotal.Text)
   printFinish = False
   DatePrintColumn = 0
   DateRow = 2
   
   For Each tempSheet In sheetHaveData
         Set m_ExcelSheet = m_ExcelApp.Sheets(tempSheet.SheetIndex)
         m_ExcelSheet.Cells(RowSetting + 2, startColumn).Value = "=" & txtColumn8.Text & txtRow8.Text                              ' tempCarkill.SheetName
         m_ExcelSheet.Cells(RowSetting, startColumn).Value = "�ѹ���"
         m_ExcelSheet.Cells(RowSetting + 1, startColumn).Value = "�.�. / �ҷ "
   Next tempSheet
   
  For Each tempSheet In sheetHaveData
      For Each TempPeriodDate In m_collPeriodDate
            Set m_ExcelSheet = m_ExcelApp.Sheets(tempSheet.SheetIndex)
            m_ExcelSheet.Cells(RowSetting, startColumn + 1).Value = DateToStringExtEx2(TempPeriodDate.FromDate) & " �֧ " & DateToStringExtEx2(TempPeriodDate.ToDate)
            m_ExcelSheet.Cells(RowSetting + 1, startColumn + 1).Value = "�.�. "
            m_ExcelSheet.Cells(RowSetting + 1, startColumn + 2).Value = "�ӹǹ�Թ"
            startColumn = startColumn + 2                                                            ' ��ͧ����¹��ǧ�ѹ����͹
''''           If printFinish = False Then
''''               Set m_ExcelSheet = m_ExcelApp.Sheets(Val(txtSheet7.Text))                  ' ��¹�ѹ����˹�� print
''''               If DatePrintColumn = 0 Then
''''                  m_ExcelSheet.Cells(DateRow, Val(txtColumn7.Text)).Value = "�ѹ���ú��˹�����"
''''                  FirstDate = Left(DateToStringExtEx2(TempPeriodDate.FromDate), 2)
''''               End If
''''               m_ExcelSheet.Cells(DateRow + 1, Val(txtColumn7.Text) + DatePrintColumn).Value = Left(DateToStringExtEx2(TempPeriodDate.FromDate), 2) & "-" & DateToStringExtEx2(TempPeriodDate.ToDate)
''''               LastDate = DateToStringExtEx2(TempPeriodDate.ToDate)
''''               DatePrintColumn = DatePrintColumn + 1
''''            End If
      Next TempPeriodDate
     
''''         If printFinish = False Then
''''         '    Set m_ExcelSheet = m_ExcelApp.Sheets(Val(txtSheet7.Text))                                         ' ��¹�ѹ����˹�� print
''''             m_ExcelSheet.Cells(DateRow, Val(txtColumn7.Text) + DatePrintColumn + 1).Value = "˹�餧�����"
''''             m_ExcelSheet.Cells(DateRow + 1, Val(txtColumn7.Text) + DatePrintColumn).Value = "��� " & FirstDate & "-" & LastDate
''''             m_ExcelSheet.Cells(DateRow + 1, Val(txtColumn7.Text) + DatePrintColumn + 1).Value = "��� " & FirstDate & "-" & LastDate
''''             printFinish = True                                                                                                      ' �ú�ͺ�á������ print �ա
''''         End If
         startColumn = Val(txtColumnTotal.Text)
  Next tempSheet
  
    For Each TempPeriodDate In m_collPeriodDate
         If Not (TempPeriodDate.m_Farm Is Nothing) Then
           For Each tempCarkill In TempPeriodDate.m_Farm
               Set m_ExcelSheet = m_ExcelApp.Sheets(tempCarkill.SheetIndex)
               m_ExcelSheet.Cells(RowSetting + 2, startColumn + 1).Value = tempCarkill.SumKilo
               m_ExcelSheet.Cells(RowSetting + 2, startColumn + 2).Value = tempCarkill.SumNetpay
           Next tempCarkill
           startColumn = startColumn + 2                               ' ��ͧ����¹��ǧ�ѹ����͹
         End If
    Next TempPeriodDate
   
   '''''' ˹�� print �ش��� 1
   startColumn = 0
'   For Each tempSheet In sheetHaveData                                                     ' ǹ�ѹ���
'       For Each TempPeriodDate In m_collPeriodDate
'
'         Set m_ExcelSheet = m_ExcelApp.Sheets(Val(txtSheet7.Text))
'         ShortName = m_ExcelSheet.Cells(Val(txtRow7.Text) + j - 1, 1).Value
'
'         m_ExcelSheet.Cells(DateRow, Val(txtColumn7.Text) + startColumn).Value = DateToStringExtEx2(TempPeriodDate.FromDate) & " �֧ " & DateToStringExtEx2(TempPeriodDate.ToDate)          ' �ѹ���������
'         startColumn = startColumn + 1                               ' ��ͧ����¹��ǧ�ѹ����͹
'
'       Next TempPeriodDate
'       startColumn = Val(txtColumnTotal.Text)
'    Next tempSheet
   
'''''      iCount = Val(txtRow2.Text) - Val(txtRow.Text)              ' ������ҷء��÷Ѵ�� 100%
'''''      Set m_ExcelSheet = m_ExcelApp.Sheets(Val(txtSheet7.Text))
'''''
'''''      For Each TempPeriodDate In m_collPeriodDate
'''''        If Not (TempPeriodDate.m_DistanceFarm Is Nothing) Then
'''''          For Each tempCarkill In TempPeriodDate.m_DistanceFarm
'''''               While (j < iCount)                                                                               ' ǹ��÷Ѵ
'''''                  j = j + 1
'''''                  ShortName = m_ExcelSheet.Cells(Val(txtRow7.Text) + j - 1, 1).Value                      ' �֧���ͨҡ��ͧ�á�ͧ 4
'''''                  ' 'debug.print ShortName
'''''                  Set GroupPeriodDate = GetObject("CXlsCarkill", TempPeriodDate.m_DistanceFarm, ShortName, False)
'''''                  If Not (GroupPeriodDate Is Nothing) Then
'''''                     m_ExcelSheet.Cells(Val(txtRow7.Text) + j - 1, Val(txtColumn7.Text) + GroupPeriodDate.DateIndex - 1).Value = GroupPeriodDate.SumNetpay
'''''                     Set tempShortname = GetObject("CXlsCarkill", shortnameHaveData, ShortName, False)
'''''                     If Not (tempShortname Is Nothing) Then
'''''                        tempShortname.SumNetpay = tempShortname.SumNetpay + GroupPeriodDate.SumNetpay
'''''                        Set tempShortname = Nothing
'''''                     End If
'''''                  End If
'''''               Wend
'''''          Next tempCarkill
'''''          printColumn = printColumn + 1                               ' ��ͧ����¹��ǧ�ѹ����͹
'''''          j = 0
'''''        End If
'''''      Next TempPeriodDate
'''''
'''''   ' ��ͧ sum 1 = ��ͧ sum 2
''''''  For Each tempShortname In shortnameHaveData
'''''         j = 0
'''''        While (j < iCount)                                                                               ' ǹ��÷Ѵ
'''''            j = j + 1
'''''            ShortName = m_ExcelSheet.Cells(Val(txtRow7.Text) + j - 1, 1).Value                      ' �֧���ͨҡ��ͧ�á
'''''            Set tempShortname = GetObject("CXlsCarkill", shortnameHaveData, ShortName, False)
'''''            If Not (tempShortname Is Nothing) Then
'''''               m_ExcelSheet.Cells(Val(txtRow7.Text) + j - 1, Val(txtColumn7.Text) + m_collPeriodDate.Count).Value = tempShortname.SumNetpay
'''''               m_ExcelSheet.Cells(Val(txtRow7.Text) + j - 1, Val(txtColumn7.Text) + m_collPeriodDate.Count + 1).Value = tempShortname.SumNetpay
'''''            End If
'''''         Wend
'''''  'Next tempShortname

      prgProgress.Value = MyDiff(j, iCount) * 100                ' �������� �����ǹ=0 ����� 0 ������ٻẺ double
      txtPercent.Text = prgProgress.Value
      Me.Refresh

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

      Call LoadXlsCarkillSum(Nothing, col_XlsSum)
      Call LoadXlsCarkillFW(Nothing, col_XlsFW)
      
      Call QueryData

'      GridEX1.ItemCount = CountItem(MainCollection)
'      GridEX1.Rebind

   End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
   If Shift = 1 And KeyCode = DUMMY_KEY Then
      glbErrorLog.LocalErrorMsg = Me.Name
      glbErrorLog.ShowUserError
      KeyCode = 0
   ElseIf Shift = 0 And KeyCode = 113 Then
      Call cmdOK_Click
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
   pnlHeader.Caption = "��ػ˹��ͧ��.��������"

   Me.Picture = LoadPicture(glbParameterObj.MainPicture)

   pnlHeader.Font.Name = GLB_FONT
   pnlHeader.Font.Bold = True
   pnlHeader.Font.Size = 19

   Call InitNormalLabel(lblName, "��¡��")
   Call InitNormalLabel(lblProgress, "�����׺˹��")
   Call InitNormalLabel(lblPercent, "����ૹ��")
   Call InitNormalLabel(Label1, "%")
   Call InitNormalLabel(lblFileName, "�������")
   Call InitNormalLabel(lblFileOutName, "������� Output")

   Call InitNormalLabel(lblRunData, "������� Excel")
   Call InitNormalLabel(lblRow2, "�֧��")
   Call InitNormalLabel(lblRow, "�������")
   Call InitNormalLabel(lblSheet, "�շ���")
   Call InitNormalLabel(lblFromDate, "�ҡ�ѹ���")
   Call InitNormalLabel(lblPeriodDate, "��ǧ�ѹ���" & vbCrLf & "(�� 7,7,7,7,3)")
   
   Call InitNormalLabel(lblReal6, "�Դ��ԧ")
   Call InitNormalLabel(lblSheet6, "�շ���")
   Call InitNormalLabel(lblRow6, "�Ƿ��")
   Call InitNormalLabel(lblReal6_0, "��ǹŴ")
   Call InitNormalLabel(lblRow6_1, "�Ƿ��")
   Call InitNormalLabel(lblColumn6_2, "���������")
   
   Call InitNormalLabel(lblColumn6, "���������")
   Call InitNormalLabel(lblRow9, "�Ƿ��")
   Call InitNormalLabel(lblColumn9, "���������")
    
   Call InitNormalLabel(lblPrint7, "��ػ(Print 1-3)")
   Call InitNormalLabel(lblSheet7, "�շ���")
   Call InitNormalLabel(lblRow7, "�Ƿ��")
   Call InitNormalLabel(lblColumn8, "���������")
   Call InitNormalLabel(lblRow8, "�Ƿ��")
   Call InitNormalLabel(lblColumn7, "���������")
   
   Call InitNormalLabel(lblTotalData, "��÷Ѵ���")
   Call InitNormalLabel(lblShortName, "�������")
   Call InitNormalLabel(lblSheetName, "���ͪշ")
   Call InitNormalLabel(lblColumnTotal, "���������")
   Call InitNormalLabel(lblRowTotal, "�Ƿ��")

   Call txtPercent.SetTextLenType(TEXT_FLOAT, glbSetting.MONEY_TYPE)
   txtPercent.Enabled = False
   Call txtPeriodDate.SetTextLenType(TEXT_STRING, glbSetting.DESC_TYPE)
   txtFileName.Enabled = False
   txtFileOutName.Enabled = False
   Call txtRow.SetTextLenType(TEXT_FLOAT, glbSetting.MONEY_TYPE)
   Call txtRow2.SetTextLenType(TEXT_FLOAT, glbSetting.MONEY_TYPE)
   Call txtSheet.SetTextLenType(TEXT_FLOAT, glbSetting.MONEY_TYPE)
   
   Call txtSheet6.SetTextLenType(TEXT_FLOAT, glbSetting.MONEY_TYPE)
   Call txtRow6.SetTextLenType(TEXT_FLOAT, glbSetting.MONEY_TYPE)
   Call txtColumn6.SetTextLenType(TEXT_FLOAT, glbSetting.MONEY_TYPE)
   Call txtRow6_1.SetTextLenType(TEXT_FLOAT, glbSetting.MONEY_TYPE)
   Call txtColumn6_2.SetTextLenType(TEXT_FLOAT, glbSetting.MONEY_TYPE)
   Call txtSheet7.SetTextLenType(TEXT_FLOAT, glbSetting.MONEY_TYPE)
   Call txtRow7.SetTextLenType(TEXT_FLOAT, glbSetting.MONEY_TYPE)
   Call txtColumn7.SetTextLenType(TEXT_FLOAT, glbSetting.MONEY_TYPE)
   Call txtRow8.SetTextLenType(TEXT_FLOAT, glbSetting.MONEY_TYPE)
   Call txtColumn8.SetTextLenType(TEXT_FLOAT, glbSetting.MONEY_TYPE)
   Call txtRow9.SetTextLenType(TEXT_FLOAT, glbSetting.MONEY_TYPE)
   Call txtColumn9.SetTextLenType(TEXT_FLOAT, glbSetting.MONEY_TYPE)
   Call txtRowTotal.SetTextLenType(TEXT_FLOAT, glbSetting.MONEY_TYPE)
   Call txtColumnTotal.SetTextLenType(TEXT_FLOAT, glbSetting.MONEY_TYPE)

   cmdExit.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdOK.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdSetting.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdStart.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdFileName.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdFileOutName.Picture = LoadPicture(glbParameterObj.NormalButton1)

   Call InitMainButton(cmdExit, MapText("¡��ԡ (ESC)"))
   Call InitMainButton(cmdOK, MapText("��ŧ (F2)"))
   Call InitMainButton(cmdStart, MapText("�����"))
   Call InitMainButton(cmdFileName, MapText("..."))
   Call InitMainButton(cmdFileOutName, MapText("..."))
   Call InitMainButton(cmdSetting, MapText("��駤��"))
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
Dim L As Long

   Call EnableForm(Me, False)
   m_HasActivate = False
   
   Set m_Rs = New ADODB.Recordset
   Set MainCollection = New Collection
   Set SearchCollection = New Collection
   Set SearchNameCollection = New Collection
   Set tempCollCarkill = New Collection
   Set sheetHaveData = New Collection
   Set shortnameHaveData = New Collection
   Set col_XlsSum = New Collection
   Set col_XlsFW = New Collection
   Set m_collRealFarm = New Collection
   Set TempRealFarm = New CXlsCarkill
   Set TempRealFarm2 = New CXlsCarkill

   For L = 1 To UBound(m_collCarkill)
      Set m_collCarkill(L) = New Collection
   Next L
   Set tempCarkill = New CXlsCarkill
   Set tempSheet = New CXlsCarkill
   Set tempShortname = New CXlsCarkill

   Set m_ExcelApp = CreateObject("Excel.application")

   m_HasActivate = False
   Call InitFormLayout
   Call EnableForm(Me, True)
End Sub

Private Sub Form_Unload(Cancel As Integer)
Dim L As Long
   Set MainCollection = Nothing
   Set SearchCollection = Nothing
   Set SearchNameCollection = Nothing
   Set tempCollCarkill = Nothing
   Set sheetHaveData = Nothing
   Set shortnameHaveData = Nothing
   Set col_XlsSum = Nothing
   Set col_XlsFW = Nothing
   Set m_collRealFarm = Nothing
   Set TempRealFarm = Nothing
   Set TempRealFarm2 = Nothing
   
   For L = 1 To UBound(m_collCarkill)
      Set m_collCarkill(L) = Nothing
   Next L
   
   Set tempCarkill = Nothing
   Set tempSheet = Nothing
   Set tempShortname = Nothing

   Call m_ExcelApp.Workbooks.Close
 '  Call m_ExcelApp.Close
End Sub

Private Sub QueryData()
Dim m_EstSetting As CXlsEstimateSetting
Dim IsOK As Boolean
Dim ItemCount As Long
Dim Rs As ADODB.Recordset

   Set m_EstSetting = New CXlsEstimateSetting
   Set Rs = New ADODB.Recordset
   
      m_EstSetting.XLS_EST_SET_ID = 2       ' �� ID = 2 ��·�����
      m_EstSetting.QueryFlag = 1
      If Not glbDaily.QueryXlsSetting(m_EstSetting, Rs, ItemCount, IsOK, glbErrorLog) Then
         glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
         Call EnableForm(Me, True)
         Exit Sub
      End If
   
   If ItemCount > 0 Then
      Call m_EstSetting.PopulateFromRS(1, Rs)
      ShowMode = SHOW_EDIT
      
      txtSheet.Text = m_EstSetting.SHEET_4
      txtRow.Text = m_EstSetting.ROW_4
      txtRow2.Text = m_EstSetting.ROW_5
      
      txtSheet6.Text = m_EstSetting.SHEET_1
      txtSheet7.Text = m_EstSetting.SHEET_2
      txtColumn6.Text = m_EstSetting.COLLUMN_1
      txtColumn7.Text = m_EstSetting.COLLUMN_2
      txtColumnTotal.Text = m_EstSetting.COLLUMN_3
      txtRow6.Text = m_EstSetting.ROW_1
      txtRow7.Text = m_EstSetting.ROW_2
      txtRowTotal.Text = m_EstSetting.ROW_3
      txtColumn8.Text = m_EstSetting.COLLUMN_6
      txtRow8.Text = m_EstSetting.ROW_6
      txtColumn9.Text = m_EstSetting.COLLUMN_7
      txtRow9.Text = m_EstSetting.ROW_7
      txtColumn6_2.Text = m_EstSetting.COLLUMN6_2
      txtRow6_1.Text = m_EstSetting.COLLUMN6_1
   Else
      ShowMode = SHOW_ADD
   End If

   If Rs.State = adStateOpen Then
      Rs.Close
   End If
   Set Rs = Nothing
End Sub
Private Function SaveData() As Boolean
Dim IsOK As Boolean
Dim m_EstSetting As CXlsEstimateSetting
Set m_EstSetting = New CXlsEstimateSetting

      m_EstSetting.AddEditMode = ShowMode
      
      m_EstSetting.XLS_EST_SET_ID = 2
      m_EstSetting.SHEET_4 = txtSheet.Text
      m_EstSetting.ROW_4 = txtRow.Text
      m_EstSetting.ROW_5 = txtRow2.Text
      
      m_EstSetting.SHEET_1 = txtSheet6.Text
      m_EstSetting.SHEET_2 = txtSheet7.Text
      m_EstSetting.COLLUMN_1 = txtColumn6.Text
      m_EstSetting.COLLUMN_2 = txtColumn7.Text
      m_EstSetting.COLLUMN_3 = txtColumnTotal.Text
      m_EstSetting.ROW_1 = txtRow6.Text
      m_EstSetting.ROW_2 = txtRow7.Text
      m_EstSetting.ROW_3 = txtRowTotal.Text
      m_EstSetting.COLLUMN_6 = txtColumn8.Text
      m_EstSetting.ROW_6 = txtRow8.Text
      m_EstSetting.COLLUMN_7 = txtColumn9.Text
      m_EstSetting.ROW_7 = txtRow9.Text
      m_EstSetting.COLLUMN6_2 = txtColumn6_2.Text
      m_EstSetting.COLLUMN6_1 = txtRow6_1.Text
      
  Call EnableForm(Me, False)
   
   If Not glbDaily.AddEditXlsSetting(m_EstSetting, IsOK, True, glbErrorLog) Then
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

Private Function StringSlashToDate(DateStr As String) As Date
Dim TempDay As Long
Dim TempMonth As Long
Dim TempYear As Long

On Error Resume Next
   'DD/M/YY
   
   Dim TempID As Long
   TempID = InStr(1, DateStr, "/")
   TempDay = CLng(Val(Left(DateStr, TempID - 1)))             ' �����Ţ���˹�� ��觫������
   DateStr = Mid(DateStr, TempID + 1)                               ' ����Ţ����á �Ѻ , �١�Ѵ����� = ����ͷ������ͷҧ���

     TempID = InStr(1, DateStr, "/")
     TempMonth = CLng(Val(Left(DateStr, TempID - 1)))        ' �����Ţ���˹�� ��觫������
     TempYear = CLng(Mid(DateStr, TempID + 1))                 ' ����Ţ����á �Ѻ , �١�Ѵ����� = ����ͷ������ͷҧ���
   
   If TempYear >= 2500 Then
      TempYear = TempYear - 543
   End If
   
   StringSlashToDate = DateSerial(TempYear, TempMonth, TempDay)
End Function

Private Sub CalculateDate(TempStr As String, DateCount As Long)
Dim TempID As Long
Dim TempStrNew  As String
   TempStrNew = Replace(TempStr, "(", "")
   TempStrNew = Replace(TempStrNew, ")", "")
   TempID = InStr(1, TempStrNew, ",")
   DateCount = 0
   While InStr(1, TempStrNew, ",") > 0
      TempID = InStr(1, TempStrNew, ",")
      DateCount = DateCount + Val(Left(TempStrNew, TempID - 1))   ' 7+7+7+7+3 =31
      TempStrNew = Mid(TempStrNew, TempID + 1)                        ' �Ѵ�ѹ��� ����������
   Wend
   DateCount = DateCount + Val(TempStrNew)
End Sub

Private Sub CalculateDatePeriod(TempStr As String, DateCount As Long)
Dim TempID As Long
   TempID = InStr(1, TempStr, ",")
   If TempID > 0 Then
      DateCount = Val(Left(TempStr, TempID - 1))        ' �����Ţ���˹�� ��觫������
      TempStr = Mid(TempStr, TempID + 1)                 ' ����Ţ����á �Ѻ , �١�Ѵ����� = ����ͷ������ͷҧ���
   Else
      DateCount = Val(TempStr)                                 ' ��� ��� , ���� ��ҵ���Ţ�������������
      TempStr = ""
   End If
End Sub

Private Function findBlankRow(startRow As Long) As Long
Dim collumn As Long
Dim itBlankRow As Long
Dim flagBlankRow As Boolean

   While itBlankRow < 3                                              ' ��ҧ�ú 3 ������ ����ԡǹ�ͧ
      collumn = 1
      flagBlankRow = True
      For collumn = 1 To 26
         If m_ExcelSheet.Cells(startRow, collumn).Value <> "" Then
            startRow = startRow + 1
            itBlankRow = 0
            flagBlankRow = False
            Exit For
         End If
      Next collumn
      If flagBlankRow = True Then
         itBlankRow = itBlankRow + 1
      End If
   Wend
   findBlankRow = startRow + 10
End Function

Private Function number2Column(Number As Long) As String
Dim numA As Long
Dim numB As Long
Dim i As Long
Dim collumnA As String
Dim collumnB As String
   
   numA = Int(Number / 26)               ' �ӹǹ��ѭ��� Eng
   numB = Int(Number Mod 26)
   If numB = 0 Then
     numA = numA - 1
     numB = 26
   End If
   
   collumnA = number2String(numA)
   collumnB = number2String(numB)
   
   number2Column = collumnA & collumnB
End Function
'Private Function num2text(column As Long, index As Long) As String
'   num2text = number2String(column - index)
'End Function

Private Function number2String(index As Long) As String
  Select Case index
      Case 1
         number2String = "A"
      Case 2
         number2String = "B"
      Case 3
         number2String = "C"
      Case 4
         number2String = "D"
       Case 5
         number2String = "E"
      Case 6
         number2String = "F"
      Case 7
         number2String = "G"
      Case 8
         number2String = "H"
      Case 9
         number2String = "I"
      Case 10
         number2String = "J"
       Case 11
         number2String = "K"
      Case 12
         number2String = "L"
      Case 13
         number2String = "M"
      Case 14
         number2String = "N"
      Case 15
         number2String = "O"
       Case 16
         number2String = "P"
      Case 17
         number2String = "Q"
      Case 18
         number2String = "R"
      Case 19
         number2String = "S"
      Case 20
         number2String = "T"
      Case 21
         number2String = "U"
      Case 22
         number2String = "V"
      Case 23
         number2String = "W"
      Case 24
         number2String = "X"
      Case 25
         number2String = "Y"
      Case 26
         number2String = "Z"
      Case Else
         number2String = ""
   End Select
End Function
