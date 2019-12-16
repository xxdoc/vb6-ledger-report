VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#2.0#0"; "THREED20.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmXlsFoodPay 
   BackColor       =   &H80000000&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   9150
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11940
   Icon            =   "frmXlsFoodPay.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9150
   ScaleWidth      =   11940
   StartUpPosition =   1  'CenterOwner
   Begin Threed.SSFrame SSFrame1 
      Height          =   9165
      Left            =   0
      TabIndex        =   25
      Top             =   0
      Width           =   11985
      _ExtentX        =   21140
      _ExtentY        =   16166
      _Version        =   131073
      PictureBackgroundStyle=   2
      Begin prjLedgerReport.uctlDate uctlFromDate 
         Height          =   375
         Left            =   3000
         TabIndex        =   2
         Top             =   1440
         Width           =   3855
         _ExtentX        =   6800
         _ExtentY        =   661
      End
      Begin MSComctlLib.ProgressBar prgProgress 
         Height          =   315
         Left            =   2400
         TabIndex        =   23
         Top             =   6960
         Width           =   7905
         _ExtentX        =   13944
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   1
      End
      Begin Threed.SSPanel pnlHeader 
         Height          =   705
         Left            =   0
         TabIndex        =   26
         Top             =   0
         Width           =   11955
         _ExtentX        =   21087
         _ExtentY        =   1244
         _Version        =   131073
         PictureBackgroundStyle=   2
      End
      Begin prjLedgerReport.uctlTextBox txtPercent 
         Height          =   465
         Left            =   2400
         TabIndex        =   24
         Top             =   7440
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
         TabIndex        =   0
         Top             =   960
         Width           =   6720
         _ExtentX        =   11853
         _ExtentY        =   767
      End
      Begin prjLedgerReport.uctlTextBox txtRow 
         Height          =   435
         Left            =   5880
         TabIndex        =   6
         Top             =   2880
         Width           =   840
         _ExtentX        =   1905
         _ExtentY        =   767
      End
      Begin prjLedgerReport.uctlTextBox txtSheet 
         Height          =   435
         Left            =   3000
         TabIndex        =   5
         Top             =   2880
         Width           =   960
         _ExtentX        =   1905
         _ExtentY        =   767
      End
      Begin prjLedgerReport.uctlTextBox txtPeriodDate 
         Height          =   435
         Left            =   3000
         TabIndex        =   4
         Top             =   2400
         Width           =   1680
         _ExtentX        =   1058
         _ExtentY        =   767
      End
      Begin prjLedgerReport.uctlTextBox txtColumn7 
         Height          =   435
         Left            =   5880
         TabIndex        =   8
         Top             =   3360
         Width           =   840
         _ExtentX        =   1270
         _ExtentY        =   767
      End
      Begin prjLedgerReport.uctlTextBox txtSheet6 
         Height          =   435
         Left            =   3960
         TabIndex        =   15
         Top             =   5640
         Width           =   600
         _ExtentX        =   1905
         _ExtentY        =   767
      End
      Begin prjLedgerReport.uctlTextBox txtRow6 
         Height          =   435
         Left            =   6000
         TabIndex        =   16
         Top             =   5640
         Width           =   600
         _ExtentX        =   1905
         _ExtentY        =   767
      End
      Begin prjLedgerReport.uctlTextBox txtColumn6 
         Height          =   435
         Left            =   8160
         TabIndex        =   17
         Top             =   5640
         Width           =   600
         _ExtentX        =   1270
         _ExtentY        =   767
      End
      Begin prjLedgerReport.uctlTextBox txtRow7 
         Height          =   435
         Left            =   3960
         TabIndex        =   13
         Top             =   5160
         Width           =   600
         _ExtentX        =   1905
         _ExtentY        =   767
      End
      Begin prjLedgerReport.uctlTextBox txtRow8 
         Height          =   435
         Left            =   8400
         TabIndex        =   7
         Top             =   2880
         Width           =   840
         _ExtentX        =   1905
         _ExtentY        =   767
      End
      Begin prjLedgerReport.uctlTextBox txtColumn9 
         Height          =   435
         Left            =   5880
         TabIndex        =   11
         Top             =   3840
         Width           =   840
         _ExtentX        =   1905
         _ExtentY        =   767
      End
      Begin prjLedgerReport.uctlTextBox txtRow6_1 
         Height          =   435
         Left            =   6000
         TabIndex        =   14
         Top             =   5160
         Width           =   600
         _ExtentX        =   1905
         _ExtentY        =   767
      End
      Begin prjLedgerReport.uctlTextBox txtColumn6_2 
         Height          =   435
         Left            =   8160
         TabIndex        =   18
         Top             =   6120
         Width           =   600
         _ExtentX        =   1270
         _ExtentY        =   767
      End
      Begin prjLedgerReport.uctlDate uctlToDate 
         Height          =   375
         Left            =   3000
         TabIndex        =   3
         Top             =   1920
         Width           =   3855
         _ExtentX        =   6800
         _ExtentY        =   661
      End
      Begin prjLedgerReport.uctlTextBox txtRow9 
         Height          =   435
         Left            =   3840
         TabIndex        =   10
         Top             =   3840
         Width           =   840
         _ExtentX        =   1270
         _ExtentY        =   767
      End
      Begin prjLedgerReport.uctlTextBox txtColumn8 
         Height          =   435
         Left            =   8400
         TabIndex        =   9
         Top             =   3360
         Width           =   840
         _ExtentX        =   1270
         _ExtentY        =   767
      End
      Begin prjLedgerReport.uctlTextBox txtShortName 
         Height          =   435
         Left            =   3000
         TabIndex        =   12
         Top             =   4680
         Width           =   1680
         _ExtentX        =   2963
         _ExtentY        =   767
      End
      Begin VB.Label lblColumn8 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   795
         Left            =   6840
         TabIndex        =   51
         Top             =   3360
         Width           =   1455
      End
      Begin VB.Label lblToDate 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   435
         Left            =   1800
         TabIndex        =   50
         Top             =   1920
         Width           =   975
      End
      Begin VB.Label lblReal6_0 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   435
         Left            =   4800
         TabIndex        =   49
         Top             =   6120
         Width           =   1815
      End
      Begin VB.Label lblColumn6_2 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   435
         Left            =   6840
         TabIndex        =   48
         Top             =   6120
         Width           =   1215
      End
      Begin VB.Label lblRow6_1 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   435
         Left            =   4800
         TabIndex        =   47
         Top             =   5160
         Width           =   1095
      End
      Begin VB.Label lblSheetName 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   435
         Left            =   960
         TabIndex        =   46
         Top             =   3840
         Width           =   1815
      End
      Begin VB.Label lblColumn9 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   435
         Left            =   4680
         TabIndex        =   45
         Top             =   3840
         Width           =   1095
      End
      Begin VB.Label lblRow9 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   435
         Left            =   2880
         TabIndex        =   44
         Top             =   3840
         Width           =   855
      End
      Begin VB.Label lblRow8 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   435
         Left            =   7440
         TabIndex        =   43
         Top             =   2880
         Width           =   855
      End
      Begin VB.Label lblShortName 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   435
         Left            =   960
         TabIndex        =   42
         Top             =   4680
         Width           =   1815
      End
      Begin VB.Label lblRow7 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   435
         Left            =   2880
         TabIndex        =   41
         Top             =   5160
         Width           =   975
      End
      Begin VB.Label lblPrint7 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   435
         Left            =   960
         TabIndex        =   40
         Top             =   5160
         Width           =   1815
      End
      Begin VB.Label lblRow6 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   435
         Left            =   4800
         TabIndex        =   39
         Top             =   5640
         Width           =   1095
      End
      Begin VB.Label lblColumn6 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   435
         Left            =   6840
         TabIndex        =   38
         Top             =   5640
         Width           =   1215
      End
      Begin VB.Label lblSheet6 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   435
         Left            =   3000
         TabIndex        =   37
         Top             =   5640
         Width           =   855
      End
      Begin VB.Label lblReal6 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   435
         Left            =   960
         TabIndex        =   36
         Top             =   5640
         Width           =   1815
      End
      Begin Threed.SSCommand cmdSetting 
         Height          =   525
         Left            =   3240
         TabIndex        =   20
         Top             =   8280
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         ButtonStyle     =   3
      End
      Begin VB.Label lblColumn7 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   435
         Left            =   3720
         TabIndex        =   35
         Top             =   3360
         Width           =   2055
      End
      Begin VB.Label lblPeriodDate 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   375
         Left            =   240
         TabIndex        =   34
         Top             =   2400
         Width           =   2535
      End
      Begin VB.Label lblFromDate 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   435
         Left            =   1800
         TabIndex        =   33
         Top             =   1440
         Width           =   975
      End
      Begin Threed.SSCommand cmdFileName 
         Height          =   435
         Left            =   9840
         TabIndex        =   1
         Top             =   960
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   767
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmXlsFoodPay.frx":27A2
         ButtonStyle     =   3
      End
      Begin VB.Label lblSheet 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   435
         Left            =   480
         TabIndex        =   32
         Top             =   2880
         Width           =   2295
      End
      Begin VB.Label lblRow 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   435
         Left            =   4680
         TabIndex        =   31
         Top             =   2880
         Width           =   1095
      End
      Begin VB.Label lblFileName 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   435
         Left            =   600
         TabIndex        =   30
         Top             =   960
         Width           =   2175
      End
      Begin Threed.SSCommand cmdStart 
         Height          =   525
         Left            =   1620
         TabIndex        =   19
         Top             =   8280
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmXlsFoodPay.frx":2ABC
         ButtonStyle     =   3
      End
      Begin VB.Label Label1 
         Caption         =   "Label1"
         Height          =   345
         Left            =   4200
         TabIndex        =   29
         Top             =   7440
         Width           =   1275
      End
      Begin VB.Label lblProgress 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   435
         Left            =   690
         TabIndex        =   28
         Top             =   6960
         Width           =   1575
      End
      Begin VB.Label lblPercent 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   435
         Left            =   690
         TabIndex        =   27
         Top             =   7440
         Width           =   1575
      End
      Begin Threed.SSCommand cmdExit 
         Height          =   525
         Left            =   9960
         TabIndex        =   22
         Top             =   8220
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   926
         _Version        =   131073
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdOK 
         Height          =   525
         Left            =   8280
         TabIndex        =   21
         Top             =   8220
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         ButtonStyle     =   3
      End
   End
End
Attribute VB_Name = "frmXlsFoodPay"
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

Private m_collCarkill(30) As Collection
Private tempCollCarkill As Collection
Private sheetHaveData As Collection
Private shortnameHaveData As Collection
Dim tempShortname As CXlsCarkill
Private col_XlsSum As Collection
Private col_XlsFW As Collection

Dim FirstDate As Date
Dim LastDate As Date
Dim D As String
Dim tempDateString As Long
Dim TempDate As String
Dim TempFromdate  As Date
Dim TempToDate   As Date
Dim m_collPeriodDate As Collection
Dim TempPeriodDate As CXlsCarkill
Dim m_collRealFarm As Collection
Dim TempRealFarm2 As CXlsCarkill

Dim tempCarkill As CXlsCarkill
Dim tempSheet As CXlsCarkill
Dim TempRealFarm As CXlsCarkill

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

 '  txtFileOutName.Text = dlgAdd.FileName
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
   lMenuChosen = oMenu.Popup("ตั้งค่าบรรทัดรวมอื่นๆ", "ตั้งค่าบรรทัดคงเหลือยกไป")
   If lMenuChosen = 0 Then
      Exit Sub
   End If
   Set oMenu = Nothing

   Call EnableForm(Me, False)
   If lMenuChosen = 1 Then
      ' เปิดฟอร์มอื่น
      'frmCarkillSettingSum.ShowMode = SHOW_EDIT
      Load frmCarkillSettingSum
      frmCarkillSettingSum.Show 1

      Unload frmCarkillSettingSum
      Set frmCarkillSettingSum = Nothing
   End If

   Call EnableForm(Me, True)
    If lMenuChosen = 2 Then
      ' เปิดฟอร์มอื่น
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
      Call MsgBox("กรุณากรอกข้อมูล ชีดให้ถูกต้องโดยไม่สามารถมากกว่า  " & MaxSheet, vbOKOnly, PROJECT_NAME)
      m_ExcelApp.Workbooks.Close
      Call EnableForm(Me, True)
      Exit Sub
   End If

   Call SaveData
   
' หาชุดวันที่
   Set tempCarkill = New CXlsCarkill
   Set m_collPeriodDate = New Collection
   tempDateString = txtPeriodDate.Text
   TempFromdate = uctlFromDate.ShowDate
   i = 0
   
   While TempToDate < uctlToDate.ShowDate
      i = i + 1
      TempToDate = DateAdd("D", tempDateString - 1, TempFromdate)
      
      tempCarkill.FromDate = TempFromdate
      If TempFromdate >= uctlToDate.ShowDate And TempToDate >= uctlToDate.ShowDate Then
         tempCarkill.ToDate = uctlToDate.ShowDate
      Else
          tempCarkill.ToDate = TempToDate
      End If
      tempCarkill.DateIndex = i
      
      Call m_collPeriodDate.Add(tempCarkill, Str(i))                                 'Trim(tempCarkill.InvNo) & "-" & Trim(tempCarkill.FeedNo)
      Set tempCarkill = Nothing
      Set tempCarkill = New CXlsCarkill
      
      TempFromdate = DateAdd("D", 1, TempToDate)
   Wend

   ' -------- จบ หาชุดวันที่

   ' Call CalculateDate(txtPeriodDate.Text, DateCount)         ' 31 วัน
'   FirstDate = uctlFromDate.ShowDate
'   LastDate = DateAdd("D", DateCount - 1, uctlFromDate.ShowDate)
   
   For RunSheet = Val(txtSheet.Text) To MaxSheet                  ' วนตั้งแต่ชีทที่เริ่ม จนถึง ชีทสุดท้าย
         Call GenCarkillRow(RunSheet)
   Next RunSheet

        'debug.print m_collPeriodDate.Count

   Call GenTotalInputSheet

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
   
   j = 0
   Set m_ExcelSheet = m_ExcelApp.Sheets(RunSheet)
   iCount = Val(txtRow8.Text) - Val(txtRow.Text)                ' เพราะเอาทุกบรรทัดเป็น 100%
   While (j < iCount)                                                      'And m_ExcelSheet.Cells(Val(txtRow.Text) + j - 1, "B").Value >= uctlFromDate.ShowDate
      collumn = 1
      j = j + 1
      
      If m_ExcelSheet.Cells(Val(txtRow.Text) + j - 1, txtColumn7.Text).Value <> "" Then
          tempShipDate = StringSlashToDate(m_ExcelSheet.Cells(Val(txtRow.Text) + j - 1, txtColumn7.Text).Value)
      Else
          tempShipDate = PrevKey5
      End If
   
      If tempShipDate <> PrevKey5 Then
      For Each tempCarkill In m_collPeriodDate
         If tempCarkill.FromDate <= tempShipDate And tempShipDate <= tempCarkill.ToDate Then         ' ดึงวันที่เพื่อมาเช็คบรรทัดนั้นเลย
           collumn = 0
           Set tempSheet = New CXlsCarkill
           While (collumn <= 1)                                                                     'And m_ExcelSheet.Cells(Val(txtRow.Text) + j - 1, "F").Value <> ""
               prgProgress.Value = MyDiff(j, iCount) * 100                ' เอามาหาร ถ้าส่วน=0 ให้เป็น 0 อยู่ในรูปแบบ double
               txtPercent.Text = prgProgress.Value
               Me.Refresh
               If collumn = 0 Then
                  tempSheet.NetPayment = m_ExcelSheet.Cells(Val(txtRow.Text) + j - 1, txtColumn8.Text).Value    ' ดึงเซลล์จาก excell
               ElseIf collumn = 1 Then
                  tempSheet.DueDate = m_ExcelSheet.Cells(Val(txtRow.Text) + j - 1, txtColumn7.Text).Value    ' ดึงเซลล์จาก excell
               End If
               collumn = collumn + 1
            Wend
         
            If tempCarkill.m_Farm Is Nothing Then
               Set tempCarkill.m_Farm = New Collection
            End If
            tempSheet.SheetIndex = RunSheet
            tempSheet.mySheetName = m_ExcelSheet.Cells(Val(txtRow9.Text), txtColumn9.Text).Value
            Set TempRealFarm = GetObject("CXlscarkill", tempCarkill.m_Farm, Trim(tempSheet.mySheetName), False)
            If Not (TempRealFarm Is Nothing) Then
               TempRealFarm.NetPayment = Val(TempRealFarm.NetPayment) + Val(tempSheet.NetPayment)
               tempCarkill.NetPayment = Val(tempCarkill.NetPayment) + Val(tempSheet.NetPayment)
            Else
               tempCarkill.NetPayment = tempSheet.NetPayment
               Call tempCarkill.m_Farm.Add(tempSheet, Trim(tempSheet.mySheetName))
            End If
            Set tempSheet = Nothing
            Set TempRealFarm = Nothing
         End If
      Next tempCarkill
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
   
   ' * tempCollCarkill คือ ไอเท็ม m_Farm
   
      For Each TempPeriodDate In m_collPeriodDate
         If TempPeriodDate.FromDate <= tempCollCarkill.ShipDate And tempCollCarkill.ShipDate <= TempPeriodDate.ToDate Then
               
               Set SumPeriodDate = GetObject("CXlsCarkill", m_collPeriodDate, Str(TempPeriodDate.DateIndex), False)                     ' ไอเทมของ m_collPeriodDate
               Set TempPeriodDate = GetObject("CXlsCarkill", SumPeriodDate.m_Farm, Str(tempCollCarkill.SheetIndex), False)           ' SumPeriodDate = สำหรับดึง ตัวเล็ก ค่า sheet มาจากข้อมูล
               
               TempPeriodDate.SumKilo = TempPeriodDate.SumKilo + Val(tempCollCarkill.Quantity / 2)
               TempPeriodDate.SumNetpay = TempPeriodDate.SumNetpay + Val(tempCollCarkill.NetPayment)
                       
          End If
      Next TempPeriodDate
      
         prgProgress.Value = MyDiff(j, iCount) * 100                ' เอามาหาร ถ้าส่วน=0 ให้เป็น 0 อยู่ในรูปแบบ double
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
Dim initStColumn As Long
Dim initColumn As String
Dim startRow As Long
Dim startColumn4Sum As Long
Dim DatePrintColumn As Long
Dim printColumn As Long
Dim dayLoop As Long
'Dim RowSetting As Long
Dim GroupPeriodDate As CXlsCarkill
Dim TempPeriodDate As CXlsCarkill
Dim ShortName As String
Dim DateRow As Long
Dim printFinish As Boolean
Dim sumNetPayment_1 As String
Dim sumNetPayment_2 As String
Dim runColumn As Long

Dim Now_Kilo As String
Dim Str_Kilo As String
Dim Now_Net As String
Dim Str_Net As String
   
   prgProgress.MAX = 100
   prgProgress.MIN = 0
   prgProgress.Value = 0
   Set GroupPeriodDate = New CXlsCarkill

   'debug.print m_collPeriodDate.Count
 '  startColumn = Val(txtColumnTotal.Text)
 '  RowSetting = Val(txtRowTotal.Text)
   
''''   startColumn4Sum = Val(txtColumnTotal.Text) + 2
''''   printFinish = False
''''   DatePrintColumn = 0
''''   DateRow = 2
   
'''   For Each tempSheet In sheetHaveData
'''         Set m_ExcelSheet = m_ExcelApp.Sheets(tempSheet.SheetIndex)
'''            m_ExcelSheet.Cells(tempSheet.ROW + 2, startColumn).Value = "=" & txtColumn8.Text & txtRow8.Text                               ' tempCarkill.SheetName
'''            m_ExcelSheet.Cells(tempSheet.ROW, startColumn).Value = "วันที่"
'''            m_ExcelSheet.Cells(tempSheet.ROW + 1, startColumn).Value = "ก.ก. / บาท "
 ''''  Next tempSheet
   
'''  For Each tempSheet In sheetHaveData
      Set m_ExcelSheet = m_ExcelApp.Sheets(Val(txtSheet6.Text))
''      Str_Kilo = "="
''      Str_Net = "="

      startColumn = Val(txtColumn6_2.Text) ' 35 = AK , Z=24
      initStColumn = Val(txtColumn6_2.Text)
      initColumn = txtColumn6.Text
      sumNetPayment_1 = "="
      sumNetPayment_2 = "="
      j = 1
     TempDate = txtShortName.Text
     Call CalculateDatePeriod(TempDate, dayLoop)
      
      For Each TempPeriodDate In m_collPeriodDate
''          m_ExcelSheet.Cells(tempSheet.ROW, startColumn).Value = "วันที่"
           startRow = Val(txtRow7.Text)
           
           If tempDateString = 1 Then
               m_ExcelSheet.Cells(startRow - 1, startColumn).Value = Left(DateToStringExtEx2(TempPeriodDate.FromDate), 2)
           Else
               m_ExcelSheet.Cells(startRow - 1, startColumn).Value = DateToStringExtEx2(TempPeriodDate.FromDate) & " ถึง " & DateToStringExtEx2(TempPeriodDate.ToDate)
           End If
            While startRow < Val(txtRow6_1.Text)
                 Set TempRealFarm = GetObject("CXlsCarkill", TempPeriodDate.m_Farm, Trim(m_ExcelSheet.Cells(startRow, initColumn).Value), False)                ' จะรวม Sum สำหรับเกิดจริง
                 If Not (TempRealFarm Is Nothing) Then
                    m_ExcelSheet.Cells(startRow, startColumn).Value = TempRealFarm.NetPayment
'                    sumNetPayment_1 = sumNetPayment_1 & "+" & number2Column(startColumn) & Str(startRow)
'                    sumNetPayment_2 = sumNetPayment_2 & "+" & number2Column(startColumn) & Str(startRow)
                 End If
                 startRow = startRow + 1
            Wend
            startColumn = startColumn + 1                                                            ' ต้องเปลี่ยนช่วงวันที่ก่อน
            j = j + 1
            

''            dayLoop = 5
''
           If (dayLoop + 1) = j Then                         'dayLoop
            startRow = Val(txtRow7.Text)
            m_ExcelSheet.Cells(startRow - 1, startColumn).Value = "รวม " & Left(DateToStringExtEx2(TempPeriodDate.FromDate), 2) & " - " & DateToStringExtEx2(TempPeriodDate.ToDate)

            While startRow < Val(txtRow6_1.Text)
              ' startRow = Val(txtRow7.Text)
''               tempDateString = txtShortName.Text
''               Call CalculateDatePeriod(tempDateString, dayLoop)
                runColumn = 1
''                                                                                                '                    tempDateString = txtShortName.Text
''                                                                                                '                    Call CalculateDatePeriod(tempDateString, dayLoop)
''                While Len(tempDateString) > 0
                 Do While runColumn < startColumn
                     sumNetPayment_1 = sumNetPayment_1 & "+" & Trim(number2Column(startColumn - runColumn)) & Trim(Str(startRow))
                     runColumn = runColumn + 1
                     If runColumn - 1 = dayLoop Then     ' dayLoop
 ''                       If startColumn = 68 Then
   ''                        'debug.print
     ''                   End If
                        m_ExcelSheet.Cells(startRow, startColumn).Value = sumNetPayment_1
                        
                        Set TempRealFarm2 = GetObject("CXlsCarkill", m_collRealFarm, Trim(Str(startRow)), False)
                        If TempRealFarm2 Is Nothing Then
                            Set TempRealFarm2 = New CXlsCarkill
                            TempRealFarm2.InvNo = "=" & Trim(number2Column(startColumn)) & Trim(Str(startRow))
                            Call m_collRealFarm.Add(TempRealFarm2, Trim(Str(startRow)))
                        Else
                           TempRealFarm2.InvNo = TempRealFarm2.InvNo & "+" & Trim(number2Column(startColumn)) & Trim(Str(startRow))
                        End If
                        
'''                        sumNetPayment_2 = m_ExcelSheet.Cells(startRow, initStColumn + 31).Value & "+" & Trim(number2Column(startColumn)) & Trim(Str(startRow))
'''                        m_ExcelSheet.Cells(startRow, initStColumn + 31).Value = sumNetPayment_2                 '  ดึงจากตั้งต้น ไม่ run
''                        Call CalculateDatePeriod(tempDateString, dayLoop)
                        sumNetPayment_1 = "="
''                        runColumn = 0
                         Exit Do
                     End If
                  Loop
''
                   startRow = startRow + 1
              Wend
''
                startColumn = startColumn + 1
                 j = 1
                 Call CalculateDatePeriod(TempDate, dayLoop)

                 If dayLoop = 0 Then                        'dayLoop
                     startRow = Val(txtRow7.Text)
                     m_ExcelSheet.Cells(startRow - 1, startColumn).Value = DateToStringExtEx2(uctlFromDate.ShowDate) & " ถึง " & DateToStringExtEx2(uctlToDate.ShowDate)
                     While startRow < Val(txtRow6_1.Text)
                         Set TempRealFarm2 = GetObject("CXlsCarkill", m_collRealFarm, Trim(Str(startRow)), False)
                         If Not (TempRealFarm2 Is Nothing) Then
                           m_ExcelSheet.Cells(startRow, startColumn).Value = Trim(TempRealFarm2.InvNo) 'initStColumn + 31
                        End If
                        startRow = startRow + 1
                     Wend
                  End If
 ''    Wend ' tempDateString
          End If

            
      Next TempPeriodDate
      
         'debug.print
      
''''      m_ExcelSheet.Cells(tempSheet.ROW, startColumn + 1).Value = "รวม " & FirstDate & "ถึง " & LastDate
''''      m_ExcelSheet.Cells(tempSheet.ROW + 1, startColumn + 1).Value = "ก.ก. "
''''
''''      startColumn = Val(txtColumnTotal.Text)
''''  Next tempSheet
  
    startColumn = 0
    prgProgress.Value = MyDiff(j, iCount) * 100                   ' เอามาหาร ถ้าส่วน=0 ให้เป็น 0 อยู่ในรูปแบบ double
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
   pnlHeader.Caption = "เอกสารจัดจ่าย"

   Me.Picture = LoadPicture(glbParameterObj.MainPicture)

   pnlHeader.Font.Name = GLB_FONT
   pnlHeader.Font.Bold = True
   pnlHeader.Font.Size = 19

   Call InitNormalLabel(lblProgress, "ความคืบหน้า")
   Call InitNormalLabel(lblPercent, "เปอร์เซนต์")
   Call InitNormalLabel(Label1, "%")
   Call InitNormalLabel(lblFileName, "ชื่อไฟล์")

   Call InitNormalLabel(lblRow, "แถวที่")
   Call InitNormalLabel(lblRow8, "ถึงแถวที่")
   Call InitNormalLabel(lblSheet, "ตั้งแต่ชีทที่")
   Call InitNormalLabel(lblFromDate, "จากวันที่")
   Call InitNormalLabel(lblToDate, "ถึงวันที่")
   Call InitNormalLabel(lblPeriodDate, "ครั้งละ")
   Call InitNormalLabel(lblColumn7, "วันที่คอลัมน์ที่")
   Call InitNormalLabel(lblColumn8, "Net Payment คอลัมน์ที่")
   Call InitNormalLabel(lblRow9, "แถวที่")
   Call InitNormalLabel(lblColumn9, "คอลัมน์ที่")
    
   Call InitNormalLabel(lblShortName, "ช่วงวันที่จัดจ่าย")
   Call InitNormalLabel(lblSheetName, "ชื่อฟาร์ม")
   Call InitNormalLabel(lblColumn6, "คอลัมน์ที่")
   Call InitNormalLabel(lblRow6, "แถวที่")
   Call InitNormalLabel(lblReal6_0, "รวมของก่อนหน้านี้")
   Call InitNormalLabel(lblRow6_1, "ถึงแถวที่")
   Call InitNormalLabel(lblColumn6_2, "คอลัมน์ที่")
   Call InitNormalLabel(lblPrint7, "เริ่มคอลัมป์แรก")
   Call InitNormalLabel(lblRow7, "แถวที่")
   Call InitNormalLabel(lblReal6, "ชื่อฟาร์ม")
   Call InitNormalLabel(lblSheet6, "ชีทที่")

   txtPercent.Enabled = False
   txtFileName.Enabled = False
   Call txtPercent.SetTextLenType(TEXT_FLOAT, glbSetting.MONEY_TYPE)
   Call txtPeriodDate.SetTextLenType(TEXT_STRING, glbSetting.DESC_TYPE)
   Call txtShortName.SetTextLenType(TEXT_STRING, glbSetting.DESC_TYPE)

   Call txtRow.SetTextLenType(TEXT_FLOAT, glbSetting.MONEY_TYPE)
   Call txtSheet.SetTextLenType(TEXT_FLOAT, glbSetting.MONEY_TYPE)
   
   Call txtSheet6.SetTextLenType(TEXT_FLOAT, glbSetting.MONEY_TYPE)
   Call txtRow6.SetTextLenType(TEXT_FLOAT, glbSetting.MONEY_TYPE)
   Call txtColumn6.SetTextLenType(TEXT_FLOAT, glbSetting.MONEY_TYPE)
   Call txtRow6_1.SetTextLenType(TEXT_FLOAT, glbSetting.MONEY_TYPE)
   Call txtColumn6_2.SetTextLenType(TEXT_FLOAT, glbSetting.MONEY_TYPE)
   Call txtRow7.SetTextLenType(TEXT_FLOAT, glbSetting.MONEY_TYPE)
   Call txtColumn7.SetTextLenType(TEXT_FLOAT, glbSetting.MONEY_TYPE)
   Call txtRow8.SetTextLenType(TEXT_FLOAT, glbSetting.MONEY_TYPE)
   Call txtRow9.SetTextLenType(TEXT_FLOAT, glbSetting.MONEY_TYPE)
   Call txtColumn9.SetTextLenType(TEXT_FLOAT, glbSetting.MONEY_TYPE)

   cmdExit.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdOK.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdSetting.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdStart.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdFileName.Picture = LoadPicture(glbParameterObj.NormalButton1)

   Call InitMainButton(cmdExit, MapText("ยกเลิก (ESC)"))
   Call InitMainButton(cmdOK, MapText("ตกลง (F2)"))
   Call InitMainButton(cmdStart, MapText("เริ่ม"))
   Call InitMainButton(cmdFileName, MapText("..."))
   Call InitMainButton(cmdSetting, MapText("ตั้งค่า"))
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
   
      m_EstSetting.XLS_EST_SET_ID = 3       ' ใช้ ID = 3 เลยทีเดียว
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
      txtShortName.Text = m_EstSetting.COLLUMN_3
      txtSheet6.Text = m_EstSetting.SHEET_1
      txtColumn6.Text = m_EstSetting.COLLUMN_1
      txtColumn7.Text = m_EstSetting.COLLUMN_2
      txtRow6.Text = m_EstSetting.ROW_1
      txtRow7.Text = m_EstSetting.ROW_2
      txtRow8.Text = m_EstSetting.ROW_6
      txtColumn9.Text = m_EstSetting.COLLUMN_7
      txtRow9.Text = m_EstSetting.ROW_7
      txtColumn6_2.Text = m_EstSetting.COLLUMN6_2
      txtRow6_1.Text = m_EstSetting.COLLUMN6_1
      txtColumn8.Text = m_EstSetting.COLLUMN_6
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
      
      m_EstSetting.XLS_EST_SET_ID = 3

      m_EstSetting.SHEET_4 = txtSheet.Text
      m_EstSetting.ROW_4 = txtRow.Text
      m_EstSetting.SHEET_1 = txtSheet6.Text
      m_EstSetting.COLLUMN_1 = txtColumn6.Text
      m_EstSetting.COLLUMN_3 = txtShortName.Text
      m_EstSetting.COLLUMN_2 = txtColumn7.Text
      m_EstSetting.COLLUMN_6 = txtColumn8.Text
      m_EstSetting.ROW_1 = txtRow6.Text
      m_EstSetting.ROW_2 = txtRow7.Text
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
   TempDay = CLng(Val(Left(DateStr, TempID - 1)))             ' ได้ตัวเลขตัวหน้า ฝั่งซ้ายมือ
   DateStr = Mid(DateStr, TempID + 1)                               ' ตัวเลขตัวแรก กับ , ถูกตัดไปแล้ว = เหลือที่เหลือทางขวา

     TempID = InStr(1, DateStr, "/")
     TempMonth = CLng(Val(Left(DateStr, TempID - 1)))        ' ได้ตัวเลขตัวหน้า ฝั่งซ้ายมือ
     TempYear = CLng(Mid(DateStr, TempID + 1))                 ' ตัวเลขตัวแรก กับ , ถูกตัดไปแล้ว = เหลือที่เหลือทางขวา
   
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
      TempStrNew = Mid(TempStrNew, TempID + 1)                        ' ตัดวันที่ ที่ใช้ไปแล้ว
   Wend
   DateCount = DateCount + Val(TempStrNew)
End Sub

Private Sub CalculateDatePeriod(TempStr As String, DateCount As Long)
Dim TempID As Long
   TempID = InStr(1, TempStr, ",")
   If TempID > 0 Then
      DateCount = Val(Left(TempStr, TempID - 1))        ' ได้ตัวเลขตัวหน้า ฝั่งซ้ายมือ
      TempStr = Mid(TempStr, TempID + 1)                 ' ตัวเลขตัวแรก กับ , ถูกตัดไปแล้ว = เหลือที่เหลือทางขวา
   Else
      DateCount = Val(TempStr)                                 ' หมด ตัว , แล้ว เอาตัวเลขที่เหลือได้เลย
      TempStr = ""
   End If
End Sub

Private Function findBlankRow(startRow As Long) As Long
Dim collumn As Long
Dim itBlankRow As Long
Dim flagBlankRow As Boolean

   While itBlankRow < 3                                              ' ว่างครบ 3 แถวแล้ว จะเลิกวนเอง
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
   
   numA = Int(Number / 26)               ' จำนวนพยัญชนะ Eng
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

