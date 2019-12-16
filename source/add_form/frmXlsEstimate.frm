VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#2.0#0"; "THREED20.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmXlsEstimate 
   BackColor       =   &H80000000&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   6780
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11715
   Icon            =   "frmXlsEstimate.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6780
   ScaleWidth      =   11715
   StartUpPosition =   1  'CenterOwner
   Begin Threed.SSFrame SSFrame1 
      Height          =   6765
      Left            =   0
      TabIndex        =   18
      Top             =   0
      Width           =   11715
      _ExtentX        =   20664
      _ExtentY        =   11933
      _Version        =   131073
      PictureBackgroundStyle=   2
      Begin MSComctlLib.ProgressBar prgProgress 
         Height          =   315
         Left            =   2640
         TabIndex        =   16
         Top             =   4920
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
         Left            =   2640
         TabIndex        =   17
         Top             =   5280
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   820
      End
      Begin MSComDlg.CommonDialog dlgAdd 
         Left            =   11160
         Top             =   720
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin prjLedgerReport.uctlTextBox txtFileName1 
         Height          =   435
         Left            =   2640
         TabIndex        =   0
         Top             =   1320
         Width           =   6720
         _ExtentX        =   11853
         _ExtentY        =   767
      End
      Begin prjLedgerReport.uctlTextBox txtFileName3 
         Height          =   435
         Left            =   2640
         TabIndex        =   13
         Top             =   2280
         Width           =   6720
         _ExtentX        =   1905
         _ExtentY        =   767
      End
      Begin prjLedgerReport.uctlTextBox txtFileName4 
         Height          =   435
         Left            =   2640
         TabIndex        =   12
         Top             =   2760
         Width           =   6720
         _ExtentX        =   1905
         _ExtentY        =   767
      End
      Begin prjLedgerReport.uctlTextBox txtFileName2 
         Height          =   435
         Left            =   2640
         TabIndex        =   11
         Top             =   1800
         Width           =   6720
         _ExtentX        =   1905
         _ExtentY        =   767
      End
      Begin prjLedgerReport.uctlTextBox txtFileName5 
         Height          =   435
         Left            =   2640
         TabIndex        =   14
         Top             =   3240
         Width           =   6720
         _ExtentX        =   1058
         _ExtentY        =   767
      End
      Begin prjLedgerReport.uctlTextBox txtFileName6 
         Height          =   435
         Left            =   2640
         TabIndex        =   15
         Top             =   3720
         Visible         =   0   'False
         Width           =   6720
         _ExtentX        =   1270
         _ExtentY        =   767
      End
      Begin prjLedgerReport.uctlTextBox txtFileOutName 
         Height          =   435
         Left            =   2640
         TabIndex        =   29
         Top             =   4320
         Width           =   6720
         _ExtentX        =   1270
         _ExtentY        =   767
      End
      Begin Threed.SSCommand cmdFileOutName 
         Height          =   435
         Left            =   9480
         TabIndex        =   7
         Top             =   4320
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   767
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmXlsEstimate.frx":27A2
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdFileName6 
         Height          =   435
         Left            =   9480
         TabIndex        =   6
         Top             =   3720
         Visible         =   0   'False
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   767
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmXlsEstimate.frx":2ABC
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdFileName5 
         Height          =   435
         Left            =   9480
         TabIndex        =   5
         Top             =   3240
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   767
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmXlsEstimate.frx":2DD6
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdFileName4 
         Height          =   435
         Left            =   9480
         TabIndex        =   4
         Top             =   2760
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   767
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmXlsEstimate.frx":30F0
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdFileName3 
         Height          =   435
         Left            =   9480
         TabIndex        =   3
         Top             =   2280
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   767
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmXlsEstimate.frx":340A
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdFileName2 
         Height          =   435
         Left            =   9480
         TabIndex        =   2
         Top             =   1800
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   767
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmXlsEstimate.frx":3724
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdSetting 
         Height          =   525
         Left            =   4320
         TabIndex        =   9
         Top             =   6000
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         ButtonStyle     =   3
      End
      Begin VB.Label lblFileName6 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   435
         Left            =   120
         TabIndex        =   30
         Top             =   3720
         Visible         =   0   'False
         Width           =   2415
      End
      Begin VB.Label lblFileName4 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   435
         Left            =   120
         TabIndex        =   28
         Top             =   2760
         Width           =   2415
      End
      Begin VB.Label lblFileOutName 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   315
         Left            =   120
         TabIndex        =   27
         Top             =   4320
         Width           =   2415
      End
      Begin VB.Label lblFileName5 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   435
         Left            =   120
         TabIndex        =   26
         Top             =   3240
         Width           =   2415
      End
      Begin Threed.SSCommand cmdFileName1 
         Height          =   435
         Left            =   9480
         TabIndex        =   1
         Top             =   1320
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   767
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmXlsEstimate.frx":3A3E
         ButtonStyle     =   3
      End
      Begin VB.Label lblFileName2 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   435
         Left            =   120
         TabIndex        =   25
         Top             =   1800
         Width           =   2415
      End
      Begin VB.Label lblFileName3 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   435
         Left            =   120
         TabIndex        =   24
         Top             =   2280
         Width           =   2415
      End
      Begin VB.Label lblFileName1 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   435
         Left            =   120
         TabIndex        =   23
         Top             =   1320
         Width           =   2415
      End
      Begin Threed.SSCommand cmdStart 
         Height          =   525
         Left            =   2640
         TabIndex        =   8
         Top             =   6000
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmXlsEstimate.frx":3D58
         ButtonStyle     =   3
      End
      Begin VB.Label Label1 
         Caption         =   "Label1"
         Height          =   345
         Left            =   4440
         TabIndex        =   22
         Top             =   5280
         Width           =   1275
      End
      Begin VB.Label lblProgress 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   435
         Left            =   960
         TabIndex        =   21
         Top             =   4920
         Width           =   1575
      End
      Begin VB.Label lblPercent 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   435
         Left            =   960
         TabIndex        =   20
         Top             =   5280
         Width           =   1575
      End
      Begin Threed.SSCommand cmdExit 
         Height          =   525
         Left            =   9720
         TabIndex        =   10
         Top             =   6000
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   926
         _Version        =   131073
         ButtonStyle     =   3
      End
   End
End
Attribute VB_Name = "frmXlsEstimate"
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

Private m_xlsSetting As CXlsEstimateSetting
Private m_XlsFood As Collection
Private m_XlsSetFarm As Collection

Private m_ExcelApp As Object
Private m_ExcelSheet As Object
Private ConFigID As Long

Private m_collectionFile1 As Collection
Private m_collectionFile2 As Collection
Private m_collectionFile3 As Collection
Private m_collectionFile4 As Collection
Private m_collectionFile5 As Collection
Private m_collectionFile6 As Collection
Private m_collectionOutFile As Collection

Dim tempCollumnBB(14) As String
Dim tempRunDate(7) As String
Dim tempFarm As CXlsFarm


Private Sub cmdFileName1_Click()
On Error Resume Next
Dim strDescription As String

   dlgAdd.Filter = "Access Files (*.xls)|*.xls;"
   dlgAdd.DialogTitle = "Select access file to import"
   dlgAdd.ShowOpen
   If dlgAdd.FileName = "" Then
      Exit Sub
   End If
    
   txtFileName1.Text = dlgAdd.FileName
   m_HasModify = True
End Sub

Private Sub cmdFileName2_Click()
On Error Resume Next
Dim strDescription As String
   
   'edit the filter to support more image types
   dlgAdd.Filter = "Access Files (*.xls)|*.xls;"
   dlgAdd.DialogTitle = "Select access file to import"
   dlgAdd.ShowOpen
   If dlgAdd.FileName = "" Then
      Exit Sub
   End If
    
   txtFileName2.Text = dlgAdd.FileName
   m_HasModify = True
End Sub
Private Sub cmdFileName3_Click()
On Error Resume Next
Dim strDescription As String
   
   'edit the filter to support more image types
   dlgAdd.Filter = "Access Files (*.xls)|*.xls;"
   dlgAdd.DialogTitle = "Select access file to import"
   dlgAdd.ShowOpen
   If dlgAdd.FileName = "" Then
      Exit Sub
   End If
    
   txtFileName3.Text = dlgAdd.FileName
   m_HasModify = True
End Sub
Private Sub cmdFileName4_Click()
On Error Resume Next
Dim strDescription As String
   
   'edit the filter to support more image types
   dlgAdd.Filter = "Access Files (*.xls)|*.xls;"
   dlgAdd.DialogTitle = "Select access file to import"
   dlgAdd.ShowOpen
   If dlgAdd.FileName = "" Then
      Exit Sub
   End If
    
   txtFileName4.Text = dlgAdd.FileName
   m_HasModify = True
End Sub
Private Sub cmdFileName5_Click()
On Error Resume Next
Dim strDescription As String
   
   'edit the filter to support more image types
   dlgAdd.Filter = "Access Files (*.xls)|*.xls;"
   dlgAdd.DialogTitle = "Select access file to import"
   dlgAdd.ShowOpen
   If dlgAdd.FileName = "" Then
      Exit Sub
   End If
    
   txtFileName5.Text = dlgAdd.FileName
   m_HasModify = True
End Sub
Private Sub cmdFileName6_Click()
On Error Resume Next
Dim strDescription As String
   
   'edit the filter to support more image types
   dlgAdd.Filter = "Access Files (*.xls)|*.xls;"
   dlgAdd.DialogTitle = "Select access file to import"
   dlgAdd.ShowOpen
   If dlgAdd.FileName = "" Then
      Exit Sub
   End If
    
   txtFileName6.Text = dlgAdd.FileName
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

Private Sub cmdsetting_Click()
Dim oMenu As cPopupMenu
Dim lMenuChosen As Long
Dim Ac As CAccountCode
Dim ItemCount As Long
Dim m_Rs As ADODB.Recordset
   Set oMenu = New cPopupMenu
   lMenuChosen = oMenu.Popup("ตั้งค่า แถว,คอลัมป์", "ตั้งค่าเบอร์อาหาร", "ตั้งค่าหน่วยและน้ำหนักต่อรอบ", "ตั้งค่าค่าขนส่ง", "ตั้งค่าชื่อฟาร์มหลัก")
   If lMenuChosen = 0 Then
      Exit Sub
   End If
   Set oMenu = Nothing
   
   Call EnableForm(Me, False)
   If lMenuChosen = 1 Then
      frmXlsEstimateSetting.ShowMode = SHOW_EDIT
      Load frmXlsEstimateSetting
      frmXlsEstimateSetting.Show 1

      Unload frmXlsEstimateSetting
      Set frmXlsEstimateSetting = Nothing
      
   ElseIf lMenuChosen = 2 Then
      Load frmXlsEstFoodNum
      frmXlsEstFoodNum.Show 1

      Unload frmXlsEstFoodNum
      Set frmXlsEstFoodNum = Nothing
      
   ElseIf lMenuChosen = 3 Then
      Load frmXlsUnit
      frmXlsUnit.Show 1

      Unload frmXlsUnit
      Set frmXlsUnit = Nothing
      
   ElseIf lMenuChosen = 4 Then
      Load frmXlsSetFarm
      frmXlsSetFarm.Show 1

      Unload frmXlsSetFarm
      Set frmXlsSetFarm = Nothing
      
    ElseIf lMenuChosen = 5 Then
      Load frmXlsMainFarm
      frmXlsMainFarm.Show 1

      Unload frmXlsMainFarm
      Set frmXlsMainFarm = Nothing
   
      
   End If
   
   Call EnableForm(Me, True)
   
End Sub

Private Sub cmdStart_Click()
Dim TempID As Long
Dim HasBegin As Boolean
Dim MaxSheet As Long
Dim Ac As CAccountCode

'   Call SaveData
'   Call EnableForm(Me, False)

'   If Val(txtSheet.Text) > MaxSheet Then
'      Call MsgBox("กรุณากรอกข้อมูล ชีดให้ถูกต้องโดยไม่สามารถมากกว่า  " & MaxSheet, vbOKOnly, PROJECT_NAME)
'      Exit Sub
'   End If

   ' 1 ดึงเงื่อนไข จากการตั้งค่าทั้ง 3 ส่วน
   Call LoadXlsSetting(m_xlsSetting)
   tempCollumnBB(0) = m_xlsSetting.COLLUMN6_1
   tempCollumnBB(1) = m_xlsSetting.COLLUMN5_1
   tempCollumnBB(2) = m_xlsSetting.COLLUMN6_2
   tempCollumnBB(3) = m_xlsSetting.COLLUMN5_2
   tempCollumnBB(4) = m_xlsSetting.COLLUMN6_3
   tempCollumnBB(5) = m_xlsSetting.COLLUMN5_3
   tempCollumnBB(6) = m_xlsSetting.COLLUMN6_4
   tempCollumnBB(7) = m_xlsSetting.COLLUMN5_4
   tempCollumnBB(8) = m_xlsSetting.COLLUMN6_5
   tempCollumnBB(9) = m_xlsSetting.COLLUMN5_5
   tempCollumnBB(10) = m_xlsSetting.COLLUMN6_6
   tempCollumnBB(11) = m_xlsSetting.COLLUMN5_6
   tempCollumnBB(12) = m_xlsSetting.COLLUMN6_7
   tempCollumnBB(13) = m_xlsSetting.COLLUMN5_7

   Call LoadXlsFood(m_XlsFood)
   Call LoadXlsSetFarm(m_XlsSetFarm)
   
   ' 2 อ่าน xls ทั้ง 6 ไฟล์ เก็บใส่ 6 collection == ใช้เงื่อนไขจาก 1 ในการอ่าน
   If txtFileName1.Text <> "" Then              ' 1
      m_ExcelApp.Workbooks.Close
      m_ExcelApp.Workbooks.Open (txtFileName1.Text)
   
      MaxSheet = m_ExcelApp.Sheets.Count
      Call ExportAccount(m_collectionFile1)

      m_ExcelApp.Workbooks.Close
   End If
   
   If txtFileName2.Text <> "" Then           ' 2
      m_ExcelApp.Workbooks.Close
      m_ExcelApp.Workbooks.Open (txtFileName2.Text)
   
      MaxSheet = m_ExcelApp.Sheets.Count
      Call ExportAccount(m_collectionFile2)

      m_ExcelApp.Workbooks.Close
   End If
   
   If txtFileName3.Text <> "" Then           ' 3
      m_ExcelApp.Workbooks.Close
      m_ExcelApp.Workbooks.Open (txtFileName3.Text)
   
      MaxSheet = m_ExcelApp.Sheets.Count
      Call ExportAccount(m_collectionFile3)

      m_ExcelApp.Workbooks.Close
   End If
   
   If txtFileName4.Text <> "" Then           ' 4
      m_ExcelApp.Workbooks.Close
      m_ExcelApp.Workbooks.Open (txtFileName4.Text)
   
      MaxSheet = m_ExcelApp.Sheets.Count
      Call ExportAccount(m_collectionFile4)

      m_ExcelApp.Workbooks.Close
   End If
   
   If txtFileName5.Text <> "" Then        ' 5
      m_ExcelApp.Workbooks.Close
      m_ExcelApp.Workbooks.Open (txtFileName5.Text)
   
      MaxSheet = m_ExcelApp.Sheets.Count
      Call ExportAccount(m_collectionFile5)

      m_ExcelApp.Workbooks.Close
   End If
   
   If txtFileName6.Text <> "" Then        '6
      m_ExcelApp.Workbooks.Close
      m_ExcelApp.Workbooks.Open (txtFileName6.Text)
   
      MaxSheet = m_ExcelApp.Sheets.Count
      Call ExportAccount(m_collectionFile6)
      
      m_ExcelApp.Workbooks.Close
   End If
   
   If Not VerifyTextControl(lblFileOutName, txtFileOutName) Then
      Exit Sub
   End If
   
   ' 3 ประมวลผล เขียนใส่ xls
   ' 3.1 รวบรวมทุกคอเล็กชั่นให้เป็น collection เดียว
   Call MixXlsFile(m_collectionFile1)
   Call MixXlsFile(m_collectionFile2)
   Call MixXlsFile(m_collectionFile3)
   Call MixXlsFile(m_collectionFile4)
   Call MixXlsFile(m_collectionFile5)
   Call MixXlsFile(m_collectionFile6)
   
   m_ExcelApp.Workbooks.Open (txtFileOutName.Text)       ' template มา
   MaxSheet = m_ExcelApp.Sheets.Count
   Call WriteXlsFile
   m_ExcelApp.Workbooks.Close                                     ' ปิด template
   
  Call EnableForm(Me, True)
 
End Sub

Private Sub ExportAccount(Optional m_collectionFile As Collection = Nothing)
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

Dim farmName1 As String
Dim farmName2 As String
Dim weekNum As String
Dim fromDateHead As String
Dim toDateHead As String

Dim runDate(14) As String
Dim FoodNum(100) As String

Dim xlsFood_temp As CXlsFood
Dim xlsSetFarm_temp As CXlsSetFarm
Dim R As Long
Dim C As Long

   prgProgress.MAX = 100
   prgProgress.MIN = 0
   
'   Call LoadGLJNLforAccountExcel(Nothing, SearchCollection, uctlFromDate.ShowDate, uctlToDate.ShowDate)
'   Call LoadGLAccSearch(Nothing, SearchNameCollection)
   
   prgProgress.MAX = 100
   prgProgress.MIN = 0
   prgProgress.Value = 0

   i = 0
   
   Set m_ExcelSheet = m_ExcelApp.Sheets(m_xlsSetting.SHEET_1)

   j = 0

'   iCount = Val(txtRow2.Text) - Val(txtRow.Text)
'   While (j < iCount)
'      j = j + 1
'      prgProgress.Value = MyDiff(j, iCount) * 100
'      txtPercent.Text = prgProgress.Value
'      Me.Refresh

      farmName1 = m_ExcelSheet.Cells(m_xlsSetting.ROW_1, m_xlsSetting.COLLUMN_1).Value        ' ดึงเซลล์จาก excell ได้สำเร็จ !!
      weekNum = m_ExcelSheet.Cells(m_xlsSetting.ROW_2, m_xlsSetting.COLLUMN_2).Value
      fromDateHead = m_ExcelSheet.Cells(m_xlsSetting.ROW_3, m_xlsSetting.COLLUMN_3).Value
      toDateHead = m_ExcelSheet.Cells(m_xlsSetting.ROW_4, m_xlsSetting.COLLUMN_4).Value
      
'      For j = 0 To UBound(tempCollumnBulk) - 1
'          runDate(j) = m_ExcelSheet.Cells(m_xlsSetting.ROW_5, tempCollumnBulk(j)).Value
'      Next j
'
'      For j = m_xlsSetting.FROMDATAROW To m_xlsSetting.TODATAROW
'          FoodNum(j) = m_ExcelSheet.Cells(j, m_xlsSetting.COLLUMNFOOD).Value
'      Next j
      
      Set m_collectionFile = New Collection
     
      For C = 0 To UBound(tempCollumnBB) - 1                                       'หลัก
      
         For R = m_xlsSetting.FROMDATAROW To m_xlsSetting.TODATAROW     'แถว
           FoodNum(R) = m_ExcelSheet.Cells(R, m_xlsSetting.COLLUMNFOOD).Value
           
           
             Set tempFarm = New CXlsFarm
             tempFarm.weekNum = m_ExcelSheet.Cells(m_xlsSetting.ROW_2, m_xlsSetting.COLLUMN_2).Value
             tempFarm.farmName1 = m_ExcelSheet.Cells(m_xlsSetting.ROW_1, m_xlsSetting.COLLUMN_1).Value
             tempFarm.fromDateHead = m_ExcelSheet.Cells(m_xlsSetting.ROW_3, m_xlsSetting.COLLUMN_3).Value
             tempFarm.toDateHead = m_ExcelSheet.Cells(m_xlsSetting.ROW_4, m_xlsSetting.COLLUMN_4).Value

             tempFarm.F_unitName = m_ExcelSheet.Cells(m_xlsSetting.ROWBB, tempCollumnBB(C)).Value      ' หัวเรื่อง Bag หรือ Bulk

             If C Mod 2 = 0 Then
                runDate(C) = m_ExcelSheet.Cells(m_xlsSetting.ROW_5, tempCollumnBB(C + 1)).Value
             Else
                runDate(C) = m_ExcelSheet.Cells(m_xlsSetting.ROW_5, tempCollumnBB(C)).Value
             End If
             
             Set xlsSetFarm_temp = GetObject("CXlsSetFarm", m_XlsSetFarm, Trim(tempFarm.farmName1) & "-" & Trim(tempFarm.F_unitName), False)     ' ราคาของสินค้าตัวนั้น
             If Not (xlsSetFarm_temp Is Nothing) Then
                tempFarm.F_TRANS_PRICE = xlsSetFarm_temp.SET_FARM_PRICE
             End If
             
             Set xlsFood_temp = GetObject("CXlsFood", m_XlsFood, Trim(FoodNum(R)) & "-" & Trim(tempFarm.F_unitName), False)     ' ราคาของสินค้าตัวนั้น
             If Not (xlsFood_temp Is Nothing) Then
                tempFarm.F_cost = xlsFood_temp.XLS_FOOD_COST
                tempFarm.F_unitValue = xlsFood_temp.XLS_UNIT_LIMIT              ' 12000 หรือ 330
                tempFarm.unit_multiply = xlsFood_temp.XLS_UNIT_MULTIPLY
             End If
             
             tempFarm.FoodNum = FoodNum(R)
             tempFarm.F_Value = m_ExcelSheet.Cells(R, tempCollumnBB(C)).Value                      ' ค่าที่อยู่ในตาราง
             tempFarm.F_date = runDate(C)
             tempFarm.F_destination = m_ExcelSheet.Cells(R, "R").Value
             
              If tempFarm.FoodNum <> "" And tempFarm.FoodNum <> "รวม" And tempFarm.FoodNum <> "รวมทั้งหมด" Then
                 Call m_collectionFile.Add(tempFarm, Trim(tempFarm.FoodNum) & "-" & Trim(tempFarm.F_destination) & "-" & Trim(tempFarm.F_date) & "-" & Trim(tempFarm.F_unitName))
              End If
              Set tempFarm = Nothing
               
           Next R
   Next C
   
'      If TempNo <> "" Then
'
'         If Trim(TempNo) = "212-1100" Then
'            'debug.print ("")
'         End If
'
'         FindCode = False
'         For Each Ac In MainCollection                                                                            ' วนจากที่คิวรี่ access มา
'            If Ac.MAIN_CODE = Trim(TempNo) Then
'               FindCode = True
'               Exit For
'            End If
'         Next Ac

 '        Set GC = GetGLAcc(SearchNameCollection, Trim(TempNo))                  ' ดึงมาจาก excell ตาราง CGLAcc

'''''         m_ExcelSheet.Cells(Val(txtrow.Text) + j - 1, Val(txtCollumn4.Text)).Value = GC.ACCNAM       ' เงินสด-ฟาร์ม  ,, txtCollumn4.Text =คอลัมป์ตรวจ 21

'         ReturnV = True
'         ReturnZ = True
'
'         Debit = 0
'         Credit = 0
'         Deposit = 0
'
'         If Not (FindCode) Then
'            Deposit = GC.BEGCUR
'            If GC.GROUP = 2 Or GC.GROUP = 3 Then
'               Deposit = Deposit * -1
'            End If
'            Set Gl = GetGLJnl(SearchCollection, Trim(TempNo) & "-0", ReturnV)
'            Debit = Gl.AMOUNT
'            Set Gl = GetGLJnl(SearchCollection, Trim(TempNo) & "-1", ReturnZ)
'            Credit = Gl.AMOUNT
'            If Deposit + Debit - Credit > 0 Then
'               m_ExcelSheet.Cells(Val(txtRow.Text) + j - 1, Val(txtCollumn2.Text)).Value = Abs(Deposit + Debit - Credit)
'            ElseIf Deposit + Debit - Credit < 0 Then
'               m_ExcelSheet.Cells(Val(txtRow.Text) + j - 1, Val(txtCollumn3.Text)).Value = Abs(Deposit + Debit - Credit)
'            End If
'         Else
'            Debit = 0
'            Credit = 0
'            Deposit = 0
'            Call Recuresive(Trim(Ac.SUB_CODE), Debit, Credit, ReturnV, ReturnZ, Deposit)
'            If ReturnV Then
'               m_ExcelSheet.Cells(Val(txtRow.Text) + j - 1, Val(txtCollumn2.Text)).Font.colorindex = 1
'            Else
'               m_ExcelSheet.Cells(Val(txtRow.Text) + j - 1, Val(txtCollumn2.Text)).Font.colorindex = 3
'            End If
'            If ReturnZ Then
'               m_ExcelSheet.Cells(Val(txtRow.Text) + j - 1, Val(txtCollumn3.Text)).Font.colorindex = 1
'            Else
'               m_ExcelSheet.Cells(Val(txtRow.Text) + j - 1, Val(txtCollumn3.Text)).Font.colorindex = 3
'            End If
'            If Deposit + Debit - Credit > 0 Then
'               m_ExcelSheet.Cells(Val(txtRow.Text) + j - 1, Val(txtCollumn2.Text)).Value = Abs(Deposit + Debit - Credit)
'            ElseIf Deposit + Debit - Credit < 0 Then
'               m_ExcelSheet.Cells(Val(txtRow.Text) + j - 1, Val(txtCollumn3.Text)).Value = Abs(Deposit + Debit - Credit)
'            End If
'         End If
'         If ReturnV Or ReturnZ Then
'            m_ExcelSheet.Cells(Val(txtRow.Text) + j - 1, Val(txtCollumn4.Text)).Font.colorindex = 1
'         Else
'            m_ExcelSheet.Cells(Val(txtRow.Text) + j - 1, Val(txtCollumn4.Text)).Font.colorindex = 3
'         End If
'      End If
'   Wend
   
   prgProgress.Value = 100
   txtPercent.Text = 100
      
   Exit Sub
   
ErrorHandler:
   glbErrorLog.LocalErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowUserError
End Sub

Private Sub MixXlsFile(Optional m_collection As Collection = Nothing)

      If Not (m_collection Is Nothing) Then                       ' 1
         For Each tempFarm In m_collection
            If tempFarm.F_Value <> 0 Then
                     Call m_collectionOutFile.Add(tempFarm, Trim(tempFarm.farmName1) & "-" & Trim(tempFarm.FoodNum) & "-" & Trim(tempFarm.F_destination) & "-" & Trim(tempFarm.F_date) & "-" & Trim(tempFarm.F_unitName))
                     Set tempFarm = Nothing
            End If
         Next tempFarm
      End If
Exit Sub

ErrorHandler:
   glbErrorLog.LocalErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowUserError
End Sub

Private Sub WriteXlsFile()
Dim firstRow As Long
Dim shortSumRow As Long
Dim headRow As Long
Dim runColumn As Long
Dim carNum As Long
'Dim fromRowFarm As Long
'Dim ToRowFarm As Long
'Dim Total1(20) As Double
'Dim Total2(20) As Double
Dim strSumbyFarm1 As String
Dim strSumbyFarm2 As String
Dim strSumbyAll1 As String
Dim strSumbyAll2 As String
Dim tempUnit As Double
Dim tempTransPrice As Double

Dim PrevKey1 As String     ' farmName1
Dim PrevKey2 As String     ' f_unitName
Dim PrevKey3 As String
'Dim m_tempCollections As Collection

   'ตั้งค่า คอลัมป์ตามเท็มเพลต ฮาร์ทโค๊ดได้เลย
   tempCollumnBB(0) = "A"     ' ชื่อฟาร์ม = สถานที่ส่ง
   tempCollumnBB(1) = "B"     ' จำนวนคัน
   tempCollumnBB(2) = "C"     ' เบอร์อาหาร
   tempCollumnBB(3) = "D"     ' จำนวน/กส.
   tempCollumnBB(4) = "E"     ' จำนวน/กก
   tempCollumnBB(5) = "F"     ' ราคา
   tempCollumnBB(6) = "G"    ' ค่าขนส่ง
   tempCollumnBB(7) = "H"     ' ราคารวม
   tempCollumnBB(8) = "I"       ' วันที่ส่ง
   tempCollumnBB(9) = "J"      ' สถานที่ส่ง
   
   Set m_ExcelSheet = m_ExcelApp.Sheets(m_xlsSetting.SHEET_1)
   firstRow = 3
   shortSumRow = 60
   headRow = 1
   carNum = 0
   tempUnit = 0
   strSumbyFarm1 = "="
   strSumbyFarm2 = "="
   strSumbyAll1 = "="
   strSumbyAll2 = "="
   
   ' 3.2 จาก coll เดียว ค่อยๆพิมพ์จนเป็น 1 ไฟล์ xls ใช้คอนเซ็ปอาหารคนละเบอร์
   For Each tempFarm In m_collectionOutFile
'      Set m_tempCollections = New Collection
'      Set m_tempCollections = m_collectionOutFile.ITEM(firstRow)
      
    'เขียนชื่อหน้ากระดาษ
    If firstRow = 3 Then
      m_ExcelSheet.Cells(headRow, tempCollumnBB(4)).Value = DateToStringExtEx2(tempFarm.fromDateHead)
      m_ExcelSheet.Cells(headRow, tempCollumnBB(6)).Value = DateToStringExtEx2(tempFarm.toDateHead)
      m_ExcelSheet.Cells(headRow, tempCollumnBB(9)).Value = tempFarm.weekNum
      
      m_ExcelSheet.Cells(shortSumRow - 2, "B").Value = DateToStringExtEx2(tempFarm.fromDateHead) & " ถึง " & DateToStringExtEx2(tempFarm.toDateHead)
      m_ExcelSheet.Cells(shortSumRow - 1, "B").Value = "กก."
      m_ExcelSheet.Cells(shortSumRow - 1, "C").Value = "จำนวนเงิน"

    End If
      
      If PrevKey1 <> tempFarm.farmName1 And firstRow <> 3 Then
         Call Sumbyfarm(firstRow, "รวม", strSumbyFarm1, strSumbyFarm2)     ' แสดงบรรทัดรวม
         Call shortSumbyfarm(shortSumRow, PrevKey1, strSumbyFarm1, strSumbyFarm2)     ' ย่อรวม
         strSumbyFarm1 = "="
         strSumbyFarm2 = "="
         carNum = 0
         tempUnit = 0
         firstRow = firstRow + 1
         shortSumRow = shortSumRow + 1
      End If

            For runColumn = 0 To 9
               If runColumn = 0 Then
                       m_ExcelSheet.Cells(firstRow, tempCollumnBB(runColumn)).Value = tempFarm.farmName1        ' เงินสด-ฟาร์ม  ,, txtCollumn4.Text =คอลัมป์ตรวจ 21
               
               ElseIf runColumn = 1 Then
                  tempTransPrice = 0
                  tempUnit = tempUnit + tempFarm.F_Value
                  If tempFarm.F_date = PrevKey3 Then
                       If tempFarm.F_unitName <> PrevKey2 Then           ' เป็นเงื่อนไขที่ต้องขึ้นคันใหม่ = มีค่าขนส่ง
                          carNum = carNum + 1
                          tempUnit = 0
                          tempUnit = tempUnit + tempFarm.F_Value
                          tempTransPrice = tempFarm.F_TRANS_PRICE
                       Else   'เป็นหน่วย Bag หรือ bulk เดียวกัน
                          If tempUnit > tempFarm.F_unitValue Then
                              carNum = carNum + 1
                              tempUnit = 0
                              tempUnit = tempUnit + tempFarm.F_Value
                              tempTransPrice = tempFarm.F_TRANS_PRICE
                          End If
                       End If
                    Else    ' คนละวันกัน
                       carNum = carNum + 1
                       tempUnit = 0
                       tempUnit = tempUnit + tempFarm.F_Value
                       tempTransPrice = tempFarm.F_TRANS_PRICE
                    End If
                       m_ExcelSheet.Cells(firstRow, tempCollumnBB(runColumn)).Value = carNum
                       
               ElseIf runColumn = 2 Then
                       m_ExcelSheet.Cells(firstRow, tempCollumnBB(runColumn)).Value = tempFarm.FoodNum
               ElseIf runColumn = 3 Then
                       m_ExcelSheet.Cells(firstRow, tempCollumnBB(runColumn)).Value = tempFarm.F_Value        ' จำนวนก.ก.
'                       Total1(runColumn) = Total1(runColumn) + tempFarm.F_unitValue
'                       Total2(runColumn) = Total2(runColumn) + tempFarm.F_unitValue
               ElseIf runColumn = 4 Then
                       m_ExcelSheet.Cells(firstRow, tempCollumnBB(runColumn)).Value = Val(tempFarm.F_Value) * Val(tempFarm.unit_multiply)
                       strSumbyFarm1 = strSumbyFarm1 & "+" & Trim(tempCollumnBB(runColumn)) & Trim(firstRow)
                       strSumbyAll1 = strSumbyAll1 & "+" & Trim(tempCollumnBB(runColumn)) & Trim(firstRow)
               ElseIf runColumn = 5 Then
                       m_ExcelSheet.Cells(firstRow, tempCollumnBB(runColumn)).Value = tempFarm.F_cost
               ElseIf runColumn = 6 Then
                     If tempTransPrice <> 0 Then
                       m_ExcelSheet.Cells(firstRow, tempCollumnBB(runColumn)).Value = tempTransPrice                 ' ค่าขนส่ง
                     End If
               ElseIf runColumn = 7 Then
                       m_ExcelSheet.Cells(firstRow, tempCollumnBB(runColumn)).Value = "=D" & Trim(Str(firstRow)) & "*F" & Trim(Str(firstRow)) & "+G" & Trim(Str(firstRow))               ' ราคารวม
'                       Total1(runColumn) = Total1(runColumn) + tempFarm.F_unitValue
'                       Total2(runColumn) = Total2(runColumn) + tempFarm.F_unitValue
                        strSumbyFarm2 = strSumbyFarm2 & "+" & Trim(tempCollumnBB(runColumn)) & Trim(firstRow)
                        strSumbyAll2 = strSumbyAll2 & "+" & Trim(tempCollumnBB(runColumn)) & Trim(firstRow)
               ElseIf runColumn = 8 Then
'                        If firstRow = 16 Or firstRow = 40 Or firstRow = 24 Then
'                           'debug.print
'                        End If
                       m_ExcelSheet.Cells(firstRow, tempCollumnBB(runColumn)).Value = DateToStringExtEx2(tempFarm.F_date)                 ' วันที่ส่ง
               ElseIf runColumn = 9 Then
                       m_ExcelSheet.Cells(firstRow, tempCollumnBB(runColumn)).Value = tempFarm.F_destination          ' หมายเหตุ
               End If
            Next runColumn
      PrevKey1 = tempFarm.farmName1
      PrevKey2 = tempFarm.F_unitName
      PrevKey3 = tempFarm.F_date
      firstRow = firstRow + 1
   Next tempFarm


    Call Sumbyfarm(firstRow, "รวม", strSumbyFarm1, strSumbyFarm2)     ' แสดงบรรทัดรวม
    Call Sumbyfarm(firstRow + 1, "รวมทั้งหมด", strSumbyAll1, strSumbyAll2)   ' แสดงบรรทัดรวม
    m_ExcelSheet.Cells(firstRow + 2, tempCollumnBB(7)).Value = "=" & Trim(tempCollumnBB(7)) & Trim(Str(firstRow + 1)) & "/" & Trim(tempCollumnBB(4)) & Trim(Str(firstRow + 1))
    
    Call shortSumbyfarm(shortSumRow, "รวมสายสุกร", strSumbyFarm1, strSumbyAll2 & "-(" & Right(strSumbyFarm2, Len(strSumbyFarm2) - 2) & ")") ' ย่อรวม รวมเฉพาะสายสุกร
    Call shortSumbyfarm(shortSumRow + 1, PrevKey1, strSumbyFarm1, strSumbyFarm2)   ' ย่อรวม ของ MM
    Call shortSumbyfarm(shortSumRow + 2, "รวม", strSumbyAll1, strSumbyAll2)   ' แสดงบรรทัดรวม
    
'      If Not (m_collectionOutFile Is Nothing) Then                       ' 1
'         For Each tempFarm In m_collectionOutFile
'            If tempFarm.F_Value <> 0 Then
'                     Call m_collectionOutFile.Add(tempFarm, Trim(tempFarm.farmName1) & "-" & Trim(tempFarm.FoodNum) & "-" & Trim(tempFarm.F_date) & "-" & Trim(tempFarm.F_unitName))
'                     Set tempFarm = Nothing
'            End If
'         Next tempFarm
'      End If
      
Exit Sub

ErrorHandler:
   glbErrorLog.LocalErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowUserError
End Sub
Private Sub Sumbyfarm(firstRow As Long, Txt0 As String, Txt1 As String, Txt2 As String)
Dim runColumn As Long

   For runColumn = 0 To 9
      If runColumn = 0 Then
            m_ExcelSheet.Cells(firstRow, tempCollumnBB(runColumn)).Value = Txt0
       ElseIf runColumn = 4 Then
            m_ExcelSheet.Cells(firstRow, tempCollumnBB(runColumn)).Value = Txt1
       ElseIf runColumn = 7 Then
            m_ExcelSheet.Cells(firstRow, tempCollumnBB(runColumn)).Value = Txt2
       End If
   Next runColumn

End Sub

Private Sub shortSumbyfarm(shortSumRow As Long, Txt0 As String, Txt1 As String, Txt2 As String)
Dim runColumn As Long

   For runColumn = 1 To 3
      If runColumn = 1 Then
            m_ExcelSheet.Cells(shortSumRow, "A").Value = Txt0
       ElseIf runColumn = 2 Then
            m_ExcelSheet.Cells(shortSumRow, "B").Value = Txt1
       ElseIf runColumn = 3 Then
            m_ExcelSheet.Cells(shortSumRow, "C").Value = Txt2
       End If
   Next runColumn

End Sub

Private Sub Recuresive(Text As String, Debit As Double, Credit As Double, ReturnV As Boolean, ReturnZ As Boolean, Deposit As Double)
Dim Gl As CGLJnl
Dim TempNo As String
Dim Pos As Long
Dim ReturnX As Boolean
Dim Ac As CGLAcc
   
'   Pos = InStr(1, Text, ",")
'   If Pos = 0 Then
'      TempNo = Text
'      Text = ""
'   Else
'      TempNo = Left(Text, Pos - 1)
'      Text = Mid(Text, Pos + 1, Len(Text) - Pos)
'   End If
'
'   Set Ac = GetGLAcc(SearchNameCollection, Trim(TempNo))
'
'   Set Gl = GetGLJnl(SearchCollection, Trim(TempNo) & "-0", ReturnX)
'   ReturnV = ReturnV And ReturnX
'   Debit = Debit + Gl.AMOUNT
'   Set Gl = GetGLJnl(SearchCollection, Trim(TempNo) & "-1")
'   ReturnZ = ReturnZ And ReturnX
'   Credit = Credit + Gl.AMOUNT
'
'   If Ac.GROUP = 2 Or Ac.GROUP = 3 Then
'      Deposit = Deposit - Ac.BEGCUR
'   Else
'      Deposit = Deposit + Ac.BEGCUR
'   End If
'
'   If Text <> "" Then
'      Call Recuresive(Text, Debit, Credit, ReturnV, ReturnZ, Deposit)
'   End If
End Sub

Private Sub Form_Activate()
Dim FromDate As Date
Dim ToDate As Date

   If Not m_HasActivate Then
      m_HasActivate = True
'      Me.Refresh
'      DoEvents
'
'      m_HasModify = False
'
'      Call LoadAccountCode(Nothing, MainCollection)
'
'      Call QueryData
'
'      GridEX1.ItemCount = CountItem(MainCollection)
'      GridEX1.Rebind
      
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
'   ElseIf Shift = 0 And KeyCode = 113 Then
'      Call cmdOK_Click
'      KeyCode = 0
   ElseIf Shift = 0 And KeyCode = 114 Then
'      Call cmdEdit_Click
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
   pnlHeader.Caption = "ประมาณอาหาร"
   
   Me.Picture = LoadPicture(glbParameterObj.MainPicture)
   
   pnlHeader.Font.Name = GLB_FONT
   pnlHeader.Font.Bold = True
   pnlHeader.Font.Size = 19
   
 '  Call InitNormalLabel(lblName, "รายการ")
   
   Call InitNormalLabel(lblProgress, "ความคืบหน้า")
   Call InitNormalLabel(lblPercent, "เปอร์เซนต์")
   Call InitNormalLabel(Label1, "%")
   Call InitNormalLabel(lblFileName1, "ไฟล์ สายสุกร1")
   Call InitNormalLabel(lblFileName2, "ไฟล์ สายสุกร2")
   Call InitNormalLabel(lblFileName3, "ไฟล์ สายสุกร3")
   Call InitNormalLabel(lblFileName4, "ไฟล์ สายสุกร4")
   Call InitNormalLabel(lblFileName5, "ไฟล์ MM")
   Call InitNormalLabel(lblFileName6, "ชื่อไฟล์6")
   
   Call InitNormalLabel(lblFileOutName, "ชื่อไฟล์ที่ save")

'   Call InitNormalLabel(lblCollumn, "คอลัมน์รหัส")
'   Call InitNormalLabel(lblRow, "แถวเริ่ม")
'   Call InitNormalLabel(lblRow2, "แถวจบ")
'   Call InitNormalLabel(lblSheet, "ชีด")
'   Call InitNormalLabel(lblCollumn2, "คอลัมน์เดบิต")
'   Call InitNormalLabel(lblCollumn3, "คอลัมน์เครดิต")
'   Call InitNormalLabel(lblCollumn4, "คอลัมน์ตรวจ")
'
'   Call InitNormalLabel(lblFromDate, "จากวันที่")
'   Call InitNormalLabel(lblToDate, "ถึงวันที่")
   
   Call txtPercent.SetTextLenType(TEXT_FLOAT, glbSetting.MONEY_TYPE)
   txtPercent.Enabled = False
   Call txtFileName1.SetTextLenType(TEXT_STRING, glbSetting.DESC_TYPE)
   Call txtFileName2.SetTextLenType(TEXT_STRING, glbSetting.DESC_TYPE)
   Call txtFileName3.SetTextLenType(TEXT_STRING, glbSetting.DESC_TYPE)
   Call txtFileName4.SetTextLenType(TEXT_STRING, glbSetting.DESC_TYPE)
   Call txtFileName5.SetTextLenType(TEXT_STRING, glbSetting.DESC_TYPE)
   Call txtFileName6.SetTextLenType(TEXT_STRING, glbSetting.DESC_TYPE)
   Call txtFileOutName.SetTextLenType(TEXT_STRING, glbSetting.DESC_TYPE)
   'txtFileName.Enabled = False
'   Call txtCollumn.SetTextLenType(TEXT_FLOAT, glbSetting.MONEY_TYPE)
'   Call txtRow.SetTextLenType(TEXT_FLOAT, glbSetting.MONEY_TYPE)
'   Call txtRow2.SetTextLenType(TEXT_FLOAT, glbSetting.MONEY_TYPE)
'   Call txtSheet.SetTextLenType(TEXT_FLOAT, glbSetting.MONEY_TYPE)
'   Call txtCollumn2.SetTextLenType(TEXT_FLOAT, glbSetting.MONEY_TYPE)
'   Call txtCollumn3.SetTextLenType(TEXT_FLOAT, glbSetting.MONEY_TYPE)
'   Call txtCollumn4.SetTextLenType(TEXT_FLOAT, glbSetting.MONEY_TYPE)
'   Call txtTemptxt.SetTextLenType(TEXT_STRING, glbSetting.CODE_TYPE)
'
'   txtTemptxt.Enabled = False
   
   cmdExit.Picture = LoadPicture(glbParameterObj.NormalButton1)
'   cmdOK.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdStart.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdFileName1.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdFileName2.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdFileName3.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdFileName4.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdFileName5.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdFileName6.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdFileOutName.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdSetting.Picture = LoadPicture(glbParameterObj.NormalButton1)
   
   Call InitMainButton(cmdExit, MapText("ยกเลิก (ESC)"))
 '  Call InitMainButton(cmdOK, MapText("ตกลง (F2)"))
   Call InitMainButton(cmdStart, MapText("เริ่ม"))
   
   Call InitMainButton(cmdFileName1, MapText("..."))
   Call InitMainButton(cmdFileName2, MapText("..."))
   Call InitMainButton(cmdFileName3, MapText("..."))
   Call InitMainButton(cmdFileName4, MapText("..."))
   Call InitMainButton(cmdFileName5, MapText("..."))
   Call InitMainButton(cmdFileName6, MapText("..."))
   Call InitMainButton(cmdFileOutName, MapText("..."))
   
   Call InitMainButton(cmdSetting, MapText("ตั้งค่า"))

'   Call InitGrid1
'   Call ResetStatus
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
   Set m_collectionFile1 = New Collection
   Set m_collectionFile2 = New Collection
   Set m_collectionFile3 = New Collection
   Set m_collectionFile4 = New Collection
   Set m_collectionFile5 = New Collection
   Set m_collectionFile6 = New Collection
   Set m_collectionOutFile = New Collection

   Set m_xlsSetting = New CXlsEstimateSetting
   Set tempFarm = New CXlsFarm
   Set m_XlsFood = New Collection
   Set m_XlsSetFarm = New Collection
   
   Set m_ExcelApp = CreateObject("Excel.application")
   
   m_HasActivate = False
   Call InitFormLayout
   Call EnableForm(Me, True)
End Sub
'
'Private Sub InitGrid1()
'Dim Col As JSColumn
'
'   GridEX1.Columns.Clear
'   GridEX1.BackColor = GLB_GRID_COLOR
'   GridEX1.ItemCount = 0
'   GridEX1.BackColorHeader = GLB_GRIDHD_COLOR
'   GridEX1.ColumnHeaderFont.Bold = True
'   GridEX1.ColumnHeaderFont.Name = GLB_FONT
'   GridEX1.TabKeyBehavior = jgexControlNavigation
'
'   Set Col = GridEX1.Columns.Add '1
'   Col.Width = 0
'   Col.Caption = "ID"
'
'   Set Col = GridEX1.Columns.Add '2
'   Col.Width = 0
'   Col.Caption = "Real ID"
'
'   Set Col = GridEX1.Columns.Add '3
'   Col.Width = 1500
'   Col.Caption = MapText("รหัสหลัก")
'
'   Set Col = GridEX1.Columns.Add '4
'   Col.Width = 12000
'   Col.Caption = MapText("รหัสย่อย")
'
'End Sub

Private Sub Form_Unload(Cancel As Integer)

   Set m_collectionFile1 = Nothing
   Set m_collectionFile2 = Nothing
   Set m_collectionFile3 = Nothing
   Set m_collectionFile4 = Nothing
   Set m_collectionFile5 = Nothing
   Set m_collectionFile6 = Nothing
   Set m_collectionOutFile = Nothing

   Set m_xlsSetting = Nothing
   Set tempFarm = Nothing
   Set m_XlsFood = Nothing
   Set m_XlsSetFarm = Nothing
   
   Call m_ExcelApp.Workbooks.Close
   'Call m_ExcelApp.Close
End Sub
'Private Sub GridEX1_ColResize(ByVal ColIndex As Integer, ByVal NewColWidth As Long, ByVal Cancel As GridEX20.JSRetBoolean)
'   'debug.print ColIndex & " " & NewColWidth
'End Sub
'Private Sub GridEX1_DblClick()
'   Call cmdEdit_Click
'End Sub
'Private Sub GridEX1_UnboundReadData(ByVal RowIndex As Long, ByVal Bookmark As Variant, ByVal Values As GridEX20.JSRowData)
'On Error GoTo ErrorHandler
'Dim RealIndex As Long
'
'   glbErrorLog.ModuleName = Me.Name
'   glbErrorLog.RoutineName = "UnboundReadData"
'
'   If MainCollection Is Nothing Then
'      Exit Sub
'   End If
'
'   If RowIndex <= 0 Then
'      Exit Sub
'   End If
'
'   Dim CR As CAccountCode
'   If MainCollection.Count <= 0 Then
'      Exit Sub
'   End If
'   Set CR = GetItem(MainCollection, RowIndex, RealIndex)
'   If CR Is Nothing Then
'      Exit Sub
'   End If
'
'   Values(1) = CR.ACCOUNT_CODE_ID
'   Values(2) = RealIndex
'   Values(3) = CR.MAIN_CODE
'   Values(4) = CR.SUB_CODE
'
'Exit Sub
'
'ErrorHandler:
'   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
'   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
'End Sub
'Public Sub RefreshGrid()
'   GridEX1.ItemCount = CountItem(MainCollection)
'   GridEX1.Rebind
'End Sub
Private Sub QueryData()
Dim AG As CAccountConfig
Dim ItemCount As Long
Dim Rs As ADODB.Recordset
   
'   Set AG = New CAccountConfig
'   Set Rs = New ADODB.Recordset
'
'   AG.ACCOUNT_CONFIG_ID = -1
'   Call AG.QueryData(Rs, ItemCount)
   If ItemCount > 0 Then
      Call AG.PopulateFromRS(1, Rs)
      ShowMode = SHOW_EDIT
      
'      ConFigID = AG.ACCOUNT_CONFIG_ID
'      txtSheet.Text = AG.SHEET
'      txtRow.Text = AG.ROW
'      txtRow2.Text = AG.ROW2
'      txtCollumn.Text = AG.COLLUMN_CODE
'      txtCollumn2.Text = AG.COLLUMN_DEBIT
'      txtCollumn3.Text = AG.COLLUMN_CREDIT
'      txtCollumn4.Text = AG.COLLUMN_CHECK
'      uctlFromDate.ShowDate = AG.FROM_DATE
'      uctlToDate.ShowDate = AG.TO_DATE
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

'   AG.AddEditMode = ShowMode
'   AG.ACCOUNT_CONFIG_ID = ConFigID
'   AG.SHEET = Val(txtSheet.Text)
'   AG.ROW = Val(txtRow.Text)
'   AG.ROW2 = Val(txtRow2.Text)
'   AG.COLLUMN_CODE = Val(txtCollumn.Text)
'   AG.COLLUMN_DEBIT = Val(txtCollumn2.Text)
'   AG.COLLUMN_CREDIT = Val(txtCollumn3.Text)
'   AG.COLLUMN_CHECK = Val(txtCollumn4.Text)
'   AG.FROM_DATE = uctlFromDate.ShowDate
'   AG.TO_DATE = uctlToDate.ShowDate
'
'   Call AG.AddEditData
End Sub

