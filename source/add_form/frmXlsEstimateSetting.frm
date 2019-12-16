VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#2.0#0"; "THREED20.OCX"
Begin VB.Form frmXlsEstimateSetting 
   BackColor       =   &H80000000&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   6540
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9810
   Icon            =   "frmXlsEstimateSetting.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6540
   ScaleWidth      =   9810
   StartUpPosition =   1  'CenterOwner
   Begin Threed.SSFrame SSFrame1 
      Height          =   6525
      Left            =   0
      TabIndex        =   34
      Top             =   0
      Width           =   9825
      _ExtentX        =   17330
      _ExtentY        =   11509
      _Version        =   131073
      PictureBackgroundStyle=   2
      Begin Threed.SSPanel pnlHeader 
         Height          =   675
         Left            =   0
         TabIndex        =   36
         Top             =   0
         Width           =   9795
         _ExtentX        =   17277
         _ExtentY        =   1191
         _Version        =   131073
         PictureBackgroundStyle=   2
      End
      Begin prjLedgerReport.uctlTextBox txtCollumn1 
         Height          =   435
         Left            =   5880
         TabIndex        =   1
         Top             =   1200
         Width           =   600
         _ExtentX        =   1905
         _ExtentY        =   767
      End
      Begin prjLedgerReport.uctlTextBox txtRow1 
         Height          =   435
         Left            =   8160
         TabIndex        =   2
         Top             =   1200
         Width           =   600
         _ExtentX        =   1905
         _ExtentY        =   767
      End
      Begin prjLedgerReport.uctlTextBox txtSheet1 
         Height          =   435
         Left            =   3600
         TabIndex        =   0
         Top             =   1200
         Width           =   600
         _ExtentX        =   1905
         _ExtentY        =   767
      End
      Begin prjLedgerReport.uctlTextBox txtCollumn2 
         Height          =   435
         Left            =   5880
         TabIndex        =   4
         Top             =   1680
         Width           =   600
         _ExtentX        =   1905
         _ExtentY        =   767
      End
      Begin prjLedgerReport.uctlTextBox txtRow2 
         Height          =   435
         Left            =   8160
         TabIndex        =   5
         Top             =   1680
         Width           =   600
         _ExtentX        =   1905
         _ExtentY        =   767
      End
      Begin prjLedgerReport.uctlTextBox txtSheet2 
         Height          =   435
         Left            =   3600
         TabIndex        =   3
         Top             =   1680
         Width           =   600
         _ExtentX        =   1905
         _ExtentY        =   767
      End
      Begin prjLedgerReport.uctlTextBox txtCollumn3 
         Height          =   435
         Left            =   5880
         TabIndex        =   7
         Top             =   2160
         Width           =   600
         _ExtentX        =   1905
         _ExtentY        =   767
      End
      Begin prjLedgerReport.uctlTextBox txtRow3 
         Height          =   435
         Left            =   8160
         TabIndex        =   8
         Top             =   2160
         Width           =   600
         _ExtentX        =   1905
         _ExtentY        =   767
      End
      Begin prjLedgerReport.uctlTextBox txtSheet3 
         Height          =   435
         Left            =   3600
         TabIndex        =   6
         Top             =   2160
         Width           =   600
         _ExtentX        =   1905
         _ExtentY        =   767
      End
      Begin prjLedgerReport.uctlTextBox txtCollumn4 
         Height          =   435
         Left            =   5880
         TabIndex        =   10
         Top             =   2640
         Width           =   600
         _ExtentX        =   1905
         _ExtentY        =   767
      End
      Begin prjLedgerReport.uctlTextBox txtRow4 
         Height          =   435
         Left            =   8160
         TabIndex        =   11
         Top             =   2640
         Width           =   600
         _ExtentX        =   1905
         _ExtentY        =   767
      End
      Begin prjLedgerReport.uctlTextBox txtSheet4 
         Height          =   435
         Left            =   3600
         TabIndex        =   9
         Top             =   2640
         Width           =   600
         _ExtentX        =   1905
         _ExtentY        =   767
      End
      Begin prjLedgerReport.uctlTextBox txtsheet5 
         Height          =   435
         Left            =   3600
         TabIndex        =   12
         Top             =   3240
         Width           =   600
         _ExtentX        =   1905
         _ExtentY        =   767
      End
      Begin prjLedgerReport.uctlTextBox txtRow5 
         Height          =   435
         Left            =   5880
         TabIndex        =   13
         Top             =   3240
         Width           =   600
         _ExtentX        =   1905
         _ExtentY        =   767
      End
      Begin prjLedgerReport.uctlTextBox txtCollumn6_1 
         Height          =   435
         Left            =   3600
         TabIndex        =   16
         Top             =   4320
         Width           =   600
         _ExtentX        =   1905
         _ExtentY        =   767
      End
      Begin prjLedgerReport.uctlTextBox txtCollumn6_2 
         Height          =   435
         Left            =   4320
         TabIndex        =   17
         Top             =   4320
         Width           =   600
         _ExtentX        =   1905
         _ExtentY        =   767
      End
      Begin prjLedgerReport.uctlTextBox txtCollumn6_3 
         Height          =   435
         Left            =   5040
         TabIndex        =   18
         Top             =   4320
         Width           =   600
         _ExtentX        =   1905
         _ExtentY        =   767
      End
      Begin prjLedgerReport.uctlTextBox txtCollumn6_4 
         Height          =   435
         Left            =   5760
         TabIndex        =   19
         Top             =   4320
         Width           =   600
         _ExtentX        =   1905
         _ExtentY        =   767
      End
      Begin prjLedgerReport.uctlTextBox txtCollumn6_5 
         Height          =   435
         Left            =   6480
         TabIndex        =   20
         Top             =   4320
         Width           =   600
         _ExtentX        =   1905
         _ExtentY        =   767
      End
      Begin prjLedgerReport.uctlTextBox txtCollumn6_6 
         Height          =   435
         Left            =   7200
         TabIndex        =   21
         Top             =   4320
         Width           =   600
         _ExtentX        =   1905
         _ExtentY        =   767
      End
      Begin prjLedgerReport.uctlTextBox txtCollumn6_7 
         Height          =   435
         Left            =   7920
         TabIndex        =   22
         Top             =   4320
         Width           =   600
         _ExtentX        =   1905
         _ExtentY        =   767
      End
      Begin prjLedgerReport.uctlTextBox txtCollumn5_1 
         Height          =   435
         Left            =   3600
         TabIndex        =   23
         Top             =   4800
         Width           =   600
         _ExtentX        =   1905
         _ExtentY        =   767
      End
      Begin prjLedgerReport.uctlTextBox txtCollumn5_2 
         Height          =   435
         Left            =   4320
         TabIndex        =   24
         Top             =   4800
         Width           =   600
         _ExtentX        =   1905
         _ExtentY        =   767
      End
      Begin prjLedgerReport.uctlTextBox txtCollumn5_3 
         Height          =   435
         Left            =   5040
         TabIndex        =   25
         Top             =   4800
         Width           =   600
         _ExtentX        =   1905
         _ExtentY        =   767
      End
      Begin prjLedgerReport.uctlTextBox txtCollumn5_4 
         Height          =   435
         Left            =   5760
         TabIndex        =   26
         Top             =   4800
         Width           =   600
         _ExtentX        =   1905
         _ExtentY        =   767
      End
      Begin prjLedgerReport.uctlTextBox txtCollumn5_5 
         Height          =   435
         Left            =   6480
         TabIndex        =   27
         Top             =   4800
         Width           =   600
         _ExtentX        =   1905
         _ExtentY        =   767
      End
      Begin prjLedgerReport.uctlTextBox txtCollumn5_6 
         Height          =   435
         Left            =   7200
         TabIndex        =   28
         Top             =   4800
         Width           =   600
         _ExtentX        =   1905
         _ExtentY        =   767
      End
      Begin prjLedgerReport.uctlTextBox txtCollumn5_7 
         Height          =   435
         Left            =   7920
         TabIndex        =   29
         Top             =   4800
         Width           =   600
         _ExtentX        =   1905
         _ExtentY        =   767
      End
      Begin prjLedgerReport.uctlTextBox txtfromDataRow 
         Height          =   435
         Left            =   5880
         TabIndex        =   31
         Top             =   5280
         Width           =   600
         _ExtentX        =   1905
         _ExtentY        =   767
      End
      Begin prjLedgerReport.uctlTextBox txttoDataRow 
         Height          =   435
         Left            =   8160
         TabIndex        =   32
         Top             =   5280
         Width           =   600
         _ExtentX        =   1905
         _ExtentY        =   767
      End
      Begin prjLedgerReport.uctlTextBox txtCollumnFood 
         Height          =   435
         Left            =   3600
         TabIndex        =   30
         Top             =   5280
         Width           =   600
         _ExtentX        =   1905
         _ExtentY        =   767
      End
      Begin prjLedgerReport.uctlTextBox txtSheetBB 
         Height          =   435
         Left            =   3600
         TabIndex        =   14
         Top             =   3840
         Width           =   600
         _ExtentX        =   1905
         _ExtentY        =   767
      End
      Begin prjLedgerReport.uctlTextBox txtRowBB 
         Height          =   435
         Left            =   5880
         TabIndex        =   15
         Top             =   3840
         Width           =   600
         _ExtentX        =   1905
         _ExtentY        =   767
      End
      Begin VB.Label lblBB 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   435
         Left            =   120
         TabIndex        =   65
         Top             =   3840
         Width           =   2175
      End
      Begin VB.Label lblSheetBB 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   435
         Left            =   2520
         TabIndex        =   64
         Top             =   3840
         Width           =   975
      End
      Begin VB.Label lblRowBB 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   435
         Left            =   4800
         TabIndex        =   63
         Top             =   3840
         Width           =   975
      End
      Begin VB.Label lblCollumnFood 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   435
         Left            =   360
         TabIndex        =   62
         Top             =   5280
         Width           =   3135
      End
      Begin VB.Label lblfromDataRow 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   435
         Left            =   3960
         TabIndex        =   59
         Top             =   5280
         Width           =   1815
      End
      Begin VB.Label lbltoDataRow 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   435
         Left            =   6600
         TabIndex        =   58
         Top             =   5280
         Width           =   1455
      End
      Begin VB.Label lblCollumn5 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   435
         Left            =   360
         TabIndex        =   57
         Top             =   4800
         Width           =   3135
      End
      Begin VB.Label lblCollumn6 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   435
         Left            =   1560
         TabIndex        =   56
         Top             =   4320
         Width           =   1935
      End
      Begin VB.Label lblRow5 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   435
         Left            =   4800
         TabIndex        =   55
         Top             =   3240
         Width           =   975
      End
      Begin VB.Label lblsheet5 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   435
         Left            =   2520
         TabIndex        =   54
         Top             =   3240
         Width           =   975
      End
      Begin VB.Label lbldate 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   435
         Left            =   480
         TabIndex        =   53
         Top             =   3240
         Width           =   1815
      End
      Begin VB.Label lblCollumn4 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   435
         Left            =   4560
         TabIndex        =   52
         Top             =   2640
         Width           =   1215
      End
      Begin VB.Label lblRow4 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   435
         Left            =   6840
         TabIndex        =   51
         Top             =   2640
         Width           =   1215
      End
      Begin VB.Label lblSheet4 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   435
         Left            =   2520
         TabIndex        =   50
         Top             =   2640
         Width           =   975
      End
      Begin VB.Label lbltodate 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   435
         Left            =   480
         TabIndex        =   49
         Top             =   2640
         Width           =   1815
      End
      Begin VB.Label lblCollumn3 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   435
         Left            =   4560
         TabIndex        =   48
         Top             =   2160
         Width           =   1215
      End
      Begin VB.Label lblRow3 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   435
         Left            =   6840
         TabIndex        =   47
         Top             =   2160
         Width           =   1215
      End
      Begin VB.Label lblSheet3 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   435
         Left            =   2520
         TabIndex        =   46
         Top             =   2160
         Width           =   975
      End
      Begin VB.Label lblfromdate 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   435
         Left            =   480
         TabIndex        =   45
         Top             =   2160
         Width           =   1815
      End
      Begin VB.Label lblCollumn2 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   435
         Left            =   4560
         TabIndex        =   44
         Top             =   1680
         Width           =   1215
      End
      Begin VB.Label lblRow2 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   435
         Left            =   6840
         TabIndex        =   43
         Top             =   1680
         Width           =   1215
      End
      Begin VB.Label lblSheet2 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   435
         Left            =   2520
         TabIndex        =   42
         Top             =   1680
         Width           =   975
      End
      Begin VB.Label lblweek2 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   435
         Left            =   480
         TabIndex        =   41
         Top             =   1680
         Width           =   1815
      End
      Begin VB.Label lblCollumn1 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   435
         Left            =   4560
         TabIndex        =   40
         Top             =   1200
         Width           =   1215
      End
      Begin VB.Label lblRow1 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   435
         Left            =   6840
         TabIndex        =   39
         Top             =   1200
         Width           =   1215
      End
      Begin VB.Label lblSheet1 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   435
         Left            =   2520
         TabIndex        =   38
         Top             =   1200
         Width           =   975
      End
      Begin VB.Label lblfarmName1 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   435
         Left            =   480
         TabIndex        =   37
         Top             =   1200
         Width           =   1815
      End
      Begin Threed.SSCommand cmdExit 
         Cancel          =   -1  'True
         Height          =   525
         Left            =   5040
         TabIndex        =   35
         Top             =   5880
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   926
         _Version        =   131073
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdOK 
         Height          =   525
         Left            =   3240
         TabIndex        =   33
         Top             =   5880
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmXlsEstimateSetting.frx":27A2
         ButtonStyle     =   3
      End
   End
   Begin prjLedgerReport.uctlTextBox uctlTextBox1 
      Height          =   435
      Left            =   3240
      TabIndex        =   60
      Top             =   0
      Width           =   600
      _ExtentX        =   1905
      _ExtentY        =   767
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "Label1"
      Height          =   435
      Left            =   0
      TabIndex        =   61
      Top             =   0
      Width           =   3135
   End
End
Attribute VB_Name = "frmXlsEstimateSetting"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private m_HasActivate As Boolean
Private m_HasModify As Boolean
Private m_Rs As ADODB.Recordset
Private m_EstSetting As CXlsEstimateSetting

Public ID As Long
Public OKClick As Boolean
Public ShowMode As SHOW_MODE_TYPE
Public HeaderText As String
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
      
      m_EstSetting.XLS_EST_SET_ID = 1       '‡ªÁπ ID =1 ‡ ¡Õ‰¥È¡—È¬
      m_EstSetting.QueryFlag = 1
      If Not glbDaily.QueryXlsSetting(m_EstSetting, m_Rs, ItemCount, IsOK, glbErrorLog) Then
         glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
         Call EnableForm(Me, True)
         Exit Sub
      End If
   End If
   
   If ItemCount > 0 Then
      Call m_EstSetting.PopulateFromRS(1, m_Rs)
      
      txtSheet1.Text = m_EstSetting.SHEET_1
      txtSheet2.Text = m_EstSetting.SHEET_2
      txtSheet3.Text = m_EstSetting.SHEET_3
      txtSheet4.Text = m_EstSetting.SHEET_4
      txtCollumn1.Text = m_EstSetting.COLLUMN_1
      txtCollumn2.Text = m_EstSetting.COLLUMN_2
      txtCollumn3.Text = m_EstSetting.COLLUMN_3
      txtCollumn4.Text = m_EstSetting.COLLUMN_4
      txtRow1.Text = m_EstSetting.ROW_1
      txtRow2.Text = m_EstSetting.ROW_2
      txtRow3.Text = m_EstSetting.ROW_3
      txtRow4.Text = m_EstSetting.ROW_4
      
      txtfromDataRow.Text = m_EstSetting.FROMDATAROW
      txttoDataRow.Text = m_EstSetting.TODATAROW
      
      txtsheet5.Text = m_EstSetting.SHEET_5
      txtRow5.Text = m_EstSetting.ROW_5
      txtCollumn5_1.Text = m_EstSetting.COLLUMN5_1
      txtCollumn5_2.Text = m_EstSetting.COLLUMN5_2
      txtCollumn5_3.Text = m_EstSetting.COLLUMN5_3
      txtCollumn5_4.Text = m_EstSetting.COLLUMN5_4
      txtCollumn5_5.Text = m_EstSetting.COLLUMN5_5
      txtCollumn5_6.Text = m_EstSetting.COLLUMN5_6
      txtCollumn5_7.Text = m_EstSetting.COLLUMN5_7
      
      txtCollumn6_1.Text = m_EstSetting.COLLUMN6_1
      txtCollumn6_2.Text = m_EstSetting.COLLUMN6_2
      txtCollumn6_3.Text = m_EstSetting.COLLUMN6_3
      txtCollumn6_4.Text = m_EstSetting.COLLUMN6_4
      txtCollumn6_5.Text = m_EstSetting.COLLUMN6_5
      txtCollumn6_6.Text = m_EstSetting.COLLUMN6_6
      txtCollumn6_7.Text = m_EstSetting.COLLUMN6_7
      
      txtCollumnFood.Text = m_EstSetting.COLLUMNFOOD
      txtSheetBB.Text = m_EstSetting.SHEETBB
      txtRowBB.Text = m_EstSetting.ROWBB
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
   
   If Not VerifyTextControl(lblfarmName1, txtSheet1, False) Then
      Exit Function
   End If
   If Not VerifyTextControl(lblfarmName1, txtCollumn1, False) Then
      Exit Function
   End If
   If Not VerifyTextControl(lblfarmName1, txtRow1, False) Then
      Exit Function
   End If
   
   If Not VerifyTextControl(lblweek2, txtSheet2, False) Then
      Exit Function
   End If
   If Not VerifyTextControl(lblweek2, txtCollumn2, False) Then
      Exit Function
   End If
   If Not VerifyTextControl(lblweek2, txtRow2, False) Then
      Exit Function
   End If
   
      If Not VerifyTextControl(lblFromDate, txtSheet3, False) Then
      Exit Function
   End If
   If Not VerifyTextControl(lblFromDate, txtCollumn3, False) Then
      Exit Function
   End If
   If Not VerifyTextControl(lblFromDate, txtRow3, False) Then
      Exit Function
   End If
   
   If Not VerifyTextControl(lblToDate, txtSheet4, False) Then
      Exit Function
   End If
   If Not VerifyTextControl(lblToDate, txtCollumn4, False) Then
      Exit Function
   End If
   If Not VerifyTextControl(lblToDate, txtRow4, False) Then
      Exit Function
   End If
   
    If Not VerifyTextControl(lblfromDataRow, txtfromDataRow, False) Then
      Exit Function
   End If
   If Not VerifyTextControl(lbltoDataRow, txttoDataRow, False) Then
      Exit Function
   End If
   If Not VerifyTextControl(lblCollumnFood, txtCollumnFood, False) Then
      Exit Function
   End If
   
   If Not VerifyTextControl(lblSheetBB, txtSheetBB, False) Then
      Exit Function
   End If
   If Not VerifyTextControl(lblRowBB, txtRowBB, False) Then
      Exit Function
   End If
   
   If Not m_HasModify Then
      SaveData = True
      Exit Function
   End If
   
   m_EstSetting.AddEditMode = SHOW_EDIT
   m_EstSetting.XLS_EST_SET_ID = 1
   m_EstSetting.SHEET_1 = txtSheet1.Text
   m_EstSetting.SHEET_2 = txtSheet2.Text
   m_EstSetting.SHEET_3 = txtSheet3.Text
   m_EstSetting.SHEET_4 = txtSheet4.Text
   m_EstSetting.COLLUMN_1 = txtCollumn1.Text
   m_EstSetting.COLLUMN_2 = txtCollumn2.Text
   m_EstSetting.COLLUMN_3 = txtCollumn3.Text
   m_EstSetting.COLLUMN_4 = txtCollumn4.Text
   m_EstSetting.ROW_1 = txtRow1.Text
   m_EstSetting.ROW_2 = txtRow2.Text
   m_EstSetting.ROW_3 = txtRow3.Text
   m_EstSetting.ROW_4 = txtRow4.Text
   m_EstSetting.FROMDATAROW = txtfromDataRow.Text
   m_EstSetting.TODATAROW = txttoDataRow.Text
   
   m_EstSetting.SHEET_5 = txtsheet5.Text
   m_EstSetting.ROW_5 = txtRow5.Text
   m_EstSetting.COLLUMN5_1 = txtCollumn5_1.Text
   m_EstSetting.COLLUMN5_2 = txtCollumn5_2.Text
   m_EstSetting.COLLUMN5_3 = txtCollumn5_3.Text
   m_EstSetting.COLLUMN5_4 = txtCollumn5_4.Text
   m_EstSetting.COLLUMN5_5 = txtCollumn5_5.Text
   m_EstSetting.COLLUMN5_6 = txtCollumn5_6.Text
   m_EstSetting.COLLUMN5_7 = txtCollumn5_7.Text

   m_EstSetting.COLLUMN6_1 = txtCollumn6_1.Text
   m_EstSetting.COLLUMN6_2 = txtCollumn6_2.Text
   m_EstSetting.COLLUMN6_3 = txtCollumn6_3.Text
   m_EstSetting.COLLUMN6_4 = txtCollumn6_4.Text
   m_EstSetting.COLLUMN6_5 = txtCollumn6_5.Text
   m_EstSetting.COLLUMN6_6 = txtCollumn6_6.Text
   m_EstSetting.COLLUMN6_7 = txtCollumn6_7.Text
   
   m_EstSetting.COLLUMNFOOD = txtCollumnFood.Text
   m_EstSetting.SHEETBB = txtSheetBB.Text
   m_EstSetting.ROWBB = txtRowBB.Text
      
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
Private Sub Form_Activate()
   If Not m_HasActivate Then
      m_HasActivate = True
      Me.Refresh
      DoEvents
      
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
   
   Me.Caption = MapText("µ—Èß§Ë“ ·∂«,§Õ≈—¡ªÏ")
   pnlHeader.Caption = Me.Caption
   
   Call InitNormalLabel(lblfarmName1, MapText("™◊ËÕø“√Ï¡"))
   Call InitNormalLabel(lblweek2, MapText(" —ª¥“ÀÏ∑’Ë"))
   Call InitNormalLabel(lblFromDate, MapText("√–À«Ë“ß«—π∑’Ë"))
   Call InitNormalLabel(lblToDate, MapText("∂÷ß«—π∑’Ë"))
   Call InitNormalLabel(lbldate, MapText("™ÿ¥«—π∑’Ë"))
   
   Call InitNormalLabel(lblSheet1, MapText("™’∑∑’Ë"))
   Call InitNormalLabel(lblSheet2, MapText("™’∑∑’Ë"))
   Call InitNormalLabel(lblSheet3, MapText("™’∑∑’Ë"))
   Call InitNormalLabel(lblSheet4, MapText("™’∑∑’Ë"))
   Call InitNormalLabel(lblsheet5, MapText("™’∑∑’Ë"))
   Call InitNormalLabel(lblCollumn1, MapText("§Õ≈—¡ªÏ∑’Ë"))
   Call InitNormalLabel(lblCollumn2, MapText("§Õ≈—¡ªÏ∑’Ë"))
   Call InitNormalLabel(lblCollumn3, MapText("§Õ≈—¡ªÏ∑’Ë"))
   Call InitNormalLabel(lblCollumn4, MapText("§Õ≈—¡ªÏ∑’Ë"))
   Call InitNormalLabel(lblCollumn5, MapText("™ÿ¥«—π∑’Ë·≈–§Õ≈—¡ªÏ BULK"))
   Call InitNormalLabel(lblCollumn6, MapText("™ÿ¥§Õ≈—¡ªÏ BAG"))
   Call InitNormalLabel(lblRow1, MapText("·∂«∑’Ë"))
   Call InitNormalLabel(lblRow2, MapText("·∂«∑’Ë"))
   Call InitNormalLabel(lblRow3, MapText("·∂«∑’Ë"))
   Call InitNormalLabel(lblRow4, MapText("·∂«∑’Ë"))
   Call InitNormalLabel(lblRow5, MapText("·∂«∑’Ë"))
   
   Call InitNormalLabel(lblfromDataRow, MapText("®“°·∂«¢ÈÕ¡Ÿ≈"))
   Call InitNormalLabel(lbltoDataRow, MapText("∂÷ß·∂«¢ÈÕ¡Ÿ≈"))
   Call InitNormalLabel(lblCollumnFood, MapText("‡∫Õ√ÏÕ“À“√"))
   
  Call InitNormalLabel(lblBB, MapText("À—«·∂« Bag , Bulk"))
  Call InitNormalLabel(lblSheetBB, MapText("™’∑∑’Ë"))
  Call InitNormalLabel(lblRowBB, MapText("·∂«∑’Ë"))
  
  Call txtSheet1.SetTextLenType(TEXT_FLOAT, glbSetting.MONEY_TYPE)
  Call txtRow1.SetTextLenType(TEXT_FLOAT, glbSetting.MONEY_TYPE)
  Call txtCollumn1.SetTextLenType(TEXT_STRING, glbSetting.CODE_TYPE)
  Call txtSheet2.SetTextLenType(TEXT_FLOAT, glbSetting.MONEY_TYPE)
  Call txtRow2.SetTextLenType(TEXT_FLOAT, glbSetting.MONEY_TYPE)
  Call txtCollumn2.SetTextLenType(TEXT_STRING, glbSetting.CODE_TYPE)
  Call txtSheet3.SetTextLenType(TEXT_FLOAT, glbSetting.MONEY_TYPE)
  Call txtRow3.SetTextLenType(TEXT_FLOAT, glbSetting.MONEY_TYPE)
  Call txtCollumn3.SetTextLenType(TEXT_STRING, glbSetting.CODE_TYPE)
  Call txtSheet4.SetTextLenType(TEXT_FLOAT, glbSetting.MONEY_TYPE)
  Call txtRow4.SetTextLenType(TEXT_FLOAT, glbSetting.MONEY_TYPE)
  Call txtCollumn4.SetTextLenType(TEXT_STRING, glbSetting.CODE_TYPE)
  
    Call txtsheet5.SetTextLenType(TEXT_FLOAT, glbSetting.MONEY_TYPE)
    Call txtRow5.SetTextLenType(TEXT_FLOAT, glbSetting.MONEY_TYPE)
    Call txtCollumn5_1.SetTextLenType(TEXT_STRING, glbSetting.CODE_TYPE)
    Call txtCollumn5_2.SetTextLenType(TEXT_STRING, glbSetting.CODE_TYPE)
    Call txtCollumn5_3.SetTextLenType(TEXT_STRING, glbSetting.CODE_TYPE)
    Call txtCollumn5_4.SetTextLenType(TEXT_STRING, glbSetting.CODE_TYPE)
    Call txtCollumn5_5.SetTextLenType(TEXT_STRING, glbSetting.CODE_TYPE)
    Call txtCollumn5_6.SetTextLenType(TEXT_STRING, glbSetting.CODE_TYPE)
    Call txtCollumn5_7.SetTextLenType(TEXT_STRING, glbSetting.CODE_TYPE)
    
    Call txtCollumn6_1.SetTextLenType(TEXT_STRING, glbSetting.CODE_TYPE)
    Call txtCollumn6_2.SetTextLenType(TEXT_STRING, glbSetting.CODE_TYPE)
    Call txtCollumn6_3.SetTextLenType(TEXT_STRING, glbSetting.CODE_TYPE)
    Call txtCollumn6_4.SetTextLenType(TEXT_STRING, glbSetting.CODE_TYPE)
    Call txtCollumn6_5.SetTextLenType(TEXT_STRING, glbSetting.CODE_TYPE)
    Call txtCollumn6_6.SetTextLenType(TEXT_STRING, glbSetting.CODE_TYPE)
    Call txtCollumn6_7.SetTextLenType(TEXT_STRING, glbSetting.CODE_TYPE)
    
    Call txtCollumnFood.SetTextLenType(TEXT_STRING, glbSetting.CODE_TYPE)
    Call txtfromDataRow.SetTextLenType(TEXT_STRING, glbSetting.CODE_TYPE)
    Call txttoDataRow.SetTextLenType(TEXT_STRING, glbSetting.CODE_TYPE)
   
    Call txtSheetBB.SetTextLenType(TEXT_FLOAT, glbSetting.MONEY_TYPE)
    Call txtRowBB.SetTextLenType(TEXT_FLOAT, glbSetting.MONEY_TYPE)
  
   Me.Picture = LoadPicture(glbParameterObj.MainPicture)
   
   pnlHeader.Font.Name = GLB_FONT
   pnlHeader.Font.Bold = True
   pnlHeader.Font.Size = 19
   
   cmdExit.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdOK.Picture = LoadPicture(glbParameterObj.NormalButton1)
   
   Call InitMainButton(cmdExit, MapText("¬°‡≈‘° (ESC)"))
   Call InitMainButton(cmdOK, MapText("µ°≈ß (F2)"))
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
   
   Set m_EstSetting = New CXlsEstimateSetting
   Set m_Rs = New ADODB.Recordset
   
   m_HasActivate = False
   Call InitFormLayout
   Call EnableForm(Me, True)
End Sub
Private Sub Form_Unload(Cancel As Integer)
   Set m_EstSetting = Nothing
End Sub

Private Sub txtsheet1_Change()
   m_HasModify = True
End Sub
Private Sub txtsheet2_Change()
   m_HasModify = True
End Sub
Private Sub txtsheet3_Change()
   m_HasModify = True
End Sub
Private Sub txtsheet4_Change()
   m_HasModify = True
End Sub
Private Sub txtCollumn1_Change()
   m_HasModify = True
End Sub
Private Sub txtCollumn2_Change()
   m_HasModify = True
End Sub
Private Sub txtCollumn3_Change()
   m_HasModify = True
End Sub
Private Sub txtCollumn4_Change()
   m_HasModify = True
End Sub
Private Sub txtRow1_Change()
   m_HasModify = True
End Sub
Private Sub txtRowt2_Change()
   m_HasModify = True
End Sub
Private Sub txtRow3_Change()
   m_HasModify = True
End Sub
Private Sub txtRow4_Change()
   m_HasModify = True
End Sub
Private Sub txtsheet5_Change()
   m_HasModify = True
End Sub
Private Sub txtRow5_Change()
   m_HasModify = True
End Sub
Private Sub txtCollumn5_1_Change()
   m_HasModify = True
End Sub
Private Sub txtCollumn5_2_Change()
   m_HasModify = True
End Sub
Private Sub txtCollumn5_3_Change()
   m_HasModify = True
End Sub
Private Sub txtCollumn5_4_Change()
   m_HasModify = True
End Sub
Private Sub txtCollumn5_5_Change()
   m_HasModify = True
End Sub
Private Sub txtCollumn5_6_Change()
   m_HasModify = True
End Sub
Private Sub txtCollumn5_7_Change()
   m_HasModify = True
End Sub
Private Sub txtCollumn6_1_Change()
   m_HasModify = True
End Sub
Private Sub txtCollumn6_2_Change()
   m_HasModify = True
End Sub
Private Sub txtCollumn6_3_Change()
   m_HasModify = True
End Sub
Private Sub txtCollumn6_4_Change()
   m_HasModify = True
End Sub
Private Sub txtCollumn6_5_Change()
   m_HasModify = True
End Sub
Private Sub txtCollumn6_6_Change()
   m_HasModify = True
End Sub
Private Sub txtCollumn6_7_Change()
   m_HasModify = True
End Sub
Private Sub txtfromDataRow_Change()
   m_HasModify = True
End Sub
Private Sub txttoDataRow_Change()
   m_HasModify = True
End Sub
Private Sub txtCollumnFood_Change()
   m_HasModify = True
End Sub
Private Sub txtSheetBB_Change()
   m_HasModify = True
End Sub
Private Sub txtRowBB_Change()
   m_HasModify = True
End Sub
