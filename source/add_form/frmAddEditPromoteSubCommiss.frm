VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#2.0#0"; "THREED20.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{E684D8A3-716C-4E59-AA94-7144C04B0074}#1.1#0"; "GridEX20.ocx"
Begin VB.Form frmAddEditPromoteSubCommiss 
   ClientHeight    =   8490
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11880
   Icon            =   "frmAddEditPromoteSubCommiss.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   Moveable        =   0   'False
   NegotiateMenus  =   0   'False
   ScaleHeight     =   8490
   ScaleWidth      =   11880
   StartUpPosition =   1  'CenterOwner
   WindowState     =   2  'Maximized
   Begin Threed.SSFrame SSFrame1 
      Height          =   8520
      Left            =   0
      TabIndex        =   11
      Top             =   0
      Width           =   11895
      _ExtentX        =   20981
      _ExtentY        =   15028
      _Version        =   131073
      PictureBackgroundStyle=   2
      Begin MSComDlg.CommonDialog dlgAdd 
         Left            =   0
         Top             =   0
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin Threed.SSPanel pnlHeader 
         Height          =   705
         Left            =   0
         TabIndex        =   12
         Top             =   0
         Width           =   11925
         _ExtentX        =   21034
         _ExtentY        =   1244
         _Version        =   131073
         PictureBackgroundStyle=   2
      End
      Begin prjLedgerReport.uctlDate uctlFromDate 
         Height          =   375
         Left            =   3840
         TabIndex        =   2
         Top             =   2760
         Width           =   3855
         _ExtentX        =   6800
         _ExtentY        =   661
      End
      Begin prjLedgerReport.uctlTextBox txtMasCommissNo 
         Height          =   450
         Left            =   3840
         TabIndex        =   0
         Top             =   840
         Width           =   2895
         _ExtentX        =   5106
         _ExtentY        =   873
      End
      Begin prjLedgerReport.uctlDate uctlToDate 
         Height          =   375
         Left            =   10080
         TabIndex        =   3
         Top             =   2760
         Width           =   3855
         _ExtentX        =   6800
         _ExtentY        =   661
      End
      Begin MSComctlLib.TabStrip TabStrip1 
         Height          =   555
         Left            =   120
         TabIndex        =   4
         Top             =   3240
         Width           =   11640
         _ExtentX        =   20532
         _ExtentY        =   979
         _Version        =   393216
         BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
            NumTabs         =   1
            BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               ImageVarType    =   2
            EndProperty
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "AngsanaUPC"
            Size            =   14.25
            Charset         =   222
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin GridEX20.GridEX GridEX1 
         Height          =   3885
         Left            =   120
         TabIndex        =   5
         Top             =   3720
         Width           =   11595
         _ExtentX        =   20452
         _ExtentY        =   6853
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
         Column(1)       =   "frmAddEditPromoteSubCommiss.frx":27A2
         Column(2)       =   "frmAddEditPromoteSubCommiss.frx":286A
         FormatStylesCount=   5
         FormatStyle(1)  =   "frmAddEditPromoteSubCommiss.frx":290E
         FormatStyle(2)  =   "frmAddEditPromoteSubCommiss.frx":2A6A
         FormatStyle(3)  =   "frmAddEditPromoteSubCommiss.frx":2B1A
         FormatStyle(4)  =   "frmAddEditPromoteSubCommiss.frx":2BCE
         FormatStyle(5)  =   "frmAddEditPromoteSubCommiss.frx":2CA6
         ImageCount      =   0
         PrinterProperties=   "frmAddEditPromoteSubCommiss.frx":2D5E
      End
      Begin prjLedgerReport.uctlTextBox txtMasterCommissDesc 
         Height          =   450
         Left            =   3840
         TabIndex        =   1
         Top             =   2280
         Width           =   6015
         _ExtentX        =   10610
         _ExtentY        =   794
      End
      Begin prjLedgerReport.uctlTextLookup uctlCustomerLookup 
         Height          =   375
         Left            =   3840
         TabIndex        =   17
         Top             =   1800
         Width           =   6015
         _ExtentX        =   11033
         _ExtentY        =   661
      End
      Begin prjLedgerReport.uctlTextLookup uctlCommissionSale 
         Height          =   375
         Left            =   3840
         TabIndex        =   19
         Top             =   1320
         Width           =   6015
         _ExtentX        =   10610
         _ExtentY        =   661
      End
      Begin VB.Label lblSaleName 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   435
         Left            =   1560
         TabIndex        =   20
         Top             =   1320
         Width           =   1935
      End
      Begin VB.Label lblCustomer 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   375
         Left            =   1920
         TabIndex        =   18
         Top             =   1800
         Width           =   1575
      End
      Begin VB.Label lblMasCommissDesc 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   2160
         TabIndex        =   16
         Top             =   2280
         Width           =   1365
      End
      Begin Threed.SSCommand cmdExit 
         Cancel          =   -1  'True
         Height          =   525
         Left            =   10080
         TabIndex        =   10
         Top             =   7800
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   926
         _Version        =   131073
         ButtonStyle     =   3
      End
      Begin VB.Label lblFromDate 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   435
         Left            =   1680
         TabIndex        =   15
         Top             =   2760
         Width           =   1755
      End
      Begin VB.Label lblToDate 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   435
         Left            =   8400
         TabIndex        =   14
         Top             =   2760
         Width           =   1365
      End
      Begin VB.Label lblMasCommissNo 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   435
         Left            =   1800
         TabIndex        =   13
         Top             =   840
         Width           =   1695
      End
      Begin Threed.SSCommand cmdOK 
         Height          =   525
         Left            =   8400
         TabIndex        =   9
         Top             =   7800
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmAddEditPromoteSubCommiss.frx":2F36
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdEdit 
         Height          =   525
         Left            =   1770
         TabIndex        =   7
         Top             =   7830
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   926
         _Version        =   131073
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdAdd 
         Height          =   525
         Left            =   150
         TabIndex        =   6
         Top             =   7830
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmAddEditPromoteSubCommiss.frx":3250
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdDelete 
         Height          =   525
         Left            =   3420
         TabIndex        =   8
         Top             =   7830
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmAddEditPromoteSubCommiss.frx":356A
         ButtonStyle     =   3
      End
   End
End
Attribute VB_Name = "frmAddEditPromoteSubCommiss"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Const ROOT_TREE = "Root"
Private m_HasActivate As Boolean
Private m_HasModify As Boolean
Private m_Rs As ADODB.Recordset

Private m_PromoteYear As CCommissMasterPromote
Private masterSub As CComMasSubPromote
Private cm5_Rs As ADODB.Recordset    '' Commiss
 
Public HeaderText As String
Public ShowMode As SHOW_MODE_TYPE
Public OKClick As Boolean

Public Commiss_PROMOTE_ID As Long
Public Commiss_TYP As String
Public STKCOD As String
Public STKDES As String
Public SLM_PERCENT As Long
Public MASTER_Commiss_ID As Long
Public MASTER_Commiss_NO As String
Public Flag As String

Public VALID_FROM As String
Public VALID_TO As String

Dim ItemCount As Long
Dim itemCountGrid5 As Long     ' incen

Private m_TableName As String
Private FileName As String
Private m_SumUnit As Double
Private m_Customer As Collection
Private FtSaleColl As Collection

Public Sub QueryData(Flag As Boolean)
Dim IsOK As Boolean

                     IsOK = True
                     If Flag Then
                            Call EnableForm(Me, False)
                            
                           If MASTER_Commiss_ID = 0 Then                                 ' กรณี Add มันต้องหาก่อนว่า เจนคีย์ล่าสุดคืออะไร
                               If Not glbDaily.QueryMasterCommiss(m_PromoteYear, m_Rs, ItemCount, IsOK, glbErrorLog) Then
                                  glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
                                  Call EnableForm(Me, True)
                                  Exit Sub
                                End If
                                 Call m_PromoteYear.PopulateFromRS(1, m_Rs)
                                 MASTER_Commiss_ID = m_PromoteYear.MASTER_Commiss_ID
                            End If
                            
                            m_PromoteYear.MASTER_Commiss_ID = MASTER_Commiss_ID
                            m_PromoteYear.VALID_FROM = -1
                            m_PromoteYear.VALID_TO = -1
                            If Not glbDaily.QueryMasterCommiss(m_PromoteYear, m_Rs, ItemCount, IsOK, glbErrorLog) Then
                                  glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
                                  Call EnableForm(Me, True)
                                  Exit Sub
                            End If
                     End If

                     If ItemCount > 0 Then
                        Call m_PromoteYear.PopulateFromRS(1, m_Rs)
                        txtMasCommissNo.Text = m_PromoteYear.MASTER_Commiss_NO    'ยังอยู่ในโหมด edit
                       uctlFromDate.ShowDate = m_PromoteYear.VALID_FROM
                         uctlToDate.ShowDate = m_PromoteYear.VALID_TO
                         uctlCommissionSale.MyCombo.Text = m_PromoteYear.SLM_NAME
                         uctlCustomerLookup.MyCombo.Text = m_PromoteYear.CUS_NAME
                         uctlCommissionSale.MyTextBox.Text = m_PromoteYear.SLM_ID
                         uctlCustomerLookup.MyTextBox.Text = m_PromoteYear.CUS_ID
                        txtMasterCommissDesc.Text = m_PromoteYear.MASTER_COMMISS_DESC
                     End If

                  If Not IsOK Then
                     glbErrorLog.ShowUserError
                     Call EnableForm(Me, True)
                     Exit Sub
                  End If
   
   Call TabStrip1_Click
   Call EnableForm(Me, True)
End Sub

Private Function SaveData() As Boolean
Dim IsOK As Boolean

   If Not VerifyTextControl(lblMasCommissNo, txtMasCommissNo, False) Then
      Exit Function
   End If
   If Not VerifyDate(lblFromDate, uctlFromDate, False) Then
      Exit Function
   End If
   If Not VerifyDate(lblToDate, uctlToDate, False) Then
      Exit Function
   End If
   
      If Not (Len(uctlCommissionSale.MyTextBox.Text) > 0) Then
      Exit Function
   End If
   
      If Not (Len(uctlCommissionSale.MyTextBox.Text) > 0) Then
       Exit Function
   End If
   
   If Not CheckUniqueNs(EXPORT_UNIQUE, txtMasCommissNo.Text, MASTER_Commiss_ID) Then
      glbErrorLog.LocalErrorMsg = MapText("มีข้อมูล") & " " & txtMasCommissNo.Text & " " & MapText("อยู่ในระบบแล้ว")
      glbErrorLog.ShowUserError
      Call txtMasCommissNo.SetFocus
      Exit Function
   End If
   
   If Not m_HasModify Then
      SaveData = True
      Exit Function
   End If
   

'If MASTER_Commiss_ID = 0 Then
m_PromoteYear.ShowMode = ShowMode                        ' ตรงนี้ ตอนบันทึก
'Else:
' m_PromoteYear.ShowMode = SHOW_EDIT
' End If
    m_PromoteYear.MASTER_Commiss_ID = MASTER_Commiss_ID
    m_PromoteYear.MASTER_Commiss_NO = txtMasCommissNo.Text
        m_PromoteYear.SLM_ID = uctlCommissionSale.MyTextBox.Text
      m_PromoteYear.SLM_NAME = uctlCommissionSale.MyCombo.Text
    m_PromoteYear.CUS_ID = uctlCustomerLookup.MyTextBox.Text
      m_PromoteYear.CUS_NAME = uctlCustomerLookup.MyCombo.Text
    m_PromoteYear.VALID_FROM = uctlFromDate.ShowDate
    m_PromoteYear.VALID_TO = uctlToDate.ShowDate
    m_PromoteYear.MASTER_COMMISS_DESC = txtMasterCommissDesc.Text
   
   Call EnableForm(Me, False)
   
   If Not glbDaily.AddEditMasterCommiss(m_PromoteYear, IsOK, True, glbErrorLog) Then
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

Private Sub cmdAdd_Click()    ' เมื่อกดเพิ่มแต่ละ case 3 case
Dim OKClick As Boolean
Dim IsOK As Boolean

If Not cmdAdd.Enabled Then
      Exit Sub
   End If
   
  If m_PromoteYear.MASTER_Commiss_ID <= 0 Then
      glbErrorLog.LocalErrorMsg = MapText("กรุณาทำการบันทึกข้อมูลให้เรียบร้อยก่อน")
      glbErrorLog.ShowUserError
      Exit Sub
   End If
   
   Call EnableForm(Me, False)
   
   
   OKClick = False
   
If TabStrip1.SelectedItem.index = 1 Then
      frmAddEditPromoteCommiss.ShowMode = SHOW_ADD
      frmAddEditPromoteCommiss.HeaderText = MapText("เพิ่ม Com ขาย (พิเศษ) ")
      frmAddEditPromoteCommiss.Credit_text = "ยอดขาย (%)"
      frmAddEditPromoteCommiss.MASTER_Commiss_ID = MASTER_Commiss_ID
      frmAddEditPromoteCommiss.CREDIT_TYP = "01"
               frmAddEditPromoteCommiss.DB_TYP = 1
          frmAddEditPromoteCommiss.COM_KEY = 1    ' เพื่อฟอร์มย่อยเล็กสุด
      Load frmAddEditPromoteCommiss
     frmAddEditPromoteCommiss.Show 1
     
  ElseIf TabStrip1.SelectedItem.index = 2 Then
      frmAddEditPromoteCommiss.ShowMode = SHOW_ADD
      frmAddEditPromoteCommiss.HeaderText = MapText("เพิ่ม Com เก็บเงิน (พิเศษ) ")
      frmAddEditPromoteCommiss.Credit_text = "เครดิตภายใน (วัน)"
      frmAddEditPromoteCommiss.MASTER_Commiss_ID = MASTER_Commiss_ID
      frmAddEditPromoteCommiss.CREDIT_TYP = "02"
               frmAddEditPromoteCommiss.DB_TYP = 2
                frmAddEditPromoteCommiss.COM_KEY = 2   ' เพื่อฟอร์มย่อยเล็กสุด
      Load frmAddEditPromoteCommiss
     frmAddEditPromoteCommiss.Show 1
      
   ElseIf TabStrip1.SelectedItem.index = 3 Then
      frmAddEditPromoteCommiss.ShowMode = SHOW_ADD
      frmAddEditPromoteCommiss.HeaderText = MapText("เพิ่ม Commiss  (พิเศษ) - สินค้า Commodity")
      frmAddEditPromoteCommiss.Credit_text = "เครดิตภายใน (วัน)"
      frmAddEditPromoteCommiss.MASTER_Commiss_ID = MASTER_Commiss_ID
      frmAddEditPromoteCommiss.CREDIT_TYP = "03"
               frmAddEditPromoteCommiss.DB_TYP = 3
      frmAddEditPromoteCommiss.COM_KEY = 3   ' เพื่อฟอร์มย่อยเล็กสุด
      Load frmAddEditPromoteCommiss
     frmAddEditPromoteCommiss.Show 1
   End If
      
      OKClick = frmAddEditPromoteCommiss.OKClick
   
      Unload frmAddEditPromoteCommiss
      Set frmAddEditPromoteCommiss = Nothing
   
'   If OKClick Then
'      Call RefreshGrid
'   End If

   If OKClick Then
      m_HasModify = True
   End If
   
          m_PromoteYear.QueryFlag = 1
      QueryData (True)
      
End Sub

Private Sub cmdDelete_Click()
Dim ID1 As Long
Dim ID2 As Long
Dim IsOK As Boolean
Dim IsCanLock As Boolean

   If Not cmdDelete.Enabled Then
      Exit Sub
   End If
   
   If Not VerifyGrid(GridEX1.Value(1)) Then
      Exit Sub
   End If
   
   If Not ConfirmDelete(GridEX1.Value(3) & " - " & GridEX1.Value(4)) Then
      Exit Sub
   End If

   ID2 = GridEX1.Value(1)
   
'   ID1 = GridEX1.Value(1)
'   If TabStrip1.SelectedItem.Index = 1 Then
'      If ID1 <= 0 Then
'         m_PromoteYear.DetailsCom1.Remove (ID2)
'      Else
'         m_PromoteYear.DetailsCom1.ITEM(ID2).Flag = "D"
'      End If
'      m_HasModify = True
'' ElseIf TabStrip1.SelectedItem.Index = 2 Then
'  End If
   
If Not glbDaily.DeleteMasterSubCommiss(ID2, IsOK, True, glbErrorLog) Then  ' Ind = 2 ลบ mastersubID ในตารางตัวเอง และ ลูก
      glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
      Call glbDatabaseMngr.UnLockTable(m_TableName, ID2, IsCanLock, glbErrorLog)
      Call EnableForm(Me, True)
      Exit Sub
End If

  
      m_PromoteYear.QueryFlag = 1
      QueryData (True)
'   Call RefreshGrid
End Sub

Private Sub cmdEdit_Click()
Dim IsOK As Boolean
Dim IsCanLock As Boolean
Dim ID As Long
Dim OKClick As Boolean

 If Not cmdEdit.Enabled Then
      Exit Sub
   End If
      
   If Not VerifyGrid(GridEX1.Value(1)) Then
      Exit Sub
   End If

   ID = Val(GridEX1.Value(1))   ' Commiss_PROMOTE_ID
   OKClick = False
   
   If TabStrip1.SelectedItem.index = 1 Then             'ไว้เมื่อแก้ไข
         frmAddEditPromoteCommiss.MASTER_COMMISS_SUB_PROMOTE_ID = ID                         ' ID ของ คอเล็คคชั้น
      ' 'debug.print m_PromoteYear.MASTER_Commiss_ID
       frmAddEditPromoteCommiss.MASTER_Commiss_ID = m_PromoteYear.MASTER_Commiss_ID
         frmAddEditPromoteCommiss.HeaderText = MapText("แก้ไขค่า Commiss ขาย (พิเศษ)")
               frmAddEditPromoteCommiss.Credit_text = "ยอดขาย (%)"
         frmAddEditPromoteCommiss.ShowMode = SHOW_EDIT
         frmAddEditPromoteCommiss.COM_KEY = 1        ' ใช้ฟอร์คอมขายในการเพิ่มข้อมูลเล็กสุด
         frmAddEditPromoteCommiss.DB_TYP = 1
         Load frmAddEditPromoteCommiss
         frmAddEditPromoteCommiss.Show 1

   ElseIf TabStrip1.SelectedItem.index = 2 Then
         frmAddEditPromoteCommiss.MASTER_COMMISS_SUB_PROMOTE_ID = ID                         ' ID ของ คอเล็คคชั้น
       frmAddEditPromoteCommiss.MASTER_Commiss_ID = m_PromoteYear.MASTER_Commiss_ID
         frmAddEditPromoteCommiss.HeaderText = MapText("แก้ไขค่า Commiss เก็บเงิน (พิเศษ)")
               frmAddEditPromoteCommiss.Credit_text = "เครดิตภายใน (วัน)"
         frmAddEditPromoteCommiss.ShowMode = SHOW_EDIT
            frmAddEditPromoteCommiss.COM_KEY = 2     ' ใช้ฟอร์คอมเก็บเงินในการเพิ่มข้อมูลเล็กสุด
         frmAddEditPromoteCommiss.DB_TYP = 2
         Load frmAddEditPromoteCommiss
         frmAddEditPromoteCommiss.Show 1
         
   ElseIf TabStrip1.SelectedItem.index = 3 Then
         frmAddEditPromoteCommiss.MASTER_COMMISS_SUB_PROMOTE_ID = ID                         ' ID ของ คอเล็คคชั้น
       frmAddEditPromoteCommiss.MASTER_Commiss_ID = m_PromoteYear.MASTER_Commiss_ID
         frmAddEditPromoteCommiss.HeaderText = MapText("แก้ไขค่า Commiss เก็บเงิน (พิเศษ) - สินค้า Commodity")
         frmAddEditPromoteCommiss.Credit_text = "เครดิตภายใน (วัน)"
         frmAddEditPromoteCommiss.ShowMode = SHOW_EDIT
         frmAddEditPromoteCommiss.COM_KEY = 3
         frmAddEditPromoteCommiss.DB_TYP = 3
         Load frmAddEditPromoteCommiss
         frmAddEditPromoteCommiss.Show 1
   End If

            OKClick = frmAddEditPromoteCommiss.OKClick
            
            Unload frmAddEditPromoteCommiss
            Set frmAddEditPromoteCommiss = Nothing

         
   If OKClick Then
'      Call RefreshGrid
      m_HasModify = True
   End If
   
          m_PromoteYear.QueryFlag = 1
      QueryData (True)

End Sub
Private Sub cmdOK_Click()
Dim oMenu As cPopupMenu
Dim lMenuChosen  As Long


   Set oMenu = New cPopupMenu
   lMenuChosen = oMenu.Popup("บันทึก", "-", "บันทึกและออก")
   If lMenuChosen = 0 Then
      Exit Sub
   End If

   If lMenuChosen = 1 Then
      If Not SaveData Then
         Exit Sub
      End If

      ShowMode = SHOW_EDIT
      MASTER_Commiss_ID = m_PromoteYear.MASTER_Commiss_ID
      m_PromoteYear.QueryFlag = 1
      QueryData (True)
      m_HasModify = False
       OKClick = True
   ElseIf lMenuChosen = 3 Then
      If Not SaveData Then
         Exit Sub
      End If

      OKClick = True
      Unload Me
   End If
   
End Sub

Private Sub Form_Activate()

         Call LoadCustomerLookup(uctlCustomerLookup.MyCombo, m_Customer)
      Set uctlCustomerLookup.MyCollection = m_Customer
      
            Call LoadSaleLookup(uctlCommissionSale.MyCombo, FtSaleColl) 'FtSaleColl, COMMISSION_TABLE
      Set uctlCommissionSale.MyCollection = FtSaleColl

   If Not m_HasActivate Then
      m_HasActivate = True
      Me.Refresh
      DoEvents
            
      Call EnableForm(Me, False)
      If (ShowMode = SHOW_EDIT) Or (ShowMode = SHOW_VIEW_ONLY) Then
         m_PromoteYear.QueryFlag = 1
         Call QueryData(True)
         Call TabStrip1_Click
      ElseIf ShowMode = SHOW_ADD Then
         m_PromoteYear.QueryFlag = 0
         Call QueryData(False)
      End If
      
      Call EnableForm(Me, True)
      m_HasModify = False
   End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
Static InUsed As Long
   If InUsed = 1 Then
      Exit Sub
   End If
   InUsed = 1
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
'   ElseIf Shift = 0 And KeyCode = 121 Then
'      Call cmdSave_Click
'      KeyCode = 0
   End If
   InUsed = 0
End Sub

Private Sub Form_Unload(Cancel As Integer)
   If m_Rs.State = adStateOpen Then
      m_Rs.Close
   End If
   Set m_Rs = Nothing
   
'   If cm5_Rs.State = adStateOpen Then
'      cm5_Rs.Close
'   End If
'   Set cm5_Rs = Nothing

Set m_Customer = Nothing
Set FtSaleColl = Nothing

   Set m_PromoteYear = Nothing
   'Set masterSub = Nothing
End Sub

Private Sub GridEX1_ColResize(ByVal ColIndex As Integer, ByVal NewColWidth As Long, ByVal Cancel As GridEX20.JSRetBoolean)
   'debug.print ColIndex & " " & NewColWidth
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

   Set Col = GridEX1.Columns.Add '1
   Col.Width = 0
   Col.Caption = "REAL_ID"

   Set Col = GridEX1.Columns.Add '3
   Col.Width = 2100
   Col.Caption = MapText("เครดิต")

   Set Col = GridEX1.Columns.Add '3
   Col.Width = 8100
   Col.Caption = MapText("รายละเอียด")
End Sub

Private Sub InitFormLayout()
   pnlHeader.PictureBackground = LoadPicture(glbParameterObj.NormalForm1)
   SSFrame1.PictureBackground = LoadPicture(glbParameterObj.NormalForm1)
   
   Me.Caption = HeaderText
   pnlHeader.Caption = Me.Caption
'
   Call InitNormalLabel(lblMasCommissNo, MapText("เลขที่"))
      Call InitNormalLabel(lblMasCommissDesc, MapText("รายละเอียด"))
   Call InitNormalLabel(lblFromDate, MapText("เริ่มใช้วันที่"))
      Call InitNormalLabel(lblToDate, MapText("ถึงวันที่"))
         Call InitNormalLabel(lblCustomer, MapText("รหัสลูกค้า"))
            Call InitNormalLabel(lblSaleName, MapText("พนักงานขาย"))
      
      Call txtMasterCommissDesc.SetTextLenType(TEXT_STRING, glbSetting.CODE_TYPE)
  Call txtMasCommissNo.SetTextLenType(TEXT_STRING, glbSetting.CODE_TYPE)
   
   Me.Picture = LoadPicture(glbParameterObj.MainPicture)

   pnlHeader.Font.Name = GLB_FONT
   pnlHeader.Font.Bold = True
   pnlHeader.Font.Size = 19


   cmdOK.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdAdd.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdEdit.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdDelete.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdExit.Picture = LoadPicture(glbParameterObj.NormalButton1)
   
 '  Call InitMainButton(cmdExit, MapText("ยกเลิก (ESC)"))
   Call InitMainButton(cmdOK, MapText("ตกลง (F2)"))
   Call InitMainButton(cmdAdd, MapText("เพิ่ม (F7)"))
   Call InitMainButton(cmdEdit, MapText("แก้ไข (F3)"))
   Call InitMainButton(cmdDelete, MapText("ลบ (F6)"))

   
   Call InitGrid1
   
   TabStrip1.Font.Bold = True
   TabStrip1.Font.Name = GLB_FONT
   TabStrip1.Font.Size = 16
   
   TabStrip1.Tabs.Clear
   TabStrip1.Tabs.Add().Caption = MapText("1.Com ขาย (พิเศษ)")
   TabStrip1.Tabs.Add().Caption = MapText("2.Com เก็บเงิน (พิเศษ)")
   TabStrip1.Tabs.Add().Caption = MapText("3.Comเก็บเงิน (พิเศษ) - สินค้า Commodity")
'      TabStrip1.Tabs.Add().Caption = MapText("4.เก็บเงิน - ตั้งค่าสินค้า Commodity")
'         TabStrip1.Tabs.Add().Caption = MapText("Commiss (พิเศษ)")
End Sub

Private Sub cmdExit_Click()
'   If Not ConfirmExit(m_HasModify) Then
'      Exit Sub
'   End If

  ' OKClick = False
   Unload Me
End Sub

Private Sub Form_Load()
   OKClick = False
    Call InitFormLayout
    
   m_HasActivate = False
   m_HasModify = False
   Set m_Rs = New ADODB.Recordset
   Set m_PromoteYear = New CCommissMasterPromote
 '  Set cm5_Rs = New ADODB.Recordset
   Set masterSub = New CComMasSubPromote

 Set m_Customer = New Collection
 Set FtSaleColl = New Collection
'   m_HasActivate = False
'   Call EnableForm(Me, True)
End Sub

Private Sub GridEX1_DblClick()
   Call cmdEdit_Click
End Sub

Private Sub GridEX1_RowFormat(RowBuffer As GridEX20.JSRowData)
   If TabStrip1.SelectedItem.index = 1 Then
      RowBuffer.RowStyle = RowBuffer.Value(1)
      ElseIf TabStrip1.SelectedItem.index = 2 Then
      RowBuffer.RowStyle = RowBuffer.Value(1)
      ElseIf TabStrip1.SelectedItem.index = 3 Then
      RowBuffer.RowStyle = RowBuffer.Value(1)
   End If
End Sub

Private Sub GridEX1_UnboundReadData(ByVal RowIndex As Long, ByVal Bookmark As Variant, ByVal Values As GridEX20.JSRowData)
On Error GoTo ErrorHandler
Dim RealIndex As Long
Dim CR As CComMasSubPromote

   glbErrorLog.ModuleName = Me.Name
   glbErrorLog.RoutineName = "UnboundReadData"

 If TabStrip1.SelectedItem.index = 1 Then
      If m_PromoteYear.DetailsCom1 Is Nothing Then
         Exit Sub
      End If
   
      If RowIndex <= 0 Then
         Exit Sub
      End If
      
      If m_PromoteYear.DetailsCom1.Count <= 0 Then
         Exit Sub
      End If
      Set CR = GetItem(m_PromoteYear.DetailsCom1, RowIndex, RealIndex)
      If CR Is Nothing Then
         Exit Sub
      End If
     '      'debug.print CR.MASTER_COMMISS_SUB_PROMOTE_ID & "-" & RealIndex
      Values(1) = CR.MASTER_COMMISS_SUB_PROMOTE_ID
      Values(2) = RealIndex
      Values(3) = CR.CREDIT_NAME
      Values(4) = CR.CREDIT_DESC
'
  ElseIf TabStrip1.SelectedItem.index = 2 Then
      If m_PromoteYear.DetailsCom2 Is Nothing Then
         Exit Sub
      End If
   
      If RowIndex <= 0 Then
         Exit Sub
      End If

      If m_PromoteYear.DetailsCom2.Count <= 0 Then
         Exit Sub
      End If
      Set CR = GetItem(m_PromoteYear.DetailsCom2, RowIndex, RealIndex)
      If CR Is Nothing Then
         Exit Sub
      End If
     '      'debug.print CR.MASTER_COMMISS_SUB_PROMOTE_ID & "-" & RealIndex
      Values(1) = CR.MASTER_COMMISS_SUB_PROMOTE_ID
      Values(2) = RealIndex
      Values(3) = CR.CREDIT_NAME
      Values(4) = CR.CREDIT_DESC
ElseIf TabStrip1.SelectedItem.index = 3 Then

      If m_PromoteYear.DetailsCom1 Is Nothing Then
         Exit Sub
      End If
   
      If RowIndex <= 0 Then
         Exit Sub
      End If
   
      If m_PromoteYear.DetailsCom3.Count <= 0 Then
         Exit Sub
      End If
      Set CR = GetItem(m_PromoteYear.DetailsCom3, RowIndex, RealIndex)
      If CR Is Nothing Then
         Exit Sub
      End If
     '      'debug.print CR.MASTER_COMMISS_SUB_PROMOTE_ID & "-" & RealIndex
      Values(1) = CR.MASTER_COMMISS_SUB_PROMOTE_ID
      Values(2) = RealIndex
      Values(3) = CR.CREDIT_NAME
      Values(4) = CR.CREDIT_DESC

   End If
      
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub

Private Sub TabStrip1_Click()
   If TabStrip1.SelectedItem Is Nothing Then
      Exit Sub
   End If
   
   If TabStrip1.SelectedItem.index = 1 Then
      Call InitGrid1
      GridEX1.ItemCount = CountItem(m_PromoteYear.DetailsCom1)
      GridEX1.Rebind
   ElseIf TabStrip1.SelectedItem.index = 2 Then
      Call InitGrid1
      GridEX1.ItemCount = CountItem(m_PromoteYear.DetailsCom2)
      GridEX1.Rebind
   ElseIf TabStrip1.SelectedItem.index = 3 Then
      Call InitGrid1
      GridEX1.ItemCount = CountItem(m_PromoteYear.DetailsCom3)
      GridEX1.Rebind
   End If
End Sub

Private Sub txtMasCommissNo_Change()
   m_HasModify = True
End Sub

Private Sub uctlToDate_HasChange()
   m_HasModify = True
End Sub

Private Sub uctlFromDate_HasChange()
   m_HasModify = True
End Sub

'Private Sub txtDocumentNo_LostFocus()
'   If Not CheckUniqueNs(IMPORT_UNIQUE, txtDocumentNo.Text, ID) Then
'      glbErrorLog.LocalErrorMsg = MapText("มีข้อมูล") & " " & txtDocumentNo.Text & " " & MapText("อยู่ในระบบแล้ว")
'      glbErrorLog.ShowUserError
'
'      'txtDocumentNo.SetFocus
'      txtDocumentNo.Text = ""
'      Exit Sub
'   End If
'End Sub

'Private Sub uctlDocumentDate_LostFocus()
'   If ShowMode = SHOW_ADD And uctlDocumentDate.ShowDate > 0 Then
'      If Not VerifyDateInterval(uctlDocumentDate.ShowDate) Then
'         uctlDocumentDate.SetFocus
'         Exit Sub
'      End If
'   ElseIf Not CheckUniqueNs(IMPORT_UNIQUE, txtDocumentNo.Text, ID) Then
'      txtDocumentNo.SetFocus
'      Exit Sub
'   ElseIf Not (uctlDocumentDate.ShowDate > 0) Then
'      uctlDocumentDate.SetFocus
'      Exit Sub
'   End If
'End Sub
Public Sub RefreshGrid()
If TabStrip1.SelectedItem.index = 1 Then
   GridEX1.ItemCount = CountItem(m_PromoteYear.DetailsCom1)
   GridEX1.Rebind
ElseIf TabStrip1.SelectedItem.index = 2 Then
   GridEX1.ItemCount = CountItem(m_PromoteYear.DetailsCom2)
   GridEX1.Rebind
ElseIf TabStrip1.SelectedItem.index = 3 Then
   GridEX1.ItemCount = CountItem(m_PromoteYear.DetailsCom3)
   GridEX1.Rebind
End If
End Sub

Private Sub Form_Resize()
On Error Resume Next
   SSFrame1.Width = ScaleWidth
   SSFrame1.Height = ScaleHeight
   pnlHeader.Width = ScaleWidth
   GridEX1.Width = ScaleWidth - 2 * GridEX1.Left
   GridEX1.Height = ScaleHeight - GridEX1.Top - 620
   TabStrip1.Width = GridEX1.Width
   cmdAdd.Top = ScaleHeight - 580
   cmdEdit.Top = ScaleHeight - 580
   cmdDelete.Top = ScaleHeight - 580
   cmdOK.Top = ScaleHeight - 580
'   cmdExit.Top = ScaleHeight - 580
'   cmdExit.Left = ScaleWidth - cmdExit.Width - 50
   cmdOK.Left = ScaleWidth - cmdOK.Width - 50        'cmdExit.Left
End Sub

Private Sub GridEX1_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = DUMMY_KEY Then
      Call cmdExit_Click
      KeyCode = 0
   ElseIf Shift = 0 And KeyCode = 113 Then
      Call cmdOK_Click
      KeyCode = 0
   End If
End Sub

Private Sub uctlCustomerLookup_HasChange()
   m_HasModify = True
End Sub

Private Sub uctlCustomerLookup_Change()
   m_HasModify = True
End Sub
Private Sub uctlCommissionSale_HasChange()
   m_HasModify = True
End Sub

Private Sub uctlCommissionSale_Change()
   m_HasModify = True
End Sub


