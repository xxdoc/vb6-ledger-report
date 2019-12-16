VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsParameter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Const WIN_MIS_REVISION = "3"
Private Const APP_REVISION = "12"

Private Const REGKEY = "Software\WIN\MittraphapReport"
Private Reg As clsRegistry

Public DBFile As String
Public UserName As String
Public Password As String

Public LicenseIP As String
Public LicensePort As Long
Public Version As String
Public Language As Long

Public AppName As String
Public AppVersion As String

Public Drive As String
Public Directory As String
Public ImportYear As Long
Public ImportSemester As Long
Public FilterExt As String
Public NewExt As String
Public UsedCount As Long
Public Temp As Long
Public Area As Long

Public Drive2 As String
Public Directory2 As String
Public ImportYear2 As Long
Public ImportSemester2 As Long
Public FilterExt2 As String
Public NewExt2 As String
Public UsedCount2 As Long
Public Temp2 As Long
Public Area2 As Long

Public MainPicture As String
Public MainButton As String
Public NormalButton1 As String
Public NormalForm1 As String
Public ButtonCursor As String
Public ApplicationPath As String
Public ReportFile As String
Public PaidVocherPic As String
Public JVVocherPic As String
Public ReceiptVocherPic As String
Public DBConfigFile As String
Public ComputerName As String
Public MgpCustomerProfilePic As String
Private Sub Class_Initialize()
Dim ErrCode As Integer

   Set Reg = New clsRegistry
   Reg.RgtRootKey = &H80000002 'HKEY_LOCAL_MACHINE
   Reg.RgtSubKey = REGKEY

   ApplicationPath = App.PATH
   ReportFile = ApplicationPath & "\" & "temp_report.html"
   PaidVocherPic = ApplicationPath & "\Pictures\" & "paid_vocher.jpg"
   JVVocherPic = ApplicationPath & "\Pictures\" & "transfer_vocher.jpg"
   ReceiptVocherPic = ApplicationPath & "\Pictures\" & "receipt_vocher.jpg"
   MgpCustomerProfilePic = ApplicationPath & "\Pictures\" & "MgpCustomerProfilePic.jpg"
   DBConfigFile = ApplicationPath & "\" & "config.mdb"
   
   AppName = "GeneticSoft WIN Pricing"
   AppVersion = App.Major & "." & App.Minor & "." & WIN_MIS_REVISION & "." & App.Revision & "   (Foxpro 7.0)"
   DBFile = "D:\Express"
   
   UserName = "SYSDBA"
   Password = "masterkey"
   
   LicenseIP = "127.0.0.1"
   LicensePort = 8888
   
   Call Reg.RgtGetStringValue("DBFile", DBFile, ErrCode)
   Call Reg.RgtGetStringValue("UserName", UserName, ErrCode)
   Call Reg.RgtGetStringValue("Password", Password, ErrCode)
   Password = DecryptText(Password)
   Call Reg.RgtGetStringValue("LicenseIP", LicenseIP, ErrCode)
   Call Reg.RgtGetLongValue("LicensePort", LicensePort, ErrCode)
   Call Reg.RgtGetLongValue("Language", Language, ErrCode)
   Language = 1
   
   Call Reg.RgtGetStringValue("Drive", Drive, ErrCode)
   Call Reg.RgtGetStringValue("Directory", Directory, ErrCode)
   Call Reg.RgtGetLongValue("ImportYear", ImportYear, ErrCode)
   Call Reg.RgtGetLongValue("ImportSemester", ImportSemester, ErrCode)
   Call Reg.RgtGetStringValue("FilterExt", FilterExt, ErrCode)
   Call Reg.RgtGetStringValue("NewExt", NewExt, ErrCode)
   Call Reg.RgtGetStringValue("DBConfigFile", DBConfigFile, ErrCode)
   Call Reg.RgtGetLongValue("Day", UsedCount, ErrCode)
   Call Reg.RgtGetLongValue("Area", Area, ErrCode)
   
   Call Reg.RgtGetStringValue("Drive2", Drive2, ErrCode)
   Call Reg.RgtGetStringValue("Directory2", Directory2, ErrCode)
   Call Reg.RgtGetLongValue("ImportYear2", ImportYear2, ErrCode)
   Call Reg.RgtGetLongValue("ImportSemester2", ImportSemester2, ErrCode)
   Call Reg.RgtGetStringValue("FilterExt2", FilterExt2, ErrCode)
   Call Reg.RgtGetStringValue("NewExt2", NewExt2, ErrCode)
   Call Reg.RgtGetLongValue("Day2", UsedCount2, ErrCode)
   Call Reg.RgtGetLongValue("Area2", Area2, ErrCode)
   
   Version = App.Major & "." & App.Minor & "." & APP_REVISION & "." & App.Revision

   Temp = 0
   
   MainPicture = App.PATH & "\Pictures\main1.jpg"
   MainButton = App.PATH & "\Pictures\retangle_button1.jpg"
   NormalButton1 = App.PATH & "\Pictures\normal_button2.jpg"
   NormalForm1 = App.PATH & "\Pictures\dialog_form1.jpg"
   ButtonCursor = App.PATH & "\Pictures\H_POINT.CUR"
   
   ComputerName = ""
End Sub

Private Sub Class_Terminate()
Dim ErrCode As Integer
   
   Call Reg.RgtSetStringValue("DBFile", DBFile, ErrCode)
   Call Reg.RgtSetStringValue("UserName", UserName, ErrCode)
   Password = EncryptText(Password)
   Call Reg.RgtSetStringValue("Password", Password, ErrCode)
   Call Reg.RgtSetStringValue("LicenseIP", LicenseIP, ErrCode)
   Call Reg.RgtSetLongValue("LicensePort", LicensePort, ErrCode)
   
   Call Reg.RgtSetStringValue("Drive", Drive, ErrCode)
   Call Reg.RgtSetStringValue("Directory", Directory, ErrCode)
   Call Reg.RgtSetLongValue("ImportYear", ImportYear, ErrCode)
   Call Reg.RgtSetLongValue("ImportSemester", ImportSemester, ErrCode)
   Call Reg.RgtSetStringValue("FilterExt", FilterExt, ErrCode)
   Call Reg.RgtSetStringValue("NewExt", NewExt, ErrCode)
   Call Reg.RgtSetLongValue("Day", UsedCount, ErrCode)
   Call Reg.RgtSetLongValue("Area", Area, ErrCode)
   Call Reg.RgtSetLongValue("Language", Language, ErrCode)
   
   Call Reg.RgtSetStringValue("Drive2", Drive2, ErrCode)
   Call Reg.RgtSetStringValue("Directory2", Directory2, ErrCode)
   Call Reg.RgtSetLongValue("ImportYear2", ImportYear2, ErrCode)
   Call Reg.RgtSetLongValue("ImportSemester2", ImportSemester2, ErrCode)
   Call Reg.RgtSetStringValue("FilterExt2", FilterExt2, ErrCode)
   Call Reg.RgtSetStringValue("NewExt2", NewExt2, ErrCode)
   Call Reg.RgtSetLongValue("Day2", UsedCount2, ErrCode)
   Call Reg.RgtSetLongValue("Area2", Area2, ErrCode)
   Call Reg.RgtSetStringValue("DBConfigFile", DBConfigFile, ErrCode)
   
   Set Reg = Nothing
End Sub
