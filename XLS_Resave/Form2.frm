VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frm 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Параметры запуска"
   ClientHeight    =   2775
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   3885
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2775
   ScaleWidth      =   3885
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdLog 
      Caption         =   "..."
      Height          =   255
      Left            =   3000
      TabIndex        =   7
      Top             =   1200
      Width           =   495
   End
   Begin VB.TextBox txtLog 
      Height          =   375
      Left            =   360
      TabIndex        =   6
      Text            =   "c:\xls2ssg.log"
      Top             =   1080
      Width           =   2535
   End
   Begin MSComDlg.CommonDialog cDlg 
      Left            =   240
      Top             =   1560
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Отмена"
      Height          =   495
      Left            =   2040
      TabIndex        =   4
      Top             =   2160
      Width           =   1575
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Height          =   495
      Left            =   240
      TabIndex        =   3
      Top             =   2160
      Width           =   1695
   End
   Begin VB.CommandButton cmdFile 
      Caption         =   "..."
      Height          =   255
      Left            =   3000
      TabIndex        =   2
      Top             =   360
      Width           =   495
   End
   Begin VB.TextBox txtFile 
      Height          =   285
      Left            =   360
      Locked          =   -1  'True
      TabIndex        =   1
      Top             =   360
      Width           =   2535
   End
   Begin VB.Label Label5 
      Caption         =   "Файл для логирования"
      Height          =   255
      Left            =   360
      TabIndex        =   5
      Top             =   720
      Width           =   2415
   End
   Begin VB.Label Label4 
      Caption         =   "Настроечный файл"
      Height          =   255
      Left            =   360
      TabIndex        =   0
      Top             =   120
      Width           =   2415
   End
End
Attribute VB_Name = "frm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public OK As Boolean

Private Sub cmdCancel_Click()
  OK = False
  Me.Hide
End Sub

Private Sub cmdFile_Click()
On Error GoTo bye
  cDlg.CancelError = True
  cDlg.ShowOpen
  txtFile = cDlg.Filename
bye:

End Sub

Private Sub cmdLog_Click()
On Error GoTo bye
  cDlg.CancelError = True
  cDlg.ShowOpen
  txtLog = cDlg.Filename
bye:
End Sub

Private Sub cmdOK_Click()
  If txtFile <> "" Then
     Me.Hide
     OK = True
     
     Call SaveSetting("MTZ", "XLSRESAVE", "File", txtFile.Text)
    Call SaveSetting("MTZ", "XLSRESAVE", "Log", txtLog.Text)
     Exit Sub
  End If
  MsgBox "Задайте параметр(ы)"
End Sub

Private Sub Form_Load()
txtFile.Text = GetSetting("MTZ", "XLSRESAVE", "File", "")
txtLog.Text = GetSetting("MTZ", "XLSRESAVE", "Log", "c:\XLSresaveer.log")
End Sub
