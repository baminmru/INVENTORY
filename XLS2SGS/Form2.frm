VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Begin VB.Form frm 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Параметры запуска"
   ClientHeight    =   4080
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   3885
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4080
   ScaleWidth      =   3885
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdLog 
      Caption         =   "..."
      Height          =   255
      Left            =   2880
      TabIndex        =   13
      Top             =   3120
      Width           =   495
   End
   Begin VB.TextBox txtLog 
      Height          =   375
      Left            =   240
      TabIndex        =   12
      Text            =   "c:\xls2ssg.log"
      Top             =   3000
      Width           =   2535
   End
   Begin MSComDlg.CommonDialog cDlg 
      Left            =   3000
      Top             =   1440
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Отмена"
      Height          =   495
      Left            =   2040
      TabIndex        =   10
      Top             =   3480
      Width           =   1575
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Height          =   495
      Left            =   240
      TabIndex        =   9
      Top             =   3480
      Width           =   1695
   End
   Begin VB.CommandButton cmdFile 
      Caption         =   "..."
      Height          =   255
      Left            =   2880
      TabIndex        =   8
      Top             =   2280
      Width           =   495
   End
   Begin VB.TextBox txtFile 
      Height          =   285
      Left            =   240
      Locked          =   -1  'True
      TabIndex        =   7
      Top             =   2280
      Width           =   2535
   End
   Begin VB.TextBox txtSite 
      Height          =   285
      Left            =   240
      TabIndex        =   5
      Text            =   "SGS_INVENTORY"
      Top             =   1560
      Width           =   2535
   End
   Begin VB.TextBox txtPWD 
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   240
      PasswordChar    =   "*"
      TabIndex        =   3
      Top             =   960
      Width           =   2535
   End
   Begin VB.TextBox txtUser 
      Height          =   285
      Left            =   240
      TabIndex        =   1
      Top             =   360
      Width           =   2535
   End
   Begin VB.Label Label5 
      Caption         =   "Файл для логирования"
      Height          =   255
      Left            =   240
      TabIndex        =   11
      Top             =   2640
      Width           =   2415
   End
   Begin VB.Label Label4 
      Caption         =   "Настроечный файл"
      Height          =   255
      Left            =   240
      TabIndex        =   6
      Top             =   2040
      Width           =   2415
   End
   Begin VB.Label Label3 
      Caption         =   "Сайт"
      Height          =   255
      Left            =   240
      TabIndex        =   4
      Top             =   1320
      Width           =   1935
   End
   Begin VB.Label Label2 
      Caption         =   "Пароль"
      Height          =   255
      Left            =   240
      TabIndex        =   2
      Top             =   720
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Пользователь"
      Height          =   255
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   1815
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
  If txtFile <> "" And txtSite <> "" And txtUser <> "" Then
     Me.Hide
     OK = True
     Call SaveSetting("MTZ", "XLS2SGS", "Site", txtSite.Text)
     Call SaveSetting("MTZ", "XLS2SGS", "File", txtFile.Text)
     Call SaveSetting("MTZ", "XLS2SGS", "User", txtUser.Text)
     Call SaveSetting("MTZ", "XLS2SGS", "Log", txtLog.Text)
     Exit Sub
  End If
  MsgBox "Задайте параметр(ы)"
End Sub

Private Sub Form_Load()
txtSite.Text = GetSetting("MTZ", "XLS2SGS", "Site", "LOADCFG")
txtFile.Text = GetSetting("MTZ", "XLS2SGS", "File", "")
txtUser.Text = GetSetting("MTZ", "XLS2SGS", "User", "supervisor")
txtLog.Text = GetSetting("MTZ", "XLS2SGS", "Log", "c:\xls2sgs.log")
End Sub
