VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmGetOS 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "ОС \ Материал"
   ClientHeight    =   1980
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   4785
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1980
   ScaleWidth      =   4785
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox txtPath 
      Height          =   375
      Left            =   120
      Locked          =   -1  'True
      TabIndex        =   6
      Top             =   1080
      Width           =   3735
   End
   Begin VB.CommandButton cmdPath 
      Caption         =   "..."
      Height          =   375
      Left            =   3840
      TabIndex        =   5
      Top             =   1080
      Width           =   495
   End
   Begin VB.CommandButton cmdInvI 
      Caption         =   "..."
      Height          =   375
      Left            =   3840
      TabIndex        =   3
      Top             =   240
      Width           =   495
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Height          =   375
      Left            =   2280
      TabIndex        =   2
      Top             =   1560
      Width           =   855
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Отмена"
      Height          =   375
      Left            =   3240
      TabIndex        =   1
      Top             =   1560
      Width           =   1095
   End
   Begin VB.TextBox txtINV 
      Height          =   405
      Left            =   120
      Locked          =   -1  'True
      TabIndex        =   0
      Top             =   240
      Width           =   3735
   End
   Begin MSComDlg.CommonDialog cdlg 
      Left            =   4320
      Top             =   1320
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Label Label2 
      Caption         =   "Файл отчета"
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   840
      Width           =   3975
   End
   Begin VB.Label Label1 
      Caption         =   "ОС \ Материал:"
      Height          =   255
      Left            =   0
      TabIndex        =   4
      Top             =   0
      Width           =   1695
   End
End
Attribute VB_Name = "frmGetOS"
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

Private Sub cmdInvI_Click()
Dim id As String
Dim brief As String
If Manager.GetObjectListDialog2(id, brief, "", "INV_OS") Then
    txtINV.Text = brief
    txtINV.Tag = id
End If

End Sub

Private Sub cmdOK_Click()
If txtINV.Tag <> "" And txtPath.Text <> "" Then
    OK = True
    Me.Hide
Else
    MsgBox "Необходимо выбрать объект для создания отчета и имя файла", "Внимание"
End If
End Sub

Private Sub cmdPath_Click()
  cdlg.CancelError = True
  cdlg.Filter = "Документ|*.doc"
  cdlg.DefaultExt = "doc"
  cdlg.FileName = GetSetting(App.Title, "Recent", "LastWord", "")
  On Error GoTo bye
  cdlg.ShowOpen
  txtPath.Text = cdlg.FileName
  Call SaveSetting(App.Title, "Recent", "LastWord", txtPath.Text)
bye:
End Sub
