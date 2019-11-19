VERSION 5.00
Begin VB.Form frmPrnSetup 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Настройка принтера"
   ClientHeight    =   2025
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4005
   Icon            =   "frmPrnSetup.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2025
   ScaleWidth      =   4005
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox txtDocPrn 
      Height          =   315
      Left            =   120
      Style           =   2  'Dropdown List
      TabIndex        =   4
      Top             =   1065
      Width           =   3780
   End
   Begin VB.ComboBox txtZPrn 
      Height          =   315
      Left            =   120
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   360
      Width           =   3780
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Отменить"
      Height          =   390
      Left            =   2400
      TabIndex        =   2
      Top             =   1440
      Width           =   1455
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   390
      Left            =   840
      TabIndex        =   1
      Top             =   1440
      Width           =   1455
   End
   Begin VB.Label Label2 
      Caption         =   "Драйвер принтера документов:"
      Height          =   315
      Left            =   135
      TabIndex        =   5
      Top             =   840
      Width           =   3705
   End
   Begin VB.Label Label1 
      Caption         =   "Драйвер принтера штрихкодов:"
      Height          =   315
      Left            =   135
      TabIndex        =   0
      Top             =   135
      Width           =   3705
   End
End
Attribute VB_Name = "frmPrnSetup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_HelpID = 680
Option Explicit
' окно настройки параметров принтеров

Private Sub cmdCancel_Click()
Me.Hide
End Sub

' сохранение настроек
Private Sub cmdOK_Click()
Me.Hide
SaveSetting "SGS", "ITTSETTINGS", "ZPRN", txtZPrn.Text
SaveSetting "SGS", "ITTSETTINGS", "DOCPRN", txtDocPrn.Text
End Sub

' загрузка
Private Sub Form_Load()


Dim P As Printer
  txtZPrn.Clear
  txtDocPrn.Clear
  For Each P In Printers
    txtZPrn.AddItem P.DeviceName
    txtDocPrn.AddItem P.DeviceName
  Next
  On Error Resume Next
  txtZPrn.Text = GetSetting("SGS", "ITTSETTINGS", "ZPRN", "")
  txtDocPrn.Text = GetSetting("SGS", "ITTSETTINGS", "DOCPRN", "")
End Sub

