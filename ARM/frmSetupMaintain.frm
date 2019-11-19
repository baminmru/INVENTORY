VERSION 5.00
Begin VB.Form frmSetupMaintain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Настройки обслуживания БД"
   ClientHeight    =   1545
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   4560
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1545
   ScaleWidth      =   4560
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox chkDo 
      Caption         =   "Проводить периодическое обслуживание базы данных при запуске программы"
      Height          =   615
      Left            =   120
      TabIndex        =   2
      Top             =   240
      Width           =   4455
   End
   Begin VB.CommandButton OKButton 
      Caption         =   "OK"
      Height          =   375
      Left            =   1920
      TabIndex        =   1
      Top             =   960
      Width           =   1215
   End
   Begin VB.CommandButton CancelButton 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   3240
      TabIndex        =   0
      Top             =   960
      Width           =   1215
   End
End
Attribute VB_Name = "frmSetupMaintain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public OK As Boolean

Private Sub CancelButton_Click()
OK = False
Unload Me
End Sub

Private Sub Form_Load()
  If GetSetting("ABOL", "INVENTORY", "DODBMAINTAIN", "FALSE") = "TRUE" Then
    chkDo.Value = vbChecked
  End If
End Sub

Private Sub OKButton_Click()
If chkDo.Value = vbChecked Then
   Call SaveSetting("ABOL", "INVENTORY", "DODBMAINTAIN", "TRUE")
Else
  Call SaveSetting("ABOL", "INVENTORY", "DODBMAINTAIN", "FALSE")
End If
Unload Me
End Sub
