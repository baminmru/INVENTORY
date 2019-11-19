VERSION 5.00
Begin VB.Form frmGetInv 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Параметр отчета"
   ClientHeight    =   1425
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   4560
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1425
   ScaleWidth      =   4560
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtINV 
      Height          =   405
      Left            =   120
      Locked          =   -1  'True
      TabIndex        =   4
      Top             =   360
      Width           =   3735
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Отмена"
      Height          =   375
      Left            =   3240
      TabIndex        =   2
      Top             =   960
      Width           =   1095
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Height          =   375
      Left            =   2280
      TabIndex        =   1
      Top             =   960
      Width           =   855
   End
   Begin VB.CommandButton cmdInvI 
      Caption         =   "..."
      Height          =   375
      Left            =   3840
      TabIndex        =   0
      Top             =   360
      Width           =   495
   End
   Begin VB.Label Label1 
      Caption         =   "Инвентаризация"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   1695
   End
End
Attribute VB_Name = "frmGetInv"
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
If Manager.GetObjectListDialogEx3(id, brief, "", "INV_INV") Then
    txtINV.Text = brief
    txtINV.Tag = id
End If

End Sub

Private Sub cmdOK_Click()
If txtINV.Tag <> "" Then
    OK = True
    Me.Hide
Else
    MsgBox "Необходимо выбрать инвентаризацию для создания отчета", "Внимание"
End If
End Sub

