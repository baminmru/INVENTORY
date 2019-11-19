VERSION 5.00
Begin VB.Form frmGetInv2 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Параметр отчета"
   ClientHeight    =   7095
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   4560
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7095
   ScaleWidth      =   4560
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtCodeMask 
      Height          =   375
      Left            =   120
      TabIndex        =   16
      Top             =   1800
      Width           =   4215
   End
   Begin VB.TextBox txtMaskE 
      Height          =   375
      Index           =   5
      Left            =   120
      TabIndex        =   14
      Top             =   6120
      Width           =   4335
   End
   Begin VB.TextBox txtMaskE 
      Height          =   375
      Index           =   4
      Left            =   120
      TabIndex        =   13
      Top             =   5640
      Width           =   4335
   End
   Begin VB.TextBox txtMaskE 
      Height          =   375
      Index           =   3
      Left            =   120
      TabIndex        =   12
      Top             =   5160
      Width           =   4335
   End
   Begin VB.TextBox txtMaskE 
      Height          =   375
      Index           =   2
      Left            =   120
      TabIndex        =   11
      Top             =   4680
      Width           =   4335
   End
   Begin VB.TextBox txtMaskE 
      Height          =   375
      Index           =   1
      Left            =   120
      TabIndex        =   10
      Top             =   4200
      Width           =   4335
   End
   Begin VB.TextBox txtMaskE 
      Height          =   375
      Index           =   0
      Left            =   120
      TabIndex        =   8
      Top             =   3720
      Width           =   4335
   End
   Begin VB.TextBox txtMask 
      Height          =   375
      Left            =   120
      TabIndex        =   7
      Top             =   2880
      Width           =   4335
   End
   Begin VB.CheckBox chkBad 
      Caption         =   "Печатать только помеченные, как плохие"
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   960
      Width           =   4095
   End
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
      Top             =   6600
      Width           =   1095
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Height          =   375
      Left            =   2280
      TabIndex        =   1
      Top             =   6600
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
   Begin VB.Label Label4 
      Caption         =   "Печатать штрихкод  по маске:"
      Height          =   375
      Left            =   120
      TabIndex        =   15
      Top             =   1440
      Width           =   4215
   End
   Begin VB.Label Label3 
      Caption         =   "Исключить из отчета  ОС по маске:"
      Height          =   255
      Left            =   120
      TabIndex        =   9
      Top             =   3360
      Width           =   4215
   End
   Begin VB.Label Label2 
      Caption         =   "Включить в отчет ОС по маске:"
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   2520
      Width           =   4215
   End
   Begin VB.Label Label1 
      Caption         =   "Инвентаризация:"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   1695
   End
End
Attribute VB_Name = "frmGetInv2"
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

