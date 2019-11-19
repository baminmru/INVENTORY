VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmNoInvFilter 
   Caption         =   "Параметры выборки"
   ClientHeight    =   2505
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   4800
   LinkTopic       =   "Form1"
   ScaleHeight     =   2505
   ScaleWidth      =   4800
   StartUpPosition =   3  'Windows Default
   Begin VB.CheckBox chkShowClosed 
      Caption         =   "Показывать списанные объекты"
      Height          =   375
      Left            =   120
      TabIndex        =   6
      Top             =   1920
      Width           =   3495
   End
   Begin VB.CommandButton OKButton 
      Caption         =   "OK"
      Height          =   375
      Left            =   3120
      TabIndex        =   3
      Top             =   360
      Width           =   1215
   End
   Begin VB.CommandButton CancelButton 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   3120
      TabIndex        =   2
      Top             =   840
      Width           =   1215
   End
   Begin MSComCtl2.DTPicker dtpTo 
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   1200
      Width           =   2775
      _ExtentX        =   4895
      _ExtentY        =   661
      _Version        =   393216
      Format          =   126353409
      CurrentDate     =   40138
   End
   Begin MSComCtl2.DTPicker dtpFrom 
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   360
      Width           =   2775
      _ExtentX        =   4895
      _ExtentY        =   661
      _Version        =   393216
      Format          =   126353409
      CurrentDate     =   40138
   End
   Begin VB.Label Label1 
      Caption         =   "С:"
      Height          =   375
      Left            =   120
      TabIndex        =   5
      Top             =   120
      Width           =   1695
   End
   Begin VB.Label Label2 
      Caption         =   "По:"
      Height          =   375
      Left            =   120
      TabIndex        =   4
      Top             =   840
      Width           =   2055
   End
End
Attribute VB_Name = "frmNoInvFilter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public OK As Boolean

Private Sub CancelButton_Click()
OK = False
Me.Hide
End Sub

Private Sub Form_Load()
OK = False
dtpFrom.Value = Date - 30
dtpTo.Value = Date
End Sub

Private Sub OKButton_Click()
OK = True
Me.Hide
End Sub

