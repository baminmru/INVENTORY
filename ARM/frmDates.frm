VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmDates 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Даты инвентаризации"
   ClientHeight    =   1875
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   6030
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1875
   ScaleWidth      =   6030
   ShowInTaskbar   =   0   'False
   Begin MSComCtl2.DTPicker dtpTo 
      Height          =   375
      Left            =   240
      TabIndex        =   5
      Top             =   1200
      Width           =   2775
      _ExtentX        =   4895
      _ExtentY        =   661
      _Version        =   393216
      Format          =   74907649
      CurrentDate     =   40138
   End
   Begin MSComCtl2.DTPicker dtpFrom 
      Height          =   375
      Left            =   240
      TabIndex        =   3
      Top             =   360
      Width           =   2775
      _ExtentX        =   4895
      _ExtentY        =   661
      _Version        =   393216
      Format          =   74907649
      CurrentDate     =   40138
   End
   Begin VB.CommandButton CancelButton 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   4680
      TabIndex        =   1
      Top             =   600
      Width           =   1215
   End
   Begin VB.CommandButton OKButton 
      Caption         =   "OK"
      Height          =   375
      Left            =   4680
      TabIndex        =   0
      Top             =   120
      Width           =   1215
   End
   Begin VB.Label Label2 
      Caption         =   "По:"
      Height          =   375
      Left            =   240
      TabIndex        =   4
      Top             =   840
      Width           =   2055
   End
   Begin VB.Label Label1 
      Caption         =   "С:"
      Height          =   375
      Left            =   240
      TabIndex        =   2
      Top             =   120
      Width           =   1695
   End
End
Attribute VB_Name = "frmDates"
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
