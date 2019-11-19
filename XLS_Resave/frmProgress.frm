VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmProgress 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Загрузка технической информации"
   ClientHeight    =   1260
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7470
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1260
   ScaleWidth      =   7470
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   735
      Left            =   120
      TabIndex        =   0
      Top             =   480
      Width           =   7215
      Begin MSComctlLib.ProgressBar pb 
         Height          =   375
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   6975
         _ExtentX        =   12303
         _ExtentY        =   661
         _Version        =   393216
         Appearance      =   1
      End
   End
   Begin VB.Label Label1 
      Height          =   495
      Left            =   240
      TabIndex        =   2
      Top             =   0
      Width           =   6735
   End
End
Attribute VB_Name = "frmProgress"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_HelpID = 180
Option Explicit


Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  If UnloadMode <> vbFormCode Then
    Cancel = True
  End If
End Sub

Public Sub NextVal(ByVal s As String)

  pb.Value = (pb.Value + 1) Mod 100
  Label1.Caption = s
  DoEvents
End Sub

