VERSION 5.00
Object = "{C8AE1F3B-9D93-4357-8B5E-74AB45B9B42F}#1.0#0"; "LoglExtender.ocx"
Begin VB.Form frmProtocol 
   Caption         =   "Протокол работы пользователя"
   ClientHeight    =   3030
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   4560
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   3030
   ScaleWidth      =   4560
   Begin LogExtender.JournalViewEx objLog 
      Height          =   2655
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4335
      _ExtentX        =   7646
      _ExtentY        =   4683
   End
End
Attribute VB_Name = "frmProtocol"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Item As Object

Private Sub Form_Load()
objLog.Init
objLog.OnInit Item, "", Me
Me.Caption = "Протокол: " & Item.Name
objLog.OnClick Item, Me
End Sub

Private Sub Form_Resize()
  On Error Resume Next
  objLog.Move 0, 0, Me.ScaleWidth, Me.ScaleHeight
End Sub

