VERSION 5.00
Begin VB.Form frmSplash 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   3975
   ClientLeft      =   255
   ClientTop       =   1410
   ClientWidth     =   7095
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "frmSplash.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3975
   ScaleWidth      =   7095
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      BackColor       =   &H8000000E&
      Height          =   4035
      Left            =   0
      TabIndex        =   0
      Top             =   -90
      Width           =   7080
      Begin VB.Image Image1 
         Height          =   2355
         Left            =   120
         Picture         =   "frmSplash.frx":000C
         Stretch         =   -1  'True
         Top             =   840
         Width           =   3300
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Платформа ""Муромец"""
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   300
         Index           =   1
         Left            =   2145
         TabIndex        =   7
         Top             =   3255
         Width           =   3555
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Платформа ""Муромец"""
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808080&
         Height          =   300
         Index           =   0
         Left            =   2160
         TabIndex        =   6
         Top             =   3270
         Width           =   3555
      End
      Begin VB.Label lblCompanyProduct 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "SGS представляет:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   18
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   435
         Index           =   1
         Left            =   2250
         TabIndex        =   5
         Top             =   225
         Width           =   4545
         WordWrap        =   -1  'True
      End
      Begin VB.Label lblWarning 
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   150
         TabIndex        =   1
         Top             =   3660
         Width           =   6855
      End
      Begin VB.Label lblVersion 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H8000000E&
         Caption         =   "Version"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   6000
         TabIndex        =   2
         Top             =   3240
         Width           =   885
      End
      Begin VB.Label lblProductName 
         Alignment       =   2  'Center
         BackColor       =   &H8000000E&
         Caption         =   "Product"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   20.25
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   2055
         Left            =   3480
         TabIndex        =   4
         Top             =   1035
         Width           =   3360
         WordWrap        =   -1  'True
      End
      Begin VB.Label lblCompanyProduct 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "SGS представляет:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   18
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808080&
         Height          =   435
         Index           =   0
         Left            =   2280
         TabIndex        =   3
         Top             =   240
         Width           =   4545
         WordWrap        =   -1  'True
      End
   End
End
Attribute VB_Name = "frmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Sub Form_KeyPress(KeyAscii As Integer)
    Unload Me
End Sub

Private Sub Form_Load()
    lblVersion.Caption = "Version " & App.Major & "." & App.Minor & "." & App.Revision
    lblProductName.Caption = App.Title
End Sub

Private Sub Frame1_Click()
    Unload Me
End Sub

Private Sub lblCopyright_Click()

End Sub

Public Sub DBMainteice()
  Dim cap As String

  Dim d As Date
  If GetSetting("ABOL", "INVENTORY", "DODBMAINTAIN", "FALSE") = "TRUE" Then
    cap = Me.Caption
    d = CDate(GetSetting("ABOL", "INVENTORY", "DBMAINTAIN", Date - 1))
    If d < Date Then
    
      On Error Resume Next
      If Session.IsPOSTGRESQL Then
        Me.Caption = "Сборка мусора"
        DoEvents
        Session.GetData ("vacuum full")
        Me.Caption = "Переиндексация"
        DoEvents
        
        Session.GetData ("reindex database test")
        Me.Caption = "Сбор статистики"
        DoEvents
        Session.GetData ("analize")
      End If
      
      Me.Caption = "Очистка лога"
        DoEvents
      Session.GetData ("truncate table syslog")
       
      On Error GoTo bye
      Me.Caption = "Разблокировка"
      DoEvents
      Dim v As NamedValues
      Set v = New NamedValues
      Call Session.Exec("AdminUnlockAll", v)
      Call SaveSetting("ABOL", "INVENTORY", "DBMAINTAIN", Date + 7)
    End If
    Me.Caption = cap
  End If
  
  Exit Sub
bye:
  MsgBox Err.Description
  Me.Caption = cap
End Sub

