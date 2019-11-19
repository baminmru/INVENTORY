VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmGetFile 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Загрузка  технических данных"
   ClientHeight    =   1125
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   6030
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1125
   ScaleWidth      =   6030
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdPath 
      Caption         =   "..."
      Height          =   375
      Left            =   3840
      TabIndex        =   4
      Top             =   480
      Width           =   495
   End
   Begin VB.TextBox txtPath 
      Height          =   375
      Left            =   120
      Locked          =   -1  'True
      TabIndex        =   3
      Top             =   480
      Width           =   3615
   End
   Begin MSComDlg.CommonDialog cdlg 
      Left            =   4200
      Top             =   720
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
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
   Begin VB.Label Label1 
      Caption         =   "Файл с данными"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   3975
   End
End
Attribute VB_Name = "frmGetFile"
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

Private Sub cmdPath_Click()
  On Error Resume Next
  
  On Error GoTo bye
  Dim fn As String
  cdlg.CancelError = True
  cdlg.Filter = "Документ |*.*"
  'cdlg.DefaultExt = "XML"
  'cdlg.FileName = App.path & "\" & item.ID & ".xml"
  cdlg.Flags = cdlOFNPathMustExist + cdlOFNHideReadOnly + cdlOFNFileMustExist
  cdlg.ShowOpen
  txtPath = cdlg.FileName
  

bye:
End Sub

Private Sub OKButton_Click()
    If txtPath.Text <> "" Then
        OK = True
        Me.Hide
    End If
End Sub
