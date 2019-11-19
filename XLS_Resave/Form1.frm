VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Загрузка данных по ОС"
   ClientHeight    =   5580
   ClientLeft      =   2745
   ClientTop       =   2070
   ClientWidth     =   7695
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5580
   ScaleWidth      =   7695
   Begin MSComDlg.CommonDialog Cdlg 
      Left            =   7560
      Top             =   3600
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.TextBox Status 
      Height          =   4065
      Left            =   135
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   4
      Top             =   1350
      Width           =   7425
   End
   Begin VB.CommandButton cmdProcess 
      Caption         =   "Загрузка данных"
      Height          =   825
      Left            =   120
      Picture         =   "Form1.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "Запустить процесс размещения счетов по директориям"
      Top             =   450
      UseMaskColor    =   -1  'True
      Width           =   7440
   End
   Begin VB.CommandButton cmdSetPath 
      Caption         =   "..."
      Height          =   285
      Left            =   6930
      TabIndex        =   1
      ToolTipText     =   "Найти настроечный файл"
      Top             =   135
      Width           =   600
   End
   Begin VB.TextBox txtFile 
      Height          =   330
      Left            =   1920
      TabIndex        =   0
      ToolTipText     =   "Путь к настроечному файлу"
      Top             =   90
      Width           =   4965
   End
   Begin VB.Label Label1 
      Caption         =   "Настроечный файл"
      Height          =   285
      Left            =   135
      TabIndex        =   3
      Top             =   90
      Width           =   1680
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit



Public Sub cmdProcess_Click()
  ProcessAll txtFile
End Sub

Private Sub cmdSetPath_Click()

  Cdlg.CancelError = True
On Error GoTo bye
  Cdlg.FileName = txtFile
  Cdlg.Filter = "Все файлы|*.*"
  Cdlg.Flags = cdlOFNFileMustExist + cdlOFNHideReadOnly + cdlOFNLongNames
  Cdlg.ShowOpen
  txtFile = Cdlg.FileName
bye:
End Sub



Public Function EraseDat(s2 As String) As String
    Dim I, s1 As String, s As String
    s = Trim(s2)
    For I = 1 To Len(s)
    If Mid(s, I, 1) = "/" Or Mid(s, I, 1) = "\" Or Mid(s, I, 1) = "." Or Mid(s, I, 1) = ":" Or Mid(s, I, 1) = " " Then
        s1 = s1 + "_"
    Else
        s1 = s1 + Mid(s, I, 1)
    End If
    Next
    EraseDat = s1
    
End Function



Private Sub Form_Unload(Cancel As Integer)
    SaveSetting App.Title, "General", "filename", txtFile.Text
End Sub
