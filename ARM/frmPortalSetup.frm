VERSION 5.00
Begin VB.Form frmPortalSetup 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "��������� ����  ������ �������"
   ClientHeight    =   3435
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   4455
   Icon            =   "frmPortalSetup.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3435
   ScaleWidth      =   4455
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtTable 
      Height          =   285
      Left            =   120
      TabIndex        =   12
      Text            =   "S25"
      Top             =   2880
      Width           =   2895
   End
   Begin VB.CommandButton cmdTest 
      Caption         =   "����"
      Height          =   375
      Left            =   3120
      TabIndex        =   10
      Top             =   1560
      Width           =   1215
   End
   Begin VB.TextBox txtPass 
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   120
      PasswordChar    =   "*"
      TabIndex        =   9
      Top             =   2160
      Width           =   2895
   End
   Begin VB.TextBox txtUsr 
      Height          =   285
      Left            =   120
      TabIndex        =   8
      Top             =   1560
      Width           =   2895
   End
   Begin VB.TextBox txtDB 
      Height          =   285
      Left            =   120
      TabIndex        =   7
      Top             =   960
      Width           =   2895
   End
   Begin VB.TextBox txtSrv 
      Height          =   285
      Left            =   120
      TabIndex        =   6
      Top             =   360
      Width           =   2895
   End
   Begin VB.CommandButton CancelButton 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   3120
      TabIndex        =   1
      Top             =   840
      Width           =   1215
   End
   Begin VB.CommandButton OKButton 
      Caption         =   "OK"
      Height          =   375
      Left            =   3120
      TabIndex        =   0
      Top             =   360
      Width           =   1215
   End
   Begin VB.Label Label5 
      Caption         =   "�������� �������"
      Height          =   255
      Left            =   120
      TabIndex        =   11
      Top             =   2520
      Width           =   2175
   End
   Begin VB.Label Label4 
      Caption         =   "������ ������������"
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   1920
      Width           =   2775
   End
   Begin VB.Label Label3 
      Caption         =   "����������� SQL"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   1320
      Width           =   2775
   End
   Begin VB.Label Label2 
      Caption         =   "���� ������"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   720
      Width           =   2775
   End
   Begin VB.Label Label1 
      Caption         =   "������ ��"
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   0
      Width           =   2655
   End
End
Attribute VB_Name = "frmPortalSetup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_HelpID = 140

Option Explicit
'���� ��������� ���������� � ����� PORTAL IMS

Private Sub CancelButton_Click()
  Me.Hide
End Sub

'�������� ����������
Private Sub cmdTest_Click()
 On Error Resume Next
  Dim conn As ADODB.Connection
  Set conn = New ADODB.Connection
  conn.Provider = "SQLoledb"
  conn.ConnectionString = "Server=" & txtSrv & ";DataBase=" & txtDB & ";UID=" & txtUsr & ";Pwd=" & txtPass & ";"
  conn.open
  If conn.State = ADODB.adStateOpen Then
    conn.Close
    MsgBox "���������� �����������"
  Else
    MsgBox "������ ���������� ����������"
  
  End If
  Set conn = Nothing
End Sub



Private Sub Form_Load()
    txtSrv = GetSetting("ABOL", "PORTALDB", "PORTALSRV", "")
    txtDB = GetSetting("ABOL", "PORTALDB", "PORTALDB", "")
    txtUsr = GetSetting("ABOL", "PORTALDB", "PORTALUSR", "")
    txtPass = GetSetting("ABOL", "PORTALDB", "PORTALPASS", "")
    txtTable = GetSetting("ABOL", "PORTALDB", "PORTALTABLE", "S25")
End Sub

Private Sub OKButton_Click()
  Dim conn As ADODB.Connection
  Set conn = New ADODB.Connection
  conn.Provider = "SQLoledb"
  conn.ConnectionString = "Server=" & txtSrv & ";DataBase=" & txtDB & ";UID=" & txtUsr & ";Pwd=" & txtPass & ";"
  conn.open
  If conn.State = ADODB.adStateOpen Then
    conn.Close
    
    SaveSetting "ABOL", "PORTALDB", "PORTALSRV", txtSrv
    SaveSetting "ABOL", "PORTALDB", "PORTALDB", txtDB
    SaveSetting "ABOL", "PORTALDB", "PORTALUSR", txtUsr
    SaveSetting "ABOL", "PORTALDB", "PORTALPASS", txtPass
    SaveSetting "ABOL", "PORTALDB", "PORTALTABLE", txtTable
    
  Else
    MsgBox "�������� ��������� ����������"
  End If
  Set conn = Nothing
  Me.Hide
End Sub
