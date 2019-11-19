VERSION 5.00
Object = "{E9DB983E-3879-4902-8162-947677DC197D}#12.0#0"; "CRViewer.dll"
Begin VB.Form frmReport 
   ClientHeight    =   6285
   ClientLeft      =   1650
   ClientTop       =   1545
   ClientWidth     =   8880
   Icon            =   "frmReport.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   6285
   ScaleWidth      =   8880
   ShowInTaskbar   =   0   'False
   Begin CrystalActiveXReportViewerLib12Ctl.CrystalActiveXReportViewer CRViewer1 
      Height          =   5535
      Left            =   120
      TabIndex        =   1
      Top             =   600
      Width           =   8535
      _cx             =   15055
      _cy             =   9763
      DisplayGroupTree=   0   'False
      DisplayToolbar  =   -1  'True
      EnableGroupTree =   -1  'True
      EnableNavigationControls=   -1  'True
      EnableStopButton=   0   'False
      EnablePrintButton=   -1  'True
      EnableZoomControl=   -1  'True
      EnableCloseButton=   0   'False
      EnableProgressControl=   0   'False
      EnableSearchControl=   -1  'True
      EnableRefreshButton=   0   'False
      EnableDrillDown =   -1  'True
      EnableAnimationControl=   0   'False
      EnableSelectExpertButton=   0   'False
      EnableToolbar   =   -1  'True
      DisplayBorder   =   -1  'True
      DisplayTabs     =   0   'False
      DisplayBackgroundEdge=   -1  'True
      SelectionFormula=   ""
      EnablePopupMenu =   -1  'True
      EnableExportButton=   -1  'True
      EnableSearchExpertButton=   0   'False
      EnableHelpButton=   0   'False
      LaunchHTTPHyperlinksInNewBrowser=   -1  'True
      EnableLogonPrompts=   -1  'True
      LocaleID        =   1049
      EnableInteractiveParameterPrompting=   0   'False
   End
   Begin VB.CommandButton cmdPrnSetup 
      Caption         =   "Настройка принтера"
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   3255
   End
End
Attribute VB_Name = "frmReport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
' Окно для вывода формы отчета

' собственно отчет
Public rpt As CRAXDDRT.Report

' настройка параметров принтера
Private Sub cmdPrnSetup_Click()
  rpt.PrinterSetupEx Me.hwnd
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  If UnloadMode = vbFormMDIForm Or UnloadMode = vbFormCode Or UnloadMode = vbAppWindows Or UnloadMode = vbAppTaskManager Then
    Cancel = False
  Else
    Cancel = True
    Me.Hide
  End If

End Sub

Private Sub Form_Resize()
    On Error Resume Next
    'CRViewer1.Top = 0
    CRViewer1.Left = 0
    CRViewer1.Height = Me.ScaleHeight - CRViewer1.Top
    CRViewer1.Width = Me.ScaleWidth
End Sub




