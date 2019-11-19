VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ReportShow"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit
'Класс для отображения отчетов

Public Caption As String
Public ReportPath As String
Public ReportSource As String
Public ReportFilter As String
Public PrinterName As String

Private m_Formulas As Formulas
Private frm As frmReport
Public EnableTree As Boolean

'Коллекция формул
'Parameters:
' параметров нет
'Returns:
'  объект класса Formulas
'  ,или Nothing
'See Also:
'  Caption
'  ExportPDF
'  ExportPDFDirectRS
'  PrinterName
'  PrintOut
'  ReportFilter
'  ReportPath
'  ReportSource
'  Run
'  RunDirectRS
'Example:
' dim variable as Formulas
' Set variable = me.Formulas()
Public Function Formulas() As Formulas
  Set Formulas = m_Formulas
End Function

Private Sub Class_Initialize()
  Set m_Formulas = New Formulas
  Set frm = New frmReport
End Sub

Private Sub Class_Terminate()
  On Error Resume Next
  Unload frm
  Set m_Formulas = Nothing
  Set frm = Nothing
End Sub

'Запуск отчета
'Parameters:
'[IN]   Modal , тип параметра: Boolean = False,
'[IN]   conn , тип параметра: ADODB.Connection = Nothing  - явное указание базы данных, источника данных
'See Also:
'  Caption
'  ExportPDF
'  ExportPDFDirectRS
'  Formulas
'  PrinterName
'  PrintOut
'  ReportFilter
'  ReportPath
'  ReportSource
'  RunDirectRS
'Example:
'  call me.Run(...параметры...)
Public Sub Run(Optional ByVal Modal As Boolean = False, Optional ByVal conn As ADODB.Connection = Nothing)
  Dim crX As CRAXDRT.Application
  Dim adoRs As ADODB.Recordset
  On Error GoTo bye
  Set crX = New CRAXDRT.Application
  Dim rpt As CRAXDRT.Report
  Set rpt = crX.OpenReport(ReportPath)
  
  rpt.DiscardSavedData
  
  Dim ff As FormulaFieldDefinition
  Dim fd As FormulaData
  For Each fd In Formulas
    Set ff = rpt.FormulaFields.GetItemByName(fd.Name)
    ff.Text = fd.Expression
  Next
  
  If conn Is Nothing Then
  
    Set adoRs = Session.GetRows(ReportSource, "", "", ReportFilter)
  Else
    If ReportFilter <> "" Then
      Set adoRs = conn.Execute("select * from " & ReportSource & " where  " & ReportFilter)
    Else
     Set adoRs = conn.Execute("select * from " & ReportSource)
    End If
  End If
  
  'Set adoRs = Session.GetRows(ReportSource, "", "", ReportFilter)
  
  rpt.Database.SetDataSource adoRs
  
  If PrinterName <> "" Then
    Dim P As Printer
    For Each P In Printers
      If P.DeviceName = PrinterName Then
        rpt.SelectPrinter P.DriverName, P.DeviceName, P.Port
        Exit For
      End If
    Next
  End If
  rpt.Database.Verify
  frm.Caption = Caption
  
  frm.CRViewer1.EnableGroupTree = EnableTree
  Set frm.rpt = rpt
  frm.CRViewer1.ReportSource = rpt
  frm.CRViewer1.ViewReport
  If Modal Then
    frm.Show vbModal
  Else
    frm.Show
  End If
  
  frm.WindowState = 0
  Exit Sub
bye:
  MsgBox Err.Description, vbOKOnly + vbCritical, "Отчет"
End Sub

'Запуск отчета с готовым рекордсетом
'Parameters:
'[IN]   rs , тип параметра: ADODB.Recordset - данные для отображения,
'[IN]   Modal , тип параметра: Boolean = False  - ...
'See Also:
'  Caption
'  ExportPDF
'  ExportPDFDirectRS
'  Formulas
'  PrinterName
'  PrintOut
'  ReportFilter
'  ReportPath
'  ReportSource
'  Run
'Example:
'  call me.RunDirectRS(...параметры...)
Public Sub RunDirectRS(ByVal rs As ADODB.Recordset, Optional ByVal Modal As Boolean = False)
  Dim crX As CRAXDRT.Application
  Dim adoRs As ADODB.Recordset
  On Error GoTo bye
  Set crX = New CRAXDRT.Application
  Dim rpt As CRAXDRT.Report
  Set rpt = crX.OpenReport(ReportPath)
  
  rpt.DiscardSavedData
  
  Dim ff As FormulaFieldDefinition
  Dim fd As FormulaData
  For Each fd In Formulas
    Set ff = rpt.FormulaFields.GetItemByName(fd.Name)
    ff.Text = fd.Expression
  Next
  
  
  rpt.Database.SetDataSource rs
  
  If PrinterName <> "" Then
    Dim P As Printer
    For Each P In Printers
      If P.DeviceName = PrinterName Then
        rpt.SelectPrinter P.DriverName, P.DeviceName, P.Port
        Exit For
      End If
    Next
  End If
  rpt.Database.Verify
  frm.Caption = Caption
  frm.CRViewer1.EnableGroupTree = EnableTree
  Set frm.rpt = rpt
  frm.CRViewer1.ReportSource = rpt
  frm.CRViewer1.ViewReport
  If Modal Then
    frm.Show vbModal
  Else
    frm.Show
  End If
  
  frm.WindowState = 0
  Exit Sub
bye:
  MsgBox Err.Description, vbOKOnly + vbCritical, "Отчет"
End Sub

' Печать отчета
'Parameters:
'[IN]   conn , тип параметра: ADODB.Connection = Nothing  - явный источник данных ( база данных)
'See Also:
'  Caption
'  ExportPDF
'  ExportPDFDirectRS
'  Formulas
'  PrinterName
'  ReportFilter
'  ReportPath
'  ReportSource
'  Run
'  RunDirectRS
'Example:
'  call me.PrintOut(...параметры...)
Public Sub PrintOut(Optional ByVal conn As ADODB.Connection = Nothing)
  Dim crX As CRAXDRT.Application
  Dim adoRs As ADODB.Recordset
  On Error GoTo bye
  Set crX = New CRAXDRT.Application
  Dim rpt As CRAXDRT.Report
  Set rpt = crX.OpenReport(ReportPath)
  rpt.DiscardSavedData
  
  Dim ff As FormulaFieldDefinition
  Dim fd As FormulaData
  For Each fd In Formulas
    Set ff = rpt.FormulaFields.GetItemByName(fd.Name)
    ff.Text = fd.Expression
  Next
  
  If conn Is Nothing Then
  
    Set adoRs = Session.GetRows(ReportSource, "", "", ReportFilter)
  Else
    If ReportFilter <> "" Then
      Set adoRs = conn.Execute("select * from " & ReportSource & " where  " & ReportFilter)
    Else
     Set adoRs = conn.Execute("select * from " & ReportSource)
    End If
  End If
  
  rpt.Database.SetDataSource adoRs
  
  
  rpt.Database.Verify
  
  If PrinterName <> "" Then
    Dim P As Printer
    For Each P In Printers
      If P.DeviceName = PrinterName Then
        
'        Set Printer = P
'        Printer.Orientation = 2
'        Dim w As Double, h As Double
'        Printer.PaperSize = vbPRPSA4
'        h = Printer.Height
'        w = Printer.Width
'
'        Printer.PaperSize = 256
'
'        Printer.Width = h
'        Printer.Height = w
 ''       rpt.SelectPrinter Printer.DriverName, Printer.DeviceName, Printer.port
        rpt.SelectPrinter P.DriverName, P.DeviceName, P.Port
        
        Exit For
      End If
    Next
  End If
  
  rpt.PrintOut False, 1
  
  frm.Caption = Caption

  frm.CRViewer1.ReportSource = rpt
  frm.CRViewer1.PrintReport
  frm.Show vbModal
  'frm.WindowState = 0
  Exit Sub
bye:
  MsgBox Err.Description, vbOKOnly + vbCritical, "Отчет"
End Sub

' вывод отчета в PDF
'Parameters:
'[IN]   FileName , тип параметра: String,
'[IN]   conn , тип параметра: ADODB.Connection = Nothing  - ...
'See Also:
'  Caption
'  ExportPDFDirectRS
'  Formulas
'  PrinterName
'  PrintOut
'  ReportFilter
'  ReportPath
'  ReportSource
'  Run
'  RunDirectRS
'Example:
'  call me.ExportPDF(...параметры...)
Public Sub ExportPDF(ByVal fileName As String, Optional ByVal conn As ADODB.Connection = Nothing)
  Dim crX As CRAXDRT.Application
  Dim adoRs As ADODB.Recordset
  On Error GoTo bye
  Set crX = New CRAXDRT.Application
  Dim rpt As CRAXDRT.Report
  Set rpt = crX.OpenReport(ReportPath)
  rpt.DiscardSavedData
  
  Dim ff As FormulaFieldDefinition
  Dim fd As FormulaData
  For Each fd In Formulas
    Set ff = rpt.FormulaFields.GetItemByName(fd.Name)
    ff.Text = fd.Expression
  Next
  
  If conn Is Nothing Then
  
    Set adoRs = Session.GetRows(ReportSource, "", "", ReportFilter)
  Else
    If ReportFilter <> "" Then
      Set adoRs = conn.Execute("select * from " & ReportSource & " where  " & ReportFilter)
    Else
     Set adoRs = conn.Execute("select * from " & ReportSource)
    End If
  End If
  
  rpt.Database.SetDataSource adoRs
    
  rpt.Database.Verify
  
    
  rpt.ExportOptions.PDFExportAllPages = True
  rpt.ExportOptions.PDFFirstPageNumber = 1
  rpt.ExportOptions.DestinationType = crEDTDiskFile
  rpt.ExportOptions.DiskFileName = fileName
  rpt.ExportOptions.FormatType = crEFTPortableDocFormat
  'rpt.ExportOptions.PromptForExportOptions
  rpt.EnableParameterPrompting = False
  rpt.DisplayProgressDialog = False
  rpt.Export False
  
  Exit Sub
bye:
  MsgBox Err.Description, vbOKOnly + vbCritical, "Отчет"
End Sub

'Вывод отчета в PDF, по готовой выборке данных
'Parameters:
'[IN]   FileName , тип параметра: String,
'[IN]   adoRs , тип параметра: ADODB.Recordset  - готовая выборка данных
'See Also:
'  Caption
'  ExportPDF
'  Formulas
'  PrinterName
'  PrintOut
'  ReportFilter
'  ReportPath
'  ReportSource
'  Run
'  RunDirectRS
'Example:
'  call me.ExportPDFDirectRS(...параметры...)
Public Sub ExportPDFDirectRS(ByVal fileName As String, ByVal adoRs As ADODB.Recordset)
  Dim crX As CRAXDRT.Application
 
  On Error GoTo bye
  Set crX = New CRAXDRT.Application
  Dim rpt As CRAXDRT.Report
  Set rpt = crX.OpenReport(ReportPath)
  rpt.DiscardSavedData
  
  Dim ff As FormulaFieldDefinition
  Dim fd As FormulaData
  For Each fd In Formulas
    Set ff = rpt.FormulaFields.GetItemByName(fd.Name)
    ff.Text = fd.Expression
  Next
   
  rpt.Database.SetDataSource adoRs
  
  rpt.Database.Verify
  
    
  rpt.ExportOptions.PDFExportAllPages = True
  rpt.ExportOptions.PDFFirstPageNumber = 1
  rpt.ExportOptions.DestinationType = crEDTDiskFile
  rpt.ExportOptions.DiskFileName = fileName
  rpt.ExportOptions.FormatType = crEFTPortableDocFormat
  rpt.EnableParameterPrompting = False
  rpt.DisplayProgressDialog = False
  rpt.Export False
  
  Exit Sub
bye:
  MsgBox Err.Description, vbOKOnly + vbCritical, "Экспорт"
End Sub


