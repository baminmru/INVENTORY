Attribute VB_Name = "Module1"


Dim ArgArray() As String
Private Type ACCDEF
  theorg As String
  FromPath As String
'  FileType As String
'  ToPath As String
'  renameTo As String
  used As Integer
End Type



Public Status As String
Dim site As String
Dim login As String
Dim password As String
Dim txtFile As String
Dim logPath As String
Dim defFile As String
Dim ad() As ACCDEF, maxAD As Long, P() As String, maxP As Long
Public W As Writer


Public Sub Main()
Dim cnt As Integer, i As Integer
Dim arr() As String
  
  logPath = "c:\XLSRESAVE.log"
  
  If Command = "" Then
   Dim f As frm
   Set f = New frm
   f.Show vbModal
   If f.OK Then
   
    defFile = f.txtFile
    If f.txtLog <> "" Then
      logPath = f.txtLog
    End If
   Else
    MsgBox "Вызов: XLSRESAVE paramfile LOG:<log file>", vbOKOnly, "Пересохранение Excel файлов"
    End
   End If
   Unload f
   Set f = Nothing
  Else
   
    
    cnt = GetCommandLine
    defFile = ArgArray(1)
    For i = 2 To cnt
      arr() = Split(ArgArray(i), ":")
      If UBound(arr) >= 1 Then
      If UCase(arr(0)) = "LOG" Then
        logPath = Mid(ArgArray(i), 5)
      End If
    End If
    Next
    
    End If
    
    Set W = New Writer
    If logPath <> "" Then
      W.SetFilePath logPath
    End If
    
    
  
    ProcessAll defFile
 
    SetStatus "Завершение работы"
    
  
    Set W = Nothing
    End
  
End Sub


Function GetCommandLine(Optional MaxArgs) As Long
    On Error Resume Next
    'Declare variables.
    Dim c, CmdLine, CmdLnLen, InArg, i, NumArgs
    'See if MaxArgs was provided.
    If IsMissing(MaxArgs) Then MaxArgs = 20
    'Make array of the correct size.
    ReDim ArgArray(MaxArgs)
    NumArgs = 0: InArg = 0
    'Get command line arguments.
    CmdLine = Command()
    CmdLnLen = Len(CmdLine)
    'Go thru command line one character
    'at a time.
    For i = 1 To CmdLnLen
        c = Mid(CmdLine, i, 1)

        'Test for space or tab.
        If (InArg < 2 And c <> " " And c <> vbTab) Or _
            (InArg = 2 And c <> """") Or _
            (InArg = 3 And c <> "'") Then
        
            'Neither space nor tab.
            'Test if already in argument.
            If InArg = 0 Then
            'New argument begins.
            'Test for too many arguments.
                If NumArgs = MaxArgs Then Exit For
                NumArgs = NumArgs + 1
                If c = """" Then
                    InArg = 2
                    GoTo nnn
                ElseIf c = "'" Then
                    InArg = 3
                    GoTo nnn
                Else
                    InArg = 1
                End If
                ArgArray(NumArgs) = ""
            End If
            'Concatenate character to current argument.
            ArgArray(NumArgs) = ArgArray(NumArgs) & c
        Else
            
            InArg = 0
        End If
nnn:
    Next i
    'Resize array just enough to hold arguments.
    ReDim Preserve ArgArray(NumArgs)
    
    'Return Array in Function name.
    GetCommandLine = NumArgs
End Function





Public Sub ProcessAll(ByVal FilePath As String)
  Dim fd As Long, i, j, mypath, myname, str
  Dim t As Boolean
  If FilePath = "" Then Exit Sub
  Status = ""
  On Error GoTo noFile
  fd = FreeFile
  Open FilePath For Input Access Read As #fd
  On Error GoTo closefile
  maxAD = 1
  SetStatus "Читаем настроечный файл"
  Do While Not EOF(fd)
    ReDim Preserve ad(maxAD)
    Input #fd, ad(maxAD).theorg, ad(maxAD).FromPath   ', ad(maxAD).renameTo
    ad(maxAD).used = 0
    SetStatus "Читаем настроечный файл строка № " & maxAD
    maxAD = maxAD + 1
  Loop
  
  Close fd
  On Error GoTo noFile

  
  'loop for all known path
  SetStatus "Сканируем директории"
  For i = 1 To maxAD - 1
    SetStatus "Сканируем директорию " & ad(i).FromPath
    mypath = ad(i).FromPath
    myname = Dir(mypath & "\*.xls", vbNormal)
    Do While myname <> ""
      If myname <> "." And myname <> ".." Then
       
        Dim spath As String
        spath = mypath & "\" & myname
        SetStatus "Файл " & spath
        
        ResaveFile spath
        
'        Dim md5 As String
'        Dim ccc As CMD5
'        Set ccc = New CMD5
'        md5 = ccc.FileMD5(spath)
'        Set ccc = Nothing
'
'
'
'        If Not IsFileLoaded(spath, md5, "МАТЕРИАЛЫ") Then
'          If Not LoadXLS_MAT(spath, ad(i).theorg) Then
'            If Not IsFileLoaded(spath, md5, "ОС") Then
'              If LoadXLS_OS(spath, ad(i).theorg) Then
'                Set ccc = New CMD5
'                md5 = ccc.FileMD5(spath)
'                Set ccc = Nothing
'                RegisterFile spath, md5, "ОС"
'                SetStatus "Регистрация файла"
'              End If
'            Else
'              SetStatus "Файл уже загружался"
'            End If
'          Else
'           Set ccc = New CMD5
'           md5 = ccc.FileMD5(spath)
'           Set ccc = Nothing
'           RegisterFile spath, md5, "МАТЕРИАЛЫ"
'           SetStatus "Регистрация файла"
'          End If
'        Else
'           SetStatus "Файл уже загружался"
'        End If
'
      End If
nxtName:
      myname = Dir  ' Get next entry.
    Loop
  Next
  SetStatus "Закончено сканирование"
  Exit Sub
  
   On Error GoTo noFile
  ' make list of source paths
  
closefile:
    Close fd
    Resume Next
    Exit Sub
noFile:
SetStatus Err.Description
 'MsgBox Err.Description
End Sub

Public Sub SetStatus(ByVal s As String)
  W.putBuf s
  Debug.Print s
  DoEvents
End Sub


Private Sub ResaveFile(path As String)
On Error Resume Next
  Dim ex As Object 'Excel.Application
  Set ex = CreateObject("Excel.Application")
  Dim wb As Object 'Excel.Workbook
  SetStatus "Open " & path
  Set wb = ex.Workbooks.Open(path)
  Dim npath As String
  npath = Replace(LCase(path), ".xls", "_new.xls")
  Call wb.SaveAs(FileName:=npath, _
        FileFormat:=-4143, password:="", WriteResPassword:="", _
        ReadOnlyRecommended:=False, CreateBackup:=False)
   SetStatus "Save " & npath
  wb.Close True
  Set wb = Nothing
  ex.Quit
  If Err.Number = 0 Then
   Kill path
   SetStatus "Delete " & path
  End If
  
End Sub



