Attribute VB_Name = "Module1"
Option Explicit
Public Manager As MTZManager.Main
Public Session As MTZSession.Session
Public UsersID As String
Public UserName As String
Public UserPassword As String
Public PrivateStoreID As String
Public SysStoreID As String
Public Site As String
Public LastChat As Date
Public NextReminder As Date
Public DeltaReminder As String
Public usr As MTZUsers.Application
Public MyUser As MTZUsers.Users
Public fProg As frmProgress


Public Declare Function AddFontResource Lib "gdi32" Alias "AddFontResourceA" (ByVal lpFileName As String) As Long
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Public Declare Function WriteProfileString Lib "kernel32" Alias "WriteProfileStringA" (ByVal lpszSection As String, ByVal lpszKeyName As String, ByVal lpszString As String) As Long
Public Const WM_FONTCHANGE = &H1D
Public Const HWND_BROADCAST = &HFFFF&





Sub Main()
Dim par() As String
Dim i As Long
Dim tst As Long
Dim UserPassword As String
Set Manager = New MTZManager.Main

tst = 0
  If Command$ <> "" Then
        par() = Split(Command, " ")
        For i = LBound(par) To UBound(par)
          If UCase(Left(par(i), 4)) = "USR:" Then
            UserName = Right(par(i), Len(par(i)) - 4)
            tst = tst + 1
          End If
          
          If UCase(Left(par(i), 4)) = "PWD:" Then
            UserPassword = Right(par(i), Len(par(i)) - 4)
            tst = tst + 1
          End If
          
          If UCase(Left(par(i), 4)) = "APP:" Then
            Site = Right(par(i), Len(par(i)) - 4)
            tst = tst + 1
          End If

        Next
        If tst = 3 Then
          Set Session = Manager.GetSession(Site)
          If Session Is Nothing Then
            GoTo useForm
          End If
          
          If Not Session.Login(UserName, UserPassword) Then
            Set Session = Nothing
            GoTo useForm
          End If
        Else
         GoTo useForm
        End If
  Else
  
useForm:
    Dim f As frmLogin
    Set f = New frmLogin

again:
    Set Session = Nothing
    Set Manager = Nothing
    Set Manager = New MTZManager.Main
    
    f.Show vbModal
    If Not f.OK Then
      Unload f
      Set f = Nothing
      Set Manager = Nothing
      Exit Sub
    End If
    Site = f.txtSite
    
    Set Session = Manager.GetSession(Site)
    If Session Is Nothing Then
      MsgBox "Ќе определен сайт с таким именем", vbCritical, "ќшибка"
      GoTo again
    End If
    
    
    
    If Not Session.Login(f.txtUserName, f.txtPassword) Then
      Set Session = Nothing
      MsgBox "Ќеверные данные регистрации", vbCritical, "ќшибка"
      GoTo again
    End If
    UserName = f.txtUserName
    UserPassword = f.txtPassword
    Unload f
    Set f = Nothing
 
 End If
 
  
  Dim rs As ADODB.Recordset
  Set rs = Manager.ListInstances(Site, "MTZUsers")
  Set usr = Manager.GetInstanceObject(rs!InstanceID)
  Manager.LockInstanceObject usr.id
  
  
  Set rs = Nothing
  Set MyUser = usr.FindRowObject("Users", Session.GetSessionUserID())
  Set rs = Nothing
  
  Set MyRole = ChooseRole()
  If MyRole Is Nothing Then
      Session.Logout
     Set Manager = Nothing
     Exit Sub
  End If
  
  Manager.LockInstanceObject MyRole.id
  
  frmSplash.Show
  frmSplash.lblWarning = "«агрузка умолчаний"
  DoEvents
  
   
  
  Dim orgid As String
  
    
   
  frmSplash.lblWarning = "«агрузка лицензий"
  DoEvents
  On Error Resume Next
   Dim intFile As Integer
   intFile = FreeFile
   Open App.path & "\Licenses.txt" For Input As #intFile
   Dim strKey As String, strprogid As String
   ' On the client machine, read the license key from the file.
   
   
   While Not EOF(intFile)
    strprogid = ""
    strKey = ""
    Input #intFile, strprogid, strKey
    If strprogid <> "" Then
      Licenses.Add strprogid, strKey
    Else
      GoTo closefile
    End If
   Wend

closefile:
   Close #intFile
   
   
  frmSplash.lblWarning = "ѕодключение документов"
  DoEvents
  
  RegisterMDIGUI
  
  frmSplash.lblWarning = "ѕересчет сроков использовани€"
  DoEvents
  CalcSrok
  
  
  frmSplash.lblWarning = "ќбслуживание базы"
  DoEvents
  frmSplash.DBMainteice
  
  
  frmSplash.lblWarning = "»нициализаци€ меню"
  DoEvents
  Load frmMain
  
  Unload frmSplash
  
  Call Manager.AddCustomObjects(MyRole, "ROLE")
  Call Manager.AddCustomObjects(MyUser, "USER")
  
  frmMain.Show
  
End Sub


Public Sub PrintGrid(gr As Object)
  
  Dim r As RECT
  Dim ph As Long, pw As Long
  Dim i As Long, j As Long
  Dim ColPerPage() As Long, HorPages As Long, curw As Long
  Dim CurRow As Long, CurCol As Long, FirstRow As Long, CellTop As Long
  Dim dx As Double, dy As Double, pcnt As Long

  ph = Printer.ScaleHeight - 1000
  pw = Printer.ScaleWidth - 200
  dx = 1.1
  dy = 1.1
  pcnt = 0

  ' считаем сколько страниц надо по ширине
  curw = 0
  HorPages = 1
  ReDim ColPerPage(HorPages)
  ColPerPage(HorPages) = 0
  For i = 0 To gr.Cols - 1
    If gr.ColWidth(i) > 0 Then curw = curw + gr.ColWidth(i) * dx

    ' ширина превысила размер страницы
    If curw > pw Then
      HorPages = HorPages + 1
      ReDim Preserve ColPerPage(HorPages)
      ColPerPage(HorPages) = IIf(i - 1 < 1, 1, i - 1)
      curw = gr.ColWidth(i) * dx
    End If

    ' если колонка очень широка€ то запихаем ее в отдельную страницу
    If i > 0 And curw > pw Then
      HorPages = HorPages + 1
      ReDim Preserve ColPerPage(HorPages)
      ColPerPage(HorPages) = i
      curw = 0
    End If
  Next
  ReDim Preserve ColPerPage(HorPages + 1)
  ColPerPage(HorPages + 1) = gr.Cols

  CurCol = 0
  CurRow = 0
  FirstRow = 0
  Printer.Font.Name = gr.Font.Name
  Printer.Font.Bold = gr.Font.Bold
  Printer.Font.Charset = gr.Font.Charset
  Printer.Font.Italic = gr.Font.Italic
  Printer.Font.Strikethrough = gr.Font.Strikethrough
  Printer.Font.Underline = gr.Font.Underline
  Printer.Font.Weight = gr.Font.Weight
  Printer.Font.Size = gr.Font.Size

  ' цикл по вертикальным блокам
  While FirstRow < gr.Rows

    ' √оризонтальный блок страниц
    For i = 1 To HorPages
      curw = 0

      ' колонки дл€ каждой из страниц
      For j = ColPerPage(i) To ColPerPage(i + 1) - 1

        ' только видимые колонки
        If gr.ColWidth(j) > 0 Then
          CellTop = 0
          CurRow = FirstRow

          ' ограничение по высоте листа
          While CellTop <= ph

              ' не проходим по высоте листа
              If CellTop + gr.RowHeight(CurRow) * dy > ph Then
                If gr.RowHeight(CurRow) * dy > ph Then
                  ' если высота колонки очень велика то мен€ем ее на меньшую
                  gr.RowHeight(CurRow) = ph / dy
                  GoTo nxtcol
                Else
                  GoTo nxtcol
                End If
              End If

              ' пересчитываем пр€моугольник дл€ отрисовки текста
              r.Left = curw / Printer.TwipsPerPixelX + 2
              r.Right = IIf((curw + gr.ColWidth(j) * dx) > pw, pw, curw + gr.ColWidth(j) * dx) _
                / Printer.TwipsPerPixelX - 2
              r.Top = CellTop / Printer.TwipsPerPixelY + 2
              r.Bottom = (CellTop + gr.RowHeight(CurRow) * dy) / Printer.TwipsPerPixelY - 2

              ' ѕервую строку отдел€ем жирной линией
              If CurRow = 0 Then
                Printer.Line (curw, (CellTop + gr.RowHeight(CurRow) * dy) - 20)- _
                  (IIf((curw + gr.ColWidth(j) * dx) > pw, pw, curw + gr.ColWidth(j) * dx), _
                  (CellTop + gr.RowHeight(CurRow) * dy)), , BF
              End If


              ' выводим рамочку
              Printer.Line (curw, CellTop)- _
                (IIf((curw + gr.ColWidth(j) * dx) > pw, pw, curw + gr.ColWidth(j) * dx), _
                (CellTop + gr.RowHeight(CurRow) * dy)), , B


              ' выводим текст в пр€моугольную область (с переносом слов)
              DrawText Printer.hdc, gr.TextMatrix(CurRow, j), Len(gr.TextMatrix(CurRow, j)), r, &H10 + &H100

              ' измен€ем позицию дл€ следующей строки
              CellTop = CellTop + gr.RowHeight(CurRow) * dy

              ' готовимс€ к следующей сторке
              CurRow = CurRow + 1
              If CurRow >= gr.Rows Then GoTo nxtcol

          Wend
nxtcol:
          ' учитываем ширину и переходим к следующей колонке
          curw = curw + gr.ColWidth(j) * dx
        End If
      Next ' цикл по колонкам


      ' печатаем номер страницы
      Printer.Line (0, ph - 20)-(Printer.ScaleWidth, ph), , B
      Printer.CurrentX = Printer.ScaleWidth / 3
      Printer.CurrentY = ph + 100
      pcnt = pcnt + 1
      Printer.Print "—траница є" & pcnt
      ' не отбиваем страницу после последнего листа
      If CurRow < gr.Rows Or i < HorPages Then Printer.NewPage
    Next
    ' готовимс€ к новому блоку горизонтальных страниц
    FirstRow = CurRow
  Wend
  Printer.EndDoc
End Sub



Private Sub RegisterMDIGUI()
 Dim g As GUI
Set g = New GUI
g.Init "INV_INV"
Manager.RegisterGUI g, "INV_INV"
Set g = New GUI
g.Init "INV_OS"
Manager.RegisterGUI g, "INV_OS"
Set g = New GUI
g.Init "INV_DIC"
Manager.RegisterGUI g, "INV_DIC"
Set g = New GUI
g.Init "INV_NUM"
Manager.RegisterGUI g, "INV_NUM"
Set g = New GUI
g.Init "INVF"
Manager.RegisterGUI g, "INVF"

End Sub


Public Sub InstallFont(ByVal FontPath As String, ByVal FontName As String, ByVal FontFileName)
    Dim ret As Integer
    On Error Resume Next
    ret = AddFontResource(FontPath)
    If ret = 1 Then
        ret = SendMessage(HWND_BROADCAST, WM_FONTCHANGE, 0, 0)
        ret = WriteProfileString("fonts", FontName + " (TrueType)", FontFileName)
    End If
End Sub


Public Sub SaveHistory(RowItem As INV_OS.INVOS_PLACE)
On Error GoTo bye
'''''''''''''''' change me

 Dim cRow As INV_OS.INVOS_PLACE
 Dim hist As INV_OS.INVOS_HIST
 Set cRow = RowItem
 Set hist = cRow.Application.INVOS_HIST.Add
 
 
 With hist
    .UntilDate = Now
    Set .ChangedBy = cRow.Application.FindRowObject("Users", cRow.Application.MTZSession.GetSessionUserID)
    'Set .TheOrg = cRow.TheOrg
    .ComplNumber = cRow.ComplNumber
    Set .Direction = cRow.Direction
    Set .Uprav = cRow.Uprav
    Set .Otdel = cRow.Otdel
    Set .TheHouse = cRow.TheHouse
    Set .MatOtv = cRow.MatOtv
    .Flow = cRow.Flow
    .Room = cRow.Room
    .WorkPlaceNum = cRow.WorkPlaceNum
    Set .TheOwner = cRow.TheOwner
    .save
 End With



 Exit Sub
bye:
  
End Sub


Public Sub CalcSrok()
  Dim rs As ADODB.Recordset
  Set rs = Session.GetData("select instanceid from v_autoinvos_info where invos_info_ismaterial_val=0 and  statusname <>'—писано' and invos_srok_recalcdate <=" & IIf(Session.IsMSSQL, MakeMSSQLDate(Date), IIf(Session.IsORACLE, MakeORACLEDate(Date), MakePGSQLDate(Date))))
  
  Dim os As INV_OS.Application
  Dim isrk As INV_OS.INVOS_SROK
  While Not rs.EOF
    Set os = Manager.GetInstanceObject(rs!InstanceID)
    With os.INVOS_SROK.Item(1)
       .RecalcDate = DateAdd("m", 1, DateSerial(Year(Date), Month(Date), 1))
       .save
    End With
    
    With os.INVOS_INFO.Item(1)
      .SrokFI = .SrokFI + 1
      If .SrokOI > 0 Then
        .SrokOI = .SrokOI - 1
      End If
      .save
    End With
    
    Manager.FreeAllInstanses
    rs.MoveNext
  Wend
  
  rs.Close
  Set rs = Nothing

End Sub

Public Function GetCompl(Name As String) As String
  Dim i As Integer
  Dim spcPos As Integer
  spcPos = -1
  For i = Len(Name) To 1 Step -1
    If Mid(Name, i, 1) = " " Then
      spcPos = i
      Exit For
    End If
  Next
  If spcPos < 0 Then
    GetCompl = ""
    Exit Function
  End If
  If Len(Name) - spcPos > 10 Then
    GetCompl = ""
    Exit Function
  End If
  
  Dim cnum As String
  cnum = Mid(Name, spcPos + 1, Len(Name) - spcPos)
  If IsNumeric(Replace(Replace(cnum, ".", ""), ",", "")) Then
    If InStr(1, cnum, ".") > 0 Or InStr(1, cnum, ",") > 0 Then
      GetCompl = Replace(cnum, ",", ".")
    End If
  End If

End Function


Public Sub RegisterFile(ByVal fileName As String, ByVal md5 As String, ByVal FileType As String)

  Dim invf As invf.Application
  Dim id As String
  id = CreateGUID2
  
  Manager.NewInstance id, "INVF", fileName
  Set invf = Manager.GetInstanceObject(id)
  With invf.INVF_DEF.Add
    .ThePath = fileName
    .TheHash = md5
    Set .TheUser = MyUser
    .TypeOfFile = FileType
    .Loaddate = Now
    .save
  End With
End Sub


Public Function IsFileLoaded(ByVal fileName As String, ByVal md5 As String, ByVal FileType As String) As Boolean
  Dim res As Boolean
  Dim rs As ADODB.Recordset
  Set rs = Session.GetData("select count(*) cnt from INVF_DEF where thePath ='" & fileName & "' and theHash='" & md5 & "' and TypeOfFile='" & FileType & "'")
  If rs Is Nothing Then
    res = False
  ElseIf rs.EOF Then
    res = False
  ElseIf rs!cnt = 0 Then
    res = False
  Else
    res = True
  End If
  rs.Close
  Set rs = Nothing
  IsFileLoaded = res
End Function



