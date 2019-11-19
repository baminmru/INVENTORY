Imports System.Data.Common
Imports System.Data
Imports System.Data.SQLite


Public Class frmScan

    Public InvCode As String
    Private inv_instanceid As String
    Private instanceid As String


    Public Sub New()

        ' This call is required by the Windows Form Designer.
        InitializeComponent()

        ' Add any initialization after the InitializeComponent() call.
        InitBC()
    End Sub

    Private Sub InitBC()
        Try
            Me.Barcode1.DecoderParameters.CODABAR = Barcode.DisabledEnabled.Disabled
            Me.Barcode1.DecoderParameters.CODE128 = Barcode.DisabledEnabled.Enabled
            Me.Barcode1.DecoderParameters.CODE39 = Barcode.DisabledEnabled.Disabled
            'Me.Barcode1.DecoderParameters.CODE39Params.Code32Prefix = False
            'Me.Barcode1.DecoderParameters.CODE39Params.Concatenation = False
            'Me.Barcode1.DecoderParameters.CODE39Params.ConvertToCode32 = False
            'Me.Barcode1.DecoderParameters.CODE39Params.FullAscii = False
            'Me.Barcode1.DecoderParameters.CODE39Params.Redundancy = False
            'Me.Barcode1.DecoderParameters.CODE39Params.ReportCheckDigit = False
            'Me.Barcode1.DecoderParameters.CODE39Params.VerifyCheckDigit = False
            Me.Barcode1.DecoderParameters.D2OF5 = Barcode.DisabledEnabled.Disabled
            Me.Barcode1.DecoderParameters.EAN13 = Barcode.DisabledEnabled.Disabled
            Me.Barcode1.DecoderParameters.EAN8 = Barcode.DisabledEnabled.Disabled
            Me.Barcode1.DecoderParameters.I2OF5 = Barcode.DisabledEnabled.Disabled
            Me.Barcode1.DecoderParameters.KOREAN_3OF5 = Barcode.DisabledEnabled.Disabled
            Me.Barcode1.DecoderParameters.MSI = Barcode.DisabledEnabled.Disabled
            Me.Barcode1.DecoderParameters.UPCA = Barcode.DisabledEnabled.Disabled
            Me.Barcode1.DecoderParameters.UPCE0 = Barcode.DisabledEnabled.Disabled
            Me.Barcode1.EnableScanner = True
            Me.Barcode1.ScanParameters.BeepFrequency = 2670
            Me.Barcode1.ScanParameters.BeepTime = 200
            Me.Barcode1.ScanParameters.CodeIdType = Barcode.CodeIdTypes.None
            Me.Barcode1.ScanParameters.LedTime = 3000
            Me.Barcode1.ScanParameters.ScanType = Barcode.ScanTypes.Foreground
            Me.Barcode1.ScanParameters.WaveFile = ""
        Catch
        End Try

        Me.Text = "Сбор данных"
        cmbStatus.SelectedIndex = 0
        Try
            sqlite_con = New SQLiteConnection("data source=""/OMK.db3""")
            sqlite_con.Open()

        Catch ex As System.Exception
            MsgBox("Не обнаружен файл  с базой данных, Произведите выгрузку данных из программы инвентаризации.", MsgBoxStyle.Critical + MsgBoxStyle.OkOnly, "Ошибка")
        End Try
    End Sub

    Private Sub Barcode1_OnRead(ByVal sender As Object, ByVal readerData As Symbol.Barcode.ReaderData) Handles Barcode1.OnRead
        If readerData.Result = Symbol.Results.SUCCESS Then


            If readerData.Type = Symbol.Barcode.DecoderTypes.CODE128 Then
                txtCode.Text = readerData.Text
            End If

        End If
    End Sub

    Private Sub ProcessCode()
        txtName.Text = txtCode.Text
    End Sub

    Private Sub MenuItem1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuItem1.Click
        If MsgBox("Закрыть программу?", MsgBoxStyle.YesNo + MsgBoxStyle.Question, "Подтвердите") = MsgBoxResult.Yes Then
            Try
                Me.Barcode1.EnableScanner = False
            Catch
            End Try

            Me.Close()
            Try
                If Not sqlite_con Is Nothing Then
                    sqlite_con.Close()
                    sqlite_con = Nothing
                End If
            Catch
            End Try
        End If
    End Sub

    

    Protected Overrides Sub Finalize()
        Me.Barcode1.EnableScanner = False
        If Not sqlite_con Is Nothing Then
            sqlite_con.Close()
            sqlite_con = Nothing
        End If
        MyBase.Finalize()
    End Sub

    Private Function CheckCode(ByVal Code As String) As Boolean
        ' lbl.TExt = "CC"
        Dim ok As Boolean
        Dim sqlite_da As SQLiteDataAdapter = New SQLiteDataAdapter()
        Dim dataSet As DataSet = New DataSet()
        Dim dr As DataRow
        ok = False
        sqlite_da.SelectCommand = New SQLiteCommand("select * from INV where visiblecode='" + Code + "'", sqlite_con)

        ' lbl.TExt = ' lbl.TExt + "O"

        sqlite_da.Fill(dataSet)
        If dataSet.Tables.Count > 0 Then
            ' lbl.TExt = ' lbl.TExt + "T" + dataSet.Tables.Count.ToString()
            If dataSet.Tables(0).Rows.Count > 0 Then
                ' lbl.TExt = ' lbl.TExt + "R" + dataSet.Tables(0).Rows.Count.ToString()
                For i = 0 To dataSet.Tables(0).Rows.Count - 1
                    ' lbl.TExt = ' lbl.TExt + "I" + i.ToString()
                    dr = dataSet.Tables(0).Rows.Item(i)
                    ok = True
                    txtName.Text = dr("INVOS_INFO_ShortName").ToString
                    txtOwner.Text = dr("INVOS_PLACE_TheOwner").ToString
                    txtComp.Text = dr("INVOS_PLACE_ComplNumber").ToString
                    txtInfo.Text = dr("INVOS_INFO_Info").ToString
                    If dr("VisibleCode").ToString.Substring(0, 1) = "M" Then
                        txtNum.Text = dr("INVOS_INFO_CArdNum").ToString
                    Else
                        txtNum.Text = dr("INVOS_INFO_InvNum").ToString
                    End If
                    inv_instanceid = dr("INV_INSTANCEID").ToString
                    instanceid = dr("INSTANCEID").ToString
                    Exit For
                Next

            End If
        End If


        ' lbl.TExt = ' lbl.TExt + "C"
        Return ok
    End Function


    Private Sub SaveCode(ByVal Code As String)

        Dim sqlite_cmd As SQLiteCommand
        sqlite_cmd = New SQLiteCommand()
        sqlite_cmd.CommandType = CommandType.Text

        sqlite_cmd.CommandText = "delete from  T where shCode='" + Code + "'"
        sqlite_cmd.Connection = sqlite_con
        sqlite_cmd.ExecuteNonQuery()

        sqlite_cmd = New SQLiteCommand()
        sqlite_cmd.CommandType = CommandType.Text
        sqlite_cmd.CommandText = "insert into T(shCode,Status,INVID,OSID,CheckTime)values ('" _
        + Code + "','" + cmbStatus.Text + "','" + inv_instanceid + "','" + instanceid + "',?)"
        sqlite_cmd.Connection = sqlite_con

        Dim p As DbParameter

        p = sqlite_cmd.CreateParameter()
        sqlite_cmd.Parameters.Add(p)
        p.Value = Now


        sqlite_cmd.ExecuteNonQuery()
        txtCode.Text = ""
        txtName.Text = ""
        txtOwner.Text = ""
        txtComp.Text = ""
        txtNum.Text = ""
        txtInfo.Text = ""

    End Sub


    Private Sub BadCode(ByVal Code As String)

        Dim sqlite_cmd As SQLiteCommand
        sqlite_cmd = New SQLiteCommand()
        sqlite_cmd.CommandType = CommandType.Text


        sqlite_cmd.CommandText = "delete from  B where shCode='" + Code + "'"
        sqlite_cmd.Connection = sqlite_con
        sqlite_cmd.ExecuteNonQuery()

        sqlite_cmd = New SQLiteCommand()
        sqlite_cmd.CommandType = CommandType.Text
        sqlite_cmd.CommandText = "insert into B(shCode,CheckTime)values ('" _
        + Code + "',?)"
        sqlite_cmd.Connection = sqlite_con

        Dim p As DbParameter

        p = sqlite_cmd.CreateParameter()
        sqlite_cmd.Parameters.Add(p)
        p.Value = Now


        sqlite_cmd.ExecuteNonQuery()
        txtCode.Text = ""
        txtName.Text = ""
        txtOwner.Text = ""
        txtComp.Text = ""
        txtNum.Text = ""
        txtInfo.Text = ""

    End Sub


    Private Sub ToExpl(ByVal Code As String)

        Dim sqlite_cmd As SQLiteCommand
        sqlite_cmd = New SQLiteCommand()
        sqlite_cmd.CommandType = CommandType.Text

      

        sqlite_cmd.CommandText = "delete from  EXPL where shCode='" + Code + "'"
        sqlite_cmd.Connection = sqlite_con
        sqlite_cmd.ExecuteNonQuery()


        sqlite_cmd = New SQLiteCommand()
        sqlite_cmd.CommandType = CommandType.Text
        sqlite_cmd.CommandText = "insert into EXPL(shCode,CheckTime)values ('" _
        + Code + "',?)"
        sqlite_cmd.Connection = sqlite_con

        Dim p As DbParameter

        p = sqlite_cmd.CreateParameter()
        sqlite_cmd.Parameters.Add(p)
        p.Value = Now


        sqlite_cmd.ExecuteNonQuery()


    End Sub

    Private Sub Remont(ByVal Code As String)

        Dim sqlite_cmd As SQLiteCommand
        sqlite_cmd = New SQLiteCommand()
        sqlite_cmd.CommandType = CommandType.Text

        Dim info As String

        info = InputBox("Состав работ", "Информация о ремонте")

        sqlite_cmd.CommandText = "delete from  REP where shCode='" + Code + "'"
        sqlite_cmd.Connection = sqlite_con
        sqlite_cmd.ExecuteNonQuery()


        sqlite_cmd = New SQLiteCommand()
        sqlite_cmd.CommandType = CommandType.Text
        sqlite_cmd.CommandText = "insert into REP(shCode,info,CheckTime)values ('" _
        + Code + "',?,?)"
        sqlite_cmd.Connection = sqlite_con

        Dim p As DbParameter

        p = sqlite_cmd.CreateParameter()
        sqlite_cmd.Parameters.Add(p)
        p.Value = info

        p = sqlite_cmd.CreateParameter()
        sqlite_cmd.Parameters.Add(p)
        p.Value = Now

        sqlite_cmd.ExecuteNonQuery()
    

    End Sub


    Private Sub RENT(ByVal Code As String)

        Dim sqlite_cmd As SQLiteCommand
        sqlite_cmd = New SQLiteCommand()
        sqlite_cmd.CommandType = CommandType.Text

        Dim info As String

        info = InputBox("Арендатор", "Информация об аренде")

        sqlite_cmd.CommandText = "delete from  RENT where shCode='" + Code + "'"
        sqlite_cmd.Connection = sqlite_con
        sqlite_cmd.ExecuteNonQuery()



        sqlite_cmd = New SQLiteCommand()
        sqlite_cmd.CommandType = CommandType.Text
        sqlite_cmd.CommandText = "insert into RENT(shCode,info,CheckTime)values ('" _
        + Code + "',?,?)"
        sqlite_cmd.Connection = sqlite_con

        Dim p As DbParameter
        p = sqlite_cmd.CreateParameter()
        sqlite_cmd.Parameters.Add(p)
        p.Value = info

        p = sqlite_cmd.CreateParameter()
        sqlite_cmd.Parameters.Add(p)
        p.Value = Now


        sqlite_cmd.ExecuteNonQuery()
      

    End Sub

   

    Private Sub cmdOK_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdOK.Click
        SaveCode(txtCode.Text)
        cmdOK.Enabled = False
        cmdEdit.Enabled = False
        cmdRep.Enabled = False
        cmdRent.Enabled = False
        cmdExpl.Enabled = False
        cmdBadCode.Enabled = False
    End Sub

    Private Sub txtCode_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtCode.TextChanged
        Dim ok As Boolean
        Dim s As String
        s = txtCode.Text
        ok = False
        If s.Length = 10 Then
            ok = True
        End If
    

        If ok Then
            If CheckCode(txtCode.Text) Then
                If chkAuto.Checked Then
                    cmdOK_Click(sender, e)
                Else
                    cmdOK.Enabled = True
                    cmdEdit.Enabled = True
                    cmdRep.Enabled = True
                    cmdRent.Enabled = True
                    cmdExpl.Enabled = True
                    cmdBadCode.Enabled = True
                End If
            Else
                MsgBox("Штрихкода нет в базе терминала", MsgBoxStyle.OkOnly, "Штрихкод не опознан")
                txtName.Text = ""
                txtOwner.Text = ""
                txtComp.Text = ""
                txtNum.Text = ""
                txtInfo.Text = ""
                cmdOK.Enabled = False
                cmdEdit.Enabled = False
                cmdRep.Enabled = False
                cmdRent.Enabled = False
                cmdExpl.Enabled = False
                cmdBadCode.Enabled = False
            End If
        Else
            cmdOK.Enabled = False
            cmdEdit.Enabled = False
            cmdRep.Enabled = False
            cmdRent.Enabled = False
            cmdExpl.Enabled = False
            cmdBadCode.Enabled = False
        End If
    End Sub

    Private Sub MenuItem2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        frmEdit.ShowDialog()
    End Sub

    Private Sub cmdClear_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdClear.Click
        cmdOK.Enabled = False
        cmdEdit.Enabled = False
        cmdRep.Enabled = False
        cmdRent.Enabled = False
        cmdExpl.Enabled = False
        cmdBadCode.Enabled = False
        txtCode.Text = ""
        txtName.Text = ""
        txtOwner.Text = ""
        txtComp.Text = ""
        txtNum.Text = ""
        txtInfo.Text = ""
    End Sub

    Private Sub cmdBadCode_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdBadCode.Click
        BadCode(txtCode.Text)
        cmdOK.Enabled = False
        cmdEdit.Enabled = False
        cmdRep.Enabled = False
        cmdRent.Enabled = False
        cmdExpl.Enabled = False
        cmdBadCode.Enabled = False
    End Sub

    Private Sub cmdRep_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdRep.Click
        Remont(txtCode.Text)
    End Sub

    Private Sub cmdRent_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdRent.Click
        RENT(txtCode.Text)
    End Sub

    Private Sub cmdEdit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdEdit.Click
        Dim f As frmEdit
        f = New frmEdit
        f.Code = txtCode.Text
        f.OSName = txtName.Text
        f.OsInfo = txtInfo.Text
        f.OwnerName = txtOwner.Text
        If f.ShowDialog() = Windows.Forms.DialogResult.OK Then
            txtInfo.Text = f.txtInfo.Text
            txtOwner.Text = f.cmbOwner.Text
        End If
        f = Nothing
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdExpl.Click
        ToExpl(txtCode.Text)
    End Sub

   

    Private Sub frmScan_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

    End Sub

    Private Sub MenuItem2_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuItem2.Click
        If cmdOK.Enabled Then
            cmdOK_Click(sender, Nothing)
        End If
    End Sub

    Private Sub cmdUnknown_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdUnknown.Click
        Dim f As frmUnknown
        f = New frmUnknown
        If f.ShowDialog() = Windows.Forms.DialogResult.OK Then
        
        End If

        f = Nothing
    End Sub
End Class