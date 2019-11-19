
Imports System.Data.Common
Imports System.Data
Imports System.Data.SQLite

Public Class frmUnknown




    Private Sub cmdCancel_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdCancel.Click
        Me.DialogResult = Windows.Forms.DialogResult.Cancel
        Me.Close()
    End Sub

    Private Sub cmdSave_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdSave.Click
        Dim sqlite_cmd As SQLiteCommand

        sqlite_cmd = New SQLiteCommand()
        sqlite_cmd.CommandType = CommandType.Text


        sqlite_cmd.CommandText = "insert into U (PLACE,NAME,INFO) values(?,?,?)"
        sqlite_cmd.Connection = sqlite_con



        Dim p As DbParameter

        p = sqlite_cmd.CreateParameter()
        sqlite_cmd.Parameters.Add(p)
        p.Value = txtPlace.Text

        p = sqlite_cmd.CreateParameter()
        sqlite_cmd.Parameters.Add(p)
        p.Value = txtName.Text

        p = sqlite_cmd.CreateParameter()
        sqlite_cmd.Parameters.Add(p)
        p.Value = txtInfo.Text


        sqlite_cmd.ExecuteNonQuery()

        Me.DialogResult = Windows.Forms.DialogResult.OK
        Me.Close()
    End Sub

    Private Sub frmUnknown_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

        Dim sqlite_da As SQLiteDataAdapter = New SQLiteDataAdapter()
        Dim dataSet As DataSet = New DataSet()
        sqlite_da.SelectCommand = New SQLiteCommand("select count(*) cnt from INV", sqlite_con)
        sqlite_da.Fill(dataSet)
        Dim dt As DataTable
        Dim cnt As Integer
        cnt = 0
        If dataSet.Tables.Count > 0 Then
            dt = dataSet.Tables(0)
            If dt.Rows.Count > 0 Then
                cnt = dt.Rows(0)("cnt")
            End If
        End If
        If cnt = 0 Then
            MsgBox("Сначала в терминал надо загрузить данные инвентаризации", MsgBoxStyle.OkOnly, "Регистрация невозможна")
            Me.Close()
        End If



    End Sub
End Class