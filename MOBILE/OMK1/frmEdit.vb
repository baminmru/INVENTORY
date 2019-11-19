Imports System.Data.Common
Imports System.Data
Imports System.Data.SQLite


Public Class frmEdit

    Public OwnerName As String
    Public Code As String
    Public OSName As String
    Public OsInfo As String

    Private Sub cmdCancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdCancel.Click
        Me.DialogResult = Windows.Forms.DialogResult.Cancel
        Me.Close()
    End Sub

    Private Sub frmEdit_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

        Dim sqlite_da As SQLiteDataAdapter = New SQLiteDataAdapter()
        Dim dataSet As DataSet = New DataSet()
        sqlite_da.SelectCommand = New SQLiteCommand("select name from owners", sqlite_con)
        sqlite_da.Fill(dataSet)
        Dim dt As DataTable

        dt = dataSet.Tables(0)
        cmbOwner.DisplayMember = "name"
        cmbOwner.ValueMember = "name"
        cmbOwner.DataSource = dt
        cmbOwner.Text = OwnerName
        txtCode.Text = Code
        txtInfo.Text = OsInfo
        txtName.Text = OSName
    End Sub

    Private Sub cmdSave_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdSave.Click

        Dim sqlite_cmd As SQLiteCommand
        sqlite_cmd = New SQLiteCommand()
        sqlite_cmd.CommandType = CommandType.Text

    
        sqlite_cmd.CommandText = "update INV set Changed =1, INVOS_PLACE_theowner=?, INVOS_INFO_info=?  where VisibleCode='" _
        + Code + "'"
        sqlite_cmd.Connection = sqlite_con



        Dim p As DbParameter

        p = sqlite_cmd.CreateParameter()
        sqlite_cmd.Parameters.Add(p)
        p.Value = cmbOwner.Text


        p = sqlite_cmd.CreateParameter()
        sqlite_cmd.Parameters.Add(p)
        p.Value = txtInfo.Text

        sqlite_cmd.ExecuteNonQuery()

        ' пометить для печати нового кода
        If chkChangeCode.Checked Then
            ChCode(True)
        Else
            ChCode(False)
        End If

        Me.DialogResult = Windows.Forms.DialogResult.OK
        Me.Close()
    End Sub


    Private Sub ChCode(ByVal save As Boolean)
        Dim sqlite_cmd As SQLiteCommand
        sqlite_cmd = New SQLiteCommand()
        sqlite_cmd.CommandType = CommandType.Text

        sqlite_cmd.CommandText = "delete from  PRT where shCode='" + Code + "'"
        sqlite_cmd.Connection = sqlite_con
        sqlite_cmd.ExecuteNonQuery()


        If save Then
            sqlite_cmd = New SQLiteCommand()
            sqlite_cmd.CommandType = CommandType.Text
            sqlite_cmd.CommandText = "insert into PRT(shCode,CheckTime)values ('" _
            + Code + "',?)"
            sqlite_cmd.Connection = sqlite_con

            Dim p As DbParameter
            p = sqlite_cmd.CreateParameter()
            sqlite_cmd.Parameters.Add(p)
            p.Value = Now

            sqlite_cmd.ExecuteNonQuery()
        End If

    End Sub
End Class