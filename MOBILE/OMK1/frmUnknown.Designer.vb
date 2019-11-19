<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Public Class frmUnknown
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()> _
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        If disposing AndAlso components IsNot Nothing Then
            components.Dispose()
        End If
        MyBase.Dispose(disposing)
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer
    Private mainMenu1 As System.Windows.Forms.MainMenu

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Me.mainMenu1 = New System.Windows.Forms.MainMenu
        Me.txtInfo = New System.Windows.Forms.TextBox
        Me.Label1 = New System.Windows.Forms.Label
        Me.cmdSave = New System.Windows.Forms.Button
        Me.cmdCancel = New System.Windows.Forms.Button
        Me.txtPlace = New System.Windows.Forms.TextBox
        Me.lblCode = New System.Windows.Forms.Label
        Me.txtName = New System.Windows.Forms.TextBox
        Me.lblName = New System.Windows.Forms.Label
        Me.SuspendLayout()
        '
        'txtInfo
        '
        Me.txtInfo.Location = New System.Drawing.Point(4, 99)
        Me.txtInfo.Multiline = True
        Me.txtInfo.Name = "txtInfo"
        Me.txtInfo.ScrollBars = System.Windows.Forms.ScrollBars.Both
        Me.txtInfo.Size = New System.Drawing.Size(222, 114)
        Me.txtInfo.TabIndex = 11
        '
        'Label1
        '
        Me.Label1.Location = New System.Drawing.Point(4, 79)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(73, 17)
        Me.Label1.Text = "Описание"
        '
        'cmdSave
        '
        Me.cmdSave.Location = New System.Drawing.Point(4, 232)
        Me.cmdSave.Name = "cmdSave"
        Me.cmdSave.Size = New System.Drawing.Size(87, 32)
        Me.cmdSave.TabIndex = 12
        Me.cmdSave.Text = "Записать"
        '
        'cmdCancel
        '
        Me.cmdCancel.Location = New System.Drawing.Point(154, 232)
        Me.cmdCancel.Name = "cmdCancel"
        Me.cmdCancel.Size = New System.Drawing.Size(82, 32)
        Me.cmdCancel.TabIndex = 14
        Me.cmdCancel.Text = "Отмена"
        '
        'txtPlace
        '
        Me.txtPlace.Location = New System.Drawing.Point(4, 18)
        Me.txtPlace.Name = "txtPlace"
        Me.txtPlace.Size = New System.Drawing.Size(222, 21)
        Me.txtPlace.TabIndex = 9
        '
        'lblCode
        '
        Me.lblCode.Location = New System.Drawing.Point(4, 4)
        Me.lblCode.Name = "lblCode"
        Me.lblCode.Size = New System.Drawing.Size(156, 19)
        Me.lblCode.Text = "Помещение"
        '
        'txtName
        '
        Me.txtName.Location = New System.Drawing.Point(4, 55)
        Me.txtName.Name = "txtName"
        Me.txtName.Size = New System.Drawing.Size(222, 21)
        Me.txtName.TabIndex = 10
        '
        'lblName
        '
        Me.lblName.Location = New System.Drawing.Point(4, 42)
        Me.lblName.Name = "lblName"
        Me.lblName.Size = New System.Drawing.Size(104, 21)
        Me.lblName.Text = "Название"
        '
        'frmUnknown
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(96.0!, 96.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Dpi
        Me.AutoScroll = True
        Me.ClientSize = New System.Drawing.Size(240, 268)
        Me.Controls.Add(Me.txtInfo)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.cmdSave)
        Me.Controls.Add(Me.cmdCancel)
        Me.Controls.Add(Me.txtPlace)
        Me.Controls.Add(Me.lblCode)
        Me.Controls.Add(Me.txtName)
        Me.Controls.Add(Me.lblName)
        Me.Menu = Me.mainMenu1
        Me.Name = "frmUnknown"
        Me.Text = "Без штрихкода"
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents txtInfo As System.Windows.Forms.TextBox
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents cmdSave As System.Windows.Forms.Button
    Friend WithEvents cmdCancel As System.Windows.Forms.Button
    Friend WithEvents txtPlace As System.Windows.Forms.TextBox
    Friend WithEvents lblCode As System.Windows.Forms.Label
    Friend WithEvents txtName As System.Windows.Forms.TextBox
    Friend WithEvents lblName As System.Windows.Forms.Label
End Class
