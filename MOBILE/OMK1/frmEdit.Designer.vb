<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Public Class frmEdit
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
        Me.cmbOwner = New System.Windows.Forms.ComboBox
        Me.lblOwner = New System.Windows.Forms.Label
        Me.lblName = New System.Windows.Forms.Label
        Me.txtName = New System.Windows.Forms.TextBox
        Me.lblCode = New System.Windows.Forms.Label
        Me.txtCode = New System.Windows.Forms.TextBox
        Me.cmdCancel = New System.Windows.Forms.Button
        Me.cmdSave = New System.Windows.Forms.Button
        Me.Label1 = New System.Windows.Forms.Label
        Me.txtInfo = New System.Windows.Forms.TextBox
        Me.chkChangeCode = New System.Windows.Forms.CheckBox
        Me.SuspendLayout()
        '
        'cmbOwner
        '
        Me.cmbOwner.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDown
        Me.cmbOwner.Location = New System.Drawing.Point(5, 61)
        Me.cmbOwner.Name = "cmbOwner"
        Me.cmbOwner.Size = New System.Drawing.Size(222, 22)
        Me.cmbOwner.TabIndex = 2
        '
        'lblOwner
        '
        Me.lblOwner.Location = New System.Drawing.Point(5, 43)
        Me.lblOwner.Name = "lblOwner"
        Me.lblOwner.Size = New System.Drawing.Size(62, 15)
        Me.lblOwner.Text = "Владелец"
        '
        'lblName
        '
        Me.lblName.Location = New System.Drawing.Point(5, 86)
        Me.lblName.Name = "lblName"
        Me.lblName.Size = New System.Drawing.Size(60, 21)
        Me.lblName.Text = "Название"
        '
        'txtName
        '
        Me.txtName.Location = New System.Drawing.Point(5, 100)
        Me.txtName.Name = "txtName"
        Me.txtName.ReadOnly = True
        Me.txtName.Size = New System.Drawing.Size(222, 21)
        Me.txtName.TabIndex = 3
        '
        'lblCode
        '
        Me.lblCode.Location = New System.Drawing.Point(5, 5)
        Me.lblCode.Name = "lblCode"
        Me.lblCode.Size = New System.Drawing.Size(58, 23)
        Me.lblCode.Text = "Код"
        '
        'txtCode
        '
        Me.txtCode.Location = New System.Drawing.Point(5, 19)
        Me.txtCode.Name = "txtCode"
        Me.txtCode.ReadOnly = True
        Me.txtCode.Size = New System.Drawing.Size(222, 21)
        Me.txtCode.TabIndex = 1
        '
        'cmdCancel
        '
        Me.cmdCancel.Location = New System.Drawing.Point(155, 233)
        Me.cmdCancel.Name = "cmdCancel"
        Me.cmdCancel.Size = New System.Drawing.Size(82, 32)
        Me.cmdCancel.TabIndex = 6
        Me.cmdCancel.Text = "Отмена"
        '
        'cmdSave
        '
        Me.cmdSave.Location = New System.Drawing.Point(5, 233)
        Me.cmdSave.Name = "cmdSave"
        Me.cmdSave.Size = New System.Drawing.Size(87, 32)
        Me.cmdSave.TabIndex = 5
        Me.cmdSave.Text = "Записать"
        '
        'Label1
        '
        Me.Label1.Location = New System.Drawing.Point(5, 124)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(73, 17)
        Me.Label1.Text = "Описание"
        '
        'txtInfo
        '
        Me.txtInfo.Location = New System.Drawing.Point(5, 144)
        Me.txtInfo.Multiline = True
        Me.txtInfo.Name = "txtInfo"
        Me.txtInfo.ScrollBars = System.Windows.Forms.ScrollBars.Both
        Me.txtInfo.Size = New System.Drawing.Size(222, 48)
        Me.txtInfo.TabIndex = 4
        '
        'chkChangeCode
        '
        Me.chkChangeCode.Location = New System.Drawing.Point(3, 198)
        Me.chkChangeCode.Name = "chkChangeCode"
        Me.chkChangeCode.Size = New System.Drawing.Size(224, 27)
        Me.chkChangeCode.TabIndex = 10
        Me.chkChangeCode.Text = "Перепечатать штрихкод"
        '
        'frmEdit
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(96.0!, 96.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Dpi
        Me.AutoScroll = True
        Me.ClientSize = New System.Drawing.Size(240, 268)
        Me.Controls.Add(Me.chkChangeCode)
        Me.Controls.Add(Me.txtInfo)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.cmdSave)
        Me.Controls.Add(Me.cmdCancel)
        Me.Controls.Add(Me.txtCode)
        Me.Controls.Add(Me.lblCode)
        Me.Controls.Add(Me.txtName)
        Me.Controls.Add(Me.lblName)
        Me.Controls.Add(Me.lblOwner)
        Me.Controls.Add(Me.cmbOwner)
        Me.Menu = Me.mainMenu1
        Me.Name = "frmEdit"
        Me.Text = "Редактирование"
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents cmbOwner As System.Windows.Forms.ComboBox
    Friend WithEvents lblOwner As System.Windows.Forms.Label
    Friend WithEvents lblName As System.Windows.Forms.Label
    Friend WithEvents txtName As System.Windows.Forms.TextBox
    Friend WithEvents lblCode As System.Windows.Forms.Label
    Friend WithEvents txtCode As System.Windows.Forms.TextBox
    Friend WithEvents cmdCancel As System.Windows.Forms.Button
    Friend WithEvents cmdSave As System.Windows.Forms.Button
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents txtInfo As System.Windows.Forms.TextBox
    Friend WithEvents chkChangeCode As System.Windows.Forms.CheckBox
End Class
