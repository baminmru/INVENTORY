<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Public Class frmScan
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
        Me.MenuItem1 = New System.Windows.Forms.MenuItem
        Me.MenuItem2 = New System.Windows.Forms.MenuItem
        Me.Label1 = New System.Windows.Forms.Label
        Me.txtCode = New System.Windows.Forms.TextBox
        Me.Label2 = New System.Windows.Forms.Label
        Me.Label3 = New System.Windows.Forms.Label
        Me.txtName = New System.Windows.Forms.TextBox
        Me.Barcode1 = New Barcode.Barcode
        Me.txtComp = New System.Windows.Forms.TextBox
        Me.cmdOK = New System.Windows.Forms.Button
        Me.Label4 = New System.Windows.Forms.Label
        Me.cmbStatus = New System.Windows.Forms.ComboBox
        Me.Label5 = New System.Windows.Forms.Label
        Me.txtOwner = New System.Windows.Forms.TextBox
        Me.Label6 = New System.Windows.Forms.Label
        Me.txtNum = New System.Windows.Forms.TextBox
        Me.chkAuto = New System.Windows.Forms.CheckBox
        Me.Label7 = New System.Windows.Forms.Label
        Me.cmdBadCode = New System.Windows.Forms.Button
        Me.cmdRep = New System.Windows.Forms.Button
        Me.cmdRent = New System.Windows.Forms.Button
        Me.cmdEdit = New System.Windows.Forms.Button
        Me.cmdClear = New System.Windows.Forms.Button
        Me.Label8 = New System.Windows.Forms.Label
        Me.txtInfo = New System.Windows.Forms.TextBox
        Me.cmdExpl = New System.Windows.Forms.Button
        Me.cmdUnknown = New System.Windows.Forms.Button
        Me.SuspendLayout()
        '
        'mainMenu1
        '
        Me.mainMenu1.MenuItems.Add(Me.MenuItem1)
        Me.mainMenu1.MenuItems.Add(Me.MenuItem2)
        '
        'MenuItem1
        '
        Me.MenuItem1.Text = "Закрыть"
        '
        'MenuItem2
        '
        Me.MenuItem2.Text = "Записать"
        '
        'Label1
        '
        Me.Label1.Location = New System.Drawing.Point(3, 0)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(67, 16)
        Me.Label1.Text = "Код"
        '
        'txtCode
        '
        Me.txtCode.Location = New System.Drawing.Point(3, 16)
        Me.txtCode.Name = "txtCode"
        Me.txtCode.Size = New System.Drawing.Size(114, 21)
        Me.txtCode.TabIndex = 1
        '
        'Label2
        '
        Me.Label2.Location = New System.Drawing.Point(3, 81)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(102, 17)
        Me.Label2.Text = "№ комплекта"
        '
        'Label3
        '
        Me.Label3.Location = New System.Drawing.Point(9, 120)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(61, 21)
        Me.Label3.Text = "Название"
        '
        'txtName
        '
        Me.txtName.Location = New System.Drawing.Point(81, 123)
        Me.txtName.Name = "txtName"
        Me.txtName.ReadOnly = True
        Me.txtName.Size = New System.Drawing.Size(155, 21)
        Me.txtName.TabIndex = 8
        '
        'Barcode1
        '
        '
        'txtComp
        '
        Me.txtComp.Location = New System.Drawing.Point(3, 96)
        Me.txtComp.Name = "txtComp"
        Me.txtComp.ReadOnly = True
        Me.txtComp.Size = New System.Drawing.Size(72, 21)
        Me.txtComp.TabIndex = 6
        '
        'cmdOK
        '
        Me.cmdOK.Enabled = False
        Me.cmdOK.Location = New System.Drawing.Point(189, 16)
        Me.cmdOK.Name = "cmdOK"
        Me.cmdOK.Size = New System.Drawing.Size(47, 20)
        Me.cmdOK.TabIndex = 3
        Me.cmdOK.Text = "OK"
        '
        'Label4
        '
        Me.Label4.Location = New System.Drawing.Point(3, 40)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(149, 17)
        Me.Label4.Text = "Состояние"
        '
        'cmbStatus
        '
        Me.cmbStatus.Items.Add("В наличии")
        Me.cmbStatus.Items.Add("Требует ремонта")
        Me.cmbStatus.Items.Add("Сломан")
        Me.cmbStatus.Location = New System.Drawing.Point(3, 56)
        Me.cmbStatus.Name = "cmbStatus"
        Me.cmbStatus.Size = New System.Drawing.Size(231, 22)
        Me.cmbStatus.TabIndex = 4
        '
        'Label5
        '
        Me.Label5.Location = New System.Drawing.Point(9, 150)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(66, 21)
        Me.Label5.Text = "Отв. лицо"
        '
        'txtOwner
        '
        Me.txtOwner.Location = New System.Drawing.Point(81, 150)
        Me.txtOwner.Name = "txtOwner"
        Me.txtOwner.ReadOnly = True
        Me.txtOwner.Size = New System.Drawing.Size(156, 21)
        Me.txtOwner.TabIndex = 9
        '
        'Label6
        '
        Me.Label6.Location = New System.Drawing.Point(123, 81)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(113, 17)
        Me.Label6.Text = "Инв. №/ № Уч. к."
        '
        'txtNum
        '
        Me.txtNum.Location = New System.Drawing.Point(81, 96)
        Me.txtNum.Name = "txtNum"
        Me.txtNum.ReadOnly = True
        Me.txtNum.Size = New System.Drawing.Size(155, 21)
        Me.txtNum.TabIndex = 7
        '
        'chkAuto
        '
        Me.chkAuto.Location = New System.Drawing.Point(153, 18)
        Me.chkAuto.Name = "chkAuto"
        Me.chkAuto.Size = New System.Drawing.Size(30, 18)
        Me.chkAuto.TabIndex = 5
        '
        'Label7
        '
        Me.Label7.Location = New System.Drawing.Point(153, -2)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(38, 17)
        Me.Label7.Text = "Авто"
        '
        'cmdBadCode
        '
        Me.cmdBadCode.Enabled = False
        Me.cmdBadCode.Font = New System.Drawing.Font("Tahoma", 9.0!, System.Drawing.FontStyle.Regular)
        Me.cmdBadCode.Location = New System.Drawing.Point(3, 225)
        Me.cmdBadCode.Name = "cmdBadCode"
        Me.cmdBadCode.Size = New System.Drawing.Size(111, 17)
        Me.cmdBadCode.TabIndex = 13
        Me.cmdBadCode.Text = "Неверный код"
        '
        'cmdRep
        '
        Me.cmdRep.Enabled = False
        Me.cmdRep.Font = New System.Drawing.Font("Tahoma", 9.0!, System.Drawing.FontStyle.Regular)
        Me.cmdRep.Location = New System.Drawing.Point(125, 225)
        Me.cmdRep.Name = "cmdRep"
        Me.cmdRep.Size = New System.Drawing.Size(109, 17)
        Me.cmdRep.TabIndex = 14
        Me.cmdRep.Text = "В ремонт"
        '
        'cmdRent
        '
        Me.cmdRent.Enabled = False
        Me.cmdRent.Font = New System.Drawing.Font("Tahoma", 9.0!, System.Drawing.FontStyle.Regular)
        Me.cmdRent.Location = New System.Drawing.Point(125, 246)
        Me.cmdRent.Name = "cmdRent"
        Me.cmdRent.Size = New System.Drawing.Size(109, 18)
        Me.cmdRent.TabIndex = 16
        Me.cmdRent.Text = "В аренду"
        '
        'cmdEdit
        '
        Me.cmdEdit.Enabled = False
        Me.cmdEdit.Font = New System.Drawing.Font("Tahoma", 9.0!, System.Drawing.FontStyle.Regular)
        Me.cmdEdit.Location = New System.Drawing.Point(3, 203)
        Me.cmdEdit.Name = "cmdEdit"
        Me.cmdEdit.Size = New System.Drawing.Size(111, 18)
        Me.cmdEdit.TabIndex = 11
        Me.cmdEdit.Text = "Редактировать"
        '
        'cmdClear
        '
        Me.cmdClear.Font = New System.Drawing.Font("Tahoma", 9.0!, System.Drawing.FontStyle.Regular)
        Me.cmdClear.Location = New System.Drawing.Point(123, 16)
        Me.cmdClear.Name = "cmdClear"
        Me.cmdClear.Size = New System.Drawing.Size(24, 19)
        Me.cmdClear.TabIndex = 2
        Me.cmdClear.Text = "X"
        '
        'Label8
        '
        Me.Label8.Location = New System.Drawing.Point(9, 177)
        Me.Label8.Name = "Label8"
        Me.Label8.Size = New System.Drawing.Size(61, 17)
        Me.Label8.Text = "Описание"
        '
        'txtInfo
        '
        Me.txtInfo.Location = New System.Drawing.Point(81, 177)
        Me.txtInfo.Name = "txtInfo"
        Me.txtInfo.ReadOnly = True
        Me.txtInfo.Size = New System.Drawing.Size(154, 21)
        Me.txtInfo.TabIndex = 10
        '
        'cmdExpl
        '
        Me.cmdExpl.Enabled = False
        Me.cmdExpl.Font = New System.Drawing.Font("Tahoma", 9.0!, System.Drawing.FontStyle.Regular)
        Me.cmdExpl.Location = New System.Drawing.Point(125, 204)
        Me.cmdExpl.Name = "cmdExpl"
        Me.cmdExpl.Size = New System.Drawing.Size(109, 17)
        Me.cmdExpl.TabIndex = 12
        Me.cmdExpl.Text = "В эксплуатацию"
        '
        'cmdUnknown
        '
        Me.cmdUnknown.Font = New System.Drawing.Font("Tahoma", 9.0!, System.Drawing.FontStyle.Regular)
        Me.cmdUnknown.Location = New System.Drawing.Point(3, 246)
        Me.cmdUnknown.Name = "cmdUnknown"
        Me.cmdUnknown.Size = New System.Drawing.Size(111, 18)
        Me.cmdUnknown.TabIndex = 15
        Me.cmdUnknown.Text = "Нет штрихкода !"
        '
        'frmScan
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(96.0!, 96.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Dpi
        Me.AutoScroll = True
        Me.ClientSize = New System.Drawing.Size(240, 268)
        Me.Controls.Add(Me.cmdUnknown)
        Me.Controls.Add(Me.cmdExpl)
        Me.Controls.Add(Me.txtInfo)
        Me.Controls.Add(Me.Label8)
        Me.Controls.Add(Me.cmdClear)
        Me.Controls.Add(Me.cmdEdit)
        Me.Controls.Add(Me.cmdRent)
        Me.Controls.Add(Me.cmdRep)
        Me.Controls.Add(Me.cmdBadCode)
        Me.Controls.Add(Me.Label7)
        Me.Controls.Add(Me.chkAuto)
        Me.Controls.Add(Me.txtNum)
        Me.Controls.Add(Me.Label6)
        Me.Controls.Add(Me.txtOwner)
        Me.Controls.Add(Me.Label5)
        Me.Controls.Add(Me.cmbStatus)
        Me.Controls.Add(Me.Label4)
        Me.Controls.Add(Me.cmdOK)
        Me.Controls.Add(Me.txtComp)
        Me.Controls.Add(Me.txtName)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.txtCode)
        Me.Controls.Add(Me.Label1)
        Me.KeyPreview = True
        Me.Menu = Me.mainMenu1
        Me.Name = "frmScan"
        Me.Text = "Сбор данных"
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents txtCode As System.Windows.Forms.TextBox
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents txtName As System.Windows.Forms.TextBox
    Friend WithEvents Barcode1 As Barcode.Barcode
    Friend WithEvents MenuItem1 As System.Windows.Forms.MenuItem
    Friend WithEvents txtComp As System.Windows.Forms.TextBox
    Friend WithEvents cmdOK As System.Windows.Forms.Button
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents cmbStatus As System.Windows.Forms.ComboBox
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents txtOwner As System.Windows.Forms.TextBox
    Friend WithEvents Label6 As System.Windows.Forms.Label
    Friend WithEvents txtNum As System.Windows.Forms.TextBox
    Friend WithEvents chkAuto As System.Windows.Forms.CheckBox
    Friend WithEvents Label7 As System.Windows.Forms.Label
    Friend WithEvents cmdBadCode As System.Windows.Forms.Button
    Friend WithEvents cmdRep As System.Windows.Forms.Button
    Friend WithEvents cmdRent As System.Windows.Forms.Button
    Friend WithEvents cmdEdit As System.Windows.Forms.Button
    Friend WithEvents cmdClear As System.Windows.Forms.Button
    Friend WithEvents Label8 As System.Windows.Forms.Label
    Friend WithEvents txtInfo As System.Windows.Forms.TextBox
    Friend WithEvents cmdExpl As System.Windows.Forms.Button
    Friend WithEvents MenuItem2 As System.Windows.Forms.MenuItem
    Friend WithEvents cmdUnknown As System.Windows.Forms.Button
End Class
