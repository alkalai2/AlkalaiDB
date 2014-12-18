<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class CreateTableForm
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()> _
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Me.btn_createFormCancel = New System.Windows.Forms.Button()
        Me.btn_createFormCreate = New System.Windows.Forms.Button()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.list_createFormAttributes = New System.Windows.Forms.ListBox()
        Me.txt_createFormTableName = New System.Windows.Forms.TextBox()
        Me.group_createFormEdit = New System.Windows.Forms.GroupBox()
        Me.check_createFormNotNull = New System.Windows.Forms.CheckBox()
        Me.check_createFormPK = New System.Windows.Forms.CheckBox()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.combo_createFormDataTypes = New System.Windows.Forms.ComboBox()
        Me.GroupBox2 = New System.Windows.Forms.GroupBox()
        Me.group_createFormLocal = New System.Windows.Forms.GroupBox()
        Me.GroupBox1 = New System.Windows.Forms.GroupBox()
        Me.check_createFormRemote = New System.Windows.Forms.CheckBox()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.Label5 = New System.Windows.Forms.Label()
        Me.Label6 = New System.Windows.Forms.Label()
        Me.Label7 = New System.Windows.Forms.Label()
        Me.Label8 = New System.Windows.Forms.Label()
        Me.txt_localServer = New System.Windows.Forms.TextBox()
        Me.txt_localPort = New System.Windows.Forms.TextBox()
        Me.txt_localDB = New System.Windows.Forms.TextBox()
        Me.txt_localUser = New System.Windows.Forms.TextBox()
        Me.txt_localPassword = New System.Windows.Forms.TextBox()
        Me.group_createFormEdit.SuspendLayout()
        Me.group_createFormLocal.SuspendLayout()
        Me.GroupBox1.SuspendLayout()
        Me.SuspendLayout()
        '
        'btn_createFormCancel
        '
        Me.btn_createFormCancel.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.btn_createFormCancel.Location = New System.Drawing.Point(361, 222)
        Me.btn_createFormCancel.Name = "btn_createFormCancel"
        Me.btn_createFormCancel.Size = New System.Drawing.Size(75, 23)
        Me.btn_createFormCancel.TabIndex = 0
        Me.btn_createFormCancel.Text = "Cancel"
        Me.btn_createFormCancel.UseVisualStyleBackColor = True
        '
        'btn_createFormCreate
        '
        Me.btn_createFormCreate.Location = New System.Drawing.Point(456, 222)
        Me.btn_createFormCreate.Name = "btn_createFormCreate"
        Me.btn_createFormCreate.Size = New System.Drawing.Size(75, 23)
        Me.btn_createFormCreate.TabIndex = 1
        Me.btn_createFormCreate.Text = "Create"
        Me.btn_createFormCreate.UseVisualStyleBackColor = True
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(201, 27)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(71, 13)
        Me.Label1.TabIndex = 2
        Me.Label1.Text = "Table Name :"
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(201, 55)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(51, 13)
        Me.Label2.TabIndex = 3
        Me.Label2.Text = "Attributes"
        '
        'list_createFormAttributes
        '
        Me.list_createFormAttributes.FormattingEnabled = True
        Me.list_createFormAttributes.Location = New System.Drawing.Point(202, 76)
        Me.list_createFormAttributes.Name = "list_createFormAttributes"
        Me.list_createFormAttributes.Size = New System.Drawing.Size(120, 134)
        Me.list_createFormAttributes.TabIndex = 4
        '
        'txt_createFormTableName
        '
        Me.txt_createFormTableName.Location = New System.Drawing.Point(292, 26)
        Me.txt_createFormTableName.Name = "txt_createFormTableName"
        Me.txt_createFormTableName.Size = New System.Drawing.Size(149, 20)
        Me.txt_createFormTableName.TabIndex = 5
        '
        'group_createFormEdit
        '
        Me.group_createFormEdit.Controls.Add(Me.check_createFormNotNull)
        Me.group_createFormEdit.Controls.Add(Me.check_createFormPK)
        Me.group_createFormEdit.Controls.Add(Me.Label3)
        Me.group_createFormEdit.Controls.Add(Me.combo_createFormDataTypes)
        Me.group_createFormEdit.Enabled = False
        Me.group_createFormEdit.Location = New System.Drawing.Point(341, 74)
        Me.group_createFormEdit.Name = "group_createFormEdit"
        Me.group_createFormEdit.Size = New System.Drawing.Size(203, 136)
        Me.group_createFormEdit.TabIndex = 11
        Me.group_createFormEdit.TabStop = False
        Me.group_createFormEdit.Text = "Edit Atributes"
        '
        'check_createFormNotNull
        '
        Me.check_createFormNotNull.AutoSize = True
        Me.check_createFormNotNull.Location = New System.Drawing.Point(9, 94)
        Me.check_createFormNotNull.Name = "check_createFormNotNull"
        Me.check_createFormNotNull.Size = New System.Drawing.Size(64, 17)
        Me.check_createFormNotNull.TabIndex = 15
        Me.check_createFormNotNull.Text = "Not Null"
        Me.check_createFormNotNull.UseVisualStyleBackColor = True
        '
        'check_createFormPK
        '
        Me.check_createFormPK.AutoSize = True
        Me.check_createFormPK.Location = New System.Drawing.Point(9, 61)
        Me.check_createFormPK.Name = "check_createFormPK"
        Me.check_createFormPK.Size = New System.Drawing.Size(81, 17)
        Me.check_createFormPK.TabIndex = 14
        Me.check_createFormPK.Text = "Primary Key"
        Me.check_createFormPK.UseVisualStyleBackColor = True
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(6, 28)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(63, 13)
        Me.Label3.TabIndex = 13
        Me.Label3.Text = "Data Type :"
        '
        'combo_createFormDataTypes
        '
        Me.combo_createFormDataTypes.FormattingEnabled = True
        Me.combo_createFormDataTypes.Location = New System.Drawing.Point(95, 25)
        Me.combo_createFormDataTypes.Name = "combo_createFormDataTypes"
        Me.combo_createFormDataTypes.Size = New System.Drawing.Size(89, 21)
        Me.combo_createFormDataTypes.TabIndex = 12
        '
        'GroupBox2
        '
        Me.GroupBox2.ForeColor = System.Drawing.SystemColors.ControlDark
        Me.GroupBox2.Location = New System.Drawing.Point(189, 4)
        Me.GroupBox2.Name = "GroupBox2"
        Me.GroupBox2.Size = New System.Drawing.Size(355, 207)
        Me.GroupBox2.TabIndex = 12
        Me.GroupBox2.TabStop = False
        Me.GroupBox2.Text = "Configure Table"
        '
        'group_createFormLocal
        '
        Me.group_createFormLocal.Controls.Add(Me.txt_localPassword)
        Me.group_createFormLocal.Controls.Add(Me.txt_localUser)
        Me.group_createFormLocal.Controls.Add(Me.txt_localDB)
        Me.group_createFormLocal.Controls.Add(Me.txt_localPort)
        Me.group_createFormLocal.Controls.Add(Me.txt_localServer)
        Me.group_createFormLocal.Controls.Add(Me.Label8)
        Me.group_createFormLocal.Controls.Add(Me.Label7)
        Me.group_createFormLocal.Controls.Add(Me.Label6)
        Me.group_createFormLocal.Controls.Add(Me.Label5)
        Me.group_createFormLocal.Controls.Add(Me.Label4)
        Me.group_createFormLocal.Enabled = False
        Me.group_createFormLocal.Location = New System.Drawing.Point(4, 55)
        Me.group_createFormLocal.Name = "group_createFormLocal"
        Me.group_createFormLocal.Size = New System.Drawing.Size(187, 156)
        Me.group_createFormLocal.TabIndex = 13
        Me.group_createFormLocal.TabStop = False
        Me.group_createFormLocal.Text = "Local Connection"
        '
        'GroupBox1
        '
        Me.GroupBox1.Controls.Add(Me.check_createFormRemote)
        Me.GroupBox1.ForeColor = System.Drawing.SystemColors.ControlDark
        Me.GroupBox1.Location = New System.Drawing.Point(4, 4)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(191, 207)
        Me.GroupBox1.TabIndex = 15
        Me.GroupBox1.TabStop = False
        Me.GroupBox1.Text = "Configure Host"
        '
        'check_createFormRemote
        '
        Me.check_createFormRemote.AutoSize = True
        Me.check_createFormRemote.Checked = True
        Me.check_createFormRemote.CheckState = System.Windows.Forms.CheckState.Checked
        Me.check_createFormRemote.ForeColor = System.Drawing.SystemColors.ActiveCaptionText
        Me.check_createFormRemote.Location = New System.Drawing.Point(6, 25)
        Me.check_createFormRemote.Name = "check_createFormRemote"
        Me.check_createFormRemote.Size = New System.Drawing.Size(142, 17)
        Me.check_createFormRemote.TabIndex = 0
        Me.check_createFormRemote.Text = "Use Remote Connection"
        Me.check_createFormRemote.UseVisualStyleBackColor = True
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Location = New System.Drawing.Point(12, 31)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(38, 13)
        Me.Label4.TabIndex = 0
        Me.Label4.Text = "Server"
        '
        'Label5
        '
        Me.Label5.AutoSize = True
        Me.Label5.Location = New System.Drawing.Point(12, 54)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(26, 13)
        Me.Label5.TabIndex = 1
        Me.Label5.Text = "Port"
        '
        'Label6
        '
        Me.Label6.AutoSize = True
        Me.Label6.Location = New System.Drawing.Point(12, 78)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(53, 13)
        Me.Label6.TabIndex = 2
        Me.Label6.Text = "Database"
        '
        'Label7
        '
        Me.Label7.AutoSize = True
        Me.Label7.Location = New System.Drawing.Point(12, 100)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(43, 13)
        Me.Label7.TabIndex = 3
        Me.Label7.Text = "User ID"
        '
        'Label8
        '
        Me.Label8.AutoSize = True
        Me.Label8.Location = New System.Drawing.Point(12, 123)
        Me.Label8.Name = "Label8"
        Me.Label8.Size = New System.Drawing.Size(53, 13)
        Me.Label8.TabIndex = 4
        Me.Label8.Text = "Password"
        '
        'txt_localServer
        '
        Me.txt_localServer.Location = New System.Drawing.Point(70, 28)
        Me.txt_localServer.Name = "txt_localServer"
        Me.txt_localServer.Size = New System.Drawing.Size(100, 20)
        Me.txt_localServer.TabIndex = 5
        '
        'txt_localPort
        '
        Me.txt_localPort.Location = New System.Drawing.Point(71, 51)
        Me.txt_localPort.Name = "txt_localPort"
        Me.txt_localPort.Size = New System.Drawing.Size(100, 20)
        Me.txt_localPort.TabIndex = 6
        '
        'txt_localDB
        '
        Me.txt_localDB.Location = New System.Drawing.Point(71, 74)
        Me.txt_localDB.Name = "txt_localDB"
        Me.txt_localDB.Size = New System.Drawing.Size(100, 20)
        Me.txt_localDB.TabIndex = 7
        '
        'txt_localUser
        '
        Me.txt_localUser.Location = New System.Drawing.Point(71, 97)
        Me.txt_localUser.Name = "txt_localUser"
        Me.txt_localUser.Size = New System.Drawing.Size(100, 20)
        Me.txt_localUser.TabIndex = 8
        '
        'txt_localPassword
        '
        Me.txt_localPassword.Location = New System.Drawing.Point(71, 120)
        Me.txt_localPassword.Name = "txt_localPassword"
        Me.txt_localPassword.Size = New System.Drawing.Size(100, 20)
        Me.txt_localPassword.TabIndex = 16
        '
        'CreateTableForm
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.CancelButton = Me.btn_createFormCancel
        Me.ClientSize = New System.Drawing.Size(552, 251)
        Me.Controls.Add(Me.group_createFormLocal)
        Me.Controls.Add(Me.txt_createFormTableName)
        Me.Controls.Add(Me.list_createFormAttributes)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.btn_createFormCreate)
        Me.Controls.Add(Me.btn_createFormCancel)
        Me.Controls.Add(Me.group_createFormEdit)
        Me.Controls.Add(Me.GroupBox2)
        Me.Controls.Add(Me.GroupBox1)
        Me.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Name = "CreateTableForm"
        Me.Text = "Create Table"
        Me.group_createFormEdit.ResumeLayout(False)
        Me.group_createFormEdit.PerformLayout()
        Me.group_createFormLocal.ResumeLayout(False)
        Me.group_createFormLocal.PerformLayout()
        Me.GroupBox1.ResumeLayout(False)
        Me.GroupBox1.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents btn_createFormCancel As System.Windows.Forms.Button
    Friend WithEvents btn_createFormCreate As System.Windows.Forms.Button
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents list_createFormAttributes As System.Windows.Forms.ListBox
    Friend WithEvents txt_createFormTableName As System.Windows.Forms.TextBox
    Friend WithEvents group_createFormEdit As System.Windows.Forms.GroupBox
    Friend WithEvents check_createFormNotNull As System.Windows.Forms.CheckBox
    Friend WithEvents check_createFormPK As System.Windows.Forms.CheckBox
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents combo_createFormDataTypes As System.Windows.Forms.ComboBox
    Friend WithEvents GroupBox2 As System.Windows.Forms.GroupBox
    Friend WithEvents group_createFormLocal As System.Windows.Forms.GroupBox
    Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
    Friend WithEvents check_createFormRemote As System.Windows.Forms.CheckBox
    Friend WithEvents txt_localPassword As System.Windows.Forms.TextBox
    Friend WithEvents txt_localUser As System.Windows.Forms.TextBox
    Friend WithEvents txt_localDB As System.Windows.Forms.TextBox
    Friend WithEvents txt_localPort As System.Windows.Forms.TextBox
    Friend WithEvents txt_localServer As System.Windows.Forms.TextBox
    Friend WithEvents Label8 As System.Windows.Forms.Label
    Friend WithEvents Label7 As System.Windows.Forms.Label
    Friend WithEvents Label6 As System.Windows.Forms.Label
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents Label4 As System.Windows.Forms.Label
End Class
