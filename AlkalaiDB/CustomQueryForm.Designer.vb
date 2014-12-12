<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class CustomQueryForm
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
        Me.label_customQuery = New System.Windows.Forms.Label()
        Me.txt_customQuery = New System.Windows.Forms.TextBox()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.btn_createCustomQuery = New System.Windows.Forms.Button()
        Me.check_customNewTable = New System.Windows.Forms.CheckBox()
        Me.label_customName = New System.Windows.Forms.Label()
        Me.txt_customName = New System.Windows.Forms.TextBox()
        Me.SuspendLayout()
        '
        'label_customQuery
        '
        Me.label_customQuery.AutoSize = True
        Me.label_customQuery.Location = New System.Drawing.Point(12, 58)
        Me.label_customQuery.Name = "label_customQuery"
        Me.label_customQuery.Size = New System.Drawing.Size(47, 13)
        Me.label_customQuery.TabIndex = 0
        Me.label_customQuery.Text = "Query :  "
        '
        'txt_customQuery
        '
        Me.txt_customQuery.Location = New System.Drawing.Point(65, 55)
        Me.txt_customQuery.Name = "txt_customQuery"
        Me.txt_customQuery.Size = New System.Drawing.Size(333, 20)
        Me.txt_customQuery.TabIndex = 1
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.ForeColor = System.Drawing.SystemColors.ControlDarkDark
        Me.Label1.ImageAlign = System.Drawing.ContentAlignment.BottomCenter
        Me.Label1.Location = New System.Drawing.Point(32, 86)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(408, 13)
        Me.Label1.TabIndex = 2
        Me.Label1.Text = "( SELECT Values FROM Table1 , Table2 WHERE Table1.Values = Table2.Values ; )"
        '
        'btn_createCustomQuery
        '
        Me.btn_createCustomQuery.Location = New System.Drawing.Point(414, 53)
        Me.btn_createCustomQuery.Name = "btn_createCustomQuery"
        Me.btn_createCustomQuery.Size = New System.Drawing.Size(62, 23)
        Me.btn_createCustomQuery.TabIndex = 3
        Me.btn_createCustomQuery.Text = "Create"
        Me.btn_createCustomQuery.UseVisualStyleBackColor = True
        '
        'check_customNewTable
        '
        Me.check_customNewTable.AutoSize = True
        Me.check_customNewTable.Location = New System.Drawing.Point(61, 12)
        Me.check_customNewTable.Name = "check_customNewTable"
        Me.check_customNewTable.Size = New System.Drawing.Size(112, 17)
        Me.check_customNewTable.TabIndex = 4
        Me.check_customNewTable.Text = "Create New Table"
        Me.check_customNewTable.UseVisualStyleBackColor = True
        '
        'label_customName
        '
        Me.label_customName.AutoSize = True
        Me.label_customName.Location = New System.Drawing.Point(224, 13)
        Me.label_customName.Name = "label_customName"
        Me.label_customName.Size = New System.Drawing.Size(41, 13)
        Me.label_customName.TabIndex = 5
        Me.label_customName.Text = "Name: "
        Me.label_customName.Visible = False
        '
        'txt_customName
        '
        Me.txt_customName.Location = New System.Drawing.Point(285, 10)
        Me.txt_customName.Name = "txt_customName"
        Me.txt_customName.Size = New System.Drawing.Size(113, 20)
        Me.txt_customName.TabIndex = 6
        Me.txt_customName.Visible = False
        '
        'CustomQueryForm
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(499, 136)
        Me.Controls.Add(Me.txt_customName)
        Me.Controls.Add(Me.label_customName)
        Me.Controls.Add(Me.check_customNewTable)
        Me.Controls.Add(Me.btn_createCustomQuery)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.txt_customQuery)
        Me.Controls.Add(Me.label_customQuery)
        Me.Name = "CustomQueryForm"
        Me.Text = "Create a Custom Query"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents label_customQuery As System.Windows.Forms.Label
    Friend WithEvents txt_customQuery As System.Windows.Forms.TextBox
    Public WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents btn_createCustomQuery As System.Windows.Forms.Button
    Friend WithEvents check_customNewTable As System.Windows.Forms.CheckBox
    Friend WithEvents label_customName As System.Windows.Forms.Label
    Friend WithEvents txt_customName As System.Windows.Forms.TextBox
End Class
