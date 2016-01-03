<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class SlabThick
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
        Me.ListBox_Slaplist = New System.Windows.Forms.ListBox()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.TextBox_thickness = New System.Windows.Forms.TextBox()
        Me.Button_addslabthk = New System.Windows.Forms.Button()
        Me.Button_RemoveSlabTHK = New System.Windows.Forms.Button()
        Me.Button_close = New System.Windows.Forms.Button()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.TextBox_addslabthk = New System.Windows.Forms.TextBox()
        Me.Button_modify = New System.Windows.Forms.Button()
        Me.Button_rename = New System.Windows.Forms.Button()
        Me.SuspendLayout()
        '
        'ListBox_Slaplist
        '
        Me.ListBox_Slaplist.FormattingEnabled = True
        Me.ListBox_Slaplist.Location = New System.Drawing.Point(15, 51)
        Me.ListBox_Slaplist.MultiColumn = True
        Me.ListBox_Slaplist.Name = "ListBox_Slaplist"
        Me.ListBox_Slaplist.Size = New System.Drawing.Size(109, 134)
        Me.ListBox_Slaplist.TabIndex = 0
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(12, 9)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(31, 13)
        Me.Label1.TabIndex = 1
        Me.Label1.Text = "Slab:"
        '
        'TextBox_thickness
        '
        Me.TextBox_thickness.Location = New System.Drawing.Point(148, 25)
        Me.TextBox_thickness.Name = "TextBox_thickness"
        Me.TextBox_thickness.Size = New System.Drawing.Size(65, 20)
        Me.TextBox_thickness.TabIndex = 2
        '
        'Button_addslabthk
        '
        Me.Button_addslabthk.Location = New System.Drawing.Point(148, 51)
        Me.Button_addslabthk.Name = "Button_addslabthk"
        Me.Button_addslabthk.Size = New System.Drawing.Size(65, 28)
        Me.Button_addslabthk.TabIndex = 3
        Me.Button_addslabthk.Text = "Add"
        Me.Button_addslabthk.UseVisualStyleBackColor = True
        '
        'Button_RemoveSlabTHK
        '
        Me.Button_RemoveSlabTHK.Location = New System.Drawing.Point(148, 153)
        Me.Button_RemoveSlabTHK.Name = "Button_RemoveSlabTHK"
        Me.Button_RemoveSlabTHK.Size = New System.Drawing.Size(65, 28)
        Me.Button_RemoveSlabTHK.TabIndex = 3
        Me.Button_RemoveSlabTHK.Text = "Remove"
        Me.Button_RemoveSlabTHK.UseVisualStyleBackColor = True
        '
        'Button_close
        '
        Me.Button_close.Location = New System.Drawing.Point(178, 199)
        Me.Button_close.Name = "Button_close"
        Me.Button_close.Size = New System.Drawing.Size(65, 28)
        Me.Button_close.TabIndex = 3
        Me.Button_close.Text = "Close"
        Me.Button_close.UseVisualStyleBackColor = True
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(220, 32)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(23, 13)
        Me.Label2.TabIndex = 4
        Me.Label2.Text = "mm"
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(145, 9)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(59, 13)
        Me.Label3.TabIndex = 5
        Me.Label3.Text = "Thickness:"
        '
        'TextBox_addslabthk
        '
        Me.TextBox_addslabthk.Location = New System.Drawing.Point(15, 25)
        Me.TextBox_addslabthk.Name = "TextBox_addslabthk"
        Me.TextBox_addslabthk.Size = New System.Drawing.Size(109, 20)
        Me.TextBox_addslabthk.TabIndex = 6
        '
        'Button_modify
        '
        Me.Button_modify.Location = New System.Drawing.Point(148, 85)
        Me.Button_modify.Name = "Button_modify"
        Me.Button_modify.Size = New System.Drawing.Size(65, 28)
        Me.Button_modify.TabIndex = 3
        Me.Button_modify.Text = "Modify"
        Me.Button_modify.UseVisualStyleBackColor = True
        '
        'Button_rename
        '
        Me.Button_rename.Location = New System.Drawing.Point(148, 119)
        Me.Button_rename.Name = "Button_rename"
        Me.Button_rename.Size = New System.Drawing.Size(65, 28)
        Me.Button_rename.TabIndex = 3
        Me.Button_rename.Text = "Rename"
        Me.Button_rename.UseVisualStyleBackColor = True
        '
        'SlabThick
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(250, 238)
        Me.Controls.Add(Me.TextBox_addslabthk)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.Button_close)
        Me.Controls.Add(Me.Button_rename)
        Me.Controls.Add(Me.Button_modify)
        Me.Controls.Add(Me.Button_RemoveSlabTHK)
        Me.Controls.Add(Me.Button_addslabthk)
        Me.Controls.Add(Me.TextBox_thickness)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.ListBox_Slaplist)
        Me.Name = "SlabThick"
        Me.Text = "Slab Thick - BOQ Tools (c) DHK"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents ListBox_Slaplist As System.Windows.Forms.ListBox
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents TextBox_thickness As System.Windows.Forms.TextBox
    Friend WithEvents Button_addslabthk As System.Windows.Forms.Button
    Friend WithEvents Button_RemoveSlabTHK As System.Windows.Forms.Button
    Friend WithEvents Button_close As System.Windows.Forms.Button
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents TextBox_addslabthk As System.Windows.Forms.TextBox
    Friend WithEvents Button_modify As System.Windows.Forms.Button
    Friend WithEvents Button_rename As System.Windows.Forms.Button
End Class
