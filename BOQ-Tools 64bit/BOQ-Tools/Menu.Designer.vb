<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class Menu
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
        Me.MenuStrip1 = New System.Windows.Forms.MenuStrip()
        Me.FilesToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.ExitToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.ToolsToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.ReinforcedConcreteToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.SteelStructureToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.FinishingWorksToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.HelpsToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.AboutAuthorToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.MenuStrip1.SuspendLayout()
        Me.SuspendLayout()
        '
        'MenuStrip1
        '
        Me.MenuStrip1.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.FilesToolStripMenuItem, Me.ToolsToolStripMenuItem, Me.HelpsToolStripMenuItem})
        Me.MenuStrip1.Location = New System.Drawing.Point(0, 0)
        Me.MenuStrip1.Name = "MenuStrip1"
        Me.MenuStrip1.Size = New System.Drawing.Size(380, 24)
        Me.MenuStrip1.TabIndex = 0
        Me.MenuStrip1.Text = "MenuStrip1"
        '
        'FilesToolStripMenuItem
        '
        Me.FilesToolStripMenuItem.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.ExitToolStripMenuItem})
        Me.FilesToolStripMenuItem.Name = "FilesToolStripMenuItem"
        Me.FilesToolStripMenuItem.ShortcutKeys = CType((System.Windows.Forms.Keys.Alt Or System.Windows.Forms.Keys.F), System.Windows.Forms.Keys)
        Me.FilesToolStripMenuItem.Size = New System.Drawing.Size(42, 20)
        Me.FilesToolStripMenuItem.Text = "&Files"
        '
        'ExitToolStripMenuItem
        '
        Me.ExitToolStripMenuItem.Name = "ExitToolStripMenuItem"
        Me.ExitToolStripMenuItem.Size = New System.Drawing.Size(92, 22)
        Me.ExitToolStripMenuItem.Text = "&Exit"
        '
        'ToolsToolStripMenuItem
        '
        Me.ToolsToolStripMenuItem.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.ReinforcedConcreteToolStripMenuItem, Me.SteelStructureToolStripMenuItem, Me.FinishingWorksToolStripMenuItem})
        Me.ToolsToolStripMenuItem.Name = "ToolsToolStripMenuItem"
        Me.ToolsToolStripMenuItem.ShortcutKeyDisplayString = "T"
        Me.ToolsToolStripMenuItem.Size = New System.Drawing.Size(48, 20)
        Me.ToolsToolStripMenuItem.Text = "&Tools"
        '
        'ReinforcedConcreteToolStripMenuItem
        '
        Me.ReinforcedConcreteToolStripMenuItem.Name = "ReinforcedConcreteToolStripMenuItem"
        Me.ReinforcedConcreteToolStripMenuItem.ShortcutKeys = CType((System.Windows.Forms.Keys.Alt Or System.Windows.Forms.Keys.R), System.Windows.Forms.Keys)
        Me.ReinforcedConcreteToolStripMenuItem.Size = New System.Drawing.Size(217, 22)
        Me.ReinforcedConcreteToolStripMenuItem.Text = "&Reinforced concrete"
        '
        'SteelStructureToolStripMenuItem
        '
        Me.SteelStructureToolStripMenuItem.Name = "SteelStructureToolStripMenuItem"
        Me.SteelStructureToolStripMenuItem.ShortcutKeys = CType((System.Windows.Forms.Keys.Alt Or System.Windows.Forms.Keys.S), System.Windows.Forms.Keys)
        Me.SteelStructureToolStripMenuItem.Size = New System.Drawing.Size(217, 22)
        Me.SteelStructureToolStripMenuItem.Text = "&Steel structure"
        '
        'FinishingWorksToolStripMenuItem
        '
        Me.FinishingWorksToolStripMenuItem.Name = "FinishingWorksToolStripMenuItem"
        Me.FinishingWorksToolStripMenuItem.Size = New System.Drawing.Size(217, 22)
        Me.FinishingWorksToolStripMenuItem.Text = "&Finishing works"
        '
        'HelpsToolStripMenuItem
        '
        Me.HelpsToolStripMenuItem.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.AboutAuthorToolStripMenuItem})
        Me.HelpsToolStripMenuItem.Name = "HelpsToolStripMenuItem"
        Me.HelpsToolStripMenuItem.ShortcutKeys = CType((System.Windows.Forms.Keys.Alt Or System.Windows.Forms.Keys.H), System.Windows.Forms.Keys)
        Me.HelpsToolStripMenuItem.Size = New System.Drawing.Size(49, 20)
        Me.HelpsToolStripMenuItem.Text = "&Helps"
        '
        'AboutAuthorToolStripMenuItem
        '
        Me.AboutAuthorToolStripMenuItem.Name = "AboutAuthorToolStripMenuItem"
        Me.AboutAuthorToolStripMenuItem.Size = New System.Drawing.Size(152, 22)
        Me.AboutAuthorToolStripMenuItem.Text = "&About author"
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(228, 239)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(140, 13)
        Me.Label1.TabIndex = 1
        Me.Label1.Text = "by Đỗ Hữu Khương 2014 (c)"
        '
        'Menu
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(380, 261)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.MenuStrip1)
        Me.MainMenuStrip = Me.MenuStrip1
        Me.Name = "Menu"
        Me.Text = "BOQ Tools (c) DHK 2014"
        Me.MenuStrip1.ResumeLayout(False)
        Me.MenuStrip1.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents MenuStrip1 As System.Windows.Forms.MenuStrip
    Friend WithEvents FilesToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents ExitToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents ToolsToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents ReinforcedConcreteToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents SteelStructureToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents FinishingWorksToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents HelpsToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents AboutAuthorToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents Label1 As System.Windows.Forms.Label
End Class
