<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmPigObjFs
    Inherits System.Windows.Forms.Form

    'Form 重写 Dispose，以清理组件列表。
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

    'Windows 窗体设计器所必需的
    Private components As System.ComponentModel.IContainer

    '注意: 以下过程是 Windows 窗体设计器所必需的
    '可以使用 Windows 窗体设计器修改它。
    '不要使用代码编辑器修改它。
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container()
        Me.ContextMenuStrip1 = New System.Windows.Forms.ContextMenuStrip(Me.components)
        Me.MenuStrip1 = New System.Windows.Forms.MenuStrip()
        Me.FileToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.ExitToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.PFileSystemObjectToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.GetFileToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.GetFolderToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.FileExistsToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.FolderExistsToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.CreateFolderToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.PTextStreamToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.ReadFileToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.WriteFileToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.HelpToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.OnlineDocumentationToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.SeowPhongStudioToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.tbMain = New System.Windows.Forms.TextBox()
        Me.MenuStrip1.SuspendLayout()
        Me.SuspendLayout()
        '
        'ContextMenuStrip1
        '
        Me.ContextMenuStrip1.Name = "ContextMenuStrip1"
        Me.ContextMenuStrip1.Size = New System.Drawing.Size(61, 4)
        '
        'MenuStrip1
        '
        Me.MenuStrip1.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.FileToolStripMenuItem, Me.PFileSystemObjectToolStripMenuItem, Me.PTextStreamToolStripMenuItem, Me.HelpToolStripMenuItem})
        Me.MenuStrip1.Location = New System.Drawing.Point(0, 0)
        Me.MenuStrip1.Name = "MenuStrip1"
        Me.MenuStrip1.Size = New System.Drawing.Size(765, 28)
        Me.MenuStrip1.TabIndex = 1
        Me.MenuStrip1.Text = "MenuStrip1"
        '
        'FileToolStripMenuItem
        '
        Me.FileToolStripMenuItem.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.ExitToolStripMenuItem})
        Me.FileToolStripMenuItem.Name = "FileToolStripMenuItem"
        Me.FileToolStripMenuItem.Size = New System.Drawing.Size(46, 24)
        Me.FileToolStripMenuItem.Text = "File"
        '
        'ExitToolStripMenuItem
        '
        Me.ExitToolStripMenuItem.Name = "ExitToolStripMenuItem"
        Me.ExitToolStripMenuItem.Size = New System.Drawing.Size(152, 24)
        Me.ExitToolStripMenuItem.Text = "Exit"
        '
        'PFileSystemObjectToolStripMenuItem
        '
        Me.PFileSystemObjectToolStripMenuItem.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.GetFileToolStripMenuItem, Me.GetFolderToolStripMenuItem, Me.FileExistsToolStripMenuItem, Me.FolderExistsToolStripMenuItem, Me.CreateFolderToolStripMenuItem})
        Me.PFileSystemObjectToolStripMenuItem.Name = "PFileSystemObjectToolStripMenuItem"
        Me.PFileSystemObjectToolStripMenuItem.Size = New System.Drawing.Size(158, 24)
        Me.PFileSystemObjectToolStripMenuItem.Text = "pFileSystemObject"
        '
        'GetFileToolStripMenuItem
        '
        Me.GetFileToolStripMenuItem.Name = "GetFileToolStripMenuItem"
        Me.GetFileToolStripMenuItem.Size = New System.Drawing.Size(173, 24)
        Me.GetFileToolStripMenuItem.Text = "GetFile"
        '
        'GetFolderToolStripMenuItem
        '
        Me.GetFolderToolStripMenuItem.Name = "GetFolderToolStripMenuItem"
        Me.GetFolderToolStripMenuItem.Size = New System.Drawing.Size(173, 24)
        Me.GetFolderToolStripMenuItem.Text = "GetFolder"
        '
        'FileExistsToolStripMenuItem
        '
        Me.FileExistsToolStripMenuItem.Name = "FileExistsToolStripMenuItem"
        Me.FileExistsToolStripMenuItem.Size = New System.Drawing.Size(173, 24)
        Me.FileExistsToolStripMenuItem.Text = "FileExists"
        '
        'FolderExistsToolStripMenuItem
        '
        Me.FolderExistsToolStripMenuItem.Name = "FolderExistsToolStripMenuItem"
        Me.FolderExistsToolStripMenuItem.Size = New System.Drawing.Size(173, 24)
        Me.FolderExistsToolStripMenuItem.Text = "FolderExists"
        '
        'CreateFolderToolStripMenuItem
        '
        Me.CreateFolderToolStripMenuItem.Name = "CreateFolderToolStripMenuItem"
        Me.CreateFolderToolStripMenuItem.Size = New System.Drawing.Size(173, 24)
        Me.CreateFolderToolStripMenuItem.Text = "CreateFolder"
        '
        'PTextStreamToolStripMenuItem
        '
        Me.PTextStreamToolStripMenuItem.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.ReadFileToolStripMenuItem, Me.WriteFileToolStripMenuItem})
        Me.PTextStreamToolStripMenuItem.Name = "PTextStreamToolStripMenuItem"
        Me.PTextStreamToolStripMenuItem.Size = New System.Drawing.Size(115, 24)
        Me.PTextStreamToolStripMenuItem.Text = "pTextStream"
        '
        'ReadFileToolStripMenuItem
        '
        Me.ReadFileToolStripMenuItem.Name = "ReadFileToolStripMenuItem"
        Me.ReadFileToolStripMenuItem.Size = New System.Drawing.Size(152, 24)
        Me.ReadFileToolStripMenuItem.Text = "Read File"
        '
        'WriteFileToolStripMenuItem
        '
        Me.WriteFileToolStripMenuItem.Name = "WriteFileToolStripMenuItem"
        Me.WriteFileToolStripMenuItem.Size = New System.Drawing.Size(152, 24)
        Me.WriteFileToolStripMenuItem.Text = "Write File"
        '
        'HelpToolStripMenuItem
        '
        Me.HelpToolStripMenuItem.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.OnlineDocumentationToolStripMenuItem, Me.SeowPhongStudioToolStripMenuItem})
        Me.HelpToolStripMenuItem.Name = "HelpToolStripMenuItem"
        Me.HelpToolStripMenuItem.Size = New System.Drawing.Size(56, 24)
        Me.HelpToolStripMenuItem.Text = "Help"
        '
        'OnlineDocumentationToolStripMenuItem
        '
        Me.OnlineDocumentationToolStripMenuItem.Name = "OnlineDocumentationToolStripMenuItem"
        Me.OnlineDocumentationToolStripMenuItem.Size = New System.Drawing.Size(241, 24)
        Me.OnlineDocumentationToolStripMenuItem.Text = "Online documentation"
        '
        'SeowPhongStudioToolStripMenuItem
        '
        Me.SeowPhongStudioToolStripMenuItem.Name = "SeowPhongStudioToolStripMenuItem"
        Me.SeowPhongStudioToolStripMenuItem.Size = New System.Drawing.Size(241, 24)
        Me.SeowPhongStudioToolStripMenuItem.Text = "Seow Phong Studio"
        '
        'tbMain
        '
        Me.tbMain.Dock = System.Windows.Forms.DockStyle.Fill
        Me.tbMain.Location = New System.Drawing.Point(0, 28)
        Me.tbMain.Multiline = True
        Me.tbMain.Name = "tbMain"
        Me.tbMain.ScrollBars = System.Windows.Forms.ScrollBars.Vertical
        Me.tbMain.Size = New System.Drawing.Size(765, 413)
        Me.tbMain.TabIndex = 2
        '
        'frmPigObjFs
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(8.0!, 15.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(765, 441)
        Me.Controls.Add(Me.tbMain)
        Me.Controls.Add(Me.MenuStrip1)
        Me.MainMenuStrip = Me.MenuStrip1
        Me.Name = "frmPigObjFs"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Piggy Objectified FileSystemObject"
        Me.MenuStrip1.ResumeLayout(False)
        Me.MenuStrip1.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents ContextMenuStrip1 As System.Windows.Forms.ContextMenuStrip
    Friend WithEvents MenuStrip1 As System.Windows.Forms.MenuStrip
    Friend WithEvents FileToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents ExitToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents PFileSystemObjectToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents GetFileToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents GetFolderToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents FileExistsToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents FolderExistsToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents CreateFolderToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents PTextStreamToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents ReadFileToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents WriteFileToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents HelpToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents OnlineDocumentationToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents SeowPhongStudioToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents tbMain As System.Windows.Forms.TextBox

End Class
