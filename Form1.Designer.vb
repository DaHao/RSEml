<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class Form1
    Inherits System.Windows.Forms.Form

    'Form 覆寫 Dispose 以清除元件清單。
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

    '為 Windows Form 設計工具的必要項
    Private components As System.ComponentModel.IContainer

    '注意: 以下為 Windows Form 設計工具所需的程序
    '可以使用 Windows Form 設計工具進行修改。
    '請不要使用程式碼編輯器進行修改。
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Me.BrowserFolder = New System.Windows.Forms.Button()
        Me.FolderPath = New System.Windows.Forms.TextBox()
        Me.FilePath = New System.Windows.Forms.TextBox()
        Me.BrowseFile = New System.Windows.Forms.Button()
        Me.MenuStrip1 = New System.Windows.Forms.MenuStrip()
        Me.說明ToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.Memo = New System.Windows.Forms.ToolStripMenuItem()
        Me.SMTPToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.SelfToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.reset = New System.Windows.Forms.Button()
        Me.exitbutton = New System.Windows.Forms.Button()
        Me.send = New System.Windows.Forms.Button()
        Me.smtpGroup = New System.Windows.Forms.GroupBox()
        Me.smtpReset = New System.Windows.Forms.Button()
        Me.smtpCancel = New System.Windows.Forms.Button()
        Me.smtpok = New System.Windows.Forms.Button()
        Me.userPw = New System.Windows.Forms.TextBox()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.userAcc = New System.Windows.Forms.TextBox()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.smtpSsl = New System.Windows.Forms.CheckBox()
        Me.smtpPort = New System.Windows.Forms.TextBox()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.smtpHost = New System.Windows.Forms.TextBox()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.ProgressBar1 = New System.Windows.Forms.ProgressBar()
        Me.TextBox1 = New System.Windows.Forms.TextBox()
        Me.MenuStrip1.SuspendLayout()
        Me.smtpGroup.SuspendLayout()
        Me.SuspendLayout()
        '
        'BrowserFolder
        '
        Me.BrowserFolder.Location = New System.Drawing.Point(263, 55)
        Me.BrowserFolder.Name = "BrowserFolder"
        Me.BrowserFolder.Size = New System.Drawing.Size(92, 23)
        Me.BrowserFolder.TabIndex = 13
        Me.BrowserFolder.Text = "BrowseFolder"
        Me.BrowserFolder.UseVisualStyleBackColor = True
        '
        'FolderPath
        '
        Me.FolderPath.Location = New System.Drawing.Point(12, 55)
        Me.FolderPath.Name = "FolderPath"
        Me.FolderPath.Size = New System.Drawing.Size(245, 22)
        Me.FolderPath.TabIndex = 12
        '
        'FilePath
        '
        Me.FilePath.Location = New System.Drawing.Point(12, 29)
        Me.FilePath.Name = "FilePath"
        Me.FilePath.Size = New System.Drawing.Size(245, 22)
        Me.FilePath.TabIndex = 10
        '
        'BrowseFile
        '
        Me.BrowseFile.Location = New System.Drawing.Point(263, 27)
        Me.BrowseFile.Name = "BrowseFile"
        Me.BrowseFile.Size = New System.Drawing.Size(92, 23)
        Me.BrowseFile.TabIndex = 11
        Me.BrowseFile.Text = "BrowseFile"
        Me.BrowseFile.UseVisualStyleBackColor = True
        '
        'MenuStrip1
        '
        Me.MenuStrip1.BackColor = System.Drawing.SystemColors.Control
        Me.MenuStrip1.Dock = System.Windows.Forms.DockStyle.None
        Me.MenuStrip1.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.說明ToolStripMenuItem, Me.SMTPToolStripMenuItem})
        Me.MenuStrip1.Location = New System.Drawing.Point(0, 0)
        Me.MenuStrip1.Name = "MenuStrip1"
        Me.MenuStrip1.Size = New System.Drawing.Size(106, 24)
        Me.MenuStrip1.TabIndex = 18
        Me.MenuStrip1.Text = "MenuStrip1"
        '
        '說明ToolStripMenuItem
        '
        Me.說明ToolStripMenuItem.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.Memo})
        Me.說明ToolStripMenuItem.Name = "說明ToolStripMenuItem"
        Me.說明ToolStripMenuItem.Size = New System.Drawing.Size(46, 20)
        Me.說明ToolStripMenuItem.Text = "Help"
        '
        'Memo
        '
        Me.Memo.Name = "Memo"
        Me.Memo.Size = New System.Drawing.Size(97, 22)
        Me.Memo.Text = "Hint"
        '
        'SMTPToolStripMenuItem
        '
        Me.SMTPToolStripMenuItem.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.SelfToolStripMenuItem})
        Me.SMTPToolStripMenuItem.Name = "SMTPToolStripMenuItem"
        Me.SMTPToolStripMenuItem.Size = New System.Drawing.Size(52, 20)
        Me.SMTPToolStripMenuItem.Text = "SMTP"
        '
        'SelfToolStripMenuItem
        '
        Me.SelfToolStripMenuItem.Name = "SelfToolStripMenuItem"
        Me.SelfToolStripMenuItem.Size = New System.Drawing.Size(98, 22)
        Me.SelfToolStripMenuItem.Text = "自訂"
        '
        'reset
        '
        Me.reset.Location = New System.Drawing.Point(148, 83)
        Me.reset.Name = "reset"
        Me.reset.Size = New System.Drawing.Size(69, 23)
        Me.reset.TabIndex = 15
        Me.reset.Text = "Clear"
        Me.reset.UseVisualStyleBackColor = True
        '
        'exitbutton
        '
        Me.exitbutton.Location = New System.Drawing.Point(229, 83)
        Me.exitbutton.Name = "exitbutton"
        Me.exitbutton.Size = New System.Drawing.Size(69, 23)
        Me.exitbutton.TabIndex = 16
        Me.exitbutton.Text = "Exit"
        Me.exitbutton.UseVisualStyleBackColor = True
        '
        'send
        '
        Me.send.Location = New System.Drawing.Point(67, 83)
        Me.send.Name = "send"
        Me.send.Size = New System.Drawing.Size(69, 23)
        Me.send.TabIndex = 14
        Me.send.Text = "Send"
        Me.send.UseVisualStyleBackColor = True
        '
        'smtpGroup
        '
        Me.smtpGroup.Controls.Add(Me.smtpReset)
        Me.smtpGroup.Controls.Add(Me.smtpCancel)
        Me.smtpGroup.Controls.Add(Me.smtpok)
        Me.smtpGroup.Controls.Add(Me.userPw)
        Me.smtpGroup.Controls.Add(Me.Label4)
        Me.smtpGroup.Controls.Add(Me.userAcc)
        Me.smtpGroup.Controls.Add(Me.Label3)
        Me.smtpGroup.Controls.Add(Me.smtpSsl)
        Me.smtpGroup.Controls.Add(Me.smtpPort)
        Me.smtpGroup.Controls.Add(Me.Label2)
        Me.smtpGroup.Controls.Add(Me.smtpHost)
        Me.smtpGroup.Controls.Add(Me.Label1)
        Me.smtpGroup.Font = New System.Drawing.Font("新細明體", 11.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(136, Byte))
        Me.smtpGroup.Location = New System.Drawing.Point(12, 112)
        Me.smtpGroup.Name = "smtpGroup"
        Me.smtpGroup.Size = New System.Drawing.Size(343, 243)
        Me.smtpGroup.TabIndex = 23
        Me.smtpGroup.TabStop = False
        Me.smtpGroup.Text = "SMTP"
        Me.smtpGroup.Visible = False
        '
        'smtpReset
        '
        Me.smtpReset.Font = New System.Drawing.Font("新細明體", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(136, Byte))
        Me.smtpReset.Location = New System.Drawing.Point(136, 195)
        Me.smtpReset.Name = "smtpReset"
        Me.smtpReset.Size = New System.Drawing.Size(69, 23)
        Me.smtpReset.TabIndex = 25
        Me.smtpReset.Text = "Reset"
        Me.smtpReset.UseVisualStyleBackColor = True
        '
        'smtpCancel
        '
        Me.smtpCancel.Font = New System.Drawing.Font("新細明體", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(136, Byte))
        Me.smtpCancel.Location = New System.Drawing.Point(211, 196)
        Me.smtpCancel.Name = "smtpCancel"
        Me.smtpCancel.Size = New System.Drawing.Size(69, 23)
        Me.smtpCancel.TabIndex = 26
        Me.smtpCancel.Text = "Cancel"
        Me.smtpCancel.UseVisualStyleBackColor = True
        '
        'smtpok
        '
        Me.smtpok.Font = New System.Drawing.Font("新細明體", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(136, Byte))
        Me.smtpok.Location = New System.Drawing.Point(55, 195)
        Me.smtpok.Name = "smtpok"
        Me.smtpok.Size = New System.Drawing.Size(75, 23)
        Me.smtpok.TabIndex = 24
        Me.smtpok.Text = "Save"
        Me.smtpok.UseVisualStyleBackColor = True
        '
        'userPw
        '
        Me.userPw.Font = New System.Drawing.Font("新細明體", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(136, Byte))
        Me.userPw.Location = New System.Drawing.Point(104, 155)
        Me.userPw.Name = "userPw"
        Me.userPw.PasswordChar = Global.Microsoft.VisualBasic.ChrW(42)
        Me.userPw.Size = New System.Drawing.Size(232, 22)
        Me.userPw.TabIndex = 8
        Me.userPw.UseSystemPasswordChar = True
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Font = New System.Drawing.Font("新細明體", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(136, Byte))
        Me.Label4.Location = New System.Drawing.Point(52, 160)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(46, 13)
        Me.Label4.TabIndex = 7
        Me.Label4.Text = "密碼："
        '
        'userAcc
        '
        Me.userAcc.Font = New System.Drawing.Font("新細明體", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(136, Byte))
        Me.userAcc.Location = New System.Drawing.Point(104, 124)
        Me.userAcc.Name = "userAcc"
        Me.userAcc.Size = New System.Drawing.Size(233, 22)
        Me.userAcc.TabIndex = 6
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Font = New System.Drawing.Font("新細明體", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(136, Byte))
        Me.Label3.Location = New System.Drawing.Point(13, 129)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(85, 13)
        Me.Label3.TabIndex = 5
        Me.Label3.Text = "使用者帳號："
        '
        'smtpSsl
        '
        Me.smtpSsl.AutoSize = True
        Me.smtpSsl.Font = New System.Drawing.Font("新細明體", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(136, Byte))
        Me.smtpSsl.Location = New System.Drawing.Point(11, 103)
        Me.smtpSsl.Name = "smtpSsl"
        Me.smtpSsl.Size = New System.Drawing.Size(99, 17)
        Me.smtpSsl.TabIndex = 4
        Me.smtpSsl.Text = "是否使用SSL"
        Me.smtpSsl.UseVisualStyleBackColor = True
        Me.smtpSsl.Visible = False
        '
        'smtpPort
        '
        Me.smtpPort.Font = New System.Drawing.Font("新細明體", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(136, Byte))
        Me.smtpPort.Location = New System.Drawing.Point(105, 55)
        Me.smtpPort.Name = "smtpPort"
        Me.smtpPort.Size = New System.Drawing.Size(38, 23)
        Me.smtpPort.TabIndex = 3
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Font = New System.Drawing.Font("新細明體", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(136, Byte))
        Me.Label2.Location = New System.Drawing.Point(27, 60)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(72, 13)
        Me.Label2.TabIndex = 2
        Me.Label2.Text = "連接埠號："
        '
        'smtpHost
        '
        Me.smtpHost.Font = New System.Drawing.Font("新細明體", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(136, Byte))
        Me.smtpHost.Location = New System.Drawing.Point(105, 24)
        Me.smtpHost.Name = "smtpHost"
        Me.smtpHost.Size = New System.Drawing.Size(232, 22)
        Me.smtpHost.TabIndex = 1
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Font = New System.Drawing.Font("新細明體", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(136, Byte))
        Me.Label1.Location = New System.Drawing.Point(8, 29)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(91, 13)
        Me.Label1.TabIndex = 0
        Me.Label1.Text = "SMTP伺服器："
        '
        'ProgressBar1
        '
        Me.ProgressBar1.Location = New System.Drawing.Point(229, 361)
        Me.ProgressBar1.Name = "ProgressBar1"
        Me.ProgressBar1.Size = New System.Drawing.Size(126, 15)
        Me.ProgressBar1.TabIndex = 24
        '
        'TextBox1
        '
        Me.TextBox1.Font = New System.Drawing.Font("新細明體", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(136, Byte))
        Me.TextBox1.Location = New System.Drawing.Point(12, 112)
        Me.TextBox1.Multiline = True
        Me.TextBox1.Name = "TextBox1"
        Me.TextBox1.ScrollBars = System.Windows.Forms.ScrollBars.Vertical
        Me.TextBox1.Size = New System.Drawing.Size(343, 243)
        Me.TextBox1.TabIndex = 17
        '
        'Form1
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 12.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(367, 384)
        Me.Controls.Add(Me.ProgressBar1)
        Me.Controls.Add(Me.reset)
        Me.Controls.Add(Me.exitbutton)
        Me.Controls.Add(Me.send)
        Me.Controls.Add(Me.BrowserFolder)
        Me.Controls.Add(Me.FolderPath)
        Me.Controls.Add(Me.FilePath)
        Me.Controls.Add(Me.BrowseFile)
        Me.Controls.Add(Me.MenuStrip1)
        Me.Controls.Add(Me.smtpGroup)
        Me.Controls.Add(Me.TextBox1)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
        Me.MaximizeBox = False
        Me.Name = "Form1"
        Me.Text = "EML寄發程式"
        Me.MenuStrip1.ResumeLayout(False)
        Me.MenuStrip1.PerformLayout()
        Me.smtpGroup.ResumeLayout(False)
        Me.smtpGroup.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents BrowserFolder As System.Windows.Forms.Button
    Friend WithEvents FolderPath As System.Windows.Forms.TextBox
    Friend WithEvents FilePath As System.Windows.Forms.TextBox
    Friend WithEvents BrowseFile As System.Windows.Forms.Button
    Friend WithEvents MenuStrip1 As System.Windows.Forms.MenuStrip
    Friend WithEvents 說明ToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents Memo As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents reset As System.Windows.Forms.Button
    Friend WithEvents exitbutton As System.Windows.Forms.Button
    Friend WithEvents send As System.Windows.Forms.Button
    Friend WithEvents SMTPToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents SelfToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents smtpGroup As System.Windows.Forms.GroupBox
    Friend WithEvents smtpHost As System.Windows.Forms.TextBox
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents smtpPort As System.Windows.Forms.TextBox
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents smtpSsl As System.Windows.Forms.CheckBox
    Friend WithEvents userPw As System.Windows.Forms.TextBox
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents userAcc As System.Windows.Forms.TextBox
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents smtpCancel As System.Windows.Forms.Button
    Friend WithEvents smtpok As System.Windows.Forms.Button
    Friend WithEvents smtpReset As System.Windows.Forms.Button
    Friend WithEvents ProgressBar1 As System.Windows.Forms.ProgressBar
    Friend WithEvents TextBox1 As System.Windows.Forms.TextBox

End Class
