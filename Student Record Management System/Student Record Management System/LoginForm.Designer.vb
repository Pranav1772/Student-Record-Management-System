<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class LoginForm
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
        Me.components = New System.ComponentModel.Container()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(LoginForm))
        Me.UserName = New System.Windows.Forms.Label()
        Me.PassWord = New System.Windows.Forms.Label()
        Me.TextBox1 = New System.Windows.Forms.TextBox()
        Me.TextBox2 = New System.Windows.Forms.TextBox()
        Me.Blogin = New System.Windows.Forms.Button()
        Me.ErrorProvider1 = New System.Windows.Forms.ErrorProvider(Me.components)
        CType(Me.ErrorProvider1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'UserName
        '
        Me.UserName.AutoSize = True
        Me.UserName.ForeColor = System.Drawing.Color.White
        Me.UserName.Location = New System.Drawing.Point(58, 45)
        Me.UserName.Name = "UserName"
        Me.UserName.Size = New System.Drawing.Size(61, 13)
        Me.UserName.TabIndex = 0
        Me.UserName.Text = "Username :"
        '
        'PassWord
        '
        Me.PassWord.AutoSize = True
        Me.PassWord.ForeColor = System.Drawing.Color.White
        Me.PassWord.Location = New System.Drawing.Point(58, 97)
        Me.PassWord.Name = "PassWord"
        Me.PassWord.Size = New System.Drawing.Size(59, 13)
        Me.PassWord.TabIndex = 1
        Me.PassWord.Text = "Password :"
        '
        'TextBox1
        '
        Me.TextBox1.Location = New System.Drawing.Point(155, 38)
        Me.TextBox1.Name = "TextBox1"
        Me.TextBox1.Size = New System.Drawing.Size(246, 20)
        Me.TextBox1.TabIndex = 2
        '
        'TextBox2
        '
        Me.TextBox2.Location = New System.Drawing.Point(155, 90)
        Me.TextBox2.Name = "TextBox2"
        Me.TextBox2.Size = New System.Drawing.Size(246, 20)
        Me.TextBox2.TabIndex = 3
        '
        'Blogin
        '
        Me.Blogin.Location = New System.Drawing.Point(213, 131)
        Me.Blogin.Name = "Blogin"
        Me.Blogin.Size = New System.Drawing.Size(103, 35)
        Me.Blogin.TabIndex = 4
        Me.Blogin.Text = "Login"
        Me.Blogin.UseVisualStyleBackColor = True
        '
        'ErrorProvider1
        '
        Me.ErrorProvider1.ContainerControl = Me
        '
        'LoginForm
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(64, Byte), Integer), CType(CType(0, Byte), Integer))
        Me.ClientSize = New System.Drawing.Size(492, 178)
        Me.Controls.Add(Me.Blogin)
        Me.Controls.Add(Me.TextBox2)
        Me.Controls.Add(Me.TextBox1)
        Me.Controls.Add(Me.PassWord)
        Me.Controls.Add(Me.UserName)
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Name = "LoginForm"
        Me.Text = "Student Record Management System : Login"
        CType(Me.ErrorProvider1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents UserName As System.Windows.Forms.Label
    Friend WithEvents PassWord As System.Windows.Forms.Label
    Friend WithEvents TextBox1 As System.Windows.Forms.TextBox
    Friend WithEvents TextBox2 As System.Windows.Forms.TextBox
    Friend WithEvents Blogin As System.Windows.Forms.Button
    Friend WithEvents ErrorProvider1 As System.Windows.Forms.ErrorProvider
End Class
