<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class Form11
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(Form11))
        Me.ShapeContainer1 = New Microsoft.VisualBasic.PowerPacks.ShapeContainer()
        Me.OvalShape1 = New Microsoft.VisualBasic.PowerPacks.OvalShape()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.TextBox1 = New System.Windows.Forms.TextBox()
        Me.TextBox2 = New System.Windows.Forms.TextBox()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.Button1 = New System.Windows.Forms.Button()
        Me.CheckBox1 = New System.Windows.Forms.CheckBox()
        Me.Button2 = New System.Windows.Forms.Button()
        Me.SuspendLayout()
        '
        'ShapeContainer1
        '
        Me.ShapeContainer1.Location = New System.Drawing.Point(0, 0)
        Me.ShapeContainer1.Margin = New System.Windows.Forms.Padding(0)
        Me.ShapeContainer1.Name = "ShapeContainer1"
        Me.ShapeContainer1.Shapes.AddRange(New Microsoft.VisualBasic.PowerPacks.Shape() {Me.OvalShape1})
        Me.ShapeContainer1.Size = New System.Drawing.Size(1317, 642)
        Me.ShapeContainer1.TabIndex = 0
        Me.ShapeContainer1.TabStop = False
        '
        'OvalShape1
        '
        Me.OvalShape1.BackColor = System.Drawing.Color.White
        Me.OvalShape1.BackgroundImage = CType(resources.GetObject("OvalShape1.BackgroundImage"), System.Drawing.Image)
        Me.OvalShape1.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch
        Me.OvalShape1.BorderColor = System.Drawing.SystemColors.Highlight
        Me.OvalShape1.FillColor = System.Drawing.Color.White
        Me.OvalShape1.Location = New System.Drawing.Point(641, 98)
        Me.OvalShape1.Name = "OvalShape1"
        Me.OvalShape1.Size = New System.Drawing.Size(154, 150)
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.BackColor = System.Drawing.Color.Transparent
        Me.Label1.Font = New System.Drawing.Font("Engravers MT", 14.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label1.ForeColor = System.Drawing.SystemColors.ButtonFace
        Me.Label1.Location = New System.Drawing.Point(604, 315)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(206, 22)
        Me.Label1.TabIndex = 1
        Me.Label1.Text = "PLEASE LOGIN"
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.BackColor = System.Drawing.Color.Transparent
        Me.Label2.Font = New System.Drawing.Font("Lucida Bright", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label2.ForeColor = System.Drawing.SystemColors.ButtonFace
        Me.Label2.Location = New System.Drawing.Point(563, 383)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(84, 15)
        Me.Label2.TabIndex = 2
        Me.Label2.Text = "USER NAME"
        '
        'TextBox1
        '
        Me.TextBox1.Location = New System.Drawing.Point(669, 380)
        Me.TextBox1.Name = "TextBox1"
        Me.TextBox1.Size = New System.Drawing.Size(141, 20)
        Me.TextBox1.TabIndex = 3
        '
        'TextBox2
        '
        Me.TextBox2.Location = New System.Drawing.Point(669, 420)
        Me.TextBox2.Name = "TextBox2"
        Me.TextBox2.Size = New System.Drawing.Size(141, 20)
        Me.TextBox2.TabIndex = 4
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.BackColor = System.Drawing.Color.Transparent
        Me.Label3.Font = New System.Drawing.Font("Lucida Bright", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label3.ForeColor = System.Drawing.SystemColors.ButtonFace
        Me.Label3.Location = New System.Drawing.Point(565, 422)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(82, 15)
        Me.Label3.TabIndex = 5
        Me.Label3.Text = "PASSWORD"
        '
        'Button1
        '
        Me.Button1.Location = New System.Drawing.Point(669, 467)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(75, 23)
        Me.Button1.TabIndex = 6
        Me.Button1.Text = "LOGIN"
        Me.Button1.UseVisualStyleBackColor = True
        '
        'CheckBox1
        '
        Me.CheckBox1.AutoSize = True
        Me.CheckBox1.BackColor = System.Drawing.Color.Transparent
        Me.CheckBox1.FlatStyle = System.Windows.Forms.FlatStyle.Popup
        Me.CheckBox1.ForeColor = System.Drawing.Color.White
        Me.CheckBox1.Location = New System.Drawing.Point(669, 444)
        Me.CheckBox1.Name = "CheckBox1"
        Me.CheckBox1.Size = New System.Drawing.Size(122, 17)
        Me.CheckBox1.TabIndex = 7
        Me.CheckBox1.Text = "Tampilkan Password"
        Me.CheckBox1.UseVisualStyleBackColor = False
        '
        'Button2
        '
        Me.Button2.Location = New System.Drawing.Point(750, 467)
        Me.Button2.Name = "Button2"
        Me.Button2.Size = New System.Drawing.Size(75, 23)
        Me.Button2.TabIndex = 8
        Me.Button2.Text = "KELUAR"
        Me.Button2.UseVisualStyleBackColor = True
        '
        'Form11
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackgroundImage = CType(resources.GetObject("$this.BackgroundImage"), System.Drawing.Image)
        Me.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch
        Me.ClientSize = New System.Drawing.Size(1317, 642)
        Me.Controls.Add(Me.Button2)
        Me.Controls.Add(Me.CheckBox1)
        Me.Controls.Add(Me.Button1)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.TextBox2)
        Me.Controls.Add(Me.TextBox1)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.ShapeContainer1)
        Me.Name = "Form11"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.Manual
        Me.Text = "Form11"
        Me.WindowState = System.Windows.Forms.FormWindowState.Maximized
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents ShapeContainer1 As Microsoft.VisualBasic.PowerPacks.ShapeContainer
    Friend WithEvents OvalShape1 As Microsoft.VisualBasic.PowerPacks.OvalShape
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents TextBox1 As System.Windows.Forms.TextBox
    Friend WithEvents TextBox2 As System.Windows.Forms.TextBox
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents Button1 As System.Windows.Forms.Button
    Friend WithEvents CheckBox1 As System.Windows.Forms.CheckBox
    Friend WithEvents Button2 As System.Windows.Forms.Button
End Class
