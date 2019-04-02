<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class newDeal
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
        Me.Label1 = New System.Windows.Forms.Label()
        Me.CustomerName = New System.Windows.Forms.TextBox()
        Me.GroupBox1 = New System.Windows.Forms.GroupBox()
        Me.HPEOption = New System.Windows.Forms.RadioButton()
        Me.DellOption = New System.Windows.Forms.RadioButton()
        Me.HPIOption = New System.Windows.Forms.RadioButton()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.DealID = New System.Windows.Forms.TextBox()
        Me.OKButton = New System.Windows.Forms.Button()
        Me.tCancelButton = New System.Windows.Forms.Button()
        Me.GroupBox2 = New System.Windows.Forms.GroupBox()
        Me.cWestcoast = New System.Windows.Forms.CheckBox()
        Me.cTechData = New System.Windows.Forms.CheckBox()
        Me.cIngram = New System.Windows.Forms.CheckBox()
        Me.GroupBox1.SuspendLayout()
        Me.GroupBox2.SuspendLayout()
        Me.SuspendLayout()
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(16, 11)
        Me.Label1.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(113, 17)
        Me.Label1.TabIndex = 0
        Me.Label1.Text = "Customer Name:"
        '
        'CustomerName
        '
        Me.CustomerName.Location = New System.Drawing.Point(20, 31)
        Me.CustomerName.Margin = New System.Windows.Forms.Padding(4)
        Me.CustomerName.Name = "CustomerName"
        Me.CustomerName.Size = New System.Drawing.Size(385, 22)
        Me.CustomerName.TabIndex = 1
        '
        'GroupBox1
        '
        Me.GroupBox1.Controls.Add(Me.HPEOption)
        Me.GroupBox1.Controls.Add(Me.DellOption)
        Me.GroupBox1.Controls.Add(Me.HPIOption)
        Me.GroupBox1.Location = New System.Drawing.Point(20, 64)
        Me.GroupBox1.Margin = New System.Windows.Forms.Padding(4)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Padding = New System.Windows.Forms.Padding(4)
        Me.GroupBox1.Size = New System.Drawing.Size(387, 58)
        Me.GroupBox1.TabIndex = 2
        Me.GroupBox1.TabStop = False
        Me.GroupBox1.Text = "Vendor"
        '
        'HPEOption
        '
        Me.HPEOption.AutoSize = True
        Me.HPEOption.Location = New System.Drawing.Point(157, 23)
        Me.HPEOption.Margin = New System.Windows.Forms.Padding(4)
        Me.HPEOption.Name = "HPEOption"
        Me.HPEOption.Size = New System.Drawing.Size(57, 21)
        Me.HPEOption.TabIndex = 2
        Me.HPEOption.TabStop = True
        Me.HPEOption.Text = "HPE"
        Me.HPEOption.UseVisualStyleBackColor = True
        '
        'DellOption
        '
        Me.DellOption.AutoSize = True
        Me.DellOption.Checked = True
        Me.DellOption.Location = New System.Drawing.Point(287, 23)
        Me.DellOption.Margin = New System.Windows.Forms.Padding(4)
        Me.DellOption.Name = "DellOption"
        Me.DellOption.Size = New System.Drawing.Size(53, 21)
        Me.DellOption.TabIndex = 1
        Me.DellOption.TabStop = True
        Me.DellOption.Text = "Dell"
        Me.DellOption.UseVisualStyleBackColor = True
        '
        'HPIOption
        '
        Me.HPIOption.AutoSize = True
        Me.HPIOption.Location = New System.Drawing.Point(33, 23)
        Me.HPIOption.Margin = New System.Windows.Forms.Padding(4)
        Me.HPIOption.Name = "HPIOption"
        Me.HPIOption.Size = New System.Drawing.Size(51, 21)
        Me.HPIOption.TabIndex = 0
        Me.HPIOption.TabStop = True
        Me.HPIOption.Text = "HPI"
        Me.HPIOption.UseVisualStyleBackColor = True
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(16, 208)
        Me.Label2.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(232, 17)
        Me.Label2.TabIndex = 4
        Me.Label2.Text = "Deal ID / SQ ID / Deal Reg Number:"
        '
        'DealID
        '
        Me.DealID.Location = New System.Drawing.Point(20, 228)
        Me.DealID.Margin = New System.Windows.Forms.Padding(4)
        Me.DealID.Name = "DealID"
        Me.DealID.Size = New System.Drawing.Size(381, 22)
        Me.DealID.TabIndex = 5
        '
        'OKButton
        '
        Me.OKButton.Location = New System.Drawing.Point(20, 260)
        Me.OKButton.Margin = New System.Windows.Forms.Padding(4)
        Me.OKButton.Name = "OKButton"
        Me.OKButton.Size = New System.Drawing.Size(184, 36)
        Me.OKButton.TabIndex = 6
        Me.OKButton.Text = "OK"
        Me.OKButton.UseVisualStyleBackColor = True
        '
        'tCancelButton
        '
        Me.tCancelButton.Location = New System.Drawing.Point(223, 260)
        Me.tCancelButton.Margin = New System.Windows.Forms.Padding(4)
        Me.tCancelButton.Name = "tCancelButton"
        Me.tCancelButton.Size = New System.Drawing.Size(184, 36)
        Me.tCancelButton.TabIndex = 7
        Me.tCancelButton.Text = "Cancel"
        Me.tCancelButton.UseVisualStyleBackColor = True
        '
        'GroupBox2
        '
        Me.GroupBox2.Controls.Add(Me.cWestcoast)
        Me.GroupBox2.Controls.Add(Me.cTechData)
        Me.GroupBox2.Controls.Add(Me.cIngram)
        Me.GroupBox2.Location = New System.Drawing.Point(17, 131)
        Me.GroupBox2.Margin = New System.Windows.Forms.Padding(4)
        Me.GroupBox2.Name = "GroupBox2"
        Me.GroupBox2.Padding = New System.Windows.Forms.Padding(4)
        Me.GroupBox2.Size = New System.Drawing.Size(384, 74)
        Me.GroupBox2.TabIndex = 9
        Me.GroupBox2.TabStop = False
        Me.GroupBox2.Text = "Distribution Partners"
        '
        'cWestcoast
        '
        Me.cWestcoast.AutoSize = True
        Me.cWestcoast.Location = New System.Drawing.Point(11, 46)
        Me.cWestcoast.Margin = New System.Windows.Forms.Padding(4)
        Me.cWestcoast.Name = "cWestcoast"
        Me.cWestcoast.Size = New System.Drawing.Size(96, 21)
        Me.cWestcoast.TabIndex = 2
        Me.cWestcoast.Text = "Westcoast"
        Me.cWestcoast.UseVisualStyleBackColor = True
        '
        'cTechData
        '
        Me.cTechData.AutoSize = True
        Me.cTechData.Location = New System.Drawing.Point(205, 17)
        Me.cTechData.Margin = New System.Windows.Forms.Padding(4)
        Me.cTechData.Name = "cTechData"
        Me.cTechData.Size = New System.Drawing.Size(90, 21)
        Me.cTechData.TabIndex = 1
        Me.cTechData.Text = "Techdata"
        Me.cTechData.UseVisualStyleBackColor = True
        '
        'cIngram
        '
        Me.cIngram.AutoSize = True
        Me.cIngram.Location = New System.Drawing.Point(11, 18)
        Me.cIngram.Margin = New System.Windows.Forms.Padding(4)
        Me.cIngram.Name = "cIngram"
        Me.cIngram.Size = New System.Drawing.Size(73, 21)
        Me.cIngram.TabIndex = 0
        Me.cIngram.Text = "Ingram"
        Me.cIngram.UseVisualStyleBackColor = True
        '
        'newDeal
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(8.0!, 16.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(423, 312)
        Me.Controls.Add(Me.GroupBox2)
        Me.Controls.Add(Me.tCancelButton)
        Me.Controls.Add(Me.OKButton)
        Me.Controls.Add(Me.DealID)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.GroupBox1)
        Me.Controls.Add(Me.CustomerName)
        Me.Controls.Add(Me.Label1)
        Me.Margin = New System.Windows.Forms.Padding(4)
        Me.Name = "newDeal"
        Me.Text = "Deal Information"
        Me.GroupBox1.ResumeLayout(False)
        Me.GroupBox1.PerformLayout()
        Me.GroupBox2.ResumeLayout(False)
        Me.GroupBox2.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents CustomerName As System.Windows.Forms.TextBox
    Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
    Friend WithEvents DellOption As System.Windows.Forms.RadioButton
    Friend WithEvents HPIOption As System.Windows.Forms.RadioButton
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents DealID As System.Windows.Forms.TextBox
    Friend WithEvents OKButton As System.Windows.Forms.Button
    Friend WithEvents tCancelButton As System.Windows.Forms.Button
    Friend WithEvents GroupBox2 As Windows.Forms.GroupBox
    Friend WithEvents cWestcoast As Windows.Forms.CheckBox
    Friend WithEvents cTechData As Windows.Forms.CheckBox
    Friend WithEvents cIngram As Windows.Forms.CheckBox
    Friend WithEvents HPEOption As Windows.Forms.RadioButton
End Class
