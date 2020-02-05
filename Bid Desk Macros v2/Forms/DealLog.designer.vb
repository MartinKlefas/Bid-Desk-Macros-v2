<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()>
Partial Class AddDeal
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()>
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
    <System.Diagnostics.DebuggerStepThrough()>
    Private Sub InitializeComponent()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.CustomerName = New System.Windows.Forms.TextBox()
        Me.VendorGroupBox = New System.Windows.Forms.GroupBox()
        Me.LenovoOption = New System.Windows.Forms.RadioButton()
        Me.btnMS = New System.Windows.Forms.RadioButton()
        Me.HPEOption = New System.Windows.Forms.RadioButton()
        Me.DellOption = New System.Windows.Forms.RadioButton()
        Me.HPIOption = New System.Windows.Forms.RadioButton()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.DealID = New System.Windows.Forms.TextBox()
        Me.OKButton = New System.Windows.Forms.Button()
        Me.tCancelButton = New System.Windows.Forms.Button()
        Me.GroupBox2 = New System.Windows.Forms.GroupBox()
        Me.chkExertis = New System.Windows.Forms.CheckBox()
        Me.cWestcoast = New System.Windows.Forms.CheckBox()
        Me.cTechData = New System.Windows.Forms.CheckBox()
        Me.cIngram = New System.Windows.Forms.CheckBox()
        Me.BackgroundWorker1 = New System.ComponentModel.BackgroundWorker()
        Me.VendorGroupBox.SuspendLayout()
        Me.GroupBox2.SuspendLayout()
        Me.SuspendLayout()
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(12, 9)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(85, 13)
        Me.Label1.TabIndex = 0
        Me.Label1.Text = "Customer Name:"
        '
        'CustomerName
        '
        Me.CustomerName.Location = New System.Drawing.Point(15, 25)
        Me.CustomerName.Name = "CustomerName"
        Me.CustomerName.Size = New System.Drawing.Size(290, 20)
        Me.CustomerName.TabIndex = 1
        '
        'VendorGroupBox
        '
        Me.VendorGroupBox.Controls.Add(Me.LenovoOption)
        Me.VendorGroupBox.Controls.Add(Me.btnMS)
        Me.VendorGroupBox.Controls.Add(Me.HPEOption)
        Me.VendorGroupBox.Controls.Add(Me.DellOption)
        Me.VendorGroupBox.Controls.Add(Me.HPIOption)
        Me.VendorGroupBox.Location = New System.Drawing.Point(15, 52)
        Me.VendorGroupBox.Name = "VendorGroupBox"
        Me.VendorGroupBox.Size = New System.Drawing.Size(290, 66)
        Me.VendorGroupBox.TabIndex = 2
        Me.VendorGroupBox.TabStop = False
        Me.VendorGroupBox.Text = "Vendor"
        '
        'LenovoOption
        '
        Me.LenovoOption.AutoSize = True
        Me.LenovoOption.Location = New System.Drawing.Point(25, 42)
        Me.LenovoOption.Name = "LenovoOption"
        Me.LenovoOption.Size = New System.Drawing.Size(61, 17)
        Me.LenovoOption.TabIndex = 4
        Me.LenovoOption.TabStop = True
        Me.LenovoOption.Text = "Lenovo"
        Me.LenovoOption.UseVisualStyleBackColor = True
        '
        'btnMS
        '
        Me.btnMS.AutoSize = True
        Me.btnMS.Location = New System.Drawing.Point(118, 42)
        Me.btnMS.Name = "btnMS"
        Me.btnMS.Size = New System.Drawing.Size(96, 17)
        Me.btnMS.TabIndex = 3
        Me.btnMS.TabStop = True
        Me.btnMS.Text = "Microsoft (HW)"
        Me.btnMS.UseVisualStyleBackColor = True
        '
        'HPEOption
        '
        Me.HPEOption.AutoSize = True
        Me.HPEOption.Location = New System.Drawing.Point(118, 19)
        Me.HPEOption.Name = "HPEOption"
        Me.HPEOption.Size = New System.Drawing.Size(47, 17)
        Me.HPEOption.TabIndex = 2
        Me.HPEOption.TabStop = True
        Me.HPEOption.Text = "HPE"
        Me.HPEOption.UseVisualStyleBackColor = True
        '
        'DellOption
        '
        Me.DellOption.AutoSize = True
        Me.DellOption.Checked = True
        Me.DellOption.Location = New System.Drawing.Point(215, 19)
        Me.DellOption.Name = "DellOption"
        Me.DellOption.Size = New System.Drawing.Size(43, 17)
        Me.DellOption.TabIndex = 1
        Me.DellOption.TabStop = True
        Me.DellOption.Text = "Dell"
        Me.DellOption.UseVisualStyleBackColor = True
        '
        'HPIOption
        '
        Me.HPIOption.AutoSize = True
        Me.HPIOption.Location = New System.Drawing.Point(25, 19)
        Me.HPIOption.Name = "HPIOption"
        Me.HPIOption.Size = New System.Drawing.Size(43, 17)
        Me.HPIOption.TabIndex = 0
        Me.HPIOption.TabStop = True
        Me.HPIOption.Text = "HPI"
        Me.HPIOption.UseVisualStyleBackColor = True
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(12, 192)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(182, 13)
        Me.Label2.TabIndex = 4
        Me.Label2.Text = "Deal ID / SQ ID / Deal Reg Number:"
        '
        'DealID
        '
        Me.DealID.Location = New System.Drawing.Point(15, 208)
        Me.DealID.Name = "DealID"
        Me.DealID.Size = New System.Drawing.Size(287, 20)
        Me.DealID.TabIndex = 5
        '
        'OKButton
        '
        Me.OKButton.Location = New System.Drawing.Point(15, 234)
        Me.OKButton.Name = "OKButton"
        Me.OKButton.Size = New System.Drawing.Size(138, 29)
        Me.OKButton.TabIndex = 6
        Me.OKButton.Text = "OK"
        Me.OKButton.UseVisualStyleBackColor = True
        '
        'tCancelButton
        '
        Me.tCancelButton.Location = New System.Drawing.Point(167, 234)
        Me.tCancelButton.Name = "tCancelButton"
        Me.tCancelButton.Size = New System.Drawing.Size(138, 29)
        Me.tCancelButton.TabIndex = 7
        Me.tCancelButton.Text = "Cancel"
        Me.tCancelButton.UseVisualStyleBackColor = True
        '
        'GroupBox2
        '
        Me.GroupBox2.Controls.Add(Me.chkExertis)
        Me.GroupBox2.Controls.Add(Me.cWestcoast)
        Me.GroupBox2.Controls.Add(Me.cTechData)
        Me.GroupBox2.Controls.Add(Me.cIngram)
        Me.GroupBox2.Location = New System.Drawing.Point(13, 129)
        Me.GroupBox2.Name = "GroupBox2"
        Me.GroupBox2.Size = New System.Drawing.Size(288, 60)
        Me.GroupBox2.TabIndex = 9
        Me.GroupBox2.TabStop = False
        Me.GroupBox2.Text = "Distribution Partners"
        '
        'chkExertis
        '
        Me.chkExertis.AutoSize = True
        Me.chkExertis.Location = New System.Drawing.Point(154, 37)
        Me.chkExertis.Name = "chkExertis"
        Me.chkExertis.Size = New System.Drawing.Size(57, 17)
        Me.chkExertis.TabIndex = 3
        Me.chkExertis.Text = "Exertis"
        Me.chkExertis.UseVisualStyleBackColor = True
        '
        'cWestcoast
        '
        Me.cWestcoast.AutoSize = True
        Me.cWestcoast.Location = New System.Drawing.Point(8, 37)
        Me.cWestcoast.Name = "cWestcoast"
        Me.cWestcoast.Size = New System.Drawing.Size(77, 17)
        Me.cWestcoast.TabIndex = 2
        Me.cWestcoast.Text = "Westcoast"
        Me.cWestcoast.UseVisualStyleBackColor = True
        '
        'cTechData
        '
        Me.cTechData.AutoSize = True
        Me.cTechData.Location = New System.Drawing.Point(154, 14)
        Me.cTechData.Name = "cTechData"
        Me.cTechData.Size = New System.Drawing.Size(72, 17)
        Me.cTechData.TabIndex = 1
        Me.cTechData.Text = "Techdata"
        Me.cTechData.UseVisualStyleBackColor = True
        '
        'cIngram
        '
        Me.cIngram.AutoSize = True
        Me.cIngram.Location = New System.Drawing.Point(8, 15)
        Me.cIngram.Name = "cIngram"
        Me.cIngram.Size = New System.Drawing.Size(58, 17)
        Me.cIngram.TabIndex = 0
        Me.cIngram.Text = "Ingram"
        Me.cIngram.UseVisualStyleBackColor = True
        '
        'BackgroundWorker1
        '
        '
        'AddDeal
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(317, 301)
        Me.Controls.Add(Me.GroupBox2)
        Me.Controls.Add(Me.tCancelButton)
        Me.Controls.Add(Me.OKButton)
        Me.Controls.Add(Me.DealID)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.VendorGroupBox)
        Me.Controls.Add(Me.CustomerName)
        Me.Controls.Add(Me.Label1)
        Me.Name = "AddDeal"
        Me.Text = "Deal Information"
        Me.VendorGroupBox.ResumeLayout(False)
        Me.VendorGroupBox.PerformLayout()
        Me.GroupBox2.ResumeLayout(False)
        Me.GroupBox2.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents CustomerName As System.Windows.Forms.TextBox
    Friend WithEvents VendorGroupBox As System.Windows.Forms.GroupBox
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
    Friend WithEvents BackgroundWorker1 As ComponentModel.BackgroundWorker
    Friend WithEvents btnMS As Windows.Forms.RadioButton
    Friend WithEvents LenovoOption As Windows.Forms.RadioButton
    Friend WithEvents chkExertis As Windows.Forms.CheckBox
End Class
