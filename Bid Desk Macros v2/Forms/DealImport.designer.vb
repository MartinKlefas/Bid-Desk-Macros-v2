<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()>
Partial Class ImportDeal
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
        Me.MSOption = New System.Windows.Forms.RadioButton()
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
        Me.NDTNumber = New System.Windows.Forms.TextBox()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.ChkTicket = New System.Windows.Forms.CheckBox()
        Me.AMMail = New System.Windows.Forms.TextBox()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.ccList = New System.Windows.Forms.TextBox()
        Me.Label5 = New System.Windows.Forms.Label()
        Me.VendorGroupBox.SuspendLayout()
        Me.GroupBox2.SuspendLayout()
        Me.SuspendLayout()
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(10, 86)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(85, 13)
        Me.Label1.TabIndex = 0
        Me.Label1.Text = "Customer Name:"
        '
        'CustomerName
        '
        Me.CustomerName.Location = New System.Drawing.Point(13, 102)
        Me.CustomerName.Name = "CustomerName"
        Me.CustomerName.Size = New System.Drawing.Size(290, 20)
        Me.CustomerName.TabIndex = 1
        '
        'VendorGroupBox
        '
        Me.VendorGroupBox.Controls.Add(Me.MSOption)
        Me.VendorGroupBox.Controls.Add(Me.HPEOption)
        Me.VendorGroupBox.Controls.Add(Me.DellOption)
        Me.VendorGroupBox.Controls.Add(Me.HPIOption)
        Me.VendorGroupBox.Location = New System.Drawing.Point(13, 128)
        Me.VendorGroupBox.Name = "VendorGroupBox"
        Me.VendorGroupBox.Size = New System.Drawing.Size(290, 76)
        Me.VendorGroupBox.TabIndex = 2
        Me.VendorGroupBox.TabStop = False
        Me.VendorGroupBox.Text = "Vendor"
        '
        'MSOption
        '
        Me.MSOption.AutoSize = True
        Me.MSOption.Location = New System.Drawing.Point(24, 42)
        Me.MSOption.Name = "MSOption"
        Me.MSOption.Size = New System.Drawing.Size(68, 17)
        Me.MSOption.TabIndex = 3
        Me.MSOption.Text = "Microsoft"
        Me.MSOption.UseVisualStyleBackColor = True
        '
        'HPEOption
        '
        Me.HPEOption.AutoSize = True
        Me.HPEOption.Location = New System.Drawing.Point(118, 19)
        Me.HPEOption.Name = "HPEOption"
        Me.HPEOption.Size = New System.Drawing.Size(47, 17)
        Me.HPEOption.TabIndex = 2
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
        Me.HPIOption.Text = "HPI"
        Me.HPIOption.UseVisualStyleBackColor = True
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(11, 273)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(182, 13)
        Me.Label2.TabIndex = 4
        Me.Label2.Text = "Deal ID / SQ ID / Deal Reg Number:"
        '
        'DealID
        '
        Me.DealID.Location = New System.Drawing.Point(14, 289)
        Me.DealID.Name = "DealID"
        Me.DealID.Size = New System.Drawing.Size(287, 20)
        Me.DealID.TabIndex = 5
        '
        'OKButton
        '
        Me.OKButton.Location = New System.Drawing.Point(11, 384)
        Me.OKButton.Name = "OKButton"
        Me.OKButton.Size = New System.Drawing.Size(138, 29)
        Me.OKButton.TabIndex = 6
        Me.OKButton.Text = "OK"
        Me.OKButton.UseVisualStyleBackColor = True
        '
        'tCancelButton
        '
        Me.tCancelButton.Location = New System.Drawing.Point(163, 384)
        Me.tCancelButton.Name = "tCancelButton"
        Me.tCancelButton.Size = New System.Drawing.Size(138, 29)
        Me.tCancelButton.TabIndex = 7
        Me.tCancelButton.Text = "Cancel"
        Me.tCancelButton.UseVisualStyleBackColor = True
        '
        'GroupBox2
        '
        Me.GroupBox2.Controls.Add(Me.cWestcoast)
        Me.GroupBox2.Controls.Add(Me.cTechData)
        Me.GroupBox2.Controls.Add(Me.cIngram)
        Me.GroupBox2.Location = New System.Drawing.Point(14, 210)
        Me.GroupBox2.Name = "GroupBox2"
        Me.GroupBox2.Size = New System.Drawing.Size(288, 60)
        Me.GroupBox2.TabIndex = 9
        Me.GroupBox2.TabStop = False
        Me.GroupBox2.Text = "Distribution Partners"
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
        'NDTNumber
        '
        Me.NDTNumber.Location = New System.Drawing.Point(14, 328)
        Me.NDTNumber.Name = "NDTNumber"
        Me.NDTNumber.Size = New System.Drawing.Size(287, 20)
        Me.NDTNumber.TabIndex = 11
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(11, 312)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(127, 13)
        Me.Label3.TabIndex = 10
        Me.Label3.Text = "NextDesk Ticket Number"
        '
        'ChkTicket
        '
        Me.ChkTicket.AutoSize = True
        Me.ChkTicket.Location = New System.Drawing.Point(14, 355)
        Me.ChkTicket.Name = "ChkTicket"
        Me.ChkTicket.Size = New System.Drawing.Size(121, 17)
        Me.ChkTicket.TabIndex = 12
        Me.ChkTicket.Text = "Add Details to ticket"
        Me.ChkTicket.UseVisualStyleBackColor = True
        '
        'AMMail
        '
        Me.AMMail.Location = New System.Drawing.Point(13, 25)
        Me.AMMail.Name = "AMMail"
        Me.AMMail.Size = New System.Drawing.Size(290, 20)
        Me.AMMail.TabIndex = 14
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Location = New System.Drawing.Point(12, 9)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(160, 13)
        Me.Label4.TabIndex = 13
        Me.Label4.Text = "Account Manager Email or Alias:"
        '
        'ccList
        '
        Me.ccList.Location = New System.Drawing.Point(11, 64)
        Me.ccList.Name = "ccList"
        Me.ccList.Size = New System.Drawing.Size(290, 20)
        Me.ccList.TabIndex = 16
        '
        'Label5
        '
        Me.Label5.AutoSize = True
        Me.Label5.Location = New System.Drawing.Point(10, 48)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(158, 13)
        Me.Label5.TabIndex = 15
        Me.Label5.Text = "Extra People to CC (; separated)"
        '
        'ImportDeal
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(317, 425)
        Me.Controls.Add(Me.ccList)
        Me.Controls.Add(Me.Label5)
        Me.Controls.Add(Me.AMMail)
        Me.Controls.Add(Me.Label4)
        Me.Controls.Add(Me.ChkTicket)
        Me.Controls.Add(Me.NDTNumber)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.GroupBox2)
        Me.Controls.Add(Me.tCancelButton)
        Me.Controls.Add(Me.OKButton)
        Me.Controls.Add(Me.DealID)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.VendorGroupBox)
        Me.Controls.Add(Me.CustomerName)
        Me.Controls.Add(Me.Label1)
        Me.Name = "ImportDeal"
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
    Friend WithEvents NDTNumber As Windows.Forms.TextBox
    Friend WithEvents Label3 As Windows.Forms.Label
    Friend WithEvents ChkTicket As Windows.Forms.CheckBox
    Friend WithEvents AMMail As Windows.Forms.TextBox
    Friend WithEvents Label4 As Windows.Forms.Label
    Friend WithEvents ccList As Windows.Forms.TextBox
    Friend WithEvents Label5 As Windows.Forms.Label
    Friend WithEvents MSOption As Windows.Forms.RadioButton
End Class
