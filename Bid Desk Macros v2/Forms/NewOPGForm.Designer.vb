<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class NewOPGForm
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
        Me.DealID = New System.Windows.Forms.TextBox()
        Me.SpecialMsg = New System.Windows.Forms.Label()
        Me.OPGBox = New System.Windows.Forms.TextBox()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.Button2 = New System.Windows.Forms.Button()
        Me.OKButton = New System.Windows.Forms.Button()
        Me.SuspendLayout()
        '
        'DealID
        '
        Me.DealID.Location = New System.Drawing.Point(12, 25)
        Me.DealID.Name = "DealID"
        Me.DealID.Size = New System.Drawing.Size(291, 20)
        Me.DealID.TabIndex = 3
        '
        'SpecialMsg
        '
        Me.SpecialMsg.AutoSize = True
        Me.SpecialMsg.Location = New System.Drawing.Point(12, 9)
        Me.SpecialMsg.Name = "SpecialMsg"
        Me.SpecialMsg.Size = New System.Drawing.Size(182, 13)
        Me.SpecialMsg.TabIndex = 2
        Me.SpecialMsg.Text = "Deal ID / SQ ID / Deal Reg Number:"
        '
        'OPGBox
        '
        Me.OPGBox.Location = New System.Drawing.Point(12, 64)
        Me.OPGBox.Name = "OPGBox"
        Me.OPGBox.Size = New System.Drawing.Size(291, 20)
        Me.OPGBox.TabIndex = 5
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(12, 48)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(149, 13)
        Me.Label1.TabIndex = 4
        Me.Label1.Text = "OPG ID / Secondary Identifier"
        '
        'Button2
        '
        Me.Button2.Location = New System.Drawing.Point(172, 90)
        Me.Button2.Name = "Button2"
        Me.Button2.Size = New System.Drawing.Size(131, 25)
        Me.Button2.TabIndex = 7
        Me.Button2.Text = "Cancel"
        Me.Button2.UseVisualStyleBackColor = True
        '
        'OKButton
        '
        Me.OKButton.Location = New System.Drawing.Point(15, 90)
        Me.OKButton.Name = "OKButton"
        Me.OKButton.Size = New System.Drawing.Size(131, 25)
        Me.OKButton.TabIndex = 6
        Me.OKButton.Text = "OK"
        Me.OKButton.UseVisualStyleBackColor = True
        '
        'NewOPGForm
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(316, 125)
        Me.Controls.Add(Me.Button2)
        Me.Controls.Add(Me.OKButton)
        Me.Controls.Add(Me.OPGBox)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.DealID)
        Me.Controls.Add(Me.SpecialMsg)
        Me.Name = "NewOPGForm"
        Me.Text = "NewOPGForm"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents DealID As Windows.Forms.TextBox
    Friend WithEvents SpecialMsg As Windows.Forms.Label
    Friend WithEvents OPGBox As Windows.Forms.TextBox
    Friend WithEvents Label1 As Windows.Forms.Label
    Friend WithEvents Button2 As Windows.Forms.Button
    Friend WithEvents OKButton As Windows.Forms.Button
End Class
