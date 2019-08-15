<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()>
Partial Class DealIdent
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
        Me.SpecialMsg = New System.Windows.Forms.Label()
        Me.DealID = New System.Windows.Forms.TextBox()
        Me.OKButton = New System.Windows.Forms.Button()
        Me.Button2 = New System.Windows.Forms.Button()
        Me.SuspendLayout()
        '
        'SpecialMsg
        '
        Me.SpecialMsg.AutoSize = True
        Me.SpecialMsg.Location = New System.Drawing.Point(12, 3)
        Me.SpecialMsg.Name = "SpecialMsg"
        Me.SpecialMsg.Size = New System.Drawing.Size(182, 13)
        Me.SpecialMsg.TabIndex = 0
        Me.SpecialMsg.Text = "Deal ID / SQ ID / Deal Reg Number:"
        '
        'DealID
        '
        Me.DealID.Location = New System.Drawing.Point(12, 19)
        Me.DealID.Name = "DealID"
        Me.DealID.Size = New System.Drawing.Size(291, 20)
        Me.DealID.TabIndex = 1
        '
        'OKButton
        '
        Me.OKButton.Location = New System.Drawing.Point(15, 45)
        Me.OKButton.Name = "OKButton"
        Me.OKButton.Size = New System.Drawing.Size(131, 25)
        Me.OKButton.TabIndex = 2
        Me.OKButton.Text = "OK"
        Me.OKButton.UseVisualStyleBackColor = True
        '
        'Button2
        '
        Me.Button2.Location = New System.Drawing.Point(172, 45)
        Me.Button2.Name = "Button2"
        Me.Button2.Size = New System.Drawing.Size(131, 25)
        Me.Button2.TabIndex = 3
        Me.Button2.Text = "Cancel"
        Me.Button2.UseVisualStyleBackColor = True
        '
        'DealIdent
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(317, 79)
        Me.Controls.Add(Me.Button2)
        Me.Controls.Add(Me.OKButton)
        Me.Controls.Add(Me.DealID)
        Me.Controls.Add(Me.SpecialMsg)
        Me.Name = "DealIdent"
        Me.Text = "Deal ID Finder"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents SpecialMsg As System.Windows.Forms.Label
    Friend WithEvents DealID As System.Windows.Forms.TextBox
    Friend WithEvents OKButton As System.Windows.Forms.Button
    Friend WithEvents Button2 As System.Windows.Forms.Button
End Class
