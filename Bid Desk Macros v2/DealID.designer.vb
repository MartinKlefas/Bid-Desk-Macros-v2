<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class DealIdent
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
        Me.SpecialMsg = New System.Windows.Forms.Label()
        Me.DealID = New System.Windows.Forms.TextBox()
        Me.Button1 = New System.Windows.Forms.Button()
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
        'Button1
        '
        Me.Button1.Location = New System.Drawing.Point(15, 45)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(131, 25)
        Me.Button1.TabIndex = 2
        Me.Button1.Text = "OK"
        Me.Button1.UseVisualStyleBackColor = True
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
        Me.Controls.Add(Me.Button1)
        Me.Controls.Add(Me.DealID)
        Me.Controls.Add(Me.SpecialMsg)
        Me.Name = "DealIdent"
        Me.Text = "Form1"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents SpecialMsg As System.Windows.Forms.Label
    Friend WithEvents DealID As System.Windows.Forms.TextBox
    Friend WithEvents Button1 As System.Windows.Forms.Button
    Friend WithEvents Button2 As System.Windows.Forms.Button
End Class
