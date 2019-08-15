<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class TicketActions
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
        Me.CnclButton = New System.Windows.Forms.Button()
        Me.OKButton = New System.Windows.Forms.Button()
        Me.NDTNum = New System.Windows.Forms.TextBox()
        Me.SpecialMsg = New System.Windows.Forms.Label()
        Me.BackgroundWorker1 = New System.ComponentModel.BackgroundWorker()
        Me.SuspendLayout()
        '
        'CnclButton
        '
        Me.CnclButton.Location = New System.Drawing.Point(172, 51)
        Me.CnclButton.Name = "CnclButton"
        Me.CnclButton.Size = New System.Drawing.Size(131, 25)
        Me.CnclButton.TabIndex = 7
        Me.CnclButton.Text = "Cancel"
        Me.CnclButton.UseVisualStyleBackColor = True
        '
        'OKButton
        '
        Me.OKButton.Location = New System.Drawing.Point(15, 51)
        Me.OKButton.Name = "OKButton"
        Me.OKButton.Size = New System.Drawing.Size(131, 25)
        Me.OKButton.TabIndex = 6
        Me.OKButton.Text = "OK"
        Me.OKButton.UseVisualStyleBackColor = True
        '
        'NDTNum
        '
        Me.NDTNum.Location = New System.Drawing.Point(12, 25)
        Me.NDTNum.Name = "NDTNum"
        Me.NDTNum.Size = New System.Drawing.Size(291, 20)
        Me.NDTNum.TabIndex = 5
        '
        'SpecialMsg
        '
        Me.SpecialMsg.AutoSize = True
        Me.SpecialMsg.Location = New System.Drawing.Point(12, 9)
        Me.SpecialMsg.Name = "SpecialMsg"
        Me.SpecialMsg.Size = New System.Drawing.Size(130, 13)
        Me.SpecialMsg.TabIndex = 4
        Me.SpecialMsg.Text = "NextDesk Ticket Number:"
        '
        'BackgroundWorker1
        '
        '
        'TicketActions
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(312, 85)
        Me.Controls.Add(Me.CnclButton)
        Me.Controls.Add(Me.OKButton)
        Me.Controls.Add(Me.NDTNum)
        Me.Controls.Add(Me.SpecialMsg)
        Me.Name = "TicketActions"
        Me.Text = "TicketActions"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents CnclButton As Windows.Forms.Button
    Friend WithEvents OKButton As Windows.Forms.Button
    Friend WithEvents NDTNum As Windows.Forms.TextBox
    Friend WithEvents SpecialMsg As Windows.Forms.Label
    Friend WithEvents BackgroundWorker1 As ComponentModel.BackgroundWorker
End Class
