<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class CloneLater
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
        Me.targetDate = New System.Windows.Forms.MonthCalendar()
        Me.btnSetReminder = New System.Windows.Forms.Button()
        Me.Button2 = New System.Windows.Forms.Button()
        Me.SuspendLayout()
        '
        'targetDate
        '
        Me.targetDate.Location = New System.Drawing.Point(9, 6)
        Me.targetDate.MaxSelectionCount = 1
        Me.targetDate.MinDate = New Date(2019, 5, 17, 0, 0, 0, 0)
        Me.targetDate.Name = "targetDate"
        Me.targetDate.TabIndex = 1
        '
        'btnSetReminder
        '
        Me.btnSetReminder.Location = New System.Drawing.Point(9, 176)
        Me.btnSetReminder.Name = "btnSetReminder"
        Me.btnSetReminder.Size = New System.Drawing.Size(111, 23)
        Me.btnSetReminder.TabIndex = 2
        Me.btnSetReminder.Text = "Set Reminder"
        Me.btnSetReminder.UseVisualStyleBackColor = True
        '
        'Button2
        '
        Me.Button2.Location = New System.Drawing.Point(125, 176)
        Me.Button2.Name = "Button2"
        Me.Button2.Size = New System.Drawing.Size(111, 23)
        Me.Button2.TabIndex = 3
        Me.Button2.Text = "Cancel"
        Me.Button2.UseVisualStyleBackColor = True
        '
        'CloneLater
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(247, 211)
        Me.Controls.Add(Me.Button2)
        Me.Controls.Add(Me.btnSetReminder)
        Me.Controls.Add(Me.targetDate)
        Me.Name = "CloneLater"
        Me.Text = "CloneLater"
        Me.ResumeLayout(False)

    End Sub

    Friend WithEvents targetDate As Windows.Forms.MonthCalendar
    Friend WithEvents btnSetReminder As Windows.Forms.Button
    Friend WithEvents Button2 As Windows.Forms.Button
End Class
