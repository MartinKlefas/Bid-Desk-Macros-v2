<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class SearchForm
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
        Me.GroupBox1 = New System.Windows.Forms.GroupBox()
        Me.SearchTerm = New System.Windows.Forms.TextBox()
        Me.ChkAM = New System.Windows.Forms.CheckBox()
        Me.ChkCustomer = New System.Windows.Forms.CheckBox()
        Me.ChkDeal = New System.Windows.Forms.CheckBox()
        Me.ChkOPG = New System.Windows.Forms.CheckBox()
        Me.CHKNDT = New System.Windows.Forms.CheckBox()
        Me.btnCancel = New System.Windows.Forms.Button()
        Me.BtnSearch = New System.Windows.Forms.Button()
        Me.GroupBox1.SuspendLayout()
        Me.SuspendLayout()
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(13, 13)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(62, 13)
        Me.Label1.TabIndex = 0
        Me.Label1.Text = "Search For:"
        '
        'GroupBox1
        '
        Me.GroupBox1.Controls.Add(Me.CHKNDT)
        Me.GroupBox1.Controls.Add(Me.ChkOPG)
        Me.GroupBox1.Controls.Add(Me.ChkDeal)
        Me.GroupBox1.Controls.Add(Me.ChkCustomer)
        Me.GroupBox1.Controls.Add(Me.ChkAM)
        Me.GroupBox1.Location = New System.Drawing.Point(13, 61)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(200, 142)
        Me.GroupBox1.TabIndex = 1
        Me.GroupBox1.TabStop = False
        Me.GroupBox1.Text = "In"
        '
        'SearchTerm
        '
        Me.SearchTerm.Location = New System.Drawing.Point(16, 30)
        Me.SearchTerm.Name = "SearchTerm"
        Me.SearchTerm.Size = New System.Drawing.Size(197, 20)
        Me.SearchTerm.TabIndex = 2
        '
        'ChkAM
        '
        Me.ChkAM.AutoSize = True
        Me.ChkAM.Location = New System.Drawing.Point(7, 20)
        Me.ChkAM.Name = "ChkAM"
        Me.ChkAM.Size = New System.Drawing.Size(42, 17)
        Me.ChkAM.TabIndex = 0
        Me.ChkAM.Text = "AM"
        Me.ChkAM.UseVisualStyleBackColor = True
        '
        'ChkCustomer
        '
        Me.ChkCustomer.AutoSize = True
        Me.ChkCustomer.Location = New System.Drawing.Point(7, 44)
        Me.ChkCustomer.Name = "ChkCustomer"
        Me.ChkCustomer.Size = New System.Drawing.Size(70, 17)
        Me.ChkCustomer.TabIndex = 1
        Me.ChkCustomer.Text = "Customer"
        Me.ChkCustomer.UseVisualStyleBackColor = True
        '
        'ChkDeal
        '
        Me.ChkDeal.AutoSize = True
        Me.ChkDeal.Location = New System.Drawing.Point(7, 68)
        Me.ChkDeal.Name = "ChkDeal"
        Me.ChkDeal.Size = New System.Drawing.Size(62, 17)
        Me.ChkDeal.TabIndex = 2
        Me.ChkDeal.Text = "Deal ID"
        Me.ChkDeal.UseVisualStyleBackColor = True
        '
        'ChkOPG
        '
        Me.ChkOPG.AutoSize = True
        Me.ChkOPG.Location = New System.Drawing.Point(7, 92)
        Me.ChkOPG.Name = "ChkOPG"
        Me.ChkOPG.Size = New System.Drawing.Size(63, 17)
        Me.ChkOPG.TabIndex = 3
        Me.ChkOPG.Text = "OPG ID"
        Me.ChkOPG.UseVisualStyleBackColor = True
        '
        'CHKNDT
        '
        Me.CHKNDT.AutoSize = True
        Me.CHKNDT.Location = New System.Drawing.Point(7, 116)
        Me.CHKNDT.Name = "CHKNDT"
        Me.CHKNDT.Size = New System.Drawing.Size(89, 17)
        Me.CHKNDT.TabIndex = 4
        Me.CHKNDT.Text = "NDT Number"
        Me.CHKNDT.UseVisualStyleBackColor = True
        '
        'btnCancel
        '
        Me.btnCancel.Location = New System.Drawing.Point(118, 210)
        Me.btnCancel.Name = "btnCancel"
        Me.btnCancel.Size = New System.Drawing.Size(96, 23)
        Me.btnCancel.TabIndex = 3
        Me.btnCancel.Text = "Cancel"
        Me.btnCancel.UseVisualStyleBackColor = True
        '
        'BtnSearch
        '
        Me.BtnSearch.Location = New System.Drawing.Point(13, 210)
        Me.BtnSearch.Name = "BtnSearch"
        Me.BtnSearch.Size = New System.Drawing.Size(96, 23)
        Me.BtnSearch.TabIndex = 4
        Me.BtnSearch.Text = "Search"
        Me.BtnSearch.UseVisualStyleBackColor = True
        '
        'SearchForm
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(226, 244)
        Me.Controls.Add(Me.BtnSearch)
        Me.Controls.Add(Me.btnCancel)
        Me.Controls.Add(Me.SearchTerm)
        Me.Controls.Add(Me.GroupBox1)
        Me.Controls.Add(Me.Label1)
        Me.Name = "SearchForm"
        Me.Text = "SearchForm"
        Me.GroupBox1.ResumeLayout(False)
        Me.GroupBox1.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents Label1 As Windows.Forms.Label
    Friend WithEvents GroupBox1 As Windows.Forms.GroupBox
    Friend WithEvents CHKNDT As Windows.Forms.CheckBox
    Friend WithEvents ChkOPG As Windows.Forms.CheckBox
    Friend WithEvents ChkDeal As Windows.Forms.CheckBox
    Friend WithEvents ChkCustomer As Windows.Forms.CheckBox
    Friend WithEvents ChkAM As Windows.Forms.CheckBox
    Friend WithEvents SearchTerm As Windows.Forms.TextBox
    Friend WithEvents btnCancel As Windows.Forms.Button
    Friend WithEvents BtnSearch As Windows.Forms.Button
End Class
