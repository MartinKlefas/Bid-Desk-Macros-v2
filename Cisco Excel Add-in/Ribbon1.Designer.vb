Partial Class Ribbon1
    Inherits Microsoft.Office.Tools.Ribbon.RibbonBase

    <System.Diagnostics.DebuggerNonUserCode()>
    Public Sub New(ByVal container As System.ComponentModel.IContainer)
        MyClass.New()

        'Required for Windows.Forms Class Composition Designer support
        If (container IsNot Nothing) Then
            container.Add(Me)
        End If

    End Sub

    <System.Diagnostics.DebuggerNonUserCode()>
    Public Sub New()
        MyBase.New(Globals.Factory.GetRibbonFactory())

        'This call is required by the Component Designer.
        InitializeComponent()

    End Sub

    'Component overrides dispose to clean up the component list.
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

    'Required by the Component Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Component Designer
    'It can be modified using the Component Designer.
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()>
    Private Sub InitializeComponent()
        Me.Tab1 = Me.Factory.CreateRibbonTab
        Me.Cisco = Me.Factory.CreateRibbonGroup
        Me.BtnLogin = Me.Factory.CreateRibbonButton
        Me.NewDeal = Me.Factory.CreateRibbonButton
        Me.BtnDLDeal = Me.Factory.CreateRibbonButton
        Me.Button1 = Me.Factory.CreateRibbonButton
        Me.Group1 = Me.Factory.CreateRibbonGroup
        Me.Button2 = Me.Factory.CreateRibbonButton
        Me.Button3 = Me.Factory.CreateRibbonButton
        Me.Button4 = Me.Factory.CreateRibbonButton
        Me.Group2 = Me.Factory.CreateRibbonGroup
        Me.Button5 = Me.Factory.CreateRibbonButton
        Me.Button6 = Me.Factory.CreateRibbonButton
        Me.Tab1.SuspendLayout()
        Me.Cisco.SuspendLayout()
        Me.Group1.SuspendLayout()
        Me.Group2.SuspendLayout()
        Me.SuspendLayout()
        '
        'Tab1
        '
        Me.Tab1.Groups.Add(Me.Cisco)
        Me.Tab1.Groups.Add(Me.Group1)
        Me.Tab1.Groups.Add(Me.Group2)
        Me.Tab1.Label = "Browser Controller Tools"
        Me.Tab1.Name = "Tab1"
        '
        'Cisco
        '
        Me.Cisco.Items.Add(Me.BtnLogin)
        Me.Cisco.Items.Add(Me.NewDeal)
        Me.Cisco.Items.Add(Me.BtnDLDeal)
        Me.Cisco.Items.Add(Me.Button1)
        Me.Cisco.Label = "Cisco"
        Me.Cisco.Name = "Cisco"
        '
        'BtnLogin
        '
        Me.BtnLogin.Label = "Log In"
        Me.BtnLogin.Name = "BtnLogin"
        '
        'NewDeal
        '
        Me.NewDeal.Label = "CreateDeal"
        Me.NewDeal.Name = "NewDeal"
        '
        'BtnDLDeal
        '
        Me.BtnDLDeal.Label = "Download Quote"
        Me.BtnDLDeal.Name = "BtnDLDeal"
        '
        'Button1
        '
        Me.Button1.Label = "Get Cisco AM Details"
        Me.Button1.Name = "Button1"
        '
        'Group1
        '
        Me.Group1.Items.Add(Me.Button2)
        Me.Group1.Items.Add(Me.Button3)
        Me.Group1.Items.Add(Me.Button4)
        Me.Group1.Label = "Lenovo"
        Me.Group1.Name = "Group1"
        '
        'Button2
        '
        Me.Button2.Label = "Log in"
        Me.Button2.Name = "Button2"
        '
        'Button3
        '
        Me.Button3.Label = "Show Deal"
        Me.Button3.Name = "Button3"
        '
        'Button4
        '
        Me.Button4.Label = "Send to Disti"
        Me.Button4.Name = "Button4"
        '
        'Group2
        '
        Me.Group2.Items.Add(Me.Button5)
        Me.Group2.Items.Add(Me.Button6)
        Me.Group2.Label = "Test Functions"
        Me.Group2.Name = "Group2"
        '
        'Button5
        '
        Me.Button5.Label = "Table to Excel file"
        Me.Button5.Name = "Button5"
        '
        'Button6
        '
        Me.Button6.Label = "Test String Trimmer"
        Me.Button6.Name = "Button6"
        '
        'Ribbon1
        '
        Me.Name = "Ribbon1"
        Me.RibbonType = "Microsoft.Excel.Workbook"
        Me.Tabs.Add(Me.Tab1)
        Me.Tab1.ResumeLayout(False)
        Me.Tab1.PerformLayout()
        Me.Cisco.ResumeLayout(False)
        Me.Cisco.PerformLayout()
        Me.Group1.ResumeLayout(False)
        Me.Group1.PerformLayout()
        Me.Group2.ResumeLayout(False)
        Me.Group2.PerformLayout()
        Me.ResumeLayout(False)

    End Sub

    Friend WithEvents Tab1 As Microsoft.Office.Tools.Ribbon.RibbonTab
    Friend WithEvents Cisco As Microsoft.Office.Tools.Ribbon.RibbonGroup
    Friend WithEvents BtnLogin As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents NewDeal As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents BtnDLDeal As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents Button1 As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents Group1 As Microsoft.Office.Tools.Ribbon.RibbonGroup
    Friend WithEvents Button2 As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents Button3 As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents Button4 As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents Group2 As Microsoft.Office.Tools.Ribbon.RibbonGroup
    Friend WithEvents Button5 As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents Button6 As Microsoft.Office.Tools.Ribbon.RibbonButton
End Class

Partial Class ThisRibbonCollection

    <System.Diagnostics.DebuggerNonUserCode()> _
    Friend ReadOnly Property Ribbon1() As Ribbon1
        Get
            Return Me.GetRibbon(Of Ribbon1)()
        End Get
    End Property
End Class
