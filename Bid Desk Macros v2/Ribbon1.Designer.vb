﻿Partial Class Ribbon1
    Inherits Microsoft.Office.Tools.Ribbon.RibbonBase

    <System.Diagnostics.DebuggerNonUserCode()> _
    Public Sub New(ByVal container As System.ComponentModel.IContainer)
        MyClass.New()

        'Required for Windows.Forms Class Composition Designer support
        If (container IsNot Nothing) Then
            container.Add(Me)
        End If

    End Sub

    <System.Diagnostics.DebuggerNonUserCode()> _
    Public Sub New()
        MyBase.New(Globals.Factory.GetRibbonFactory())

        'This call is required by the Component Designer.
        InitializeComponent()

    End Sub

    'Component overrides dispose to clean up the component list.
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

    'Required by the Component Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Component Designer
    'It can be modified using the Component Designer.
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Me.Tab1 = Me.Factory.CreateRibbonTab
        Me.Group1 = Me.Factory.CreateRibbonGroup
        Me.Button1 = Me.Factory.CreateRibbonButton
        Me.MoveBtn = Me.Factory.CreateRibbonButton
        Me.ReplyToBidBtn = Me.Factory.CreateRibbonButton
        Me.FwdPrice = Me.Factory.CreateRibbonButton
        Me.HPFwd = Me.Factory.CreateRibbonButton
        Me.FwdDecision = Me.Factory.CreateRibbonButton
        Me.ExpireButton = Me.Factory.CreateRibbonButton
        Me.ExtensionBtn = Me.Factory.CreateRibbonButton
        Me.Tab1.SuspendLayout()
        Me.Group1.SuspendLayout()
        Me.SuspendLayout()
        '
        'Tab1
        '
        Me.Tab1.Groups.Add(Me.Group1)
        Me.Tab1.Label = "Bid Tools"
        Me.Tab1.Name = "Tab1"
        '
        'Group1
        '
        Me.Group1.Items.Add(Me.Button1)
        Me.Group1.Items.Add(Me.MoveBtn)
        Me.Group1.Items.Add(Me.ReplyToBidBtn)
        Me.Group1.Items.Add(Me.FwdPrice)
        Me.Group1.Items.Add(Me.HPFwd)
        Me.Group1.Items.Add(Me.FwdDecision)
        Me.Group1.Items.Add(Me.ExpireButton)
        Me.Group1.Items.Add(Me.ExtensionBtn)
        Me.Group1.Label = "Group1"
        Me.Group1.Name = "Group1"
        '
        'Button1
        '
        Me.Button1.Label = "Button1"
        Me.Button1.Name = "Button1"
        '
        'MoveBtn
        '
        Me.MoveBtn.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge
        Me.MoveBtn.Image = Global.Bid_Desk_Macros_v2.My.Resources.Resources.folder_download_down_decrease_arrow
        Me.MoveBtn.Label = "Move by Deal ID"
        Me.MoveBtn.Name = "MoveBtn"
        Me.MoveBtn.ShowImage = True
        '
        'ReplyToBidBtn
        '
        Me.ReplyToBidBtn.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge
        Me.ReplyToBidBtn.Image = Global.Bid_Desk_Macros_v2.My.Resources.Resources._128__1_
        Me.ReplyToBidBtn.Label = "Reply to Bid Request"
        Me.ReplyToBidBtn.Name = "ReplyToBidBtn"
        Me.ReplyToBidBtn.ShowImage = True
        '
        'FwdPrice
        '
        Me.FwdPrice.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge
        Me.FwdPrice.Image = Global.Bid_Desk_Macros_v2.My.Resources.Resources.right_arrow
        Me.FwdPrice.Label = "Forward Pricing"
        Me.FwdPrice.Name = "FwdPrice"
        Me.FwdPrice.ShowImage = True
        '
        'HPFwd
        '
        Me.HPFwd.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge
        Me.HPFwd.Image = Global.Bid_Desk_Macros_v2.My.Resources.Resources._128
        Me.HPFwd.Label = "Forward HP Response"
        Me.HPFwd.Name = "HPFwd"
        Me.HPFwd.ShowImage = True
        '
        'FwdDecision
        '
        Me.FwdDecision.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge
        Me.FwdDecision.Image = Global.Bid_Desk_Macros_v2.My.Resources.Resources.money_dollar_coins_cash_piggy_bank_finance_business
        Me.FwdDecision.Label = "Forward DR Decision"
        Me.FwdDecision.Name = "FwdDecision"
        Me.FwdDecision.ShowImage = True
        '
        'ExpireButton
        '
        Me.ExpireButton.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge
        Me.ExpireButton.Image = Global.Bid_Desk_Macros_v2.My.Resources.Resources.exclamation_mark_red
        Me.ExpireButton.Label = "Create Expiry Messages"
        Me.ExpireButton.Name = "ExpireButton"
        Me.ExpireButton.ShowImage = True
        '
        'ExtensionBtn
        '
        Me.ExtensionBtn.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge
        Me.ExtensionBtn.Image = Global.Bid_Desk_Macros_v2.My.Resources.Resources.actions_view_calendar_tasks
        Me.ExtensionBtn.Label = "Send Extension Message"
        Me.ExtensionBtn.Name = "ExtensionBtn"
        Me.ExtensionBtn.ShowImage = True
        '
        'Ribbon1
        '
        Me.Name = "Ribbon1"
        Me.RibbonType = "Microsoft.Outlook.Explorer"
        Me.Tabs.Add(Me.Tab1)
        Me.Tab1.ResumeLayout(False)
        Me.Tab1.PerformLayout()
        Me.Group1.ResumeLayout(False)
        Me.Group1.PerformLayout()
        Me.ResumeLayout(False)

    End Sub

    Friend WithEvents Tab1 As Microsoft.Office.Tools.Ribbon.RibbonTab
    Friend WithEvents Group1 As Microsoft.Office.Tools.Ribbon.RibbonGroup
    Friend WithEvents Button1 As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents MoveBtn As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents ReplyToBidBtn As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents ExpireButton As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents FwdDecision As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents FwdPrice As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents HPFwd As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents ExtensionBtn As Microsoft.Office.Tools.Ribbon.RibbonButton
End Class

Partial Class ThisRibbonCollection

    <System.Diagnostics.DebuggerNonUserCode()> _
    Friend ReadOnly Property Ribbon1() As Ribbon1
        Get
            Return Me.GetRibbon(Of Ribbon1)()
        End Get
    End Property
End Class
