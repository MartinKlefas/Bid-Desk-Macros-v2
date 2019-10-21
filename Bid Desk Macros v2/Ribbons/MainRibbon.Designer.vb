Partial Class MainRibbon
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
        Dim Tab2 As Microsoft.Office.Tools.Ribbon.RibbonTab
        Me.Group5 = Me.Factory.CreateRibbonGroup
        Me.Tab1 = Me.Factory.CreateRibbonTab
        Me.Group1 = Me.Factory.CreateRibbonGroup
        Me.Group3 = Me.Factory.CreateRibbonGroup
        Me.Group2 = Me.Factory.CreateRibbonGroup
        Me.Group4 = Me.Factory.CreateRibbonGroup
        Me.Button2 = Me.Factory.CreateRibbonButton
        Me.btnAutoAll = Me.Factory.CreateRibbonButton
        Me.ReplyToBidBtn = Me.Factory.CreateRibbonButton
        Me.MoveBtn = Me.Factory.CreateRibbonButton
        Me.MvAttach = Me.Factory.CreateRibbonButton
        Me.FwdPrice = Me.Factory.CreateRibbonButton
        Me.HPFwd = Me.Factory.CreateRibbonButton
        Me.FwdDecision = Me.Factory.CreateRibbonButton
        Me.ExpireButton = Me.Factory.CreateRibbonButton
        Me.ExtensionBtn = Me.Factory.CreateRibbonButton
        Me.WonBtn = Me.Factory.CreateRibbonButton
        Me.DeadBtn = Me.Factory.CreateRibbonButton
        Me.BtnLater = Me.Factory.CreateRibbonButton
        Me.btnOnOff = Me.Factory.CreateRibbonButton
        Me.btnHoliday = Me.Factory.CreateRibbonButton
        Me.BtnAddtoDB = Me.Factory.CreateRibbonButton
        Me.ImprtLots = Me.Factory.CreateRibbonButton
        Me.addOPG = Me.Factory.CreateRibbonButton
        Me.btnLookup = Me.Factory.CreateRibbonButton
        Me.btnChangeAM = Me.Factory.CreateRibbonButton
        Me.Button1 = Me.Factory.CreateRibbonButton
        Me.BtnAutoAll_TabMail = Me.Factory.CreateRibbonButton
        Me.UnSortedMails = Me.Factory.CreateRibbonButton
        Me.UnsortedMails2 = Me.Factory.CreateRibbonButton
        Tab2 = Me.Factory.CreateRibbonTab
        Tab2.SuspendLayout()
        Me.Group5.SuspendLayout()
        Me.Tab1.SuspendLayout()
        Me.Group1.SuspendLayout()
        Me.Group3.SuspendLayout()
        Me.Group2.SuspendLayout()
        Me.Group4.SuspendLayout()
        Me.SuspendLayout()
        '
        'Tab2
        '
        Tab2.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office
        Tab2.ControlId.OfficeId = "TabMail"
        Tab2.Groups.Add(Me.Group5)
        Tab2.Label = "TabMail"
        Tab2.Name = "Tab2"
        '
        'Group5
        '
        Me.Group5.Items.Add(Me.BtnAutoAll_TabMail)
        Me.Group5.Items.Add(Me.UnSortedMails)
        Me.Group5.Label = "Bid Tools"
        Me.Group5.Name = "Group5"
        '
        'Tab1
        '
        Me.Tab1.Groups.Add(Me.Group1)
        Me.Tab1.Groups.Add(Me.Group3)
        Me.Tab1.Groups.Add(Me.Group2)
        Me.Tab1.Groups.Add(Me.Group4)
        Me.Tab1.Label = "Bid Tools"
        Me.Tab1.Name = "Tab1"
        '
        'Group1
        '
        Me.Group1.Items.Add(Me.btnAutoAll)
        Me.Group1.Items.Add(Me.UnsortedMails2)
        Me.Group1.Items.Add(Me.ReplyToBidBtn)
        Me.Group1.Items.Add(Me.MoveBtn)
        Me.Group1.Items.Add(Me.MvAttach)
        Me.Group1.Items.Add(Me.FwdPrice)
        Me.Group1.Items.Add(Me.HPFwd)
        Me.Group1.Items.Add(Me.FwdDecision)
        Me.Group1.Items.Add(Me.ExpireButton)
        Me.Group1.Items.Add(Me.ExtensionBtn)
        Me.Group1.Items.Add(Me.WonBtn)
        Me.Group1.Items.Add(Me.DeadBtn)
        Me.Group1.Items.Add(Me.BtnLater)
        Me.Group1.Label = "Processing Actions"
        Me.Group1.Name = "Group1"
        '
        'Group3
        '
        Me.Group3.Items.Add(Me.btnOnOff)
        Me.Group3.Items.Add(Me.btnHoliday)
        Me.Group3.Label = "Automation"
        Me.Group3.Name = "Group3"
        '
        'Group2
        '
        Me.Group2.Items.Add(Me.BtnAddtoDB)
        Me.Group2.Items.Add(Me.ImprtLots)
        Me.Group2.Items.Add(Me.addOPG)
        Me.Group2.Items.Add(Me.btnLookup)
        Me.Group2.Items.Add(Me.btnChangeAM)
        Me.Group2.Label = "Database Actions"
        Me.Group2.Name = "Group2"
        '
        'Group4
        '
        Me.Group4.Items.Add(Me.Button1)
        Me.Group4.Items.Add(Me.Button2)
        Me.Group4.Label = "Ticket Tools"
        Me.Group4.Name = "Group4"
        '
        'Button2
        '
        Me.Button2.Label = "Button2"
        Me.Button2.Name = "Button2"
        '
        'btnAutoAll
        '
        Me.btnAutoAll.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge
        Me.btnAutoAll.Image = Global.Bid_Desk_Macros_v2.My.Resources.Resources.robot_pngrepo_com
        Me.btnAutoAll.Label = "Auto Process"
        Me.btnAutoAll.Name = "btnAutoAll"
        Me.btnAutoAll.ShowImage = True
        '
        'ReplyToBidBtn
        '
        Me.ReplyToBidBtn.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge
        Me.ReplyToBidBtn.Image = Global.Bid_Desk_Macros_v2.My.Resources.Resources._128__1_
        Me.ReplyToBidBtn.Label = "Reply to Bid Request"
        Me.ReplyToBidBtn.Name = "ReplyToBidBtn"
        Me.ReplyToBidBtn.ShowImage = True
        '
        'MoveBtn
        '
        Me.MoveBtn.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge
        Me.MoveBtn.Image = Global.Bid_Desk_Macros_v2.My.Resources.Resources.folder_download_down_decrease_arrow
        Me.MoveBtn.Label = "Move by Deal ID"
        Me.MoveBtn.Name = "MoveBtn"
        Me.MoveBtn.ShowImage = True
        '
        'MvAttach
        '
        Me.MvAttach.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge
        Me.MvAttach.Image = Global.Bid_Desk_Macros_v2.My.Resources.Resources.output_onlinepngtools
        Me.MvAttach.Label = "Move && Attach"
        Me.MvAttach.Name = "MvAttach"
        Me.MvAttach.ShowImage = True
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
        Me.ExtensionBtn.Label = "Extension Requested"
        Me.ExtensionBtn.Name = "ExtensionBtn"
        Me.ExtensionBtn.ShowImage = True
        '
        'WonBtn
        '
        Me.WonBtn.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge
        Me.WonBtn.Image = Global.Bid_Desk_Macros_v2.My.Resources.Resources.smiley_devilish_face
        Me.WonBtn.Label = "Mark Won"
        Me.WonBtn.Name = "WonBtn"
        Me.WonBtn.ShowImage = True
        '
        'DeadBtn
        '
        Me.DeadBtn.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge
        Me.DeadBtn.Image = Global.Bid_Desk_Macros_v2.My.Resources.Resources.face_sad_smiley
        Me.DeadBtn.Label = "Mark Dead"
        Me.DeadBtn.Name = "DeadBtn"
        Me.DeadBtn.ShowImage = True
        '
        'BtnLater
        '
        Me.BtnLater.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge
        Me.BtnLater.Image = Global.Bid_Desk_Macros_v2.My.Resources.Resources.kissclipart_calendar_icon_orange_png_clipart_computer_icons_133adbff6cfb7003
        Me.BtnLater.Label = "Clone Later"
        Me.BtnLater.Name = "BtnLater"
        Me.BtnLater.ShowImage = True
        '
        'btnOnOff
        '
        Me.btnOnOff.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge
        Me.btnOnOff.Image = Global.Bid_Desk_Macros_v2.My.Resources.Resources.off
        Me.btnOnOff.Label = "Automation Off"
        Me.btnOnOff.Name = "btnOnOff"
        Me.btnOnOff.ShowImage = True
        '
        'btnHoliday
        '
        Me.btnHoliday.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge
        Me.btnHoliday.Image = Global.Bid_Desk_Macros_v2.My.Resources.Resources.OfficeWork_Icon
        Me.btnHoliday.Label = "At Work"
        Me.btnHoliday.Name = "btnHoliday"
        Me.btnHoliday.ShowImage = True
        '
        'BtnAddtoDB
        '
        Me.BtnAddtoDB.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge
        Me.BtnAddtoDB.Image = Global.Bid_Desk_Macros_v2.My.Resources.Resources.database_add_insert_21836
        Me.BtnAddtoDB.Label = "Add Single Deal"
        Me.BtnAddtoDB.Name = "BtnAddtoDB"
        Me.BtnAddtoDB.ShowImage = True
        '
        'ImprtLots
        '
        Me.ImprtLots.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge
        Me.ImprtLots.Image = Global.Bid_Desk_Macros_v2.My.Resources.Resources._0074af6a9c
        Me.ImprtLots.Label = "Add Multiple Deals"
        Me.ImprtLots.Name = "ImprtLots"
        Me.ImprtLots.ShowImage = True
        '
        'addOPG
        '
        Me.addOPG.Image = Global.Bid_Desk_Macros_v2.My.Resources.Resources.database_add_insert_21836
        Me.addOPG.Label = "Add OPG"
        Me.addOPG.Name = "addOPG"
        Me.addOPG.ShowImage = True
        '
        'btnLookup
        '
        Me.btnLookup.Image = Global.Bid_Desk_Macros_v2.My.Resources.Resources.Magnifying_Glass
        Me.btnLookup.Label = "Lookup Deal"
        Me.btnLookup.Name = "btnLookup"
        Me.btnLookup.ShowImage = True
        '
        'btnChangeAM
        '
        Me.btnChangeAM.Image = Global.Bid_Desk_Macros_v2.My.Resources.Resources.shift_change_pngrepo_com
        Me.btnChangeAM.Label = "Replace AM"
        Me.btnChangeAM.Name = "btnChangeAM"
        Me.btnChangeAM.ShowImage = True
        '
        'Button1
        '
        Me.Button1.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge
        Me.Button1.Image = Global.Bid_Desk_Macros_v2.My.Resources.Resources.question_mark
        Me.Button1.Label = "More Info (MS)"
        Me.Button1.Name = "Button1"
        Me.Button1.ShowImage = True
        '
        'BtnAutoAll_TabMail
        '
        Me.BtnAutoAll_TabMail.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge
        Me.BtnAutoAll_TabMail.Image = Global.Bid_Desk_Macros_v2.My.Resources.Resources.robot_pngrepo_com
        Me.BtnAutoAll_TabMail.Label = "Auto Process"
        Me.BtnAutoAll_TabMail.Name = "BtnAutoAll_TabMail"
        Me.BtnAutoAll_TabMail.ShowImage = True
        '
        'UnSortedMails
        '
        Me.UnSortedMails.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge
        Me.UnSortedMails.Image = Global.Bid_Desk_Macros_v2.My.Resources.Resources.Green_robot
        Me.UnSortedMails.Label = "Reprocess Not Defined"
        Me.UnSortedMails.Name = "UnSortedMails"
        Me.UnSortedMails.ShowImage = True
        '
        'UnsortedMails2
        '
        Me.UnsortedMails2.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge
        Me.UnsortedMails2.Image = Global.Bid_Desk_Macros_v2.My.Resources.Resources.Green_robot
        Me.UnsortedMails2.Label = "Reprocess Not Defined"
        Me.UnsortedMails2.Name = "UnsortedMails2"
        Me.UnsortedMails2.ShowImage = True
        '
        'MainRibbon
        '
        Me.Name = "MainRibbon"
        Me.RibbonType = "Microsoft.Outlook.Explorer"
        Me.Tabs.Add(Me.Tab1)
        Me.Tabs.Add(Tab2)
        Tab2.ResumeLayout(False)
        Tab2.PerformLayout()
        Me.Group5.ResumeLayout(False)
        Me.Group5.PerformLayout()
        Me.Tab1.ResumeLayout(False)
        Me.Tab1.PerformLayout()
        Me.Group1.ResumeLayout(False)
        Me.Group1.PerformLayout()
        Me.Group3.ResumeLayout(False)
        Me.Group3.PerformLayout()
        Me.Group2.ResumeLayout(False)
        Me.Group2.PerformLayout()
        Me.Group4.ResumeLayout(False)
        Me.Group4.PerformLayout()
        Me.ResumeLayout(False)

    End Sub

    Friend WithEvents Tab1 As Microsoft.Office.Tools.Ribbon.RibbonTab
    Friend WithEvents Group1 As Microsoft.Office.Tools.Ribbon.RibbonGroup
    Friend WithEvents MoveBtn As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents ReplyToBidBtn As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents ExpireButton As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents FwdDecision As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents FwdPrice As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents HPFwd As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents ExtensionBtn As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents WonBtn As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents DeadBtn As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents btnAutoAll As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents Group2 As Microsoft.Office.Tools.Ribbon.RibbonGroup
    Friend WithEvents BtnAddtoDB As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents ImprtLots As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents btnOnOff As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents Group3 As Microsoft.Office.Tools.Ribbon.RibbonGroup
    Friend WithEvents addOPG As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents btnLookup As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents BtnLater As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents btnChangeAM As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents MvAttach As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents btnHoliday As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents Group4 As Microsoft.Office.Tools.Ribbon.RibbonGroup
    Friend WithEvents Button1 As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents Group5 As Microsoft.Office.Tools.Ribbon.RibbonGroup
    Friend WithEvents BtnAutoAll_TabMail As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents Button2 As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents UnSortedMails As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents UnsortedMails2 As Microsoft.Office.Tools.Ribbon.RibbonButton
End Class

Partial Class ThisRibbonCollection

    <System.Diagnostics.DebuggerNonUserCode()> _
    Friend ReadOnly Property Ribbon1() As MainRibbon
        Get
            Return Me.GetRibbon(Of MainRibbon)()
        End Get
    End Property
End Class
