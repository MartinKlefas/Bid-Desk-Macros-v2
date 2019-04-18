Imports System.Diagnostics
Imports Microsoft.Office.Tools.Ribbon

Public Class Ribbon1
    Public AutoInbound As Boolean

    Private Sub Ribbon1_Load(ByVal sender As System.Object, ByVal e As RibbonUIEventArgs) Handles MyBase.Load
        AutoInbound = False
    End Sub

    Public Sub DisableRibbon()
        For Each control In Me.Group1.Items
            control.Enabled = False
        Next

    End Sub

    Private Sub Button1_Click(sender As Object, e As RibbonControlEventArgs) Handles Button1.Click
        Dim dealForm As New DealIdent(Nothing, "")
        Debug.WriteLine(dealForm.ReadExcel("C:\Users\mklefass\Downloads\P001373741-01.xlsx", "Sheet1", 1, 1))
    End Sub

    Private Sub Button2_Click(sender As Object, e As RibbonControlEventArgs) Handles MoveBtn.Click
        Globals.ThisAddIn.MoveBasedOnDealID()
    End Sub

    Private Sub Button3_Click(sender As Object, e As RibbonControlEventArgs) Handles ReplyToBidBtn.Click
        Globals.ThisAddIn.ReplyToBidRequest()
    End Sub

    Private Sub ExpireButton_Click(sender As Object, e As RibbonControlEventArgs) Handles ExpireButton.Click
        Globals.ThisAddIn.ExpiryMessages()
    End Sub

    Private Sub FwdDecision_Click(sender As Object, e As RibbonControlEventArgs) Handles FwdDecision.Click
        Globals.ThisAddIn.FwdDRDecision()
    End Sub

    Private Sub FwdPrice_Click(sender As Object, e As RibbonControlEventArgs) Handles FwdPrice.Click
        Globals.ThisAddIn.FwdPricing()
    End Sub

    Private Sub HPFwd_Click(sender As Object, e As RibbonControlEventArgs) Handles HPFwd.Click
        Globals.ThisAddIn.FwdHPResponse()
    End Sub

    Private Sub ExtensionBtn_Click(sender As Object, e As RibbonControlEventArgs) Handles ExtensionBtn.Click
        Globals.ThisAddIn.ExtensionMessage()
    End Sub

    Private Sub WonBtn_Click(sender As Object, e As RibbonControlEventArgs) Handles WonBtn.Click
        Globals.ThisAddIn.MarkedWon()
    End Sub

    Private Sub DeadBtn_Click(sender As Object, e As RibbonControlEventArgs) Handles DeadBtn.Click
        Globals.ThisAddIn.MarkDead()

    End Sub

    Private Sub Button2_Click_1(sender As Object, e As RibbonControlEventArgs) Handles btnAutoAll.Click
        Dim Selection As Outlook.Selection = Globals.ThisAddIn.GetSelection()
        Dim MessageList As New List(Of Outlook.MailItem)

        For Each item In Selection
            If TypeName(item) = "MailItem" Then
                MessageList.Add(item)
            End If
        Next

        Dim autoForm As New NewMailForm(MessageList)
        autoForm.Show()

    End Sub

    Private Sub BtnAddtoDB_Click(sender As Object, e As RibbonControlEventArgs) Handles BtnAddtoDB.Click
        Dim Selection As Outlook.Selection = Globals.ThisAddIn.GetSelection()
        Dim frmAddtoSql As ImportDeal

        If Selection.Count = 1 AndAlso TypeName(Selection.Item(1)) = "MailItem" Then
            Dim msg As Outlook.MailItem = Selection.Item(1)
            Dim senderEmail As String
            If Not msg.SenderEmailAddress.ToLower.Contains("@") AndAlso msg.SenderEmailAddress.ToLower.Contains("/") AndAlso msg.SenderEmailAddress.ToLower.Contains("recipients") Then
                senderEmail = msg.Sender.GetExchangeUser.PrimarySmtpAddress
            Else
                senderEmail = msg.SenderEmailAddress

            End If
            frmAddtoSql = New ImportDeal(senderEmail)
        Else
            frmAddtoSql = New ImportDeal()
        End If
        frmAddtoSql.Show()

    End Sub

    Private Sub ImprtLots_Click(sender As Object, e As RibbonControlEventArgs) Handles ImprtLots.Click


        Dim frmBulkImport As New BulkImport
        frmBulkImport.Show()
    End Sub

    Private Sub BtnOnOff_Click(sender As Object, e As RibbonControlEventArgs) Handles btnOnOff.Click
        If AutoInbound Then
            AutoInbound = False
            btnOnOff.Image = My.Resources.off
        Else
            AutoInbound = True
            btnOnOff.Image = My.Resources._on
        End If
    End Sub


End Class
