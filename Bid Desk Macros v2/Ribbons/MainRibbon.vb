Imports System.Diagnostics
Imports Microsoft.Office.Tools.Ribbon
Imports String_Extensions

Public Class MainRibbon
    Public AutoInbound As Boolean
    Public Shared OnHoliday As Boolean

    Private Sub Ribbon1_Load(ByVal sender As System.Object, ByVal e As RibbonUIEventArgs) Handles MyBase.Load
        AutoInbound = False
        OnHoliday = False
    End Sub

    Public Shared Function WriteHolidayMessage() As String
        If OnHoliday Then
            Return HolidayMessage
        Else
            Return ""
        End If
    End Function

    Public Sub DisableRibbon()
        For Each control In Me.Group1.Items
            control.Enabled = False
        Next

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
            Dim senderEmail As String, ndt As String
            If Not msg.SenderEmailAddress.ToLower.Contains("@") AndAlso msg.SenderEmailAddress.ToLower.Contains("/") AndAlso msg.SenderEmailAddress.ToLower.Contains("recipients") Then
                senderEmail = msg.Sender.GetExchangeUser.PrimarySmtpAddress


            Else
                    senderEmail = msg.SenderEmailAddress
                If senderEmail.ToLower.Equals("tim.lee@insight.com") OrElse senderEmail.ToLower.Equals("richard.west@insight.com") Then
                    senderEmail = FindOnBehalfOf(msg.Body)
                End If
            End If

            If msg.Subject.StartsWith("[nextDesk]", ThisAddIn.searchType) Then
                ndt = msg.Subject.Substring(InStr(msg.Subject, "#"), 7)
            Else
                ndt = ""
            End If

            frmAddtoSql = New ImportDeal(senderEmail, ndt)
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
            btnOnOff.Label = "Automation Off"
        Else
            AutoInbound = True
            btnOnOff.Label = "Automation On"
            btnOnOff.Image = My.Resources._on
        End If
    End Sub

    Private Sub BtnHoliday_Click(sender As Object, e As RibbonControlEventArgs) Handles btnHoliday.Click
        If OnHoliday Then
            OnHoliday = False
            btnHoliday.Image = My.Resources.OfficeWork_Icon
            btnHoliday.Label = "At Work"
        Else
            OnHoliday = True
            btnHoliday.Label = "On Holiday"
            btnHoliday.Image = My.Resources.Vacation_Icon
        End If
    End Sub


    Private Sub AddOPG_Click(sender As Object, e As RibbonControlEventArgs) Handles addOPG.Click


        Call Globals.ThisAddIn.FindOPG()
    End Sub

    Private Sub BtnLookup_Click(sender As Object, e As RibbonControlEventArgs) Handles btnLookup.Click
        Dim frmSearch As New SearchForm
        frmSearch.Show()
    End Sub

    Private Sub BtnLater_Click(sender As Object, e As RibbonControlEventArgs) Handles BtnLater.Click
        Call Globals.ThisAddIn.CloneLater()
    End Sub

    Private Sub Button1_Click(sender As Object, e As RibbonControlEventArgs) Handles btnChangeAM.Click
        Dim frmNewAm As New ChangeAM
        frmNewAm.Show()
    End Sub

    Private Sub MvAttach_Click(sender As Object, e As RibbonControlEventArgs) Handles MvAttach.Click
        Globals.ThisAddIn.MoveBasedOnDealID(attach:=True)
    End Sub

    Private Sub Button1_Click_1(sender As Object, e As RibbonControlEventArgs) Handles Button1.Click
        Dim Selection As Outlook.Selection = Globals.ThisAddIn.GetSelection()
        Dim ndt As String = ""

        If Selection.Count = 1 AndAlso TypeName(Selection.Item(1)) = "MailItem" Then
            Dim msg As Outlook.MailItem = Selection.Item(1)

            If msg.Subject.StartsWith("[nextDesk]", ThisAddIn.searchType) Then
                ndt = msg.Subject.Substring(InStr(msg.Subject, "#"), 7)

            End If

        End If



        Dim ticketform As New TicketActions("MS-More-Info", ndt)
        ticketform.Show()
    End Sub



    Private Sub Button3_Click_1(sender As Object, e As RibbonControlEventArgs) Handles BtnAutoAll_TabMail.Click
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

    Private Function FindOnBehalfOf(MessageBody As String) As String

        Try
            Dim startPos As Integer = InStr(MessageBody, "Sales e-mail address")
            FindOnBehalfOf = Split(Mid(MessageBody, startPos + Len("Sales e-mail address") + 2), vbLf)(0)
        Catch ex As Exception
            FindOnBehalfOf = ""
        End Try

    End Function

End Class
