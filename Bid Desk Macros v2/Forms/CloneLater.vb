Public Class CloneLater
    Private ReadOnly CurrentItem As Outlook.MailItem

    Public Sub New(targetDate As Date, email As Outlook.MailItem)
        Me.InitializeComponent()
        Me.targetDate.SelectionStart = targetDate
        Me.targetDate.SelectionEnd = targetDate

        Me.CurrentItem = email
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles btnSetReminder.Click

        'set reminder flag
        With CurrentItem
            .MarkAsTask(Microsoft.Office.Interop.Outlook.OlMarkInterval.olMarkNoDate)
            .TaskDueDate = Me.targetDate.SelectionEnd
            .ReminderSet = True
            .ReminderTime = Me.targetDate.SelectionStart
            .Save()
        End With

        'reply
        Dim msgReply As Outlook.MailItem = CurrentItem.ReplyAll

        msgReply.HTMLBody = CloneLaterMessage.Replace("%CLONEDATE%", Me.targetDate.SelectionEnd.ToShortDateString) & msgReply.HTMLBody

        msgReply.Display()

        Me.Close()

        'set do not remind again flag
        Dim MessagesList As New List(Of Outlook.MailItem) From {
            CurrentItem
        }

        Dim DealIDForm As New DealIdent(MessagesList, "CloneLater", True)
        DealIDForm.Show()

    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Me.Close()
    End Sub


End Class