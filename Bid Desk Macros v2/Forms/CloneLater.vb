Public Class CloneLater
    Private ReadOnly CurrentItem As Outlook.MailItem
    Private ReadOnly fullAutonomy As Boolean
    Private ReadOnly ReplyToThis As Outlook.MailItem
    Private ReadOnly DealID As String

    Public Sub New(targetDate As Date, email As Outlook.MailItem, Optional ReplyToThisMail As Outlook.MailItem = Nothing, Optional fullAutonomy As Boolean = False, Optional DealID As String = "")
        Me.InitializeComponent()
        Me.targetDate.SelectionStart = targetDate
        Me.targetDate.SelectionEnd = targetDate

        Me.CurrentItem = email
        Me.fullAutonomy = fullAutonomy
        Me.ReplyToThis = ReplyToThisMail
        Me.DealID = DealID
        If fullAutonomy Then

            Call Button1_Click(Nothing, Nothing)
        End If
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
        Dim msgReply As Outlook.MailItem

        If ReplyToThis Is Nothing Then
            'reply
            msgReply = CurrentItem.ReplyAll
        Else
            msgReply = ReplyToThis.ReplyAll
        End If

        msgReply.HTMLBody = CloneLaterMessage.Replace("%CLONEDATE%", Me.targetDate.SelectionEnd.ToShortDateString) & MainRibbon.WriteHolidayMessage() & msgReply.HTMLBody
        If fullAutonomy Then
            msgReply.Display()
        Else
            msgReply.Display()
        End If
        Me.Close()

        'set do not remind again flag
        Dim MessagesList As New List(Of Outlook.MailItem) From {
            CurrentItem
        }

        Dim DealIDForm As New DealIdent(MessagesList, "CloneLater", True, DealID:=Me.DealID)
        DealIDForm.Show()

    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Me.Close()
    End Sub


End Class