Public Class CloneLater
    Private CurrentItem As Outlook.MailItem

    Public Sub New(targetDate As Date, email As Outlook.MailItem)
        Me.InitializeComponent()
        Me.targetDate.SelectionStart = targetDate
        Me.targetDate.SelectionEnd = targetDate

        Me.CurrentItem = email
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click

        'set reminder flag

        'reply

        'move

    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Me.Close()
    End Sub
End Class