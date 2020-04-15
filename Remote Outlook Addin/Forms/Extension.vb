Imports Microsoft.Office.Interop.Outlook

Public Class Extension
    Private ReadOnly msg As MailItem
    Private ReadOnly dealform As DealIdent

    Public Sub New(msg As MailItem)
        InitializeComponent()
        Me.msg = msg
        dealform = New DealIdent(msg, "", False)
        Me.txtDealId.Text = dealform.DealID.Text
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Me.Close()
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Dim action As String = ""
        If RadioButton1.Checked Then
            action = "Extended"
        End If
        If RadioButton2.Checked Then
            action = "Clone"
            Dim msgReply As Outlook.MailItem = Me.msg.ReplyAll

            msgReply.HTMLBody = CloneLaterMessage.Replace("%CLONEDATE%", "on the day of its' expiry") & msgReply.HTMLBody

            msgReply.Display()
        End If
        dealform.DealID.Text = Me.txtDealId.Text
        dealform.Mode = action
        Call dealform.Button1_Click()

        Me.Close()


    End Sub
End Class