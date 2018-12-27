Imports System.Windows.Forms

Public Class DealIdent




    Private Sub DealID_KeyDown(sender As Object, e As KeyEventArgs) Handles DealID.KeyDown
        If e.KeyCode = Keys.Enter Then
            Button1_Click()
        End If
    End Sub

    Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
        Call Button2_Click()
    End Sub

    Private Sub Button1_Click() Handles OKButton.Click
        DealID.Text = Trim(DealID.Text)
        Me.DialogResult = DialogResult.OK
        Me.Hide()
    End Sub

    Private Sub Button2_Click() Handles Button2.Click
        DealID.Text = ""
        Me.DialogResult = DialogResult.Cancel
        Me.Hide()
    End Sub

    Private Sub DealID_MouseDown1(sender As Object, e As MouseEventArgs) Handles DealID.MouseDown
        If e.Button = MouseButtons.Right Then

            DealID.Text = My.Computer.Clipboard.GetText
        End If
    End Sub


End Class