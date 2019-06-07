Public Class ChangeAM
    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Me.Close()
    End Sub

    Private Sub OKButton_Click(sender As Object, e As EventArgs) Handles OKButton.Click
        Dim changed As Integer = Globals.ThisAddIn.ChangeAM(OldAM.Text, NewAM.Text)

        If changed > 0 Then
            MsgBox(changed & " deals successfully updated")
        Else
            MsgBox("Some kind of error occurred")
        End If
        Me.Close()
    End Sub
End Class