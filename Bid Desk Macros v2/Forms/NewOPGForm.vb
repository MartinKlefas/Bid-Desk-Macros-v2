Public Class NewOPGForm


    Public Sub New(foundDealID As String)
        InitializeComponent()
        OPGBox.Text = foundDealID

    End Sub

    Private Sub OKButton_Click(sender As Object, e As EventArgs) Handles OKButton.Click
        If Globals.ThisAddIn.AddOPG(DealID.Text, OPGBox.Text) > 0 Then
            MsgBox("Successfully Updated")
        Else
            MsgBox("Some kind of error occurred")
        End If
        Me.Close()
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Me.Close()
    End Sub
End Class