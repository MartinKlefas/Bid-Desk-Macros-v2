Imports String_Extensions
Public Class SearchForm

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles btnCancel.Click
        Me.Close()
    End Sub

    Private Sub BtnSearch_Click(sender As Object, e As EventArgs) Handles BtnSearch.Click
        SearchTerm.Text = SearchTerm.Text.trimextended()
        Dim ResultsFrm As New SearchResults(SearchTerm.Text, ChkAM.Checked, ChkCustomer.Checked, ChkDeal.Checked, ChkOPG.Checked, CHKNDT.Checked)

        ResultsFrm.Show()
        Me.Close()
    End Sub
End Class