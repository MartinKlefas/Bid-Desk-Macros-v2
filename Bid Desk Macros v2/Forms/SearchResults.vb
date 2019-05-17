Public Class SearchResults
    Private SearchIn As New Dictionary(Of String, Boolean) From {
        {"AM", False},
        {"Customer", False},
        {"DealID", False},
        {"OPGID", False},
        {"NDT", False}
        }
    Private SearchTerm As String

    Public Sub New(SearchTerm As String, AM As Boolean, Customer As Boolean, Deal As Boolean, OPG As Boolean, NDT As Boolean)
        Me.SearchTerm = SearchTerm
        SearchIn("AM") = AM
        SearchIn("Customer") = Customer
        SearchIn("DealID") = Deal
        SearchIn("OPGID") = OPG
        SearchIn("NDT") = NDT
    End Sub

    Private Sub BtnClose_Click(sender As Object, e As EventArgs) Handles btnClose.Click
        Me.Close()
    End Sub
End Class