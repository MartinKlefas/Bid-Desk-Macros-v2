Public Class SearchResults
    Private ReadOnly SearchIn As New Dictionary(Of String, Boolean) From {
        {"AM", False},
        {"Customer", False},
        {"DealID", False},
        {"OPGID", False},
        {"NDT", False}
        }
    Private ReadOnly SearchTerm As String

    Public Sub New(SearchTerm As String, AM As Boolean, Customer As Boolean, Deal As Boolean, OPG As Boolean, NDT As Boolean)
        InitializeComponent()
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

    Private Sub SearchResults_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        Dim Columns As List(Of String) = Globals.ThisAddIn.sqlInterface.FindColumns()

        For Each column In Columns
            DataGridView1.Columns.Add(column, column)
        Next

        Dim GlobalResults As New List(Of Dictionary(Of String, String))

        For Each column As String In SearchIn.Keys
            If SearchIn(column) Then
                Dim where As String = column & " LIKE '%" & SearchTerm & "%'"

                Dim results As New List(Of Dictionary(Of String, String))

                results = Globals.ThisAddIn.sqlInterface.SelectData_Dict("*", where)

                If results.Count > 0 Then
                    GlobalResults.AddRange(results)
                End If
            End If
        Next

        If GlobalResults.Count > 0 Then
            DataGridView1.Rows.Add(GlobalResults.Count)
            Dim i As Integer = 0
            For Each tResult As Dictionary(Of String, String) In GlobalResults
                For Each tkey In tResult.Keys
                    DataGridView1.Item(tkey, i).Value = tResult(tkey)
                Next
                i += 1
            Next
        Else
            DataGridView1.Rows.Add(1)
            DataGridView1.Item(0, 0).Value = "No results found"
        End If
    End Sub
End Class