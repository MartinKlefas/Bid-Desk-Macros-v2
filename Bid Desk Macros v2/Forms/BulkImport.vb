Imports System.Diagnostics
Imports System.IO
Imports System.Windows.Forms
Imports String_Extensions

Public Class BulkImport
    Private Sub BulkImport_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Dim Columns As List(Of String) = Globals.ThisAddIn.sqlInterface.FindColumns()

        For Each column In Columns
            DataGridView1.Columns.Add(column, column)
        Next
    End Sub

    Private Sub LinkLabel1_LinkClicked(sender As Object, e As Windows.Forms.LinkLabelLinkClickedEventArgs) Handles LinkLabel1.LinkClicked
        Dim Columns As List(Of String) = Globals.ThisAddIn.sqlInterface.FindColumns()

        Dim txtLine As String = ""

        For Each column In Columns
            txtLine.Append(column & ",")
        Next

        txtLine = Strings.Left(txtLine, Len(txtLine) - 1)

        Dim sFileName As String = Path.GetTempPath() & "BulkImportTemplate.csv"
        System.IO.File.WriteAllText(sFileName, txtLine)
        Try
            System.Diagnostics.Process.Start(sFileName)

        Catch ex As Exception
            Debug.WriteLine("Unable to open file")
        End Try
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Me.Close()
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        For Each row As DataGridViewRow In DataGridView1.Rows
            Dim rowDict As New Dictionary(Of String, String)
            For Each cell As DataGridViewCell In row.Cells
                rowDict.Add(cell.OwningColumn.HeaderText, cell.Value)
            Next
            Globals.ThisAddIn.sqlInterface.Add_Data(rowDict)
        Next
    End Sub
End Class