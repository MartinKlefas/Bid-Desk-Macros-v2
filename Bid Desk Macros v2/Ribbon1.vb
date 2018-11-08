Imports System.Diagnostics
Imports Microsoft.Office.Tools.Ribbon

Public Class Ribbon1

    Private Sub Ribbon1_Load(ByVal sender As System.Object, ByVal e As RibbonUIEventArgs) Handles MyBase.Load

    End Sub

    Private Sub Button1_Click(sender As Object, e As RibbonControlEventArgs) Handles Button1.Click
        Dim sqlInterface As New clsDatabase(ThisAddIn.server, ThisAddIn.user,
                                   ThisAddIn.database, ThisAddIn.port)
        Dim tmp As String
        tmp = sqlInterface.SelectData("AM", "DealID = 16859207")
        Debug.WriteLine(tmp)
    End Sub
End Class
