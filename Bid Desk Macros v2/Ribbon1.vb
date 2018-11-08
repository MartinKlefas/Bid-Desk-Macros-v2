Imports Microsoft.Office.Tools.Ribbon

Public Class Ribbon1

    Private Sub Ribbon1_Load(ByVal sender As System.Object, ByVal e As RibbonUIEventArgs) Handles MyBase.Load

    End Sub

    Private Sub Button1_Click(sender As Object, e As RibbonControlEventArgs) Handles Button1.Click
        Dim sqlInterface As New clsDatabase(ThisAddIn.server, ThisAddIn.user,
                                   ThisAddIn.database, ThisAddIn.port)
        sqlInterface.SelectData("AM", "where DealID = 16859207")
    End Sub
End Class
