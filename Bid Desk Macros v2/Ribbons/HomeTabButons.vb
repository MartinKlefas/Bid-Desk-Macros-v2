Imports Microsoft.Office.Tools.Ribbon

Public Class HomeTabButons

    Private Sub HomeTabButons_Load(ByVal sender As System.Object, ByVal e As RibbonUIEventArgs) Handles MyBase.Load

    End Sub

    Private Sub Button1_Click(sender As Object, e As RibbonControlEventArgs) Handles Button1.Click
        Dim Selection As Outlook.Selection = Globals.ThisAddIn.GetSelection()
        Dim MessageList As New List(Of Outlook.MailItem)

        For Each item In Selection
            If TypeName(item) = "MailItem" Then
                MessageList.Add(item)
            End If
        Next

        Dim autoForm As New NewMailForm(MessageList)
        autoForm.Show()

    End Sub
End Class
