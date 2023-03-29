Imports Microsoft.Office.Tools.Ribbon

Public Class Ribbon1

    Private Sub Ribbon1_Load(ByVal sender As System.Object, ByVal e As RibbonUIEventArgs) Handles MyBase.Load

    End Sub

    Private Sub ReplyToBidBtn_Click(sender As Object, e As RibbonControlEventArgs) Handles ReplyToBidBtn.Click

        Dim obj As Object
        Dim msg As Outlook.MailItem

        If Globals.ThisAddIn.GetSelection().Count > 1 Then
            MsgBox("This can only be used with one bid request at a time")
            Exit Sub
        End If

        obj = Globals.ThisAddIn.GetCurrentItem()
        If obj IsNot Nothing AndAlso TypeName(obj) = "MailItem" Then
            msg = obj

            Dim newDeal As New AddDeal(msg)
            newDeal.Show()
        End If


    End Sub

    Private Sub ExtensionBtn_Click(sender As Object, e As RibbonControlEventArgs) Handles ExtensionBtn.Click
        Dim obj As Object
        Dim msg As Outlook.MailItem

        If Globals.ThisAddIn.GetSelection().Count > 1 Then
            MsgBox("This can only be used with one bid request at a time")
            Exit Sub
        End If

        obj = Globals.ThisAddIn.GetCurrentItem()
        If obj IsNot Nothing AndAlso TypeName(obj) = "MailItem" Then
            msg = obj

            Dim newDeal As New Extension(msg)
            newDeal.Show()
        End If
    End Sub
End Class
