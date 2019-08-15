Imports System.IO
Imports Microsoft.Office.Tools.Ribbon

Public Class Ribbon1

    Private Sub Ribbon1_Load(ByVal sender As System.Object, ByVal e As RibbonUIEventArgs) Handles MyBase.Load

    End Sub

    Private Sub BtnLogin_Click(sender As Object, e As RibbonControlEventArgs) Handles BtnLogin.Click
        Dim frm As New BrowserController("Login")
        frm.Show()
    End Sub

    Private Sub NewDeal_Click(sender As Object, e As RibbonControlEventArgs) Handles NewDeal.Click
        Dim frm As New BrowserController("NewDeal")
        frm.Show()
    End Sub

    Private Sub Button1_Click(sender As Object, e As RibbonControlEventArgs) Handles BtnDLDeal.Click
        Dim frm As New BrowserController("DownloadQuote")
        frm.Show()
    End Sub
End Class
