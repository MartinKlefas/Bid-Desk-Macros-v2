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
        Dim frm As New BrowserController("DownloadQuote", "6645831")
        frm.Show()
    End Sub

    Private Sub Button1_Click_1(sender As Object, e As RibbonControlEventArgs) Handles Button1.Click
        Dim frm As New FindCiscoAM("45378469")
        frm.Show()
    End Sub

    Private Sub Button2_Click(sender As Object, e As RibbonControlEventArgs) Handles Button2.Click
        Dim frm As New LenovoBrowserController("Login")
        frm.Show()
    End Sub
End Class
