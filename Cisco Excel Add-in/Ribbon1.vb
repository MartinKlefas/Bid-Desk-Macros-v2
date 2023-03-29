Imports System.IO
Imports Microsoft.Office.Tools.Ribbon
Imports String_Extensions

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

    Private Sub Button3_Click(sender As Object, e As RibbonControlEventArgs) Handles Button3.Click
        Dim frm As New LenovoBrowserController("ShowBid", "BBR-01516191")
        frm.Show()
    End Sub

    Private Sub Button4_Click(sender As Object, e As RibbonControlEventArgs) Handles Button4.Click
        Dim frm As New LenovoBrowserController("SendToDisti", "BBR-01516191")
        frm.Show()
    End Sub

    Private Sub Button5_Click(sender As Object, e As RibbonControlEventArgs) Handles Button5.Click

        Dim ws As Excel.Worksheet = Globals.ThisAddIn.Application.ActiveSheet

        Dim myData As List(Of Dictionary(Of String, String))

        Dim Headers As List(Of String)


        Dim mObj As Object = ReadTable(ws)
        myData = mObj(0)
        Headers = mObj(1)


        ws.Cells(10, 1) = ExcelFile(myData, Headers)

    End Sub

    Private Sub Button6_Click(sender As Object, e As RibbonControlEventArgs) Handles Button6.Click
        Dim testCell As Excel.Range = Globals.ThisAddIn.Application.ActiveSheet.cells(1, 1)

        testCell.Value = TrimNumbers(testCell.Value)
    End Sub
End Class
