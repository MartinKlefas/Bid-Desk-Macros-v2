﻿Imports Microsoft.Office.Tools.Ribbon

Public Class Ribbon1

    Private Sub Ribbon1_Load(ByVal sender As System.Object, ByVal e As RibbonUIEventArgs) Handles MyBase.Load

    End Sub

    Private Sub BtnLogin_Click(sender As Object, e As RibbonControlEventArgs) Handles BtnLogin.Click
        Dim frm As New BrowserController("Login")
        frm.Show()
    End Sub
End Class
