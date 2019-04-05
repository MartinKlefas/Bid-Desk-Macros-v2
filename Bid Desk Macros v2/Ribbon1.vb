Imports System.Diagnostics
Imports Microsoft.Office.Tools.Ribbon

Public Class Ribbon1

    Private Sub Ribbon1_Load(ByVal sender As System.Object, ByVal e As RibbonUIEventArgs) Handles MyBase.Load

    End Sub

    Private Sub Button1_Click(sender As Object, e As RibbonControlEventArgs) Handles Button1.Click
        Dim sqlInterface As New ClsDatabase(ThisAddIn.server, ThisAddIn.user,
                                   ThisAddIn.database, ThisAddIn.password)
        Dim tmp As String
        'tmp = sqlInterface.SelectData("Customer", "DealID = 'E002540241'")
        tmp = sqlInterface.ValueExists("E002540241")
        Debug.WriteLine(tmp)
        tmp = sqlInterface.ValueExists("E002540251")
        Debug.WriteLine(tmp)

    End Sub

    Private Sub Button2_Click(sender As Object, e As RibbonControlEventArgs) Handles MoveBtn.Click
        Globals.ThisAddIn.MoveBasedOnDealID()
    End Sub

    Private Sub Button3_Click(sender As Object, e As RibbonControlEventArgs) Handles ReplyToBidBtn.Click
        Globals.ThisAddIn.ReplyToBidRequest()
    End Sub

    Private Sub ExpireButton_Click(sender As Object, e As RibbonControlEventArgs) Handles ExpireButton.Click
        Globals.ThisAddIn.ExpiryMessages()
    End Sub

    Private Sub FwdDecision_Click(sender As Object, e As RibbonControlEventArgs) Handles FwdDecision.Click
        Globals.ThisAddIn.fwdDRDecision()
    End Sub

    Private Sub FwdPrice_Click(sender As Object, e As RibbonControlEventArgs) Handles FwdPrice.Click
        Globals.ThisAddIn.FwdPricing()
    End Sub

    Private Sub HPFwd_Click(sender As Object, e As RibbonControlEventArgs) Handles HPFwd.Click
        Globals.ThisAddIn.FwdHPResponse()
    End Sub

    Private Sub ExtensionBtn_Click(sender As Object, e As RibbonControlEventArgs) Handles ExtensionBtn.Click
        Globals.ThisAddIn.ExtensionMessage()
    End Sub

    Private Sub WonBtn_Click(sender As Object, e As RibbonControlEventArgs) Handles WonBtn.Click
        Globals.ThisAddIn.MarkedWon()
    End Sub

    Private Sub DeadBtn_Click(sender As Object, e As RibbonControlEventArgs) Handles DeadBtn.Click
        Globals.ThisAddIn.MarkDead()

    End Sub
End Class
