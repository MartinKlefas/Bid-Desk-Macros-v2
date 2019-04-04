Imports String_Extensions.StringExtensions

Public Class NewMailForm
    Private entryIDCollection As String

    Public Sub New(entryIDCollection As String)
        Me.entryIDCollection = entryIDCollection
        Me.Label1.Text = "Determining the appropriate action for " & entryIDCollection.CountCharacter(",") & " new emails."
        InitializeComponent()

    End Sub

    Private Sub BackgroundWorker1_DoWork(sender As Object, e As ComponentModel.DoWorkEventArgs) Handles BackgroundWorker1.DoWork

        Dim msg As Outlook.MailItem

        For Each itemID In Split(entryIDCollection, ",")
            Dim item = Globals.ThisAddIn.Application.Session.GetItemFromID(itemID)
            If TypeName(item) = "MailItem" Then
                msg = item
            End If
        Next

    End Sub
End Class