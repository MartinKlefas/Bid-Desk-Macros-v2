Imports Microsoft.Office.Interop.Outlook
Imports String_Extensions.StringExtensions

Public Class NewMailForm
    Private entryIDCollection As String
    Private searchType As StringComparison = ThisAddIn.searchType

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
                If isExpiryNotice(msg) Then

                End If
                If isDRDecision(msg) Then

                End If
                If isPricing(msg) Then

                End If
            End If
        Next

    End Sub

    Private Function IsPricing(msg As MailItem) As Boolean
        Throw New NotImplementedException()
    End Function

    Private Function IsDRDecision(msg As MailItem) As Boolean
        If msg.SenderEmailAddress.Equals("", searchType) And msg.Subject.StartsWith("Opportunity", searchType) Then
            Return True
        ElseIf msg.Subject.StartsWith("Deal Registration REGE", searchType) And msg.Subject.endswith("review complete", searchtype) Then
            Return True
        ElseIf msg.Subject.StartsWith("Deal Registration REGI", searchtype) AndAlso msg.Body.tolower.Contains("the review for the deal registration") Then
            Return True
        Else
            Return False
        End If
    End Function

    Private Function IsExpiryNotice(newMail As Outlook.MailItem) As Boolean

        If newMail.Subject.StartsWith("Deal Registration", searchType) And newMail.Subject.EndsWith("Expiring", searchType) Then
            Return True
        ElseIf newMail.Subject.Equals("A Reminder that your Approved Deal is about to Expire", searchType) Then
            Return True
        ElseIf newMail.Subject.tolower.Contains("your quote expiration reminder mail") Then
            Return True
        Else
            Return False
        End If

    End Function
End Class